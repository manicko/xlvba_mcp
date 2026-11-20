"""Unit tests for the PivotTables / PivotCaches / PivotFields inventory (task_048).

These tests exercise the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` and
:class:`~excel_vba_mcp.adapters.excel.pivot_tables_adapter.PivotTablesAdapter`
against a pure-Python fake ``Workbook`` object graph (a fake STA worker that runs
the submitted callable inline). No Excel / COM is required.

Covered:

* ``pivot_tables_list`` worksheet enumeration across every sheet (the
  ``Workbook.PivotTables`` collection is unusable on the target build, so the
  sweep is per-worksheet);
* shared-cache identity — two PivotTables on one ``CacheIndex`` share a cache,
  reported once with ``affected_pivot_tables``;
* cache metadata — source type mapping (``xlDatabase`` -> ``database``,
  ``xlExternal`` -> ``external``, ``xlConsolidation`` -> ``consolidation``,
  ``xlScenario`` -> ``scenario``, ``xlPivotTable`` -> ``pivot_table``), OLAP
  detection, refresh metadata;
* field layout — orientation, 1-based position within orientation, data-field
  aggregation read via ``DataFields``, manual filter state;
* unsupported source — external/OLAP/consolidation/scenario/pivot_table sources
  report ``source_status == "unsupported_source"``;
* the service resolves ``workbook_id`` and submits ``READ_ONLY``;
* ``pivot_table_get`` / ``pivot_cache_get`` / ``pivot_fields_list`` resolution
  and error handling for a missing name/index.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import PivotTablesAdapter
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldInfo,
    PivotInventory,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotSourceType,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio

# ``XlPivotFieldOrientation`` / ``XlPivotTableSourceType`` / ``XlConsolidationFunction``
# values used to build the fakes (mirrors the adapter's own maps).
XL_HIDDEN = 0
XL_ROW = 1
XL_COLUMN = 2
XL_PAGE = 3
XL_DATA = 4

XL_SOURCE_DATABASE = 1
XL_SOURCE_EXTERNAL = 2
XL_SOURCE_CONSOLIDATION = 3
XL_SOURCE_SCENARIO = 4
XL_SOURCE_PIVOT_TABLE = -4148

XL_FUNC_SUM = -4157
XL_FUNC_AVERAGE = -4106
XL_FUNC_COUNT = -4112

# ``PivotField.Position`` sentinel for a hidden/unplaced field (#N/A).
POSITION_SENTINEL = -2146826246


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel PivotTables object graph
# ---------------------------------------------------------------------------
class FakePivotItem:
    def __init__(self, value: str, visible: bool = True) -> None:
        self.Value = value
        self.Visible = visible


class FakePivotItems:
    def __init__(self, items: list[FakePivotItem]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakePivotItem:  # noqa: N802 - COM casing
        return self._items[index - 1]


class FakePivotField:
    def __init__(
        self,
        source_name: str,
        *,
        caption: str | None = None,
        orientation: int = XL_HIDDEN,
        position: int | None = None,
        show_all_items: bool = False,
        all_items_visible: bool = True,
        items: list[FakePivotItem] | None = None,
    ) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = caption if caption is not None else source_name
        self.Orientation = orientation
        self.Position = POSITION_SENTINEL if position is None else position
        self.ShowAllItems = show_all_items
        self.AllItemsVisible = all_items_visible
        self._items = items or []

    def PivotItems(self) -> FakePivotItems:  # noqa: N802 - COM casing
        return FakePivotItems(self._items)


class FakeDataField:
    def __init__(
        self,
        source_name: str,
        caption: str,
        function: int,
        *,
        number_format: str = "General",
        position: int | None = None,
    ) -> None:
        self.SourceName = source_name
        self.Name = caption
        self.Function = function
        self.NumberFormat = number_format
        self.Position = POSITION_SENTINEL if position is None else position


class FakePivotFields:
    def __init__(self, fields: list[FakePivotField]) -> None:
        self._fields = fields

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> FakePivotField:  # noqa: N802 - COM casing
        return self._fields[index - 1]

    def __iter__(self):
        return iter(self._fields)


class FakeDataFields:
    def __init__(self, fields: list[FakeDataField]) -> None:
        self._fields = fields

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> FakeDataField:  # noqa: N802 - COM casing
        return self._fields[index - 1]


class FakePivotCache:
    def __init__(
        self,
        index: int,
        *,
        source_type: int = XL_SOURCE_DATABASE,
        source_data: Any = "Data!R1C1:R6C5",
        refresh_date: datetime | None = None,
        refresh_on_file_open: bool = False,
        enable_refresh: bool = True,
        record_count: int | None = 6,
        olap: bool = False,
    ) -> None:
        self.Index = index
        self.SourceType = source_type
        self.SourceData = source_data
        self.RefreshDate = refresh_date
        self.RefreshOnFileOpen = refresh_on_file_open
        self.EnableRefresh = enable_refresh
        self.RecordCount = record_count
        self.OLAP = olap


class FakeRange:
    def __init__(self, address: str) -> None:
        self.Address = address


class FakePivotTable:
    def __init__(
        self,
        name: str,
        sheet_name: str,
        cache: FakePivotCache,
        *,
        index: int = 1,
        table_range1: str = "$A$1:$B$7",
        table_range2: str = "$A$3:$B$7",
        data_body_range: str | None = "$B$5:$B$7",
        row_grand: bool | None = True,
        column_grand: bool | None = True,
        refresh_date: datetime | None = None,
        fields: list[FakePivotField] | None = None,
        data_fields: list[FakeDataField] | None = None,
    ) -> None:
        self.Name = name
        self._sheet_name = sheet_name
        self._cache = cache
        self.Index = index
        self.CacheIndex = cache.Index
        self.TableRange1 = FakeRange(table_range1)
        self.TableRange2 = FakeRange(table_range2)
        self.DataBodyRange = FakeRange(data_body_range) if data_body_range else None
        self.RowGrand = row_grand
        self.ColumnGrand = column_grand
        self.RefreshDate = refresh_date
        self._fields = fields or []
        self._data_fields = data_fields or []

    @property
    def Parent(self) -> FakeSheet:  # noqa: N802 - COM casing
        return FakeSheet(self._sheet_name)

    def PivotCache(self) -> FakePivotCache:  # noqa: N802 - COM casing
        return self._cache

    def PivotFields(self, *_args: Any) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields(self._fields)

    @property
    def DataFields(self) -> FakeDataFields:  # noqa: N802 - COM casing
        return FakeDataFields(self._data_fields)

    @property
    def RowFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self._fields if f.Orientation == XL_ROW])

    @property
    def ColumnFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self._fields if f.Orientation == XL_COLUMN])

    @property
    def PageFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self._fields if f.Orientation == XL_PAGE])


class FakePivotTables:
    def __init__(self, pivots: list[FakePivotTable]) -> None:
        self._pivots = pivots

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._pivots)

    def Item(self, index: int) -> FakePivotTable:  # noqa: N802 - COM casing
        return self._pivots[index - 1]

    def __iter__(self):
        return iter(self._pivots)


class FakeSheet:
    def __init__(self, name: str, pivots: list[FakePivotTable] | None = None) -> None:
        self.Name = name
        self._pivots = pivots or []

    def PivotTables(self) -> FakePivotTables:  # noqa: N802 - COM casing
        return FakePivotTables(self._pivots)


class FakeSheetCollection:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets

    def Item(self, key: Any) -> FakeSheet:  # noqa: N802 - COM casing
        if isinstance(key, int):
            return self._sheets[key - 1]
        for sheet in self._sheets:
            if sheet.Name == key:
                return sheet
        raise KeyError(key)

    def __iter__(self):
        return iter(self._sheets)


class FakeWorkbook:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets
        self.FullName = ""
        self.ReadOnly = False

    @property
    def Worksheets(self) -> FakeSheetCollection:  # noqa: N802 - COM casing
        return FakeSheetCollection(self._sheets)

    def sheet(self, name: str) -> FakeSheet:
        for sheet in self._sheets:
            if sheet.Name == name:
                return sheet
        raise KeyError(name)


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook([FakeSheet("Data")])
        wb.FullName = path
        wb.ReadOnly = bool(kwargs.get("ReadOnly", False))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)


class FakeApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.Workbooks = FakeWorkbooks()


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self) -> None:
        self.submits: list[dict[str, Any]] = []

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append(
            {"name": getattr(fn, "__name__", "?"), "section": section, "op_class": op_class}
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# ---------------------------------------------------------------------------
# Rig — real service stack, fake worker/Excel
# ---------------------------------------------------------------------------
class Rig:
    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.app = FakeApp()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = FakeWorker()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.service = PivotTablesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.audit,
        )

    async def open(self, *, read_only: bool = False) -> Any:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action="excel_open"),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )
        return await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=read_only
        )

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


# ---------------------------------------------------------------------------
# Fixture builders
# ---------------------------------------------------------------------------
def _shared_cache_pivots() -> tuple[FakePivotCache, list[FakePivotTable]]:
    """Two PivotTables (ptA, ptB) sharing cache 1 with a row/column/page/data layout."""

    cache = FakePivotCache(
        1,
        source_type=XL_SOURCE_DATABASE,
        source_data="Data!R1C1:R6C5",
        refresh_date=datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC),
        record_count=6,
    )
    pt_a_fields = [
        FakePivotField("Region", orientation=XL_ROW, position=1),
        FakePivotField("Product", orientation=XL_COLUMN, position=1),
        FakePivotField("Quarter", orientation=XL_PAGE, position=1),
        FakePivotField("Sales", orientation=XL_HIDDEN, position=None),
        FakePivotField("Units", orientation=XL_HIDDEN, position=None),
    ]
    pt_a_data = [
        FakeDataField("Sales", "Sum of Sales", XL_FUNC_SUM, number_format="#,##0.00", position=1),
        FakeDataField("Units", "Average of Units", XL_FUNC_AVERAGE, position=2),
    ]
    pt_a = FakePivotTable(
        "ptA",
        "PivotReport",
        cache,
        index=1,
        fields=pt_a_fields,
        data_fields=pt_a_data,
    )

    pt_b_fields = [
        FakePivotField("Region", orientation=XL_ROW, position=1),
        FakePivotField("Product", orientation=XL_HIDDEN, position=None),
        FakePivotField("Quarter", orientation=XL_HIDDEN, position=None),
        FakePivotField("Sales", orientation=XL_HIDDEN, position=None),
        FakePivotField(
            "Units",
            orientation=XL_HIDDEN,
            position=None,
            all_items_visible=False,
            items=[
                FakePivotItem("North", True),
                FakePivotItem("South", True),
                FakePivotItem("East", True),
                FakePivotItem("West", False),
            ],
        ),
    ]
    pt_b_data = [FakeDataField("Units", "Sum of Units", XL_FUNC_SUM, position=1)]
    pt_b = FakePivotTable(
        "ptB",
        "PivotReport",
        cache,
        index=2,
        fields=pt_b_fields,
        data_fields=pt_b_data,
    )
    return cache, [pt_a, pt_b]


def _listobject_cache_pivot() -> tuple[FakePivotCache, FakePivotTable]:
    """A third PivotTable (ptC) on a ListObject-backed cache 2 with a count agg."""

    cache = FakePivotCache(
        2,
        source_type=XL_SOURCE_DATABASE,
        source_data="TableData!R1C1:R6C5",
        record_count=6,
    )
    fields = [
        FakePivotField("Product", orientation=XL_ROW, position=1),
        FakePivotField("Region", orientation=XL_HIDDEN, position=None),
        FakePivotField("Sales", orientation=XL_HIDDEN, position=None),
        FakePivotField("Units", orientation=XL_HIDDEN, position=None),
        FakePivotField("Quarter", orientation=XL_HIDDEN, position=None),
    ]
    data = [FakeDataField("Sales", "Count of Sales", XL_FUNC_COUNT, position=1)]
    pt_c = FakePivotTable(
        "ptC",
        "PivotReport",
        cache,
        index=3,
        table_range1="$S$1:$T$5",
        table_range2="$S$3:$T$5",
        data_body_range="$T$4:$T$5",
        fields=fields,
        data_fields=data,
    )
    return cache, pt_c


def _seed_shared_workbook(wb: FakeWorkbook) -> None:
    """Attach ptA/ptB/ptC to the default workbook's sheets."""

    shared_cache, shared_pivots = _shared_cache_pivots()
    _cache_two, pt_c = _listobject_cache_pivot()
    data_sheet = FakeSheet("Data")
    report_sheet = FakeSheet("PivotReport", shared_pivots + [pt_c])
    wb._sheets = [data_sheet, report_sheet]


def _external_cache_pivot() -> FakePivotTable:
    """A PivotTable on an external (unsupported) source cache."""

    cache = FakePivotCache(
        1,
        source_type=XL_SOURCE_EXTERNAL,
        source_data=("SomeExternalArray",),
        record_count=100,
    )
    fields = [FakePivotField("Region", orientation=XL_ROW, position=1)]
    return FakePivotTable("ptExt", "PivotReport", cache, fields=fields)


def _olap_cache_pivot() -> FakePivotTable:
    """A PivotTable on an OLAP (unsupported) source cache."""

    cache = FakePivotCache(
        1,
        source_type=XL_SOURCE_EXTERNAL,
        source_data="OLAP$" * 3,
        olap=True,
        record_count=500,
    )
    fields = [FakePivotField("Region", orientation=XL_ROW, position=1)]
    return FakePivotTable("ptOlap", "PivotReport", cache, fields=fields)


# ---------------------------------------------------------------------------
# Adapter-level: worksheet enumeration + cache sharing + metadata
# ---------------------------------------------------------------------------
def test_enumerate_pivot_tables_sweeps_every_worksheet() -> None:
    shared_cache, shared_pivots = _shared_cache_pivots()
    _cache_two, pt_c = _listobject_cache_pivot()
    wb = FakeWorkbook(
        [
            FakeSheet("Data"),
            FakeSheet("PivotReport", shared_pivots + [pt_c]),
        ]
    )
    adapter = PivotTablesAdapter()

    tables = adapter.enumerate_pivot_tables(wb)

    # All three PivotTables are found, sorted by (sheet_index, name).
    assert [pt.name for pt in tables] == ["ptA", "ptB", "ptC"]
    assert all(pt.sheet_index == 2 for pt in tables)


def test_enumerate_pivot_tables_ignores_workbook_pivottables_collection() -> None:
    """The sweep must use ``ws.PivotTables()`` — ``Workbook.PivotTables`` is unusable."""

    shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    # A Workbook-level PivotTables collection that reports 0 (as the real build does)
    # must not affect the worksheet sweep.
    wb.PivotTables = FakePivotTables([])  # type: ignore[attr-defined]
    adapter = PivotTablesAdapter()

    tables = adapter.enumerate_pivot_tables(wb)
    assert [pt.name for pt in tables] == ["ptA", "ptB"]


def test_enumerate_pivot_tables_sorts_by_sheet_then_name() -> None:
    cache = FakePivotCache(1)
    sheet1 = FakeSheet(
        "Zeta",
        [
            FakePivotTable("ptZ", "Zeta", cache, index=1),
            FakePivotTable("ptA", "Zeta", cache, index=2),
        ],
    )
    sheet2 = FakeSheet("Alpha", [FakePivotTable("ptM", "Alpha", cache, index=3)])
    wb = FakeWorkbook([sheet1, sheet2])
    adapter = PivotTablesAdapter()

    tables = adapter.enumerate_pivot_tables(wb)
    # Sheet order (Zeta=1, Alpha=2) dominates; within a sheet the name breaks ties.
    assert [(pt.sheet_index, pt.name) for pt in tables] == [
        (1, "ptA"),
        (1, "ptZ"),
        (2, "ptM"),
    ]


def test_pivot_table_info_metadata_and_field_layout() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    pt_a = next(pt for pt in adapter.enumerate_pivot_tables(wb) if pt.name == "ptA")

    assert pt_a.sheet_name == "PivotReport"
    assert pt_a.cache_index == 1
    assert pt_a.table_range1 == "$A$1:$B$7"
    assert pt_a.table_range2 == "$A$3:$B$7"
    assert pt_a.data_body_range == "$B$5:$B$7"
    assert pt_a.row_grand is True
    assert pt_a.source_type is PivotSourceType.DATABASE
    assert pt_a.source_status == "supported"

    # Layout buckets.
    assert [f.source_name for f in pt_a.row_fields] == ["Region"]
    assert [f.source_name for f in pt_a.column_fields] == ["Product"]
    assert [f.source_name for f in pt_a.page_fields] == ["Quarter"]
    assert pt_a.hidden_fields == []

    # Data fields: orientation DATA, aggregation from DataFields, position.
    assert [f.source_name for f in pt_a.data_fields] == ["Sales", "Units"]
    sales = next(f for f in pt_a.data_fields if f.source_name == "Sales")
    assert sales.orientation is PivotFieldOrientation.DATA
    assert sales.function is PivotAggregation.SUM
    assert sales.data_field_name == "Sum of Sales"
    assert sales.number_format == "#,##0.00"
    assert sales.position == 1
    units = next(f for f in pt_a.data_fields if f.source_name == "Units")
    assert units.function is PivotAggregation.AVERAGE
    assert units.position == 2


def test_hidden_field_position_is_none() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    tables = adapter.enumerate_pivot_tables(wb)
    pt_a = next(pt for pt in tables if pt.name == "ptA")
    # A hidden source field carries no position.
    for field in pt_a.row_fields + pt_a.column_fields + pt_a.page_fields:
        assert field.position is not None


def test_manual_filter_state() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    fields = adapter.enumerate_pivot_fields(wb, "ptB")
    units = next(f for f in fields if f.source_name == "Units")
    assert units.all_items_visible is False
    assert units.filter_state is PivotFilterKind.MANUAL

    # A field with all items visible reports NONE.
    region = next(f for f in fields if f.source_name == "Region")
    assert region.all_items_visible is True
    assert region.filter_state is PivotFilterKind.NONE


# ---------------------------------------------------------------------------
# Adapter-level: cache sharing + cache metadata + source-type mapping
# ---------------------------------------------------------------------------
def test_shared_cache_identity() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    pivot_tables = adapter.enumerate_pivot_tables(wb)
    cache_one = adapter.read_pivot_cache(wb, 1)

    assert cache_one.cache_index == 1
    assert cache_one.shared_cache is True
    assert sorted(cache_one.affected_pivot_tables) == ["ptA", "ptB"]
    # ptA and ptB report the same cache index.
    assert {pt.cache_index for pt in pivot_tables if pt.name in ("ptA", "ptB")} == {1}


def test_cache_metadata_and_source_data() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    cache = adapter.read_pivot_cache(wb, 1)
    assert cache.source_type is PivotSourceType.DATABASE
    assert cache.source_data == "Data!R1C1:R6C5"
    assert cache.source_data_kind == "text"
    assert cache.refresh_date == datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC)
    assert cache.record_count == 6
    assert cache.olap is False


def test_source_type_mapping() -> None:
    """Each ``XlPivotTableSourceType`` maps to its ``PivotSourceType``."""

    cases = [
        (XL_SOURCE_DATABASE, PivotSourceType.DATABASE, "supported"),
        (XL_SOURCE_EXTERNAL, PivotSourceType.EXTERNAL, "unsupported_source"),
        (XL_SOURCE_CONSOLIDATION, PivotSourceType.CONSOLIDATION, "unsupported_source"),
        (XL_SOURCE_SCENARIO, PivotSourceType.SCENARIO, "unsupported_source"),
        (XL_SOURCE_PIVOT_TABLE, PivotSourceType.PIVOT_TABLE, "unsupported_source"),
    ]
    for source_type, expected_type, expected_status in cases:
        cache = FakePivotCache(1, source_type=source_type, source_data="x")
        pivot = FakePivotTable("pt", "PivotReport", cache)
        wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
        adapter = PivotTablesAdapter()

        table = adapter.enumerate_pivot_tables(wb)[0]
        assert table.source_type is expected_type, source_type
        assert table.source_status == expected_status, source_type


def test_unsupported_external_source_status() -> None:
    wb = FakeWorkbook([FakeSheet("PivotReport", [_external_cache_pivot()])])
    adapter = PivotTablesAdapter()

    table = adapter.enumerate_pivot_tables(wb)[0]
    assert table.source_type is PivotSourceType.EXTERNAL
    assert table.source_status == "unsupported_source"


def test_olap_source_detected_as_unsupported() -> None:
    wb = FakeWorkbook([FakeSheet("PivotReport", [_olap_cache_pivot()])])
    adapter = PivotTablesAdapter()

    table = adapter.enumerate_pivot_tables(wb)[0]
    assert table.source_type is PivotSourceType.OLAP
    assert table.source_status == "unsupported_source"


def test_external_source_data_array_kind() -> None:
    """An external source's Variant-array ``SourceData`` is surfaced as ``array``."""

    wb = FakeWorkbook([FakeSheet("PivotReport", [_external_cache_pivot()])])
    adapter = PivotTablesAdapter()

    cache = adapter.read_pivot_cache(wb, 1)
    assert cache.source_data_kind == "array"


# ---------------------------------------------------------------------------
# Adapter-level: field orientation / position / aggregation
# ---------------------------------------------------------------------------
def test_field_orientation_and_position() -> None:
    _shared_cache, shared_pivots = _shared_cache_pivots()
    wb = FakeWorkbook([FakeSheet("PivotReport", shared_pivots)])
    adapter = PivotTablesAdapter()

    fields = adapter.enumerate_pivot_fields(wb, "ptA")
    by_name = {f.source_name: f for f in fields}

    assert by_name["Region"].orientation is PivotFieldOrientation.ROW
    assert by_name["Region"].position == 1
    assert by_name["Product"].orientation is PivotFieldOrientation.COLUMN
    assert by_name["Quarter"].orientation is PivotFieldOrientation.PAGE
    # A data field is DATA orientation with its aggregation.
    assert by_name["Sales"].orientation is PivotFieldOrientation.DATA
    assert by_name["Sales"].function is PivotAggregation.SUM
    assert by_name["Sales"].is_data_field is True


def test_field_aggregation_codes() -> None:
    """Each accepted ``XlConsolidationFunction`` code maps to its aggregation."""

    cases = [
        (XL_FUNC_SUM, PivotAggregation.SUM),
        (XL_FUNC_AVERAGE, PivotAggregation.AVERAGE),
        (XL_FUNC_COUNT, PivotAggregation.COUNT),
    ]
    for function_code, expected in cases:
        cache = FakePivotCache(1)
        fields = [FakePivotField("Sales", orientation=XL_HIDDEN)]
        data = [FakeDataField("Sales", "Sum of Sales", function_code, position=1)]
        pivot = FakePivotTable("pt", "PivotReport", cache, fields=fields, data_fields=data)
        wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
        adapter = PivotTablesAdapter()

        result = adapter.enumerate_pivot_fields(wb, "pt")
        assert result[0].function is expected, function_code


def test_field_not_found_raises() -> None:
    wb = FakeWorkbook([FakeSheet("PivotReport", [])])
    adapter = PivotTablesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.enumerate_pivot_fields(wb, "missing")
    assert exc_info.value.context["pivot_table"] == "missing"


def test_cache_not_found_raises() -> None:
    wb = FakeWorkbook([FakeSheet("PivotReport", [])])
    adapter = PivotTablesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.read_pivot_cache(wb, 99)
    assert exc_info.value.context["cache_index"] == 99


# ---------------------------------------------------------------------------
# Service-level: read paths + op_class + workbook resolution
# ---------------------------------------------------------------------------
async def test_service_list_pivot_tables(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())

    inventory = await rig.service.list_pivot_tables(info.workbook_id)

    assert isinstance(inventory, PivotInventory)
    assert [pt.name for pt in inventory.pivot_tables] == ["ptA", "ptB", "ptC"]
    # Two caches: one shared (ptA/ptB), one for ptC.
    assert {c.cache_index for c in inventory.caches} == {1, 2}
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_tables_list")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_get_pivot_table(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())

    pt_a = await rig.service.get_pivot_table(info.workbook_id, "ptA")
    assert isinstance(pt_a, PivotTableInfo)
    assert pt_a.name == "ptA"
    assert pt_a.cache_index == 1
    assert pt_a.source_status == "supported"

    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_table_get")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_get_pivot_table_missing_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.get_pivot_table(info.workbook_id, "missing")
    assert exc_info.value.context["pivot_table"] == "missing"


async def test_service_get_pivot_cache(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())

    cache = await rig.service.get_pivot_cache(info.workbook_id, 1)
    assert isinstance(cache, PivotCacheInfo)
    assert cache.cache_index == 1
    assert cache.shared_cache is True
    assert sorted(cache.affected_pivot_tables) == ["ptA", "ptB"]

    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_cache_get")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_list_pivot_fields(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())

    fields = await rig.service.list_pivot_fields(info.workbook_id, "ptA")
    assert all(isinstance(f, PivotFieldInfo) for f in fields)
    assert {f.source_name for f in fields} >= {"Region", "Product", "Quarter", "Sales", "Units"}

    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_fields_list")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.list_pivot_tables("missing")


async def test_service_unsupported_source_passthrough(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    wb = rig.workbook()
    wb._sheets = [FakeSheet("PivotReport", [_external_cache_pivot()])]

    inventory = await rig.service.list_pivot_tables(info.workbook_id)
    assert len(inventory.pivot_tables) == 1
    assert inventory.pivot_tables[0].source_type is PivotSourceType.EXTERNAL
    assert inventory.pivot_tables[0].source_status == "unsupported_source"


# ---------------------------------------------------------------------------
# Tool-level: thin delegation + argument validation
# ---------------------------------------------------------------------------
async def test_tools_validate_and_delegate(tmp_path: Path) -> None:
    from excel_vba_mcp.api import pivot_tools

    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_shared_workbook(rig.workbook())
    pivot_tools.set_pivot_tables_service(rig.service)
    try:
        inventory = await pivot_tools.pivot_tables_list(info.workbook_id)
        assert len(inventory.pivot_tables) == 3

        pt_a = await pivot_tools.pivot_table_get(info.workbook_id, "ptA")
        assert pt_a.name == "ptA"

        cache = await pivot_tools.pivot_cache_get(info.workbook_id, 1)
        assert cache.cache_index == 1

        fields = await pivot_tools.pivot_fields_list(info.workbook_id, "ptA")
        assert fields
    finally:
        pivot_tools.set_pivot_tables_service(None)


async def test_tools_reject_empty_workbook_id() -> None:
    from excel_vba_mcp.api import pivot_tools

    with pytest.raises(ExcelVbaMcpError):
        await pivot_tools.pivot_tables_list("")


async def test_tools_reject_bad_cache_index() -> None:
    from excel_vba_mcp.api import pivot_tools

    rig_service = object()  # not reached: validation fires first
    pivot_tools.set_pivot_tables_service(rig_service)  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError):
            await pivot_tools.pivot_cache_get("wb", 0)
    finally:
        pivot_tools.set_pivot_tables_service(None)


def test_error_code_on_missing_pivot() -> None:
    wb = FakeWorkbook([FakeSheet("PivotReport", [])])
    adapter = PivotTablesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.enumerate_pivot_fields(wb, "nope")
    # A missing PivotTable is a lookup failure, not an unsupported source.
    assert exc_info.value.code is ErrorCode.COM_BUSY
