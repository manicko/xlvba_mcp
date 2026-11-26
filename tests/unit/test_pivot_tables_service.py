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

import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest
import pywintypes

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import (
    SHEET_STATUS_CAPPED,
    SHEET_STATUS_OK,
    SHEET_STATUS_TOO_LARGE,
    PivotTablesAdapter,
)
from excel_vba_mcp.application.pivot_tables_service import (
    PivotTablesService,
    _overall_sheet_status,
)
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
class _ReadCounter:
    """Block B spy: counts actual ``.Item()`` reads across the fake graph.

    ``pivot_fields`` / ``data_fields`` / ``pivot_items`` count the item reads;
    ``total`` is their sum. A tick-BEFORE-op budget stops before the op that would
    exceed it, so the count equals the budget; a tick-after-op budget would read
    one extra item (budget + 1).
    """

    def __init__(self) -> None:
        self.pivot_fields = 0
        self.data_fields = 0
        self.pivot_items = 0

    @property
    def total(self) -> int:
        return self.pivot_fields + self.data_fields + self.pivot_items


class FakePivotItem:
    def __init__(self, value: str, visible: bool = True) -> None:
        self.Value = value
        self.Visible = visible


class FakePivotItems:
    def __init__(self, items: list[FakePivotItem], counter: _ReadCounter | None = None) -> None:
        self._items = items
        self._counter = counter if counter is not None else _ReadCounter()

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakePivotItem:  # noqa: N802 - COM casing
        self._counter.pivot_items += 1
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
        read_counter: _ReadCounter | None = None,
    ) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = caption if caption is not None else source_name
        self.Orientation = orientation
        self.Position = POSITION_SENTINEL if position is None else position
        self.ShowAllItems = show_all_items
        self.AllItemsVisible = all_items_visible
        self._items = items or []
        self._counter = read_counter if read_counter is not None else _ReadCounter()
        # Block B spy: how many times ``PivotItems()`` was enumerated on this
        # field (the SourceType gate must keep this at 0 for unsupported pivots).
        self.pivot_items_calls = 0

    def PivotItems(self) -> FakePivotItems:  # noqa: N802 - COM casing
        self.pivot_items_calls += 1
        return FakePivotItems(self._items, self._counter)


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
    def __init__(
        self, fields: list[FakePivotField], counter: _ReadCounter | None = None
    ) -> None:
        self._fields = fields
        self._counter = counter if counter is not None else _ReadCounter()

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> FakePivotField:  # noqa: N802 - COM casing
        self._counter.pivot_fields += 1
        return self._fields[index - 1]

    def __iter__(self):
        return iter(self._fields)


class FakeDataFields:
    def __init__(
        self, fields: list[FakeDataField], counter: _ReadCounter | None = None
    ) -> None:
        self._fields = fields
        self._counter = counter if counter is not None else _ReadCounter()

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> FakeDataField:  # noqa: N802 - COM casing
        self._counter.data_fields += 1
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
        self._source_data = source_data
        # Block B spy: how many times the cache's source metadata was resolved
        # (``SourceData`` read). The ``CacheIndex`` memo resolves each distinct
        # cache once, so this equals the distinct-cache count, not the pivot count.
        self.metadata_reads = 0
        self.RefreshDate = refresh_date
        self.RefreshOnFileOpen = refresh_on_file_open
        self.EnableRefresh = enable_refresh
        self.RecordCount = record_count
        self.OLAP = olap

    @property
    def SourceData(self) -> Any:  # noqa: N802 - COM casing
        self.metadata_reads += 1
        return self._source_data

    @SourceData.setter
    def SourceData(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._source_data = value


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
        read_counter: _ReadCounter | None = None,
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
        self._counter = read_counter if read_counter is not None else _ReadCounter()
        # Block B spy: how many times ``PivotCache()`` was resolved on this table.
        self.pivot_cache_calls = 0

    @property
    def Parent(self) -> FakeSheet:  # noqa: N802 - COM casing
        return FakeSheet(self._sheet_name)

    def PivotCache(self) -> FakePivotCache:  # noqa: N802 - COM casing
        self.pivot_cache_calls += 1
        return self._cache

    def PivotFields(self, *_args: Any) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields(self._fields, self._counter)

    @property
    def DataFields(self) -> FakeDataFields:  # noqa: N802 - COM casing
        return FakeDataFields(self._data_fields, self._counter)

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
    def __init__(
        self,
        name: str,
        pivots: list[FakePivotTable] | None = None,
        *,
        used_range: str = "A1:B7",
    ) -> None:
        self.Name = name
        self._pivots = pivots or []
        # Block B: a large/small ``UsedRange`` the adapter must NEVER key on.
        self.used_range = used_range
        # Block B spy: how many times ``PivotTables()`` was called on this sheet
        # (one worksheet sweep == exactly one call per sheet).
        self.pivot_tables_calls = 0

    def PivotTables(self) -> FakePivotTables:  # noqa: N802 - COM casing
        self.pivot_tables_calls += 1
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
        # Block B: a file size the adapter must NEVER key on.
        self.file_size_bytes = 0

    @property
    def Worksheets(self) -> FakeSheetCollection:  # noqa: N802 - COM casing
        return FakeSheetCollection(self._sheets)

    def sheet(self, name: str) -> FakeSheet:
        for sheet in self._sheets:
            if sheet.Name == name:
                return sheet
        raise KeyError(name)

    def pivot_tables_sweep_calls(self) -> int:
        """Total ``sheet.PivotTables()`` calls across every sheet (Block B spy)."""

        return sum(sheet.pivot_tables_calls for sheet in self._sheets)


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


# ---------------------------------------------------------------------------
# Block B — single-sweep inventory + per-sheet measured budget (P6 gate)
# ---------------------------------------------------------------------------
def _many_field_pivot(
    name: str,
    sheet_name: str,
    cache: FakePivotCache,
    field_count: int,
    read_counter: _ReadCounter | None = None,
) -> FakePivotTable:
    """A PivotTable with ``field_count`` hidden source fields (a "big sheet")."""

    fields = [
        FakePivotField(f"f{index}", orientation=XL_HIDDEN, read_counter=read_counter)
        for index in range(field_count)
    ]
    return FakePivotTable(name, sheet_name, cache, fields=fields, read_counter=read_counter)


async def test_single_sweep_uses_one_worksheet_pass(tmp_path: Path) -> None:
    """The inventory is ONE worksheet sweep, not the former ``2 + N`` re-enumeration."""

    rig = Rig(tmp_path)
    info = await rig.open()
    wb = rig.workbook()
    _seed_shared_workbook(wb)  # 2 sheets (Data + PivotReport), 3 pivots, 2 caches
    sheet_count = len(wb._sheets)
    cache_count = 2

    inventory = await rig.service.list_pivot_tables(info.workbook_id)

    calls = wb.pivot_tables_sweep_calls()
    # One worksheet sweep == exactly one ``PivotTables()`` call per sheet.
    assert calls == sheet_count
    # Strictly fewer than the former ``2 + 2*caches`` re-enumeration.
    assert calls < 2 + 2 * cache_count
    # The inventory is still complete and correct.
    assert len(inventory.pivot_tables) == 3
    assert {cache.cache_index for cache in inventory.caches} == {1, 2}


def test_source_type_gate_skips_field_enumeration() -> None:
    """The SourceType gate runs FIRST: an unsupported pivot enumerates no fields/items."""

    ext_cache = FakePivotCache(1, source_type=XL_SOURCE_EXTERNAL, source_data=("x",))
    ext_field = FakePivotField("Region", orientation=XL_ROW, items=[FakePivotItem("A", True)])
    ext_pivot = FakePivotTable("ptExt", "PivotReport", ext_cache, fields=[ext_field])

    db_cache = FakePivotCache(2, source_type=XL_SOURCE_DATABASE)
    db_field = FakePivotField("Region", orientation=XL_ROW, items=[FakePivotItem("A", True)])
    db_pivot = FakePivotTable("ptDb", "PivotReport", db_cache, fields=[db_field])

    wb = FakeWorkbook([FakeSheet("PivotReport", [ext_pivot, db_pivot])])
    adapter = PivotTablesAdapter()

    sweep = adapter.inventory_single_sweep(wb, sheet_budget=None)

    by_name = {pt.name: pt for pt in sweep["pivot_tables"]}
    assert by_name["ptExt"].source_status == "unsupported_source"
    assert by_name["ptDb"].source_status == "supported"
    # Gate FIRST: the external pivot's field/items were never enumerated.
    assert ext_field.pivot_items_calls == 0
    # The supported pivot's field WAS enumerated (its items were read).
    assert db_field.pivot_items_calls > 0


def test_sheet_budget_caps_partial_results() -> None:
    """A sheet whose budget is hit after a completed PivotTable is ``capped`` (partial)."""

    cache = FakePivotCache(1)
    first = _many_field_pivot("pt1", "S", cache, 6)
    second = _many_field_pivot("pt2", "S", cache, 6)
    wb = FakeWorkbook([FakeSheet("S", [first, second])])
    adapter = PivotTablesAdapter()

    sweep = adapter.inventory_single_sweep(wb, sheet_budget=10)

    status = sweep["sheet_statuses"][0]
    assert status["status"] == SHEET_STATUS_CAPPED
    assert status["pivot_table_count"] == 1  # partial: the first pivot is kept
    assert [pt.name for pt in sweep["pivot_tables"]] == ["pt1"]
    # The second (too expensive) pivot contributed no fields.
    assert sweep["pivot_tables"][0].hidden_fields


def test_sheet_budget_too_large_names_sheet() -> None:
    """A sheet whose FIRST PivotTable blows the budget is ``too_large`` and names the sheet."""

    cache = FakePivotCache(1)
    big = _many_field_pivot("ptBig", "HugeSheet", cache, 50)
    wb = FakeWorkbook([FakeSheet("HugeSheet", [big])])
    adapter = PivotTablesAdapter()

    sweep = adapter.inventory_single_sweep(wb, sheet_budget=10)

    status = sweep["sheet_statuses"][0]
    assert status["status"] == SHEET_STATUS_TOO_LARGE
    assert status["sheet"] == "HugeSheet"  # names the offending sheet
    assert status["pivot_table_count"] == 0  # no partial
    assert sweep["pivot_tables"] == []


def test_capped_and_too_large_are_distinct() -> None:
    """``capped`` keeps partial results; ``too_large`` keeps none and names the sheet."""

    adapter = PivotTablesAdapter()
    cache = FakePivotCache(1)

    capped_sheet = FakeSheet(
        "S", [_many_field_pivot("a", "S", cache, 6), _many_field_pivot("b", "S", cache, 6)]
    )
    capped_wb = FakeWorkbook([capped_sheet])
    capped = adapter.inventory_single_sweep(capped_wb, sheet_budget=10)["sheet_statuses"][0]

    huge_wb = FakeWorkbook([FakeSheet("Big", [_many_field_pivot("a", "Big", cache, 50)])])
    huge = adapter.inventory_single_sweep(huge_wb, sheet_budget=10)["sheet_statuses"][0]

    assert capped["status"] == SHEET_STATUS_CAPPED
    assert capped["pivot_table_count"] > 0
    assert huge["status"] == SHEET_STATUS_TOO_LARGE
    assert huge["pivot_table_count"] == 0
    assert huge["sheet"] == "Big"
    assert capped["status"] != huge["status"]


def test_budget_not_keyed_on_file_size_rows_or_used_range() -> None:
    """The budget is MEASURED (field count), never file size / row count / UsedRange."""

    adapter = PivotTablesAdapter()
    cache = FakePivotCache(1)

    # Small file + tiny UsedRange, but a big sheet (many fields) -> too_large.
    small = FakeWorkbook(
        [FakeSheet("S", [_many_field_pivot("p", "S", cache, 50)], used_range="A1:B2")]
    )
    small.file_size_bytes = 10
    small_status = adapter.inventory_single_sweep(small, sheet_budget=10)["sheet_statuses"][0]

    # Big file + huge UsedRange / row count, but small pivots (1 field) -> ok.
    big = FakeWorkbook(
        [FakeSheet("S", [_many_field_pivot("p", "S", cache, 1)], used_range="A1:XFD1048576")]
    )
    big.file_size_bytes = 5_500_000
    big_status = adapter.inventory_single_sweep(big)["sheet_statuses"][0]

    assert small_status["status"] == SHEET_STATUS_TOO_LARGE
    assert big_status["status"] == SHEET_STATUS_OK
    # Orthogonal: the outcome tracks measured field count, not file/row/UsedRange.
    assert small_status["status"] != big_status["status"]


async def test_service_records_per_sheet_status_audit(tmp_path: Path) -> None:
    """``list_pivot_tables`` records the per-sheet sweep status in the audit trail."""

    rig = Rig(tmp_path)
    info = await rig.open()
    wb = rig.workbook()
    wb._sheets = [
        FakeSheet("Data"),
        FakeSheet("PivotReport", [_many_field_pivot("ptA", "PivotReport", FakePivotCache(1), 4)]),
    ]

    await rig.service.list_pivot_tables(info.workbook_id)

    entries = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if "pivot_tables_list" in line
    ]
    assert entries
    result = entries[-1]["result"]
    assert result["status"] == SHEET_STATUS_OK
    assert {sheet["sheet"] for sheet in result["sheets"]} == {"Data", "PivotReport"}
    assert all(sheet["status"] == SHEET_STATUS_OK for sheet in result["sheets"])
    # Block B criterion (c): the measured per-sheet cost is reported in the audit.
    costs = {sheet["sheet"]: sheet["cost"] for sheet in result["sheets"]}
    assert costs["Data"] == 0  # no PivotTables -> no measured reads
    assert costs["PivotReport"] == 4  # four field reads


async def test_service_too_large_names_sheet_in_audit(tmp_path: Path) -> None:
    """End-to-end: the default budget yields ``too_large`` naming the sheet, no partial."""

    rig = Rig(tmp_path)
    info = await rig.open()
    wb = rig.workbook()
    huge = _many_field_pivot("ptBig", "HugeSheet", FakePivotCache(1), 2500)
    wb._sheets = [FakeSheet("HugeSheet", [huge])]

    inventory = await rig.service.list_pivot_tables(info.workbook_id)

    # too_large: no partial PivotTable is returned.
    assert inventory.pivot_tables == []
    entries = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if "pivot_tables_list" in line
    ]
    result = entries[-1]["result"]
    assert result["status"] == SHEET_STATUS_TOO_LARGE
    assert result["sheets"][0]["sheet"] == "HugeSheet"
    assert result["sheets"][0]["status"] == SHEET_STATUS_TOO_LARGE
    # Default budget (2000) reached (tick-before-op); no partial.
    assert result["sheets"][0]["cost"] == 2000


# ---------------------------------------------------------------------------
# Block B — targeted follow-up tests (validator APPROVED; test-engineer request)
# ---------------------------------------------------------------------------
def test_budget_read_count_is_tick_before_op() -> None:
    """Criterion (c): the budget counts ACTUAL item reads and stops BEFORE the op.

    A single 50-field pivot with ``sheet_budget=10`` must read exactly 10
    ``PivotFields.Item`` calls (positions 1..10); the 11th tick raises **before**
    ``Item(11)``. A tick-after-op implementation would read 11 -> this test fails
    on that mutation, proving the budget never interrupts an in-flight call.
    """

    counter = _ReadCounter()
    cache = FakePivotCache(1)
    big = _many_field_pivot("ptBig", "S", cache, 50, read_counter=counter)
    wb = FakeWorkbook([FakeSheet("S", [big])])
    adapter = PivotTablesAdapter()

    sweep = adapter.inventory_single_sweep(wb, sheet_budget=10)

    status = sweep["sheet_statuses"][0]
    assert status["status"] == SHEET_STATUS_TOO_LARGE
    # tick-BEFORE-op: exactly 10 field reads; no data-field / item reads (no items).
    assert counter.pivot_fields == 10
    assert counter.data_fields == 0
    assert counter.pivot_items == 0
    assert counter.total == 10
    # The reported cost equals the reads performed (the raised tick did not count).
    assert status["cost"] == 10


def test_cache_index_memo_resolves_each_cache_once() -> None:
    """The ``CacheIndex`` memo resolves each distinct cache once, not once per pivot.

    Three PivotTables share two caches (ptA/ptB -> cache 1, ptC -> cache 2). The
    cache source metadata (``SourceData``) is read exactly twice (the distinct
    caches), never three times (the pivots). A deleted memo would re-resolve per
    pivot (3 reads) and fail this test. (``PivotCache()`` itself is also read once
    per pivot by the identity path, so the metadata resolution is the memo
    discriminator — the "re-resolve per pivot" the memo prevents.)
    """

    cache1 = FakePivotCache(1)
    cache2 = FakePivotCache(2)
    region = FakePivotField("Region", orientation=XL_ROW)
    product = FakePivotField("Product", orientation=XL_ROW)
    pt_a = FakePivotTable("ptA", "PivotReport", cache1, index=1, fields=[region])
    pt_b = FakePivotTable("ptB", "PivotReport", cache1, index=2, fields=[region])
    pt_c = FakePivotTable("ptC", "PivotReport", cache2, index=3, fields=[product])
    wb = FakeWorkbook([FakeSheet("PivotReport", [pt_a, pt_b, pt_c])])
    adapter = PivotTablesAdapter()

    adapter.inventory_single_sweep(wb, sheet_budget=None)

    # Each distinct cache resolved exactly once (memoized by CacheIndex).
    assert cache1.metadata_reads == 1
    assert cache2.metadata_reads == 1
    total_metadata_reads = cache1.metadata_reads + cache2.metadata_reads
    assert total_metadata_reads == 2  # distinct caches
    assert total_metadata_reads < 3  # NOT per-pivot (a deleted memo would be 3)


def test_sheet_status_reports_measured_cost() -> None:
    """Criterion (c): the per-sheet ``cost`` is the measured read count at stop."""

    adapter = PivotTablesAdapter()
    cache = FakePivotCache(1)

    # capped: 3 pivots x 3 fields, budget 6 -> stops before the 4th pivot, cost 6.
    capped_wb = FakeWorkbook(
        [
            FakeSheet(
                "S",
                [
                    _many_field_pivot("a", "S", cache, 3),
                    _many_field_pivot("b", "S", cache, 3),
                    _many_field_pivot("c", "S", cache, 3),
                ],
            )
        ]
    )
    capped = adapter.inventory_single_sweep(capped_wb, sheet_budget=6)["sheet_statuses"][0]
    assert capped["status"] == SHEET_STATUS_CAPPED
    assert capped["cost"] == 6
    assert capped["pivot_table_count"] == 2

    # too_large: 1 pivot x 50 fields, budget 10 -> cost 10, no partial.
    huge_wb = FakeWorkbook([FakeSheet("S", [_many_field_pivot("a", "S", cache, 50)])])
    huge = adapter.inventory_single_sweep(huge_wb, sheet_budget=10)["sheet_statuses"][0]
    assert huge["status"] == SHEET_STATUS_TOO_LARGE
    assert huge["cost"] == 10


def test_overall_sheet_status_mixed_severity() -> None:
    """``_overall_sheet_status`` reduces mixed severities (too_large > capped > ok)."""

    ok = {"status": SHEET_STATUS_OK}
    capped = {"status": SHEET_STATUS_CAPPED}
    too_large = {"status": SHEET_STATUS_TOO_LARGE}

    # ok + capped + too_large -> too_large (most severe), order-independent.
    assert _overall_sheet_status([ok, capped, too_large]) == SHEET_STATUS_TOO_LARGE
    assert _overall_sheet_status([too_large, ok, capped]) == SHEET_STATUS_TOO_LARGE
    # ok + capped -> capped.
    assert _overall_sheet_status([ok, capped]) == SHEET_STATUS_CAPPED
    assert _overall_sheet_status([capped, ok]) == SHEET_STATUS_CAPPED
    # all ok / empty -> ok.
    assert _overall_sheet_status([ok, ok]) == SHEET_STATUS_OK
    assert _overall_sheet_status([]) == SHEET_STATUS_OK


def test_none_meter_enumerates_all_fields() -> None:
    """A ``meter=None`` path (``enumerate_pivot_fields``) is unbounded: all fields."""

    cache = FakePivotCache(1)
    big = _many_field_pivot("ptBig", "S", cache, 2500)
    wb = FakeWorkbook([FakeSheet("S", [big])])
    adapter = PivotTablesAdapter()

    fields = adapter.enumerate_pivot_fields(wb, "ptBig")

    # No cap: every one of the 2500 fields is returned.
    assert len(fields) == 2500


async def test_none_meter_field_filters_returns(tmp_path: Path) -> None:
    """A ``meter=None`` path (``get_field_filters``) reads a field without a cap."""

    rig = Rig(tmp_path)
    info = await rig.open()
    wb = rig.workbook()
    big = _many_field_pivot("ptBig", "PivotReport", FakePivotCache(1), 2500)
    wb._sheets = [FakeSheet("PivotReport", [big])]

    filters = await rig.service.get_pivot_field_filters(info.workbook_id, "ptBig", "f0")

    assert filters.field == "f0"
    assert filters.filter.kind is PivotFilterKind.NONE


def test_refresh_uses_two_sweeps() -> None:
    """Pin current refresh behavior: ``refresh_pivot_table`` = 2 worksheet sweeps.

    ``_find_pivot`` (1 sweep) + ``enumerate_pivot_tables`` (1 sweep). This
    documents the refresh path is intentionally NOT the single-sweep inventory
    (that is ``list_pivot_tables``); it guards against an accidental change.
    """

    cache = FakePivotCache(1)
    pt = FakePivotTable(
        "ptA", "PivotReport", cache, fields=[FakePivotField("Region", orientation=XL_ROW)]
    )
    wb = FakeWorkbook([FakeSheet("PivotReport", [pt])])
    adapter = PivotTablesAdapter()

    adapter.refresh_pivot_table(wb, "ptA")

    assert wb.pivot_tables_sweep_calls() == 2


# ---------------------------------------------------------------------------
# Block B — external-cache SourceData 1004 (0x800A03EC) hardening
# ---------------------------------------------------------------------------
def _com_error_1004() -> pywintypes.com_error:
    """Build the exact ``com_error`` Excel raises for runtime 1004.

    Outer ``DISP_E_EXCEPTION`` (``0x80020009``) wrapping the inner Excel
    "Application-defined or object-defined error" (``0x800A03EC`` = 1004), as
    observed on the real ``report_w`` external cache's ``SourceData`` getter.
    """

    return pywintypes.com_error(
        -2147352567,  # 0x80020009 DISP_E_EXCEPTION
        "Application-defined or object-defined error.",
        (0, None, None, None, 0, -2146827284),  # EXCEPINFO inner = 0x800A03EC (1004)
        None,
    )


class _SourceData1004Cache(FakePivotCache):
    """A cache whose ``SourceData`` getter raises Excel 1004 (the traced failure)."""

    @property
    def SourceData(self) -> Any:  # noqa: N802 - COM casing
        self.metadata_reads += 1
        raise _com_error_1004()

    @SourceData.setter
    def SourceData(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._source_data = value


class _PivotCache1004Table(FakePivotTable):
    """A PivotTable whose structural ``PivotCache()`` identity call raises 1004."""

    def PivotCache(self) -> FakePivotCache:  # noqa: N802 - COM casing
        self.pivot_cache_calls += 1
        raise _com_error_1004()


def test_external_cache_source_data_1004_degrades_to_unavailable() -> None:
    """A cache whose ``SourceData`` getter raises 1004 must NOT abort the inventory.

    Reproduces the real ``report_w`` case: a workbook with a huge (566k-row) sheet
    that holds **0** PivotTables plus one external (``xlExternal``) cache whose
    ``SourceData`` getter raises ``DISP_E_EXCEPTION`` / 1004 (``0x800A03EC``). The
    single sweep must complete, report the external pivot ``unsupported_source``
    (SourceType gate first), and surface the cache with ``source_data=None`` /
    ``source_data_kind="unavailable"`` — never a generic failure.
    """

    cache = _SourceData1004Cache(1, source_type=XL_SOURCE_EXTERNAL, record_count=51760)
    ext_pivot = FakePivotTable(
        "PivotTable1",
        "clean_data_full",
        cache,
        fields=[FakePivotField("Region", orientation=XL_ROW)],
    )
    big_empty = FakeSheet("dashboard_data", [], used_range="A1:B2:566386")
    wb = FakeWorkbook([FakeSheet("clean_data_full", [ext_pivot]), big_empty])
    adapter = PivotTablesAdapter()

    sweep = adapter.inventory_single_sweep(wb)  # must NOT raise

    # The external pivot is still inventoried; the SourceType gate reports it
    # unsupported without enumerating its fields.
    assert [pt.name for pt in sweep["pivot_tables"]] == ["PivotTable1"]
    pivot = sweep["pivot_tables"][0]
    assert pivot.source_type is PivotSourceType.EXTERNAL
    assert pivot.source_status == "unsupported_source"

    # The cache degrades to "unavailable" instead of failing the whole sweep.
    assert len(sweep["caches"]) == 1
    cache_info = sweep["caches"][0]
    assert cache_info.cache_index == 1
    assert cache_info.source_type is PivotSourceType.EXTERNAL
    assert cache_info.source_data is None
    assert cache_info.source_data_kind == "unavailable"
    assert cache_info.record_count == 51760  # sibling metadata still read

    # Both sheets report ok — the 566k 0-pivot sheet included (no too_large/capped).
    statuses = {entry["sheet"]: entry for entry in sweep["sheet_statuses"]}
    assert statuses["clean_data_full"]["status"] == SHEET_STATUS_OK
    assert statuses["dashboard_data"]["status"] == SHEET_STATUS_OK
    assert statuses["dashboard_data"]["pivot_table_count"] == 0
    assert statuses["dashboard_data"]["cost"] == 0


async def test_service_survives_source_data_1004(tmp_path: Path) -> None:
    """End-to-end: ``list_pivot_tables`` returns the inventory (not COM_BUSY) on 1004."""

    rig = Rig(tmp_path)
    info = await rig.open()
    cache = _SourceData1004Cache(1, source_type=XL_SOURCE_EXTERNAL, record_count=51760)
    ext_pivot = FakePivotTable("PivotTable1", "clean_data_full", cache)
    rig.workbook()._sheets = [
        FakeSheet("clean_data_full", [ext_pivot]),
        FakeSheet("dashboard_data", [], used_range="A1:B2:566386"),
    ]

    inventory = await rig.service.list_pivot_tables(info.workbook_id)

    assert [pt.name for pt in inventory.pivot_tables] == ["PivotTable1"]
    assert inventory.pivot_tables[0].source_status == "unsupported_source"
    assert inventory.caches[0].source_data_kind == "unavailable"
    # The audit records an overall ok status (not a failure).
    entries = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if "pivot_tables_list" in line
    ]
    assert entries[-1]["result"]["status"] == SHEET_STATUS_OK


def test_generic_1004_on_structural_call_is_not_swallowed() -> None:
    """A 1004 from a structural identity call (``PivotCache()``) must still propagate.

    The SourceData hardening is deliberately narrow: only the **optional**
    ``SourceData`` read is guarded. A 1004 from the required ``PivotTable.
    PivotCache()`` identity call must NOT be silently converted into an empty
    inventory — otherwise a real, unexplained Excel error would be masked.
    """

    cache = FakePivotCache(1, source_type=XL_SOURCE_EXTERNAL)
    pivot = _PivotCache1004Table("pt", "PivotReport", cache)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(pywintypes.com_error):
        adapter.inventory_single_sweep(wb)
