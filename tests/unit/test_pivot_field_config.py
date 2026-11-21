"""Unit tests for the PivotField configuration surface (task_050).

These tests drive :class:`~excel_vba_mcp.adapters.excel.pivot_tables_adapter.PivotTablesAdapter`,
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` and
the ``pivot_field_set_*`` MCP tools against a **fake COM graph** — no Excel, no
COM, no ``pywin32``. They assert:

* **Orientation / layout** — only the layout buckets are assignable;
  ``orientation="data"`` is rejected because assigning ``xlDataField`` never
  creates a data field (that is ``AddDataField``); ``hidden`` removes a field
  non-destructively (``Orientation = 0``) while ``delete=True`` removes it
  destructively (``PivotField.Delete``);
* **Position** — 1-based, validated against the orientation's field count before
  any COM mutation; a hidden (unplaced) field is rejected; the resulting order
  is returned;
* **Aggregation** — an existing data field is updated in place
  (``DataFields[i].Function``), a missing one is added with ``AddDataField``;
  only the eleven accepted codes are allowed and legacy codes are rejected;
* **Captions** — a unique caption is written; a duplicate caption (another data
  field or a source field name) is rejected before COM;
* **Properties** — a sparse patch writes only the explicitly set fields; an
  empty patch, a blank caption/number format and a non-data field are rejected;
* **Field identity** — an unknown or ambiguous field is rejected;
* **Guards** — writable session → snapshot → approval, in that order, and every
  mutation is dispatched with ``OperationClass.MUTATING``;
* **Postcondition** — a fake that silently ignores the write is reported as
  ``verified=False`` with an explanatory note, never as success;
* **Unsupported source** — an OLAP / Data Model cache is rejected with
  ``UNSUPPORTED_PIVOT_SOURCE`` before any field is touched.

Run::

    uv run pytest tests/unit/test_pivot_field_config.py -q
"""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import (
    PivotTablesAdapter,
    validate_field_properties,
)
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotFieldInfo, PivotFieldProperties
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
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

XL_FUNC_SUM = -4157
XL_FUNC_AVERAGE = -4106
XL_FUNC_COUNT = -4112
XL_FUNC_MAX = -4136

# ``PivotField.Position`` sentinel for a hidden/unplaced field (#N/A).
POSITION_SENTINEL = -2146826246

_ACTION_OPEN = "excel_open"
_ACTION_LAYOUT = "pivot_field_set_layout"
_ACTION_POSITION = "pivot_field_set_position"
_ACTION_AGGREGATION = "pivot_field_set_aggregation"
_ACTION_PROPERTIES = "pivot_field_set_properties"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake COM graph (mutable: layout / position / aggregation / properties)
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
    """A mutable source field (orientation / position / items / delete).

    The fake emulates Excel's swap semantics: assigning ``Position`` exchanges
    places with the sibling that currently holds that position, and assigning
    ``Orientation = hidden`` resets ``Position`` to the ``#N/A`` sentinel.
    """

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
        ignore_writes: bool = False,
        owner: FakePivotTable | None = None,
    ) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = caption if caption is not None else source_name
        self._orientation = orientation
        self._position = POSITION_SENTINEL if position is None else position
        self.ShowAllItems = show_all_items
        self.AllItemsVisible = all_items_visible
        self._items = items or []
        self.deleted = False
        self._ignore_writes = ignore_writes
        self._owner = owner

    @property
    def Orientation(self) -> int:  # noqa: N802 - COM casing
        return self._orientation

    @Orientation.setter
    def Orientation(self, value: int) -> None:  # noqa: N802 - COM casing
        if self._ignore_writes:
            return
        self._orientation = value
        if value == XL_HIDDEN:
            self._position = POSITION_SENTINEL

    @property
    def Position(self) -> int:  # noqa: N802 - COM casing
        return self._position

    @Position.setter
    def Position(self, value: int) -> None:  # noqa: N802 - COM casing
        if self._ignore_writes:
            return
        previous = self._position
        self._position = value
        owner = self._owner
        if owner is None or value == POSITION_SENTINEL:
            return
        for sibling in owner.source_fields:
            if (
                sibling is not self
                and sibling.Orientation == self._orientation
                and sibling.Position == value
            ):
                sibling._position = previous

    def Delete(self) -> None:  # noqa: N802 - COM casing
        self.deleted = True

    def PivotItems(self) -> FakePivotItems:  # noqa: N802 - COM casing
        return FakePivotItems(self._items)


class FakeDataField:
    """A mutable data field (function / caption / number format / position)."""

    def __init__(
        self,
        source_name: str,
        caption: str,
        function: int,
        *,
        number_format: str = "General",
        position: int | None = None,
        ignore_writes: bool = False,
        owner: FakePivotTable | None = None,
    ) -> None:
        self.SourceName = source_name
        self._caption = caption
        self.Function = function
        self.NumberFormat = number_format
        self._position = POSITION_SENTINEL if position is None else position
        self._ignore_writes = ignore_writes
        self._owner = owner

    @property
    def Name(self) -> str:  # noqa: N802 - COM casing
        return self._caption

    @Name.setter
    def Name(self, value: str) -> None:  # noqa: N802 - COM casing
        if self._ignore_writes:
            return
        self._caption = value

    @property
    def Caption(self) -> str:  # noqa: N802 - COM casing
        return self._caption

    @Caption.setter
    def Caption(self, value: str) -> None:  # noqa: N802 - COM casing
        if self._ignore_writes:
            return
        self._caption = value

    @property
    def Position(self) -> int:  # noqa: N802 - COM casing
        return self._position

    @Position.setter
    def Position(self, value: int) -> None:  # noqa: N802 - COM casing
        if self._ignore_writes:
            return
        previous = self._position
        self._position = value
        owner = self._owner
        if owner is None or value == POSITION_SENTINEL:
            return
        for sibling in owner.data_fields:
            if sibling is not self and sibling.Position == value:
                sibling._position = previous

    def __setattr__(self, name: str, value: Any) -> None:
        if name in {"Function", "NumberFormat"} and getattr(self, "_ignore_writes", False):
            return
        object.__setattr__(self, name, value)


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
        index: int = 1,
        *,
        source_type: int = XL_SOURCE_DATABASE,
        olap: bool = False,
    ) -> None:
        self.Index = index
        self.SourceType = source_type
        self.SourceData = "Data!R1C1:R6C5"
        self.RefreshDate = datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC)
        self.OLAP = olap
        self.RecordCount = 6


class FakeRange:
    def __init__(self, address: str) -> None:
        self.Address = address


class FakePivotTable:
    """A mutable PivotTable (fields + data fields + AddDataField/Delete)."""

    def __init__(
        self,
        name: str = "ptA",
        *,
        cache: FakePivotCache | None = None,
        fields: list[FakePivotField] | None = None,
        data_fields: list[FakeDataField] | None = None,
    ) -> None:
        self.Name = name
        self.Index = 1
        self._cache = cache or FakePivotCache()
        self.CacheIndex = self._cache.Index
        self.TableRange1 = FakeRange("$A$1:$B$7")
        self.TableRange2 = FakeRange("$A$3:$B$7")
        self.DataBodyRange = FakeRange("$B$5:$B$7")
        self.RowGrand = True
        self.ColumnGrand = True
        self.RefreshDate = None
        self._fields = fields or []
        self._data_fields = data_fields or []
        for field in self._fields:
            field._owner = self
        for data_field in self._data_fields:
            data_field._owner = self

    @property
    def source_fields(self) -> list[FakePivotField]:
        return [f for f in self._fields if not f.deleted]

    @property
    def data_fields(self) -> list[FakeDataField]:
        return list(self._data_fields)

    def PivotCache(self) -> FakePivotCache:  # noqa: N802 - COM casing
        return self._cache

    def PivotFields(self, *_args: Any) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields(self.source_fields)

    @property
    def DataFields(self) -> FakeDataFields:  # noqa: N802 - COM casing
        return FakeDataFields(self._data_fields)

    @property
    def RowFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self.source_fields if f.Orientation == XL_ROW])

    @property
    def ColumnFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self.source_fields if f.Orientation == XL_COLUMN])

    @property
    def PageFields(self) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields([f for f in self.source_fields if f.Orientation == XL_PAGE])

    def AddDataField(  # noqa: N802 - COM casing
        self,
        field: FakePivotField,
        caption: str | None = None,
        function: int | None = None,
    ) -> None:
        text = caption if caption else f"Sum of {field.SourceName}"
        existing = {entry.Name.strip().lower() for entry in self._data_fields}
        source_names = {f.SourceName.strip().lower() for f in self._fields}
        # Excel rejects a data-field caption that duplicates another data field or
        # a source field name.
        if text.strip().lower() in existing or text.strip().lower() in source_names:
            raise RuntimeError("PivotTable field name already exists")
        self._data_fields.append(
            FakeDataField(
                field.SourceName,
                text,
                function if function is not None else XL_FUNC_SUM,
                position=len(self._data_fields) + 1,
                owner=self,
            )
        )


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


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook([FakeSheet("PivotReport")])
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
    """Fake STA worker: runs the submitted callable inline (no thread, no COM)."""

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
            {
                "name": getattr(fn, "__name__", "?"),
                "section": section,
                "op_class": op_class,
            }
        )
        return fn()

    def get_status(self) -> Any:
        return WorkbookSessionState.RUNNING


class FakeSnapshotGate:
    """Minimal snapshot gate: only an explicitly created snapshot satisfies it."""

    def __init__(self, snapshots: set[str] | None = None) -> None:
        self.snapshots = snapshots if snapshots is not None else set()

    def require_snapshot(self, workbook_id: str) -> None:
        if workbook_id not in self.snapshots:
            raise ExcelVbaMcpError(
                "Workbook has no snapshot.",
                code=ErrorCode.SNAPSHOT_FAILED,
                context={"workbook_id": workbook_id},
            )


# ---------------------------------------------------------------------------
# Rig — real service stack over the fake worker/Excel
# ---------------------------------------------------------------------------
class Rig:
    def __init__(self, tmp_path: Path, *, with_snapshot: bool = False) -> None:
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
        self.snapshots = {with_snapshot}
        self.pivots = PivotTablesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.audit,
            snapshot_service=FakeSnapshotGate(self._existing_snapshots()),
            approval=self.approval,
        )

    def _existing_snapshots(self) -> set[str]:
        return set()

    async def open(self, *, read_only: bool = False) -> Any:
        self.grant(_ACTION_OPEN)
        return await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=read_only
        )

    def grant(self, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


# ---------------------------------------------------------------------------
# Fake graph builders
# ---------------------------------------------------------------------------
def _pt_a() -> FakePivotTable:
    """ptA-like layout: Region(row), Product(column), Quarter(page), 2 data fields."""

    fields = [
        FakePivotField("Region", orientation=XL_ROW, position=1),
        FakePivotField("Product", orientation=XL_COLUMN, position=1),
        FakePivotField("Quarter", orientation=XL_PAGE, position=1),
        FakePivotField("Sales", orientation=XL_HIDDEN, position=None),
        FakePivotField("Units", orientation=XL_HIDDEN, position=None),
    ]
    data_fields = [
        FakeDataField("Sales", "Sum of Sales", XL_FUNC_SUM, position=1),
        FakeDataField("Units", "Average of Units", XL_FUNC_AVERAGE, position=2),
    ]
    return FakePivotTable("ptA", fields=fields, data_fields=data_fields)


def _two_row_pivot() -> FakePivotTable:
    """A pivot with two row fields, used to verify position swap semantics."""

    fields = [
        FakePivotField("Region", orientation=XL_ROW, position=1),
        FakePivotField("Quarter", orientation=XL_ROW, position=2),
    ]
    return FakePivotTable("ptRows", fields=fields)


def _seed(rig: Rig, pivot: FakePivotTable) -> FakePivotTable:
    """Install ``pivot`` on the (single) PivotReport sheet of the open workbook."""

    wb = rig.workbook()
    sheet = wb.Worksheets.Item("PivotReport")
    sheet._pivots = [pivot]
    return pivot


def _olap_pivot() -> FakePivotTable:
    """A pivot whose cache is an OLAP / Data Model source (unsupported)."""

    cache = FakePivotCache(1, source_type=XL_SOURCE_DATABASE, olap=True)
    return FakePivotTable(
        "ptOlap", cache=cache, fields=[FakePivotField("Region", orientation=XL_ROW, position=1)]
    )


# ---------------------------------------------------------------------------
# Adapter: orientation / layout
# ---------------------------------------------------------------------------
def test_set_layout_moves_field_to_column() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_layout(wb, "ptA", "Region", PivotFieldOrientation.COLUMN)

    region = next(f for f in pivot._fields if f.SourceName == "Region")
    assert region.Orientation == XL_COLUMN
    assert result["verified"] is True
    assert result["notes"] == []
    # The resulting column order is Region then Product (swap semantics).
    assert result["orientation_order"] == ["Region", "Product"]
    assert isinstance(result["field"], PivotFieldInfo)
    assert result["field"].orientation is PivotFieldOrientation.COLUMN
    # The readback agrees with the inventory read.
    described = {f.source_name: f for f in adapter.enumerate_pivot_fields(wb, "ptA")}
    assert described["Region"].orientation is PivotFieldOrientation.COLUMN


def test_set_layout_rejects_data_orientation() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_layout(wb, "ptA", "Units", PivotFieldOrientation.DATA)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_orientation"
    # The field is untouched: no data field was created.
    units = next(f for f in pivot._fields if f.SourceName == "Units")
    assert units.Orientation == XL_HIDDEN
    assert len(pivot._data_fields) == 2


def test_set_layout_hidden_is_non_destructive() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_layout(wb, "ptA", "Product", PivotFieldOrientation.HIDDEN)

    product = next(f for f in pivot._fields if f.SourceName == "Product")
    assert product.Orientation == XL_HIDDEN
    assert product.deleted is False
    # The field is still present in PivotFields (non-destructive).
    assert any(f.SourceName == "Product" for f in pivot.PivotFields())
    assert result["verified"] is True
    assert result["field"].orientation is PivotFieldOrientation.HIDDEN
    assert result["field"].position is None


def test_set_layout_delete_is_destructive() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_layout(
        wb, "ptA", "Product", PivotFieldOrientation.HIDDEN, delete=True
    )

    assert not any(f.SourceName == "Product" for f in pivot.PivotFields())
    assert result["verified"] is True
    # The DTO describes the last-known state as hidden/unplaced.
    assert result["field"].orientation is PivotFieldOrientation.HIDDEN
    assert result["field"].position is None


def test_set_layout_delete_requires_hidden_orientation() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_layout(wb, "ptA", "Product", PivotFieldOrientation.ROW, delete=True)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_orientation"


def test_set_layout_rejects_olap_source() -> None:
    pivot = _olap_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_layout(wb, "ptOlap", "Region", PivotFieldOrientation.ROW)

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE


def test_set_layout_rejects_unknown_field() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_layout(wb, "ptA", "Nope", PivotFieldOrientation.ROW)

    assert exc_info.value.context["reason"] == "field_not_found"


# ---------------------------------------------------------------------------
# Adapter: position
# ---------------------------------------------------------------------------
def test_set_position_swaps_within_orientation() -> None:
    pivot = _two_row_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_position(wb, "ptRows", "Quarter", 1)

    quarter = next(f for f in pivot._fields if f.SourceName == "Quarter")
    assert quarter.Position == 1
    assert result["verified"] is True
    assert result["orientation_order"] == ["Quarter", "Region"]
    assert result["field"].position == 1


def test_set_position_out_of_range_rejected_before_write() -> None:
    pivot = _two_row_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_position(wb, "ptRows", "Quarter", 3)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "position_out_of_range"
    assert exc_info.value.context["count"] == 2
    # Nothing changed.
    quarter = next(f for f in pivot._fields if f.SourceName == "Quarter")
    assert quarter.Position == 2


def test_set_position_rejects_hidden_field() -> None:
    adapter = PivotTablesAdapter()
    # A genuinely hidden field with no data field has no position to set.
    pivot2 = FakePivotTable(
        "ptHidden",
        fields=[FakePivotField("Notes", orientation=XL_HIDDEN, position=None)],
        data_fields=[],
    )
    wb2 = FakeWorkbook([FakeSheet("PivotReport", [pivot2])])
    with pytest.raises(ExcelVbaMcpError) as exc_info2:
        adapter.set_field_position(wb2, "ptHidden", "Notes", 1)
    assert exc_info2.value.context["reason"] == "field_not_placed"


def test_set_position_on_data_field() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_position(wb, "ptA", "Units", 1)

    units_data = pivot._data_fields[1]
    assert units_data.Position == 1
    assert result["verified"] is True
    # Data-field order is reported by caption (the data-field identity).
    assert result["orientation_order"] == ["Average of Units", "Sum of Sales"]


def test_set_position_reports_unverified_on_silent_failure() -> None:
    pivot = _two_row_pivot()
    # A fake that ignores writes: the postcondition must catch it.
    for field in pivot._fields:
        field._ignore_writes = True
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_position(wb, "ptRows", "Quarter", 1)

    assert result["verified"] is False
    assert any("Postcondition mismatch" in note for note in result["notes"])
    assert result["field"].position == 2  # unchanged


# ---------------------------------------------------------------------------
# Adapter: aggregation + captions
# ---------------------------------------------------------------------------
def test_set_aggregation_updates_existing_data_field() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_aggregation(wb, "ptA", "Sales", PivotAggregation.MAX)

    assert pivot._data_fields[0].Function == XL_FUNC_MAX
    assert result["verified"] is True
    assert result["field"].function is PivotAggregation.MAX
    assert len(pivot._data_fields) == 2  # updated in place, not added


def test_set_aggregation_adds_data_field_via_add_data_field() -> None:
    pivot = FakePivotTable(
        "ptAdd",
        fields=[
            FakePivotField("Region", orientation=XL_ROW, position=1),
            FakePivotField("Units", orientation=XL_HIDDEN, position=None),
        ],
        data_fields=[],
    )
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_aggregation(
        wb, "ptAdd", "Units", PivotAggregation.COUNT, caption="Count of Units"
    )

    assert len(pivot._data_fields) == 1
    assert pivot._data_fields[0].SourceName == "Units"
    assert pivot._data_fields[0].Name == "Count of Units"
    assert pivot._data_fields[0].Function == XL_FUNC_COUNT
    assert result["verified"] is True
    # The source field's own orientation is untouched (no Orientation = 4).
    units = next(f for f in pivot._fields if f.SourceName == "Units")
    assert units.Orientation == XL_HIDDEN


def test_set_aggregation_rejects_legacy_function_code() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_aggregation(wb, "ptA", "Sales", -4114)  # type: ignore[arg-type]

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_function"
    assert pivot._data_fields[0].Function == XL_FUNC_SUM  # untouched


def test_set_aggregation_duplicate_caption_rejected() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    # "Sum of Sales" already belongs to the Sales data field; asking for it on a
    # different source field must be rejected, not silently applied to Sales.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_aggregation(
            wb, "ptA", "Units", PivotAggregation.COUNT, caption="Sum of Sales"
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "duplicate_caption"
    assert pivot._data_fields[0].Function == XL_FUNC_SUM


def test_set_aggregation_caption_equal_to_source_name_rejected() -> None:
    pivot = FakePivotTable(
        "ptAdd",
        fields=[
            FakePivotField("Region", orientation=XL_ROW, position=1),
            FakePivotField("Units", orientation=XL_HIDDEN, position=None),
        ],
        data_fields=[],
    )
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_aggregation(
            wb, "ptAdd", "Units", PivotAggregation.COUNT, caption="Region"
        )

    assert exc_info.value.context["reason"] == "duplicate_caption"
    assert pivot._data_fields == []


def test_set_aggregation_ambiguous_without_caption() -> None:
    pivot = FakePivotTable(
        "ptMulti",
        fields=[
            FakePivotField("Region", orientation=XL_ROW, position=1),
            FakePivotField("Units", orientation=XL_HIDDEN, position=None),
        ],
        data_fields=[
            FakeDataField("Units", "Sum of Units", XL_FUNC_SUM, position=1),
            FakeDataField("Units", "Average of Units", XL_FUNC_AVERAGE, position=2),
        ],
    )
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_aggregation(wb, "ptMulti", "Units", PivotAggregation.MAX)

    assert exc_info.value.context["reason"] == "ambiguous_data_field"


def test_set_aggregation_addmaps_duplicate_to_structured_error() -> None:
    """A duplicate caption that only Excel can reject is surfaced truthfully."""

    pivot = FakePivotTable(
        "ptAdd",
        fields=[
            FakePivotField("Region", orientation=XL_ROW, position=1),
            FakePivotField("Other", orientation=XL_HIDDEN, position=None),
        ],
        data_fields=[FakeDataField("Region", "Region total", XL_FUNC_SUM, position=1)],
    )
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    # "Region total" exists for Region; asking for it on another source field is
    # rejected (Excel would reject the duplicate caption too).
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_aggregation(
            wb, "ptAdd", "Other", PivotAggregation.COUNT, caption="Region total"
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "duplicate_caption"


def test_every_accepted_aggregation_code_is_accepted() -> None:
    cases = {
        PivotAggregation.SUM: XL_FUNC_SUM,
        PivotAggregation.AVERAGE: XL_FUNC_AVERAGE,
        PivotAggregation.COUNT: XL_FUNC_COUNT,
        PivotAggregation.MAX: XL_FUNC_MAX,
    }
    for aggregation, expected_code in cases.items():
        pivot = FakePivotTable(
            "ptAgg",
            fields=[FakePivotField("Sales", orientation=XL_HIDDEN, position=None)],
            data_fields=[FakeDataField("Sales", "Sum of Sales", XL_FUNC_SUM, position=1)],
        )
        wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
        adapter = PivotTablesAdapter()

        adapter.set_field_aggregation(wb, "ptAgg", "Sales", aggregation)

        assert pivot._data_fields[0].Function == expected_code, aggregation


# ---------------------------------------------------------------------------
# Adapter: properties
# ---------------------------------------------------------------------------
def test_set_properties_caption_and_number_format() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_properties(
        wb,
        "ptA",
        "Sales",
        PivotFieldProperties(caption="Total Sales", number_format="0.00"),
    )

    assert pivot._data_fields[0].Name == "Total Sales"
    assert pivot._data_fields[0].NumberFormat == "0.00"
    assert result["verified"] is True
    assert result["field"].data_field_name == "Total Sales"
    assert result["field"].number_format == "0.00"


def test_set_properties_show_all_items_targets_source_field() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.set_field_properties(
        wb, "ptA", "Region", PivotFieldProperties(show_all_items=True)
    )

    region = next(f for f in pivot._fields if f.SourceName == "Region")
    assert region.ShowAllItems is True
    assert result["verified"] is True
    assert result["field"].show_all_items is True


def test_set_properties_empty_patch_rejected() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_properties(wb, "ptA", "Sales", PivotFieldProperties())

    assert exc_info.value.context["reason"] == "empty_patch"


def test_set_properties_duplicate_caption_rejected() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_properties(
            wb, "ptA", "Sales", PivotFieldProperties(caption="Average of Units")
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "duplicate_caption"


def test_set_properties_caption_on_non_data_field_rejected() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_properties(wb, "ptA", "Quarter", PivotFieldProperties(caption="Qtr"))

    assert exc_info.value.context["reason"] == "not_a_data_field"


def test_set_properties_blank_caption_rejected() -> None:
    pivot = _pt_a()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_properties(wb, "ptA", "Sales", PivotFieldProperties(caption="  "))

    assert exc_info.value.context["reason"] == "invalid_caption"


def test_set_properties_rejects_olap_source() -> None:
    pivot = _olap_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_field_properties(
            wb, "ptOlap", "Region", PivotFieldProperties(show_all_items=True)
        )

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE


def test_validate_field_properties_reports_applied_fields() -> None:
    applied = validate_field_properties(PivotFieldProperties(caption="X", show_all_items=False))
    assert applied == ["caption", "show_all_items"]


# ---------------------------------------------------------------------------
# Service: guard chain + audit + op_class
# ---------------------------------------------------------------------------
async def test_service_set_layout_success_is_guarded_and_audited(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    pivot = _seed(rig, _pt_a())
    rig.grant(_ACTION_LAYOUT)
    rig.pivots._snapshot.snapshots.add(info.workbook_id)  # type: ignore[union-attr]

    result = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
    )

    assert isinstance(result, PivotFieldInfo)
    assert result.orientation is PivotFieldOrientation.COLUMN
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_field_set_layout")
    assert submit["op_class"] is OperationClass.MUTATING
    records = [json.loads(line) for line in (tmp_path / "audit.jsonl").read_text().splitlines()]
    entry = next(r for r in records if r["tool"] == "pivot_field_set_layout")
    assert entry["operation_class"] == "mutating"
    assert entry["result"]["verified"] is True
    assert entry["result"]["before"]["orientation"] == "row"
    assert entry["result"]["after"]["orientation"] == "column"
    assert entry["result"]["orientation_order"] == ["Region", "Product"]
    assert pivot is not None


async def test_service_rejects_read_only_open(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.grant(_ACTION_LAYOUT)
    rig.pivots._snapshot.snapshots.add(info.workbook_id)  # type: ignore[union-attr]

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_service_rejects_without_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant(_ACTION_LAYOUT)
    # No snapshot created.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_service_rejects_without_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.pivots._snapshot.snapshots.add(info.workbook_id)  # type: ignore[union-attr]
    # The layout action is deliberately not granted.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.COM_BUSY


async def test_service_rejects_data_orientation_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout("wb", "ptA", "Units", PivotFieldOrientation.DATA)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_orientation"
    # Pure argument validation: nothing was submitted to the worker.
    assert not [s for s in rig.worker.submits if s["name"].startswith("pivot_field_set_")]


async def test_service_rejects_bad_position_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_position("wb", "ptA", "Region", 0)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert not [s for s in rig.worker.submits if s["name"].startswith("pivot_field_set_")]


async def test_service_rejects_legacy_aggregation_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_aggregation(
            "wb",
            "ptA",
            "Sales",
            -4114,  # type: ignore[arg-type]
        )

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_function"
    assert not [s for s in rig.worker.submits if s["name"].startswith("pivot_field_set_")]


async def test_service_rejects_empty_property_patch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_properties("wb", "ptA", "Sales", PivotFieldProperties())

    assert exc_info.value.context["reason"] == "empty_patch"
    assert not [s for s in rig.worker.submits if s["name"].startswith("pivot_field_set_")]


async def test_service_set_aggregation_and_position_are_mutations(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _pt_a())
    rig.grant(_ACTION_AGGREGATION)
    rig.grant(_ACTION_POSITION)
    rig.grant(_ACTION_PROPERTIES)
    rig.pivots._snapshot.snapshots.add(info.workbook_id)  # type: ignore[union-attr]

    aggregation = await rig.pivots.set_pivot_field_aggregation(
        info.workbook_id, "ptA", "Sales", PivotAggregation.MAX
    )
    position = await rig.pivots.set_pivot_field_position(info.workbook_id, "ptA", "Units", 1)
    properties = await rig.pivots.set_pivot_field_properties(
        info.workbook_id, "ptA", "Sales", PivotFieldProperties(caption="Total Sales")
    )

    assert aggregation.function is PivotAggregation.MAX
    assert position.position == 1
    assert properties.data_field_name == "Total Sales"
    names = [s["name"] for s in rig.worker.submits]
    assert "pivot_field_set_aggregation" in names
    assert "pivot_field_set_position" in names
    assert "pivot_field_set_properties" in names
    assert all(
        s["op_class"] is OperationClass.MUTATING
        for s in rig.worker.submits
        if s["name"].startswith("pivot_field_set_")
    )


# ---------------------------------------------------------------------------
# Tools: thin delegation + argument validation
# ---------------------------------------------------------------------------
async def test_tools_delegate_to_the_service(tmp_path: Path) -> None:
    from excel_vba_mcp.api import pivot_tools

    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _pt_a())
    for action in (_ACTION_LAYOUT, _ACTION_POSITION, _ACTION_AGGREGATION, _ACTION_PROPERTIES):
        rig.grant(action)
    rig.pivots._snapshot.snapshots.add(info.workbook_id)  # type: ignore[union-attr]
    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        layout = await pivot_tools.pivot_field_set_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )
        position = await pivot_tools.pivot_field_set_position(info.workbook_id, "ptA", "Units", 1)
        aggregation = await pivot_tools.pivot_field_set_aggregation(
            info.workbook_id, "ptA", "Sales", PivotAggregation.AVERAGE
        )
        properties = await pivot_tools.pivot_field_set_properties(
            info.workbook_id,
            "ptA",
            "Sales",
            PivotFieldProperties(caption="Total Sales"),
        )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert layout.orientation is PivotFieldOrientation.COLUMN
    assert position.position == 1
    assert aggregation.function is PivotAggregation.AVERAGE
    assert properties.data_field_name == "Total Sales"


async def test_tools_reject_invalid_orientation() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await pivot_tools.pivot_field_set_layout(
                "wb",
                "ptA",
                "Region",
                "diagonal",  # type: ignore[arg-type]
            )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert exc_info.value.context["reason"] == "invalid_orientation"


async def test_tools_reject_legacy_aggregation() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await pivot_tools.pivot_field_set_aggregation(
                "wb",
                "ptA",
                "Sales",
                9,  # type: ignore[arg-type]
            )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert exc_info.value.context["reason"] == "invalid_function"


async def test_tools_reject_bad_position() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError):
            await pivot_tools.pivot_field_set_position("wb", "ptA", "Region", 0)
    finally:
        pivot_tools.set_pivot_tables_service(None)


async def test_tools_reject_non_dto_properties() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError):
            await pivot_tools.pivot_field_set_properties(
                "wb",
                "ptA",
                "Sales",
                {"caption": "X"},  # type: ignore[arg-type]
            )
    finally:
        pivot_tools.set_pivot_tables_service(None)
