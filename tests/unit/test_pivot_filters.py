"""Unit tests for the PivotField filter surface (task_051).

These tests drive :class:`~excel_vba_mcp.adapters.excel.pivot_tables_adapter.PivotTablesAdapter`,
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` and
the ``pivot_field_*_filters`` MCP tools against a **fake COM graph** — no Excel,
no COM, no ``pywin32``. They assert:

* **Read** — ``all_items_visible``, the visible/hidden ``PivotItems`` (manual
  filter state), the classified filter kind and any detected ``PivotFilters``
  entries are reported truthfully;
* **Manual apply** — ``PivotItems[i].Visible`` hides/shows the named items and the
  postcondition re-read matches the request; an unknown, blank, duplicate or
  empty item set is rejected **before** any mutation;
* **Unsupported filters** — a label / value / date / top-bottom request returns an
  explicit ``unsupported`` result with a reason, never touches Excel and is never
  converted into a manual filter (``PivotFilters.Add`` is not called);
* **Clear scope** — ``scope="field"`` clears one field, ``scope="table"`` clears
  the whole PivotTable; both are confirmed by a re-read, and a silent no-op is
  reported as a postcondition failure rather than success;
* **Field identity** — an unknown field is rejected;
* **Guards** — writable session → snapshot → approval, in that order, and every
  mutation is dispatched with ``OperationClass.MUTATING``;
* **Unsupported source** — an OLAP / Data Model cache is rejected with
  ``UNSUPPORTED_PIVOT_SOURCE`` before any field is touched.

Run::

    uv run pytest tests/unit/test_pivot_filters.py -q
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
    normalize_filter_scope,
    normalize_filter_spec,
)
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotFieldFilters, PivotFilterSpec
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotFilterKind,
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

# ``XlPivotFieldOrientation`` / ``XlPivotTableSourceType`` values used by the fakes.
XL_HIDDEN = 0
XL_ROW = 1
XL_COLUMN = 2
XL_PAGE = 3

XL_SOURCE_DATABASE = 1

#: ``XlPivotFilterType`` codes used to prove unsupported filters are reported.
XL_FILTER_VALUE_EQUALS = 7
XL_FILTER_CAPTION_EQUALS = 15

#: ``PivotField.Position`` sentinel for a hidden/unplaced field (#N/A).
POSITION_SENTINEL = -2146826246

_ACTION_OPEN = "excel_open"
_ACTION_APPLY = "pivot_field_apply_filter"
_ACTION_CLEAR = "pivot_field_clear_filters"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake COM graph (mutable: item visibility / filters / clear)
# ---------------------------------------------------------------------------
class FakePivotItem:
    """A PivotItem whose ``Visible`` flag can be made write-opaque."""

    def __init__(self, value: str, visible: bool = True, *, ignore_writes: bool = False) -> None:
        object.__setattr__(self, "Value", value)
        object.__setattr__(self, "Visible", visible)
        object.__setattr__(self, "_ignore_writes", ignore_writes)

    def __setattr__(self, name: str, value: Any) -> None:
        if name == "Visible" and getattr(self, "_ignore_writes", False):
            return
        object.__setattr__(self, name, value)


class FakePivotItems:
    def __init__(self, items: list[FakePivotItem]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakePivotItem:  # noqa: N802 - COM casing
        return self._items[index - 1]


class FakePivotFilter:
    """A single ``PivotFilters`` entry (only its ``Type`` is read)."""

    def __init__(self, filter_type: int) -> None:
        self.Type = filter_type


class FakePivotFilters:
    """A ``PivotFilters`` collection that fails loudly if ``Add`` is called."""

    def __init__(self, filters: list[FakePivotFilter], owner: FakePivotField) -> None:
        self._filters = filters
        self._owner = owner

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._filters)

    def Item(self, index: int) -> FakePivotFilter:  # noqa: N802 - COM casing
        return self._filters[index - 1]

    def Add(self, *_args: Any, **_kwargs: Any) -> None:  # noqa: N802 - COM casing
        # The adapter must never call ``PivotFilters.Add`` on this build.
        self._owner.add_filter_calls += 1
        raise RuntimeError("PivotFilters.Add is not supported on the target build")


class FakePivotField:
    """A source field with items, filter state and a clear operation."""

    def __init__(
        self,
        source_name: str,
        *,
        orientation: int = XL_HIDDEN,
        position: int | None = None,
        items: list[FakePivotItem] | None = None,
        filters: list[FakePivotFilter] | None = None,
        clear_is_noop: bool = False,
    ) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = source_name
        self._orientation = orientation
        self._position = POSITION_SENTINEL if position is None else position
        self.ShowAllItems = False
        self._items = items or []
        self._filters = filters or []
        self._clear_is_noop = clear_is_noop
        self.add_filter_calls = 0
        self.clear_all_filters_calls = 0
        self.deleted = False

    @property
    def Orientation(self) -> int:  # noqa: N802 - COM casing
        return self._orientation

    @Orientation.setter
    def Orientation(self, value: int) -> None:  # noqa: N802 - COM casing
        self._orientation = value
        if value == XL_HIDDEN:
            self._position = POSITION_SENTINEL

    @property
    def Position(self) -> int:  # noqa: N802 - COM casing
        return self._position

    @property
    def AllItemsVisible(self) -> bool:  # noqa: N802 - COM casing
        """Excel computes this from the item visibility flags."""

        return all(item.Visible for item in self._items)

    @property
    def PivotFilters(self) -> FakePivotFilters:  # noqa: N802 - COM casing
        return FakePivotFilters(self._filters, self)

    def PivotItems(self) -> FakePivotItems:  # noqa: N802 - COM casing
        return FakePivotItems(self._items)

    def ClearAllFilters(self) -> None:  # noqa: N802 - COM casing
        self.clear_all_filters_calls += 1
        if self._clear_is_noop:
            return
        for item in self._items:
            item.Visible = True
        self._filters = []

    def Delete(self) -> None:  # noqa: N802 - COM casing
        self.deleted = True


class FakeDataField:
    """A minimal data field (only the metadata the adapter reads)."""

    def __init__(
        self,
        source_name: str,
        caption: str,
        function: int,
        *,
        position: int | None = None,
    ) -> None:
        self.SourceName = source_name
        self.Name = caption
        self.Caption = caption
        self.Function = function
        self.NumberFormat = "General"
        self._position = POSITION_SENTINEL if position is None else position

    @property
    def Position(self) -> int:  # noqa: N802 - COM casing
        return self._position


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
    def __init__(self, index: int = 1, *, olap: bool = False) -> None:
        self.Index = index
        self.SourceType = XL_SOURCE_DATABASE
        self.SourceData = "Data!R1C1:R7C5"
        self.RefreshDate = datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC)
        self.OLAP = olap
        self.RecordCount = 6


class FakeRange:
    def __init__(self, address: str) -> None:
        self.Address = address


class FakePivotTable:
    """A PivotTable whose ``ClearAllFilters`` can be made a silent no-op."""

    def __init__(
        self,
        name: str,
        *,
        cache: FakePivotCache | None = None,
        fields: list[FakePivotField] | None = None,
        data_fields: list[FakeDataField] | None = None,
        clear_is_noop: bool = False,
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
        self._clear_is_noop = clear_is_noop
        self.clear_all_filters_calls = 0

    @property
    def source_fields(self) -> list[FakePivotField]:
        return [field for field in self._fields if not field.deleted]

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

    def ClearAllFilters(self) -> None:  # noqa: N802 - COM casing
        self.clear_all_filters_calls += 1
        if self._clear_is_noop:
            return
        for field in self.source_fields:
            field.ClearAllFilters()


class FakePivotTables:
    def __init__(self, pivots: list[FakePivotTable]) -> None:
        self._pivots = pivots

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._pivots)

    def Item(self, index: int) -> FakePivotTable:  # noqa: N802 - COM casing
        return self._pivots[index - 1]


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
        workbook = FakeWorkbook([FakeSheet("PivotReport")])
        workbook.FullName = path
        workbook.ReadOnly = bool(kwargs.get("ReadOnly", False))
        self.items.append(workbook)
        return workbook

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
        self.pivots = PivotTablesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.audit,
            snapshot_service=FakeSnapshotGate(),
            approval=self.approval,
        )

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

    def add_snapshot(self, workbook_id: str) -> None:
        gate = self.pivots._snapshot
        assert isinstance(gate, FakeSnapshotGate)
        gate.snapshots.add(workbook_id)

    def workbook(self) -> FakeWorkbook:
        for workbook in self.app.Workbooks:
            return workbook
        raise AssertionError("no workbook open")


# ---------------------------------------------------------------------------
# Fake graph builders
# ---------------------------------------------------------------------------
def _items(*values: str, hidden: tuple[str, ...] = ()) -> list[FakePivotItem]:
    return [FakePivotItem(value, value not in hidden) for value in values]


def _region_pivot(
    *,
    name: str = "ptA",
    hidden: tuple[str, ...] = (),
    ignore_item_writes: bool = False,
    filters: list[FakePivotFilter] | None = None,
    clear_is_noop: bool = False,
    olap: bool = False,
) -> FakePivotTable:
    region = FakePivotField(
        "Region",
        orientation=XL_ROW,
        position=1,
        items=_items("North", "South", "East", "West", hidden=hidden),
        filters=filters,
        clear_is_noop=clear_is_noop,
    )
    if ignore_item_writes:
        for item in region._items:
            object.__setattr__(item, "_ignore_writes", True)
    return FakePivotTable(
        name,
        cache=FakePivotCache(1, olap=olap),
        fields=[region],
        data_fields=[FakeDataField("Units", "Sum of Units", -4157, position=1)],
    )


def _seed(rig: Rig, pivot: FakePivotTable) -> FakePivotTable:
    workbook = rig.workbook()
    sheet = workbook.Worksheets.Item("PivotReport")
    sheet._pivots = [pivot]
    return pivot


def _region(pivot: FakePivotTable) -> FakePivotField:
    field = next(f for f in pivot._fields if f.SourceName == "Region")
    assert isinstance(field, FakePivotField)
    return field


def _manual_spec(*items: str) -> PivotFilterSpec:
    return PivotFilterSpec(kind=PivotFilterKind.MANUAL, visible_items=list(items))


# ---------------------------------------------------------------------------
# Adapter: read
# ---------------------------------------------------------------------------
def test_get_filters_reports_manual_hidden_items() -> None:
    pivot = _region_pivot(hidden=("West",))
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    filters = PivotTablesAdapter().get_field_filters(wb, "ptA", "Region")

    assert isinstance(filters, PivotFieldFilters)
    assert filters.pivot_table == "ptA"
    assert filters.field == "Region"
    assert filters.all_items_visible is False
    assert filters.visible_items == ["North", "South", "East"]
    assert filters.hidden_items == ["West"]
    assert filters.filter.kind is PivotFilterKind.MANUAL
    assert filters.filter.supported is True
    assert filters.unsupported == []


def test_get_filters_reports_all_visible() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    filters = PivotTablesAdapter().get_field_filters(wb, "ptA", "Region")

    assert filters.all_items_visible is True
    assert filters.hidden_items == []
    assert filters.filter.kind is PivotFilterKind.NONE
    assert filters.visible_items == ["North", "South", "East", "West"]


def test_get_filters_reports_detected_pivot_filters() -> None:
    pivot = _region_pivot(filters=[FakePivotFilter(XL_FILTER_VALUE_EQUALS)])
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    filters = PivotTablesAdapter().get_field_filters(wb, "ptA", "Region")

    assert filters.unsupported == [f"value({XL_FILTER_VALUE_EQUALS})"]
    # The manual state is still reported alongside the unsupported entries.
    assert filters.filter.kind is PivotFilterKind.NONE


def test_get_filters_rejects_unknown_field() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().get_field_filters(wb, "ptA", "Nope")

    assert exc_info.value.context["reason"] == "field_not_found"


def test_get_filters_never_calls_pivot_filters_add() -> None:
    pivot = _region_pivot(filters=[FakePivotFilter(XL_FILTER_CAPTION_EQUALS)])
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    PivotTablesAdapter().get_field_filters(wb, "ptA", "Region")

    assert _region(pivot).add_filter_calls == 0


# ---------------------------------------------------------------------------
# Adapter: manual apply
# ---------------------------------------------------------------------------
def test_apply_manual_filter_hides_unlisted_items() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.apply_field_filter(wb, "ptA", "Region", _manual_spec("North", "South"))

    region = _region(pivot)
    assert [item.Value for item in region._items if item.Visible] == ["North", "South"]
    assert [item.Value for item in region._items if not item.Visible] == ["East", "West"]
    assert result["verified"] is True
    assert result["applied"] is True
    assert result["notes"] == []
    assert result["filters"].visible_items == ["North", "South"]
    assert result["filters"].hidden_items == ["East", "West"]
    assert result["filters"].filter.kind is PivotFilterKind.MANUAL
    assert result["filters"].all_items_visible is False
    # The unsupported ``PivotFilters.Add`` path was never touched.
    assert region.add_filter_calls == 0


def test_apply_manual_filter_shows_all_listed_items() -> None:
    pivot = _region_pivot(hidden=("West",))
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    result = PivotTablesAdapter().apply_field_filter(
        wb, "ptA", "Region", _manual_spec("North", "South", "East", "West")
    )

    assert result["filters"].hidden_items == []
    assert result["filters"].all_items_visible is True
    assert result["filters"].filter.kind is PivotFilterKind.NONE
    assert _region(pivot).add_filter_calls == 0


def test_apply_manual_filter_rejects_unknown_item() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(
            wb, "ptA", "Region", _manual_spec("North", "Nowhere")
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "unknown_filter_item"
    assert exc_info.value.context["unknown_items"] == ["nowhere"]
    # Nothing changed.
    assert all(item.Visible for item in _region(pivot)._items)


def test_apply_manual_filter_rejects_duplicate_items() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(wb, "ptA", "Region", _manual_spec("North", "north"))

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "ambiguous_filter_item"


def test_apply_manual_filter_rejects_blank_item() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(wb, "ptA", "Region", _manual_spec(" "))

    assert exc_info.value.context["reason"] == "invalid_filter_item"


def test_apply_manual_filter_rejects_empty_item_set() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(wb, "ptA", "Region", _manual_spec())

    assert exc_info.value.context["reason"] == "empty_filter"
    assert all(item.Visible for item in _region(pivot)._items)


def test_apply_unsupported_label_filter_is_explicit() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()
    spec = PivotFilterSpec(
        kind=PivotFilterKind.LABEL,
        visible_items=[],
        supported=False,
        unsupported_reason=None,
    )

    result = adapter.apply_field_filter(wb, "ptA", "Region", spec)

    assert result["applied"] is False
    assert result["filters"].filter.supported is False
    assert result["filters"].filter.kind is PivotFilterKind.LABEL
    assert result["filters"].unsupported == ["label"]
    assert "PivotFilters.Add is unsupported" in (result["filters"].filter.unsupported_reason or "")
    # No Excel mutation and no ``PivotFilters.Add`` attempt.
    assert all(item.Visible for item in _region(pivot)._items)
    assert _region(pivot).add_filter_calls == 0


@pytest.mark.parametrize(
    ("kind", "expected"),
    [
        (PivotFilterKind.VALUE, "value"),
        (PivotFilterKind.DATE, "date"),
        (PivotFilterKind.TOP_BOTTOM, "top_bottom"),
        (PivotFilterKind.UNSUPPORTED, "unsupported"),
    ],
)
def test_every_non_manual_kind_is_unsupported(kind: PivotFilterKind, expected: str) -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    result = PivotTablesAdapter().apply_field_filter(
        wb, "ptA", "Region", PivotFilterSpec(kind=kind)
    )

    assert result["applied"] is False
    assert result["filters"].unsupported == [expected]
    assert _region(pivot).add_filter_calls == 0


def test_apply_filter_rejects_olap_source() -> None:
    pivot = _region_pivot(olap=True)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(wb, "ptA", "Region", _manual_spec("North"))

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    assert all(item.Visible for item in _region(pivot)._items)


def test_apply_manual_filter_reports_postcondition_mismatch() -> None:
    pivot = _region_pivot(ignore_item_writes=True)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().apply_field_filter(wb, "ptA", "Region", _manual_spec("North"))

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "postcondition_mismatch"
    # The fake ignored the write, so nothing is hidden.
    assert all(item.Visible for item in _region(pivot)._items)


# ---------------------------------------------------------------------------
# Adapter: clear
# ---------------------------------------------------------------------------
def test_clear_field_scope_restores_all_items() -> None:
    pivot = _region_pivot(hidden=("West",))
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])
    adapter = PivotTablesAdapter()

    result = adapter.clear_field_filters(wb, "ptA", "Region", scope="field")

    assert _region(pivot).clear_all_filters_calls == 1
    assert pivot.clear_all_filters_calls == 0
    assert result["scope"] == "field"
    assert result["field"] == "Region"
    assert result["cleared_fields"] == ["Region"]
    assert result["verified"] is True
    assert result["filters"].hidden_items == []
    assert result["filters"].all_items_visible is True


def test_clear_table_scope_clears_every_field() -> None:
    pivot = _region_pivot(hidden=("West",))
    other = FakePivotField("Product", orientation=XL_COLUMN, position=1, items=_items("A", "B"))
    object.__setattr__(other._items[1], "Visible", False)
    pivot._fields.append(other)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    result = PivotTablesAdapter().clear_field_filters(wb, "ptA", scope="table")

    assert pivot.clear_all_filters_calls == 1
    assert _region(pivot).clear_all_filters_calls == 1
    assert result["scope"] == "table"
    assert result["field"] == ""
    assert result["cleared_fields"] == ["Product", "Region"]
    assert result["verified"] is True
    assert all(item.Visible for item in other._items)
    # The table-level result is explicit about its scope through the empty field.
    assert result["filters"].pivot_table == "ptA"


def test_clear_table_scope_restores_page_field_selection_state() -> None:
    """A table clear resets every field, not only the row field with hidden items."""

    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    result = PivotTablesAdapter().clear_field_filters(wb, "ptA", scope="table")

    assert result["cleared_fields"] == ["Region"]
    assert result["filters"].field == ""


def test_clear_field_scope_requires_field() -> None:
    pivot = _region_pivot()
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().clear_field_filters(wb, "ptA", None, scope="field")

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "field_required"


def test_clear_rejects_invalid_scope() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        normalize_filter_scope("everything")

    assert exc_info.value.context["reason"] == "invalid_scope"


def test_clear_field_scope_reports_postcondition_mismatch() -> None:
    pivot = _region_pivot(hidden=("West",), clear_is_noop=True)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().clear_field_filters(wb, "ptA", "Region", scope="field")

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "postcondition_mismatch"
    assert exc_info.value.context["hidden_items"] == ["West"]


def test_clear_table_scope_reports_postcondition_mismatch() -> None:
    pivot = _region_pivot(hidden=("West",), clear_is_noop=True)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().clear_field_filters(wb, "ptA", scope="table")

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "postcondition_mismatch"
    assert exc_info.value.context["still_filtered"] == ["Region"]


def test_clear_rejects_olap_source() -> None:
    pivot = _region_pivot(hidden=("West",), olap=True)
    wb = FakeWorkbook([FakeSheet("PivotReport", [pivot])])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        PivotTablesAdapter().clear_field_filters(wb, "ptA", "Region", scope="field")

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    assert _region(pivot).clear_all_filters_calls == 0


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------
def test_normalize_filter_spec_cleans_items() -> None:
    spec = normalize_filter_spec(_manual_spec(" North ", "South"))

    assert spec.visible_items == ["North", "South"]
    assert spec.kind is PivotFilterKind.MANUAL


def test_normalize_filter_spec_rejects_non_dto() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        normalize_filter_spec({"kind": "manual"})  # type: ignore[arg-type]

    assert exc_info.value.context["reason"] == "invalid_filter_spec"


# ---------------------------------------------------------------------------
# Service: read + guard chain + audit + op_class
# ---------------------------------------------------------------------------
async def test_service_get_filters_is_read_only_and_audited(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    pivot = _seed(rig, _region_pivot(hidden=("West",)))

    filters = await rig.pivots.get_pivot_field_filters(info.workbook_id, "ptA", "Region")

    assert filters.hidden_items == ["West"]
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_field_get_filters")
    assert submit["op_class"] is OperationClass.READ_ONLY
    records = [json.loads(line) for line in (tmp_path / "audit.jsonl").read_text().splitlines()]
    entry = next(r for r in records if r["tool"] == "pivot_field_get_filters")
    assert entry["operation_class"] == "read_only"
    assert entry["result"]["hidden_items"] == ["West"]
    assert pivot is not None


async def test_service_apply_filter_is_guarded_and_audited(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _region_pivot())
    rig.grant(_ACTION_APPLY)
    rig.add_snapshot(info.workbook_id)

    filters = await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, "ptA", "Region", _manual_spec("North")
    )

    assert filters.hidden_items == ["South", "East", "West"]
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_field_apply_filter")
    assert submit["op_class"] is OperationClass.MUTATING
    records = [json.loads(line) for line in (tmp_path / "audit.jsonl").read_text().splitlines()]
    entry = next(r for r in records if r["tool"] == "pivot_field_apply_filter")
    assert entry["operation_class"] == "mutating"
    assert entry["result"]["applied"] is True
    assert entry["result"]["verified"] is True
    assert entry["result"]["before"]["hidden_items"] == []


async def test_service_clear_filter_is_guarded_and_audited(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _region_pivot(hidden=("West",)))
    rig.grant(_ACTION_CLEAR)
    rig.add_snapshot(info.workbook_id)

    filters = await rig.pivots.clear_pivot_field_filters(info.workbook_id, "ptA", scope="table")

    assert filters.field == ""
    assert filters.hidden_items == []
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_field_clear_filters")
    assert submit["op_class"] is OperationClass.MUTATING
    records = [json.loads(line) for line in (tmp_path / "audit.jsonl").read_text().splitlines()]
    entry = next(r for r in records if r["tool"] == "pivot_field_clear_filters")
    assert entry["operation_class"] == "mutating"
    assert entry["result"]["scope"] == "table"
    assert entry["result"]["cleared_fields"] == ["Region"]
    assert entry["result"]["verified"] is True


async def test_service_apply_rejects_read_only_open(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.grant(_ACTION_APPLY)
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, "ptA", "Region", _manual_spec("North")
        )

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_service_apply_rejects_without_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant(_ACTION_APPLY)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, "ptA", "Region", _manual_spec("North")
        )

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_service_apply_rejects_without_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    # The apply action is deliberately not granted.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, "ptA", "Region", _manual_spec("North")
        )

    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED


async def test_service_clear_rejects_invalid_scope_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.clear_pivot_field_filters("wb", "ptA", "Region", scope="everything")

    assert exc_info.value.context["reason"] == "invalid_scope"
    assert not [s for s in rig.worker.submits if s["name"] == "pivot_field_clear_filters"]


async def test_service_clear_field_scope_requires_field_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    await rig.open()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.clear_pivot_field_filters("wb", "ptA", None, scope="field")

    assert exc_info.value.context["reason"] == "field_required"
    assert not [s for s in rig.worker.submits if s["name"] == "pivot_field_clear_filters"]


async def test_service_unsupported_filter_returns_before_dispatch(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _region_pivot())
    # No snapshot and no approval: an unsupported request must not need them.

    filters = await rig.pivots.apply_pivot_field_filter(
        info.workbook_id,
        "ptA",
        "Region",
        PivotFilterSpec(kind=PivotFilterKind.VALUE),
    )

    assert filters.filter.supported is False
    assert filters.unsupported == ["value"]
    assert not [s for s in rig.worker.submits if s["name"] == "pivot_field_apply_filter"]
    records = [json.loads(line) for line in (tmp_path / "audit.jsonl").read_text().splitlines()]
    entry = next(r for r in records if r["tool"] == "pivot_field_apply_filter")
    assert entry["result"]["applied"] is False
    assert entry["result"]["status"] == "unsupported"


# ---------------------------------------------------------------------------
# Tools: thin delegation + argument validation
# ---------------------------------------------------------------------------
async def test_tools_delegate_to_the_service(tmp_path: Path) -> None:
    from excel_vba_mcp.api import pivot_tools

    rig = Rig(tmp_path)
    info = await rig.open()
    _seed(rig, _region_pivot(hidden=("West",)))
    rig.grant(_ACTION_APPLY)
    rig.grant(_ACTION_CLEAR)
    rig.add_snapshot(info.workbook_id)
    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        read = await pivot_tools.pivot_field_get_filters(info.workbook_id, "ptA", "Region")
        applied = await pivot_tools.pivot_field_apply_filter(
            info.workbook_id, "ptA", "Region", _manual_spec("North")
        )
        cleared = await pivot_tools.pivot_field_clear_filters(
            info.workbook_id, "ptA", scope="table"
        )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert read.hidden_items == ["West"]
    assert applied.hidden_items == ["South", "East", "West"]
    assert cleared.hidden_items == []


async def test_tools_reject_non_dto_filter_spec() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await pivot_tools.pivot_field_apply_filter(
                "wb",
                "ptA",
                "Region",
                {"kind": "manual"},  # type: ignore[arg-type]
            )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert exc_info.value.context["reason"] == "invalid_filter_spec"


async def test_tools_reject_invalid_clear_scope() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await pivot_tools.pivot_field_clear_filters("wb", "ptA", "Region", scope="everything")
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert exc_info.value.context["reason"] == "invalid_scope"


async def test_tools_reject_field_scope_without_field() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError):
            await pivot_tools.pivot_field_clear_filters("wb", "ptA", scope="field")
    finally:
        pivot_tools.set_pivot_tables_service(None)


async def test_tools_reject_empty_workbook_id() -> None:
    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(object())  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError):
            await pivot_tools.pivot_field_get_filters("", "ptA", "Region")
    finally:
        pivot_tools.set_pivot_tables_service(None)
