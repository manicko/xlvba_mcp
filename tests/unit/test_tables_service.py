"""Unit tests for the Excel Table (ListObject) surface (task_026).

These tests exercise the production
:class:`~excel_vba_mcp.application.tables_service.TablesService` and
:class:`~excel_vba_mcp.adapters.excel.tables_adapter.TablesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

The fake models an Excel worksheet as a dense grid of cells (the single source of
truth). A ``FakeListObject`` occupies a header row plus a contiguous data body in
that grid; ``ListRows`` / ``ListColumns`` mutations insert/delete grid rows and
columns, so every read (``Range.Value`` / ``DataBodyRange.Cells(...).Value``)
reflects the mutations.

Covered:

* ``tables_list`` / ``table_get`` metadata (range, header/body/total ranges,
  totals, style, rows, columns).
* ``table_read_rows`` typed values (``str`` / ``int`` / ``float`` / ``bool`` /
  ``None`` / ISO date / Excel error) and offset/limit + response-size bound.
* ``table_append_rows`` / ``table_insert_row`` / ``table_update_rows`` /
  ``table_delete_rows`` via ``ListRows`` (structure preserved).
* ``table_add_column`` / ``table_delete_column`` via ``ListColumns`` with header
  uniqueness.
* ``table_create`` / ``table_resize``.
* Guard chain: no snapshot -> ``SNAPSHOT_FAILED``; no approval -> denied;
  read-only session -> ``WORKBOOK_READ_ONLY``.
* Ambiguous ``where`` update/delete rejected with a ``preview`` and no mutation.
* Non-mutating reads submit ``READ_ONLY``; mutations submit ``MUTATING``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.tables_adapter import TableMutationResult, TablesAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    NameConflictError,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    OperationClass,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph — a worksheet is a dense grid; a ListObject owns a
# header row + contiguous data body within that grid.
# ---------------------------------------------------------------------------
def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _parse_ref(ref: str) -> tuple[int, int]:
    cleaned = ref.replace("$", "")
    letters = "".join(ch for ch in cleaned if ch.isalpha())
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return int(digits), col


class FakeSheet:
    """A fake worksheet whose dense grid is the single source of truth."""

    def __init__(self, name: str) -> None:
        self.Name = name
        self.CodeName = f"{name}_Code"
        self.ListObjects = FakeListObjects(self)
        self._grid: list[list[Any]] = []
        self.Cells = FakeCells(self)

    # 1-based grid access ------------------------------------------------
    def get(self, row: int, col: int) -> Any:
        if 1 <= row <= len(self._grid):
            line = self._grid[row - 1]
            if 1 <= col <= len(line):
                return line[col - 1]
        return None

    def set(self, row: int, col: int, value: Any) -> None:
        self._ensure(row, col)
        self._grid[row - 1][col - 1] = value

    def _ensure(self, row: int, col: int) -> None:
        while len(self._grid) < row:
            self._grid.append([])
        while len(self._grid[row - 1]) < col:
            self._grid[row - 1].append(None)

    def seed(self, rows: list[list[Any]]) -> None:
        self._grid = [list(row) for row in rows]

    # structural grid ops used by ListRows / ListColumns -----------------
    def grid_row_count(self) -> int:
        return len(self._grid)

    def insert_row(self, at: int) -> None:
        self._grid.insert(at - 1, [])

    def delete_row(self, at: int) -> None:
        del self._grid[at - 1]

    def insert_col(self, at: int) -> None:
        for line in self._grid:
            while len(line) < at:
                line.append(None)
            line.insert(at - 1, None)

    def delete_col(self, at: int) -> None:
        for line in self._grid:
            if 1 <= at <= len(line):
                del line[at - 1]

    # COM-ish range resolution -------------------------------------------
    def Range(self, *args: Any) -> Any:  # noqa: N802 - COM casing
        if len(args) == 1 and isinstance(args[0], str):
            start, _, end = args[0].partition(":")
            top, left = _parse_ref(start)
            bottom, right = _parse_ref(end) if end else (top, left)
            return FakeRange(self, top, left, bottom - top + 1, right - left + 1)
        first, second = args
        return FakeRange(
            self,
            first.row,
            first.col,
            second.row - first.row + 1,
            second.col - first.col + 1,
        )


class FakeCell:
    """A cell proxy bound to a sheet + 1-based position; ``Value`` reads/writes."""

    def __init__(self, sheet: FakeSheet, row: int, col: int) -> None:
        self._sheet = sheet
        self.row = row
        self.col = col

    @property
    def Value(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.get(self.row, self.col)

    @Value.setter
    def Value(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.set(self.row, self.col, value)


class FakeCells:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet

    def __call__(self, row: int, col: int) -> FakeCell:
        return FakeCell(self._sheet, row, col)


class FakeRange:
    def __init__(self, sheet: FakeSheet, top: int, left: int, height: int, width: int) -> None:
        self._sheet = sheet
        self.Top = top
        self.Left = left
        self.Height = height
        self.Width = width

    @property
    def Row(self) -> int:  # noqa: N802 - COM casing
        return self.Top

    @property
    def Column(self) -> int:  # noqa: N802 - COM casing
        return self.Left

    @property
    def Worksheet(self) -> FakeSheet:  # noqa: N802 - COM casing
        return self._sheet

    @property
    def Address(self) -> str:  # noqa: N802 - COM casing
        return (
            f"${_col_letter(self.Left)}${self.Top}:"
            f"${_col_letter(self.Left + self.Width - 1)}${self.Top + self.Height - 1}"
        )

    @property
    def Value(self) -> Any:  # noqa: N802 - COM casing
        grid = tuple(
            tuple(self._sheet.get(r, c) for c in range(self.Left, self.Left + self.Width))
            for r in range(self.Top, self.Top + self.Height)
        )
        if self.Height == 1 and self.Width == 1:
            return grid[0][0]
        if self.Height == 1:
            return grid[0]
        return grid

    def Cells(self, row: int, col: int) -> FakeCell:  # noqa: N802 - COM casing
        return FakeCell(self._sheet, self.Top + row - 1, self.Left + col - 1)


class FakeListRow:
    def __init__(self, list_object: FakeListObject, index: int) -> None:
        self._lo = list_object
        self._index = index

    @property
    def Range(self) -> FakeRange:  # noqa: N802 - COM casing
        # A single-row range positioned over this data row in the grid.
        sheet = self._lo.sheet
        return FakeRange(sheet, self._lo.data_top + self._index - 1, 1, 1, self._lo.column_count)

    def Delete(self) -> None:  # noqa: N802 - COM
        self._lo.sheet.delete_row(self._lo.data_top + self._index - 1)
        self._lo._row_count -= 1


class FakeListRows:
    def __init__(self, list_object: FakeListObject) -> None:
        self._lo = list_object

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self._lo.row_count

    def Add(self, position: int | None = None) -> FakeListRow:  # noqa: N802, N803 - COM
        sheet = self._lo.sheet
        if position is None:
            sheet.insert_row(sheet.grid_row_count() + 1)
            self._lo._row_count += 1
            index = self._lo._row_count
        else:
            at = self._lo.data_top + position - 1
            sheet.insert_row(at)
            self._lo._row_count += 1
            index = position
        return FakeListRow(self._lo, index)

    def Item(self, index: int) -> FakeListRow:  # noqa: N802 - COM casing
        return FakeListRow(self._lo, index)


class FakeListColumn:
    def __init__(self, list_object: FakeListObject, index: int) -> None:
        self._lo = list_object
        self._index = index

    @property
    def Name(self) -> str:  # noqa: N802 - COM casing
        return self._lo._columns[self._index - 1]

    @Name.setter
    def Name(self, value: str) -> None:  # noqa: N802 - COM casing
        self._lo._columns[self._index - 1] = value
        self._lo.sheet.set(self._lo.header_top, self._index, value)

    @property
    def DataBodyRange(self) -> FakeRange | None:  # noqa: N802 - COM casing
        if self._lo.row_count == 0:
            return None
        return FakeRange(self._lo.sheet, self._lo.data_top, self._index, self._lo.row_count, 1)

    @property
    def TotalsCalculation(self) -> int:  # noqa: N802 - COM casing
        return -4143

    def Delete(self) -> None:  # noqa: N802 - COM
        self._lo.sheet.delete_col(self._index)
        self._lo._columns.pop(self._index - 1)
        self._lo.sheet.set(self._lo.header_top, self._index, self._lo._columns[self._index - 1])


class FakeListColumns:
    def __init__(self, list_object: FakeListObject) -> None:
        self._lo = list_object

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self._lo.column_count

    def Add(self, position: int | None = None) -> FakeListColumn:  # noqa: N802, N803 - COM
        at = self._lo.column_count + 1 if position is None else position
        self._lo.sheet.insert_col(at)
        self._lo._columns.insert(at - 1, "")
        self._lo.sheet.set(self._lo.header_top, at, "")
        return FakeListColumn(self._lo, at)

    def Item(self, index: int) -> FakeListColumn:  # noqa: N802 - COM casing
        return FakeListColumn(self._lo, index)


class FakeListObject:
    def __init__(
        self,
        sheet: FakeSheet,
        name: str,
        *,
        show_totals: bool = False,
        style: str | None = "TableStyleMedium2",
    ) -> None:
        self.sheet = sheet
        self.Name = name
        self.DisplayName = name
        self.ShowTotals = show_totals
        self.TableStyle = style
        self.header_top = 1
        self._columns: list[str] = []
        self._data_top = 2
        self._row_count = 0

    # geometry (explicit state, maintained by the mutating helpers) -----
    @property
    def data_top(self) -> int:
        return self._data_top

    @property
    def row_count(self) -> int:
        return self._row_count

    @property
    def column_count(self) -> int:
        return len(self._columns)

    # COM surface --------------------------------------------------------
    def header_row_range(self) -> FakeRange:
        return FakeRange(self.sheet, self.header_top, 1, 1, self.column_count)

    def data_body_range(self) -> FakeRange | None:
        if self._row_count == 0:
            return None
        return FakeRange(self.sheet, self._data_top, 1, self._row_count, self.column_count)

    def total_row_range(self) -> FakeRange | None:
        if not self.ShowTotals:
            return None
        return FakeRange(self.sheet, self._data_top + self._row_count, 1, 1, self.column_count)

    @property
    def Range(self) -> FakeRange:  # noqa: N802 - COM casing
        height = 1 + self._row_count + (1 if self.ShowTotals else 0)
        return FakeRange(self.sheet, self.header_top, 1, height, self.column_count)

    @property
    def HeaderRowRange(self) -> FakeRange:  # noqa: N802 - COM casing
        return self.header_row_range()

    @property
    def DataBodyRange(self) -> FakeRange | None:  # noqa: N802 - COM casing
        return self.data_body_range()

    @property
    def TotalRowRange(self) -> FakeRange | None:  # noqa: N802 - COM casing
        return self.total_row_range()

    @property
    def ListRows(self) -> FakeListRows:  # noqa: N802 - COM casing
        return FakeListRows(self)

    @property
    def ListColumns(self) -> FakeListColumns:  # noqa: N802 - COM casing
        return FakeListColumns(self)

    def Resize(self, new_range: FakeRange) -> None:  # noqa: N802, N803 - COM
        # Grow/shrink the grid to the new range; headers stay on the top row.
        self.header_top = new_range.Top
        while self.sheet.grid_row_count() < new_range.Top + new_range.Height - 1:
            self.sheet.insert_row(self.sheet.grid_row_count() + 1)
        while self.sheet.grid_row_count() > new_range.Top + new_range.Height - 1:
            self.sheet.delete_row(self.sheet.grid_row_count())
        width = new_range.Width
        for line in self.sheet._grid:
            while len(line) < width:
                line.append(None)
        self._row_count = max(0, new_range.Height - 1)
        while len(self._columns) < width:
            self._columns.append("")


class FakeNames:
    def __call__(self, key: str) -> Any:
        raise KeyError(key)


class FakeListObjects:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet
        self._items: list[FakeListObject] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Add(self, **kwargs: Any) -> FakeListObject:  # noqa: N802, N803 - COM
        source = kwargs.get("Source")
        if not isinstance(source, FakeRange):  # pragma: no cover - defensive
            raise AssertionError("FakeListObjects.Add expects a FakeRange Source")
        name = f"Table{len(self._items) + 1}"
        list_object = FakeListObject(self._sheet, name, show_totals=False)
        list_object.header_top = source.Top
        list_object._data_top = source.Top + 1
        # Header from the source's first row; data rows from the rest.
        list_object._columns = [
            "" if v is None else str(v)
            for v in (
                self._sheet.get(source.Top, c)
                for c in range(source.Left, source.Left + source.Width)
            )
        ]
        list_object._row_count = max(0, source.Height - 1)
        self._items.append(list_object)
        return list_object

    def Item(self, index: int) -> FakeListObject:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


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
        self.Names = FakeNames()
        self.ReadOnly = False
        self.FullName = ""

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
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
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
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )
        self.service = TablesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
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
        manifest = SnapshotManifest(
            snapshot_id="snap-test",
            created_at=datetime.now(UTC),
            source_path=self.canonical,
            file_name="snap.xlsm",
            stored_path=str(self.tmp_path / "snap.xlsm"),
            sha256="0" * 64,
            size=0,
        )
        self.snapshot._latest_snapshot[workbook_id] = manifest
        self.snapshot._known_sources[workbook_id] = self.canonical

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


# ---------------------------------------------------------------------------
# Fake table helpers
# ---------------------------------------------------------------------------
def _new_table(
    wb: FakeWorkbook,
    sheet_name: str = "Data",
    *,
    name: str = "Tbl",
    columns: list[str],
    rows: list[list[Any]],
    show_totals: bool = False,
    style: str | None = "TableStyleMedium2",
) -> FakeListObject:
    """Seed a sheet grid and attach a FakeListObject over it (test helper)."""

    sheet = wb.sheet(sheet_name)
    grid: list[list[Any]] = [list(columns)]
    for row in rows:
        grid.append(list(row))
    sheet.seed(grid)
    list_object = FakeListObject(sheet, name, show_totals=show_totals, style=style)
    list_object.header_top = 1
    list_object._data_top = 2
    list_object._columns = list(columns)
    list_object._row_count = len(rows)
    sheet.ListObjects._items.append(list_object)
    return list_object


# ---------------------------------------------------------------------------
# Adapter-level: inventory + typed read
# ---------------------------------------------------------------------------
def test_list_and_get_table_metadata() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(
        wb,
        columns=["Name", "Qty"],
        rows=[["a", 1], ["b", 2]],
        show_totals=True,
        style="TableStyleMedium4",
    )

    adapter = TablesAdapter()
    tables = adapter.list_tables(wb)
    assert len(tables) == 1
    info = tables[0]
    assert info.name == "Tbl"
    assert info.sheet_name == "Data"
    assert info.sheet_code_name == "Data_Code"
    assert info.row_count == 2
    assert [c.name for c in info.columns] == ["Name", "Qty"]
    assert info.show_totals is True
    assert info.table_style == "TableStyleMedium4"
    assert info.header_row_range is not None
    assert info.data_body_range is not None
    assert info.total_row_range is not None

    fetched = adapter.get_table(wb, "tbl")  # case-insensitive
    assert fetched.name == "Tbl"


def test_get_table_missing_raises() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.get_table(wb, "Nope")


def test_read_rows_typed_values() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(
        wb,
        columns=["Name", "Qty", "Price", "Active", "When", "Err", "ErrCode"],
        rows=[
            ["widget", 3, 1.5, True, datetime(2026, 1, 2, 3, 4, 5), "#DIV/0!", -2146826281],
            [None, 0, 0.0, False, None, None, None],
        ],
    )
    adapter = TablesAdapter()

    result = adapter.read_rows(wb, "Tbl")

    assert result.total_rows == 2
    assert result.returned_rows == 2
    assert result.columns == ["Name", "Qty", "Price", "Active", "When", "Err", "ErrCode"]
    first = result.rows[0]
    assert first[0].kind is CellValueKind.STRING and first[0].value == "widget"
    assert first[1].kind is CellValueKind.NUMBER and first[1].value == 3
    assert first[2].kind is CellValueKind.NUMBER and first[2].value == 1.5
    assert first[3].kind is CellValueKind.BOOLEAN and first[3].value is True
    assert first[4].kind is CellValueKind.DATE and first[4].value == "2026-01-02T03:04:05"
    # Error literal text and the raw OLE error code both classify as ERROR.
    assert first[5].kind is CellValueKind.ERROR and first[5].value == "#DIV/0!"
    assert first[6].kind is CellValueKind.ERROR and first[6].value == "#DIV/0!"
    second = result.rows[1]
    assert second[0].kind is CellValueKind.EMPTY
    assert second[3].kind is CellValueKind.BOOLEAN and second[3].value is False


def test_read_rows_offset_and_limit() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["r1"], ["r2"], ["r3"], ["r4"]])
    adapter = TablesAdapter()

    page = adapter.read_rows(wb, "Tbl", offset=1, limit=2)
    assert page.total_rows == 4
    assert page.returned_rows == 2
    assert [row[0].value for row in page.rows] == ["r2", "r3"]


def test_read_rows_invalid_offset() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["r1"]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.read_rows(wb, "Tbl", offset=-1)


# ---------------------------------------------------------------------------
# Adapter-level: row mutations via ListRows
# ---------------------------------------------------------------------------
def test_append_rows_preserves_structure() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1]])
    adapter = TablesAdapter()

    result = adapter.append_rows(wb, "Tbl", [{"Name": "b", "Qty": 2}])

    assert result.action == "append_rows"
    assert result.before is not None and result.before.row_count == 1
    assert result.after.row_count == 2
    assert result.affected_rows == [2]
    assert result.after.name == "Tbl"
    assert [c.name for c in result.after.columns] == ["Name", "Qty"]
    rows = adapter.read_rows(wb, "Tbl").rows
    assert [row[0].value for row in rows] == ["a", "b"]
    assert rows[1][1].value == 2


def test_append_rows_unknown_header_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.append_rows(wb, "Tbl", [{"Nope": 1}])


def test_insert_row_position() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"], ["c"]])
    adapter = TablesAdapter()

    result = adapter.insert_row(wb, "Tbl", 2, {"Name": "b"})
    assert result.affected_rows == [2]
    names = [row[0].value for row in adapter.read_rows(wb, "Tbl").rows]
    assert names == ["a", "b", "c"]


def test_insert_row_out_of_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.insert_row(wb, "Tbl", 5)


def test_update_rows_by_index() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1], ["b", 2]])
    adapter = TablesAdapter()

    result = adapter.update_rows(wb, "Tbl", [{"index": 2, "values": {"Qty": 20}}])

    assert result.affected_rows == [2]
    rows = adapter.read_rows(wb, "Tbl").rows
    assert rows[1][1].value == 20
    assert rows[1][0].value == "b"  # untouched column preserved


def test_update_rows_where_ambiguous_rejected_with_preview() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1], ["a", 2]])
    adapter = TablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.update_rows(wb, "Tbl", [{"where": {"Name": "a"}, "values": {"Qty": 9}}])
    ctx = exc_info.value.context
    assert ctx["reason"] == "ambiguous"
    assert ctx["match_count"] == 2
    assert ctx["preview"]  # a non-empty preview is attached
    rows = adapter.read_rows(wb, "Tbl").rows
    assert [row[1].value for row in rows] == [1, 2]  # no mutation


def test_update_rows_where_unique_succeeds() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1], ["b", 2]])
    adapter = TablesAdapter()

    result = adapter.update_rows(wb, "Tbl", [{"where": {"Name": "b"}, "values": {"Qty": 7}}])
    assert result.affected_rows == [2]
    assert adapter.read_rows(wb, "Tbl").rows[1][1].value == 7


def test_delete_rows_by_indices() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"], ["b"], ["c"]])
    adapter = TablesAdapter()

    result = adapter.delete_rows(wb, "Tbl", indices=[1, 3])
    assert result.affected_rows == [1, 3]
    assert [row[0].value for row in adapter.read_rows(wb, "Tbl").rows] == ["b"]


def test_delete_rows_where_ambiguous_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"], ["a"]])
    adapter = TablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.delete_rows(wb, "Tbl", where={"Name": "a"})
    assert exc_info.value.context["reason"] == "ambiguous"
    assert adapter.read_rows(wb, "Tbl").total_rows == 2  # nothing deleted


def test_delete_rows_out_of_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.delete_rows(wb, "Tbl", indices=[5])


# ---------------------------------------------------------------------------
# Adapter-level: column mutations via ListColumns
# ---------------------------------------------------------------------------
def test_add_column_unique_header() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1]])
    adapter = TablesAdapter()

    result = adapter.add_column(wb, "Tbl", "Note")
    assert [c.name for c in result.after.columns] == ["Name", "Qty", "Note"]


def test_add_column_duplicate_header_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.add_column(wb, "Tbl", "name")  # case-insensitive duplicate
    assert "unique" in str(exc_info.value).lower()


def test_add_column_position() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["A", "B"], rows=[[1, 2]])
    adapter = TablesAdapter()

    result = adapter.add_column(wb, "Tbl", "Mid", position=2)
    assert [c.name for c in result.after.columns] == ["A", "Mid", "B"]


def test_delete_column() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["A", "B", "C"], rows=[[1, 2, 3]])
    adapter = TablesAdapter()

    result = adapter.delete_column(wb, "Tbl", "B")
    assert [c.name for c in result.after.columns] == ["A", "C"]
    rows = adapter.read_rows(wb, "Tbl").rows
    assert rows[0][0].value == 1
    assert rows[0][1].value == 3


def test_delete_column_missing_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["A"], rows=[[1]])
    adapter = TablesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.delete_column(wb, "Tbl", "Z")


# ---------------------------------------------------------------------------
# Adapter-level: create + resize
# ---------------------------------------------------------------------------
def test_create_table_and_resize() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.seed([["Name", "Qty"], ["a", 1], ["b", 2], [None, None]])

    adapter = TablesAdapter()
    created = adapter.create_table(wb, "Data", "A1:B4")
    assert created.action == "create_table"
    assert created.before is None
    assert created.after.row_count == 3
    assert [c.name for c in created.after.columns] == ["Name", "Qty"]

    resized = adapter.resize(wb, created.after.name, "A1:B5")
    assert resized.action == "resize"
    assert resized.after.row_count == 4


# ---------------------------------------------------------------------------
# Service-level: read paths + op_class
# ---------------------------------------------------------------------------
async def test_service_list_and_get_tables(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name"], rows=[["a"]])

    tables = await rig.service.list_tables(info.workbook_id)
    assert [t.name for t in tables] == ["Tbl"]

    one = await rig.service.get_table(info.workbook_id, "Tbl")
    assert one.name == "Tbl"

    read_submit = next(s for s in rig.worker.submits if s["name"] == "tables_list")
    assert read_submit["op_class"] == OperationClass.READ_ONLY


async def test_service_read_rows(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name", "Qty"], rows=[["a", 1], ["b", 2]])

    rows = await rig.service.read_table_rows(info.workbook_id, "Tbl", offset=1, limit=1)
    assert rows.returned_rows == 1
    assert rows.rows[0][0].value == "b"


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.list_tables("missing")


# ---------------------------------------------------------------------------
# Service-level: mutation guards
# ---------------------------------------------------------------------------
async def test_append_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name"], rows=[])
    rig.grant("table_append_rows")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.append_table_rows(info.workbook_id, "Tbl", [{"Name": "a"}])
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "table_append_rows"] == []


async def test_append_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name"], rows=[])
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.append_table_rows(info.workbook_id, "Tbl", [{"Name": "a"}])
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "table_append_rows"] == []


async def test_append_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    _new_table(rig.workbook(), columns=["Name"], rows=[])
    rig.add_snapshot(info.workbook_id)
    rig.grant("table_append_rows")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.append_table_rows(info.workbook_id, "Tbl", [{"Name": "a"}])
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_append_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name", "Qty"], rows=[["a", 1]])
    rig.add_snapshot(info.workbook_id)
    rig.grant("table_append_rows")

    result = await rig.service.append_table_rows(info.workbook_id, "Tbl", [{"Name": "b", "Qty": 2}])
    assert isinstance(result, TableMutationResult)
    assert result.after.row_count == 2
    mutate = next(s for s in rig.worker.submits if s["name"] == "table_append_rows")
    assert mutate["op_class"] == OperationClass.MUTATING


async def test_update_rows_ambiguous_rejected_at_service(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _new_table(rig.workbook(), columns=["Name", "Qty"], rows=[["a", 1], ["a", 2]])
    rig.add_snapshot(info.workbook_id)
    rig.grant("table_update_rows")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.update_table_rows(
            info.workbook_id, "Tbl", [{"where": {"Name": "a"}, "values": {"Qty": 9}}]
        )
    assert exc_info.value.context["reason"] == "ambiguous"


async def test_create_and_resize_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    sheet = rig.workbook().sheet("Data")
    sheet.seed([["Name", "Qty"], ["a", 1], [None, None]])

    rig.add_snapshot(info.workbook_id)
    rig.grant("table_create")
    created = await rig.service.create_table(info.workbook_id, "Data", "A1:B3")
    assert created.after.name.startswith("Table")

    rig.grant("table_resize")
    resized = await rig.service.resize_table(info.workbook_id, created.after.name, "A1:B4")
    assert resized.after.row_count == 3


# ---------------------------------------------------------------------------
# WP1: raise-site error taxonomy — deterministic cases must NOT be COM_BUSY
# ---------------------------------------------------------------------------
def _assert_non_retryable_taxonomy(
    error: ExcelVbaMcpError, *, code: str, category: str
) -> None:
    """Assert the serialized envelope is a specific, non-retryable, corrective error."""

    payload = error.to_dict()
    assert payload["errorCode"] == code
    assert payload["errorCode"] not in {"COM_BUSY", "RPC_E_CALL_REJECTED"}
    assert payload["errorCategory"] == category
    assert payload["errorCategory"] != "transient_com_failure"
    assert payload["isRetryable"] is False
    actions = payload["suggestedNextActions"]
    assert actions and all(isinstance(action, str) and action for action in actions)
    assert "retry" not in " ".join(actions).lower()


def test_missing_table_is_not_found() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().get_table(wb, "Nope")
    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context["reason"] == "not_found"


def test_unknown_header_is_not_found() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().append_rows(wb, "Tbl", [{"Nope": 1}])
    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")


def test_invalid_row_header_data_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().insert_row(wb, "Tbl", 5)  # out-of-range insert position
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "out_of_range"


def test_invalid_offset_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["r1"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().read_rows(wb, "Tbl", offset=-1)
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_offset"


def test_invalid_limit_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["r1"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().read_rows(wb, "Tbl", limit=0)
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_limit"


def test_out_of_range_index_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().delete_rows(wb, "Tbl", indices=[5])
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "out_of_range"


def test_ambiguous_row_selector_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name", "Qty"], rows=[["a", 1], ["a", 2]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().update_rows(wb, "Tbl", [{"where": {"Name": "a"}, "values": {"Qty": 9}}])
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "ambiguous"


def test_invalid_update_input_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().update_rows(
            wb, "Tbl", [{"index": 1, "where": {"Name": "a"}, "values": {"Name": "b"}}]
        )
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_selector"


def test_invalid_delete_input_is_invalid_range() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        TablesAdapter().delete_rows(wb, "Tbl")  # neither indices nor where
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_selector"


def test_duplicate_column_header_is_name_conflict() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _new_table(wb, columns=["Name"], rows=[["a"]])
    with pytest.raises(NameConflictError) as exc_info:
        TablesAdapter().add_column(wb, "Tbl", "name")  # case-insensitive duplicate
    _assert_non_retryable_taxonomy(exc_info.value, code="NAME_CONFLICT", category="validation")
    assert exc_info.value.context["reason"] == "duplicate"
