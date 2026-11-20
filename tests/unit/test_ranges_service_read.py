"""Unit tests for the read-only Range/Cells surface (task_042).

These tests exercise the production
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and
:class:`~excel_vba_mcp.adapters.excel.ranges_adapter.RangesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* ``range_get`` metadata for single-cell, row, column and block addresses;
* address normalization (multi-area / invalid / whole-row / whole-column);
* ``range_read_values`` 1-cell (scalar) vs multi-cell (2-D) shape and typed
  classification (string / number / bool / blank / date / error), with date
  typing driven strictly by ``NumberFormat`` (never guessed from a serial);
* bounds: oversized request rejected before reading, oversized response
  truncated with ``truncated``;
* ``range_read_formulas`` ``Formula2`` primary / ``Formula`` fallback and
  ``_xlfn`` prefix flagging;
* invalid sheet / invalid address -> typed ``InvalidRangeError``;
* service submits ``READ_ONLY`` and resolves ``workbook_id``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.ranges_adapter import RangesAdapter
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    MAX_RANGE_CELLS,
    MAX_RESPONSE_CELLS,
    RangeValues,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    OperationClass,
    RangeShape,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph — a worksheet owns a dense grid of cells plus a
# parallel grid of number formats; a FakeRange resolves an A1 address against it.
# ---------------------------------------------------------------------------
def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _col_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _parse_ref(ref: str) -> tuple[int, int]:
    cleaned = ref.replace("$", "")
    letters = "".join(ch for ch in cleaned if ch.isalpha())
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    return int(digits), _col_index(letters)


class FakeCells:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet

    def __call__(self, row: int, col: int) -> FakeCell:
        return FakeCell(self._sheet, row, col)


class FakeCell:
    def __init__(self, sheet: FakeSheet, row: int, col: int) -> None:
        self._sheet = sheet
        self.row = row
        self.col = col

    @property
    def Value2(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.get(self.row, self.col)

    @Value2.setter
    def Value2(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.set(self.row, self.col, value)

    @property
    def NumberFormat(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.get_format(self.row, self.col)

    @NumberFormat.setter
    def NumberFormat(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.set_format(self.row, self.col, value)


class FakeRange:
    def __init__(self, sheet: FakeSheet, top: int, left: int, height: int, width: int) -> None:
        self._sheet = sheet
        self._areas = 1
        self.Top = top
        self.Left = left
        self.Height = height
        self.Width = width

    # COM-ish surface ----------------------------------------------------
    @property
    def Row(self) -> int:  # noqa: N802 - COM casing
        return self.Top

    @property
    def Column(self) -> int:  # noqa: N802 - COM casing
        return self.Left

    @property
    def Rows(self) -> FakeRange:  # noqa: N802 - COM casing
        # Excel returns a range whose ``Count`` is the row count.
        return FakeRange(self._sheet, self.Top, self.Left, self.Height, 1)

    @property
    def Columns(self) -> FakeRange:  # noqa: N802 - COM casing
        # Excel returns a range whose ``Count`` is the column count.
        return FakeRange(self._sheet, self.Top, self.Left, 1, self.Width)

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self.Height * self.Width

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
    def Areas(self) -> FakeAreas:  # noqa: N802 - COM casing
        return FakeAreas(self._areas)

    @property
    def Value2(self) -> Any:  # noqa: N802 - COM casing
        grid = tuple(
            tuple(self._sheet.get(r, c) for c in range(self.Left, self.Left + self.Width))
            for r in range(self.Top, self.Top + self.Height)
        )
        if self.Height == 1 and self.Width == 1:
            return grid[0][0]
        return grid

    @property
    def Formula2(self) -> Any:  # noqa: N802 - COM casing
        grid = tuple(
            tuple(self._sheet.get_formula(r, c) for c in range(self.Left, self.Left + self.Width))
            for r in range(self.Top, self.Top + self.Height)
        )
        if self.Height == 1 and self.Width == 1:
            return grid[0][0]
        return grid

    @property
    def Formula(self) -> Any:  # noqa: N802 - COM casing
        return self.Formula2

    @property
    def NumberFormat(self) -> Any:  # noqa: N802 - COM casing
        formats = {
            self._sheet.get_format(r, c)
            for r in range(self.Top, self.Top + self.Height)
            for c in range(self.Left, self.Left + self.Width)
        }
        if len(formats) == 1:
            return formats.pop()
        return None

    def Cells(self, row: int, col: int) -> FakeCell:  # noqa: N802 - COM casing
        return FakeCell(self._sheet, self.Top + row - 1, self.Left + col - 1)


class FakeAreas:
    def __init__(self, count: int) -> None:
        self._count = count

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self._count


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name
        self._grid: dict[tuple[int, int], Any] = {}
        self._formats: dict[tuple[int, int], Any] = {}
        self._formulas: dict[tuple[int, int], Any] = {}
        self.Cells = FakeCells(self)

    def get(self, row: int, col: int) -> Any:
        return self._grid.get((row, col))

    def set(self, row: int, col: int, value: Any) -> None:
        self._grid[(row, col)] = value

    def get_format(self, row: int, col: int) -> Any:
        return self._formats.get((row, col), "General")

    def set_format(self, row: int, col: int, value: Any) -> None:
        self._formats[(row, col)] = value

    def get_formula(self, row: int, col: int) -> Any:
        return self._formulas.get((row, col), self._grid.get((row, col)))

    def set_formula(self, row: int, col: int, value: Any) -> None:
        self._formulas[(row, col)] = value

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
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )
        self.service = RangesService(
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

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


# ---------------------------------------------------------------------------
# Fake sheet helpers
# ---------------------------------------------------------------------------
def _seed_cells(sheet: FakeSheet) -> None:
    """Populate a sheet with one cell of each classification kind.

    Layout (column B, rows 1..9) mirrors the fixture design so the unit test is
    self-describing: string, number, bool, blank, date-formatted number, a
    date-serial-looking number with a General format, an OLE error code and a
    literal error string.
    """

    sheet.set(1, 2, "hello world")  # B1 string
    sheet.set(2, 2, 42.0)  # B2 number
    sheet.set(3, 2, True)  # B3 boolean
    # B4 intentionally blank
    sheet.set(5, 2, 46024.0)  # B5 date-formatted number
    sheet.set_format(5, 2, "yyyy-mm-dd")
    sheet.set(6, 2, 45000.0)  # B6 date-shaped serial, General format -> number
    sheet.set(7, 2, -2146826281)  # B7 OLE error code (#DIV/0!)
    sheet.set(8, 2, "#N/A")  # B8 literal error string
    sheet.set(9, 2, 3.14)  # B9 number


def _seed_block(sheet: FakeSheet) -> None:
    """Populate a 3x3 numeric block at F1:H3."""

    for r in range(3):
        for c in range(3):
            sheet.set(1 + r, 6 + c, float(1 + r * 3 + c))


# ---------------------------------------------------------------------------
# Adapter-level: range_get metadata
# ---------------------------------------------------------------------------
def test_range_get_scalar_metadata() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    wb.sheet("Data").set(1, 1, 7)
    adapter = RangesAdapter()

    resolved = adapter.resolve_range(wb, "Data", "A1")
    assert resolved.sheet == "Data"
    assert resolved.rows == 1
    assert resolved.cols == 1
    assert resolved.shape is RangeShape.SCALAR


def test_range_get_block_metadata() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    resolved = adapter.resolve_range(wb, "Data", "A1:C10")
    assert resolved.rows == 10
    assert resolved.cols == 3
    assert resolved.shape is RangeShape.BLOCK
    assert resolved.address.replace("$", "") == "A1:C10"


def test_range_get_row_and_column_shapes() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    row = adapter.resolve_range(wb, "Data", "A1:C1")
    assert row.shape is RangeShape.ROW
    column = adapter.resolve_range(wb, "Data", "A1:A3")
    assert column.shape is RangeShape.COLUMN


def test_range_get_invalid_sheet_raises() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.resolve_range(wb, "Nope", "A1")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_range_get_invalid_address_raises() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.resolve_range(wb, "Data", "NotAnAddress")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Adapter-level: read_values typing + shape
# ---------------------------------------------------------------------------
def test_read_values_scalar_is_single_cell() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set(2, 2, "hello world")
    adapter = RangesAdapter()

    result = adapter.read_values(wb, "Data", "B2")
    assert isinstance(result, RangeValues)
    assert result.rows == 1
    assert result.cols == 1
    assert result.shape is RangeShape.SCALAR
    assert result.values[0][0].kind is CellValueKind.STRING
    assert result.values[0][0].value == "hello world"


def test_read_values_classifies_each_kind() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    _seed_cells(sheet)
    adapter = RangesAdapter()

    result = adapter.read_values(wb, "Data", "B1:B9")
    assert result.rows == 9
    assert result.cols == 1
    assert result.shape is RangeShape.COLUMN
    kinds = [row[0].kind for row in result.values]
    assert kinds == [
        CellValueKind.STRING,
        CellValueKind.NUMBER,
        CellValueKind.BOOLEAN,
        CellValueKind.EMPTY,
        CellValueKind.DATE,
        CellValueKind.NUMBER,  # 45000.0 with General format is NOT a date
        CellValueKind.ERROR,
        CellValueKind.ERROR,
        CellValueKind.NUMBER,
    ]
    assert result.values[4][0].value == "2026-01-02T00:00:00"
    assert result.values[6][0].value == "#DIV/0!"
    assert result.values[7][0].value == "#N/A"


def test_read_values_date_only_when_date_formatted() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    # Same serial, two formats: date-formatted -> DATE, General -> NUMBER.
    sheet.set(1, 1, 46024.0)
    sheet.set_format(1, 1, "yyyy-mm-dd")
    sheet.set(2, 1, 46024.0)
    sheet.set_format(2, 1, "General")
    adapter = RangesAdapter()

    result = adapter.read_values(wb, "Data", "A1:A2")
    assert result.values[0][0].kind is CellValueKind.DATE
    assert result.values[1][0].kind is CellValueKind.NUMBER


def test_read_values_2d_block() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    _seed_block(sheet)
    adapter = RangesAdapter()

    result = adapter.read_values(wb, "Data", "F1:H3")
    assert result.rows == 3
    assert result.cols == 3
    assert result.shape is RangeShape.BLOCK
    assert [row[0].value for row in result.values] == [1.0, 4.0, 7.0]
    assert [row[2].value for row in result.values] == [3.0, 6.0, 9.0]


# ---------------------------------------------------------------------------
# Adapter-level: read_formulas
# ---------------------------------------------------------------------------
def test_read_formulas_formula2_primary() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_formula(1, 1, "=1+2")
    sheet.set(1, 1, 3.0)
    adapter = RangesAdapter()

    result = adapter.read_formulas(wb, "Data", "A1")
    assert result.rows == 1
    assert result.formulas[0][0].formula2 == "=1+2"
    assert result.formulas[0][0].formula is None  # Formula == Formula2 -> omitted
    assert result.notes == []


def test_read_formulas_flags_xlfn_prefix() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_formula(1, 1, "=_xlfn.CONCAT(A2:A3)")
    adapter = RangesAdapter()

    result = adapter.read_formulas(wb, "Data", "A1")
    assert result.formulas[0][0].formula2 == "=_xlfn.CONCAT(A2:A3)"
    assert any("_xlfn" in note for note in result.notes)


# ---------------------------------------------------------------------------
# Adapter-level: bounds / truncation
# ---------------------------------------------------------------------------
def test_read_values_rejects_oversized_request_before_reading() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    # A range larger than MAX_RANGE_CELLS must be rejected up front.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.read_values(wb, "Data", "A1:XFD1048576")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells"] > MAX_RANGE_CELLS


def test_read_values_truncates_response() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    # A block just over MAX_RESPONSE_CELLS (rows exceed the response cell budget
    # when multiplied by the 3 columns).
    rows = MAX_RESPONSE_CELLS // 3 + 2
    for r in range(1, rows + 1):
        for c in range(1, 4):
            sheet.set(r, c, float(r))
    adapter = RangesAdapter()

    result = adapter.read_values(wb, "Data", f"A1:C{rows}")
    assert result.truncated is True
    assert result.rows * result.cols <= MAX_RESPONSE_CELLS


def test_read_values_offset_and_limit() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    for r in range(1, 5):
        sheet.set(r, 1, f"r{r}")
    adapter = RangesAdapter()

    page = adapter.read_values(wb, "Data", "A1:A4", offset=1, limit=2)
    assert page.rows == 2
    assert [row[0].value for row in page.values] == ["r2", "r3"]


def test_read_values_invalid_offset() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.read_values(wb, "Data", "A1", offset=-1)


def test_read_values_multi_area_rejected() -> None:
    class MultiAreaSheet(FakeSheet):
        def Range(self, *args: Any) -> Any:  # noqa: N802 - COM casing
            range_obj = super().Range(*args)
            range_obj._areas = 2
            return range_obj

    sheet = MultiAreaSheet("Data")
    wb2 = FakeWorkbook([sheet])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.read_values(wb2, "Data", "A1:B2")
    assert exc_info.value.context["areas"] == 2


# ---------------------------------------------------------------------------
# Service-level: read paths + op_class + workbook resolution
# ---------------------------------------------------------------------------
async def test_service_get_range(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.workbook().sheet("Data").set(1, 1, 7)

    resolved = await rig.service.get_range(info.workbook_id, "Data", "A1:B2")
    assert resolved.rows == 2
    assert resolved.cols == 2

    submit = next(s for s in rig.worker.submits if s["name"] == "range_get")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_read_values(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_cells(rig.workbook().sheet("Data"))

    result = await rig.service.read_values(info.workbook_id, "Data", "B1:B9")
    assert result.rows == 9
    assert result.values[0][0].kind is CellValueKind.STRING
    submit = next(s for s in rig.worker.submits if s["name"] == "range_read_values")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_read_formulas(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.workbook().sheet("Data").set_formula(1, 1, "=1+2")

    result = await rig.service.read_formulas(info.workbook_id, "Data", "A1")
    assert result.formulas[0][0].formula2 == "=1+2"


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.get_range("missing", "Data", "A1")


async def test_service_invalid_sheet_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.get_range(info.workbook_id, "Nope", "A1")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
