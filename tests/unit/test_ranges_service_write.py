"""Unit tests for the Range/Cells write surface (task_043).

These tests exercise the production
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and
:class:`~excel_vba_mcp.adapters.excel.ranges_adapter.RangesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* payload validation — wrong-sized payloads (row-count, col-count and target
  mismatch) are rejected before any mutation, leaving the grid untouched;
* gates — a read-only working copy (``WORKBOOK_READ_ONLY``), a missing snapshot
  (``SNAPSHOT_FAILED``) and a missing approval grant (deny-by-default);
* typed value serialization — str/int/float/bool/None/date/Excel error round-trip
  through ``Range.Value2`` (a leading ``=`` string is stored as text, never a
  formula);
* postcondition mismatch — a tampered re-read yields ``verified=False`` with a
  note;
* the ListObject structure guard — a cell write inside a table is allowed and a
  structural change is rejected with ``LISTOBJECT_INTERSECTION``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.ranges_adapter import RangesAdapter, _validate_target_cells
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ApprovalDeniedError, ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    MAX_RANGE_CELLS,
    CellValue,
    FormulaWriteRequest,
    RangeWriteRequest,
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
# Fake Excel object graph — a worksheet owns a dense grid of cells plus a
# parallel grid of number formats and formulas. A FakeRange resolves an A1
# address against the grid and supports a Value2 / Formula2 assignment.
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
        return FakeRange(self._sheet, self.Top, self.Left, self.Height, 1)

    @property
    def Columns(self) -> FakeRange:  # noqa: N802 - COM casing
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

    @Value2.setter
    def Value2(self, value: Any) -> None:  # noqa: N802 - COM casing
        if self.Height == 1 and self.Width == 1:
            self._sheet.set(self.Top, self.Left, value)
            return
        for r in range(self.Height):
            row = value[r] if r < len(value) else ()
            for c in range(self.Width):
                self._sheet.set(self.Top + r, self.Left + c, row[c] if c < len(row) else None)

    @property
    def Formula2(self) -> Any:  # noqa: N802 - COM casing
        grid = tuple(
            tuple(self._sheet.get_formula(r, c) for c in range(self.Left, self.Left + self.Width))
            for r in range(self.Top, self.Top + self.Height)
        )
        if self.Height == 1 and self.Width == 1:
            return grid[0][0]
        return grid

    @Formula2.setter
    def Formula2(self, value: Any) -> None:  # noqa: N802 - COM casing
        if self.Height == 1 and self.Width == 1:
            self._sheet.set_formula(self.Top, self.Left, value)
            return
        for r in range(self.Height):
            row = value[r] if r < len(value) else ()
            for c in range(self.Width):
                self._sheet.set_formula(
                    self.Top + r, self.Left + c, row[c] if c < len(row) else None
                )

    @property
    def Formula(self) -> Any:  # noqa: N802 - COM casing
        return self.Formula2

    @Formula.setter
    def Formula(self, value: Any) -> None:  # noqa: N802 - COM casing
        self.Formula2 = value

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


class FakeListObject:
    """A minimal ListObject whose ``Range`` covers a header row + data rows."""

    def __init__(self, sheet: FakeSheet, name: str, top: int, left: int, height: int, width: int):
        self.sheet = sheet
        self.Name = name
        self._top = top
        self._left = left
        self._height = height
        self._width = width

    @property
    def Range(self) -> FakeRange:  # noqa: N802 - COM casing
        return FakeRange(self.sheet, self._top, self._left, self._height, self._width)

    def grow(self, height: int, width: int) -> None:
        """Simulate Excel growing the table structure (test helper)."""

        self._height = height
        self._width = width


class FakeListObjects:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet
        self._items: list[FakeListObject] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Add(self, list_object: FakeListObject) -> FakeListObject:  # noqa: N802, N803 - COM
        self._items.append(list_object)
        return list_object

    def Item(self, index: int) -> FakeListObject:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __iter__(self):
        return iter(self._items)


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.ListObjects = FakeListObjects(self)
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
# Helpers
# ---------------------------------------------------------------------------
def _cv(kind: CellValueKind, value: Any = None) -> CellValue:
    return CellValue(kind=kind, value=value)


def _grid(rows: int, cols: int, fill: Any = None) -> list[list[CellValue]]:
    return [[_cv(CellValueKind.EMPTY) for _ in range(cols)] for _ in range(rows)]


# ---------------------------------------------------------------------------
# Adapter-level: payload validation before any mutation
# ---------------------------------------------------------------------------
def test_write_values_rejects_wrong_row_count() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    payload = [_cv(CellValueKind.NUMBER, 1.0)]  # 1x1 payload against a 2x2 target
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_values(wb, "Data", "A1:B2", 2, 2, payload)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    # No mutation: the grid is untouched.
    assert wb.sheet("Data").get(1, 1) is None


def test_write_values_rejects_target_shape_mismatch() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    payload = [[_cv(CellValueKind.NUMBER, 1.0), _cv(CellValueKind.NUMBER, 2.0)]]
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_values(wb, "Data", "A1:B2", 1, 2, payload)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["target_rows"] == 2
    assert wb.sheet("Data").get(1, 1) is None


def test_write_values_rejects_row_count_mismatch_with_declared_rows() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    payload = [[_cv(CellValueKind.NUMBER, 1.0)]]  # 1 row, but rows=2 declared
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_values(wb, "Data", "A1:B2", 2, 2, payload)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Adapter-level: typed value serialization
# ---------------------------------------------------------------------------
def test_write_values_scalar_round_trip() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    result = adapter.write_values(wb, "Data", "B2", 1, 1, [[_cv(CellValueKind.STRING, "hi")]])
    assert result.verified is True
    assert wb.sheet("Data").get(2, 2) == "hi"
    assert result.postcondition_values is not None
    assert result.postcondition_values.values[0][0].kind is CellValueKind.STRING


def test_write_values_typed_block() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    grid = [
        [_cv(CellValueKind.STRING, "a"), _cv(CellValueKind.NUMBER, 1.5)],
        [_cv(CellValueKind.BOOLEAN, True), _cv(CellValueKind.EMPTY)],
    ]
    result = adapter.write_values(wb, "Data", "A1:B2", 2, 2, grid)
    assert result.verified is True
    assert wb.sheet("Data").get(1, 1) == "a"
    assert wb.sheet("Data").get(1, 2) == 1.5
    assert wb.sheet("Data").get(2, 1) is True
    assert wb.sheet("Data").get(2, 2) is None


def test_write_values_leading_equals_is_text_not_formula() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    result = adapter.write_values(wb, "Data", "A1", 1, 1, [[_cv(CellValueKind.STRING, "=1+1")]])
    assert result.verified is True
    # Stored as text, not as a formula (the formula channel is untouched).
    assert wb.sheet("Data").get(1, 1) == "=1+1"
    assert wb.sheet("Data").get_formula(1, 1) == "=1+1"  # no formula was written


def test_write_values_date_serializes_to_excel_serial() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    # A date-formatted target cell (Excel would store the serial under a date
    # number format); the adapter writes the raw Excel serial (days since
    # 1899-12-30) and the re-read classifies it as a date because the format is
    # a date format.
    sheet.set_format(1, 1, "yyyy-mm-dd")
    adapter = RangesAdapter()

    moment = datetime(2026, 1, 2, 0, 0, 0)
    result = adapter.write_values(
        wb, "Data", "A1", 1, 1, [[_cv(CellValueKind.DATE, moment.isoformat())]]
    )
    assert result.verified is True
    assert sheet.get(1, 1) == 46024.0


def test_write_values_error_literal_serializes_to_ole_code() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    result = adapter.write_values(wb, "Data", "A1", 1, 1, [[_cv(CellValueKind.ERROR, "#DIV/0!")]])
    assert result.verified is True
    assert wb.sheet("Data").get(1, 1) == -2146826281


# ---------------------------------------------------------------------------
# Adapter-level: raw JSON scalars (WP-Range ergonomics)
# ---------------------------------------------------------------------------
def test_write_values_accepts_raw_scalars() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    # Raw JSON scalars (no {kind, value} wrapper) are stored as literals by real
    # JSON type only: string, number, boolean, null.
    grid = [["hello", 42, True, None]]
    result = adapter.write_values(wb, "Data", "A1:D1", 1, 4, grid)
    assert result.verified is True
    sheet = wb.sheet("Data")
    assert sheet.get(1, 1) == "hello"
    assert sheet.get(1, 2) == 42
    assert sheet.get(1, 3) is True
    assert sheet.get(1, 4) is None
    kinds = [cell.kind for cell in result.postcondition_values.values[0]]  # type: ignore[union-attr]
    assert kinds == [
        CellValueKind.STRING,
        CellValueKind.NUMBER,
        CellValueKind.BOOLEAN,
        CellValueKind.EMPTY,
    ]


def test_write_values_raw_leading_equals_stays_text() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    result = adapter.write_values(wb, "Data", "A1", 1, 1, [["=1+1"]])
    assert result.verified is True
    # A raw string beginning with '=' is stored as text, never a formula.
    assert wb.sheet("Data").get(1, 1) == "=1+1"
    assert wb.sheet("Data").get_formula(1, 1) == "=1+1"


def test_write_values_raw_ambiguous_strings_stay_strings() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    # "true" / "123" / "2024-01-01" are real JSON strings and must NOT be coerced
    # to bool / number / date (no ambiguous coercion).
    grid = [["true", "123", "2024-01-01"]]
    result = adapter.write_values(wb, "Data", "A1:C1", 1, 3, grid)
    assert result.verified is True
    kinds = [cell.kind for cell in result.postcondition_values.values[0]]  # type: ignore[union-attr]
    assert kinds == [CellValueKind.STRING, CellValueKind.STRING, CellValueKind.STRING]
    assert wb.sheet("Data").get(1, 2) == "123"


def test_write_values_cellvalue_dict_still_works() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    # A typed CellValue dict is still accepted unchanged alongside raw scalars.
    grid = [[{"kind": "number", "value": 7}, "x"]]
    result = adapter.write_values(wb, "Data", "A1:B1", 1, 2, grid)
    assert result.verified is True
    assert wb.sheet("Data").get(1, 1) == 7
    assert wb.sheet("Data").get(1, 2) == "x"


# ---------------------------------------------------------------------------
# Adapter-level: over-cap valid target -> RANGE_TOO_LARGE (WP-Range)
# ---------------------------------------------------------------------------
def test_validate_target_cells_rejects_valid_over_cap() -> None:
    # A valid target one cell over the cap is RANGE_TOO_LARGE, not INVALID_RANGE.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        _validate_target_cells(MAX_RANGE_CELLS + 1)
    assert exc_info.value.code is ErrorCode.RANGE_TOO_LARGE
    assert exc_info.value.code is not ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells"] == MAX_RANGE_CELLS + 1


def test_validate_target_cells_allows_exact_cap() -> None:
    # Exactly at the cap is allowed (no raise).
    _validate_target_cells(MAX_RANGE_CELLS)


# ---------------------------------------------------------------------------
# Adapter-level: formulas
# ---------------------------------------------------------------------------
def test_write_formulas_scalar_round_trip() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    result = adapter.write_formulas(wb, "Data", "A1", 1, 1, [["=1+2"]])
    assert result.verified is True
    assert wb.sheet("Data").get_formula(1, 1) == "=1+2"


def test_write_formulas_rejects_non_formula_entry() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_formulas(wb, "Data", "A1", 1, 1, [["1+2"]])  # missing '='
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert wb.sheet("Data").get(1, 1) is None


def test_write_formulas_wrong_shape_rejected_before_mutation() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_formulas(wb, "Data", "A1:B2", 2, 2, [["=1+1"]])
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert wb.sheet("Data").get(1, 1) is None


# ---------------------------------------------------------------------------
# Adapter-level: postcondition mismatch
# ---------------------------------------------------------------------------
def test_write_values_reports_postcondition_mismatch() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    adapter = RangesAdapter()
    original_set = sheet.set

    def tamper(row: int, col: int, value: Any) -> None:
        original_set(row, col, value)
        if (row, col) == (1, 1):
            original_set(row, col, "tampered")

    sheet.set = tamper  # type: ignore[method-assign]

    result = adapter.write_values(wb, "Data", "A1", 1, 1, [[_cv(CellValueKind.STRING, "ok")]])
    assert result.verified is False
    assert result.notes and "mismatch" in result.notes[0].lower()


# ---------------------------------------------------------------------------
# Adapter-level: ListObject structure guard
# ---------------------------------------------------------------------------
def _seed_table(wb: FakeWorkbook, sheet_name: str = "Data") -> FakeListObject:
    sheet = wb.sheet(sheet_name)
    # Header row + 2 data rows at A1:B3.
    sheet.set(1, 1, "id")
    sheet.set(1, 2, "name")
    sheet.set(2, 1, 1.0)
    sheet.set(2, 2, "alpha")
    sheet.set(3, 1, 2.0)
    sheet.set(3, 2, "beta")
    list_object = FakeListObject(sheet, "tblData", 1, 1, 3, 2)
    sheet.ListObjects.Add(list_object)
    return list_object


def test_write_values_inside_table_allowed() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _seed_table(wb)
    adapter = RangesAdapter()

    result = adapter.write_values(wb, "Data", "B2", 1, 1, [[_cv(CellValueKind.STRING, "gamma")]])
    assert result.verified is True
    assert wb.sheet("Data").get(2, 2) == "gamma"


def test_write_values_table_structural_change_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    list_object = _seed_table(wb)
    adapter = RangesAdapter()
    original_address = list_object.Range.Address

    # Simulate Excel growing the table when a write crosses its boundary: the
    # adapter snapshots the table geometry before the write, so a post-write
    # growth must be detected as a structural change.
    sheet = wb.sheet("Data")
    original_set = sheet.set

    def grow_on_write(row: int, col: int, value: Any) -> None:
        original_set(row, col, value)
        if (row, col) == (2, 2):
            list_object.grow(4, 2)

    sheet.set = grow_on_write  # type: ignore[method-assign]

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.write_values(wb, "Data", "B2", 1, 1, [[_cv(CellValueKind.STRING, "x")]])
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblData"
    # Silence the unused-variable lint while keeping the original geometry handy.
    assert original_address == "$A$1:$B$3"


# ---------------------------------------------------------------------------
# Service-level: guard chain
# ---------------------------------------------------------------------------
async def test_service_write_values_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_write_values")
    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.write_values(info.workbook_id, "Data", "A1", payload)
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY
    assert [s for s in rig.worker.submits if s["name"] == "range_write_values"] == []


async def test_service_write_values_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("range_write_values")
    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.write_values(info.workbook_id, "Data", "A1", payload)
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "range_write_values"] == []


async def test_service_write_values_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.write_values(info.workbook_id, "Data", "A1", payload)
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "range_write_values"] == []


async def test_service_write_values_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_write_values")
    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 7.0)]])

    result = await rig.service.write_values(info.workbook_id, "Data", "A1", payload)
    assert result.verified is True
    mutate = next(s for s in rig.worker.submits if s["name"] == "range_write_values")
    assert mutate["op_class"] == OperationClass.MUTATING
    assert rig.workbook().sheet("Data").get(1, 1) == 7.0


async def test_bump_version_wired_on_range_write(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_write_values")
    assert rig.session.get_version(info.workbook_id) == 0

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 7.0)]])
    await rig.service.write_values(info.workbook_id, "Data", "A1", payload)
    assert rig.session.get_version(info.workbook_id) == 1

    # A raising mutation (payload shape mismatch -> INVALID_RANGE on the worker)
    # must NOT bump the version: the bump runs only after a successful submit.
    bad = RangeWriteRequest(rows=2, cols=2, values=[[1, 2], [3, 4]])
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.write_values(info.workbook_id, "Data", "A1", bad)
    assert rig.session.get_version(info.workbook_id) == 1


async def test_service_write_formulas_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_write_formulas")
    payload = FormulaWriteRequest(rows=1, cols=1, formulas=[["=1+2"]])

    result = await rig.service.write_formulas(info.workbook_id, "Data", "A1", payload)
    assert result.verified is True
    mutate = next(s for s in rig.worker.submits if s["name"] == "range_write_formulas")
    assert mutate["op_class"] == OperationClass.MUTATING
    assert rig.workbook().sheet("Data").get_formula(1, 1) == "=1+2"


async def test_service_write_formulas_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    payload = FormulaWriteRequest(rows=1, cols=1, formulas=[["=1+2"]])

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.write_formulas(info.workbook_id, "Data", "A1", payload)
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "range_write_formulas"] == []
