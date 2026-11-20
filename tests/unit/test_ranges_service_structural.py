"""Unit tests for the structural Range/Cells surface (task_045).

These tests exercise the production
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and
:class:`~excel_vba_mcp.adapters.excel.ranges_adapter.RangesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* explicit destination (copy) and explicit shift (insert/delete) validation —
  a missing/blank/unknown argument is rejected before any mutation;
* copy to an anchor cell, copy to a matching block and cross-sheet copy, with a
  value/dimension postcondition;
* the hard ListObject guard — copy/insert/delete/merge targets that intersect a
  ``ListObject.Range`` are rejected with ``LISTOBJECT_INTERSECTION`` and name the
  owning table;
* the merge data-loss gate — non-anchor data blocks the merge and the preview
  lists the affected cells until ``allow_data_loss`` is set;
* unmerge only when a merged area actually intersects the target;
* protected sheets, postcondition mismatch reporting and the
  snapshot / read-only / approval gates.
"""

from __future__ import annotations

import inspect
import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.ranges_adapter import RangesAdapter
from excel_vba_mcp.api import range_tools
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
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
# Fake Excel object graph — a worksheet owns a dense value/formula/format grid
# plus a merged-cell map (cell -> anchor). A FakeRange resolves an A1 address
# against the grid and supports Copy / Insert / Delete / Merge / UnMerge.
# ---------------------------------------------------------------------------
_MAX_ROW = 200
_MAX_COL = 50


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


def _addr(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


class FakeCell:
    """One cell proxy (Value2 / Formula / NumberFormat / MergeCells / MergeArea)."""

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
    def Formula(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.get_formula(self.row, self.col)

    @Formula.setter
    def Formula(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.set_formula(self.row, self.col, value)

    @property
    def NumberFormat(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.get_format(self.row, self.col)

    @NumberFormat.setter
    def NumberFormat(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.set_format(self.row, self.col, value)

    @property
    def MergeCells(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.is_merged(self.row, self.col)

    @property
    def MergeArea(self) -> Any:  # noqa: N802 - COM casing
        anchor = self._sheet.merge_anchor(self.row, self.col)
        if anchor is None:
            return FakeRange(self._sheet, self.row, self.col, 1, 1)
        top, left, height, width = self._sheet.merge_area(anchor)
        return FakeRange(self._sheet, top, left, height, width)


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
        if self.Height == 1 and self.Width == 1:
            return f"${_col_letter(self.Left)}${self.Top}"
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
        if self.Height == 1:
            return grid[0]
        return grid

    @Value2.setter
    def Value2(self, value: Any) -> None:  # noqa: N802 - COM casing
        if self.Height == 1 and self.Width == 1:
            self._sheet.set(self.Top, self.Left, value)
            return
        if self.Height == 1:
            for c in range(self.Width):
                self._sheet.set(self.Top, self.Left + c, value[c] if c < len(value) else None)
            return
        for r in range(self.Height):
            row = value[r] if r < len(value) else ()
            for c in range(self.Width):
                self._sheet.set(self.Top + r, self.Left + c, row[c] if c < len(row) else None)

    @property
    def Formula(self) -> Any:  # noqa: N802 - COM casing
        grid = tuple(
            tuple(self._sheet.formula_text(r, c) for c in range(self.Left, self.Left + self.Width))
            for r in range(self.Top, self.Top + self.Height)
        )
        if self.Height == 1 and self.Width == 1:
            return grid[0][0]
        return grid

    @Formula.setter
    def Formula(self, value: Any) -> None:  # noqa: N802 - COM casing
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
    def NumberFormat(self) -> Any:  # noqa: N802 - COM casing
        formats = {
            self._sheet.get_format(r, c)
            for r in range(self.Top, self.Top + self.Height)
            for c in range(self.Left, self.Left + self.Width)
        }
        if len(formats) == 1:
            return formats.pop()
        return None

    @property
    def MergeCells(self) -> Any:  # noqa: N802 - COM casing
        cells = [
            (r, c)
            for r in range(self.Top, self.Top + self.Height)
            for c in range(self.Left, self.Left + self.Width)
        ]
        merged = [(r, c) for r, c in cells if self._sheet.is_merged(r, c)]
        if not merged:
            return False
        anchors = {self._sheet.merge_anchor(r, c) for r, c in merged}
        if len(anchors) == 1 and len(merged) == len(cells):
            return True
        return None

    def Cells(self, row: int, col: int) -> FakeCell:  # noqa: N802 - COM casing
        return FakeCell(self._sheet, self.Top + row - 1, self.Left + col - 1)

    # Structural operations ----------------------------------------------
    def Copy(self, destination: FakeRange) -> None:  # noqa: N802 - COM casing
        destination._sheet.copy_block_from(  # noqa: SLF001 - test double
            self._sheet,
            self.Top,
            self.Left,
            self.Height,
            self.Width,
            destination.Top,
            destination.Left,
        )

    def Insert(self, Shift: int | None = None, CopyOrigin: int | None = None) -> None:  # noqa: N802, N803
        del CopyOrigin  # the fake formats from the left/above by definition
        self._sheet.shift_cells(
            self.Top, self.Left, self.Height, self.Width, shift=Shift, inserting=True
        )

    def Delete(self, Shift: int | None = None) -> None:  # noqa: N802, N803
        self._sheet.shift_cells(
            self.Top, self.Left, self.Height, self.Width, shift=Shift, inserting=False
        )

    def Merge(self, Across: bool = False) -> None:  # noqa: N802, N803 - COM casing
        self._sheet.merge_block(self.Top, self.Left, self.Height, self.Width, across=Across)

    def UnMerge(self) -> None:  # noqa: N802 - COM casing
        self._sheet.unmerge_block(self.Top, self.Left, self.Height, self.Width)


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


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.ListObjects = FakeListObjects(self)
        self.ProtectContents = False
        self._grid: dict[tuple[int, int], Any] = {}
        self._formats: dict[tuple[int, int], Any] = {}
        self._formulas: dict[tuple[int, int], Any] = {}
        self._merged: dict[tuple[int, int], tuple[int, int]] = {}
        self.Cells = FakeCells(self)  # noqa: N815 - COM casing

    # Grid access --------------------------------------------------------
    def get(self, row: int, col: int) -> Any:
        return self._grid.get((row, col))

    def set(self, row: int, col: int, value: Any) -> None:
        if value is None:
            self._grid.pop((row, col), None)
        else:
            self._grid[(row, col)] = value

    def get_format(self, row: int, col: int) -> Any:
        return self._formats.get((row, col), "General")

    def set_format(self, row: int, col: int, value: Any) -> None:
        self._formats[(row, col)] = value

    def get_formula(self, row: int, col: int) -> Any:
        return self._formulas.get((row, col))

    def set_formula(self, row: int, col: int, value: Any) -> None:
        if value is None:
            self._formulas.pop((row, col), None)
        else:
            self._formulas[(row, col)] = value

    def formula_text(self, row: int, col: int) -> Any:
        """Return the formula if any, else the value (Excel ``Formula`` semantics)."""

        formula = self._formulas.get((row, col))
        if formula is not None:
            return formula
        return self._grid.get((row, col))

    # Merged cells -------------------------------------------------------
    def is_merged(self, row: int, col: int) -> bool:
        return (row, col) in self._merged

    def merge_anchor(self, row: int, col: int) -> tuple[int, int] | None:
        return self._merged.get((row, col))

    def merge_area(self, anchor: tuple[int, int]) -> tuple[int, int, int, int]:
        """Return ``(top, left, height, width)`` for the merged area at ``anchor``."""

        cells = [cell for cell, target in self._merged.items() if target == anchor]
        rows = [cell[0] for cell in cells]
        cols = [cell[1] for cell in cells]
        return min(rows), min(cols), max(rows) - min(rows) + 1, max(cols) - min(cols) + 1

    def merge_block(self, top: int, left: int, height: int, width: int, *, across: bool) -> None:
        """Merge the block (one area, or one area per row when ``across``)."""

        for r in range(top, top + height):
            for c in range(left, left + width):
                anchor = (r, left) if across else (top, left)
                self._merged[(r, c)] = anchor
                if (r, c) != anchor:
                    self._grid.pop((r, c), None)
                    self._formulas.pop((r, c), None)

    def unmerge_block(self, top: int, left: int, height: int, width: int) -> None:
        """Unmerge every merged area that intersects the block (Excel semantics)."""

        anchors = {
            anchor
            for (row, col), anchor in self._merged.items()
            if top <= row < top + height and left <= col < left + width
        }
        if not anchors:
            return
        self._merged = {
            cell: anchor for cell, anchor in self._merged.items() if anchor not in anchors
        }

    # Structural shifts --------------------------------------------------
    def copy_block_from(
        self,
        source_sheet: FakeSheet,
        src_top: int,
        src_left: int,
        height: int,
        width: int,
        dst_top: int,
        dst_left: int,
    ) -> None:
        for r in range(height):
            for c in range(width):
                source = (src_top + r, src_left + c)
                target = (dst_top + r, dst_left + c)
                self._grid[target] = source_sheet._grid.get(source)  # noqa: SLF001 - test double
                self._formats[target] = source_sheet._formats.get(source, "General")  # noqa: SLF001
                formula = source_sheet._formulas.get(source)  # noqa: SLF001 - test double
                if formula is None:
                    self._formulas.pop(target, None)
                else:
                    self._formulas[target] = formula
        self._grid = {key: value for key, value in self._grid.items() if value is not None}

    def shift_cells(
        self, top: int, left: int, height: int, width: int, *, shift: int | None, inserting: bool
    ) -> None:
        if shift is None:
            raise AssertionError("the fake never receives an omitted shift")
        # -4161 / -4159 -> horizontal; -4121 / -4162 -> vertical.
        horizontal = shift in (-4161, -4159)
        if inserting:
            if horizontal:
                for r in range(top, top + height):
                    for c in range(_MAX_COL, left + width - 1, -1):
                        self._move(r, c - width, r, c)
                    for c in range(left, left + width):
                        self._clear(r, c)
            else:
                for c in range(left, left + width):
                    for r in range(_MAX_ROW, top + height - 1, -1):
                        self._move(r - height, c, r, c)
                    for r in range(top, top + height):
                        self._clear(r, c)
            return
        if horizontal:
            for r in range(top, top + height):
                for c in range(left, _MAX_COL - width + 1):
                    self._move(r, c + width, r, c)
                for c in range(_MAX_COL - width + 1, _MAX_COL + 1):
                    self._clear(r, c)
        else:
            for c in range(left, left + width):
                for r in range(top, _MAX_ROW - height + 1):
                    self._move(r + height, c, r, c)
                for r in range(_MAX_ROW - height + 1, _MAX_ROW + 1):
                    self._clear(r, c)

    def _move(self, from_row: int, from_col: int, to_row: int, to_col: int) -> None:
        value = self._grid.pop((from_row, from_col), None)
        if value is None:
            self._grid.pop((to_row, to_col), None)
        else:
            self._grid[(to_row, to_col)] = value
        fmt = self._formats.pop((from_row, from_col), None)
        if fmt is None:
            self._formats.pop((to_row, to_col), None)
        else:
            self._formats[(to_row, to_col)] = fmt
        formula = self._formulas.pop((from_row, from_col), None)
        if formula is None:
            self._formulas.pop((to_row, to_col), None)
        else:
            self._formulas[(to_row, to_col)] = formula

    def _clear(self, row: int, col: int) -> None:
        self._grid.pop((row, col), None)
        self._formats.pop((row, col), None)
        self._formulas.pop((row, col), None)

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

    @property
    def UsedRange(self) -> FakeRange:  # noqa: N802 - COM casing
        rows = [cell[0] for cell in self._grid if self._grid[cell] is not None]
        cols = [cell[1] for cell in self._grid if self._grid[cell] is not None]
        for row, col in self._merged:
            rows.append(row)
            cols.append(col)
        if not rows or not cols:
            return FakeRange(self, 1, 1, 1, 1)
        top, left = min(rows), min(cols)
        return FakeRange(self, top, left, max(rows) - top + 1, max(cols) - left + 1)


class FakeCells:
    def __init__(self, sheet: FakeSheet) -> None:
        self._sheet = sheet

    def __call__(self, row: int, col: int) -> FakeCell:
        return FakeCell(self._sheet, row, col)


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
        wb = FakeWorkbook([FakeSheet("Data"), FakeSheet("Other")])
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
def _seed_block(sheet: FakeSheet, top: int, left: int, values: list[list[Any]]) -> None:
    for r, row in enumerate(values):
        for c, value in enumerate(row):
            sheet.set(top + r, left + c, value)


def _seed_table(sheet: FakeSheet) -> FakeListObject:
    """Add a ListObject over T1:V3 (header + 2 data rows)."""

    table = FakeListObject(sheet, "tblData", 1, 20, 3, 3)
    sheet.ListObjects.Add(table)
    return table


def _merged_sheet() -> FakeWorkbook:
    """A workbook whose ``Data`` sheet holds a block, a merged area and a table."""

    data = FakeSheet("Data")
    _seed_block(data, 2, 1, [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]])
    data.merge_block(1, 10, 2, 3, across=False)
    data.set(1, 10, "merged title")
    _seed_table(data)
    other = FakeSheet("Other")
    return FakeWorkbook([data, other])


# ---------------------------------------------------------------------------
# Adapter — explicit destination / shift validation (no mutation)
# ---------------------------------------------------------------------------
def test_copy_requires_explicit_destination() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "Data", "")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    # No mutation: the destination block is untouched.
    assert wb.sheet("Data").get(2, 4) is None


def test_copy_rejects_blank_destination_sheet() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "  ", "D2")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_copy_rejects_overlapping_source_and_destination() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "Data", "B3")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert "overlap" in str(exc_info.value)


def test_copy_rejects_shape_mismatch() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "Data", "D2:E3")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert wb.sheet("Data").get(2, 4) is None


def test_insert_rejects_missing_shift() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    for shift in ("", "   ", None):  # type: ignore[list-item]
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            adapter.insert_range(wb, "Data", "B2", shift)
        assert exc_info.value.code is ErrorCode.INVALID_RANGE
    # No mutation: the target cell still holds its value.
    assert wb.sheet("Data").get(2, 2) == 2.0


def test_insert_rejects_unknown_shift() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.insert_range(wb, "Data", "B2", "sideways")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_delete_rejects_missing_shift() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.delete_range(wb, "Data", "B2", "")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert wb.sheet("Data").get(2, 2) == 2.0


def test_delete_rejects_insert_only_shift() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.delete_range(wb, "Data", "B2", "right")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Adapter — copy
# ---------------------------------------------------------------------------
def test_copy_to_anchor_cell() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    result = adapter.copy_range(wb, "Data", "A2:B4", "Data", "D2")

    assert result.action == "copy"
    assert result.verified is True
    assert result.range_after is not None
    assert result.range_after.address == "$D$2:$E$4"
    sheet = wb.sheet("Data")
    assert [[sheet.get(r, c) for c in (4, 5)] for r in (2, 3, 4)] == [
        [1.0, 2.0],
        [3.0, 4.0],
        [5.0, 6.0],
    ]


def test_copy_to_matching_block() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    result = adapter.copy_range(wb, "Data", "A2:B3", "Data", "G7:H8")

    assert result.verified is True
    sheet = wb.sheet("Data")
    assert [[sheet.get(r, c) for c in (7, 8)] for r in (7, 8)] == [[1.0, 2.0], [3.0, 4.0]]


def test_copy_cross_sheet() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    result = adapter.copy_range(wb, "Data", "A2:B4", "Other", "A1")

    assert result.verified is True
    assert result.source is not None and result.source.sheet == "Data"
    other = wb.sheet("Other")
    assert [[other.get(r, c) for c in (1, 2)] for r in (1, 2, 3)] == [
        [1.0, 2.0],
        [3.0, 4.0],
        [5.0, 6.0],
    ]


def test_copy_reports_postcondition_mismatch() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    original_copy = sheet.copy_block_from

    def tamper(*args: Any) -> None:
        original_copy(*args)
        sheet.set(2, 4, "tampered")

    sheet.copy_block_from = tamper  # type: ignore[method-assign]

    result = adapter.copy_range(wb, "Data", "A2:B4", "Data", "D2")
    assert result.verified is False
    assert result.notes and "mismatch" in result.notes[0].lower()


# ---------------------------------------------------------------------------
# Adapter — insert / delete
# ---------------------------------------------------------------------------
def test_insert_shift_right_moves_cells() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    result = adapter.insert_range(wb, "Data", "B2", "right")

    assert result.action == "insert"
    assert result.shift == "right"
    assert result.verified is True
    sheet = wb.sheet("Data")
    assert sheet.get(2, 2) is None  # the inserted cell is blank
    assert sheet.get(2, 3) == 2.0  # the old B2 value moved right


def test_insert_shift_down_moves_cells() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    result = adapter.insert_range(wb, "Data", "B2", "down")

    assert result.verified is True
    sheet = wb.sheet("Data")
    assert sheet.get(2, 2) is None
    assert sheet.get(3, 2) == 2.0


def test_insert_reports_postcondition_mismatch() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    original_shift = sheet.shift_cells

    def tamper(*args: Any, **kwargs: Any) -> None:
        original_shift(*args, **kwargs)
        sheet.set(2, 2, "tampered")

    sheet.shift_cells = tamper  # type: ignore[method-assign]

    result = adapter.insert_range(wb, "Data", "B2", "right")
    assert result.verified is False


def test_delete_shift_left_moves_cells() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    sheet.set(2, 3, "C2")  # the cell to the right of the deleted B2

    result = adapter.delete_range(wb, "Data", "B2", "left")

    assert result.action == "delete"
    assert result.shift == "left"
    assert result.verified is True
    assert sheet.get(2, 2) == "C2"
    assert sheet.get(2, 3) is None


def test_delete_shift_up_moves_cells() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    sheet.set(3, 2, "B3")

    result = adapter.delete_range(wb, "Data", "B2", "up")

    assert result.verified is True
    assert sheet.get(2, 2) == "B3"
    # The column shifted up: the old B4 value now sits in B3.
    assert sheet.get(3, 2) == 6.0


# ---------------------------------------------------------------------------
# Adapter — merge / unmerge
# ---------------------------------------------------------------------------
def test_merge_blocks_non_anchor_data() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    _seed_block(sheet, 6, 1, [["a", "b", "c"], ["d", "e", "f"]])

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.merge_range(wb, "Data", "A6:C7")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells_with_data"] == ["B6", "C6", "A7", "B7", "C7"]
    assert exc_info.value.context["anchor"] == "A6"
    # No merge happened: the cells are still unmerged and hold their values.
    assert sheet.get(7, 3) == "f"
    assert sheet.merge_anchor(6, 1) is None


def test_merge_allowed_with_allow_data_loss() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    _seed_block(sheet, 6, 1, [["a", "b", "c"], ["d", "e", "f"]])

    result = adapter.merge_range(wb, "Data", "A6:C7", allow_data_loss=True)

    assert result.verified is True
    assert result.affected_addresses == ["B6", "C6", "A7", "B7", "C7"]
    assert sheet.merge_anchor(6, 1) == (6, 1)
    assert sheet.get(6, 1) == "a"  # the anchor value survives
    assert sheet.get(6, 2) is None  # non-anchor values are discarded
    assert result.notes and "allow_data_loss" in result.notes[0]


def test_merge_across_keeps_one_anchor_per_row() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")
    _seed_block(sheet, 8, 1, [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]])

    result = adapter.merge_range(wb, "Data", "A8:C9", across=True, allow_data_loss=True)

    assert result.verified is True
    assert result.preconditions["across"] is True
    assert sheet.merge_anchor(8, 1) == (8, 1)
    assert sheet.merge_anchor(9, 1) == (9, 1)
    assert sheet.get(8, 1) == 1.0
    assert sheet.get(9, 1) == 4.0
    assert sheet.get(8, 2) is None


def test_merge_single_cell_range_is_a_no_op_merge() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    # A 1x1 range has no non-anchor data, so the data-loss gate does not fire.
    result = adapter.merge_range(wb, "Data", "B4", allow_data_loss=False)

    assert result.verified is True
    assert result.preconditions["cells_with_data"] == []


def test_unmerge_requires_a_merged_target() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.unmerge_range(wb, "Data", "A2:B4")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert "nothing to unmerge" in str(exc_info.value)


def test_unmerge_success() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")

    result = adapter.unmerge_range(wb, "Data", "J1:L2")

    assert result.verified is True
    assert result.affected_addresses == ["$J$1:$L$2"]
    assert sheet.merge_anchor(1, 10) is None
    assert sheet.get(1, 10) == "merged title"


def test_unmerge_partial_overlap_uses_merge_area() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()
    sheet = wb.sheet("Data")

    result = adapter.unmerge_range(wb, "Data", "K1")

    assert result.verified is True
    assert result.affected_addresses == ["$J$1:$L$2"]
    assert sheet.merge_anchor(1, 11) is None


# ---------------------------------------------------------------------------
# Adapter — ListObject hard guard
# ---------------------------------------------------------------------------
def test_copy_destination_intersecting_table_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "Data", "T1")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblData"
    assert exc_info.value.context["table_range"] == "$T$1:$V$3"
    assert exc_info.value.context["intersection_address"] == "T1"


def test_copy_source_intersecting_table_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "T1:V3", "Data", "A20")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblData"


def test_insert_intersecting_table_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.insert_range(wb, "Data", "U2", "right")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["operation"] == "insert"


def test_delete_intersecting_table_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.delete_range(wb, "Data", "U2:V3", "up")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblData"


def test_merge_intersecting_table_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.merge_range(wb, "Data", "T1:V3", allow_data_loss=True)
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION


# ---------------------------------------------------------------------------
# Adapter — merged cells and protected sheets
# ---------------------------------------------------------------------------
def test_insert_over_merged_cells_requires_allow_merged() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.insert_range(wb, "Data", "J1:L2", "right")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["merged_areas"] == ["$J$1:$L$2"]

    result = adapter.insert_range(wb, "Data", "J1:L2", "right", allow_merged=True)
    assert result.verified is True


def test_delete_over_merged_cells_requires_allow_merged() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.delete_range(wb, "Data", "J1:L2", "up")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_copy_onto_merged_destination_rejected() -> None:
    wb = _merged_sheet()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.copy_range(wb, "Data", "A2:B4", "Data", "J1")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["merged_areas"] == ["$J$1:$L$2"]


def test_protected_sheet_rejected_for_structural_ops() -> None:
    wb = _merged_sheet()
    wb.sheet("Data").ProtectContents = True
    adapter = RangesAdapter()

    for call in (
        lambda: adapter.insert_range(wb, "Data", "B2", "right"),
        lambda: adapter.delete_range(wb, "Data", "B2", "left"),
        lambda: adapter.merge_range(wb, "Data", "B4", allow_data_loss=True),
        lambda: adapter.unmerge_range(wb, "Data", "J1:L2"),
        lambda: adapter.copy_range(wb, "Data", "A2:B4", "Data", "D2"),
    ):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            call()
        assert exc_info.value.code is ErrorCode.PROTECTED_TARGET


# ---------------------------------------------------------------------------
# Service — guard chain
# ---------------------------------------------------------------------------
async def _open_seeded(rig: Rig, *, read_only: bool = False) -> str:
    """Open a working copy and seed a small block on its ``Data`` sheet."""

    info = await rig.open(read_only=read_only)
    sheet = rig.workbook().sheet("Data")
    _seed_block(sheet, 1, 1, [["x", 2.0], ["y", 4.0]])
    sheet.merge_block(9, 1, 1, 2, across=False)
    sheet.set(9, 1, "unmerge me")
    return info.workbook_id


async def test_service_copy_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig, read_only=True)
    rig.add_snapshot(wid)
    rig.grant("range_copy")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.copy_range(wid, "Data", "A1:B2", "Data", "D1")
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY
    assert [s for s in rig.worker.submits if s["name"] == "range_copy"] == []


async def test_service_copy_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.grant("range_copy")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.copy_range(wid, "Data", "A1:B2", "Data", "D1")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "range_copy"] == []


async def test_service_copy_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.service.copy_range(wid, "Data", "A1:B2", "Data", "D1")
    assert [s for s in rig.worker.submits if s["name"] == "range_copy"] == []


async def test_service_copy_success_and_audit(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_copy")

    result = await rig.service.copy_range(wid, "Data", "A1:B2", "Data", "D1")

    assert result.verified is True
    assert result.destination is not None and result.destination.address == "$D$1"
    mutate = next(s for s in rig.worker.submits if s["name"] == "range_copy")
    assert mutate["op_class"] == OperationClass.MUTATING
    assert rig.workbook().sheet("Data").get(1, 4) == "x"

    rig.audit.flush()
    events = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    copy_events = [event for event in events if event["tool"] == "range_copy"]
    assert copy_events, "the structural mutation must be audited"
    assert copy_events[-1]["operation_class"] == "mutating"
    assert copy_events[-1]["result"]["verified"] is True


async def test_service_insert_requires_explicit_shift_before_com(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_insert")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_range(wid, "Data", "A1", "")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert [s for s in rig.worker.submits if s["name"] == "range_insert"] == []


async def test_service_insert_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_insert")

    result = await rig.service.insert_range(wid, "Data", "B1", "right")

    assert result.verified is True
    assert result.shift == "right"
    assert rig.workbook().sheet("Data").get(1, 3) == 2.0


async def test_service_delete_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_delete")

    result = await rig.service.delete_range(wid, "Data", "B1", "left")

    assert result.verified is True
    assert result.shift == "left"
    assert rig.workbook().sheet("Data").get(1, 2) is None


async def test_service_merge_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)

    with pytest.raises(ExcelVbaMcpError):
        await rig.service.merge_range(wid, "Data", "A1:B2", allow_data_loss=True)
    assert [s for s in rig.worker.submits if s["name"] == "range_merge"] == []


async def test_service_merge_data_loss_gate(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_merge")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.merge_range(wid, "Data", "A1:B2")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells_with_data"] == ["B1", "A2", "B2"]
    # The workbook is unchanged: the merge did not happen.
    sheet = rig.workbook().sheet("Data")
    assert sheet.get(2, 2) == 4.0
    assert sheet.merge_anchor(1, 1) is None

    result = await rig.service.merge_range(wid, "Data", "A1:B2", allow_data_loss=True)
    assert result.verified is True
    assert result.affected_addresses == ["B1", "A2", "B2"]
    assert sheet.merge_anchor(1, 1) == (1, 1)


async def test_service_unmerge_requires_merged_target(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_unmerge")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.unmerge_range(wid, "Data", "A1:B2")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_service_unmerge_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    wid = await _open_seeded(rig)
    rig.add_snapshot(wid)
    rig.grant("range_unmerge")

    result = await rig.service.unmerge_range(wid, "Data", "A9:B9")

    assert result.verified is True
    assert result.affected_addresses == ["$A$9:$B$9"]
    assert rig.workbook().sheet("Data").merge_anchor(9, 1) is None


# ---------------------------------------------------------------------------
# API tools — thin async surface
# ---------------------------------------------------------------------------
def test_tool_insert_requires_shift_argument() -> None:
    """The shift parameter has no default: an omitted shift is a contract error."""

    signature = inspect.signature(range_tools.range_insert)
    assert signature.parameters["shift"].default is inspect.Parameter.empty

    delete_signature = inspect.signature(range_tools.range_delete)
    assert delete_signature.parameters["shift"].default is inspect.Parameter.empty


async def test_tool_insert_rejects_blank_shift_before_service_lookup() -> None:
    range_tools.set_ranges_service(None)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await range_tools.range_insert("wb", "Data", "A1", "")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_tool_delete_rejects_unknown_shift() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await range_tools.range_delete("wb", "Data", "A1", "down")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_tool_copy_rejects_blank_destination() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await range_tools.range_copy("wb", "Data", "A1", "Data", " ")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_tool_merge_rejects_non_boolean_flags() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await range_tools.range_merge("wb", "Data", "A1", allow_data_loss="yes")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_tool_copy_delegates_to_service() -> None:
    from excel_vba_mcp.domain.models import StructuralMutationResult

    seen: dict[str, Any] = {}

    class _Stub:
        async def copy_range(self, *args: Any) -> StructuralMutationResult:
            seen["args"] = args
            return StructuralMutationResult(action="copy", verified=True)

    range_tools.set_ranges_service(_Stub())  # type: ignore[arg-type]
    try:
        result = await range_tools.range_copy("wb", "Data", "A1:B2", "Data", "D1")
    finally:
        range_tools.set_ranges_service(None)

    assert result.verified is True
    assert seen["args"] == ("wb", "Data", "A1:B2", "Data", "D1")
