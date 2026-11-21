"""Unit tests for the bounded Range formatting surface (task_044).

These tests exercise the production
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and
:class:`~excel_vba_mcp.adapters.excel.ranges_adapter.RangesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* tri-state read — a uniform range returns concrete values, a mixed range
  reports ``None`` + the property name in ``mixed_properties`` (never
  flattened), and ``unsupported`` is always reported;
* partial update — only the fields the caller set are written, absent fields
  are left untouched;
* validation before mutation — a bad enum / color / empty string / null on a
  non-clearable field is rejected before any COM mutation;
* clearing — a present ``None`` clears ``number_format`` / ``font_color`` /
  ``fill_color`` to automatic where Excel allows;
* postcondition — a tampered re-read yields ``verified=False`` with a note;
* gates — a read-only working copy (``WORKBOOK_READ_ONLY``), a missing snapshot
  (``SNAPSHOT_FAILED``) and a missing approval grant (deny-by-default);
* the service submits ``READ_ONLY`` for ``get_format`` and ``MUTATING`` for
  ``set_format`` and resolves ``workbook_id``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.ranges_adapter import (
    RangesAdapter,
    validate_format_patch,
)
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ApprovalDeniedError, ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import FormatInfo, RangeFormatPatch
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

# Excel color-index sentinels used by the fake (mirror the adapter constants).
_XL_COLOR_INDEX_AUTOMATIC = -4105
_XL_COLOR_INDEX_NONE = -4142

# Excel enum ints mirrored from the adapter for building expectations.
_XL_HALIGN_GENERAL = 1
_XL_VALIGN_BOTTOM = -4107
_XL_PATTERN_NONE = -4142
_XL_PATTERN_SOLID = 1

_DEFAULT_FONT: dict[str, Any] = {
    "name": "Calibri",
    "bold": False,
    "italic": False,
    "size": 11.0,
    "color": 0,
    "color_index": _XL_COLOR_INDEX_AUTOMATIC,
}
_DEFAULT_FILL: dict[str, Any] = {
    "color": None,
    "color_index": _XL_COLOR_INDEX_NONE,
    "pattern": _XL_PATTERN_NONE,
}
_DEFAULT_ALIGN: dict[str, Any] = {
    "horizontal": _XL_HALIGN_GENERAL,
    "vertical": _XL_VALIGN_BOTTOM,
    "wrap_text": False,
}


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph — a worksheet owns a value/formula grid plus per-cell
# number format, font, fill and alignment style. A FakeRange spans a block and
# computes the "common" value across its cells (Excel returns Null when they
# disagree), while a setter applies the value to every spanned cell.
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


def _common(values: list[Any]) -> Any:
    """Return the shared value, or ``None`` when the cells disagree (Excel Null)."""

    if not values:
        return None
    first = values[0]
    for value in values[1:]:
        if value != first:
            return None
    return first


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


class FakeFont:
    """A ``Range.Font`` proxy computing/setting across the spanned cells."""

    def __init__(self, sheet: FakeSheet, top: int, left: int, height: int, width: int) -> None:
        self._sheet = sheet
        self._top = top
        self._left = left
        self._height = height
        self._width = width

    def _cells(self) -> list[tuple[int, int]]:
        return [
            (r, c)
            for r in range(self._top, self._top + self._height)
            for c in range(self._left, self._left + self._width)
        ]

    def _get(self, key: str) -> Any:
        return _common([self._sheet.get_font(r, c, key) for r, c in self._cells()])

    def _set(self, key: str, value: Any) -> None:
        for r, c in self._cells():
            self._sheet.set_font(r, c, key, value)

    @property
    def Name(self) -> Any:  # noqa: N802 - COM casing
        return self._get("name")

    @Name.setter
    def Name(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("name", value)

    @property
    def Bold(self) -> Any:  # noqa: N802 - COM casing
        return self._get("bold")

    @Bold.setter
    def Bold(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("bold", value)

    @property
    def Italic(self) -> Any:  # noqa: N802 - COM casing
        return self._get("italic")

    @Italic.setter
    def Italic(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("italic", value)

    @property
    def Size(self) -> Any:  # noqa: N802 - COM casing
        return self._get("size")

    @Size.setter
    def Size(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("size", value)

    @property
    def Color(self) -> Any:  # noqa: N802 - COM casing
        return self._get("color")

    @Color.setter
    def Color(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("color", value)
        self._set("color_index", 0)

    @property
    def ColorIndex(self) -> Any:  # noqa: N802 - COM casing
        return self._get("color_index")

    @ColorIndex.setter
    def ColorIndex(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("color_index", value)


class FakeInterior:
    """A ``Range.Interior`` proxy computing/setting fill across spanned cells."""

    def __init__(self, sheet: FakeSheet, top: int, left: int, height: int, width: int) -> None:
        self._sheet = sheet
        self._top = top
        self._left = left
        self._height = height
        self._width = width

    def _cells(self) -> list[tuple[int, int]]:
        return [
            (r, c)
            for r in range(self._top, self._top + self._height)
            for c in range(self._left, self._left + self._width)
        ]

    def _get(self, key: str) -> Any:
        return _common([self._sheet.get_fill(r, c, key) for r, c in self._cells()])

    def _set(self, key: str, value: Any) -> None:
        for r, c in self._cells():
            self._sheet.set_fill(r, c, key, value)

    @property
    def Color(self) -> Any:  # noqa: N802 - COM casing
        return self._get("color")

    @Color.setter
    def Color(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("color", value)
        self._set("color_index", 0)
        self._set("pattern", _XL_PATTERN_SOLID)

    @property
    def ColorIndex(self) -> Any:  # noqa: N802 - COM casing
        return self._get("color_index")

    @ColorIndex.setter
    def ColorIndex(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("color_index", value)
        if value == _XL_COLOR_INDEX_NONE:
            self._set("color", None)
            self._set("pattern", _XL_PATTERN_NONE)

    @property
    def Pattern(self) -> Any:  # noqa: N802 - COM casing
        return self._get("pattern")

    @Pattern.setter
    def Pattern(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._set("pattern", value)


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
    def Address(self) -> str:  # noqa: N802 - COM casing
        return (
            f"${_col_letter(self.Left)}${self.Top}:"
            f"${_col_letter(self.Left + self.Width - 1)}${self.Top + self.Height - 1}"
        )

    @property
    def Areas(self) -> FakeAreas:  # noqa: N802 - COM casing
        return FakeAreas(self._areas)

    def _cells(self) -> list[tuple[int, int]]:
        return [
            (r, c)
            for r in range(self.Top, self.Top + self.Height)
            for c in range(self.Left, self.Left + self.Width)
        ]

    @property
    def Font(self) -> FakeFont:  # noqa: N802 - COM casing
        return FakeFont(self._sheet, self.Top, self.Left, self.Height, self.Width)

    @property
    def Interior(self) -> FakeInterior:  # noqa: N802 - COM casing
        return FakeInterior(self._sheet, self.Top, self.Left, self.Height, self.Width)

    @property
    def NumberFormat(self) -> Any:  # noqa: N802 - COM casing
        return _common([self._sheet.get_format(r, c) for r, c in self._cells()])

    @NumberFormat.setter
    def NumberFormat(self, value: Any) -> None:  # noqa: N802 - COM casing
        for r, c in self._cells():
            self._sheet.set_format(r, c, value)

    @property
    def HorizontalAlignment(self) -> Any:  # noqa: N802 - COM casing
        return _common([self._sheet.get_align(r, c, "horizontal") for r, c in self._cells()])

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value: Any) -> None:  # noqa: N802 - COM casing
        for r, c in self._cells():
            self._sheet.set_align(r, c, "horizontal", value)

    @property
    def VerticalAlignment(self) -> Any:  # noqa: N802 - COM casing
        return _common([self._sheet.get_align(r, c, "vertical") for r, c in self._cells()])

    @VerticalAlignment.setter
    def VerticalAlignment(self, value: Any) -> None:  # noqa: N802 - COM casing
        for r, c in self._cells():
            self._sheet.set_align(r, c, "vertical", value)

    @property
    def WrapText(self) -> Any:  # noqa: N802 - COM casing
        return _common([self._sheet.get_align(r, c, "wrap_text") for r, c in self._cells()])

    @WrapText.setter
    def WrapText(self, value: Any) -> None:  # noqa: N802 - COM casing
        for r, c in self._cells():
            self._sheet.set_align(r, c, "wrap_text", value)


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
        self._font: dict[tuple[int, int], dict[str, Any]] = {}
        self._fill: dict[tuple[int, int], dict[str, Any]] = {}
        self._align: dict[tuple[int, int], dict[str, Any]] = {}
        self.Cells = FakeCells(self)

    # Values -------------------------------------------------------------
    def get(self, row: int, col: int) -> Any:
        return self._grid.get((row, col))

    def set(self, row: int, col: int, value: Any) -> None:
        self._grid[(row, col)] = value

    # Number format ------------------------------------------------------
    def get_format(self, row: int, col: int) -> Any:
        return self._formats.get((row, col), "General")

    def set_format(self, row: int, col: int, value: Any) -> None:
        self._formats[(row, col)] = value

    # Font ---------------------------------------------------------------
    def get_font(self, row: int, col: int, key: str) -> Any:
        return {**_DEFAULT_FONT, **self._font.get((row, col), {})}[key]

    def set_font(self, row: int, col: int, key: str, value: Any) -> None:
        style = {**_DEFAULT_FONT, **self._font.get((row, col), {})}
        style[key] = value
        self._font[(row, col)] = style

    # Fill ---------------------------------------------------------------
    def get_fill(self, row: int, col: int, key: str) -> Any:
        return {**_DEFAULT_FILL, **self._fill.get((row, col), {})}[key]

    def set_fill(self, row: int, col: int, key: str, value: Any) -> None:
        style = {**_DEFAULT_FILL, **self._fill.get((row, col), {})}
        style[key] = value
        self._fill[(row, col)] = style

    # Alignment ----------------------------------------------------------
    def get_align(self, row: int, col: int, key: str) -> Any:
        return {**_DEFAULT_ALIGN, **self._align.get((row, col), {})}[key]

    def set_align(self, row: int, col: int, key: str, value: Any) -> None:
        style = {**_DEFAULT_ALIGN, **self._align.get((row, col), {})}
        style[key] = value
        self._align[(row, col)] = style

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
# Adapter-level: tri-state read
# ---------------------------------------------------------------------------
def test_get_format_uniform_single_cell() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_format(1, 1, "0.00")
    adapter = RangesAdapter()

    info = adapter.get_format(wb, "Data", "A1")

    assert isinstance(info, FormatInfo)
    assert info.number_format == "0.00"
    assert info.font.name == "Calibri"
    assert info.font.bold is False
    assert info.font.italic is False
    assert info.font.size == 11.0
    assert info.font.color == 0
    assert info.fill.color is None  # no fill
    assert info.fill.pattern == "xlNone"
    assert info.alignment.horizontal == "xlGeneral"
    assert info.alignment.vertical == "xlBottom"
    assert info.alignment.wrap_text is False
    assert info.mixed_properties == []
    # Unsupported aspects are always reported, never silently dropped.
    assert "borders" in info.unsupported
    assert "number_format_local" in info.unsupported


def test_get_format_uniform_block_not_mixed() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    for r in range(1, 3):
        for c in range(1, 3):
            sheet.set_font(r, c, "bold", True)
            sheet.set_format(r, c, "0.00")
    adapter = RangesAdapter()

    info = adapter.get_format(wb, "Data", "A1:B2")

    assert info.font.bold is True
    assert info.number_format == "0.00"
    assert info.mixed_properties == []


def test_get_format_mixed_reports_tristate() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    # Two cells that disagree on number_format, bold and fill color.
    sheet.set_format(1, 1, "0.00")
    sheet.set_format(2, 1, "General")
    sheet.set_font(1, 1, "bold", True)
    sheet.set_font(2, 1, "bold", False)
    sheet.set_fill(1, 1, "color", 255)
    sheet.set_fill(1, 1, "color_index", 0)
    # cell (2,1) stays no-fill.
    adapter = RangesAdapter()

    info = adapter.get_format(wb, "Data", "A1:A2")

    assert info.number_format is None
    assert info.font.bold is None
    assert info.fill.color is None
    assert "number_format" in info.mixed_properties
    assert "font.bold" in info.mixed_properties
    assert "fill.color" in info.mixed_properties
    # Properties that agree are still reported concretely.
    assert info.font.name == "Calibri"
    assert info.alignment.horizontal == "xlGeneral"


# ---------------------------------------------------------------------------
# Adapter-level: partial write
# ---------------------------------------------------------------------------
def test_set_format_partial_leaves_other_props_untouched() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_font(1, 1, "italic", True)
    sheet.set_font(1, 1, "size", 14.0)
    adapter = RangesAdapter()

    result = adapter.set_format(wb, "Data", "A1", RangeFormatPatch(font_bold=True))

    assert result.verified is True
    assert result.applied == ["font_bold"]
    # Only bold was written; italic and size are unchanged.
    assert sheet.get_font(1, 1, "bold") is True
    assert sheet.get_font(1, 1, "italic") is True
    assert sheet.get_font(1, 1, "size") == 14.0


def test_set_format_number_fill_alignment() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    patch = RangeFormatPatch(
        number_format="0.000",
        font_bold=True,
        font_color=255,
        fill_color=65280,
        alignment_horizontal="xlCenter",
        alignment_vertical="xlTop",
        alignment_wrap_text=True,
    )
    result = adapter.set_format(wb, "Data", "A1", patch)

    assert result.verified is True
    info = result.result_format
    assert info is not None
    assert info.number_format == "0.000"
    assert info.font.bold is True
    assert info.font.color == 255
    assert info.fill.color == 65280
    assert info.alignment.horizontal == "xlCenter"
    assert info.alignment.vertical == "xlTop"
    assert info.alignment.wrap_text is True


def test_set_format_applied_lists_only_present_fields() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    patch = RangeFormatPatch(font_bold=True, fill_color=10)
    result = adapter.set_format(wb, "Data", "A1", patch)

    assert result.applied == ["fill_color", "font_bold"]
    assert result.ignored == []


# ---------------------------------------------------------------------------
# Adapter-level: clearing to automatic
# ---------------------------------------------------------------------------
def test_set_format_clear_number_format_to_general() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_format(1, 1, "0.00")
    adapter = RangesAdapter()

    result = adapter.set_format(wb, "Data", "A1", RangeFormatPatch(number_format=None))

    assert result.verified is True
    assert sheet.get_format(1, 1) == "General"


def test_set_format_clear_colors_to_automatic() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    sheet.set_font(1, 1, "color", 255)
    sheet.set_font(1, 1, "color_index", 0)
    sheet.set_fill(1, 1, "color", 4660)
    sheet.set_fill(1, 1, "color_index", 0)
    adapter = RangesAdapter()

    result = adapter.set_format(
        wb, "Data", "A1", RangeFormatPatch(font_color=None, fill_color=None)
    )

    assert result.verified is True
    assert sheet.get_font(1, 1, "color_index") == _XL_COLOR_INDEX_AUTOMATIC
    assert sheet.get_fill(1, 1, "color_index") == _XL_COLOR_INDEX_NONE


# ---------------------------------------------------------------------------
# Adapter-level: validation before mutation
# ---------------------------------------------------------------------------
def test_set_format_rejects_invalid_alignment_enum() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_format(wb, "Data", "A1", RangeFormatPatch(alignment_horizontal="xlBogus"))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    # No mutation: the alignment is untouched.
    assert sheet.get_align(1, 1, "horizontal") == _XL_HALIGN_GENERAL


def test_set_format_rejects_out_of_range_color() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_format(wb, "Data", "A1", RangeFormatPatch(font_color=0x1000000))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_set_format_rejects_empty_number_format() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_format(wb, "Data", "A1", RangeFormatPatch(number_format="   "))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_set_format_rejects_null_on_non_clearable_field() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_format(wb, "Data", "A1", RangeFormatPatch(font_bold=None))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_set_format_rejects_bad_font_size() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_format(wb, "Data", "A1", RangeFormatPatch(font_size=0))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Adapter-level: postcondition mismatch
# ---------------------------------------------------------------------------
def test_set_format_reports_postcondition_mismatch() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    sheet = wb.sheet("Data")
    adapter = RangesAdapter()
    original_set_font = sheet.set_font

    def tamper(row: int, col: int, key: str, value: Any) -> None:
        original_set_font(row, col, key, value)
        if key == "bold":
            # Simulate Excel not honouring the write.
            original_set_font(row, col, key, not value)

    sheet.set_font = tamper  # type: ignore[method-assign]

    result = adapter.set_format(wb, "Data", "A1", RangeFormatPatch(font_bold=True))

    assert result.verified is False
    assert result.notes and "mismatch" in result.notes[0].lower()


# ---------------------------------------------------------------------------
# validate_format_patch — direct unit coverage
# ---------------------------------------------------------------------------
def test_validate_format_patch_accepts_empty_patch() -> None:
    # No fields set -> nothing to validate.
    validate_format_patch(RangeFormatPatch())


def test_validate_format_patch_accepts_supported_enums() -> None:
    validate_format_patch(
        RangeFormatPatch(
            fill_pattern="xlSolid",
            alignment_horizontal="xlLeft",
            alignment_vertical="xlCenter",
        )
    )


def test_validate_format_patch_rejects_unknown_pattern() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_format_patch(RangeFormatPatch(fill_pattern="xlDiagonal"))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_format_patch_rejects_overlong_name() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_format_patch(RangeFormatPatch(font_name="x" * 300))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Service-level: read path
# ---------------------------------------------------------------------------
async def test_service_get_format(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.workbook().sheet("Data").set_format(1, 1, "0.00")

    result = await rig.service.get_format(info.workbook_id, "Data", "A1")
    assert result.number_format == "0.00"

    submit = next(s for s in rig.worker.submits if s["name"] == "range_get_format")
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_get_format_unknown_workbook(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.get_format("missing", "Data", "A1")


# ---------------------------------------------------------------------------
# Service-level: guard chain for set_format
# ---------------------------------------------------------------------------
async def test_service_set_format_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_set_format")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_format(
            info.workbook_id, "Data", "A1", RangeFormatPatch(font_bold=True)
        )
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY
    assert [s for s in rig.worker.submits if s["name"] == "range_set_format"] == []


async def test_service_set_format_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("range_set_format")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_format(
            info.workbook_id, "Data", "A1", RangeFormatPatch(font_bold=True)
        )
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "range_set_format"] == []


async def test_service_set_format_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    # No grant.

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.set_format(
            info.workbook_id, "Data", "A1", RangeFormatPatch(font_bold=True)
        )
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "range_set_format"] == []


async def test_service_set_format_validates_before_gates(tmp_path: Path) -> None:
    """A malformed patch is rejected (INVALID_RANGE) even without a snapshot."""

    rig = Rig(tmp_path)
    info = await rig.open()
    # No snapshot, no approval — but the invalid enum must still fail first.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_format(
            info.workbook_id, "Data", "A1", RangeFormatPatch(alignment_horizontal="xlNope")
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert [s for s in rig.worker.submits if s["name"] == "range_set_format"] == []


async def test_service_set_format_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_set_format")

    result = await rig.service.set_format(
        info.workbook_id, "Data", "A1", RangeFormatPatch(font_bold=True, fill_color=255)
    )
    assert result.verified is True
    assert result.applied == ["fill_color", "font_bold"]

    mutate = next(s for s in rig.worker.submits if s["name"] == "range_set_format")
    assert mutate["op_class"] == OperationClass.MUTATING
    sheet = rig.workbook().sheet("Data")
    assert sheet.get_font(1, 1, "bold") is True
    assert sheet.get_fill(1, 1, "color") == 255
