"""STA-local Excel Range/Cells adapter (task_042, spec §14 Phase 2 — Range/Cells).

:class:`RangesAdapter` is the single place where the Excel ``Worksheet.Range`` /
``Range.Value2`` / ``Range.Formula2`` object model is touched for the read-only
Phase-2 surface (``range_get``, ``range_read_values``, ``range_read_formulas``).
It is a **pure STA-local helper**: every method is synchronous, receives the live
``Workbook`` COM proxy (``Any``) and returns only Pydantic DTOs / primitives, so
results cross the single-STA worker boundary safely
(:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never imports
``pywin32`` at module import time and never calls ``.Activate`` / ``.Select`` on
any object (addendum BG03/BG04/BG05).

Read semantics (task_038 decision record, ``docs/research/ranges_cells_api.md``):

* **Address.** An explicit ``sheet`` + A1-style ``address``. Whole-row (``1:3``)
  and whole-column (``A:C``) targets are supported **only within the configured
  bounds**; an invalid, ambiguous (multi-area) or oversized target is rejected
  with a typed :class:`~excel_vba_mcp.domain.errors.InvalidRangeError`.
* **Values.** ``Range.Value2`` is the raw channel (scalar for one cell, 2-D for
  any multi-cell block). Dates are typed **only** from ``Range.NumberFormat``
  (never guessed from a bare serial). Classification reuses the Stage-1
  :func:`~excel_vba_mcp.adapters.excel.cell_values.classify_cell_value`.
* **Formulas.** ``Range.Formula2`` is the primary channel on build 14332;
  ``Range.Formula`` is the fallback. ``_xlfn`` / ``_xlws`` / ``_xlpm`` prefixes
  are passed through verbatim and flagged in ``notes``.

Bounds (task_038 §3.2/§3.3): ``MAX_RANGE_CELLS`` / ``MAX_RESPONSE_CELLS`` are
enforced **before** materializing a response; an oversized request is rejected
and an oversized response is truncated (``truncated=True``). The whole
``UsedRange`` is never read unless it is explicitly requested within limits.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from excel_vba_mcp.adapters.excel.cell_values import classify_cell_value
from excel_vba_mcp.domain.errors import InvalidRangeError
from excel_vba_mcp.domain.models import (
    MAX_RANGE_CELLS,
    MAX_RESPONSE_CELLS,
    CellValue,
    FormulaCell,
    RangeAddress,
    RangeFormulas,
    RangeValues,
)
from excel_vba_mcp.domain.statuses import CellValueKind, RangeShape

#: Formula prefixes that mark a function newer than the reading build
#: (task_038 §2.1). Presence is surfaced in ``notes``; the string itself is
#: passed through verbatim and never stripped or rewritten.
_FORMULA_PREFIXES: tuple[str, ...] = (
    "_xlfn._xlws.",
    "_xlfn._xlpm.",
    "_xlfn.",
    "_xlws.",
    "_xlpm.",
)

#: ``Range.NumberFormat`` token detection for date/time formats (task_038 §1.4).
#: A numeric value is typed as a date only when its number format carries one of
#: these tokens; ``General`` / numeric formats never produce a date (no guessing).
_DATE_FORMAT_TOKENS: tuple[str, ...] = ("yy", "dd", "mmm")

#: Excel's 1900 date system epoch (``serial 0`` renders as 1899-12-30).
_EXCEL_EPOCH = datetime(1899, 12, 30)


# ---------------------------------------------------------------------------
# Safe COM readers (never raise; no proxy leaks)
# ---------------------------------------------------------------------------
def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively (never raises)."""

    if value is None:
        return default
    try:
        text = str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    return text if text else default


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively (never raises)."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


# ---------------------------------------------------------------------------
# Address parsing / normalization (no COM)
# ---------------------------------------------------------------------------
def _col_index(letters: str) -> int:
    """Return the 1-based column index for Excel column ``letters`` (e.g. ``A`` -> 1)."""

    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _col_letter(index: int) -> str:
    """Return the Excel column letter(s) for a 1-based column ``index``."""

    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _cell_address(row: int, col: int) -> str:
    """Return an A1-style address (no ``$``) for a 1-based ``(row, col)`` pair."""

    return f"{_col_letter(col)}{row}"


def _require_sheet(wb: Any, name: str) -> Any:
    """Resolve ``name`` to a worksheet proxy or raise a structured error."""

    if not isinstance(name, str) or not name.strip():
        raise InvalidRangeError(
            "Argument 'sheet' must be a non-empty string.", context={"sheet": name}
        )
    try:
        return wb.Worksheets.Item(name)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            f"Worksheet {name!r} was not found in the workbook.", context={"sheet": name}
        ) from exc


def _resolve_range(worksheet: Any, address: str, *, sheet_name: str) -> Any:
    """Resolve an A1 ``address`` to a single-area Range on ``worksheet``.

    Rejects multi-area and unparseable addresses. Cell-count bounds are enforced
    by the callers (:meth:`RangesAdapter.read_values` / ``read_formulas``) before
    any cell is read.
    """

    if not isinstance(address, str) or not address.strip():
        raise InvalidRangeError(
            "Argument 'address' must be a non-empty A1-style string.",
            context={"sheet": sheet_name, "address": address},
        )
    try:
        range_obj = worksheet.Range(address.strip())
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            f"Range {address!r} is not valid on sheet {sheet_name!r}.",
            context={"sheet": sheet_name, "address": address},
        ) from exc

    # Multi-area guard (task_038 §7): reject rather than guess an area.
    try:
        area_count = int(range_obj.Areas.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        area_count = 1
    if area_count > 1:
        raise InvalidRangeError(
            f"Range {address!r} spans {area_count} areas (ambiguous); a single area is required.",
            context={"sheet": sheet_name, "address": address, "areas": area_count},
        )
    return range_obj


def _describe_range(range_obj: Any, *, sheet: str) -> RangeAddress:
    """Build a :class:`RangeAddress` from a resolved Range."""

    rows = _safe_int(range_obj.Rows.Count, default=0)
    cols = _safe_int(range_obj.Columns.Count, default=0)
    address = _safe_str(range_obj.Address)
    return RangeAddress(
        sheet=sheet,
        address=address,
        rows=rows,
        cols=cols,
        shape=_shape_for(rows, cols),
        areas=1,
    )


def _shape_for(rows: int, cols: int) -> RangeShape:
    """Return the normalized :class:`RangeShape` tag for a dimensions pair."""

    if rows <= 1 and cols <= 1:
        return RangeShape.SCALAR
    if rows <= 1:
        return RangeShape.ROW
    if cols <= 1:
        return RangeShape.COLUMN
    return RangeShape.BLOCK


# ---------------------------------------------------------------------------
# Value / formula reading (STA-local)
# ---------------------------------------------------------------------------
def _normalize_grid(raw: Any) -> tuple[tuple[Any, ...], ...]:
    """Normalize a ``Range.Value2`` / ``Formula2`` read into a 2-D grid.

    Excel returns a scalar for a single-cell range and a 2-D array for any
    multi-cell range (task_038 §1.4); both are normalized to 2-D here.
    """

    if not isinstance(raw, tuple):
        return ((raw,),)
    if not raw:
        return ((),)
    if isinstance(raw[0], tuple):
        return tuple(tuple(row) for row in raw)
    return (tuple(raw),)


def _grid_cell(grid: tuple[tuple[Any, ...], ...], row: int, col: int) -> Any:
    """Return ``grid[row][col]`` defensively (``None`` when out of range)."""

    if row < len(grid):
        line = grid[row]
        if col < len(line):
            return line[col]
    return None


def _is_date_format(number_format: Any) -> bool:
    """Return ``True`` only when ``number_format`` is a date/time format.

    ``None`` (mixed) and ``General`` / numeric formats are never dates, so a
    bare numeric serial is never guessed into a date (task_038 §1.4).
    """

    if number_format is None:
        return False
    lowered = str(number_format).lower()
    if lowered in ("general", ""):
        return False
    return any(token in lowered for token in _DATE_FORMAT_TOKENS)


def _classify_cell(value: Any, *, address: str, number_format: Any) -> CellValue:
    """Classify a raw ``Value2`` value into a typed :class:`CellValue`.

    A number is promoted to :attr:`CellValueKind.DATE` (ISO-8601 string) **only**
    when its number format is a date/time format and the value is a plausible
    Excel serial; otherwise the Stage-1 ``classify_cell_value`` order applies
    (bool before int, OLE error codes -> error, etc.).
    """

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and _is_date_format(number_format)
    ):
        moment = _EXCEL_EPOCH + timedelta(days=float(value))
        return CellValue(kind=CellValueKind.DATE, value=moment.isoformat(), address=address)
    return classify_cell_value(value, address=address)


def _classify_grid(
    grid: tuple[tuple[Any, ...], ...],
    *,
    sub_range: Any,
    top_row: int,
    left_col: int,
) -> list[list[CellValue]]:
    """Classify a whole ``Value2`` grid using a single ``NumberFormat`` read.

    ``Range.NumberFormat`` is read **once** for the bounded sub-range. When it is
    ``Null`` (the cells disagree, task_038 §1.1) only then does the adapter fall
    back to per-cell ``NumberFormat`` reads for the returned cells, so a uniform
    date/number column costs one format read instead of one per cell.
    """

    rows = len(grid)
    try:
        range_format = sub_range.NumberFormat
    except Exception:  # noqa: BLE001 - defensive COM read
        range_format = None

    if range_format is not None:
        return [
            [
                _classify_cell(
                    _grid_cell(grid, r, c),
                    address=_cell_address(top_row + r, left_col + c),
                    number_format=range_format,
                )
                for c in range(len(grid[r]) if r < rows else 0)
            ]
            for r in range(rows)
        ]

    # Mixed formats: read per-cell only for the cells being returned (bounded).
    result: list[list[CellValue]] = []
    for r in range(rows):
        row_values: list[CellValue] = []
        for c in range(len(grid[r])):
            try:
                cell_format = sub_range.Cells(r + 1, c + 1).NumberFormat
            except Exception:  # noqa: BLE001 - defensive per-cell COM read
                cell_format = None
            row_values.append(
                _classify_cell(
                    _grid_cell(grid, r, c),
                    address=_cell_address(top_row + r, left_col + c),
                    number_format=cell_format,
                )
            )
        result.append(row_values)
    return result


def _sub_range(worksheet: Any, range_obj: Any, start_row: int, row_count: int, cols: int) -> Any:
    """Return the bounded sub-range for a window starting at ``start_row``."""

    top_row = _safe_int(range_obj.Row) + start_row
    left_col = _safe_int(range_obj.Column)
    return worksheet.Range(
        worksheet.Cells(top_row, left_col),
        worksheet.Cells(top_row + row_count - 1, left_col + cols - 1),
    )


def _read_formula2_grid(range_obj: Any) -> tuple[tuple[Any, ...], ...]:
    """Read ``range_obj.Formula2`` as a 2-D grid, falling back to ``Formula``.

    On the target build (14332) ``Formula2`` is available; a build without it
    (or an OLAP source) raises and the adapter transparently uses ``Formula``.
    """

    try:
        raw = range_obj.Formula2
    except Exception:  # noqa: BLE001 - Formula2 unavailable / OLAP -> fallback
        raw = range_obj.Formula
    return _normalize_grid(raw)


def _read_formula_grid(range_obj: Any) -> tuple[tuple[Any, ...], ...]:
    """Read the secondary ``range_obj.Formula`` (IIE) grid (empty on failure)."""

    try:
        raw = range_obj.Formula
    except Exception:  # noqa: BLE001 - Formula unavailable (e.g. OLAP)
        return ((),)
    return _normalize_grid(raw)


def _build_formula_cell(address: str, raw2: Any, raw1: Any) -> FormulaCell:
    """Assemble one :class:`FormulaCell` (formula2 primary, formula when different)."""

    formula2 = "" if raw2 is None else str(raw2)
    formula1 = "" if raw1 is None else str(raw1)
    formula_secondary = formula1 if formula1 != formula2 else None
    return FormulaCell(address=address, formula2=formula2, formula=formula_secondary)


def _flag_formula_prefix(cell: FormulaCell, notes: list[str]) -> None:
    """Append a note when a formula carries a newer-function prefix (verbatim)."""

    for prefix in _FORMULA_PREFIXES:
        if prefix in cell.formula2 or (cell.formula and prefix in cell.formula):
            notes.append(f"{cell.address}: formula uses the {prefix!r} prefix (newer function).")
            return


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class RangesAdapter:
    """STA-local Excel Range read/metadata operations (no COM at import)."""

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------
    def resolve_range(self, wb: Any, sheet: str, address: str) -> RangeAddress:
        """Resolve an explicit ``sheet`` + A1 ``address`` to :class:`RangeAddress`.

        Validates the sheet exists, the address parses, the target is a single
        area and within the configured bounds; rejects invalid/ambiguous/
        oversized targets with a typed :class:`InvalidRangeError`.
        """

        worksheet = _require_sheet(wb, sheet)
        range_obj = _resolve_range(worksheet, address, sheet_name=_safe_str(worksheet.Name))
        return _describe_range(range_obj, sheet=_safe_str(worksheet.Name))

    # ------------------------------------------------------------------
    # Values (read-only, typed, bounded)
    # ------------------------------------------------------------------
    def read_values(
        self,
        wb: Any,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeValues:
        """Read typed cell values from a bounded ``Range.Value2`` window.

        ``offset`` is a 0-based row index into the requested range; ``limit``
        caps the returned rows. The response is bounded by
        :data:`~excel_vba_mcp.domain.models.MAX_RESPONSE_CELLS` (an oversized
        window is truncated and ``truncated`` is set) and by
        :data:`~excel_vba_mcp.domain.models.MAX_RANGE_CELLS` (an oversized
        request is rejected before any cell is read).
        """

        _validate_window(offset, limit)
        worksheet, range_obj, resolved = self._open(wb, sheet, address)

        top_row = _safe_int(range_obj.Row)
        left_col = _safe_int(range_obj.Column)
        total_rows = resolved.rows
        total_cols = resolved.cols

        start, end, truncated = _window(total_rows, total_cols, offset, limit)
        row_count = end - start
        shape = _shape_for(total_rows, total_cols)

        if row_count <= 0 or total_cols <= 0:
            return RangeValues(
                range=resolved,
                shape=shape,
                rows=0,
                cols=total_cols,
                values=[],
                truncated=truncated or start >= total_rows,
            )

        sub_range = _sub_range(worksheet, range_obj, start, row_count, total_cols)
        grid = _normalize_grid(sub_range.Value2)
        values = _classify_grid(
            grid, sub_range=sub_range, top_row=top_row + start, left_col=left_col
        )

        return RangeValues(
            range=resolved,
            shape=shape,
            rows=row_count,
            cols=total_cols,
            values=values,
            truncated=truncated,
        )

    # ------------------------------------------------------------------
    # Formulas (read-only, 2-D normalized)
    # ------------------------------------------------------------------
    def read_formulas(
        self,
        wb: Any,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeFormulas:
        """Read typed formulas from a bounded window (``Formula2`` primary).

        ``Formula2`` is read first (Dynamic-Arrays form, build 14332); when the
        property is unavailable the adapter falls back to ``Formula``. A
        ``formula`` secondary field is included only when it differs from
        ``formula2`` (an ``@`` was added/removed by IIE). ``_xlfn`` / ``_xlws``
        / ``_xlpm`` prefixes are surfaced in ``notes``; the formula strings are
        never stripped or rewritten.
        """

        _validate_window(offset, limit)
        worksheet, range_obj, resolved = self._open(wb, sheet, address)

        top_row = _safe_int(range_obj.Row)
        left_col = _safe_int(range_obj.Column)
        total_rows = resolved.rows
        total_cols = resolved.cols

        start, end, truncated = _window(total_rows, total_cols, offset, limit)
        row_count = end - start
        shape = _shape_for(total_rows, total_cols)
        notes: list[str] = []

        if row_count <= 0 or total_cols <= 0:
            return RangeFormulas(
                range=resolved,
                shape=shape,
                rows=0,
                cols=total_cols,
                formulas=[],
                truncated=truncated or start >= total_rows,
                notes=notes,
            )

        sub_range = _sub_range(worksheet, range_obj, start, row_count, total_cols)
        formula2_grid = _read_formula2_grid(sub_range)
        formula_grid = _read_formula_grid(sub_range)

        formulas: list[list[FormulaCell]] = []
        for row_offset in range(row_count):
            excel_row = top_row + start + row_offset
            row_cells: list[FormulaCell] = []
            for col_offset in range(total_cols):
                cell_address = _cell_address(excel_row, left_col + col_offset)
                raw2 = _grid_cell(formula2_grid, row_offset, col_offset)
                raw1 = _grid_cell(formula_grid, row_offset, col_offset)
                cell = _build_formula_cell(cell_address, raw2, raw1)
                _flag_formula_prefix(cell, notes)
                row_cells.append(cell)
            formulas.append(row_cells)

        return RangeFormulas(
            range=resolved,
            shape=shape,
            rows=row_count,
            cols=total_cols,
            formulas=formulas,
            truncated=truncated,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _open(self, wb: Any, sheet: str, address: str) -> tuple[Any, Any, RangeAddress]:
        """Resolve ``sheet``/``address`` to ``(worksheet, range, RangeAddress)``."""

        worksheet = _require_sheet(wb, sheet)
        range_obj = _resolve_range(worksheet, address, sheet_name=_safe_str(worksheet.Name))
        return worksheet, range_obj, _describe_range(range_obj, sheet=_safe_str(worksheet.Name))


# ---------------------------------------------------------------------------
# Window / bound helpers (no COM)
# ---------------------------------------------------------------------------
def _validate_window(offset: int, limit: int | None) -> None:
    """Reject an invalid ``offset`` / ``limit`` before any COM access."""

    if offset < 0:
        raise InvalidRangeError("Argument 'offset' must be >= 0.", context={"offset": offset})
    if limit is not None and limit < 1:
        raise InvalidRangeError(
            "Argument 'limit' must be >= 1 when provided.", context={"limit": limit}
        )


def _window(
    total_rows: int, total_cols: int, offset: int, limit: int | None
) -> tuple[int, int, bool]:
    """Compute the bounded ``(start, end, truncated)`` row window.

    Rejects a request whose full size exceeds
    :data:`~excel_vba_mcp.domain.models.MAX_RANGE_CELLS` before any cell is read,
    then bounds the returned rows by
    :data:`~excel_vba_mcp.domain.models.MAX_RESPONSE_CELLS` (setting
    ``truncated`` when the response is cut short).
    """

    if total_rows * total_cols > MAX_RANGE_CELLS:
        raise InvalidRangeError(
            f"Range {total_rows}x{total_cols}={total_rows * total_cols} exceeds "
            f"MAX_RANGE_CELLS ({MAX_RANGE_CELLS}).",
            context={
                "rows": total_rows,
                "cols": total_cols,
                "cells": total_rows * total_cols,
                "limit": MAX_RANGE_CELLS,
            },
        )

    start = min(offset, total_rows)
    end = total_rows if limit is None else min(total_rows, offset + limit)
    requested = end - start

    truncated = requested * total_cols > MAX_RESPONSE_CELLS
    if truncated:
        max_rows = max(0, MAX_RESPONSE_CELLS // max(1, total_cols))
        end = start + min(requested, max_rows)
    return start, end, truncated


__all__ = ["RangesAdapter"]
