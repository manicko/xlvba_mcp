"""STA-local Excel Table (ListObject) adapter (task_026, spec §7.6 / §6.5).

:class:`TablesAdapter` is the single place where the Excel ``ListObjects`` /
``ListRows`` / ``ListColumns`` object model is touched. It is a **pure STA-local
helper**: every method is synchronous, receives the live ``Workbook`` COM proxy
(``Any``) and returns only Pydantic DTOs / primitives, so results cross the
single-STA worker boundary safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`).
It never imports ``pywin32`` at module import time and never calls
``.Activate`` / ``.Select`` on any object.

Why ListRows / ListColumns (spec §7.6):

* An Excel Table is a ``ListObject``, not a plain range. Rows and columns are
  added **through** ``ListRows.Add`` / ``ListColumns.Add(Position)`` so the table
  keeps its name, filters, totals and structured references. Writing adjacent
  cells directly would silently break the table structure, so this adapter never
  does that for structural changes.
* Updates / deletes address an **explicit, unambiguous** row index. When a
  request carries a ``where`` filter that matches zero or more than one row, the
  adapter raises a structured :class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError`
  with a ``preview`` in ``context`` instead of guessing.

Typed values (spec §7.6): ``str`` / ``int`` / ``float`` / ``bool`` / ``None`` /
``datetime`` are preserved; dates are surfaced as ``CellValueKind.DATE`` with an
ISO-8601 string (the domain :class:`~excel_vba_mcp.domain.models.CellValue`
``value`` union has no ``datetime`` member); Excel cell errors are surfaced as
``CellValueKind.ERROR`` with the Excel error literal, never as an ordinary string.

All mutating methods assume the caller already validated the session
(writable), a snapshot exists and approval was granted; the adapter only executes
the Excel object-model calls on the STA thread and captures before/after metadata.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    CellValue,
    DomainModel,
    TableColumnInfo,
    TableInfo,
)
from excel_vba_mcp.domain.statuses import CellValueKind

# Upper bound on the number of cells a single ``read_rows`` response returns. A
# larger table is truncated (``TableRows.truncated`` is set) rather than dumping
# an unbounded payload across the worker boundary.
_MAX_RESPONSE_CELLS = 2000

# Upper bound on the number of preview rows attached to an ambiguity error.
_MAX_PREVIEW_ROWS = 10

# Excel error literals returned by ``Range.Value`` for error cells (spec §7.6:
# "представлять ошибки ячеек отдельным типизированным значением").
_EXCEL_ERROR_CODES: frozenset[str] = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
    }
)

# Excel error cells read through ``Range.Value`` surface as signed OLE error
# codes (``xlErr*``), not as the literal text. Map the code -> literal so error
# cells are typed as ``CellValueKind.ERROR`` instead of a bogus number.
_EXCEL_ERROR_CODE_TO_LITERAL: dict[int, str] = {
    -2147221244: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
}

# ``ListColumn.TotalsCalculation`` codes -> stable function names.
_TOTALS_FUNCTIONS: dict[int, str] = {
    -4143: "none",
    -4157: "sum",
    -4112: "count",
    -4114: "average",
    -4136: "max",
    -4139: "min",
    -4155: "stddev",
    -4156: "var",
    -4113: "count_nums",
    9: "custom",
}

# ``ListObjects.Add`` source type: xlSrcRange.
_XL_SRC_RANGE = 1


# ---------------------------------------------------------------------------
# DTOs (no COM)
# ---------------------------------------------------------------------------
class TableRows(DomainModel):
    """A typed, bounded slice of a table's data rows (spec §7.6).

    ``rows`` holds one inner list per data row, each inner list positionally
    aligned with ``columns`` (the header names). ``offset`` is the 0-based index
    of the first returned data row; ``limit`` is the requested cap (``None`` when
    no cap was given). ``truncated`` is ``True`` when the response-size limit or
    the requested ``limit`` cut the result short.
    """

    name: str
    sheet_name: str
    columns: list[str]
    offset: int
    limit: int | None
    total_rows: int
    returned_rows: int
    truncated: bool
    rows: list[list[CellValue]]


class TableMutationResult(DomainModel):
    """Before/after metadata for one ListObject mutation (spec §7.6).

    ``before`` / ``after`` are the :class:`TableInfo` snapshots captured on the
    STA thread around the mutation (``before`` is ``None`` for ``table_create``).
    ``affected_rows`` lists the 1-based data-row indices the mutation touched.
    """

    name: str
    action: str
    before: TableInfo | None
    after: TableInfo
    affected_rows: list[int] = []
    notes: list[str] = []


# ---------------------------------------------------------------------------
# Value typing helpers (no COM beyond the supplied proxy)
# ---------------------------------------------------------------------------
def _column_letter(index: int) -> str:
    """Return the Excel column letter(s) for a 1-based column ``index``."""

    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively (never raises)."""

    try:
        text = value
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    if text is None:
        return default
    if isinstance(text, str):
        return text
    return str(text)


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively (never raises)."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_bool(value: Any, default: bool = False) -> bool:
    """Read a COM boolean property defensively (never raises)."""

    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_address(range_obj: Any) -> str | None:
    """Return ``range_obj.Address`` as a string, or ``None`` when absent."""

    if range_obj is None:
        return None
    try:
        address = range_obj.Address
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    if address is None:
        return None
    text = str(address)
    return text or None


def _cell_value(value: Any, *, address: str) -> CellValue:
    """Classify a raw cell value into a typed :class:`CellValue`.

    Order matters: ``bool`` is checked before ``int`` (``bool`` is an ``int``
    subclass), and error codes/literals before the generic number / string case.
    """

    if value is None:
        return CellValue(kind=CellValueKind.EMPTY, value=None, address=address)
    if isinstance(value, bool):
        return CellValue(kind=CellValueKind.BOOLEAN, value=value, address=address)
    if isinstance(value, datetime):
        return CellValue(kind=CellValueKind.DATE, value=value.isoformat(), address=address)
    if isinstance(value, int) and value in _EXCEL_ERROR_CODE_TO_LITERAL:
        return CellValue(
            kind=CellValueKind.ERROR, value=_EXCEL_ERROR_CODE_TO_LITERAL[value], address=address
        )
    if isinstance(value, (int, float)):
        return CellValue(kind=CellValueKind.NUMBER, value=value, address=address)
    if isinstance(value, str):
        if value in _EXCEL_ERROR_CODES:
            return CellValue(kind=CellValueKind.ERROR, value=value, address=address)
        return CellValue(kind=CellValueKind.STRING, value=value, address=address)
    # Unknown COM variant: fall back to its string form rather than leaking it.
    return CellValue(kind=CellValueKind.STRING, value=str(value), address=address)


def _cell_kind_name(value: Any) -> str:
    """Return a coarse type name for a sampled cell (``TableColumnInfo.data_type``)."""

    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, datetime):
        return "date"
    if isinstance(value, int) and value in _EXCEL_ERROR_CODE_TO_LITERAL:
        return "error"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str) and value in _EXCEL_ERROR_CODES:
        return "error"
    return "string"


def _normalize_grid(raw: Any) -> tuple[tuple[Any, ...], ...]:
    """Normalize a ``Range.Value`` read into a tuple-of-tuples grid."""

    if not isinstance(raw, tuple):
        return ((raw,),)
    if not raw:
        return ((),)
    if isinstance(raw[0], tuple):
        return tuple(tuple(row) for row in raw)
    return (tuple(raw),)


# ---------------------------------------------------------------------------
# Table description (read-only)
# ---------------------------------------------------------------------------
def _sheet_code_name(sheet: Any) -> str | None:
    """Read ``Worksheet.CodeName`` (read-only identifier, never written)."""

    try:
        value = sheet.CodeName
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    if value is None:
        return None
    text = str(value)
    return text or None


def _describe_column(wb: Any, list_column: Any, index: int) -> TableColumnInfo:
    """Build a :class:`TableColumnInfo` for one ``ListColumn`` (STA-local)."""

    name = _safe_str(list_column.Name)
    data_type: str | None = None
    try:
        body = list_column.DataBodyRange
    except Exception:  # noqa: BLE001 - defensive COM read
        body = None
    if body is not None:
        try:
            data_type = _cell_kind_name(body.Cells(1, 1).Value)
        except Exception:  # noqa: BLE001 - defensive COM read
            data_type = None

    totals_code = _safe_int(list_column.TotalsCalculation, default=-4143)
    totals_function = _TOTALS_FUNCTIONS.get(totals_code)
    has_totals = totals_code != -4143
    return TableColumnInfo(
        name=name,
        index=index,
        data_type=data_type,
        has_totals_calculation=has_totals,
        totals_function=totals_function,
    )


def describe_table(wb: Any, sheet: Any, list_object: Any) -> TableInfo:
    """Build a :class:`TableInfo` DTO for one ``ListObject`` (STA-local).

    Args:
        wb: The owning ``Workbook`` proxy (unused for the description itself but
            kept for signature symmetry with the other adapters).
        sheet: The worksheet proxy that owns ``list_object``.
        list_object: One ``ListObjects`` item proxy.

    Returns:
        The populated :class:`TableInfo` (name, sheet, range metadata, totals,
        style, row count and columns).
    """

    name = _safe_str(list_object.Name)
    display_name = _safe_str(getattr(list_object, "DisplayName", None)) or None
    sheet_name = _safe_str(sheet.Name)
    range_address = _safe_str(list_object.Range.Address)
    header_row_range = _safe_address(getattr(list_object, "HeaderRowRange", None))
    data_body_range = _safe_address(getattr(list_object, "DataBodyRange", None))
    total_row_range = _safe_address(getattr(list_object, "TotalRowRange", None))
    show_totals = _safe_bool(getattr(list_object, "ShowTotals", False), default=False)

    style_value = getattr(list_object, "TableStyle", None)
    table_style = _safe_str(style_value) if style_value is not None else None
    if not table_style:
        table_style = None

    try:
        row_count = _safe_int(list_object.ListRows.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        row_count = 0

    columns: list[TableColumnInfo] = []
    try:
        list_columns = list_object.ListColumns
        for index in range(1, _safe_int(list_columns.Count) + 1):
            try:
                column = list_columns.Item(index)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
            columns.append(_describe_column(wb, column, index))
    except Exception:  # noqa: BLE001 - defensive COM read
        pass

    return TableInfo(
        name=name,
        display_name=display_name,
        sheet_name=sheet_name,
        sheet_code_name=_sheet_code_name(sheet),
        range_address=range_address,
        header_row_range=header_row_range,
        data_body_range=data_body_range,
        total_row_range=total_row_range,
        show_totals=show_totals,
        table_style=table_style,
        row_count=row_count,
        columns=columns,
    )


def _iter_list_objects(wb: Any) -> list[tuple[Any, Any]]:
    """Return ``(sheet, list_object)`` pairs for every table on every worksheet."""

    pairs: list[tuple[Any, Any]] = []
    try:
        worksheets = wb.Worksheets
    except Exception:  # noqa: BLE001 - defensive COM read
        return pairs
    for sheet in worksheets:
        try:
            list_objects = sheet.ListObjects
            for index in range(1, _safe_int(list_objects.Count) + 1):
                try:
                    pairs.append((sheet, list_objects.Item(index)))
                except Exception:  # noqa: BLE001 - defensive per-item COM read
                    continue
        except Exception:  # noqa: BLE001 - sheet without ListObjects
            continue
    return pairs


def _require_list_object(wb: Any, name: str) -> tuple[Any, Any]:
    """Resolve ``name`` to ``(sheet, list_object)`` or raise a structured error."""

    target = name.strip().lower()
    for sheet, list_object in _iter_list_objects(wb):
        candidate = _safe_str(list_object.Name).lower()
        display = _safe_str(getattr(list_object, "DisplayName", "")).lower()
        if target in (candidate, display):
            return sheet, list_object
    raise ExcelVbaMcpError(
        f"Table {name!r} was not found in the workbook.",
        context={"table": name},
    )


def _header_index(list_object: Any, header: str) -> int:
    """Return the 1-based column index for ``header`` (case-insensitive)."""

    target = header.strip().lower()
    try:
        list_columns = list_object.ListColumns
        for index in range(1, _safe_int(list_columns.Count) + 1):
            column = list_columns.Item(index)
            if _safe_str(column.Name).strip().lower() == target:
                return index
    except Exception:  # noqa: BLE001 - defensive COM read
        pass
    raise ExcelVbaMcpError(
        f"Column {header!r} was not found in the table.",
        context={"header": header},
    )


def _header_names(list_object: Any) -> list[str]:
    """Return the ordered header names of ``list_object``."""

    names: list[str] = []
    try:
        list_columns = list_object.ListColumns
        for index in range(1, _safe_int(list_columns.Count) + 1):
            names.append(_safe_str(list_columns.Item(index).Name))
    except Exception:  # noqa: BLE001 - defensive COM read
        pass
    return names


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class TablesAdapter:
    """STA-local Excel Table inventory + row/column operations (no COM at import)."""

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def list_tables(self, wb: Any) -> list[TableInfo]:
        """Return one :class:`TableInfo` per ``ListObject`` on every worksheet."""

        result: list[TableInfo] = []
        for sheet, list_object in _iter_list_objects(wb):
            result.append(describe_table(wb, sheet, list_object))
        return result

    def get_table(self, wb: Any, name: str) -> TableInfo:
        """Return the :class:`TableInfo` for a single table by name/display name."""

        sheet, list_object = _require_list_object(wb, name)
        return describe_table(wb, sheet, list_object)

    # ------------------------------------------------------------------
    # Rows (read-only, typed, bounded)
    # ------------------------------------------------------------------
    def read_rows(
        self,
        wb: Any,
        name: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> TableRows:
        """Read a typed, bounded slice of a table's data rows (spec §7.6).

        ``offset`` is a 0-based data-row index; ``limit`` caps the returned rows.
        The response is additionally capped at ``_MAX_RESPONSE_CELLS`` cells so a
        huge table cannot flood the worker boundary (``truncated`` is set).
        """

        if offset < 0:
            raise ExcelVbaMcpError(
                "Argument 'offset' must be >= 0.",
                context={"offset": offset},
            )
        if limit is not None and limit < 1:
            raise ExcelVbaMcpError(
                "Argument 'limit' must be >= 1 when provided.",
                context={"limit": limit},
            )

        sheet, list_object = _require_list_object(wb, name)
        columns = _header_names(list_object)
        total_rows = _safe_int(list_object.ListRows.Count)
        body = getattr(list_object, "DataBodyRange", None)

        if body is None or total_rows == 0 or not columns:
            return TableRows(
                name=_safe_str(list_object.Name),
                sheet_name=_safe_str(sheet.Name),
                columns=columns,
                offset=offset,
                limit=limit,
                total_rows=total_rows,
                returned_rows=0,
                truncated=False,
                rows=[],
            )

        start = min(offset, total_rows)
        end = total_rows if limit is None else min(total_rows, offset + limit)
        requested = end - start
        column_count = len(columns)

        # Enforce the response-size ceiling by shrinking the row window.
        max_rows_by_budget = max(0, _MAX_RESPONSE_CELLS // max(1, column_count))
        truncated = requested > max_rows_by_budget
        row_count = min(requested, max_rows_by_budget)
        end = start + row_count

        if row_count == 0:
            return TableRows(
                name=_safe_str(list_object.Name),
                sheet_name=_safe_str(sheet.Name),
                columns=columns,
                offset=offset,
                limit=limit,
                total_rows=total_rows,
                returned_rows=0,
                truncated=truncated or offset >= total_rows,
                rows=[],
            )

        top_row = _safe_int(body.Row) + start
        left_col = _safe_int(body.Column)
        sub_range = body.Worksheet.Range(
            body.Worksheet.Cells(top_row, left_col),
            body.Worksheet.Cells(top_row + row_count - 1, left_col + column_count - 1),
        )
        raw = sub_range.Value
        grid = _normalize_grid(raw)
        if len(grid) < row_count:  # single-row / single-column normalization guard
            grid = tuple(grid) + ((),) * (row_count - len(grid))

        rows: list[list[CellValue]] = []
        for row_offset in range(row_count):
            grid_row = grid[row_offset] if row_offset < len(grid) else ()
            excel_row = top_row + row_offset
            values: list[CellValue] = []
            for col_offset in range(column_count):
                raw_value = grid_row[col_offset] if col_offset < len(grid_row) else None
                address = f"{_column_letter(left_col + col_offset)}{excel_row}"
                values.append(_cell_value(raw_value, address=address))
            rows.append(values)

        return TableRows(
            name=_safe_str(list_object.Name),
            sheet_name=_safe_str(sheet.Name),
            columns=columns,
            offset=offset,
            limit=limit,
            total_rows=total_rows,
            returned_rows=len(rows),
            truncated=truncated,
            rows=rows,
        )

    # ------------------------------------------------------------------
    # Rows (mutations via ListRows)
    # ------------------------------------------------------------------
    def append_rows(self, wb: Any, name: str, rows: list[dict[str, Any]]) -> TableMutationResult:
        """Append data rows through ``ListRows.Add`` (structure-preserving).

        Each ``rows`` entry maps header name -> value. Unknown headers raise a
        structured error before any row is added. ``None`` clears the cell.
        """

        if not rows:
            raise ExcelVbaMcpError(
                "Argument 'rows' must contain at least one row.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)
        TablesAdapter._validate_row_headers(list_object, rows)

        first_new_index = _safe_int(list_object.ListRows.Count) + 1
        for row in rows:
            new_row = list_object.ListRows.Add()
            _write_row_values(new_row, list_object, row)

        after = describe_table(wb, sheet, list_object)
        new_count = _safe_int(list_object.ListRows.Count)
        affected = list(range(first_new_index, new_count + 1))
        return TableMutationResult(
            name=before.name,
            action="append_rows",
            before=before,
            after=after,
            affected_rows=affected,
            notes=[f"appended {len(rows)} row(s)"],
        )

    def insert_row(
        self,
        wb: Any,
        name: str,
        at: int,
        values: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        """Insert one data row at 1-based position ``at`` via ``ListRows.Add``.

        ``at`` must be in ``1..count+1`` (``count+1`` appends). Optional
        ``values`` maps header name -> value for the new row.
        """

        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)
        count = _safe_int(list_object.ListRows.Count)
        if at < 1 or at > count + 1:
            raise ExcelVbaMcpError(
                f"Insert position {at} is out of range; expected 1..{count + 1}.",
                context={"table": name, "at": at, "row_count": count},
            )
        if values:
            TablesAdapter._validate_row_headers(list_object, [values])
        new_row = list_object.ListRows.Add(at)
        if values:
            _write_row_values(new_row, list_object, values)

        after = describe_table(wb, sheet, list_object)
        return TableMutationResult(
            name=before.name,
            action="insert_row",
            before=before,
            after=after,
            affected_rows=[at],
            notes=[f"inserted a row at position {at}"],
        )

    def update_rows(
        self,
        wb: Any,
        name: str,
        updates: list[dict[str, Any]],
    ) -> TableMutationResult:
        """Update rows by an explicit 1-based index or an unambiguous ``where``.

        Each ``updates`` entry is either ``{"index": int, "values": {...}}`` or
        ``{"where": {header: value, ...}, "values": {...}}`` (never both). A
        ``where`` filter must match **exactly one** data row; otherwise a
        structured ambiguity error with a ``preview`` is raised and no cell is
        written (spec §7.6: never auto-guess similar rows).
        """

        if not updates:
            raise ExcelVbaMcpError(
                "Argument 'updates' must contain at least one update.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)
        resolved = self._resolve_updates(list_object, updates)
        for index, values in resolved:
            self._write_row_cells(list_object, index, values)

        after = describe_table(wb, sheet, list_object)
        affected = sorted({index for index, _ in resolved})
        return TableMutationResult(
            name=before.name,
            action="update_rows",
            before=before,
            after=after,
            affected_rows=affected,
            notes=[f"updated {len(resolved)} row(s)"],
        )

    def delete_rows(
        self,
        wb: Any,
        name: str,
        indices: list[int] | None = None,
        where: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        """Delete data rows by explicit 1-based indices or an unambiguous filter.

        ``indices`` must be unique and within ``1..count``; a ``where`` filter
        must match exactly one data row. Either selection is validated up front
        so a bad request never mutates data; duplicates/out-of-range raise a
        structured error listing the valid range / a preview.
        """

        if (indices is None) == (where is None):
            raise ExcelVbaMcpError(
                "Provide exactly one of 'indices' or 'where' for delete_rows.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)

        if indices is not None:
            if not indices:
                raise ExcelVbaMcpError(
                    "Argument 'indices' must contain at least one row index.",
                    context={"table": name},
                )
            targets = self._validate_indices(list_object, indices)
        else:
            assert where is not None  # narrowed by the xor check above
            targets = [self._resolve_where(list_object, where)]

        for index in sorted(set(targets), reverse=True):
            list_object.ListRows.Item(index).Delete()

        after = describe_table(wb, sheet, list_object)
        return TableMutationResult(
            name=before.name,
            action="delete_rows",
            before=before,
            after=after,
            affected_rows=sorted(targets),
            notes=[f"deleted {len(set(targets))} row(s)"],
        )

    # ------------------------------------------------------------------
    # Columns (mutations via ListColumns)
    # ------------------------------------------------------------------
    def add_column(
        self,
        wb: Any,
        name: str,
        header: str,
        position: int | None = None,
    ) -> TableMutationResult:
        """Add a column via ``ListColumns.Add(Position)`` with a unique header.

        ``header`` must be non-empty and unique (case-insensitive) across the
        table; ``position`` is a 1-based column index (append when ``None``).
        """

        if not isinstance(header, str) or not header.strip():
            raise ExcelVbaMcpError(
                "Argument 'header' must be a non-empty string.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)
        existing = [h.strip().lower() for h in _header_names(list_object)]
        if header.strip().lower() in existing:
            raise ExcelVbaMcpError(
                f"Column header {header!r} already exists (headers must be unique).",
                context={"table": name, "header": header, "existing": existing},
            )
        column_count = len(existing)
        if position is not None and (position < 1 or position > column_count + 1):
            raise ExcelVbaMcpError(
                f"Column position {position} is out of range; expected 1..{column_count + 1}.",
                context={"table": name, "position": position, "column_count": column_count},
            )

        if position is None:
            new_column = list_object.ListColumns.Add()
        else:
            new_column = list_object.ListColumns.Add(position)
        new_column.Name = header

        after = describe_table(wb, sheet, list_object)
        return TableMutationResult(
            name=before.name,
            action="add_column",
            before=before,
            after=after,
            notes=[f"added column {header!r}"],
        )

    def delete_column(self, wb: Any, name: str, header: str) -> TableMutationResult:
        """Delete the column identified by ``header`` via ``ListColumns``.

        The header must resolve to exactly one column; a missing or duplicated
        header raises a structured error (with a preview for duplicates) instead
        of guessing.
        """

        if not isinstance(header, str) or not header.strip():
            raise ExcelVbaMcpError(
                "Argument 'header' must be a non-empty string.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)

        target = header.strip().lower()
        matches = [
            index
            for index, column_name in enumerate(_header_names(list_object), start=1)
            if column_name.strip().lower() == target
        ]
        if not matches:
            raise ExcelVbaMcpError(
                f"Column {header!r} was not found in the table.",
                context={"table": name, "header": header, "headers": _header_names(list_object)},
            )
        if len(matches) > 1:
            raise ExcelVbaMcpError(
                f"Column header {header!r} is ambiguous (matched {matches}); "
                "refusing to guess which column to delete.",
                context={
                    "table": name,
                    "header": header,
                    "reason": "ambiguous",
                    "preview": matches,
                },
            )

        list_object.ListColumns.Item(matches[0]).Delete()
        after = describe_table(wb, sheet, list_object)
        return TableMutationResult(
            name=before.name,
            action="delete_column",
            before=before,
            after=after,
            notes=[f"deleted column {header!r}"],
        )

    # ------------------------------------------------------------------
    # Table lifecycle
    # ------------------------------------------------------------------
    def create_table(
        self,
        wb: Any,
        sheet: str,
        range_or_name: str,
        *,
        style: str | None = None,
    ) -> TableMutationResult:
        """Create a ListObject on ``sheet`` over ``range_or_name`` via ``Add``.

        ``range_or_name`` is either an A1 range address on ``sheet`` or a defined
        name that resolves to a range. The header row is taken from the source
        range (``xlSrcRange`` + ``xlYes``). An optional ``style`` sets
        ``TableStyle``.
        """

        if not isinstance(sheet, str) or not sheet.strip():
            raise ExcelVbaMcpError(
                "Argument 'sheet' must be a non-empty string.",
                context={"sheet": sheet},
            )
        if not isinstance(range_or_name, str) or not range_or_name.strip():
            raise ExcelVbaMcpError(
                "Argument 'range_or_name' must be a non-empty string.",
                context={"sheet": sheet},
            )
        worksheet = _require_sheet(wb, sheet)
        source = _resolve_source_range(wb, worksheet, range_or_name)

        try:
            list_object = worksheet.ListObjects.Add(
                SourceType=_XL_SRC_RANGE,
                Source=source,
                XlListObjectHasHeaders=1,
            )
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not create a table on {sheet!r} over {range_or_name!r}.",
                context={"sheet": sheet, "range_or_name": range_or_name},
            ) from exc

        if style:
            try:
                list_object.TableStyle = style
            except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
                raise ExcelVbaMcpError(
                    f"Could not apply table style {style!r}.",
                    context={"sheet": sheet, "style": style},
                ) from exc

        after = describe_table(wb, worksheet, list_object)
        return TableMutationResult(
            name=after.name,
            action="create_table",
            before=None,
            after=after,
            notes=[f"created table on {sheet!r} over {range_or_name!r}"],
        )

    def resize(self, wb: Any, name: str, new_range: str) -> TableMutationResult:
        """Resize a table to ``new_range`` via ``ListObject.Resize``.

        ``new_range`` is an A1 range address on the table's own worksheet and
        must include the header row.
        """

        if not isinstance(new_range, str) or not new_range.strip():
            raise ExcelVbaMcpError(
                "Argument 'new_range' must be a non-empty A1 range.",
                context={"table": name},
            )
        sheet, list_object = _require_list_object(wb, name)
        before = describe_table(wb, sheet, list_object)
        try:
            target = sheet.Range(new_range)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Range {new_range!r} is not valid on sheet {sheet.Name!r}.",
                context={"table": name, "new_range": new_range},
            ) from exc
        try:
            list_object.Resize(target)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not resize table {name!r} to {new_range!r}.",
                context={"table": name, "new_range": new_range},
            ) from exc

        after = describe_table(wb, sheet, list_object)
        return TableMutationResult(
            name=before.name,
            action="resize",
            before=before,
            after=after,
            notes=[f"resized to {new_range!r}"],
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _validate_row_headers(list_object: Any, rows: list[dict[str, Any]]) -> None:
        """Reject unknown headers before any row is added."""

        headers = {h.strip().lower() for h in _header_names(list_object)}
        for row in rows:
            for header in row:
                if header.strip().lower() not in headers:
                    raise ExcelVbaMcpError(
                        f"Row references unknown header {header!r}; "
                        f"available headers are {sorted(headers)}.",
                        context={"table": _safe_str(list_object.Name), "header": header},
                    )

    @staticmethod
    def _resolve_updates(
        list_object: Any, updates: list[dict[str, Any]]
    ) -> list[tuple[int, dict[str, Any]]]:
        """Resolve each update entry to ``(row_index, values)`` (no mutation)."""

        count = _safe_int(list_object.ListRows.Count)
        resolved: list[tuple[int, dict[str, Any]]] = []
        seen: set[int] = set()
        for entry in updates:
            has_index = "index" in entry and entry["index"] is not None
            has_where = "where" in entry and entry["where"] is not None
            if has_index == has_where:
                raise ExcelVbaMcpError(
                    "Each update must provide exactly one of 'index' or 'where'.",
                    context={"table": _safe_str(list_object.Name), "entry": _preview(entry)},
                )
            values = entry.get("values") or {}
            if has_index:
                index = _safe_int(entry["index"], default=0)
                if index < 1 or index > count:
                    raise ExcelVbaMcpError(
                        f"Update row index {index} is out of range; expected 1..{count}.",
                        context={
                            "table": _safe_str(list_object.Name),
                            "index": index,
                            "row_count": count,
                        },
                    )
            else:
                index = TablesAdapter._resolve_where(list_object, entry["where"])
            if index in seen:
                raise ExcelVbaMcpError(
                    f"Row index {index} is updated more than once; refusing to guess.",
                    context={"table": _safe_str(list_object.Name), "index": index},
                )
            seen.add(index)
            resolved.append((index, dict(values)))
        return resolved

    @staticmethod
    def _validate_indices(list_object: Any, indices: list[int]) -> list[int]:
        """Validate explicit row indices (range + uniqueness) before deleting."""

        count = _safe_int(list_object.ListRows.Count)
        if len(set(indices)) != len(indices):
            raise ExcelVbaMcpError(
                "Argument 'indices' contains duplicate row indices.",
                context={"table": _safe_str(list_object.Name), "indices": list(indices)},
            )
        for index in indices:
            if index < 1 or index > count:
                raise ExcelVbaMcpError(
                    f"Delete row index {index} is out of range; expected 1..{count}.",
                    context={
                        "table": _safe_str(list_object.Name),
                        "index": index,
                        "row_count": count,
                    },
                )
        return list(indices)

    @staticmethod
    def _resolve_where(list_object: Any, where: dict[str, Any]) -> int:
        """Resolve a ``where`` filter to exactly one 1-based row index or raise."""

        if not where:
            raise ExcelVbaMcpError(
                "A 'where' filter must reference at least one header.",
                context={"table": _safe_str(list_object.Name)},
            )
        header_index = {
            h.strip().lower(): i for i, h in enumerate(_header_names(list_object), start=1)
        }
        for key in where:
            if key.strip().lower() not in header_index:
                raise ExcelVbaMcpError(
                    f"Filter references unknown header {key!r}; "
                    f"available headers are {sorted(header_index)}.",
                    context={"table": _safe_str(list_object.Name), "header": key},
                )

        count = _safe_int(list_object.ListRows.Count)
        body = getattr(list_object, "DataBodyRange", None)
        if body is None or count == 0:
            raise ExcelVbaMcpError(
                "The 'where' filter matched no rows (the table has no data).",
                context={
                    "table": _safe_str(list_object.Name),
                    "reason": "ambiguous",
                    "match_count": 0,
                    "preview": [],
                },
            )

        matches: list[int] = []
        preview: list[dict[str, Any]] = []
        for row_index in range(1, count + 1):
            row_values: dict[str, Any] = {}
            matched = True
            for key, expected in where.items():
                col_index = header_index[key.strip().lower()]
                try:
                    actual = body.Cells(row_index, col_index).Value
                except Exception:  # noqa: BLE001 - defensive per-cell COM read
                    actual = None
                row_values[key] = _cell_value(actual, address="").model_dump(mode="json")
                if not _value_matches(actual, expected):
                    matched = False
            if matched:
                matches.append(row_index)
                if len(preview) < _MAX_PREVIEW_ROWS:
                    preview.append({"index": row_index, "values": row_values})

        if len(matches) != 1:
            raise ExcelVbaMcpError(
                f"The 'where' filter matched {len(matches)} rows; an unambiguous "
                "single-row match is required (refusing to guess).",
                context={
                    "table": _safe_str(list_object.Name),
                    "reason": "ambiguous",
                    "match_count": len(matches),
                    "preview": preview,
                },
            )
        return matches[0]

    @staticmethod
    def _write_row_cells(list_object: Any, index: int, values: dict[str, Any]) -> None:
        """Write ``{header: value}`` into an existing data row (1-based index)."""

        body = getattr(list_object, "DataBodyRange", None)
        if body is None:
            raise ExcelVbaMcpError(
                "The table has no data body to update.",
                context={"table": _safe_str(list_object.Name), "index": index},
            )
        for header, value in values.items():
            col_index = _header_index(list_object, header)
            body.Cells(index, col_index).Value = value


def _write_row_values(list_row: Any, list_object: Any, row: dict[str, Any]) -> None:
    """Write a ``{header: value}`` mapping into a ``ListRow`` (module helper)."""

    row_range = list_row.Range
    for header, value in row.items():
        col_index = _header_index(list_object, header)
        row_range.Cells(1, col_index).Value = value


def _value_matches(actual: Any, expected: Any) -> bool:
    """Return whether a cell value matches a filter expectation."""

    if isinstance(actual, bool) or isinstance(expected, bool):
        return bool(actual) == bool(expected)
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        return float(actual) == float(expected)
    if actual is None:
        return expected is None
    return str(actual).strip().lower() == str(expected).strip().lower()


def _preview(entry: dict[str, Any]) -> dict[str, Any]:
    """Return a small, serializable preview of an update entry."""

    return {str(key): value for key, value in entry.items()}


def _require_sheet(wb: Any, name: str) -> Any:
    """Resolve ``name`` to a worksheet proxy or raise a structured error."""

    try:
        return wb.Worksheets.Item(name)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise ExcelVbaMcpError(
            f"Worksheet {name!r} was not found in the workbook.",
            context={"sheet": name},
        ) from exc


def _resolve_source_range(wb: Any, worksheet: Any, range_or_name: str) -> Any:
    """Resolve ``range_or_name`` to a Range on ``worksheet`` (A1 or defined name)."""

    candidate = range_or_name.strip()
    try:
        return worksheet.Range(candidate)
    except Exception:  # noqa: BLE001 - fall through to defined-name resolution
        pass
    try:
        name = wb.Names(candidate)
        refers = name.RefersToRange
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise ExcelVbaMcpError(
            f"Range or defined name {range_or_name!r} could not be resolved on "
            f"sheet {worksheet.Name!r}.",
            context={"range_or_name": range_or_name},
        ) from exc
    return refers


__all__ = [
    "TableMutationResult",
    "TableRows",
    "TablesAdapter",
    "describe_table",
]
