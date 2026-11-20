"""Excel Table (ListObject) MCP tools (task_026, spec §7.6 / §6.5).

Thin, async tool handlers for the table surface: ``tables_list``, ``table_get``,
``table_read_rows``, ``table_append_rows``, ``table_insert_row``,
``table_update_rows``, ``table_delete_rows``, ``table_add_column``,
``table_delete_column``, ``table_create`` and ``table_resize``. They are
deliberately **thin**: each validates its arguments, delegates to
:class:`~excel_vba_mcp.application.tables_service.TablesService` and returns a
domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` /
``win32com`` import lives in this module; all Excel access happens on the
single-STA worker inside the service.

Mutating tools (``table_append_rows`` / ``table_insert_row`` /
``table_update_rows`` / ``table_delete_rows`` / ``table_add_column`` /
``table_delete_column`` / ``table_create`` / ``table_resize``) only succeed when
the service guard chain passes: writable session, snapshot, approval
(deny-by-default).

Like the other tool modules, the active service is published here via
:func:`set_tables_service` (called from the server lifespan) and read via
:func:`get_tables_service`.
"""

from __future__ import annotations

from typing import Any

from excel_vba_mcp.adapters.excel.tables_adapter import TableMutationResult, TableRows
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import TableInfo

# The active tables service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_tables_service: TablesService | None = None


def set_tables_service(service: TablesService | None) -> None:
    """Publish (or clear) the active tables service used by the tools."""

    global _tables_service
    _tables_service = service


def get_tables_service() -> TablesService:
    """Return the active tables service or raise if it is not configured."""

    if _tables_service is None:
        raise ExcelVbaMcpError(
            "The tables service is not initialized.",
            context={"tool": "table_tools"},
        )
    return _tables_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


def _require_int(value: Any, field: str, *, minimum: int | None = None) -> int:
    """Validate that an argument is an integer (optionally with a minimum)."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be an integer.",
            context={"field": field},
        )
    if minimum is not None and value < minimum:
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be >= {minimum}.",
            context={"field": field, "value": value},
        )
    return value


def _require_row_dicts(value: Any, field: str) -> list[dict[str, Any]]:
    """Validate that an argument is a list of ``{header: value}`` dicts."""

    if not isinstance(value, list) or not value:
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty list of row objects.",
            context={"field": field},
        )
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(value):
        if not isinstance(row, dict):
            raise ExcelVbaMcpError(
                f"Argument {field!r}[{index}] must be an object mapping header -> value.",
                context={"field": field, "index": index},
            )
        rows.append(dict(row))
    return rows


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------
async def tables_list(workbook_id: str) -> list[TableInfo]:
    """List every Excel Table (ListObject) of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`TableInfo` per ``ListObject`` on every worksheet, with range
        metadata, totals, style, row count and columns.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_tables_service()
    return await service.list_tables(workbook_id)


async def table_get(workbook_id: str, name: str) -> TableInfo:
    """Return the :class:`TableInfo` for a single table by name.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table ``Name`` / ``DisplayName``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_tables_service()
    return await service.get_table(workbook_id, name)


async def table_read_rows(
    workbook_id: str,
    name: str,
    offset: int = 0,
    limit: int | None = None,
) -> TableRows:
    """Read a typed, bounded slice of a table's data rows.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        offset: 0-based data-row index of the first row to return.
        limit: Optional cap on the number of rows returned.

    Returns:
        A :class:`TableRows` with typed values (``str`` / ``int`` / ``float`` /
        ``bool`` / ``None`` / ISO date / typed Excel error).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    offset = _require_int(offset, "offset", minimum=0)
    if limit is not None:
        limit = _require_int(limit, "limit", minimum=1)
    service = get_tables_service()
    return await service.read_table_rows(workbook_id, name, offset=offset, limit=limit)


# ---------------------------------------------------------------------------
# Mutating tools (guarded)
# ---------------------------------------------------------------------------
async def table_append_rows(
    workbook_id: str, name: str, rows: list[dict[str, Any]]
) -> TableMutationResult:
    """Append rows through ``ListRows.Add`` (structure-preserving).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        rows: A non-empty list of ``{header: value}`` objects.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    row_dicts = _require_row_dicts(rows, "rows")
    service = get_tables_service()
    return await service.append_table_rows(workbook_id, name, row_dicts)


async def table_insert_row(
    workbook_id: str,
    name: str,
    at: int,
    values: dict[str, Any] | None = None,
) -> TableMutationResult:
    """Insert one data row at 1-based position ``at`` via ``ListRows.Add``.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        at: 1-based insert position (``1..count+1``).
        values: Optional ``{header: value}`` object for the new row.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    at = _require_int(at, "at", minimum=1)
    row_values: dict[str, Any] | None = None
    if values is not None:
        if not isinstance(values, dict):
            raise ExcelVbaMcpError(
                "Argument 'values' must be an object mapping header -> value.",
                context={"field": "values"},
            )
        row_values = dict(values)
    service = get_tables_service()
    return await service.insert_table_row(workbook_id, name, at, row_values)


async def table_update_rows(
    workbook_id: str, name: str, updates: list[dict[str, Any]]
) -> TableMutationResult:
    """Update rows by explicit index or an unambiguous ``where`` filter.

    Each update is ``{"index": int, "values": {...}}`` or
    ``{"where": {header: value, ...}, "values": {...}}`` (exactly one selector).
    An ambiguous ``where`` is rejected with a preview; no cell is written.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        updates: A non-empty list of update objects.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    update_dicts = _require_row_dicts(updates, "updates")
    service = get_tables_service()
    return await service.update_table_rows(workbook_id, name, update_dicts)


async def table_delete_rows(
    workbook_id: str,
    name: str,
    indices: list[int] | None = None,
    where: dict[str, Any] | None = None,
) -> TableMutationResult:
    """Delete rows by explicit 1-based indices or an unambiguous ``where``.

    Provide exactly one of ``indices`` or ``where``. An ambiguous ``where`` is
    rejected with a preview; no row is deleted.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        indices: A list of unique 1-based data-row indices.
        where: A ``{header: value}`` filter matching exactly one row.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    if (indices is None) == (where is None):
        raise ExcelVbaMcpError(
            "Provide exactly one of 'indices' or 'where' for table_delete_rows.",
            context={"field": "indices/where"},
        )
    index_list: list[int] | None = None
    if indices is not None:
        if not isinstance(indices, list):
            raise ExcelVbaMcpError(
                "Argument 'indices' must be a list of integers.",
                context={"field": "indices"},
            )
        index_list = [_require_int(value, "indices[]", minimum=1) for value in indices]
    where_dict: dict[str, Any] | None = None
    if where is not None:
        if not isinstance(where, dict):
            raise ExcelVbaMcpError(
                "Argument 'where' must be an object mapping header -> value.",
                context={"field": "where"},
            )
        where_dict = dict(where)
    service = get_tables_service()
    return await service.delete_table_rows(workbook_id, name, index_list, where_dict)


async def table_add_column(
    workbook_id: str,
    name: str,
    header: str,
    position: int | None = None,
) -> TableMutationResult:
    """Add a column via ``ListColumns.Add(Position)`` with a unique header.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). ``header`` must be unique (case-insensitive).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        header: The new column header (must be unique).
        position: Optional 1-based column index (append when omitted).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(header, "header")
    if position is not None:
        position = _require_int(position, "position", minimum=1)
    service = get_tables_service()
    return await service.add_table_column(workbook_id, name, header, position)


async def table_delete_column(workbook_id: str, name: str, header: str) -> TableMutationResult:
    """Delete the column identified by ``header`` via ``ListColumns``.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). A missing/ambiguous header is rejected.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        header: The column header to delete.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(header, "header")
    service = get_tables_service()
    return await service.delete_table_column(workbook_id, name, header)


async def table_create(
    workbook_id: str,
    sheet: str,
    range_or_name: str,
    style: str | None = None,
) -> TableMutationResult:
    """Create a ListObject on ``sheet`` over ``range_or_name``.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        sheet: The worksheet that will own the table.
        range_or_name: An A1 range address on ``sheet`` or a workbook defined name.
        style: Optional ``TableStyle`` name to apply.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(range_or_name, "range_or_name")
    service = get_tables_service()
    return await service.create_table(workbook_id, sheet, range_or_name, style=style)


async def table_resize(workbook_id: str, name: str, new_range: str) -> TableMutationResult:
    """Resize a table to ``new_range`` (must include the header row).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The table name.
        new_range: The new A1 range address on the table's own worksheet.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(new_range, "new_range")
    service = get_tables_service()
    return await service.resize_table(workbook_id, name, new_range)


__all__ = [
    "get_tables_service",
    "set_tables_service",
    "table_add_column",
    "table_append_rows",
    "table_create",
    "table_delete_column",
    "table_delete_rows",
    "table_get",
    "table_insert_row",
    "table_read_rows",
    "table_resize",
    "table_update_rows",
    "tables_list",
]
