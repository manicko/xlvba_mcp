"""Sheets inventory + worksheet operation MCP tools (task_024, spec §7.5/§6.5).

Thin, async tool handlers for the sheet surface: ``sheets_list``, ``sheet_get``,
``sheet_add``, ``sheet_rename``, ``sheet_copy``, ``sheet_delete`` and
``sheet_set_visibility``. They are deliberately **thin**: each validates its
arguments, delegates to
:class:`~excel_vba_mcp.application.sheets_service.SheetsService` and returns a
domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` / ``win32com``
import lives in this module; all Excel access happens on the single-STA worker
inside the service.

Mutating tools (``sheet_add`` / ``sheet_rename`` / ``sheet_copy`` /
``sheet_delete`` / ``sheet_set_visibility``) only succeed when the service guard
chain passes: writable session, snapshot, approval (deny-by-default) and — for
delete — the visible-sheet guard.

Like the other tool modules, the active service is published here via
:func:`set_sheets_service` (called from the server lifespan) and read via
:func:`get_sheets_service`.
"""

from __future__ import annotations

from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import SheetInfo

# The active sheets service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_sheets_service: SheetsService | None = None


def set_sheets_service(service: SheetsService | None) -> None:
    """Publish (or clear) the active sheets service used by the tools."""

    global _sheets_service
    _sheets_service = service


def get_sheets_service() -> SheetsService:
    """Return the active sheets service or raise if it is not configured."""

    if _sheets_service is None:
        raise ExcelVbaMcpError(
            "The sheets service is not initialized.",
            context={"tool": "sheet_tools"},
        )
    return _sheets_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def sheets_list(workbook_id: str) -> list[SheetInfo]:
    """List every sheet (all types) of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`SheetInfo` per ``Sheets`` item (Worksheet / Chart / Macro /
        Dialog), each with name, index, object_type, visibility and — for
        worksheets — code_name, protected, table_count and used_range.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_sheets_service()
    return await service.list_sheets(workbook_id)


async def sheet_get(workbook_id: str, name: str) -> SheetInfo:
    """Return the :class:`SheetInfo` for a single sheet by name.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The sheet name (as shown in Excel).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_sheets_service()
    return await service.get_sheet(workbook_id, name)


async def sheet_add(workbook_id: str, name: str, sheet_type: str = "worksheet") -> SheetInfo:
    """Add a worksheet named ``name`` (MVP: Worksheet only).

    Guarded mutation: requires a writable session, a successful snapshot and an
    approval grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The new sheet name.
        sheet_type: Only ``"worksheet"`` is supported in the MVP.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_sheets_service()
    return await service.add_sheet(workbook_id, name, sheet_type)


async def sheet_rename(workbook_id: str, name: str, new_name: str) -> SheetInfo:
    """Rename a sheet (``CodeName`` is never changed).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The current sheet name.
        new_name: The new sheet name.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(new_name, "new_name")
    service = get_sheets_service()
    return await service.rename_sheet(workbook_id, name, new_name)


async def sheet_copy(workbook_id: str, name: str) -> SheetInfo:
    """Copy a sheet and return the new sheet's DTO.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The sheet to copy.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_sheets_service()
    return await service.copy_sheet(workbook_id, name)


async def sheet_delete(workbook_id: str, name: str) -> SheetInfo:
    """Delete a sheet (destructive; fully guarded).

    Requires a writable session, a snapshot and an approval grant; the delete is
    rejected when it would leave no visible worksheet. Returns the pre-delete
    :class:`SheetInfo` of the removed sheet.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The sheet to delete.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_sheets_service()
    return await service.delete_sheet(workbook_id, name)


async def sheet_set_visibility(workbook_id: str, name: str, visible: bool) -> SheetInfo:
    """Set a sheet's visibility (``xlSheetVisible`` / ``xlSheetHidden``).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). Hiding the last visible worksheet is rejected.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The sheet whose visibility changes.
        visible: ``True`` for ``xlSheetVisible`` (0), ``False`` for
            ``xlSheetHidden`` (1).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_sheets_service()
    return await service.set_visibility(workbook_id, name, bool(visible))


__all__ = [
    "get_sheets_service",
    "set_sheets_service",
    "sheet_add",
    "sheet_copy",
    "sheet_delete",
    "sheet_get",
    "sheet_rename",
    "sheet_set_visibility",
    "sheets_list",
]
