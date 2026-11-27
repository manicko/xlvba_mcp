"""Sheets inventory + worksheet operation MCP tools (task_024, spec §7.5/§6.5).

Thin, async tool handlers for the sheet surface: the ``sheet`` /
``sheet_read`` flat-signature dispatchers (WP2) over the per-verb handlers
``sheets_list``, ``sheet_get``, ``sheet_add``, ``sheet_rename``, ``sheet_copy``,
``sheet_delete`` and ``sheet_set_visibility``. They are deliberately **thin**:
each validates its arguments, delegates to
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

from typing import Literal

from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
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


async def sheet(
    action: Literal["add", "rename", "copy", "set-visibility"],
    workbook_id: str | None = None,
    name: str | None = None,
    new_name: str | None = None,
    sheet_type: str = "worksheet",
    visible: bool | None = None,
) -> SheetInfo:
    """Sheet mutation dispatch: ``add``, ``rename``, ``copy`` or ``set-visibility`` (WP2).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`sheet_add` / :func:`sheet_rename` /
    :func:`sheet_copy` / :func:`sheet_set_visibility` handlers unchanged and
    forwards only the validated arguments. Per-action required parameters are
    enforced by the target handler (``add`` / ``rename`` / ``copy`` need
    ``name``; ``rename`` also needs ``new_name``; ``set-visibility`` needs
    ``name`` and ``visible``) so a missing required argument surfaces as a
    structured :class:`ExcelVbaMcpError` rather than a crash. An unknown
    ``action`` raises a structured :class:`ExcelVbaMcpError`.

    Args:
        action: ``add`` (new worksheet), ``rename``, ``copy`` or
            ``set-visibility``.
        workbook_id: The id returned by ``open``; required for every action.
        name: The sheet name; required for every action.
        new_name: The new sheet name; required for ``rename``.
        sheet_type: Only ``"worksheet"`` is supported (``add`` only).
        visible: ``True`` for ``xlSheetVisible``, ``False`` for ``xlSheetHidden``
            (``set-visibility`` only).

    Returns:
        The affected :class:`SheetInfo` (the new sheet for ``add`` / ``copy``,
        the renamed sheet for ``rename``, the updated sheet for
        ``set-visibility``).
    """

    if action == "add":
        return await sheet_add(workbook_id or "", name or "", sheet_type)
    if action == "rename":
        return await sheet_rename(workbook_id or "", name or "", new_name or "")
    if action == "copy":
        return await sheet_copy(workbook_id or "", name or "")
    if action == "set-visibility":
        if visible is None:
            raise ExcelVbaMcpError(
                "Argument 'visible' is required for the 'set-visibility' action.",
                context={"tool": "sheet", "action": action, "field": "visible"},
            )
        return await sheet_set_visibility(workbook_id or "", name or "", visible)
    raise ExcelVbaMcpError(
        f"Unknown sheet action {action!r}; expected 'add', 'rename', 'copy' or "
        "'set-visibility'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "sheet", "action": action},
    )


async def sheet_read(
    action: Literal["list", "get"],
    workbook_id: str | None = None,
    name: str | None = None,
) -> list[SheetInfo] | SheetInfo:
    """Sheet read dispatch: ``list`` or ``get`` (WP2).

    Thin flat-signature dispatcher (option (a)): routes to the existing
    :func:`sheets_list` / :func:`sheet_get` handlers unchanged. ``list`` requires
    only ``workbook_id``; ``get`` also requires ``name`` (enforced by the target
    handler). An unknown ``action`` raises a structured
    :class:`ExcelVbaMcpError`.

    Args:
        action: ``list`` (every sheet) or ``get`` (one sheet by name).
        workbook_id: The id returned by ``open``; required for both actions.
        name: The sheet name; required for ``get``.

    Returns:
        A list of :class:`SheetInfo` (``list``) or a single :class:`SheetInfo`
        (``get``).
    """

    if action == "list":
        return await sheets_list(workbook_id or "")
    if action == "get":
        return await sheet_get(workbook_id or "", name or "")
    raise ExcelVbaMcpError(
        f"Unknown sheet_read action {action!r}; expected 'list' or 'get'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "sheet_read", "action": action},
    )


__all__ = [
    "get_sheets_service",
    "set_sheets_service",
    "sheet",
    "sheet_add",
    "sheet_copy",
    "sheet_delete",
    "sheet_get",
    "sheet_read",
    "sheet_rename",
    "sheet_set_visibility",
    "sheets_list",
]
