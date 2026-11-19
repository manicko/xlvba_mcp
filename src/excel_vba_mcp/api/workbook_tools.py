"""Base workbook / session MCP tools (task_012).

These are the thin, async MCP tool handlers for the workbook lifecycle:
``excel_open``, ``excel_close``, ``excel_save``, ``excel_get_status`` and
``excel_list_workbooks``. They are deliberately **thin**: each one validates its
arguments, delegates to :class:`~excel_vba_mcp.application.session_service.SessionService`,
and returns a domain DTO. They **never touch COM** (ARCH-009): no
``pythoncom`` / ``win32com`` import lives in this module, and all Excel access
happens inside the session service on the dedicated STA worker.

The handlers need a configured :class:`SessionService`. Because a STDIO MCP
server has a single runtime owned by the server lifespan, the active service is
published here via :func:`set_session_service` (called from the lifespan) and
read via :func:`get_session_service`. This keeps the public tool signatures
exactly as specified while remaining straightforward to unit-test with a fake
service.
"""

from __future__ import annotations

from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import SessionStatus, WorkbookInfo
from excel_vba_mcp.domain.statuses import SessionMode

# The active session service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band
# (e.g. in a unit test that has not installed a fake).
_session_service: SessionService | None = None


def set_session_service(service: SessionService | None) -> None:
    """Publish (or clear) the active session service used by the tools."""

    global _session_service
    _session_service = service


def get_session_service() -> SessionService:
    """Return the active session service or raise if it is not configured."""

    if _session_service is None:
        raise ExcelVbaMcpError(
            "The workbook session service is not initialized.",
            context={"tool": "workbook_tools"},
        )
    return _session_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def excel_open(
    path: str,
    mode: SessionMode = SessionMode.INSPECT,
    read_only: bool = False,
) -> WorkbookInfo:
    """Open an allowlisted workbook as a hidden working copy.

    Args:
        path: Path to the workbook; validated against the allowlist.
        mode: ``inspect`` (macros disabled) or ``execute`` (macros by UI).
        read_only: Open the working copy read-only.

    Returns:
        A :class:`WorkbookInfo` describing the opened working copy.
    """

    _require_non_empty(path, "path")
    session_mode = SessionMode(mode)
    service = get_session_service()
    return await service.open_workbook(path, mode=session_mode, read_only=read_only)


async def excel_close(workbook_id: str, save_changes: bool = False) -> SessionStatus:
    """Close an open workbook (default: no save).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        save_changes: Persist changes on close. Default ``False`` so an open is
            never silently saved.

    Returns:
        A :class:`SessionStatus` describing the session after close.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_session_service()
    return await service.close_workbook(workbook_id, save_changes=save_changes)


async def excel_save(workbook_id: str) -> WorkbookInfo:
    """Explicitly save an open workbook's working copy.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        A refreshed :class:`WorkbookInfo`.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_session_service()
    return await service.save_workbook(workbook_id)


async def excel_get_status(workbook_id: str | None = None) -> SessionStatus:
    """Return session / workbook status without touching COM.

    Args:
        workbook_id: When given, report that workbook's status; otherwise report
            session-level status.

    Returns:
        A :class:`SessionStatus`.
    """

    if workbook_id is not None:
        _require_non_empty(workbook_id, "workbook_id")
    service = get_session_service()
    return await service.get_status(workbook_id)


async def excel_list_workbooks() -> list[WorkbookInfo]:
    """List the currently open workbooks from session bookkeeping (no COM)."""

    service = get_session_service()
    return await service.list_workbooks()


__all__ = [
    "excel_close",
    "excel_get_status",
    "excel_list_workbooks",
    "excel_open",
    "excel_save",
    "get_session_service",
    "set_session_service",
]
