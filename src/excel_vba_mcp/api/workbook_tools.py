"""Base workbook / session MCP tools (task_012, task_032).

These are the thin, async MCP tool handlers for the workbook lifecycle:
``excel_open``, ``excel_close``, ``excel_save``, ``excel_get_status``,
``excel_list_workbooks`` and ``workbook_get_model``. They are deliberately
**thin**: each one validates its arguments, delegates to an application service,
and returns a domain DTO. They **never touch COM** (ARCH-009): no
``pythoncom`` / ``win32com`` import lives in this module, and all Excel access
happens inside the services on the dedicated STA worker.

The handlers need configured services. Because a STDIO MCP server has a single
runtime owned by the server lifespan, the active services are published here via
``set_session_service`` / ``set_workbook_model_service`` (called from the
lifespan) and read via ``get_session_service`` / ``get_workbook_model_service``.
This keeps the public tool signatures exactly as specified while remaining
straightforward to unit-test with fake services.
"""

from __future__ import annotations

from functools import partial

from excel_vba_mcp.api.job_tools import start_long_op
from excel_vba_mcp.application.job_registry import RunningJob
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.workbook_model_service import WorkbookModelService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import SessionStatus, WorkbookInfo, WorkbookModel
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode

# The active session service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band
# (e.g. in a unit test that has not installed a fake).
_session_service: SessionService | None = None

# The active workbook model service, published by the server lifespan. ``None``
# before startup and after shutdown; ``workbook_get_model`` raises if unset.
_workbook_model_service: WorkbookModelService | None = None


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


def set_workbook_model_service(service: WorkbookModelService | None) -> None:
    """Publish (or clear) the active workbook model service used by the tools."""

    global _workbook_model_service
    _workbook_model_service = service


def get_workbook_model_service() -> WorkbookModelService:
    """Return the active workbook model service or raise if it is not configured."""

    if _workbook_model_service is None:
        raise ExcelVbaMcpError(
            "The workbook model service is not initialized.",
            context={"tool": "workbook_tools"},
        )
    return _workbook_model_service


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
    mode: SessionMode = SessionMode.EXECUTE,
    read_only: bool = False,
    *,
    background: bool = False,
) -> WorkbookInfo | RunningJob:
    """Open an allowlisted workbook as a hidden working copy.

    By default this creates a writable, execution-capable session
    (``mode=execute``, ``read_only=False``). Pass ``mode=inspect`` and/or
    ``read_only=True`` to opt into a macros-disabled or read-only session.

    Args:
        path: Path to the workbook; validated against the allowlist.
        mode: ``inspect`` (macros disabled) or ``execute`` (macros by UI).
            Default ``execute`` so a normal open is writable and
            execution-capable.
        read_only: Open the working copy read-only. Default ``False``.
        background: When ``True`` (Block F), schedule the open on the job registry
            and return ``{status: "running", job_id, poll_after_s}`` promptly
            instead of blocking; retrieve the :class:`WorkbookInfo` with
            ``job_get``. Default ``False`` (unchanged blocking behaviour).

    Returns:
        A :class:`WorkbookInfo` describing the opened working copy, or a
        :class:`RunningJob` handle when ``background=True``.
    """

    _require_non_empty(path, "path")
    session_mode = SessionMode(mode)
    service = get_session_service()
    if background:
        return start_long_op(
            partial(service.open_workbook, path, mode=session_mode, read_only=read_only),
            timeout=300.0,
            section="open",
            op_class=OperationClass.MUTATING,
        )
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


async def workbook_get_model(workbook_id: str, include_vba: bool = True) -> WorkbookModel:
    """Return the aggregated Stage 1 workbook model for an open workbook.

    The model aggregates sheets, tables, names, shapes, controls and — when
    ``include_vba`` is true — the VBA project, each with a section status. It
    never reads all-cells / Range data and never aborts because VBA is
    unavailable (``vba_access`` reports the VBA state instead).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        include_vba: When ``True`` (default) the VBA project is aggregated.

    Returns:
        A :class:`WorkbookModel` with a content ``model_hash``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_workbook_model_service()
    return await service.get_workbook_model(workbook_id, include_vba=include_vba)


__all__ = [
    "excel_close",
    "excel_get_status",
    "excel_list_workbooks",
    "excel_open",
    "excel_save",
    "get_session_service",
    "get_workbook_model_service",
    "set_session_service",
    "set_workbook_model_service",
    "workbook_get_model",
]
