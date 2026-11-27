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
from typing import Literal

from excel_vba_mcp.api.job_tools import start_long_op
from excel_vba_mcp.application.job_registry import RunningJob
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.workbook_model_service import WorkbookModelService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    SaveAsResult,
    SessionStatus,
    WorkbookInfo,
    WorkbookModel,
)
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


async def excel_close(
    workbook_id: str,
    save_changes: bool = False,
    discard_changes: bool = False,
) -> SessionStatus:
    """Close an open workbook (refuse-by-default when there are unsaved changes).

    The unsaved-change disposition is enforced in the application service
    (:meth:`SessionService.close_workbook`) using the real dirty signal
    (``WorkbookAdapter.is_dirty`` -> ``not wb.Saved``):

    * a clean workbook closes normally;
    * dirty + ``save_changes`` saves then closes;
    * dirty + ``discard_changes`` discards (closes without saving);
    * dirty + neither **refuses** with a structured :class:`ExcelVbaMcpError`;
    * ``save_changes`` and ``discard_changes`` together are rejected as
      contradictory.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        save_changes: Persist changes on close. Default ``False`` so an open is
            never silently saved.
        discard_changes: Discard unsaved changes on close. Default ``False``.
            Mutually exclusive with ``save_changes``.

    Returns:
        A :class:`SessionStatus` describing the session after close.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_session_service()
    return await service.close_workbook(
        workbook_id,
        save_changes=save_changes,
        discard_changes=discard_changes,
    )


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


async def excel_save_as(
    workbook_id: str,
    path: str,
    *,
    overwrite_existing: bool = False,
    overwrite_source: bool = False,
) -> SaveAsResult:
    """Save an open workbook's working copy to a (possibly new) target path.

    Thin, no-COM handler (U7; plan_05 §5.5): validates arguments and delegates to
    :meth:`SessionService.save_as_workbook`. The session stays bound to its
    working copy — this never rebinds the open workbook. Fail-closed: an existing
    target needs ``overwrite_existing=true``; the original source needs
    ``overwrite_source=true`` and an unchanged source hash; ``outcome="ok"`` only
    after the final size + SHA-256 verify.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        path: Destination path; validated against the allowlist (MUTATING).
        overwrite_existing: Replace an existing target (writes a ``.bak``).
        overwrite_source: Allow the target to be the original source itself
            (requires the source hash to still match; writes a ``.bak``).

    Returns:
        A :class:`SaveAsResult` (fail-closed ``outcome``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(path, "path")
    service = get_session_service()
    return await service.save_as_workbook(
        workbook_id,
        path,
        overwrite_existing=overwrite_existing,
        overwrite_source=overwrite_source,
    )


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


async def workbook(
    action: Literal["open", "close", "save"],
    path: str | None = None,
    workbook_id: str | None = None,
    mode: SessionMode = SessionMode.EXECUTE,
    read_only: bool = False,
    save_changes: bool = False,
    discard_changes: bool = False,
    *,
    background: bool = False,
) -> WorkbookInfo | SessionStatus | RunningJob:
    """Workbook lifecycle dispatch: ``open``, ``close`` or ``save`` (WP1).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`excel_open` /
    :func:`excel_close` / :func:`excel_save` handlers unchanged and forwards
    only the validated arguments. Per-action required parameters are enforced by
    the target handler (``open`` needs ``path``; ``close`` / ``save`` need
    ``workbook_id``) so a missing required argument surfaces as a structured
    :class:`ExcelVbaMcpError` rather than a crash. An unknown ``action`` raises
    a structured :class:`ExcelVbaMcpError`.

    Args:
        action: ``open`` (open a working copy), ``close`` (release it) or
            ``save`` (persist it in place).
        path: Workbook path; required for ``open``.
        workbook_id: The id returned by ``open``; required for ``close`` /
            ``save``.
        mode: ``execute`` (default) or ``inspect`` (macros disabled) for
            ``open``.
        read_only: Open the working copy read-only (for ``open``).
        save_changes: Save then close (for ``close``).
        discard_changes: Discard unsaved changes (for ``close``).
        background: Schedule the ``open`` on the job registry (for ``open``).

    Returns:
        A :class:`WorkbookInfo` (``open`` / ``save``), a :class:`SessionStatus`
        (``close``), or a :class:`RunningJob` when ``background=True``.
    """

    if action == "open":
        return await excel_open(
            path or "",
            mode=mode,
            read_only=read_only,
            background=background,
        )
    if action == "close":
        return await excel_close(
            workbook_id or "",
            save_changes=save_changes,
            discard_changes=discard_changes,
        )
    if action == "save":
        return await excel_save(workbook_id or "")
    raise ExcelVbaMcpError(
        f"Unknown workbook action {action!r}; expected 'open', 'close' or 'save'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "workbook", "action": action},
    )


async def workbook_read(
    action: Literal["status", "list", "model"],
    workbook_id: str | None = None,
    include_vba: bool = True,
) -> SessionStatus | list[WorkbookInfo] | WorkbookModel:
    """Workbook read dispatch: ``status``, ``list`` or ``model`` (WP1).

    Thin flat-signature dispatcher (option (a)): routes to the existing
    :func:`excel_get_status` / :func:`excel_list_workbooks` /
    :func:`workbook_get_model` handlers unchanged. ``status`` accepts an
    optional ``workbook_id``; ``model`` requires one (enforced by the target
    handler). An unknown ``action`` raises a structured
    :class:`ExcelVbaMcpError`.

    Args:
        action: ``status`` (session/workbook state), ``list`` (open workbooks)
            or ``model`` (aggregated Stage 1 model).
        workbook_id: The id returned by ``open``; required for ``model``,
            optional for ``status``.
        include_vba: Aggregate the VBA project into the model (``model`` only).

    Returns:
        A :class:`SessionStatus` (``status``), a list of :class:`WorkbookInfo`
        (``list``) or a :class:`WorkbookModel` (``model``).
    """

    if action == "status":
        return await excel_get_status(workbook_id)
    if action == "list":
        return await excel_list_workbooks()
    if action == "model":
        return await workbook_get_model(workbook_id or "", include_vba=include_vba)
    raise ExcelVbaMcpError(
        f"Unknown workbook_read action {action!r}; expected 'status', 'list' or 'model'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "workbook_read", "action": action},
    )


__all__ = [
    "excel_close",
    "excel_get_status",
    "excel_list_workbooks",
    "excel_open",
    "excel_save",
    "excel_save_as",
    "get_session_service",
    "get_workbook_model_service",
    "set_session_service",
    "set_workbook_model_service",
    "workbook",
    "workbook_get_model",
    "workbook_read",
]
