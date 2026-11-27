"""Controls inventory + guarded ``OnAction`` MCP tools (task_030, spec §7.6 / §12.2-A10).

Thin, async tool handlers for the Form Control / ActiveX surface:
``controls_list``, ``control_get`` and ``control_set_on_action``. They are
deliberately **thin**: each validates its arguments, delegates to
:class:`~excel_vba_mcp.application.controls_service.ControlsService` and returns a
domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` / ``win32com``
import lives in this module; all Excel access happens on the single-STA worker
inside the service.

``control_set_on_action`` writes ``OnAction`` on a Form Control only — the
assigned macro is never executed and ActiveX controls (no ``OnAction``) are
rejected — and only succeeds when the service guard chain passes: writable
session, snapshot, approval (deny-by-default).

Like the other tool modules, the active service is published here via
:func:`set_controls_service` (called from the server lifespan) and read via
:func:`get_controls_service`.
"""

from __future__ import annotations

from typing import Literal

from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import ControlEntry

# The active controls service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_controls_service: ControlsService | None = None


def set_controls_service(service: ControlsService | None) -> None:
    """Publish (or clear) the active controls service used by the tools."""

    global _controls_service
    _controls_service = service


def get_controls_service() -> ControlsService:
    """Return the active controls service or raise if it is not configured."""

    if _controls_service is None:
        raise ExcelVbaMcpError(
            "The controls service is not initialized.",
            context={"tool": "control_tools"},
        )
    return _controls_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------
async def controls_list(workbook_id: str) -> list[ControlEntry]:
    """List every Form Control and ActiveX control of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`ControlEntry` per Form Control (``kind=form_control`` with
        ``control_type`` / ``caption`` / ``linked_cell`` / ``on_action``) and per
        ActiveX control (``kind=activex`` with ``prog_id`` and an **inferred,
        probable** ``probable_procedure`` / ``macro_mapping=probable``). ActiveX
        event mappings are best-effort guesses, never confirmed.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_controls_service()
    return await service.list_controls(workbook_id)


async def control_get(workbook_id: str, name: str) -> ControlEntry:
    """Return the :class:`ControlEntry` for one control by name.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The control name (case-insensitive).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_controls_service()
    return await service.get_control(workbook_id, name)


# ---------------------------------------------------------------------------
# Mutating tool (guarded)
# ---------------------------------------------------------------------------
async def control_set_on_action(workbook_id: str, name: str, macro_name: str) -> ControlEntry:
    """Assign ``OnAction`` to a Form Control (macro is NOT executed).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). An empty ``macro_name`` clears the assignment.
    ActiveX controls are rejected (they have no ``OnAction`` property).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The control name (case-insensitive).
        macro_name: The macro to assign (``""`` clears it).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_controls_service()
    return await service.set_on_action(workbook_id, name, macro_name)


# ---------------------------------------------------------------------------
# Flat-signature dispatcher (WP5, option (a))
# ---------------------------------------------------------------------------
async def control(
    action: Literal["list", "get", "set-on-action"],
    workbook_id: str | None = None,
    name: str | None = None,
    macro_name: str | None = None,
) -> list[ControlEntry] | ControlEntry:
    """Control dispatch: ``list``, ``get`` or ``set-on-action`` (WP5).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`controls_list` /
    :func:`control_get` / :func:`control_set_on_action` handlers unchanged and
    forwards only the validated arguments. Per-action required parameters are
    enforced by the target handler (``get`` / ``set-on-action`` both need
    ``name``; ``workbook_id`` is required for every action) so a missing
    required argument surfaces as a structured :class:`ExcelVbaMcpError` rather
    than a crash. ``macro_name`` is optional for ``set-on-action`` — an omitted
    value clears the assignment (the target handler's contract). An unknown
    ``action`` raises a structured :class:`ExcelVbaMcpError`.

    ``list`` and ``get`` are reads; ``set-on-action`` is a guarded mutation that
    writes ``OnAction`` on a Form Control only (ActiveX controls are rejected) —
    the assigned macro is never executed.

    Args:
        action: ``list`` (every control), ``get`` (one by name) or
            ``set-on-action`` (assign/clear ``OnAction``).
        workbook_id: The id returned by ``open``; required for every action.
        name: The control name (case-insensitive); required for ``get`` and
            ``set-on-action``.
        macro_name: The macro to assign; optional for ``set-on-action`` (``""``
            or omitted clears the assignment).

    Returns:
        A list of :class:`ControlEntry` (``list``) or a single
        :class:`ControlEntry` (``get`` / ``set-on-action``).
    """

    if action == "list":
        return await controls_list(workbook_id or "")
    if action == "get":
        return await control_get(workbook_id or "", name or "")
    if action == "set-on-action":
        return await control_set_on_action(workbook_id or "", name or "", macro_name or "")
    raise ExcelVbaMcpError(
        f"Unknown control action {action!r}; expected 'list', 'get' or 'set-on-action'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "control", "action": action},
    )


__all__ = [
    "control",
    "control_get",
    "control_set_on_action",
    "controls_list",
    "get_controls_service",
    "set_controls_service",
]
