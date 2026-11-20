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

from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
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


__all__ = [
    "control_get",
    "control_set_on_action",
    "controls_list",
    "get_controls_service",
    "set_controls_service",
]
