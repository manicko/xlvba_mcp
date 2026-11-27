"""Shapes inventory + guarded ``OnAction`` MCP tools (task_030, spec §7.6 / §12.2-A10).

Thin, async tool handlers for the drawing-shape surface: ``shapes_list``,
``shape_get`` and ``shape_set_on_action``. They are deliberately **thin**: each
validates its arguments, delegates to
:class:`~excel_vba_mcp.application.shapes_service.ShapesService` and returns a
domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` / ``win32com``
import lives in this module; all Excel access happens on the single-STA worker
inside the service.

``shape_set_on_action`` writes ``Shape.OnAction`` only — the assigned macro is
never executed — and only succeeds when the service guard chain passes: writable
session, snapshot, approval (deny-by-default).

Like the other tool modules, the active service is published here via
:func:`set_shapes_service` (called from the server lifespan) and read via
:func:`get_shapes_service`.
"""

from __future__ import annotations

from typing import Literal

from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import ShapeInfo

# The active shapes service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_shapes_service: ShapesService | None = None


def set_shapes_service(service: ShapesService | None) -> None:
    """Publish (or clear) the active shapes service used by the tools."""

    global _shapes_service
    _shapes_service = service


def get_shapes_service() -> ShapesService:
    """Return the active shapes service or raise if it is not configured."""

    if _shapes_service is None:
        raise ExcelVbaMcpError(
            "The shapes service is not initialized.",
            context={"tool": "shape_tools"},
        )
    return _shapes_service


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
async def shapes_list(workbook_id: str) -> list[ShapeInfo]:
    """List every drawing shape of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`ShapeInfo` per ``Worksheet.Shapes`` item with name, kind
        (AutoShape / TextBox / Group / Picture / Line / Form Control / ActiveX /
        …), ``on_action`` (the assigned macro, read verbatim — never executed),
        visibility, position/size and best-effort text.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_shapes_service()
    return await service.list_shapes(workbook_id)


async def shape_get(workbook_id: str, name: str) -> ShapeInfo:
    """Return the :class:`ShapeInfo` for one shape by name.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The shape name (case-insensitive).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_shapes_service()
    return await service.get_shape(workbook_id, name)


# ---------------------------------------------------------------------------
# Mutating tool (guarded)
# ---------------------------------------------------------------------------
async def shape_set_on_action(workbook_id: str, name: str, macro_name: str) -> ShapeInfo:
    """Assign ``Shape.OnAction = macro_name`` (macro is NOT executed).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). An empty ``macro_name`` clears the assignment.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The shape name (case-insensitive).
        macro_name: The macro to assign (``""`` clears it).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    service = get_shapes_service()
    return await service.set_on_action(workbook_id, name, macro_name)


# ---------------------------------------------------------------------------
# Flat-signature dispatcher (WP5, option (a))
# ---------------------------------------------------------------------------
async def shape(
    action: Literal["list", "get", "set-on-action"],
    workbook_id: str | None = None,
    name: str | None = None,
    macro_name: str | None = None,
) -> list[ShapeInfo] | ShapeInfo:
    """Shape dispatch: ``list``, ``get`` or ``set-on-action`` (WP5).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`shapes_list` / :func:`shape_get`
    / :func:`shape_set_on_action` handlers unchanged and forwards only the
    validated arguments. Per-action required parameters are enforced by the
    target handler (``get`` / ``set-on-action`` both need ``name``;
    ``workbook_id`` is required for every action) so a missing required
    argument surfaces as a structured :class:`ExcelVbaMcpError` rather than a
    crash. ``macro_name`` is optional for ``set-on-action`` — an omitted value
    clears the assignment (the target handler's contract). An unknown
    ``action`` raises a structured :class:`ExcelVbaMcpError`.

    ``list`` and ``get`` are reads; ``set-on-action`` is a guarded mutation that
    writes ``Shape.OnAction`` only — the assigned macro is never executed.

    Args:
        action: ``list`` (every shape), ``get`` (one by name) or
            ``set-on-action`` (assign/clear ``OnAction``).
        workbook_id: The id returned by ``open``; required for every action.
        name: The shape name (case-insensitive); required for ``get`` and
            ``set-on-action``.
        macro_name: The macro to assign; optional for ``set-on-action`` (``""``
            or omitted clears the assignment).

    Returns:
        A list of :class:`ShapeInfo` (``list``) or a single :class:`ShapeInfo`
        (``get`` / ``set-on-action``).
    """

    if action == "list":
        return await shapes_list(workbook_id or "")
    if action == "get":
        return await shape_get(workbook_id or "", name or "")
    if action == "set-on-action":
        return await shape_set_on_action(workbook_id or "", name or "", macro_name or "")
    raise ExcelVbaMcpError(
        f"Unknown shape action {action!r}; expected 'list', 'get' or 'set-on-action'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "shape", "action": action},
    )


__all__ = [
    "get_shapes_service",
    "set_shapes_service",
    "shape",
    "shape_get",
    "shape_set_on_action",
    "shapes_list",
]
