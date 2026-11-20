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

from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
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


__all__ = [
    "get_shapes_service",
    "set_shapes_service",
    "shape_get",
    "shape_set_on_action",
    "shapes_list",
]
