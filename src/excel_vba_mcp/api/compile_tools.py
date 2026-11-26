"""VBE compile MCP tools (task_020, spec §6.3 / §8.2).

Thin, async tool handlers for the compile surface: ``vba_compile_diagnostics`` and
``vba_get_compile_diagnostics``. They are deliberately **thin**: each validates
its arguments, delegates to
:class:`~excel_vba_mcp.application.compile_service.CompileService` and returns a
domain DTO. They **never touch COM** (ARCH-009).

On the validated Excel build (16.0.14332) non-activating compilation is
infeasible, so :func:`vba_compile_diagnostics` always returns
``status=COMPILE_INDETERMINATE`` with ``reason=NON_ACTIVATING_COMPILE_UNAVAILABLE``
and partial, explicitly non-compile diagnostics. It never returns ``compiled`` or
``compile_error``.

Like the other tool modules, the active service is published here via
:func:`set_compile_service` (called from the server lifespan) and read via
:func:`get_compile_service`.
"""

from __future__ import annotations

from excel_vba_mcp.application.compile_service import (
    CompileDiagnostics,
    CompileResult,
    CompileService,
)
from excel_vba_mcp.domain.errors import ExcelVbaMcpError

# The active compile service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_compile_service: CompileService | None = None


def set_compile_service(service: CompileService | None) -> None:
    """Publish (or clear) the active compile service used by the tools."""

    global _compile_service
    _compile_service = service


def get_compile_service() -> CompileService:
    """Return the active compile service or raise if it is not configured."""

    if _compile_service is None:
        raise ExcelVbaMcpError(
            "The compile service is not initialized.",
            context={"tool": "compile_tools"},
        )
    return _compile_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def vba_compile_diagnostics(workbook_id: str) -> CompileResult:
    """Return a fail-closed compile diagnostic result for an open workbook.

    The call never compiles and never returns ``compiled`` / ``compile_error``. On
    this build it returns ``status=COMPILE_INDETERMINATE`` with reason
    ``NON_ACTIVATING_COMPILE_UNAVAILABLE`` plus partial, non-intrusive
    diagnostics (``is_compile=false``): broken references, project protection,
    component inventory and VBOM availability.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        A :class:`CompileResult` with ``status=COMPILE_INDETERMINATE``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_compile_service()
    return await service.compile_project(workbook_id)


async def vba_get_compile_diagnostics(workbook_id: str) -> CompileDiagnostics:
    """Return the last compile diagnostics without forcing a compile re-run.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        The cached :class:`CompileDiagnostics` (``is_compile=false``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_compile_service()
    return await service.get_last_diagnostics(workbook_id)


__all__ = [
    "get_compile_service",
    "set_compile_service",
    "vba_compile_diagnostics",
    "vba_get_compile_diagnostics",
]
