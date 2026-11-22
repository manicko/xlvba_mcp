"""Policy-controlled macro execution MCP tools (task_022, spec §6.3 / §7.3).

Thin, async tool handlers for the macro surface: ``macro_list``, ``macro_run`` and
``macro_get_last_result``. They are deliberately **thin**: each validates its
arguments, delegates to :class:`~excel_vba_mcp.application.macro_service.MacroService`
and returns a domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` /
``win32com`` import lives in this module; all Excel/VBIDE access happens on the
single-STA worker inside the service.

``macro_run`` only runs a macro that is approved (deny-by-default), classified
``background_safe`` for the current workbook hash (BG06) and passes the
interactive-UI preflight (R13); a blocked macro is returned as a truthful
``MacroRunResult`` with ``status=blocked`` and is never dispatched.

Like the other tool modules, the active service is published here via
:func:`set_macro_service` (called from the server lifespan) and read via
:func:`get_macro_service`.
"""

from __future__ import annotations

from typing import Any

from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyClassification,
    MacroInfo,
    MacroService,
)
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import MacroRunResult

# The active macro service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_macro_service: MacroService | None = None


def set_macro_service(service: MacroService | None) -> None:
    """Publish (or clear) the active macro service used by the tools."""

    global _macro_service
    _macro_service = service


def get_macro_service() -> MacroService:
    """Return the active macro service or raise if it is not configured."""

    if _macro_service is None:
        raise ExcelVbaMcpError(
            "The macro service is not initialized.",
            context={"tool": "macro_tools"},
        )
    return _macro_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def macro_list(workbook_id: str) -> list[MacroInfo]:
    """List the callable macro candidates of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`MacroInfo` per public ``Sub`` / ``Function`` procedure, each
        annotated with its ``background_safe`` classification for the current
        workbook hash.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_macro_service()
    return await service.list_macros(workbook_id)


async def macro_run(
    workbook_id: str,
    macro_name: str,
    args: list[Any] | None = None,
) -> MacroRunResult:
    """Run one approved, background-safe macro of an open workbook.

    The macro is executed only when it is approved (deny-by-default), classified
    ``background_safe`` for the current workbook hash (BG06) and passes the
    interactive-UI preflight (R13). Otherwise a truthful result with
    ``status=blocked`` (or ``not_found``) is returned and the macro is **not**
    dispatched.

    Args:
        workbook_id: The id returned by :func:`excel_open` (execute-mode session).
        macro_name: The procedure to run, qualified as ``Module.Procedure`` or a
            bare ``Procedure`` name.
        args: Positional arguments (``str`` / ``int`` / ``float`` / ``bool`` /
            ``None`` / ``datetime`` and 1-D / 2-D lists of those).

    Returns:
        A :class:`MacroRunResult` with ``status`` one of ``success`` /
        ``not_found`` / ``runtime_error`` / ``blocked`` / ``timeout``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(macro_name, "macro_name")
    service = get_macro_service()
    return await service.run_macro(workbook_id, macro_name, args)


async def macro_classify_background(
    workbook_id: str,
    procedure: str,
    background_safe: bool,
    expected_project_source_hash: str | None = None,
) -> BackgroundSafetyClassification:
    """Classify a procedure's ``background_safe`` for the current source revision.

    Thin, UNGATED bookkeeping handler (approach A): it validates its arguments and
    delegates to :meth:`MacroService.classify_background`, returning the
    :class:`BackgroundSafetyClassification` DTO. It never touches COM and is not the
    security boundary — :func:`macro_run` keeps the deny-by-default approval gate and
    the deterministic R13 veto. The classification is bound to the workbook hash and
    the live VBA ``project_source_hash`` and is invalidated on source change.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        procedure: The procedure to classify (``Module.Procedure`` or a bare unique
            ``Procedure`` name).
        background_safe: The classification to record.
        expected_project_source_hash: Optional optimistic-concurrency guard; when it
            no longer matches the live source hash the write is rejected
            (``STALE_TARGET``).

    Returns:
        A :class:`BackgroundSafetyClassification` describing what was recorded.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(procedure, "procedure")
    service = get_macro_service()
    return await service.classify_background(
        workbook_id,
        procedure,
        background_safe,
        expected_project_source_hash=expected_project_source_hash,
    )


async def macro_get_last_result(workbook_id: str) -> MacroRunResult | None:
    """Return the cached last macro run result without re-running the macro.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        The cached :class:`MacroRunResult`, or ``None`` when no macro has been run.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_macro_service()
    return await service.get_last_result(workbook_id)


__all__ = [
    "get_macro_service",
    "macro_classify_background",
    "macro_get_last_result",
    "macro_list",
    "macro_run",
    "set_macro_service",
]
