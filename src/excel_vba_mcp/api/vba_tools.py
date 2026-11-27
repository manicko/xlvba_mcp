"""Read-only VBA project MCP tools (task_016, spec §6.2 / §7.2).

Thin, async MCP tool handlers for the read-only VBA project surface:
``vba_project_info``, ``vba_list_components``, ``vba_list_procedures``,
``vba_read_component``, ``vba_export_project``, ``vba_get_references`` and
``vba_search_code``. They are deliberately **thin**: each validates its
arguments, delegates to
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService` and
returns a domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` /
``win32com`` import lives in this module; all Excel/VBIDE access happens on the
single-STA worker inside the service.

Like :mod:`excel_vba_mcp.api.workbook_tools`, the active service is published
here via :func:`set_vba_project_service` (called from the server lifespan) and
read via :func:`get_vba_project_service`.
"""

from __future__ import annotations

from typing import Literal

from excel_vba_mcp.application.compile_service import CompileDiagnostics
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaExportManifest,
    VbaMutationResult,
    VbaProcedureInfo,
    VbaProjectInfo,
    VbaReferenceInfo,
    VbaSearchMatch,
)

# The active VBA project service, published by the server lifespan. ``None``
# before startup and after shutdown; the tools raise a clear error if used out of
# band (e.g. in a unit test that has not installed a fake).
_vba_project_service: VBAProjectService | None = None


def set_vba_project_service(service: VBAProjectService | None) -> None:
    """Publish (or clear) the active VBA project service used by the tools."""

    global _vba_project_service
    _vba_project_service = service


def get_vba_project_service() -> VBAProjectService:
    """Return the active VBA project service or raise if it is not configured."""

    if _vba_project_service is None:
        raise ExcelVbaMcpError(
            "The VBA project service is not initialized.",
            context={"tool": "vba_tools"},
        )
    return _vba_project_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def vba_project_info(workbook_id: str) -> VbaProjectInfo:
    """Return the VBA project summary (components, references, protection).

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        A :class:`VbaProjectInfo` with the component inventory and references.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_vba_project_service()
    return await service.get_project_info(workbook_id)


async def vba_list_components(workbook_id: str) -> list[VbaComponentInfo]:
    """List the VBA components of an open workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`VbaComponentInfo` per component (type, lines, source hash).
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_vba_project_service()
    return await service.list_components(workbook_id)


async def vba_list_procedures(workbook_id: str, component: str) -> list[VbaProcedureInfo]:
    """List the procedures of one VBA component.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        component: The ``VBComponent.Name`` to inspect.

    Returns:
        The procedures found via the CodeModule API (heuristic ones are flagged).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(component, "component")
    service = get_vba_project_service()
    return await service.list_procedures(workbook_id, component)


async def vba_read_component(workbook_id: str, component: str) -> VbaComponentSource:
    """Read the code module of one VBA component.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        component: The ``VBComponent.Name`` to read.

    Returns:
        A :class:`VbaComponentSource` with the code text and normalized hash.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(component, "component")
    service = get_vba_project_service()
    return await service.read_component(workbook_id, component)


async def vba_export_project(workbook_id: str, out_dir: str | None = None) -> VbaExportManifest:
    """Export every VBA component of an open workbook.

    Standard/Class modules export as ``.bas`` / ``.cls``; UserForms as
    ``.frm`` + ``.frx``; Document modules as ``.cls`` (never removed). When
    ``out_dir`` is omitted a temporary directory is used.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        out_dir: Optional destination directory inside an allowlisted artifact
            root.

    Returns:
        A :class:`VbaExportManifest` describing the exported files and hashes.
    """

    _require_non_empty(workbook_id, "workbook_id")
    if out_dir is not None:
        _require_non_empty(out_dir, "out_dir")
    service = get_vba_project_service()
    return await service.export_project(workbook_id, out_dir)


async def vba_get_references(workbook_id: str) -> list[VbaReferenceInfo]:
    """Return the VBA project references (broken ones flagged).

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`VbaReferenceInfo` per reference; ``is_broken`` marks missing
        references (also reported in ``VbaProjectInfo.missing_references``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_vba_project_service()
    return await service.get_references(workbook_id)


async def vba_search_code(
    workbook_id: str, pattern: str, regex: bool = False
) -> list[VbaSearchMatch]:
    """Search every component's code for a substring or regular expression.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        pattern: Substring (default) or regular expression (``regex=True``).
        regex: Treat ``pattern`` as a regular expression.

    Returns:
        A bounded list of :class:`VbaSearchMatch` hits.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pattern, "pattern")
    service = get_vba_project_service()
    return await service.search_code(workbook_id, pattern, regex=regex)


async def vba_read(
    action: Literal[
        "project-info",
        "list-components",
        "list-procedures",
        "read-component",
        "get-references",
        "search-code",
        "get-compile-diagnostics",
    ],
    workbook_id: str | None = None,
    component: str | None = None,
    pattern: str | None = None,
    regex: bool = False,
) -> (
    VbaProjectInfo
    | list[VbaComponentInfo]
    | list[VbaProcedureInfo]
    | VbaComponentSource
    | list[VbaReferenceInfo]
    | list[VbaSearchMatch]
    | CompileDiagnostics
):
    """VBA read/diagnostic dispatch (WP9).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`vba_project_info` /
    :func:`vba_list_components` / :func:`vba_list_procedures` /
    :func:`vba_read_component` / :func:`vba_get_references` /
    :func:`vba_search_code` handlers and the compile service's cached
    diagnostics unchanged, and forwards only the validated arguments. Per-action
    required parameters are enforced handler-side (``workbook_id`` is required
    for every action; ``list-procedures`` / ``read-component`` need
    ``component``; ``search-code`` needs ``pattern``) so a missing required
    argument surfaces as a structured :class:`ExcelVbaMcpError` rather than a
    crash. This tool is read-only.

    Args:
        action: The read/diagnostic verb to perform.
        workbook_id: The id returned by ``excel_open`` (required).
        component: The ``VBComponent.Name`` (``list-procedures`` /
            ``read-component``).
        pattern: Substring or regular expression (``search-code``).
        regex: Treat ``pattern`` as a regular expression (``search-code``).

    Returns:
        A :class:`VbaProjectInfo` (``project-info``), ``list[VbaComponentInfo]``
        (``list-components``), ``list[VbaProcedureInfo]`` (``list-procedures``),
        a :class:`VbaComponentSource` (``read-component``),
        ``list[VbaReferenceInfo]`` (``get-references``),
        ``list[VbaSearchMatch]`` (``search-code``) or a
        :class:`CompileDiagnostics` (``get-compile-diagnostics``).
    """

    if action == "project-info":
        return await vba_project_info(workbook_id or "")
    if action == "list-components":
        return await vba_list_components(workbook_id or "")
    if action == "list-procedures":
        return await vba_list_procedures(workbook_id or "", component or "")
    if action == "read-component":
        return await vba_read_component(workbook_id or "", component or "")
    if action == "get-references":
        return await vba_get_references(workbook_id or "")
    if action == "search-code":
        return await vba_search_code(workbook_id or "", pattern or "", regex=regex)
    if action == "get-compile-diagnostics":
        from excel_vba_mcp.api.compile_tools import get_compile_service

        _require_non_empty(workbook_id or "", "workbook_id")
        service = get_compile_service()
        return await service.get_last_diagnostics(workbook_id or "")
    raise ExcelVbaMcpError(
        f"Unknown vba_read action {action!r}; expected 'project-info', "
        "'list-components', 'list-procedures', 'read-component', "
        "'get-references', 'search-code' or 'get-compile-diagnostics'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "vba_read", "action": action},
    )


async def vba_replace_component_source(
    workbook_id: str,
    component: str,
    new_source: str,
    expected_source_hash: str,
) -> VbaMutationResult:
    """Replace the entire code module of a Standard/Class VBA component.

    Guarded mutation (spec §7.2): requires a writable session, a successful
    snapshot, a matching ``expected_source_hash`` (revision check) and an
    approval grant. Document modules and UserForms are rejected. Returns a
    :class:`VbaMutationResult` with before/after hashes and export paths.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        component: The ``VBComponent.Name`` to replace (Standard/Class only).
        new_source: The new VBA source text.
        expected_source_hash: The SHA-256 of the **current** normalized code
            body (stale-revision guard).

    Returns:
        A :class:`VbaMutationResult` describing the mutation.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(component, "component")
    _require_non_empty(expected_source_hash, "expected_source_hash")
    service = get_vba_project_service()
    return await service.replace_component_source(
        workbook_id,
        component,
        new_source,
        expected_source_hash=expected_source_hash,
    )


async def vba_insert_code(
    workbook_id: str,
    component: str,
    code: str,
    position: str,
    expected_source_hash: str,
) -> VbaMutationResult:
    """Insert ``code`` at a semantic position in a Standard/Class VBA component.

    ``position`` must be ``"declarations"``, ``"before_procedure:<name>"`` or
    ``"after_procedure:<name>"``. The line is resolved via the VBIDE CodeModule
    API; ambiguous text heuristics are rejected. Same guard chain as
    :func:`vba_replace_component_source`.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        component: The ``VBComponent.Name`` to insert into (Standard/Class only).
        code: The VBA code text to insert.
        position: Insertion position spec.
        expected_source_hash: The SHA-256 of the **current** normalized code
            body (stale-revision guard).

    Returns:
        A :class:`VbaMutationResult` describing the mutation.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(component, "component")
    _require_non_empty(position, "position")
    _require_non_empty(expected_source_hash, "expected_source_hash")
    service = get_vba_project_service()
    return await service.insert_code(
        workbook_id,
        component,
        code,
        position=position,
        expected_source_hash=expected_source_hash,
    )


__all__ = [
    "get_vba_project_service",
    "set_vba_project_service",
    "vba_export_project",
    "vba_get_references",
    "vba_insert_code",
    "vba_list_components",
    "vba_list_procedures",
    "vba_project_info",
    "vba_read",
    "vba_read_component",
    "vba_replace_component_source",
    "vba_search_code",
]
