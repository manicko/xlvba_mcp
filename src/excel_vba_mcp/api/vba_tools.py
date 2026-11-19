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

from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaExportManifest,
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


__all__ = [
    "get_vba_project_service",
    "set_vba_project_service",
    "vba_export_project",
    "vba_get_references",
    "vba_list_components",
    "vba_list_procedures",
    "vba_project_info",
    "vba_read_component",
    "vba_search_code",
]
