"""PivotTables / PivotCaches / PivotFields MCP tool contracts (task_040, task_048).

This module defines the **signatures, Pydantic DTOs and docstrings** for the
Phase 2 PivotTables surface. The read-only inventory tools
(``pivot_tables_list``, ``pivot_table_get``, ``pivot_cache_get``,
``pivot_fields_list``) are implemented by ``task_048`` and delegate to
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`; the
mutating tools (refresh / layout / aggregation / filters) remain frozen contract
stubs implemented by ``task_049``–``task_051`` and raise
:class:`NotImplementedError`.

It imports **no Excel/COM** — only domain DTOs; all Excel access happens on the
single-STA worker inside the pivot service/adapter.

Tool groups:

* Read-only (no snapshot/approval): ``pivot_tables_list``, ``pivot_table_get``,
  ``pivot_cache_get``, ``pivot_fields_list``, ``pivot_field_get_filters``.
* Mutating (snapshot + approval + audit + postcondition): ``pivot_table_refresh``,
  ``pivot_cache_refresh``, ``pivot_field_set_layout``,
  ``pivot_field_set_position``, ``pivot_field_set_aggregation``,
  ``pivot_field_set_properties``, ``pivot_field_apply_filter``,
  ``pivot_field_clear_filters``.

Only worksheet/Table-backed caches (``SourceType == database``) are supported;
external, OLAP and Data Model sources are reported as ``unsupported_source`` and
never refreshed. ``Workbook.RefreshAll`` is never used.
"""

from __future__ import annotations

from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFieldProperties,
    PivotFilterSpec,
    PivotInventory,
    PivotRefreshResult,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import PivotAggregation, PivotFieldOrientation

# The active pivot tables service, published by the server lifespan. ``None``
# before startup and after shutdown; the read tools raise a clear error if used
# out of band (e.g. in a unit test that has not installed a fake).
_pivot_tables_service: PivotTablesService | None = None

_NOT_IMPLEMENTED = (
    "PivotTables tool {name} is a frozen contract (task_040) and is implemented "
    "in a later task; it performs no Excel/COM access here."
)


def _not_implemented(name: str) -> NotImplementedError:
    """Return the canonical not-yet-implemented error for a contract stub."""

    return NotImplementedError(_NOT_IMPLEMENTED.format(name=name))


def set_pivot_tables_service(service: PivotTablesService | None) -> None:
    """Publish (or clear) the active pivot tables service used by the tools."""

    global _pivot_tables_service
    _pivot_tables_service = service


def get_pivot_tables_service() -> PivotTablesService:
    """Return the active pivot tables service or raise if it is not configured."""

    if _pivot_tables_service is None:
        raise ExcelVbaMcpError(
            "The pivot tables service is not initialized.",
            context={"tool": "pivot_tools"},
        )
    return _pivot_tables_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


def _require_int(value: object, field: str, *, minimum: int | None = None) -> int:
    """Validate that an argument is an integer (optionally with a minimum)."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be an integer.",
            context={"field": field},
        )
    if minimum is not None and value < minimum:
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be >= {minimum}.",
            context={"field": field, "value": value},
        )
    return value


# ---------------------------------------------------------------------------
# Read-only tools (implemented: task_048)
# ---------------------------------------------------------------------------
async def pivot_tables_list(workbook_id: str) -> PivotInventory:
    """List every PivotTable and its cache across all worksheets.

    Enumeration is worksheet-scoped (a full-workbook sweep) because
    ``Workbook.PivotTables`` is not a usable inventory on the target build.
    Shared caches are reported once with their affected tables.

    Args:
        workbook_id: The id returned by ``excel_open``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_pivot_tables_service()
    return await service.list_pivot_tables(workbook_id)


async def pivot_table_get(workbook_id: str, pivot_table: str) -> PivotTableInfo:
    """Return one PivotTable's metadata by name.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    service = get_pivot_tables_service()
    return await service.get_pivot_table(workbook_id, pivot_table)


async def pivot_cache_get(workbook_id: str, cache_index: int) -> PivotCacheInfo:
    """Return one PivotCache's metadata by 1-based cache index.

    Args:
        workbook_id: The id returned by ``excel_open``.
        cache_index: The 1-based ``CacheIndex`` linking a table to its cache.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_int(cache_index, "cache_index", minimum=1)
    service = get_pivot_tables_service()
    return await service.get_pivot_cache(workbook_id, cache_index)


async def pivot_fields_list(workbook_id: str, pivot_table: str) -> list[PivotFieldInfo]:
    """List a PivotTable's fields with layout, aggregation and filter state.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    service = get_pivot_tables_service()
    return await service.list_pivot_fields(workbook_id, pivot_table)


async def pivot_field_get_filters(
    workbook_id: str, pivot_table: str, field: str
) -> PivotFieldFilters:
    """Read a PivotField's current filter state (manual visible/hidden items).

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` or data-field caption.
    """

    raise _not_implemented("pivot_field_get_filters")


# ---------------------------------------------------------------------------
# Mutating tools (guarded: writable session + snapshot + approval + audit)
# ---------------------------------------------------------------------------
async def pivot_table_refresh(workbook_id: str, pivot_table: str) -> PivotRefreshResult:
    """Refresh one PivotTable via ``PivotTable.RefreshTable`` (table-scoped).

    The cache source type is checked before invocation; unsupported sources
    return ``unsupported_source`` and are not refreshed. The result reports every
    table sharing the cache. Guarded mutation: writable session + snapshot +
    approval + audit.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
    """

    raise _not_implemented("pivot_table_refresh")


async def pivot_cache_refresh(workbook_id: str, cache_index: int) -> PivotRefreshResult:
    """Refresh one PivotCache via ``PivotCache.Refresh`` (cache-scoped).

    A cache is shared (N:1 with PivotTables), so this affects every table using
    the cache; ``affected_pivot_tables`` reports them before/after. Only a
    database/Table-backed cache is refreshed; external/OLAP/Data Model return
    ``unsupported_source``. Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        cache_index: The 1-based ``CacheIndex``.
    """

    raise _not_implemented("pivot_cache_refresh")


async def pivot_field_set_layout(
    workbook_id: str,
    pivot_table: str,
    field: str,
    orientation: PivotFieldOrientation,
) -> PivotFieldInfo:
    """Set a field's orientation (hidden/row/column/page).

    Moving a field into the data area is ``AddDataField``, not ``xlDataField``;
    only the layout buckets (hidden/row/column/page) are set through orientation.
    Guarded mutation: writable session + snapshot + approval + postcondition.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName``.
        orientation: The target orientation bucket.
    """

    raise _not_implemented("pivot_field_set_layout")


async def pivot_field_set_position(
    workbook_id: str,
    pivot_table: str,
    field: str,
    position: int,
) -> PivotFieldInfo:
    """Set a field's 1-based position within its current orientation.

    Out-of-range positions are rejected before mutation. Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName``.
        position: The new 1-based position within the orientation.
    """

    raise _not_implemented("pivot_field_set_position")


async def pivot_field_set_aggregation(
    workbook_id: str,
    pivot_table: str,
    field: str,
    function: PivotAggregation,
    *,
    caption: str | None = None,
) -> PivotFieldInfo:
    """Add/update a data field's aggregation via ``AddDataField``/``Function``.

    ``caption`` must be unique when supplied (duplicates are rejected by Excel).
    Only the eleven accepted aggregation codes are allowed. Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The source field ``SourceName``.
        function: The target aggregation function.
        caption: Optional unique data-field caption.
    """

    raise _not_implemented("pivot_field_set_aggregation")


async def pivot_field_set_properties(
    workbook_id: str,
    pivot_table: str,
    field: str,
    properties: PivotFieldProperties,
) -> PivotFieldInfo:
    """Set bounded PivotField properties (caption/number format/show-all-items).

    Guarded mutation: writable session + snapshot + approval + postcondition.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` or data-field caption.
        properties: The sparse property update (:class:`PivotFieldProperties`).
    """

    raise _not_implemented("pivot_field_set_properties")


async def pivot_field_apply_filter(
    workbook_id: str,
    pivot_table: str,
    field: str,
    spec: PivotFilterSpec,
) -> PivotFieldFilters:
    """Apply a supported (manual, visible-items) filter to one field.

    Label/value/date/top-bottom filters are unsupported on the target build and
    return an explicit ``unsupported`` result rather than being silently
    converted. Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName``.
        spec: The filter request (:class:`PivotFilterSpec`).
    """

    raise _not_implemented("pivot_field_apply_filter")


async def pivot_field_clear_filters(
    workbook_id: str,
    pivot_table: str,
    field: str | None = None,
    *,
    scope: str = "field",
) -> PivotFieldFilters:
    """Clear filters for one field or the whole PivotTable.

    ``scope`` is ``"field"`` (default; requires ``field``) or ``"table"``
    (clears all filters of the PivotTable). Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field to clear when ``scope == "field"``.
        scope: ``"field"`` or ``"table"``.
    """

    raise _not_implemented("pivot_field_clear_filters")


__all__ = [
    "get_pivot_tables_service",
    "pivot_cache_get",
    "pivot_cache_refresh",
    "pivot_field_apply_filter",
    "pivot_field_clear_filters",
    "pivot_field_get_filters",
    "pivot_field_set_aggregation",
    "pivot_field_set_layout",
    "pivot_field_set_position",
    "pivot_field_set_properties",
    "pivot_fields_list",
    "pivot_table_get",
    "pivot_table_refresh",
    "pivot_tables_list",
    "set_pivot_tables_service",
]
