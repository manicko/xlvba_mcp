"""PivotTables / PivotCaches / PivotFields MCP tool contracts (task_040).

This module defines the **signatures, Pydantic DTOs and docstrings** for the
Phase 2 PivotTables surface. It is a frozen contract only (task_040): the bodies
raise :class:`NotImplementedError` and are implemented by later tasks
(``task_048``–``task_051``). It imports **no Excel/COM** — only domain DTOs; all
Excel access happens on the single-STA worker inside the pivot service/adapter.

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

_NOT_IMPLEMENTED = (
    "PivotTables tool {name} is a frozen contract (task_040) and is implemented "
    "in a later task; it performs no Excel/COM access here."
)


def _not_implemented(name: str) -> NotImplementedError:
    """Return the canonical not-yet-implemented error for a contract stub."""

    return NotImplementedError(_NOT_IMPLEMENTED.format(name=name))


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------


async def pivot_tables_list(workbook_id: str) -> PivotInventory:
    """List every PivotTable and its cache across all worksheets.

    Enumeration is worksheet-scoped (a full-workbook sweep) because
    ``Workbook.PivotTables`` is not a usable inventory on the target build.
    Shared caches are reported once with their affected tables.

    Args:
        workbook_id: The id returned by ``excel_open``.
    """

    raise _not_implemented("pivot_tables_list")


async def pivot_table_get(workbook_id: str, pivot_table: str) -> PivotTableInfo:
    """Return one PivotTable's metadata by name.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
    """

    raise _not_implemented("pivot_table_get")


async def pivot_cache_get(workbook_id: str, cache_index: int) -> PivotCacheInfo:
    """Return one PivotCache's metadata by 1-based cache index.

    Args:
        workbook_id: The id returned by ``excel_open``.
        cache_index: The 1-based ``CacheIndex`` linking a table to its cache.
    """

    raise _not_implemented("pivot_cache_get")


async def pivot_fields_list(workbook_id: str, pivot_table: str) -> list[PivotFieldInfo]:
    """List a PivotTable's fields with layout, aggregation and filter state.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
    """

    raise _not_implemented("pivot_fields_list")


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
]
