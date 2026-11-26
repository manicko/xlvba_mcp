"""PivotTables / PivotCaches / PivotFields MCP tool contracts (task_040, task_048–task_051).

This module defines the **signatures, Pydantic DTOs and docstrings** for the
Phase 2 PivotTables surface. The read-only inventory tools
(``pivot_tables_list``, ``pivot_table_get``, ``pivot_cache_get``,
``pivot_fields_list``) are implemented by ``task_048``, the targeted refresh
tools (``pivot_table_refresh``, ``pivot_cache_refresh``) by ``task_049``, the
PivotField configuration tools (``pivot_field_set_layout``,
``pivot_field_set_position``, ``pivot_field_set_aggregation``,
``pivot_field_set_properties``) by ``task_050`` and the filter tools
(``pivot_field_get_filters``, ``pivot_field_apply_filter``,
``pivot_field_clear_filters``) by ``task_051``; all of them delegate to
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`.

It imports **no Excel/COM** — only domain DTOs; all Excel access happens on the
single-STA worker inside the pivot service/adapter.

Tool groups:

* Read-only (no snapshot/approval): ``pivot_tables_list``, ``pivot_table_get``,
  ``pivot_cache_get``, ``pivot_fields_list``, ``pivot_field_get_filters``.
* Mutating (writable + snapshot + approval + audit + postcondition):
  ``pivot_table_refresh``, ``pivot_cache_refresh`` (task_049),
  ``pivot_field_set_layout``, ``pivot_field_set_position``,
  ``pivot_field_set_aggregation``, ``pivot_field_set_properties`` (task_050),
  ``pivot_field_apply_filter``, ``pivot_field_clear_filters`` (task_051).

Only worksheet/Table-backed caches (``SourceType == database``) are supported;
external, OLAP and Data Model sources are reported as ``unsupported_source`` and
never refreshed or reconfigured. ``Workbook.RefreshAll`` is never used.
"""

from __future__ import annotations

from functools import partial

from excel_vba_mcp.api.job_tools import start_long_op
from excel_vba_mcp.application.job_registry import RunningJob
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
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
from excel_vba_mcp.domain.statuses import OperationClass, PivotAggregation, PivotFieldOrientation

# The active pivot tables service, published by the server lifespan. ``None``
# before startup and after shutdown; the read tools raise a clear error if used
# out of band (e.g. in a unit test that has not installed a fake).
_pivot_tables_service: PivotTablesService | None = None


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


def _require_non_empty(value: object, field: str) -> str:
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


def _require_orientation(value: object) -> PivotFieldOrientation:
    """Validate a target orientation (accepts the enum or its string name)."""

    if isinstance(value, PivotFieldOrientation):
        return value
    if isinstance(value, str):
        try:
            return PivotFieldOrientation(value.strip().lower())
        except ValueError as exc:
            raise ExcelVbaMcpError(
                f"Argument 'orientation' {value!r} is not a valid orientation; use "
                "one of 'hidden', 'row', 'column', 'page' or 'data'.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "orientation": value,
                    "accepted": [member.value for member in PivotFieldOrientation],
                    "reason": "invalid_orientation",
                },
            ) from exc
    raise ExcelVbaMcpError(
        f"Argument 'orientation' must be a PivotFieldOrientation or its string "
        f"name, got {value!r}.",
        code=ErrorCode.INVALID_RANGE,
        context={"orientation": repr(value), "reason": "invalid_orientation"},
    )


def _require_aggregation(value: object) -> PivotAggregation:
    """Validate an aggregation (accepts the enum or its string name).

    Legacy integer codes (``9``, ``-4114``, ``-4153``, ``-4154``) are refused
    here: the accepted vocabulary is the eleven :class:`PivotAggregation` members.
    """

    if isinstance(value, PivotAggregation):
        return value
    if isinstance(value, str):
        try:
            return PivotAggregation(value.strip().lower())
        except ValueError as exc:
            raise ExcelVbaMcpError(
                f"Argument 'function' {value!r} is not an accepted aggregation; use "
                f"one of {[member.value for member in PivotAggregation]}. Legacy "
                "codes (9, -4114, -4153, -4154) are rejected.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "function": value,
                    "accepted": [member.value for member in PivotAggregation],
                    "reason": "invalid_function",
                },
            ) from exc
    raise ExcelVbaMcpError(
        f"Argument 'function' must be a PivotAggregation or its string name, got "
        f"{value!r}. Legacy codes (9, -4114, -4153, -4154) are rejected.",
        code=ErrorCode.INVALID_RANGE,
        context={
            "function": repr(value),
            "accepted": [member.value for member in PivotAggregation],
            "reason": "invalid_function",
        },
    )


def _require_filter_spec(value: object) -> PivotFilterSpec:
    """Validate that a filter argument is a :class:`PivotFilterSpec`."""

    if not isinstance(value, PivotFilterSpec):
        raise ExcelVbaMcpError(
            "Argument 'spec' must be a PivotFilterSpec.",
            code=ErrorCode.INVALID_RANGE,
            context={"spec": repr(value), "reason": "invalid_filter_spec"},
        )
    return value


def _require_scope(value: object) -> str:
    """Validate a clear scope (``"field"`` or ``"table"``)."""

    if isinstance(value, str) and value.strip().lower() in {"field", "table"}:
        return value.strip().lower()
    raise ExcelVbaMcpError(
        f"Argument 'scope' must be 'field' or 'table', got {value!r}.",
        code=ErrorCode.INVALID_RANGE,
        context={
            "scope": repr(value),
            "accepted": ["field", "table"],
            "reason": "invalid_scope",
        },
    )


# ---------------------------------------------------------------------------
# Read-only tools (implemented: task_048)
# ---------------------------------------------------------------------------
async def pivot_tables_list(
    workbook_id: str,
    *,
    background: bool = False,
) -> PivotInventory | RunningJob:
    """List every PivotTable and its cache across all worksheets.

    Enumeration is worksheet-scoped (a full-workbook sweep) because
    ``Workbook.PivotTables`` is not a usable inventory on the target build.
    Shared caches are reported once with their affected tables.

    Args:
        workbook_id: The id returned by ``excel_open``.
        background: When ``True`` (Block F), schedule the sweep on the job registry
            and return ``{status: "running", job_id, poll_after_s}`` promptly
            instead of blocking; retrieve the :class:`PivotInventory` with
            ``job_get``. Default ``False`` (unchanged blocking behaviour).
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_pivot_tables_service()
    if background:
        return start_long_op(
            partial(service.list_pivot_tables, workbook_id),
            timeout=60.0,
            section="pivot",
            op_class=OperationClass.READ_ONLY,
        )
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

    Reports ``all_items_visible``, the visible and hidden ``PivotItems`` (the
    manual filter), the classified filter kind and any detected
    ``PivotFilters`` entries. Read-only; ``PivotFilters.Add`` is never called.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` (or a data-field caption).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    service = get_pivot_tables_service()
    return await service.get_pivot_field_filters(workbook_id, pivot_table, field)


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

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    service = get_pivot_tables_service()
    return await service.refresh_pivot_table(workbook_id, pivot_table)


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

    _require_non_empty(workbook_id, "workbook_id")
    _require_int(cache_index, "cache_index", minimum=1)
    service = get_pivot_tables_service()
    return await service.refresh_pivot_cache(workbook_id, cache_index)


async def pivot_field_set_layout(
    workbook_id: str,
    pivot_table: str,
    field: str,
    orientation: PivotFieldOrientation,
    *,
    delete: bool = False,
) -> PivotFieldInfo:
    """Set a field's orientation (hidden/row/column/page), or remove it.

    Moving a field into the data area is ``AddDataField``, not ``xlDataField``:
    ``orientation="data"`` is rejected here and the caller is pointed at
    ``pivot_field_set_aggregation``. Removal is explicit:

    * ``orientation="hidden"`` (default) removes the field from the layout
      **non-destructively** (``PivotField.Orientation = 0``); the field and its
      items/filter state stay in the PivotTable.
    * ``orientation="hidden"`` with ``delete=True`` removes it **destructively**
      (``PivotField.Delete``).

    Guarded mutation: writable session + snapshot + approval + postcondition.
    The returned :class:`PivotFieldInfo` is the post-mutation state re-read from
    the PivotTable (for a destructive removal, the last-known state).

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName``.
        orientation: The target orientation bucket (hidden/row/column/page).
        delete: ``True`` for the destructive ``PivotField.Delete`` removal;
            requires ``orientation="hidden"``. ``False`` (default) is the
            non-destructive ``Orientation = 0``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    normalized = _require_orientation(orientation)
    service = get_pivot_tables_service()
    return await service.set_pivot_field_layout(
        workbook_id, pivot_table, field, normalized, delete=delete
    )


async def pivot_field_set_position(
    workbook_id: str,
    pivot_table: str,
    field: str,
    position: int,
) -> PivotFieldInfo:
    """Set a field's 1-based position within its current orientation.

    Out-of-range positions and hidden (unplaced) fields are rejected before any
    COM mutation. The returned :class:`PivotFieldInfo` is the post-mutation state
    re-read from the PivotTable; the resulting orientation order is recorded in
    the audit entry. Guarded mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` (or a data field's caption).
        position: The new 1-based position within the orientation.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    _require_int(position, "position", minimum=1)
    service = get_pivot_tables_service()
    return await service.set_pivot_field_position(workbook_id, pivot_table, field, position)


async def pivot_field_set_aggregation(
    workbook_id: str,
    pivot_table: str,
    field: str,
    function: PivotAggregation,
    *,
    caption: str | None = None,
) -> PivotFieldInfo:
    """Add/update a data field's aggregation via ``AddDataField``/``Function``.

    Only the eleven accepted aggregation codes are allowed; legacy codes
    (``9``, ``-4114``, ``-4153``, ``-4154``) are rejected. ``caption`` must be
    unique when supplied (duplicates are rejected by Excel and pre-validated
    here). An existing data field on the same source field is updated in place;
    a source field without a data field is added with ``AddDataField``
    (assigning ``Orientation = xlDataField`` does not create one). Guarded
    mutation.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The source field ``SourceName``.
        function: The target aggregation function.
        caption: Optional unique data-field caption.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    _require_aggregation(function)
    if caption is not None:
        _require_non_empty(caption, "caption")
    service = get_pivot_tables_service()
    return await service.set_pivot_field_aggregation(
        workbook_id, pivot_table, field, function, caption=caption
    )


async def pivot_field_set_properties(
    workbook_id: str,
    pivot_table: str,
    field: str,
    properties: PivotFieldProperties,
) -> PivotFieldInfo:
    """Set bounded PivotField properties (caption/number format/show-all-items).

    Only the properties explicitly set on ``properties`` are written;
    ``caption`` / ``number_format`` target the data field (a source field without
    one is rejected) and ``show_all_items`` targets the source ``PivotField``.
    An empty patch, a blank caption/number format or a duplicate caption is
    rejected before any COM mutation. Guarded mutation: writable session +
    snapshot + approval + postcondition.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` or data-field caption.
        properties: The sparse property update (:class:`PivotFieldProperties`).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    if not isinstance(properties, PivotFieldProperties):
        raise ExcelVbaMcpError(
            "Argument 'properties' must be a PivotFieldProperties patch.",
            context={"properties": repr(properties)},
        )
    service = get_pivot_tables_service()
    return await service.set_pivot_field_properties(workbook_id, pivot_table, field, properties)


async def pivot_field_apply_filter(
    workbook_id: str,
    pivot_table: str,
    field: str,
    spec: PivotFilterSpec,
) -> PivotFieldFilters:
    """Apply a supported (manual, visible-items) filter to one field.

    Only a supported ``manual`` filter is applied (``PivotItems[i].Visible``).
    Label / value / date / top-bottom filters map to ``PivotFilters.Add``, which
    is unsupported on the target build: they return an explicit ``unsupported``
    result (``filter.supported = False`` with a reason) and are never silently
    converted into a manual filter. The request is validated before any mutation.
    Guarded mutation: writable session + snapshot + approval + postcondition.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field ``SourceName`` (or a data-field caption).
        spec: The filter request (:class:`PivotFilterSpec`).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    _require_non_empty(field, "field")
    normalized = _require_filter_spec(spec)
    service = get_pivot_tables_service()
    return await service.apply_pivot_field_filter(workbook_id, pivot_table, field, normalized)


async def pivot_field_clear_filters(
    workbook_id: str,
    pivot_table: str,
    field: str | None = None,
    *,
    scope: str = "field",
) -> PivotFieldFilters:
    """Clear filters for one field or the whole PivotTable.

    ``scope="field"`` (default; requires ``field``) calls
    ``PivotField.ClearAllFilters``; ``scope="table"`` calls
    ``PivotTable.ClearAllFilters`` and returns a table-level result whose
    ``field`` is empty. The scope and the field requirement are validated before
    any mutation. Guarded mutation: writable session + snapshot + approval +
    postcondition.

    Args:
        workbook_id: The id returned by ``excel_open``.
        pivot_table: The stable PivotTable ``Name``.
        field: The field to clear when ``scope == "field"``.
        scope: ``"field"`` or ``"table"``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(pivot_table, "pivot_table")
    normalized_scope = _require_scope(scope)
    if normalized_scope == "field":
        _require_non_empty(field, "field")
    service = get_pivot_tables_service()
    return await service.clear_pivot_field_filters(
        workbook_id, pivot_table, field, scope=normalized_scope
    )


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
