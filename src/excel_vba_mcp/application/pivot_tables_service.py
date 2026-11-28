"""Excel PivotTables / PivotCaches / PivotFields service (task_048, task_049, task_050).

:class:`PivotTablesService` is the async application-layer orchestration for the
Phase 2 PivotTables surface. The read-only inventory tools
(``pivot_tables_list``, ``pivot_table_get``, ``pivot_cache_get``,
``pivot_fields_list``) are implemented by ``task_048``; the targeted refresh
tools (``pivot_table_refresh``, ``pivot_cache_refresh``) by ``task_049``; the
PivotField configuration tools (``set_pivot_field_layout``,
``set_pivot_field_position``, ``set_pivot_field_aggregation``,
``set_pivot_field_properties``) by ``task_050``; and the filter tools
(``get_pivot_field_filters``, ``apply_pivot_field_filter``,
``clear_pivot_field_filters``) by ``task_051``. It is deliberately **COM-free**:
every method resolves the ``workbook_id`` through the session bookkeeping and
then submits a single STA-local callable to the dedicated single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only domain
DTOs / plain containers cross the worker boundary (ARCH-009).

Read-only path (task_048, task_051 read): no snapshot / approval guard chain is
required for ``get_pivot_field_filters`` (it never mutates a PivotTable,
PivotCache or PivotField). Every call is recorded in the append-only audit trail
as ``read_only``.

Mutation path (task_049, task_050, task_051): ``refresh_pivot_table`` /
``refresh_pivot_cache``, the four ``set_pivot_field_*`` tools,
``apply_pivot_field_filter`` and ``clear_pivot_field_filters`` run the full
guard chain — writable working copy, successful snapshot, deny-by-default
approval — then a single mutating STA dispatch, recorded as ``mutating``. Only a
worksheet/Table-backed (``database``) cache is refreshed, reconfigured or
filtered; external / OLAP / Data Model sources are rejected with
``UNSUPPORTED_PIVOT_SOURCE`` and never touched. ``Workbook.RefreshAll`` is never
used.

Design (task_039 decision record, ``docs/research/pivot_tables_api.md``):

* **Worksheet sweep.** The adapter enumerates PivotTables per worksheet
  (``ws.PivotTables()``) because ``Workbook.PivotTables`` is unusable on the
  target build. The service returns the full inventory sorted by
  ``(sheet_index, pivot_table_name)``.
* **Cache identity.** A cache is identified by its 1-based ``CacheIndex`` and
  resolved through the PivotTable that references it. Shared caches are reported
  once with ``affected_pivot_tables`` listing every table on that cache.
* **Field identity.** ``list_pivot_fields`` resolves a PivotTable by its stable
  ``Name`` and returns each field's source name, caption, orientation, 1-based
  position within its orientation, data-field aggregation and manual filter
  state.
* **Refresh.** ``refresh_pivot_table`` → ``PivotTable.RefreshTable`` (returns a
  Boolean truth signal); ``refresh_pivot_cache`` → ``PivotCache.Refresh`` (no
  Boolean, confirmed by an advanced ``RefreshDate``). Both report every table
  sharing the cache and are confirmed by Excel's actual result state, not by the
  absence of a COM exception (task_039 §4).
* **Unsupported sources.** External / consolidation / scenario / pivot-table /
  OLAP / Data Model sources return ``unsupported_source`` and are never refreshed.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from typing import Any, TypeVar, cast

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import (
    SHEET_STATUS_CAPPED,
    SHEET_STATUS_OK,
    SHEET_STATUS_TOO_LARGE,
    PivotTablesAdapter,
    normalize_filter_scope,
    normalize_filter_spec,
    unsupported_filter_reason,
    validate_field_properties,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import (
    AmbiguousFilterError,
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    StaleCursorError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFieldProperties,
    PivotFilterSpec,
    PivotFilterState,
    PivotInventory,
    PivotRefreshResult,
    PivotResultPage,
    PivotTableInfo,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotFilterRequestKind,
)
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger

#: Default per-call soft ceiling (seconds) for a pivot inventory/refresh/config
#: call on the STA worker. Inventory reads are bounded and fast; a targeted
#: refresh of a local cache is also bounded, so the same ceiling as the other
#: services is used.
_DEFAULT_TIMEOUT_S = 60.0

#: Approval action names (stable identities for the pivot mutation surface).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_REFRESH_CACHE = "pivot_cache_refresh"
_ACTION_SET_LAYOUT = "pivot_field_set_layout"
_ACTION_SET_POSITION = "pivot_field_set_position"
_ACTION_SET_AGGREGATION = "pivot_field_set_aggregation"
_ACTION_SET_PROPERTIES = "pivot_field_set_properties"
_ACTION_APPLY_FILTER = "pivot_field_apply_filter"
_ACTION_CLEAR_FILTERS = "pivot_field_clear_filters"

#: The orientation buckets a field's ``Orientation`` may be assigned to.
_LAYOUT_ORIENTATIONS: frozenset[PivotFieldOrientation] = frozenset(
    {
        PivotFieldOrientation.HIDDEN,
        PivotFieldOrientation.ROW,
        PivotFieldOrientation.COLUMN,
        PivotFieldOrientation.PAGE,
    }
)

#: Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class PivotTablesService:
    """Async PivotTables/PivotCaches/PivotFields inventory + refresh (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        audit: Append-only JSONL audit trail.
        snapshot_service: Pre-mutation snapshot gate for the refresh tools.
            ``None`` in a read-only rig; the refresh tools then fail closed with
            ``SNAPSHOT_FAILED`` (deny-by-default).
        approval: Deny-by-default approval gate for the refresh tools. ``None``
            in a read-only rig; the refresh tools then fail closed.
        large_data: Unified large-data limits (single source). Defaults to
            :class:`~excel_vba_mcp.config.LargeDataConfig`; the server passes
            ``config.large_data``. ``max_pivot_result_cells`` bounds one
            rendered-result page.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        audit: AuditLogger,
        snapshot_service: SnapshotService | None = None,
        approval: ApprovalPolicy | None = None,
        large_data: LargeDataConfig | None = None,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._audit = audit
        self._snapshot = snapshot_service
        self._approval = approval
        self._large_data = large_data if large_data is not None else LargeDataConfig()
        self._adapter = PivotTablesAdapter()

    # ------------------------------------------------------------------
    # Public API — read-only (async; no COM here)
    # ------------------------------------------------------------------
    async def list_pivot_tables(self, workbook_id: str) -> PivotInventory:
        """Return every PivotTable and its cache across all worksheets.

        Enumeration is a **single worksheet sweep** (Block B): one pass feeds both
        the PivotTable inventory and the PivotCache inventory (linked by a
        ``CacheIndex`` memo), replacing the former ``2 + N`` re-enumeration. The
        source type is gated first per PivotTable, so an external / OLAP / Data
        Model cache is reported with its identity only. A per-sheet **measured**
        budget (field / ``PivotItems`` reads, never file size / row count /
        ``UsedRange``) stops the sweep between ops and reports the sheet as
        ``capped`` (partial) or ``too_large`` (named) in the audit record.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_tables_list",
            lambda wb: self._adapter.inventory_single_sweep(wb),
            record,
        )
        inventory = PivotInventory(
            workbook_id=workbook_id,
            pivot_tables=result["pivot_tables"],
            caches=result["caches"],
        )
        sheet_statuses = result["sheet_statuses"]
        self._record_audit(
            "pivot_tables_list",
            "read_only",
            record,
            {"workbook_id": workbook_id},
            {
                "pivot_table_count": len(inventory.pivot_tables),
                "cache_count": len(inventory.caches),
                "status": _overall_sheet_status(sheet_statuses),
                "sheets": sheet_statuses,
            },
        )
        return inventory

    async def get_pivot_table(self, workbook_id: str, pivot_table: str) -> PivotTableInfo:
        """Return one PivotTable's metadata by name.

        The full inventory is read and filtered by the stable PivotTable
        ``Name``; a missing name raises a structured error.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_table_get",
            lambda wb: self._find_pivot_table(wb, pivot_table),
            record,
        )
        self._record_audit(
            "pivot_table_get",
            "read_only",
            record,
            {"workbook_id": workbook_id, "pivot_table": pivot_table},
            {
                "pivot_table": result.name,
                "sheet": result.sheet_name,
                "cache_index": result.cache_index,
                "source_status": result.source_status,
            },
        )
        return result

    async def get_pivot_cache(self, workbook_id: str, cache_index: int) -> PivotCacheInfo:
        """Return one PivotCache's metadata by 1-based cache index.

        The cache is resolved through the PivotTable that references it (the
        ``PivotCaches`` collection is unusable for enumeration). Shared caches
        report ``affected_pivot_tables``.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_cache_get",
            lambda wb: self._adapter.read_pivot_cache(wb, cache_index),
            record,
        )
        self._record_audit(
            "pivot_cache_get",
            "read_only",
            record,
            {"workbook_id": workbook_id, "cache_index": cache_index},
            {
                "cache_index": result.cache_index,
                "source_type": result.source_type.value,
                "shared": result.shared_cache,
                "affected_pivot_tables": result.affected_pivot_tables,
            },
        )
        return result

    async def list_pivot_fields(self, workbook_id: str, pivot_table: str) -> list[PivotFieldInfo]:
        """List a PivotTable's fields with layout, aggregation and filter state.

        ``pivot_table`` is the stable PivotTable ``Name``; a missing name raises
        a structured error.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_fields_list",
            lambda wb: self._adapter.enumerate_pivot_fields(wb, pivot_table),
            record,
        )
        self._record_audit(
            "pivot_fields_list",
            "read_only",
            record,
            {"workbook_id": workbook_id, "pivot_table": pivot_table},
            {
                "pivot_table": pivot_table,
                "field_count": len(result),
            },
        )
        return result

    async def get_pivot_field_filters(
        self, workbook_id: str, pivot_table: str, field: str
    ) -> PivotFieldFilters:
        """Read one field's filter state (manual visible/hidden items).

        Read-only: the adapter reports ``AllItemsVisible``, the visible/hidden
        ``PivotItems`` (the manual filter), the classified filter kind and any
        detected ``PivotFilters`` entries. ``PivotFilters.Add`` is never called.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_field_get_filters",
            lambda wb: self._adapter.get_field_filters(wb, pivot_table, field),
            record,
        )
        self._record_audit(
            "pivot_field_get_filters",
            "read_only",
            record,
            {"workbook_id": workbook_id, "pivot_table": pivot_table, "field": field},
            {
                "pivot_table": result.pivot_table,
                "field": result.field,
                "all_items_visible": result.all_items_visible,
                "visible_count": len(result.visible_items),
                "hidden_items": list(result.hidden_items),
                "kind": result.filter.kind.value,
                "unsupported": list(result.unsupported),
            },
        )
        return result

    async def read_pivot_result(
        self,
        workbook_id: str,
        pivot_table: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
        cursor_fingerprint: str | None = None,
    ) -> PivotResultPage:
        """Read a bounded page of a PivotTable's rendered output (read-only).

        The displayed ``TableRange1`` cells are read through the shared bounded
        Range reader on the STA worker, for every source type (database,
        external, OLAP / Data Model — rendered cells are read, never the external
        source). A large PivotTable is paged through successive ``offset`` reads;
        the returned :class:`~excel_vba_mcp.domain.models.PageInfo` reports
        ``has_more`` / ``complete`` / ``total`` and the workbook content
        ``version``. Pass ``cursor_version`` (the prior page's ``page.version``)
        to fail with ``STALE_CURSOR`` when the workbook changed between pages.

        Pass ``cursor_fingerprint`` (the prior page's ``fingerprint``) to also
        fail with ``STALE_CURSOR`` when the PivotTable's structure changed
        without a version bump (an out-of-band refresh that does not go through a
        mutating service call, contract §3.2). The version check is preserved and
        runs first; the fingerprint check only fires when a fingerprint cursor is
        supplied (the first page has none).
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_result_read",
            lambda wb: self._adapter.read_pivot_result(
                wb,
                pivot_table,
                offset=offset,
                limit=limit,
                version_getter=lambda: self._session.get_version(workbook_id),
                max_cells=self._large_data.max_pivot_result_cells,
            ),
            record,
        )
        if cursor_version is not None and result.page.version != cursor_version:
            raise StaleCursorError(
                "The workbook changed since the first page; restart paging.",
                context={
                    "expected_version": cursor_version,
                    "actual_version": result.page.version,
                    "kind": "pivot_result",
                },
            )
        if cursor_fingerprint is not None and result.fingerprint != cursor_fingerprint:
            raise StaleCursorError(
                "The PivotTable structure changed since the first page; restart paging.",
                context={
                    "expected_fingerprint": cursor_fingerprint,
                    "actual_fingerprint": result.fingerprint,
                    "kind": "pivot",
                },
            )
        self._record_audit(
            "pivot_result_read",
            "read_only",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "offset": offset,
                "limit": limit,
            },
            {
                "pivot_table": result.pivot_table,
                "sheet": result.sheet_name,
                "source_type": result.source_type.value,
                "rows": result.rows,
                "cols": result.cols,
                "returned": result.page.returned,
                "has_more": result.page.has_more,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — mutations (guarded; task_049)
    # ------------------------------------------------------------------
    async def refresh_pivot_table(self, workbook_id: str, pivot_table: str) -> PivotRefreshResult:
        """Refresh one PivotTable via ``PivotTable.RefreshTable`` (guarded).

        The guard chain (writable working copy → successful snapshot → approval)
        runs before any STA dispatch. The adapter checks the cache source type
        before invocation and returns ``unsupported_source`` for external / OLAP /
        Data Model caches; success is confirmed by Excel's truth signal (the
        ``RefreshTable`` Boolean and/or an advanced ``RefreshDate``).
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_REFRESH_TABLE)

        result = await self._submit_mutation(
            workbook_id,
            "pivot_table_refresh",
            lambda wb: self._adapter.refresh_pivot_table(wb, pivot_table),
            record,
        )
        self._record_audit(
            "pivot_table_refresh",
            "mutating",
            record,
            {"workbook_id": workbook_id, "pivot_table": pivot_table},
            {
                "pivot_table": result.pivot_table,
                "cache_index": result.cache_index,
                "method": result.method.value,
                "status": result.status.value,
                "refreshed": result.refreshed,
                "refresh_return": result.refresh_return,
                "affected_pivot_tables": result.affected_pivot_tables,
                "duration_ms": result.duration_ms,
            },
        )
        return result

    async def refresh_pivot_cache(self, workbook_id: str, cache_index: int) -> PivotRefreshResult:
        """Refresh one PivotCache via ``PivotCache.Refresh`` (guarded).

        A cache is shared N:1 with PivotTables, so ``affected_pivot_tables``
        reports every table on the cache. The adapter checks the source type
        before invocation and returns ``unsupported_source`` for external / OLAP /
        Data Model caches; ``PivotCache.Refresh`` returns no Boolean, so success
        is confirmed by the advanced ``RefreshDate``.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_REFRESH_CACHE)

        result = await self._submit_mutation(
            workbook_id,
            "pivot_cache_refresh",
            lambda wb: self._adapter.refresh_pivot_cache(wb, cache_index),
            record,
        )
        self._record_audit(
            "pivot_cache_refresh",
            "mutating",
            record,
            {"workbook_id": workbook_id, "cache_index": cache_index},
            {
                "pivot_table": result.pivot_table,
                "cache_index": result.cache_index,
                "method": result.method.value,
                "status": result.status.value,
                "refreshed": result.refreshed,
                "affected_pivot_tables": result.affected_pivot_tables,
                "duration_ms": result.duration_ms,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — PivotField configuration (guarded; task_050)
    # ------------------------------------------------------------------
    async def set_pivot_field_layout(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        orientation: PivotFieldOrientation,
        *,
        delete: bool = False,
    ) -> PivotFieldInfo:
        """Set one field's layout bucket, or remove it (guarded).

        ``orientation`` must be one of ``hidden`` / ``row`` / ``column`` /
        ``page``; ``data`` is rejected before the STA dispatch because assigning
        ``Orientation = xlDataField`` never creates a data field — that is
        :meth:`set_pivot_field_aggregation` (``PivotTable.AddDataField``).

        Removal is exposed two ways: ``orientation="hidden"`` (default) is the
        **non-destructive** ``Orientation = 0``; ``delete=True`` (which requires
        ``orientation="hidden"``) is the **destructive** ``PivotField.Delete``.
        The guard chain (writable → snapshot → approval) runs before any STA
        dispatch and the adapter re-reads the field afterwards.
        """

        normalized = _require_layout_orientation(orientation)
        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_SET_LAYOUT)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_set_layout",
            lambda wb: self._adapter.set_field_layout(
                wb, pivot_table, field, normalized, delete=delete
            ),
            record,
        )
        field_info = cast("PivotFieldInfo", outcome["field"])
        self._record_audit(
            "pivot_field_set_layout",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "orientation": normalized.value,
                "delete": delete,
            },
            {
                "pivot_table": pivot_table,
                "field": field_info.source_name,
                "orientation": field_info.orientation.value,
                "position": field_info.position,
                "verified": outcome["verified"],
                "orientation_order": outcome["orientation_order"],
                "before": _field_state(outcome["before"]),
                "after": _field_state(field_info),
            },
        )
        return field_info

    async def set_pivot_field_position(
        self, workbook_id: str, pivot_table: str, field: str, position: int
    ) -> PivotFieldInfo:
        """Set one field's 1-based position within its orientation (guarded).

        The position is validated against the orientation's field count inside
        the adapter **before** any COM mutation, so an out-of-range position is
        rejected without touching Excel. The resulting order of the orientation
        is recorded in the audit entry.
        """

        _require_position(position)
        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_SET_POSITION)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_set_position",
            lambda wb: self._adapter.set_field_position(wb, pivot_table, field, position),
            record,
        )
        field_info = cast("PivotFieldInfo", outcome["field"])
        self._record_audit(
            "pivot_field_set_position",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "position": position,
            },
            {
                "pivot_table": pivot_table,
                "field": field_info.source_name,
                "orientation": field_info.orientation.value,
                "position": field_info.position,
                "verified": outcome["verified"],
                "orientation_order": outcome["orientation_order"],
                "before": _field_state(outcome["before"]),
                "after": _field_state(field_info),
            },
        )
        return field_info

    async def set_pivot_field_aggregation(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        function: PivotAggregation,
        *,
        caption: str | None = None,
    ) -> PivotFieldInfo:
        """Add or update one data field's aggregation via ``AddDataField`` (guarded).

        An existing data field on the same source field is updated in place; a
        source field without a data field is added with
        ``PivotTable.AddDataField`` (assigning ``Orientation = xlDataField`` does
        not create one). ``caption`` identifies the data field and must be
        unique; only the eleven accepted aggregation codes are allowed (legacy
        codes are rejected before the STA dispatch).
        """

        _require_aggregation(function)
        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_SET_AGGREGATION)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_set_aggregation",
            lambda wb: self._adapter.set_field_aggregation(
                wb, pivot_table, field, function, caption=caption
            ),
            record,
        )
        field_info = cast("PivotFieldInfo", outcome["field"])
        self._record_audit(
            "pivot_field_set_aggregation",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "function": _aggregation_value(function),
                "caption": caption,
            },
            {
                "pivot_table": pivot_table,
                "field": field_info.source_name,
                "function": field_info.function.value if field_info.function else None,
                "caption": field_info.data_field_name,
                "verified": outcome["verified"],
                "data_field_order": outcome["orientation_order"],
                "before": _field_state(outcome["before"]),
                "after": _field_state(field_info),
            },
        )
        return field_info

    async def set_pivot_field_properties(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        properties: PivotFieldProperties,
    ) -> PivotFieldInfo:
        """Set bounded PivotField properties (caption/number format/show-all) (guarded).

        The sparse patch is validated before the STA dispatch (an empty patch, a
        blank caption or number format, or a duplicate caption is rejected
        without touching Excel); the adapter writes only the properties the
        caller explicitly set and re-reads them as a postcondition.
        """

        validate_field_properties(properties)
        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_SET_PROPERTIES)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_set_properties",
            lambda wb: self._adapter.set_field_properties(wb, pivot_table, field, properties),
            record,
        )
        field_info = cast("PivotFieldInfo", outcome["field"])
        self._record_audit(
            "pivot_field_set_properties",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "applied": sorted(properties.model_fields_set),
            },
            {
                "pivot_table": pivot_table,
                "field": field_info.source_name,
                "caption": field_info.data_field_name,
                "number_format": field_info.number_format,
                "show_all_items": field_info.show_all_items,
                "verified": outcome["verified"],
                "before": _field_state(outcome["before"]),
                "after": _field_state(field_info),
            },
        )
        return field_info

    # ------------------------------------------------------------------
    # Public API — PivotField filters (guarded; task_051)
    # ------------------------------------------------------------------
    async def apply_pivot_field_filter(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        spec: PivotFilterSpec,
    ) -> PivotFieldFilters:
        """Apply a supported manual filter to one field (guarded).

        The request is validated before anything else (a non-PivotFilterSpec, an
        unknown kind, a blank/duplicate/oversized item list or an empty manual
        filter is rejected). A label / value / date / top-bottom request maps to
        ``PivotFilters.Add``, which is unsupported on the target build: it returns
        an explicit ``unsupported`` result with a reason and **never** touches
        Excel or converts the request into a manual filter (task_039 §6).

        A supported manual filter then runs the guard chain (writable working copy
        → successful snapshot → approval) before the single STA dispatch. The
        adapter re-reads the visible/hidden items as a postcondition; a mismatch
        is a structured error, never a false success.
        """

        normalized = normalize_filter_spec(spec)
        reason = unsupported_filter_reason(normalized)
        record = self._resolve_record(workbook_id)
        if reason is not None:
            # Unsupported filter kind: no snapshot/approval needed because no
            # mutation is attempted. Audited as a refused mutating request.
            result = _unsupported_filter_dto(pivot_table, field, normalized, reason)
            self._record_audit(
                "pivot_field_apply_filter",
                "mutating",
                record,
                {
                    "workbook_id": workbook_id,
                    "pivot_table": pivot_table,
                    "field": field,
                    "kind": normalized.kind.value,
                    "visible_items": list(normalized.visible_items),
                },
                {
                    "pivot_table": pivot_table,
                    "field": field,
                    "applied": False,
                    "status": "unsupported",
                    "reason": reason,
                },
            )
            return result

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_APPLY_FILTER)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_apply_filter",
            lambda wb: self._adapter.apply_field_filter(wb, pivot_table, field, normalized),
            record,
        )
        filters = cast("PivotFieldFilters", outcome["filters"])
        self._record_audit(
            "pivot_field_apply_filter",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "kind": normalized.kind.value,
                "visible_items": list(normalized.visible_items),
            },
            {
                "pivot_table": filters.pivot_table,
                "field": filters.field,
                "applied": outcome["applied"],
                "verified": outcome["verified"],
                "kind": filters.filter.kind.value,
                "visible_items": list(filters.visible_items),
                "hidden_items": list(filters.hidden_items),
                "before": _filters_audit_state(outcome["before"]),
            },
        )
        return filters

    async def clear_pivot_field_filters(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str | None = None,
        *,
        scope: str = "field",
    ) -> PivotFieldFilters:
        """Clear filters for one field or the whole PivotTable (guarded).

        ``scope="field"`` (the default) requires ``field`` and calls
        ``PivotField.ClearAllFilters`` on that field; ``scope="table"`` calls
        ``PivotTable.ClearAllFilters`` for the whole PivotTable. The scope and the
        field requirement are validated before the guard chain, so an ambiguous
        request never reaches Excel. The adapter re-reads the field state as a
        postcondition; a mismatch is reported truthfully.
        """

        normalized_scope = normalize_filter_scope(scope)
        if normalized_scope == "field" and (not isinstance(field, str) or not field.strip()):
            raise AmbiguousFilterError(
                "Argument 'field' is required when scope is 'field'.",
                context={
                    "pivot_table": pivot_table,
                    "scope": normalized_scope,
                    "reason": "field_required",
                },
            )

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_CLEAR_FILTERS)

        outcome = await self._submit_mutation(
            workbook_id,
            "pivot_field_clear_filters",
            lambda wb: self._adapter.clear_field_filters(
                wb, pivot_table, field, scope=normalized_scope
            ),
            record,
        )
        filters = cast("PivotFieldFilters", outcome["filters"])
        self._record_audit(
            "pivot_field_clear_filters",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "pivot_table": pivot_table,
                "field": field,
                "scope": normalized_scope,
            },
            {
                "pivot_table": filters.pivot_table,
                "field": filters.field,
                "scope": outcome["scope"],
                "cleared_fields": list(outcome["cleared_fields"]),
                "verified": outcome["verified"],
                "before": _filter_states_audit(outcome["before"]),
            },
        )
        return filters

    # ------------------------------------------------------------------
    # Guards (non-COM)
    # ------------------------------------------------------------------
    def _require_writable(self, workbook_id: str) -> WorkbookInfo:
        """Resolve the record and reject a read-only working copy."""

        record = self._resolve_record(workbook_id)
        if record.read_only:
            raise ExcelVbaMcpError(
                "A pivot mutation requires a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": workbook_id, "path": record.path},
            )
        return record

    def _require_snapshot(self, workbook_id: str) -> None:
        """Block the mutation with ``SNAPSHOT_FAILED`` when no snapshot exists.

        Fails closed when the snapshot service is not wired (read-only rig).
        """

        if self._snapshot is None:
            raise ExcelVbaMcpError(
                "The snapshot service is not configured; a pivot mutation requires it.",
                code=ErrorCode.SNAPSHOT_FAILED,
                context={"tool": "pivot_mutation"},
            )
        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating pivot operation (deny-by-default)."""

        if self._approval is None:
            raise ApprovalDeniedError(
                "The approval policy is not configured; a pivot mutation requires it.",
                context={"tool": "pivot_mutation", "action": action},
            )
        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"Pivot mutation {action!r} denied by approval policy: {decision.reason}",
                context={
                    "action": action,
                    "workbook_id": record.workbook_id,
                    "reason": decision.reason,
                },
            )

    async def _submit_mutation(
        self,
        workbook_id: str,
        tool: str,
        fn: Callable[[Any], _T],
        record: WorkbookInfo,
    ) -> _T:
        """Run a mutating ``fn(wb)`` on the STA worker (never auto-retried)."""

        call = self._bind_workbook(fn, record, workbook_id)
        call.__name__ = tool
        result = cast(
            "_T",
            await self._worker.submit(
                call,
                timeout=_DEFAULT_TIMEOUT_S,
                section="diagnostic",
                op_class=OperationClass.MUTATING,
            ),
        )
        self._session.bump_version(workbook_id)
        return result

    # ------------------------------------------------------------------
    # STA submission helpers
    # ------------------------------------------------------------------
    def _resolve_record(self, workbook_id: str) -> WorkbookInfo:
        """Resolve ``workbook_id`` to its open :class:`WorkbookInfo` (no COM)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise WorkbookNotFoundError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            context={"workbook_id": workbook_id},
        )

    def _app_adapter(self) -> ExcelAppAdapter:
        """Return the owned hidden Excel application adapter (never a new one)."""

        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "pivots"},
            )
        return adapter

    async def _submit_read(
        self,
        workbook_id: str,
        tool: str,
        fn: Callable[[Any], _T],
        record: WorkbookInfo,
    ) -> _T:
        """Run a read-only ``fn(wb)`` on the STA worker."""

        call = self._bind_workbook(fn, record, workbook_id)
        call.__name__ = tool
        return cast(
            "_T",
            await self._worker.submit(
                call,
                timeout=_DEFAULT_TIMEOUT_S,
                section="diagnostic",
                op_class=OperationClass.READ_ONLY,
            ),
        )

    def _bind_workbook(
        self,
        fn: Callable[[Any], Any],
        record: WorkbookInfo,
        workbook_id: str,
    ) -> Callable[[], Any]:
        """Wrap ``fn`` so it re-locates the working copy on the STA thread."""

        app_adapter = self._app_adapter()
        working_path = record.path

        def run() -> Any:
            wb = _find_open_workbook(app_adapter, working_path)
            if wb is None:
                raise WorkbookNotFoundError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": working_path},
                )
            return fn(wb)

        return run

    # ------------------------------------------------------------------
    # Adapter helpers (non-COM)
    # ------------------------------------------------------------------
    def _find_pivot_table(self, wb: Any, pivot_table: str) -> PivotTableInfo:
        """Return the :class:`PivotTableInfo` matching ``pivot_table`` by name."""

        target = pivot_table.strip().lower()
        for info in self._adapter.enumerate_pivot_tables(wb):
            if info.name.lower() == target:
                return info
        raise ExcelVbaMcpError(
            f"PivotTable {pivot_table!r} was not found in the workbook.",
            context={"pivot_table": pivot_table},
        )

    # ------------------------------------------------------------------
    # Audit (non-COM)
    # ------------------------------------------------------------------
    def _record_audit(
        self,
        tool: str,
        operation_class: str,
        record: WorkbookInfo,
        params: dict[str, Any],
        result: dict[str, Any],
    ) -> None:
        canonical = ""
        with contextlib.suppress(ExcelVbaMcpError):
            canonical = self._resolve_record(record.workbook_id).source_path or ""
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": canonical},
            result=result,
        )


def _require_layout_orientation(orientation: PivotFieldOrientation) -> PivotFieldOrientation:
    """Validate a target orientation for ``set_pivot_field_layout`` (non-COM).

    Only the layout buckets are assignable; ``data`` is rejected here (before any
    STA dispatch) because assigning ``Orientation = xlDataField`` never creates a
    data field — that is ``PivotTable.AddDataField``.
    """

    if not isinstance(orientation, PivotFieldOrientation):
        raise ExcelVbaMcpError(
            f"Argument 'orientation' must be a PivotFieldOrientation, got {orientation!r}.",
            code=ErrorCode.INVALID_RANGE,
            context={"orientation": repr(orientation), "reason": "invalid_orientation"},
        )
    if orientation not in _LAYOUT_ORIENTATIONS:
        raise ExcelVbaMcpError(
            "Orientation 'data' cannot be assigned: it does not create a data "
            "field. Use pivot_field_set_aggregation (AddDataField) instead; the "
            "layout buckets are 'hidden', 'row', 'column' and 'page'.",
            code=ErrorCode.INVALID_RANGE,
            context={
                "orientation": orientation.value,
                "reason": "invalid_orientation",
            },
        )
    return orientation


def _require_position(position: int) -> int:
    """Validate a 1-based field position (non-COM)."""

    if isinstance(position, bool) or not isinstance(position, int):
        raise ExcelVbaMcpError(
            f"Argument 'position' must be an integer, got {position!r}.",
            code=ErrorCode.INVALID_RANGE,
            context={"position": repr(position), "reason": "position_out_of_range"},
        )
    if position < 1:
        raise ExcelVbaMcpError(
            f"Argument 'position' must be >= 1, got {position}.",
            code=ErrorCode.INVALID_RANGE,
            context={"position": position, "reason": "position_out_of_range"},
        )
    return position


def _require_aggregation(function: PivotAggregation) -> PivotAggregation:
    """Validate an aggregation request (non-COM).

    Only the eleven accepted ``PivotAggregation`` members are allowed; legacy
    codes the target build rejects cannot be expressed and are refused by the
    adapter's code table as well.
    """

    if not isinstance(function, PivotAggregation):
        raise ExcelVbaMcpError(
            f"Argument 'function' must be a PivotAggregation, got {function!r}. "
            "Legacy codes (9, -4114, -4153, -4154) are not accepted.",
            code=ErrorCode.INVALID_RANGE,
            context={
                "function": repr(function),
                "accepted": [member.value for member in PivotAggregation],
                "reason": "invalid_function",
            },
        )
    return function


def _aggregation_value(function: PivotAggregation) -> str:
    """Return the audit-friendly value of an aggregation argument."""

    return function.value if isinstance(function, PivotAggregation) else repr(function)


def _field_state(info: PivotFieldInfo) -> dict[str, Any]:
    """Return the audit-friendly subset of a field's post-mutation state."""

    return {
        "source_name": info.source_name,
        "caption": info.caption,
        "orientation": info.orientation.value,
        "position": info.position,
        "is_data_field": info.is_data_field,
        "function": info.function.value if info.function else None,
        "data_field_name": info.data_field_name,
        "number_format": info.number_format,
        "show_all_items": info.show_all_items,
    }


def _unsupported_filter_dto(
    pivot_table: str, field: str, spec: PivotFilterSpec, reason: str
) -> PivotFieldFilters:
    """Build the explicit unsupported result for a non-manual filter request."""

    kind = (
        PivotFilterKind(spec.kind.value)
        if isinstance(spec.kind, PivotFilterRequestKind)
        else PivotFilterKind.UNSUPPORTED
    )
    return PivotFieldFilters(
        pivot_table=pivot_table,
        field=field,
        filter=PivotFilterState(
            kind=kind,
            visible_items=[],
            supported=False,
            unsupported_reason=reason,
        ),
        unsupported=[spec.kind.value],
    )


def _filters_audit_state(filters: PivotFieldFilters | None) -> dict[str, Any] | None:
    """Return the audit-friendly subset of one field's filter state."""

    if filters is None:
        return None
    return {
        "field": filters.field,
        "all_items_visible": filters.all_items_visible,
        "visible_items": list(filters.visible_items),
        "hidden_items": list(filters.hidden_items),
        "kind": filters.filter.kind.value,
        "unsupported": list(filters.unsupported),
    }


def _filter_states_audit(states: Any) -> Any:
    """Return an audit-friendly view of a before-state (DTO or name→DTO map)."""

    if isinstance(states, PivotFieldFilters):
        return _filters_audit_state(states)
    if isinstance(states, dict):
        return {name: _filters_audit_state(value) for name, value in states.items()}
    return None


def _overall_sheet_status(sheet_statuses: list[dict[str, Any]]) -> str:
    """Reduce the per-sheet sweep statuses to one overall status (Block B).

    Severity is ``too_large`` > ``capped`` > ``ok``: a workbook with any
    ``too_large`` sheet is reported ``too_large`` (the most severe), else
    ``capped`` if any sheet was capped, else ``ok``.
    """

    statuses = {entry.get("status") for entry in sheet_statuses}
    if SHEET_STATUS_TOO_LARGE in statuses:
        return SHEET_STATUS_TOO_LARGE
    if SHEET_STATUS_CAPPED in statuses:
        return SHEET_STATUS_CAPPED
    return SHEET_STATUS_OK


def _find_open_workbook(app_adapter: ExcelAppAdapter, working_path: str) -> Any | None:
    """Return the open ``Workbook`` proxy matching ``working_path`` (STA thread)."""

    app = app_adapter.app
    if app is None:
        return None
    target = working_path.lower()
    try:
        for wb in app.Workbooks:
            try:
                if str(wb.FullName).lower() == target:
                    return wb
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return None


__all__ = ["PivotTablesService"]
