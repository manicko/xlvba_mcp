"""Excel PivotTables / PivotCaches / PivotFields service (task_048, task_049).

:class:`PivotTablesService` is the async application-layer orchestration for the
Phase 2 PivotTables surface. The read-only inventory tools
(``pivot_tables_list``, ``pivot_table_get``, ``pivot_cache_get``,
``pivot_fields_list``) are implemented by ``task_048``; the targeted refresh
tools (``pivot_table_refresh``, ``pivot_cache_refresh``) are implemented by
``task_049``. It is deliberately **COM-free**: every method resolves the
``workbook_id`` through the session bookkeeping and then submits a single
STA-local callable to the dedicated single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only domain
DTOs cross the worker boundary (ARCH-009).

Read-only path (task_048): no snapshot / approval guard chain is required (this
task never mutates a PivotTable, PivotCache or PivotField). Every call is
recorded in the append-only audit trail as ``read_only``.

Mutation path (task_049): ``refresh_pivot_table`` / ``refresh_pivot_cache`` run
the full guard chain — writable working copy, successful snapshot, deny-by-default
approval — then a single mutating STA dispatch, recorded as ``mutating``. Only a
worksheet/Table-backed (``database``) cache is refreshed; external / OLAP / Data
Model sources return ``unsupported_source`` and are never refreshed.
``Workbook.RefreshAll`` is never used.

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
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import PivotTablesAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldInfo,
    PivotInventory,
    PivotRefreshResult,
    PivotTableInfo,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger

#: Default per-call soft ceiling (seconds) for a pivot inventory/refresh call on
#: the STA worker. Inventory reads are bounded and fast; a targeted refresh of a
#: local cache is also bounded, so the same ceiling as the other services is used.
_DEFAULT_TIMEOUT_S = 60.0

#: Approval action names (stable identities for the pivot mutation surface).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_REFRESH_CACHE = "pivot_cache_refresh"

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
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        audit: AuditLogger,
        snapshot_service: SnapshotService | None = None,
        approval: ApprovalPolicy | None = None,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._audit = audit
        self._snapshot = snapshot_service
        self._approval = approval
        self._adapter = PivotTablesAdapter()

    # ------------------------------------------------------------------
    # Public API — read-only (async; no COM here)
    # ------------------------------------------------------------------
    async def list_pivot_tables(self, workbook_id: str) -> PivotInventory:
        """Return every PivotTable and its cache across all worksheets.

        Enumeration is worksheet-scoped (a full-workbook sweep) because
        ``Workbook.PivotTables`` is not a usable inventory on the target build.
        Shared caches are reported once with their affected tables.
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "pivot_tables_list",
            lambda wb: PivotInventory(
                workbook_id=workbook_id,
                pivot_tables=self._adapter.enumerate_pivot_tables(wb),
                caches=self._collect_caches(wb),
            ),
            record,
        )
        self._record_audit(
            "pivot_tables_list",
            "read_only",
            record,
            {"workbook_id": workbook_id},
            {
                "pivot_table_count": len(result.pivot_tables),
                "cache_count": len(result.caches),
            },
        )
        return result

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
    # Guards (non-COM)
    # ------------------------------------------------------------------
    def _require_writable(self, workbook_id: str) -> WorkbookInfo:
        """Resolve the record and reject a read-only working copy."""

        record = self._resolve_record(workbook_id)
        if record.read_only:
            raise ExcelVbaMcpError(
                "Pivot refresh requires a writable working copy.",
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
                "The snapshot service is not configured; refresh requires it.",
                code=ErrorCode.SNAPSHOT_FAILED,
                context={"tool": "pivot_refresh"},
            )
        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating pivot operation (deny-by-default)."""

        if self._approval is None:
            raise ExcelVbaMcpError(
                "The approval policy is not configured; refresh requires it.",
                code=ErrorCode.COM_BUSY,
                context={"tool": "pivot_refresh", "action": action},
            )
        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ExcelVbaMcpError(
                f"Pivot refresh {action!r} denied by approval policy: {decision.reason}",
                code=ErrorCode.COM_BUSY,
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
        return cast(
            "_T",
            await self._worker.submit(
                call,
                timeout=_DEFAULT_TIMEOUT_S,
                section="diagnostic",
                op_class=OperationClass.MUTATING,
            ),
        )

    # ------------------------------------------------------------------
    # STA submission helpers
    # ------------------------------------------------------------------
    def _resolve_record(self, workbook_id: str) -> WorkbookInfo:
        """Resolve ``workbook_id`` to its open :class:`WorkbookInfo` (no COM)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise ExcelVbaMcpError(
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
                raise ExcelVbaMcpError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": working_path},
                )
            return fn(wb)

        return run

    # ------------------------------------------------------------------
    # Adapter helpers (non-COM)
    # ------------------------------------------------------------------
    def _collect_caches(self, wb: Any) -> list[PivotCacheInfo]:
        """Build one :class:`PivotCacheInfo` per distinct ``CacheIndex`` (STA-local)."""

        pivot_tables = self._adapter.enumerate_pivot_tables(wb)
        cache_indices = sorted({pt.cache_index for pt in pivot_tables})
        caches: list[PivotCacheInfo] = []
        for cache_index in cache_indices:
            caches.append(self._adapter.read_pivot_cache(wb, cache_index))
        return caches

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
