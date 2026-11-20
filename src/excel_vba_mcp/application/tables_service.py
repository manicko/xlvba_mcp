"""Excel Table (ListObject) inventory + row/column operations service (task_026).

:class:`TablesService` is the async application-layer orchestration for the table
surface (``tables_list``, ``table_get``, ``table_read_rows`` and the mutating
``table_append_rows`` / ``table_insert_row`` / ``table_update_rows`` /
``table_delete_rows`` / ``table_add_column`` / ``table_delete_column`` /
``table_create`` / ``table_resize`` tools). It is deliberately **COM-free**: every
method resolves the ``workbook_id`` through the session bookkeeping, enforces the
non-COM guard chain and then submits a single STA-local callable to the dedicated
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only
domain DTOs cross the worker boundary.

Guard chain for mutating operations (deny-by-default, spec §7.6 / §9.1):

1. **Writable session** — a read-only working copy is rejected with
   ``WORKBOOK_READ_ONLY``.
2. **Snapshot** — ``snapshot_service.require_snapshot`` blocks with
   ``SNAPSHOT_FAILED`` when no successful snapshot exists.
3. **Approval** — ``ApprovalPolicy.authorize_vba_write`` (a ``MUTATING``
   classification) denies unless a live grant/allowlist entry matches the action +
   workbook identity.

Every mutating adapter call captures before/after metadata (a
:class:`~excel_vba_mcp.adapters.excel.tables_adapter.TableMutationResult`) so the
returned DTO reflects the actual post-mutation table, and every tool invocation is
recorded in the append-only audit trail.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from typing import Any, TypeVar, cast

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.tables_adapter import (
    TableMutationResult,
    TableRows,
    TablesAdapter,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import TableInfo, WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds) for a table operation on the STA worker.
_DEFAULT_TIMEOUT_S = 60.0

# Approval action names (stable identities for the table mutation surface).
_ACTION_APPEND = "table_append_rows"
_ACTION_INSERT = "table_insert_row"
_ACTION_UPDATE = "table_update_rows"
_ACTION_DELETE = "table_delete_rows"
_ACTION_ADD_COLUMN = "table_add_column"
_ACTION_DELETE_COLUMN = "table_delete_column"
_ACTION_CREATE = "table_create"
_ACTION_RESIZE = "table_resize"

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class TablesService:
    """Async Excel Table inventory + row/column operations (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        snapshot_service: Pre-mutation snapshot gate.
        approval: Deny-by-default approval gate for mutating table operations.
        path_policy: Allowlist policy (used to validate the working-copy path
            identity for approval keys).
        audit: Append-only JSONL audit trail.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        snapshot_service: SnapshotService,
        approval: ApprovalPolicy,
        path_policy: PathPolicy,
        audit: AuditLogger,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._snapshot = snapshot_service
        self._approval = approval
        self._path_policy = path_policy
        self._audit = audit
        self._adapter = TablesAdapter()

    # ------------------------------------------------------------------
    # Public API — read-only (async; no COM here)
    # ------------------------------------------------------------------
    async def list_tables(self, workbook_id: str) -> list[TableInfo]:
        """Return the table inventory of ``workbook_id`` (all worksheets)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id, "tables_list", lambda wb: self._adapter.list_tables(wb), record
        )
        self._record_audit(
            "tables_list",
            "read_only",
            record,
            {"workbook_id": workbook_id},
            {"count": len(result)},
        )
        return result

    async def get_table(self, workbook_id: str, name: str) -> TableInfo:
        """Return the :class:`TableInfo` for a single table by name."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "table_get",
            lambda wb: self._adapter.get_table(wb, name),
            record,
        )
        self._record_audit(
            "table_get",
            "read_only",
            record,
            {"workbook_id": workbook_id, "table": name},
            {"table": result.name, "sheet": result.sheet_name, "rows": result.row_count},
        )
        return result

    async def read_table_rows(
        self,
        workbook_id: str,
        name: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> TableRows:
        """Read a typed, bounded slice of a table's data rows (read-only)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "table_read_rows",
            lambda wb: self._adapter.read_rows(wb, name, offset=offset, limit=limit),
            record,
        )
        self._record_audit(
            "table_read_rows",
            "read_only",
            record,
            {
                "workbook_id": workbook_id,
                "table": name,
                "offset": offset,
                "limit": limit,
            },
            {
                "table": result.name,
                "returned_rows": result.returned_rows,
                "total_rows": result.total_rows,
                "truncated": result.truncated,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — mutations (guarded)
    # ------------------------------------------------------------------
    async def append_table_rows(
        self, workbook_id: str, name: str, rows: list[dict[str, Any]]
    ) -> TableMutationResult:
        """Append rows through ``ListRows.Add`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_APPEND,
            tool="table_append_rows",
            params={"workbook_id": workbook_id, "table": name, "row_count": len(rows)},
            fn=lambda wb: self._adapter.append_rows(wb, name, rows),
        )

    async def insert_table_row(
        self,
        workbook_id: str,
        name: str,
        at: int,
        values: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        """Insert one row at 1-based position ``at`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_INSERT,
            tool="table_insert_row",
            params={"workbook_id": workbook_id, "table": name, "at": at},
            fn=lambda wb: self._adapter.insert_row(wb, name, at, values),
        )

    async def update_table_rows(
        self, workbook_id: str, name: str, updates: list[dict[str, Any]]
    ) -> TableMutationResult:
        """Update rows by explicit index / unambiguous filter (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_UPDATE,
            tool="table_update_rows",
            params={"workbook_id": workbook_id, "table": name, "update_count": len(updates)},
            fn=lambda wb: self._adapter.update_rows(wb, name, updates),
        )

    async def delete_table_rows(
        self,
        workbook_id: str,
        name: str,
        indices: list[int] | None = None,
        where: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        """Delete rows by explicit indices / unambiguous filter (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_DELETE,
            tool="table_delete_rows",
            params={
                "workbook_id": workbook_id,
                "table": name,
                "indices": list(indices) if indices is not None else None,
            },
            fn=lambda wb: self._adapter.delete_rows(wb, name, indices, where),
        )

    async def add_table_column(
        self,
        workbook_id: str,
        name: str,
        header: str,
        position: int | None = None,
    ) -> TableMutationResult:
        """Add a column via ``ListColumns.Add`` with a unique header (guarded)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_ADD_COLUMN,
            tool="table_add_column",
            params={
                "workbook_id": workbook_id,
                "table": name,
                "header": header,
                "position": position,
            },
            fn=lambda wb: self._adapter.add_column(wb, name, header, position),
        )

    async def delete_table_column(
        self, workbook_id: str, name: str, header: str
    ) -> TableMutationResult:
        """Delete the column identified by ``header`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_DELETE_COLUMN,
            tool="table_delete_column",
            params={"workbook_id": workbook_id, "table": name, "header": header},
            fn=lambda wb: self._adapter.delete_column(wb, name, header),
        )

    async def create_table(
        self,
        workbook_id: str,
        sheet: str,
        range_or_name: str,
        *,
        style: str | None = None,
    ) -> TableMutationResult:
        """Create a ListObject on ``sheet`` over ``range_or_name`` (guarded)."""

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_CREATE)

        def create_fn(wb: Any) -> TableMutationResult:
            return self._adapter.create_table(wb, sheet, range_or_name, style=style)

        result = await self._submit_mutation(workbook_id, "table_create", create_fn, record)
        self._record_audit(
            "table_create",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "range_or_name": range_or_name,
                "style": style,
            },
            {
                "table": result.after.name,
                "sheet": result.after.sheet_name,
                "rows": result.after.row_count,
            },
        )
        return result

    async def resize_table(
        self, workbook_id: str, name: str, new_range: str
    ) -> TableMutationResult:
        """Resize a table to ``new_range`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_RESIZE,
            tool="table_resize",
            params={"workbook_id": workbook_id, "table": name, "new_range": new_range},
            fn=lambda wb: self._adapter.resize(wb, name, new_range),
        )

    # ------------------------------------------------------------------
    # Guards (non-COM)
    # ------------------------------------------------------------------
    def _require_writable(self, workbook_id: str) -> WorkbookInfo:
        """Resolve the record and reject a read-only working copy."""

        record = self._resolve_record(workbook_id)
        if record.read_only:
            raise ExcelVbaMcpError(
                "Table mutations require a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": workbook_id, "path": record.path},
            )
        return record

    def _require_snapshot(self, workbook_id: str) -> None:
        """Block the mutation with ``SNAPSHOT_FAILED`` when no snapshot exists."""

        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating table operation (deny-by-default)."""

        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ExcelVbaMcpError(
                f"Table mutation {action!r} denied by approval policy: {decision.reason}",
                code=ErrorCode.COM_BUSY,
                context={
                    "action": action,
                    "workbook_id": record.workbook_id,
                    "reason": decision.reason,
                },
            )

    # ------------------------------------------------------------------
    # Mutation orchestration (guarded)
    # ------------------------------------------------------------------
    async def _mutate(
        self,
        workbook_id: str,
        name: str,
        *,
        action: str,
        tool: str,
        params: dict[str, Any],
        fn: Callable[[Any], TableMutationResult],
    ) -> TableMutationResult:
        """Run the guard chain then a single mutating STA dispatch + audit."""

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, action)

        result = await self._submit_mutation(workbook_id, tool, fn, record)
        self._record_audit(
            tool,
            "mutating",
            record,
            params,
            {
                "table": result.after.name,
                "action": result.action,
                "row_count": result.after.row_count,
                "affected_rows": result.affected_rows,
            },
        )
        return result

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
                context={"tool": "tables"},
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


__all__ = ["TablesService"]
