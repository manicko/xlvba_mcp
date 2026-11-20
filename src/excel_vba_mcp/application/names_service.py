"""Defined-name inventory + scoped CRUD service (task_028, spec §7.7 / §6.5).

:class:`NamesService` is the async application-layer orchestration for the
defined-name surface (``names_list``, ``name_get``, ``name_create``,
``name_update`` and ``name_delete``). It is deliberately **COM-free**: every
method resolves the ``workbook_id`` through the session bookkeeping, enforces the
non-COM guard chain and then submits a single STA-local callable to the dedicated
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only
domain DTOs cross the worker boundary.

Guard chain for mutating operations (deny-by-default, spec §7.7 / §9.1):

1. **Writable session** — a read-only working copy is rejected with
   ``WORKBOOK_READ_ONLY``.
2. **Snapshot** — ``snapshot_service.require_snapshot`` blocks with
   ``SNAPSHOT_FAILED`` when no successful snapshot exists.
3. **Approval** — ``ApprovalPolicy.authorize_vba_write`` (a ``MUTATING``
   classification) denies unless a live grant/allowlist entry matches the action +
   workbook identity.

Reads (``names_list`` / ``name_get``) are ``READ_ONLY`` and require no approval.
Every tool invocation is recorded in the append-only audit trail.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from typing import Any, TypeVar, cast

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.names_adapter import NamesAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import DefinedNameInfo, WorkbookInfo
from excel_vba_mcp.domain.statuses import NameScope, OperationClass
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds) for a name operation on the STA worker.
_DEFAULT_TIMEOUT_S = 60.0

# Approval action names (stable identities for the name mutation surface).
_ACTION_CREATE = "name_create"
_ACTION_UPDATE = "name_update"
_ACTION_DELETE = "name_delete"

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class NamesService:
    """Async defined-name inventory + scoped CRUD (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        snapshot_service: Pre-mutation snapshot gate.
        approval: Deny-by-default approval gate for mutating name operations.
        path_policy: Allowlist policy (kept for signature symmetry with the other
            guarded services; approval keys use the canonical source path).
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
        self._adapter = NamesAdapter()

    # ------------------------------------------------------------------
    # Public API — read-only (async; no COM here)
    # ------------------------------------------------------------------
    async def list_names(self, workbook_id: str) -> list[DefinedNameInfo]:
        """Return every defined name of ``workbook_id`` (all scopes)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id, "names_list", lambda wb: self._adapter.list_names(wb), record
        )
        self._record_audit(
            "names_list",
            "read_only",
            record,
            {"workbook_id": workbook_id},
            {"count": len(result)},
        )
        return result

    async def get_name(
        self,
        workbook_id: str,
        name: str,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Return the :class:`DefinedNameInfo` for one name in ``scope``."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "name_get",
            lambda wb: self._adapter.get_name(wb, name, scope, sheet),
            record,
        )
        self._record_audit(
            "name_get",
            "read_only",
            record,
            {
                "workbook_id": workbook_id,
                "name": name,
                "scope": str(scope),
                "sheet": sheet,
            },
            {
                "name": result.name,
                "scope": result.scope.value,
                "target_kind": result.target_kind.value,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — mutations (guarded)
    # ------------------------------------------------------------------
    async def create_name(
        self,
        workbook_id: str,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Create a defined name in ``scope`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_CREATE,
            tool="name_create",
            params={
                "workbook_id": workbook_id,
                "name": name,
                "scope": str(scope),
                "sheet": sheet,
            },
            fn=lambda wb: self._adapter.create_name(wb, name, refers_to, scope=scope, sheet=sheet),
            result_keys=("name", "scope", "target_kind"),
        )

    async def update_name(
        self,
        workbook_id: str,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Update a name's ``RefersTo`` in ``scope`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_UPDATE,
            tool="name_update",
            params={
                "workbook_id": workbook_id,
                "name": name,
                "scope": str(scope),
                "sheet": sheet,
            },
            fn=lambda wb: self._adapter.update_name(wb, name, refers_to, scope=scope, sheet=sheet),
            result_keys=("name", "scope", "target_kind"),
        )

    async def delete_name(
        self,
        workbook_id: str,
        name: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Delete a name in ``scope`` (guarded mutation)."""

        return await self._mutate(
            workbook_id,
            name,
            action=_ACTION_DELETE,
            tool="name_delete",
            params={
                "workbook_id": workbook_id,
                "name": name,
                "scope": str(scope),
                "sheet": sheet,
            },
            fn=lambda wb: self._adapter.delete_name(wb, name, scope=scope, sheet=sheet),
            result_keys=("name", "scope", "target_kind"),
        )

    # ------------------------------------------------------------------
    # Guards (non-COM)
    # ------------------------------------------------------------------
    def _require_writable(self, workbook_id: str) -> WorkbookInfo:
        """Resolve the record and reject a read-only working copy."""

        record = self._resolve_record(workbook_id)
        if record.read_only:
            raise ExcelVbaMcpError(
                "Name mutations require a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": workbook_id, "path": record.path},
            )
        return record

    def _require_snapshot(self, workbook_id: str) -> None:
        """Block the mutation with ``SNAPSHOT_FAILED`` when no snapshot exists."""

        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating name operation (deny-by-default)."""

        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ExcelVbaMcpError(
                f"Name mutation {action!r} denied by approval policy: {decision.reason}",
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
        fn: Callable[[Any], DefinedNameInfo],
        result_keys: tuple[str, ...],
    ) -> DefinedNameInfo:
        """Run the guard chain then a single mutating STA dispatch + audit."""

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, action)

        result = await self._submit_mutation(workbook_id, tool, fn, record)
        summary: dict[str, Any] = {}
        for key in result_keys:
            value = getattr(result, key, None)
            summary[key] = getattr(value, "value", value)
        self._record_audit(tool, "mutating", record, params, summary)
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
                context={"tool": "names"},
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


__all__ = ["NamesService"]
