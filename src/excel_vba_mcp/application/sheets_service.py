"""Sheets inventory + worksheet operations service (task_024, spec §7.5/§6.5).

:class:`SheetsService` is the async application-layer orchestration for the
sheet surface. It is deliberately **COM-free**: every method resolves the
``workbook_id`` through the session bookkeeping, enforces the non-COM guard
chain (writable session, snapshot, approval, visible-sheet guard) and then
submits a single STA-local callable to the dedicated single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only domain
DTOs cross the worker boundary.

Guard chain for mutating operations (deny-by-default):

1. **Writable session** — a read-only working copy is rejected with
   ``WORKBOOK_READ_ONLY``.
2. **Snapshot** — ``snapshot_service.require_snapshot`` blocks with
   ``SNAPSHOT_FAILED`` when no successful snapshot exists.
3. **Approval** — ``ApprovalPolicy.authorize_vba_write`` (a ``MUTATING``
   classification) denies unless a live grant/allowlist entry matches the
   action + workbook identity.
4. **Visible-sheet guard** (delete only) — deleting the last visible worksheet
   is rejected with a structured error so a workbook always keeps at least one
   visible sheet.

Every mutating adapter call is followed by a postcondition read (the collection
state) so the returned DTO reflects the actual post-mutation workbook, and every
tool invocation is recorded in the append-only audit trail.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from typing import Any, TypeVar, cast

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.sheets_adapter import SheetsAdapter, visible_worksheet_count
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import SheetInfo, WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds) for a sheet operation on the STA worker.
_DEFAULT_TIMEOUT_S = 60.0

# Approval action names (stable identities for the sheet mutation surface).
_ACTION_ADD = "sheet_add"
_ACTION_RENAME = "sheet_rename"
_ACTION_COPY = "sheet_copy"
_ACTION_DELETE = "sheet_delete"
_ACTION_VISIBILITY = "sheet_set_visibility"

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class SheetsService:
    """Async Sheets inventory + worksheet operations (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        snapshot_service: Pre-mutation snapshot gate.
        approval: Deny-by-default approval gate for mutating sheet operations.
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
        self._adapter = SheetsAdapter()

    # ------------------------------------------------------------------
    # Public API (async; no COM here — all COM on the STA worker)
    # ------------------------------------------------------------------
    async def list_sheets(self, workbook_id: str) -> list[SheetInfo]:
        """Return the sheet inventory of ``workbook_id`` (all sheet types)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id, "sheets_list", lambda wb: self._adapter.list_sheets(wb), record
        )
        self._record_audit(
            "sheets_list",
            "read_only",
            record,
            {"workbook_id": workbook_id},
            {"count": len(result)},
        )
        return result

    async def get_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        """Return the :class:`SheetInfo` for a single sheet by name."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "sheet_get",
            lambda wb: self._adapter.get_sheet(wb, name),
            record,
        )
        self._record_audit(
            "sheet_get",
            "read_only",
            record,
            {"workbook_id": workbook_id, "sheet": name},
            {"sheet": result.name, "object_type": result.object_type.value},
        )
        return result

    async def add_sheet(
        self, workbook_id: str, name: str, sheet_type: str = "worksheet"
    ) -> SheetInfo:
        """Add a worksheet named ``name`` to ``workbook_id``.

        Guarded mutation: writable session + snapshot + approval. Postcondition:
        the returned DTO is re-read from the live collection after the add.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_ADD)

        def add_fn(wb: Any) -> SheetInfo:
            self._adapter.add_sheet(wb, name, sheet_type=sheet_type)
            return self._adapter.get_sheet(wb, name)

        result = await self._submit_mutation(workbook_id, "sheet_add", add_fn, record)
        self._record_audit(
            "sheet_add",
            "mutating",
            record,
            {"workbook_id": workbook_id, "sheet": name, "sheet_type": sheet_type},
            {"sheet": result.name, "object_type": result.object_type.value},
        )
        return result

    async def rename_sheet(self, workbook_id: str, name: str, new_name: str) -> SheetInfo:
        """Rename ``name`` to ``new_name`` in ``workbook_id``.

        Guarded mutation: writable session + snapshot + approval. ``CodeName`` is
        never modified (the adapter only sets the user-visible ``Name``).
        Postcondition: the returned DTO is re-read under the new name.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_RENAME)

        def rename_fn(wb: Any) -> SheetInfo:
            self._adapter.rename_sheet(wb, name, new_name)
            return self._adapter.get_sheet(wb, new_name)

        result = await self._submit_mutation(workbook_id, "sheet_rename", rename_fn, record)
        self._record_audit(
            "sheet_rename",
            "mutating",
            record,
            {"workbook_id": workbook_id, "old_name": name, "new_name": new_name},
            {"sheet": result.name, "code_name": result.code_name},
        )
        return result

    async def copy_sheet(
        self, workbook_id: str, name: str, *, target_name: str | None = None
    ) -> SheetInfo:
        """Copy ``name`` in ``workbook_id`` and return the new sheet's DTO.

        Guarded mutation: writable session + snapshot + approval. Postcondition:
        the returned DTO is the newly created sheet (resolved by the adapter).

        When ``target_name`` is given the copy is renamed to it (validated by the
        adapter: non-empty, ``<= 31`` chars, unique; an invalid value raises
        ``INVALID_RANGE``). When ``target_name`` is ``None`` the copy keeps the
        Excel-generated name (unchanged behaviour).
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_COPY)

        def copy_fn(wb: Any) -> SheetInfo:
            return self._adapter.copy_sheet(wb, name, to_end=True, target_name=target_name)

        result = await self._submit_mutation(workbook_id, "sheet_copy", copy_fn, record)
        self._record_audit(
            "sheet_copy",
            "mutating",
            record,
            {"workbook_id": workbook_id, "source": name, "target_name": target_name},
            {"sheet": result.name, "index": result.index},
        )
        return result

    async def set_visibility(self, workbook_id: str, name: str, visible: bool) -> SheetInfo:
        """Set ``name`` visible/hidden in ``workbook_id``.

        Guarded mutation: writable session + snapshot + approval. Hiding the last
        visible worksheet is rejected with a structured error (also enforced by
        the adapter). Postcondition: the returned DTO reflects the new visibility.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_VISIBILITY)

        def visibility_fn(wb: Any) -> SheetInfo:
            return self._adapter.set_visibility(wb, name, visible)

        result = await self._submit_mutation(
            workbook_id, "sheet_set_visibility", visibility_fn, record
        )
        self._record_audit(
            "sheet_set_visibility",
            "mutating",
            record,
            {"workbook_id": workbook_id, "sheet": name, "visible": visible},
            {"sheet": result.name, "visibility": result.visibility.value},
        )
        return result

    async def delete_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        """Delete ``name`` from ``workbook_id``.

        The full destructive guard chain runs **before** the STA dispatch:

        1. writable session (``WORKBOOK_READ_ONLY``);
        2. snapshot (``SNAPSHOT_FAILED``);
        3. existence pre-check on the STA thread (a missing sheet is reported as
           "not found" before the approval gate);
        4. approval (deny-by-default, ``MUTATING``);
        5. visible-sheet guard — deleting the sheet must leave at least one
           visible worksheet, otherwise the delete is rejected. This is
           re-evaluated inside the mutation so a stale pre-check cannot race
           another mutation.

        Postcondition: the deleted sheet is gone and the returned DTO is the
        pre-delete snapshot of the removed sheet, with the remaining collection
        state verified on the STA thread.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)

        # Existence pre-check on the STA thread so a missing sheet is reported as
        # "not found" before the approval gate.
        def precheck_fn(wb: Any) -> SheetInfo:
            return self._adapter.get_sheet(wb, name)

        await self._submit_read(workbook_id, "sheet_delete_precheck", precheck_fn, record)

        self._require_approval(record, _ACTION_DELETE)

        def delete_fn(wb: Any) -> SheetInfo:
            self._guard_visible_sheet(wb, name)
            info = self._adapter.get_sheet(wb, name)
            # Excel shows a "permanently delete this sheet?" confirmation when
            # DisplayAlerts is True; with alerts suppressed the default (Yes) is
            # taken and the sheet is actually removed.
            with self._app_adapter().alerts_suppressed():
                wb.Sheets.Item(name).Delete()
            # Postcondition: the sheet is gone.
            remaining = {sheet.name for sheet in self._adapter.list_sheets(wb)}
            if info.name in remaining:
                raise ExcelVbaMcpError(
                    f"Postcondition failed: sheet {name!r} still exists after delete; "
                    f"remaining={sorted(remaining)}.",
                    context={"sheet": name, "remaining": sorted(remaining)},
                )
            return info

        result = await self._submit_mutation(workbook_id, "sheet_delete", delete_fn, record)
        self._record_audit(
            "sheet_delete",
            "mutating",
            record,
            {"workbook_id": workbook_id, "sheet": name},
            {"deleted": result.name, "index": result.index},
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
                "Sheet mutations require a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": workbook_id, "path": record.path},
            )
        return record

    def _require_snapshot(self, workbook_id: str) -> None:
        """Block the mutation with ``SNAPSHOT_FAILED`` when no snapshot exists."""

        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating sheet operation (deny-by-default)."""

        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"Sheet mutation {action!r} denied by approval policy: {decision.reason}",
                context={
                    "action": action,
                    "workbook_id": record.workbook_id,
                    "reason": decision.reason,
                },
            )

    def _guard_visible_sheet(self, wb: Any, name: str) -> None:
        """Reject deleting ``name`` when it is the last visible worksheet."""

        target = self._adapter.get_sheet(wb, name)
        if target.visibility.value != "visible":
            return
        if visible_worksheet_count(wb) <= 1:
            raise ExcelVbaMcpError(
                "Deleting the last visible worksheet is not allowed; at least "
                "one worksheet must remain visible.",
                context={"sheet": name},
            )

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
                context={"tool": "sheets"},
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


__all__ = ["SheetsService"]
