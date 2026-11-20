"""Excel Range/Cells service (task_042 read, task_043 write, spec §14 Phase 2 — Range/Cells).

:class:`RangesService` is the async application-layer orchestration for the
Range/Cells surface (``range_get``, ``range_read_values``,
``range_read_formulas``, ``range_write_values``, ``range_write_formulas``). It is
deliberately **COM-free**: every method resolves the ``workbook_id`` through the
session bookkeeping, enforces the non-COM guard chain and then submits a single
STA-local callable to the dedicated single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only domain
DTOs cross the worker boundary (ARCH-009).

Read-only path (task_042): no snapshot / approval / audit guard chain is
required; every read is recorded in the append-only audit trail as ``read_only``.

Mutating path (task_043, deny-by-default, spec §7.6 / §9.1):

1. **Writable session** — a read-only working copy is rejected with
   ``WORKBOOK_READ_ONLY``.
2. **Snapshot** — ``snapshot_service.require_snapshot`` blocks with
   ``SNAPSHOT_FAILED`` when no successful snapshot exists.
3. **Approval** — ``ApprovalPolicy.authorize_vba_write`` (a ``MUTATING``
   classification) denies unless a live grant/allowlist entry matches the action +
   workbook identity.

The mutating adapter call captures before/after metadata (a
:class:`~excel_vba_mcp.domain.models.RangeMutationResult`) whose postcondition is
the re-read value/formula grid, so the returned DTO reflects the actual
post-mutation Range; every tool invocation is recorded in the audit trail.
"""

from __future__ import annotations

import contextlib
from collections.abc import Callable
from typing import Any, TypeVar, cast

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.ranges_adapter import RangesAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    FormulaWriteRequest,
    RangeAddress,
    RangeFormulas,
    RangeMutationResult,
    RangeValues,
    RangeWriteRequest,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

#: Default per-call soft ceiling (seconds) for a range operation on the STA worker.
_DEFAULT_TIMEOUT_S = 60.0

#: Approval action names (stable identities for the range mutation surface).
_ACTION_WRITE_VALUES = "range_write_values"
_ACTION_WRITE_FORMULAS = "range_write_formulas"

#: Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class RangesService:
    """Async Excel Range metadata/value/formula reads and guarded writes (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        snapshot_service: Pre-mutation snapshot gate for the mutating tools.
        approval: Deny-by-default approval gate for mutating range operations.
        path_policy: Allowlist policy (used to validate the working-copy path
            identity for approval keys of the mutating tools).
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
        self._adapter = RangesAdapter()

    # ------------------------------------------------------------------
    # Public API — read-only (async; no COM here)
    # ------------------------------------------------------------------
    async def get_range(self, workbook_id: str, sheet: str, address: str) -> RangeAddress:
        """Resolve an explicit worksheet + A1 address to Range metadata."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "range_get",
            lambda wb: self._adapter.resolve_range(wb, sheet, address),
            record,
        )
        self._record_audit(
            "range_get",
            "read_only",
            record,
            {"workbook_id": workbook_id, "sheet": sheet, "address": address},
            {
                "sheet": result.sheet,
                "address": result.address,
                "rows": result.rows,
                "cols": result.cols,
            },
        )
        return result

    async def read_values(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeValues:
        """Read typed cell values from a bounded Range window (read-only)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "range_read_values",
            lambda wb: self._adapter.read_values(wb, sheet, address, offset=offset, limit=limit),
            record,
        )
        self._record_audit(
            "range_read_values",
            "read_only",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "offset": offset,
                "limit": limit,
            },
            {
                "sheet": result.range.sheet,
                "address": result.range.address,
                "rows": result.rows,
                "cols": result.cols,
                "truncated": result.truncated,
            },
        )
        return result

    async def read_formulas(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeFormulas:
        """Read typed formulas from a bounded Range window (``Formula2`` primary)."""

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "range_read_formulas",
            lambda wb: self._adapter.read_formulas(wb, sheet, address, offset=offset, limit=limit),
            record,
        )
        self._record_audit(
            "range_read_formulas",
            "read_only",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "offset": offset,
                "limit": limit,
            },
            {
                "sheet": result.range.sheet,
                "address": result.range.address,
                "rows": result.rows,
                "cols": result.cols,
                "truncated": result.truncated,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — mutations (guarded; task_043)
    # ------------------------------------------------------------------
    async def write_values(
        self, workbook_id: str, sheet: str, address: str, payload: RangeWriteRequest
    ) -> RangeMutationResult:
        """Write typed cell values to a bounded Range (guarded mutation).

        The payload shape is validated against the target before any COM
        mutation; a wrong-sized payload is rejected with ``INVALID_RANGE``.
        Requires a writable working copy, a successful snapshot and an approval
        grant, and verifies the written values by re-reading them.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_WRITE_VALUES)

        def write_fn(wb: Any) -> RangeMutationResult:
            return self._adapter.write_values(
                wb, sheet, address, payload.rows, payload.cols, payload.values
            )

        result = await self._submit_mutation(workbook_id, "range_write_values", write_fn, record)
        self._record_audit(
            "range_write_values",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "rows": payload.rows,
                "cols": payload.cols,
            },
            {
                "action": result.action,
                "sheet": result.range_after.sheet if result.range_after else sheet,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
            },
        )
        return result

    async def write_formulas(
        self, workbook_id: str, sheet: str, address: str, payload: FormulaWriteRequest
    ) -> RangeMutationResult:
        """Write formulas (``Formula2``) to a bounded Range (guarded mutation).

        The payload shape is validated against the target before any COM
        mutation; a wrong-sized payload or a non-formula entry is rejected with
        ``INVALID_RANGE``. Requires a writable working copy, a successful
        snapshot and an approval grant, and verifies the written formulas by
        re-reading them.
        """

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_WRITE_FORMULAS)

        def write_fn(wb: Any) -> RangeMutationResult:
            return self._adapter.write_formulas(
                wb, sheet, address, payload.rows, payload.cols, payload.formulas
            )

        result = await self._submit_mutation(workbook_id, "range_write_formulas", write_fn, record)
        self._record_audit(
            "range_write_formulas",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "rows": payload.rows,
                "cols": payload.cols,
            },
            {
                "action": result.action,
                "sheet": result.range_after.sheet if result.range_after else sheet,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
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
                "Range mutations require a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": workbook_id, "path": record.path},
            )
        return record

    def _require_snapshot(self, workbook_id: str) -> None:
        """Block the mutation with ``SNAPSHOT_FAILED`` when no snapshot exists."""

        self._snapshot.require_snapshot(workbook_id)

    def _require_approval(self, record: WorkbookInfo, action: str) -> None:
        """Authorize a mutating range operation (deny-by-default)."""

        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=None,
            action=action,
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ExcelVbaMcpError(
                f"Range mutation {action!r} denied by approval policy: {decision.reason}",
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
                context={"tool": "ranges"},
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


__all__ = ["RangesService"]
