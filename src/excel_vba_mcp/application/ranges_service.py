"""Excel Range/Cells service (task_042 read, task_043 write, task_044 format,
task_045 structural).

:class:`RangesService` is the async application-layer orchestration for the
Range/Cells surface (``range_get``, ``range_read_values``,
``range_read_formulas``, ``range_write_values``, ``range_write_formulas``,
``range_get_format``, ``range_set_format``, ``range_copy``, ``range_insert``,
``range_delete``, ``range_merge``, ``range_unmerge``). It is
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
from excel_vba_mcp.adapters.excel.ranges_adapter import RangesAdapter, validate_format_patch
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    FormatInfo,
    FormatMutationResult,
    FormulaWriteRequest,
    RangeAddress,
    RangeFormatPatch,
    RangeFormulas,
    RangeMutationResult,
    RangeValues,
    RangeWriteRequest,
    StructuralMutationResult,
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
_ACTION_SET_FORMAT = "range_set_format"
_ACTION_COPY = "range_copy"
_ACTION_INSERT = "range_insert"
_ACTION_DELETE = "range_delete"
_ACTION_MERGE = "range_merge"
_ACTION_UNMERGE = "range_unmerge"

#: Explicit shift directions accepted by ``insert_range`` / ``delete_range``
#: (task_038 §5.2 item 2). The adapter resolves the names to the Excel enum
#: values; this set only validates the *contract* before any COM access.
_INSERT_SHIFTS: frozenset[str] = frozenset({"right", "down"})
_DELETE_SHIFTS: frozenset[str] = frozenset({"left", "up"})

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

    async def get_format(self, workbook_id: str, sheet: str, address: str) -> FormatInfo:
        """Read the bounded, tri-state formatting of a Range (read-only).

        Returns a :class:`~excel_vba_mcp.domain.models.FormatInfo` whose values
        are tri-state (``value`` / ``mixed`` / ``unsupported``): a range with
        mixed formatting lists the affected properties in ``mixed_properties``
        instead of flattening them to one cell's value (task_038 §4.2).
        """

        record = self._resolve_record(workbook_id)
        result = await self._submit_read(
            workbook_id,
            "range_get_format",
            lambda wb: self._adapter.get_format(wb, sheet, address),
            record,
        )
        self._record_audit(
            "range_get_format",
            "read_only",
            record,
            {"workbook_id": workbook_id, "sheet": sheet, "address": address},
            {
                "sheet": result.range.sheet,
                "address": result.range.address,
                "mixed_properties": result.mixed_properties,
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

    async def set_format(
        self, workbook_id: str, sheet: str, address: str, patch: RangeFormatPatch
    ) -> FormatMutationResult:
        """Apply a partial formatting update to a bounded Range (guarded mutation).

        Only the fields the caller explicitly set (``patch.model_fields_set``) are
        written; absent fields are left untouched (task_038 §4.3). The patch is
        validated before any COM mutation. Requires a writable working copy, a
        successful snapshot and an approval grant; the targeted properties are
        re-read after the write and reported through ``verified``.
        """

        # Validate the patch first (pure, pre-COM) so malformed input is rejected
        # before the guard chain and before the STA worker touches Excel.
        validate_format_patch(patch)

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_SET_FORMAT)

        def apply_fn(wb: Any) -> FormatMutationResult:
            return self._adapter.set_format(wb, sheet, address, patch)

        result = await self._submit_mutation(workbook_id, "range_set_format", apply_fn, record)
        self._record_audit(
            "range_set_format",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "applied": result.applied,
            },
            {
                "action": result.action,
                "sheet": result.range.sheet,
                "address": result.range.address,
                "applied": result.applied,
                "verified": result.verified,
            },
        )
        return result

    # ------------------------------------------------------------------
    # Public API — structural operations (guarded; task_045)
    # ------------------------------------------------------------------
    # Copy / insert / delete / merge / unmerge all follow the same guard chain
    # (writable -> snapshot -> approval), then run one STA-local adapter call that
    # performs the preflight guards (ListObject intersection, merged cells,
    # protected sheet, spill ranges, merge data loss) and verifies the resulting
    # address / dimensions / table metadata. Every call is audited.
    async def copy_range(
        self,
        workbook_id: str,
        sheet: str,
        source: str,
        destination_sheet: str,
        destination: str,
    ) -> StructuralMutationResult:
        """Copy a range to an explicit destination (no clipboard).

        The destination is mandatory: a missing/blank destination or source sheet
        is rejected before the guard chain and before any COM access.
        """

        _require_text(source, "source")
        _require_text(destination, "destination")
        _require_text(destination_sheet, "destination_sheet")
        _require_text(sheet, "sheet")

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_COPY)

        def copy_fn(wb: Any) -> StructuralMutationResult:
            return self._adapter.copy_range(wb, sheet, source, destination_sheet, destination)

        result = await self._submit_mutation(workbook_id, "range_copy", copy_fn, record)
        self._record_audit(
            "range_copy",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "source": source,
                "destination_sheet": destination_sheet,
                "destination": destination,
            },
            {
                "action": result.action,
                "source": result.source.address if result.source else source,
                "destination": result.range_after.address if result.range_after else destination,
                "verified": result.verified,
                "affected_addresses": result.affected_addresses,
            },
        )
        return result

    async def insert_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        """Insert cells and shift them in an explicit direction (guarded mutation)."""

        _require_text(sheet, "sheet")
        _require_text(address, "address")
        shift_name = _require_shift(shift, _INSERT_SHIFTS, "insert")

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_INSERT)

        def insert_fn(wb: Any) -> StructuralMutationResult:
            return self._adapter.insert_range(
                wb, sheet, address, shift_name, allow_merged=allow_merged
            )

        result = await self._submit_mutation(workbook_id, "range_insert", insert_fn, record)
        self._record_audit(
            "range_insert",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "shift": shift_name,
                "allow_merged": allow_merged,
            },
            {
                "action": result.action,
                "shift": result.shift,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
                "preconditions": result.preconditions,
            },
        )
        return result

    async def delete_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        """Delete cells and shift the remainder in an explicit direction."""

        _require_text(sheet, "sheet")
        _require_text(address, "address")
        shift_name = _require_shift(shift, _DELETE_SHIFTS, "delete")

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_DELETE)

        def delete_fn(wb: Any) -> StructuralMutationResult:
            return self._adapter.delete_range(
                wb, sheet, address, shift_name, allow_merged=allow_merged
            )

        result = await self._submit_mutation(workbook_id, "range_delete", delete_fn, record)
        self._record_audit(
            "range_delete",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "shift": shift_name,
                "allow_merged": allow_merged,
            },
            {
                "action": result.action,
                "shift": result.shift,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
                "preconditions": result.preconditions,
            },
        )
        return result

    async def merge_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        across: bool = False,
        allow_data_loss: bool = False,
    ) -> StructuralMutationResult:
        """Merge a range into one cell (guarded, preview-gated mutation).

        Merging discards non-anchor data; the adapter previews the affected cells
        and blocks the merge with ``INVALID_RANGE`` unless ``allow_data_loss`` is
        set. The flag is validated as a real boolean before the guard chain.
        """

        _require_text(sheet, "sheet")
        _require_text(address, "address")
        _require_bool(across, "across")
        _require_bool(allow_data_loss, "allow_data_loss")

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_MERGE)

        def merge_fn(wb: Any) -> StructuralMutationResult:
            return self._adapter.merge_range(
                wb, sheet, address, across=across, allow_data_loss=allow_data_loss
            )

        result = await self._submit_mutation(workbook_id, "range_merge", merge_fn, record)
        self._record_audit(
            "range_merge",
            "mutating",
            record,
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "across": across,
                "allow_data_loss": allow_data_loss,
            },
            {
                "action": result.action,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
                "affected_addresses": result.affected_addresses,
                "preconditions": result.preconditions,
            },
        )
        return result

    async def unmerge_range(
        self, workbook_id: str, sheet: str, address: str
    ) -> StructuralMutationResult:
        """Unmerge any merged area inside an explicit target (guarded mutation)."""

        _require_text(sheet, "sheet")
        _require_text(address, "address")

        record = self._require_writable(workbook_id)
        self._require_snapshot(workbook_id)
        self._require_approval(record, _ACTION_UNMERGE)

        def unmerge_fn(wb: Any) -> StructuralMutationResult:
            return self._adapter.unmerge_range(wb, sheet, address)

        result = await self._submit_mutation(workbook_id, "range_unmerge", unmerge_fn, record)
        self._record_audit(
            "range_unmerge",
            "mutating",
            record,
            {"workbook_id": workbook_id, "sheet": sheet, "address": address},
            {
                "action": result.action,
                "address": result.range_after.address if result.range_after else address,
                "verified": result.verified,
                "affected_addresses": result.affected_addresses,
                "preconditions": result.preconditions,
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


# ---------------------------------------------------------------------------
# Structural argument validation (pure; task_045)
# ---------------------------------------------------------------------------
def _require_text(value: object, field: str) -> str:
    """Reject a blank/non-string address or sheet argument before any COM access."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} is required and must be a non-empty string.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "value": value},
        )
    return value.strip()


def _require_shift(shift: object, allowed: frozenset[str], operation: str) -> str:
    """Validate an explicit shift name (Excel never guesses it from the shape)."""

    if not isinstance(shift, str) or not shift.strip():
        raise ExcelVbaMcpError(
            f"Argument 'shift' is required for {operation} and must never be guessed.",
            code=ErrorCode.INVALID_RANGE,
            context={"shift": shift, "operation": operation},
        )
    normalized = shift.strip().lower()
    if normalized not in allowed:
        raise ExcelVbaMcpError(
            f"Argument 'shift' must be one of {sorted(allowed)} for {operation}.",
            code=ErrorCode.INVALID_RANGE,
            context={"shift": shift, "operation": operation, "allowed": sorted(allowed)},
        )
    return normalized


def _require_bool(value: object, field: str) -> bool:
    """Reject a non-boolean flag before the guard chain (defense in depth)."""

    if not isinstance(value, bool):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a boolean.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "value": value},
        )
    return value


__all__ = ["RangesService"]
