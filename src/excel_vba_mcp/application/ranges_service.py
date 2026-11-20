"""Excel Range/Cells read service (task_042, spec §14 Phase 2 — Range/Cells).

:class:`RangesService` is the async application-layer orchestration for the
read-only range surface (``range_get``, ``range_read_values``,
``range_read_formulas``). It is deliberately **COM-free**: every method resolves
the ``workbook_id`` through the session bookkeeping, enforces the non-COM bounds
and then submits a single STA-local callable to the dedicated single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`. Only domain
DTOs cross the worker boundary (ARCH-009).

Read-only path: no snapshot / approval / audit guard chain is required (those
belong to the mutating range tools, task_043+). Every read is still recorded in
the append-only audit trail as ``read_only``.
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
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import RangeAddress, RangeFormulas, RangeValues, WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

#: Default per-call soft ceiling (seconds) for a range operation on the STA worker.
_DEFAULT_TIMEOUT_S = 60.0

#: Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class RangesService:
    """Async Excel Range metadata/value/formula reads (no COM here).

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (owns ``workbook_id``
            bookkeeping and the owned app adapter).
        snapshot_service: Present for guard-chain symmetry with the mutating
            range tools (task_043+); unused by the read-only surface.
        approval: Deny-by-default approval gate (unused by reads).
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
        record: WorkbookInfo,
        params: dict[str, Any],
        result: dict[str, Any],
    ) -> None:
        canonical = ""
        with contextlib.suppress(ExcelVbaMcpError):
            canonical = self._resolve_record(record.workbook_id).source_path or ""
        self._audit.record(
            tool=tool,
            operation_class="read_only",
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
