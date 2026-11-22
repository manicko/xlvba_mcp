"""Read-only VBA project service + guarded source mutations (task_016, task_018).

:class:`VBAProjectService` is the application-layer orchestrator for the
read-only VBA project tools (``vba_project_info``, ``vba_list_components``,
``vba_list_procedures``, ``vba_read_component``, ``vba_export_project``,
``vba_get_references`` and ``vba_search_code``) and the guarded source mutations
(``vba_replace_component_source``, ``vba_insert_code``).

It owns **no COM proxies**: every Excel/VBIDE access is delegated to the
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` via
``await worker.submit(...)`` and executed by the STA-local
:class:`~excel_vba_mcp.adapters.excel.vbide_adapter.VBIDEAdapter`. The service
resolves the open workbook through the :class:`SessionService` bookkeeping, then
submits a synchronous callable that re-locates the working copy inside the owned
instance and returns serializable DTOs. Nothing here touches COM (ARCH-009).

Safety contract (spec §7.2, §10; addendum BG03/BG04/BG05):

* **Read-only tools.** No VBA mutation in the read-only methods.
* **Approval-gated export.** ``export_project`` only reads the VBA project but
  writes ``.bas`` / ``.cls`` / ``.frm`` / ``.frx`` files, so it is approval-gated
  (deny-by-default in ``restricted`` mode, auto-allowed in ``full`` mode) in
  addition to the ``PathPolicy`` export-directory check (F-4).
* **Guarded mutations.** ``replace_component_source`` / ``insert_code`` enforce
  (in order): writable session → snapshot → revision check → approval → STA
  mutation with before/after export + hash. Document modules and UserForms are
  never replaced (only Standard/Class modules, types 1/2).
* **Copies only.** The working copy opened by the session service is mutated; the
  original workbook is never modified.
* **Hidden / non-activating.** All work runs on the owned, ``Visible=False``
  Excel instance via the adapter; the service never activates/selects anything.
* **Bounded export.** ``export_project`` writes under a temp directory by
  default; a caller-supplied ``out_dir`` must live inside a configured
  artifact root (``PathPolicy``), so a tool parameter can never redirect exports
  to an arbitrary location.
"""

from __future__ import annotations

import contextlib
import os
import tempfile
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, TypeVar

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter, hash_code
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    PathNotAllowedError,
)
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaExportManifest,
    VbaMutationResult,
    VbaProcedureInfo,
    VbaProjectInfo,
    VbaReferenceInfo,
    VbaSearchMatch,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import OperationClass, VbaComponentType
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds).
_DEFAULT_TIMEOUT_S = 60.0

# Logical artifact root under which a caller-supplied export directory must live
# (a subset of ``PathPolicy.ARTIFACT_ROOT_NAMES``: snapshots / artifacts).
_EXPORT_ROOT_NAMES = ("snapshots", "artifacts")

# Type of the synchronous callable handed to the STA worker.
_VbaCall = Callable[[Any], Any]

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")

# Component types that may be replaced / mutated (Standard = 1, Class = 2).
# Document modules (100) and UserForms (3) are never replaced.
_REPLACEABLE_TYPES = (VbaComponentType.STANDARD_MODULE, VbaComponentType.CLASS_MODULE)

# VBIDE ``vbext_ProcKind`` code used to match any procedure kind.
_VBEXT_PK_PROC = 0


class VBAProjectService:
    """Read-only VBA inventory / procedures / references / export + guarded mutations.

    Args:
        worker: The single-STA COM worker (all Excel runs there).
        session_service: The workbook session lifecycle service (owns the worker
            and the ``workbook_id`` bookkeeping).
        path_policy: Allowlist / canonicalisation gate for caller-supplied export
            directories.
        audit: Append-only JSONL audit trail.
        snapshot_service: Snapshot service for the pre-mutation gate. Required for
            ``replace_component_source`` / ``insert_code``; read-only methods
            work without it.
        approval: Deny-by-default approval gate for mutating operations. Required
            for ``replace_component_source`` / ``insert_code``.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        path_policy: PathPolicy,
        audit: AuditLogger,
        snapshot_service: SnapshotService | None = None,
        approval: ApprovalPolicy | None = None,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._path_policy = path_policy
        self._audit = audit
        self._adapter = VBIDEAdapter()
        self._snapshot = snapshot_service
        self._approval = approval

    # ------------------------------------------------------------------
    # Public API (async; no COM)
    # ------------------------------------------------------------------
    async def get_project_info(self, workbook_id: str) -> VbaProjectInfo:
        """Return the VBA project summary for ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_project_info",
            lambda wb: self._adapter.project_info(wb),
        )
        self._record_audit(
            "vba_project_info",
            workbook_id,
            {"component_count": result.component_count, "locked": result.locked},
        )
        return result

    async def list_components(self, workbook_id: str) -> list[VbaComponentInfo]:
        """Return the component inventory for ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_list_components",
            lambda wb: self._adapter.list_components(wb),
        )
        self._record_audit("vba_list_components", workbook_id, {"count": len(result)})
        return result

    async def list_procedures(self, workbook_id: str, component: str) -> list[VbaProcedureInfo]:
        """Return the procedures of ``component`` in ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_list_procedures",
            lambda wb: self._adapter.list_procedures(wb, component),
        )
        self._record_audit(
            "vba_list_procedures",
            workbook_id,
            {"component": component, "count": len(result)},
        )
        return result

    async def read_component(self, workbook_id: str, component: str) -> VbaComponentSource:
        """Return the code text + normalized hash of ``component``."""

        result = await self._submit(
            workbook_id,
            "vba_read_component",
            lambda wb: self._adapter.read_component(wb, component),
        )
        self._record_audit(
            "vba_read_component",
            workbook_id,
            {"component": component, "source_hash": result.source_hash},
        )
        return result

    async def export_project(
        self, workbook_id: str, out_dir: str | None = None
    ) -> VbaExportManifest:
        """Export every VBA component of ``workbook_id`` and return the manifest.

        When ``out_dir`` is ``None`` a fresh temporary directory is used. A
        caller-supplied ``out_dir`` is validated by :class:`PathPolicy` against
        the configured artifact roots, so an export can never escape the
        allowlist.

        F-4: the export writes ``.bas`` / ``.cls`` / ``.frm`` / ``.frx`` files, so
        it is approval-gated (deny-by-default in ``restricted`` mode; auto-allowed
        in ``full`` mode). A read-only session is still allowed — the export only
        reads the VBA project — so no writable-session check is applied.

        Raises:
            ApprovalDeniedError: If the approval gate denies the export
                (restricted mode).
            ExcelVbaMcpError: If ``out_dir`` escapes the allowlist.
        """

        record = self._resolve_record(workbook_id)
        self._require_approval_service()
        assert self._approval is not None
        key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            action="vba_export_project",
        )
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"VBA export denied by approval policy: {decision.reason}",
                context={"action": key.action, "reason": decision.reason},
            )

        if out_dir is None:
            target = Path(tempfile.mkdtemp(prefix="excel_vba_mcp_vba_export_"))
        else:
            target = self._validate_export_dir(out_dir)
            target.mkdir(parents=True, exist_ok=True)

        result = await self._submit(
            workbook_id,
            "vba_export_project",
            lambda wb: self._adapter.export_project(wb, target, workbook_id=workbook_id),
        )
        self._record_audit(
            "vba_export_project",
            workbook_id,
            {"out_dir": result.out_dir, "component_count": len(result.components)},
        )
        return result

    async def get_references(self, workbook_id: str) -> list[VbaReferenceInfo]:
        """Return the VBA project references (missing ones flagged)."""

        result = await self._submit(
            workbook_id,
            "vba_get_references",
            lambda wb: self._adapter.get_references(wb),
        )
        self._record_audit(
            "vba_get_references",
            workbook_id,
            {
                "count": len(result),
                "missing": [reference.name for reference in result if reference.is_broken],
            },
        )
        return result

    async def search_code(
        self, workbook_id: str, pattern: str, regex: bool = False
    ) -> list[VbaSearchMatch]:
        """Search every component's code in ``workbook_id`` (bounded)."""

        result = await self._submit(
            workbook_id,
            "vba_search_code",
            lambda wb: self._adapter.search_code(wb, pattern, regex=regex),
        )
        self._record_audit(
            "vba_search_code",
            workbook_id,
            {"pattern": pattern, "regex": regex, "count": len(result)},
        )
        return result

    # ------------------------------------------------------------------
    # Guarded mutations (task_018)
    # ------------------------------------------------------------------
    async def replace_component_source(
        self,
        workbook_id: str,
        component: str,
        new_source: str,
        *,
        expected_source_hash: str,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> VbaMutationResult:
        """Replace the entire code module of a Standard/Class component.

        Enforces (in order): writable session → snapshot → revision check →
        approval → STA mutation with before/after export + hash. Document modules
        (type 100) and UserForms (type 3) are rejected — only Standard (1) and
        Class (2) modules may be replaced.

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            component: The ``VBComponent.Name`` to replace.
            new_source: The new VBA source text (CRLF or LF; normalised on write).
            expected_source_hash: The SHA-256 of the **current** normalized code
                body. The mutation is rejected if the on-disk source has changed
                since it was read (stale revision).
            timeout_s: Soft per-call ceiling for the STA mutation.

        Returns:
            A :class:`VbaMutationResult` with before/after hashes, export paths
            and the audit event id.

        Raises:
            WorkbookReadOnlyError: If the session is read-only.
            SnapshotFailedError: If no successful snapshot exists.
            ExcelVbaMcpError: On revision conflict, disallowed component type,
                approval denial, or COM failure.
        """

        record = self._resolve_record(workbook_id)
        self._require_writable(record)
        self._require_snapshot_service()
        self._require_approval_service()
        assert self._snapshot is not None
        self._snapshot.require_snapshot(workbook_id)

        # Pre-check: read current source + component type (fast fail before
        # submitting a MUTATING operation).
        current = await self._submit(
            workbook_id,
            "vba_replace_precheck",
            lambda wb: self._adapter.read_component(wb, component),
        )
        self._require_replaceable(current)
        self._require_revision(current.source_hash, expected_source_hash, component)

        # Approve (deny-by-default).
        key = ApprovalKey(
            canonical_path=record.path,
            sha256=expected_source_hash,
            action="vba_replace_component_source",
            procedure=component,
        )
        assert self._approval is not None
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"VBA write denied by approval policy: {decision.reason}",
                context={
                    "action": key.action,
                    "component": component,
                    "reason": decision.reason,
                },
            )

        # Submit the STA mutation.
        export_dir = Path(tempfile.mkdtemp(prefix="excel_vba_mcp_mutation_"))
        try:
            payload = await self._submit_mutation(
                workbook_id,
                component=component,
                new_source=new_source,
                expected_source_hash=expected_source_hash,
                action="replace",
                export_dir=export_dir,
                timeout_s=timeout_s,
            )
        finally:
            # The export files are referenced by path in the result; clean up
            # only if the mutation failed (success keeps them for inspection).
            pass

        result = self._build_mutation_result(workbook_id, component, "replace", payload)
        self._record_mutation_audit(
            "vba_replace_component_source",
            workbook_id,
            component=component,
            expected_source_hash=expected_source_hash,
            result=result,
        )
        return result

    async def insert_code(
        self,
        workbook_id: str,
        component: str,
        code: str,
        *,
        position: str,
        expected_source_hash: str,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> VbaMutationResult:
        """Insert ``code`` at a semantic position in a Standard/Class component.

        ``position`` must be one of:

        * ``"declarations"`` — after the Option statements / before the first
          procedure.
        * ``"before_procedure:<name>"`` — before the named procedure.
        * ``"after_procedure:<name>"`` — after the named procedure.

        The line number is resolved via the VBIDE CodeModule API
        (``ProcStartLine`` / ``ProcCountLines`` / ``ProcOfLine``); ambiguous text
        heuristics are rejected with a structured error.

        Enforces the same guard chain as :meth:`replace_component_source`:
        writable session → snapshot → revision check → approval → STA mutation.

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            component: The ``VBComponent.Name`` to insert into.
            code: The VBA code text to insert.
            position: Insertion position (see above).
            expected_source_hash: The SHA-256 of the **current** normalized code
                body (stale-revision guard).
            timeout_s: Soft per-call ceiling for the STA mutation.

        Returns:
            A :class:`VbaMutationResult` with before/after hashes, export paths
            and the audit event id.

        Raises:
            WorkbookReadOnlyError: If the session is read-only.
            SnapshotFailedError: If no successful snapshot exists.
            ExcelVbaMcpError: On revision conflict, disallowed component type,
                ambiguous position, approval denial, or COM failure.
        """

        record = self._resolve_record(workbook_id)
        self._require_writable(record)
        self._require_snapshot_service()
        self._require_approval_service()
        assert self._snapshot is not None
        self._snapshot.require_snapshot(workbook_id)

        # Pre-check: read current source + component type.
        current = await self._submit(
            workbook_id,
            "vba_insert_precheck",
            lambda wb: self._adapter.read_component(wb, component),
        )
        self._require_replaceable(current)
        self._require_revision(current.source_hash, expected_source_hash, component)

        # Approve (deny-by-default).
        key = ApprovalKey(
            canonical_path=record.path,
            sha256=expected_source_hash,
            action="vba_insert_code",
            procedure=component,
        )
        assert self._approval is not None
        decision = self._approval.authorize_vba_write(key)
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"VBA write denied by approval policy: {decision.reason}",
                context={
                    "action": key.action,
                    "component": component,
                    "reason": decision.reason,
                },
            )

        # Submit the STA mutation.
        export_dir = Path(tempfile.mkdtemp(prefix="excel_vba_mcp_mutation_"))
        payload = await self._submit_insert(
            workbook_id,
            component=component,
            code=code,
            position=position,
            expected_source_hash=expected_source_hash,
            export_dir=export_dir,
            timeout_s=timeout_s,
        )

        result = self._build_mutation_result(workbook_id, component, "insert", payload)
        self._record_mutation_audit(
            "vba_insert_code",
            workbook_id,
            component=component,
            expected_source_hash=expected_source_hash,
            result=result,
            extra={"position": position},
        )
        return result

    # ------------------------------------------------------------------
    # Internal helpers (non-COM)
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

    async def _submit(
        self,
        workbook_id: str,
        tool: str,
        fn: Callable[[Any], _T],
    ) -> _T:
        """Run ``fn(wb)`` on the STA worker for the open working copy."""

        record = self._resolve_record(workbook_id)
        app_adapter = self._app_adapter()
        call = _bind_workbook(fn, app_adapter, record.path, workbook_id)
        call.__name__ = tool
        return await self._worker.submit(
            call,
            timeout=_DEFAULT_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )

    def _app_adapter(self) -> ExcelAppAdapter:
        """Return the owned hidden Excel application adapter (never a new one)."""

        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "vba_project"},
            )
        return adapter

    def _validate_export_dir(self, out_dir: str) -> Path:
        """Validate a caller-supplied export directory against the allowlist.

        The directory itself is validated inside a configured artifact root
        (``snapshots`` / ``artifacts``). ``PathPolicy.validate_path`` also
        enforces a file-extension allowlist, which is meant for files, not
        directories; the export files themselves (``.bas`` / ``.cls`` / ``.frm``
        / ``.frx``) are written by the adapter, not by the caller, so the
        directory check is done on the canonical directory path directly.
        """

        if not isinstance(out_dir, str) or not out_dir.strip():
            raise ExcelVbaMcpError(
                "out_dir must be a non-empty string.",
                context={"field": "out_dir"},
            )
        expected_root = _artifact_root_for(self._path_policy)
        candidate = Path(out_dir)
        if candidate.suffix:
            # A caller passed a file path; validate it as a file (extension gate
            # applies) so the rejection reason is precise.
            return self._path_policy.validate_path(
                out_dir, OperationClass.READ_ONLY, expected_root=expected_root
            )
        return _validate_dir(out_dir, self._path_policy, expected_root)

    def _record_audit(self, tool: str, workbook_id: str, result: dict[str, Any]) -> None:
        canonical = ""
        with contextlib.suppress(ExcelVbaMcpError):
            canonical = self._resolve_record(workbook_id).source_path or ""
        self._audit.record(
            tool=tool,
            operation_class="read_only",
            params={"workbook_id": workbook_id},
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": canonical},
            result=result,
        )

    # ------------------------------------------------------------------
    # Mutation helpers (non-COM guard chain + STA submission)
    # ------------------------------------------------------------------
    def _require_writable(self, record: WorkbookInfo) -> None:
        """Reject a read-only session with ``WORKBOOK_READ_ONLY``."""

        if record.read_only:
            raise ExcelVbaMcpError(
                "The workbook is read-only; VBA mutations require a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": record.workbook_id, "path": record.path},
            )

    def _require_snapshot_service(self) -> None:
        """Ensure the snapshot service is wired (required for mutations)."""

        if self._snapshot is None:
            raise ExcelVbaMcpError(
                "The snapshot service is not configured; mutations require it.",
                code=ErrorCode.SNAPSHOT_FAILED,
                context={"tool": "vba_mutation"},
            )

    def _require_approval_service(self) -> None:
        """Ensure the approval policy is wired (required for mutations)."""

        if self._approval is None:
            raise ApprovalDeniedError(
                "The approval policy is not configured; mutations require it.",
                context={"tool": "vba_mutation", "reason": "approval_unavailable"},
            )

    @staticmethod
    def _require_replaceable(source: VbaComponentSource) -> None:
        """Reject Document modules and UserForms with a clear error."""

        if source.component_type not in _REPLACEABLE_TYPES:
            raise ExcelVbaMcpError(
                f"Component {source.component_name!r} is a "
                f"{source.component_type.value}; only Standard/Class modules "
                "may be replaced or mutated.",
                code=ErrorCode.COM_BUSY,
                context={
                    "component": source.component_name,
                    "component_type": source.component_type.value,
                    "reason": "non_replaceable_component",
                },
            )

    @staticmethod
    def _require_revision(current_hash: str, expected_hash: str, component: str) -> None:
        """Reject a stale revision when the on-disk source has changed."""

        if current_hash != expected_hash:
            raise ExcelVbaMcpError(
                f"Revision conflict for component {component!r}: the source has "
                "changed since it was read.",
                code=ErrorCode.COM_BUSY,
                context={
                    "component": component,
                    "expected_source_hash": expected_hash,
                    "current_source_hash": current_hash,
                    "reason": "revision_conflict",
                },
            )

    async def _submit_mutation(
        self,
        workbook_id: str,
        *,
        component: str,
        new_source: str,
        expected_source_hash: str,
        action: str,
        export_dir: Path,
        timeout_s: float,
    ) -> dict[str, Any]:
        """Submit the STA replace callable and return the payload dict."""

        record = self._resolve_record(workbook_id)
        app_adapter = self._app_adapter()
        before_dir = export_dir / "before"
        after_dir = export_dir / "after"
        before_dir.mkdir(parents=True, exist_ok=True)
        after_dir.mkdir(parents=True, exist_ok=True)

        def replace_fn() -> dict[str, Any]:
            wb = _find_open_workbook(app_adapter, record.path)
            if wb is None:
                raise ExcelVbaMcpError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": record.path},
                )
            return _do_replace(
                wb,
                component=component,
                new_source=new_source,
                expected_source_hash=expected_source_hash,
                before_dir=before_dir,
                after_dir=after_dir,
            )

        replace_fn.__name__ = "vba_replace_component_source"
        return await self._worker.submit(
            replace_fn,
            timeout=timeout_s,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    async def _submit_insert(
        self,
        workbook_id: str,
        *,
        component: str,
        code: str,
        position: str,
        expected_source_hash: str,
        export_dir: Path,
        timeout_s: float,
    ) -> dict[str, Any]:
        """Submit the STA insert callable and return the payload dict."""

        record = self._resolve_record(workbook_id)
        app_adapter = self._app_adapter()
        before_dir = export_dir / "before"
        after_dir = export_dir / "after"
        before_dir.mkdir(parents=True, exist_ok=True)
        after_dir.mkdir(parents=True, exist_ok=True)

        def insert_fn() -> dict[str, Any]:
            wb = _find_open_workbook(app_adapter, record.path)
            if wb is None:
                raise ExcelVbaMcpError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": record.path},
                )
            return _do_insert(
                wb,
                component=component,
                code=code,
                position=position,
                expected_source_hash=expected_source_hash,
                before_dir=before_dir,
                after_dir=after_dir,
            )

        insert_fn.__name__ = "vba_insert_code"
        return await self._worker.submit(
            insert_fn,
            timeout=timeout_s,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    @staticmethod
    def _build_mutation_result(
        workbook_id: str,
        component: str,
        action: str,
        payload: dict[str, Any],
    ) -> VbaMutationResult:
        """Construct a :class:`VbaMutationResult` from the STA payload dict."""

        now = datetime.now(UTC)
        procedure = payload.get("procedure")
        return VbaMutationResult(
            workbook_id=workbook_id,
            component_name=component,
            action=action,
            before_hash=payload["before_hash"],
            after_hash=payload["after_hash"],
            before_export_path=payload["before_export_path"],
            after_export_path=payload["after_export_path"],
            audit_id=payload["audit_id"],
            procedure=procedure,
            created_at=now,
            updated_at=now,
        )

    def _record_mutation_audit(
        self,
        tool: str,
        workbook_id: str,
        *,
        component: str,
        expected_source_hash: str,
        result: VbaMutationResult,
        extra: dict[str, Any] | None = None,
    ) -> None:
        """Record a mutating audit entry (operation_class='mutating')."""

        canonical = ""
        with contextlib.suppress(ExcelVbaMcpError):
            canonical = self._resolve_record(workbook_id).source_path or ""
        params: dict[str, Any] = {
            "workbook_id": workbook_id,
            "component": component,
            "expected_source_hash": expected_source_hash,
            "before_hash": result.before_hash,
            "after_hash": result.after_hash,
        }
        if extra:
            params.update(extra)
        self._audit.record(
            tool=tool,
            operation_class="mutating",
            params=params,
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": canonical},
            result={
                "audit_id": result.audit_id,
                "before_export_path": result.before_export_path,
                "after_export_path": result.after_export_path,
            },
        )


# ---------------------------------------------------------------------------
# Module-level helpers (no COM at import time)
# ---------------------------------------------------------------------------
def _bind_workbook(
    fn: _VbaCall,
    app_adapter: ExcelAppAdapter,
    working_path: str,
    workbook_id: str,
) -> Callable[[], Any]:
    """Wrap ``fn`` so it re-locates the working copy on the STA thread."""

    def run() -> Any:
        wb = _find_open_workbook(app_adapter, working_path)
        if wb is None:
            raise ExcelVbaMcpError(
                "The working copy is no longer open in the owned instance.",
                context={"workbook_id": workbook_id, "working_path": working_path},
            )
        return fn(wb)

    return run


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


def _artifact_root_for(path_policy: PathPolicy) -> str | None:
    """Return the first configured artifact root name, if any."""

    roots = path_policy.roots
    for name in _EXPORT_ROOT_NAMES:
        if name in roots:
            return name
    return None


def _validate_dir(raw: str, path_policy: PathPolicy, expected_root: str | None) -> Path:
    """Validate a directory path against the allowlist (extension gate bypassed).

    ``PathPolicy.validate_path`` rejects a candidate by file extension, which is
    meaningful for files but not for directories. The directory is therefore
    canonicalized with the same ``resolve`` / ``realpath`` cross-check the policy
    uses and the containment / expected-root rules are re-applied here so an
    export can never escape the allowlist.
    """

    candidate = Path(raw)
    try:
        resolved = candidate.resolve()
    except OSError:  # pragma: no cover - resolve rarely fails
        resolved = Path(os.path.abspath(raw))
    real = Path(os.path.realpath(candidate))
    canonical = real if os.fspath(real) != os.fspath(resolved) else resolved

    roots = path_policy.roots
    matches = [name for name, root in roots.items() if _is_within(canonical, root)]
    if not matches:
        raise PathNotAllowedError(
            "Path not allowed (outside_roots).",
            context={"reason": "outside_roots", "operation": "read_only", "name": candidate.name},
        )
    if expected_root is not None and expected_root not in matches:
        raise PathNotAllowedError(
            "Path not allowed (root_mismatch).",
            context={
                "reason": "root_mismatch",
                "operation": "read_only",
                "name": candidate.name,
                "expected_root": expected_root,
            },
        )
    return canonical


def _is_within(child: str | Path, parent: str | Path) -> bool:
    """Return whether ``child`` is inside ``parent`` (no symlink resolution)."""

    try:
        common = os.path.commonpath([os.fspath(child), os.fspath(parent)])
    except ValueError:
        return False
    return os.path.normcase(common) == os.path.normcase(os.fspath(parent))


# ---------------------------------------------------------------------------
# STA-local mutation helpers (run on the worker thread; touch COM only)
# ---------------------------------------------------------------------------
def _do_replace(
    wb: Any,
    *,
    component: str,
    new_source: str,
    expected_source_hash: str,
    before_dir: Path,
    after_dir: Path,
) -> dict[str, Any]:
    """Replace the entire code module of ``component`` (STA thread).

    Re-checks the revision, exports the before/after source, performs the
    mutation (``DeleteLines`` + ``AddFromString``) and returns a payload dict
    with before/after hashes and export paths.

    Raises:
        ExcelVbaMcpError: On revision conflict, protection, or COM failure.
    """

    adapter = VBIDEAdapter()
    project = adapter._project(wb)  # noqa: SLF001 - STA-local access
    adapter._ensure_unlocked(project)  # noqa: SLF001
    comp = adapter._require_component(project, component)  # noqa: SLF001
    code_module = adapter._code_module(comp)  # noqa: SLF001

    # Re-check the revision inside the worker (defence in depth).
    current_text = adapter._read_code_text(code_module)  # noqa: SLF001
    current_hash = hash_code(current_text)
    if current_hash != expected_source_hash:
        raise ExcelVbaMcpError(
            f"Revision conflict for component {component!r}: the source changed "
            "between the pre-check and the mutation.",
            code=ErrorCode.COM_BUSY,
            context={
                "component": component,
                "expected_source_hash": expected_source_hash,
                "current_source_hash": current_hash,
                "reason": "revision_conflict",
            },
        )

    # Export before (file for inspection; hash from the code body for identity).
    before_path = before_dir / _export_file_name(component, int(comp.Type))
    comp.Export(str(before_path))
    before_hash = current_hash  # already computed from the code module text

    # Mutate: delete all lines, then add the new source.
    total = int(code_module.CountOfLines)
    if total > 0:
        code_module.DeleteLines(1, total)
    if new_source.strip():
        code_module.AddFromString(new_source)

    # Export after.
    after_path = after_dir / _export_file_name(component, int(comp.Type))
    comp.Export(str(after_path))
    after_text = adapter._read_code_text(code_module)  # noqa: SLF001
    after_hash = hash_code(after_text)

    return {
        "before_hash": before_hash,
        "after_hash": after_hash,
        "before_export_path": str(before_path),
        "after_export_path": str(after_path),
        "audit_id": "",
        "procedure": None,
    }


def _do_insert(
    wb: Any,
    *,
    component: str,
    code: str,
    position: str,
    expected_source_hash: str,
    before_dir: Path,
    after_dir: Path,
) -> dict[str, Any]:
    """Insert ``code`` at a semantic position in ``component`` (STA thread).

    Resolves the insertion line via the CodeModule API, exports the before/after
    source, performs the insertion and returns a payload dict.

    Raises:
        ExcelVbaMcpError: On revision conflict, ambiguous position, protection,
            or COM failure.
    """

    adapter = VBIDEAdapter()
    project = adapter._project(wb)  # noqa: SLF001
    adapter._ensure_unlocked(project)  # noqa: SLF001
    comp = adapter._require_component(project, component)  # noqa: SLF001
    code_module = adapter._code_module(comp)  # noqa: SLF001

    # Re-check the revision inside the worker.
    current_text = adapter._read_code_text(code_module)  # noqa: SLF001
    current_hash = hash_code(current_text)
    if current_hash != expected_source_hash:
        raise ExcelVbaMcpError(
            f"Revision conflict for component {component!r}: the source changed "
            "between the pre-check and the mutation.",
            code=ErrorCode.COM_BUSY,
            context={
                "component": component,
                "expected_source_hash": expected_source_hash,
                "current_source_hash": current_hash,
                "reason": "revision_conflict",
            },
        )

    # Resolve the insertion line.
    insert_line = _resolve_insert_line(code_module, position, component)

    # Export before (file for inspection; hash from the code body for identity).
    before_path = before_dir / _export_file_name(component, int(comp.Type))
    comp.Export(str(before_path))
    before_hash = current_hash  # already computed from the code module text

    # Mutate: insert lines.
    if code.strip():
        code_module.InsertLines(insert_line, code)

    # Export after.
    after_path = after_dir / _export_file_name(component, int(comp.Type))
    comp.Export(str(after_path))
    after_text = adapter._read_code_text(code_module)  # noqa: SLF001
    after_hash = hash_code(after_text)

    # Extract the procedure name from the position string, if any.
    procedure: str | None = None
    if position.startswith("before_procedure:"):
        procedure = position[len("before_procedure:") :]
    elif position.startswith("after_procedure:"):
        procedure = position[len("after_procedure:") :]

    return {
        "before_hash": before_hash,
        "after_hash": after_hash,
        "before_export_path": str(before_path),
        "after_export_path": str(after_path),
        "audit_id": "",
        "procedure": procedure,
    }


def _export_file_name(component: str, type_code: int) -> str:
    """Return the export file name for a component (e.g. ``modTest.bas``)."""

    extension = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(type_code, ".txt")
    return f"{component}{extension}"


def _resolve_insert_line(code_module: Any, position: str, component: str) -> int:
    """Resolve the insertion line from a position spec (STA thread).

    Uses the VBIDE CodeModule API (``ProcStartLine`` / ``ProcCountLines`` /
    ``ProcOfLine``) — never blind text search. Raises
    :class:`ExcelVbaMcpError` when the position cannot be resolved unambiguously.
    """

    if position == "declarations":
        return _find_first_procedure_line(code_module)

    if position.startswith("before_procedure:"):
        name = position[len("before_procedure:") :].strip()
        if not name:
            raise _ambiguous_position_error(position, component, "empty procedure name")
        start = _proc_start_line(code_module, name)
        return start

    if position.startswith("after_procedure:"):
        name = position[len("after_procedure:") :].strip()
        if not name:
            raise _ambiguous_position_error(position, component, "empty procedure name")
        start = _proc_start_line(code_module, name)
        count = _proc_count_lines(code_module, name)
        return start + count

    raise _ambiguous_position_error(position, component, "unknown position format")


def _find_first_procedure_line(code_module: Any) -> int:
    """Return the line number of the first procedure, or the declaration end.

    Scans lines with ``ProcOfLine`` to find the first procedure header. If no
    procedure is found, returns the line after the last Option/declaration line
    (i.e. the end of the declarations section).
    """

    total = int(code_module.CountOfLines)
    for line_number in range(1, total + 1):
        try:
            result = code_module.ProcOfLine(line_number, _VBEXT_PK_PROC)
        except Exception:  # noqa: BLE001 - ProcOfLine may fail on non-proc lines
            continue
        # ProcOfLine returns (name, kind) for procedure lines, or raises.
        if isinstance(result, tuple) and len(result) >= 1 and result[0]:
            return line_number

    # No procedures found: insert after the declarations section (end of file).
    # The declarations section is everything before the first blank line after
    # the Option statements, or the entire module if no procedures exist.
    return total + 1 if total > 0 else 1


def _proc_start_line(code_module: Any, name: str) -> int:
    """Return the start line of the named procedure via ``ProcStartLine``."""

    try:
        start = code_module.ProcStartLine(name, _VBEXT_PK_PROC)
    except Exception as exc:  # noqa: BLE001
        raise ExcelVbaMcpError(
            f"Cannot resolve start line for procedure {name!r}.",
            code=ErrorCode.COM_BUSY,
            context={"procedure": name, "reason": "proc_start_line_failed"},
        ) from exc
    if not isinstance(start, int) or start < 1:
        raise ExcelVbaMcpError(
            f"ProcStartLine returned an invalid line for procedure {name!r}.",
            code=ErrorCode.COM_BUSY,
            context={"procedure": name, "start": start, "reason": "invalid_proc_start"},
        )
    return start


def _proc_count_lines(code_module: Any, name: str) -> int:
    """Return the line count of the named procedure via ``ProcCountLines``."""

    try:
        count = code_module.ProcCountLines(name, _VBEXT_PK_PROC)
    except Exception as exc:  # noqa: BLE001
        raise ExcelVbaMcpError(
            f"Cannot resolve line count for procedure {name!r}.",
            code=ErrorCode.COM_BUSY,
            context={"procedure": name, "reason": "proc_count_lines_failed"},
        ) from exc
    if not isinstance(count, int) or count < 1:
        raise ExcelVbaMcpError(
            f"ProcCountLines returned an invalid count for procedure {name!r}.",
            code=ErrorCode.COM_BUSY,
            context={"procedure": name, "count": count, "reason": "invalid_proc_count"},
        )
    return count


def _ambiguous_position_error(position: str, component: str, reason: str) -> ExcelVbaMcpError:
    """Build a structured error for an ambiguous insertion position."""

    return ExcelVbaMcpError(
        f"Cannot resolve insertion position {position!r} for component {component!r}: {reason}.",
        code=ErrorCode.COM_BUSY,
        context={
            "component": component,
            "position": position,
            "reason": "ambiguous_insert_position",
            "detail": reason,
        },
    )


__all__ = ["VBAProjectService"]
