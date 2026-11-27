"""Immutable snapshot + restore orchestration (task_014, spec §9.1 / §9.4 / §6.4).

:class:`SnapshotService` is the application-layer orchestrator for workbook
snapshots. It owns **no COM proxies**: every Excel call is delegated to the
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` via
``await worker.submit(...)``. The pure-filesystem
:class:`~excel_vba_mcp.storage.snapshot_store.SnapshotStore` persists the
immutable snapshot + manifest.

Snapshot branches (research §5.1):

* **Open workbook** — ``Workbook.SaveCopyAs`` runs on the STA worker and captures
  the current in-memory state into a *separate* file without replacing the open
  workbook. The store then hashes/finalizes it (``snapshot_method="SaveCopyAs"``).
* **Closed workbook** — a plain file copy of the on-disk file
  (``snapshot_method="file_copy"``).

Restore (spec §9.4) closes the current working copy **without saving** and opens a
fresh working copy from the immutable snapshot baseline. It reverts the workbook
**file only** and never reverts external macro side effects — this is stated in
``external_side_effects_not_undone`` on every manifest and recorded in the audit
trail.

Pre-mutation gate (spec §9.1 / §6.4): :meth:`require_snapshot` raises
``SNAPSHOT_FAILED`` when a workbook has no successful snapshot, so mutating
entry points (VBA replace/insert, macro execution, destructive sheet ops, …) can
refuse to proceed until a snapshot exists. Wiring that guard into each mutation
tool is the responsibility of those tools (tasks 016/018/022/…); this service
exposes the guard and records the latest successful snapshot per workbook.
"""

from __future__ import annotations

import contextlib
import shutil
import tempfile
from pathlib import Path
from typing import Any, Literal
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    SnapshotFailedError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import SnapshotRestoreResult, WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

# Default per-call soft ceiling (seconds).
_DEFAULT_TIMEOUT_S = 60.0

# Manifest fields that are inherently volatile and excluded from a diff's
# added/removed/changed metadata report.
_VOLATILE_MANIFEST_FIELDS = frozenset({"snapshot_id", "created_at", "stored_path"})


class SnapshotDiff(BaseModel):
    """Structured comparison of two snapshots (spec §6.4 ``snapshot_diff``).

    Compares manifests and file hashes only — never a full binary diff.
    """

    model_config = ConfigDict(extra="forbid")

    snapshot_a: str
    snapshot_b: str
    sha256_a: str
    sha256_b: str
    file_identical: bool = Field(description="True when both snapshot files hash equal.")
    added: list[str] = Field(default_factory=list, description="Manifest keys only in B.")
    removed: list[str] = Field(default_factory=list, description="Manifest keys only in A.")
    changed: dict[str, dict[str, Any]] = Field(
        default_factory=dict,
        description="Manifest keys whose values differ: {key: {'a': ..., 'b': ...}}.",
    )
    notes: list[str] = Field(default_factory=list)


class SnapshotService:
    """Create / restore / diff / list immutable workbook snapshots.

    Args:
        store: Pure-filesystem snapshot store.
        session_service: The workbook session lifecycle service (owns the worker).
        worker: The single-STA COM worker (all Excel runs there).
        path_policy: Allowlist / canonicalisation gate for source paths.
        audit: Append-only JSONL audit trail.
        approval: Deny-by-default approval gate for the snapshot-creation
            filesystem write. When ``None`` (legacy/test construction only) the
            approval gate is a no-op; the production server always injects one
            (``server._build_runtime``).
    """

    def __init__(
        self,
        store: SnapshotStore,
        session_service: SessionService,
        worker: ExcelComWorker,
        path_policy: PathPolicy,
        audit: AuditLogger,
        approval: ApprovalPolicy | None = None,
    ) -> None:
        self._store = store
        self._session = session_service
        self._worker = worker
        self._path_policy = path_policy
        self._audit = audit
        self._approval = approval
        # The most recent successful snapshot per workbook id (pre-mutation gate).
        self._latest_snapshot: dict[str, SnapshotManifest] = {}
        # Last-known canonical source path per workbook id (closed-branch lookup).
        self._known_sources: dict[str, str] = {}
        # The snapshot id whose restored copy is currently active, keyed by the
        # rebound canonical (the ORIGINAL source path). Used for a deterministic
        # same-snapshot repeat error (plan_06 R3) instead of a generic COM_BUSY.
        self._restored_snapshot_by_canonical: dict[str, str] = {}

    # ------------------------------------------------------------------
    # Properties (non-COM)
    # ------------------------------------------------------------------
    @property
    def store(self) -> SnapshotStore:
        return self._store

    # ------------------------------------------------------------------
    # Create
    # ------------------------------------------------------------------
    async def create_snapshot(
        self,
        workbook_id: str,
        *,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SnapshotManifest:
        """Create an immutable snapshot of ``workbook_id``.

        Open workbooks are snapshotted with ``Workbook.SaveCopyAs`` (capturing the
        current in-memory state without replacing the open workbook); closed
        workbooks fall back to a file copy of the last-known source path.

        Args:
            workbook_id: The id of an open workbook, or a previously-seen id for
                the closed branch.
            timeout_s: Soft per-call ceiling for the STA SaveCopyAs.

        Returns:
            The written :class:`SnapshotManifest`.

        Raises:
            ApprovalDeniedError: If the approval gate denies the snapshot write
                (deny-by-default in ``restricted`` mode).
            ExcelVbaMcpError: If the workbook id is unknown.
            SnapshotFailedError: If the snapshot could not be produced.
        """

        info = await self._find_open(workbook_id)
        # F-4: a snapshot is a filesystem write (SaveCopyAs / file copy). Gate it
        # with the same deny-by-default approval policy the other mutating tools
        # use. A snapshot create cannot require a pre-existing snapshot, so only
        # the approval gate applies (no ``require_snapshot``).
        self._require_create_approval(workbook_id, info)
        snapshot_id = self._new_snapshot_id()

        if info is not None:
            manifest = await self._create_open_branch(
                snapshot_id, info, workbook_id=workbook_id, timeout_s=timeout_s
            )
        else:
            manifest = self._create_closed_branch(snapshot_id, workbook_id)

        self._latest_snapshot[workbook_id] = manifest
        self._known_sources[workbook_id] = manifest.source_path
        self._record_audit(
            tool="snapshot_create",
            operation_class="mutating",
            params={"workbook_id": workbook_id, "snapshot_id": snapshot_id},
            workbook={"canonical_path": manifest.source_path},
            snapshot_id=snapshot_id,
            result={
                "snapshot_id": snapshot_id,
                "sha256": manifest.sha256,
                "snapshot_method": manifest.snapshot_method,
            },
        )
        return manifest

    async def _create_open_branch(
        self,
        snapshot_id: str,
        info: WorkbookInfo,
        *,
        workbook_id: str,
        timeout_s: float,
    ) -> SnapshotManifest:
        """SaveCopyAs branch: capture the open workbook's in-memory state."""

        source_path = info.source_path or info.path
        working_path = info.path
        tmp_dir = Path(tempfile.mkdtemp(prefix="excel_vba_mcp_snap_"))
        try:
            tmp_save = tmp_dir / (Path(source_path).name or "snapshot.xlsm")
            tmp_exports = tmp_dir / "exports"
            app_adapter = self._app_adapter()

            def save_copy_as_fn() -> list[dict[str, Any]] | None:
                wb = _find_open_workbook(app_adapter, working_path)
                if wb is None:
                    raise WorkbookNotFoundError(
                        "The working copy is no longer open in the owned instance.",
                        context={"workbook_id": workbook_id, "working_path": working_path},
                    )
                # Suppress Excel's own alerts while SaveCopyAs writes the copy.
                with app_adapter.alerts_suppressed():
                    wb.SaveCopyAs(str(tmp_save))
                return _export_vba_components(wb, tmp_exports)

            save_copy_as_fn.__name__ = "save_copy_as"
            vba_exports = await self._worker.submit(
                save_copy_as_fn,
                timeout=timeout_s,
                section="diagnostic",
                op_class=OperationClass.READ_ONLY,
            )

            def place_copy(dest: Path) -> None:
                shutil.move(str(tmp_save), str(dest))
                if tmp_exports.is_dir():
                    shutil.move(str(tmp_exports), str(dest.parent / "exports"))

            return self._store.finalize_open_branch(
                snapshot_id,
                place_copy,
                source_path=source_path,
                excel_version=info.excel_version,
                workbook_model=_model_from_info(info),
                vba_exports=vba_exports,
            )
        except SnapshotFailedError:
            raise
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - wrap any failure as SNAPSHOT_FAILED
            raise SnapshotFailedError(
                f"Failed to create snapshot for workbook {workbook_id!r}.",
                context={"workbook_id": workbook_id, "error": str(exc)},
            ) from exc
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def _create_closed_branch(self, snapshot_id: str, workbook_id: str) -> SnapshotManifest:
        """Closed branch: file-copy the last-known on-disk source."""

        source_path = self._known_sources.get(workbook_id)
        if source_path is None:
            raise ExcelVbaMcpError(
                f"Unknown workbook_id {workbook_id!r} and no known source path.",
                context={"workbook_id": workbook_id},
            )
        try:
            return self._store.create(
                snapshot_id,
                source_path=source_path,
                excel_version=None,
                workbook_model={},
                vba_exports=None,
                method="file_copy",
            )
        except SnapshotFailedError:
            raise
        except Exception as exc:  # noqa: BLE001 - wrap any failure as SNAPSHOT_FAILED
            raise SnapshotFailedError(
                f"Failed to create snapshot for workbook {workbook_id!r}.",
                context={"workbook_id": workbook_id, "error": str(exc)},
            ) from exc

    # ------------------------------------------------------------------
    # Restore
    # ------------------------------------------------------------------
    async def restore_snapshot(
        self,
        snapshot_id: str,
        *,
        workbook_id: str | None = None,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SnapshotRestoreResult:
        """Restore a working session from an immutable snapshot.

        Restore adds **no second approval gate**: its approval point is inherited
        from the writable ``open_workbook`` call below (``SessionMode.EXECUTE`` →
        ``EXECUTING``), which is deny-by-default in ``restricted`` mode and
        auto-allowed in ``full`` mode. The snapshot baseline itself is only read,
        so no additional gate is needed.

        Lifecycle order (WP3; never close both copies, never zero): the fresh
        working copy is opened from the snapshot baseline **first**, while the
        previous working copy (when ``workbook_id`` is given and still open) is
        left untouched. The restored copy is then **validated** (a lightweight
        STA inspect confirms it is actually open in the owned instance). Only
        **after** validation succeeds is the obsolete previous working copy
        closed **without saving** (discarding its mutations). The restored copy
        remains the single active session record. If validation of the restored
        copy fails, the previous copy is kept open and a structured
        ``SNAPSHOT_FAILED`` error is raised (the failed restored copy is cleaned
        up best-effort so exactly one working copy — the previous one — remains).

        Additive hardening (plan_06; the WP3 order above is unchanged):

        * **R1 — ``previous_close`` status.** A failed previous-copy close never
          fails the restore and never invalidates the restored session: the
          result carries ``previous_close="failed"`` with safe ``previous_close_error``
          text and the previous copy remains open. ``workbook_id=None`` (or the
          restored id) yields ``previous_close="not_requested"``; a clean close
          yields ``"closed"``.
        * **R3 — identity rebind.** After a successful open+validate the restored
          record's ``canonical`` is rebound to ``manifest.source_path`` (the
          ORIGINAL) and its ``source_sha256`` to ``manifest.sha256`` (the
          baseline), so ``list_workbooks`` / reopen reflect the original, never
          the immutable snapshot file. The working copy stays ``work/<newid>.ext``.
        * **R3 — same-snapshot repeat.** Restoring the snapshot that is already
          the active restored baseline for the same original raises a
          deterministic validation error (``reason="same_snapshot"``,
          ``isRetryable=false``) instead of a generic ``COM_BUSY``. A different
          snapshot is allowed.

        The freshly opened restored session is **execute mode** and **writable**
        (``read_only=False``), matching the default ``excel_open``; this lets the
        Full-Edit workflow continue editing and running macros after a restore
        (an inspect-mode restore would disable macros and make ``macro_run`` fail
        with ``MACRO_BLOCKED``). The immutable snapshot file is never modified
        (the session service copies it into ``work/`` before opening).

        External macro side effects are **not** undone — only the workbook file is
        restored (spec §9.4).

        Args:
            snapshot_id: The snapshot to restore.
            workbook_id: The currently-open working copy to discard (no save) once
                the restored copy has been validated. ``None`` yields
                ``previous_close="not_requested"`` (no previous copy is closed).
            timeout_s: Soft per-call ceiling for the STA open/close/validate.

        Returns:
            A :class:`~excel_vba_mcp.domain.models.SnapshotRestoreResult` whose
            ``workbook`` is the freshly opened working copy in execute mode
            (``mode == SessionMode.EXECUTE``) and writable (``read_only is False``)
            with ``source_path`` rebound to the original, plus the
            ``previous_close`` disposition (R1).

        Raises:
            ApprovalDeniedError: If the inherited ``excel_open`` approval gate
                denies the writable open (restricted mode).
            SnapshotFailedError: If the snapshot does not exist, or the restored
                working copy could not be validated as open (the previous copy is
                kept open in that case).
            ExcelVbaMcpError: With ``code='INVALID_RANGE'`` and
                ``reason='same_snapshot'`` when the requested snapshot is already
                the active restored baseline for the same original (R3).
        """

        manifest = self._store.load(snapshot_id)
        snapshot_file = self._store.path_for(snapshot_id) / manifest.file_name

        # R3 (same-snapshot repeat): a deterministic validation error when the
        # requested snapshot is already the active restored baseline for the same
        # original. Raised BEFORE any open so no duplicate working copy is created.
        previous_canonical = self._canonical_for(workbook_id)
        if (
            previous_canonical is not None
            and self._restored_snapshot_by_canonical.get(previous_canonical) == snapshot_id
        ):
            raise ExcelVbaMcpError(
                "This snapshot is already the active restored baseline for this "
                "workbook; restore a different snapshot or close the restored copy "
                "first.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "reason": "same_snapshot",
                    "snapshot_id": snapshot_id,
                    "workbook_id": workbook_id,
                    "canonical_path": previous_canonical,
                },
            )

        # 1. Open the restored working copy FIRST, while the previous copy (if any)
        #    is still open. If this open fails, the previous copy is untouched and
        #    the error propagates — both copies are never closed.
        info = await self._session.open_workbook(
            str(snapshot_file),
            mode=SessionMode.EXECUTE,
            read_only=False,
            timeout_s=timeout_s,
        )
        # A freshly restored working copy is clean.
        with contextlib.suppress(Exception):
            info = info.model_copy(update={"source_path": manifest.source_path})

        # 2. Validate the restored working copy BEFORE closing the previous copy.
        #    On failure the previous copy is kept open; the just-opened (invalid)
        #    restored copy is discarded best-effort so exactly one copy remains.
        try:
            await self._validate_restored_copy(info, timeout_s=timeout_s)
        except ExcelVbaMcpError:
            with contextlib.suppress(ExcelVbaMcpError):
                await self._session.close_workbook(
                    info.workbook_id,
                    save_changes=False,
                    discard_changes=True,
                    timeout_s=timeout_s,
                )
            raise

        # 3. Validation succeeded: close the obsolete previous working copy
        #    WITHOUT saving (discard mutations). The restored copy remains the
        #    single active session record (never both, never zero). R1: a close
        #    failure is surfaced as previous_close="failed" and NEVER changes the
        #    restore outcome (still success) or invalidates the restored session.
        previous_close, previous_close_error = await self._close_previous_copy(
            workbook_id, restored_id=info.workbook_id, timeout_s=timeout_s
        )

        # 4. R3: rebind the restored session identity to the ORIGINAL (the
        #    manifest source path) so list_workbooks / reopen reflect the original,
        #    never the immutable snapshot file. The working copy stays
        #    work/<newid>.ext; the snapshot file is never the session identity.
        #    Done AFTER close-previous so the original canonical key is free (no
        #    collision in the session's per-path active/lock maps).
        self._session.rebind_restored_identity(
            info.workbook_id,
            canonical=manifest.source_path,
            source_sha256=manifest.sha256,
        )
        self._restored_snapshot_by_canonical[manifest.source_path] = snapshot_id

        self._latest_snapshot[info.workbook_id] = manifest
        self._known_sources[info.workbook_id] = manifest.source_path
        self._record_audit(
            tool="snapshot_restore",
            operation_class="mutating",
            params={"snapshot_id": snapshot_id, "workbook_id": workbook_id},
            workbook={"canonical_path": manifest.source_path},
            snapshot_id=snapshot_id,
            result={
                "restored_workbook_id": info.workbook_id,
                "sha256": manifest.sha256,
                "external_side_effects_not_undone": True,
                "previous_close": previous_close,
                "previous_close_error": previous_close_error,
            },
        )
        return SnapshotRestoreResult(
            workbook=info,
            previous_close=previous_close,
            previous_close_error=previous_close_error,
        )

    async def _close_previous_copy(
        self,
        workbook_id: str | None,
        *,
        restored_id: str,
        timeout_s: float,
    ) -> tuple[Literal["closed", "failed", "not_requested"], str | None]:
        """Close the previous working copy without saving; never fail the restore (R1).

        Returns ``(previous_close, previous_close_error)``:

        * ``("not_requested", None)`` when no distinct previous copy was supplied
          (``workbook_id is None`` or it equals the restored id);
        * ``("closed", None)`` when the previous copy was closed without saving;
        * ``("failed", <safe text>)`` when the close raised — the restore outcome
          is unchanged (still success) and the previous copy remains open.

        A close failure is caught broadly (never leaks out) so a transient COM
        failure on close cannot turn a successful restore into an error.
        """

        if workbook_id is None or workbook_id == restored_id:
            return "not_requested", None
        try:
            await self._session.close_workbook(
                workbook_id,
                save_changes=False,
                discard_changes=True,
                timeout_s=timeout_s,
            )
        except Exception as exc:  # noqa: BLE001 - R1: a close failure never fails restore
            return "failed", _safe_error_text(exc)
        return "closed", None

    def _canonical_for(self, workbook_id: str | None) -> str | None:
        """Return the rebound canonical (original source path) for an open workbook id."""

        if workbook_id is None:
            return None
        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info.source_path or info.path
        return None

    async def _validate_restored_copy(self, info: WorkbookInfo, *, timeout_s: float) -> None:
        """Confirm the restored working copy is actually open in the owned instance.

        Runs a lightweight, read-only inspect on the single-STA worker: it locates
        the restored workbook by its working path (``wb.FullName``) and touches a
        harmless property to prove the instance responds. It raises
        :class:`~excel_vba_mcp.domain.errors.SnapshotFailedError` (a structured
        ``SNAPSHOT_FAILED``) when the restored copy cannot be found or inspected,
        so :meth:`restore_snapshot` never closes the previous working copy until
        the restored one is known-good.
        """

        app_adapter = self._app_adapter()
        working_path = info.path

        def inspect_restored_fn() -> bool:
            wb = _find_open_workbook(app_adapter, working_path)
            if wb is None:
                return False
            # Touch a harmless property to confirm the instance actually responds.
            _ = wb.FullName
            return True

        inspect_restored_fn.__name__ = "validate_restored"
        try:
            valid = await self._worker.submit(
                inspect_restored_fn,
                timeout=timeout_s,
                section="diagnostic",
                op_class=OperationClass.READ_ONLY,
            )
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - wrap any inspect failure
            raise SnapshotFailedError(
                "The restored working copy could not be validated as open.",
                context={
                    "restored_workbook_id": info.workbook_id,
                    "working_path": working_path,
                    "error": str(exc),
                },
            ) from exc
        if not valid:
            raise SnapshotFailedError(
                "The restored working copy is not open in the owned instance.",
                context={
                    "restored_workbook_id": info.workbook_id,
                    "working_path": working_path,
                },
            )

    # ------------------------------------------------------------------
    # Diff / list
    # ------------------------------------------------------------------
    async def diff_snapshots(self, snapshot_id_a: str, snapshot_id_b: str) -> SnapshotDiff:
        """Compare two snapshots by manifest metadata and file hash.

        Args:
            snapshot_id_a: First snapshot id.
            snapshot_id_b: Second snapshot id.

        Returns:
            A :class:`SnapshotDiff`.

        Raises:
            SnapshotFailedError: If either snapshot does not exist.
        """

        a = self._store.load(snapshot_id_a)
        b = self._store.load(snapshot_id_b)
        return _diff_manifests(a, b)

    async def list_snapshots(self) -> list[SnapshotManifest]:
        """Return every snapshot manifest known to the store."""

        manifests = self._store.list()
        self._record_audit(
            tool="snapshot_list",
            operation_class="read_only",
            params={},
            result={"count": len(manifests)},
        )
        return manifests

    # ------------------------------------------------------------------
    # Pre-mutation gate (spec §9.1 / §6.4)
    # ------------------------------------------------------------------
    def require_snapshot(self, workbook_id: str) -> SnapshotManifest:
        """Return the latest successful snapshot for ``workbook_id`` or raise.

        Mutation entry points (VBA replace/insert, macro execution, destructive
        sheet operations, …) must call this **before** mutating a workbook. When
        no snapshot has been created for the workbook it raises
        :class:`~excel_vba_mcp.domain.errors.SnapshotFailedError`
        (``SNAPSHOT_FAILED``), which blocks the mutation.

        Args:
            workbook_id: The workbook about to be mutated.

        Returns:
            The latest successful :class:`SnapshotManifest` for the workbook.

        Raises:
            SnapshotFailedError: If no successful snapshot exists for the workbook.
        """

        manifest = self._latest_snapshot.get(workbook_id)
        if manifest is None:
            raise SnapshotFailedError(
                "A successful snapshot is required before mutating this workbook.",
                context={"workbook_id": workbook_id},
            )
        return manifest

    def has_snapshot(self, workbook_id: str) -> bool:
        """Return whether a successful snapshot exists for ``workbook_id``."""

        return workbook_id in self._latest_snapshot

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    async def _find_open(self, workbook_id: str) -> WorkbookInfo | None:
        for info in await self._session.list_workbooks():
            if info.workbook_id == workbook_id:
                return info
        return None

    def _app_adapter(self) -> ExcelAppAdapter:
        # The session service owns the single hidden Excel application adapter;
        # the snapshot service reuses it (never creates a second instance).
        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "snapshot_create"},
            )
        return adapter

    @staticmethod
    def _new_snapshot_id() -> str:
        return f"snap-{uuid4().hex}"

    def _require_create_approval(self, workbook_id: str, info: WorkbookInfo | None) -> None:
        """Deny-by-default approval gate for the snapshot filesystem write (F-4).

        A snapshot create is a ``MUTATING`` operation for approval purposes: it
        writes a snapshot file (and optional VBA exports). In ``restricted`` mode
        with no allowlist entry / live grant it raises ``APPROVAL_DENIED``; in the
        default ``full`` mode the policy auto-allows. When no policy is wired
        (legacy/test construction) the gate is a no-op.
        """

        if self._approval is None:
            return
        if info is not None:
            canonical = info.source_path or info.path
        else:
            canonical = self._known_sources.get(workbook_id, "")
        key = ApprovalKey(canonical_path=canonical, action="snapshot_create")
        decision = self._approval.check(
            key,
            session_id=getattr(self._session, "session_id", None),
            operation_class=OperationClass.MUTATING,
        )
        if not decision.allowed:
            raise ApprovalDeniedError(
                f"snapshot_create denied by approval policy: {decision.reason}",
                context={
                    "action": key.action,
                    "canonical_path": canonical,
                    "reason": decision.reason,
                },
            )

    def _record_audit(
        self,
        *,
        tool: str,
        operation_class: str,
        params: dict[str, Any],
        workbook: dict[str, Any] | None = None,
        snapshot_id: str | None = None,
        result: dict[str, Any] | None = None,
    ) -> None:
        session_id = getattr(self._session, "session_id", None)
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=session_id,
            workbook=workbook,
            snapshot_id=snapshot_id,
            result=result,
        )


# ---------------------------------------------------------------------------
# Module-level helpers (no COM at import time)
# ---------------------------------------------------------------------------
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


def _safe_error_text(exc: Exception) -> str:
    """Return safe, human-readable text for a close failure (R1; no secrets).

    Uses the domain ``message`` when available (never the context, which may hold
    paths), falling back to ``str(exc)`` for a generic exception, capped to a
    reasonable length so a volatile COM string cannot bloat the result.
    """

    text = getattr(exc, "message", None) or str(exc) or exc.__class__.__name__
    return text[:500]


def _model_from_info(info: WorkbookInfo) -> dict[str, Any]:
    """Build a compact, serializable workbook model from a :class:`WorkbookInfo`."""

    return {
        "name": info.name,
        "format": info.format,
        "path": info.path,
        "source_path": info.source_path,
        "has_vba": info.has_vba,
        "excel_version": info.excel_version,
        "read_only": info.read_only,
    }


def _export_vba_components(wb: Any, exports_dir: Path) -> list[dict[str, Any]] | None:
    """Best-effort VBA component export to ``exports_dir`` (STA thread only).

    Returns a list of ``{"component_name", "type", "path", "export_file_sha256"}``
    entries, or ``None`` when VBIDE is unavailable / access is denied / the project
    is locked. Never raises: exporting VBA is a best-effort enrichment of the
    manifest.

    Each ``export_file_sha256`` is the SHA-256 of the **exported file bytes** (the
    ``.bas``/``.cls``/``.frm`` written by ``VBComponent.Export``). It is a
    *file* hash and is **not** interchangeable with the normalized
    ``source_hash`` produced by ``vba_read_component``
    (:func:`~excel_vba_mcp.adapters.excel.vbide_adapter.hash_code`), which is the
    only token the VBA mutation guards accept as ``expected_source_hash`` (F7).
    """

    try:
        project = wb.VBProject
        components = project.VBComponents
    except Exception:  # noqa: BLE001 - VBOM disabled / locked / no project
        return None

    exports_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    try:
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            ctype = int(component.Type)
            name = str(component.Name)
            extension = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(ctype, ".txt")
            target = exports_dir / f"{name}{extension}"
            component.Export(str(target))
            entry: dict[str, Any] = {
                "component_name": name,
                "type": ctype,
                "path": f"exports/{target.name}",
                "export_file_sha256": _sha256(target),
            }
            frx = target.with_suffix(".frx")
            if frx.is_file():
                entry["frx"] = f"exports/{frx.name}"
                entry["frx_sha256"] = _sha256(frx)
            results.append(entry)
    except Exception:  # noqa: BLE001 - a locked/partial export degrades to partial
        return results or None
    return results


def _sha256(path: Path) -> str:
    import hashlib

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _diff_manifests(a: SnapshotManifest, b: SnapshotManifest) -> SnapshotDiff:
    """Compare two manifests + hashes (no full binary diff)."""

    dump_a = a.model_dump(mode="json")
    dump_b = b.model_dump(mode="json")
    keys = (set(dump_a) | set(dump_b)) - _VOLATILE_MANIFEST_FIELDS
    added = sorted(k for k in keys if k not in dump_a)
    removed = sorted(k for k in keys if k not in dump_b)
    changed: dict[str, dict[str, Any]] = {}
    for key in sorted(keys):
        if key in dump_a and key in dump_b and dump_a[key] != dump_b[key]:
            changed[key] = {"a": dump_a[key], "b": dump_b[key]}

    notes: list[str] = []
    if a.sha256 == b.sha256:
        notes.append("Snapshot files are byte-identical (SHA-256 match).")
    else:
        notes.append("Snapshot files differ (SHA-256 mismatch).")
    return SnapshotDiff(
        snapshot_a=a.snapshot_id,
        snapshot_b=b.snapshot_id,
        sha256_a=a.sha256,
        sha256_b=b.sha256,
        file_identical=a.sha256 == b.sha256,
        added=added,
        removed=removed,
        changed=changed,
        notes=notes,
    )


__all__ = ["SnapshotDiff", "SnapshotService"]
