"""Workbook session lifecycle service (task_012).

:class:`SessionService` is the application-layer orchestrator for the base
workbook/session tools (``excel_open``, ``excel_close``, ``excel_save``,
``excel_get_status``, ``excel_list_workbooks``). It owns **no COM proxies**: all
COM access is delegated to the single-STA :class:`ExcelComWorker` via
``await worker.submit(...)``. The service only builds STA-local callables and
serializable DTOs; the ``api/*`` layer awaits this service and never touches COM
(ARCH-009).

Lifecycle guarantees (spec §7.1, addendum BG03/BG04/BG08):

* **Copies, never the original.** ``open_workbook`` copies the allowlisted
  original into a per-workbook working file under ``work/`` and opens *that*
  copy. The original file is never modified.
* **Hidden / non-activating.** The owned Excel instance is always
  ``Visible=False`` (owned by the adapter); this service never activates/selects
  and only changes ``AutomationSecurity`` / ``DisplayAlerts`` inside the
  adapter's ``finally``-restoring context managers, on the STA thread.
* **Mode → macro trust.** ``inspect`` opens with ``AutomationSecurity =
  ForceDisable`` (3) so workbook macros never run; ``execute`` applies
  ``ByUI`` (2) immediately before ``Open`` (a trusted working copy + an approved
  context is required by the approval gate).
* **Explicit saves only.** ``close_workbook`` defaults to ``SaveChanges=False``;
  changes are persisted only by an explicit ``save_workbook`` (or
  ``close_workbook(..., save_changes=True)``).
* **Per-path locking.** A per-canonical-path ``asyncio.Lock`` serializes opens of
  the same file and prevents duplicate concurrent opens.
* **Deny-by-default approval.** Mutating/executing opens require an allowlist
  entry or live grant; inspect read-only opens are auto-allowed by policy.
"""

from __future__ import annotations

import asyncio
import contextlib
import dataclasses
import hashlib
import os
import shutil
import tempfile
from collections.abc import Callable, Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any
from uuid import uuid4

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker, WorkerStatus
from excel_vba_mcp.adapters.excel.workbook_adapter import WorkbookAdapter
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelHungError,
    ExcelVbaMcpError,
    ProtectedTargetError,
    StaleTargetError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import (
    SaveAsResult,
    SessionPolicy,
    SessionStatus,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.macro_policy import automation_security_for
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.owned_pid_store import (
    OwnedPidStore,
    build_owned_record,
)

if TYPE_CHECKING:
    from excel_vba_mcp.application.job_registry import JobRegistry

# Default per-call soft ceilings (seconds) — see task_010 research §Q2.
_DEFAULT_TIMEOUT_S = 60.0

# Working-copy subtree name (spec §7.1 "рабочие копии в work/").
_WORK_SUBDIR = "work"


@dataclasses.dataclass
class _SessionRecord:
    """Bookkeeping for one open workbook (no COM proxies)."""

    workbook_id: str
    canonical: str
    working_path: str
    mode: SessionMode
    read_only: bool
    dirty: bool
    info: WorkbookInfo
    lock: asyncio.Lock
    # Source identity captured at open (U7; plan_05 §9): the streaming SHA-256
    # and byte size of the canonical original, taken AFTER path validation and
    # BEFORE the working copy is materialized. ``save_as_workbook`` re-verifies
    # both before any source-overwrite so a drifted source is fail-closed.
    source_sha256: str
    source_size: int
    # Monotonic workbook content version, bumped on each successful mutation;
    # reads never bump it (WP-A large-data version primitive).
    version: int = 0


def _utc_now() -> datetime:
    return datetime.now(UTC)


def _sha256_file(path: Path) -> tuple[str, int]:
    """Return ``(sha256_hex, size_bytes)`` of ``path`` (streaming, never whole-file)."""

    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


class SessionService:
    """Hidden workbook open/close/save + session state; owns no COM proxies.

    Args:
        worker: The single-STA COM worker (all COM runs there).
        app_adapter: Hidden, non-activating application wrapper (STA thread).
        path_policy: Allowlist / canonicalisation gate for the original path.
        approval: Deny-by-default approval gate.
        audit: Append-only JSONL audit trail (redacted).
        work_dir: Directory for working copies. Defaults to a ``work/`` folder
            under the system temp directory so the repository is never polluted;
            the server may override it with a configured artifact root.
        open_timeout_s: Server-side deadline (seconds) for the ``excel_open`` STA
            open (Block E). This is the per-op-class server authority for the open
            path.
        owned_pid_dir: Directory for the startup-reconcile ownership marker. When
            ``None`` it defaults to ``work_dir`` (isolated temp in tests); the
            server passes a stable per-user directory so the marker survives a
            restart / logoff-logon.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        app_adapter: ExcelAppAdapter,
        path_policy: PathPolicy,
        approval: ApprovalPolicy,
        audit: AuditLogger,
        *,
        work_dir: Path | None = None,
        open_timeout_s: float = _DEFAULT_TIMEOUT_S,
        owned_pid_dir: Path | None = None,
    ) -> None:
        self._worker = worker
        self._app_adapter = app_adapter
        self._path_policy = path_policy
        self._approval = approval
        self._audit = audit
        self._open_timeout_s = open_timeout_s
        self._workbook_adapter = WorkbookAdapter()
        self._session_id = str(uuid4())
        self._session_mode: SessionMode = SessionMode.INSPECT
        # True once close_session() has released the owned Excel instance; a
        # closed session refuses new work until a new worker is started.
        self._closed = False
        # False while recover_session() is rebuilding the instance, so no new
        # work slips in behind the recovery.
        self._accepting = True
        self._records: dict[str, _SessionRecord] = {}
        self._active_by_canonical: dict[str, str] = {}
        # Per-canonical-path locks, created once per path and reused so a second
        # open of the same file serializes behind the first (held for the
        # workbook lifetime; released by close_workbook).
        self._locks: dict[str, asyncio.Lock] = {}
        self._work_dir = work_dir or (Path(tempfile.gettempdir()) / "excel_vba_mcp" / _WORK_SUBDIR)
        # Startup-reconcile store for the owned Excel pid (Block A): written after a
        # successful (re)launch and cleared on close, so the next server start can
        # detect a leftover live instance and refuse to start a second Excel. The
        # marker lives in a stable per-user dir (owned_pid_dir) so it survives a
        # restart; tests pass an isolated temp dir.
        self._owned_pid_store = OwnedPidStore(owned_pid_dir or self._work_dir)
        # Optional job registry (Block F): published by the server so a successful
        # recovery can finalize in-flight jobs to INTERRUPTED_UNKNOWN. ``None`` in
        # minimal/embedder use (recovery then leaves jobs untouched).
        self._jobs: JobRegistry | None = None
        # Optional in-memory snapshot gate (U2; plan_03 §6.2). Published by the
        # server as ``SnapshotService.has_snapshot`` so the ``excel_get_status``
        # policy block can report ``has_snapshot`` from the in-memory
        # ``_latest_snapshot`` (never the on-disk store). ``None`` in
        # minimal/embedder use (the policy then reports ``has_snapshot=False``).
        self._snapshot_gate: Callable[[str], bool] | None = None

    # ------------------------------------------------------------------
    # Properties (non-COM)
    # ------------------------------------------------------------------

    @property
    def session_id(self) -> str:
        """The owning session id."""

        return self._session_id

    @property
    def work_dir(self) -> Path:
        """The working-copy root directory."""

        return self._work_dir

    @property
    def accepting(self) -> bool:
        """Whether this session accepts new work (non-COM bookkeeping).

        ``False`` while recovering (task_034) or after :meth:`close_session`.
        """

        return self._accepting

    def set_job_registry(self, jobs: JobRegistry | None) -> None:
        """Publish the job registry used to finalize jobs on a successful recovery.

        The server calls this once at startup. When set, a successful
        :meth:`recover_owned_process` marks the session's in-flight jobs as
        ``INTERRUPTED_UNKNOWN`` (definitive; never completed, never replayed).
        """

        self._jobs = jobs

    def set_snapshot_gate(self, gate: Callable[[str], bool] | None) -> None:
        """Publish the in-memory snapshot gate used by the status policy block (U2).

        The server calls this once at startup with
        :meth:`SnapshotService.has_snapshot`. The gate is read-only and changes no
        behavior — it only lets :meth:`excel_get_status` report ``has_snapshot``
        truthfully from the in-memory ``_latest_snapshot`` (never the store).
        """

        self._snapshot_gate = gate

    def rebind_restored_identity(
        self,
        workbook_id: str,
        *,
        canonical: str,
        source_sha256: str,
    ) -> None:
        """Rebind a restored session's identity to the ORIGINAL source (plan_06 R3).

        After a snapshot restore the freshly opened working copy is initially
        registered under the snapshot **file** path (its open source). This method
        rebinds the live record so its ``canonical`` (and the ``source_path`` on
        the cached :class:`WorkbookInfo`) is the manifest's ``source_path`` — the
        ORIGINAL — and its ``source_sha256`` is the snapshot baseline. The working
        copy on disk (``work/<id>.ext``) is unchanged and the snapshot file is
        never the session identity, so ``list_workbooks`` / reopen reflect the
        original.

        The per-canonical-path active map and lock are migrated to the rebound
        canonical so a later ``excel_open`` of the original fast-rejects against
        the restored id and ``close_workbook`` releases the right lock. The caller
        must invoke this only AFTER the previous working copy has been closed (so
        the original canonical key is free).

        Args:
            workbook_id: The restored (open) workbook id.
            canonical: The original source path to rebind to.
            source_sha256: The baseline source hash to record.
        """

        record = self._require_record(workbook_id)
        old_canonical = record.canonical
        record.canonical = canonical
        record.source_sha256 = source_sha256
        record.info = record.info.model_copy(update={"source_path": canonical})
        if old_canonical != canonical:
            # Migrate the per-path active map + lock to the rebound canonical.
            self._active_by_canonical.pop(old_canonical, None)
            self._locks.pop(old_canonical, None)
            self._active_by_canonical[canonical] = workbook_id
            self._locks[canonical] = record.lock

    # ------------------------------------------------------------------
    # Open
    # ------------------------------------------------------------------

    async def open_workbook(
        self,
        path: str,
        *,
        mode: SessionMode = SessionMode.INSPECT,
        read_only: bool = False,
        timeout_s: float | None = None,
    ) -> WorkbookInfo:
        """Open an allowlisted workbook as a hidden working copy.

        The original file is copied into ``work/`` and only the copy is opened;
        the original is never modified. ``inspect`` mode disables macros
        (``ForceDisable``); ``execute`` mode applies ``ByUI`` immediately before
        ``Open`` and requires an approved context (deny-by-default).

        Args:
            path: Caller-supplied path; validated against the allowlist.
            mode: ``inspect`` (macros disabled) or ``execute`` (macros by UI).
            read_only: Open the working copy read-only.
            timeout_s: Soft per-call ceiling for the STA open. Defaults to the
                server-configured ``open_timeout_s`` (Block E authority).

        Returns:
            A :class:`WorkbookInfo` describing the opened working copy.

        Raises:
            PathNotAllowedError: If ``path`` is outside the allowlisted roots.
            ApprovalDeniedError: If the approval gate denies the open.
            ExcelVbaMcpError: If the workbook is already open on this path.
            ExcelHungError: If the STA open times out / needs recovery.
        """

        mode = SessionMode(mode)
        # Resolve the per-call open deadline to the server authority (Block E).
        effective_timeout_s = self._open_timeout_s if timeout_s is None else timeout_s
        # Lazy auto-recovery fallback (F2b / strategy C): if the owned session was
        # left in NEEDS_RECOVERY with a dead owned pid (e.g. the process was
        # terminated after a hung macro), self-heal with ONE in-process recovery so
        # the next open does not require a server restart.
        await self._maybe_auto_recover()
        self._require_accepting_new_work("excel_open")
        canonical = self._path_policy.validate_path(path, OperationClass.MUTATING)
        canonical_str = str(canonical)

        # Deny-by-default approval BEFORE any mutation (copy / open). Inspect
        # read-only opens are auto-allowed by the policy; execute opens always
        # require an approved context; a writable inspect open is a mutation.
        if mode == SessionMode.INSPECT and read_only:
            approval_op = OperationClass.READ_ONLY
        elif mode == SessionMode.EXECUTE:
            approval_op = OperationClass.EXECUTING
        else:
            approval_op = OperationClass.MUTATING
        self._require_approval(
            self._approval_key(canonical_str, "excel_open"),
            approval_op,
        )

        # Fast reject: if this canonical path is already open, error clearly
        # instead of blocking forever on the held per-path lock.
        existing_id = self._active_by_canonical.get(canonical_str)
        if existing_id is not None:
            raise ExcelVbaMcpError(
                f"Workbook is already open for {canonical.name!r}.",
                context={"canonical_path": canonical_str, "workbook_id": existing_id},
            )

        # Acquire the per-canonical-path lock and HOLD it for the workbook
        # lifetime (released by close_workbook). The lock is shared per path, so
        # a concurrent open of the same file serializes behind the first,
        # preventing duplicate concurrent opens.
        lock = self._locks.setdefault(canonical_str, asyncio.Lock())
        await lock.acquire()
        try:
            # Re-check under the lock (another open may have won the race).
            existing_id = self._active_by_canonical.get(canonical_str)
            if existing_id is not None:
                raise ExcelVbaMcpError(
                    f"Workbook is already open for {canonical.name!r}.",
                    context={"canonical_path": canonical_str, "workbook_id": existing_id},
                )

            workbook_id = str(uuid4())
            # Capture source identity AFTER path validation and BEFORE the working
            # copy is materialized (U7; plan_05 §9): a streaming SHA-256 + byte
            # size of the canonical original, re-verified at Save As so a source
            # that drifted or was replaced since open is fail-closed.
            source_sha256, source_size = _sha256_file(canonical)
            working_path = self._materialize_working_copy(canonical, workbook_id)

            try:
                info = await self._open_on_worker(
                    working_path=working_path,
                    canonical_str=canonical_str,
                    workbook_id=workbook_id,
                    mode=mode,
                    read_only=read_only,
                    timeout_s=effective_timeout_s,
                )
            except Exception:
                # Never leak the per-path lock or working copy on a failed open.
                self._cleanup_working_copy(working_path)
                raise
        except Exception:
            lock.release()
            raise

        record = _SessionRecord(
            workbook_id=workbook_id,
            canonical=canonical_str,
            working_path=working_path,
            mode=mode,
            read_only=read_only,
            # A freshly opened working copy has no unsaved changes.
            dirty=False,
            info=info,
            lock=lock,
            source_sha256=source_sha256,
            source_size=source_size,
        )
        # The lock is intentionally still held for the workbook lifetime; it is
        # released by close_workbook. Register bookkeeping now.
        self._records[workbook_id] = record
        self._active_by_canonical[canonical_str] = workbook_id
        self._session_mode = mode

        self._record_audit(
            tool="excel_open",
            operation_class="mutating",
            params={"path": canonical_str, "mode": mode.value, "read_only": read_only},
            workbook={"canonical_path": canonical_str},
            result={"workbook_id": workbook_id, "path": working_path},
        )
        return info

    async def _open_on_worker(
        self,
        *,
        working_path: str,
        canonical_str: str,
        workbook_id: str,
        mode: SessionMode,
        read_only: bool,
        timeout_s: float,
    ) -> WorkbookInfo:
        """Build and submit the STA-local open callable."""

        app_adapter = self._app_adapter
        workbook_adapter = self._workbook_adapter
        session_id = self._session_id
        security = int(automation_security_for(mode))

        def open_fn() -> WorkbookInfo:
            # Apply the mode's automation security immediately before Open and
            # restore it in finally (adapter context managers, STA thread).
            with app_adapter.automation_security(security), app_adapter.alerts_suppressed():
                wb = app_adapter.open_workbook(
                    working_path,
                    read_only=read_only,
                    update_links=0,
                    notify=False,
                    add_to_mru=False,
                )
                return workbook_adapter.describe(
                    wb,
                    workbook_id=workbook_id,
                    session_id=session_id,
                    mode=mode,
                    source_path=canonical_str,
                )

        open_fn.__name__ = "open_workbook"
        return await self._worker.submit(
            open_fn,
            timeout=timeout_s,
            section="open",
            op_class=OperationClass.MUTATING,
        )

    # ------------------------------------------------------------------
    # Close
    # ------------------------------------------------------------------

    async def close_workbook(
        self,
        workbook_id: str,
        *,
        save_changes: bool = False,
        discard_changes: bool = False,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SessionStatus:
        """Close an open workbook and release its lock (refuse-by-default).

        Unsaved-change disposition (D1 / R-Q1 / R-Q8 / R-Q9), decided from the
        **real** dirty signal (:meth:`WorkbookAdapter.is_dirty` -> ``not
        wb.Saved``), never ``record.dirty`` (which is never set ``True``):

        * **clean** workbook -> closes normally;
        * dirty + ``save_changes`` -> save then close;
        * dirty + ``discard_changes`` -> discard (close without saving);
        * dirty + neither -> **refuse** with a structured
          :class:`ExcelVbaMcpError` (refuse-by-default);
        * ``save_changes`` and ``discard_changes`` together -> rejected as
          contradictory (:class:`ExcelVbaMcpError`).

        Args:
            workbook_id: The id returned by :meth:`open_workbook`.
            save_changes: Persist changes on close. Default ``False`` so an open
                is never silently saved (spec §7.1).
            discard_changes: Discard unsaved changes on close. Default
                ``False``; mutually exclusive with ``save_changes``.
            timeout_s: Soft per-call ceiling for the STA close.

        Returns:
            A :class:`SessionStatus` describing the session after close.

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown, the flags are
                contradictory, or the workbook has unsaved changes and neither
                ``save_changes`` nor ``discard_changes`` was given.
            ApprovalDeniedError: If a save-on-close is not approved.
        """

        record = self._require_record(workbook_id)

        # Contradictory flags are a caller error — reject before any COM work.
        if save_changes and discard_changes:
            raise ExcelVbaMcpError(
                "save_changes and discard_changes are contradictory; pass at most one.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "reason": "contradictory_flags",
                    "workbook_id": workbook_id,
                },
            )

        # Refuse-by-default: a dirty workbook needs an explicit disposition. The
        # dirty check is a STA-local read of the live ``wb.Saved`` (R-Q8).
        dirty = await self._is_workbook_dirty(record, timeout_s)
        if dirty and not save_changes and not discard_changes:
            raise ExcelVbaMcpError(
                "The workbook has unsaved changes; close refuses by default. "
                "Pass save_changes=true to save then close, or discard_changes=true "
                "to discard them.",
                code=ErrorCode.CONFIRMATION_REQUIRED,
                context={
                    "reason": "unsaved_changes_refused",
                    "workbook_id": workbook_id,
                },
            )

        # A save-on-close is a mutation; a plain close / discard is read-only
        # (auto-allowed). Discarding only drops the disposable working copy, so
        # the original is never touched.
        close_op = OperationClass.MUTATING if save_changes else OperationClass.READ_ONLY
        self._require_approval(
            self._approval_key(record.canonical, "excel_close"),
            close_op,
        )

        app_adapter = self._app_adapter

        def close_fn() -> bool:
            wb = _find_open_workbook(app_adapter, record.working_path)
            if wb is not None:
                app_adapter.close_workbook(wb, save_changes=save_changes)
            return True

        close_fn.__name__ = "close_workbook"
        await self._worker.submit(
            close_fn,
            timeout=timeout_s,
            section="close",
            op_class=OperationClass.MUTATING,
        )

        # Release bookkeeping + the per-path lock (held since open). Drop the
        # shared lock entry once released so the registry does not grow unbounded.
        self._records.pop(workbook_id, None)
        self._active_by_canonical.pop(record.canonical, None)
        self._cleanup_working_copy(record.working_path)
        record.lock.release()
        self._locks.pop(record.canonical, None)

        self._record_audit(
            tool="excel_close",
            operation_class="mutating" if save_changes else "read_only",
            params={
                "workbook_id": workbook_id,
                "save_changes": save_changes,
                "discard_changes": discard_changes,
            },
            workbook={"canonical_path": record.canonical},
            result={
                "workbook_id": workbook_id,
                "closed": True,
                "saved": save_changes,
                "discarded": discard_changes and not save_changes,
            },
        )
        return self._build_status(
            workbook_id=None,
            mode=record.mode,
            read_only=record.read_only,
            dirty=False,
            notes=["workbook closed"],
        )

    async def _is_workbook_dirty(self, record: _SessionRecord, timeout_s: float) -> bool:
        """Return whether the live workbook has unsaved changes (STA-local read).

        Uses the real :meth:`WorkbookAdapter.is_dirty` (``not wb.Saved``); a
        workbook that is no longer open in the owned instance is treated as
        clean (nothing to lose). Runs on the single-STA worker; never touches the
        event loop with a COM proxy.
        """

        app_adapter = self._app_adapter
        workbook_adapter = self._workbook_adapter

        def dirty_fn() -> bool:
            wb = _find_open_workbook(app_adapter, record.working_path)
            if wb is None:
                return False
            return workbook_adapter.is_dirty(wb)

        dirty_fn.__name__ = "is_workbook_dirty"
        return await self._worker.submit(
            dirty_fn,
            timeout=timeout_s,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    async def save_workbook(
        self,
        workbook_id: str,
        *,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> WorkbookInfo:
        """Explicitly save an open workbook's working copy.

        Args:
            workbook_id: The id returned by :meth:`open_workbook`.
            timeout_s: Soft per-call ceiling for the STA save.

        Returns:
            A refreshed :class:`WorkbookInfo` (``dirty`` cleared).

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown.
            ApprovalDeniedError: If the save is not approved.
        """

        record = self._require_record(workbook_id)
        self._require_approval(
            self._approval_key(record.canonical, "excel_save"),
            OperationClass.MUTATING,
        )

        app_adapter = self._app_adapter
        workbook_adapter = self._workbook_adapter
        session_id = self._session_id

        def save_fn() -> WorkbookInfo:
            wb = _find_open_workbook(app_adapter, record.working_path)
            if wb is None:
                raise WorkbookNotFoundError(
                    "The workbook is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id},
                )
            with app_adapter.alerts_suppressed():
                wb.Save()
            return workbook_adapter.describe(
                wb,
                workbook_id=workbook_id,
                session_id=session_id,
                mode=record.mode,
                source_path=record.canonical,
            )

        save_fn.__name__ = "save_workbook"
        info = await self._worker.submit(
            save_fn,
            timeout=timeout_s,
            section="save",
            op_class=OperationClass.MUTATING,
        )
        record.dirty = False
        record.info = info
        # A successful save is a content mutation: bump the record's own version
        # so a later paged read detects the change (WP-B large-data cursor).
        record.version += 1

        self._record_audit(
            tool="excel_save",
            operation_class="mutating",
            params={"workbook_id": workbook_id},
            workbook={"canonical_path": record.canonical},
            result={"workbook_id": workbook_id, "saved": True},
        )
        return info

    # ------------------------------------------------------------------
    # Save As (U7; plan_05 §5.5, §9) — new target file, session stays bound
    # ------------------------------------------------------------------

    async def save_as_workbook(
        self,
        workbook_id: str,
        path: str,
        *,
        overwrite_existing: bool = False,
        overwrite_source: bool = False,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SaveAsResult:
        """Save an open workbook's working copy to a (possibly new) target path.

        Fail-closed Save As (plan_05 §5.5, §9): the working copy is flushed in
        place with ``wb.Save()`` on the STA worker, then copied to the target by
        a pure-filesystem atomic temp→``os.replace`` on the event loop. The
        session is **never rebound** — ``wb.SaveAs`` is never called (it would
        repoint ``FullName`` and break ``_find_open_workbook`` / original
        protection); ``workbook_id`` / ``working_path`` / ``canonical`` are
        unchanged and a later ``excel_save`` still writes the working copy.

        The source identity (streaming SHA-256 + size) captured at
        :meth:`open_workbook` is re-verified before any source-overwrite, so a
        source that was replaced or modified since open is rejected. A single
        ``<target>.bak`` preserves the previous target across an overwrite.

        Args:
            workbook_id: The id returned by :meth:`open_workbook`.
            path: Destination path; validated against the allowlist (MUTATING).
            overwrite_existing: Replace an existing target (writes a ``.bak``).
            overwrite_source: Allow the target to be the original source itself
                (requires the source hash to still match; writes a ``.bak``).
            timeout_s: Soft per-call ceiling for the STA ``wb.Save()`` flush.

        Returns:
            A :class:`SaveAsResult` whose ``outcome`` is ``"ok"`` only after the
            final size + SHA-256 of the target were verified; ``"unknown"`` /
            ``"error"`` fail-closed (never success), with ``backup_path`` when a
            previous target was preserved.

        Raises:
            PathNotAllowedError: If ``path`` is outside the allowlisted roots.
            ProtectedTargetError: If the target is the working copy, the source
                (without ``overwrite_source``), or an existing target (without
                ``overwrite_existing``).
            StaleTargetError: If the source was replaced or modified since open.
            ApprovalDeniedError: If the MUTATING approval is denied.
            WorkbookNotFoundError: If the workbook is no longer open.
        """

        record = self._require_record(workbook_id)

        # Path gate: allowlist + extension + symlink/UNC/device (no writes).
        canonical_t = self._path_policy.validate_path(path, OperationClass.MUTATING)
        # The working copy is the live session file; Save As must never target it.
        if os.path.normcase(str(canonical_t)) == os.path.normcase(record.working_path):
            raise ProtectedTargetError(
                "The Save As target is the open working copy; choose a different path.",
                context={
                    "reason": "target_is_working_copy",
                    "target_path": str(canonical_t),
                },
            )

        # Approval MUTATING on the target canonical (== source canonical when
        # T==S, so a single check covers both).
        self._require_approval(
            self._approval_key(str(canonical_t), "excel_save_as"),
            OperationClass.MUTATING,
        )

        # Re-validate the stored source canonical and re-hash it (fail-closed).
        re_canonical = self._path_policy.validate_path(record.canonical, OperationClass.READ_ONLY)
        if os.path.normcase(str(re_canonical)) != os.path.normcase(record.canonical):
            raise StaleTargetError(
                "The original source path was replaced since the workbook was opened.",
                context={
                    "reason": "source_replaced",
                    "stored_source": record.canonical,
                    "current_source": str(re_canonical),
                },
            )
        source_sha256, source_size = _sha256_file(re_canonical)
        if source_sha256 != record.source_sha256 or source_size != record.source_size:
            raise StaleTargetError(
                "The original source file was modified since the workbook was opened.",
                context={
                    "reason": "source_modified_since_open",
                    "expected_sha256": record.source_sha256,
                    "actual_sha256": source_sha256,
                },
            )

        same = os.path.normcase(str(canonical_t)) == os.path.normcase(record.canonical)
        exists = canonical_t.exists()

        # Decision matrix (plan_05 §5.5).
        if same and not overwrite_source:
            raise ProtectedTargetError(
                "The target is the original source; pass overwrite_source=true to replace it.",
                context={"reason": "source_protected", "target_path": str(canonical_t)},
            )
        if not same and exists and not overwrite_existing:
            raise ProtectedTargetError(
                "The target already exists; pass overwrite_existing=true to replace it.",
                context={"reason": "target_exists", "target_path": str(canonical_t)},
            )

        # A pre-existing ``<target>.bak`` (from a prior interrupted save) is a
        # DETERMINISTIC conflict detected before the destination is replaced, not
        # an indeterminate interruption. Reject it pre-mutation (no temp, no
        # replace; the existing ``.bak`` and the destination are left untouched)
        # so the caller gets an actionable isError instead of ``outcome="unknown"``
        # (no auto-resume).
        if exists and Path(f"{canonical_t}.bak").exists():
            raise ExcelVbaMcpError(
                "A backup already exists for the target; refusing to overwrite it "
                "(no auto-resume).",
                code=ErrorCode.PROTECTED_TARGET,
                context={
                    "reason": "backup_exists",
                    "target_path": str(canonical_t),
                    "backup_path": f"{canonical_t}.bak",
                },
            )

        # 1. Flush the working copy in place (never wb.SaveAs).
        await self._flush_working_copy(record, timeout_s)

        # 2-6. Pure-filesystem atomic copy + single .bak + final verify.
        write = self._atomic_save_as_copy(
            working_path=record.working_path,
            canonical_t=canonical_t,
            target_exists=exists,
        )

        result = SaveAsResult(
            workbook_id=workbook_id,
            target_path=str(canonical_t),
            source_path=record.canonical,
            source_overwritten=same,
            overwrote_existing=exists,
            backup_path=write["backup_path"],
            sha256=write["sha256"],
            size=write["size"],
            outcome=write["outcome"],
            mode=record.mode,
            notes=write["notes"],
            created_at=_utc_now(),
            updated_at=_utc_now(),
        )
        self._record_audit(
            tool="excel_save_as",
            operation_class="mutating",
            params={
                "workbook_id": workbook_id,
                "target_path": str(canonical_t),
                "overwrite_existing": overwrite_existing,
                "overwrite_source": overwrite_source,
            },
            workbook={"canonical_path": record.canonical},
            result={
                "target_path": str(canonical_t),
                "sha256": write["sha256"],
                "size": write["size"],
                "backup_path": write["backup_path"],
                "source_overwritten": same,
                "outcome": write["outcome"],
            },
        )
        return result

    async def _flush_working_copy(self, record: _SessionRecord, timeout_s: float) -> None:
        """Flush the working copy in place via ``wb.Save()`` on the STA worker."""

        app_adapter = self._app_adapter

        def flush_fn() -> None:
            wb = _find_open_workbook(app_adapter, record.working_path)
            if wb is None:
                raise WorkbookNotFoundError(
                    "The workbook is no longer open in the owned instance.",
                    context={"workbook_id": record.workbook_id},
                )
            with app_adapter.alerts_suppressed():
                wb.Save()

        flush_fn.__name__ = "save_as_flush"
        # Await the single-STA FIFO submit so the flush completes before the
        # pure-filesystem copy runs on the event loop.
        await self._worker.submit(
            flush_fn,
            timeout=timeout_s,
            section="save",
            op_class=OperationClass.MUTATING,
        )

    def _atomic_save_as_copy(
        self,
        *,
        working_path: str,
        canonical_t: Path,
        target_exists: bool,
    ) -> dict[str, Any]:
        """Atomic temp→replace copy with a single ``.bak`` and final verify.

        Pure filesystem (no COM); runs on the event loop. Never reports success
        from "no exception": ``outcome="ok"`` only after the staged copy and the
        final target both match the working copy by size + SHA-256. Any
        interruption after the backup boundary yields ``outcome="unknown"`` with
        ``backup_path`` (no auto-resume); a clean pre-backup failure yields
        ``outcome="error"`` with the destination untouched. A pre-existing ``.bak``
        from a prior interrupted save is rejected earlier in
        :meth:`save_as_workbook` (deterministic ``PROTECTED_TARGET`` /
        ``backup_exists``) before this method runs, so it is never overwritten
        here and never reaches the destination.
        """

        working = Path(working_path)
        working_sha, working_size = _sha256_file(working)
        notes: list[str] = []
        backup_path: str | None = None
        descriptor = -1
        temp_path: Path | None = None
        backup_created = False
        try:
            canonical_t.parent.mkdir(parents=True, exist_ok=True)
            descriptor, temp_name = tempfile.mkstemp(
                prefix=f".{canonical_t.name}.", suffix=".tmp", dir=str(canonical_t.parent)
            )
            temp_path = Path(temp_name)
            # Copy working -> temp, flush + fsync.
            with os.fdopen(descriptor, "wb") as out_handle:
                descriptor = -1
                with working.open("rb") as in_handle:
                    shutil.copyfileobj(in_handle, out_handle, length=1 << 20)
                out_handle.flush()
                os.fsync(out_handle.fileno())
            # Verify the staged copy matches the working copy (size + sha).
            staged_sha, staged_size = _sha256_file(temp_path)
            if staged_size != working_size or staged_sha != working_sha:
                notes.append("staged copy verification failed; destination untouched")
                return {
                    "outcome": "error",
                    "sha256": working_sha,
                    "size": working_size,
                    "backup_path": None,
                    "notes": notes,
                }
            # Preserve the current target (single .bak, overwritten each save).
            if target_exists:
                backup_path = f"{canonical_t}.bak"
                self._copy_and_fsync(canonical_t, Path(backup_path))
                backup_created = True
                notes.append("previous target preserved at .bak")
            # Atomic publish (same volume).
            os.replace(temp_path, canonical_t)
            temp_path = None
            # Final verify of the published target.
            final_sha, final_size = _sha256_file(canonical_t)
            if final_size != working_size or final_sha != working_sha:
                notes.append("final target verification failed")
                rolled_back = self._rollback_from_backup(canonical_t, backup_path)
                notes.append("rolled back from .bak" if rolled_back else "rollback unavailable")
                return {
                    "outcome": "unknown",
                    "sha256": final_sha,
                    "size": final_size,
                    "backup_path": backup_path,
                    "notes": notes,
                }
            notes.append("target verified (size + sha256)")
            return {
                "outcome": "ok",
                "sha256": final_sha,
                "size": final_size,
                "backup_path": backup_path,
                "notes": notes,
            }
        except OSError as exc:
            # Interruption / unprovable write: never report success. If the backup
            # boundary was crossed the destination state is ambiguous -> unknown.
            notes.append(f"save-as interrupted: {exc.__class__.__name__}")
            outcome = "unknown" if backup_created else "error"
            return {
                "outcome": outcome,
                "sha256": working_sha,
                "size": working_size,
                "backup_path": backup_path,
                "notes": notes,
            }
        finally:
            if descriptor >= 0:
                with contextlib.suppress(OSError):
                    os.close(descriptor)
            if temp_path is not None:
                with contextlib.suppress(OSError):
                    temp_path.unlink(missing_ok=True)

    @staticmethod
    def _copy_and_fsync(source: Path, destination: Path) -> None:
        """Copy ``source``→``destination`` (overwriting) with flush + fsync."""

        with source.open("rb") as in_handle, destination.open("wb") as out_handle:
            shutil.copyfileobj(in_handle, out_handle, length=1 << 20)
            out_handle.flush()
            os.fsync(out_handle.fileno())

    @staticmethod
    def _rollback_from_backup(canonical_t: Path, backup_path: str | None) -> bool:
        """Best-effort restore of ``canonical_t`` from ``backup_path`` (never raises)."""

        if not backup_path:
            return False
        backup = Path(backup_path)
        if not backup.is_file():
            return False
        try:
            descriptor, temp_name = tempfile.mkstemp(
                prefix=f".{canonical_t.name}.", suffix=".tmp", dir=str(canonical_t.parent)
            )
            temp_path = Path(temp_name)
            with os.fdopen(descriptor, "wb") as out_handle:
                descriptor = -1
                with backup.open("rb") as in_handle:
                    shutil.copyfileobj(in_handle, out_handle, length=1 << 20)
                out_handle.flush()
                os.fsync(out_handle.fileno())
            os.replace(temp_path, canonical_t)
            return True
        except OSError:
            return False

    # ------------------------------------------------------------------
    # Status / listing (no COM)
    # ------------------------------------------------------------------

    async def get_status(self, workbook_id: str | None = None) -> SessionStatus:
        """Return session / workbook status from non-COM bookkeeping.

        Merges the worker's non-COM :class:`WorkerStatus` snapshot with session
        bookkeeping; never submits to the worker.

        Args:
            workbook_id: When given, report that workbook's mode/read-only/dirty
                state; otherwise report session-level status.
        """

        worker_status = self._worker.get_status()
        now = _utc_now()

        if workbook_id is not None:
            record = self._records.get(workbook_id)
            if record is None:
                return self._build_status(
                    workbook_id=workbook_id,
                    mode=self._session_mode,
                    read_only=False,
                    dirty=False,
                    notes=[f"workbook_id {workbook_id!r} is not currently open"],
                    worker_status=worker_status,
                )
            return self._build_status(
                workbook_id=workbook_id,
                mode=record.mode,
                read_only=record.read_only,
                dirty=record.dirty,
                worker_status=worker_status,
            )

        return self._build_status(
            workbook_id=None,
            mode=self._session_mode,
            read_only=False,
            dirty=False,
            worker_status=worker_status,
            created_at=now,
        )

    async def list_workbooks(self) -> list[WorkbookInfo]:
        """Return the currently open workbooks from bookkeeping (no COM)."""

        return [record.info for record in self._records.values()]

    def list_workbooks_sync(self) -> list[WorkbookInfo]:
        """Return the currently open workbooks without awaiting (no COM).

        Synchronous, non-COM convenience for application services that already
        run on the event loop and only need the session bookkeeping (e.g. the
        read-only VBA project service resolving a ``workbook_id`` before it
        submits work to the STA worker).
        """

        return [record.info for record in self._records.values()]

    def get_version(self, workbook_id: str) -> int:
        """Return the workbook's monotonic content version (no COM).

        The version starts at ``0`` on open and is bumped by :meth:`bump_version`
        on each successful mutation; reads never bump it (WP-A large-data version
        primitive). Raises :class:`WorkbookNotFoundError` for an unknown id.
        """

        return self._require_record(workbook_id).version

    def bump_version(self, workbook_id: str) -> None:
        """Increment the workbook's content version by one (no COM).

        Called after a successful mutation so a later paged read can detect that
        the workbook changed. Raises :class:`WorkbookNotFoundError` for an unknown
        id.
        """

        self._require_record(workbook_id).version += 1

    # ------------------------------------------------------------------
    # Recovery / session lifecycle (task_034; spec §5.6, §9.4, §11)
    # ------------------------------------------------------------------

    def worker_status(self) -> WorkerStatus:
        """Return the worker's non-COM status snapshot (never submits work)."""

        return self._worker.get_status()

    async def recover_session(
        self,
        *,
        reason: str = "manual",
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SessionStatus:
        """Rebuild the owned hidden Excel instance after a failure (spec §9.4).

        Recovery is the safe exit from a timed-out / lost-Excel state. The
        sequence is:

        1. **Stop accepting new work** (session-level ``accepting`` guard);
        2. **close every owned workbook** with ``Close(SaveChanges=False)`` and
           **Quit** / release the owned proxy — the worker's reset path drains
           the owned instance (``ExcelComWorker.reset`` -> close-section drain);
        3. **release the STA apartment and rebuild** the hidden instance via
           ``worker.reset()`` on the same single-STA thread.

        Only the automation-owned instance is touched: a user-owned Excel process
        is never enumerated, activated or terminated (BG05).

        The session bookkeeping (open workbooks, per-path locks, working copies)
        is dropped because the rebuilt instance has no open workbooks. The method
        **never reports a false success**: if the reset fails it returns
        ``state=NEEDS_RECOVERY`` (or the worker's truthful ``closed``/``error``
        state) with diagnostics in ``notes`` and keeps refusing new work.

        Args:
            reason: Short label recorded in the audit entry / notes.
            timeout_s: Accepted for signature symmetry with the other session
                operations; the worker's own bounded reset join is authoritative.

        Returns:
            A :class:`SessionStatus` whose state is ``running`` after a
            successful rebuild, or ``needs_recovery`` / ``error`` with
            diagnostics when the rebuild failed.
        """

        worker = self._worker
        self._accepting = False
        self._state = WorkbookSessionState.NEEDS_RECOVERY

        before = worker.get_status()
        notes: list[str] = [f"recovery requested (reason={reason})"]
        notes.append(f"worker_state_before={before.state.value}")
        if before.owned_pid is not None:
            notes.append(f"owned_pid_before={before.owned_pid}")
        if before.last_error:
            notes.append(f"last_error_before={before.last_error}")

        discarded = self._discard_open_workbooks()
        notes.append(f"discarded_open_workbooks={discarded}")

        reset_error: str | None = None
        if before.state == WorkbookSessionState.CLOSED:
            # A stopped worker cannot be reset on its thread; report truthfully
            # instead of waiting for a bounded join that can never complete.
            reset_error = "worker is closed; a new worker must be started"
        else:
            try:
                await worker.reset()
            except ExcelHungError as exc:
                reset_error = f"worker reset timed out: {exc}"
            except Exception as exc:  # noqa: BLE001 - recovery reports, never raises
                reset_error = f"worker reset failed: {exc.__class__.__name__}: {exc}"
        notes.append(reset_error or "worker reset completed")

        after = worker.get_status()
        if reset_error is None and after.state == WorkbookSessionState.RUNNING:
            self._state = WorkbookSessionState.RUNNING
            self._accepting = True
            self._closed = False
            # Persist the rebuilt owned pid so the next server start's reconcile
            # ("no second Excel") stays airtight after a lazy rebuild (Block A).
            self._persist_owned_pid(after.owned_pid)
            notes.append("owned Excel instance rebuilt (hidden)")
        else:
            failed_state = (
                after.state
                if after.state
                in (
                    WorkbookSessionState.NEEDS_RECOVERY,
                    WorkbookSessionState.ERROR,
                    WorkbookSessionState.CLOSED,
                )
                else WorkbookSessionState.NEEDS_RECOVERY
            )
            self._state = failed_state
            self._accepting = False
            notes.append(reset_error or f"reset did not reach running (state={after.state.value})")
            notes.append("new work is refused until recovery succeeds")
        if after.owned_pid is not None:
            notes.append(f"owned_pid_after={after.owned_pid}")

        status = self._build_status(
            workbook_id=None,
            mode=self._session_mode,
            read_only=False,
            dirty=False,
            notes=notes,
            worker_status=after,
        )
        self._record_audit(
            tool="recover_session",
            operation_class="executing",
            params={"reason": reason, "discarded_workbooks": discarded},
            result={
                "needs_recovery": self._state != WorkbookSessionState.RUNNING,
                "state": self._state.value,
                "owned_pid_after": after.owned_pid,
                "reset_error": reset_error,
            },
        )
        return status

    async def recover_owned_process(
        self,
        *,
        reason: str = "agent_directed",
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> dict[str, Any]:
        """Agent-directed recovery of the owned Excel process (``excel_recover``).

        This is the Block A recovery path (report_07 §5, plan_02 Work item A). It is
        **agent-directed** (the agent calls ``excel_recover``; the server never
        auto-terminates a live pid on a deadline) and **fail-closed**:

        1. **Approve (RECOVERY class).** ``ApprovalPolicy.check_recovery`` is
           auto-allowed in BOTH full and restricted mode (audited); the key is
           pid-scoped (``"pid:<owned_pid>"``), never a workbook wildcard.
        2. **Fail closed** when the worker is closed or no live, identity-verified
           handle is held (``owned_handle`` is ``None`` / reduced to ``int``): the
           outcome reports ``recovery_possible=False`` / ``ok=False`` and the
           documented fallback is a **server restart** (U8). **No pid-only kill.**
        3. **Terminate + confirm via the SAME held handle**
           (``ExcelAppAdapter.terminate_owned_handle`` ->
           ``TerminateProcess`` -> ``WaitForSingleObject`` == ``WAIT_OBJECT_0``).
        4. **Discard working copies only** (originals are never modified) and
           **rebuild** the hidden instance via ``worker.reset()`` on the single STA.
        5. **Interrupted mutation = unknown, never auto-retried:** ``prior_job`` is
           ``"unknown"`` when an op was in flight / deadline-exceeded /
           needs-recovery; the caller must not replay it.

        Args:
            reason: Short label recorded in the audit entry / outcome.
            timeout_s: Accepted for signature symmetry; the worker's bounded reset
                join is authoritative.

        Returns:
            A JSON-serializable dict with the recovery contract fields: ``ok``,
            ``terminated``, ``confirmed``, ``recovery_possible``, ``process_restarted``,
            ``prior_job`` (``completed`` / ``failed`` / ``unknown``), ``what_lost``,
            ``reason``, ``owned_pid`` and ``approval_source``. The payload states it
            is destructive to **working copies only** and originals are safe.
        """

        del timeout_s  # the worker's bounded reset join is authoritative
        worker = self._worker
        adapter = self._app_adapter
        before = worker.get_status()
        owned_pid = getattr(before, "owned_pid", None)
        worker_state = getattr(before, "state", None)
        active_operation = getattr(before, "active_operation", None)

        # RECOVERY approval (auto-allowed FULL + RESTRICTED, pid-scoped, audited).
        recovery_key = ApprovalKey(
            canonical_path=f"pid:{owned_pid if owned_pid is not None else 'none'}",
            action="excel_recover",
        )
        decision = self._approval.check_recovery(recovery_key)
        approval_source = decision.source.value if decision.source is not None else None

        # An in-flight / deadline-exceeded / needs-recovery op is UNKNOWN (never
        # auto-retried); otherwise the prior op completed.
        prior_job = (
            "unknown"
            if active_operation
            or worker_state
            in (WorkbookSessionState.DEADLINE_EXCEEDED, WorkbookSessionState.NEEDS_RECOVERY)
            else "completed"
        )

        # Fail closed: a closed worker cannot be recovered in-process.
        if worker_state == WorkbookSessionState.CLOSED:
            return self._finish_recovery(
                reason="worker is closed; a server restart is required (U8)",
                owned_pid=owned_pid,
                prior_job=prior_job,
                approval_source=approval_source,
                discarded=0,
            )

        # Fail closed: recovery requires a live, identity-verified held handle.
        if (
            owned_pid is None
            or not adapter.recovery_available
            or adapter.owned_handle is None
            or isinstance(adapter.owned_handle, int)
        ):
            return self._finish_recovery(
                reason=(
                    "owned process handle unavailable or identity unverified; "
                    "recovery unavailable — server restart is the fallback (U8)"
                ),
                owned_pid=owned_pid,
                prior_job=prior_job,
                approval_source=approval_source,
                discarded=0,
            )

        # Recovery is possible: stop accepting, discard working copies (originals
        # untouched), terminate+confirm via the SAME handle, then rebuild.
        self._accepting = False
        self._state = WorkbookSessionState.NEEDS_RECOVERY
        discarded = self._discard_open_workbooks()

        # Terminate+confirm via the SAME held handle, off the event loop (the handle
        # is a free-thread kernel object; works while the STA is blocked in COM).
        terminated = await asyncio.to_thread(adapter.terminate_owned_handle)
        confirmed = terminated  # success only on WAIT_OBJECT_0

        reset_error: str | None = None
        try:
            await worker.reset()
        except ExcelHungError as exc:
            reset_error = f"worker reset timed out: {exc}"
        except Exception as exc:  # noqa: BLE001 - recovery reports, never raises
            reset_error = f"worker reset failed: {exc.__class__.__name__}: {exc}"

        after = worker.get_status()
        restarted = (
            reset_error is None and getattr(after, "state", None) == WorkbookSessionState.RUNNING
        )
        new_pid = getattr(after, "owned_pid", None)
        if restarted:
            self._state = WorkbookSessionState.RUNNING
            self._accepting = True
            self._closed = False
            self._persist_owned_pid(new_pid)
            # The owned process is confirmed restarted: any in-flight /
            # deadline-exceeded job is now definitively INTERRUPTED_UNKNOWN (never
            # completed, never replayed). Best-effort — must never fail a recovery.
            self._mark_jobs_interrupted()
        else:
            self._state = WorkbookSessionState.NEEDS_RECOVERY
            self._accepting = False

        outcome: dict[str, Any] = {
            "ok": restarted,
            "terminated": terminated,
            "confirmed": confirmed,
            "recovery_possible": True,
            "process_restarted": restarted,
            "prior_job": prior_job,
            "what_lost": (f"working copies discarded ({discarded}); originals are never modified"),
            "reason": reset_error,
            "owned_pid": new_pid if restarted else owned_pid,
            "approval_source": approval_source,
        }
        self._record_audit(
            tool="excel_recover",
            operation_class="recovery",
            params={"reason": reason, "owned_pid": owned_pid, "discarded_workbooks": discarded},
            result={
                "ok": restarted,
                "terminated": terminated,
                "confirmed": confirmed,
                "process_restarted": restarted,
                "prior_job": prior_job,
                "owned_pid_after": new_pid if restarted else owned_pid,
                "reset_error": reset_error,
            },
        )
        return outcome

    def _mark_jobs_interrupted(self) -> None:
        """Finalize the session's interruptible jobs to INTERRUPTED_UNKNOWN.

        Called only after a **successful** recovery (owned process confirmed
        restarted). Best-effort: a registry/persistence error must never fail the
        recovery outcome.
        """

        jobs = self._jobs
        if jobs is None:
            return
        with contextlib.suppress(Exception):
            jobs.mark_interrupted(reason="recovery")

    def _persist_owned_pid(self, owned_pid: int | None) -> None:
        """Persist the current owned pid + owner identity for startup reconcile.

        Centralized so every rebuild path (``recover_session`` lazy rebuild and
        ``recover_owned_process``) keeps the on-disk marker fresh; a ``None`` pid is
        ignored. The full owner+excel identity is recorded (create-time + image +
        session) so the next startup reconcile can verify an orphan before killing
        it and never kill on pid-alone. Best-effort: a persistence error must never
        fail a recovery.
        """

        if owned_pid is None:
            return
        with contextlib.suppress(Exception):
            owner_pid = os.getpid()
            o_ct, o_exe, o_sess = ExcelAppAdapter._capture_process_identity(owner_pid)
            e_ct, e_exe, e_sess = ExcelAppAdapter._capture_process_identity(owned_pid)
            record = build_owned_record(
                owner_pid=owner_pid,
                excel_pid=owned_pid,
                owner_create_time=o_ct,
                owner_exe=o_exe,
                owner_session_id=o_sess,
                excel_create_time=e_ct,
                excel_exe=e_exe,
                excel_session_id=e_sess,
            )
            self._owned_pid_store.write_record(record)

    def _finish_recovery(
        self,
        *,
        reason: str,
        owned_pid: int | None,
        prior_job: str,
        approval_source: str | None,
        discarded: int,
    ) -> dict[str, Any]:
        """Build + audit a fail-closed ``excel_recover`` outcome (no kill, no reset)."""

        outcome: dict[str, Any] = {
            "ok": False,
            "terminated": False,
            "confirmed": False,
            "recovery_possible": False,
            "process_restarted": False,
            "prior_job": prior_job,
            "what_lost": "none; originals are never modified",
            "reason": reason,
            "owned_pid": owned_pid,
            "approval_source": approval_source,
        }
        self._record_audit(
            tool="excel_recover",
            operation_class="recovery",
            params={"owned_pid": owned_pid, "discarded_workbooks": discarded},
            result={
                "ok": False,
                "recovery_possible": False,
                "prior_job": prior_job,
                "reason": reason,
            },
        )
        return outcome

    async def close_session(self, *, timeout_s: float = _DEFAULT_TIMEOUT_S) -> SessionStatus:
        """Close the owned Excel session (workbooks + Quit + release); idempotent.

        Closes every owned workbook with ``Close(SaveChanges=False)``, quits the
        owned instance and releases it, then marks the session closed. A second
        call is a no-op that still returns a truthful ``closed`` status. Only the
        automation-owned instance is touched; a user-owned Excel process is never
        enumerated, activated or terminated (BG05).

        Args:
            timeout_s: Accepted for signature symmetry; the worker's own bounded
                join is authoritative.

        Returns:
            A :class:`SessionStatus` with state ``closed`` (or the worker's
            error/closed state if the release failed).
        """

        if self._closed:
            return self._build_status(
                workbook_id=None,
                mode=self._session_mode,
                read_only=False,
                dirty=False,
                notes=["session already closed"],
            )

        self._accepting = False
        self._state = WorkbookSessionState.CLOSING
        discarded = self._discard_open_workbooks()
        with contextlib.suppress(Exception):
            await self._worker.shutdown()
        after = self._worker.get_status()
        self._closed = True
        self._state = after.state
        # The owned process is gone: clear the startup-reconcile pid marker so the
        # next server start does not mistake a dead pid for a live leftover (Block A).
        self._owned_pid_store.clear()
        self._record_audit(
            tool="close_session",
            operation_class="mutating",
            params={"discarded_workbooks": discarded},
            result={"needs_recovery": False, "state": after.state.value},
        )
        return self._build_status(
            workbook_id=None,
            mode=self._session_mode,
            read_only=False,
            dirty=False,
            notes=[f"closed owned workbooks={discarded}"],
            worker_status=after,
        )

    def _require_accepting_new_work(self, tool: str) -> None:
        """Refuse new work while recovering or after the session was closed."""

        if self._accepting and not self._closed:
            return
        raise ExcelHungError(
            "The Excel session is not accepting new work.",
            context={
                "tool": tool,
                "needs_recovery": True,
                "state": self._state.value,
                "closed": self._closed,
            },
        )

    async def _maybe_auto_recover(self) -> None:
        """Lazy auto-recovery fallback (F2b / strategy C).

        If the session is in ``NEEDS_RECOVERY`` because its automation-owned Excel
        PID is dead/stale (the process was terminated after a hung macro), attempt
        exactly ONE in-process :meth:`recover_session` so the next ``excel_open``
        self-heals without a server restart. When the owned PID is still alive the
        call is genuinely hung (an uninterruptible COM call in flight) and is left to
        the explicit recovery path — a reset cannot reap a busy STA thread.

        Best-effort: any failure is swallowed so :meth:`_require_accepting_new_work`
        reports the truthful state and never masks the original condition. This guard
        is deliberately narrow (open only), not sprinkled across every service.
        """

        if self._closed:
            return
        worker_status = self._worker.get_status()
        # ``getattr`` keeps this guard a no-op for minimal fakes whose ``get_status``
        # returns a bare state rather than a full ``WorkerStatus`` (the rest of the
        # service reads worker status the same defensive way).
        if getattr(worker_status, "state", None) != WorkbookSessionState.NEEDS_RECOVERY:
            return
        owned_pid = getattr(worker_status, "owned_pid", None)
        # Only a dead/stale owned pid is self-healed here; a live hung STA is not.
        if owned_pid is None or self._owned_pid_alive(owned_pid):
            return
        with contextlib.suppress(Exception):
            await self.recover_session(reason="lazy_auto_recovery")

    @staticmethod
    def _owned_pid_alive(pid: int) -> bool:
        """Read-only liveness probe delegating to the adapter (never raises)."""

        return ExcelAppAdapter._pid_alive(pid)

    def _discard_open_workbooks(self) -> int:
        """Drop all workbook bookkeeping (the owned instance is being reset).

        Releases each held per-path lock, removes the working copies and clears
        the lock registry. Returns the number of workbooks that were tracked.
        """

        records = list(self._records.values())
        self._records.clear()
        self._active_by_canonical.clear()
        for record in records:
            with contextlib.suppress(RuntimeError):
                if record.lock.locked():
                    record.lock.release()
            self._cleanup_working_copy(record.working_path)
        self._locks.clear()
        return len(records)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _build_status(
        self,
        *,
        workbook_id: str | None,
        mode: SessionMode,
        read_only: bool,
        dirty: bool,
        notes: list[str] | None = None,
        worker_status: Any = None,
        created_at: datetime | None = None,
    ) -> SessionStatus:
        """Assemble a :class:`SessionStatus` from worker + bookkeeping."""

        if worker_status is None:
            worker_status = self._worker.get_status()
        now = created_at or _utc_now()
        merged_notes = list(getattr(worker_status, "notes", []) or [])
        if notes:
            merged_notes.extend(notes)
        active_operation = getattr(worker_status, "active_operation", None)
        state = getattr(worker_status, "state", WorkbookSessionState.RUNNING)
        busy = bool(active_operation) or state == WorkbookSessionState.BUSY
        return SessionStatus(
            session_id=self._session_id,
            workbook_id=workbook_id,
            state=state,
            mode=mode,
            busy=busy,
            read_only=read_only,
            dirty=dirty,
            active_operation=active_operation,
            access_mode=self._approval.access_mode,
            policy=self._build_policy(workbook_id),
            notes=merged_notes,
            created_at=now,
            updated_at=_utc_now(),
        )

    def _build_policy(self, workbook_id: str | None) -> SessionPolicy:
        """Assemble the additive per-workbook ``SessionPolicy`` (U2; plan_03 §6.2).

        ``has_snapshot`` reads the in-memory gate (``SnapshotService.has_snapshot``
        -> ``_latest_snapshot``) when a workbook is identified; with no workbook it
        is ``False``. ``classification_is_authorization`` is always ``False`` and
        ``r13_preflight`` is always ``"veto"`` (enforced by the DTO literal types).
        """

        access_mode = self._approval.access_mode
        requires_grant = access_mode == AccessMode.RESTRICTED
        has_snapshot = (
            self._snapshot_gate(workbook_id)
            if workbook_id is not None and self._snapshot_gate is not None
            else False
        )
        return SessionPolicy(
            approval_mode=access_mode,
            mutating_requires_grant=requires_grant,
            executing_requires_grant=requires_grant,
            recovery_auto_allowed=True,
            snapshot_required_for_mutations=True,
            has_snapshot=has_snapshot,
            classification_is_authorization=False,
            r13_preflight="veto",
        )

    def _require_record(self, workbook_id: str) -> _SessionRecord:
        record = self._records.get(workbook_id)
        if record is None:
            raise WorkbookNotFoundError(
                f"Unknown workbook_id {workbook_id!r} (not open).",
                context={"workbook_id": workbook_id},
            )
        return record

    def _approval_key(self, canonical: str, action: str) -> ApprovalKey:
        return ApprovalKey(canonical_path=canonical, action=action)

    def _require_approval(
        self,
        key: ApprovalKey,
        operation_class: OperationClass,
    ) -> None:
        decision = self._approval.check(
            key,
            session_id=self._session_id,
            operation_class=operation_class,
        )
        if not decision.allowed:
            self._record_audit(
                tool=key.action,
                operation_class=operation_class.value,
                params={"action": key.action, "canonical_path": key.canonical_path},
                approval={
                    "allowed": False,
                    "source": decision.source.value if decision.source else None,
                    "reason": decision.reason,
                },
                result={"error_code": ErrorCode.APPROVAL_DENIED.value},
            )
            raise ApprovalDeniedError(
                f"Operation {key.action!r} denied by approval policy: {decision.reason}",
                context={
                    "action": key.action,
                    "canonical_path": key.canonical_path,
                    "reason": decision.reason,
                },
            )

    def _materialize_working_copy(self, canonical: Path, workbook_id: str) -> str:
        """Copy the original into ``work/`` and return the working copy path."""

        self._work_dir.mkdir(parents=True, exist_ok=True)
        suffix = canonical.suffix or ".xlsm"
        working = self._work_dir / f"{workbook_id}{suffix}"
        shutil.copy2(canonical, working)
        return str(working)

    def _cleanup_working_copy(self, working_path: str) -> None:
        """Best-effort removal of a working copy (never the original)."""

        # A leftover working copy is not fatal; the original is untouched.
        with contextlib.suppress(OSError):
            Path(working_path).unlink(missing_ok=True)

    def _record_audit(
        self,
        *,
        tool: str,
        operation_class: str,
        params: dict[str, Any],
        workbook: dict[str, Any] | None = None,
        approval: Mapping[str, Any] | None = None,
        result: dict[str, Any] | None = None,
    ) -> None:
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=self._session_id,
            workbook=workbook,
            approval=approval,
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


__all__ = ["SessionService"]
