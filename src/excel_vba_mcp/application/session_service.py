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
import shutil
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker, WorkerStatus
from excel_vba_mcp.adapters.excel.workbook_adapter import WorkbookAdapter
from excel_vba_mcp.domain.errors import ExcelHungError, ExcelVbaMcpError, MacroBlockedError
from excel_vba_mcp.domain.models import SessionStatus, WorkbookInfo
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.macro_policy import automation_security_for
from excel_vba_mcp.security.path_policy import PathPolicy

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


def _utc_now() -> datetime:
    return datetime.now(UTC)


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
    ) -> None:
        self._worker = worker
        self._app_adapter = app_adapter
        self._path_policy = path_policy
        self._approval = approval
        self._audit = audit
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

    # ------------------------------------------------------------------
    # Open
    # ------------------------------------------------------------------

    async def open_workbook(
        self,
        path: str,
        *,
        mode: SessionMode = SessionMode.INSPECT,
        read_only: bool = False,
        timeout_s: float = _DEFAULT_TIMEOUT_S,
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
            timeout_s: Soft per-call ceiling for the STA open.

        Returns:
            A :class:`WorkbookInfo` describing the opened working copy.

        Raises:
            PathNotAllowedError: If ``path`` is outside the allowlisted roots.
            MacroBlockedError: If the approval gate denies the open.
            ExcelVbaMcpError: If the workbook is already open on this path.
            ExcelHungError: If the STA open times out / needs recovery.
        """

        mode = SessionMode(mode)
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
            working_path = self._materialize_working_copy(canonical, workbook_id)

            try:
                info = await self._open_on_worker(
                    working_path=working_path,
                    canonical_str=canonical_str,
                    workbook_id=workbook_id,
                    mode=mode,
                    read_only=read_only,
                    timeout_s=timeout_s,
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
        timeout_s: float = _DEFAULT_TIMEOUT_S,
    ) -> SessionStatus:
        """Close an open workbook (default: no save) and release its lock.

        Args:
            workbook_id: The id returned by :meth:`open_workbook`.
            save_changes: Persist changes on close. Default ``False`` so an open
                is never silently saved (spec §7.1).
            timeout_s: Soft per-call ceiling for the STA close.

        Returns:
            A :class:`SessionStatus` describing the session after close.

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown.
            MacroBlockedError: If a save-on-close is not approved.
        """

        record = self._require_record(workbook_id)

        # A save-on-close is a mutation; a plain close is read-only (auto-allowed).
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
            params={"workbook_id": workbook_id, "save_changes": save_changes},
            workbook={"canonical_path": record.canonical},
            result={"workbook_id": workbook_id, "closed": True, "saved": save_changes},
        )
        return self._build_status(
            workbook_id=None,
            mode=record.mode,
            read_only=record.read_only,
            dirty=False,
            notes=["workbook closed"],
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
            MacroBlockedError: If the save is not approved.
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
                raise ExcelVbaMcpError(
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

        self._record_audit(
            tool="excel_save",
            operation_class="mutating",
            params={"workbook_id": workbook_id},
            workbook={"canonical_path": record.canonical},
            result={"workbook_id": workbook_id, "saved": True},
        )
        return info

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
            notes=merged_notes,
            created_at=now,
            updated_at=_utc_now(),
        )

    def _require_record(self, workbook_id: str) -> _SessionRecord:
        record = self._records.get(workbook_id)
        if record is None:
            raise ExcelVbaMcpError(
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
            raise MacroBlockedError(
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
        result: dict[str, Any] | None = None,
    ) -> None:
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=self._session_id,
            workbook=workbook,
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
