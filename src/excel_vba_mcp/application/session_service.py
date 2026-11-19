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
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.workbook_adapter import WorkbookAdapter
from excel_vba_mcp.domain.errors import ExcelVbaMcpError, MacroBlockedError
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
