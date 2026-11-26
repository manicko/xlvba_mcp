"""Dedicated single-STA Excel COM worker (task_010).

One ``threading.Thread`` owns the COM apartment for the whole server lifetime.
Every Excel COM call runs on that thread; MCP/API layers only ever
:meth:`ExcelComWorker.submit` awaitable callables and receive serializable
DTOs back. This satisfies ARCH-009 (sync tools run on anyio pool threads, so
they must never touch COM directly) while preserving a single-STA guarantee.

Concurrency model
-----------------

* **Apartment:** the worker thread calls
  ``pythoncom.CoInitializeEx(COINIT_APARTMENTTHREADED)`` exactly once at start
  and ``CoUninitialize`` exactly once at the very end. No other thread touches
  COM.
* **Queue:** a ``queue.SimpleQueue`` of :class:`_WorkItem`. ``SimpleQueue`` is
  FIFO and lock-light; the worker is the only consumer.
* **Awaitable submit:** each item carries a ``concurrent.futures.Future``.
  :meth:`submit` enqueues the item and awaits ``asyncio.wrap_future(fut)`` under
  ``asyncio.wait_for`` — the event loop thread never blocks.
* **Message pump:** the loop uses a bounded ``queue.get(timeout=0.05)`` and
  calls ``pythoncom.PumpWaitingMessages()`` on each idle timeout. It **never**
  pumps inside a running ``fn`` (that would reenter user code).

Timeout / recovery
------------------

Each item has an absolute ``time.monotonic()`` deadline (the per-call ``timeout``
clamped to the server liveness ceiling — the single per-op-class authority,
Block E). The deadline is split into two phases:

* the **queue-wait** phase — the item never started because the worker was busy
  on an earlier one. On expiry :meth:`submit` raises ``ComBusyError`` (``busy``)
  with a ``retry_after_s`` hint and **never** changes the worker state; and
* the **execution** phase — the item started (``_WorkItem.started`` is ``True``)
  but ``fn`` is still running. On expiry the worker moves to
  ``DEADLINE_EXCEEDED`` and raises a truthful ``EXCEL_HUNG``.

A deadline-exceeded call is **not** cancelled (an uninterruptible COM call cannot
be), the owned pid is **not** terminated (the server never auto-kills a live pid
— C1), and mutating/executing items are **never** replayed. When the call finally
returns, ``_clear_active`` resumes ``RUNNING`` (DEADLINE_EXCEEDED is not a poison
state). ``NEEDS_RECOVERY`` is reserved for genuine recovery (Block A); the only
exits from it are the hung ``fn`` returning or an explicit :meth:`reset` /
``close``-section drain.

Proxy-leak guard
----------------

A submitted ``fn`` may only return a Pydantic ``DomainModel``, a primitive, or
a container of those. :func:`_ensure_dto` validates the return boundary: a COM
proxy or any non-JSON-serializable object raises and the worker records
``PROXY_LEAK_SUSPECTED`` without storing it.

COM is imported lazily inside the worker thread / helper methods so importing
this module never requires ``pywin32`` and stays safe on non-Windows unit-test
environments.
"""

from __future__ import annotations

import asyncio
import contextlib
import dataclasses
import queue
import threading
import time
from collections.abc import Callable
from concurrent.futures import Future
from typing import Any, Literal, TypeVar

from pydantic import BaseModel

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.domain.errors import (
    ComBusyError,
    ErrorCode,
    ExcelHungError,
    ExcelVbaMcpError,
    RpcCallRejectedError,
    error_from_hresult,
)
from excel_vba_mcp.domain.models import DomainModel
from excel_vba_mcp.domain.statuses import OperationClass, WorkbookSessionState

T = TypeVar("T")

#: The section tag attached to every work item. ``close`` is special: it is the
#: recovery/shutdown drain and is allowed through even in ``NEEDS_RECOVERY``.
Section = Literal["open", "close", "save", "diagnostic", "execute"]

# Well-known COM HRESULTs (kept as plain ints; the domain layer stays COM-free).
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A

# READ_ONLY bounded-retry policy for RPC_E_CALL_REJECTED.
_READ_ONLY_MAX_RETRIES = 3
_RETRY_BACKOFFS_S = (0.5, 1.0, 2.0)

# Idle pump interval and shutdown orphan join default.
_IDLE_PUMP_TIMEOUT_S = 0.05

# How long ``reset`` waits for a hung ``fn`` to finish before it gives up and
# rebuilds anyway (the old thread is abandoned as a last resort).
_RESET_JOIN_TIMEOUT_S = 3.0

# Retry hint (seconds) returned to the caller on a queue-wait timeout (the item
# never started because the worker was busy on an earlier call). This is a soft
# hint only; the worker is single-STA and the caller should re-submit later.
_BUSY_RETRY_AFTER_S = 1.0


class WorkerUnavailableError(ExcelVbaMcpError):
    """The worker is not started / is stopped / is shutting down."""

    default_code = ErrorCode.EXCEL_HUNG
    default_message = "The Excel COM worker is unavailable."


# ---------------------------------------------------------------------------
# Non-COM bookkeeping DTO
# ---------------------------------------------------------------------------


class WorkerStatus(DomainModel):
    """Serializable snapshot of worker bookkeeping (never touches COM)."""

    state: WorkbookSessionState
    accepting: bool
    active_section: Section | None = None
    active_operation: str | None = None
    last_hresult: int | None = None
    owned_pid: int | None = None
    last_error: str | None = None
    notes: list[str] = []


@dataclasses.dataclass(frozen=True)
class _StatusSnapshot:
    """Immutable, lock-guarded status snapshot (the single source of truth).

    Every field :meth:`ExcelComWorker.get_status` returns lives in one frozen
    object. Writers build a *new* snapshot and swap it under ``_lock``; readers
    copy the snapshot reference under the same lock. Because the active triple
    (``active_token`` / ``active_section`` / ``active_operation``) is published
    and cleared as part of a single atomic swap, a concurrent reader can never
    observe a torn active op or a completed op alongside a stale active op
    (report_07 §2.4 D1 / plan_02 Work item D).
    """

    state: WorkbookSessionState
    accepting: bool
    active_section: Section | None = None
    active_operation: str | None = None
    active_token: int | None = None
    last_hresult: int | None = None
    owned_pid: int | None = None
    last_error: str | None = None
    notes: tuple[str, ...] = ()


# ---------------------------------------------------------------------------
# Work item
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class _WorkItem:
    """A single queued unit of work bound for the STA thread."""

    fn: Callable[[], Any]
    future: Future[Any]
    deadline: float  # time.monotonic()
    section: Section
    op_class: OperationClass
    token: int
    is_sentinel: bool = False
    # Set to ``True`` by the worker the moment ``_process_item`` begins handling
    # this item (before ``fn`` runs). The awaiter in :meth:`ExcelComWorker.submit`
    # reads it on a wait timeout to tell a never-started item (queue-wait timeout
    # -> ``busy``) apart from a started-but-slow one (exec deadline ->
    # ``DEADLINE_EXCEEDED``). Block E.
    started: bool = False

    @property
    def is_close_drain(self) -> bool:
        return self.section == "close"


# ---------------------------------------------------------------------------
# Return-boundary proxy-leak guard
# ---------------------------------------------------------------------------

_ALLOWED_PRIMITIVES = (str, int, float, bool, type(None), bytes)


def _ensure_dto(result: Any) -> Any:
    """Validate that ``result`` is a safe, JSON-serializable return value.

    A submitted ``fn`` may only return a Pydantic ``BaseModel``, a primitive, or
    a (possibly nested) container of those. A COM proxy or any other
    non-JSON-serializable object raises ``ExcelVbaMcpError`` so the worker can
    record ``PROXY_LEAK_SUSPECTED`` and refuse to store it.

    Returns:
        The validated result unchanged.
    """

    if isinstance(result, BaseModel):
        # Force serialization now so a proxy hidden inside a model field fails
        # at this boundary rather than later in the MCP layer.
        result.model_dump(mode="json")
        return result
    if isinstance(result, _ALLOWED_PRIMITIVES):
        return result
    if isinstance(result, dict):
        for key, value in result.items():
            if not isinstance(key, str):
                raise ExcelVbaMcpError(
                    "Non-string dict key in worker return value (possible proxy leak)",
                    code=ErrorCode.COM_BUSY,
                    context={"key_type": type(key).__name__},
                )
            _ensure_dto(value)
        return result
    if isinstance(result, (list, tuple, set, frozenset)):
        for item in result:
            _ensure_dto(item)
        return result
    raise ExcelVbaMcpError(
        f"Worker returned a non-serializable object of type {type(result).__name__} "
        "(possible COM proxy leak)",
        code=ErrorCode.COM_BUSY,
        context={"type": type(result).__name__},
    )


# ---------------------------------------------------------------------------
# HRESULT extraction
# ---------------------------------------------------------------------------


def _hresult_of(exc: BaseException) -> int | None:
    """Extract an HRESULT from a ``com_error``-like exception, if present."""

    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        try:
            return int(hresult) & 0xFFFFFFFF
        except (TypeError, ValueError):  # pragma: no cover - defensive
            return None
    args = getattr(exc, "args", ())
    if args:
        try:
            return int(args[0]) & 0xFFFFFFFF
        except (TypeError, ValueError):
            return None
    return None


# ---------------------------------------------------------------------------
# Worker
# ---------------------------------------------------------------------------


class ExcelComWorker:
    """Single dedicated COM STA thread; serialized FIFO; awaitable submit.

    The worker is created and owned by the server lifespan
    (``MCPServer(lifespan=...)``); see ARCH-005. :meth:`start` spawns the STA
    thread and awaits readiness; :meth:`submit` enqueues awaitable work;
    :meth:`shutdown` runs the ordered close→quit→release→uninitialize sequence.
    """

    def __init__(
        self,
        app_adapter: ExcelAppAdapter,
        *,
        orphan_timeout_s: float = 10.0,
        liveness_ceiling_s: float | None = None,
    ) -> None:
        self._adapter = app_adapter
        self._orphan_timeout_s = orphan_timeout_s
        # Server-wide upper bound applied to every per-call COM deadline (Block E).
        # ``None`` means no clamp (used by unit tests that construct the worker
        # directly); the server threads ``ServerConfig.liveness_ceiling_s`` here.
        self._liveness_ceiling_s = liveness_ceiling_s

        # Lifecycle state (guarded by ``_lock`` where mutated from the loop side).
        self._lock = threading.Lock()

        # Single lock-guarded immutable status snapshot — the ONLY source of
        # truth for ``get_status`` (report_07 §2.4 D1). Active bookkeeping,
        # lifecycle state, last error/hresult and owned pid all live here and are
        # mutated only by swapping a fresh snapshot under ``_lock``.
        self._snapshot = _StatusSnapshot(
            state=WorkbookSessionState.STARTING,
            accepting=False,
        )

        # Loop-side lifecycle flags (not part of the status snapshot).
        self._started = False
        self._stopped = False

        # Proxy-leak / diagnostic counters.
        self._leak_suspected = False

        # Owned-instance caption mirror (not surfaced by ``get_status``).
        self._owned_caption: str | None = None

        # Work queue + thread + startup handshake.
        self._queue: queue.SimpleQueue[_WorkItem | None] = queue.SimpleQueue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._startup_error: BaseException | None = None

        # Shutdown idempotency.
        self._shutdown_started = False
        self._drained = False

        # Monotonic token source for active-operation correlation.
        self._token_seq = 0

    # ------------------------------------------------------------------
    # Public API — lifecycle (loop side)
    # ------------------------------------------------------------------

    async def start(self) -> None:
        """Spawn the STA thread and wait until it is ready (or fails).

        Ready means: ``CoInitializeEx`` succeeded, ``create_hidden()`` set the
        instance ``Visible=False``, and the worker entered its run loop.

        Raises:
            WorkerUnavailableError: If the thread failed to start (COM init or
                ``DispatchEx`` error); the stored ``__cause__`` carries the
                original exception.
        """

        if self._started:
            return
        self._update_snapshot(state=WorkbookSessionState.STARTING)
        self._ready.clear()
        self._startup_error = None
        self._thread = threading.Thread(target=self._run, name="excel-com-sta", daemon=True)
        self._thread.start()

        # Await readiness without blocking the event loop. Poll the event in a
        # short loop so a crashed thread surfaces its startup error promptly.
        deadline = time.monotonic() + self._orphan_timeout_s
        while not self._ready.wait(timeout=0.05):
            if self._startup_error is not None:
                raise WorkerUnavailableError(
                    "Excel COM worker failed to start",
                    context={"reason": str(self._startup_error)},
                ) from self._startup_error
            if time.monotonic() > deadline:
                raise WorkerUnavailableError(
                    "Timed out waiting for the Excel COM worker to start",
                )

        if self._startup_error is not None:
            raise WorkerUnavailableError(
                "Excel COM worker failed to start",
                context={"reason": str(self._startup_error)},
            ) from self._startup_error

        self._started = True
        self._update_snapshot(
            accepting=True,
            state=WorkbookSessionState.RUNNING,
            owned_pid=self._adapter.owned_pid,
        )
        self._owned_caption = self._adapter.owned_caption

    async def submit(
        self,
        fn: Callable[[], T],
        *,
        timeout: float,
        section: Section = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> T:
        """Enqueue ``fn`` on the STA thread and await its result.

        The per-call ``timeout`` is a **server-side soft hint** (no client
        contract): it is clamped to :attr:`_liveness_ceiling_s` and becomes the
        single per-op-class server deadline that is the authority for this call
        (Block E). The deadline is split into two phases:

        * **Queue-wait** — the item never started because the worker was busy on
          an earlier call. On expiry this raises :class:`ComBusyError` (``busy``)
          with a ``retry_after_s`` hint and **never** changes the worker state
          (no ``NEEDS_RECOVERY`` / no poison).
        * **Execution** — the item started (``_WorkItem.started`` is ``True``) but
          ``fn`` is still running at the deadline. On expiry this moves the worker
          to ``DEADLINE_EXCEEDED`` and raises a truthful ``EXCEL_HUNG``. The
          still-running call is **not** cancelled (an uninterruptible COM call
          cannot be) and the owned pid is **not** terminated; the worker resumes
          ``RUNNING`` when the call returns (clear-before-complete).

        Args:
            fn: A zero-argument callable that runs on the STA thread. It must
                touch only the owned Excel instance and return a serializable
                value (DTO / primitive / container).
            timeout: Soft per-call deadline hint in seconds (clamped to the
                liveness ceiling).
            section: Work section tag (drives settings scope and recovery rules).
            op_class: Operation classification (drives retry eligibility).

        Returns:
            The value returned by ``fn`` (validated as JSON-serializable).

        Raises:
            WorkerUnavailableError: If the worker is not started, is stopping,
                or is shutting down.
            ExcelHungError: If a prior call timed out and the worker is in
                ``NEEDS_RECOVERY`` (fail fast — never enqueue), or if this call
                itself exceeds its execution deadline.
            ComBusyError: If the item never started before the queue-wait
                deadline (the worker was busy); ``context['retry_after_s']`` is a
                soft retry hint and the worker state is unchanged.
        """

        # Fail fast: not started / stopped / not accepting.
        snapshot = self._snapshot
        if not self._started or self._stopped or not snapshot.accepting:
            raise WorkerUnavailableError(
                "Excel COM worker is not available",
                context={"state": snapshot.state.value, "accepting": snapshot.accepting},
            )
        # Fail fast: a prior call may still be running on the STA thread.
        if snapshot.state == WorkbookSessionState.NEEDS_RECOVERY and section != "close":
            raise ExcelHungError(
                "Excel COM worker needs recovery; refusing new work",
                code=ErrorCode.EXCEL_HUNG,
                context={"state": snapshot.state.value, "section": section},
            )

        # The server deadline is the authority: clamp the soft hint to the
        # liveness ceiling so no single op's deadline exceeds it (Block E).
        effective_timeout = (
            timeout
            if self._liveness_ceiling_s is None
            else min(timeout, self._liveness_ceiling_s)
        )

        self._token_seq += 1
        token = self._token_seq
        fut: Future[T] = Future()
        item: _WorkItem = _WorkItem(
            fn=fn,
            future=fut,
            deadline=time.monotonic() + effective_timeout,
            section=section,
            op_class=op_class,
            token=token,
        )
        self._queue.put(item)

        # ``wait_for`` bounds the queue-wait; the worker bounds execution via the
        # item deadline. Never blocks the loop.
        try:
            return await asyncio.wait_for(asyncio.wrap_future(fut), timeout=effective_timeout)
        except TimeoutError:
            # Distinguish never-started (queue-wait) from started+slow (exec) via
            # the ``started`` flag the worker sets in ``_process_item`` (Block E).
            if not item.started:
                # Queue-wait timeout: the item never reached the STA thread. This
                # is a transient ``busy`` condition; NEVER poison the worker state
                # and never set NEEDS_RECOVERY. The abandoned future is completed
                # later by the worker (its exception is suppressed if set).
                self._note(
                    f"submit queue-wait timeout section={section} token={token} (busy)"
                )
                raise ComBusyError(
                    "The Excel COM worker is busy; the operation did not start before "
                    "its deadline. Retry later.",
                    code=ErrorCode.COM_BUSY,
                    context={
                        "section": section,
                        "token": token,
                        "op_class": op_class.value,
                        "reason": "queue_wait_timeout",
                        "retry_after_s": _BUSY_RETRY_AFTER_S,
                    },
                ) from None
            # Execution deadline exceeded: the call started but is still running.
            # Publish the active triple as one snapshot swap, then DEADLINE_EXCEEDED
            # as a second swap, BEFORE raising so get_status() is truthful. The
            # running call is not cancelled and not replayed; the worker resumes
            # RUNNING when it returns (see ``_clear_active``).
            self._set_active(token, section, getattr(fn, "__name__", None))
            self._update_snapshot(state=WorkbookSessionState.DEADLINE_EXCEEDED)
            self._note(f"submit deadline exceeded section={section} token={token}")
            raise ExcelHungError(
                "The Excel COM operation exceeded its deadline and may still be running",
                code=ErrorCode.EXCEL_HUNG,
                context={
                    "section": section,
                    "token": token,
                    "op_class": op_class.value,
                    "deadline_exceeded": True,
                    "verdict": "SUSPECTED_STALLED",
                },
            ) from None

    def get_status(self) -> WorkerStatus:
        """Return a non-COM snapshot of worker bookkeeping (safe to call anytime).

        The active triple and every other field are read from a single immutable
        :class:`_StatusSnapshot` copied under ``_lock`` — never lock-free — so a
        concurrent writer can never expose a torn or stale active op.
        """

        with self._lock:
            snap = self._snapshot
        return WorkerStatus(
            state=snap.state,
            accepting=snap.accepting,
            active_section=snap.active_section,
            active_operation=snap.active_operation,
            last_hresult=snap.last_hresult,
            owned_pid=snap.owned_pid,
            last_error=snap.last_error,
            notes=list(snap.notes),
        )

    async def shutdown(self) -> None:
        """Ordered, idempotent shutdown of the worker and owned Excel instance.

        Sequence (research §5): stop accepting → enqueue ``close`` drain +
        sentinel → complete queued futures with ``WorkerUnavailableError`` →
        close owned workbooks (alerts suppressed) → ``Quit`` (which verifies the
        owned process death, with a pid-scoped fallback) → release refs →
        ``CoUninitialize`` → bounded ``thread.join(orphan_timeout)`` → a final
        pid-scoped reap on the calling thread if the daemon STA thread never
        finished the drain. The ``close`` section is allowed even in
        ``NEEDS_RECOVERY``.
        """

        if self._shutdown_started:
            return
        self._shutdown_started = True

        # 1. Stop accepting new work.
        self._update_snapshot(
            accepting=False,
            state=WorkbookSessionState.CLOSING,
        )

        # 2. Enqueue the close drain (allowed even in NEEDS_RECOVERY) + sentinel.
        close_item = self._make_internal_item(section="close")
        self._queue.put(close_item)
        self._queue.put(None)  # sentinel ends the run loop

        # Wait for the run loop to finish the drain + uninitialize (bounded).
        thread = self._thread
        if thread is not None:
            await asyncio.to_thread(thread.join, self._orphan_timeout_s)
            if thread.is_alive():
                self._note("shutdown timed out joining worker thread")
                # Do not hang the server; the daemon thread is the last resort.
                self._update_snapshot(state=WorkbookSessionState.ERROR)

        # 3. Complete any futures still pending (defensive; the run loop should
        #    already have drained the queue).
        self._fail_pending_futures()

        # 4. Enforce the owned-instance invariant as the last resort (FE-17).
        #    ``_drain_close`` normally reaps the owned process through the
        #    adapter's verified quit; if the daemon STA thread is still alive
        #    after the bounded join it may never have run, so reap the owned pid
        #    here. The pid is the authoritative identity: killing the process
        #    does not need the (possibly dead) STA thread, and only the
        #    adapter's own pid is ever passed.
        pid = self._snapshot.owned_pid
        if pid is not None:
            reaped = await asyncio.to_thread(self._adapter.reap_owned_pid, pid)
            if not reaped:
                self._note(f"owned pid {pid} still alive after the shutdown fallback")

        self._started = False
        self._owned_caption = None
        self._update_snapshot(owned_pid=None)
        if thread is None or not thread.is_alive():
            self._update_snapshot(state=WorkbookSessionState.CLOSED)

    async def reset(self) -> None:
        """Rebuild the owned Excel instance on the same STA thread.

        The only exit from ``NEEDS_RECOVERY`` besides the hung ``fn`` returning.
        Closes/Quits the current instance, drops refs, and creates a fresh
        ``create_hidden()`` instance — all inside the worker thread.
        """

        # Allow a close-section reset even while not accepting.
        fut: Future[bool] = Future()
        item: _WorkItem = _WorkItem(
            fn=self._do_reset,
            future=fut,
            deadline=time.monotonic() + self._orphan_timeout_s,
            section="close",
            op_class=OperationClass.MUTATING,
            token=self._next_token(),
        )
        self._queue.put(item)
        # reset itself is the recovery action; await it directly (bounded).
        try:
            await asyncio.wait_for(asyncio.wrap_future(fut), timeout=self._orphan_timeout_s)
        except TimeoutError:
            raise ExcelHungError(
                "Timed out while resetting the Excel COM worker",
                code=ErrorCode.EXCEL_HUNG,
            ) from None
        # Only mark RUNNING if the reset succeeded and we are still accepting.
        with self._lock:
            snap = self._snapshot
            if snap.accepting and snap.state != WorkbookSessionState.CLOSING:
                self._snapshot = dataclasses.replace(
                    snap,
                    state=WorkbookSessionState.RUNNING,
                    owned_pid=self._adapter.owned_pid,
                )
            else:
                self._snapshot = dataclasses.replace(
                    snap, owned_pid=self._adapter.owned_pid
                )
        self._owned_caption = self._adapter.owned_caption

    # ------------------------------------------------------------------
    # Worker thread (internal)
    # ------------------------------------------------------------------

    def _run(self) -> None:
        """Worker thread body: CoInitializeEx → ready → loop → drain → uninit."""

        pythoncom = self._import_pythoncom()

        # Apartment init (once).
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as exc:  # RPC_E_CHANGED_MODE is a hard error here
            self._startup_error = exc
            self._ready.set()
            return

        # Create the owned hidden instance (still on this thread).
        try:
            self._adapter.create_hidden()
        except Exception as exc:
            self._startup_error = exc
            self._cleanup_com(pythoncom)
            self._ready.set()
            return

        # Ready: signal start() and enter the loop.
        self._ready.set()

        try:
            self._loop(pythoncom)
        finally:
            self._cleanup_com(pythoncom)

    def _loop(self, pythoncom: Any) -> None:
        """Process queued work items until the shutdown sentinel arrives."""

        while True:
            try:
                item = self._queue.get(timeout=_IDLE_PUMP_TIMEOUT_S)
            except queue.Empty:
                # Idle: service cross-apartment/queued calls without reentering
                # user code. Never pump inside a running fn.
                with contextlib.suppress(Exception):
                    pythoncom.PumpWaitingMessages()
                if self._shutdown_started:
                    # Sentinel may be queued behind a slow producer; keep waiting.
                    continue
                continue

            if item is None:  # shutdown sentinel
                # Drain any close-section items already queued before finishing.
                self._drain_remaining()
                return

            self._process_item(item, pythoncom)

    def _drain_remaining(self) -> None:
        """After the sentinel, run any close-section drain items still queued."""

        while True:
            try:
                item = self._queue.get_nowait()
            except queue.Empty:
                return
            if item is None:
                continue
            if item.is_close_drain:
                self._process_item(item, self._import_pythoncom())
            else:
                with contextlib.suppress(Exception):
                    item.future.set_exception(WorkerUnavailableError("Worker shutting down"))

    def _process_item(self, item: _WorkItem, pythoncom: Any) -> None:
        """Execute one work item with deadline / recovery / retry handling."""

        # Expired before start: do not invoke fn.
        if item.is_sentinel:  # pragma: no cover - sentinel handled earlier
            return
        # Mark the item as started the moment the worker begins handling it, so a
        # concurrent awaiter can tell "never started" (queue-wait -> busy) from
        # "started + slow" (exec deadline -> DEADLINE_EXCEEDED) (Block E). This is
        # set before the pre-start expiry check so an item the worker picked up but
        # found already expired is still reported as started (it was dequeued).
        item.started = True
        if time.monotonic() > item.deadline and not item.is_close_drain:
            self._complete_exception(
                item,
                ExcelHungError(
                    "The work item expired before it could start",
                    code=ErrorCode.EXCEL_HUNG,
                    context={"section": item.section, "token": item.token},
                ),
            )
            return

        # Recovery hold: refuse non-close items while NEEDS_RECOVERY.
        if self._snapshot.state == WorkbookSessionState.NEEDS_RECOVERY and item.section != "close":
            self._complete_exception(
                item,
                WorkerUnavailableError(
                    "Worker needs recovery; refusing new work",
                    context={"section": item.section, "token": item.token},
                ),
            )
            return

        # Record active operation before invoking fn (truthful get_status) via a
        # single atomic snapshot swap.
        self._set_active(item.token, item.section, getattr(item.fn, "__name__", None))

        try:
            result = self._run_with_retries(item, pythoncom)
            validated = _ensure_dto(result)
        except ExcelVbaMcpError as exc:
            self._record_error(exc.hresult, str(exc))
            # Clear-before-complete: publish the cleared active snapshot BEFORE
            # delivering the result, so a concurrent reader never observes a
            # completed op alongside a stale active op (report_07 §2.4 D1).
            self._clear_active()
            self._complete_exception(item, exc)
            return
        except Exception as exc:  # noqa: BLE001 - map any COM error
            self._clear_active()
            self._handle_unexpected(item, exc)
            return

        # Success: clear active BEFORE publishing the result (clear-before-complete).
        self._clear_active()
        self._complete_result(item, validated)

    def _run_with_retries(self, item: _WorkItem, pythoncom: Any) -> Any:
        """Invoke ``item.fn`` with the READ_ONLY bounded retry policy."""

        attempts = 0
        while True:
            try:
                return item.fn()
            except ExcelVbaMcpError:
                # Domain errors already carry a precise code (preconditions, guards,
                # policy). Never re-wrap them into a generic COM error, which would
                # mask the real reason (e.g. "last visible sheet" -> COM_BUSY).
                raise
            except Exception as exc:  # noqa: BLE001 - classified below
                hresult = _hresult_of(exc)
                if hresult == RPC_E_CALL_REJECTED:
                    # READ_ONLY: bounded retry with pump-between; else no replay.
                    if (
                        item.op_class == OperationClass.READ_ONLY
                        and attempts < _READ_ONLY_MAX_RETRIES
                    ):
                        backoff = _RETRY_BACKOFFS_S[min(attempts, len(_RETRY_BACKOFFS_S) - 1)]
                        self._sleep_pumping(pythoncom, backoff)
                        attempts += 1
                        continue
                    raise RpcCallRejectedError(
                        "Excel rejected the COM call (RPC_E_CALL_REJECTED)",
                        code=ErrorCode.RPC_E_CALL_REJECTED,
                        hresult=hresult,
                        context={
                            "attempts": attempts + 1,
                            "class": item.op_class.value,
                            "replayed": item.op_class == OperationClass.READ_ONLY,
                        },
                    ) from exc
                if hresult == RPC_E_SERVERCALL_RETRYLATER:
                    raise ComBusyError(
                        "Excel is busy (RPC_E_SERVERCALL_RETRYLATER)",
                        code=ErrorCode.COM_BUSY,
                        hresult=hresult,
                        context={"class": item.op_class.value},
                    ) from exc
                # Unknown HRESULT / non-COM error: map via domain helper.
                raise error_from_hresult(
                    hresult if hresult is not None else 0,
                    str(exc) or exc.__class__.__name__,
                    context={"class": item.op_class.value},
                ) from exc

    def _handle_unexpected(self, item: _WorkItem, exc: Exception) -> None:
        """Map an unexpected exception from a work item to a domain error."""

        hresult = _hresult_of(exc)
        if hresult is not None:
            error: ExcelVbaMcpError = error_from_hresult(
                hresult,
                str(exc) or exc.__class__.__name__,
                context={"class": item.op_class.value},
            )
        else:
            # No HRESULT (e.g. a plain Python error or a hang-shaped failure).
            error = ExcelVbaMcpError(
                str(exc) or exc.__class__.__name__,
                code=ErrorCode.COM_BUSY,
                context={"class": item.op_class.value, "type": type(exc).__name__},
            )
        self._record_error(error.hresult, str(error))
        self._complete_exception(item, error)

    # ------------------------------------------------------------------
    # Shutdown / reset internals (run on the STA thread)
    # ------------------------------------------------------------------

    def _make_internal_item(self, *, section: Section) -> _WorkItem:
        """Build a close-section drain item that closes the instance + quits."""

        fut: Future[bool] = Future()
        return _WorkItem(
            fn=self._drain_close,
            future=fut,
            deadline=time.monotonic() + self._orphan_timeout_s,
            section=section,
            op_class=OperationClass.MUTATING,
            token=self._next_token(),
        )

    def _drain_close(self) -> bool:
        """Close owned workbooks, Quit, release refs, drop pid (STA thread).

        Idempotent: a second call (e.g. from :meth:`_cleanup_com`) is a no-op so
        ``Quit`` runs at most once. The worker keeps the owned pid until the
        adapter's quit has *verified* the process death (poll + pid-scoped
        fallback), so teardown never drops the identity of a surviving process.
        """

        if self._drained:
            return True
        self._drained = True

        adapter = self._adapter
        # Close owned workbooks best-effort (alerts suppressed).
        wb = getattr(adapter, "_app", None)
        if wb is not None:
            with contextlib.suppress(Exception):
                workbooks = wb.Workbooks
                # Snapshot names so closing does not mutate the collection mid-loop.
                names = [w.Name for w in workbooks]
                for name in names:
                    with contextlib.suppress(Exception):
                        target = workbooks(name)
                        adapter.close_workbook(target, save_changes=False)
        # Quit (enforces the owned process death) + release. Mirror the adapter
        # identity only after the verified quit: a still-surviving pid is kept
        # so :meth:`shutdown` can retry the pid-scoped fallback.
        with contextlib.suppress(Exception):
            adapter.quit()
        pid = adapter.owned_pid
        self._owned_caption = adapter.owned_caption
        self._update_snapshot(owned_pid=pid)
        if pid is not None:
            self._note(f"owned pid {pid} not verified dead by adapter quit")
        with contextlib.suppress(Exception):
            adapter.release()
        self._note("drain_close complete")
        return True

    def _do_reset(self) -> bool:
        """Rebuild the owned instance on the same STA thread (recovery exit)."""

        # Tear down the current instance first.
        with contextlib.suppress(Exception):
            self._drain_close()
        # Create a fresh hidden instance.
        self._drained = False
        self._adapter.create_hidden()
        self._owned_caption = self._adapter.owned_caption
        # One atomic swap: fresh pid, cleared error + active bookkeeping.
        self._update_snapshot(
            owned_pid=self._adapter.owned_pid,
            last_hresult=None,
            last_error=None,
            active_token=None,
            active_section=None,
            active_operation=None,
        )
        self._note("reset rebuilt owned instance")
        return True

    def _cleanup_com(self, pythoncom: Any) -> None:
        """Best-effort teardown + CoUninitialize on the worker thread."""

        with contextlib.suppress(Exception):
            self._drain_close()
        with contextlib.suppress(Exception):
            self._adapter.release()
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _import_pythoncom() -> Any:
        import pythoncom  # type: ignore[import-untyped]  # noqa: PLC0415

        return pythoncom

    def _sleep_pumping(self, pythoncom: Any, seconds: float) -> None:
        """Sleep ``seconds`` while pumping messages (bounded, READ_ONLY only)."""

        deadline = time.monotonic() + seconds
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            with contextlib.suppress(Exception):
                pythoncom.PumpWaitingMessages()
            time.sleep(min(0.05, remaining))

    def _fail_pending_futures(self) -> None:
        """Complete any still-pending futures with WorkerUnavailableError."""

        while True:
            try:
                item = self._queue.get_nowait()
            except queue.Empty:
                return
            if item is None:
                continue
            with contextlib.suppress(Exception):
                item.future.set_exception(WorkerUnavailableError("Worker is shutting down"))

    def _complete_result(self, item: _WorkItem, result: Any) -> None:
        with contextlib.suppress(Exception):
            item.future.set_result(result)

    def _complete_exception(self, item: _WorkItem, exc: BaseException) -> None:
        with contextlib.suppress(Exception):
            item.future.set_exception(exc)

    # -- Atomic status-snapshot mutators (all swap under ``_lock``) ----------

    def _update_snapshot(self, **changes: Any) -> None:
        """Build a fresh snapshot with ``changes`` applied and swap it atomically."""

        with self._lock:
            self._snapshot = dataclasses.replace(self._snapshot, **changes)

    def _set_active(self, token: int, section: Section, operation: str | None) -> None:
        """Publish the active operation triple in one atomic snapshot swap."""

        self._update_snapshot(
            active_token=token,
            active_section=section,
            active_operation=operation,
        )

    def _record_error(self, hresult: int | None, message: str) -> None:
        """Record the last hresult/error in one atomic snapshot swap."""

        self._update_snapshot(last_hresult=hresult, last_error=message)

    def _clear_active(self) -> None:
        # Clear the active triple atomically. If the hung fn finally returned
        # while NEEDS_RECOVERY, keep NEEDS_RECOVERY (reset()/next close drain
        # proceeds; the tool decides recovery) and record the transition in the
        # SAME swap so the note and the cleared triple publish together.
        #
        # If the call was in DEADLINE_EXCEEDED (Block E: started + slow, no kill),
        # the deadline state is NOT a poison: clear it and resume RUNNING (when
        # still accepting) in the same swap, so a slow-but-alive op never leaves a
        # stale deadline state behind.
        with self._lock:
            snap = self._snapshot
            notes = snap.notes
            new_state = snap.state
            if snap.state == WorkbookSessionState.NEEDS_RECOVERY:
                notes = (notes + ("active call returned after NEEDS_RECOVERY",))[-50:]
            elif snap.state == WorkbookSessionState.DEADLINE_EXCEEDED:
                new_state = (
                    WorkbookSessionState.RUNNING if snap.accepting else snap.state
                )
                notes = (notes + ("deadline-exceeded call returned; resumed",))[-50:]
            self._snapshot = dataclasses.replace(
                snap,
                active_token=None,
                active_section=None,
                active_operation=None,
                state=new_state,
                notes=notes,
            )

    def _next_token(self) -> int:
        self._token_seq += 1
        return self._token_seq

    def _note(self, message: str) -> None:
        """Append a bounded diagnostic note (non-COM string only)."""

        with self._lock:
            notes = (self._snapshot.notes + (message,))[-50:]
            self._snapshot = dataclasses.replace(self._snapshot, notes=notes)

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        snap = self._snapshot
        return (
            f"ExcelComWorker(state={snap.state.value}, accepting={snap.accepting}, "
            f"started={self._started})"
        )


# ``ExcelHungError`` is referenced by name in the public contract; it is
# imported above with the other domain errors and re-exported here.
__all__ = [
    "ExcelComWorker",
    "ExcelHungError",
    "Section",
    "WorkerStatus",
    "WorkerUnavailableError",
]
