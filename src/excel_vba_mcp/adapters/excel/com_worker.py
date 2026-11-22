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

Each item has an absolute ``time.monotonic()`` deadline. Two independent
deadlines exist:

* the **queue-wait** deadline enforced by ``asyncio.wait_for`` in :meth:`submit`
  (the item never started because the worker was busy on an earlier one), and
* the **execution** deadline the worker checks before it starts a queued item.

When a call exceeds its deadline the worker sets ``NEEDS_RECOVERY``, records the
active token/section, and raises a truthful ``EXCEL_HUNG`` to the awaiting
tool. The still-running call is **not** cancelled (an uninterruptible COM call
cannot be), and mutating/executing items are **never** replayed. The only exits
from ``NEEDS_RECOVERY`` are the hung ``fn`` returning or an explicit
:meth:`reset` / ``close``-section drain.

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
    ) -> None:
        self._adapter = app_adapter
        self._orphan_timeout_s = orphan_timeout_s

        # Lifecycle state (guarded by ``_lock`` where mutated from the loop side).
        self._lock = threading.Lock()
        self._state: WorkbookSessionState = WorkbookSessionState.STARTING
        self._accepting = False
        self._started = False
        self._stopped = False

        # Active operation bookkeeping (written by the worker thread).
        self._active_section: Section | None = None
        self._active_operation: str | None = None
        self._active_token: int | None = None
        self._last_hresult: int | None = None
        self._last_error: str | None = None
        self._notes: list[str] = []

        # Proxy-leak / diagnostic counters.
        self._leak_suspected = False

        # Owned-instance identity mirrors (set once create_hidden succeeds).
        self._owned_pid: int | None = None
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
        self._state = WorkbookSessionState.STARTING
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
        with self._lock:
            self._accepting = True
            self._state = WorkbookSessionState.RUNNING
        self._owned_pid = self._adapter.owned_pid
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

        Args:
            fn: A zero-argument callable that runs on the STA thread. It must
                touch only the owned Excel instance and return a serializable
                value (DTO / primitive / container).
            timeout: Soft deadline in seconds. Exceeding it yields a truthful
                ``EXCEL_HUNG`` and moves the worker to ``NEEDS_RECOVERY``.
            section: Work section tag (drives settings scope and recovery rules).
            op_class: Operation classification (drives retry eligibility).

        Returns:
            The value returned by ``fn`` (validated as JSON-serializable).

        Raises:
            WorkerUnavailableError: If the worker is not started, is stopping,
                or is shutting down.
            ExcelHungError: If a prior call timed out and the worker is in
                ``NEEDS_RECOVERY`` (fail fast — never enqueue), or if this call
                itself exceeds its deadline.
            ComBusyError / RpcCallRejectedError: For COM busy/rejected outcomes.
        """

        # Fail fast: not started / stopped / not accepting.
        if not self._started or self._stopped or not self._accepting:
            raise WorkerUnavailableError(
                "Excel COM worker is not available",
                context={"state": self._state.value, "accepting": self._accepting},
            )
        # Fail fast: a prior call may still be running on the STA thread.
        if self._state == WorkbookSessionState.NEEDS_RECOVERY and section != "close":
            raise ExcelHungError(
                "Excel COM worker needs recovery; refusing new work",
                code=ErrorCode.EXCEL_HUNG,
                context={"state": self._state.value, "section": section},
            )

        self._token_seq += 1
        token = self._token_seq
        fut: Future[T] = Future()
        item: _WorkItem = _WorkItem(
            fn=fn,
            future=fut,
            deadline=time.monotonic() + timeout,
            section=section,
            op_class=op_class,
            token=token,
        )
        self._queue.put(item)

        # ``wait_for`` bounds the queue-wait; the worker bounds execution via the
        # item deadline. Never blocks the loop.
        try:
            return await asyncio.wait_for(asyncio.wrap_future(fut), timeout=timeout)
        except TimeoutError:
            # Record active token/section BEFORE raising so get_status() is
            # truthful. Do not cancel the running call; do not replay.
            with self._lock:
                self._state = WorkbookSessionState.NEEDS_RECOVERY
            self._active_token = token
            self._active_section = section
            self._active_operation = getattr(fn, "__name__", None)
            self._note(f"submit timeout section={section} token={token}")
            raise ExcelHungError(
                "The Excel COM operation exceeded its deadline and may still be running",
                code=ErrorCode.EXCEL_HUNG,
                context={
                    "section": section,
                    "token": token,
                    "op_class": op_class.value,
                },
            ) from None

    def get_status(self) -> WorkerStatus:
        """Return a non-COM snapshot of worker bookkeeping (safe to call anytime)."""

        with self._lock:
            state = self._state
            accepting = self._accepting
        return WorkerStatus(
            state=state,
            accepting=accepting,
            active_section=self._active_section,
            active_operation=self._active_operation,
            last_hresult=self._last_hresult,
            owned_pid=self._owned_pid,
            last_error=self._last_error,
            notes=list(self._notes),
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
        with self._lock:
            self._accepting = False
            self._state = WorkbookSessionState.CLOSING

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
                with self._lock:
                    self._state = WorkbookSessionState.ERROR

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
        pid = self._owned_pid
        if pid is not None:
            reaped = await asyncio.to_thread(self._adapter.reap_owned_pid, pid)
            if not reaped:
                self._note(f"owned pid {pid} still alive after the shutdown fallback")

        self._started = False
        self._owned_pid = None
        self._owned_caption = None
        if thread is None or not thread.is_alive():
            with self._lock:
                self._state = WorkbookSessionState.CLOSED

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
            if self._accepting and self._state != WorkbookSessionState.CLOSING:
                self._state = WorkbookSessionState.RUNNING
        self._owned_pid = self._adapter.owned_pid
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
        if self._state == WorkbookSessionState.NEEDS_RECOVERY and item.section != "close":
            self._complete_exception(
                item,
                WorkerUnavailableError(
                    "Worker needs recovery; refusing new work",
                    context={"section": item.section, "token": item.token},
                ),
            )
            return

        # Record active operation before invoking fn (truthful get_status).
        self._active_token = item.token
        self._active_section = item.section
        self._active_operation = getattr(item.fn, "__name__", None)

        try:
            result = self._run_with_retries(item, pythoncom)
            validated = _ensure_dto(result)
            self._complete_result(item, validated)
        except ExcelVbaMcpError as exc:
            self._last_hresult = exc.hresult
            self._last_error = str(exc)
            self._complete_exception(item, exc)
        except Exception as exc:  # noqa: BLE001 - map any COM error
            self._handle_unexpected(item, exc)
        finally:
            self._clear_active()

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
        self._last_hresult = error.hresult
        self._last_error = str(error)
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
        self._owned_pid = adapter.owned_pid
        self._owned_caption = adapter.owned_caption
        if self._owned_pid is not None:
            self._note(f"owned pid {self._owned_pid} not verified dead by adapter quit")
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
        self._owned_pid = self._adapter.owned_pid
        self._owned_caption = self._adapter.owned_caption
        self._last_hresult = None
        self._last_error = None
        self._active_token = None
        self._active_section = None
        self._active_operation = None
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

    def _clear_active(self) -> None:
        self._active_token = None
        self._active_section = None
        self._active_operation = None
        # The hung fn finally returned: leave NEEDS_RECOVERY so reset()/next
        # close drain can proceed; the tool decides recovery explicitly.
        if self._state == WorkbookSessionState.NEEDS_RECOVERY:
            # Keep NEEDS_RECOVERY until an explicit reset succeeds; just record
            # that the call is no longer running.
            self._note("active call returned after NEEDS_RECOVERY")

    def _next_token(self) -> int:
        self._token_seq += 1
        return self._token_seq

    def _note(self, message: str) -> None:
        """Append a bounded diagnostic note (non-COM string only)."""

        self._notes.append(message)
        if len(self._notes) > 50:
            del self._notes[: len(self._notes) - 50]

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        return (
            f"ExcelComWorker(state={self._state.value}, accepting={self._accepting}, "
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
