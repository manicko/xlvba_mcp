"""Unit tests for the single-STA Excel COM worker (task_010).

These tests never open Excel. They inject a fake ``ExcelAppAdapter`` so the
worker thread's ``CoInitializeEx`` / ``DispatchEx`` path is exercised without a
live Excel instance. A fake COM error (``FakeComError``) carrying an
``hresult`` is used to drive the busy/rejected classification, since
``pythoncom.com_error`` may not be importable everywhere.

Covered:
* FIFO ordering and serialization on a single STA thread.
* Awaitable submit does not block the event loop (slow fn + concurrent tick).
* Block E: a started-but-slow call -> ``DEADLINE_EXCEEDED`` + truthful
  ``EXCEL_HUNG`` (no poison, worker resumes RUNNING); a never-started item ->
  ``busy``/``retry_after_s`` (``ComBusyError``) with the worker state unchanged.
* Shutdown ordering and idempotency (fake adapter records calls).
* Shutdown reaps the owned pid only after the drain's verified quit (FE-17):
  no fallback when the quit ended the process, a pid-scoped fallback otherwise.
* Proxy-leak guard raises on a non-serializable return value.
* READ_ONLY bounded retry (<=3) vs MUTATING no-retry for ``RPC_E_CALL_REJECTED``.
* Settings context managers restore on exception (fake app object).
"""

from __future__ import annotations

import asyncio
import threading
import time
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import (
    ExcelComWorker,
    WorkerStatus,
    WorkerUnavailableError,
)
from excel_vba_mcp.domain.errors import (
    ComBusyError,
    ErrorCode,
    ExcelHungError,
    ExcelVbaMcpError,
    RpcCallRejectedError,
)
from excel_vba_mcp.domain.statuses import OperationClass, WorkbookSessionState

# HRESULTs exercised by the retry policy.
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A

pythoncom = pytest.importorskip(
    "pythoncom", reason="pywin32 required for the COM STA worker unit tests"
)


# ---------------------------------------------------------------------------
# Fakes
# ---------------------------------------------------------------------------


class FakeComError(Exception):
    """Minimal stand-in for ``pythoncom.com_error`` carrying an ``hresult``."""

    def __init__(self, hresult: int) -> None:
        super().__init__(f"com_error 0x{hresult & 0xFFFFFFFF:08X}")
        self.hresult = hresult


class FakeAppObject:
    """Fake ``Application`` for the settings-restore test."""

    def __init__(self) -> None:
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 3


class FakeAdapter(ExcelAppAdapter):
    """Records lifecycle calls without touching COM / Excel."""

    def __init__(self) -> None:
        super().__init__()
        self.create_count = 0
        self.quit_count = 0
        self.release_count = 0
        self.create_hook: Any = None

    def create_hidden(self) -> None:
        self.create_count += 1
        self._owned_pid = 4242
        self._owned_caption = "fake"
        if self.create_hook is not None:
            self.create_hook()

    def quit(self) -> None:
        self.quit_count += 1
        self._app = None
        self._owned_pid = None
        self._owned_caption = None

    def release(self) -> None:
        self.release_count += 1
        self._app = None
        self._owned_pid = None
        self._owned_caption = None


class VerifyingAdapter(FakeAdapter):
    """Fake adapter whose ``quit`` may leave the owned process resident (FE-17).

    ``quit_ends_process`` models the two outcomes of a real ``Application.Quit``:
    ``True`` (the process dies, the adapter clears its identity) and ``False``
    (the worst case: the process survives, so ``reap_owned_pid`` must terminate
    it). ``reap_owned_pid`` records the pid the worker still holds when it runs,
    which proves teardown does not drop the identity before the verified
    shutdown.
    """

    def __init__(self, *, quit_ends_process: bool = False) -> None:
        super().__init__()
        self.quit_ends_process = quit_ends_process
        self.reap_calls: list[int] = []
        self.worker_owned_pid_at_reap: list[int | None] = []
        self.worker: ExcelComWorker | None = None

    def quit(self) -> None:
        self.quit_count += 1
        self._app = None
        if self.quit_ends_process:
            self._owned_pid = None
            self._owned_caption = None

    def reap_owned_pid(self, pid: int) -> bool:
        self.reap_calls.append(pid)
        if self.worker is not None:
            self.worker_owned_pid_at_reap.append(self.worker.get_status().owned_pid)
        self._owned_pid = None
        self._owned_caption = None
        return True


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


async def _start(adapter: FakeAdapter | None = None) -> tuple[FakeAdapter, ExcelComWorker]:
    adapter = adapter or FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()
    return adapter, worker


# ---------------------------------------------------------------------------
# FIFO ordering + serialization
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_requests_complete_in_fifo_order_on_one_thread() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()

    order: list[str] = []
    thread_ids: list[int] = []

    def make(name: str):
        def fn() -> str:
            thread_ids.append(threading.get_ident())
            order.append(name)
            return name

        return fn

    first = await worker.submit(make("first"), timeout=2)
    second = await worker.submit(make("second"), timeout=2)
    third = await worker.submit(make("third"), timeout=2)

    assert (first, second, third) == ("first", "second", "third")
    assert order == ["first", "second", "third"]
    assert len(set(thread_ids)) == 1, "all calls must run on one STA thread"

    await worker.shutdown()


@pytest.mark.anyio
async def test_concurrent_submits_are_serialized() -> None:
    """Two concurrent submits run one-at-a-time in submission order."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    active = 0
    max_active = 0
    seq: list[int] = []
    lock = threading.Lock()

    def make(idx: int):
        def fn() -> int:
            nonlocal active, max_active
            with lock:
                active += 1
                max_active = max(max_active, active)
            time.sleep(0.05)
            with lock:
                seq.append(idx)
                active -= 1
            return idx

        return fn

    results = await asyncio.gather(
        worker.submit(make(0), timeout=3),
        worker.submit(make(1), timeout=3),
        worker.submit(make(2), timeout=3),
    )

    assert results == [0, 1, 2]
    assert seq == [0, 1, 2]
    assert max_active == 1, "the worker must never run two items concurrently"

    await worker.shutdown()


# ---------------------------------------------------------------------------
# Awaitable submit does not block the event loop
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_slow_fn_does_not_block_event_loop() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    ticks = 0

    async def ticker() -> None:
        nonlocal ticks
        for _ in range(5):
            await asyncio.sleep(0.02)
            ticks += 1

    def slow() -> str:
        time.sleep(0.3)
        return "slow-done"

    task = asyncio.create_task(ticker())
    result = await worker.submit(slow, timeout=3)
    await task

    assert result == "slow-done"
    assert ticks == 5, "the event loop must keep ticking while fn runs"

    await worker.shutdown()


# ---------------------------------------------------------------------------
# Block E — split queue-wait (busy) vs execution deadline (DEADLINE_EXCEEDED)
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_exec_deadline_sets_deadline_exceeded_and_raises_hung() -> None:
    """A started-but-slow call sets DEADLINE_EXCEEDED (not NEEDS_RECOVERY)."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def slow() -> str:
        time.sleep(1.0)
        return "done"

    with pytest.raises(ExcelHungError) as excinfo:
        await worker.submit(slow, timeout=0.1)

    assert excinfo.value.code == ErrorCode.EXCEL_HUNG
    assert excinfo.value.context.get("deadline_exceeded") is True
    assert excinfo.value.context.get("verdict") == "SUSPECTED_STALLED"
    status = worker.get_status()
    assert status.state == WorkbookSessionState.DEADLINE_EXCEEDED
    # The deadline state is NOT a poison: the worker resumes RUNNING once the
    # uninterruptible call finally returns (Block E, no kill).
    await asyncio.sleep(1.2)
    assert worker.get_status().state == WorkbookSessionState.RUNNING

    await worker.shutdown()


@pytest.mark.anyio
async def test_queue_wait_timeout_returns_busy_without_needs_recovery() -> None:
    """A never-started item yields busy/retry_after_s and never poisons the worker."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    release = threading.Event()

    def blocker() -> str:
        release.wait(2.0)
        return "blocked"

    # Occupy the single STA with a long op so the second item cannot start.
    blocker_task = asyncio.create_task(worker.submit(blocker, timeout=5.0))
    await asyncio.sleep(0.1)

    started = threading.Event()

    def should_not_run() -> str:
        started.set()
        return "x"

    with pytest.raises(ComBusyError) as excinfo:
        await worker.submit(should_not_run, timeout=0.1)

    assert excinfo.value.code == ErrorCode.COM_BUSY
    assert excinfo.value.context.get("reason") == "queue_wait_timeout"
    assert "retry_after_s" in excinfo.value.context
    # The worker state is NOT poisoned by a queue-wait timeout.
    assert worker.get_status().state != WorkbookSessionState.NEEDS_RECOVERY
    assert worker.get_status().state != WorkbookSessionState.DEADLINE_EXCEEDED
    await asyncio.sleep(0.2)
    assert not started.is_set(), "the queue-wait-timeout item must not have started"

    release.set()
    await blocker_task
    await worker.shutdown()


@pytest.mark.anyio
async def test_worker_usable_after_deadline_exceeded() -> None:
    """After a deadline the worker resumes and accepts new work (no recovery needed)."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def slow() -> str:
        time.sleep(1.0)
        return "done"

    with pytest.raises(ExcelHungError):
        await worker.submit(slow, timeout=0.1)
    assert worker.get_status().state == WorkbookSessionState.DEADLINE_EXCEEDED

    # Let the slow call return so the worker resumes RUNNING.
    await asyncio.sleep(1.2)
    assert worker.get_status().state == WorkbookSessionState.RUNNING

    # The rebuilt worker is usable again without an explicit recovery.
    result = await worker.submit(lambda: "ok", timeout=2.0)
    assert result == "ok"

    await worker.shutdown()


@pytest.mark.anyio
async def test_liveness_ceiling_clamps_the_per_call_deadline() -> None:
    """The liveness ceiling is the authority: min(timeout, ceiling) bounds the call."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0, liveness_ceiling_s=0.2)
    await worker.start()

    def slow() -> str:
        time.sleep(1.0)
        return "done"

    start = time.monotonic()
    with pytest.raises(ExcelHungError) as excinfo:
        # The soft hint is 5.0s but the 0.2s ceiling must win.
        await worker.submit(slow, timeout=5.0)
    elapsed = time.monotonic() - start

    assert excinfo.value.context.get("deadline_exceeded") is True
    # Bounded by the ceiling (~0.2s), NOT the 5.0s hint.
    assert 0.15 <= elapsed < 0.9, f"deadline not clamped to the ceiling: {elapsed:.3f}s"
    assert worker.get_status().state == WorkbookSessionState.DEADLINE_EXCEEDED

    # Let the slow call return so the worker resumes; then shut down.
    await asyncio.sleep(1.2)
    await worker.shutdown()


def test_process_item_marks_started_and_expires_before_start() -> None:
    """A directly-driven item with a past deadline is 'started' + exec/expired (not busy)."""

    from concurrent.futures import Future

    from excel_vba_mcp.adapters.excel.com_worker import _WorkItem

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    item: _WorkItem = _WorkItem(
        fn=lambda: "x",
        future=Future(),
        deadline=time.monotonic() - 1.0,  # already past
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
        token=1,
    )

    # Drive the worker's own processing path directly (no thread needed for the
    # expired-before-start branch).
    worker._process_item(item, pythoncom)

    # The worker began handling it -> started; the future got a truthful
    # ExcelHungError (exec/expired), never a busy result.
    assert item.started is True
    assert item.future.done()
    exc = item.future.exception()
    assert isinstance(exc, ExcelHungError)


@pytest.mark.anyio
async def test_submit_before_start_raises_unavailable() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=1.0)

    with pytest.raises(WorkerUnavailableError):
        await worker.submit(lambda: 1, timeout=1.0)


@pytest.mark.anyio
async def test_submit_after_shutdown_raises_unavailable() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=1.0)
    await worker.start()
    await worker.shutdown()

    with pytest.raises(WorkerUnavailableError):
        await worker.submit(lambda: 1, timeout=1.0)


# ---------------------------------------------------------------------------
# Shutdown ordering + idempotency
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_shutdown_is_ordered_and_idempotent() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()

    assert adapter.create_count == 1

    await worker.shutdown()
    await worker.shutdown()  # second call is a no-op

    assert adapter.quit_count == 1, "Quit must run exactly once"
    assert adapter.release_count >= 1, "release must run during drain/cleanup"

    status = worker.get_status()
    assert status.state in (WorkbookSessionState.CLOSED, WorkbookSessionState.ERROR)


# ---------------------------------------------------------------------------
# Shutdown reaps the owned pid (FE-17)
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_shutdown_skips_the_fallback_when_the_drain_verified_the_quit() -> None:
    """A drain whose quit ended the process leaves nothing to reap."""

    adapter = VerifyingAdapter(quit_ends_process=True)
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    adapter.worker = worker
    await worker.start()
    assert worker.get_status().owned_pid == 4242

    await worker.shutdown()

    assert adapter.reap_calls == [], "no fallback is needed once the drain verified the quit"
    assert worker.get_status().owned_pid is None


@pytest.mark.anyio
async def test_shutdown_reaps_the_owned_pid_after_an_unverified_drain() -> None:
    """If the STA drain left the process alive, shutdown reaps it on the caller thread."""

    adapter = VerifyingAdapter(quit_ends_process=False)
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    adapter.worker = worker
    await worker.start()
    assert worker.get_status().owned_pid == 4242

    await worker.shutdown()

    assert adapter.reap_calls == [4242], "shutdown must fall back to a pid-scoped reap"
    assert adapter.worker_owned_pid_at_reap == [4242], (
        "_owned_pid must not be cleared before the fallback ran"
    )
    assert worker.get_status().owned_pid is None, "the worker drops the pid after the reap"


# ---------------------------------------------------------------------------
# Proxy-leak guard
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_proxy_leak_guard_rejects_non_serializable_result() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()

    class FakeProxy:
        """Stands in for a leaked COM proxy."""

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await worker.submit(lambda: FakeProxy(), timeout=2.0)

    assert excinfo.value.code == ErrorCode.COM_BUSY
    assert "non-serializable" in str(excinfo.value).lower()

    await worker.shutdown()


@pytest.mark.anyio
async def test_proxy_leak_guard_accepts_domain_model() -> None:
    from excel_vba_mcp.domain.models import WorkbookInfo
    from excel_vba_mcp.domain.statuses import SessionMode

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()

    dto = WorkbookInfo(
        workbook_id="wb1",
        name="Book.xlsm",
        format="xlsm",
        path="C:/tmp/Book.xlsm",
        read_only=False,
        has_vba=True,
        excel_version="16.0",
        session_id="s1",
        mode=SessionMode.INSPECT,
    )

    result = await worker.submit(lambda: dto, timeout=2.0)
    assert isinstance(result, WorkbookInfo)

    await worker.shutdown()


# ---------------------------------------------------------------------------
# READ_ONLY retry vs MUTATING no-retry for RPC_E_CALL_REJECTED
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_read_only_retries_rpc_e_call_rejected() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    calls = {"n": 0}

    def flaky() -> str:
        calls["n"] += 1
        if calls["n"] <= 3:
            raise FakeComError(RPC_E_CALL_REJECTED)
        return "ok"

    result = await worker.submit(flaky, timeout=5.0, op_class=OperationClass.READ_ONLY)
    assert result == "ok"
    assert calls["n"] == 4, "READ_ONLY should retry up to 3 times then succeed"

    await worker.shutdown()


@pytest.mark.anyio
async def test_read_only_exhausts_retries_then_raises_rpc_rejected() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    calls = {"n": 0}

    def always_reject() -> str:
        calls["n"] += 1
        raise FakeComError(RPC_E_CALL_REJECTED)

    with pytest.raises(RpcCallRejectedError) as excinfo:
        await worker.submit(always_reject, timeout=5.0, op_class=OperationClass.READ_ONLY)

    assert excinfo.value.code == ErrorCode.RPC_E_CALL_REJECTED
    # 1 initial + 3 retries.
    assert calls["n"] == 4

    await worker.shutdown()


@pytest.mark.anyio
async def test_mutating_does_not_retry_rpc_e_call_rejected() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    calls = {"n": 0}

    def mutating() -> str:
        calls["n"] += 1
        raise FakeComError(RPC_E_CALL_REJECTED)

    with pytest.raises(RpcCallRejectedError):
        await worker.submit(mutating, timeout=5.0, op_class=OperationClass.MUTATING)

    assert calls["n"] == 1, "MUTATING must not be replayed"

    await worker.shutdown()


@pytest.mark.anyio
async def test_server_call_retrylater_maps_to_com_busy() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def busy() -> str:
        raise FakeComError(RPC_E_SERVERCALL_RETRYLATER)

    with pytest.raises(ComBusyError) as excinfo:
        await worker.submit(busy, timeout=5.0, op_class=OperationClass.READ_ONLY)

    assert excinfo.value.code == ErrorCode.COM_BUSY

    await worker.shutdown()


# ---------------------------------------------------------------------------
# Settings context managers restore on exception (fake app object)
# ---------------------------------------------------------------------------


def test_context_managers_restore_on_exception() -> None:
    app = FakeAppObject()
    adapter = ExcelAppAdapter()
    adapter._app = app

    with pytest.raises(RuntimeError), adapter.alerts_suppressed():
        assert app.DisplayAlerts is False
        raise RuntimeError("boom")
    assert app.DisplayAlerts is True

    with pytest.raises(RuntimeError), adapter.screenupdating_off():
        assert app.ScreenUpdating is False
        raise RuntimeError("boom")
    assert app.ScreenUpdating is True

    with pytest.raises(RuntimeError), adapter.automation_security(1):
        assert app.AutomationSecurity == 1
        raise RuntimeError("boom")
    assert app.AutomationSecurity == 3


# ---------------------------------------------------------------------------
# get_status is a non-COM snapshot
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_get_status_reflects_running_state() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()

    status = worker.get_status()
    assert isinstance(status, WorkerStatus)
    assert status.state == WorkbookSessionState.RUNNING
    assert status.accepting is True
    assert status.owned_pid == 4242

    # Must be JSON-serializable (no COM).
    dumped = status.model_dump(mode="json")
    assert dumped["state"] == "running"

    await worker.shutdown()


# ---------------------------------------------------------------------------
# reset(): rebuild the owned instance (F2b recovery exit)
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_reset_after_hung_fn_released_rebuilds_and_runs() -> None:
    """After the hung fn returns, ``reset()`` rebuilds a fresh owned instance."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=2.0)
    await worker.start()
    assert adapter.create_count == 1
    assert worker.get_status().owned_pid == 4242

    def slow() -> str:
        time.sleep(0.5)
        return "done"

    with pytest.raises(ExcelHungError):
        await worker.submit(slow, timeout=0.1)
    assert worker.get_status().state == WorkbookSessionState.DEADLINE_EXCEEDED

    # Let the hung fn finish so the close-section reset item can be processed.
    await asyncio.sleep(0.7)

    await worker.reset()

    # A fresh hidden instance was created and the worker is RUNNING again.
    assert adapter.create_count == 2
    assert worker.get_status().state == WorkbookSessionState.RUNNING
    assert worker.get_status().accepting is True

    # The rebuilt worker is usable again.
    result = await worker.submit(lambda: "ok", timeout=2.0)
    assert result == "ok"

    await worker.shutdown()


@pytest.mark.anyio
async def test_reset_while_fn_hung_times_out_without_deadlock() -> None:
    """``reset()`` while the fn is still hung is bounded by ``orphan_timeout_s``."""

    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=0.5)
    await worker.start()

    def very_slow() -> str:
        time.sleep(1.5)
        return "done"

    with pytest.raises(ExcelHungError):
        await worker.submit(very_slow, timeout=0.1)
    assert worker.get_status().state == WorkbookSessionState.DEADLINE_EXCEEDED

    # The fn is still running on the STA thread, so the close-section reset item is
    # stuck behind it: reset() gives up after the bounded orphan timeout (no block).
    started = time.monotonic()
    with pytest.raises(ExcelHungError):
        await worker.reset()
    elapsed = time.monotonic() - started
    assert elapsed < 3.0, f"reset blocked {elapsed:.2f}s (must be bounded by orphan_timeout_s)"

    await worker.shutdown()


# ---------------------------------------------------------------------------
# Block D — atomic status snapshot (report_07 §2.4 D1 / plan_02 Work item D)
#
# Deterministic, non-flaky tests on the PRODUCTION ``ExcelComWorker`` snapshot
# methods (no COM, no Excel, no STA thread). They use Events to force the exact
# interleavings that the old lock-free pattern got wrong:
#   * active-exposed read is internally consistent and the result is not yet done;
#   * a discriminating check: the OLD lock-free publish tears in the gap, the
#     atomic snapshot stays consistent (the regression guard D must have);
#   * clear-before-complete under a forced pause (no stale active after clear);
#   * snapshot immutability + a coverage guard (reads>0 AND active_observed>0,
#     so the concurrency check cannot pass vacuously).
# ---------------------------------------------------------------------------


def _token_of(operation: str | None) -> int | None:
    """Recover the op index from an ``op{token}`` operation tag (test helper).

    ``WorkerStatus.active_section`` is a ``Section`` literal, so the op index is
    correlated via ``active_operation`` (``op{token}``) while ``active_section``
    stays a valid ``Section`` value.
    """

    if operation is None or not operation.startswith("op"):
        return None
    try:
        return int(operation[2:])
    except ValueError:
        return None


def _make_item(token: int):
    """Build a real ``_WorkItem`` so ``_complete_result`` exercises production."""

    from concurrent.futures import Future

    from excel_vba_mcp.adapters.excel.com_worker import _WorkItem

    return _WorkItem(
        fn=lambda: token,
        future=Future(),
        deadline=time.monotonic() + 10.0,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
        token=token,
    )


def _new_worker() -> ExcelComWorker:
    return ExcelComWorker(FakeAdapter(), orphan_timeout_s=2.0)


def test_active_exposed_read_is_consistent_and_not_done() -> None:
    """While active is exposed (writer paused), the read is consistent + not done."""

    worker = _new_worker()
    item = _make_item(7)
    ev_active = threading.Event()
    ev_release = threading.Event()

    def writer() -> None:
        worker._set_active(7, "diagnostic", "op7")
        ev_active.set()
        ev_release.wait(2.0)
        worker._clear_active()
        worker._complete_result(item, 7)

    wt = threading.Thread(target=writer, name="d_active_writer", daemon=True)
    wt.start()
    try:
        assert ev_active.wait(2.0), "writer never published the active op"
        status = worker.get_status()
        # Internally consistent (section is a valid Section, operation from the
        # same op) ...
        assert status.active_section == "diagnostic"
        assert status.active_operation == "op7"
        assert _token_of(status.active_operation) == 7
        # ... and the result is NOT delivered while the op is still exposed.
        assert item.future.done() is False
    finally:
        ev_release.set()
        wt.join(2.0)

    after = worker.get_status()
    assert after.active_section is None and after.active_operation is None
    assert item.future.done() is True


def test_lockfree_publish_tears_but_atomic_snapshot_is_consistent() -> None:
    """Discriminating guard: OLD lock-free publish tears in the gap; atomic does not.

    This FAILS on the old lock-free pattern (section written, pause, then
    operation) and PASSES on the atomic snapshot — the regression D lacked.
    """

    # --- OLD pattern: three separate writes with a gap between section/op. ---
    old = {"section": None, "operation": None}
    ev_old_section = threading.Event()
    ev_old_go = threading.Event()

    def old_writer() -> None:
        old["section"] = "sec5"  # section published ...
        ev_old_section.set()
        ev_old_go.wait(2.0)  # ... gap ...
        old["operation"] = "op5"  # ... operation published later.

    ot = threading.Thread(target=old_writer, name="d_old_writer", daemon=True)
    ot.start()
    try:
        assert ev_old_section.wait(2.0)
        # In the gap the section is published but the operation is not yet, so a
        # lock-free reader sees a mismatched (torn) pair.
        old_torn = old["section"] == "sec5" and old["operation"] != "op5"
        assert old_torn, "the old lock-free pattern must expose a torn read in the gap"
    finally:
        ev_old_go.set()
        ot.join(2.0)

    # --- ATOMIC production snapshot: consistent across the same gap. ---
    worker = _new_worker()
    ev_active = threading.Event()
    ev_release = threading.Event()

    def atomic_writer() -> None:
        worker._set_active(5, "diagnostic", "op5")  # single atomic swap
        ev_active.set()
        ev_release.wait(2.0)

    at = threading.Thread(target=atomic_writer, name="d_atomic_writer", daemon=True)
    at.start()
    try:
        assert ev_active.wait(2.0)
        status = worker.get_status()
        atomic_torn = not (
            status.active_section == "diagnostic"
            and status.active_operation == "op5"
            and _token_of(status.active_operation) == 5
        )
        assert not atomic_torn, "atomic snapshot must be consistent in the gap"
    finally:
        ev_release.set()
        at.join(2.0)


def test_clear_precedes_complete_under_forced_pause() -> None:
    """After clear (gate closed) there is NO active op; done only after complete."""

    worker = _new_worker()
    item = _make_item(9)
    ev_cleared = threading.Event()
    ev_complete_gate = threading.Event()

    def writer() -> None:
        worker._set_active(9, "diagnostic", "op9")
        worker._clear_active()
        ev_cleared.set()
        ev_complete_gate.wait(2.0)  # hold completion back
        worker._complete_result(item, 9)

    wt = threading.Thread(target=writer, name="d_cbc_writer", daemon=True)
    wt.start()
    try:
        assert ev_cleared.wait(2.0), "writer never cleared the active op"
        status = worker.get_status()
        # Cleared active is visible ...
        assert status.active_section is None and status.active_operation is None
        # ... while the result is still withheld (gate closed).
        assert item.future.done() is False
    finally:
        ev_complete_gate.set()
        wt.join(2.0)

    assert item.future.done() is True


def test_status_snapshot_is_immutable_and_concurrency_is_consistent() -> None:
    """Returned DTO mutation is isolated; concurrent reads are consistent + observed."""

    worker = _new_worker()

    # Immutability: mutating a returned WorkerStatus must not affect the snapshot.
    first = worker.get_status()
    first.notes.append("mutated-note")
    first.active_section = "hacked"
    first.active_operation = "hacked"
    second = worker.get_status()
    assert "mutated-note" not in second.notes
    assert second.active_section is None and second.active_operation is None

    # Coverage guard: readers + one writer; every read must be consistent and the
    # test must actually observe the active window (cannot pass vacuously). The
    # writer is bounded by a short wall-clock budget so the test stays < 0.5 s.
    stop = threading.Event()
    counts = {"reads": 0, "active_observed": 0}
    violations: list[str] = []
    lock = threading.Lock()
    n_readers = 2
    budget_s = 0.3

    def reader() -> None:
        while not stop.is_set():
            status = worker.get_status()
            operation = status.active_operation
            section = status.active_section
            with lock:
                counts["reads"] += 1
                if operation is not None:
                    counts["active_observed"] += 1
                    token = _token_of(operation)
                    if token is None or section != "diagnostic":
                        violations.append(f"torn: {section!r}/{operation!r}")

    def writer() -> None:
        deadline = time.monotonic() + budget_s
        token = 0
        while time.monotonic() < deadline:
            worker._update_snapshot(
                active_section="diagnostic",
                active_operation=f"op{token}",
                active_token=token,
            )
            time.sleep(0.0005)  # widen the active window so readers observe it
            worker._update_snapshot(
                active_section=None, active_operation=None, active_token=None
            )
            token += 1

    readers = [
        threading.Thread(target=reader, name=f"d_cov_reader_{i}", daemon=True)
        for i in range(n_readers)
    ]
    wt = threading.Thread(target=writer, name="d_cov_writer", daemon=True)
    for rt in readers:
        rt.start()
    wt.start()
    wt.join(5.0)
    stop.set()
    for rt in readers:
        rt.join(5.0)

    assert not violations, f"consistency violations: {violations[:3]}"
    assert counts["reads"] > 0, "readers must have observed the worker"
    assert counts["active_observed"] > 0, "the active window must have been observed"
