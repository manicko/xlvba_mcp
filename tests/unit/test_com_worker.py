"""Unit tests for the single-STA Excel COM worker (task_010).

These tests never open Excel. They inject a fake ``ExcelAppAdapter`` so the
worker thread's ``CoInitializeEx`` / ``DispatchEx`` path is exercised without a
live Excel instance. A fake COM error (``FakeComError``) carrying an
``hresult`` is used to drive the busy/rejected classification, since
``pythoncom.com_error`` may not be importable everywhere.

Covered:
* FIFO ordering and serialization on a single STA thread.
* Awaitable submit does not block the event loop (slow fn + concurrent tick).
* Timeout -> ``NEEDS_RECOVERY`` + truthful ``EXCEL_HUNG``; second submit fails
  fast; a ``close``-section item is allowed through.
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
# Timeout -> NEEDS_RECOVERY + EXCEL_HUNG; fail fast; close allowed
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_timeout_sets_needs_recovery_and_raises_hung() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def slow() -> str:
        time.sleep(1.0)
        return "done"

    with pytest.raises(ExcelHungError) as excinfo:
        await worker.submit(slow, timeout=0.1)

    assert excinfo.value.code == ErrorCode.EXCEL_HUNG
    status = worker.get_status()
    assert status.state == WorkbookSessionState.NEEDS_RECOVERY

    await worker.shutdown()


@pytest.mark.anyio
async def test_second_submit_fails_fast_in_recovery() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def slow() -> str:
        time.sleep(1.0)
        return "done"

    with pytest.raises(ExcelHungError):
        await worker.submit(slow, timeout=0.1)

    # The hung fn is still running; a new non-close submit must fail fast.
    started = threading.Event()

    def should_not_run() -> str:
        started.set()
        return "x"

    with pytest.raises(ExcelHungError):
        await worker.submit(should_not_run, timeout=1.0)

    await asyncio.sleep(0.2)
    assert not started.is_set(), "the fail-fast submit must not start fn"

    await worker.shutdown()


@pytest.mark.anyio
async def test_close_section_allowed_in_recovery() -> None:
    adapter = FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=3.0)
    await worker.start()

    def slow() -> str:
        time.sleep(0.3)
        return "done"

    with pytest.raises(ExcelHungError):
        await worker.submit(slow, timeout=0.1)

    # Let the hung fn return so the close item can run.
    await asyncio.sleep(0.5)

    result = await worker.submit(lambda: "closed", timeout=2.0, section="close")
    assert result == "closed"

    await worker.shutdown()


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
    assert worker.get_status().state == WorkbookSessionState.NEEDS_RECOVERY

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
    assert worker.get_status().state == WorkbookSessionState.NEEDS_RECOVERY

    # The fn is still running on the STA thread, so the close-section reset item is
    # stuck behind it: reset() gives up after the bounded orphan timeout (no block).
    started = time.monotonic()
    with pytest.raises(ExcelHungError):
        await worker.reset()
    elapsed = time.monotonic() - started
    assert elapsed < 3.0, f"reset blocked {elapsed:.2f}s (must be bounded by orphan_timeout_s)"

    await worker.shutdown()
