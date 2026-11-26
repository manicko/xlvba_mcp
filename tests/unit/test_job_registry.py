"""Unit tests for the minimal in-memory job registry (Block E — E/F boundary).

These tests never open Excel and never import COM. They drive the real
:class:`~excel_vba_mcp.application.job_registry.JobRegistry` against a pure-Python
fake worker whose ``submit`` either blocks on an event (a long op) or raises a
typed error. They prove the Block E contract:

* the registry OWNS the result: a job's result is retrievable by ``job_id`` even
  after the request handle is abandoned (no awaiter);
* ``job_get`` transitions ``running -> succeeded`` (and ``expired`` / ``failed``);
* ``job_cancel`` is intent-only (the op still runs to completion);
* an unknown ``job_id`` is a structured ``NOT_FOUND``;
* ``poll_after_s`` is clamped to ``[1, 30]``.
"""

from __future__ import annotations

import asyncio
import threading
from typing import Any

import pytest

from excel_vba_mcp.application.job_registry import JobRegistry, JobState
from excel_vba_mcp.domain.errors import ComBusyError, ErrorCode, ExcelHungError, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import OperationClass

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class BlockingWorker:
    """Fake STA worker whose ``submit`` blocks until ``release`` is set."""

    def __init__(self, result: Any = "ok") -> None:
        self.started = threading.Event()
        self.release = threading.Event()
        self.result = result
        self.submits: list[dict[str, Any]] = []

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append({"section": section, "op_class": op_class, "timeout": timeout})
        self.started.set()
        while not self.release.is_set():
            await asyncio.sleep(0.005)
        return self.result


class RaisingWorker:
    """Fake STA worker whose ``submit`` raises a fixed exception."""

    def __init__(self, exc: BaseException) -> None:
        self.exc = exc

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        raise self.exc


async def _wait_started(worker: BlockingWorker) -> None:
    while not worker.started.is_set():
        await asyncio.sleep(0.005)


# ---------------------------------------------------------------------------
# running -> succeeded + registry owns the result
# ---------------------------------------------------------------------------
async def test_job_get_transitions_running_to_succeeded() -> None:
    worker = BlockingWorker(result={"v": 42})
    registry = JobRegistry(worker)

    handle = registry.start(lambda: "x", timeout=5.0)

    assert handle.status == "running"
    assert handle.job_id
    await _wait_started(worker)
    running = registry.status(handle.job_id)
    assert running.status is JobState.RUNNING
    assert running.started is True

    worker.release.set()
    final = await registry.join(handle.job_id, timeout=2.0)

    assert final.status is JobState.SUCCEEDED
    assert final.result == {"v": 42}
    assert final.error_code is None


async def test_registry_owns_result_after_request_abandoned() -> None:
    """The result survives the request handle being abandoned (no awaiter)."""

    worker = BlockingWorker(result="kept")
    registry = JobRegistry(worker)

    handle = registry.start(lambda: "x", timeout=5.0)
    job_id = handle.job_id
    # While running, the registry owns the asyncio task (not the caller).
    assert job_id in registry._tasks
    # Simulate the request coroutine going away: drop the handle and never await
    # the start task. The registry still owns the running task and its result.
    del handle
    await asyncio.sleep(0)
    # The job is still owned/known after the request is gone.
    assert registry.status(job_id).status in (JobState.QUEUED, JobState.RUNNING)

    worker.release.set()
    status = None
    for _ in range(400):
        status = registry.status(job_id)
        if status.status is JobState.SUCCEEDED:
            break
        await asyncio.sleep(0.005)

    assert status is not None
    assert status.status is JobState.SUCCEEDED
    assert status.result == "kept"


async def test_cancel_unknown_raises_and_terminal_is_noop() -> None:
    worker = BlockingWorker(result="done")
    registry = JobRegistry(worker)

    # cancel of an unknown id -> structured NOT_FOUND.
    with pytest.raises(ExcelVbaMcpError) as excinfo:
        registry.cancel("nope")
    assert excinfo.value.code is ErrorCode.NOT_FOUND

    # cancel on a TERMINAL job is a no-op (does not flip cancel_requested).
    handle = registry.start(lambda: "x", timeout=5.0)
    worker.release.set()
    final = await registry.join(handle.job_id, timeout=2.0)
    assert final.status is JobState.SUCCEEDED

    after = registry.cancel(handle.job_id)
    assert after.cancel_requested is False
    assert after.status is JobState.SUCCEEDED


# ---------------------------------------------------------------------------
# job_cancel is intent-only
# ---------------------------------------------------------------------------
async def test_job_cancel_is_intent_only() -> None:
    worker = BlockingWorker(result="done")
    registry = JobRegistry(worker)

    handle = registry.start(lambda: "x", timeout=5.0)
    await _wait_started(worker)

    cancelled = registry.cancel(handle.job_id)
    assert cancelled.cancel_requested is True
    assert cancelled.status is JobState.RUNNING

    # The uninterruptible op is NOT interrupted: it still runs to completion.
    worker.release.set()
    final = await registry.join(handle.job_id, timeout=2.0)

    assert final.status is JobState.SUCCEEDED
    assert final.result == "done"
    assert final.cancel_requested is True


# ---------------------------------------------------------------------------
# deadline -> expired; busy -> failed; unknown -> NOT_FOUND
# ---------------------------------------------------------------------------
async def test_deadline_marks_expired() -> None:
    registry = JobRegistry(RaisingWorker(ExcelHungError("deadline", code=ErrorCode.EXCEL_HUNG)))

    handle = registry.start(lambda: "x", timeout=5.0)
    final = await registry.join(handle.job_id, timeout=2.0)

    assert final.status is JobState.EXPIRED
    assert final.error_code == ErrorCode.EXCEL_HUNG.value
    assert final.started is True


async def test_queue_wait_busy_marks_failed_with_retry_hint() -> None:
    registry = JobRegistry(
        RaisingWorker(ComBusyError("busy", code=ErrorCode.COM_BUSY, context={"retry_after_s": 1.0}))
    )

    handle = registry.start(lambda: "x", timeout=5.0)
    final = await registry.join(handle.job_id, timeout=2.0)

    assert final.status is JobState.FAILED
    assert final.error_code == ErrorCode.COM_BUSY.value
    assert final.retry_after_s == 1.0


async def test_unknown_job_raises_not_found() -> None:
    registry = JobRegistry(BlockingWorker())

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        registry.status("does-not-exist")

    assert excinfo.value.code is ErrorCode.NOT_FOUND


# ---------------------------------------------------------------------------
# poll_after_s clamp [1, 30] default 5
# ---------------------------------------------------------------------------
async def test_poll_after_is_clamped() -> None:
    worker = BlockingWorker()
    high = JobRegistry(worker, poll_default_s=99.0)
    handle_high = high.start(lambda: "x", timeout=5.0)
    assert handle_high.poll_after_s == 30.0

    low = JobRegistry(worker, poll_default_s=0.0)
    handle_low = low.start(lambda: "x", timeout=5.0)
    assert handle_low.poll_after_s == 1.0

    default = JobRegistry(worker)
    handle_default = default.start(lambda: "x", timeout=5.0)
    assert handle_default.poll_after_s == 5.0

    # Clean up the blocked tasks.
    worker.release.set()
    for registry, handle in ((high, handle_high), (low, handle_low), (default, handle_default)):
        await registry.join(handle.job_id, timeout=2.0)
