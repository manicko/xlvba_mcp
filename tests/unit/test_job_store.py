"""Unit tests for the Block F job persistence store (no Excel, no COM).

These tests drive the real :class:`~excel_vba_mcp.storage.job_store.JobStore`
and the real :class:`~excel_vba_mcp.application.job_registry.JobRegistry` against
pure-Python fake workers. They prove the Block F persistence contract:

* a RUNNING marker is written at start and the record is finalized on terminal
  (atomically — valid JSON, no temp files left behind);
* a startup reconcile turns a persisted non-terminal record into
  ``interrupted_unknown`` (never completed, never retried);
* a persisted completed result is returned after a restart (new store + registry);
* idempotency dedupes a repeated key while running and returns the stored result
  once completed;
* polling works with no progress notifications (no dependency);
* an interrupted job is never auto-retried.
"""

from __future__ import annotations

import asyncio
import json
import threading
from collections.abc import Iterator
from datetime import UTC, datetime
from typing import Any

import pytest

from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.job_tools import job_get, set_job_registry, start_long_op
from excel_vba_mcp.application.job_registry import JobRegistry, JobState, RunningJob
from excel_vba_mcp.domain.models import WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
from excel_vba_mcp.storage.job_store import JobRecord, JobStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _reset_registry() -> Iterator[None]:
    set_job_registry(None)
    yield
    set_job_registry(None)


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class BlockingWorker:
    """Fake STA worker whose ``submit`` blocks until ``release`` is set."""

    def __init__(self, result: Any = "ok") -> None:
        self.started = threading.Event()
        self.release = threading.Event()
        self.result = result

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.started.set()
        while not self.release.is_set():
            await asyncio.sleep(0.005)
        return self.result


class ImmediateWorker:
    """Fake STA worker whose ``submit`` returns ``fn()`` immediately."""

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        return fn()


class AsyncProducer:
    """Async long-op producer that blocks until ``release`` (no progress API)."""

    def __init__(self, result: Any = None) -> None:
        self.started = threading.Event()
        self.release = threading.Event()
        self.result = result if result is not None else {"done": True}

    async def run(self) -> Any:
        self.started.set()
        while not self.release.is_set():
            await asyncio.sleep(0.005)
        return self.result


async def _wait_started(worker: BlockingWorker) -> None:
    while not worker.started.is_set():
        await asyncio.sleep(0.005)


async def _wait_producer_started(producer: AsyncProducer) -> None:
    while not producer.started.is_set():
        await asyncio.sleep(0.005)


def _record(job_id: str, state: str, **overrides: Any) -> JobRecord:
    now = datetime.now(UTC)
    payload: dict[str, Any] = {
        "job_id": job_id,
        "state": state,
        "created_at": now,
        "updated_at": now,
        "op_class": OperationClass.READ_ONLY.value,
        "section": "diagnostic",
        "poll_after_s": 5.0,
    }
    payload.update(overrides)
    return JobRecord(**payload)


# ---------------------------------------------------------------------------
# RUNNING marker at start + finalize on terminal (atomic)
# ---------------------------------------------------------------------------
async def test_running_marker_written_at_start_and_finalized_on_terminal(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    worker = BlockingWorker(result={"v": 1})
    registry = JobRegistry(worker, store=store)

    handle = registry.start(lambda: "x", timeout=5.0)
    await _wait_started(worker)

    running = store.get(handle.job_id)
    assert running is not None
    assert running.state == JobState.RUNNING.value

    # Atomic write: the document is valid JSON and no temp file is left behind.
    path = store.path_for(handle.job_id)
    assert path.is_file()
    json.loads(path.read_text(encoding="utf-8"))
    assert list(tmp_path.glob("*.tmp")) == []

    worker.release.set()
    final = await registry.join(handle.job_id, timeout=2.0)
    assert final.status is JobState.SUCCEEDED

    finalized = store.get(handle.job_id)
    assert finalized is not None
    assert finalized.state == JobState.SUCCEEDED.value
    assert finalized.result == {"v": 1}
    assert list(tmp_path.glob("*.tmp")) == []


# ---------------------------------------------------------------------------
# Startup reconcile: RUNNING -> interrupted_unknown (never completed/retried)
# ---------------------------------------------------------------------------
def test_startup_reconcile_running_to_interrupted_unknown(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    store.save(_record("job-old", "running"))

    store2 = JobStore(tmp_path)

    assert len(store2.reconciled) == 1
    reconciled = store2.get("job-old")
    assert reconciled is not None
    assert reconciled.state == "interrupted_unknown"
    assert reconciled.state not in ("succeeded", "completed")
    assert reconciled.outcome == "unknown"
    assert reconciled.retryable is False
    # Never retried: exactly one record remains, still interrupted.
    assert len(store2.all_records()) == 1
    assert store2.get("job-old").state == "interrupted_unknown"  # type: ignore[union-attr]


def test_startup_reconcile_leaves_completed_untouched(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    store.save(_record("job-done", "succeeded", result={"ok": 42}))

    store2 = JobStore(tmp_path)

    assert store2.reconciled == []
    done = store2.get("job-done")
    assert done is not None
    assert done.state == "succeeded"
    assert done.result == {"ok": 42}


# ---------------------------------------------------------------------------
# Restart retrieval: a new registry + store loads a completed result by job_id
# ---------------------------------------------------------------------------
async def test_restart_retrieval_loads_completed_result(tmp_path: Any) -> None:
    store1 = JobStore(tmp_path)
    registry1 = JobRegistry(ImmediateWorker(), store=store1)
    handle = registry1.start(lambda: "kept", timeout=5.0)
    await registry1.join(handle.job_id, timeout=2.0)

    # Simulate a restart: a brand-new store + registry over the same root.
    store2 = JobStore(tmp_path)
    registry2 = JobRegistry(ImmediateWorker(), store=store2)

    status = registry2.status(handle.job_id)
    assert status.status is JobState.SUCCEEDED
    assert status.result == "kept"

    # The job_get tool serves the same persisted result after the restart.
    set_job_registry(registry2)
    try:
        polled = await job_get(handle.job_id)
        assert polled.status is JobState.SUCCEEDED
        assert polled.result == "kept"
    finally:
        set_job_registry(None)


# ---------------------------------------------------------------------------
# Idempotency: dedupe while running; stored result once completed
# ---------------------------------------------------------------------------
async def test_idempotency_dedupes_while_running(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    worker = BlockingWorker(result="r")
    registry = JobRegistry(worker, store=store)

    first = registry.start(lambda: "x", timeout=5.0, idempotency_key="k1")
    second = registry.start(lambda: "x", timeout=5.0, idempotency_key="k1")

    assert first.job_id == second.job_id
    assert len(registry._jobs) == 1
    assert len(store.all_records()) == 1

    worker.release.set()
    await registry.join(first.job_id, timeout=2.0)


async def test_idempotency_completed_returns_stored_result(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    registry = JobRegistry(ImmediateWorker(), store=store)

    first = registry.start(lambda: "value", timeout=5.0, idempotency_key="k2")
    await registry.join(first.job_id, timeout=2.0)

    second = registry.start(lambda: "value", timeout=5.0, idempotency_key="k2")

    assert second.job_id == first.job_id
    assert second.status == JobState.SUCCEEDED.value
    assert second.result == "value"
    assert len(registry._jobs) == 1


# ---------------------------------------------------------------------------
# Polling is independent of progress notifications
# ---------------------------------------------------------------------------
async def test_polling_works_without_progress_notifications(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    worker = BlockingWorker(result="done")
    registry = JobRegistry(worker, store=store)
    set_job_registry(registry)
    try:
        handle = registry.start(lambda: "x", timeout=5.0)

        # No progress notification exists or is consumed; job_get works regardless.
        running = await job_get(handle.job_id)
        assert running.status in (JobState.QUEUED, JobState.RUNNING)
        assert running.poll_after_s == 5.0

        worker.release.set()
        final = await registry.join(handle.job_id, timeout=2.0)
        assert final.status is JobState.SUCCEEDED

        polled = await job_get(handle.job_id)
        assert polled.status is JobState.SUCCEEDED
        assert polled.result == "done"
    finally:
        set_job_registry(None)


# ---------------------------------------------------------------------------
# Producer returns promptly + job_get retrieves the terminal result end-to-end
# ---------------------------------------------------------------------------
async def test_producer_returns_promptly_and_job_get_retrieves_terminal(
    tmp_path: Any,
) -> None:
    store = JobStore(tmp_path)
    registry = JobRegistry(BlockingWorker(), store=store)
    set_job_registry(registry)
    try:
        producer = AsyncProducer()
        handle = start_long_op(
            producer.run,
            timeout=5.0,
            section="test",
            op_class=OperationClass.READ_ONLY,
        )

        # Prompt: the producer returns a running handle without blocking.
        assert handle.status == "running"
        assert handle.job_id
        await _wait_producer_started(producer)

        mid = await job_get(handle.job_id)
        assert mid.status is JobState.RUNNING

        producer.release.set()
        final = await registry.join(handle.job_id, timeout=2.0)
        assert final.status is JobState.SUCCEEDED

        got = await job_get(handle.job_id)
        assert got.status is JobState.SUCCEEDED
        assert got.result == {"done": True}
    finally:
        set_job_registry(None)


# ---------------------------------------------------------------------------
# Producer wiring: a real tool with background=True returns promptly
# ---------------------------------------------------------------------------
class _ImmediateSession:
    """Fake SessionService whose ``open_workbook`` returns immediately (no COM)."""

    async def open_workbook(
        self, path: str, *, mode: SessionMode, read_only: bool
    ) -> WorkbookInfo:
        return WorkbookInfo(
            workbook_id="wb1",
            name="Book.xlsm",
            format="xlsm",
            path=path,
            read_only=read_only,
            has_vba=True,
            excel_version="16.0",
            session_id="s1",
            mode=mode,
        )


async def test_excel_open_background_returns_promptly_and_job_get_retrieves(
    tmp_path: Any,
) -> None:
    store = JobStore(tmp_path)
    registry = JobRegistry(BlockingWorker(), store=store)
    set_job_registry(registry)
    workbook_tools.set_session_service(_ImmediateSession())  # type: ignore[arg-type]
    try:
        handle = await workbook_tools.excel_open("C:/data/Book.xlsm", background=True)

        assert isinstance(handle, RunningJob)
        assert handle.status == "running"
        assert handle.job_id

        final = await registry.join(handle.job_id, timeout=2.0)
        assert final.status is JobState.SUCCEEDED
        assert final.result.workbook_id == "wb1"

        got = await job_get(handle.job_id)
        assert got.status is JobState.SUCCEEDED
        assert got.result.workbook_id == "wb1"
    finally:
        workbook_tools.set_session_service(None)
        set_job_registry(None)


# ---------------------------------------------------------------------------
# No auto-retry of INTERRUPTED_UNKNOWN
# ---------------------------------------------------------------------------
async def test_interrupted_unknown_is_never_retried(tmp_path: Any) -> None:
    store = JobStore(tmp_path)
    store.save(_record("job-interrupted", "running"))

    store2 = JobStore(tmp_path)
    registry = JobRegistry(BlockingWorker(), store=store2)

    status = registry.status("job-interrupted")
    assert status.status is JobState.INTERRUPTED_UNKNOWN
    assert status.outcome == "unknown"
    assert status.retryable is False
    assert status.phase == "interrupted"

    # No task was scheduled and no new record was written (no replay).
    assert registry._tasks == {}
    assert len(store2.all_records()) == 1
    assert store2.get("job-interrupted").state == "interrupted_unknown"  # type: ignore[union-attr]