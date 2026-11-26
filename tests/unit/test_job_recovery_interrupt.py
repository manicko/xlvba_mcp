"""Block F recovery wiring: a successful recovery finalizes jobs definitively.

No Excel / COM is launched. The single-STA worker and the process handle are
faked, so the suite runs on any host. Coverage (integration_report §4):

* a job left ``suspected_stalled`` (deadline-exceeded) is finalized to
  ``INTERRUPTED_UNKNOWN`` after a **successful** recovery — ``outcome='unknown'``,
  ``retryable=False``, prior state appended to ``events`` (never overwritten);
* a late completion from the still-running STA call cannot overwrite the
  interrupted outcome (never ``completed``);
* no replay (no new task / record);
* a failed / aborted recovery does **not** mark jobs interrupted;
* after a restart the JobStore still shows ``interrupted_unknown``.
"""

from __future__ import annotations

import asyncio
import threading
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.api.job_tools import job_get, set_job_registry
from excel_vba_mcp.application.job_registry import JobRegistry, JobState
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.job_store import JobStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class _FakeHandle:
    """A live, non-int handle object (stands in for a ``PyHANDLE``)."""


class _ImmediateWorker:
    """Fake STA worker whose ``submit`` returns a fixed result immediately."""

    def __init__(self, result: Any) -> None:
        self.result = result

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        return self.result


class _BlockingWorker:
    """Fake STA worker whose ``submit`` blocks until ``release`` is set."""

    def __init__(self) -> None:
        self.started = threading.Event()
        self.release = threading.Event()

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
        return "late"


class _FakeWorker:
    """Fake single-STA worker for the recovery path (reset can fail or restart)."""

    def __init__(self, adapter: ExcelAppAdapter, *, restart: bool) -> None:
        self.adapter = adapter
        self.state = WorkbookSessionState.RUNNING
        self.accepting = True
        self.owned_pid = 4242
        self.active_operation: str | None = "macro_run"
        self.restart = restart
        self.reset_count = 0

    def get_status(self) -> Any:
        from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus

        return WorkerStatus(
            state=self.state,
            accepting=self.accepting,
            owned_pid=self.owned_pid,
            notes=[],
            active_operation=self.active_operation,
        )

    async def reset(self) -> bool:
        self.reset_count += 1
        if self.restart:
            self.state = WorkbookSessionState.RUNNING
            self.accepting = True
            self.owned_pid = 9999
            self.adapter._owned_pid = 9999
            self.adapter._owned_handle = _FakeHandle()
        else:
            self.state = WorkbookSessionState.NEEDS_RECOVERY
            self.accepting = False
        return self.restart


def _make_session(tmp_path: Path, *, restart: bool) -> tuple[SessionService, AuditLogger]:
    adapter = ExcelAppAdapter()
    adapter._owned_pid = 4242
    adapter._owned_handle = _FakeHandle()
    adapter.terminate_owned_handle = lambda: True  # type: ignore[method-assign]
    worker = _FakeWorker(adapter, restart=restart)
    audit = AuditLogger(tmp_path / "audit.jsonl")
    session = SessionService(
        worker,  # type: ignore[arg-type]
        adapter,
        PathPolicy({str(tmp_path): str(tmp_path)}),
        ApprovalPolicy(mode=AccessMode.RESTRICTED),
        audit,
        work_dir=tmp_path / "work",
    )
    return session, audit


async def _wait_started(worker: _BlockingWorker) -> None:
    while not worker.started.is_set():
        await asyncio.sleep(0.005)


# ---------------------------------------------------------------------------
# Successful recovery -> definitive INTERRUPTED_UNKNOWN (+ event history)
# ---------------------------------------------------------------------------
async def test_successful_recovery_marks_suspected_stalled_interrupted(tmp_path: Path) -> None:
    store = JobStore(tmp_path / "jobs")
    registry = JobRegistry(_ImmediateWorker({"status": "suspected_stalled"}), store=store)
    handle = registry.start(lambda: None, timeout=5.0)
    await registry.join(handle.job_id, timeout=2.0)
    assert registry.status(handle.job_id).status is JobState.SUCCEEDED

    session, audit = _make_session(tmp_path, restart=True)
    session.set_job_registry(registry)
    try:
        outcome = await session.recover_owned_process()
        assert outcome["ok"] is True

        status = registry.status(handle.job_id)
        assert status.status is JobState.INTERRUPTED_UNKNOWN
        assert status.status is not JobState.SUCCEEDED
        assert status.outcome == "unknown"
        assert status.retryable is False
        # Prior suspected_stalled is retained in the additive event history.
        assert any(event.result_status == "suspected_stalled" for event in status.events)

        set_job_registry(registry)
        got = await job_get(handle.job_id)
        assert got.status is JobState.INTERRUPTED_UNKNOWN
        assert got.outcome == "unknown"
        assert got.retryable is False
        assert any(event.result_status == "suspected_stalled" for event in got.events)

        # No replay: one job, one persisted record.
        assert len(registry._jobs) == 1
        assert len(store.all_records()) == 1
    finally:
        set_job_registry(None)
        audit.close()


async def test_running_job_marked_and_late_completion_does_not_overwrite(
    tmp_path: Path,
) -> None:
    store = JobStore(tmp_path / "jobs")
    worker = _BlockingWorker()
    registry = JobRegistry(worker, store=store)
    handle = registry.start(lambda: None, timeout=5.0)
    await _wait_started(worker)

    marked = registry.mark_interrupted(reason="recovery")
    assert len(marked) == 1
    assert marked[0].status is JobState.INTERRUPTED_UNKNOWN
    assert marked[0].outcome == "unknown"
    assert marked[0].retryable is False
    assert marked[0].events[0].state == "running"

    # The still-running STA call returns late; it must NOT overwrite the outcome.
    worker.release.set()
    final = await registry.join(handle.job_id, timeout=2.0)
    assert final.status is JobState.INTERRUPTED_UNKNOWN
    assert final.result is None
    assert len(registry._jobs) == 1
    assert len(store.all_records()) == 1


# ---------------------------------------------------------------------------
# Failed / aborted recovery does NOT mark interrupted
# ---------------------------------------------------------------------------
async def test_failed_recovery_does_not_mark_interrupted(tmp_path: Path) -> None:
    store = JobStore(tmp_path / "jobs")
    registry = JobRegistry(_ImmediateWorker({"status": "suspected_stalled"}), store=store)
    handle = registry.start(lambda: None, timeout=5.0)
    await registry.join(handle.job_id, timeout=2.0)

    session, audit = _make_session(tmp_path, restart=False)
    session.set_job_registry(registry)
    try:
        outcome = await session.recover_owned_process()
        assert outcome["ok"] is False

        status = registry.status(handle.job_id)
        assert status.status is JobState.SUCCEEDED
        assert status.events == []
        assert len(store.all_records()) == 1
    finally:
        audit.close()


async def test_fail_closed_recovery_does_not_mark_interrupted(tmp_path: Path) -> None:
    store = JobStore(tmp_path / "jobs")
    registry = JobRegistry(_ImmediateWorker({"status": "suspected_stalled"}), store=store)
    handle = registry.start(lambda: None, timeout=5.0)
    await registry.join(handle.job_id, timeout=2.0)

    session, audit = _make_session(tmp_path, restart=True)
    # No live, identity-verified handle => fail closed (no reset, no mark).
    session._app_adapter._owned_handle = None  # type: ignore[attr-defined]
    session.set_job_registry(registry)
    try:
        outcome = await session.recover_owned_process()
        assert outcome["ok"] is False
        assert outcome["recovery_possible"] is False

        status = registry.status(handle.job_id)
        assert status.status is JobState.SUCCEEDED
        assert status.events == []
    finally:
        audit.close()


# ---------------------------------------------------------------------------
# Restart: JobStore still shows interrupted_unknown
# ---------------------------------------------------------------------------
async def test_restart_store_still_interrupted(tmp_path: Path) -> None:
    store = JobStore(tmp_path / "jobs")
    registry = JobRegistry(_ImmediateWorker({"status": "suspected_stalled"}), store=store)
    handle = registry.start(lambda: None, timeout=5.0)
    await registry.join(handle.job_id, timeout=2.0)
    registry.mark_interrupted(reason="recovery")

    # A brand-new store + registry over the same root (a restart).
    store2 = JobStore(tmp_path / "jobs")
    record = store2.get(handle.job_id)
    assert record is not None
    assert record.state == "interrupted_unknown"
    assert record.outcome == "unknown"
    assert record.retryable is False
    assert any(event["result_status"] == "suspected_stalled" for event in record.events)

    registry2 = JobRegistry(_ImmediateWorker(None), store=store2)
    status = registry2.status(handle.job_id)
    assert status.status is JobState.INTERRUPTED_UNKNOWN
    assert status.outcome == "unknown"
    assert status.retryable is False