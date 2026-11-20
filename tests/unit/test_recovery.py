"""Unit tests for session recovery / close (task_034; spec §5.6, §9.4, §11).

A fake single-STA worker (no COM) exercises the recovery contract:

* ``recover_session`` closes the owned workbooks (``Close(SaveChanges=False)``),
  resets the worker to rebuild the hidden instance, drops the session
  bookkeeping (including per-path locks) and reports ``running``;
* a failed reset returns ``needs_recovery`` with diagnostics and keeps refusing
  new work — never a false success;
* a COM/UIA timeout surfaces as a structured ``EXCEL_HUNG`` and moves the worker
  into ``needs_recovery``;
* ``close_session`` is idempotent and closes the owned workbooks without saving;
* ``excel_health`` reports the truth without submitting any work to the worker
  (no forced re-run).
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.health_tools import excel_health
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelHungError
from excel_vba_mcp.domain.statuses import (
    HealthStatus,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel / worker (no COM)
# ---------------------------------------------------------------------------
class FakeWorkbook:
    """Minimal ``Workbook`` proxy recording how it was closed."""

    def __init__(self, app: FakeApp, path: str) -> None:
        self._app = app
        self.FullName = path
        self.Name = Path(path).name
        self.ReadOnly = False
        self.Saved = True
        self.Application = app
        self.closed = False
        self.saved_on_close: bool | None = None

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM name
        self.saved_on_close = SaveChanges
        self.closed = True
        if self in self._app.Workbooks:
            self._app.Workbooks.remove(self)


class FakeWorkbooks:
    """Minimal ``Workbooks`` collection."""

    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self._items: list[FakeWorkbook] = []

    def Open(self, path: str, **_kwargs: Any) -> FakeWorkbook:  # noqa: N802, N803 - COM name
        workbook = FakeWorkbook(self._app, path)
        self._items.append(workbook)
        return workbook

    def __iter__(self) -> Iterator[FakeWorkbook]:
        return iter(list(self._items))

    def __len__(self) -> int:
        return len(self._items)

    def __contains__(self, item: object) -> bool:
        return item in self._items

    def remove(self, item: FakeWorkbook) -> None:
        self._items.remove(item)


class FakeApp:
    """Minimal hidden ``Application`` proxy."""

    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(self)
        self.quit_called = False

    def Quit(self) -> None:  # noqa: N802 - COM name
        self.quit_called = True


class FakeWorker:
    """Fake single-STA worker: records calls, never runs a real COM thread."""

    def __init__(self, adapter: ExcelAppAdapter) -> None:
        self.adapter = adapter
        self.state = WorkbookSessionState.RUNNING
        self.accepting = True
        self.owned_pid = 4242
        self.notes: list[str] = []
        self.active_operation: str | None = None
        self.last_error: str | None = None
        self.submits: list[str] = []
        self.reset_count = 0
        self.shutdown_count = 0
        self.reset_fails = False
        self.fail_sections: set[str] = set()
        self._next_pid = 5000

    # -- worker API --------------------------------------------------------
    def get_status(self) -> WorkerStatus:
        return WorkerStatus(
            state=self.state,
            accepting=self.accepting,
            owned_pid=self.owned_pid,
            notes=list(self.notes),
            active_operation=self.active_operation,
            last_error=self.last_error,
        )

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float = 60.0,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        del timeout, op_class  # unused by the fake
        self.submits.append(section)
        if section in self.fail_sections:
            self.state = WorkbookSessionState.NEEDS_RECOVERY
            self.accepting = False
            self.notes.append(f"simulated {section} hang")
            raise ExcelHungError(f"simulated {section} hang")
        if section == "close":
            self._drain()
            return True
        return fn()

    async def reset(self) -> bool:
        self.reset_count += 1
        self._drain()
        if self.reset_fails:
            self.state = WorkbookSessionState.NEEDS_RECOVERY
            self.accepting = False
            self.notes.append("simulated reset timeout")
            raise ExcelHungError("simulated reset timeout")
        self.adapter._app = FakeApp()
        self._next_pid += 1
        self.adapter._owned_pid = self._next_pid
        self.owned_pid = self._next_pid
        self.state = WorkbookSessionState.RUNNING
        self.accepting = True
        self.notes = []
        return True

    async def shutdown(self) -> None:
        self.shutdown_count += 1
        self._drain()
        self.state = WorkbookSessionState.CLOSED
        self.accepting = False
        self.owned_pid = None

    def _drain(self) -> None:
        """Close owned workbooks (no save) + Quit + release, like the worker drain."""

        app = self.adapter.app
        if app is not None:
            for workbook in list(app.Workbooks):
                self.adapter.close_workbook(workbook, save_changes=False)
            self.adapter.quit()
        self.adapter.release()


# ---------------------------------------------------------------------------
# Rig
# ---------------------------------------------------------------------------
class Rig:
    """A session service backed by the fake worker and a real temp workbook."""

    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"PK\x03\x04fake-workbook")
        self.canonical = str(self.original.resolve())
        self.path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
        self.approval = ApprovalPolicy()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.adapter = ExcelAppAdapter()
        self.app = FakeApp()
        self.adapter._app = self.app
        self.adapter._owned_pid = 4242
        self.worker = FakeWorker(self.adapter)
        self.session = SessionService(
            self.worker,
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action="excel_open"),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id="grant-open",
            )
        )


@pytest.fixture
def rig(tmp_path: Path) -> Iterator[Rig]:
    instance = Rig(tmp_path)
    yield instance
    instance.audit.close()


async def _open(rig: Rig) -> Any:
    return await rig.session.open_workbook(str(rig.original), mode=SessionMode.EXECUTE, timeout_s=5)


# ---------------------------------------------------------------------------
# recover_session
# ---------------------------------------------------------------------------
async def test_recover_closes_owned_workbooks_and_resets(rig: Rig) -> None:
    info = await _open(rig)
    workbook = rig.app.Workbooks._items[0]
    assert len(rig.session.list_workbooks_sync()) == 1

    status = await rig.session.recover_session()

    # Truthful success: the hidden instance was rebuilt and is accepting work.
    assert status.state is WorkbookSessionState.RUNNING
    assert rig.session.accepting is True
    assert rig.worker.reset_count == 1
    # The owned workbook was closed without saving and the bookkeeping was reset.
    assert workbook.saved_on_close is False
    assert workbook.closed is True
    assert rig.session.list_workbooks_sync() == []
    # A fresh hidden instance exists (still hidden), with a new owned pid.
    assert rig.adapter.app is not rig.app
    assert rig.adapter.app.Visible is False
    assert rig.worker.owned_pid != 4242
    assert any("rebuilt" in note for note in status.notes)
    # Per-path locks were released, so the same path can be opened again.
    again = await _open(rig)
    assert again.workbook_id != info.workbook_id


async def test_recover_failure_returns_needs_recovery_without_false_success(
    rig: Rig,
) -> None:
    await _open(rig)
    rig.worker.reset_fails = True

    status = await rig.session.recover_session()

    assert status.state is WorkbookSessionState.NEEDS_RECOVERY
    assert rig.session.accepting is False
    assert any("reset" in note for note in status.notes)
    # New work is refused before it reaches the worker (no new COM call).
    submits_before = len(rig.worker.submits)
    with pytest.raises(ExcelHungError):
        await _open(rig)
    assert len(rig.worker.submits) == submits_before


async def test_recover_is_safe_when_worker_already_needs_recovery(rig: Rig) -> None:
    rig.worker.state = WorkbookSessionState.NEEDS_RECOVERY
    rig.worker.accepting = False
    rig.worker.last_error = "EXCEL_HUNG: worker reset required"

    status = await rig.session.recover_session(reason="after_hang")

    assert status.state is WorkbookSessionState.RUNNING
    assert rig.worker.reset_count == 1
    assert any("EXCEL_HUNG" in note for note in status.notes)


# ---------------------------------------------------------------------------
# COM/UIA timeout -> structured error + needs_recovery
# ---------------------------------------------------------------------------
async def test_hung_call_maps_to_excel_hung_and_needs_recovery(rig: Rig) -> None:
    rig.worker.fail_sections = {"open"}

    with pytest.raises(ExcelHungError) as excinfo:
        await _open(rig)

    assert excinfo.value.code is ErrorCode.EXCEL_HUNG
    assert rig.worker.state is WorkbookSessionState.NEEDS_RECOVERY
    assert (await rig.session.get_status()).state is WorkbookSessionState.NEEDS_RECOVERY


# ---------------------------------------------------------------------------
# close_session
# ---------------------------------------------------------------------------
async def test_close_session_is_idempotent(rig: Rig) -> None:
    await _open(rig)
    workbook = rig.app.Workbooks._items[0]

    first = await rig.session.close_session()

    assert first.state is WorkbookSessionState.CLOSED
    assert workbook.saved_on_close is False
    assert workbook.closed is True
    assert rig.worker.shutdown_count == 1
    assert rig.adapter.owned_pid is None

    second = await rig.session.close_session()

    assert second.state is WorkbookSessionState.CLOSED
    assert any("already closed" in note for note in second.notes)
    # Idempotent: the worker was shut down exactly once.
    assert rig.worker.shutdown_count == 1
    # A closed session refuses new work.
    with pytest.raises(ExcelHungError):
        await _open(rig)


async def test_close_session_without_open_workbook_is_safe(rig: Rig) -> None:
    status = await rig.session.close_session()

    assert status.state is WorkbookSessionState.CLOSED
    assert rig.worker.shutdown_count == 1


async def test_recover_after_close_reports_truthful_closed_state(rig: Rig) -> None:
    await rig.session.close_session()

    status = await rig.session.recover_session()

    # A closed worker cannot be reset: report the truthful closed state instead of
    # waiting for a join that can never complete, and keep refusing new work.
    assert status.state is WorkbookSessionState.CLOSED
    assert rig.session.accepting is False
    assert rig.worker.reset_count == 0
    assert any("closed" in note for note in status.notes)


# ---------------------------------------------------------------------------
# excel_health (no forced re-run)
# ---------------------------------------------------------------------------
async def test_excel_health_is_truthful_without_forced_rerun(rig: Rig) -> None:
    workbook_tools.set_session_service(rig.session)
    try:
        before = len(rig.worker.submits)

        health = await excel_health()

        assert health.status is HealthStatus.OK
        assert health.accepting is True
        assert health.open_workbooks == 0
        assert health.owned_pid == 4242
        # The health check never submitted work to the worker.
        assert len(rig.worker.submits) == before

        rig.worker.state = WorkbookSessionState.NEEDS_RECOVERY
        rig.worker.accepting = False
        rig.worker.notes = ["EXCEL_HUNG: worker reset required"]

        degraded = await excel_health()

        assert degraded.status is HealthStatus.NEEDS_RECOVERY
        assert degraded.needs_recovery is True
        assert degraded.accepting is False
        assert degraded.notes == ["EXCEL_HUNG: worker reset required"]
        # Still no forced re-run while the worker is hung.
        assert len(rig.worker.submits) == before
    finally:
        workbook_tools.set_session_service(None)


async def test_excel_health_without_service_is_unavailable() -> None:
    workbook_tools.set_session_service(None)

    health = await excel_health()

    assert health.status is HealthStatus.UNAVAILABLE
    assert health.accepting is False
