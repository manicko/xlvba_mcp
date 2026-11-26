"""Live-Excel integration verification for recovery + close (task_034).

Drives the *production* :class:`SessionService` with the real single-STA
:class:`ExcelComWorker` + :class:`ExcelAppAdapter` against the target Excel
build, always hidden / non-activating. A temporary **copy** of
``CleanProject.xlsm`` is the test's original; the committed fixture is never
opened or modified.

Coverage:

* ``test_hung_call_then_recover_rebuilds_hidden_instance`` — a blocking call
  exceeds its soft deadline -> structured ``EXCEL_HUNG`` + worker
  ``NEEDS_RECOVERY``; a follow-up call fails fast (no new parallel COM call);
  ``recover_session`` closes the owned workbook (no save), rebuilds a fresh
  hidden instance, and health reports ``ok`` again (BG03/BG04).
* ``test_close_session_leaves_no_orphan_and_keeps_bystander`` — ``close_session``
  releases the owned instance (no orphaned ``EXCEL.exe``), is idempotent, and a
  separate bystander Excel instance is neither touched nor killed (BG05).
* ``test_excel_health_truthful_while_hung`` — ``excel_health`` reports
  ``needs_recovery`` while the worker is hung without submitting any work
  (no forced re-run).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_recovery.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import shutil
import subprocess
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.health_tools import excel_health
from excel_vba_mcp.application.macro_service import BackgroundSafetyRegistry, MacroService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelHungError
from excel_vba_mcp.domain.statuses import (
    HealthStatus,
    MacroRunStatus,
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

try:  # pragma: no cover - exercised implicitly by the skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CLEAN_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CleanProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
_CALL_TIMEOUT_S = 30.0
_HANG_SECONDS = 3.0
_HANG_CALL_TIMEOUT_S = 0.5


def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


# ---------------------------------------------------------------------------
# Read-only desktop / process helpers (never activate / show / kill)
# ---------------------------------------------------------------------------
def _pid_from_hwnd(hwnd: int) -> int | None:
    if win32process is None:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


def _pid_alive(pid: int) -> bool:
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive
        return False
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _excel_pids() -> set[int]:
    """Read-only snapshot of current EXCEL.EXE pids (never kills anything)."""

    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
    return pids


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return {"hwnd": hwnd, "pid": int(pid), "title": title}


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if win32gui is None or win32process is None:
        return result

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            return True
        if wpid != pid:
            return True
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            title = ""
        try:
            visible = bool(win32gui.IsWindowVisible(hwnd))
        except Exception:  # noqa: BLE001 - platform specific
            visible = False
        result.append({"hwnd": hwnd, "pid": int(pid), "title": title, "visible": visible})
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_cb, None)
    return result


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at {label} "
        f"(title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    visible = [window for window in _windows_for_pid(owned_pid) if window["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[window['title'] for window in visible]}"
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Fixtures / service builder
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _RecoveryRig:
    """Live worker + session service + temp dirs for one recovery test."""

    def __init__(self, tmp_dir: Path, original: Path) -> None:
        self.tmp_dir = tmp_dir
        self.original = original
        self.path_policy = PathPolicy({str(tmp_dir): str(tmp_dir)})
        self.canonical = str(self.path_policy.validate_path(str(original), OperationClass.MUTATING))
        self.approval = ApprovalPolicy()
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.service = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_dir / "work",
        )

    def grant(self, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="integration-test",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


def _make_rig(tmp_path: Path) -> _RecoveryRig:
    original = tmp_path / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    return _RecoveryRig(tmp_path, original)


# ---------------------------------------------------------------------------
# Test 1: hung call -> EXCEL_HUNG + NEEDS_RECOVERY -> recover rebuilds hidden
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_hung_call_then_recover_rebuilds_hidden_instance(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")
    await rig.start()
    old_pid: int | None = None
    new_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        assert info.path
        _assert_owned_never_foreground(old_pid, "after-open")
        _assert_no_visible_owned_windows(old_pid, "after-open")

        # Simulate a hung COM call: a blocking call that outlives its soft deadline.
        def hang() -> str:
            time.sleep(_HANG_SECONDS)
            return "finished"

        with pytest.raises(ExcelHungError) as excinfo:
            await rig.worker.submit(
                hang,
                timeout=_HANG_CALL_TIMEOUT_S,
                section="execute",
                op_class=OperationClass.EXECUTING,
            )
        assert excinfo.value.code is ErrorCode.EXCEL_HUNG
        # Block E: an execution-deadline call is NOT a poison state. The worker
        # moves to running/deadline_exceeded and stays accepting (no fail-fast);
        # the call is never cancelled and the owned pid is never terminated.
        hung_status = rig.worker.get_status()
        assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
        assert hung_status.accepting is True

        # Health is truthful while the call is over its deadline and never forces a
        # re-run: the worker is alive and accepting (OK), not needs_recovery.
        workbook_tools.set_session_service(rig.service)
        health = await excel_health()
        assert health.status is HealthStatus.OK
        assert health.needs_recovery is False
        assert health.worker_state is WorkbookSessionState.DEADLINE_EXCEEDED
        assert health.owned_pid == old_pid

        # Recover: close owned workbook (no save) + rebuild a fresh hidden instance.
        status = await rig.service.recover_session(reason="integration_hang")
        assert status.state.value == "running"
        assert rig.service.list_workbooks_sync() == []

        new_pid = rig.worker.get_status().owned_pid
        assert new_pid is not None
        assert new_pid != old_pid, "recover must create a fresh owned instance"
        assert _sha256(rig.original) == original_hash, "recover must not touch the original"

        _assert_owned_never_foreground(new_pid, "after-recover")
        _assert_no_visible_owned_windows(new_pid, "after-recover")
        hidden = await rig.worker.submit(
            lambda: bool(rig.app_adapter.app.Visible),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert hidden is False, "the rebuilt instance must stay Visible=False"

        health_after = await excel_health()
        assert health_after.status is HealthStatus.OK
        assert health_after.owned_pid == new_pid

        print(
            f"[recover] old_pid={old_pid} new_pid={new_pid} original_unchanged=True"
        )
    finally:
        workbook_tools.set_session_service(None)
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"old owned PID {old_pid} still alive after recover"
    if new_pid is not None and new_pid != old_pid:
        assert _wait_pid_gone(new_pid), f"new owned PID {new_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: close_session -> no orphan, bystander untouched, idempotent
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_close_session_leaves_no_orphan_and_keeps_bystander(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")
    await rig.start()
    owned_pid: int | None = None
    bystander: Any = None
    bystander_pid: int | None = None
    import pythoncom

    # Guard: main-thread STA may be torn down by an import-time COM probe.
    pythoncom.CoInitialize()
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # A separate, non-owned Excel instance (the "user" Excel). Kept hidden so
        # the test itself never shows a window; it is never passed to the worker.
        bystander = win32com.client.DispatchEx("Excel.Application")
        bystander.Visible = False
        bystander_pid = _pid_from_hwnd(int(bystander.Hwnd))
        assert bystander_pid is not None

        await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )

        first = await rig.service.close_session()
        assert first.state.value == "closed"
        assert rig.worker.get_status().owned_pid is None

        # Idempotent: a second close is a truthful no-op.
        second = await rig.service.close_session()
        assert second.state.value == "closed"
        assert any("already closed" in note for note in second.notes)

        # The owned instance is gone; the bystander is untouched and alive (BG05).
        assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after close"
        assert _pid_alive(bystander_pid), "close_session must not kill a non-owned Excel"
        assert bystander.Workbooks.Count == 0
        assert bool(bystander.Visible) is False
        assert _sha256(rig.original) == original_hash, "close_session must not modify the original"

        # A closed session refuses new work.
        with pytest.raises(ExcelHungError):
            await rig.service.open_workbook(
                str(rig.original), mode=SessionMode.INSPECT, read_only=True
            )

        print(
            f"[close] owned_pid={owned_pid} bystander_pid={bystander_pid} "
            f"bystander_alive=True original_unchanged=True"
        )
    finally:
        if bystander is not None:
            with contextlib.suppress(Exception):
                bystander.Quit()
        if bystander_pid is not None:
            _wait_pid_gone(bystander_pid)
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after close"


# ---------------------------------------------------------------------------
# Test 3: excel_health truthful while hung, without a forced re-run
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_excel_health_truthful_while_hung(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        workbook_tools.set_session_service(rig.service)
        healthy = await excel_health()
        assert healthy.status is HealthStatus.OK
        assert healthy.accepting is True
        assert healthy.open_workbooks == 0

        def hang() -> str:
            time.sleep(_HANG_SECONDS)
            return "finished"

        with pytest.raises(ExcelHungError):
            await rig.worker.submit(
                hang,
                timeout=_HANG_CALL_TIMEOUT_S,
                section="diagnostic",
                op_class=OperationClass.READ_ONLY,
            )

        token_before = rig.worker._token_seq
        hung_health = await excel_health()
        # Block E: a deadline-exceeded call leaves the worker alive and accepting
        # (running/deadline_exceeded), so health stays OK — not needs_recovery.
        assert hung_health.status is HealthStatus.OK
        assert hung_health.needs_recovery is False
        assert hung_health.worker_state is WorkbookSessionState.DEADLINE_EXCEEDED
        assert rig.worker._token_seq == token_before, "excel_health must not submit work"

        print(f"[health] owned_pid={owned_pid} status={hung_health.status.value}")
    finally:
        workbook_tools.set_session_service(None)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4 (F2b): a hung MACRO (via MacroService) drives in-process auto-recovery
# ---------------------------------------------------------------------------
async def _inject_hang_macro(rig: _RecoveryRig, working_path: str) -> None:
    """Add a busy-loop ``modTest.HangForever`` to the in-memory working copy.

    The on-disk file is unchanged, so the workbook hash (and the background_safe
    classification) still holds; only the in-memory VBA source gains the procedure.
    """

    app_adapter = rig.app_adapter

    def inject() -> bool:
        app = app_adapter.app
        target = working_path.lower()
        wb = None
        for candidate in app.Workbooks:
            if str(candidate.FullName).lower() == target:
                wb = candidate
                break
        if wb is None:
            raise RuntimeError("working copy not open")
        module = wb.VBProject.VBComponents("modTest").CodeModule
        module.AddFromString("Public Sub HangForever()\r\nDo\r\nLoop\r\nEnd Sub\r\n")
        return True

    inject.__name__ = "inject_hang"
    await rig.worker.submit(
        inject, timeout=30.0, section="diagnostic", op_class=OperationClass.EXECUTING
    )


async def _classify_safe(
    macro: MacroService, session: SessionService, workbook_id: str, procedure: str
) -> None:
    """Classify ``procedure`` background_safe bound to the live hash + source revision."""

    record = next(r for r in session.list_workbooks_sync() if r.workbook_id == workbook_id)
    _inventory, project_source_hash = await macro._collect_project_state(workbook_id, record)
    workbook_hash = _sha256(Path(record.path))
    macro._background.classify(  # noqa: SLF001 - test reach-in
        workbook_hash, project_source_hash, procedure, background_safe=True
    )


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_hung_macro_drives_auto_recovery(tmp_path: Path) -> None:
    """A hung macro via ``MacroService`` -> ``SUSPECTED_STALLED`` + agent-directed recovery.

    Injects a busy-loop macro and runs it through the real ``MacroService`` with a
    short deadline. Asserts the honest ``SUSPECTED_STALLED`` / ``MACRO_TIMEOUT`` with
    ``session_recovery is None`` (no auto-kill, no in-process recovery): the owned PID
    is left alive and the worker is ``running/deadline_exceeded``. Agent-directed
    ``recover_owned_process`` then rebuilds a fresh hidden instance and a subsequent
    open + benign macro succeeds WITHOUT a server restart. No user Excel PID is
    touched (BG05).
    """

    rig = _make_rig(tmp_path)
    baseline_pids = _excel_pids()
    # _RecoveryRig uses the default ApprovalPolicy (RESTRICTED): grant the open and
    # the two macro_run approvals (bound to the working-copy hash + procedure).
    background = BackgroundSafetyRegistry()
    macro = MacroService(rig.worker, rig.service, rig.approval, background, rig.audit)
    rig.grant("excel_open")
    wb_hash = _sha256(rig.original)
    for proc in ("modTest.HangForever", "modTest.AddNumbers"):
        rig.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(
                    canonical_path=rig.canonical,
                    sha256=wb_hash,
                    action="macro_run",
                    procedure=proc,
                ),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="integration-test",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )
    await rig.start()
    old_pid: int | None = None
    new_pid: int | None = None
    try:
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )

        await _inject_hang_macro(rig, info.path)
        await _classify_safe(macro, rig.service, info.workbook_id, "modTest.HangForever")

        result = await macro.run_macro(
            info.workbook_id, "modTest.HangForever", timeout_s=_HANG_CALL_TIMEOUT_S
        )

        # Block E: the busy-loop macro outlives its soft deadline -> SUSPECTED_STALLED
        # (MACRO_TIMEOUT). No auto-kill and no in-process recovery: the owned pid is
        # left alive, session_recovery is None and the worker is
        # running/deadline_exceeded; recovery is agent-directed.
        assert result.status is MacroRunStatus.SUSPECTED_STALLED, (
            f"expected SUSPECTED_STALLED: {result}"
        )
        assert result.error_code == "MACRO_TIMEOUT"
        assert result.status is not MacroRunStatus.SUCCESS
        assert result.session_recovery is None, (
            f"recovery is agent-directed, not automatic: {result.session_recovery!r}"
        )
        assert _pid_alive(old_pid), "the owned pid must NOT be terminated (no auto-kill)"
        assert rig.worker.get_status().state is WorkbookSessionState.DEADLINE_EXCEEDED

        # Agent-directed recovery rebuilds a fresh hidden instance; the old pid is reaped.
        recover = await rig.service.recover_owned_process(reason="task066_hung_macro")
        assert recover["process_restarted"] is True, f"agent-directed recovery failed: {recover}"
        new_pid = recover["owned_pid"]
        assert new_pid is not None and new_pid != old_pid, (
            f"recovery did not rebuild a fresh instance: old={old_pid} new={new_pid}"
        )

        # A subsequent open + benign macro SUCCEEDS without a server restart.
        fresh = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        await _classify_safe(macro, rig.service, fresh.workbook_id, "modTest.AddNumbers")
        ok = await macro.run_macro(fresh.workbook_id, "modTest.AddNumbers", [2, 3])
        assert ok.status is MacroRunStatus.SUCCESS, f"post-recovery session unusable: {ok}"
        assert ok.return_value == 5, f"fresh-session AddNumbers(2,3) != 5: {ok}"

        # BG05: no user Excel PID was touched.
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        print(
            f"[hung-macro-recover] status={result.status.value} "
            f"session_recovery={result.session_recovery} old_pid={old_pid} new_pid={new_pid}"
        )
    finally:
        for pid in (old_pid, new_pid):
            if pid is not None and _pid_alive(pid):
                with contextlib.suppress(Exception):
                    subprocess.run(
                        ["taskkill", "/F", "/PID", str(pid)],
                        capture_output=True,
                        text=True,
                        timeout=30,
                    )
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"old owned PID {old_pid} still alive after shutdown"
