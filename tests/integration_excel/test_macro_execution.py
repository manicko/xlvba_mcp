"""Live-Excel verification of guarded macro execution (task_022).

This is the task_022 independent verification of the production
:class:`~excel_vba_mcp.application.macro_service.MacroService`,
:class:`~excel_vba_mcp.adapters.excel.macro_runner.MacroRunner`,
:class:`~excel_vba_mcp.application.macro_service.BackgroundSafetyRegistry` and the
guarded-approval gate, against the target Excel build (Windows 10 LTSC 2019
17763 x64, Excel 16.0.14332 x64, UI 1033, VBOM enabled). It drives the real
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a
real :class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter` and a real
:class:`~excel_vba_mcp.application.session_service.SessionService`, always hidden /
non-activating, and only ever touches a temporary **copy** of the committed
``CleanProject.xlsm`` fixture. The committed fixture is never opened or modified.

Contract verified (research ``g1-poc-research-2026-11-19.md`` §4 / task_009 §3):

* ``AddNumbers(2,3)`` -> ``SUCCESS`` with return value ``5``; ``NoArgs`` ->
  ``"no-args"``; ``MakeArray`` -> ``[1,2,3]`` (Variant array serialized).
* A missing macro -> ``MACRO_NOT_FOUND``.
* ``ShowMsg`` / ``ShowForm`` are blocked by the interactive-UI preflight
  (``INTERACTIVE_UI_NOT_SUPPORTED``) and are **never dispatched** (no modal, no
  visible window).
* ``RaiseError`` -> ``MACRO_RUNTIME_ERROR``: on this build an unhandled VBA
  runtime error raises a modal and blocks ``Application.Run``; the service
  terminates **only the owned PID** (no dialog dismissal, ``dismissed=false``) and
  never reports success. This is an honest BG04 limitation.
* A hung macro (injected ``HangForever``) -> ``MACRO_TIMEOUT`` / needs-recovery;
  no retry. The owned PID is terminated by the test harness to unblock shutdown.
* BG03/BG04: no automation-owned window becomes visible or foreground during the
  clean run; BG05: user-owned Excel processes are untouched; the owned
  ``EXCEL.exe`` is reaped after shutdown.
* The committed fixture and the temp original are byte-identical afterwards.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_macro_execution.py -v -s
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
from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyRegistry,
    MacroService,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.statuses import (
    MacroRunStatus,
    OperationClass,
    SessionMode,
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

# Optional COM dependencies (None when unavailable so the module imports and the
# COM tests skip on non-Windows hosts).
try:  # pragma: no cover - exercised implicitly by skip logic
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
# RaiseError shows a VBA modal that blocks Application.Run; keep the soft deadline
# short so the watchdog fires quickly.
_MODAL_TIMEOUT_S = 15.0
# A hung (busy-loop) macro trips the soft deadline fast.
_HANG_TIMEOUT_S = 3.0


# ---------------------------------------------------------------------------
# Skip helpers
# ---------------------------------------------------------------------------
def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(
            f"fixture missing: {path}; run `uv run python tests/fixtures/build_fixtures.py` first"
        )
    return path


# ---------------------------------------------------------------------------
# Read-only desktop / process helpers (never activate / show / kill)
# ---------------------------------------------------------------------------
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


def _terminate_owned_pid(pid: int) -> bool:
    """Terminate ONLY the automation-owned Excel PID (owned-only, no wildcard)."""

    try:  # pragma: no cover - depends on host
        completed = subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return completed.returncode == 0
    except Exception:  # noqa: BLE001
        return False


def _excel_pids() -> set[int]:
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
    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[w['title'] for w in visible]}"
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _grant(
    approval: ApprovalPolicy,
    key: ApprovalKey,
) -> None:
    approval.grant(
        ApprovalGrant(
            key=key,
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="task022-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session + macro service + background registry + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {"work": str(self.work_dir), "projects": str(self.projects_dir)}
        )
        self.approval = ApprovalPolicy()
        self.background = BackgroundSafetyRegistry()
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )
        self.service = MacroService(
            self.worker, self.session, self.approval, self.background, self.audit
        )
        self._original: Path | None = None

    def bind_original(self, original: Path) -> None:
        self._original = original

    @property
    def canonical(self) -> str:
        assert self._original is not None
        return str(self.path_policy.validate_path(str(self._original), OperationClass.MUTATING))

    @property
    def workbook_hash(self) -> str:
        assert self._original is not None
        return _sha256(self._original)

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    def grant_open(self) -> None:
        _grant(
            self.approval,
            ApprovalKey(canonical_path=self.canonical, action="excel_open"),
        )

    def grant_macro_run(self, procedure: str) -> None:
        _grant(
            self.approval,
            ApprovalKey(
                canonical_path=self.canonical,
                sha256=self.workbook_hash,
                action="macro_run",
                procedure=procedure,
            ),
        )

    def classify_safe(self, procedure: str) -> None:
        self.background.classify(self.workbook_hash, procedure, background_safe=True)

    async def open(self) -> Any:
        self.grant_open()
        assert self._original is not None
        return await self.session.open_workbook(
            str(self._original), mode=SessionMode.EXECUTE, read_only=False
        )


def _make_rig(tmp_path: Path, fixture: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original)
    return rig


async def _inject_hang(rig: _Rig, working_path: str) -> None:
    """Add a busy-loop ``modTest.HangForever`` to the in-memory working copy.

    Used only to exercise the timeout path: a pure VBA infinite loop never returns
    from ``Application.Run`` and raises no dialog. The on-disk file is unchanged,
    so the workbook hash (and the background_safe classification) still holds.
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


# ---------------------------------------------------------------------------
# Test 1: success / not-found / preflight blocks + BG03/04/05 + no orphan
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_macro_execution_success_and_preflight(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    baseline_pids = _excel_pids()
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]
        fg_before = _foreground()

        info = await rig.open()
        assert info.has_vba is True
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- Success cases (approved + background_safe classified). ---
        rig.grant_macro_run("modTest.AddNumbers")
        rig.classify_safe("modTest.AddNumbers")
        add = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert add.status is MacroRunStatus.SUCCESS, f"AddNumbers failed: {add}"
        assert add.return_value == 5, f"AddNumbers(2,3) != 5: {add}"

        rig.grant_macro_run("modTest.NoArgs")
        rig.classify_safe("modTest.NoArgs")
        no_args = await rig.service.run_macro(info.workbook_id, "modTest.NoArgs")
        assert no_args.status is MacroRunStatus.SUCCESS
        assert no_args.return_value == "no-args"

        rig.grant_macro_run("modTest.MakeArray")
        rig.classify_safe("modTest.MakeArray")
        arr = await rig.service.run_macro(info.workbook_id, "modTest.MakeArray")
        assert arr.status is MacroRunStatus.SUCCESS
        assert arr.return_value == [1, 2, 3]

        # --- Missing macro -> NOT_FOUND (never dispatched). ---
        missing = await rig.service.run_macro(info.workbook_id, "modTest.Nope")
        assert missing.status is MacroRunStatus.NOT_FOUND

        # --- BG06 deny-by-default: approved but NOT classified -> blocked. ---
        rig.grant_macro_run("modTest.Echo")
        unclassified = await rig.service.run_macro(info.workbook_id, "modTest.Echo", ["hi"])
        assert unclassified.status is MacroRunStatus.BLOCKED
        assert unclassified.error_code == "BACKGROUND_SAFETY_NOT_VERIFIED"

        # --- R13 preflight: ShowMsg / ShowForm blocked even if wrongly classified. ---
        for macro in ("modTest.ShowMsg", "modTest.ShowForm"):
            rig.grant_macro_run(macro)
            rig.classify_safe(macro)
            blocked = await rig.service.run_macro(info.workbook_id, macro)
            assert blocked.status is MacroRunStatus.BLOCKED, f"{macro} was not blocked: {blocked}"
            assert blocked.error_code == "INTERACTIVE_UI_NOT_SUPPORTED"

        # get_last_result returns the cache without re-running.
        last = await rig.service.get_last_result(info.workbook_id)
        assert last is not None
        assert last.run_id == blocked.run_id

        # --- BG03/BG04/BG05 while Excel is alive (clean run only). ---
        _assert_owned_never_foreground(owned_pid, "macro-clean")
        _assert_no_visible_owned_windows(owned_pid, "macro-clean")

        # Original + committed fixtures untouched.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"
        fg_after = _foreground()

        print(
            f"[macro-clean] AddNumbers={add.return_value} NoArgs={no_args.return_value!r} "
            f"MakeArray={arr.return_value} missing={missing.status.value} "
            f"unclassified={unclassified.error_code} owned_pid={owned_pid} "
            f"fg_before={None if fg_before is None else fg_before['pid']} "
            f"fg_after={None if fg_after is None else fg_after['pid']}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: RaiseError -> RUNTIME_ERROR via owned-PID terminate (BG04 limitation)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_macro_runtime_error_modal_owned_terminate(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        rig.grant_macro_run("modTest.RaiseError")
        rig.classify_safe("modTest.RaiseError")

        result = await rig.service.run_macro(
            info.workbook_id, "modTest.RaiseError", timeout_s=_MODAL_TIMEOUT_S
        )

        # Honest classification: never success.
        assert result.status is MacroRunStatus.RUNTIME_ERROR, (
            f"RaiseError did not classify as RUNTIME_ERROR: {result}"
        )
        assert result.status is not MacroRunStatus.SUCCESS
        # The modal path terminates ONLY the owned PID and records dismissed=false.
        assert "dismissed" in (result.error_text or "").lower(), (
            f"expected the owned-PID-terminate modal path: {result.error_text!r}"
        )
        # Owned PID was terminated by the service (no orphan).
        assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} survived the modal terminate"

        print(
            f"[macro-runtime] status={result.status.value} error_code={result.error_code} "
            f"owned_pid={owned_pid} error_text={result.error_text!r}"
        )
    finally:
        # Defensive: ensure the owned instance is gone before shutdown.
        if owned_pid is not None and _pid_alive(owned_pid):
            _terminate_owned_pid(owned_pid)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: hung macro -> MACRO_TIMEOUT / needs-recovery (no retry)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_macro_timeout_needs_recovery(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        await _inject_hang(rig, info.path)
        rig.grant_macro_run("modTest.HangForever")
        rig.classify_safe("modTest.HangForever")

        result = await rig.service.run_macro(
            info.workbook_id, "modTest.HangForever", timeout_s=_HANG_TIMEOUT_S
        )

        # A busy-loop macro has no modal -> plain timeout / needs-recovery.
        assert result.status is MacroRunStatus.TIMEOUT, (
            f"HangForever did not classify as TIMEOUT: {result}"
        )
        assert result.error_code == "MACRO_TIMEOUT"
        assert "needs recovery" in (result.error_text or "").lower()
        assert result.status is not MacroRunStatus.SUCCESS

        print(
            f"[macro-timeout] status={result.status.value} error_code={result.error_code} "
            f"owned_pid={owned_pid}"
        )
    finally:
        # The service does NOT terminate on a plain timeout; the worker is stuck in
        # the busy loop. Terminate ONLY the owned PID to unblock app.Run, then stop.
        if owned_pid is not None and _pid_alive(owned_pid):
            _terminate_owned_pid(owned_pid)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
