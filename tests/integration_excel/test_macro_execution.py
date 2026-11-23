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
* A hung macro (injected ``HangForever``) -> ``MACRO_TIMEOUT``; no retry. The
  service terminates **only the owned PID** and recovers the session in-process
  (F2b), so the worker is RUNNING again with a fresh instance and a subsequent
  open + benign macro succeeds without a server restart.
* BG03/BG04: no automation-owned window becomes visible or foreground during the
  clean run; BG05: user-owned Excel processes are untouched; the owned
  ``EXCEL.exe`` is reaped after shutdown.
* The committed fixture and the temp original are byte-identical afterwards.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_macro_execution.py -v -s
"""

from __future__ import annotations

import asyncio
import contextlib
import hashlib
import json
import shutil
import subprocess
import sys
import time
from collections.abc import Iterator
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyClassification,
    BackgroundSafetyRegistry,
    MacroService,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.statuses import (
    AccessMode,
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

    def __init__(self, tmp_dir: Path, *, mode: AccessMode = AccessMode.RESTRICTED) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {"work": str(self.work_dir), "projects": str(self.projects_dir)}
        )
        self.approval = ApprovalPolicy(mode=mode)
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
        self.workbook_id: str | None = None

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

    async def classify_safe(self, procedure: str) -> None:
        """Classify ``procedure`` bound to the live hash + VBA source revision.

        Computes the ``project_source_hash`` live (so it reflects an in-memory
        ``_inject_hang`` edit made before the call) and classifies directly in the
        registry. Direct (not via :meth:`MacroService.classify_background`) on
        purpose: the R13 tests must be able to (wrongly) mark an interactive
        procedure ``background_safe`` to prove the deterministic preflight veto
        still blocks it. The success test exercises ``classify_background`` itself.
        """

        assert self.workbook_id is not None
        record = next(
            r for r in self.session.list_workbooks_sync() if r.workbook_id == self.workbook_id
        )
        _inventory, project_source_hash = await self.service._collect_project_state(
            self.workbook_id, record
        )
        workbook_hash = _sha256(Path(record.path))
        self.background.classify(
            workbook_hash, project_source_hash, procedure, background_safe=True
        )

    async def open(self) -> Any:
        self.grant_open()
        assert self._original is not None
        info = await self.session.open_workbook(
            str(self._original), mode=SessionMode.EXECUTE, read_only=False
        )
        self.workbook_id = info.workbook_id
        return info


def _make_rig(tmp_path: Path, fixture: Path, *, mode: AccessMode = AccessMode.RESTRICTED) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path, mode=mode)
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


async def _mutate_vba_source(rig: _Rig, working_path: str) -> None:
    """Append a harmless comment line to ``modTest`` in the in-memory working copy.

    The on-disk file is unchanged (the workbook hash holds), but the VBA source
    revision changes, so a ``background_safe`` classification bound to the prior
    source hash must no longer match (-> ``BACKGROUND_SAFETY_NOT_VERIFIED``). This
    proves the source-revision binding invalidates on unsaved VBA edits.
    """

    app_adapter = rig.app_adapter

    def mutate() -> bool:
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
        module.AddFromString("' b5 source-revision probe\r\n")
        return True

    mutate.__name__ = "mutate_vba_source"
    await rig.worker.submit(
        mutate, timeout=30.0, section="diagnostic", op_class=OperationClass.EXECUTING
    )


@contextlib.contextmanager
def _count_execute_dispatches(rig: _Rig) -> Iterator[dict[str, int]]:
    """Count ``section="execute"`` worker submissions (single-dispatch evidence).

    Wraps the shared worker's ``submit`` so a macro test can assert a macro was
    dispatched exactly once (no auto-retry) or never dispatched at all (blocked /
    not-found). Other sections (diagnostic / open / close) are ignored. This is a
    read-only test probe: it only counts, it never changes what the worker does.
    """

    original = rig.worker.submit
    counts = {"execute": 0}

    async def wrapped(
        fn: Any,
        *,
        timeout: float,
        section: Any = "diagnostic",
        op_class: Any = OperationClass.READ_ONLY,
    ) -> Any:
        if section == "execute":
            counts["execute"] += 1
        return await original(fn, timeout=timeout, section=section, op_class=op_class)

    rig.worker.submit = wrapped  # type: ignore[method-assign]
    try:
        yield counts
    finally:
        del rig.worker.submit


async def _run_macro_sampling_foreground(
    rig: _Rig,
    workbook_id: str,
    macro_name: str,
    *,
    timeout_s: float,
    poll_interval_s: float = 0.1,
) -> tuple[Any, list[dict[str, Any]]]:
    """Run a macro while sampling the foreground window (BG03 read-only evidence).

    Samples ``GetForegroundWindow`` from the event-loop thread while the macro
    dispatch is in flight and records any sample whose PID equals the automation-
    owned PID. It never activates, focuses, clicks, types or dismisses anything.

    Returns:
        ``(result, owned_foreground_samples)`` where ``owned_foreground_samples``
        holds every foreground sample owned by the automation PID during the run.
    """

    owned_pid = rig.worker.get_status().owned_pid
    task = asyncio.create_task(rig.service.run_macro(workbook_id, macro_name, timeout_s=timeout_s))
    owned_foreground: list[dict[str, Any]] = []
    deadline = time.monotonic() + timeout_s + 30.0
    while not task.done():
        if time.monotonic() > deadline:
            break
        fg = _foreground()
        if fg is not None and owned_pid is not None and fg["pid"] == owned_pid:
            owned_foreground.append(fg)
        await asyncio.sleep(poll_interval_s)
    result = await task
    return result, owned_foreground


async def _read_guard_props(rig: _Rig) -> dict[str, Any] | None:
    """Best-effort: read ``Interactive`` / ``DisplayAlerts`` from the owned instance.

    Reads the F2a guard properties through the adapter's owned ``Application``
    proxy, on the STA worker (COM is never touched from the event-loop thread).
    Returns ``None`` — never raises — when the owned instance is unavailable
    (e.g. terminated by the modal path or the worker is recovering), so a caller
    can treat the read-back as optional evidence.
    """

    app_adapter = rig.app_adapter

    def read() -> dict[str, Any]:
        app = app_adapter.app
        if app is None:
            raise RuntimeError("owned app is gone")
        return {
            "Interactive": bool(app.Interactive),
            "DisplayAlerts": bool(app.DisplayAlerts),
        }

    read.__name__ = "read_guard_props"
    try:
        return await rig.worker.submit(
            read, timeout=10.0, section="diagnostic", op_class=OperationClass.READ_ONLY
        )
    except Exception:  # noqa: BLE001 - best-effort probe; absence is not a failure
        return None


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
        await rig.classify_safe("modTest.AddNumbers")
        add = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert add.status is MacroRunStatus.SUCCESS, f"AddNumbers failed: {add}"
        assert add.return_value == 5, f"AddNumbers(2,3) != 5: {add}"

        rig.grant_macro_run("modTest.NoArgs")
        await rig.classify_safe("modTest.NoArgs")
        no_args = await rig.service.run_macro(info.workbook_id, "modTest.NoArgs")
        assert no_args.status is MacroRunStatus.SUCCESS
        assert no_args.return_value == "no-args"

        rig.grant_macro_run("modTest.MakeArray")
        await rig.classify_safe("modTest.MakeArray")
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
            await rig.classify_safe(macro)
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
    old_pid: int | None = None
    new_pid: int | None = None
    try:
        info = await rig.open()
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        rig.grant_macro_run("modTest.RaiseError")
        await rig.classify_safe("modTest.RaiseError")

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
        assert "dismissed=false" in (result.error_text or "").lower(), (
            f"expected dismissed=false (dialog not dismissed): {result.error_text!r}"
        )
        # F2b: the session recovered in-process and rebuilt a fresh owned instance.
        assert result.session_recovery == "recovered", (
            f"expected in-process recovery: session_recovery={result.session_recovery!r}"
        )
        # The old owned PID was terminated by the service (no orphan) and a NEW
        # owned instance now exists; the worker is RUNNING again.
        assert _wait_pid_gone(old_pid), f"owned PID {old_pid} survived the modal terminate"
        new_pid = rig.worker.get_status().owned_pid
        assert new_pid is not None and new_pid != old_pid, (
            f"recovery did not rebuild a fresh owned instance: old={old_pid} new={new_pid}"
        )
        assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

        print(
            f"[macro-runtime] status={result.status.value} error_code={result.error_code} "
            f"old_pid={old_pid} new_pid={new_pid} session_recovery={result.session_recovery} "
            f"error_text={result.error_text!r}"
        )
    finally:
        # Defensive: ensure every automation-owned instance is gone before shutdown.
        for pid in (old_pid, new_pid):
            if pid is not None and _pid_alive(pid):
                _terminate_owned_pid(pid)
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"owned PID {old_pid} still alive after shutdown"


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
    old_pid: int | None = None
    new_pid: int | None = None
    try:
        info = await rig.open()
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        await _inject_hang(rig, info.path)
        rig.grant_macro_run("modTest.HangForever")
        await rig.classify_safe("modTest.HangForever")

        result = await rig.service.run_macro(
            info.workbook_id, "modTest.HangForever", timeout_s=_HANG_TIMEOUT_S
        )

        # A busy-loop macro has no modal -> plain timeout. F2b: the owned PID is
        # terminated and the session recovers in-process (no server restart).
        assert result.status is MacroRunStatus.TIMEOUT, (
            f"HangForever did not classify as TIMEOUT: {result}"
        )
        assert result.error_code == "MACRO_TIMEOUT"
        assert result.session_recovery == "recovered", (
            f"expected in-process recovery: session_recovery={result.session_recovery!r}"
        )
        assert result.status is not MacroRunStatus.SUCCESS

        # Recovery rebuilt the hidden instance: worker RUNNING with a NEW owned pid.
        new_pid = rig.worker.get_status().owned_pid
        assert new_pid is not None and new_pid != old_pid, (
            f"recovery did not rebuild a fresh owned instance: old={old_pid} new={new_pid}"
        )
        assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

        # The old owned PID was terminated (no orphan).
        assert _wait_pid_gone(old_pid), f"owned PID {old_pid} survived the timeout terminate"

        # A subsequent open + benign macro SUCCEEDS without a server restart. The
        # prior workbook_id is invalid after recovery, so re-open the fixture copy.
        fresh = await rig.open()
        new_pid = rig.worker.get_status().owned_pid
        rig.grant_macro_run("modTest.AddNumbers")
        await rig.classify_safe("modTest.AddNumbers")
        ok = await rig.service.run_macro(fresh.workbook_id, "modTest.AddNumbers", [2, 3])
        assert ok.status is MacroRunStatus.SUCCESS, f"post-recovery session unusable: {ok}"
        assert ok.return_value == 5

        print(
            f"[macro-timeout] status={result.status.value} error_code={result.error_code} "
            f"session_recovery={result.session_recovery} old_pid={old_pid} new_pid={new_pid}"
        )
    finally:
        # Defensive: ensure any owned instance is gone before shutdown.
        for pid in (old_pid, new_pid):
            if pid is not None and _pid_alive(pid):
                _terminate_owned_pid(pid)
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"owned PID {old_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4 (task_023): successful macro -> return value + duration + serialized
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_successful_macro_returns_value_and_duration(tmp_path: Path) -> None:
    """A background-safe macro returns its value, a positive duration and JSON output.

    Runs ``AddNumbers(2,3)`` and ``MakeArray()`` on a writable EXECUTE session and
    asserts: the scalar and Variant-array returns are captured, the result DTO is
    JSON-serializable, and ``duration_ms`` is recorded (> 0). Each macro is
    dispatched exactly once (no hidden retry). BG03/04/05 and the fixture hash hold.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # Capture the user-owned Excel baseline BEFORE the automation starts so an
    # owned PID (which the modal path may terminate mid-test) is never counted as
    # a disappeared user process (BG05).
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        assert info.has_vba is True
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        with _count_execute_dispatches(rig) as counts:
            rig.grant_macro_run("modTest.AddNumbers")
            await rig.classify_safe("modTest.AddNumbers")
            add = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
            assert add.status is MacroRunStatus.SUCCESS, f"AddNumbers(2,3) failed: {add}"
            assert add.return_value == 5, f"AddNumbers(2,3) != 5: {add}"
            assert add.duration_ms > 0, f"duration not recorded: {add}"

            rig.grant_macro_run("modTest.MakeArray")
            await rig.classify_safe("modTest.MakeArray")
            arr = await rig.service.run_macro(info.workbook_id, "modTest.MakeArray")
            assert arr.status is MacroRunStatus.SUCCESS, f"MakeArray failed: {arr}"
            assert arr.return_value == [1, 2, 3], f"MakeArray != [1,2,3]: {arr}"
            assert arr.duration_ms > 0, f"duration not recorded: {arr}"

        # Exactly one Application.Run dispatch per macro (no auto-retry on success).
        assert counts["execute"] == 2, f"expected 2 execute dispatches, got {counts['execute']}"

        # The result DTO is JSON-serializable and carries the return value.
        add_payload = json.loads(add.model_dump_json())
        assert add_payload["return_value"] == 5
        assert add_payload["status"] == MacroRunStatus.SUCCESS.value

        # BG03/BG04 while Excel is alive; BG05 user session preserved.
        _assert_owned_never_foreground(owned_pid, "macro-success")
        _assert_no_visible_owned_windows(owned_pid, "macro-success")
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        # Disposable copies only: the temp original and committed fixture are intact.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[macro-success] AddNumbers={add.return_value} duration_ms={add.duration_ms} "
            f"MakeArray={arr.return_value} duration_ms={arr.duration_ms} "
            f"owned_pid={owned_pid} dispatches={counts['execute']}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 5 (task_023): blocked macros never run (preflight + BG06) + BG03/04/05
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_blocked_macro_never_runs(tmp_path: Path) -> None:
    """Interactive-UI and unclassified macros are blocked BEFORE dispatch.

    ``ShowMsg`` (MsgBox) and ``ShowForm`` (UserForm.Show) are classified safe and
    approved so they reach the interactive-UI preflight (R13); both must be blocked
    with ``INTERACTIVE_UI_NOT_SUPPORTED`` and **never dispatched** (no visible
    window, no side effect). An approved-but-unclassified procedure must be blocked
    with ``BACKGROUND_SAFETY_NOT_VERIFIED`` (BG06). BG03/04/05 hold throughout.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # Capture the user-owned Excel baseline BEFORE the automation starts so an
    # owned PID (which the modal path may terminate mid-test) is never counted as
    # a disappeared user process (BG05).
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        with _count_execute_dispatches(rig) as counts:
            # R13 preflight: known interactive-UI macros are blocked even when
            # approved and (wrongly) classified background_safe.
            for macro in ("modTest.ShowMsg", "modTest.ShowForm"):
                rig.grant_macro_run(macro)
                await rig.classify_safe(macro)
                blocked = await rig.service.run_macro(info.workbook_id, macro)
                assert blocked.status is MacroRunStatus.BLOCKED, (
                    f"{macro} was not blocked: {blocked}"
                )
                assert blocked.error_code == "INTERACTIVE_UI_NOT_SUPPORTED", (
                    f"{macro} error_code={blocked.error_code}"
                )

            # BG06 deny-by-default: approved but NOT classified -> blocked.
            rig.grant_macro_run("modTest.Echo")
            unclassified = await rig.service.run_macro(info.workbook_id, "modTest.Echo", ["hi"])
            assert unclassified.status is MacroRunStatus.BLOCKED
            assert unclassified.error_code == "BACKGROUND_SAFETY_NOT_VERIFIED"

        # None of the blocked procedures were dispatched (no macro body ran).
        assert counts["execute"] == 0, (
            f"a blocked macro was dispatched {counts['execute']} time(s); it must never run"
        )

        # No visible owned window and no focus theft (BG03/BG04) after the blocks.
        _assert_owned_never_foreground(owned_pid, "macro-blocked")
        _assert_no_visible_owned_windows(owned_pid, "macro-blocked")

        # BG05: user-owned Excel processes are untouched.
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        # No side effect on the disposable copies.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[macro-blocked] blocked=2 unclassified=1 dispatches={counts['execute']} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 6 (task_023): missing macro -> MACRO_NOT_FOUND, never dispatched/retried
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_not_found_macro(tmp_path: Path) -> None:
    """An unknown procedure -> ``MACRO_NOT_FOUND`` and is never dispatched.

    ``modTest.Nope`` does not exist in the inventory, so the service resolves it
    before any dispatch: the result is ``MACRO_NOT_FOUND`` with zero ``execute``
    submissions (no retry), and the worker stays usable (not NEEDS_RECOVERY).
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        with _count_execute_dispatches(rig) as counts:
            missing = await rig.service.run_macro(info.workbook_id, "modTest.Nope")

        assert missing.status is MacroRunStatus.NOT_FOUND, f"expected NOT_FOUND: {missing}"
        assert missing.error_code == "MACRO_NOT_FOUND", f"error_code={missing.error_code}"
        assert missing.status is not MacroRunStatus.SUCCESS
        # Never dispatched => no retry was possible.
        assert counts["execute"] == 0, (
            f"a missing macro was dispatched {counts['execute']} time(s); it must not run"
        )

        # The not-found path is a clean resolve miss, not a hung/recovery state.
        assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

        print(
            f"[macro-not-found] status={missing.status.value} "
            f"error_code={missing.error_code} dispatches={counts['execute']} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 7 (task_023): runtime error classified honestly, no false precision,
#                    BG03 preserved, and a fresh session is usable afterwards
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_runtime_error_classified_without_false_precision(tmp_path: Path) -> None:
    """``RaiseError`` -> ``MACRO_RUNTIME_ERROR`` via owned-PID terminate (BG04).

    The unhandled VBA runtime error raises a modal that blocks ``Application.Run``;
    the service classifies ``MACRO_RUNTIME_ERROR`` and terminates **only the owned
    PID** without dismissing the dialog (``dismissed=false``). It is never reported
    as success, the VBA error number/description is **not fabricated** (Excel does
    not forward it), BG03 (no focus theft) holds during the modal, and the session
    recovers in-process (F2b) so a fresh session runs a macro successfully.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # Capture the user-owned baseline BEFORE start: the modal path terminates the
    # owned PID mid-test, so it must not be counted as a vanished user process.
    baseline_pids = _excel_pids()
    await rig.start()
    old_pid: int | None = None
    fresh_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        rig.grant_macro_run("modTest.RaiseError")
        await rig.classify_safe("modTest.RaiseError")

        with _count_execute_dispatches(rig) as counts:
            result, owned_foreground = await _run_macro_sampling_foreground(
                rig, info.workbook_id, "modTest.RaiseError", timeout_s=_MODAL_TIMEOUT_S
            )

        # Honest classification: never success; single dispatch (no retry).
        assert result.status is MacroRunStatus.RUNTIME_ERROR, (
            f"RaiseError did not classify as RUNTIME_ERROR: {result}"
        )
        assert result.status is not MacroRunStatus.SUCCESS
        assert result.error_code == "MACRO_RUNTIME_ERROR"
        assert counts["execute"] == 1, (
            f"expected exactly 1 execute dispatch, got {counts['execute']}"
        )

        # The owned PID was terminated (modal not dismissed; honest BG04 limitation).
        assert "dismissed=false" in (result.error_text or "").lower(), (
            f"expected the owned-PID-terminate modal path: {result.error_text!r}"
        )
        # No false precision: Excel does not forward the VBA error number/description,
        # so the service must not invent them.
        error_text = result.error_text or ""
        assert "boom" not in error_text.lower(), f"fabricated VBA description: {error_text!r}"
        assert "1234" not in error_text, f"fabricated VBA error number: {error_text!r}"

        # BG03: the automation-owned PID never held foreground during the modal.
        assert not owned_foreground, (
            f"BG03 violation: owned PID {old_pid} held foreground during the modal: "
            f"{[w['title'] for w in owned_foreground]}"
        )

        # The service terminated ONLY the owned PID; it is gone (no orphan).
        assert _wait_pid_gone(old_pid), f"owned PID {old_pid} survived the modal terminate"

        # --- Recovery: the modal path already recovered in-process (F2b). ---
        assert result.session_recovery == "recovered", (
            f"expected in-process recovery: session_recovery={result.session_recovery!r}"
        )
        assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

        # The prior workbook_id is invalid after recovery (open books discarded), so
        # re-open the fixture copy and prove a fresh session runs a macro.
        fresh = await rig.open()
        fresh_pid = rig.worker.get_status().owned_pid
        assert fresh_pid is not None and fresh_pid != old_pid

        rig.grant_macro_run("modTest.AddNumbers")
        await rig.classify_safe("modTest.AddNumbers")
        ok = await rig.service.run_macro(fresh.workbook_id, "modTest.AddNumbers", [2, 3])
        assert ok.status is MacroRunStatus.SUCCESS, f"fresh session unusable: {ok}"
        assert ok.return_value == 5, f"fresh-session AddNumbers(2,3) != 5: {ok}"

        # Fresh instance stays hidden / non-intrusive (BG03/BG04).
        _assert_owned_never_foreground(fresh_pid, "runtime-fresh")
        _assert_no_visible_owned_windows(fresh_pid, "runtime-fresh")

        # BG05: no user Excel PID was touched.
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        # Disposable copies intact.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[macro-runtime] status={result.status.value} error_code={result.error_code} "
            f"dispatches={counts['execute']} old_pid={old_pid} fresh_pid={fresh_pid} "
            f"owned_foreground_samples={len(owned_foreground)}"
        )
    finally:
        # Defensive: ensure every automation-owned instance is gone before shutdown.
        for pid in (old_pid, fresh_pid):
            if pid is not None and _pid_alive(pid):
                _terminate_owned_pid(pid)
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"old owned PID {old_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 8 (task_023): hung macro -> MACRO_TIMEOUT / needs-recovery, attempted once
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_timeout_not_retried(tmp_path: Path) -> None:
    """A hung macro -> ``MACRO_TIMEOUT`` + in-process recovery, attempted once.

    An injected ``modTest.HangForever`` busy loop exceeds the short soft deadline.
    The service classifies ``MACRO_TIMEOUT``, records diagnostics, never auto-retries
    (exactly one ``execute`` dispatch), terminates **only the owned PID** and
    recovers the session in-process (F2b): the worker is RUNNING again with a fresh
    owned instance and ``session_recovery == "recovered"``. No server restart.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # Capture the user-owned Excel baseline BEFORE the automation starts so an
    # owned PID (which the modal path may terminate mid-test) is never counted as
    # a disappeared user process (BG05).
    baseline_pids = _excel_pids()
    await rig.start()
    old_pid: int | None = None
    new_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        old_pid = rig.worker.get_status().owned_pid
        assert old_pid is not None

        # Inject a busy-loop macro into the in-memory working copy (file untouched).
        await _inject_hang(rig, info.path)
        rig.grant_macro_run("modTest.HangForever")
        await rig.classify_safe("modTest.HangForever")

        with _count_execute_dispatches(rig) as counts:
            result = await rig.service.run_macro(
                info.workbook_id, "modTest.HangForever", timeout_s=_HANG_TIMEOUT_S
            )

        # Plain timeout (no modal); never success.
        assert result.status is MacroRunStatus.TIMEOUT, (
            f"HangForever did not classify as TIMEOUT: {result}"
        )
        assert result.error_code == "MACRO_TIMEOUT"
        assert result.status is not MacroRunStatus.SUCCESS
        # Exactly one attempt: no auto-retry of a hung macro.
        assert counts["execute"] == 1, (
            f"expected exactly 1 execute dispatch (no retry), got {counts['execute']}"
        )
        # Diagnostics: the run waited ~the soft deadline before giving up.
        assert result.duration_ms >= int(_HANG_TIMEOUT_S * 1000 * 0.8), (
            f"duration {result.duration_ms}ms does not reflect the soft deadline {_HANG_TIMEOUT_S}s"
        )

        # F2b: in-process recovery rebuilt the owned instance -> RUNNING + fresh pid.
        assert result.session_recovery == "recovered", (
            f"expected in-process recovery: session_recovery={result.session_recovery!r}"
        )
        assert rig.worker.get_status().state is WorkbookSessionState.RUNNING
        new_pid = rig.worker.get_status().owned_pid
        assert new_pid is not None and new_pid != old_pid, (
            f"recovery did not rebuild a fresh owned instance: old={old_pid} new={new_pid}"
        )
        # The old owned PID was terminated by the service (no orphan).
        assert _wait_pid_gone(old_pid), f"owned PID {old_pid} survived the timeout terminate"

        # No user Excel PID was touched (BG05); disposable copies intact.
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[macro-timeout] status={result.status.value} error_code={result.error_code} "
            f"session_recovery={result.session_recovery} duration_ms={result.duration_ms} "
            f"dispatches={counts['execute']} old_pid={old_pid} new_pid={new_pid}"
        )
    finally:
        # Defensive: the service already terminated the old owned PID and recovered;
        # ensure any owned instance is gone before shutdown.
        for pid in (old_pid, new_pid):
            if pid is not None and _pid_alive(pid):
                _terminate_owned_pid(pid)
        await rig.stop()

    assert old_pid is not None
    assert _wait_pid_gone(old_pid), f"owned PID {old_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 9 (task_b5): classify a real procedure via the service => SUCCESS, and an
# unsaved VBA source edit invalidates it (BACKGROUND_SAFETY_NOT_VERIFIED)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_service_classify_then_source_edit_invalidates(tmp_path: Path) -> None:
    """``classify_background`` binds to the live source revision.

    Classify ``modTest.AddNumbers`` via :meth:`MacroService.classify_background` (no
    per-run grant in FULL mode) -> ``SUCCESS``. Then mutate the VBA source in memory
    (no save): the workbook hash is unchanged but the ``project_source_hash`` changes,
    so the re-run is ``BACKGROUND_SAFETY_NOT_VERIFIED`` (never a false success).
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture, mode=AccessMode.FULL)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        dto = await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
        assert isinstance(dto, BackgroundSafetyClassification)
        assert dto.background_safe is True
        assert dto.project_source_hash

        ok = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert ok.status is MacroRunStatus.SUCCESS, f"classify+run failed: {ok}"
        assert ok.return_value == 5

        # Unsaved VBA edit -> new source revision -> classification invalidated.
        await _mutate_vba_source(rig, info.path)
        blocked = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert blocked.status is MacroRunStatus.BLOCKED
        assert blocked.error_code == "BACKGROUND_SAFETY_NOT_VERIFIED"

        # Disposable copies intact.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[b5-classify] status={ok.status.value} rv={ok.return_value} "
            f"after_edit={blocked.error_code} owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 10 (task_b5): FULL mode runs a macro with no per-run grant
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_full_mode_macro_runs_without_grant(tmp_path: Path) -> None:
    """In FULL access mode, a policy-approved + background_safe macro runs with no
    ``macro_run`` grant (the empty-registry / per-run-approval dead end is gone)."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture, mode=AccessMode.FULL)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # No grant_macro_run: FULL mode authorizes EXECUTING by policy.
        await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
        result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert result.status is MacroRunStatus.SUCCESS, f"FULL-mode run failed: {result}"
        assert result.return_value == 5

        print(f"[b5-full-mode] status={result.status.value} rv={result.return_value}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 11 (F2a): the Interactive/DisplayAlerts guard brackets Application.Run
# and restores the pre-run values. Best-effort read-back via the adapter.
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_guard_restores_props_around_live_macro(tmp_path: Path) -> None:
    """F2a defense-in-depth: ``Interactive`` / ``DisplayAlerts`` are restored.

    Runs a background-safe macro (``AddNumbers``) and reads the guard properties
    back through the adapter: they must equal their pre-run values (the guard sets
    them ``False`` for the dispatch and restores them). Then runs ``RaiseError``
    and asserts it is still honestly classified ``RUNTIME_ERROR`` (the guard never
    converts a failure into success).

    Best-effort by design: on the target build an unhandled VBA runtime error
    raises a modal that ``DisplayAlerts = False`` does **not** suppress, so the
    service terminates only the owned PID. When the owned instance survives, the
    read-back is asserted; when it was terminated (or is otherwise unavailable),
    the read-back is skipped — never a hard failure.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- Successful macro: the instance survives, so the read-back is real. ---
        pre = await _read_guard_props(rig)
        rig.grant_macro_run("modTest.AddNumbers")
        await rig.classify_safe("modTest.AddNumbers")
        ok = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
        assert ok.status is MacroRunStatus.SUCCESS, f"AddNumbers failed: {ok}"
        assert ok.return_value == 5

        post = await _read_guard_props(rig)
        if pre is not None and post is not None:
            assert post == pre, f"guard did not restore: pre={pre} post={post}"
        else:
            print("[guard-live] success read-back skipped (guard props unavailable)")

        # --- Runtime-error macro: still classified RUNTIME_ERROR (never success). ---
        rig.grant_macro_run("modTest.RaiseError")
        await rig.classify_safe("modTest.RaiseError")
        pre_err = await _read_guard_props(rig)

        result = await rig.service.run_macro(
            info.workbook_id, "modTest.RaiseError", timeout_s=_MODAL_TIMEOUT_S
        )

        assert result.status is MacroRunStatus.RUNTIME_ERROR, (
            f"RaiseError did not classify as RUNTIME_ERROR: {result}"
        )
        assert result.status is not MacroRunStatus.SUCCESS

        # Best-effort read-back: only meaningful if the owned instance survived
        # (the modal path terminates it, so this is typically skipped here).
        post_err = await _read_guard_props(rig)
        if pre_err is not None and post_err is not None and _pid_alive(owned_pid):
            assert post_err == pre_err, (
                f"guard did not restore around runtime error: pre={pre_err} post={post_err}"
            )
        else:
            print("[guard-live] runtime-error read-back skipped (owned instance terminated)")

        # BG05: user-owned Excel processes are untouched.
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        print(
            f"[guard-live] ok_rv={ok.return_value} ok_restored={pre == post} "
            f"runtime_status={result.status.value} owned_pid={owned_pid}"
        )
    finally:
        # The modal path may have terminated the owned PID; reap defensively.
        if owned_pid is not None and _pid_alive(owned_pid):
            _terminate_owned_pid(owned_pid)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
