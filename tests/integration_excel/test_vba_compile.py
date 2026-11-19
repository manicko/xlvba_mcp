"""Live-Excel verification of VBA compile outcome classification (task_021).

Independent verification of :class:`CompileService.compile_project` on the
validated build (Windows 10 LTSC 2019 17763 x64, Excel 16.0.14332 x64, UI 1033,
VBOM enabled). This module deliberately states its own expectations instead of
reusing the task_020 assertions: it drives the production single-STA
:class:`ExcelComWorker`, :class:`ExcelAppAdapter`, :class:`SessionService`,
:class:`VBAProjectService` and :class:`CompileService` against a temporary
**copy** of the committed ``CleanProject.xlsm`` / ``CompileErrorProject.xlsm``
fixtures. The committed fixtures are never opened or modified.

Contract under verification (research ``hidden-compile-research-2026-11-19``):

* ``compile_project`` never returns ``COMPILED`` and never fabricates a
  ``COMPILE_ERROR``; it returns ``COMPILE_INDETERMINATE`` with reason
  ``NON_ACTIVATING_COMPILE_UNAVAILABLE`` and diagnostics that are explicitly
  ``is_compile=False`` and still carry the component / reference inventory;
* command discovery records no usable non-activating compile route
  (``has_compile_entry=False``, ``route_usable=False``,
  ``uia_expand_attempted=False``, ``executed=False``);
* after the probe a subsequent ``READ_ONLY`` COM operation completes promptly,
  proving no modal dialog was left blocking the STA worker;
* BG03/BG04: no automation-owned window becomes visible or foreground during or
  after the probe;
* BG05: the original fixtures are byte-identical afterwards, the owned instance
  is a newly created process (never a user instance), and no pre-existing user
  Excel process is terminated;
* the owned ``EXCEL.exe`` is reaped on shutdown (no orphan).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vba_compile.py -v -s
"""

from __future__ import annotations

import contextlib
import csv
import hashlib
import os
import re
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
from excel_vba_mcp.application.compile_service import (
    REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE,
    CompileService,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.statuses import (
    CompileStatus,
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
_COMPILE_ERROR_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CompileErrorProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
_CALL_TIMEOUT_S = 30.0
# A subsequent COM call that cannot finish within this ceiling indicates a hung
# modal / blocked STA worker (BG07/BG09 failure shape).
_PROMPT_LIMIT_S = 20.0

# Win32 dialog class used by Excel/VBE modal prompts.
_DIALOG_CLASS = "#32770"

# Compiled status members that must never be fabricated on this build.
_FORBIDDEN_STATUSES = (CompileStatus.COMPILED, CompileStatus.COMPILE_ERROR)


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


def _excel_pids() -> set[int] | None:
    """Return the current ``EXCEL.exe`` PIDs, or ``None`` when unreadable.

    Read-only ``tasklist`` query used for the BG05 user-session-preservation
    check. Never kills or signals any process.
    """

    if sys.platform != "win32":
        return None
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive
        return None
    pids: set[int] = set()
    for row in csv.reader(out.stdout.splitlines()):
        if len(row) >= 2 and row[0].strip().lower() == "excel.exe":
            with contextlib.suppress(ValueError):
                pids.add(int(row[1].strip()))
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


def _pid_from_hwnd(hwnd: int) -> int | None:
    if win32process is None:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


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


def _modal_dialogs_for_pid(pid: int) -> list[dict[str, Any]]:
    """Return visible ``#32770`` modal windows owned by ``pid`` (read-only)."""

    dialogs: list[dict[str, Any]] = []
    if win32gui is None or win32process is None:
        return dialogs

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != _DIALOG_CLASS:
                return True
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            title = win32gui.GetWindowText(hwnd)
        except Exception:  # noqa: BLE001 - defensive per-window read
            return True
        if int(wpid) == pid:
            dialogs.append({"hwnd": int(hwnd), "pid": int(pid), "title": title})
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_cb, None)
    return dialogs


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


def _assert_no_owned_modal(owned_pid: int, label: str) -> None:
    dialogs = _modal_dialogs_for_pid(owned_pid)
    assert not dialogs, (
        f"BG04 violation: owned automation PID {owned_pid} shows a modal dialog at {label}: "
        f"{[d['title'] for d in dialogs]}"
    )


def _assert_no_foreign_excel_killed(before: set[int] | None, owned_pid: int) -> None:
    """BG05: the owned instance is new and no pre-existing user Excel is gone."""

    assert owned_pid != os.getpid(), (
        "the automation instance must be a separate EXCEL.exe process, not this test process"
    )
    if before is None:
        return
    assert owned_pid not in before, (
        "BG05 violation: the owned automation instance reused an already-running "
        "user Excel process instead of creating a new one"
    )
    after = _excel_pids()
    if after is None:
        return
    disappeared = before - after
    assert not disappeared, (
        f"BG05 violation: pre-existing user Excel processes disappeared: {sorted(disappeared)}"
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _grant(
    approval: ApprovalPolicy, canonical: str, action: str, *, sha256: str | None = None
) -> None:
    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, sha256=sha256, action=action),
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="task021-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# ---------------------------------------------------------------------------
# Bystander Excel (represents the user session; created/closed only by a test)
# ---------------------------------------------------------------------------
@contextlib.contextmanager
def _bystander_excel() -> Iterator[tuple[Any, int]]:
    """Create a *separate* hidden Excel instance to represent the user session.

    Yields ``(app, pid)``. The context manager always ``Quit``-s the bystander;
    it is never activated, shown or used as the automation target. COM is
    initialized on the calling thread for the bystander's lifetime because the
    synchronous PoC tests may already have torn the loop thread's apartment down.
    """

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        pid = _pid_from_hwnd(int(app.Hwnd))
        assert pid is not None, "could not resolve the bystander Excel PID"
        yield app, pid
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
            with contextlib.suppress(Exception):
                del app
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session + VBA project + compile service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {"work": str(self.work_dir), "projects": str(self.projects_dir)}
        )
        self.approval = ApprovalPolicy()
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
        self.vba = VBAProjectService(self.worker, self.session, self.path_policy, self.audit)
        self.service = CompileService(
            self.worker, self.session, self.vba, self.path_policy, self.audit
        )
        self.original: Path | None = None

    def bind_original(self, original: Path) -> None:
        self.original = original

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    def grant_open(self) -> None:
        assert self.original is not None
        canonical = str(self.path_policy.validate_path(str(self.original), OperationClass.MUTATING))
        _grant(self.approval, canonical, "excel_open", sha256=_sha256(self.original))

    async def open(self) -> Any:
        self.grant_open()
        assert self.original is not None
        return await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=True
        )

    async def read_application_version(self) -> str:
        """Read ``Application.Version`` through the worker (READ_ONLY)."""

        adapter = self.app_adapter

        def read_version() -> str:
            app = adapter.app
            if app is None:
                raise RuntimeError("owned Excel application is not available")
            return str(app.Version)

        return await self.worker.submit(
            read_version,
            timeout=_CALL_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )

    async def read_cell_a1(self, working_path: str) -> str:
        """Read ``A1`` of the open working copy through the worker (READ_ONLY)."""

        adapter = self.app_adapter

        def read_a1() -> str:
            app = adapter.app
            if app is None:
                raise RuntimeError("owned Excel application is not available")
            target = working_path.lower()
            for workbook in app.Workbooks:
                if str(workbook.FullName).lower() == target:
                    return str(workbook.Worksheets(1).Range("A1").Value)
            raise RuntimeError(f"working copy is not open: {working_path}")

        return await self.worker.submit(
            read_a1,
            timeout=_CALL_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )


def _make_rig(tmp_path: Path, fixture: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original)
    return rig


# ---------------------------------------------------------------------------
# Test 1: a clean project yields a truthful indeterminate result (never compiled)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_clean_project_compile_status(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    before_pids = _excel_pids()
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)  # type: ignore[arg-type]
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None, "worker did not record an owned PID"

        info = await rig.open()
        assert info.has_vba is True, "CleanProject.xlsm must expose a VBA project"

        result = await rig.service.compile_project(info.workbook_id)

        # The clean project must NOT be reported as compiled.
        assert result.status is CompileStatus.INDETERMINATE, (
            f"expected COMPILE_INDETERMINATE on this build, got {result.status}"
        )
        assert result.status is not CompileStatus.COMPILED
        assert result.reason == REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE, (
            f"unexpected reason on a VBOM-enabled build: {result.reason}"
        )

        diagnostics = result.diagnostics
        assert diagnostics.is_compile is False, "diagnostics must never claim a compile"
        assert any("NOT a compile" in note for note in diagnostics.notes)
        assert diagnostics.component_count == len(diagnostics.components)
        assert diagnostics.components, "clean project must still yield a component inventory"
        assert diagnostics.references, "clean project must still yield references"

        # BG03/BG04 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "clean-compile")
        _assert_no_visible_owned_windows(owned_pid, "clean-compile")

        # Originals untouched.
        assert _sha256(rig.original) == original_hash, "temp original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[clean] status={result.status.value} reason={result.reason} "
            f"components={diagnostics.component_count} references={len(diagnostics.references)} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
    _assert_no_foreign_excel_killed(before_pids, owned_pid)


# ---------------------------------------------------------------------------
# Test 2: the compile-error fixture must not produce a false compiled/compile_error
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_error_fixture_no_false_compiled(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_COMPILE_ERROR_PROJECT)
    before_pids = _excel_pids()
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)  # type: ignore[arg-type]
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None, "worker did not record an owned PID"

        info = await rig.open()
        assert info.has_vba is True, "CompileErrorProject.xlsm must expose a VBA project"

        result = await rig.service.compile_project(info.workbook_id)

        # Never a fabricated success and never a fabricated compile_error: the
        # service cannot observe the VBE compile result on this build.
        for forbidden in _FORBIDDEN_STATUSES:
            assert result.status is not forbidden, (
                f"status {result.status} must not be reported on this build"
            )
        assert result.status is CompileStatus.INDETERMINATE, (
            f"expected COMPILE_INDETERMINATE, got {result.status}"
        )
        assert result.reason == REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE

        # Diagnostics are preserved (the negative fixture still inventories).
        diagnostics = result.diagnostics
        assert diagnostics.is_compile is False
        assert diagnostics.component_count == len(diagnostics.components)
        component_names = {component.component_name for component in diagnostics.components}
        assert {"modBad", "ThisWorkbook", "Sheet1"}.issubset(component_names), (
            f"component inventory not preserved for the error fixture: {sorted(component_names)}"
        )
        assert diagnostics.references, "reference inventory must be preserved"
        assert all(reference.name for reference in diagnostics.references)

        _assert_owned_never_foreground(owned_pid, "error-fixture-compile")
        _assert_no_visible_owned_windows(owned_pid, "error-fixture-compile")

        assert _sha256(rig.original) == original_hash, "temp original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[compile_error] status={result.status.value} reason={result.reason} "
            f"components={sorted(component_names)} references={len(diagnostics.references)} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
    _assert_no_foreign_excel_killed(before_pids, owned_pid)


# ---------------------------------------------------------------------------
# Test 3: no usable non-activating compile route exists (command discovery)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_indeterminate_without_ui_route(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    before_pids = _excel_pids()
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)  # type: ignore[arg-type]
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None, "worker did not record an owned PID"

        info = await rig.open()
        result = await rig.service.compile_project(info.workbook_id)

        assert result.status is CompileStatus.INDETERMINATE
        assert result.diagnostics.is_compile is False

        discovery = result.command_discovery
        assert discovery.attempted is True, "discovery must have been attempted"
        assert discovery.has_compile_entry is False, (
            f"unexpected Compile entry on this build: {discovery.compile_caption_scan}"
        )
        assert discovery.route_usable is False, (
            "no usable non-activating compile route may be reported on this build"
        )
        assert discovery.uia_expand_attempted is False, (
            "UIA menu expansion must never be attempted by the compile probe"
        )
        assert discovery.executed is False, "no CommandBarControl may be executed"
        assert discovery.execute_reason, "a truthful reason must explain the dead route"
        assert discovery.find_control, "no FindControl evidence recorded"
        assert all(probe.error for probe in discovery.find_control), (
            f"expected FindControl failures on this build: {discovery.find_control}"
        )

        _assert_owned_never_foreground(owned_pid, "no-ui-route")
        _assert_no_visible_owned_windows(owned_pid, "no-ui-route")
        assert _sha256(rig.original) == original_hash, "temp original must be unchanged"

        print(
            f"[no-ui-route] status={result.status.value} "
            f"attempted={discovery.attempted} has_compile_entry={discovery.has_compile_entry} "
            f"route_usable={discovery.route_usable} "
            f"uia_expand_attempted={discovery.uia_expand_attempted} "
            f"executed={discovery.executed} find_control_probes={len(discovery.find_control)}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
    _assert_no_foreign_excel_killed(before_pids, owned_pid)


# ---------------------------------------------------------------------------
# Test 4: no modal blocks the next COM operation (and no orphaned owned Excel)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_no_modal_blocks_next_com_operation(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_COMPILE_ERROR_PROJECT)
    before_pids = _excel_pids()
    rig = _make_rig(tmp_path, fixture)
    owned_pid: int | None = None
    with _bystander_excel() as (_bystander_app, bystander_pid):
        try:
            await rig.start()
            committed_hash = _sha256(fixture)
            original_hash = _sha256(rig.original)  # type: ignore[arg-type]
            owned_pid = rig.worker.get_status().owned_pid
            assert owned_pid is not None, "worker did not record an owned PID"
            assert owned_pid != bystander_pid, (
                "the automation instance must be distinct from the user-session bystander"
            )

            info = await rig.open()
            result = await rig.service.compile_project(info.workbook_id)
            assert result.status is CompileStatus.INDETERMINATE

            # No owned modal dialog was left behind by the compile probe.
            _assert_no_owned_modal(owned_pid, "after-compile")

            # A subsequent READ_ONLY COM op must succeed promptly (no hung modal).
            read_started = time.monotonic()
            version = await rig.read_application_version()
            read_elapsed = time.monotonic() - read_started
            assert re.match(r"^\d+\.\d+", version), f"unexpected Version: {version!r}"
            assert read_elapsed < _PROMPT_LIMIT_S, (
                f"READ_ONLY Application.Version took {read_elapsed:.1f}s after compile; "
                "a modal dialog is likely blocking the STA worker"
            )

            # The workbook itself is still reachable through COM.
            a1 = await rig.read_cell_a1(info.path)
            assert isinstance(a1, str), f"unexpected A1 type: {type(a1).__name__}"

            # The worker is healthy (not stuck in NEEDS_RECOVERY).
            worker_state = rig.worker.get_status().state
            assert worker_state is WorkbookSessionState.RUNNING, (
                f"worker left a healthy state after the probe: {worker_state}"
            )

            # BG03/BG04 while Excel is alive.
            _assert_owned_never_foreground(owned_pid, "post-compile-read")
            _assert_no_visible_owned_windows(owned_pid, "post-compile-read")
            _assert_no_owned_modal(owned_pid, "post-compile-read")

            # BG05: the user-session bystander survives untouched.
            assert _pid_alive(bystander_pid), (
                "BG05 violation: the user-session bystander Excel was terminated"
            )

            assert _sha256(rig.original) == original_hash, "temp original must be unchanged"
            assert _sha256(fixture) == committed_hash, "committed fixture must never change"

            print(
                f"[no-modal] version={version} read_elapsed={read_elapsed:.2f}s "
                f"a1={a1!r} worker_state={worker_state.value} "
                f"owned_pid={owned_pid} bystander_pid={bystander_pid}"
            )
        finally:
            await rig.stop()

        # The bystander is still alive after the owned instance shut down.
        assert _pid_alive(bystander_pid), (
            "BG05 violation: the user-session bystander did not survive the owned shutdown"
        )

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown (orphan)"
    _assert_no_foreign_excel_killed(before_pids, owned_pid)
