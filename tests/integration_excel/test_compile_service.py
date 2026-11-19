"""Live-Excel verification of the fail-closed compile service (task_020).

This is the task_020 independent verification of the production
:class:`~excel_vba_mcp.application.compile_service.CompileService` against the
target Excel build (Windows 10 LTSC 2019 17763 x64, Excel 16.0.14332 x64, UI
1033, VBOM enabled). It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter` and a real
:class:`~excel_vba_mcp.application.session_service.SessionService`, always
hidden / non-activating, and only ever touches a temporary **copy** of the
committed ``CleanProject.xlsm`` / ``CompileErrorProject.xlsm`` fixtures. The
committed fixtures are never opened or modified.

Contract verified on both fixtures (research ``hidden-compile-research-2026-11-19``):

* ``compile_project`` returns ``COMPILE_INDETERMINATE`` with reason
  ``NON_ACTIVATING_COMPILE_UNAVAILABLE`` — never ``compiled`` / ``compile_error``;
* the diagnostics are partial and explicitly ``is_compile=false`` and include
  references, ``Protection`` and the component inventory;
* command discovery is read-only and shows **no** Compile route: no
  ``has_compile_entry``, ``route_usable=false``, no UIA expansion, no Execute;
* BG03/BG04: no automation-owned window becomes visible or foreground during or
  after the probe;
* the committed fixtures and the temp originals are byte-identical afterwards;
* no orphaned owned ``EXCEL.exe`` remains after shutdown.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_compile_service.py -v -s
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
    approval: ApprovalPolicy, canonical: str, action: str, *, sha256: str | None = None
) -> None:
    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, sha256=sha256, action=action),
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="task020-verification",
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
        self._original: Path | None = None

    def bind_original(self, original: Path) -> None:
        self._original = original

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    def grant_open(self) -> None:
        assert self._original is not None
        canonical = str(
            self.path_policy.validate_path(str(self._original), OperationClass.MUTATING)
        )
        _grant(self.approval, canonical, "excel_open", sha256=_sha256(self._original))

    async def open(self) -> Any:
        self.grant_open()
        assert self._original is not None
        return await self.session.open_workbook(
            str(self._original), mode=SessionMode.INSPECT, read_only=True
        )


def _make_rig(tmp_path: Path, fixture: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original)
    return rig


# ---------------------------------------------------------------------------
# Test: both fixtures yield a truthful indeterminate result
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "fixture",
    [_CLEAN_PROJECT, _COMPILE_ERROR_PROJECT],
    ids=["clean", "compile_error"],
)
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_compile_is_indeterminate_on_both_fixtures(tmp_path: Path, fixture: Path) -> None:
    fixture = _fixture_or_skip(fixture)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert info.has_vba is True

        result = await rig.service.compile_project(info.workbook_id)

        # Fail-closed status: never a false compiled / compile_error.
        assert result.status is CompileStatus.INDETERMINATE, (
            f"expected COMPILE_INDETERMINATE, got {result.status}"
        )
        assert result.status is not CompileStatus.COMPILED
        assert result.status is not CompileStatus.COMPILE_ERROR
        assert result.reason == REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE, (
            f"unexpected reason on a VBOM-enabled build: {result.reason}"
        )

        # Partial diagnostics are explicitly NOT a compile.
        diagnostics = result.diagnostics
        assert diagnostics.is_compile is False
        assert any("NOT a compile" in note for note in diagnostics.notes)
        assert diagnostics.vbom_available is True
        assert diagnostics.protection == 0
        assert diagnostics.locked is False
        assert diagnostics.references, "no references enumerated"
        assert all(reference.name for reference in diagnostics.references)
        assert diagnostics.components, "no component inventory returned"
        assert diagnostics.component_count == len(diagnostics.components)
        component_names = {component.component_name for component in diagnostics.components}
        assert {"ThisWorkbook", "Sheet1"}.issubset(component_names)

        # Read-only command discovery: no Compile route, no Execute, no UIA expand.
        discovery = result.command_discovery
        assert discovery.attempted is True
        assert discovery.executed is False
        assert discovery.uia_expand_attempted is False
        assert discovery.has_compile_entry is False, (
            f"unexpected Compile entry on this build: {discovery.compile_caption_scan}"
        )
        assert discovery.route_usable is False
        assert discovery.find_control, "no FindControl evidence recorded"
        assert all(probe.error for probe in discovery.find_control), (
            f"expected FindControl failures on this build: {discovery.find_control}"
        )

        # Cached diagnostics must be the same partial diagnostics (no compile).
        diagnostics_cached = await rig.service.get_last_diagnostics(info.workbook_id)
        assert diagnostics_cached == diagnostics
        assert diagnostics_cached.is_compile is False

        # BG03/BG04 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "compile-probe")
        _assert_no_visible_owned_windows(owned_pid, "compile-probe")

        # Original and committed fixtures are untouched.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"

        print(
            f"[compile:{fixture.stem}] status={result.status.value} "
            f"reason={result.reason} references={len(diagnostics.references)} "
            f"components={diagnostics.component_count} "
            f"has_compile_entry={discovery.has_compile_entry} "
            f"route_usable={discovery.route_usable} "
            f"uia_expand_attempted={discovery.uia_expand_attempted} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
