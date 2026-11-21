"""Live-Excel independent verification of workbook session modes and safe-close.

This module is the task_013 *independent* verification of the production
:class:`~excel_vba_mcp.application.session_service.SessionService` against the
target Excel build. It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` plus a recording
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter`, always
hidden / non-activating, and only ever touches a temporary **copy** of the
committed ``CleanProject.xlsm`` fixture. The committed fixture is never opened or
modified.

The assertions are written independently of the task_012 happy-path checks and
target the security / lifecycle invariants of spec §7.1, §10.1 and acceptance
items A01 / A02 / A15:

* ``test_inspect_and_execute_modes`` — ``inspect`` opens with
  ``AutomationSecurity = ForceDisable`` (macros never run) while ``execute``
  applies ``ByUI`` and requires an approved context (deny-by-default). The owned
  instance stays ``Visible=False`` and never takes foreground / shows a window
  (BG03/BG04) with the interactive session preserved (BG05).
* ``test_forbidden_path_rejected`` — a path outside the allowlisted roots raises
  ``PathNotAllowedError`` (``PATH_NOT_ALLOWED``) before any workbook/instance is
  created for it; the original is untouched.
* ``test_close_save_semantics`` — ``close_workbook(save_changes=False)`` never
  persists an in-memory edit to the working copy, while
  ``close_workbook(save_changes=True)`` does; the original fixture is unchanged
  in both cases.
* ``test_application_settings_restored`` — ``DisplayAlerts`` /
  ``ScreenUpdating`` / ``AutomationSecurity`` are restored to their originals
  after both a successful open and a failing open (BG08), because the adapter's
  context managers restore in ``finally``.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_workbook_session.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, PathNotAllowedError
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.macro_policy import AutomationSecurityMode
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
_CALL_TIMEOUT_S = 30.0

# MsoAutomationSecurity values (mirror the Office enum; kept as plain ints so the
# recording adapter can compare without importing Office).
_MSO_FORCE_DISABLE = int(AutomationSecurityMode.FORCE_DISABLE)  # 3
_MSO_BY_UI = int(AutomationSecurityMode.BY_UI)  # 2


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


def _assert_user_session_preserved(
    owned_pid: int, fg_before: dict[str, Any] | None, label: str
) -> None:
    """BG05: the interactive foreground is not stolen by the owned instance.

    When a non-owned window held the foreground before the operation, it must
    still hold it afterwards; the owned automation instance must never appear as
    the foreground window. Best-effort: absence of evidence is not a violation.
    """

    fg_after = _foreground()
    if fg_after is not None:
        assert fg_after["pid"] != owned_pid, (
            f"BG05 violation: owned automation PID {owned_pid} became foreground at {label} "
            f"(title={fg_after['title']!r})"
        )
    if fg_before is not None and fg_before["pid"] != owned_pid and fg_after is not None:
        assert fg_after["hwnd"] == fg_before["hwnd"], (
            f"BG05 violation: interactive foreground changed across {label} "
            f"(before={fg_before['title']!r} after={fg_after['title']!r})"
        )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Recording adapter: real Excel, but records the security value applied per open
# and the on-disk working-copy hash immediately after each close (before the
# service deletes the working copy). This is the "record the value passed" seam
# used to independently observe macro-trust and save semantics.
# ---------------------------------------------------------------------------
@dataclass
class _PostClose:
    """On-disk state of a working copy captured right after a real Close."""

    path: str
    save_changes: bool
    sha256: str | None


class _RecordingAppAdapter(ExcelAppAdapter):
    """A real :class:`ExcelAppAdapter` that records security/save evidence."""

    def __init__(self) -> None:
        super().__init__()
        self.security_calls: list[int] = []
        self.post_close: list[_PostClose] = []

    @contextlib.contextmanager
    def automation_security(self, mso: int, app: Any = None) -> Any:
        # Record the exact AutomationSecurity value the service applies for open,
        # then delegate to the real restore-in-finally implementation.
        self.security_calls.append(int(mso))
        with super().automation_security(mso, app):
            yield

    def close_workbook(self, wb: Any, *, save_changes: bool = False) -> None:
        # Capture the working-copy path before Close invalidates the proxy, then
        # read its on-disk hash right after Close but before SessionService's
        # unconditional working-copy cleanup. This distinguishes a real save from
        # a no-op close without touching production code.
        path: str | None = None
        if wb is not None:
            try:
                path = str(wb.FullName)
            except Exception:  # noqa: BLE001 - defensive COM read
                path = None
        super().close_workbook(wb, save_changes=save_changes)
        if path is not None:
            digest: str | None
            try:
                candidate = Path(path)
                digest = _sha256(candidate) if candidate.exists() else None
            except OSError:
                digest = None
            self.post_close.append(_PostClose(path=path, save_changes=save_changes, sha256=digest))


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _SessionRig:
    """A live worker + recording adapter + session service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {"work": str(self.work_dir), "projects": str(self.projects_dir)}
        )
        self.approval = ApprovalPolicy()
        self.app_adapter = _RecordingAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.service = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )

    def grant(self, action: str, *, sha256: str | None = None) -> None:
        canonical = str(self.path_policy.validate_path(self._original, OperationClass.MUTATING))
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=canonical, sha256=sha256, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task013-verification",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def bind_original(self, original: Path) -> None:
        self._original = original
        self.original = original
        self.canonical = str(self.path_policy.validate_path(str(original), OperationClass.MUTATING))

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


def _make_rig(tmp_path: Path) -> _SessionRig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _SessionRig(tmp_path)
    original = rig.projects_dir / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    rig.bind_original(original)
    return rig


async def _edit_a1(rig: _SessionRig, path: str, value: str) -> None:
    """Write a cell into the open working copy on the STA worker thread."""

    def edit() -> str:
        app = rig.app_adapter.app
        target = path.lower()
        for wb in app.Workbooks:
            if str(wb.FullName).lower() == target:
                wb.Worksheets(1).Range("A1").Value = value
                return "edited"
        raise RuntimeError("working copy not found in owned instance")

    await rig.worker.submit(edit, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.MUTATING)


async def _read_settings(rig: _SessionRig) -> dict[str, Any]:
    """Read the owned instance's Application settings (STA read-only)."""

    def read() -> dict[str, Any]:
        app = rig.app_adapter.app
        return {
            "DisplayAlerts": bool(app.DisplayAlerts),
            "ScreenUpdating": bool(app.ScreenUpdating),
            "AutomationSecurity": int(app.AutomationSecurity),
        }

    return await rig.worker.submit(read, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.READ_ONLY)


async def _workbook_count(rig: _SessionRig) -> int:
    def count() -> int:
        app = rig.app_adapter.app
        try:
            return int(app.Workbooks.Count)
        except Exception:  # noqa: BLE001 - defensive COM read
            return -1

    return await rig.worker.submit(
        count, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.READ_ONLY
    )


# ---------------------------------------------------------------------------
# Test 1: inspect (macros disabled) vs execute (ByUI + deny-by-default)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_inspect_and_execute_modes(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig.original)
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- Inspect: macros must NOT be enabled (ForceDisable during open). ---
        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=True
        )
        assert info.workbook_id
        assert info.mode == SessionMode.INSPECT
        assert info.read_only is True
        assert info.has_vba is True, "fixture must carry a VBA project"
        assert Path(info.path) != rig.original.resolve(), (
            "must open a working copy, not the original"
        )
        # The AutomationSecurity value applied for this open was ForceDisable (3).
        assert rig.app_adapter.security_calls[-1] == _MSO_FORCE_DISABLE, (
            "inspect open must apply msoAutomationSecurityForceDisable (macros disabled)"
        )

        status = await rig.service.get_status(info.workbook_id)
        assert status.workbook_id == info.workbook_id
        assert status.mode == SessionMode.INSPECT
        assert status.read_only is True

        visible = await rig.worker.submit(
            lambda: bool(rig.app_adapter.app.Visible),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert visible is False, "owned Excel instance must stay Visible=False"
        _assert_owned_never_foreground(owned_pid, "inspect-open")
        _assert_no_visible_owned_windows(owned_pid, "inspect-open")
        _assert_user_session_preserved(owned_pid, fg_before, "inspect-open")

        await rig.service.close_workbook(info.workbook_id)
        assert _sha256(rig.original) == original_hash, "inspect close must not modify the original"

        # --- Execute without a grant is denied (deny-by-default). ---
        with pytest.raises(Exception) as denied:  # noqa: B017 - assert type below
            await rig.service.open_workbook(
                str(rig.original), mode=SessionMode.EXECUTE, read_only=False
            )
        assert denied.type.__name__ == "ApprovalDeniedError", (
            f"execute without a grant must be denied by approval policy, got {denied.type.__name__}"
        )
        assert _sha256(rig.original) == original_hash, (
            "a denied execute must not touch the original"
        )
        assert await rig.service.list_workbooks() == []

        # --- Execute with an approved grant opens (ByUI applied). ---
        rig.grant("excel_open", sha256=original_hash)
        exec_info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        assert exec_info.mode == SessionMode.EXECUTE
        assert exec_info.read_only is False
        # The AutomationSecurity value applied for this open was ByUI (2).
        assert rig.app_adapter.security_calls[-1] == _MSO_BY_UI, (
            "execute open must apply msoAutomationSecurityByUI (trusted copy)"
        )

        visible_after = await rig.worker.submit(
            lambda: bool(rig.app_adapter.app.Visible),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert visible_after is False, (
            "owned Excel instance must stay Visible=False in execute mode"
        )
        _assert_owned_never_foreground(owned_pid, "execute-open")
        _assert_no_visible_owned_windows(owned_pid, "execute-open")
        _assert_user_session_preserved(owned_pid, fg_before, "execute-open")

        await rig.service.close_workbook(exec_info.workbook_id)
        assert _sha256(rig.original) == original_hash, "execute close must not modify the original"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        print(
            f"[modes] inspect_security={_MSO_FORCE_DISABLE} execute_security={_MSO_BY_UI} "
            f"owned_pid={owned_pid} deny_by_default=True original_unchanged=True"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: a path outside the allowlisted roots is rejected before any COM work
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_forbidden_path_rejected(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    outside_dir: Path | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # A file OUTSIDE both configured roots (work / projects).
        outside_dir = tmp_path.parent / f"outside_{uuid4().hex}"
        outside_dir.mkdir(parents=True, exist_ok=True)
        outside_file = outside_dir / "Forbidden.xlsm"
        shutil.copy2(_CLEAN_PROJECT, outside_file)
        outside_hash = _sha256(outside_file)

        assert await _workbook_count(rig) == 0

        with pytest.raises(PathNotAllowedError) as exc_info:
            await rig.service.open_workbook(
                str(outside_file), mode=SessionMode.INSPECT, read_only=True
            )
        assert exc_info.value.code == ErrorCode.PATH_NOT_ALLOWED

        # No workbook / working copy was created for the rejected path.
        assert await rig.service.list_workbooks() == []
        assert await _workbook_count(rig) == 0, "no workbook may be opened for a rejected path"
        leftover = list(rig.work_dir.glob("*")) if rig.work_dir.exists() else []
        assert not leftover, f"a rejected open must not leave a working copy: {leftover}"

        # The out-of-roots original is byte-for-byte untouched.
        assert _sha256(outside_file) == outside_hash, "rejected path's file must be untouched"

        print(f"[forbidden] code={exc_info.value.code.value} workbooks_open=0 file_unchanged=True")
    finally:
        if outside_dir is not None:
            with contextlib.suppress(OSError):
                shutil.rmtree(outside_dir, ignore_errors=True)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: safe-close semantics — no silent save; explicit save persists
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_close_save_semantics(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")  # writable inspect open
    rig.grant("excel_close")  # save-on-close is a mutation
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- close(save_changes=False): an in-memory edit must NOT hit disk. ---
        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=False
        )
        working_copy = Path(info.path)
        work_hash_at_open = _sha256(working_copy)
        assert work_hash_at_open == original_hash, "working copy starts identical to the original"

        await _edit_a1(rig, info.path, "task013-nosave")
        # The edit is only in Excel memory; the on-disk working copy is unchanged.
        assert _sha256(working_copy) == work_hash_at_open, (
            "an in-memory edit must not be persisted to disk before an explicit save"
        )

        await rig.service.close_workbook(info.workbook_id, save_changes=False)

        # The recording adapter captured the on-disk hash right after Close and
        # before SessionService removed the working copy: it must equal the
        # original (no silent save happened).
        last = rig.app_adapter.post_close[-1]
        assert last.save_changes is False
        assert last.sha256 == work_hash_at_open, (
            "close(save_changes=False) must not write the in-memory edit to disk"
        )
        assert _sha256(rig.original) == original_hash, (
            "original must be unchanged after no-save close"
        )
        assert not working_copy.exists(), "working copy is cleaned up after close"

        # --- close(save_changes=True): the explicit save must persist. ---
        info2 = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=False
        )
        working_copy2 = Path(info2.path)
        work_hash2_at_open = _sha256(working_copy2)
        assert work_hash2_at_open == original_hash

        await _edit_a1(rig, info2.path, "task013-save")
        await rig.service.close_workbook(info2.workbook_id, save_changes=True)

        saved = rig.app_adapter.post_close[-1]
        assert saved.save_changes is True
        assert saved.sha256 is not None, "saved working copy must exist on disk at close time"
        assert saved.sha256 != work_hash2_at_open, (
            "close(save_changes=True) must persist the edit to the working copy"
        )
        assert _sha256(rig.original) == original_hash, (
            "even an explicit save must not modify the original fixture"
        )
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        print(
            "[close] nosave_unchanged=True save_changed=True "
            f"original_unchanged=True owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4: Application settings are restored to their originals (BG08)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_application_settings_restored(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    outside_dir: Path | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        baseline = await _read_settings(rig)
        assert set(baseline) == {"DisplayAlerts", "ScreenUpdating", "AutomationSecurity"}

        # --- A successful open then restores settings on the way out. ---
        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=True
        )
        await rig.service.close_workbook(info.workbook_id)
        after_success = await _read_settings(rig)
        assert after_success == baseline, (
            f"settings not restored after a successful open: {after_success} != {baseline}"
        )

        # --- A failing open (forbidden path) also leaves settings untouched. ---
        outside_dir = tmp_path.parent / f"outside_{uuid4().hex}"
        outside_dir.mkdir(parents=True, exist_ok=True)
        outside_file = outside_dir / "Forbidden.xlsm"
        shutil.copy2(_CLEAN_PROJECT, outside_file)
        with pytest.raises(PathNotAllowedError):
            await rig.service.open_workbook(
                str(outside_file), mode=SessionMode.INSPECT, read_only=True
            )
        after_forbidden = await _read_settings(rig)
        assert after_forbidden == baseline, (
            f"settings not restored after a failing open: {after_forbidden} != {baseline}"
        )

        # --- Direct proof the adapter context managers restore in finally. ---
        # Force an exception *inside* the security/alerts context managers on the
        # real owned instance; the finally-restore must put settings back.
        def boom() -> str:
            with (
                rig.app_adapter.automation_security(_MSO_FORCE_DISABLE),
                rig.app_adapter.alerts_suppressed(),
            ):
                raise RuntimeError("forced failure inside security context")

        with pytest.raises(ExcelVbaMcpError):
            await rig.worker.submit(boom, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.MUTATING)
        after_boom = await _read_settings(rig)
        assert after_boom == baseline, (
            f"settings not restored after an in-context failure: {after_boom} != {baseline}"
        )

        # The failing paths never touched the original.
        assert _sha256(rig.original) == original_hash

        print(
            "[settings] restored_after_success/failure=True "
            f"baseline={baseline} owned_pid={owned_pid}"
        )
    finally:
        if outside_dir is not None:
            with contextlib.suppress(OSError):
                shutil.rmtree(outside_dir, ignore_errors=True)
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
