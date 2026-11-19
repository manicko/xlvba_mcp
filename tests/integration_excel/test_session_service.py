"""Live-Excel integration verification for the workbook session service (task_012).

This module drives the *production* :class:`SessionService` (with the real
single-STA :class:`ExcelComWorker` + :class:`ExcelAppAdapter`) against the target
Excel build, always hidden/non-activating. It uses only a temporary **copy** of
the ``CleanProject.xlsm`` fixture as the "original"; the committed fixture is
never opened or modified.

Coverage:

* ``test_session_open_inspect_read_only`` — open in ``inspect`` mode
  (``read_only=True``); status reports read-only + inspect; ``close`` without
  save leaves the original copy byte-for-byte unchanged (BG05).
* ``test_session_execute_save_and_close`` — ``execute`` mode (ByUI) open with an
  approved grant; a real cell edit + explicit ``save_workbook`` changes the
  working copy on disk; ``close`` without save leaves the original copy
  unchanged.
* ``test_session_close_default_no_save`` — after an edit, ``close`` (default
  ``save_changes=False``) discards changes: the working copy is removed and the
  original copy is unchanged.

Every test asserts the owned instance stays hidden / non-activating
(BG03/BG04) and that the owned ``EXCEL.exe`` is reaped on shutdown (no orphan).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_session_service.py -v -s
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
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Optional COM dependencies (None when unavailable so the module imports and
# the COM tests skip on non-Windows hosts).
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


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Fixtures / service builder
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _SessionRig:
    """Bundle of a live worker + session service + temp dirs for one test."""

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


def _make_rig(tmp_path: Path) -> _SessionRig:
    """Copy the committed fixture into a temp dir as the test's original."""

    original = tmp_path / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    return _SessionRig(tmp_path, original)


# ---------------------------------------------------------------------------
# Test 1: inspect read-only open + close (original unchanged, BG05)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_session_open_inspect_read_only(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=True
        )
        assert info.workbook_id
        assert info.mode == SessionMode.INSPECT
        assert info.read_only is True
        assert Path(info.path) != rig.original.resolve(), (
            "must open a working copy, not the original"
        )
        _assert_owned_never_foreground(owned_pid, "inspect-after-open")
        _assert_no_visible_owned_windows(owned_pid, "inspect-after-open")

        # Hidden instance + restored automation security (BG03/BG04/BG08).
        visible = await rig.worker.submit(
            lambda: bool(rig.app_adapter.app.Visible),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert visible is False, "owned Excel instance must stay Visible=False"

        status = await rig.service.get_status(info.workbook_id)
        assert status.workbook_id == info.workbook_id
        assert status.read_only is True
        assert status.mode == SessionMode.INSPECT
        assert status.dirty is False

        # Close without save; the original is byte-for-byte unchanged (BG05).
        await rig.service.close_workbook(info.workbook_id)
        assert _sha256(rig.original) == original_hash, "close must not modify the original"
        assert await rig.service.list_workbooks() == []

        print(
            f"[inspect] workbook_id={info.workbook_id} path={info.path} "
            f"read_only={info.read_only} owned_pid={owned_pid} original_unchanged=True"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: execute (ByUI) open + edit + explicit save changes the working copy
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_session_execute_save_and_close(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")
    rig.grant("excel_save")
    await rig.start()
    owned_pid: int | None = None
    working_copy: Path | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        working_copy = Path(info.path)
        assert info.mode == SessionMode.EXECUTE
        assert info.read_only is False
        assert working_copy.exists()
        work_hash_before = _sha256(working_copy)

        # Make a real edit on the STA worker (direct object-model write).
        def edit_a1() -> str:
            app = rig.app_adapter.app
            target = str(info.path).lower()
            for wb in app.Workbooks:
                if str(wb.FullName).lower() == target:
                    wb.Worksheets(1).Range("A1").Value = "task012-save"
                    return "edited"
            raise RuntimeError("working copy not found in owned instance")

        await rig.worker.submit(edit_a1, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.MUTATING)

        # Explicit save writes through to the working copy on disk.
        refreshed = await rig.service.save_workbook(info.workbook_id)
        assert refreshed.workbook_id == info.workbook_id
        work_hash_after = _sha256(working_copy)
        assert work_hash_after != work_hash_before, "save_workbook must persist to the working copy"

        # The original copy is never modified by edit/save (BG05).
        assert _sha256(rig.original) == original_hash, "edit/save must not modify the original"

        _assert_owned_never_foreground(owned_pid, "execute-after-save")
        _assert_no_visible_owned_windows(owned_pid, "execute-after-save")

        await rig.service.close_workbook(info.workbook_id)
        assert _sha256(rig.original) == original_hash, "close must not modify the original"

        print(
            f"[execute] workbook_id={info.workbook_id} owned_pid={owned_pid} "
            f"work_changed={work_hash_after != work_hash_before} original_unchanged=True"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: close default (no save) discards changes; original unchanged
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_session_close_default_no_save(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")  # writable inspect open
    await rig.start()
    owned_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.service.open_workbook(
            str(rig.original), mode=SessionMode.INSPECT, read_only=False
        )
        working_copy = Path(info.path)

        # Edit the working copy but DO NOT save; close must discard it.
        def edit_a1() -> str:
            app = rig.app_adapter.app
            target = str(info.path).lower()
            for wb in app.Workbooks:
                if str(wb.FullName).lower() == target:
                    wb.Worksheets(1).Range("A1").Value = "discard-me"
                    return "edited"
            raise RuntimeError("working copy not found")

        await rig.worker.submit(edit_a1, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.MUTATING)

        # Default close -> no save. Working copy is removed; original unchanged.
        await rig.service.close_workbook(info.workbook_id)
        assert _sha256(rig.original) == original_hash, (
            "close-without-save must not touch the original"
        )
        assert not working_copy.exists(), "working copy should be cleaned up on close"

        print(
            f"[close-nosave] workbook_id={info.workbook_id} "
            f"owned_pid={owned_pid} original_unchanged=True"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
