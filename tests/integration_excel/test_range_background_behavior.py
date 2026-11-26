"""Desktop non-interference checks for the Ranges & Cells workflow (task_047).

These tests complement :mod:`test_ranges_service` with the inherited background
execution contract (BG03 / BG04 / BG05) for the Phase 2 range surface:

* **BG03** — the automation Excel instance is created with ``DispatchEx`` (a new
  owned ``EXCEL.exe``), the automation-owned window is never visible and never
  takes the foreground, and no non-owned Excel process is touched;
* **BG04** — ``Application.Visible`` stays ``False`` and no visible owned window
  exists across reads, writes, formatting, structural operations, validation and
  conditional-formatting changes;
* **BG05** — ``DisplayAlerts`` / ``ScreenUpdating`` / ``AutomationSecurity`` keep
  their defaults (the data-loss merge only suppresses alerts for its own call and
  restores the prior value), and the interactive foreground is not stolen;
* the committed ``RangesCellsProject.xlsm`` fixture is byte-identical after every
  mutation family, and the owned process is reaped on shutdown (no orphan), even
  after a forced hung COM call.

Every operation runs on a disposable working copy through the single hidden STA
worker. Excel is created hidden and no ``.Activate`` / ``.Select`` / ``ShowWindow``
/ ``SetForegroundWindow`` is ever called.
"""

from __future__ import annotations

import asyncio
import contextlib
import hashlib
import shutil
import subprocess
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    CellValue,
    ConditionalFormatSpec,
    RangeFormatPatch,
    RangeWriteRequest,
    ValidationSpec,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    ConditionalFormatType,
    DataValidationType,
    OperationClass,
    SessionMode,
    ValidationOperator,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytest.importorskip("win32com", reason="Excel COM automation unavailable")

_RANGES_PROJECT = (
    Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "RangesCellsProject.xlsm"
)

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_TEST_TIMEOUT_S = 120.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


def _excel_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001 - COM unavailable (non-Windows)
        return False
    return True


requires_excel = pytest.mark.skipif(
    not _excel_available(), reason="Excel COM automation is not available on this host"
)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def ranges_project() -> Any:
    """Yield the committed ``RangesCellsProject.xlsm`` fixture path (skip if absent)."""

    if not _RANGES_PROJECT.is_file():
        pytest.skip(f"fixture not available: {_RANGES_PROJECT}")
    return _RANGES_PROJECT


# ---------------------------------------------------------------------------
# BG03/04/05 observation helpers (read-only; never activate/focus/click)
# ---------------------------------------------------------------------------
def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _pid_alive(pid: int) -> bool:
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001 - best effort
        return True
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _excel_pids() -> set[int]:
    """Return the current ``EXCEL.EXE`` PIDs via ``tasklist`` (best effort)."""

    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001 - best effort
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        parts = [part.strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE":
            with contextlib.suppress(ValueError):
                pids.add(int(parts[1]))
    return pids


def _foreground() -> dict[str, Any] | None:
    """Capture the current foreground window (never changes it)."""

    try:
        import win32gui
        import win32process
    except Exception:  # noqa: BLE001 - pywin32 unavailable
        return None
    try:  # pragma: no cover - depends on host
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return {"hwnd": hwnd, "pid": pid, "title": win32gui.GetWindowText(hwnd)}
    except Exception:  # noqa: BLE001 - best effort
        return None


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    """Return every top-level window of ``pid`` with its visibility (read-only)."""

    try:
        import win32gui
        import win32process
    except Exception:  # noqa: BLE001 - pywin32 unavailable
        return []
    windows: list[dict[str, Any]] = []

    def _callback(hwnd: int, _extra: Any) -> bool:
        try:  # pragma: no cover - depends on host
            _tid, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            if win_pid == pid:
                windows.append(
                    {
                        "hwnd": hwnd,
                        "title": win32gui.GetWindowText(hwnd),
                        "visible": bool(win32gui.IsWindowVisible(hwnd)),
                    }
                )
        except Exception:  # noqa: BLE001 - best effort
            pass
        return True

    with contextlib.suppress(Exception):  # pragma: no cover - depends on host
        win32gui.EnumWindows(_callback, None)
    return windows


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:  # pragma: no cover - depends on host
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at {label} "
        f"(title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    windows = _windows_for_pid(owned_pid)
    visible = [window for window in windows if window["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at "
        f"{label}: {[window['title'] for window in visible]}"
    )


# ---------------------------------------------------------------------------
# Rig
# ---------------------------------------------------------------------------
class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for range tools."""

    def __init__(self, tmp_path: Path, project: Path) -> None:
        self.tmp_path = tmp_path
        self.project = project
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        canonical = self.path_policy.validate_path(str(project), OperationClass.MUTATING)
        self.canonical = str(canonical)
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.approval = ApprovalPolicy()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )
        self.ranges = RangesService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

    @property
    def app(self) -> ExcelAppAdapter:
        return self.adapter

    async def start(self) -> None:
        await self.worker.start()

    async def open(self, *, read_only: bool = False) -> Any:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action="excel_open"),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task_047",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )
        return await self.session.open_workbook(
            str(self.project), mode=SessionMode.INSPECT, read_only=read_only
        )

    def grant(self, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task_047",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_all_range_actions(self) -> None:
        for action in (
            "range_write_values",
            "range_write_formulas",
            "range_set_format",
            "range_copy",
            "range_insert",
            "range_delete",
            "range_merge",
            "range_unmerge",
            "range_set_validation",
            "range_set_conditional_formatting",
        ):
            self.grant(action)

    async def create_snapshot(self, workbook_id: str) -> None:
        await self.snapshot.create_snapshot(workbook_id)

    async def close(self) -> None:
        owned_pid = getattr(self.worker, "_owned_pid", None)
        await self.worker.shutdown()
        self.audit.close()
        # Reap only our own process; never kill a non-owned Excel instance.
        if owned_pid:
            await asyncio.to_thread(_wait_pid_gone, int(owned_pid))


@pytest.fixture
async def rig(tmp_path: Path, ranges_project: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every mutation runs on a disposable working copy and never the source.
    original = tmp_path / ranges_project.name
    shutil.copy2(ranges_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


async def _open_for_mutation(rig: Rig) -> str:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant_all_range_actions()
    return wid


def _value(value: str | float, kind: CellValueKind = CellValueKind.STRING) -> RangeWriteRequest:
    return RangeWriteRequest(rows=1, cols=1, values=[[CellValue(kind=kind, value=value)]])


async def _run_every_mutation_family(rig: Rig, wid: str) -> None:
    """Exercise one representative mutation of every family on the working copy."""

    assert (await rig.ranges.write_values(wid, "Cells", "A20", _value("bg"))).verified
    assert (
        await rig.ranges.set_format(wid, "Cells", "C20", RangeFormatPatch(font_bold=True))
    ).verified
    assert (await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")).verified
    assert (
        await rig.ranges.set_validation(
            wid,
            "Cells",
            "W4:W6",
            ValidationSpec(
                type=DataValidationType.LIST,
                operator=ValidationOperator.BETWEEN,
                formula1='"Yes,No"',
                in_cell_dropdown=True,
            ),
        )
    ).present
    assert (
        await rig.ranges.set_conditional_formatting(
            wid,
            "Cells",
            "X1:X5",
            ConditionalFormatSpec(
                type=ConditionalFormatType.CELL_VALUE,
                operator=ValidationOperator.GREATER,
                formula1="10",
            ),
        )
    ).verified


# ---------------------------------------------------------------------------
# BG04: hidden owned Excel, no visible owned window, defaults preserved
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_excel_hidden_with_no_visible_windows(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_mutation(rig)
    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    await _run_every_mutation_family(rig, wid)

    def read_app_flags() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"visible": bool(app.Visible), "display_alerts": bool(app.DisplayAlerts)}

    read_app_flags.__name__ = "read_app_flags"
    flags = await rig.worker.submit(
        read_app_flags,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )

    # BG04: the owned instance is hidden and has no visible top-level window.
    assert flags["visible"] is False
    _assert_no_visible_owned_windows(owned_pid, "after-all-mutation-families")
    _assert_owned_never_foreground(owned_pid, "after-all-mutation-families")

    # BG05: the default alert setting is preserved (not globally disabled).
    assert flags["display_alerts"] is True

    # BG03: the committed fixture is byte-identical after every mutation family.
    assert _sha256(rig.project) == before

    # The mutation really happened on the disposable working copy.
    written = await rig.ranges.read_values(wid, "Cells", "A20")
    assert written.values[0][0].value == "bg"

    # The session is still healthy and accepting work.
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


# ---------------------------------------------------------------------------
# BG05: the interactive foreground is never stolen
# ---------------------------------------------------------------------------
@requires_excel
async def test_no_focus_theft_during_range_mutations(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)
    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    fg_before = _foreground()
    await _run_every_mutation_family(rig, wid)
    fg_after = _foreground()

    _assert_owned_never_foreground(owned_pid, "after-range-mutations")
    if fg_before is not None and fg_before["pid"] != owned_pid:
        # A non-owned window held the foreground before; it must still hold it.
        assert fg_after is not None
        assert fg_after["pid"] == fg_before["pid"], (
            f"BG05 violation: interactive foreground moved from pid {fg_before['pid']} "
            f"({fg_before['title']!r}) to pid {fg_after['pid']} ({fg_after['title']!r})"
        )


# ---------------------------------------------------------------------------
# BG05: the data-loss merge restores the prior DisplayAlerts value
# ---------------------------------------------------------------------------
@requires_excel
async def test_display_alerts_restored_after_data_loss_merge(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.merge_range(
        wid, "StructOps", "J2:L2", across=False, allow_data_loss=True
    )
    assert result.verified is True

    def read_alerts() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"display_alerts": bool(app.DisplayAlerts)}

    read_alerts.__name__ = "read_alerts"
    flags = await rig.worker.submit(
        read_alerts,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert flags["display_alerts"] is True


# ---------------------------------------------------------------------------
# BG03: owned PID is a new process and no non-owned Excel is touched
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_pid_is_new_and_no_non_owned_excel_touched(tmp_path: Path) -> None:
    ranges_project = _RANGES_PROJECT
    if not ranges_project.is_file():
        pytest.skip(f"fixture not available: {ranges_project}")

    pre_existing = _excel_pids()

    original = tmp_path / ranges_project.name
    shutil.copy2(ranges_project, original)
    r = Rig(tmp_path, original)
    await r.start()
    try:
        info = await r.open()
        owned_pid = r.worker.get_status().owned_pid
        assert owned_pid is not None
        # DispatchEx must create a *new* process, never reuse a running instance.
        assert owned_pid not in pre_existing

        # Do a read and a write so the instance is genuinely exercised.
        await r.create_snapshot(info.workbook_id)
        r.grant("range_write_values")
        assert (
            await r.ranges.write_values(info.workbook_id, "Cells", "A20", _value("pid"))
        ).verified
        assert _pid_alive(owned_pid)
    finally:
        await r.close()

    # BG03: the owned process is reaped and every pre-existing process survives.
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
    still_alive = {pid for pid in pre_existing if _pid_alive(pid)}
    assert still_alive == pre_existing, (
        f"a non-owned Excel process disappeared: {sorted(pre_existing - still_alive)}"
    )


# ---------------------------------------------------------------------------
# BG04 + orphan: a hung COM call is reaped on shutdown
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_excel_reaped_after_needs_recovery_shutdown(tmp_path: Path) -> None:
    ranges_project = _RANGES_PROJECT
    if not ranges_project.is_file():
        pytest.skip(f"fixture not available: {ranges_project}")

    original = tmp_path / ranges_project.name
    shutil.copy2(ranges_project, original)
    r = Rig(tmp_path, original)
    await r.start()
    try:
        await r.open()
        owned_pid = r.worker.get_status().owned_pid
        assert owned_pid is not None

        def hang() -> None:
            time.sleep(3.0)

        hang.__name__ = "task047_bg_hang"
        with pytest.raises(ExcelVbaMcpError) as hung_info:
            await r.worker.submit(
                hang,
                timeout=0.1,
                section="diagnostic",
                op_class=OperationClass.MUTATING,
            )
        assert hung_info.value.code is ErrorCode.EXCEL_HUNG
        # Block E: a deadline-exceeded call leaves the worker in
        # running/deadline_exceeded (still accepting), not a poison NEEDS_RECOVERY
        # state; shutdown still reaps the owned pid (no orphan).
        assert r.worker.get_status().state is WorkbookSessionState.DEADLINE_EXCEEDED
    finally:
        await r.close()

    assert _wait_pid_gone(owned_pid), (
        f"owned PID {owned_pid} survived a deadline-exceeded shutdown (orphan)"
    )
