"""Live-Excel verification of Sheets inventory + worksheet ops (task_024).

This is the task_024 independent verification of the production
:class:`~excel_vba_mcp.application.sheets_service.SheetsService` and
:class:`~excel_vba_mcp.adapters.excel.sheets_adapter.SheetsAdapter` against the
target Excel build. It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter` (hidden /
non-activating) and a real
:class:`~excel_vba_mcp.application.session_service.SessionService`, always
hidden / non-activating, and only ever touches a temporary **copy** of the
committed ``CleanProject.xlsm`` fixture. The committed fixture is never opened
or modified.

Covered:

* ``sheets_list`` returns worksheets **and** a non-worksheet (Chart) sheet; a
  Chart sheet is added via the adapter so the ``Sheets`` vs ``Worksheets``
  distinction is exercised on the real object model.
* ``sheet_get`` metadata: name, index, object_type, visibility, code_name,
  protected, table_count, used_range.
* ``sheet_add`` / ``sheet_rename`` / ``sheet_copy`` / ``sheet_set_visibility``
  succeed and their postconditions hold.
* ``sheet_delete`` requires snapshot + approval; deleting the last visible sheet
  is rejected; the original fixture is unchanged.
* ``CodeName`` is unchanged by rename.
* BG03/BG04/BG05: the owned instance is never foreground / has no visible
  windows / user Excel processes are untouched; the owned PID is reaped after
  shutdown.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_sheets_service.py -v -s
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

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-untyped]
    import win32com  # type: ignore[import-untyped]
    import win32gui  # type: ignore[import-untyped]
    import win32process  # type: ignore[import-untyped]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    SheetObjectType,
    SheetVisibility,
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

# ---------------------------------------------------------------------------
# Constants / skip helpers
# ---------------------------------------------------------------------------
_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

# Poll window for the owned-PID reaping check after shutdown.
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


def _pywin32_available() -> bool:
    return (
        sys.platform == "win32"
        and pythoncom is not None
        and win32com is not None
        and win32gui is not None
    )


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(f"fixture not available: {path}")
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# BG03/04/05 helpers (read-only; never activate/focus/click)
# ---------------------------------------------------------------------------
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


def _excel_pids() -> set[int]:
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
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
    return pids


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:  # pragma: no cover - depends on host
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        title = win32gui.GetWindowText(hwnd)
        return {"hwnd": hwnd, "pid": pid, "title": title}
    except Exception:  # noqa: BLE001 - best effort
        return None


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    if win32gui is None or win32process is None:
        return []
    windows: list[dict[str, Any]] = []

    def _cb(hwnd: int, _extra: Any) -> bool:
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
        win32gui.EnumWindows(_cb, None)
    return windows


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at "
        f"{label} (title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at "
        f"{label}: {[w['title'] for w in visible]}"
    )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session + snapshot + sheets service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.snapshots_dir = tmp_dir / "snapshots"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.snapshots_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {
                "work": str(self.work_dir),
                "projects": str(self.projects_dir),
                "snapshots": str(self.snapshots_dir),
            }
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
        self.snapshot = SnapshotService(
            SnapshotStore(self.snapshots_dir),
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )
        self.service = SheetsService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self._original: Path | None = None
        self._writable = False

    def bind_original(self, original: Path, *, writable: bool = False) -> None:
        self._original = original
        self._writable = writable

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

    async def open(self) -> Any:
        assert self._original is not None
        _grant(
            self.approval,
            ApprovalKey(canonical_path=self.canonical, action="excel_open"),
        )
        return await self.session.open_workbook(
            str(self._original),
            mode=SessionMode.INSPECT,
            read_only=not self._writable,
        )

    def grant_mutation(self, action: str) -> None:
        _grant(
            self.approval,
            ApprovalKey(canonical_path=self.canonical, action=action),
            scope=ApprovalScope.SESSION,
        )


def _grant(
    approval: ApprovalPolicy,
    key: ApprovalKey,
    *,
    scope: ApprovalScope = ApprovalScope.SESSION,
) -> None:
    approval.grant(
        ApprovalGrant(
            key=key,
            scope=scope,
            granted_at=datetime.now(UTC),
            granted_by="task024-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


def _make_rig(tmp_path: Path, *, writable: bool = True) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original, writable=writable)
    return rig


def _working_path(rig: _Rig, workbook_id: str) -> str:
    record = rig.session._records[workbook_id]
    return record.working_path


# ---------------------------------------------------------------------------
# Test 1: list (worksheet + non-worksheet), get metadata, BG03/04/05 + no orphan
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_list_sheets_includes_chart_and_metadata(tmp_path: Path) -> None:
    """A Chart sheet appears via ``Sheets`` but not ``Worksheets``; metadata holds."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    baseline_pids = _excel_pids()
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        assert info.read_only is False
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # Inject a Chart sheet on the STA thread so a non-worksheet type exists.
        def add_chart() -> None:
            app = rig.app_adapter.app
            target = _working_path(rig, info.workbook_id).lower()
            for wb in app.Workbooks:
                if str(wb.FullName).lower() == target:
                    wb.Charts.Add().Name = "VerifyChart"
                    return
            raise AssertionError("working copy not open")

        await rig.worker.submit(
            add_chart, timeout=60.0, section="diagnostic", op_class=OperationClass.MUTATING
        )

        sheets = await rig.service.list_sheets(info.workbook_id)
        by_name = {s.name: s for s in sheets}
        assert "VerifyChart" in by_name
        assert by_name["VerifyChart"].object_type is SheetObjectType.CHART

        worksheet_names = [
            n for n, s in by_name.items() if s.object_type is SheetObjectType.WORKSHEET
        ]
        assert worksheet_names, "expected at least one worksheet"

        one = await rig.service.get_sheet(info.workbook_id, worksheet_names[0])
        assert one.name == worksheet_names[0]
        assert one.index >= 1
        assert one.visibility in (SheetVisibility.VISIBLE, SheetVisibility.HIDDEN)
        # CodeName is a read-only identifier and must be present for a worksheet.
        assert one.code_name is not None

        _assert_owned_never_foreground(owned_pid, "sheets-list")
        _assert_no_visible_owned_windows(owned_pid, "sheets-list")
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        print(
            f"[sheets-list] sheets={sorted(by_name)} owned_pid={owned_pid} "
            f"worksheets={worksheet_names}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: add / rename (CodeName unchanged) / copy / set_visibility
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_add_rename_copy_visibility(tmp_path: Path) -> None:
    """add/rename/copy/set_visibility succeed; rename keeps CodeName; BG holds."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    baseline_pids = _excel_pids()
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        snapshot = await rig.snapshot.create_snapshot(info.workbook_id)
        assert snapshot.snapshot_id

        # add
        rig.grant_mutation("sheet_add")
        added = await rig.service.add_sheet(info.workbook_id, "VerifyAdded")
        assert added.name == "VerifyAdded"
        assert added.object_type is SheetObjectType.WORKSHEET
        added_code_name = added.code_name

        # rename keeps CodeName
        rig.grant_mutation("sheet_rename")
        renamed = await rig.service.rename_sheet(info.workbook_id, "VerifyAdded", "VerifyRenamed")
        assert renamed.name == "VerifyRenamed"
        assert renamed.code_name == added_code_name, "CodeName must not change on rename"

        # copy
        rig.grant_mutation("sheet_copy")
        copied = await rig.service.copy_sheet(info.workbook_id, "VerifyRenamed")
        assert copied.name.startswith("VerifyRenamed")
        assert copied.name != "VerifyRenamed"

        # set_visibility (hide the copy; it is not the last visible sheet)
        rig.grant_mutation("sheet_set_visibility")
        hidden = await rig.service.set_visibility(info.workbook_id, copied.name, False)
        assert hidden.visibility is SheetVisibility.HIDDEN

        _assert_owned_never_foreground(owned_pid, "sheets-mutations")
        _assert_no_visible_owned_windows(owned_pid, "sheets-mutations")
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"

        print(
            f"[sheets-mutations] added={added.name} renamed={renamed.name} "
            f"copied={copied.name} hidden={hidden.visibility.value} owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: delete guards (snapshot, approval, last-visible) + success
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_delete_guards_and_success(tmp_path: Path) -> None:
    """delete requires snapshot + approval; last visible sheet rejected; success."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    baseline_pids = _excel_pids()
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        # 1. No snapshot -> SNAPSHOT_FAILED (snapshot is the first guard).
        rig.grant_mutation("sheet_delete")
        with pytest.raises(Exception) as no_snapshot:
            await rig.service.delete_sheet(info.workbook_id, "VerifyNone")
        assert (
            "SNAPSHOT_FAILED" in str(no_snapshot.value)
            or "snapshot" in str(no_snapshot.value).lower()
        )

        # Provide a snapshot.
        snapshot = await rig.snapshot.create_snapshot(info.workbook_id)
        assert snapshot.snapshot_id

        # 2. Snapshot present but no approval -> denied by the approval gate.
        #    Use a disposable sheet on a *separate* rig whose workbook has more
        #    than one visible worksheet, so the visible-sheet guard does not fire
        #    and the approval gate is the only guard that can block.
        rig2 = _make_rig(tmp_path / "rig2", writable=True)
        await rig2.start()
        try:
            info2 = await rig2.open()
            snap2 = await rig2.snapshot.create_snapshot(info2.workbook_id)
            assert snap2.snapshot_id
            rig2.grant_mutation("sheet_add")
            # Add two sheets so >1 visible worksheet exists (visible guard won't fire).
            await rig2.service.add_sheet(info2.workbook_id, "VerifyKeep")
            added0 = await rig2.service.add_sheet(info2.workbook_id, "VerifyNoApproval")
            assert added0.name == "VerifyNoApproval"
            with pytest.raises(Exception) as no_approval:
                await rig2.service.delete_sheet(info2.workbook_id, "VerifyNoApproval")
            assert (
                "denied" in str(no_approval.value).lower()
                or "approval" in str(no_approval.value).lower()
            )
        finally:
            await rig2.stop()

        # Add a disposable sheet and delete it successfully.
        rig.grant_mutation("sheet_add")
        added = await rig.service.add_sheet(info.workbook_id, "VerifyDelete")
        assert added.name == "VerifyDelete"

        rig.grant_mutation("sheet_delete")
        deleted = await rig.service.delete_sheet(info.workbook_id, "VerifyDelete")
        assert deleted.name == "VerifyDelete"
        remaining = {s.name for s in await rig.service.list_sheets(info.workbook_id)}
        assert "VerifyDelete" not in remaining

        # 3. Deleting the last visible sheet is rejected. Make every sheet hidden
        #    except one worksheet, then attempt to delete that last visible one.
        def isolate_last_visible() -> str:
            app = rig.app_adapter.app
            target = _working_path(rig, info.workbook_id).lower()
            for wb in app.Workbooks:
                if str(wb.FullName).lower() != target:
                    continue
                visible = [s for s in wb.Worksheets if int(s.Visible) == 0]
                # Hide all but the first visible worksheet.
                for sheet in visible[1:]:
                    sheet.Visible = 1
                # Hide chart sheets too (non-worksheet, not counted anyway).
                return str(visible[0].Name) if visible else ""
            raise AssertionError("working copy not open")

        last_visible = await rig.worker.submit(
            isolate_last_visible,
            timeout=60.0,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
        if last_visible:
            with pytest.raises(Exception) as last_guard:
                await rig.service.delete_sheet(info.workbook_id, last_visible)
            assert "visible" in str(last_guard.value).lower()
            remaining = {s.name for s in await rig.service.list_sheets(info.workbook_id)}
            assert last_visible in remaining, "last visible sheet must survive"

        _assert_owned_never_foreground(owned_pid, "sheets-delete")
        _assert_no_visible_owned_windows(owned_pid, "sheets-delete")
        missing_user_pids = baseline_pids - _excel_pids()
        assert not missing_user_pids, f"user Excel PIDs disappeared: {missing_user_pids}"
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        print(
            f"[sheets-delete] deleted={deleted.name} last_visible={last_visible!r} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
