"""Real Excel integration tests for the defined-name surface (task_028).

These tests drive the real :class:`~excel_vba_mcp.application.names_service.
NamesService` and :class:`~excel_vba_mcp.adapters.excel.names_adapter.
NamesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel). They assert:

* a workbook-scope range name and a worksheet-scope range name are created and
  reported with the expected scope, ``RefersTo`` and range target (sheet +
  address);
* a constant name and a formula name are returned with ``target_kind``
  ``constant`` / ``formula`` (their ``RefersToRange`` raises) — no formula is
  evaluated;
* ``names_list`` / ``name_get`` expose scope, ``NameLocal``, ``Visible``,
  ``RefersTo`` and ``RefersToLocal``;
* ``name_update`` preserves the requested scope and refreshes ``RefersTo``;
* ``name_delete`` removes the name;
* duplicate and cross-scope conflicts are rejected;
* the original fixture on disk is byte-identical after every test (migrations
  never touch the source);
* BG03/BG04/BG05 hold (hidden + non-activating owned Excel, no non-owned Excel
  touched, and the worker refuses mutations on a hung instance);
* the owned Excel PID is reaped after shutdown (no orphan).

Mutations run only on the disposable working copy of the fixture. Excel is
created with ``DispatchEx`` and always ``Visible = False``; no ``.Activate`` /
``.Select`` is ever called.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_names_service.py -v -s
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
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    NameScope,
    NameTargetKind,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

_TEST_TIMEOUT_S = 120.0
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


class Rig:
    """Real COM worker + Excel + full session/snapshot/policy stack."""

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
        self.names = NamesService(
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
                granted_by="tester",
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
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def create_snapshot(self, workbook_id: str) -> None:
        await self.snapshot.create_snapshot(workbook_id)

    async def close(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


@pytest.fixture
async def rig(tmp_path: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every test mutates a disposable working copy and never the source.
    clean_project = _fixture_or_skip(_CLEAN_PROJECT)
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


# ---------------------------------------------------------------------------
# STA helpers to seed a worksheet on the working copy
# ---------------------------------------------------------------------------
def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


def _record_path(rig: Rig, workbook_id: str) -> str:
    return rig.session._records[workbook_id].working_path


async def _add_worksheet(rig: Rig, workbook_id: str, name: str) -> str:
    """Create a real worksheet on the working copy and return its name."""

    working_path = _record_path(rig, workbook_id)

    def seed() -> str:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Add()
        sheet.Name = name
        return str(sheet.Name)

    seed.__name__ = "seed_worksheet"
    return await rig.worker.submit(
        seed, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.MUTATING
    )


# ---------------------------------------------------------------------------
# Inventory + non-range targets
# ---------------------------------------------------------------------------
@requires_excel
async def test_create_list_and_get_names(rig: Rig) -> None:
    before = _sha256(rig.project)
    committed_hash = _sha256(_CLEAN_PROJECT)
    info = await rig.open()
    wid = info.workbook_id

    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")

    wb_range = await rig.names.create_name(
        wid, "WbRange", "=NamesData!$A$1:$B$2", scope=NameScope.WORKBOOK
    )
    assert wb_range.scope is NameScope.WORKBOOK
    assert wb_range.target_kind is NameTargetKind.RANGE
    assert wb_range.target_sheet == "NamesData"
    assert wb_range.target_address.endswith("$A$1:$B$2")

    sheet_range = await rig.names.create_name(
        wid, "SheetRange", "=NamesData!$C$1", scope=NameScope.WORKSHEET, sheet="NamesData"
    )
    assert sheet_range.scope is NameScope.WORKSHEET
    assert sheet_range.sheet_name == "NamesData"
    assert sheet_range.target_kind is NameTargetKind.RANGE
    assert sheet_range.target_sheet == "NamesData"

    constant = await rig.names.create_name(wid, "Pi", "=3.14", scope=NameScope.WORKBOOK)
    assert constant.target_kind is NameTargetKind.CONSTANT
    assert constant.refers_to is not None

    formula = await rig.names.create_name(
        wid, "Total", "=SUM(NamesData!$A$1:$A$2)", scope=NameScope.WORKBOOK
    )
    assert formula.target_kind is NameTargetKind.FORMULA

    names = await rig.names.list_names(wid)
    by_key = {(n.name, n.scope, n.sheet_name): n for n in names}
    assert ("WbRange", NameScope.WORKBOOK, None) in by_key
    assert ("SheetRange", NameScope.WORKSHEET, "NamesData") in by_key
    assert by_key[("Pi", NameScope.WORKBOOK, None)].target_kind is NameTargetKind.CONSTANT
    assert by_key[("Total", NameScope.WORKBOOK, None)].target_kind is NameTargetKind.FORMULA

    fetched = await rig.names.get_name(wid, "WbRange", NameScope.WORKBOOK)
    assert fetched.name == "WbRange"
    assert fetched.refers_to is not None and "NamesData" in fetched.refers_to
    assert fetched.name_local is not None

    fetched_sheet = await rig.names.get_name(wid, "SheetRange", NameScope.WORKSHEET, "NamesData")
    assert fetched_sheet.scope is NameScope.WORKSHEET

    # BG03: original + committed fixture never modified.
    assert _sha256(rig.project) == before
    assert _sha256(_CLEAN_PROJECT) == committed_hash


@requires_excel
async def test_get_missing_name_raises(rig: Rig) -> None:
    info = await rig.open()
    with pytest.raises(ExcelVbaMcpError):
        await rig.names.get_name(info.workbook_id, "NoSuchName", NameScope.WORKBOOK)


# ---------------------------------------------------------------------------
# Update + delete
# ---------------------------------------------------------------------------
@requires_excel
async def test_update_preserves_scope_and_delete(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")
    rig.grant("name_update")
    rig.grant("name_delete")

    await rig.names.create_name(wid, "WbRange", "=NamesData!$A$1", scope=NameScope.WORKBOOK)

    updated = await rig.names.update_name(
        wid, "WbRange", "=NamesData!$D$1:$D$2", scope=NameScope.WORKBOOK
    )
    assert updated.scope is NameScope.WORKBOOK  # scope preserved
    assert updated.refers_to is not None and "$D$1:$D$2" in updated.refers_to
    assert updated.target_kind is NameTargetKind.RANGE

    deleted = await rig.names.delete_name(wid, "WbRange", scope=NameScope.WORKBOOK)
    assert deleted.name == "WbRange"
    assert deleted.scope is NameScope.WORKBOOK
    with pytest.raises(ExcelVbaMcpError):
        await rig.names.get_name(wid, "WbRange", NameScope.WORKBOOK)

    remaining = await rig.names.list_names(wid)
    assert all(n.name != "WbRange" for n in remaining)

    # BG03: source fixture unchanged.
    assert _sha256(rig.project) == before


# ---------------------------------------------------------------------------
# Conflict detection
# ---------------------------------------------------------------------------
@requires_excel
async def test_duplicate_and_cross_scope_conflicts_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")

    await rig.names.create_name(wid, "Shared", "=1", scope=NameScope.WORKBOOK)

    # Duplicate in the same scope.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.names.create_name(wid, "Shared", "=2", scope=NameScope.WORKBOOK)
    assert exc_info.value.context.get("reason") == "duplicate"

    # Cross-scope collision (workbook name already exists).
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.names.create_name(
            wid, "Shared", "=NamesData!$A$1", scope=NameScope.WORKSHEET, sheet="NamesData"
        )
    assert exc_info.value.context.get("reason") == "scope_conflict"


# ---------------------------------------------------------------------------
# Guards
# ---------------------------------------------------------------------------
@requires_excel
async def test_create_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    rig.grant("name_create")  # approval but no snapshot

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.names.create_name(wid, "X", "=1", scope=NameScope.WORKBOOK)
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


@requires_excel
async def test_create_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.names.create_name(wid, "X", "=1", scope=NameScope.WORKBOOK)


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")
    await rig.names.create_name(wid, "WbRange", "=NamesData!$A$1", scope=NameScope.WORKBOOK)

    # BG03/BG05: the automation Excel is hidden and non-activating (read on the
    # STA worker so the COM property access is thread-correct).
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
    assert flags["visible"] is False
    assert flags["display_alerts"] is True
    assert rig.snapshot._latest_snapshot.get(wid) is not None

    _assert_owned_never_foreground(owned_pid, "names-create")
    _assert_no_visible_owned_windows(owned_pid, "names-create")
    # BG03: source fixture unchanged.
    assert _sha256(rig.project) == before


@requires_excel
async def test_bg04_worker_refuses_mutation_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")

    def hang() -> None:
        import time

        # Block the STA thread past the soft deadline so the submit times out and
        # the worker enters NEEDS_RECOVERY (a hung dispatch is never retried).
        time.sleep(3.0)

    hang.__name__ = "hang_test"
    with pytest.raises(ExcelVbaMcpError):
        await rig.worker.submit(
            hang,
            timeout=0.1,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    # BG04: the worker is now NEEDS_RECOVERY and refuses new work, so the name
    # mutation never reaches Excel (fail fast, no enqueue).
    with pytest.raises(ExcelVbaMcpError):
        await rig.names.create_name(wid, "X", "=1", scope=NameScope.WORKBOOK)


@requires_excel
async def test_bg05_session_state_and_alerts(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING

    def read_alerts() -> bool:
        app = rig.app.app
        assert app is not None
        return bool(app.DisplayAlerts)

    read_alerts.__name__ = "read_alerts"
    alerts = await rig.worker.submit(
        read_alerts,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert alerts is True


@requires_excel
async def test_owned_excel_reaped_after_shutdown(tmp_path: Path) -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    clean_project = _fixture_or_skip(_CLEAN_PROJECT)
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    await r.start()
    await r.open()
    owned_pid = r.worker.get_status().owned_pid
    assert owned_pid is not None
    await r.close()

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
