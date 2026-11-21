"""Independent desktop non-interference checks for the PivotTables workflow (task_052).

This module verifies the inherited background-execution requirements
(``plan_addendum_background_execution.md``) for the Phase 2 PivotTables surface
against a **hidden, automation-created Excel instance** and a disposable working
copy of the committed ``PivotTablesProject.xlsm`` fixture:

* **BG03 — no focus theft:** the owned automation PID never becomes the
  foreground window while the full pivot workflow (inventory, refresh, layout,
  filter) runs, sampled continuously during execution;
* **BG04 — no visible popup:** the owned Excel stays ``Visible = False`` and no
  window owned by the automation PID is visible at any checkpoint;
* **BG05 — user-session preservation:** the committed fixture is byte-identical
  afterwards, the owned PID is reaped after shutdown (no orphan), and a separate
  user-owned Excel instance is neither touched nor killed;
* **recovery:** a forced hung COM call is reported as a structured
  ``EXCEL_HUNG`` error, the worker enters ``NEEDS_RECOVERY``, a follow-up call
  fails fast, and ``recover_session`` rebuilds a fresh hidden instance while
  leaving the original fixture unchanged.

All Excel access is serialised through the single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`; no
``.Activate`` / ``.Select`` is ever called and the owned instance is always
``Visible = False``.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_background_behavior.py -v -s
"""

from __future__ import annotations

import asyncio
import contextlib
import hashlib
import shutil
import subprocess
import sys
import time
from collections.abc import Awaitable, Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelHungError
from excel_vba_mcp.domain.models import PivotFilterSpec
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
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

_REPO_ROOT = Path(__file__).resolve().parents[2]
_PIVOTS_FIXTURE = _REPO_ROOT / "tests" / "fixtures" / "PivotTablesProject.xlsm"

_TEST_TIMEOUT_S = 180.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
_HANG_SECONDS = 3.0
_HANG_CALL_TIMEOUT_S = 0.5

_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_SET_LAYOUT = "pivot_field_set_layout"
_ACTION_SET_AGGREGATION = "pivot_field_set_aggregation"
_ACTION_APPLY_FILTER = "pivot_field_apply_filter"


def _pywin32_available() -> bool:
    return (
        sys.platform == "win32"
        and pythoncom is not None
        and win32com is not None
        and win32gui is not None
        and win32process is not None
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
# Read-only desktop / process helpers (never activate / show / kill)
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


def _pid_from_hwnd(hwnd: int) -> int | None:
    if win32process is None:
        return None
    try:  # pragma: no cover - depends on host
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


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
    visible = [window for window in _windows_for_pid(owned_pid) if window["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at "
        f"{label}: {[window['title'] for window in visible]}"
    )


async def _sample_foreground_while(
    owned_pid: int, operation: Callable[[], Awaitable[Any]]
) -> list[dict[str, Any]]:
    """Run ``operation`` while sampling the foreground window (BG03 evidence)."""

    task = asyncio.ensure_future(operation())
    samples: list[dict[str, Any]] = []
    while not task.done():
        fg = _foreground()
        if fg is not None and fg["pid"] == owned_pid:
            samples.append(fg)
        await asyncio.sleep(0.05)
    await task
    return samples


# ---------------------------------------------------------------------------
# Fixtures / rig — real COM worker + Excel + pivot service
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService for one background check."""

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
        self.pivots = PivotTablesService(
            self.worker,
            self.session,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )

    @property
    def app(self) -> ExcelAppAdapter:
        return self.adapter

    async def start(self) -> None:
        await self.worker.start()

    def grant(self, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task-052-verifier",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def open(self, *, read_only: bool = False) -> Any:
        self.grant("excel_open")
        return await self.session.open_workbook(
            str(self.project), mode=SessionMode.INSPECT, read_only=read_only
        )

    async def open_for_mutation(self) -> Any:
        info = await self.open(read_only=False)
        self.grant(_ACTION_REFRESH_TABLE)
        self.grant(_ACTION_SET_LAYOUT)
        self.grant(_ACTION_SET_AGGREGATION)
        self.grant(_ACTION_APPLY_FILTER)
        return info

    async def close(self) -> None:
        owned_pid = self.worker.get_status().owned_pid
        await self.worker.shutdown()
        self.audit.close()
        if owned_pid is not None:
            _wait_pid_gone(owned_pid, timeout=_ORPHAN_POLL_TIMEOUT_S)


@pytest.fixture
async def rig(tmp_path: Path) -> Any:
    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


def _manual(*items: str) -> PivotFilterSpec:
    return PivotFilterSpec(kind=PivotFilterKind.MANUAL, visible_items=list(items))


async def _read_app_flags(rig: Rig) -> dict[str, Any]:
    def read() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"visible": bool(app.Visible), "display_alerts": bool(app.DisplayAlerts)}

    read.__name__ = "read_app_flags"
    return dict(
        await rig.worker.submit(
            read,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )


async def _run_pivot_workflow(rig: Rig, workbook_id: str) -> None:
    """Exercise the full pivot read + mutation surface in one workflow."""

    await rig.pivots.list_pivot_tables(workbook_id)
    await rig.pivots.get_pivot_table(workbook_id, "ptA")
    await rig.pivots.get_pivot_cache(workbook_id, 1)
    await rig.pivots.list_pivot_fields(workbook_id, "ptA")
    await rig.pivots.refresh_pivot_table(workbook_id, "ptA")
    await rig.pivots.set_pivot_field_layout(
        workbook_id, "ptA", "Quarter", PivotFieldOrientation.ROW
    )
    await rig.pivots.set_pivot_field_aggregation(workbook_id, "ptA", "Sales", PivotAggregation.MAX)
    await rig.pivots.apply_pivot_field_filter(workbook_id, "ptA", "Region", _manual("North"))
    await rig.pivots.get_pivot_field_filters(workbook_id, "ptA", "Region")


# ---------------------------------------------------------------------------
# BG03 — no focus theft while the pivot workflow runs
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_excel_never_takes_foreground(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    samples = await _sample_foreground_while(
        owned_pid, lambda: _run_pivot_workflow(rig, info.workbook_id)
    )

    assert not samples, f"BG03 violation: owned PID {owned_pid} held foreground: {samples}"
    _assert_owned_never_foreground(owned_pid, "pivot-workflow")

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG04 — owned Excel stays hidden with no visible window
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_excel_stays_hidden_with_no_visible_windows(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    checkpoints: list[tuple[str, Callable[[], Awaitable[Any]]]] = [
        ("after-inventory", lambda: rig.pivots.list_pivot_tables(info.workbook_id)),
        ("after-refresh", lambda: rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")),
        (
            "after-layout",
            lambda: rig.pivots.set_pivot_field_layout(
                info.workbook_id, "ptA", "Quarter", PivotFieldOrientation.ROW
            ),
        ),
        (
            "after-filter",
            lambda: rig.pivots.apply_pivot_field_filter(
                info.workbook_id, "ptA", "Region", _manual("North")
            ),
        ),
    ]

    for label, operation in [("after-open", None), *checkpoints]:
        if operation is not None:
            await operation()
        flags = await _read_app_flags(rig)
        assert flags["visible"] is False, f"owned Excel became visible at {label}"
        _assert_no_visible_owned_windows(owned_pid, label)
        _assert_owned_never_foreground(owned_pid, label)

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG05 — fixture unchanged, owned PID reaped (no orphan)
# ---------------------------------------------------------------------------
@requires_excel
async def test_fixture_unchanged_and_owned_pid_reaped(tmp_path: Path) -> None:
    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    committed_hash = _sha256(pivot_project)

    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)
    original_hash = _sha256(original)

    r = Rig(tmp_path, original)
    owned_pid: int | None = None
    try:
        await r.start()
        owned_pid = r.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await r.open_for_mutation()
        await r.snapshot.create_snapshot(info.workbook_id)
        await _run_pivot_workflow(r, info.workbook_id)
        await r.session.close_workbook(info.workbook_id)

        # The working copy is disposable; the committed fixture is never touched.
        assert _sha256(original) == original_hash
        assert _sha256(pivot_project) == committed_hash
    finally:
        await r.close()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown (orphan)"


# ---------------------------------------------------------------------------
# BG05 — a user-owned Excel instance is neither touched nor killed
# ---------------------------------------------------------------------------
@requires_excel
async def test_user_owned_excel_instance_is_untouched(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None
    original_hash = _sha256(rig.project)
    committed_hash = _sha256(_PIVOTS_FIXTURE)

    # A separate, non-owned Excel instance (the "user" Excel). Kept hidden so the
    # test itself never shows a window; it is never passed to the worker.
    bystander = win32com.client.DispatchEx("Excel.Application")
    bystander_pid: int | None = None
    try:
        bystander.Visible = False
        bystander_pid = _pid_from_hwnd(int(bystander.Hwnd))
        assert bystander_pid is not None

        await _run_pivot_workflow(rig, info.workbook_id)

        # The owned session closes without touching the bystander process.
        status = await rig.session.close_session()
        assert status.state is WorkbookSessionState.CLOSED

        assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after close"
        assert _pid_alive(bystander_pid), "close_session must not kill a non-owned Excel"
        assert int(bystander.Workbooks.Count) == 0
        assert bool(bystander.Visible) is False
        assert _sha256(rig.project) == original_hash
        assert _sha256(_PIVOTS_FIXTURE) == committed_hash

        _assert_owned_never_foreground(owned_pid, "after-close")
    finally:
        with contextlib.suppress(Exception):
            bystander.Quit()
        if bystander_pid is not None:
            _wait_pid_gone(bystander_pid)


# ---------------------------------------------------------------------------
# Recovery — a hung COM call is reported truthfully and recovers to hidden
# ---------------------------------------------------------------------------
@requires_excel
async def test_hung_com_reports_excel_hung_and_recovers_hidden(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    await rig.pivots.list_pivot_tables(info.workbook_id)

    old_pid = rig.worker.get_status().owned_pid
    assert old_pid is not None
    original_hash = _sha256(rig.project)

    # Force a hung COM call: a blocking callable that outlives its soft deadline.
    def hang() -> str:
        time.sleep(_HANG_SECONDS)
        return "finished"

    with pytest.raises(ExcelHungError) as hung_exc:
        await rig.worker.submit(
            hang,
            timeout=_HANG_CALL_TIMEOUT_S,
            section="execute",
            op_class=OperationClass.EXECUTING,
        )
    assert hung_exc.value.code is ErrorCode.EXCEL_HUNG
    assert rig.worker.get_status().state is WorkbookSessionState.NEEDS_RECOVERY

    # Fail fast: no new parallel COM call is enqueued while NEEDS_RECOVERY.
    submits_before = rig.worker._token_seq  # read-only diagnostic
    started = time.monotonic()
    with pytest.raises(ExcelHungError):
        await rig.worker.submit(
            lambda: "should-not-run",
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    elapsed = time.monotonic() - started
    assert elapsed < 2.0, f"fail-fast submit took {elapsed:.2f}s (must not block on the hang)"
    assert rig.worker._token_seq == submits_before, "no work item may be queued while hung"

    # Recovery rebuilds a fresh hidden instance and discards the open workbook.
    status = await rig.session.recover_session(reason="task_052_pivot_hung")
    assert status.state is WorkbookSessionState.RUNNING
    assert rig.session.list_workbooks_sync() == []

    new_pid = rig.worker.get_status().owned_pid
    assert new_pid is not None
    assert new_pid != old_pid, "recover must create a fresh owned instance"
    assert _sha256(rig.project) == original_hash, "recover must not modify the original"

    flags = await _read_app_flags(rig)
    assert flags["visible"] is False, "the rebuilt instance must stay Visible=False"
    _assert_owned_never_foreground(new_pid, "after-recover")
    _assert_no_visible_owned_windows(new_pid, "after-recover")

    assert _wait_pid_gone(old_pid), f"old owned PID {old_pid} still alive after recover"
