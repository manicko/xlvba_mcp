"""Real Excel integration tests for the targeted PivotTable / PivotCache refresh (task_049).

These tests drive the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
against a **hidden, automation-created Excel instance** (never the user's Excel)
and a disposable working copy of the committed ``PivotTablesProject.xlsm``
fixture. They assert:

* ``refresh_pivot_table("ptA")`` → ``refreshed`` with ``RefreshDate`` advanced
  (``RefreshTable`` returns a Boolean truth signal);
* ``refresh_pivot_cache(1)`` (the shared cache) → both ``ptA`` and ``ptB`` are
  reported in ``affected_pivot_tables`` and the cache ``RefreshDate`` advanced;
* ``refresh_pivot_table("ptC")`` (the ListObject-backed cache 2) → ``refreshed``;
* the unsupported-source path is asserted only as "no unsupported source is
  present" (all fixture caches are local ``xlDatabase``); the mapping is covered
  by unit tests against a fake workbook;
* the guard chain — a read-only open, a missing snapshot and a missing approval
  each block the refresh before any STA dispatch;
* BG03/BG04/BG05 — the automation Excel stays hidden and non-activating, the
  owned PID never holds foreground, no owned window is visible, and the owned PID
  is reaped after shutdown (no orphan);
* the committed fixture and the temp original are byte-identical afterwards.

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.

task_052 adds independent acceptance coverage to this module: the refreshed
result range is re-read from Excel (address, shape and a changed source value),
a targeted ``ptC`` refresh is shown to leave cache 1 untouched, the public
``pivot_tools`` layer is checked, and a real consolidation-sourced cache is
rejected with ``UNSUPPORTED_PIVOT_SOURCE`` for the cache-scoped entry point.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_refresh.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import pivot_tools
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotRefreshResult
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotRefreshMethod,
    PivotRefreshStatus,
    PivotSourceType,
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

#: Approval action names (must match the service constants).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_REFRESH_CACHE = "pivot_cache_refresh"


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
# Fixtures / rig — real COM worker + Excel + PivotTablesService (refresh-wired)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService wired for guarded refresh."""

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
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def open(self, *, read_only: bool = False) -> Any:
        self.grant("excel_open")
        return await self.session.open_workbook(
            str(self.project), mode=SessionMode.INSPECT, read_only=read_only
        )

    async def open_for_refresh(self) -> Any:
        """Open writable and grant both refresh actions (writable copy + gates)."""

        info = await self.open(read_only=False)
        self.grant(_ACTION_REFRESH_TABLE)
        self.grant(_ACTION_REFRESH_CACHE)
        return info

    async def close(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


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


# ---------------------------------------------------------------------------
# RefreshTable (table-scoped) — ptA on the shared cache 1
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_table_pt_a_advances_refresh_date(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)

    before = await rig.pivots.get_pivot_table(info.workbook_id, "ptA")

    result = await rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")

    assert isinstance(result, PivotRefreshResult)
    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.refreshed is True
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.refresh_return is True
    # ptA and ptB share cache 1, so both are reported as affected.
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]

    # RefreshDate advanced (the truth signal), not merely "no exception".
    assert result.refresh_date_before is not None
    assert result.refresh_date_after is not None
    assert result.refresh_date_after > result.refresh_date_before

    after = await rig.pivots.get_pivot_table(info.workbook_id, "ptA")
    assert after.refresh_date is not None
    assert before.refresh_date is not None
    assert after.refresh_date > before.refresh_date

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# RefreshCache (cache-scoped) — the shared cache 1 backs both ptA and ptB
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_shared_cache_reports_both_tables(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)

    before = await rig.pivots.get_pivot_cache(info.workbook_id, 1)

    result = await rig.pivots.refresh_pivot_cache(info.workbook_id, 1)

    assert result.method is PivotRefreshMethod.REFRESH_CACHE
    assert result.refreshed is True
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.refresh_return is None  # PivotCache.Refresh returns no Boolean
    # The cache is shared: both tables are reported as affected.
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]
    assert result.cache_index == 1

    after = await rig.pivots.get_pivot_cache(info.workbook_id, 1)
    assert after.refresh_date is not None
    assert before.refresh_date is not None
    assert after.refresh_date > before.refresh_date

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# RefreshTable (table-scoped) — ptC on the ListObject-backed cache 2
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_table_pt_c_listobject_cache(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)

    result = await rig.pivots.refresh_pivot_table(info.workbook_id, "ptC")

    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.refreshed is True
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.cache_index == 2
    # ptC is alone on cache 2.
    assert result.affected_pivot_tables == ["ptC"]

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Unsupported source: the fixture has none (all local xlDatabase)
# ---------------------------------------------------------------------------
@requires_excel
async def test_no_unsupported_source_in_fixture(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    inventory = await rig.pivots.list_pivot_tables(info.workbook_id)

    from excel_vba_mcp.domain.statuses import PivotSourceType

    assert all(pt.source_type is PivotSourceType.DATABASE for pt in inventory.pivot_tables)
    assert all(pt.source_status == "supported" for pt in inventory.pivot_tables)

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Guard chain: read-only, no snapshot, no approval
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_rejected_on_read_only_open(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    rig.grant(_ACTION_REFRESH_TABLE)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_refresh_rejected_without_snapshot(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    # No snapshot created: the guard must block before any refresh.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_refresh_rejected_without_approval(rig: Rig) -> None:
    # Open writable but do NOT grant the refresh actions.
    rig.grant("excel_open")
    info = await rig.session.open_workbook(
        str(rig.project), mode=SessionMode.INSPECT, read_only=False
    )
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.refresh_pivot_cache(info.workbook_id, 1)

    assert exc_info.value.code is ErrorCode.COM_BUSY

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 + original unchanged + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_hidden_owned_excel_during_refresh(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    # Run the refresh (the mutating path under test) while hidden.
    result = await rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")
    assert result.refreshed is True

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

    # BG03/BG04: owned Excel hidden + non-activating, no visible owned window.
    assert flags["visible"] is False
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING
    _assert_owned_never_foreground(owned_pid, "pivot-refresh")
    _assert_no_visible_owned_windows(owned_pid, "pivot-refresh")

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_committed_fixture_unchanged(tmp_path: Path) -> None:
    """BG05: the committed fixture is byte-identical after the refresh runs."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    before = _sha256(pivot_project)
    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        info = await r.open_for_refresh()
        await r.snapshot.create_snapshot(info.workbook_id)
        await r.pivots.refresh_pivot_table(info.workbook_id, "ptA")
        await r.pivots.refresh_pivot_cache(info.workbook_id, 1)
        await r.session.close_workbook(info.workbook_id)
    finally:
        await r.close()

    assert _sha256(pivot_project) == before


@requires_excel
async def test_owned_excel_reaped_after_shutdown() -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    tmp_path = Path(tempfile.mkdtemp(prefix="pivot_refresh_"))
    try:
        original = tmp_path / pivot_project.name
        shutil.copy2(pivot_project, original)
        r = Rig(tmp_path, original)
        await r.start()
        try:
            info = await r.open_for_refresh()
            await r.snapshot.create_snapshot(info.workbook_id)
            await r.pivots.refresh_pivot_table(info.workbook_id, "ptA")
            owned_pid = r.worker.get_status().owned_pid
            assert owned_pid is not None
        finally:
            await r.close()
    finally:
        shutil.rmtree(tmp_path, ignore_errors=True)

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# STA-local readers (independent evidence from the live Excel object model)
# ---------------------------------------------------------------------------
async def _read_pivot_body(rig: Rig, pivot_table: str) -> dict[str, Any]:
    """Read ``DataBodyRange`` geometry and the grand-total cell from Excel."""

    def read() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        target = pivot_table.strip().lower()
        for workbook in app.Workbooks:
            for sheet in workbook.Worksheets:
                try:
                    pivots = sheet.PivotTables()
                    count = int(pivots.Count)
                except Exception:  # noqa: BLE001 - sheet without PivotTables
                    continue
                for index in range(1, count + 1):
                    pivot = pivots.Item(index)
                    if str(pivot.Name).lower() != target:
                        continue
                    body = pivot.DataBodyRange
                    rows = int(body.Rows.Count)
                    cols = int(body.Columns.Count)
                    return {
                        "address": str(body.Address),
                        "rows": rows,
                        "cols": cols,
                        "grand_total": body.Cells(rows, cols).Value,
                        "sales_total": _sales_grand_total(pivot),
                        "refresh_date": str(pivot.RefreshDate),
                    }
        raise AssertionError(f"pivot {pivot_table!r} not found")

    read.__name__ = "read_pivot_body"
    return dict(
        await rig.worker.submit(
            read,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )


def _sales_grand_total(pivot: Any) -> Any:
    """Return the "Total Sum of Sales" row's grand-total cell (STA-local).

    ``ptA`` stacks two data fields vertically, so the bottom-right body cell is
    the *average* total. The Sales total is the last column of the
    ``Total Sum of Sales`` row; ``None`` when the label is absent.
    """

    table = pivot.TableRange1
    rows = int(table.Rows.Count)
    cols = int(table.Columns.Count)
    for row in range(1, rows + 1):
        label = table.Cells(row, 1).Value
        if isinstance(label, str) and label.startswith("Total Sum of Sales"):
            return table.Cells(row, cols).Value
    return None


async def _read_cache_refresh_date(rig: Rig, pivot_table: str) -> datetime:
    """Read the owning cache's ``RefreshDate`` through a PivotTable name.

    The cache index is deliberately *not* used: Excel renumbers
    ``PivotCaches`` after a refresh (the refreshed cache moves to the end of the
    collection), so post-refresh reads must resolve the cache by PivotTable name.
    """

    def read() -> str:
        app = rig.app.app
        assert app is not None
        target = pivot_table.strip().lower()
        for workbook in app.Workbooks:
            for sheet in workbook.Worksheets:
                try:
                    pivots = sheet.PivotTables()
                    count = int(pivots.Count)
                except Exception:  # noqa: BLE001 - sheet without PivotTables
                    continue
                for index in range(1, count + 1):
                    pivot = pivots.Item(index)
                    if str(pivot.Name).lower() == target:
                        return str(pivot.PivotCache().RefreshDate)
        raise AssertionError(f"pivot {pivot_table!r} not found")

    read.__name__ = "read_cache_refresh_date"
    payload = await rig.worker.submit(
        read,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    return datetime.fromisoformat(payload)


async def _replace_source_sales(rig: Rig, value: float) -> float:
    """Write ``Data!D2`` (the first Sales value) and return the previous value."""

    def write() -> float:
        app = rig.app.app
        assert app is not None
        workbook = app.Workbooks(1)
        sheet = workbook.Worksheets("Data")
        previous = float(sheet.Range("D2").Value)
        sheet.Range("D2").Value = value
        return previous

    write.__name__ = "replace_source_sales"
    return float(
        await rig.worker.submit(
            write,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    )


async def _build_consolidation_pivot(rig: Rig) -> dict[str, Any]:
    """Add a consolidation-sourced PivotTable to the open working copy (STA)."""

    def build() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        workbook = app.Workbooks(1)
        sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = "Unsupported"
        cache = workbook.PivotCaches().Create(3, ("Data!R1C1:R7C2",))
        pivot = cache.CreatePivotTable(
            TableDestination=sheet.Range("B3"), TableName="ptConsolidated"
        )
        return {
            "name": str(pivot.Name),
            "cache_index": int(pivot.CacheIndex),
            "source_type": int(pivot.PivotCache().SourceType),
        }

    build.__name__ = "build_consolidation_pivot"
    return dict(
        await rig.worker.submit(
            build,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    )


# ---------------------------------------------------------------------------
# 1. RefreshTable: truth signal + advanced RefreshDate + result range
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_table_advances_refresh_date_and_keeps_result_range(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    before = await rig.pivots.get_pivot_table(wid, "ptA")
    body_before = await _read_pivot_body(rig, "ptA")

    result = await rig.pivots.refresh_pivot_table(wid, "ptA")

    assert isinstance(result, PivotRefreshResult)
    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.refreshed is True
    assert result.refresh_return is True
    assert result.cache_index == 1
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]
    assert result.duration_ms is not None
    assert result.duration_ms >= 0

    # The success signal is Excel's own RefreshDate, not "no exception".
    assert result.refresh_date_before is not None
    assert result.refresh_date_after is not None
    assert result.refresh_date_after > result.refresh_date_before

    after = await rig.pivots.get_pivot_table(wid, "ptA")
    assert before.refresh_date is not None
    assert after.refresh_date is not None
    assert after.refresh_date > before.refresh_date

    # The result range keeps its address/shape and its values (no source change).
    body_after = await _read_pivot_body(rig, "ptA")
    assert body_after["address"] == body_before["address"] == "$C$5:$F$14"
    assert (body_after["rows"], body_after["cols"]) == (body_before["rows"], body_before["cols"])
    assert body_after["grand_total"] == pytest.approx(body_before["grand_total"])
    assert body_after["sales_total"] == pytest.approx(body_before["sales_total"])

    # A sibling table on the same cache keeps its own (unchanged) result range.
    sibling = await _read_pivot_body(rig, "ptB")
    assert sibling["address"] == "$K$5:$K$8"
    assert sibling["grand_total"] == pytest.approx(24.0)

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 2. RefreshTable re-reads changed source data into the result range
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_reflects_new_source_values_in_the_result_range(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    body_before = await _read_pivot_body(rig, "ptA")
    assert body_before["sales_total"] == pytest.approx(1350.0)
    previous = await _replace_source_sales(rig, 200.0)
    assert previous == pytest.approx(100.0)

    # Without a refresh the pivot output is stale.
    stale = await _read_pivot_body(rig, "ptA")
    assert stale["sales_total"] == pytest.approx(body_before["sales_total"])

    result = await rig.pivots.refresh_pivot_table(wid, "ptA")
    assert result.status is PivotRefreshStatus.REFRESHED
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]

    body_after = await _read_pivot_body(rig, "ptA")
    assert body_after["address"] == body_before["address"]
    assert (body_after["rows"], body_after["cols"]) == (body_before["rows"], body_before["cols"])
    # The North/Alpha Sales row grew by 100, so the Sales grand total grows by 100.
    assert body_after["sales_total"] == pytest.approx(body_before["sales_total"] + 100.0)
    # The average data field is unaffected, so the bottom-right body cell is not.
    assert body_after["grand_total"] == pytest.approx(body_before["grand_total"])

    # The sibling table's own data field (Sum of Units) is unaffected.
    sibling = await _read_pivot_body(rig, "ptB")
    assert sibling["grand_total"] == pytest.approx(24.0)

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 3. RefreshCache: shared cache reports both tables and advances RefreshDate
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_cache_refreshes_shared_cache_and_advances_date(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    cache_before = await rig.pivots.get_pivot_cache(wid, 1)
    refresh_date_before = await _read_cache_refresh_date(rig, "ptA")

    result = await rig.pivots.refresh_pivot_cache(wid, 1)

    assert result.method is PivotRefreshMethod.REFRESH_CACHE
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.refreshed is True
    assert result.refresh_return is None  # PivotCache.Refresh returns no Boolean
    assert result.cache_index == 1
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]

    # Read the refreshed cache back by PivotTable name: Excel renumbers the
    # PivotCaches collection after a refresh, so index 1 no longer identifies
    # the same cache.
    refresh_date_after = await _read_cache_refresh_date(rig, "ptA")
    assert cache_before.refresh_date is not None
    assert refresh_date_after > refresh_date_before
    # The refreshed cache really is the one ptA/ptB share (not ptC's cache).
    assert await _read_cache_refresh_date(rig, "ptB") == refresh_date_after

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 4. RefreshTable is targeted: refreshing ptC leaves cache 1 untouched
# ---------------------------------------------------------------------------
@requires_excel
async def test_refresh_listobject_cache_is_targeted(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    pta_before = await rig.pivots.get_pivot_table(wid, "ptA")
    ptb_before = await rig.pivots.get_pivot_table(wid, "ptB")
    ptc_before = await rig.pivots.get_pivot_table(wid, "ptC")

    result = await rig.pivots.refresh_pivot_table(wid, "ptC")

    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.cache_index == 2
    assert result.affected_pivot_tables == ["ptC"]

    ptc_after = await rig.pivots.get_pivot_table(wid, "ptC")
    assert ptc_before.refresh_date is not None
    assert ptc_after.refresh_date is not None
    assert ptc_after.refresh_date > ptc_before.refresh_date

    # Cache 1 (ptA/ptB) was not refreshed by the targeted ptC call.
    pta_after = await rig.pivots.get_pivot_table(wid, "ptA")
    ptb_after = await rig.pivots.get_pivot_table(wid, "ptB")
    assert pta_after.refresh_date == pta_before.refresh_date
    assert ptb_after.refresh_date == ptb_before.refresh_date

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 5. Public tool layer reports the same refresh result
# ---------------------------------------------------------------------------
@requires_excel
async def test_api_tools_refresh_reports_the_same_result(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        # Cache 1 is refreshed first: a refresh renumbers the PivotCaches
        # collection, so the table-scoped call below must resolve ptC by name.
        cache_result = await pivot_tools.pivot_cache_refresh(wid, 1)
        table_result = await pivot_tools.pivot_table_refresh(wid, "ptC")
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert isinstance(cache_result, PivotRefreshResult)
    assert cache_result.status is PivotRefreshStatus.REFRESHED
    assert cache_result.method is PivotRefreshMethod.REFRESH_CACHE
    assert cache_result.cache_index == 1
    assert sorted(cache_result.affected_pivot_tables) == ["ptA", "ptB"]

    assert isinstance(table_result, PivotRefreshResult)
    assert table_result.status is PivotRefreshStatus.REFRESHED
    assert table_result.method is PivotRefreshMethod.REFRESH_TABLE
    assert table_result.affected_pivot_tables == ["ptC"]

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 6. Unsupported (consolidation) cache: cache-scoped refresh is refused
# ---------------------------------------------------------------------------
@requires_excel
async def test_unsupported_cache_refresh_is_reported(rig: Rig) -> None:
    info = await rig.open_for_refresh()
    wid = info.workbook_id
    built = await _build_consolidation_pivot(rig)
    assert built["source_type"] == 3  # xlSourceTypeConsolidation

    cache = await rig.pivots.get_pivot_cache(wid, built["cache_index"])
    assert cache.source_type is PivotSourceType.CONSOLIDATION

    await rig.snapshot.create_snapshot(wid)
    result = await rig.pivots.refresh_pivot_cache(wid, built["cache_index"])

    assert result.status is PivotRefreshStatus.UNSUPPORTED_SOURCE
    assert result.refreshed is False
    assert result.refresh_return is None
    assert result.error is not None
    assert result.error["code"] == "UNSUPPORTED_PIVOT_SOURCE"
    assert result.refresh_date_before == result.refresh_date_after

    await rig.session.close_workbook(wid)
