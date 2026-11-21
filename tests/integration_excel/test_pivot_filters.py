"""Real Excel integration tests for PivotField filters (task_051).

These tests drive the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
against a **hidden, automation-created Excel instance** (never the user's Excel)
and a disposable working copy of the committed ``PivotTablesProject.xlsm``
fixture. They assert:

* **Read** — ``ptB``'s Region manual filter (West hidden) is reported with its
  visible/hidden items, while ``ptA``'s unfiltered Region reports all items
  visible;
* **Manual apply** — ``PivotItems[i].Visible`` hides the unlisted items and the
  change is confirmed by a re-read of the PivotTable; an unknown item is
  rejected and leaves the field untouched;
* **Unsupported filters** — a label filter returns an explicit ``unsupported``
  result, leaves every item visible and never creates a ``PivotFilters`` entry
  (``PivotFilters.Count == 0``);
* **Clear scope** — ``scope="field"`` clears one field only; ``scope="table"``
  clears the whole PivotTable and leaves sibling PivotTables (even on a shared
  cache) untouched;
* **Guards** — read-only open → ``WORKBOOK_READ_ONLY``, no snapshot →
  ``SNAPSHOT_FAILED``, no approval → ``APPROVAL_DENIED``;
* **BG03/BG04/BG05** — the automation Excel stays hidden and non-activating, the
  owned PID never holds foreground, no owned window is visible, and the owned PID
  is reaped after shutdown (no orphan);
* the committed fixture is byte-identical afterwards (mutations run only on the
  disposable working copy).

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_filters.py -v -s
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
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotFieldFilters, PivotFilterSpec
from excel_vba_mcp.domain.statuses import (
    OperationClass,
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

_TEST_TIMEOUT_S = 120.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5

# Approval actions for the two PivotField filter mutations.
_ACTION_APPLY = "pivot_field_apply_filter"
_ACTION_CLEAR = "pivot_field_clear_filters"

#: PivotTable whose Region field carries a manual filter in the fixture (West
#: hidden); used to prove a sibling table on the same cache is unaffected.
_FILTERED_PIVOT = "ptB"

#: PivotTable used for the apply / clear round-trips (single row field).
_ROUNDTRIP_PIVOT = "ptA"


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
# Fixtures / rig — real COM worker + Excel + PivotTablesService (filter-wired)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService wired for guarded filters."""

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

    async def open_for_filters(self) -> Any:
        """Open writable and grant both filter mutation actions."""

        info = await self.open(read_only=False)
        self.grant(_ACTION_APPLY)
        self.grant(_ACTION_CLEAR)
        return info

    async def close(self) -> None:
        owned_pid = self.worker.get_status().owned_pid
        await self.worker.shutdown()
        self.audit.close()
        # Wait for the OS process to actually exit so a following test never
        # observes this instance as a lingering "user-owned" Excel PID.
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


async def _filters(rig: Rig, workbook_id: str, pivot_table: str, field: str) -> PivotFieldFilters:
    return await rig.pivots.get_pivot_field_filters(workbook_id, pivot_table, field)


async def _pivot_filters_count(rig: Rig, pivot_table: str, field: str) -> int:
    """Read ``PivotField.PivotFilters.Count`` directly (STA-local evidence)."""

    def read_count() -> int:
        app = rig.app.app
        assert app is not None
        target = pivot_table.strip().lower()
        wanted = field.strip().lower()
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
                    fields = pivot.PivotFields()
                    for field_index in range(1, int(fields.Count) + 1):
                        candidate = fields.Item(field_index)
                        if str(candidate.SourceName).lower() == wanted:
                            return int(candidate.PivotFilters.Count)
        return -1

    read_count.__name__ = "read_pivot_filters_count"
    return int(
        await rig.worker.submit(
            read_count,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )


# ---------------------------------------------------------------------------
# Read
# ---------------------------------------------------------------------------
@requires_excel
async def test_get_filters_reads_manual_state(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    ptb = await _filters(rig, info.workbook_id, _FILTERED_PIVOT, "Region")
    assert ptb.all_items_visible is False
    assert ptb.filter.kind is PivotFilterKind.MANUAL
    assert set(ptb.visible_items) == {"East", "North", "South"}
    assert ptb.hidden_items == ["West"]
    assert ptb.unsupported == []

    pta = await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Region")
    assert pta.all_items_visible is True
    assert pta.filter.kind is PivotFilterKind.NONE
    assert pta.hidden_items == []

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_get_filters_rejects_unknown_field(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Nope")

    assert exc_info.value.context["reason"] == "field_not_found"

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Manual apply
# ---------------------------------------------------------------------------
@requires_excel
async def test_apply_manual_filter_and_readback(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    applied = await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North", "South")
    )

    assert applied.filter.kind is PivotFilterKind.MANUAL
    assert set(applied.visible_items) == {"North", "South"}
    assert set(applied.hidden_items) == {"East", "West"}
    assert applied.all_items_visible is False

    # Readback through a fresh read agrees with the postcondition.
    readback = await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Region")
    assert set(readback.visible_items) == {"North", "South"}
    assert set(readback.hidden_items) == {"East", "West"}
    assert readback.filter.kind is PivotFilterKind.MANUAL

    # A sibling PivotTable on the same shared cache keeps its own filter state.
    ptb = await _filters(rig, info.workbook_id, _FILTERED_PIVOT, "Region")
    assert set(ptb.visible_items) == {"East", "North", "South"}
    assert ptb.hidden_items == ["West"]

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_apply_manual_filter_rejects_unknown_item(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North", "Nowhere")
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "unknown_filter_item"

    # The field is untouched: every item is still visible.
    after = await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Region")
    assert after.hidden_items == []

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_apply_unsupported_label_filter_is_explicit(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    result = await rig.pivots.apply_pivot_field_filter(
        info.workbook_id,
        _ROUNDTRIP_PIVOT,
        "Region",
        PivotFilterSpec(kind=PivotFilterKind.LABEL),
    )

    assert result.filter.supported is False
    assert result.filter.kind is PivotFilterKind.LABEL
    assert result.unsupported == ["label"]
    assert "PivotFilters.Add is unsupported" in (result.filter.unsupported_reason or "")

    # No manual filter was silently applied and no PivotFilters entry was added.
    after = await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Region")
    assert after.hidden_items == []
    assert after.filter.kind is PivotFilterKind.NONE
    assert after.unsupported == []
    assert await _pivot_filters_count(rig, _ROUNDTRIP_PIVOT, "Region") == 0

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Clear
# ---------------------------------------------------------------------------
@requires_excel
async def test_clear_field_scope_only_clears_one_field(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
    )

    cleared = await rig.pivots.clear_pivot_field_filters(
        info.workbook_id, _ROUNDTRIP_PIVOT, "Region", scope="field"
    )

    assert cleared.field == "Region"
    assert cleared.hidden_items == []
    assert cleared.all_items_visible is True

    # The sibling table keeps its fixture filter (field scope is one field only).
    ptb = await _filters(rig, info.workbook_id, _FILTERED_PIVOT, "Region")
    assert ptb.hidden_items == ["West"]

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_clear_table_scope_only_clears_that_pivot_table(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    # Filter both ptA and ptB, then clear only ptB's table.
    await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
    )
    await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, _FILTERED_PIVOT, "Region", _manual("North")
    )

    cleared = await rig.pivots.clear_pivot_field_filters(
        info.workbook_id, _FILTERED_PIVOT, scope="table"
    )

    assert cleared.field == ""
    assert cleared.hidden_items == []
    ptb = await _filters(rig, info.workbook_id, _FILTERED_PIVOT, "Region")
    assert ptb.hidden_items == []
    assert ptb.all_items_visible is True

    # ptA shares cache 1 with ptB but keeps its own manual filter.
    pta = await _filters(rig, info.workbook_id, _ROUNDTRIP_PIVOT, "Region")
    assert pta.hidden_items == ["East", "South", "West"]

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_clear_field_scope_requires_field(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.clear_pivot_field_filters(
            info.workbook_id, _ROUNDTRIP_PIVOT, None, scope="field"
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "field_required"

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_clear_rejects_invalid_scope(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.clear_pivot_field_filters(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", scope="everything"
        )

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_scope"

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Guard chain
# ---------------------------------------------------------------------------
@requires_excel
async def test_filters_rejected_on_read_only_open(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    rig.grant(_ACTION_APPLY)
    rig.grant(_ACTION_CLEAR)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
        )

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_filters_rejected_without_snapshot(rig: Rig) -> None:
    info = await rig.open_for_filters()
    # No snapshot created: the guard must block before any mutation.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.clear_pivot_field_filters(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", scope="field"
        )

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_filters_rejected_without_approval(rig: Rig) -> None:
    # Open writable but do NOT grant the filter actions.
    rig.grant("excel_open")
    info = await rig.session.open_workbook(
        str(rig.project), mode=SessionMode.INSPECT, read_only=False
    )
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.apply_pivot_field_filter(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
        )

    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 + original unchanged + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_hidden_owned_excel_during_filter(rig: Rig) -> None:
    info = await rig.open_for_filters()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    # Run a mutating filter change (the path under test) while hidden.
    applied = await rig.pivots.apply_pivot_field_filter(
        info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
    )
    assert applied.hidden_items == ["East", "South", "West"]

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
    _assert_owned_never_foreground(owned_pid, "pivot-filters")
    _assert_no_visible_owned_windows(owned_pid, "pivot-filters")

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_committed_fixture_unchanged(tmp_path: Path) -> None:
    """BG05: the committed fixture is byte-identical after the mutations run."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    before = _sha256(pivot_project)
    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        info = await r.open_for_filters()
        await r.snapshot.create_snapshot(info.workbook_id)
        await r.pivots.apply_pivot_field_filter(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
        )
        await r.pivots.apply_pivot_field_filter(
            info.workbook_id, _FILTERED_PIVOT, "Region", _manual("North")
        )
        await r.pivots.clear_pivot_field_filters(
            info.workbook_id, _ROUNDTRIP_PIVOT, "Region", scope="field"
        )
        await r.pivots.clear_pivot_field_filters(info.workbook_id, _FILTERED_PIVOT, scope="table")
        await r.session.close_workbook(info.workbook_id)
    finally:
        await r.close()

    assert _sha256(pivot_project) == before


@requires_excel
async def test_owned_excel_reaped_after_shutdown() -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    tmp_path = Path(tempfile.mkdtemp(prefix="pivot_filters_"))
    try:
        original = tmp_path / pivot_project.name
        shutil.copy2(pivot_project, original)
        r = Rig(tmp_path, original)
        await r.start()
        try:
            info = await r.open_for_filters()
            await r.snapshot.create_snapshot(info.workbook_id)
            await r.pivots.apply_pivot_field_filter(
                info.workbook_id, _ROUNDTRIP_PIVOT, "Region", _manual("North")
            )
            owned_pid = r.worker.get_status().owned_pid
            assert owned_pid is not None
        finally:
            await r.close()
    finally:
        shutil.rmtree(tmp_path, ignore_errors=True)

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
