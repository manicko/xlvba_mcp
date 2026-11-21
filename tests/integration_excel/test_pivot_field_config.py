"""Real Excel integration tests for PivotField configuration (task_050).

These tests drive the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
against a **hidden, automation-created Excel instance** (never the user's Excel)
and a disposable working copy of the committed ``PivotTablesProject.xlsm``
fixture. They assert:

* **Orientation / layout** — a field is moved between the row / column / page
  buckets and the change is confirmed by a re-read of the PivotTable;
  ``orientation="data"`` is rejected because assigning ``xlDataField`` never
  creates a data field (that is ``AddDataField``);
* **Position** — a 1-based position is applied within its orientation and the
  resulting order is read back; an out-of-range position is rejected;
* **Aggregation** — an existing data field's ``Function`` is updated
  (sum / average / count) and read back through ``DataFields``;
* **Caption / properties** — a unique data-field caption and number format are
  written and read back; a duplicate caption is rejected;
* **Removal** — ``orientation="hidden"`` removes a field non-destructively (it
  is still listed, as hidden);
* **Rejections** — an invalid orientation / position / legacy aggregation code
  is rejected with a structured error;
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

    uv run pytest tests/integration_excel/test_pivot_field_config.py -v -s
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
from excel_vba_mcp.domain.models import PivotFieldInfo, PivotFieldProperties
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
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

# Approval actions for the four PivotField configuration tools.
_ACTION_LAYOUT = "pivot_field_set_layout"
_ACTION_POSITION = "pivot_field_set_position"
_ACTION_AGGREGATION = "pivot_field_set_aggregation"
_ACTION_PROPERTIES = "pivot_field_set_properties"

#: PivotTable with a single data field: every orientation transition is accepted
#: (a second data field inserts a synthetic "Data" placeholder that locks the row
#: axis on the target build).
_SINGLE_DATA_PIVOT = "ptB"

#: PivotTable with two data fields (page/row/column layout, shared cache 1).
_MULTI_DATA_PIVOT = "ptA"


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
# Fixtures / rig — real COM worker + Excel + PivotTablesService (config-wired)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService wired for guarded config."""

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

    async def open_for_config(self) -> Any:
        """Open writable and grant all four PivotField config actions."""

        info = await self.open(read_only=False)
        for action in (_ACTION_LAYOUT, _ACTION_POSITION, _ACTION_AGGREGATION, _ACTION_PROPERTIES):
            self.grant(action)
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


def _by_name(fields: list[PivotFieldInfo]) -> dict[str, PivotFieldInfo]:
    return {field.source_name: field for field in fields}


# ---------------------------------------------------------------------------
# Orientation / layout
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_orientation_row_column_page(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    # ptB has a single data field (no synthetic "Data" placeholder locking the
    # row axis), so every orientation transition is accepted by Excel.
    moved = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, _SINGLE_DATA_PIVOT, "Region", PivotFieldOrientation.COLUMN
    )
    assert moved.orientation is PivotFieldOrientation.COLUMN
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _SINGLE_DATA_PIVOT))
    assert fields["Region"].orientation is PivotFieldOrientation.COLUMN

    # ...then to the page axis.
    moved = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, _SINGLE_DATA_PIVOT, "Region", PivotFieldOrientation.PAGE
    )
    assert moved.orientation is PivotFieldOrientation.PAGE
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _SINGLE_DATA_PIVOT))
    assert fields["Region"].orientation is PivotFieldOrientation.PAGE

    # ...and back to the row axis.
    moved = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, _SINGLE_DATA_PIVOT, "Region", PivotFieldOrientation.ROW
    )
    assert moved.orientation is PivotFieldOrientation.ROW
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _SINGLE_DATA_PIVOT))
    assert fields["Region"].orientation is PivotFieldOrientation.ROW

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_locked_axis_transition_reports_truthful_error(rig: Rig) -> None:
    """A transition Excel refuses is reported, never claimed as success.

    On ptA the row axis holds the synthetic "Data" placeholder (two data
    fields), so Excel rejects moving a row field into the already-occupied
    column axis (``Unable to set the Orientation property``). The adapter must
    surface that as a structured ``COM_BUSY`` error and leave the field as it
    was — it must not report a false success.
    """

    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, _MULTI_DATA_PIVOT, "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.COM_BUSY
    # The field is still on the row axis.
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _MULTI_DATA_PIVOT))
    assert fields["Region"].orientation is PivotFieldOrientation.ROW

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_set_data_orientation_rejected(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Units", PivotFieldOrientation.DATA
        )

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_orientation"
    # No data field was created by the rejected request.
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert fields["Units"].orientation is PivotFieldOrientation.DATA  # still a data field
    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_remove_field_non_destructive(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    removed = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, "ptA", "Product", PivotFieldOrientation.HIDDEN
    )
    assert removed.orientation is PivotFieldOrientation.HIDDEN
    assert removed.position is None

    # Non-destructive: the field is still listed (as hidden).
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert "Product" in fields
    assert fields["Product"].orientation is PivotFieldOrientation.HIDDEN
    assert fields["Product"].position is None

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Position
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_position_within_orientation(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    # ptB has one row field; move Product to the row axis so the row orientation
    # has two fields (and no synthetic "Data" placeholder).
    await rig.pivots.set_pivot_field_layout(
        info.workbook_id, _SINGLE_DATA_PIVOT, "Product", PivotFieldOrientation.ROW
    )
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _SINGLE_DATA_PIVOT))
    assert fields["Product"].position == 2

    moved = await rig.pivots.set_pivot_field_position(
        info.workbook_id, _SINGLE_DATA_PIVOT, "Product", 1
    )
    assert moved.position == 1

    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _SINGLE_DATA_PIVOT))
    assert fields["Product"].position == 1
    assert fields["Region"].position == 2

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_set_data_field_position(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    before = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert before["Sales"].position == 1
    assert before["Units"].position == 2

    moved = await rig.pivots.set_pivot_field_position(info.workbook_id, "ptA", "Units", 1)
    assert moved.position == 1

    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert fields["Units"].position == 1
    assert fields["Sales"].position == 2

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_set_position_out_of_range_rejected(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_position(info.workbook_id, "ptA", "Region", 99)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "position_out_of_range"
    # The field keeps its original position.
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert fields["Region"].position == 1

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Aggregation
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_aggregation_sum_average_count(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    before = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _MULTI_DATA_PIVOT))
    assert before["Sales"].function in (PivotAggregation.SUM, PivotAggregation.AVERAGE)

    # SUM
    summed = await rig.pivots.set_pivot_field_aggregation(
        info.workbook_id, _MULTI_DATA_PIVOT, "Sales", PivotAggregation.SUM
    )
    assert summed.function is PivotAggregation.SUM

    # AVERAGE
    updated = await rig.pivots.set_pivot_field_aggregation(
        info.workbook_id, _MULTI_DATA_PIVOT, "Sales", PivotAggregation.AVERAGE
    )
    assert updated.function is PivotAggregation.AVERAGE

    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _MULTI_DATA_PIVOT))
    assert fields["Sales"].function is PivotAggregation.AVERAGE

    # COUNT
    updated = await rig.pivots.set_pivot_field_aggregation(
        info.workbook_id, _MULTI_DATA_PIVOT, "Units", PivotAggregation.COUNT
    )
    assert updated.function is PivotAggregation.COUNT
    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, _MULTI_DATA_PIVOT))
    assert fields["Units"].function is PivotAggregation.COUNT

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_set_aggregation_rejects_legacy_code(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    from excel_vba_mcp.api import pivot_tools

    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await pivot_tools.pivot_field_set_aggregation(
                info.workbook_id,
                "ptA",
                "Sales",
                9,  # type: ignore[arg-type]
            )
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["reason"] == "invalid_function"

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Captions / properties
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_caption_and_number_format(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    updated = await rig.pivots.set_pivot_field_properties(
        info.workbook_id,
        "ptA",
        "Sales",
        PivotFieldProperties(caption="Total Sales", number_format="0.00"),
    )
    assert updated.data_field_name == "Total Sales"
    assert updated.number_format == "0.00"

    fields = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert fields["Sales"].data_field_name == "Total Sales"
    assert fields["Sales"].number_format == "0.00"

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_duplicate_caption_rejected(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    existing = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    taken = existing["Units"].data_field_name
    assert taken  # e.g. "Average of Units"

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_properties(
            info.workbook_id, "ptA", "Sales", PivotFieldProperties(caption=taken)
        )

    assert exc_info.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert exc_info.value.context["reason"] == "duplicate_caption"

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Guard chain
# ---------------------------------------------------------------------------
@requires_excel
async def test_config_rejected_on_read_only_open(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    rig.grant(_ACTION_LAYOUT)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_config_rejected_without_snapshot(rig: Rig) -> None:
    info = await rig.open_for_config()
    # No snapshot created: the guard must block before any mutation.

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Region", PivotFieldOrientation.COLUMN
        )

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_config_rejected_without_approval(rig: Rig) -> None:
    # Open writable but do NOT grant the config actions.
    rig.grant("excel_open")
    info = await rig.session.open_workbook(
        str(rig.project), mode=SessionMode.INSPECT, read_only=False
    )
    await rig.snapshot.create_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.pivots.set_pivot_field_aggregation(
            info.workbook_id, "ptA", "Sales", PivotAggregation.MAX
        )

    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 + original unchanged + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_hidden_owned_excel_during_config(rig: Rig) -> None:
    info = await rig.open_for_config()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    # Run a mutating field change (the path under test) while hidden.
    result = await rig.pivots.set_pivot_field_layout(
        info.workbook_id, "ptA", "Quarter", PivotFieldOrientation.ROW
    )
    assert result.orientation is PivotFieldOrientation.ROW

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
    _assert_owned_never_foreground(owned_pid, "pivot-field-config")
    _assert_no_visible_owned_windows(owned_pid, "pivot-field-config")

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
        info = await r.open_for_config()
        await r.snapshot.create_snapshot(info.workbook_id)
        await r.pivots.set_pivot_field_layout(
            info.workbook_id, "ptA", "Quarter", PivotFieldOrientation.ROW
        )
        await r.pivots.set_pivot_field_aggregation(
            info.workbook_id, "ptA", "Sales", PivotAggregation.MAX
        )
        await r.pivots.set_pivot_field_properties(
            info.workbook_id, "ptA", "Sales", PivotFieldProperties(caption="Total Sales")
        )
        await r.session.close_workbook(info.workbook_id)
    finally:
        await r.close()

    assert _sha256(pivot_project) == before


@requires_excel
async def test_owned_excel_reaped_after_shutdown() -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    tmp_path = Path(tempfile.mkdtemp(prefix="pivot_field_"))
    try:
        original = tmp_path / pivot_project.name
        shutil.copy2(pivot_project, original)
        r = Rig(tmp_path, original)
        await r.start()
        try:
            info = await r.open_for_config()
            await r.snapshot.create_snapshot(info.workbook_id)
            await r.pivots.set_pivot_field_position(info.workbook_id, "ptA", "Region", 1)
            owned_pid = r.worker.get_status().owned_pid
            assert owned_pid is not None
        finally:
            await r.close()
    finally:
        shutil.rmtree(tmp_path, ignore_errors=True)

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
