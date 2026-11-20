"""Real Excel integration tests for the structural Range/Cells surface (task_045).

These tests drive the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and a disposable copy of the committed ``RangesCellsProject.xlsm``
fixture (task_041). They assert:

* ``range_copy`` copies to an explicit destination (same sheet and cross-sheet)
  with a value/dimension postcondition and never touches the clipboard;
* ``range_insert`` / ``range_delete`` apply an explicit ``right`` / ``down`` /
  ``left`` / ``up`` shift and verify the resulting geometry;
* ``range_merge`` blocks non-anchor data until ``allow_data_loss`` is passed
  after the preview, then merges (also ``across=True``); ``range_unmerge``
  restores the merged area;
* insert/delete/copy targets that intersect the fixture's ``tblCells`` ListObject
  are rejected with ``LISTOBJECT_INTERSECTION``;
* an ambiguous shift and a protected sheet are rejected;
* the snapshot / approval / read-only gates reject a structural mutation;
* BG03 / BG04 / BG05 hold (the owned Excel is hidden + non-activating, the source
  fixture is unchanged, and the worker refuses work on a hung instance).

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.
"""

from __future__ import annotations

import asyncio
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
from excel_vba_mcp.domain.statuses import (
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

pytest.importorskip("win32com", reason="Excel COM automation unavailable")

_RANGES_PROJECT = (
    Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "RangesCellsProject.xlsm"
)

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_TEST_TIMEOUT_S = 120.0


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


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _pid_alive(pid: int) -> bool:
    """Return whether ``pid`` still exists (Windows ``tasklist`` probe)."""

    probe = subprocess.run(  # noqa: S603 - fixed system tool, no user input
        ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
        capture_output=True,
        text=True,
        check=False,
    )
    return str(pid) in probe.stdout


def _ensure_process_gone(pid: int, timeout_s: float = 15.0) -> None:
    """Wait for ``pid`` to exit; force-kill only our own leftover test instance."""

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return
        time.sleep(0.1)
    subprocess.run(  # noqa: S603 - fixed system tool, no user input
        ["taskkill", "/F", "/PID", str(pid)],
        capture_output=True,
        check=False,
    )


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for range writes."""

    def __init__(self, tmp_path: Path, project: Path) -> None:
        self.tmp_path = tmp_path
        self.project = project
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        canonical = self.path_policy.validate_path(str(project), OperationClass.MUTATING)
        self.canonical = str(canonical)
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter)
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

    def working_path(self, workbook_id: str) -> str:
        return self.session._records[workbook_id].working_path

    async def close(self) -> None:
        owned_pid = getattr(self.worker, "_owned_pid", None)
        await self.worker.shutdown()
        self.audit.close()
        # Excel can take a moment to finish exiting after ``Quit``; wait for the
        # owned process to disappear so this module never leaves a half-dead
        # EXCEL.exe behind (which would trip later suites' orphan assertions).
        if owned_pid:
            await asyncio.to_thread(_ensure_process_gone, int(owned_pid))


@pytest.fixture
async def rig(tmp_path: Path, ranges_project: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every test mutates a disposable working copy and never the source.
    original = tmp_path / ranges_project.name
    shutil.copy2(ranges_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


async def _read_cell(rig: Rig, workbook_id: str, sheet: str, address: str) -> Any:
    """Read one cell's value/formula/merge state through the STA worker."""

    working_path = rig.working_path(workbook_id)

    def probe() -> dict[str, Any]:
        wb = _open_workbook(rig, working_path)
        cell = wb.Worksheets(sheet).Range(address)
        return {
            "value": cell.Value2,
            "formula": cell.Formula,
            "merge_cells": bool(cell.MergeCells),
            "merge_area": str(cell.MergeArea.Address),
        }

    probe.__name__ = "read_cell"
    return await rig.worker.submit(
        probe,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )


async def _protect_sheet(rig: Rig, workbook_id: str, sheet: str) -> None:
    """Protect ``sheet`` on the disposable working copy (test setup)."""

    working_path = rig.working_path(workbook_id)

    def protect() -> None:
        wb = _open_workbook(rig, working_path)
        wb.Worksheets(sheet).Protect()

    protect.__name__ = "protect_sheet"
    await rig.worker.submit(
        protect,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.MUTATING,
    )


def _number(cell: Any) -> float:
    """Extract a numeric value from a ``_read_cell`` payload or a ``CellValue``."""

    value = cell["value"] if isinstance(cell, dict) else cell.value
    return float(value)


# ---------------------------------------------------------------------------
# Copy
# ---------------------------------------------------------------------------
async def test_copy_range_same_sheet(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_copy")

    result = await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")

    assert result.verified is True
    assert result.range_after is not None
    assert result.range_after.address == "$D$2:$E$4"
    assert result.table_metadata_unchanged is True

    post = await rig.ranges.read_values(wid, "StructOps", "D2:E4")
    matrix = [[_number(cell) for cell in row] for row in post.values]
    assert matrix == [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]

    # BG03: the source fixture is never modified.
    assert _sha256(rig.project) == before


async def test_copy_range_cross_sheet(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_copy")

    result = await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "Cells", "X1")

    assert result.verified is True
    assert result.source is not None and result.source.sheet == "StructOps"
    assert result.range_after is not None and result.range_after.sheet == "Cells"

    post = await rig.ranges.read_values(wid, "Cells", "X1:Y3")
    matrix = [[_number(cell) for cell in row] for row in post.values]
    assert matrix == [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]


async def test_copy_requires_explicit_destination(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_copy")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", " ")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Insert / delete
# ---------------------------------------------------------------------------
async def test_insert_shift_right(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")

    result = await rig.ranges.insert_range(wid, "StructOps", "B2", "right")

    assert result.verified is True
    assert result.shift == "right"
    assert result.range_after is not None and result.range_after.address == "$B$2"
    assert (await _read_cell(rig, wid, "StructOps", "B2"))["value"] is None
    assert _number(await _read_cell(rig, wid, "StructOps", "C2")) == 2.0


async def test_insert_shift_down(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")

    result = await rig.ranges.insert_range(wid, "StructOps", "B2", "down")

    assert result.verified is True
    assert result.shift == "down"
    assert (await _read_cell(rig, wid, "StructOps", "B2"))["value"] is None
    assert _number(await _read_cell(rig, wid, "StructOps", "B3")) == 2.0


async def test_delete_shift_up(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_delete")

    result = await rig.ranges.delete_range(wid, "StructOps", "B2", "up")

    assert result.verified is True
    assert result.shift == "up"
    # The column shifted up: B2 takes the old B3 value and B3 the old B4 value.
    assert _number(await _read_cell(rig, wid, "StructOps", "B2")) == 4.0
    assert _number(await _read_cell(rig, wid, "StructOps", "B3")) == 6.0


async def test_delete_shift_left(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_delete")

    result = await rig.ranges.delete_range(wid, "StructOps", "B2", "left")

    assert result.verified is True
    assert result.shift == "left"
    # A2 is untouched and the deleted B2 did not reappear.
    assert _number(await _read_cell(rig, wid, "StructOps", "A2")) == 1.0
    assert (await _read_cell(rig, wid, "StructOps", "B2"))["value"] is None


async def test_ambiguous_shift_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")
    rig.grant("range_delete")

    for shift in ("", "   "):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.ranges.insert_range(wid, "StructOps", "B2", shift)
        assert exc_info.value.code is ErrorCode.INVALID_RANGE

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.insert_range(wid, "StructOps", "B2", "sideways")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.delete_range(wid, "StructOps", "B2", "right")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    # No mutation: B2 still holds its original value.
    assert _number(await _read_cell(rig, wid, "StructOps", "B2")) == 2.0


# ---------------------------------------------------------------------------
# Merge / unmerge
# ---------------------------------------------------------------------------
async def test_merge_blocks_non_anchor_data_then_allows(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_merge")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.merge_range(wid, "StructOps", "J2:L2")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells_with_data"] == ["K2", "L2"]
    # No merge happened.
    assert (await _read_cell(rig, wid, "StructOps", "K2"))["merge_cells"] is False

    result = await rig.ranges.merge_range(wid, "StructOps", "J2:L2", allow_data_loss=True)

    assert result.verified is True
    assert result.affected_addresses == ["K2", "L2"]
    assert result.notes and "allow_data_loss" in result.notes[0]

    anchor = await _read_cell(rig, wid, "StructOps", "J2")
    assert anchor["value"] == "anchor"
    assert anchor["merge_cells"] is True
    assert anchor["merge_area"] == "$J$2:$L$2"
    # The discarded non-anchor values are gone.
    assert (await _read_cell(rig, wid, "StructOps", "K2"))["value"] is None


async def test_merge_across_keeps_row_anchors(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_merge")

    result = await rig.ranges.merge_range(
        wid, "StructOps", "J2:L3", across=True, allow_data_loss=True
    )

    assert result.verified is True
    assert result.preconditions["across"] is True
    assert (await _read_cell(rig, wid, "StructOps", "J2"))["merge_area"] == "$J$2:$L$2"
    assert (await _read_cell(rig, wid, "StructOps", "J3"))["merge_area"] == "$J$3:$L$3"


async def test_unmerge_merged_area(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_unmerge")

    result = await rig.ranges.unmerge_range(wid, "StructOps", "J5:L5")

    assert result.verified is True
    assert result.affected_addresses == ["$J$5:$L$5"]

    anchor = await _read_cell(rig, wid, "StructOps", "J5")
    assert anchor["value"] == "unmerge me"
    assert anchor["merge_cells"] is False
    assert anchor["merge_area"] == "$J$5"


async def test_unmerge_requires_merged_target(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_unmerge")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.unmerge_range(wid, "StructOps", "A2:B4")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# ListObject hard guard
# ---------------------------------------------------------------------------
async def test_insert_intersecting_table_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")

    before = await rig.ranges.get_range(wid, "Cells", "T1:V4")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.insert_range(wid, "Cells", "T2:V4", "right")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblCells"

    after = await rig.ranges.get_range(wid, "Cells", "T1:V4")
    assert (after.rows, after.cols, after.address) == (before.rows, before.cols, before.address)


async def test_delete_intersecting_table_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_delete")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.delete_range(wid, "Cells", "U2:V3", "up")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblCells"


async def test_copy_destination_intersecting_table_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_copy")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "Cells", "T1")
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblCells"


async def test_merge_intersecting_table_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_merge")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.merge_range(wid, "Cells", "T1:V4", allow_data_loss=True)
    assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION
    assert exc_info.value.context["table"] == "tblCells"


# ---------------------------------------------------------------------------
# Protection and guard chain
# ---------------------------------------------------------------------------
async def test_protected_sheet_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")
    rig.grant("range_merge")

    await _protect_sheet(rig, wid, "StructOps")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.insert_range(wid, "StructOps", "B2", "right")
    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.merge_range(wid, "StructOps", "J2:L2", allow_data_loss=True)
    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET


async def test_structural_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    rig.grant("range_copy")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_structural_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.ranges.insert_range(wid, "StructOps", "B2", "right")


async def test_structural_requires_writable(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_delete")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.delete_range(wid, "StructOps", "B2", "up")
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_hidden_owned_excel_and_original_unchanged(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_copy")

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

    await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")

    # BG03: the source fixture is unchanged after a mutation on the working copy.
    assert _sha256(rig.project) == before


async def test_bg04_worker_refuses_write_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_insert")

    def hang() -> None:
        import time

        time.sleep(3.0)

    hang.__name__ = "hang_test"
    with pytest.raises(ExcelVbaMcpError):
        await rig.worker.submit(
            hang,
            timeout=0.1,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    # BG04: the worker is NEEDS_RECOVERY and refuses new work, so the insert never
    # reaches Excel (fail fast, no enqueue).
    with pytest.raises(ExcelVbaMcpError):
        await rig.ranges.insert_range(wid, "StructOps", "B2", "right")


async def test_bg05_session_state_running(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING
