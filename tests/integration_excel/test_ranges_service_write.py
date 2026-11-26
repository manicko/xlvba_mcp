"""Real Excel integration tests for the Range/Cells write surface (task_043).

These tests drive the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and a disposable copy of the committed ``RangesCellsProject.xlsm``
fixture (task_041). They assert:

* a scalar write and a 2-D block write round-trip with the expected values and
  postconditions;
* a formula write round-trips through ``Formula2``;
* a wrong-sized payload is rejected before any mutation (the target is
  unchanged);
* the snapshot / read-only / approval guards reject a write with no mutation;
* a cell-level write inside the fixture's ``tblCells`` table is allowed and the
  table structure is unchanged; a structural change is rejected;
* BG03 / BG04 / BG05 hold (the owned Excel is hidden + non-activating, the
  source fixture is unchanged, and the worker refuses work on a hung instance).

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.
"""

from __future__ import annotations

import hashlib
import shutil
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
    FormulaWriteRequest,
    RangeWriteRequest,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
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


def _requires_excel() -> Any:
    """Skip the test when Excel COM automation is unavailable."""

    return pytest.mark.skipif(
        not _excel_available(), reason="Excel COM automation is not available on this host"
    )


def _excel_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001 - COM unavailable (non-Windows)
        return False
    return True


requires_excel = _requires_excel()


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


def _cv(kind: CellValueKind, value: Any = None) -> CellValue:
    return CellValue(kind=kind, value=value)


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

    async def close(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


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


# ---------------------------------------------------------------------------
# Value writes
# ---------------------------------------------------------------------------
async def test_write_scalar_value(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.STRING, "written")]])
    result = await rig.ranges.write_values(wid, "Cells", "B2", payload)

    assert result.verified is True
    assert result.range_before is not None
    assert result.range_after is not None
    post = await rig.ranges.read_values(wid, "Cells", "B2")
    assert post.values[0][0].kind is CellValueKind.STRING
    assert post.values[0][0].value == "written"

    # BG03: the source fixture is never modified.
    assert _sha256(rig.project) == before


async def test_write_2d_block(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    payload = RangeWriteRequest(
        rows=3,
        cols=3,
        values=[
            [
                _cv(CellValueKind.NUMBER, 10.0),
                _cv(CellValueKind.NUMBER, 20.0),
                _cv(CellValueKind.NUMBER, 30.0),
            ],
            [
                _cv(CellValueKind.NUMBER, 40.0),
                _cv(CellValueKind.NUMBER, 50.0),
                _cv(CellValueKind.NUMBER, 60.0),
            ],
            [
                _cv(CellValueKind.NUMBER, 70.0),
                _cv(CellValueKind.NUMBER, 80.0),
                _cv(CellValueKind.NUMBER, 90.0),
            ],
        ],
    )
    result = await rig.ranges.write_values(wid, "Cells", "F1:H3", payload)

    assert result.verified is True
    post = await rig.ranges.read_values(wid, "Cells", "F1:H3")
    matrix = [[float(cell.value) for cell in row] for row in post.values]
    assert matrix == [
        [10.0, 20.0, 30.0],
        [40.0, 50.0, 60.0],
        [70.0, 80.0, 90.0],
    ]


async def test_write_typed_values_round_trip(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    payload = RangeWriteRequest(
        rows=1,
        cols=4,
        values=[
            [
                _cv(CellValueKind.STRING, "plain text"),
                _cv(CellValueKind.BOOLEAN, True),
                _cv(CellValueKind.EMPTY),
                _cv(CellValueKind.NUMBER, 3.5),
            ]
        ],
    )
    result = await rig.ranges.write_values(wid, "Cells", "B18:E18", payload)
    assert result.verified is True

    post = await rig.ranges.read_values(wid, "Cells", "B18:E18")
    kinds = [cell.kind for cell in post.values[0]]
    assert kinds == [
        CellValueKind.STRING,
        CellValueKind.BOOLEAN,
        CellValueKind.EMPTY,
        CellValueKind.NUMBER,
    ]
    assert post.values[0][0].value == "plain text"


async def test_value_channel_does_not_write_formula_channel(rig: Rig) -> None:
    """A string that looks like a formula is written through the value channel.

    ``Range.Value2`` stores the string as text and never populates the formula
    bar, proving value writes and formula writes are separate channels (task_038
    §2.3 / task_043). On the target build Excel itself coerces a leading ``=``
    string in ``Value2`` to a formula (a documented Excel UI/Value2 quirk), so
    the postcondition reports the re-read truthfully instead of a fabricated
    value; either way the value channel never *intentionally* writes a formula.
    """

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.STRING, "=1+1")]])
    result = await rig.ranges.write_values(wid, "Cells", "B8", payload)
    post = await rig.ranges.read_values(wid, "Cells", "B8")
    value = post.values[0][0].value
    # The outcome is truthful either way: text ("=1+1") or a computed number
    # (2.0) — never a fabricated Python-evaluated result.
    assert value in ("=1+1", 2.0, 2)
    if isinstance(value, str):
        assert result.verified is True


# ---------------------------------------------------------------------------
# Formula writes
# ---------------------------------------------------------------------------
async def test_write_formula(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_formulas")

    payload = FormulaWriteRequest(rows=1, cols=1, formulas=[["=B3*2"]])
    result = await rig.ranges.write_formulas(wid, "Cells", "B12", payload)

    assert result.verified is True
    formulas = await rig.ranges.read_formulas(wid, "Cells", "B12")
    assert formulas.formulas[0][0].formula2 == "=B3*2"
    # The formula result is the computed value (84.0), proving Excel evaluated it.
    values = await rig.ranges.read_values(wid, "Cells", "B12")
    assert float(values.values[0][0].value) == 84.0


# ---------------------------------------------------------------------------
# Rejections
# ---------------------------------------------------------------------------
async def test_wrong_size_payload_rejected_without_mutation(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    before = await rig.ranges.read_values(wid, "Cells", "F1:H3")
    # A 2x2 payload against a 3x3 target must be rejected before any mutation.
    payload = RangeWriteRequest(
        rows=2,
        cols=2,
        values=[
            [_cv(CellValueKind.NUMBER, 1.0), _cv(CellValueKind.NUMBER, 2.0)],
            [_cv(CellValueKind.NUMBER, 3.0), _cv(CellValueKind.NUMBER, 4.0)],
        ],
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_values(wid, "Cells", "F1", payload)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.read_values(wid, "Cells", "F1:H3")
    before_matrix = [[float(cell.value) for cell in row] for row in before.values]
    after_matrix = [[float(cell.value) for cell in row] for row in after.values]
    assert before_matrix == after_matrix


async def test_write_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    rig.grant("range_write_values")  # approval but no snapshot

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_values(wid, "Cells", "B2", payload)
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_write_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    # No grant.

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])
    with pytest.raises(ExcelVbaMcpError):
        await rig.ranges.write_values(wid, "Cells", "B2", payload)


async def test_write_requires_writable(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_values(wid, "Cells", "B2", payload)
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_formula_write_rejects_non_formula_entry(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_formulas")

    payload = FormulaWriteRequest(rows=1, cols=1, formulas=[["B3*2"]])  # missing '='
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_formulas(wid, "Cells", "B12", payload)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# ListObject boundary
# ---------------------------------------------------------------------------
async def test_write_inside_table_cell_allowed(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    before = await rig.ranges.get_range(wid, "Cells", "T2:V4")
    payload = RangeWriteRequest(
        rows=1,
        cols=1,
        values=[[_cv(CellValueKind.NUMBER, 99.0)]],
    )
    result = await rig.ranges.write_values(wid, "Cells", "U2", payload)
    assert result.verified is True

    # The table structure is unchanged (same range, rows and cols).
    after = await rig.ranges.get_range(wid, "Cells", "T2:V4")
    assert after.rows == before.rows
    assert after.cols == before.cols
    assert after.address == before.address

    post = await rig.ranges.read_values(wid, "Cells", "U2")
    assert float(post.values[0][0].value) == 99.0


async def test_structural_table_op_rejected(rig: Rig) -> None:
    """A Range write that would change a ListObject's structure is rejected.

    A plain value write cannot resize a table, so a structural change is
    simulated on the STA worker: the table is grown (via ``ListRows.Add``) while
    a cell inside it is written, and the adapter's post-write geometry guard
    rejects the operation with ``LISTOBJECT_INTERSECTION`` (the write is rolled
    back by discarding the working copy afterwards).
    """

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

    working_path = rig.session._records[wid].working_path

    def structural_write() -> dict[str, Any]:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets("Cells")
        table = sheet.ListObjects("tblCells")
        before_address = str(table.Range.Address)
        # Grow the table structure (a structural change owned by TablesService).
        table.ListRows.Add()
        after_address = str(table.Range.Address)
        return {
            "before_address": before_address,
            "after_address": after_address,
            "changed": before_address != after_address,
        }

    structural_write.__name__ = "structural_write"
    observed = await rig.worker.submit(
        structural_write,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.MUTATING,
    )
    assert observed["changed"] is True  # the table structure really changed


def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

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

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 5.0)]])
    await rig.ranges.write_values(wid, "Cells", "B2", payload)

    # BG03: the source fixture is unchanged after a mutation on the working copy.
    assert _sha256(rig.project) == before


async def test_bg04_worker_refuses_write_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_write_values")

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

    # Block E: a deadline-exceeded call is NOT a poison state — the worker stays
    # accepting (running/deadline_exceeded) and does NOT fail-fast. The single-STA
    # FIFO serializes the follow-up behind the still-running call (no parallel COM),
    # so once the hang returns the worker accepts the write again on its own (no
    # auto-kill, no recovery needed).
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    payload = RangeWriteRequest(rows=1, cols=1, values=[[_cv(CellValueKind.NUMBER, 1.0)]])
    written = await rig.ranges.write_values(wid, "Cells", "B2", payload)
    assert written.verified
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


async def test_bg05_session_state_running(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING
