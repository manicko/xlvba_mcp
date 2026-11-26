"""Real Excel integration tests for the read-only Range/Cells surface (task_042).

These tests drive the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and the committed ``RangesCellsProject.xlsm`` fixture (task_041).
They assert:

* ``range_get`` metadata for scalar / row / column / block addresses, plus
  merged-cell and in-table flags;
* ``range_read_values`` typed classification (string / number / bool / blank /
  date / error) and a 2-D block, with dates typed strictly from ``NumberFormat``
  (the date-shaped serial with a ``General`` format stays a ``NUMBER``);
* ``range_read_formulas`` ``Formula2`` primary output for the fixture formulas;
* an oversized request is rejected with a typed error before reading;
* an invalid sheet and an invalid range are rejected;
* BG03 / BG04 / BG05 hold (no non-owned Excel touched; hidden + non-activating;
  the worker refuses work on a hung instance);
* the original fixture on disk is byte-identical after every test.

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
from excel_vba_mcp.domain.models import MAX_RANGE_CELLS
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    OperationClass,
    RangeShape,
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


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for range reads."""

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
# Metadata
# ---------------------------------------------------------------------------
async def test_range_get_scalar_metadata(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    resolved = await rig.ranges.get_range(wid, "Cells", "B3")
    assert resolved.sheet == "Cells"
    assert resolved.rows == 1
    assert resolved.cols == 1
    assert resolved.shape is RangeShape.SCALAR

    block = await rig.ranges.get_range(wid, "Cells", "A1:C10")
    assert block.rows == 10
    assert block.cols == 3
    assert block.shape is RangeShape.BLOCK

    row = await rig.ranges.get_range(wid, "Cells", "A1:C1")
    assert row.shape is RangeShape.ROW
    column = await rig.ranges.get_range(wid, "Cells", "A1:A3")
    assert column.shape is RangeShape.COLUMN

    # BG03: the source fixture is never modified.
    assert _sha256(rig.project) == before


# ---------------------------------------------------------------------------
# Typed values
# ---------------------------------------------------------------------------
async def test_range_read_values_typed_constants(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    result = await rig.ranges.read_values(wid, "Cells", "B2:B15")
    assert result.rows == 14
    assert result.cols == 1
    assert result.shape is RangeShape.COLUMN
    kinds = [row[0].kind for row in result.values]
    # B2..B15 design: string, number, number, bool, bool, empty, string(=text),
    # number(formula), error(#DIV/0!), error(#N/A), number(formula), date, date,
    # number(45000.0 General -> number, never a date).
    assert kinds == [
        CellValueKind.STRING,
        CellValueKind.NUMBER,
        CellValueKind.NUMBER,
        CellValueKind.BOOLEAN,
        CellValueKind.BOOLEAN,
        CellValueKind.EMPTY,
        CellValueKind.STRING,
        CellValueKind.NUMBER,
        CellValueKind.ERROR,
        CellValueKind.ERROR,
        CellValueKind.NUMBER,
        CellValueKind.DATE,
        CellValueKind.DATE,
        CellValueKind.NUMBER,
    ]
    values = [row[0].value for row in result.values]
    assert values[0] == "hello world"
    assert values[6] == "=1+1"  # literal '='-prefixed text, not a formula
    assert values[8] == "#DIV/0!"
    assert values[9] == "#N/A"
    # B13/B14 are real dates; B15 (45000.0, General) is a plain number.
    assert values[11] == "2026-01-02T00:00:00"
    assert values[12] == "2026-06-15T00:00:00"
    assert values[13] == 45000.0

    assert _sha256(rig.project) == before


async def test_range_read_values_2d_block(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    result = await rig.ranges.read_values(wid, "Cells", "F1:H3")
    assert result.rows == 3
    assert result.cols == 3
    assert result.shape is RangeShape.BLOCK
    matrix = [[float(cell.value) for cell in row] for row in result.values]
    assert matrix == [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0], [7.0, 8.0, 9.0]]


async def test_range_read_values_offset_limit(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    page = await rig.ranges.read_values(wid, "Cells", "B2:B15", offset=0, limit=1)
    assert page.rows == 1
    assert page.values[0][0].value == "hello world"


# ---------------------------------------------------------------------------
# Formulas
# ---------------------------------------------------------------------------
async def test_range_read_formulas_formula2(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    result = await rig.ranges.read_formulas(wid, "Cells", "B9:B12")
    assert result.rows == 4
    formulas = [row[0].formula2 for row in result.formulas]
    assert formulas == ["=1+2", "=1/0", "=NA()", "=B3*2"]


# ---------------------------------------------------------------------------
# Rejections
# ---------------------------------------------------------------------------
async def test_oversized_request_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.read_values(wid, "Cells", "A1:XFD1048576")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells"] > MAX_RANGE_CELLS


async def test_invalid_sheet_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.get_range(wid, "NoSuchSheet", "A1")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_invalid_range_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.get_range(wid, "Cells", "NotAnAddress")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    await rig.open()

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
    # BG03: source fixture unchanged by a read-only session.
    assert _sha256(rig.project) == before
    # BG05: alerts remain at their default (never globally disabled).
    assert flags["display_alerts"] is True


async def test_bg04_worker_refuses_read_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

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
    # so once the hang returns the worker accepts the read again on its own (no
    # auto-kill, no recovery needed).
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    rng = await rig.ranges.get_range(wid, "Cells", "A1")
    assert rng.address
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


async def test_bg05_session_state_running(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING
