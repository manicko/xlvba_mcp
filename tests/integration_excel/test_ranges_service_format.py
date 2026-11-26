"""Real Excel integration tests for the bounded Range formatting surface (task_044).

These tests drive the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and a disposable copy of the committed ``RangesCellsProject.xlsm``
fixture (task_041). They assert:

* ``range_get_format`` returns a bounded, tri-state ``FormatInfo`` — the
  fixture's number formats (``B3`` -> ``0``, ``B4`` -> ``0.00``) plus concrete
  font / fill / alignment values for a uniform range;
* a range with mixed formatting reports the affected properties as ``None`` +
  ``mixed_properties`` (never flattened);
* ``range_set_format`` applies bold / fill / number_format / alignment and the
  targeted properties read back; a partial update leaves the other properties
  untouched;
* an invalid enum / value is rejected with ``INVALID_RANGE`` and no mutation;
* BG03 / BG04 / BG05 hold (the owned Excel is hidden + non-activating, the
  source fixture is byte-identical, and the worker refuses work on a hung
  instance).

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
from excel_vba_mcp.domain.models import RangeFormatPatch
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


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for range formats."""

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


async def _open_for_format(rig: Rig) -> str:
    """Open the disposable copy writable, snapshot it and grant the format action."""

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_set_format")
    return wid


# ---------------------------------------------------------------------------
# Read format (tri-state)
# ---------------------------------------------------------------------------
@requires_excel
async def test_get_format_number_format_from_fixture(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    b3 = await rig.ranges.get_format(wid, "Cells", "B3")
    assert b3.number_format == "0"
    assert b3.mixed_properties == []

    b4 = await rig.ranges.get_format(wid, "Cells", "B4")
    assert b4.number_format == "0.00"

    # BG03: a read never touches the source fixture.
    assert _sha256(rig.project) == before


@requires_excel
async def test_get_format_font_fill_alignment_uniform(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    fmt = await rig.ranges.get_format(wid, "Cells", "B3")
    # Font is concrete for a single cell.
    assert isinstance(fmt.font.name, str) and fmt.font.name
    assert isinstance(fmt.font.size, float) and fmt.font.size > 0
    assert fmt.font.bold is False
    assert fmt.font.italic is False
    # Alignment enums are concrete for a single cell.
    assert fmt.alignment.horizontal in {"xlGeneral", "xlLeft", "xlCenter", "xlRight"}
    assert fmt.alignment.vertical in {"xlTop", "xlBottom", "xlCenter"}
    assert fmt.alignment.wrap_text is False
    # Unsupported aspects are always reported, never silently dropped.
    assert "borders" in fmt.unsupported


@requires_excel
async def test_get_format_mixed_range_is_not_flattened(rig: Rig) -> None:
    wid = await _open_for_format(rig)

    # Give A1 and A2 different number formats and bold states.
    await rig.ranges.set_format(wid, "Cells", "A1", RangeFormatPatch(number_format="0.00"))
    await rig.ranges.set_format(wid, "Cells", "A1", RangeFormatPatch(font_bold=True))
    # A2 keeps the default (General, not bold).

    mixed = await rig.ranges.get_format(wid, "Cells", "A1:A2")
    assert mixed.number_format is None
    assert mixed.font.bold is None
    assert "number_format" in mixed.mixed_properties
    assert "font.bold" in mixed.mixed_properties


# ---------------------------------------------------------------------------
# Set format (partial writes) + postcondition
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_format_bold_fill_number_alignment(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_format(rig)

    patch = RangeFormatPatch(
        font_bold=True,
        font_color=255,
        fill_color=65280,
        number_format="0.000",
        alignment_horizontal="xlCenter",
        alignment_vertical="xlTop",
        alignment_wrap_text=True,
    )
    result = await rig.ranges.set_format(wid, "Cells", "C7", patch)
    assert result.verified is True

    fmt = await rig.ranges.get_format(wid, "Cells", "C7")
    assert fmt.font.bold is True
    assert fmt.font.color == 255
    assert fmt.fill.color == 65280
    assert fmt.number_format == "0.000"
    assert fmt.alignment.horizontal == "xlCenter"
    assert fmt.alignment.vertical == "xlTop"
    assert fmt.alignment.wrap_text is True

    # BG03: the source fixture is unchanged after a format mutation.
    assert _sha256(rig.project) == before


@requires_excel
async def test_set_format_partial_update_leaves_other_props(rig: Rig) -> None:
    wid = await _open_for_format(rig)

    # Establish a baseline: bold + italic + size on one cell.
    await rig.ranges.set_format(
        wid, "Cells", "D7", RangeFormatPatch(font_bold=True, font_italic=True, font_size=14.0)
    )
    before = await rig.ranges.get_format(wid, "Cells", "D7")

    # Update ONLY the number format.
    result = await rig.ranges.set_format(wid, "Cells", "D7", RangeFormatPatch(number_format="0.00"))
    assert result.verified is True
    assert result.applied == ["number_format"]

    after = await rig.ranges.get_format(wid, "Cells", "D7")
    assert after.number_format == "0.00"
    # Untouched properties are unchanged.
    assert after.font.bold == before.font.bold is True
    assert after.font.italic == before.font.italic is True
    assert after.font.size == before.font.size


@requires_excel
async def test_set_format_clear_number_format_to_general(rig: Rig) -> None:
    wid = await _open_for_format(rig)

    await rig.ranges.set_format(wid, "Cells", "E7", RangeFormatPatch(number_format="0.000"))
    result = await rig.ranges.set_format(wid, "Cells", "E7", RangeFormatPatch(number_format=None))
    assert result.verified is True
    fmt = await rig.ranges.get_format(wid, "Cells", "E7")
    assert fmt.number_format == "General"


# ---------------------------------------------------------------------------
# Rejections (no mutation)
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_format_invalid_enum_rejected_without_mutation(rig: Rig) -> None:
    wid = await _open_for_format(rig)

    before = await rig.ranges.get_format(wid, "Cells", "F7")
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_format(
            wid, "Cells", "F7", RangeFormatPatch(alignment_horizontal="xlBogus")
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_format(wid, "Cells", "F7")
    assert after.alignment.horizontal == before.alignment.horizontal


@requires_excel
async def test_set_format_out_of_range_color_rejected(rig: Rig) -> None:
    wid = await _open_for_format(rig)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_format(wid, "Cells", "G7", RangeFormatPatch(font_color=0x1000000))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ---------------------------------------------------------------------------
# Gates
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_format_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    rig.grant("range_set_format")  # approval but no snapshot

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_format(wid, "Cells", "A1", RangeFormatPatch(font_bold=True))
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


@requires_excel
async def test_set_format_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.ranges.set_format(wid, "Cells", "A1", RangeFormatPatch(font_bold=True))


@requires_excel
async def test_set_format_requires_writable(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_set_format")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_format(wid, "Cells", "A1", RangeFormatPatch(font_bold=True))
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_format(rig)

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

    await rig.ranges.set_format(wid, "Cells", "H7", RangeFormatPatch(font_bold=True))

    # BG03: the source fixture is unchanged after a format mutation.
    assert _sha256(rig.project) == before


@requires_excel
async def test_bg04_worker_refuses_format_on_hung_excel(rig: Rig) -> None:
    wid = await _open_for_format(rig)

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
    # so once the hang returns the worker accepts the format write again on its own
    # (no auto-kill, no recovery needed).
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    fmt = await rig.ranges.set_format(wid, "Cells", "H7", RangeFormatPatch(font_bold=True))
    assert fmt.verified
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


@requires_excel
async def test_bg05_session_state_running(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING
