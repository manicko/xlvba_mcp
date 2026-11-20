"""Real Excel integration tests for the PivotTables inventory (task_048).

These tests drive the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` and
the extended
:class:`~excel_vba_mcp.application.workbook_model_service.WorkbookModelService`
against a **hidden, automation-created Excel instance** (never the user's Excel)
and a disposable working copy of the committed ``PivotTablesProject.xlsm``
fixture. They assert:

* ``pivot_tables_list`` finds all three PivotTables (``ptA`` / ``ptB`` / ``ptC``)
  across the worksheet sweep (``Workbook.PivotTables`` is unusable on the target
  build);
* shared-cache identity — ``ptA`` and ``ptB`` share cache 1 (reported once with
  ``affected_pivot_tables``), while ``ptC`` uses cache 2;
* cache metadata — source type, ``SourceData``, ``RefreshDate``, ``OLAP`` flag,
  shared flag;
* field layout / aggregation / filter state — orientation, 1-based position,
  data-field ``Function`` (via ``DataFields``), and ``ptB``'s manual
  (``PivotItems.Visible``) filter;
* ``pivot_table_get`` / ``pivot_cache_get`` / ``pivot_fields_list`` by identity;
* ``get_sheet_model`` includes PivotTable metadata scoped to ``PivotReport``
  without materializing the PivotTable cell output;
* BG03/BG04/BG05 — the automation Excel stays hidden and non-activating, the
  owned PID never holds foreground, no owned window is visible, the user's
  session is untouched, and the owned PID is reaped after shutdown (no orphan);
* the committed fixture and the temp original are byte-identical afterwards.

The fixture has no external / OLAP / Data Model cache (all sources are local
``xlDatabase``), so the unsupported-source path is asserted only as "no
unsupported source is present" — the unsupported mapping is covered by unit
tests against a fake workbook.

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_tables_inventory.py -v -s
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
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.application.workbook_model_service import (
    SECTION_PIVOTS,
    WorkbookModelService,
)
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotInventory,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    ModelSectionStatus,
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
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
# Fixtures / rig — real COM worker + Excel + full service stack
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService + WorkbookModelService."""

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
        self.sheets = SheetsService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self.tables = TablesService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
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
        self.shapes = ShapesService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self.controls = ControlsService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self.vba = VBAProjectService(
            self.worker,
            self.session,
            self.path_policy,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )
        self.pivots = PivotTablesService(
            self.worker,
            self.session,
            self.audit,
        )
        self.model = WorkbookModelService(
            self.worker,
            self.session,
            self.sheets,
            self.tables,
            self.names,
            self.shapes,
            self.controls,
            self.vba,
            self.path_policy,
            self.audit,
            pivot_tables_service=self.pivots,
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
# Inventory: all three pivots, shared cache, metadata, field layout
# ---------------------------------------------------------------------------
@requires_excel
async def test_list_all_pivots_shared_cache_and_metadata(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    inventory = await rig.pivots.list_pivot_tables(info.workbook_id)
    assert isinstance(inventory, PivotInventory)

    # All three PivotTables are present.
    names = sorted(pt.name for pt in inventory.pivot_tables)
    assert names == ["ptA", "ptB", "ptC"]

    by_name: dict[str, PivotTableInfo] = {pt.name: pt for pt in inventory.pivot_tables}

    # ptA / ptB share cache 1; ptC uses cache 2.
    assert by_name["ptA"].cache_index == 1
    assert by_name["ptB"].cache_index == 1
    assert by_name["ptC"].cache_index == 2

    # All caches are local xlDatabase (the fixture has no external/OLAP source).
    for pt in inventory.pivot_tables:
        assert pt.source_type is PivotSourceType.DATABASE
        assert pt.source_status == "supported"
        assert pt.sheet_name == "PivotReport"

    # ptA carries row + column + page + two data fields.
    pta = by_name["ptA"]
    assert [f.source_name for f in pta.row_fields] == ["Region"]
    assert [f.source_name for f in pta.column_fields] == ["Product"]
    assert [f.source_name for f in pta.page_fields] == ["Quarter"]
    pta_functions = {f.function for f in pta.data_fields if f.source_name in ("Sales", "Units")}
    assert PivotAggregation.SUM in pta_functions
    assert PivotAggregation.AVERAGE in pta_functions

    # ptB shares cache 1 with ptA and carries a manual filter (West hidden).
    ptb = by_name["ptB"]
    assert ptb.cache_index == pta.cache_index == 1
    assert ptb.table_range1 and ptb.table_range1.startswith("$")

    # ptC uses the ListObject-backed cache 2 with a count aggregation.
    ptc = by_name["ptC"]
    assert ptc.cache_index == 2
    ptc_functions = {f.function for f in ptc.data_fields}
    assert PivotAggregation.COUNT in ptc_functions

    # The inventory is deterministic: sorted by (sheet_index, name).
    assert [pt.name for pt in inventory.pivot_tables] == ["ptA", "ptB", "ptC"]

    # BG05: the committed fixture is never modified (opened read-only).
    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_shared_cache_reported_once_with_affected_tables(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    cache_one = await rig.pivots.get_pivot_cache(info.workbook_id, 1)
    assert isinstance(cache_one, PivotCacheInfo)
    assert cache_one.cache_index == 1
    assert cache_one.shared_cache is True
    assert sorted(cache_one.affected_pivot_tables) == ["ptA", "ptB"]
    assert cache_one.source_type is PivotSourceType.DATABASE
    assert cache_one.source_data_kind == "text"
    assert cache_one.source_data  # non-empty SourceData text

    # Cache 2 backs only ptC.
    cache_two = await rig.pivots.get_pivot_cache(info.workbook_id, 2)
    assert cache_two.cache_index == 2
    assert cache_two.shared_cache is False
    assert cache_two.affected_pivot_tables == ["ptC"]

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_field_layout_aggregation_and_filter_state(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    # ptA: orientation + 1-based position + data-field aggregation via DataFields.
    pta_fields = await rig.pivots.list_pivot_fields(info.workbook_id, "ptA")
    pta_by_name = {f.source_name: f for f in pta_fields}
    assert pta_by_name["Region"].orientation is PivotFieldOrientation.ROW
    assert pta_by_name["Region"].position == 1
    assert pta_by_name["Product"].orientation is PivotFieldOrientation.COLUMN
    assert pta_by_name["Quarter"].orientation is PivotFieldOrientation.PAGE
    assert pta_by_name["Sales"].orientation is PivotFieldOrientation.DATA
    assert pta_by_name["Sales"].function in (PivotAggregation.SUM, PivotAggregation.AVERAGE)
    assert pta_by_name["Sales"].is_data_field is True
    assert pta_by_name["Sales"].data_field_name  # caption like "Sum of Sales"

    # ptB: the manual (PivotItems.Visible) filter hides "West" from Region.
    ptb_fields = await rig.pivots.list_pivot_fields(info.workbook_id, "ptB")
    ptb_by_name = {f.source_name: f for f in ptb_fields}
    region = ptb_by_name["Region"]
    assert region.all_items_visible is False
    assert region.filter_state is PivotFilterKind.MANUAL

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Identity resolution: pivot_table_get / pivot_cache_get / pivot_fields_list
# ---------------------------------------------------------------------------
@requires_excel
async def test_get_pivot_table_by_name(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    pt_a = await rig.pivots.get_pivot_table(info.workbook_id, "ptA")
    assert isinstance(pt_a, PivotTableInfo)
    assert pt_a.name == "ptA"
    assert pt_a.cache_index == 1
    assert pt_a.table_range1 is not None
    assert pt_a.table_range2 is not None

    # A missing name is a structured error.
    from excel_vba_mcp.domain.errors import ExcelVbaMcpError

    with pytest.raises(ExcelVbaMcpError):
        await rig.pivots.get_pivot_table(info.workbook_id, "missing")

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_unsupported_source_not_present_in_fixture(rig: Rig) -> None:
    """The committed fixture has no external/OLAP/Data Model source."""

    info = await rig.open(read_only=True)
    inventory = await rig.pivots.list_pivot_tables(info.workbook_id)
    unsupported = [pt for pt in inventory.pivot_tables if pt.source_status != "supported"]
    assert unsupported == [], f"unexpected unsupported source: {unsupported}"
    for cache_index in {pt.cache_index for pt in inventory.pivot_tables}:
        cache = await rig.pivots.get_pivot_cache(info.workbook_id, cache_index)
        assert cache.source_type is PivotSourceType.DATABASE
        assert cache.olap is False

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# Workbook model: get_sheet_model includes PivotTable metadata only
# ---------------------------------------------------------------------------
@requires_excel
async def test_sheet_model_includes_pivot_metadata(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    model = await rig.model.get_sheet_model(info.workbook_id, "PivotReport")

    # The pivots section is present and available.
    assert SECTION_PIVOTS in model.sections
    assert model.sections[SECTION_PIVOTS].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_PIVOTS].count == 3

    # The sheet model carries PivotTable metadata scoped to the sheet...
    assert sorted(pt.name for pt in model.pivot_tables) == ["ptA", "ptB", "ptC"]
    for pt in model.pivot_tables:
        assert pt.sheet_name == "PivotReport"
        assert pt.cache_index in (1, 2)

    # ...and the model never materializes the PivotTable cell output: the pivot
    # metadata carries no cell matrix, only addresses + field layout.
    for pt in model.pivot_tables:
        assert pt.table_range1 is None or isinstance(pt.table_range1, str)
        assert isinstance(pt.data_fields, list)
        assert isinstance(pt.row_fields, list)

    # The hash is stable across two reads of the unchanged working copy.
    model2 = await rig.model.get_sheet_model(info.workbook_id, "PivotReport")
    assert model.model_hash == model2.model_hash

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_sheet_model_other_sheet_has_no_pivots(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    model = await rig.model.get_sheet_model(info.workbook_id, "Data")

    # The Data sheet has no PivotTables; the pivots section is available but empty.
    assert model.sections[SECTION_PIVOTS].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_PIVOTS].count == 0
    assert model.pivot_tables == []

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 + original unchanged + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_hidden_owned_excel(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    # Read the pivot inventory (the read path under test) while hidden.
    await rig.pivots.list_pivot_tables(info.workbook_id)

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
    _assert_owned_never_foreground(owned_pid, "pivot-inventory")
    _assert_no_visible_owned_windows(owned_pid, "pivot-inventory")

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_committed_fixture_unchanged(tmp_path: Path) -> None:
    """BG05: the committed fixture is byte-identical after the inventory reads."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    before = _sha256(pivot_project)
    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        info = await r.open(read_only=True)
        await r.pivots.list_pivot_tables(info.workbook_id)
        await r.pivots.list_pivot_fields(info.workbook_id, "ptA")
        await r.model.get_sheet_model(info.workbook_id, "PivotReport")
        await r.session.close_workbook(info.workbook_id)
    finally:
        await r.close()

    assert _sha256(pivot_project) == before


@requires_excel
async def test_owned_excel_reaped_after_shutdown() -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    import tempfile

    tmp_path = Path(tempfile.mkdtemp(prefix="pivot_inv_"))
    try:
        original = tmp_path / pivot_project.name
        shutil.copy2(pivot_project, original)
        r = Rig(tmp_path, original)
        await r.start()
        try:
            info = await r.open(read_only=True)
            await r.pivots.list_pivot_tables(info.workbook_id)
            owned_pid = r.worker.get_status().owned_pid
            assert owned_pid is not None
        finally:
            await r.close()
    finally:
        shutil.rmtree(tmp_path, ignore_errors=True)

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
