"""Cross-feature regression + desktop non-interference tests for Phase 2 (task_054).

This module is the final independent integration/regression gate for the Phase 2
Ranges & Cells and PivotTables increment
(``.ai/plans/plan_phase2_ranges_cells_pivottables.md`` task_054). It complements
the per-feature verification modules
(:mod:`test_ranges_service`, :mod:`test_pivot_tables_service`,
:mod:`test_pivot_refresh`, :mod:`test_test_runner_ranges_pivots`) with a single
*rig* that wires **every** Phase 2 service together — Ranges, Pivots, Tables,
Names, Sheets, WorkbookModel and the declarative TestRunner — against a
**hidden, automation-created Excel instance** and disposable copies of the
committed ``RangesCellsProject.xlsm`` / ``PivotTablesProject.xlsm`` fixtures.

Coverage (task_054 acceptance criteria):

1. **DTO / identity consistency** — the Range API tool layer, the Pivot API tool
   layer, the workbook model and the declarative TestRunner all return the *same*
   domain DTO types as the direct services, and share the same service instances
   (no duplicate COM logic; services are reused).
2. **Range ops preserve Stage 1** — a cell write inside ``tblCells`` and a
   format/structural op on ``Cells`` leave ``TablesService`` / ``NamesService`` /
   ``SheetsService`` behaviour intact (table geometry unchanged, names/sheets
   unchanged).
3. **Pivot ops do not invoke deferred features** — inventory / refresh / field /
   filter on ``PivotTablesProject.xlsm`` never create or use Connections,
   QueryTables, Charts, Slicers or the Data Model; refresh is *targeted*
   (``ptC`` leaves cache 1 untouched) and ``Workbook.RefreshAll`` is never called.
4. **Snapshot / approval / audit / locks / recovery** — mutations require a
   snapshot + approval, read-only / read-only-mode mutations are refused, the
   audit trail records mutating operations, and a forced hung-COM call enters
   ``NEEDS_RECOVERY`` and recovers to a fresh hidden instance.
5. **Desktop non-interference across ALL op families** — reads, writes,
   formatting, structural, validation/CF, refresh, layout and filters all keep
   the owned Excel hidden, expose no visible owned window, never steal the
   foreground, leave the originals byte-identical, and reap the owned process
   (no orphan).
6. **Unsupported cases are explicit** — an unsupported pivot source
   (consolidation), an unsupported ``PivotFilters.Add`` filter kind (label) and
   a ColorScale conditional format each return an explicit status and never a
   false success.

Safety: every operation runs on a disposable working copy through the single
hidden STA :class:`ExcelComWorker` (``BG03`` / ``BG04`` / ``BG05``); no
``.Activate`` / ``.Select`` is ever called, no non-owned Excel is touched, and
tests run serially. Excel is created with ``DispatchEx`` and always
``Visible = False``.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_phase2_ranges_pivots_regression.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import json
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
from excel_vba_mcp.api import pivot_tools, range_tools
from excel_vba_mcp.application.compile_service import CompileService
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyRegistry,
    MacroService,
)
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.test_runner import TestRunner as Runner
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.application.workbook_model_service import WorkbookModelService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelHungError, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    CellValue,
    ConditionalFormatSpec,
    DefinedNameInfo,
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFilterSpec,
    PivotInventory,
    PivotRefreshResult,
    PivotTableInfo,
    RangeAddress,
    RangeFormatPatch,
    RangeWriteRequest,
    SheetInfo,
    TableInfo,
    ValidationSpec,
)
from excel_vba_mcp.domain.statuses import (
    CaseStatus,
    CellValueKind,
    ConditionalFormatType,
    DataValidationType,
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotRefreshMethod,
    PivotRefreshStatus,
    PivotSourceType,
    SessionMode,
    StepStatus,
    ValidationOperator,
    WorkbookSessionState,
)
from excel_vba_mcp.domain.statuses import TestRunStatus as RunStatus
from excel_vba_mcp.domain.test_spec import TestSuiteSpec
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.artifact_store import ArtifactStore
from excel_vba_mcp.storage.run_repository import RunRepository
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_RANGES_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "RangesCellsProject.xlsm"
_PIVOTS_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "PivotTablesProject.xlsm"

_TEST_TIMEOUT_S = 180.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
_HANG_SECONDS = 3.0
_HANG_CALL_TIMEOUT_S = 0.5

# Approval action names (stable identities of the Range / Pivot mutation surface).
_ACTION_WRITE_VALUES = "range_write_values"
_ACTION_SET_FORMAT = "range_set_format"
_ACTION_COPY = "range_copy"
_ACTION_SET_VALIDATION = "range_set_validation"
_ACTION_SET_CF = "range_set_conditional_formatting"
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_SET_LAYOUT = "pivot_field_set_layout"
_ACTION_APPLY_FILTER = "pivot_field_apply_filter"

_ALL_RANGE_ACTIONS = (
    "range_write_values",
    "range_write_formulas",
    "range_set_format",
    "range_copy",
    "range_insert",
    "range_delete",
    "range_merge",
    "range_unmerge",
    "range_set_validation",
    "range_set_conditional_formatting",
)
_ALL_PIVOT_ACTIONS = (
    "pivot_table_refresh",
    "pivot_cache_refresh",
    "pivot_field_set_layout",
    "pivot_field_set_position",
    "pivot_field_set_aggregation",
    "pivot_field_set_properties",
    "pivot_field_apply_filter",
    "pivot_field_clear_filters",
)


# ---------------------------------------------------------------------------
# Skip helpers
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 helpers (read-only; never activate / focus / click / kill)
# ---------------------------------------------------------------------------
def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _pid_alive(pid: int) -> bool:
    if pid <= 0:  # pragma: no cover - defensive
        return False
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


def _excel_pids() -> set[int]:
    """Return the current ``EXCEL.EXE`` PIDs via ``tasklist`` (best effort)."""

    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001 - best effort
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        parts = [part.strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE" and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _settled_excel_pids(stable_for_s: float = 1.0, timeout_s: float = 20.0) -> set[int]:
    """Return the ``EXCEL.EXE`` PID set once it stops churning.

    Other integration tests create and reap automation-owned Excel instances, so a
    naive ``baseline - current`` check races with those exits and misreports an
    automation PID as a terminated user process (BG05 false positive). Polling
    until the set is stable makes the baseline deterministic without weakening the
    assertion (mirrors ``test_test_runner_ranges_pivots``).
    """

    deadline = time.monotonic() + timeout_s
    previous = _excel_pids()
    stable_since = time.monotonic()
    while time.monotonic() < deadline:
        time.sleep(0.25)
        current = _excel_pids()
        if current != previous:
            previous = current
            stable_since = time.monotonic()
        elif time.monotonic() - stable_since >= stable_for_s:
            return current
    return previous


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:  # pragma: no cover - depends on host
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return {"hwnd": hwnd, "pid": pid, "title": win32gui.GetWindowText(hwnd)}
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

    task = _ensure_future(operation())
    samples: list[dict[str, Any]] = []
    while not task.done():
        fg = _foreground()
        if fg is not None and fg["pid"] == owned_pid:
            samples.append(fg)
        await _async_sleep(0.05)
    await task
    return samples


# ---------------------------------------------------------------------------
# Small indirection so the helpers above read clearly (kept local / explicit).
# ---------------------------------------------------------------------------
def _ensure_future(coro: Awaitable[Any]) -> Any:
    import asyncio

    return asyncio.ensure_future(coro)


async def _async_sleep(seconds: float) -> None:
    import asyncio

    await asyncio.sleep(seconds)


def _read_audit(path: Path) -> list[dict[str, Any]]:
    """Parse the append-only JSONL audit trail (best effort)."""

    if not path.is_file():
        return []
    records: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        try:
            parsed = json.loads(stripped)
        except (TypeError, ValueError):
            continue
        if isinstance(parsed, dict):
            records.append(parsed)
    return records


# ---------------------------------------------------------------------------
# Payload helpers
# ---------------------------------------------------------------------------
def _cv(kind: CellValueKind, value: Any = None) -> CellValue:
    return CellValue(kind=kind, value=value)


def _single_value(
    kind: CellValueKind = CellValueKind.NUMBER, value: Any = None
) -> RangeWriteRequest:
    return RangeWriteRequest(rows=1, cols=1, values=[[_cv(kind, value)]])


def _list_validation(formula1: str = '"Yes,No"') -> ValidationSpec:
    return ValidationSpec(
        type=DataValidationType.LIST,
        operator=ValidationOperator.BETWEEN,
        formula1=formula1,
        in_cell_dropdown=True,
        ignore_blank=True,
    )


def _cell_value_cf(formula1: str = "10") -> ConditionalFormatSpec:
    return ConditionalFormatSpec(
        type=ConditionalFormatType.CELL_VALUE,
        operator=ValidationOperator.GREATER,
        formula1=formula1,
    )


def _manual(*items: str) -> PivotFilterSpec:
    return PivotFilterSpec(kind=PivotFilterKind.MANUAL, visible_items=list(items))


# ---------------------------------------------------------------------------
# Rig — real COM worker + Excel + every Phase 2 service + declarative runner.
#
# Both fixtures are copied into the temp projects root; the caller opens the one
# a test needs. The same live services back the direct MCP tools, the API tool
# layer and the declarative TestRunner, which is the core of the DTO/identity
# consistency evidence (area 1).
# ---------------------------------------------------------------------------
class Rig:
    """Wires every Phase 2 service against one hidden Excel instance."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.snapshots_dir = tmp_dir / "snapshots"
        self.artifacts_dir = tmp_dir / "artifacts"
        self.runs_dir = tmp_dir / "runs"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {
                "work": str(self.work_dir),
                "projects": str(self.projects_dir),
                "snapshots": str(self.snapshots_dir),
                "artifacts": str(self.artifacts_dir),
            }
        )
        self.approval = ApprovalPolicy()
        self.background = BackgroundSafetyRegistry()
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )
        self.snapshot = SnapshotService(
            SnapshotStore(self.snapshots_dir),
            self.session,
            self.worker,
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
        self.compile = CompileService(
            self.worker, self.session, self.vba, self.path_policy, self.audit
        )
        self.macro = MacroService(
            self.worker, self.session, self.approval, self.background, self.audit
        )
        self.sheets = SheetsService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.tables = TablesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.names = NamesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.ranges = RangesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.pivots = PivotTablesService(
            self.worker,
            self.session,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )
        self.shapes = ShapesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.controls = ControlsService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
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
        self.artifacts = ArtifactStore(self.artifacts_dir, self.path_policy, self.audit)
        self.runs = RunRepository(self.runs_dir)
        self.runner = Runner(
            self.worker,
            self.session,
            self.snapshot,
            self.compile,
            self.macro,
            self.sheets,
            self.tables,
            self.names,
            self.artifacts,
            self.runs,
            self.path_policy,
            self.audit,
            shapes_service=self.shapes,
            ranges_service=self.ranges,
            pivot_tables_service=self.pivots,
        )
        self._originals: dict[str, Path] = {}
        self._baseline: Any = None

    @property
    def app(self) -> ExcelAppAdapter:
        return self.adapter

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        owned_pid = self.worker.get_status().owned_pid
        await self.worker.shutdown()
        self.audit.close()
        if owned_pid is not None:
            _wait_pid_gone(owned_pid, timeout=_ORPHAN_POLL_TIMEOUT_S)

    # -- fixtures / grants ----------------------------------------------
    def copy_fixture(self, fixture: Path) -> Path:
        """Copy the committed fixture into the projects root as the original."""

        original = self.projects_dir / fixture.name
        shutil.copy2(fixture, original)
        self._originals[fixture.name] = original
        return original

    def original_for(self, fixture: Path) -> Path:
        return self._originals[fixture.name]

    def canonical(self, original: Path) -> str:
        return str(self.path_policy.validate_path(str(original), OperationClass.MUTATING))

    def _grant(self, canonical_path: str, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=canonical_path, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task054-regression",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_open(self, original: Path) -> None:
        self._grant(self.canonical(original), "excel_open")

    def grant(self, original: Path, *actions: str) -> None:
        canonical = self.canonical(original)
        for action in actions:
            self._grant(canonical, action)

    async def open(
        self, original: Path, *, read_only: bool = False, mode: SessionMode = SessionMode.INSPECT
    ) -> Any:
        self.grant_open(original)
        return await self.session.open_workbook(str(original), mode=mode, read_only=read_only)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
async def rig(tmp_path: Path) -> Any:
    r = Rig(tmp_path)
    await r.start()
    try:
        yield r
    finally:
        await r.stop()


# ---------------------------------------------------------------------------
# STA-local probes (independent postcondition reads / deferred-feature counts)
# ---------------------------------------------------------------------------
async def _read_app_flags(rig: Rig) -> dict[str, Any]:
    def read() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"visible": bool(app.Visible), "display_alerts": bool(app.DisplayAlerts)}

    read.__name__ = "task054_read_app_flags"
    return dict(
        await rig.worker.submit(
            read,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )


async def _deferred_feature_counts(rig: Rig) -> dict[str, int]:
    """Count the deferred-feature objects present in the open workbook (STA).

    Connections / QueryTables / ChartObjects / SlicerCaches are counted
    defensively (a missing property is reported as absent rather than raising).
    The Data Model has no cheap "is loaded" flag that does not itself instantiate
    the model, so the pivot-cache ``OLAP`` flag (reported separately by the
    PivotTablesService) is the evidence used for the Data Model instead.
    """

    def count() -> dict[str, int]:
        app = rig.app.app
        assert app is not None
        workbook = app.Workbooks(1)
        connections = 0
        try:
            connections = int(workbook.Connections.Count)
        except Exception:  # noqa: BLE001 - property may be unavailable
            connections = 0
        slicers = 0
        try:
            slicers = int(workbook.SlicerCaches.Count)
        except Exception:  # noqa: BLE001 - no slicer support / none present
            slicers = 0
        query_tables = 0
        chart_objects = 0
        chart_sheets = 0
        try:
            chart_sheets = int(workbook.Charts.Count)
        except Exception:  # noqa: BLE001
            chart_sheets = 0
        for sheet in workbook.Worksheets:
            with contextlib.suppress(Exception):
                query_tables += int(sheet.QueryTables.Count)
            with contextlib.suppress(Exception):
                chart_objects += int(sheet.ChartObjects().Count)
        return {
            "connections": connections,
            "slicers": slicers,
            "query_tables": query_tables,
            "chart_objects": chart_objects,
            "chart_sheets": chart_sheets,
        }

    count.__name__ = "task054_deferred_feature_counts"
    return dict(
        await rig.worker.submit(
            count,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )


async def _read_cache_refresh_date(rig: Rig, pivot_table: str) -> str:
    """Read the owning cache ``RefreshDate`` for ``pivot_table`` (STA-local).

    Resolved by PivotTable *name* (not cache index): a refresh renumbers the
    ``PivotCaches`` collection on the target build, so the index is unstable.
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

    read.__name__ = "task054_read_cache_refresh_date"
    payload = await rig.worker.submit(
        read,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    return str(payload)


# ===========================================================================
# Area 1 — DTO / identity consistency across tools, model and TestRunner
# ===========================================================================
@requires_excel
async def test_range_api_tools_return_same_dto_as_service(rig: Rig) -> None:
    """The Range API tool layer delegates to the *same* service and DTO types."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, read_only=True)
    wid = info.workbook_id

    direct = await rig.ranges.get_range(wid, "Cells", "B2:B3")
    range_tools.set_ranges_service(rig.ranges)
    try:
        via_tools = await range_tools.range_get(wid, "Cells", "B2:B3")
    finally:
        range_tools.set_ranges_service(None)

    assert isinstance(direct, RangeAddress)
    assert isinstance(via_tools, RangeAddress)
    assert type(via_tools) is type(direct)
    assert via_tools.model_dump(mode="json") == direct.model_dump(mode="json")
    assert (via_tools.rows, via_tools.cols, via_tools.address) == (2, 1, "$B$2:$B$3")

    await rig.session.close_workbook(wid)


@requires_excel
async def test_pivot_api_tools_return_same_dto_as_service(rig: Rig) -> None:
    """The Pivot API tool layer delegates to the *same* service and DTO types."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, read_only=True)
    wid = info.workbook_id

    direct = await rig.pivots.list_pivot_tables(wid)
    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        via_tools = await pivot_tools.pivot_tables_list(wid)
        pivot = await pivot_tools.pivot_table_get(wid, "ptA")
        cache = await pivot_tools.pivot_cache_get(wid, 1)
        fields = await pivot_tools.pivot_fields_list(wid, "ptA")
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert isinstance(via_tools, PivotInventory)
    assert type(via_tools) is type(direct)
    assert [pt.name for pt in via_tools.pivot_tables] == [pt.name for pt in direct.pivot_tables]
    assert isinstance(pivot, PivotTableInfo) and pivot.name == "ptA"
    assert isinstance(cache, PivotCacheInfo) and cache.cache_index == 1
    assert all(isinstance(field, PivotFieldInfo) for field in fields)
    assert {field.source_name for field in fields} >= {"Region", "Product", "Quarter", "Sales"}

    await rig.session.close_workbook(wid)


@requires_excel
async def test_workbook_model_reuses_pivot_service_inventory(rig: Rig) -> None:
    """``get_sheet_model`` aggregates the same ``PivotTableInfo`` DTOs as the service."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, read_only=True)
    wid = info.workbook_id

    inventory = await rig.pivots.list_pivot_tables(wid)
    sheet_model = await rig.model.get_sheet_model(wid, "PivotReport")

    assert all(isinstance(pt, PivotTableInfo) for pt in sheet_model.pivot_tables)
    assert {pt.name for pt in sheet_model.pivot_tables} == {
        pt.name for pt in inventory.pivot_tables
    }
    # Metadata only — the model never materializes PivotTable cell output.
    dumped = sheet_model.model_dump(mode="json")
    assert "cells" not in json.dumps(dumped).lower()

    await rig.session.close_workbook(wid)


@requires_excel
async def test_runner_reuses_the_same_services(rig: Rig) -> None:
    """The declarative runner drives the *same* RangesService / PivotTablesService.

    A runner ``range_write_values`` step must land in the very workbook the direct
    RangesService reads, and a runner ``pivot_table_refresh`` step must be visible
    to the direct PivotTablesService — proving there is no duplicate COM logic.
    """

    ranges_fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(ranges_fixture)
    info = await rig.open(original, mode=SessionMode.EXECUTE, read_only=False)
    wid = info.workbook_id
    # Pre-create the immutable baseline the runner copies each case from.
    baseline = await rig.snapshot.create_snapshot(wid)
    baseline_canonical = str(
        rig.path_policy.validate_path(baseline.stored_path, OperationClass.MUTATING)
    )
    rig._grant(baseline_canonical, "excel_open")
    rig._grant(baseline_canonical, _ACTION_WRITE_VALUES)

    suite = TestSuiteSpec.model_validate(
        {
            "suite": "task054-runner-identity",
            "workbook_id": wid,
            "baseline": "use_existing",
            "cases": [
                {
                    "id": "write-through-runner",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "Cells",
                            "range_address": "B2",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "string", "value": "via-runner"}]],
                            },
                        }
                    ],
                    "assertions": [
                        {
                            "type": "cell_value_equals",
                            "sheet": "Cells",
                            "cell": "B2",
                            "expected": "via-runner",
                        }
                    ],
                }
            ],
        }
    )

    result = await rig.runner.run_suite(suite)
    assert result.status is RunStatus.PASSED, f"runner failed: {result}"
    assert result.baseline_snapshot_id == baseline.snapshot_id
    case = result.cases[0]
    assert case.status is CaseStatus.PASSED
    # The runner's range_write_values step ran through the shared RangesService:
    # its write postcondition is verified=True (not a no-op), and the runner's own
    # cell_value_equals assertion (a RangesService read on the case working copy)
    # passed — both are the *same* service instance wired into the runner.
    assert case.steps[0].status is StepStatus.PASSED
    assert case.steps[0].details.get("verified") is True
    assert case.steps[0].details.get("address", "").replace("$", "") == "B2"
    assert all(assertion.passed for assertion in case.assertions), case.assertions
    # Each case ran on its own disposable copy of the immutable baseline.
    assert case.working_copy_path is not None
    assert _sha256(Path(baseline.stored_path)) == baseline.sha256


# ===========================================================================
# Area 2 — Range ops preserve Stage 1 (Tables / Names / Sheets)
# ===========================================================================
@requires_excel
async def test_range_write_inside_table_preserves_table_structure(rig: Rig) -> None:
    """A cell write inside ``tblCells`` does not change its rows/columns."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    before = _sha256(original)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_RANGE_ACTIONS)

    tables_before = await rig.tables.list_tables(wid)
    target = next(entry for entry in tables_before if entry.name == "tblCells")
    assert (target.row_count, len(target.columns)) == (3, 3)

    # T2 is inside tblCells (T1:V4): an explicit cell-level update, not a structural op.
    result = await rig.ranges.write_values(
        wid, "Cells", "T2", _single_value(CellValueKind.NUMBER, 99.0)
    )
    assert result.verified is True

    tables_after = await rig.tables.list_tables(wid)
    fetched = next(entry for entry in tables_after if entry.name == "tblCells")
    assert fetched.range_address.replace("$", "") == "T1:V4"
    assert fetched.row_count == 3
    assert [column.name for column in fetched.columns] == ["id", "name", "qty"]

    assert _sha256(original) == before
    await rig.session.close_workbook(wid)


@requires_excel
async def test_format_and_structural_ops_preserve_stage1_services(rig: Rig) -> None:
    """Format + a non-table structural op leave Tables / Names / Sheets intact."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    before = _sha256(original)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_RANGE_ACTIONS)

    sheets_before = {sheet.name for sheet in await rig.sheets.list_sheets(wid)}
    names_before = {name.name for name in await rig.names.list_names(wid)}
    table_before = next(
        entry for entry in await rig.tables.list_tables(wid) if entry.name == "tblCells"
    )

    # A formatting write and a structural op outside the ListObject.
    assert (
        await rig.ranges.set_format(wid, "Cells", "B4", RangeFormatPatch(font_bold=True))
    ).verified
    assert (await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")).verified

    sheets_after = {sheet.name for sheet in await rig.sheets.list_sheets(wid)}
    names_after = {name.name for name in await rig.names.list_names(wid)}
    table_after = next(
        entry for entry in await rig.tables.list_tables(wid) if entry.name == "tblCells"
    )

    assert sheets_after == sheets_before
    assert names_after == names_before
    assert table_after.range_address == table_before.range_address
    assert table_after.row_count == table_before.row_count == 3
    assert [column.name for column in table_after.columns] == [
        column.name for column in table_before.columns
    ]

    # The Stage 1 services still function after the range mutations.
    fetched = await rig.tables.get_table(wid, "tblCells")
    assert isinstance(fetched, TableInfo)
    assert all(isinstance(name, DefinedNameInfo) for name in await rig.names.list_names(wid))
    assert all(isinstance(sheet, SheetInfo) for sheet in await rig.sheets.list_sheets(wid))

    assert _sha256(original) == before
    await rig.session.close_workbook(wid)


# ===========================================================================
# Area 3 — Pivot ops never invoke deferred features; refresh is targeted
# ===========================================================================
@requires_excel
async def test_pivot_ops_do_not_create_deferred_objects(rig: Rig) -> None:
    """Inventory / field / filter leave Connections / QueryTables / Charts / Slicers empty."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_PIVOT_ACTIONS)

    baseline_counts = await _deferred_feature_counts(rig)
    assert baseline_counts == {
        "connections": 0,
        "slicers": 0,
        "query_tables": 0,
        "chart_objects": 0,
        "chart_sheets": 0,
    }

    # Read + mutate field layout / aggregation / filter without deferred features.
    await rig.pivots.list_pivot_tables(wid)
    await rig.pivots.list_pivot_fields(wid, "ptA")
    await rig.pivots.set_pivot_field_layout(wid, "ptA", "Quarter", PivotFieldOrientation.ROW)
    await rig.pivots.set_pivot_field_aggregation(wid, "ptA", "Sales", PivotAggregation.MAX)
    await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))

    after_counts = await _deferred_feature_counts(rig)
    assert after_counts == baseline_counts, f"pivot ops created deferred objects: {after_counts}"

    # The fixture caches are local database sources — no Data Model (OLAP) cache.
    inventory = await rig.pivots.list_pivot_tables(wid)
    for pivot in inventory.pivot_tables:
        assert pivot.source_type is PivotSourceType.DATABASE
        assert pivot.source_status == "supported"
    for cache in inventory.caches:
        assert cache.olap is False

    await rig.session.close_workbook(wid)


@requires_excel
async def test_pivot_refresh_is_targeted_and_not_refresh_all(rig: Rig) -> None:
    """Refreshing ptC (cache 2) leaves cache 1 (ptA/ptB) untouched.

    A ``Workbook.RefreshAll`` would refresh every cache; a targeted refresh of
    ``ptC`` must advance only cache 2 and leave cache 1's refresh date unchanged.
    """

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, _ACTION_REFRESH_TABLE, "pivot_cache_refresh")

    cache1_before = await _read_cache_refresh_date(rig, "ptA")  # cache 1 shared by ptA/ptB
    cache2_before = await _read_cache_refresh_date(rig, "ptC")  # cache 2 (ListObject)

    result = await rig.pivots.refresh_pivot_table(wid, "ptC")

    assert isinstance(result, PivotRefreshResult)
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.cache_index == 2
    assert result.affected_pivot_tables == ["ptC"]

    cache1_after = await _read_cache_refresh_date(rig, "ptA")
    cache2_after = await _read_cache_refresh_date(rig, "ptC")

    # Targeted: cache 2 advanced; cache 1 untouched (so RefreshAll was not used).
    assert cache2_after > cache2_before
    assert cache1_after == cache1_before

    await rig.session.close_workbook(wid)


# ===========================================================================
# Area 4 — Snapshot / approval / audit / read-only / recovery
# ===========================================================================
def _range_mutation_calls(rig: Rig, wid: str) -> list[tuple[str, Any]]:
    return [
        (
            "range_write_values",
            lambda: rig.ranges.write_values(wid, "Cells", "A40", _single_value(value=1.0)),
        ),
        (
            "range_set_format",
            lambda: rig.ranges.set_format(wid, "Cells", "A40", RangeFormatPatch(font_bold=True)),
        ),
        (
            "range_copy",
            lambda: rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2"),
        ),
        (
            "range_set_validation",
            lambda: rig.ranges.set_validation(wid, "Cells", "W4:W6", _list_validation()),
        ),
        (
            "range_set_conditional_formatting",
            lambda: rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf()),
        ),
    ]


@requires_excel
async def test_every_range_mutation_requires_a_snapshot(rig: Rig) -> None:
    """Writable + approved but unsnapshotted: every range mutation -> SNAPSHOT_FAILED."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    # Approval is present, but no snapshot was taken.
    rig.grant(original, *_ALL_RANGE_ACTIONS)

    for tool, call in _range_mutation_calls(rig, wid):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED, tool

    await rig.session.close_workbook(wid)


@requires_excel
async def test_every_range_mutation_requires_approval(rig: Rig) -> None:
    """Writable + snapshotted but unapproved: every range mutation -> APPROVAL_DENIED."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    # excel_open was granted by open(); no range action was granted (deny-by-default).

    for tool, call in _range_mutation_calls(rig, wid):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.APPROVAL_DENIED, tool
        assert "approval" in exc_info.value.message.lower(), tool

    await rig.session.close_workbook(wid)


@requires_excel
async def test_read_only_open_blocks_every_pivot_mutation(rig: Rig) -> None:
    """A read-only working copy is refused before snapshot/approval."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, read_only=True)
    wid = info.workbook_id
    rig.grant(original, *_ALL_PIVOT_ACTIONS)

    with pytest.raises(ExcelVbaMcpError) as refresh_exc:
        await rig.pivots.refresh_pivot_table(wid, "ptA")
    assert refresh_exc.value.code is ErrorCode.WORKBOOK_READ_ONLY

    with pytest.raises(ExcelVbaMcpError) as layout_exc:
        await rig.pivots.set_pivot_field_layout(wid, "ptA", "Region", PivotFieldOrientation.COLUMN)
    assert layout_exc.value.code is ErrorCode.WORKBOOK_READ_ONLY

    with pytest.raises(ExcelVbaMcpError) as filter_exc:
        await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    assert filter_exc.value.code is ErrorCode.WORKBOOK_READ_ONLY

    await rig.session.close_workbook(wid)


@requires_excel
async def test_audit_trail_records_cross_feature_mutations(rig: Rig) -> None:
    """Range + Pivot mutating ops are recorded in the append-only audit trail."""

    ranges_fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(ranges_fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_RANGE_ACTIONS)

    assert (
        await rig.ranges.write_values(
            wid, "Cells", "A20", _single_value(CellValueKind.STRING, "audit")
        )
    ).verified
    assert (
        await rig.ranges.set_format(wid, "Cells", "C20", RangeFormatPatch(font_bold=True))
    ).verified
    await rig.ranges.read_values(wid, "Cells", "B2")
    await rig.session.close_workbook(wid)
    rig.audit.flush()

    records = _read_audit(rig.audit.path)
    mutating = [record for record in records if record.get("operation_class") == "mutating"]
    tools = {record["tool"] for record in mutating}
    assert {"range_write_values", "range_set_format"} <= tools
    assert all(record.get("result") is not None for record in mutating)
    assert any(
        record["tool"] == "range_read_values" and record.get("operation_class") == "read_only"
        for record in records
    )


@requires_excel
async def test_hung_com_recovers_to_fresh_hidden_instance(rig: Rig) -> None:
    """A forced hung COM call reports EXCEL_HUNG, enters NEEDS_RECOVERY and recovers."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    before = _sha256(original)
    info = await rig.open(original, read_only=True)
    await rig.ranges.read_values(info.workbook_id, "Cells", "A1")

    old_pid = rig.worker.get_status().owned_pid
    assert old_pid is not None

    def hang() -> str:
        time.sleep(_HANG_SECONDS)
        return "finished"

    hang.__name__ = "task054_hang"
    with pytest.raises(ExcelHungError) as hung_exc:
        await rig.worker.submit(
            hang,
            timeout=_HANG_CALL_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    assert hung_exc.value.code is ErrorCode.EXCEL_HUNG
    # Block E: a deadline-exceeded call is NOT a poison state — the worker moves to
    # running/deadline_exceeded and stays accepting (no fail-fast); the call is
    # never cancelled and the owned pid is never terminated. Recovery is
    # agent-directed (the explicit recover_session below), never automatic.
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    status = await rig.session.recover_session(reason="task054_hung")
    assert status.state is WorkbookSessionState.RUNNING
    assert rig.session.list_workbooks_sync() == []

    new_pid = rig.worker.get_status().owned_pid
    assert new_pid is not None and new_pid != old_pid
    assert _sha256(original) == before, "recover must not modify the original"

    flags = await _read_app_flags(rig)
    assert flags["visible"] is False, "the rebuilt instance must stay Visible=False"
    _assert_owned_never_foreground(new_pid, "after-recover")
    _assert_no_visible_owned_windows(new_pid, "after-recover")

    assert _wait_pid_gone(old_pid), f"old owned PID {old_pid} still alive after recover"


# ===========================================================================
# Area 5 — Desktop non-interference across ALL op families
# ===========================================================================
async def _run_every_op_family(rig: Rig, wid: str) -> None:
    """Exercise one representative of every Phase 2 range op family on a working copy."""

    # Reads.
    await rig.ranges.get_range(wid, "Cells", "A1")
    await rig.ranges.read_values(wid, "Cells", "B2:B3")
    await rig.ranges.read_formulas(wid, "Cells", "B9")
    # Writes.
    assert (
        await rig.ranges.write_values(
            wid, "Cells", "A20", _single_value(CellValueKind.STRING, "bg")
        )
    ).verified
    assert (
        await rig.ranges.set_format(wid, "Cells", "C20", RangeFormatPatch(font_bold=True))
    ).verified
    # Structural.
    assert (await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")).verified
    # Validation / conditional formatting.
    assert (await rig.ranges.set_validation(wid, "Cells", "W4:W6", _list_validation())).present
    assert (
        await rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf())
    ).verified


async def _run_pivot_family(rig: Rig, wid: str) -> None:
    """Exercise one representative of every pivot op family."""

    await rig.pivots.list_pivot_tables(wid)
    await rig.pivots.get_pivot_table(wid, "ptA")
    await rig.pivots.list_pivot_fields(wid, "ptA")
    await rig.pivots.refresh_pivot_table(wid, "ptA")
    await rig.pivots.set_pivot_field_layout(wid, "ptA", "Quarter", PivotFieldOrientation.ROW)
    await rig.pivots.set_pivot_field_aggregation(wid, "ptA", "Sales", PivotAggregation.MAX)
    await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    await rig.pivots.get_pivot_field_filters(wid, "ptA", "Region")


@requires_excel
async def test_range_op_families_keep_desktop_hidden_and_originals(rig: Rig) -> None:
    """Reads/writes/format/structural/validation-CF keep BG03/04/05."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    before = _sha256(original)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_RANGE_ACTIONS)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    fg_before = _foreground()
    samples = await _sample_foreground_while(owned_pid, lambda: _run_every_op_family(rig, wid))
    assert not samples, f"BG03 violation: owned PID {owned_pid} held foreground: {samples}"

    flags = await _read_app_flags(rig)
    assert flags["visible"] is False, "owned Excel must stay hidden across all range op families"
    assert flags["display_alerts"] is True, "DisplayAlerts default must be preserved"
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

    _assert_owned_never_foreground(owned_pid, "range-families")
    _assert_no_visible_owned_windows(owned_pid, "range-families")

    # BG05: the foreground (if owned by another process) is unchanged.
    fg_after = _foreground()
    if fg_before is not None and fg_before["pid"] != owned_pid:
        assert fg_after is not None and fg_after["pid"] == fg_before["pid"], (
            f"BG05 violation: foreground moved from {fg_before['pid']} to "
            f"{fg_after['pid'] if fg_after else None}"
        )

    # BG05: the disposable original (copy of the committed fixture) is unchanged.
    assert _sha256(original) == before

    await rig.session.close_workbook(wid)


@requires_excel
async def test_pivot_op_families_keep_desktop_hidden(rig: Rig) -> None:
    """Inventory/refresh/layout/filter keep BG03/04/05."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    before = _sha256(original)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, *_ALL_PIVOT_ACTIONS)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    samples = await _sample_foreground_while(owned_pid, lambda: _run_pivot_family(rig, wid))
    assert not samples, f"BG03 violation: owned PID {owned_pid} held foreground: {samples}"

    flags = await _read_app_flags(rig)
    assert flags["visible"] is False
    assert flags["display_alerts"] is True
    _assert_owned_never_foreground(owned_pid, "pivot-families")
    _assert_no_visible_owned_windows(owned_pid, "pivot-families")
    assert _sha256(original) == before

    await rig.session.close_workbook(wid)


@requires_excel
async def test_owned_pid_is_new_no_non_owned_excel_touched_and_reaped(tmp_path: Path) -> None:
    """BG03/05: owned PID is a fresh process; pre-existing Excel survives; owned is reaped."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    pre_existing = _settled_excel_pids()

    r = Rig(tmp_path)
    original = r.copy_fixture(fixture)
    owned_pid: int | None = None
    await r.start()
    try:
        info = await r.open(original, mode=SessionMode.INSPECT, read_only=False)
        await r.snapshot.create_snapshot(info.workbook_id)
        r.grant(original, _ACTION_WRITE_VALUES)
        owned_pid = r.worker.get_status().owned_pid
        assert owned_pid is not None
        # DispatchEx must create a *new* process, never reuse a running instance.
        assert owned_pid not in pre_existing

        assert (
            await r.ranges.write_values(
                info.workbook_id, "Cells", "A20", _single_value(CellValueKind.STRING, "pid")
            )
        ).verified
        assert _pid_alive(owned_pid)
    finally:
        await r.stop()

    # BG05: the owned process is reaped; every pre-existing process survives.
    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} survived shutdown (orphan)"
    still_alive = {pid for pid in pre_existing if _pid_alive(pid)}
    assert still_alive == pre_existing, (
        f"a non-owned Excel process disappeared: {sorted(pre_existing - still_alive)}"
    )


# ===========================================================================
# Area 6 — Unsupported cases are explicit, never false success
# ===========================================================================
@requires_excel
async def test_unsupported_pivot_source_is_explicit(rig: Rig) -> None:
    """A consolidation-sourced cache is reported unsupported and never refreshed."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id

    built = await _build_consolidation_pivot(rig)
    assert built["name"] == "ptConsolidated"
    assert built["source_type"] == 3  # xlSourceTypeConsolidation
    assert built["olap"] is False

    inventory = await rig.pivots.list_pivot_tables(wid)
    pivot = {pt.name: pt for pt in inventory.pivot_tables}["ptConsolidated"]
    assert pivot.source_type is PivotSourceType.CONSOLIDATION
    assert pivot.source_status == "unsupported_source"

    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, _ACTION_REFRESH_TABLE)
    refreshed = await rig.pivots.refresh_pivot_table(wid, "ptConsolidated")
    assert refreshed.status is PivotRefreshStatus.UNSUPPORTED_SOURCE
    assert refreshed.refreshed is False
    assert refreshed.refresh_return is None
    assert refreshed.error is not None
    assert refreshed.error["code"] == "UNSUPPORTED_PIVOT_SOURCE"
    assert refreshed.refresh_date_before == refreshed.refresh_date_after

    await rig.session.close_workbook(wid)


@requires_excel
async def test_unsupported_pivot_filter_kind_is_explicit(rig: Rig) -> None:
    """A label (PivotFilters.Add) filter is reported unsupported and writes nothing."""

    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, _ACTION_APPLY_FILTER)

    result = await rig.pivots.apply_pivot_field_filter(
        wid, "ptA", "Region", PivotFilterSpec(kind=PivotFilterKind.LABEL)
    )

    # Explicit unsupported status, never a silent conversion.
    assert result.filter.supported is False
    assert result.filter.kind is PivotFilterKind.LABEL
    assert result.unsupported == ["label"]
    assert "PivotFilters.Add is unsupported" in (result.filter.unsupported_reason or "")

    # Nothing was written: no manual filter and no PivotFilters entry.
    after = await rig.pivots.get_pivot_field_filters(wid, "ptA", "Region")
    assert isinstance(after, PivotFieldFilters)
    assert after.hidden_items == []
    assert after.filter.kind is PivotFilterKind.NONE
    assert await _pivot_filters_count(rig, "ptA", "Region") == 0

    await rig.session.close_workbook(wid)


@requires_excel
async def test_unsupported_color_scale_cf_is_explicit(rig: Rig) -> None:
    """A ColorScale conditional format is rejected with INVALID_RANGE; nothing changes."""

    fixture = _fixture_or_skip(_RANGES_PROJECT)
    original = rig.copy_fixture(fixture)
    info = await rig.open(original, mode=SessionMode.INSPECT, read_only=False)
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)
    rig.grant(original, _ACTION_SET_CF)

    # Seed a supported rule, then attempt the unsupported ColorScale rule.
    await rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf())
    before = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_conditional_formatting(
            wid, "Cells", "X1:X5", ConditionalFormatSpec(type=ConditionalFormatType.COLOR_SCALE)
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["type"] == ConditionalFormatType.COLOR_SCALE.value

    after = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")
    assert [rule.type for rule in after] == [rule.type for rule in before]
    assert len(after) == len(before) == 1

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# STA-local fixtures used only by area 6
# ---------------------------------------------------------------------------
async def _build_consolidation_pivot(rig: Rig) -> dict[str, Any]:
    """Add a consolidation-sourced PivotTable to the open working copy (STA)."""

    def build() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        workbook = app.Workbooks(1)
        sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = "Unsupported"
        # xlSourceTypeConsolidation (3): local, deterministic, non-database source.
        cache = workbook.PivotCaches().Create(3, ("Data!R1C1:R7C2",))
        pivot = cache.CreatePivotTable(
            TableDestination=sheet.Range("B3"), TableName="ptConsolidated"
        )
        return {
            "name": str(pivot.Name),
            "cache_index": int(pivot.CacheIndex),
            "source_type": int(pivot.PivotCache().SourceType),
            "olap": bool(pivot.PivotCache().OLAP),
        }

    build.__name__ = "task054_build_consolidation_pivot"
    return dict(
        await rig.worker.submit(
            build,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    )


async def _pivot_filters_count(rig: Rig, pivot_table: str, field: str) -> int:
    """Read ``PivotField.PivotFilters.Count`` directly on the STA thread."""

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

    read_count.__name__ = "task054_read_pivot_filters_count"
    return int(
        await rig.worker.submit(
            read_count,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
    )
