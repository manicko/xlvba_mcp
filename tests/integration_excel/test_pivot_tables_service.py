"""Independent end-to-end verification of the PivotTables surface (task_052).

This module is the *independent* acceptance test for the Phase 2 PivotTables
feature. It drives the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` (and,
for one workflow, the public :mod:`excel_vba_mcp.api.pivot_tools` contract
functions) against a **hidden, automation-created Excel instance** and a
disposable working copy of the committed ``PivotTablesProject.xlsm`` fixture.
The assertions are written from the fixture manifest / observed Excel state, not
copied from the implementation tests:

* **Inventory** — all three PivotTables (``ptA`` / ``ptB`` / ``ptC``) with their
  sheet, 1-based sheet index, stable name, ``TableRange1`` / ``TableRange2`` /
  ``DataBodyRange`` geometry, cache indices, row/column/page/data field layout,
  data-field aggregation and manual filter state;
* **Cache sharing** — ``ptA`` / ``ptB`` share cache 1 (``shared_cache`` and the
  ``affected_pivot_tables`` list), ``ptC`` is standalone on cache 2, and every
  fixture cache is a local ``database`` source;
* **Field configuration** — orientation / position / aggregation / caption and
  number format round-trip through a fresh read, removal is non-destructive,
  and an invalid orientation / position / legacy aggregation code is rejected;
* **Filters** — a manual filter applies and reads back, ``field`` vs ``table``
  clear scopes differ, and an unsupported ``PivotFilters.Add`` kind is reported
  explicitly without touching Excel;
* **Unsupported sources** — a real consolidation-sourced PivotTable is built in
  the disposable working copy and rejected with ``UNSUPPORTED_PIVOT_SOURCE``;
* **Guards / audit** — read-only, missing-snapshot and missing-approval mutations
  are refused, and every accepted mutation is recorded as ``mutating`` in the
  append-only audit trail;
* **BG03 / BG04 / BG05** — the owned Excel stays hidden, never holds foreground,
  exposes no visible window, and the committed fixture is byte-identical.

All Excel access is serialised through the single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`; no
``.Activate`` / ``.Select`` is ever called and the owned instance is always
``Visible = False``.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_tables_service.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import json
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
from excel_vba_mcp.api import pivot_tools
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFieldProperties,
    PivotFilterSpec,
    PivotInventory,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotRefreshStatus,
    PivotSourceType,
    SessionMode,
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

#: Approval action names (stable identities of the pivot mutation surface).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_REFRESH_CACHE = "pivot_cache_refresh"
_ACTION_SET_LAYOUT = "pivot_field_set_layout"
_ACTION_SET_POSITION = "pivot_field_set_position"
_ACTION_SET_AGGREGATION = "pivot_field_set_aggregation"
_ACTION_SET_PROPERTIES = "pivot_field_set_properties"
_ACTION_APPLY_FILTER = "pivot_field_apply_filter"
_ACTION_CLEAR_FILTERS = "pivot_field_clear_filters"

#: Every mutating pivot action, granted by :meth:`Rig.open_for_mutation`.
_MUTATING_ACTIONS = (
    _ACTION_REFRESH_TABLE,
    _ACTION_REFRESH_CACHE,
    _ACTION_SET_LAYOUT,
    _ACTION_SET_POSITION,
    _ACTION_SET_AGGREGATION,
    _ACTION_SET_PROPERTIES,
    _ACTION_APPLY_FILTER,
    _ACTION_CLEAR_FILTERS,
)

#: Expected fixture geometry (recorded in ``tests/fixtures/phase2_manifest.json``
#: and re-observed independently here from the live Excel object model).
_EXPECTED_PIVOTS: dict[str, dict[str, Any]] = {
    "ptA": {
        "cache_index": 1,
        "table_range1": "$A$3:$F$14",
        "table_range2": "$A$1:$F$14",
        "data_body_range": "$C$5:$F$14",
    },
    "ptB": {
        "cache_index": 1,
        "table_range1": "$J$3:$K$8",
        "table_range2": "$J$3:$K$8",
        "data_body_range": "$K$5:$K$8",
    },
    "ptC": {
        "cache_index": 2,
        "table_range1": "$S$3:$T$8",
        "table_range2": "$S$3:$T$8",
        "data_body_range": "$T$5:$T$8",
    },
}


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
    visible = [window for window in _windows_for_pid(owned_pid) if window["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at "
        f"{label}: {[window['title'] for window in visible]}"
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _by_name(fields: list[PivotFieldInfo]) -> dict[str, PivotFieldInfo]:
    return {field.source_name: field for field in fields}


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


def _manual(*items: str) -> PivotFilterSpec:
    return PivotFilterSpec(kind=PivotFilterKind.MANUAL, visible_items=list(items))


# ---------------------------------------------------------------------------
# Fixtures / rig — real COM worker + Excel + pivot service
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + PivotTablesService for one verification run."""

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
        """Open writable and grant every pivot mutation action."""

        info = await self.open(read_only=False)
        for action in _MUTATING_ACTIONS:
            self.grant(action)
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
# 1. Inventory: all three pivots, geometry, cache indices
# ---------------------------------------------------------------------------
@requires_excel
async def test_inventory_reports_all_pivots_with_fixture_geometry(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    inventory = await rig.pivots.list_pivot_tables(info.workbook_id)
    assert isinstance(inventory, PivotInventory)

    # The worksheet sweep finds every PivotTable (Workbook.PivotTables is empty
    # on the target build).
    assert [pt.name for pt in inventory.pivot_tables] == ["ptA", "ptB", "ptC"]

    by_name: dict[str, PivotTableInfo] = {pt.name: pt for pt in inventory.pivot_tables}
    for name, expected in _EXPECTED_PIVOTS.items():
        pivot = by_name[name]
        assert pivot.sheet_name == "PivotReport"
        assert pivot.sheet_index == 3  # Data, TableData, PivotReport
        assert pivot.cache_index == expected["cache_index"]
        assert pivot.table_range1 == expected["table_range1"]
        assert pivot.table_range2 == expected["table_range2"]
        assert pivot.data_body_range == expected["data_body_range"]
        assert pivot.row_grand is True
        assert pivot.column_grand is True
        assert pivot.source_type is PivotSourceType.DATABASE
        assert pivot.source_status == "supported"
        assert pivot.refresh_date is not None

    # Two distinct caches back the three tables (shared cache 1 + standalone 2).
    assert [cache.cache_index for cache in inventory.caches] == [1, 2]

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# 2. Field layout / aggregation / filter state
# ---------------------------------------------------------------------------
@requires_excel
async def test_field_layout_aggregation_and_filter_state(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    pta = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptA"))
    assert [
        field.source_name
        for field in pta.values()
        if field.orientation is PivotFieldOrientation.ROW
    ] == ["Region"]
    assert pta["Region"].position == 1
    assert pta["Product"].orientation is PivotFieldOrientation.COLUMN
    assert pta["Quarter"].orientation is PivotFieldOrientation.PAGE
    assert pta["Sales"].orientation is PivotFieldOrientation.DATA
    assert pta["Units"].orientation is PivotFieldOrientation.DATA
    assert pta["Sales"].is_data_field is True
    assert pta["Units"].is_data_field is True
    assert pta["Sales"].function is PivotAggregation.SUM
    assert pta["Sales"].data_field_name == "Sum of Sales"
    assert pta["Units"].function is PivotAggregation.AVERAGE
    assert pta["Units"].data_field_name == "Average of Units"

    ptb = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptB"))
    assert [
        field.source_name
        for field in ptb.values()
        if field.orientation is PivotFieldOrientation.ROW
    ] == ["Region"]
    assert ptb["Region"].all_items_visible is False
    assert ptb["Region"].filter_state is PivotFilterKind.MANUAL
    assert ptb["Units"].function is PivotAggregation.SUM
    assert ptb["Units"].data_field_name == "Sum of Units"

    ptc = _by_name(await rig.pivots.list_pivot_fields(info.workbook_id, "ptC"))
    assert ptc["Product"].orientation is PivotFieldOrientation.ROW
    assert ptc["Sales"].function is PivotAggregation.COUNT
    assert ptc["Sales"].data_field_name == "Count of Sales"

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# 3. Cache sharing / metadata / source type
# ---------------------------------------------------------------------------
@requires_excel
async def test_shared_cache_identity_and_standalone_cache(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    pta = await rig.pivots.get_pivot_table(info.workbook_id, "ptA")
    ptb = await rig.pivots.get_pivot_table(info.workbook_id, "ptB")
    ptc = await rig.pivots.get_pivot_table(info.workbook_id, "ptC")

    # ptA / ptB share cache 1; ptC is alone on cache 2.
    assert pta.cache_index == 1
    assert ptb.cache_index == 1
    assert ptc.cache_index == 2
    assert pta.cache_index == ptb.cache_index

    shared = await rig.pivots.get_pivot_cache(info.workbook_id, 1)
    assert isinstance(shared, PivotCacheInfo)
    assert shared.cache_index == 1
    assert shared.shared_cache is True
    assert shared.affected_pivot_tables == ["ptA", "ptB"]
    assert shared.source_type is PivotSourceType.DATABASE
    assert shared.source_data_kind == "text"
    assert shared.source_data is not None
    assert "Data!" in shared.source_data
    assert shared.olap is False
    assert shared.refresh_date is not None

    standalone = await rig.pivots.get_pivot_cache(info.workbook_id, 2)
    assert standalone.cache_index == 2
    assert standalone.shared_cache is False
    assert standalone.affected_pivot_tables == ["ptC"]
    assert standalone.source_type is PivotSourceType.DATABASE
    assert standalone.source_data is not None
    assert "TableData!" in standalone.source_data
    assert standalone.olap is False

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# 4. Public API tool layer (pivot_tools) reports the same inventory
# ---------------------------------------------------------------------------
@requires_excel
async def test_api_tools_layer_reports_the_same_inventory(rig: Rig) -> None:
    info = await rig.open(read_only=True)

    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        inventory = await pivot_tools.pivot_tables_list(info.workbook_id)
        pivot = await pivot_tools.pivot_table_get(info.workbook_id, "ptA")
        cache = await pivot_tools.pivot_cache_get(info.workbook_id, 1)
        fields = await pivot_tools.pivot_fields_list(info.workbook_id, "ptA")
    finally:
        pivot_tools.set_pivot_tables_service(None)

    assert isinstance(inventory, PivotInventory)
    assert [pt.name for pt in inventory.pivot_tables] == ["ptA", "ptB", "ptC"]
    assert isinstance(pivot, PivotTableInfo)
    assert pivot.name == "ptA"
    assert pivot.cache_index == 1
    assert isinstance(cache, PivotCacheInfo)
    assert cache.shared_cache is True
    assert cache.affected_pivot_tables == ["ptA", "ptB"]
    assert {field.source_name for field in fields} >= {"Region", "Product", "Quarter", "Sales"}

    await rig.session.close_workbook(info.workbook_id)


# ---------------------------------------------------------------------------
# 5. Field configuration round-trip
# ---------------------------------------------------------------------------
@requires_excel
async def test_field_configuration_round_trip(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    # Layout: ptB has a single data field, so every orientation transition is
    # accepted by Excel (no synthetic "Data" placeholder locking the row axis).
    moved = await rig.pivots.set_pivot_field_layout(
        wid, "ptB", "Region", PivotFieldOrientation.COLUMN
    )
    assert moved.orientation is PivotFieldOrientation.COLUMN
    assert _by_name(await rig.pivots.list_pivot_fields(wid, "ptB"))["Region"].orientation is (
        PivotFieldOrientation.COLUMN
    )

    moved = await rig.pivots.set_pivot_field_layout(
        wid, "ptB", "Region", PivotFieldOrientation.PAGE
    )
    assert moved.orientation is PivotFieldOrientation.PAGE
    assert _by_name(await rig.pivots.list_pivot_fields(wid, "ptB"))["Region"].orientation is (
        PivotFieldOrientation.PAGE
    )

    # Position: put Region back on the row axis, move Product to the row axis
    # (position 2) and then promote Product to position 1.
    await rig.pivots.set_pivot_field_layout(wid, "ptB", "Region", PivotFieldOrientation.ROW)
    await rig.pivots.set_pivot_field_layout(wid, "ptB", "Product", PivotFieldOrientation.ROW)
    product = _by_name(await rig.pivots.list_pivot_fields(wid, "ptB"))["Product"]
    assert product.position == 2
    moved = await rig.pivots.set_pivot_field_position(wid, "ptB", "Product", 1)
    assert moved.position == 1
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptB"))
    assert fields["Product"].position == 1
    assert fields["Region"].position == 2

    # Aggregation: sum -> average -> count on the shared-cache ptA.
    summed = await rig.pivots.set_pivot_field_aggregation(wid, "ptA", "Sales", PivotAggregation.SUM)
    assert summed.function is PivotAggregation.SUM
    averaged = await rig.pivots.set_pivot_field_aggregation(
        wid, "ptA", "Sales", PivotAggregation.AVERAGE
    )
    assert averaged.function is PivotAggregation.AVERAGE
    counted = await rig.pivots.set_pivot_field_aggregation(
        wid, "ptA", "Units", PivotAggregation.COUNT
    )
    assert counted.function is PivotAggregation.COUNT
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptA"))
    assert fields["Sales"].function is PivotAggregation.AVERAGE
    assert fields["Units"].function is PivotAggregation.COUNT

    # Caption / number format readback.
    updated = await rig.pivots.set_pivot_field_properties(
        wid, "ptA", "Sales", PivotFieldProperties(caption="Total Sales", number_format="0.00")
    )
    assert updated.data_field_name == "Total Sales"
    assert updated.number_format == "0.00"
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptA"))
    assert fields["Sales"].data_field_name == "Total Sales"
    assert fields["Sales"].number_format == "0.00"

    # Non-destructive removal: the field stays listed as hidden, position None.
    removed = await rig.pivots.set_pivot_field_layout(
        wid, "ptA", "Product", PivotFieldOrientation.HIDDEN
    )
    assert removed.orientation is PivotFieldOrientation.HIDDEN
    assert removed.position is None
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptA"))
    assert "Product" in fields
    assert fields["Product"].orientation is PivotFieldOrientation.HIDDEN
    assert fields["Product"].position is None

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 6. Invalid field requests are rejected without mutating Excel
# ---------------------------------------------------------------------------
@requires_excel
async def test_invalid_field_requests_are_rejected(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    # ``orientation="data"`` never creates a data field -> rejected up front.
    with pytest.raises(ExcelVbaMcpError) as layout_exc:
        await rig.pivots.set_pivot_field_layout(wid, "ptA", "Units", PivotFieldOrientation.DATA)
    assert layout_exc.value.code is ErrorCode.INVALID_RANGE
    assert layout_exc.value.context["reason"] == "invalid_orientation"

    # An out-of-range 1-based position is rejected and leaves the field in place.
    with pytest.raises(ExcelVbaMcpError) as position_exc:
        await rig.pivots.set_pivot_field_position(wid, "ptA", "Region", 99)
    assert position_exc.value.code is ErrorCode.INVALID_RANGE
    assert position_exc.value.context["reason"] == "position_out_of_range"
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptA"))
    assert fields["Region"].position == 1

    # A legacy aggregation code is refused at the public tool boundary.
    pivot_tools.set_pivot_tables_service(rig.pivots)
    try:
        with pytest.raises(ExcelVbaMcpError) as legacy_exc:
            await pivot_tools.pivot_field_set_aggregation(wid, "ptA", "Sales", 9)  # type: ignore[arg-type]
    finally:
        pivot_tools.set_pivot_tables_service(None)
    assert legacy_exc.value.code is ErrorCode.INVALID_RANGE
    assert legacy_exc.value.context["reason"] == "invalid_function"

    # A duplicate data-field caption is ambiguous and rejected.
    fields = _by_name(await rig.pivots.list_pivot_fields(wid, "ptA"))
    taken = fields["Units"].data_field_name
    assert taken
    with pytest.raises(ExcelVbaMcpError) as duplicate_exc:
        await rig.pivots.set_pivot_field_properties(
            wid, "ptA", "Sales", PivotFieldProperties(caption=taken)
        )
    assert duplicate_exc.value.code is ErrorCode.AMBIGUOUS_FILTER
    assert duplicate_exc.value.context["reason"] == "duplicate_caption"

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 7. Filters: manual round-trip, clear scopes, unsupported kind
# ---------------------------------------------------------------------------
@requires_excel
async def test_manual_filter_round_trip_and_clear_scopes(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    applied = await rig.pivots.apply_pivot_field_filter(
        wid, "ptA", "Region", _manual("North", "South")
    )
    assert applied.filter.kind is PivotFilterKind.MANUAL
    assert set(applied.visible_items) == {"North", "South"}
    assert set(applied.hidden_items) == {"East", "West"}
    assert applied.all_items_visible is False

    readback = await rig.pivots.get_pivot_field_filters(wid, "ptA", "Region")
    assert isinstance(readback, PivotFieldFilters)
    assert set(readback.visible_items) == {"North", "South"}
    assert set(readback.hidden_items) == {"East", "West"}
    assert readback.filter.kind is PivotFilterKind.MANUAL

    # A sibling PivotTable on the same shared cache keeps its own filter state.
    sibling = await rig.pivots.get_pivot_field_filters(wid, "ptB", "Region")
    assert set(sibling.visible_items) == {"East", "North", "South"}
    assert sibling.hidden_items == ["West"]

    # scope="field" clears exactly one field.
    cleared = await rig.pivots.clear_pivot_field_filters(wid, "ptA", "Region", scope="field")
    assert cleared.field == "Region"
    assert cleared.hidden_items == []
    assert cleared.all_items_visible is True

    # scope="table" clears the whole PivotTable, not its cache siblings.
    await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    await rig.pivots.apply_pivot_field_filter(wid, "ptB", "Region", _manual("North"))
    table_cleared = await rig.pivots.clear_pivot_field_filters(wid, "ptB", scope="table")
    assert table_cleared.field == ""
    assert table_cleared.hidden_items == []
    assert (await rig.pivots.get_pivot_field_filters(wid, "ptB", "Region")).hidden_items == []
    # ptA shares cache 1 with ptB but keeps its own manual filter.
    assert (await rig.pivots.get_pivot_field_filters(wid, "ptA", "Region")).hidden_items == [
        "East",
        "South",
        "West",
    ]

    await rig.session.close_workbook(wid)


@requires_excel
async def test_unsupported_label_filter_is_explicit(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    result = await rig.pivots.apply_pivot_field_filter(
        wid, "ptA", "Region", PivotFilterSpec(kind=PivotFilterKind.LABEL)
    )

    # The request is reported as unsupported, never silently converted.
    assert result.filter.supported is False
    assert result.filter.kind is PivotFilterKind.LABEL
    assert result.unsupported == ["label"]
    assert "PivotFilters.Add is unsupported" in (result.filter.unsupported_reason or "")

    # Nothing was written: no manual filter and no PivotFilters entry.
    after = await rig.pivots.get_pivot_field_filters(wid, "ptA", "Region")
    assert after.hidden_items == []
    assert after.filter.kind is PivotFilterKind.NONE
    assert after.unsupported == []
    assert await _pivot_filters_count(rig, "ptA", "Region") == 0

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 8. Guard chain + audit trail
# ---------------------------------------------------------------------------
@requires_excel
async def test_read_only_open_blocks_every_mutation(rig: Rig) -> None:
    """A read-only working copy is rejected before snapshot/approval."""

    info = await rig.open(read_only=True)
    wid = info.workbook_id

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
async def test_missing_snapshot_blocks_every_mutation(rig: Rig) -> None:
    """Writable but unsnapshotted: every mutation is refused with SNAPSHOT_FAILED."""

    info = await rig.open_for_mutation()
    wid = info.workbook_id

    with pytest.raises(ExcelVbaMcpError) as refresh_exc:
        await rig.pivots.refresh_pivot_table(wid, "ptA")
    assert refresh_exc.value.code is ErrorCode.SNAPSHOT_FAILED

    with pytest.raises(ExcelVbaMcpError) as layout_exc:
        await rig.pivots.set_pivot_field_layout(wid, "ptA", "Region", PivotFieldOrientation.COLUMN)
    assert layout_exc.value.code is ErrorCode.SNAPSHOT_FAILED

    with pytest.raises(ExcelVbaMcpError) as filter_exc:
        await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    assert filter_exc.value.code is ErrorCode.SNAPSHOT_FAILED

    await rig.session.close_workbook(wid)


@requires_excel
async def test_missing_approval_blocks_every_mutation(rig: Rig) -> None:
    """Writable + snapshotted but unapproved: every mutation is refused (APPROVAL_DENIED)."""

    # Open writable without granting any pivot mutation action.
    rig.grant("excel_open")
    info = await rig.session.open_workbook(
        str(rig.project), mode=SessionMode.INSPECT, read_only=False
    )
    wid = info.workbook_id
    await rig.snapshot.create_snapshot(wid)

    with pytest.raises(ExcelVbaMcpError) as layout_exc:
        await rig.pivots.set_pivot_field_layout(wid, "ptA", "Region", PivotFieldOrientation.COLUMN)
    assert layout_exc.value.code is ErrorCode.APPROVAL_DENIED

    with pytest.raises(ExcelVbaMcpError) as filter_exc:
        await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    assert filter_exc.value.code is ErrorCode.APPROVAL_DENIED

    with pytest.raises(ExcelVbaMcpError) as refresh_exc:
        await rig.pivots.refresh_pivot_table(wid, "ptA")
    assert refresh_exc.value.code is ErrorCode.APPROVAL_DENIED

    await rig.session.close_workbook(wid)


@requires_excel
async def test_audit_trail_records_mutating_operations(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)
    wid = info.workbook_id

    await rig.pivots.refresh_pivot_table(wid, "ptA")
    await rig.pivots.set_pivot_field_layout(wid, "ptA", "Quarter", PivotFieldOrientation.ROW)
    await rig.pivots.apply_pivot_field_filter(wid, "ptA", "Region", _manual("North"))
    await rig.pivots.list_pivot_tables(wid)

    await rig.session.close_workbook(wid)
    rig.audit.flush()
    records = _read_audit(rig.audit.path)

    mutating = [record for record in records if record.get("operation_class") == "mutating"]
    tools = {record["tool"] for record in mutating}
    assert {
        "pivot_table_refresh",
        "pivot_field_set_layout",
        "pivot_field_apply_filter",
    } <= tools
    assert all(record.get("result") is not None for record in mutating)
    assert any(
        record["tool"] == "pivot_tables_list" and record.get("operation_class") == "read_only"
        for record in records
    )


# ---------------------------------------------------------------------------
# 9. Unsupported sources: a real consolidation cache is rejected
# ---------------------------------------------------------------------------
async def _build_consolidation_pivot(rig: Rig) -> dict[str, Any]:
    """Add a consolidation-sourced PivotTable to the open working copy (STA)."""

    def build() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        workbook = app.Workbooks(1)
        sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = "Unsupported"
        # xlSourceTypeConsolidation (3): multiple consolidation ranges are a
        # local, deterministic, non-database source (task_039 §7).
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

    build.__name__ = "build_consolidation_pivot"
    return dict(
        await rig.worker.submit(
            build,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    )


@requires_excel
async def test_consolidation_source_is_rejected(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    wid = info.workbook_id

    built = await _build_consolidation_pivot(rig)
    assert built["name"] == "ptConsolidated"
    assert built["source_type"] == 3  # xlSourceTypeConsolidation
    assert built["olap"] is False

    inventory = await rig.pivots.list_pivot_tables(wid)
    pivot = {pt.name: pt for pt in inventory.pivot_tables}["ptConsolidated"]
    assert pivot.source_type is PivotSourceType.CONSOLIDATION
    assert pivot.source_status == "unsupported_source"

    cache = await rig.pivots.get_pivot_cache(wid, built["cache_index"])
    assert cache.source_type is PivotSourceType.CONSOLIDATION
    assert cache.shared_cache is False
    assert cache.affected_pivot_tables == ["ptConsolidated"]

    # Refresh is refused with an explicit unsupported-source result (no refresh).
    await rig.snapshot.create_snapshot(wid)
    refreshed = await rig.pivots.refresh_pivot_table(wid, "ptConsolidated")
    assert refreshed.status is PivotRefreshStatus.UNSUPPORTED_SOURCE
    assert refreshed.refreshed is False
    assert refreshed.refresh_return is None
    assert refreshed.error is not None
    assert refreshed.error["code"] == "UNSUPPORTED_PIVOT_SOURCE"
    assert refreshed.refresh_date_before == refreshed.refresh_date_after

    # Field configuration is rejected before any field is touched.
    fields = await rig.pivots.list_pivot_fields(wid, "ptConsolidated")
    assert fields, "the consolidation PivotTable must expose source fields"
    target = fields[0].source_name
    before = _by_name(fields)[target]
    with pytest.raises(ExcelVbaMcpError) as layout_exc:
        await rig.pivots.set_pivot_field_layout(
            wid, "ptConsolidated", target, PivotFieldOrientation.COLUMN
        )
    assert layout_exc.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    assert layout_exc.value.context["reason"] == "unsupported_pivot_source"
    after = _by_name(await rig.pivots.list_pivot_fields(wid, "ptConsolidated"))[target]
    assert after.orientation is before.orientation
    assert after.position == before.position

    await rig.session.close_workbook(wid)


# ---------------------------------------------------------------------------
# 10. BG03 / BG04 / BG05
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_excel_hidden_and_never_foreground(rig: Rig) -> None:
    info = await rig.open_for_mutation()
    await rig.snapshot.create_snapshot(info.workbook_id)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    def read_app_flags() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"visible": bool(app.Visible), "display_alerts": bool(app.DisplayAlerts)}

    read_app_flags.__name__ = "read_app_flags"
    baseline = await rig.worker.submit(
        read_app_flags,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert baseline["visible"] is False

    # Exercise every read and mutation path while hidden.
    await rig.pivots.list_pivot_tables(info.workbook_id)
    await rig.pivots.refresh_pivot_table(info.workbook_id, "ptA")
    await rig.pivots.set_pivot_field_layout(
        info.workbook_id, "ptA", "Quarter", PivotFieldOrientation.ROW
    )
    await rig.pivots.apply_pivot_field_filter(info.workbook_id, "ptA", "Region", _manual("North"))

    flags = await rig.worker.submit(
        read_app_flags,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )

    # BG03/BG04: the owned instance is still hidden and no window is visible.
    assert flags["visible"] is False
    # BG08: temporary application settings are restored after every operation.
    assert flags["display_alerts"] == baseline["display_alerts"]
    _assert_owned_never_foreground(owned_pid, "pivot-workflow")
    _assert_no_visible_owned_windows(owned_pid, "pivot-workflow")

    await rig.session.close_workbook(info.workbook_id)


@requires_excel
async def test_committed_fixture_is_byte_identical(tmp_path: Path) -> None:
    """BG05: the committed fixture is never modified by the workflow."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    before = _sha256(pivot_project)
    original = tmp_path / pivot_project.name
    shutil.copy2(pivot_project, original)

    r = Rig(tmp_path, original)
    try:
        await r.start()
        info = await r.open_for_mutation()
        await r.snapshot.create_snapshot(info.workbook_id)
        await r.pivots.refresh_pivot_table(info.workbook_id, "ptA")
        await r.pivots.set_pivot_field_aggregation(
            info.workbook_id, "ptA", "Sales", PivotAggregation.MAX
        )
        await r.pivots.apply_pivot_field_filter(info.workbook_id, "ptA", "Region", _manual("North"))
        await r.session.close_workbook(info.workbook_id)
    finally:
        await r.close()

    assert _sha256(pivot_project) == before


@requires_excel
async def test_owned_excel_is_reaped_after_shutdown() -> None:
    """BG05: the automation PID is gone after shutdown (no orphan)."""

    pivot_project = _fixture_or_skip(_PIVOTS_FIXTURE)
    import tempfile

    tmp_path = Path(tempfile.mkdtemp(prefix="pivot_verify_"))
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
