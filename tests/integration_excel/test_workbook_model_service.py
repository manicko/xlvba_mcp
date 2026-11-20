"""Real Excel integration tests for the Stage 1 workbook model (task_032).

These tests drive the production
:class:`~excel_vba_mcp.application.workbook_model_service.WorkbookModelService`
against a **hidden, automation-created Excel instance** (never the user's Excel)
and a disposable working copy of the committed ``CleanProject.xlsm`` fixture.
They assert:

* the full model contains sheets, tables, names, shapes, controls and VBA
  components (A03), each with a section status and a content hash;
* the content hash is stable across reads of the unchanged working copy;
* ``get_sheet_model`` is scoped to one sheet;
* the read-only MCP resources (model / VBA component inventory / one component's
  source) return the model and VBA source **without modifying the workbook and
  without running VBA**;
* BG03/BG04/BG05 hold (hidden + non-activating owned Excel, no visible owned
  windows, owned PID never foreground, alerts restored);
* the committed fixture and the temp original are byte-identical afterwards;
* the owned Excel PID is reaped after shutdown (no orphan).

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called and no macro is executed. Seeding of
the table / name / shape / form control runs on the STA worker against the
disposable working copy only.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_workbook_model_service.py -v -s
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
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.application.workbook_model_service import WorkbookModelService
from excel_vba_mcp.domain.statuses import (
    ModelSectionStatus,
    OperationClass,
    SessionMode,
    VbaAccessStatus,
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

_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

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


def _assert_no_all_cells(value: Any) -> None:
    """Recursively assert there is no all-cells / range payload in the model."""

    if isinstance(value, dict):
        for key, child in value.items():
            assert key not in {"cells", "ranges", "cell_values"}, key
            _assert_no_all_cells(child)
    elif isinstance(value, list):
        for child in value:
            _assert_no_all_cells(child)


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + full session/snapshot/policy stack."""

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
    clean_project = _fixture_or_skip(_CLEAN_PROJECT)
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


# ---------------------------------------------------------------------------
# STA helper to seed the disposable working copy
# ---------------------------------------------------------------------------
def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


def _record_path(rig: Rig, workbook_id: str) -> str:
    return rig.session._records[workbook_id].working_path


async def _seed_model_fixture(rig: Rig, workbook_id: str) -> dict[str, Any]:
    """Seed a table, a defined name, a shape and a Form Control on the copy.

    Runs on the STA worker against the disposable working copy only. Nothing is
    saved and no macro is executed.
    """

    working_path = _record_path(rig, workbook_id)

    def seed() -> dict[str, Any]:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Item(1)
        sheet.Name = "Data"
        sheet.Range("A1:B3").Value2 = [[1, 2], [3, 4], [5, 6]]

        table = sheet.ListObjects.Add(1, sheet.Range("A1:B3"), False, 1)
        table.Name = "TblData"

        wb.Names.Add("BookName", "=Data!$A$1")

        shape = sheet.Shapes.AddShape(1, 10, 10, 120, 40)
        shape.Name = "Rect 1"

        button = sheet.Shapes.AddFormControl(0, 10, 80, 120, 30)
        button.Name = "Button 1"
        button.OnAction = "modTest.NoArgs"

        return {
            "sheet": str(sheet.Name),
            "table": str(table.Name),
            "name": "BookName",
            "shape": str(shape.Name),
            "control": str(button.Name),
        }

    seed.__name__ = "seed_model_fixture"
    return await rig.worker.submit(
        seed,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.MUTATING,
    )


# ---------------------------------------------------------------------------
# Full model (A03)
# ---------------------------------------------------------------------------
@requires_excel
async def test_workbook_model_contains_all_stage1_sections(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_model_fixture(rig, wid)

    model = await rig.model.get_workbook_model(wid)

    # Sheets.
    assert seeded["sheet"] in {sheet.name for sheet in model.sheets}
    # Tables.
    table_names = {table.name for table in model.tables}
    assert seeded["table"] in table_names
    # Names.
    name_names = {name.name for name in model.names}
    assert seeded["name"] in name_names
    # Shapes.
    shape_names = {shape.name for shape in model.shapes}
    assert seeded["shape"] in shape_names
    # Controls.
    control_names = {control.name for control in model.controls}
    assert seeded["control"] in control_names

    # VBA components (when VBOM trust is available). The model never aborts when
    # it is not: it reports vba_access and keeps the other sections.
    if model.vba_access is VbaAccessStatus.AVAILABLE:
        assert model.vba_project is not None
        component_names = {c.component_name for c in model.vba_project.components}
        assert "modTest" in component_names
        assert model.vba_procedures, "no VBA procedures aggregated"
    else:
        assert model.vba_project is None
        assert model.sections["vba"].status is ModelSectionStatus.UNAVAILABLE

    # Every non-VBA section is available.
    for section in ("sheets", "tables", "names", "shapes", "controls"):
        assert model.sections[section].status is ModelSectionStatus.AVAILABLE, section

    # Content hash + no all-cells payload.
    assert model.model_hash is not None and len(model.model_hash) == 64
    _assert_no_all_cells(model.model_dump(mode="json"))

    # BG03: the committed fixture is untouched.
    assert _sha256(rig.project) == before


@requires_excel
async def test_workbook_model_hash_is_stable(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_model_fixture(rig, wid)

    first = await rig.model.get_workbook_model(wid)
    second = await rig.model.get_workbook_model(wid)

    assert first.model_hash == second.model_hash
    # Timestamps are present (the hash excludes them, so stability is real).
    assert first.created_at is not None


# ---------------------------------------------------------------------------
# Sheet model
# ---------------------------------------------------------------------------
@requires_excel
async def test_sheet_model_is_scoped(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_model_fixture(rig, wid)

    sheet_model = await rig.model.get_sheet_model(wid, seeded["sheet"])

    assert sheet_model.sheet.name == seeded["sheet"]
    assert seeded["table"] in {table.name for table in sheet_model.tables}
    assert seeded["name"] in {name.name for name in sheet_model.names}
    assert seeded["shape"] in {shape.name for shape in sheet_model.shapes}
    assert seeded["control"] in {control.name for control in sheet_model.controls}
    assert sheet_model.model_hash is not None and len(sheet_model.model_hash) == 64
    _assert_no_all_cells(sheet_model.model_dump(mode="json"))


# ---------------------------------------------------------------------------
# Read-only MCP resources (model / VBA inventory / one component's source)
# ---------------------------------------------------------------------------
@requires_excel
async def test_resources_read_without_mutation(rig: Rig) -> None:
    import excel_vba_mcp.server as server
    from excel_vba_mcp.api import vba_tools, workbook_tools

    info = await rig.open()
    wid = info.workbook_id
    await _seed_model_fixture(rig, wid)

    working_path = Path(_record_path(rig, wid))
    working_hash_before = _sha256(working_path)
    original_hash_before = _sha256(rig.project)

    workbook_tools.set_workbook_model_service(rig.model)
    vba_tools.set_vba_project_service(rig.vba)
    try:
        # Model resource.
        model_contents = list(await server.mcp.read_resource(f"excel://workbooks/{wid}/model"))
        assert model_contents, "model resource returned no content"
        model_payload = json.loads(model_contents[0].content)
        assert model_payload["workbook"]["workbook_id"] == wid
        assert model_payload["model_hash"]

        # VBA component inventory resource.
        component_contents = list(
            await server.mcp.read_resource(f"excel://workbooks/{wid}/vba/components")
        )
        component_payload = json.loads(component_contents[0].content)
        assert component_payload["workbook_id"] == wid
        assert "vba_access" in component_payload
        if component_payload["vba_access"] == VbaAccessStatus.AVAILABLE.value:
            names = {c["component_name"] for c in component_payload["components"]}
            assert "modTest" in names

        # Single component source resource (only when VBA is readable).
        if component_payload["vba_access"] == VbaAccessStatus.AVAILABLE.value:
            source_contents = list(
                await server.mcp.read_resource(f"excel://workbooks/{wid}/vba/components/modTest")
            )
            assert source_contents, "source resource returned no content"
            assert "AddNumbers" in source_contents[0].content
    finally:
        workbook_tools.set_workbook_model_service(None)
        vba_tools.set_vba_project_service(None)

    # BG03: reads never modified the working copy or the original fixture.
    assert _sha256(working_path) == working_hash_before
    assert _sha256(rig.project) == original_hash_before


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05) + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_05_hidden_owned_excel_and_alerts(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_model_fixture(rig, wid)
    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    fg_before = _foreground()
    await rig.model.get_workbook_model(wid)

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
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING

    _assert_owned_never_foreground(owned_pid, "workbook-model")
    _assert_no_visible_owned_windows(owned_pid, "workbook-model")
    if fg_before is not None and fg_before["pid"] != owned_pid:
        fg_after = _foreground()
        if fg_after is not None:
            assert fg_after["pid"] != owned_pid


@requires_excel
async def test_owned_excel_reaped_after_shutdown(tmp_path: Path) -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    clean_project = _fixture_or_skip(_CLEAN_PROJECT)
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    await r.start()
    await r.open()
    owned_pid = r.worker.get_status().owned_pid
    assert owned_pid is not None
    await r.close()

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
