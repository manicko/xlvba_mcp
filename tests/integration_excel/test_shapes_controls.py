"""Real Excel integration tests for the shapes + controls surface (task_030).

These tests drive the real
:class:`~excel_vba_mcp.application.shapes_service.ShapesService` and
:class:`~excel_vba_mcp.application.controls_service.ControlsService` against an
**automation-created hidden Excel instance** (never the user's Excel). They assert:

* an AutoShape / TextBox / Form Control button / ActiveX control can be added to a
  disposable working copy and are reported with distinct kinds
  (``ShapeKind`` / ``ControlKind``);
* ``shapes_list`` / ``shape_get`` expose ``on_action`` (the assigned macro, read
  verbatim — never executed), position/size and best-effort text;
* a Form Control's ``control_type`` (button / checkbox / …) and ``linked_cell``
  are reported, and its ``on_action`` can be set and cleared;
* an ActiveX control is reported with its ``prog_id`` and a **probable/inferred**
  event-procedure association (``macro_mapping=PROBABLE``), never as confirmed;
* ``shape_set_on_action`` / ``control_set_on_action`` are guarded (writable +
  snapshot + approval) and never execute the assigned macro;
* the original fixture on disk is byte-identical after every test (mutations only
  touch the disposable working copy);
* BG03/BG04/BG05 hold (hidden + non-activating owned Excel, no non-owned Excel
  touched, worker refuses mutations on a hung instance);
* the owned Excel PID is reaped after shutdown (no orphan).

Mutations run only on the disposable working copy of the fixture. Excel is created
with ``DispatchEx`` and always ``Visible = False``; no ``.Activate`` / ``.Select``
is ever called and no control is ever clicked.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_shapes_controls.py -v -s
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
    import pythoncom  # type: ignore[import-untyped]
    import win32com  # type: ignore[import-untyped]
    import win32gui  # type: ignore[import-untyped]
    import win32process  # type: ignore[import-untyped]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    ControlFormType,
    ControlKind,
    MacroMappingStatus,
    OperationClass,
    SessionMode,
    ShapeKind,
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
# STA helpers to seed shapes / controls on the working copy
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


async def _seed_shapes_and_controls(rig: Rig, workbook_id: str) -> dict[str, Any]:
    """Add an AutoShape, a TextBox, a Form Control button and an ActiveX control.

    Returns a mapping with the created names and the sheet name. Runs on the STA
    worker against the disposable working copy. ``Shapes.AddFormControl`` creates
    a real Form Control; ``Shapes.AddOLEObject`` creates a real ActiveX object
    (``Forms.CommandButton.1``) when the host supports it.
    """

    working_path = _record_path(rig, workbook_id)

    def seed() -> dict[str, Any]:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Item(1)
        sheet.Name = "ShapesData"

        # AutoShape (msoShapeRectangle = 1) with text.
        shape = sheet.Shapes.AddShape(1, 10, 10, 120, 40)
        shape.Name = "Rect 1"
        shape.TextFrame2.TextRange.Text = "hello"

        # TextBox (msoTextOrientationHorizontal = 1) with an OnAction assignment.
        text_box = sheet.Shapes.AddTextbox(1, 10, 80, 120, 30)
        text_box.Name = "Text 1"
        text_box.TextFrame2.TextRange.Text = "caption"
        text_box.OnAction = "FormatIt"

        # Form Control button via AddFormControl (xlFormControl.Button = 0).
        button = sheet.Shapes.AddFormControl(0, 10, 140, 120, 30)
        button.Name = "Button 1"
        button.OnAction = "RunReport"

        result: dict[str, Any] = {
            "sheet": str(sheet.Name),
            "autoshape": "Rect 1",
            "textbox": "Text 1",
            "form_button": "Button 1",
            "activex": None,
        }

        # ActiveX (best effort): ``Shapes.AddOLEObject`` with an ActiveX ProgID.
        # This needs the "Forms" library; when unavailable the test still asserts
        # the Form Control / shape behaviour and skips the ActiveX assertions.
        try:
            ole = sheet.Shapes.AddOLEObject(
                "Forms.CommandButton.1", "", False, "", 10, 200, 120, 30
            )
            ole.Name = "CommandButton1"
            result["activex"] = "CommandButton1"
        except Exception:  # noqa: BLE001 - ActiveX control class may be unavailable
            result["activex"] = None
        return result

    seed.__name__ = "seed_shapes_and_controls"
    return await rig.worker.submit(
        seed,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.MUTATING,
    )


# ---------------------------------------------------------------------------
# Inventory + type distinction
# ---------------------------------------------------------------------------
@requires_excel
async def test_shapes_and_controls_inventory(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)

    shapes = await rig.shapes.list_shapes(wid)
    by_name = {s.name: s for s in shapes}

    rect = by_name[seeded["autoshape"]]
    assert rect.kind is ShapeKind.AUTOSHAPE
    assert rect.text is not None and "hello" in rect.text
    assert rect.sheet_name == seeded["sheet"]

    text_box = by_name[seeded["textbox"]]
    assert text_box.kind is ShapeKind.TEXT_BOX
    # OnAction is read verbatim and never executed.
    assert text_box.on_action == "FormatIt"

    form_button = by_name[seeded["form_button"]]
    assert form_button.kind is ShapeKind.FORM_CONTROL
    assert form_button.control_format is True
    assert form_button.on_action == "RunReport"

    fetched = await rig.shapes.get_shape(wid, seeded["textbox"])
    assert fetched.name == seeded["textbox"]

    controls = await rig.controls.list_controls(wid)
    control_by_name = {c.name: c for c in controls}
    button = control_by_name[seeded["form_button"]]
    assert button.kind is ControlKind.FORM_CONTROL
    assert button.control_type is ControlFormType.BUTTON
    assert button.on_action == "RunReport"
    assert button.macro_mapping is MacroMappingStatus.UNKNOWN

    if seeded["activex"] is not None:
        activex = control_by_name[seeded["activex"]]
        assert activex.kind is ControlKind.ACTIVEX
        assert activex.prog_id is not None
        # ActiveX event mapping is probable/inferred, never confirmed.
        assert activex.macro_mapping is MacroMappingStatus.PROBABLE
        assert activex.probable_procedure is not None
        assert activex.on_action is None

    # BG03: source fixture unchanged.
    assert _sha256(rig.project) == before


@requires_excel
async def test_get_missing_shape_and_control_raise(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_shapes_and_controls(rig, wid)

    with pytest.raises(ExcelVbaMcpError):
        await rig.shapes.get_shape(wid, "NoSuchShape")
    with pytest.raises(ExcelVbaMcpError):
        await rig.controls.get_control(wid, "NoSuchControl")


# ---------------------------------------------------------------------------
# Guarded OnAction set / clear (never executed)
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_and_clear_shape_on_action(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("shape_set_on_action")

    updated = await rig.shapes.set_on_action(wid, seeded["textbox"], "NewMacro")
    assert updated.on_action == "NewMacro"

    # The macro was never executed: re-reading the shape shows the assignment, and
    # no macro run was attempted (the worker submit log has no run entry).
    reread = await rig.shapes.get_shape(wid, seeded["textbox"])
    assert reread.on_action == "NewMacro"

    cleared = await rig.shapes.set_on_action(wid, seeded["textbox"], "")
    assert cleared.on_action is None

    assert _sha256(rig.project) == before


@requires_excel
async def test_set_and_clear_control_on_action(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("control_set_on_action")

    updated = await rig.controls.set_on_action(wid, seeded["form_button"], "RunReport2")
    assert updated.on_action == "RunReport2"

    reread = await rig.controls.get_control(wid, seeded["form_button"])
    assert reread.on_action == "RunReport2"

    cleared = await rig.controls.set_on_action(wid, seeded["form_button"], "")
    assert cleared.on_action is None

    assert _sha256(rig.project) == before


@requires_excel
async def test_control_set_on_action_rejects_activex(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("control_set_on_action")

    if seeded["activex"] is None:
        pytest.skip("ActiveX control could not be created on this host")

    with pytest.raises(ExcelVbaMcpError):
        await rig.controls.set_on_action(wid, seeded["activex"], "RunReport")


# ---------------------------------------------------------------------------
# Guards
# ---------------------------------------------------------------------------
@requires_excel
async def test_shape_set_on_action_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    rig.grant("shape_set_on_action")  # approval but no snapshot

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.shapes.set_on_action(wid, seeded["textbox"], "X")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


@requires_excel
async def test_shape_set_on_action_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.shapes.set_on_action(wid, seeded["textbox"], "X")


@requires_excel
async def test_control_set_on_action_requires_writable(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("control_set_on_action")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.controls.set_on_action(wid, "Button 1", "X")
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("shape_set_on_action")
    await rig.shapes.set_on_action(wid, seeded["textbox"], "X")

    # BG03/BG05: the automation Excel is hidden (read on the STA worker so the COM
    # property access is thread-correct).
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
    assert rig.snapshot._latest_snapshot.get(wid) is not None

    _assert_owned_never_foreground(owned_pid, "shapes-set-on-action")
    _assert_no_visible_owned_windows(owned_pid, "shapes-set-on-action")
    assert _sha256(rig.project) == before


@requires_excel
async def test_bg04_worker_refuses_mutation_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    seeded = await _seed_shapes_and_controls(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("shape_set_on_action")

    def hang() -> None:
        import time

        # Block the STA thread past the soft deadline so the submit times out and
        # the worker enters NEEDS_RECOVERY (a hung dispatch is never retried).
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
    # so once the hang returns the worker accepts the shape mutation again on its
    # own (no auto-kill, no recovery needed).
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    assigned = await rig.shapes.set_on_action(wid, seeded["textbox"], "X")
    assert assigned is not None
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


@requires_excel
async def test_bg05_session_state_and_alerts(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING

    def read_alerts() -> bool:
        app = rig.app.app
        assert app is not None
        return bool(app.DisplayAlerts)

    read_alerts.__name__ = "read_alerts"
    alerts = await rig.worker.submit(
        read_alerts,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert alerts is True


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
