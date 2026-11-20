"""Unit tests for the shapes + controls inventory surface (task_030).

These tests exercise the production
:class:`~excel_vba_mcp.application.shapes_service.ShapesService`,
:class:`~excel_vba_mcp.application.controls_service.ControlsService` and their
STA adapters (:class:`~excel_vba_mcp.adapters.excel.shapes_adapter.ShapesAdapter`,
:class:`~excel_vba_mcp.adapters.excel.controls_adapter.ControlsAdapter`) against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

The fake models ``Worksheet.Shapes`` (AutoShape / TextBox / Form Control shapes)
and ``Worksheet.OLEObjects`` (ActiveX controls), plus the workbook's
``Worksheets`` collection. A fake ``Shape`` exposes a settable ``OnAction`` and a
``ControlFormat`` object only for Form Controls, so the adapters' type mapping and
``ControlFormat`` probing are exercised without Excel.

Covered:

* shape type mapping (AutoShape / TextBox / Form Control / ActiveX / unknown);
* form control ``control_type`` mapping and ``linked_cell`` / ``caption``;
* ActiveX ``prog_id`` + **probable/inferred** event-procedure association;
* missing/unreadable properties degrade to diagnosable ``None`` + notes;
* ``OnAction`` set/get on shapes and Form Controls (never executed);
* ActiveX ``OnAction`` rejection and ambiguous/missing name errors;
* guard chain on the service layer: no snapshot -> ``SNAPSHOT_FAILED``, no
  approval -> denied, read-only session -> ``WORKBOOK_READ_ONLY``;
* reads submit ``READ_ONLY``; mutations submit ``MUTATING``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.controls_adapter import ControlsAdapter
from excel_vba_mcp.adapters.excel.shapes_adapter import ShapesAdapter
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import ControlEntry, ShapeInfo
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
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# MsoShapeType / xlFormControl codes used by the fake object graph
# ---------------------------------------------------------------------------
_MSO_AUTO_SHAPE = 1
_MSO_TEXT_BOX = 17
_MSO_FORM_CONTROL = 8
_MSO_OLE_CONTROL = 9  # msoOleControlObject (the ActiveX host shape type)

# xlFormControl codes.
_XL_FORM_BUTTON = 0
_XL_FORM_CHECKBOX = 1
_XL_FORM_LIST_BOX = 4


class FakeControlFormat:
    """A fake ``ControlFormat`` (Form Controls only).

    ``control_type`` may be ``None`` to emulate a host that does not expose
    ``ControlFormat.Type`` (the adapter then falls back to ``FormControlType`` or
    the shape name).
    """

    def __init__(self, control_type: int | None, linked_cell: str | None = None) -> None:
        if control_type is not None:
            self.Type = control_type
        self.LinkedCell = linked_cell


class FakeShape:
    """A fake ``Shape`` proxy.

    ``kind`` is the ``MsoShapeType`` numeric code returned by ``Type``. A
    ``control_format`` (only for Form Controls) makes ``ControlFormat`` readable;
    otherwise the attribute is absent so the adapter's probe falls through.
    """

    def __init__(
        self,
        name: str,
        kind: int,
        *,
        on_action: str = "",
        text: str | None = None,
        visible: bool = True,
        top: float = 10.0,
        left: float = 20.0,
        width: float = 100.0,
        height: float = 30.0,
        alternative_text: str | None = None,
        control_format: FakeControlFormat | None = None,
        shape_id: int = 1,
    ) -> None:
        self.Name = name
        self.Type = kind
        self.OnAction = on_action
        self.Visible = visible
        self.Top = top
        self.Left = left
        self.Width = width
        self.Height = height
        self.AlternativeText = alternative_text
        self.Id = shape_id
        if control_format is not None:
            self.ControlFormat = control_format
        if text is not None:
            self.TextFrame2 = type(
                "_Frame2",
                (),
                {"TextRange": type("_Range", (), {"Text": text})()},
            )()


class FakeOleObject:
    """A fake ``OLEObject`` (ActiveX control)."""

    def __init__(
        self,
        name: str,
        prog_id: str = "Forms.CommandButton.1",
        *,
        visible: bool = True,
        top: float = 5.0,
        left: float = 6.0,
        width: float = 50.0,
        height: float = 20.0,
        shape: Any = None,
    ) -> None:
        self.Name = name
        self.ProgId = prog_id
        self.Visible = visible
        self.Top = top
        self.Left = left
        self.Width = width
        self.Height = height
        self.Shape = shape


class FakeShapes:
    def __init__(self, items: list[FakeShape]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakeShape:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __iter__(self):
        return iter(list(self._items))


class FakeOleObjects:
    def __init__(self, items: list[FakeOleObject]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakeOleObject:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __iter__(self):
        return iter(list(self._items))


class FakeWorksheet:
    def __init__(
        self,
        name: str,
        shapes: list[FakeShape] | None = None,
        ole_objects: list[FakeOleObject] | None = None,
        *,
        code_name: str | None = None,
    ) -> None:
        self.Name = name
        self.CodeName = code_name if code_name is not None else name
        self.Shapes = FakeShapes(shapes or [])
        self.OLEObjects = FakeOleObjects(ole_objects or [])


class FakeWorksheets:
    def __init__(self, sheets: list[FakeWorksheet]) -> None:
        self._sheets = sheets

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._sheets)

    def Item(self, index: int) -> FakeWorksheet:  # noqa: N802 - COM casing
        return self._sheets[index - 1]

    def __iter__(self):
        return iter(list(self._sheets))


class FakeWorkbook:
    def __init__(self, sheets: list[FakeWorksheet] | None = None) -> None:
        self.Name = "Book.xlsm"
        self.FullName = "C:/tmp/Book.xlsm"
        self.ReadOnly = False
        self._sheets = sheets if sheets is not None else [FakeWorksheet("Data")]

    @property
    def Worksheets(self) -> FakeWorksheets:  # noqa: N802 - COM casing
        return FakeWorksheets(self._sheets)

    def sheet(self, name: str) -> FakeWorksheet:
        for sheet in self._sheets:
            if sheet.Name == name:
                return sheet
        raise KeyError(name)


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook([FakeWorksheet("Data")])
        wb.FullName = path
        wb.ReadOnly = bool(kwargs.get("ReadOnly", False))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)


class FakeApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks()


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self) -> None:
        self.submits: list[dict[str, Any]] = []

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append(
            {"name": getattr(fn, "__name__", "?"), "section": section, "op_class": op_class}
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# ---------------------------------------------------------------------------
# Fixtures / builders
# ---------------------------------------------------------------------------
def _populated_workbook() -> FakeWorkbook:
    """A workbook with one AutoShape, one TextBox, one Form Control and one ActiveX."""

    data = FakeWorksheet("Data", code_name="Sheet1")
    data.Shapes = FakeShapes(
        [
            FakeShape("Rect 1", _MSO_AUTO_SHAPE, text="hello", alternative_text="a rect"),
            FakeShape("Text 1", _MSO_TEXT_BOX, text="caption", on_action="FormatIt"),
            FakeShape(
                "Button 1",
                _MSO_FORM_CONTROL,
                text="Click me",
                control_format=FakeControlFormat(_XL_FORM_BUTTON, linked_cell="Data!$B$1"),
            ),
            # An ActiveX host shape (msoOLEControlObject) whose OLEObject lives in
            # the worksheet's OLEObjects collection.
            FakeShape("CommandButton1", _MSO_OLE_CONTROL, text="OK"),
        ]
    )
    data.OLEObjects = FakeOleObjects([FakeOleObject("CommandButton1", "Forms.CommandButton.1")])
    other = FakeWorksheet("Other", code_name="Sheet2")
    other.Shapes = FakeShapes(
        [
            FakeShape(
                "Check 1", _MSO_FORM_CONTROL, control_format=FakeControlFormat(_XL_FORM_CHECKBOX)
            )
        ]
    )
    return FakeWorkbook([data, other])


# ---------------------------------------------------------------------------
# Adapter: shapes
# ---------------------------------------------------------------------------
def test_shapes_list_maps_kinds_and_on_action() -> None:
    infos = ShapesAdapter().list_shapes(_populated_workbook())
    by_name = {info.name: info for info in infos}

    rect = by_name["Rect 1"]
    assert rect.kind is ShapeKind.AUTOSHAPE
    assert rect.text == "hello"
    assert rect.alternative_text == "a rect"
    assert rect.on_action is None
    assert rect.sheet_name == "Data"

    text = by_name["Text 1"]
    assert text.kind is ShapeKind.TEXT_BOX
    # OnAction is read verbatim, never executed.
    assert text.on_action == "FormatIt"

    # A Form Control shape reports control_format=True and its own kind.
    button = by_name["Button 1"]
    assert button.kind is ShapeKind.FORM_CONTROL
    assert button.control_format is True

    # The ActiveX host shape is classified as ACTIVEX.
    assert by_name["CommandButton1"].kind is ShapeKind.ACTIVEX

    # The second sheet's checkbox is listed too.
    assert "Check 1" in by_name
    assert by_name["Check 1"].kind is ShapeKind.FORM_CONTROL


def test_shapes_list_unreadable_properties_are_diagnosable() -> None:
    class _Broken:
        """A shape whose property reads raise, to prove diagnosable degradation."""

        Name = "Broken"

        @property
        def Type(self) -> int:  # noqa: N802 - COM casing
            raise RuntimeError("COM failure")

        @property
        def Visible(self) -> bool:  # noqa: N802 - COM casing
            raise RuntimeError("COM failure")

        @property
        def Top(self) -> float:  # noqa: N802 - COM casing
            raise RuntimeError("COM failure")

    wb = FakeWorkbook([FakeWorksheet("Data", shapes=[_Broken()])])  # type: ignore[list-item]
    info = ShapesAdapter().list_shapes(wb)[0]
    assert isinstance(info, ShapeInfo)
    assert info.kind is ShapeKind.UNKNOWN
    assert info.visible is None
    assert info.top is None
    # The inventory still returned a diagnosable entry instead of raising.
    assert info.name == "Broken"


def test_get_shape_by_name_and_errors() -> None:
    adapter = ShapesAdapter()
    wb = _populated_workbook()

    assert adapter.get_shape(wb, "rect 1").name == "Rect 1"  # case-insensitive

    with pytest.raises(ExcelVbaMcpError):
        adapter.get_shape(wb, "NoSuchShape")

    with pytest.raises(ExcelVbaMcpError):
        adapter.get_shape(wb, "")


def test_get_shape_ambiguous_rejected() -> None:
    wb = FakeWorkbook(
        [
            FakeWorksheet("Data", shapes=[FakeShape("Dup", _MSO_AUTO_SHAPE)]),
            FakeWorksheet("Other", shapes=[FakeShape("Dup", _MSO_AUTO_SHAPE)]),
        ]
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        ShapesAdapter().get_shape(wb, "Dup")
    assert exc_info.value.context.get("reason") == "ambiguous"


def test_set_shape_on_action_writes_and_returns_reread() -> None:
    wb = _populated_workbook()
    adapter = ShapesAdapter()

    updated = adapter.set_on_action(wb, "Text 1", "NewMacro")
    assert updated.on_action == "NewMacro"
    # The DTO is re-read from the live shape.
    assert adapter.get_shape(wb, "Text 1").on_action == "NewMacro"

    # Clearing (empty string) is supported.
    cleared = adapter.set_on_action(wb, "Text 1", "")
    assert cleared.on_action is None


# ---------------------------------------------------------------------------
# Adapter: controls
# ---------------------------------------------------------------------------
def test_controls_list_distinguishes_form_and_activex() -> None:
    entries = ControlsAdapter().list_controls(_populated_workbook())
    by_name = {entry.name: entry for entry in entries}

    button = by_name["Button 1"]
    assert button.kind is ControlKind.FORM_CONTROL
    assert button.control_type is ControlFormType.BUTTON
    assert button.linked_cell == "Data!$B$1"
    assert button.caption == "Click me"
    assert button.on_action is None
    assert button.macro_mapping is MacroMappingStatus.UNKNOWN
    assert button.probable_procedure is None

    # ActiveX: prog id + probable/inferred event association.
    activex = by_name["CommandButton1"]
    assert activex.kind is ControlKind.ACTIVEX
    assert activex.prog_id == "Forms.CommandButton.1"
    assert activex.macro_mapping is MacroMappingStatus.PROBABLE
    # Inferred from <sheet_code_name>_<control>_<Event>; never confirmed.
    assert activex.probable_procedure == "Sheet1_CommandButton1_Click"
    assert activex.on_action is None

    # The checkbox on the second sheet keeps its own control type.
    check = by_name["Check 1"]
    assert check.kind is ControlKind.FORM_CONTROL
    assert check.control_type is ControlFormType.CHECKBOX

    # Shapes that are not controls are not listed as controls.
    assert "Rect 1" not in by_name
    assert "Text 1" not in by_name


def test_controls_list_falls_back_when_control_format_type_missing() -> None:
    """A Form Control whose ``ControlFormat.Type`` is absent resolves via name."""

    wb = FakeWorkbook(
        [
            FakeWorksheet(
                "Data",
                shapes=[
                    FakeShape(
                        "Combo 1",
                        _MSO_FORM_CONTROL,
                        control_format=FakeControlFormat(None, linked_cell="Data!$C$1"),
                    )
                ],
            )
        ]
    )
    entry = ControlsAdapter().list_controls(wb)[0]
    assert entry.kind is ControlKind.FORM_CONTROL
    assert entry.control_type is ControlFormType.COMBO_BOX
    assert entry.linked_cell == "Data!$C$1"
    assert any("inferred from name" in note for note in entry.notes)


def test_controls_list_unreadable_properties_are_diagnosable() -> None:
    class _Broken:
        """An OLEObject whose property reads raise."""

        Name = "BrokenActiveX"

        @property
        def ProgId(self) -> str:  # noqa: N802 - COM casing
            raise RuntimeError("COM failure")

        @property
        def Shape(self) -> Any:  # noqa: N802 - COM casing
            raise RuntimeError("COM failure")

    wb = FakeWorkbook([FakeWorksheet("Data", ole_objects=[_Broken()])])  # type: ignore[list-item]
    entry = ControlsAdapter().list_controls(wb)[0]
    assert isinstance(entry, ControlEntry)
    assert entry.kind is ControlKind.ACTIVEX
    assert entry.prog_id is None
    assert entry.macro_mapping is MacroMappingStatus.PROBABLE
    assert any("ProgId unavailable" in note for note in entry.notes)


def test_get_control_and_errors() -> None:
    adapter = ControlsAdapter()
    wb = _populated_workbook()

    assert adapter.get_control(wb, "button 1").name == "Button 1"

    with pytest.raises(ExcelVbaMcpError):
        adapter.get_control(wb, "NoSuchControl")


def test_set_control_on_action_writes_and_returns_reread() -> None:
    wb = _populated_workbook()
    adapter = ControlsAdapter()

    updated = adapter.set_on_action(wb, "Button 1", "RunReport")
    assert updated.on_action == "RunReport"
    assert adapter.get_control(wb, "Button 1").on_action == "RunReport"

    cleared = adapter.set_on_action(wb, "Button 1", "")
    assert cleared.on_action is None


def test_set_control_on_action_rejects_activex() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        ControlsAdapter().set_on_action(wb, "CommandButton1", "RunReport")
    assert exc_info.value.context.get("kind") == "activex"


def test_set_control_on_action_rejects_plain_shape() -> None:
    wb = FakeWorkbook([FakeWorksheet("Data", shapes=[FakeShape("Rect 1", _MSO_AUTO_SHAPE)])])
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        ControlsAdapter().set_on_action(wb, "Rect 1", "RunReport")
    assert exc_info.value.context.get("kind") == "shape"


# ---------------------------------------------------------------------------
# Service-level rig
# ---------------------------------------------------------------------------
class Rig:
    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.app = FakeApp()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = FakeWorker()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )
        self.shapes = ShapesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self.controls = ControlsService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

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
            str(self.original), mode=SessionMode.INSPECT, read_only=read_only
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

    def add_snapshot(self, workbook_id: str) -> None:
        manifest = SnapshotManifest(
            snapshot_id="snap-test",
            created_at=datetime.now(UTC),
            source_path=self.canonical,
            file_name="snap.xlsm",
            stored_path=str(self.tmp_path / "snap.xlsm"),
            sha256="0" * 64,
            size=0,
        )
        self.snapshot._latest_snapshot[workbook_id] = manifest
        self.snapshot._known_sources[workbook_id] = self.canonical

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


def _seed_workbook(rig: Rig) -> None:
    """Replace the freshly opened fake workbook with a populated one."""

    wb = _populated_workbook()
    wb.FullName = rig.workbook().FullName
    wb.ReadOnly = rig.workbook().ReadOnly
    rig.app.Workbooks.items = [wb]


# ---------------------------------------------------------------------------
# Service: reads (READ_ONLY, no approval)
# ---------------------------------------------------------------------------
async def test_service_shapes_list_and_get(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    wid = info.workbook_id

    shapes = await rig.shapes.list_shapes(wid)
    assert any(s.name == "Text 1" and s.on_action == "FormatIt" for s in shapes)

    one = await rig.shapes.get_shape(wid, "Text 1")
    assert one.kind is ShapeKind.TEXT_BOX

    read_submit = next(s for s in rig.worker.submits if s["name"] == "shapes_list")
    assert read_submit["op_class"] == OperationClass.READ_ONLY


async def test_service_controls_list_and_get(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    wid = info.workbook_id

    controls = await rig.controls.list_controls(wid)
    activex = next(c for c in controls if c.name == "CommandButton1")
    assert activex.kind is ControlKind.ACTIVEX
    assert activex.macro_mapping is MacroMappingStatus.PROBABLE

    one = await rig.controls.get_control(wid, "Button 1")
    assert one.control_type is ControlFormType.BUTTON

    read_submit = next(s for s in rig.worker.submits if s["name"] == "controls_list")
    assert read_submit["op_class"] == OperationClass.READ_ONLY


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.shapes.list_shapes("missing")
    with pytest.raises(ExcelVbaMcpError):
        await rig.controls.list_controls("missing")


# ---------------------------------------------------------------------------
# Service: guards on the mutation
# ---------------------------------------------------------------------------
async def test_shape_set_on_action_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    rig.grant("shape_set_on_action")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.shapes.set_on_action(info.workbook_id, "Text 1", "X")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "shape_set_on_action"] == []


async def test_shape_set_on_action_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    rig.add_snapshot(info.workbook_id)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.shapes.set_on_action(info.workbook_id, "Text 1", "X")
    assert [s for s in rig.worker.submits if s["name"] == "shape_set_on_action"] == []


async def test_shape_set_on_action_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    _seed_workbook(rig)
    rig.add_snapshot(info.workbook_id)
    rig.grant("shape_set_on_action")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.shapes.set_on_action(info.workbook_id, "Text 1", "X")
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_control_set_on_action_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    rig.grant("control_set_on_action")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.controls.set_on_action(info.workbook_id, "Button 1", "X")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_control_set_on_action_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError):
        await rig.controls.set_on_action(info.workbook_id, "Button 1", "X")


async def test_control_set_on_action_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    _seed_workbook(rig)
    rig.add_snapshot(info.workbook_id)
    rig.grant("control_set_on_action")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.controls.set_on_action(info.workbook_id, "Button 1", "X")
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


# ---------------------------------------------------------------------------
# Service: guarded mutation success (MUTATING, macro not executed)
# ---------------------------------------------------------------------------
async def test_shape_set_on_action_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    wid = info.workbook_id
    rig.add_snapshot(wid)
    rig.grant("shape_set_on_action")

    updated = await rig.shapes.set_on_action(wid, "Text 1", "RunReport")
    assert isinstance(updated, ShapeInfo)
    assert updated.on_action == "RunReport"
    mutate = next(s for s in rig.worker.submits if s["name"] == "shape_set_on_action")
    assert mutate["op_class"] == OperationClass.MUTATING


async def test_control_set_on_action_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_workbook(rig)
    wid = info.workbook_id
    rig.add_snapshot(wid)
    rig.grant("control_set_on_action")

    updated = await rig.controls.set_on_action(wid, "Button 1", "RunReport")
    assert isinstance(updated, ControlEntry)
    assert updated.on_action == "RunReport"
    mutate = next(s for s in rig.worker.submits if s["name"] == "control_set_on_action")
    assert mutate["op_class"] == OperationClass.MUTATING

    # The assigned macro was never executed: the only submissions are the
    # mutation itself (no macro/run submit).
    assert all("run" not in s["name"] for s in rig.worker.submits)
