"""Unit tests for the Stage 1 workbook model aggregation service (task_032).

These tests never touch COM / Excel. They drive the production
:class:`~excel_vba_mcp.application.workbook_model_service.WorkbookModelService`
against pure-Python fake inventory services and assert:

* the model aggregates sheets, tables, names, shapes, controls and VBA;
* the aggregation order is deterministic (sheets → tables → names → shapes →
  controls → VBA);
* a failing section is isolated (status ``unavailable``) and never aborts the
  rest of the model;
* VBA access failures (VBOM disabled / locked / generic) set ``vba_access`` and
  keep the non-VBA sections;
* ``include_vba=False`` skips VBA and reports ``vba_access='skipped'``;
* the content hash is stable across reads and changes with the content;
* the model carries no all-cells / Range data and the service never touches a
  ``Range`` / ``Cells`` COM property;
* the sheet model is scoped to one sheet;
* the thin ``workbook_get_model`` tool validates and delegates.
"""

from __future__ import annotations

import ast
from datetime import UTC, datetime
from pathlib import Path

import pytest

from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.workbook_tools import workbook_get_model
from excel_vba_mcp.application.workbook_model_service import (
    SECTION_CONTROLS,
    SECTION_NAMES,
    SECTION_SHAPES,
    SECTION_SHEETS,
    SECTION_TABLES,
    SECTION_VBA,
    WorkbookModelService,
    content_hash,
)
from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ExcelVbaMcpError,
    VbaProjectLockedError,
)
from excel_vba_mcp.domain.models import (
    ControlEntry,
    DefinedNameInfo,
    ShapeInfo,
    SheetInfo,
    SheetModel,
    TableInfo,
    VbaComponentInfo,
    VbaProcedureInfo,
    VbaProjectInfo,
    WorkbookInfo,
    WorkbookModel,
)
from excel_vba_mcp.domain.statuses import (
    ControlKind,
    ModelSectionStatus,
    NameScope,
    NameTargetKind,
    ProcedureType,
    SessionMode,
    ShapeKind,
    SheetObjectType,
    SheetVisibility,
    VbaAccessStatus,
    VbaComponentType,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio

_SERVICE_SOURCE = (
    Path(__file__).resolve().parents[2]
    / "src"
    / "excel_vba_mcp"
    / "application"
    / "workbook_model_service.py"
)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fakes
# ---------------------------------------------------------------------------
class FakeSession:
    def __init__(self, info: WorkbookInfo) -> None:
        self._info = info
        self.session_id = info.session_id

    def list_workbooks_sync(self) -> list[WorkbookInfo]:
        return [self._info]


class FakeSheets:
    def __init__(self, sheets: list[SheetInfo], calls: list[str]) -> None:
        self._sheets = sheets
        self._calls = calls

    async def list_sheets(self, workbook_id: str) -> list[SheetInfo]:
        self._calls.append(SECTION_SHEETS)
        return list(self._sheets)

    async def get_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        self._calls.append(f"{SECTION_SHEETS}:get")
        for sheet in self._sheets:
            if sheet.name == name:
                return sheet
        raise ExcelVbaMcpError(f"sheet {name!r} not found")


class FakeTables:
    def __init__(self, tables: list[TableInfo], calls: list[str]) -> None:
        self._tables = tables
        self._calls = calls

    async def list_tables(self, workbook_id: str) -> list[TableInfo]:
        self._calls.append(SECTION_TABLES)
        return list(self._tables)


class FakeNames:
    def __init__(self, names: list[DefinedNameInfo], calls: list[str]) -> None:
        self._names = names
        self._calls = calls

    async def list_names(self, workbook_id: str) -> list[DefinedNameInfo]:
        self._calls.append(SECTION_NAMES)
        return list(self._names)


class FakeShapes:
    def __init__(self, shapes: list[ShapeInfo], calls: list[str]) -> None:
        self._shapes = shapes
        self._calls = calls

    async def list_shapes(self, workbook_id: str) -> list[ShapeInfo]:
        self._calls.append(SECTION_SHAPES)
        return list(self._shapes)


class FakeControls:
    def __init__(self, controls: list[ControlEntry], calls: list[str]) -> None:
        self._controls = controls
        self._calls = calls

    async def list_controls(self, workbook_id: str) -> list[ControlEntry]:
        self._calls.append(SECTION_CONTROLS)
        return list(self._controls)


class FakeVba:
    def __init__(
        self,
        calls: list[str],
        *,
        project: VbaProjectInfo | None = None,
        error: Exception | None = None,
        procedures: dict[str, list[VbaProcedureInfo]] | None = None,
        procedure_errors: dict[str, Exception] | None = None,
    ) -> None:
        self._calls = calls
        self._project = project
        self._error = error
        self._procedures = procedures or {}
        self._procedure_errors = procedure_errors or {}

    async def get_project_info(self, workbook_id: str) -> VbaProjectInfo:
        self._calls.append(f"{SECTION_VBA}:project")
        if self._error is not None:
            raise self._error
        assert self._project is not None
        return self._project

    async def list_procedures(self, workbook_id: str, component: str) -> list[VbaProcedureInfo]:
        self._calls.append(f"{SECTION_VBA}:procedures:{component}")
        if component in self._procedure_errors:
            raise self._procedure_errors[component]
        return list(self._procedures.get(component, []))


def _workbook_info() -> WorkbookInfo:
    return WorkbookInfo(
        workbook_id="wb1",
        name="Book.xlsm",
        format="xlsm",
        path="C:/work/wb1.xlsm",
        source_path="C:/data/Book.xlsm",
        read_only=True,
        has_vba=True,
        excel_version="16.0",
        session_id="s1",
        mode=SessionMode.INSPECT,
    )


def _sheet(name: str = "Sheet1") -> SheetInfo:
    return SheetInfo(
        name=name,
        code_name=f"{name}_Code",
        index=1,
        object_type=SheetObjectType.WORKSHEET,
        visibility=SheetVisibility.VISIBLE,
    )


def _table(name: str, sheet: str) -> TableInfo:
    return TableInfo(name=name, sheet_name=sheet, range_address="A1:B2")


def _name(name: str, *, scope: NameScope, sheet: str | None, target: str | None) -> DefinedNameInfo:
    return DefinedNameInfo(
        name=name,
        scope=scope,
        refers_to="=Sheet1!$A$1",
        target_kind=NameTargetKind.RANGE,
        target_sheet=target,
        target_address="A1",
        sheet_name=sheet,
    )


def _shape(name: str, sheet: str) -> ShapeInfo:
    return ShapeInfo(name=name, kind=ShapeKind.AUTOSHAPE, sheet_name=sheet)


def _control(name: str, sheet: str) -> ControlEntry:
    return ControlEntry(name=name, kind=ControlKind.FORM_CONTROL, sheet_name=sheet)


def _component(name: str, component_type: VbaComponentType) -> VbaComponentInfo:
    return VbaComponentInfo(component_name=name, component_type=component_type)


def _procedure(component: str, name: str) -> VbaProcedureInfo:
    return VbaProcedureInfo(
        component_name=component,
        name=name,
        procedure_type=ProcedureType.SUB,
        start_line=1,
        end_line=2,
    )


def _project(components: list[VbaComponentInfo]) -> VbaProjectInfo:
    return VbaProjectInfo(
        available=True,
        locked=False,
        name="VBAProject",
        component_count=len(components),
        components=components,
    )


class Rig:
    """Builds a WorkbookModelService over fakes with a shared call log."""

    def __init__(self, tmp_path: Path) -> None:
        self.calls: list[str] = []
        self.info = _workbook_info()
        self.session = FakeSession(self.info)
        self.sheets = FakeSheets([_sheet("Sheet1"), _sheet("Sheet2")], self.calls)
        self.tables = FakeTables([_table("Tbl1", "Sheet1"), _table("Tbl2", "Sheet2")], self.calls)
        self.names = FakeNames(
            [
                _name("BookName", scope=NameScope.WORKBOOK, sheet=None, target="Sheet1"),
                _name("OtherBook", scope=NameScope.WORKBOOK, sheet=None, target="Sheet2"),
                _name("Sheet1Name", scope=NameScope.WORKSHEET, sheet="Sheet1", target=None),
            ],
            self.calls,
        )
        self.shapes = FakeShapes([_shape("Rect 1", "Sheet1")], self.calls)
        self.controls = FakeControls([_control("Button 1", "Sheet1")], self.calls)
        self.vba = FakeVba(
            self.calls,
            project=_project([_component("modTest", VbaComponentType.STANDARD_MODULE)]),
            procedures={"modTest": [_procedure("modTest", "Main")]},
        )
        self.path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.service = WorkbookModelService(
            None,  # type: ignore[arg-type]
            self.session,  # type: ignore[arg-type]
            self.sheets,  # type: ignore[arg-type]
            self.tables,  # type: ignore[arg-type]
            self.names,  # type: ignore[arg-type]
            self.shapes,  # type: ignore[arg-type]
            self.controls,  # type: ignore[arg-type]
            self.vba,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )

    def close(self) -> None:
        self.audit.close()


# ---------------------------------------------------------------------------
# Aggregation
# ---------------------------------------------------------------------------
async def test_get_workbook_model_aggregates_all_sections(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.workbook.workbook_id == "wb1"
    assert [sheet.name for sheet in model.sheets] == ["Sheet1", "Sheet2"]
    assert [table.name for table in model.tables] == ["Tbl1", "Tbl2"]
    assert {name.name for name in model.names} == {"BookName", "OtherBook", "Sheet1Name"}
    assert [shape.name for shape in model.shapes] == ["Rect 1"]
    assert [control.name for control in model.controls] == ["Button 1"]
    assert model.vba_access is VbaAccessStatus.AVAILABLE
    assert model.vba_project is not None
    assert [component.component_name for component in model.vba_project.components] == ["modTest"]
    assert [procedure.name for procedure in model.vba_procedures] == ["Main"]
    assert model.model_hash is not None and len(model.model_hash) == 64

    # Every section was read successfully.
    assert model.sections[SECTION_SHEETS].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_TABLES].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_NAMES].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_SHAPES].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_CONTROLS].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_VBA].status is ModelSectionStatus.AVAILABLE
    assert model.sections[SECTION_SHEETS].count == 2


async def test_aggregation_order_is_deterministic(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    # Sheets → tables → names → shapes → controls → VBA project → VBA procedures.
    assert rig.calls[:5] == [
        SECTION_SHEETS,
        SECTION_TABLES,
        SECTION_NAMES,
        SECTION_SHAPES,
        SECTION_CONTROLS,
    ]
    assert rig.calls[5] == f"{SECTION_VBA}:project"
    assert rig.calls[6] == f"{SECTION_VBA}:procedures:modTest"


# ---------------------------------------------------------------------------
# VBA access handling
# ---------------------------------------------------------------------------
async def test_vbom_disabled_keeps_other_sections(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    rig.vba = FakeVba(rig.calls, error=AccessVbomDisabledError("trust access off"))
    rig.service._vba = rig.vba  # type: ignore[attr-defined]
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.vba_access is VbaAccessStatus.VBOM_DISABLED
    assert model.vba_project is None
    assert model.vba_procedures == []
    assert model.sections[SECTION_VBA].status is ModelSectionStatus.UNAVAILABLE
    # Non-VBA sections are intact.
    assert [sheet.name for sheet in model.sheets] == ["Sheet1", "Sheet2"]
    assert [table.name for table in model.tables] == ["Tbl1", "Tbl2"]
    assert len(model.controls) == 1


async def test_locked_project_keeps_other_sections(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    rig.vba = FakeVba(rig.calls, error=VbaProjectLockedError("locked"))
    rig.service._vba = rig.vba  # type: ignore[attr-defined]
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.vba_access is VbaAccessStatus.LOCKED
    assert model.sections[SECTION_VBA].status is ModelSectionStatus.UNAVAILABLE
    assert len(model.sheets) == 2


async def test_generic_vba_unavailable_keeps_other_sections(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    rig.vba = FakeVba(rig.calls, error=ExcelVbaMcpError("boom"))
    rig.service._vba = rig.vba  # type: ignore[attr-defined]
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.vba_access is VbaAccessStatus.UNAVAILABLE
    assert len(model.names) == 3


async def test_include_vba_false_skips_vba(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        model = await rig.service.get_workbook_model("wb1", include_vba=False)
    finally:
        rig.close()

    assert model.vba_access is VbaAccessStatus.SKIPPED
    assert model.vba_project is None
    assert model.vba_procedures == []
    assert model.sections[SECTION_VBA].status is ModelSectionStatus.AVAILABLE
    # The VBA service was never called.
    assert not any(call.startswith(SECTION_VBA) for call in rig.calls)


async def test_partial_vba_when_one_component_procedures_fail(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    rig.vba = FakeVba(
        rig.calls,
        project=_project(
            [
                _component("modTest", VbaComponentType.STANDARD_MODULE),
                _component("clsTest", VbaComponentType.CLASS_MODULE),
            ]
        ),
        procedures={"modTest": [_procedure("modTest", "Main")]},
        procedure_errors={"clsTest": ExcelVbaMcpError("cannot read")},
    )
    rig.service._vba = rig.vba  # type: ignore[attr-defined]
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.vba_access is VbaAccessStatus.AVAILABLE
    assert [procedure.name for procedure in model.vba_procedures] == ["Main"]
    assert model.sections[SECTION_VBA].status is ModelSectionStatus.PARTIAL
    assert model.sections[SECTION_VBA].count == 2
    assert model.sections[SECTION_VBA].diagnostics


# ---------------------------------------------------------------------------
# Section isolation
# ---------------------------------------------------------------------------
async def test_failing_section_is_isolated(tmp_path: Path) -> None:
    rig = Rig(tmp_path)

    async def boom(workbook_id: str) -> list[ShapeInfo]:
        raise ExcelVbaMcpError("shapes unavailable")

    rig.shapes.list_shapes = boom  # type: ignore[method-assign]
    try:
        model = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert model.shapes == []
    assert model.sections[SECTION_SHAPES].status is ModelSectionStatus.UNAVAILABLE
    assert model.sections[SECTION_SHAPES].diagnostics
    # Other sections remain available.
    assert len(model.sheets) == 2
    assert len(model.tables) == 2
    assert len(model.controls) == 1


# ---------------------------------------------------------------------------
# Content hash
# ---------------------------------------------------------------------------
async def test_content_hash_is_stable_and_content_sensitive(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        first = await rig.service.get_workbook_model("wb1")
        second = await rig.service.get_workbook_model("wb1")
    finally:
        rig.close()

    assert first.model_hash == second.model_hash

    # A content change yields a different hash.
    changed = first.model_copy(update={"model_hash": None})
    changed.sheets = []

    assert content_hash(changed) != first.model_hash


# ---------------------------------------------------------------------------
# No all-cells / Range data
# ---------------------------------------------------------------------------
def test_model_has_no_all_cells_fields() -> None:
    names = set(WorkbookModel.model_fields) | set(SheetModel.model_fields)
    forbidden = {"cells", "ranges", "range", "used_cells", "cell_values"}
    assert names & forbidden == set(), names & forbidden


def test_service_never_reads_range_or_cells_com() -> None:
    tree = ast.parse(_SERVICE_SOURCE.read_text(encoding="utf-8"))
    attributes = {node.attr for node in ast.walk(tree) if isinstance(node, ast.Attribute)}
    assert "Range" not in attributes
    assert "Cells" not in attributes


# ---------------------------------------------------------------------------
# Sheet model
# ---------------------------------------------------------------------------
async def test_get_sheet_model_is_scoped_to_sheet(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        model = await rig.service.get_sheet_model("wb1", "Sheet1")
    finally:
        rig.close()

    assert model.sheet.name == "Sheet1"
    assert [table.name for table in model.tables] == ["Tbl1"]
    # Worksheet-scoped name on Sheet1 + workbook-scope name targeting Sheet1.
    assert {name.name for name in model.names} == {"BookName", "Sheet1Name"}
    assert [shape.name for shape in model.shapes] == ["Rect 1"]
    assert [control.name for control in model.controls] == ["Button 1"]
    assert model.model_hash is not None and len(model.model_hash) == 64
    assert model.sections[SECTION_TABLES].status is ModelSectionStatus.AVAILABLE


async def test_get_sheet_model_isolates_failures(tmp_path: Path) -> None:
    rig = Rig(tmp_path)

    async def boom(workbook_id: str) -> list[ShapeInfo]:
        raise ExcelVbaMcpError("shapes unavailable")

    rig.shapes.list_shapes = boom  # type: ignore[method-assign]
    try:
        model = await rig.service.get_sheet_model("wb1", "Sheet1")
    finally:
        rig.close()

    assert model.shapes == []
    assert model.sections[SECTION_SHAPES].status is ModelSectionStatus.UNAVAILABLE
    assert [table.name for table in model.tables] == ["Tbl1"]


async def test_get_sheet_model_unknown_sheet_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        with pytest.raises(ExcelVbaMcpError):
            await rig.service.get_sheet_model("wb1", "Nope")
    finally:
        rig.close()


async def test_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    try:
        with pytest.raises(ExcelVbaMcpError):
            await rig.service.get_workbook_model("missing")
    finally:
        rig.close()


# ---------------------------------------------------------------------------
# Thin tool: workbook_get_model
# ---------------------------------------------------------------------------
class FakeModelService:
    def __init__(self, model: WorkbookModel) -> None:
        self._model = model
        self.calls: list[tuple[str, bool]] = []

    async def get_workbook_model(
        self, workbook_id: str, *, include_vba: bool = True
    ) -> WorkbookModel:
        self.calls.append((workbook_id, include_vba))
        return self._model


def _empty_model() -> WorkbookModel:
    now = datetime.now(UTC)
    return WorkbookModel(workbook=_workbook_info(), created_at=now, updated_at=now)


async def test_workbook_get_model_delegates() -> None:
    fake = FakeModelService(_empty_model())
    workbook_tools.set_workbook_model_service(fake)  # type: ignore[arg-type]
    try:
        model = await workbook_get_model("wb1", include_vba=False)
    finally:
        workbook_tools.set_workbook_model_service(None)

    assert model.workbook.workbook_id == "wb1"
    assert fake.calls == [("wb1", False)]


async def test_workbook_get_model_defaults_include_vba_true() -> None:
    fake = FakeModelService(_empty_model())
    workbook_tools.set_workbook_model_service(fake)  # type: ignore[arg-type]
    try:
        await workbook_get_model("wb1")
    finally:
        workbook_tools.set_workbook_model_service(None)

    assert fake.calls == [("wb1", True)]


async def test_workbook_get_model_rejects_empty_id() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_get_model("")


async def test_workbook_get_model_requires_service() -> None:
    workbook_tools.set_workbook_model_service(None)
    with pytest.raises(ExcelVbaMcpError):
        await workbook_get_model("wb1")


# ---------------------------------------------------------------------------
# Server registration (no Excel)
# ---------------------------------------------------------------------------
async def test_server_registers_model_tool_and_resources() -> None:
    import excel_vba_mcp.server as server

    tools = {tool.name for tool in await server.mcp.list_tools()}
    assert "workbook_get_model" in tools

    templates = {template.uri_template for template in await server.mcp.list_resource_templates()}
    assert "excel://workbooks/{workbook_id}/model" in templates
    assert "excel://workbooks/{workbook_id}/vba/components" in templates
    assert "excel://workbooks/{workbook_id}/vba/components/{component_name}" in templates
