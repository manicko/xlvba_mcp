"""Unit tests for the fail-closed compile service (task_020).

These tests never open Excel and never import COM. They drive the real
:class:`~excel_vba_mcp.application.compile_service.CompileService` and the real
:class:`~excel_vba_mcp.adapters.windows_ui.vbe_compile_ui.VBECompileUI` against a
pure-Python fake ``Workbook`` / ``VBProject`` / ``VBE`` object graph (plus a fake
STA worker that runs the submitted callable inline).

They prove the binding task_020 contract on the validated build:

* ``compile_project`` returns ``COMPILE_INDETERMINATE`` with reason
  ``NON_ACTIVATING_COMPILE_UNAVAILABLE`` and never ``compiled`` /
  ``compile_error`` — even for a project whose source contains a compile error;
* the diagnostics are partial and explicitly ``is_compile=False`` (references,
  broken references, protection, component inventory, VBOM availability);
* a locked project maps to ``VBA_PROJECT_LOCKED``; VBOM disabled maps to
  ``ACCESS_VBOM_DISABLED``; both remain indeterminate;
* the STA item is submitted as ``section="diagnostic"`` /
  ``op_class=EXECUTING``;
* command discovery is read-only: no ``Execute``, no UIA expansion, no Compile
  route on this build;
* ``get_last_diagnostics`` uses the cached result without a re-run.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.windows_ui.vbe_compile_ui import VBECompileUI
from excel_vba_mcp.application.compile_service import (
    REASON_ACCESS_VBOM_DISABLED,
    REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE,
    REASON_VBA_PROJECT_LOCKED,
    CompileService,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    CompileStatus,
    OperationClass,
    SessionMode,
    VbaComponentType,
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

pytestmark = pytest.mark.anyio

# ---------------------------------------------------------------------------
# Fakes (no Excel / no COM)
# ---------------------------------------------------------------------------
MOD_GOOD_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function
"""

MOD_BAD_SOURCE = """\
Option Explicit

Public Sub Bad()
    x = UndefinedIdentifier
End Sub
"""


class FakeCodeModule:
    """Minimal ``CodeModule`` for the inventory read."""

    def __init__(self, text: str) -> None:
        self._lines = text.split("\n")

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - COM casing
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:  # noqa: N802 - COM casing
        return "\n".join(self._lines[start - 1 : start - 1 + count])


class FakeComponent:
    def __init__(self, name: str, type_code: int, code: str) -> None:
        self.Name = name
        self.Type = type_code
        self.CodeModule = FakeCodeModule(code)
        self.Designer = None

    def Export(self, path: str) -> None:  # noqa: N802 - COM casing
        Path(path).write_text(self.CodeModule.Lines(1, self.CodeModule.CountOfLines))


class FakeReference:
    def __init__(self, name: str, *, broken: bool = False) -> None:
        self.Name = name
        self.GUID = "{00000000-0000-0000-0000-000000000000}"
        self.Major = 1
        self.Minor = 0
        self.FullPath = rf"C:\lib\{name}.dll"
        self.IsBroken = broken


class FakeComponents:
    def __init__(self, components: list[FakeComponent]) -> None:
        self._items = components

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakeComponent:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __call__(self, name: str) -> FakeComponent:
        for component in self._items:
            if component.Name == name:
                return component
        raise KeyError(name)

    def __iter__(self) -> Any:
        return iter(self._items)


class FakeProject:
    def __init__(self, components: list[FakeComponent], *, protection: int = 0) -> None:
        self.Name = "FakeProject"
        self.Protection = protection
        self.VBComponents = FakeComponents(components)
        self.References = [
            FakeReference("VBA"),
            FakeReference("Excel"),
            FakeReference("MissingLib", broken=True),
        ]


class FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` with an EXCEPINFO-style shape."""

    def __init__(self, message: str, excepinfo: tuple[Any, ...]) -> None:
        super().__init__(message, "Exception occurred.", excepinfo)
        self.hresult = 0x80020009 - 0x100000000  # DISP_E_EXCEPTION (signed)


class FakeWorkbook:
    def __init__(self, project: Any, *, vbom_disabled: bool = False) -> None:
        self._project = project
        self._vbom_disabled = vbom_disabled

    @property
    def VBProject(self) -> Any:  # noqa: N802 - COM casing
        if self._vbom_disabled:
            excepinfo = (
                0,
                "Microsoft Excel",
                "Programmatic access to Visual Basic Project is not trusted.",
                "xlmain11.chm",
                0,
                0x800A17B4 - 0x100000000,
            )
            raise FakeComError(
                "Programmatic access to Visual Basic Project is not trusted.", excepinfo
            )
        return self._project


class FakeWorkbooks:
    def __init__(self, project: Any, *, vbom_disabled: bool = False) -> None:
        self._project = project
        self._vbom_disabled = vbom_disabled
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **_kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(self._project, vbom_disabled=self._vbom_disabled)
        wb.FullName = path  # type: ignore[attr-defined]
        self.items.append(wb)
        return wb

    def __iter__(self) -> Any:
        return iter(self.items)


class FakeCommandBarControl:
    def __init__(self, control_id: int, caption: str, *, enabled: bool = True) -> None:
        self.Id = control_id
        self.Caption = caption
        self.Enabled = enabled
        self.executed = False

    def Execute(self) -> None:  # noqa: N802 - COM casing
        self.executed = True
        raise AssertionError("Execute must never be called by read-only discovery")


class FakeCommandBar:
    def __init__(self, name: str, controls: list[FakeCommandBarControl]) -> None:
        self.Name = name
        self.Controls = controls

    def FindControl(self, *_args: Any) -> None:  # noqa: N802 - COM casing
        raise RuntimeError("E_INVALIDARG (build-specific FindControl failure)")


class FakeCommandBars:
    def __init__(self, bars: list[FakeCommandBar]) -> None:
        self._bars = bars

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._bars)

    def __getitem__(self, name: str) -> FakeCommandBar:
        for bar in self._bars:
            if bar.Name == name:
                return bar
        raise KeyError(name)

    def __iter__(self) -> Any:
        return iter(self._bars)


class FakeVBE:
    def __init__(self, bars: FakeCommandBars) -> None:
        self.CommandBars = bars


class FakeApp:
    def __init__(self, project: Any, *, vbom_disabled: bool = False) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(project, vbom_disabled=vbom_disabled)
        debug_bar = FakeCommandBar(
            "Debug",
            [
                FakeCommandBarControl(212, "Design Mode"),
                FakeCommandBarControl(5415, "Run Project"),
            ],
        )
        self.VBE = FakeVBE(FakeCommandBars([debug_bar]))


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
            {
                "name": getattr(fn, "__name__", "?"),
                "section": section,
                "op_class": op_class,
                "timeout": timeout,
            }
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


_DEFAULT_PROJECT = FakeProject(
    [
        FakeComponent("modTest", 1, MOD_GOOD_SOURCE),
        FakeComponent("clsTest", 2, "Option Explicit\n"),
        FakeComponent("ThisWorkbook", 100, "Option Explicit\n"),
    ]
)


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    def __init__(
        self,
        tmp_path: Path,
        *,
        project: Any = _DEFAULT_PROJECT,
        vbom_disabled: bool = False,
    ) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.app = FakeApp(project, vbom_disabled=vbom_disabled)
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
        self.vba = VBAProjectService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.path_policy,
            self.audit,
        )
        self.service = CompileService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.vba,
            self.path_policy,
            self.audit,
        )

    async def open(self) -> Any:
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
            str(self.original), mode=SessionMode.INSPECT, read_only=True
        )


# ---------------------------------------------------------------------------
# Compile result: fail-closed indeterminate
# ---------------------------------------------------------------------------
async def test_compile_project_is_indeterminate_never_compiled(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    result = await rig.service.compile_project(info.workbook_id)

    assert result.status is CompileStatus.INDETERMINATE
    assert result.reason == REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE
    assert result.status is not CompileStatus.COMPILED
    assert result.status is not CompileStatus.COMPILE_ERROR
    assert result.diagnostics.is_compile is False
    assert any("NOT a compile" in note for note in result.diagnostics.notes)


async def test_compile_submits_executing_diagnostic_section(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    await rig.service.compile_project(info.workbook_id)

    submit = rig.worker.submits[-1]
    assert submit["name"] == "vba_compile"
    assert submit["section"] == "diagnostic"
    assert submit["op_class"] is OperationClass.EXECUTING


async def test_partial_diagnostics_fields(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    result = await rig.service.compile_project(info.workbook_id)
    diagnostics = result.diagnostics

    assert diagnostics.is_compile is False
    assert diagnostics.vbom_available is True
    assert diagnostics.protection == 0
    assert diagnostics.locked is False
    assert diagnostics.broken_references == ["MissingLib"]
    reference_names = {reference.name for reference in diagnostics.references}
    assert reference_names == {"VBA", "Excel", "MissingLib"}
    component_names = {component.component_name for component in diagnostics.components}
    assert component_names >= {"modTest", "clsTest", "ThisWorkbook"}
    assert diagnostics.component_count == len(diagnostics.components)
    assert {component.component_type for component in diagnostics.components} >= {
        VbaComponentType.STANDARD_MODULE,
        VbaComponentType.DOCUMENT_MODULE,
    }


async def test_compile_error_source_is_still_indeterminate(tmp_path: Path) -> None:
    """A source-level compile error must never produce a false ``compile_error``."""

    project = FakeProject([FakeComponent("modBad", 1, MOD_BAD_SOURCE)])
    rig = Rig(tmp_path, project=project)
    info = await rig.open()

    result = await rig.service.compile_project(info.workbook_id)

    assert result.status is CompileStatus.INDETERMINATE
    assert result.status is not CompileStatus.COMPILE_ERROR
    assert result.reason == REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE
    assert result.diagnostics.is_compile is False


# ---------------------------------------------------------------------------
# Locked / VBOM-disabled mapping
# ---------------------------------------------------------------------------
async def test_locked_project_maps_to_locked_reason(tmp_path: Path) -> None:
    project = FakeProject(
        [FakeComponent("modTest", 1, MOD_GOOD_SOURCE)],
        protection=1,
    )
    rig = Rig(tmp_path, project=project)
    info = await rig.open()

    result = await rig.service.compile_project(info.workbook_id)

    assert result.status is CompileStatus.INDETERMINATE
    assert result.reason == REASON_VBA_PROJECT_LOCKED
    assert result.diagnostics.locked is True
    assert result.diagnostics.protection == 1
    assert result.diagnostics.components == []
    assert any("locked" in note.lower() for note in result.diagnostics.notes)


async def test_vbom_disabled_maps_to_access_vbom_disabled(tmp_path: Path) -> None:
    rig = Rig(tmp_path, vbom_disabled=True)
    info = await rig.open()

    result = await rig.service.compile_project(info.workbook_id)

    assert result.status is CompileStatus.INDETERMINATE
    assert result.reason == REASON_ACCESS_VBOM_DISABLED
    assert result.diagnostics.vbom_available is False
    assert result.diagnostics.vbom_error == "ACCESS_VBOM_DISABLED"


# ---------------------------------------------------------------------------
# Command discovery is read-only
# ---------------------------------------------------------------------------
def test_command_discovery_is_read_only_and_route_unusable() -> None:
    app = FakeApp(_DEFAULT_PROJECT)
    discovery = VBECompileUI().discover_command(app.VBE)

    assert discovery.attempted is True
    assert discovery.executed is False
    assert discovery.uia_expand_attempted is False
    assert discovery.route_usable is False
    assert discovery.has_compile_entry is False
    assert discovery.bars_count == 1
    assert discovery.debug_controls_count == 2
    # FindControl probes recorded their build-specific failure.
    assert discovery.find_control
    assert all(probe.error for probe in discovery.find_control)
    # No control was executed.
    for bar in app.VBE.CommandBars:
        assert all(control.executed is False for control in bar.Controls)


def test_observe_dialog_rejects_interactive_mode() -> None:
    with pytest.raises(ValueError):
        VBECompileUI().observe_dialog(read_only=False)


# ---------------------------------------------------------------------------
# Cached diagnostics (no forced re-run)
# ---------------------------------------------------------------------------
async def test_get_last_diagnostics_uses_cache_without_rerun(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    first = await rig.service.compile_project(info.workbook_id)
    submit_count = len(rig.worker.submits)

    diagnostics = await rig.service.get_last_diagnostics(info.workbook_id)

    assert diagnostics is first.diagnostics
    assert len(rig.worker.submits) == submit_count


async def test_get_last_diagnostics_collects_once_when_absent(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    diagnostics = await rig.service.get_last_diagnostics(info.workbook_id)

    assert diagnostics.is_compile is False
    compile_submits = [s for s in rig.worker.submits if s["name"] == "vba_compile"]
    assert len(compile_submits) == 1


async def test_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.compile_project("ghost")
