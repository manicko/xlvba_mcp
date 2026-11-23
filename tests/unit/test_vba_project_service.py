"""Unit tests for the read-only VBA project service (task_016).

These tests never open Excel and never import COM. They drive the real
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService` and the
real :class:`~excel_vba_mcp.adapters.excel.vbide_adapter.VBIDEAdapter` against a
pure-Python fake ``Workbook`` / ``VBProject`` object graph (plus a fake STA
worker that runs the submitted callable inline). That covers the binding logic
that does not need Excel:

* component type mapping (1/2/3/100) and ``has_designer`` / ``is_document_module``;
* procedure discovery via the CodeModule API (``source="code_module"``) and the
  ``Find``-less heuristic fallback (``is_heuristic=True``/``source="heuristic"``);
* export manifest: ``.bas``/``.cls``/``.frm`` + ``.frx``, ``export_status``,
  normalized ``source_hash`` vs exported file hash;
* references: valid vs ``IsBroken`` → ``missing_references``;
* locked project → ``VbaProjectLockedError``; VBOM disabled → ``AccessVbomDisabledError``;
* the service resolves the open workbook, submits READ_ONLY diagnostic work and
  returns DTOs.
"""

from __future__ import annotations

import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter, hash_code
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    VbaProjectLockedError,
)
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    ProcedureType,
    ReferenceStatus,
    SessionMode,
    VbaComponentType,
    VbaExportStatus,
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
MOD_TEST_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Private Sub Helper()
    ' helper body
End Sub
"""

CLS_TEST_SOURCE = """\
Option Explicit

Public FullName As String
"""

FRM_TEST_SOURCE = """\
Option Explicit
"""


class FakeCodeModule:
    """Fake ``CodeModule`` supporting the VBIDE procedure APIs.

    ``ProcOfLine`` / ``ProcStartLine`` / ``ProcCountLines`` accept the
    ``vbext_ProcKind`` argument (always ``vbext_pk_Proc`` = 0 here) and classify
    by the declaration keyword at the requested line.
    """

    def __init__(self, text: str, *, fail: bool = False) -> None:
        self._text = text
        self._lines = text.split("\n")
        self._fail = fail

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        selected = self._lines[start - 1 : start - 1 + count]
        return "\n".join(selected)

    def ProcOfLine(self, line: int, kind: int = 0) -> tuple[str, int]:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        header = self._lines[line - 1].strip()
        lowered = header.casefold()
        # Return (name, kind) as the real VBIDE does; the name is the token before "(".
        name = header.split("(")[0].split()[-1]
        if "function" in lowered:
            return name, 1
        if "property" in lowered:
            return name, 2
        return name, 0

    def ProcStartLine(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        for index, line in enumerate(self._lines, start=1):
            is_declaration = line.strip().startswith(
                ("Sub ", "Function ", "Property ", "Public ", "Private ")
            )
            if is_declaration and name in line:
                return index
        return 1

    def ProcCountLines(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        # Count until the matching End (Sub/Function/Property).
        start = 1
        for index, line in enumerate(self._lines, start=1):
            is_declaration = line.strip().startswith(
                ("Sub ", "Function ", "Property ", "Public ", "Private ")
            )
            if is_declaration and name in line:
                start = index
                break
        depth = 0
        for index in range(start - 1, len(self._lines)):
            stripped = self._lines[index].strip()
            if stripped.casefold().startswith(("sub ", "function ", "property ")):
                depth += 1
            if stripped.casefold().startswith("end ") and depth > 0:
                depth -= 1
                if depth == 0:
                    return index - (start - 1) + 1
        return 1

    def Find(self, target: str, *_args: Any, **_kwargs: Any) -> bool:  # noqa: N802 - COM casing
        if self._fail:
            raise RuntimeError("code module unavailable")
        return any(target in line for line in self._lines)


class _ProcOfLineFailingModule(FakeCodeModule):
    """A ``CodeModule`` whose ``ProcOfLine`` always raises (heuristic fallback)."""

    def ProcOfLine(self, line: int, kind: int = 0) -> tuple[str, int]:  # noqa: N802 - COM casing
        raise RuntimeError("ProcOfLine unavailable")


class _ProcStartLineFailingModule(FakeCodeModule):
    """A ``CodeModule`` whose ``ProcStartLine`` always raises (heuristic fallback)."""

    def ProcStartLine(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        raise RuntimeError("ProcStartLine unavailable")


class _ProcCountLinesFailingModule(FakeCodeModule):
    """A ``CodeModule`` whose ``ProcCountLines`` always raises (heuristic fallback)."""

    def ProcCountLines(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        raise RuntimeError("ProcCountLines unavailable")


class FakeDesigner:
    def __init__(self) -> None:
        self.Controls: list[Any] = []


class FakeComponent:
    """Fake ``VBComponent`` (Type 1/2/3/100) with an exportable file body."""

    def __init__(self, name: str, type_code: int, code: str, *, fail_code: bool = False) -> None:
        self.Name = name
        self.Type = type_code
        self.CodeModule = FakeCodeModule(code, fail=fail_code)
        self.Designer = FakeDesigner() if type_code == 3 else None
        self.exported_to: str | None = None

    def Export(self, path: str) -> None:  # noqa: N802 - COM casing
        self.exported_to = path
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        extension = target.suffix.lower()
        if extension == ".frm":
            header = (
                "VERSION 5.00\r\n"
                "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTest\r\n"
                "End\r\n"
                'Attribute VB_Name = "frmTest"\r\n'
                "Attribute VB_GlobalNameSpace = False\r\n"
            )
            target.write_text(
                header + self.CodeModule.Lines(1, self.CodeModule.CountOfLines), encoding="utf-8"
            )
            target.with_suffix(".frx").write_bytes(b"FRX-BYTES")
        else:
            target.write_text(
                self.CodeModule.Lines(1, self.CodeModule.CountOfLines), encoding="utf-8"
            )


class FakeReference:
    def __init__(self, name: str, *, broken: bool = False) -> None:
        self.Name = name
        self.GUID = "{00000000-0000-0000-0000-000000000000}"
        self.Major = 1
        self.Minor = 0
        self.FullPath = rf"C:\lib\{name}.dll"
        self.IsBroken = broken


class FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` with an EXCEPINFO-style shape."""

    def __init__(self, message: str, excepinfo: tuple[Any, ...]) -> None:
        super().__init__(message, "Exception occurred.", excepinfo)
        self.hresult = 0x80020009 - 0x100000000  # DISP_E_EXCEPTION (signed)


class FakeProject:
    """Fake ``VBProject`` (protection, components, references)."""

    def __init__(self, components: list[FakeComponent], *, protection: int = 0) -> None:
        self.Name = "FakeProject"
        self.Protection = protection
        self.VBComponents = FakeComponents(components)
        self.References = [
            FakeReference("VBA"),
            FakeReference("Excel"),
            FakeReference("MissingLib", broken=True),
        ]


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

    def __iter__(self):
        return iter(self._items)


class FakeWorkbook:
    """Fake ``Workbook`` whose ``VBProject`` may raise a VBOM-disabled error."""

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
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **_kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(_PROJECT)
        wb.FullName = path  # type: ignore[attr-defined]
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
            {
                "name": getattr(fn, "__name__", "?"),
                "section": section,
                "op_class": op_class,
            }
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# A shared project graph is enough: the service only reads it.
_PROJECT = FakeProject(
    [
        FakeComponent("modTest", 1, MOD_TEST_SOURCE),
        FakeComponent("clsTest", 2, CLS_TEST_SOURCE),
        FakeComponent("frmTest", 3, FRM_TEST_SOURCE),
        FakeComponent("ThisWorkbook", 100, "Option Explicit\n"),
        FakeComponent("Sheet1", 100, "Option Explicit\n"),
    ]
)

# A project whose CodeModule API fails so the heuristic fallback is exercised.
_HEURISTIC_PROJECT = FakeProject(
    [
        FakeComponent("modHeuristic", 1, MOD_TEST_SOURCE, fail_code=True),
    ]
)

# A project whose ProcOfLine API fails so the heuristic procedure scan is used.
_PROC_OF_LINE_FAILING_PROJECT = FakeProject(
    [
        FakeComponent("modHeuristic", 1, MOD_TEST_SOURCE),
    ]
)
_PROC_OF_LINE_FAILING_PROJECT.VBComponents.Item(1).CodeModule = _ProcOfLineFailingModule(
    MOD_TEST_SOURCE
)

# A project whose ProcStartLine API fails so the heuristic procedure scan is used.
_PROC_START_LINE_FAILING_PROJECT = FakeProject(
    [
        FakeComponent("modHeuristic", 1, MOD_TEST_SOURCE),
    ]
)
_PROC_START_LINE_FAILING_PROJECT.VBComponents.Item(1).CodeModule = _ProcStartLineFailingModule(
    MOD_TEST_SOURCE
)

# A project whose ProcCountLines API fails so the heuristic procedure scan is used.
_PROC_COUNT_LINES_FAILING_PROJECT = FakeProject(
    [
        FakeComponent("modHeuristic", 1, MOD_TEST_SOURCE),
    ]
)
_PROC_COUNT_LINES_FAILING_PROJECT.VBComponents.Item(1).CodeModule = _ProcCountLinesFailingModule(
    MOD_TEST_SOURCE
)


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    def __init__(self, tmp_path: Path, *, project: Any = _PROJECT) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy(mode=AccessMode.FULL)
        self.app = FakeApp()
        self.app.Workbooks.items.clear()
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
        self.service = VBAProjectService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.path_policy,
            self.audit,
            approval=self.approval,
        )
        self._project = project

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
        info = await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=True
        )
        # Point the fake workbook at this rig's project graph.
        wb = self.app.Workbooks.items[-1]
        wb._project = self._project
        return info


# ---------------------------------------------------------------------------
# Adapter-level: component type mapping
# ---------------------------------------------------------------------------
def test_component_type_mapping_and_flags() -> None:
    adapter = VBIDEAdapter()
    components = adapter.list_components(FakeWorkbook(_PROJECT))

    by_name = {component.component_name: component for component in components}
    assert by_name["modTest"].component_type is VbaComponentType.STANDARD_MODULE
    assert by_name["clsTest"].component_type is VbaComponentType.CLASS_MODULE
    assert by_name["frmTest"].component_type is VbaComponentType.USER_FORM
    assert by_name["ThisWorkbook"].component_type is VbaComponentType.DOCUMENT_MODULE
    assert by_name["Sheet1"].component_type is VbaComponentType.DOCUMENT_MODULE

    assert by_name["frmTest"].has_designer is True
    assert by_name["modTest"].has_designer is False
    assert by_name["ThisWorkbook"].is_document_module is True
    assert by_name["ThisWorkbook"].is_exportable is True
    assert by_name["modTest"].source_hash is not None
    assert len(by_name["modTest"].source_hash) == 64


def test_project_info_reports_components_and_missing_references() -> None:
    adapter = VBIDEAdapter()
    info = adapter.project_info(FakeWorkbook(_PROJECT))

    assert info.available is True
    assert info.locked is False
    assert info.component_count == 5
    assert {reference.name for reference in info.references} == {"VBA", "Excel", "MissingLib"}
    broken = [reference for reference in info.references if reference.is_broken]
    assert [reference.name for reference in broken] == ["MissingLib"]
    assert broken[0].status is ReferenceStatus.BROKEN
    assert info.missing_references == ["MissingLib"]
    assert any("missing_references" in warning for warning in info.warnings)


def test_locked_project_raises_vba_project_locked() -> None:
    adapter = VBIDEAdapter()
    locked = FakeProject([FakeComponent("modTest", 1, MOD_TEST_SOURCE)], protection=1)
    with pytest.raises(VbaProjectLockedError) as exc_info:
        adapter.project_info(FakeWorkbook(locked))
    assert exc_info.value.code is ErrorCode.VBA_PROJECT_LOCKED


def test_vbom_disabled_raises_access_vbom_disabled() -> None:
    adapter = VBIDEAdapter()
    with pytest.raises(AccessVbomDisabledError) as exc_info:
        adapter.list_components(FakeWorkbook(_PROJECT, vbom_disabled=True))
    assert exc_info.value.code is ErrorCode.ACCESS_VBOM_DISABLED
    assert exc_info.value.hresult == 0x800A17B4


# ---------------------------------------------------------------------------
# Adapter-level: procedures
# ---------------------------------------------------------------------------
def test_list_procedures_via_code_module() -> None:
    adapter = VBIDEAdapter()
    procedures = adapter.list_procedures(FakeWorkbook(_PROJECT), "modTest")

    by_name = {procedure.name: procedure for procedure in procedures}
    assert set(by_name) == {"AddNumbers", "Helper"}
    add = by_name["AddNumbers"]
    assert add.procedure_type is ProcedureType.FUNCTION
    assert add.visibility.value == "public"
    assert add.is_heuristic is False
    assert add.source == "code_module"
    assert add.start_line >= 1
    assert add.end_line >= add.start_line
    assert "AddNumbers" in add.code

    helper = by_name["Helper"]
    assert helper.procedure_type is ProcedureType.SUB
    assert helper.visibility.value == "private"
    assert helper.source == "code_module"


def test_list_procedures_heuristic_fallback_is_flagged() -> None:
    adapter = VBIDEAdapter()
    cases = [
        _PROC_OF_LINE_FAILING_PROJECT,
        _PROC_START_LINE_FAILING_PROJECT,
        _PROC_COUNT_LINES_FAILING_PROJECT,
    ]
    for project in cases:
        procedures = adapter.list_procedures(FakeWorkbook(project), "modHeuristic")

        assert procedures, "heuristic scan must still find the declared procedures"
        assert all(procedure.is_heuristic is True for procedure in procedures)
        assert all(procedure.source == "heuristic" for procedure in procedures)
        names = {procedure.name for procedure in procedures}
        assert "AddNumbers" in names
        assert "Helper" in names

    # When the whole CodeModule API is unavailable the source text is empty and
    # the adapter returns no procedures (nothing to scan) rather than failing.
    assert adapter.list_procedures(FakeWorkbook(_HEURISTIC_PROJECT), "modHeuristic") == []


def test_read_component_returns_source_and_hash() -> None:
    adapter = VBIDEAdapter()
    source = adapter.read_component(FakeWorkbook(_PROJECT), "modTest")

    assert source.component_name == "modTest"
    assert source.component_type is VbaComponentType.STANDARD_MODULE
    assert "AddNumbers" in source.source
    assert len(source.source_hash) == 64
    assert source.source_hash == hash_code(MOD_TEST_SOURCE)


# ---------------------------------------------------------------------------
# Adapter-level: export
# ---------------------------------------------------------------------------
def test_export_project_manifest_and_files(tmp_path: Path) -> None:
    adapter = VBIDEAdapter()
    out_dir = tmp_path / "exports"
    manifest = adapter.export_project(FakeWorkbook(_PROJECT), out_dir, workbook_id="wb")

    assert manifest.workbook_id == "wb"
    by_name = {entry.component_name: entry for entry in manifest.components}
    assert by_name["modTest"].extension == ".bas"
    assert by_name["clsTest"].extension == ".cls"
    assert by_name["frmTest"].extension == ".frm"
    assert by_name["ThisWorkbook"].extension == ".cls"
    assert all(entry.export_status is VbaExportStatus.OK for entry in manifest.components)

    # Files exist with the correct extensions.
    assert (out_dir / "modTest.bas").is_file()
    assert (out_dir / "clsTest.cls").is_file()
    assert (out_dir / "frmTest.frm").is_file()
    assert (out_dir / "frmTest.frx").is_file()
    assert (out_dir / "ThisWorkbook.cls").is_file()

    # UserForm companion resource recorded; normalized source hash matches read.
    frm_entry = by_name["frmTest"]
    assert frm_entry.resource_file_name == "frmTest.frx"
    assert frm_entry.resource_file_sha256 is not None
    assert frm_entry.file_sha256 is not None
    assert frm_entry.source_hash == hash_code(FRM_TEST_SOURCE)

    # Combined manifest hash is stable.
    assert manifest.source_hash is not None
    assert len(manifest.source_hash) == 64


# ---------------------------------------------------------------------------
# Service-level: submission, resolution, export validation
# ---------------------------------------------------------------------------
async def test_service_project_info_submits_read_only_diagnostic(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    result = await rig.service.get_project_info(info.workbook_id)

    assert result.component_count == 5
    submit = next(s for s in rig.worker.submits if s["name"] == "vba_project_info")
    assert submit["section"] == "diagnostic"
    assert submit["op_class"] == OperationClass.READ_ONLY


async def test_service_list_and_read_components(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    components = await rig.service.list_components(info.workbook_id)
    assert {component.component_name for component in components} >= {
        "modTest",
        "clsTest",
        "frmTest",
        "ThisWorkbook",
        "Sheet1",
    }

    procedures = await rig.service.list_procedures(info.workbook_id, "modTest")
    assert any(procedure.name == "AddNumbers" for procedure in procedures)

    source = await rig.service.read_component(info.workbook_id, "modTest")
    assert source.source_hash


async def test_service_export_project_default_dir(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    manifest = await rig.service.export_project(info.workbook_id)

    assert manifest.out_dir
    assert Path(manifest.out_dir).is_dir()
    # The default (out_dir omitted) still returns absolute paths (unchanged).
    assert Path(manifest.out_dir).is_absolute()
    assert {entry.component_name for entry in manifest.components} >= {"modTest", "frmTest"}


async def test_service_export_project_validates_out_dir(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    # A path policy whose "artifacts" logical root resolves to an absolute dir
    # (mirrors the server's EXCEL_VBA_MCP_ARTIFACT_DIR registration).
    artifacts_root = tmp_path / "artifacts"
    artifacts_root.mkdir()
    path_policy = PathPolicy(
        {str(tmp_path): str(tmp_path), "artifacts": str(artifacts_root)}
    )
    service = VBAProjectService(
        rig.worker,  # type: ignore[arg-type]
        rig.session,
        path_policy,
        rig.audit,
        approval=rig.approval,
    )

    # A valid out_dir under the resolved artifacts root exports successfully.
    manifest = await service.export_project(info.workbook_id, str(artifacts_root / "vba"))
    assert Path(manifest.out_dir).is_dir()
    assert Path(manifest.out_dir).is_relative_to(artifacts_root.resolve())

    expected_root = os.path.realpath(artifacts_root)

    # A path outside every allowlisted root is refused; the error context reports
    # the ABSOLUTE artifacts root (not the bare symbolic "artifacts") so a caller
    # can retry. code / reason are unchanged.
    outside = tmp_path.parent / f"outside_{uuid4().hex}"
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await service.export_project(info.workbook_id, str(outside / "vba"))
    error = exc_info.value
    assert error.code is ErrorCode.PATH_NOT_ALLOWED
    assert error.context["reason"] == "outside_roots"
    assert error.context["allowed_root"] == expected_root
    assert error.context["allowed_root"] != "artifacts"
    assert Path(error.context["allowed_root"]).is_absolute()

    # A directory inside an allowlisted root but NOT the artifacts root yields
    # root_mismatch, with the symbolic expected_root replaced by the absolute
    # artifacts-root path (still PATH_NOT_ALLOWED / root_mismatch).
    other = tmp_path / "work"
    other.mkdir(exist_ok=True)
    with pytest.raises(ExcelVbaMcpError) as mismatch:
        await service.export_project(info.workbook_id, str(other / "vba"))
    assert mismatch.value.code is ErrorCode.PATH_NOT_ALLOWED
    assert mismatch.value.context["reason"] == "root_mismatch"
    assert mismatch.value.context["expected_root"] == expected_root
    assert mismatch.value.context["expected_root"] != "artifacts"


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.get_project_info("ghost")


async def test_service_references_and_search(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    references = await rig.service.get_references(info.workbook_id)
    broken = [reference.name for reference in references if reference.is_broken]
    assert broken == ["MissingLib"]

    hits = await rig.service.search_code(info.workbook_id, "AddNumbers")
    assert any(hit.component_name == "modTest" for hit in hits)

    regex_hits = await rig.service.search_code(info.workbook_id, r"^Public Function", regex=True)
    assert any(hit.component_name == "modTest" for hit in regex_hits)


# ---------------------------------------------------------------------------
# F-4: export_project is approval-gated (filesystem write)
# ---------------------------------------------------------------------------
async def test_service_export_project_restricted_requires_approval(tmp_path: Path) -> None:
    """In restricted mode an export is denied by default (F-4)."""

    rig = Rig(tmp_path)
    info = await rig.open()

    restricted = ApprovalPolicy(mode=AccessMode.RESTRICTED)
    service = VBAProjectService(
        rig.worker,  # type: ignore[arg-type]
        rig.session,
        rig.path_policy,
        rig.audit,
        approval=restricted,
    )

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await service.export_project(info.workbook_id)
    assert exc_info.value.code.value == "APPROVAL_DENIED"


async def test_service_export_project_full_mode_succeeds(tmp_path: Path) -> None:
    """In full mode an export is auto-allowed (no behavior change)."""

    rig = Rig(tmp_path)
    info = await rig.open()

    manifest = await rig.service.export_project(info.workbook_id)

    assert Path(manifest.out_dir).is_dir()
    assert {entry.component_name for entry in manifest.components} >= {"modTest", "frmTest"}
