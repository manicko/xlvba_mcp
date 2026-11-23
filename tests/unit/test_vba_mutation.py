"""Unit tests for guarded VBA source mutations (task_018).

These tests never open Excel and never import COM. They drive the real
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService` mutation
methods (``replace_component_source`` / ``insert_code``) and the real
:class:`~excel_vba_mcp.adapters.excel.vbide_adapter.VBIDEAdapter` against a
pure-Python fake ``Workbook`` / ``VBProject`` object graph (plus a fake STA
worker that runs the submitted callable inline).

Coverage:

* revision mismatch rejected without write;
* read-only session rejected with ``WORKBOOK_READ_ONLY``;
* snapshot-required gate (``SNAPSHOT_FAILED`` when no snapshot);
* semantic insertion position resolution (declarations / before / after);
* Document / UserForm rejection;
* approval deny-by-default.
"""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.vbide_adapter import hash_code
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, SnapshotFailedError
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

pytestmark = pytest.mark.anyio

# ---------------------------------------------------------------------------
# Test source strings
# ---------------------------------------------------------------------------
MOD_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Private Sub Helper()
    ' helper body
End Sub
"""

MOD_SOURCE_NEW = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Private Sub Helper()
    ' helper body v2
End Sub

Public Sub BrandNew()
    ' brand new procedure
End Sub
"""

CLS_SOURCE = """\
Option Explicit

Public FullName As String
"""


# ---------------------------------------------------------------------------
# Mutable fake CodeModule (supports DeleteLines / InsertLines / AddFromString)
# ---------------------------------------------------------------------------
class FakeMutableCodeModule:
    """Fake ``CodeModule`` that supports the VBIDE mutation APIs.

    Stores the code as a list of lines. ``DeleteLines``, ``InsertLines`` and
    ``AddFromString`` mutate the list in place so subsequent ``Lines`` reads
    reflect the changes.
    """

    def __init__(self, text: str) -> None:
        # Normalise to \n and keep as a list of lines (no trailing newline).
        normalised = text.replace("\r\n", "\n").replace("\r", "\n")
        self._lines: list[str] = normalised.split("\n")
        # Remove a trailing empty string from a final newline.
        if self._lines and self._lines[-1] == "":
            self._lines.pop()

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - COM casing
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:  # noqa: N802 - COM casing
        selected = self._lines[start - 1 : start - 1 + count]
        return "\n".join(selected)

    def DeleteLines(self, start: int, count: int) -> None:  # noqa: N802 - COM casing
        del self._lines[start - 1 : start - 1 + count]

    def InsertLines(self, line: int, text: str) -> None:  # noqa: N802 - COM casing
        new_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        for offset, new_line in enumerate(new_lines):
            self._lines.insert(line - 1 + offset, new_line)

    def AddFromString(self, text: str) -> None:  # noqa: N802 - COM casing
        new_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if self._lines and self._lines[-1] != "":
            self._lines.append("")
        self._lines.extend(new_lines)

    def ProcOfLine(self, line: int, kind: int = 0) -> tuple[str, int]:  # noqa: N802 - COM casing
        header = self._lines[line - 1].strip()
        lowered = header.casefold()
        name = header.split("(")[0].split()[-1]
        if "function" in lowered:
            return name, 1
        if "property" in lowered:
            return name, 2
        return name, 0

    def ProcStartLine(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        for index, line in enumerate(self._lines, start=1):
            is_declaration = line.strip().startswith(
                ("Sub ", "Function ", "Property ", "Public ", "Private ")
            )
            if is_declaration and name in line:
                return index
        raise RuntimeError(f"Procedure {name!r} not found")

    def ProcCountLines(self, name: str, kind: int = 0) -> int:  # noqa: N802 - COM casing
        start = 1
        for index, line in enumerate(self._lines, start=1):
            is_declaration = line.strip().startswith(
                ("Sub ", "Function ", "Property ", "Public ", "Private ")
            )
            if is_declaration and name in line:
                start = index
                break
        else:
            raise RuntimeError(f"Procedure {name!r} not found")
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
        return any(target in line for line in self._lines)


# ---------------------------------------------------------------------------
# Mutable fakes (Component / Project / Workbook / App / Worker)
# ---------------------------------------------------------------------------
class FakeMutableComponent:
    """Fake ``VBComponent`` with a mutable code module and exportable file body."""

    def __init__(self, name: str, type_code: int, code: str) -> None:
        self.Name = name
        self.Type = type_code
        self.CodeModule = FakeMutableCodeModule(code)

    def Export(self, path: str) -> None:  # noqa: N802 - COM casing
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        text = self.CodeModule.Lines(1, self.CodeModule.CountOfLines)
        target.write_text(text, encoding="utf-8")


class FakeMutableComponents:
    def __init__(self, components: list[FakeMutableComponent]) -> None:
        self._items = components

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> FakeMutableComponent:  # noqa: N802 - COM casing
        return self._items[index - 1]

    def __call__(self, name: str) -> FakeMutableComponent:
        for component in self._items:
            if component.Name == name:
                return component
        raise KeyError(name)

    def __iter__(self):
        return iter(self._items)


class FakeMutableProject:
    """Fake ``VBProject`` (protection, components)."""

    def __init__(self, components: list[FakeMutableComponent], *, protection: int = 0) -> None:
        self.Name = "FakeProject"
        self.Protection = protection
        self.VBComponents = FakeMutableComponents(components)


class FakeMutableWorkbook:
    """Fake ``Workbook`` with a mutable ``VBProject``."""

    def __init__(self, project: Any, *, read_only: bool = False) -> None:
        self._project = project
        self.ReadOnly = read_only
        self.FullName = ""

    @property
    def VBProject(self) -> Any:  # noqa: N802 - COM casing
        return self._project


class FakeMutableWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeMutableWorkbook] = []

    def Open(  # noqa: N802 - COM casing
        self, path: str, **kwargs: Any
    ) -> FakeMutableWorkbook:
        read_only = bool(kwargs.get("ReadOnly", False))
        wb = FakeMutableWorkbook(_MUTABLE_PROJECT, read_only=read_only)
        wb.FullName = path  # type: ignore[attr-defined]
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)


class FakeMutableApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeMutableWorkbooks()


class FakeMutationWorker:
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


# A shared mutable project graph (each test gets its own Rig, so no cross-test
# contamination).
_MUTABLE_PROJECT = FakeMutableProject(
    [
        FakeMutableComponent("modTest", 1, MOD_SOURCE),
        FakeMutableComponent("clsTest", 2, CLS_SOURCE),
        FakeMutableComponent("frmTest", 3, "Option Explicit\n"),
        FakeMutableComponent("ThisWorkbook", 100, "Option Explicit\n"),
        FakeMutableComponent("Sheet1", 100, "Option Explicit\n"),
    ]
)


# ---------------------------------------------------------------------------
# Fake snapshot service (minimal: only ``require_snapshot`` is needed)
# ---------------------------------------------------------------------------
class FakeSnapshotService:
    """Minimal fake for the pre-mutation snapshot gate."""

    def __init__(self, *, has_snapshot: bool = True) -> None:
        self._has_snapshot = has_snapshot
        self.require_snapshot_calls: list[str] = []

    def require_snapshot(self, workbook_id: str) -> None:
        self.require_snapshot_calls.append(workbook_id)
        if not self._has_snapshot:
            raise SnapshotFailedError(
                "A successful snapshot is required before mutating this workbook.",
                context={"workbook_id": workbook_id},
            )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class MutationRig:
    """A rig with a writable session, snapshot service, approval and audit."""

    def __init__(
        self,
        tmp_path: Path,
        *,
        read_only: bool = False,
        has_snapshot: bool = True,
        project: Any = _MUTABLE_PROJECT,
    ) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.app = FakeMutableApp()
        self.app.Workbooks.items.clear()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = FakeMutationWorker()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.snapshot = FakeSnapshotService(has_snapshot=has_snapshot)
        self.service = VBAProjectService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.path_policy,
            self.audit,
            snapshot_service=self.snapshot,  # type: ignore[arg-type]
            approval=self.approval,
        )
        self._project = project
        self._read_only = read_only

    async def open(self) -> Any:
        # Grant the open approval.
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
            str(self.original),
            mode=SessionMode.INSPECT,
            read_only=self._read_only,
        )
        # Point the fake workbook at this rig's project graph.
        wb = self.app.Workbooks.items[-1]
        wb._project = self._project
        return info

    def grant_mutation(self, component: str, expected_hash: str, action: str) -> None:
        """Grant a one-time approval for a specific mutation."""

        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(
                    canonical_path=self.canonical,
                    sha256=expected_hash,
                    action=action,
                    procedure=component,
                ),
                scope=ApprovalScope.ONE_TIME,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )


def _make_rig(tmp_path: Path, **kwargs: Any) -> MutationRig:
    """Create a rig with a fresh project graph (avoid cross-test contamination)."""

    project = FakeMutableProject(
        [
            FakeMutableComponent("modTest", 1, MOD_SOURCE),
            FakeMutableComponent("clsTest", 2, CLS_SOURCE),
            FakeMutableComponent("frmTest", 3, "Option Explicit\n"),
            FakeMutableComponent("ThisWorkbook", 100, "Option Explicit\n"),
            FakeMutableComponent("Sheet1", 100, "Option Explicit\n"),
        ]
    )
    return MutationRig(tmp_path, project=project, **kwargs)


# ---------------------------------------------------------------------------
# Tests: revision mismatch rejected without write
# ---------------------------------------------------------------------------
async def test_revision_mismatch_rejected_without_write(tmp_path: Path) -> None:
    """A stale expected_source_hash is rejected and the source is unchanged."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    actual_hash = source.source_hash

    # Pre-check reads the current hash; a wrong expected hash fails before write.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            MOD_SOURCE_NEW,
            expected_source_hash="0" * 64,
        )
    assert exc_info.value.context.get("reason") == "revision_conflict"

    # The source must be unchanged.
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash == actual_hash
    assert "BrandNew" not in after.source

    # No MUTATING submission was made (only the read-only pre-check).
    mutating_submits = [
        s
        for s in rig.worker.submits
        if s["op_class"] == OperationClass.MUTATING and s["name"] not in ("open_workbook",)
    ]
    assert mutating_submits == [], f"unexpected MUTATING submits: {mutating_submits}"


# ---------------------------------------------------------------------------
# Tests: read-only session rejected
# ---------------------------------------------------------------------------
async def test_read_only_session_rejected(tmp_path: Path) -> None:
    """A read-only session is rejected with ``WORKBOOK_READ_ONLY``."""

    rig = _make_rig(tmp_path, read_only=True)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            MOD_SOURCE_NEW,
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY

    # The source must be unchanged.
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash == source.source_hash


# ---------------------------------------------------------------------------
# Tests: snapshot-required gate
# ---------------------------------------------------------------------------
async def test_snapshot_required_gate_blocks_mutation(tmp_path: Path) -> None:
    """A mutation without a snapshot is blocked with ``SNAPSHOT_FAILED``."""

    rig = _make_rig(tmp_path, has_snapshot=False)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(SnapshotFailedError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            MOD_SOURCE_NEW,
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

    # The snapshot gate was checked.
    assert info.workbook_id in rig.snapshot.require_snapshot_calls

    # No MUTATING submission was made.
    mutating_submits = [
        s
        for s in rig.worker.submits
        if s["op_class"] == OperationClass.MUTATING and s["name"] not in ("open_workbook",)
    ]
    assert mutating_submits == []


# ---------------------------------------------------------------------------
# Tests: Document / UserForm rejection
# ---------------------------------------------------------------------------
async def test_document_module_rejected(tmp_path: Path) -> None:
    """Replacing a Document module (type 100) is rejected."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "ThisWorkbook")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "ThisWorkbook",
            "Option Explicit\n",
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.context.get("reason") == "non_replaceable_component"

    # No MUTATING submission was made.
    mutating_submits = [
        s
        for s in rig.worker.submits
        if s["op_class"] == OperationClass.MUTATING and s["name"] not in ("open_workbook",)
    ]
    assert mutating_submits == []


async def test_userform_rejected(tmp_path: Path) -> None:
    """Replacing a UserForm (type 3) is rejected."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "frmTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "frmTest",
            "Option Explicit\n",
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.context.get("reason") == "non_replaceable_component"


# ---------------------------------------------------------------------------
# Tests: approval deny-by-default
# ---------------------------------------------------------------------------
async def test_approval_deny_by_default(tmp_path: Path) -> None:
    """A mutation without an approval grant is denied."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            MOD_SOURCE_NEW,
            expected_source_hash=source.source_hash,
        )
    assert (
        exc_info.value.context.get("reason")
        == "denied by default: no matching allowlist entry or live grant"
    )

    # No MUTATING submission was made.
    mutating_submits = [
        s
        for s in rig.worker.submits
        if s["op_class"] == OperationClass.MUTATING and s["name"] not in ("open_workbook",)
    ]
    assert mutating_submits == []


# ---------------------------------------------------------------------------
# Tests: successful replace
# ---------------------------------------------------------------------------
async def test_replace_component_source_success(tmp_path: Path) -> None:
    """A successful replace produces before/after hashes and export paths."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    # Grant a one-time approval for this specific mutation.
    rig.grant_mutation("modTest", expected_hash, "vba_replace_component_source")

    result = await rig.service.replace_component_source(
        info.workbook_id,
        "modTest",
        MOD_SOURCE_NEW,
        expected_source_hash=expected_hash,
    )

    # Before hash matches the original.
    assert result.before_hash == expected_hash
    # After hash is different (new source was written).
    assert result.after_hash != expected_hash
    assert result.after_hash == hash_code(MOD_SOURCE_NEW)

    # Export paths exist.
    before_path = Path(result.before_export_path)
    after_path = Path(result.after_export_path)
    assert before_path.is_file()
    assert after_path.is_file()

    # The source actually changed.
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert "BrandNew" in after.source
    assert after.source_hash == result.after_hash

    # A MUTATING submission was made.
    mutating_submits = [
        s
        for s in rig.worker.submits
        if s["op_class"] == OperationClass.MUTATING and s["name"] not in ("open_workbook",)
    ]
    assert len(mutating_submits) == 1
    assert mutating_submits[0]["name"] == "vba_replace_component_source"


# ---------------------------------------------------------------------------
# Tests: semantic insertion position resolution
# ---------------------------------------------------------------------------
async def test_insert_declarations(tmp_path: Path) -> None:
    """Insert at ``declarations`` inserts before the first procedure."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_insert_code")

    insert_text = "Public Const MAX_RETRIES As Long = 3"
    result = await rig.service.insert_code(
        info.workbook_id,
        "modTest",
        insert_text,
        position="declarations",
        expected_source_hash=expected_hash,
    )

    after = await rig.service.read_component(info.workbook_id, "modTest")
    # The inserted constant must appear before the first procedure.
    const_line = after.source.find(insert_text)
    add_line = after.source.find("AddNumbers")
    assert const_line != -1, "inserted text not found"
    assert add_line != -1, "AddNumbers not found"
    assert const_line < add_line, "inserted text must appear before the first procedure"

    # Before/after hashes differ.
    assert result.before_hash == expected_hash
    assert result.after_hash != expected_hash


async def test_insert_before_procedure(tmp_path: Path) -> None:
    """Insert before a named procedure inserts before its start line."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_insert_code")

    insert_text = "' inserted comment before Helper"
    result = await rig.service.insert_code(
        info.workbook_id,
        "modTest",
        insert_text,
        position="before_procedure:Helper",
        expected_source_hash=expected_hash,
    )

    after = await rig.service.read_component(info.workbook_id, "modTest")
    insert_line = after.source.find(insert_text)
    helper_line = after.source.find("Helper")
    assert insert_line != -1, "inserted text not found"
    assert helper_line != -1, "Helper not found"
    assert insert_line < helper_line, "inserted text must appear before Helper"

    assert result.before_hash == expected_hash
    assert result.after_hash != expected_hash


async def test_insert_after_procedure(tmp_path: Path) -> None:
    """Insert after a named procedure inserts after its end line."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_insert_code")

    insert_text = "' inserted comment after AddNumbers"
    result = await rig.service.insert_code(
        info.workbook_id,
        "modTest",
        insert_text,
        position="after_procedure:AddNumbers",
        expected_source_hash=expected_hash,
    )

    after = await rig.service.read_component(info.workbook_id, "modTest")
    insert_line = after.source.find(insert_text)
    helper_line = after.source.find("Helper")
    add_line = after.source.find("AddNumbers")
    assert insert_line != -1, "inserted text not found"
    assert add_line != -1, "AddNumbers not found"
    assert helper_line != -1, "Helper not found"
    # The insert must appear after AddNumbers but before Helper.
    assert insert_line > add_line, "inserted text must appear after AddNumbers"
    assert insert_line < helper_line, "inserted text must appear before Helper"

    assert result.before_hash == expected_hash
    assert result.after_hash != expected_hash


async def test_insert_ambiguous_position_rejected(tmp_path: Path) -> None:
    """An unresolvable position is rejected with a structured error."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_insert_code")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            "' comment",
            position="before_procedure:NonExistent",
            expected_source_hash=expected_hash,
        )
    assert exc_info.value.context.get("reason") == "proc_start_line_failed"

    # The source must be unchanged.
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash == expected_hash


# ---------------------------------------------------------------------------
# Tests: insert also enforces snapshot + revision + approval
# ---------------------------------------------------------------------------
async def test_insert_revision_mismatch_rejected(tmp_path: Path) -> None:
    """Insert with a stale expected_source_hash is rejected without write."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            "' comment",
            position="declarations",
            expected_source_hash="0" * 64,
        )
    assert exc_info.value.context.get("reason") == "revision_conflict"

    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash == source.source_hash


async def test_insert_read_only_rejected(tmp_path: Path) -> None:
    """Insert into a read-only session is rejected."""

    rig = _make_rig(tmp_path, read_only=True)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            "' comment",
            position="declarations",
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_insert_snapshot_required(tmp_path: Path) -> None:
    """Insert without a snapshot is blocked."""

    rig = _make_rig(tmp_path, has_snapshot=False)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(SnapshotFailedError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            "' comment",
            position="declarations",
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_insert_document_module_rejected(tmp_path: Path) -> None:
    """Insert into a Document module is rejected."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "ThisWorkbook")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "ThisWorkbook",
            "' comment",
            position="declarations",
            expected_source_hash=source.source_hash,
        )
    assert exc_info.value.context.get("reason") == "non_replaceable_component"


async def test_insert_approval_deny_by_default(tmp_path: Path) -> None:
    """Insert without an approval grant is denied."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            "' comment",
            position="declarations",
            expected_source_hash=source.source_hash,
        )
    assert (
        exc_info.value.context.get("reason")
        == "denied by default: no matching allowlist entry or live grant"
    )


# ---------------------------------------------------------------------------
# Tests: original file is never modified
# ---------------------------------------------------------------------------
async def test_original_file_unchanged_after_mutation(tmp_path: Path) -> None:
    """The original file's bytes are unchanged after a successful mutation."""

    rig = _make_rig(tmp_path)
    original_hash = hashlib.sha256(rig.original.read_bytes()).hexdigest()

    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_replace_component_source")
    await rig.service.replace_component_source(
        info.workbook_id,
        "modTest",
        MOD_SOURCE_NEW,
        expected_source_hash=expected_hash,
    )

    # The original file on disk is untouched (only the working copy was mutated).
    after_hash = hashlib.sha256(rig.original.read_bytes()).hexdigest()
    assert after_hash == original_hash


# ---------------------------------------------------------------------------
# Tests: one-time grant is consumed after use
# ---------------------------------------------------------------------------
async def test_one_time_grant_consumed_after_use(tmp_path: Path) -> None:
    """A ONE_TIME grant is consumed after a successful mutation."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_replace_component_source")
    await rig.service.replace_component_source(
        info.workbook_id,
        "modTest",
        MOD_SOURCE_NEW,
        expected_source_hash=expected_hash,
    )

    # The source has changed, so a second mutation with the old hash fails on
    # revision, and the one-time grant has been consumed (a new grant is needed).
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash != expected_hash

    with pytest.raises(ExcelVbaMcpError):
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            "Option Explicit\n",
            expected_source_hash=expected_hash,
        )


# ---------------------------------------------------------------------------
# Tests: F7 hash clarity — the snapshot export FILE hash is NOT a valid revision
# token; only the normalized source_hash from vba_read_component is accepted.
# ---------------------------------------------------------------------------
async def test_snapshot_export_file_hash_rejected_as_source_hash(tmp_path: Path) -> None:
    """A snapshot export file hash yields REVISION_CONFLICT naming source_hash.

    F7: ``snapshot_create.vba_exports[].export_file_sha256`` is the SHA-256 of the
    exported ``.bas`` **file bytes** (here: the ``Attribute VB_Name`` header plus the
    source). It differs from the normalized ``source_hash`` that the mutation guard
    compares against, so using it as ``expected_source_hash`` must be rejected with
    a clear revision conflict whose ``current_source_hash`` is the normalized hash.
    """

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")

    # Simulate the exported file bytes (header + body); its hash is the snapshot
    # export file hash a confused caller might pass as the revision token.
    export_file_bytes = ('Attribute VB_Name = "modTest"\r\n' + MOD_SOURCE).encode("utf-8")
    export_file_hash = hashlib.sha256(export_file_bytes).hexdigest()

    # The file hash and the normalized source hash genuinely differ (the F7 trap).
    assert export_file_hash != source.source_hash

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            MOD_SOURCE_NEW,
            expected_source_hash=export_file_hash,
        )

    context = exc_info.value.context
    assert context.get("reason") == "revision_conflict"
    # The conflict names the normalized source_hash the guard actually uses.
    assert context.get("current_source_hash") == source.source_hash
    assert context.get("expected_source_hash") == export_file_hash

    # The source is unchanged (rejected before any write).
    after = await rig.service.read_component(info.workbook_id, "modTest")
    assert after.source_hash == source.source_hash
    assert "BrandNew" not in after.source


async def test_normalized_source_hash_succeeds(tmp_path: Path) -> None:
    """The correct normalized ``source_hash`` is accepted (F7 companion check)."""

    rig = _make_rig(tmp_path)
    info = await rig.open()
    source = await rig.service.read_component(info.workbook_id, "modTest")
    expected_hash = source.source_hash

    rig.grant_mutation("modTest", expected_hash, "vba_replace_component_source")
    result = await rig.service.replace_component_source(
        info.workbook_id,
        "modTest",
        MOD_SOURCE_NEW,
        expected_source_hash=expected_hash,
    )

    assert result.before_hash == expected_hash
    assert result.after_hash == hash_code(MOD_SOURCE_NEW)
