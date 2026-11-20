"""Unit tests for the bounded declarative test runner (task_035).

The runner is exercised against the **real** session / snapshot / artifact /
run-repository services and a pure-Python fake Excel object graph (a fake STA
worker that runs each submitted callable inline). No Excel / COM is required.

Covered:

* the closed vocabulary rejects an unknown action / assertion **before** any
  Excel call (no worker submit, no service call);
* every case runs on its own disposable copy of one immutable baseline, so one
  case's mutation cannot leak into the next and the baseline file is unchanged;
* a failed compile step blocks the case's macro steps (the macro is not called);
  an ``indeterminate`` compile does **not** block (no compile error was proven);
* a policy-blocked macro blocks the step and the case (``blocked``, no failure);
* status mapping (passed / failed / blocked / indeterminate) and run persistence
  (run document, per-case diagnostics, artifact references, ``test_get_run``);
* actions outside the runner's dependency set fail closed (no service call).
"""

from __future__ import annotations

import hashlib
import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest
from pydantic import ValidationError

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.tables_adapter import TableMutationResult
from excel_vba_mcp.adapters.windows_ui.vbe_compile_ui import CommandDiscovery
from excel_vba_mcp.application.compile_service import CompileDiagnostics, CompileResult
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.test_runner import (
    ERROR_COMPILE_FAILED,
    ERROR_MACROS_BLOCKED_BY_COMPILE,
    ERROR_UNSUPPORTED_ACTION,
    EXTERNAL_SIDE_EFFECTS_NOTE,
)
from excel_vba_mcp.application.test_runner import TestRunner as Runner
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, PathNotAllowedError
from excel_vba_mcp.domain.models import (
    AssertionResult,
    MacroRunResult,
    SheetInfo,
    TableColumnInfo,
    TableInfo,
)
from excel_vba_mcp.domain.statuses import (
    CaseStatus,
    CompileStatus,
    MacroRunStatus,
    OperationClass,
    SessionMode,
    SheetObjectType,
    SheetVisibility,
    StepStatus,
    WorkbookSessionState,
)
from excel_vba_mcp.domain.statuses import (
    TestRunStatus as RunStatus,
)
from excel_vba_mcp.domain.test_spec import AssertionType, TestAction, TestSuiteSpec
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
from excel_vba_mcp.storage.run_repository import TestRunNotFoundError as RunNotFoundError
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = pytest.mark.anyio

#: Marker prefix of the fake workbook file format (JSON state after it).
_FAKE_PREFIX = b"FAKE-XLSM::"

#: Excel visibility codes (xlSheetVisible = -1, xlSheetHidden = 0).
_XL_VISIBLE = -1
_XL_HIDDEN = 0


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph (file-backed, so a copy really is a fresh baseline)
# ---------------------------------------------------------------------------
def _encode(state: dict[str, dict[str, Any]]) -> bytes:
    """Serialize a fake workbook state (``{sheet: {cell: value}}``)."""

    return _FAKE_PREFIX + json.dumps(state, sort_keys=True).encode("utf-8")


def _decode(raw: bytes) -> dict[str, dict[str, Any]]:
    """Deserialize a fake workbook state (empty when the file is opaque)."""

    if not raw.startswith(_FAKE_PREFIX):
        return {}
    try:
        payload = json.loads(raw[len(_FAKE_PREFIX) :].decode("utf-8"))
    except ValueError:
        return {}
    if not isinstance(payload, dict):
        return {}
    return {str(name): dict(cells) for name, cells in payload.items()}


class _FakeRange:
    """A fake ``Range`` exposing a single cell's ``Value``."""

    def __init__(self, sheet: _FakeSheet, address: str) -> None:
        self._sheet = sheet
        self.Address = address

    @property
    def Value(self) -> Any:  # noqa: N802 - COM casing
        return self._sheet.cells.get(self.Address.upper())

    @Value.setter
    def Value(self, value: Any) -> None:  # noqa: N802 - COM casing
        self._sheet.cells[self.Address.upper()] = value


class _FakeSheet:
    """A fake worksheet (name, visibility and a sparse cell map)."""

    def __init__(
        self, name: str, cells: dict[str, Any] | None = None, visible: int = _XL_VISIBLE
    ) -> None:
        self.Name = name
        self.CodeName = f"{name}_Code"
        self.Type = -4167  # xlWorksheet
        self.Visible = visible
        self.ProtectContents = False
        self.Index = 0
        self.cells: dict[str, Any] = dict(cells or {})

    def Range(self, address: str) -> _FakeRange:  # noqa: N802 - COM casing
        return _FakeRange(self, address)


class _FakeSheets:
    """A fake ``Worksheets`` / ``Sheets`` collection (iterable and callable)."""

    def __init__(self, sheets: list[_FakeSheet]) -> None:
        self._items = sheets

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def __call__(self, key: Any) -> _FakeSheet:
        if isinstance(key, int):
            return self._items[key - 1]
        for sheet in self._items:
            if sheet.Name == key:
                return sheet
        raise KeyError(key)

    def Item(self, key: Any) -> _FakeSheet:  # noqa: N802 - COM casing
        return self(key)

    def __iter__(self) -> Any:
        return iter(self._items)

    def Add(self, Type: Any = None) -> _FakeSheet:  # noqa: N802, N803 - COM casing
        name = f"Sheet{len(self._items) + 1}"
        sheet = _FakeSheet(name)
        self._items.append(sheet)
        for index, item in enumerate(self._items, start=1):
            item.Index = index
        return sheet


class _FakeWorkbook:
    """A fake ``Workbook`` whose state round-trips through the file on disk."""

    def __init__(self, app: _FakeApp, path: str, state: dict[str, dict[str, Any]]) -> None:
        self._app = app
        self.FullName = path
        self.Name = Path(path).name
        self.ReadOnly = False
        self.Saved = True
        self.Application = app
        self._sheets = [_FakeSheet(name, cells) for name, cells in state.items()]

    @property
    def Sheets(self) -> _FakeSheets:  # noqa: N802 - COM casing
        return _FakeSheets(self._sheets)

    @property
    def Worksheets(self) -> _FakeSheets:  # noqa: N802 - COM casing
        return _FakeSheets(self._sheets)

    def _state(self) -> dict[str, dict[str, Any]]:
        return {sheet.Name: dict(sheet.cells) for sheet in self._sheets}

    def SaveCopyAs(self, filename: Any) -> None:  # noqa: N802 - COM casing
        Path(str(filename)).write_bytes(_encode(self._state()))

    def Save(self) -> None:  # noqa: N802 - COM casing
        Path(self.FullName).write_bytes(_encode(self._state()))

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM
        self._app.Workbooks.forget(self)


class _FakeWorkbooks:
    """A fake ``Workbooks`` collection that loads state from the file on disk."""

    def __init__(self, app: _FakeApp) -> None:
        self._app = app
        self.items: list[_FakeWorkbook] = []

    def Open(self, path: Any, **_kwargs: Any) -> _FakeWorkbook:  # noqa: N802 - COM
        target = str(path)
        raw = Path(target).read_bytes() if Path(target).is_file() else b""
        workbook = _FakeWorkbook(self._app, target, _decode(raw))
        self.items.append(workbook)
        return workbook

    def forget(self, workbook: _FakeWorkbook) -> None:
        if workbook in self.items:
            self.items.remove(workbook)

    def __iter__(self) -> Any:
        return iter(self.items)


class _FakeApp:
    """A fake Excel ``Application`` (hidden, never activates anything)."""

    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = _FakeWorkbooks(self)


class _FakeWorker:
    """Fake STA worker: runs the submitted callable inline (no thread, no COM)."""

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
        self.submits.append({"name": getattr(fn, "__name__", "?"), "op_class": op_class})
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# ---------------------------------------------------------------------------
# Service stubs (only the surface the runner dispatches to)
# ---------------------------------------------------------------------------
class _CompileStub:
    """Compile service stub with an injectable status / error."""

    def __init__(
        self,
        status: CompileStatus = CompileStatus.INDETERMINATE,
        *,
        error: ExcelVbaMcpError | None = None,
    ) -> None:
        self.status = status
        self.error = error
        self.calls: list[str] = []

    async def compile_project(self, workbook_id: str, **_kwargs: Any) -> CompileResult:
        self.calls.append(workbook_id)
        if self.error is not None:
            raise self.error
        now = datetime.now(UTC)
        return CompileResult(
            workbook_id=workbook_id,
            status=self.status,
            reason="STUB",
            diagnostics=CompileDiagnostics(),
            command_discovery=CommandDiscovery(),
            created_at=now,
            updated_at=now,
        )


class _MacroStub:
    """Macro service stub returning a canned result and recording every call."""

    def __init__(self, status: MacroRunStatus = MacroRunStatus.SUCCESS) -> None:
        self.status = status
        self.calls: list[tuple[str, str, list[Any]]] = []

    async def run_macro(
        self, workbook_id: str, macro_name: str, args: list[Any] | None = None, **_kwargs: Any
    ) -> MacroRunResult:
        self.calls.append((workbook_id, macro_name, list(args or [])))
        now = datetime.now(UTC)
        return MacroRunResult(
            run_id=uuid4().hex,
            macro_name=macro_name,
            status=self.status,
            return_value=None,
            started_at=now,
            finished_at=now,
            duration_ms=1,
            created_at=now,
            updated_at=now,
        )


class _SheetsStub:
    """Sheets service stub that really mutates the fake workbook on disk."""

    def __init__(self, app: _FakeApp, session: SessionService) -> None:
        self._app = app
        self._session = session
        self.calls: list[tuple[str, str]] = []
        self.open_workbook_ids: list[str] = []
        #: Sheet names seen the first time a workbook is touched (isolation probe).
        self.first_sheets: dict[str, list[str]] = {}
        #: File hash of each workbook right after a mutation.
        self.post_mutation_hash: dict[str, str] = {}

    def _workbook(self, workbook_id: str) -> _FakeWorkbook:
        record = self._session._records[workbook_id]
        for workbook in self._app.Workbooks:
            if str(workbook.FullName).lower() == record.working_path.lower():
                if workbook_id not in self.first_sheets:
                    self.first_sheets[workbook_id] = [sheet.Name for sheet in workbook.Worksheets]
                    self.open_workbook_ids.append(workbook_id)
                return workbook
        raise AssertionError(f"workbook {workbook_id} is not open")

    def _hash_workbook(self, workbook_id: str) -> str:
        record = self._session._records[workbook_id]
        return hashlib.sha256(Path(record.working_path).read_bytes()).hexdigest()

    async def add_sheet(
        self, workbook_id: str, name: str, sheet_type: str = "worksheet"
    ) -> SheetInfo:
        self.calls.append((workbook_id, name))
        workbook = self._workbook(workbook_id)
        workbook.Worksheets.Add()
        added = workbook.Worksheets(workbook.Worksheets.Count)
        added.Name = name
        workbook.Save()
        self.post_mutation_hash[workbook_id] = self._hash_workbook(workbook_id)
        return SheetInfo(
            name=name,
            index=added.Index,
            object_type=SheetObjectType.WORKSHEET,
            visibility=SheetVisibility.VISIBLE,
        )

    async def get_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        workbook = self._workbook(workbook_id)
        sheet = workbook.Worksheets(name)
        return SheetInfo(
            name=sheet.Name,
            index=sheet.Index or 1,
            object_type=SheetObjectType.WORKSHEET,
            visibility=(
                SheetVisibility.VISIBLE if sheet.Visible == _XL_VISIBLE else SheetVisibility.HIDDEN
            ),
        )

    async def set_visibility(self, workbook_id: str, name: str, visible: bool) -> SheetInfo:
        workbook = self._workbook(workbook_id)
        sheet = workbook.Worksheets(name)
        sheet.Visible = _XL_VISIBLE if visible else _XL_HIDDEN
        return SheetInfo(
            name=sheet.Name,
            index=sheet.Index or 1,
            object_type=SheetObjectType.WORKSHEET,
            visibility=SheetVisibility.VISIBLE if visible else SheetVisibility.HIDDEN,
        )

    async def rename_sheet(self, workbook_id: str, name: str, new_name: str) -> SheetInfo:
        workbook = self._workbook(workbook_id)
        sheet = workbook.Worksheets(name)
        sheet.Name = new_name
        return SheetInfo(
            name=new_name,
            index=sheet.Index or 1,
            object_type=SheetObjectType.WORKSHEET,
            visibility=SheetVisibility.VISIBLE,
        )

    async def copy_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        workbook = self._workbook(workbook_id)
        workbook.Worksheets(name)
        sheet = workbook.Worksheets.Add()
        return SheetInfo(
            name=sheet.Name,
            index=sheet.Index or 1,
            object_type=SheetObjectType.WORKSHEET,
            visibility=SheetVisibility.VISIBLE,
        )

    async def delete_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        workbook = self._workbook(workbook_id)
        sheet = workbook.Worksheets(name)
        return SheetInfo(
            name=sheet.Name,
            index=sheet.Index or 1,
            object_type=SheetObjectType.WORKSHEET,
            visibility=SheetVisibility.VISIBLE,
        )

    async def list_sheets(self, workbook_id: str) -> list[SheetInfo]:
        workbook = self._workbook(workbook_id)
        return [
            SheetInfo(
                name=sheet.Name,
                index=sheet.Index,
                object_type=SheetObjectType.WORKSHEET,
                visibility=SheetVisibility.VISIBLE,
            )
            for sheet in workbook.Worksheets
        ]


class _TablesStub:
    """Tables service stub returning a one-row table (no real ListObjects)."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, ...]] = []
        self.rows = 1

    async def get_table(self, workbook_id: str, name: str) -> TableInfo:
        self.calls.append(("get_table", workbook_id, name))
        return TableInfo(
            name=name,
            sheet_name="Sheet1",
            range_address="$A$1:$B$2",
            row_count=self.rows,
            columns=[TableColumnInfo(index=1, name="Key"), TableColumnInfo(index=2, name="Value")],
        )

    async def append_table_rows(
        self, workbook_id: str, name: str, rows: list[dict[str, Any]]
    ) -> TableMutationResult:
        self.calls.append(("append_table_rows", workbook_id, name, str(len(rows))))
        self.rows += len(rows)
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="append_rows", before=None, after=after, affected_rows=[2]
        )

    async def insert_table_row(
        self, workbook_id: str, name: str, at: int, values: dict[str, Any] | None = None
    ) -> TableMutationResult:
        self.calls.append(("insert_table_row", workbook_id, name, str(at)))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="insert_row", before=None, after=after, affected_rows=[at]
        )

    async def update_table_rows(
        self, workbook_id: str, name: str, updates: list[dict[str, Any]]
    ) -> TableMutationResult:
        self.calls.append(("update_table_rows", workbook_id, name, str(len(updates))))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="update_rows", before=None, after=after, affected_rows=[1]
        )

    async def delete_table_rows(
        self,
        workbook_id: str,
        name: str,
        indices: list[int] | None = None,
        where: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        self.calls.append(("delete_table_rows", workbook_id, name, str(indices)))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="delete_rows", before=None, after=after, affected_rows=[1]
        )

    async def add_table_column(
        self, workbook_id: str, name: str, header: str, position: int | None = None
    ) -> TableMutationResult:
        self.calls.append(("add_column", workbook_id, name, header))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="add_column", before=None, after=after, affected_rows=[]
        )

    async def delete_table_column(
        self, workbook_id: str, name: str, header: str
    ) -> TableMutationResult:
        self.calls.append(("delete_column", workbook_id, name, header))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="delete_column", before=None, after=after, affected_rows=[]
        )

    async def create_table(
        self,
        workbook_id: str,
        sheet_name: str,
        range_or_name: str,
        *,
        style: str | None = None,
    ) -> TableMutationResult:
        self.calls.append(("create_table", workbook_id, sheet_name, range_or_name))
        return TableMutationResult(
            name=range_or_name.replace("$", ""),
            action="create",
            before=None,
            after=await self.get_table(workbook_id, "t"),
            affected_rows=[],
        )

    async def resize_table(
        self, workbook_id: str, name: str, new_range: str
    ) -> TableMutationResult:
        self.calls.append(("resize", workbook_id, name, new_range))
        after = await self.get_table(workbook_id, name)
        return TableMutationResult(
            name=name, action="resize", before=None, after=after, affected_rows=[]
        )

    async def list_tables(self, workbook_id: str) -> list[TableInfo]:
        return []


class _NamesStub:
    """Names service stub recording calls."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, ...]] = []

    async def list_names(self, workbook_id: str) -> list[Any]:
        self.calls.append(("list_names", workbook_id))
        return []

    async def get_name(self, workbook_id: str, name: str, scope: Any, *, sheet: str | None = None):
        self.calls.append(("get_name", workbook_id, name))
        raise ExcelVbaMcpError("stub: name not found", code=ErrorCode.PATH_NOT_ALLOWED)


# ---------------------------------------------------------------------------
# Rig
# ---------------------------------------------------------------------------
class Rig:
    """Real session / snapshot / artifact / run services over a fake Excel graph."""

    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.work_dir = tmp_path / "work"
        self.snapshots_dir = tmp_path / "snapshots"
        self.artifacts_dir = tmp_path / "artifacts"
        self.runs_dir = tmp_path / "runs"
        self.path_policy = PathPolicy(
            {str(tmp_path): str(tmp_path), "artifacts": str(self.artifacts_dir)}
        )
        self.approval = ApprovalPolicy()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.app = _FakeApp()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = _FakeWorker()
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )
        self.snapshot_service = SnapshotService(
            SnapshotStore(self.snapshots_dir),
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )
        self.artifacts = ArtifactStore(self.artifacts_dir, self.path_policy, self.audit)
        self.runs = RunRepository(self.runs_dir)
        self.compile = _CompileStub()
        self.macro = _MacroStub()
        self.sheets = _SheetsStub(self.app, self.session)
        self.tables = _TablesStub()
        self.names = _NamesStub()
        self.runner = Runner(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot_service,
            self.compile,  # type: ignore[arg-type]
            self.macro,  # type: ignore[arg-type]
            self.sheets,  # type: ignore[arg-type]
            self.tables,  # type: ignore[arg-type]
            self.names,  # type: ignore[arg-type]
            self.artifacts,
            self.runs,
            self.path_policy,
            self.audit,
        )

    # -- setup helpers ---------------------------------------------------
    def make_original(self) -> Path:
        """Write a fake macro-enabled workbook into the allowlisted root."""

        original = self.tmp_path / "fixture.xlsm"
        original.write_bytes(_encode({"Sheet1": {"A1": "base", "B1": 7}}))
        return original

    def grant(self, canonical_path: str, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=canonical_path, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task035-tests",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def open_original(self) -> Any:
        """Open the original in EXECUTE mode (with the required grant)."""

        original = self.make_original()
        canonical = str(self.path_policy.validate_path(str(original), OperationClass.MUTATING))
        self.grant(canonical, "excel_open")
        return await self.session.open_workbook(
            canonical, mode=SessionMode.EXECUTE, read_only=False
        )

    async def create_baseline(self, workbook_id: str) -> Any:
        """Create a baseline snapshot and grant the runner the right to open it."""

        manifest = await self.snapshot_service.create_snapshot(workbook_id)
        canonical = str(
            self.path_policy.validate_path(manifest.stored_path, OperationClass.MUTATING)
        )
        self.grant(canonical, "excel_open")
        return manifest


def _suite(cases: list[dict[str, Any]], *, baseline: str = "create") -> TestSuiteSpec:
    """Build a validated suite document."""

    return TestSuiteSpec.model_validate(
        {
            "suite": "unit-suite",
            "workbook_id": "wb-placeholder",
            "baseline": baseline,
            "cases": cases,
        }
    )


# ---------------------------------------------------------------------------
# Closed vocabulary: rejected before any Excel call
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "step",
    [
        {"action": "os.system", "sheet": "x"},
        {"action": "shell_exec", "source": "calc.exe"},
        {"action": "eval_python", "source": "1+1"},
        {"action": "macro_run", "macro": "modTest.Bad", "code": "Shell 'calc'"},
    ],
)
def test_unknown_action_is_rejected_by_the_schema(step: dict[str, Any]) -> None:
    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate(
            {
                "suite": "s",
                "workbook_id": "w",
                "cases": [{"id": "c", "steps": [step], "assertions": []}],
            }
        )


@pytest.mark.parametrize(
    "assertion",
    [
        {"type": "python_eval", "expected": "1"},
        {"type": "shell_exec", "expected": "calc"},
        {"type": "read_cell", "expected": "1", "code": "__import__('os')"},
    ],
)
def test_unknown_assertion_is_rejected_by_the_schema(assertion: dict[str, Any]) -> None:
    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate(
            {
                "suite": "s",
                "workbook_id": "w",
                "cases": [{"id": "c", "steps": [], "assertions": [assertion]}],
            }
        )


async def test_unknown_action_never_reaches_excel(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    submits_before = len(rig.worker.submits)

    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate(
            {
                "suite": "s",
                "workbook_id": info.workbook_id,
                "cases": [{"id": "c", "steps": [{"action": "os.system"}], "assertions": []}],
            }
        )

    # No snapshot, no worker submit, no service call happened.
    assert rig.snapshot_service.store.list() == []
    assert len(rig.worker.submits) == submits_before
    assert rig.macro.calls == []
    assert rig.sheets.calls == []


async def test_assertion_allowlist_is_closed(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()

    assert TestAction.MACRO_RUN.value == "macro_run"
    assert TestAction.SHAPE_SET_ON_ACTION.value == "shape_set_on_action"
    assert AssertionType.READ_CELL.value == "read_cell"
    assert AssertionType.TABLE_ROW_COUNT.value == "table_row_count"
    assert AssertionType.TABLE_COLUMN_COUNT.value == "table_column_count"

    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate(
            {
                "suite": "s",
                "workbook_id": info.workbook_id,
                "cases": [
                    {
                        "id": "c",
                        "steps": [],
                        "assertions": [{"type": "unlink_or_rm_rf", "path": "x"}],
                    }
                ],
            }
        )


# ---------------------------------------------------------------------------
# Snapshot isolation: one disposable copy per case
# ---------------------------------------------------------------------------
async def test_each_case_gets_its_own_disposable_copy(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    baseline = await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "case-a",
                "steps": [{"action": "sheet_add", "sheet": "A"}],
                "assertions": [{"type": "sheet_exists", "name": "A"}],
            },
            {
                "id": "case-b",
                "steps": [{"action": "sheet_add", "sheet": "B"}],
                "assertions": [{"type": "sheet_exists", "name": "Sheet1"}],
            },
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)

    assert result.status is RunStatus.PASSED
    assert result.baseline_snapshot_id == baseline.snapshot_id
    # Two distinct workbook ids: one disposable copy per case.
    assert len(set(rig.sheets.open_workbook_ids)) == 2
    # Every case really started from the baseline hash.
    for case in result.cases:
        assert f"working_copy_sha256={baseline.sha256}" in " ".join(case.notes)
    # Every case's working copy is closed afterwards (only the original remains).
    assert [w.workbook_id for w in rig.session.list_workbooks_sync()] == [info.workbook_id]
    # The immutable baseline file is byte-identical to what it was.
    assert _sha256(Path(baseline.stored_path)) == baseline.sha256


async def test_one_cases_mutation_cannot_leak_into_the_next(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    baseline = await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "mutate",
                "steps": [{"action": "sheet_add", "sheet": "Leaked"}],
                "assertions": [{"type": "sheet_exists", "name": "Leaked"}],
            },
            {
                "id": "verify",
                "steps": [],
                "assertions": [
                    {"type": "read_cell", "sheet": "Sheet1", "cell": "A1", "expected": "base"},
                    {"type": "sheet_exists", "name": "Sheet1"},
                ],
            },
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)

    mutate_case, verify_case = result.cases
    assert mutate_case.status is CaseStatus.PASSED
    assert verify_case.status is CaseStatus.PASSED
    # Each case ran on a different workbook, and both started from the pristine
    # baseline sheet set (the "Leaked" sheet never reached the second case).
    mutate_wb, verify_wb = rig.sheets.open_workbook_ids
    assert mutate_wb != verify_wb
    assert rig.sheets.first_sheets[mutate_wb] == ["Sheet1"]
    assert rig.sheets.first_sheets[verify_wb] == ["Sheet1"]
    # The mutation was real on the mutating case's own copy ...
    assert rig.sheets.post_mutation_hash[mutate_wb] != baseline.sha256
    # ... and the second case never mutated its copy.
    assert verify_wb not in rig.sheets.post_mutation_hash
    # Both cases recorded that they started from the baseline hash.
    for case in result.cases:
        assert f"working_copy_sha256={baseline.sha256}" in " ".join(case.notes)
    # The immutable baseline file and the original were never touched.
    assert _sha256(Path(baseline.stored_path)) == baseline.sha256
    assert _sha256(Path(info.source_path)) == _sha256(Path(baseline.stored_path))


# ---------------------------------------------------------------------------
# Compile gates macros
# ---------------------------------------------------------------------------
async def test_compile_failure_blocks_macro_steps(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)
    rig.compile = _CompileStub(CompileStatus.COMPILE_ERROR)
    rig.runner = _rebuild_runner(rig)

    suite = _suite(
        [
            {
                "id": "gated",
                "steps": [
                    {"action": "vba_compile"},
                    {"action": "macro_run", "macro": "modTest.AddNumbers", "args": [1, 2]},
                ],
                "assertions": [{"type": "macro_status", "expected": "success"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]
    compile_step, macro_step = case.steps

    assert compile_step.status is StepStatus.FAILED
    assert compile_step.error_code == ERROR_COMPILE_FAILED
    assert macro_step.status is StepStatus.SKIPPED
    assert macro_step.details["blocked"] is True
    assert macro_step.error_code == ERROR_MACROS_BLOCKED_BY_COMPILE
    # The macro was never dispatched.
    assert rig.macro.calls == []
    # The case is failed (the compile failure is the primary defect).
    assert case.status is CaseStatus.FAILED


async def test_compile_error_from_the_probe_blocks_macros(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)
    rig.compile = _CompileStub(error=ExcelVbaMcpError("probe exploded"))
    rig.runner = _rebuild_runner(rig)

    suite = _suite(
        [
            {
                "id": "gated",
                "steps": [
                    {"action": "vba_compile"},
                    {"action": "macro_run", "macro": "modTest.AddNumbers"},
                ],
                "assertions": [{"type": "macro_status"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.steps[1].details["blocked"] is True
    assert rig.macro.calls == []
    assert case.status is CaseStatus.FAILED


async def test_indeterminate_compile_does_not_block_macros(rig_tmp: tuple[Rig, Path]) -> None:
    """No compile error was proven, so macro steps run (documented behaviour)."""

    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "probe",
                "steps": [
                    {"action": "vba_compile"},
                    {"action": "macro_run", "macro": "modTest.AddNumbers", "args": [1, 2]},
                ],
                "assertions": [{"type": "compile_status"}, {"type": "macro_status"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.steps[0].details["compile_status"] == CompileStatus.INDETERMINATE.value
    assert "no compile error proven" in " ".join(case.notes)
    assert len(rig.macro.calls) == 1
    assert case.status is CaseStatus.PASSED


async def test_policy_blocked_macro_blocks_the_case(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)
    rig.macro = _MacroStub(MacroRunStatus.BLOCKED)
    rig.runner = _rebuild_runner(rig)

    suite = _suite(
        [
            {
                "id": "blocked-macro",
                "steps": [{"action": "macro_run", "macro": "modTest.RaiseError"}],
                "assertions": [{"type": "macro_status", "expected": "success"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.steps[0].status is StepStatus.SKIPPED
    assert case.steps[0].details["blocked"] is True
    assert case.status is CaseStatus.BLOCKED
    assert result.status is RunStatus.BLOCKED


# ---------------------------------------------------------------------------
# Status mapping
# ---------------------------------------------------------------------------
async def test_missing_macro_fails_the_case(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)
    rig.macro = _MacroStub(MacroRunStatus.NOT_FOUND)
    rig.runner = _rebuild_runner(rig)

    suite = _suite(
        [
            {
                "id": "missing",
                "steps": [{"action": "macro_run", "macro": "modTest.Nope"}],
                "assertions": [{"type": "macro_status", "expected": "success"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.steps[0].status is StepStatus.FAILED
    assert case.steps[0].error_code == MacroRunStatus.NOT_FOUND.value
    assert case.status is CaseStatus.FAILED


async def test_failed_assertion_fails_the_case(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "bad-expectation",
                "steps": [],
                "assertions": [
                    {"type": "read_cell", "sheet": "Sheet1", "cell": "A1", "expected": "wrong"}
                ],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.assertions[0].passed is False
    assert case.status is CaseStatus.FAILED
    assert result.status is RunStatus.FAILED


async def test_case_without_assertions_is_indeterminate(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [{"id": "no-assertions", "steps": [], "assertions": []}], baseline="use_existing"
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)

    assert result.cases[0].status is CaseStatus.INDETERMINATE
    assert result.status is RunStatus.INDETERMINATE


async def test_suite_cannot_open_its_baseline_marks_cases_blocked(
    rig_tmp: tuple[Rig, Path],
) -> None:
    """Deny-by-default: no grant for the baseline -> every case is blocked."""

    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.snapshot_service.create_snapshot(info.workbook_id)
    # Deliberately no grant for the baseline path.

    suite = _suite([{"id": "no-approval", "steps": [], "assertions": []}])
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    case = result.cases[0]

    assert case.status is CaseStatus.BLOCKED
    assert result.status is RunStatus.BLOCKED
    assert "excel_open" in " ".join(case.notes)
    # No step was dispatched.
    assert rig.sheets.calls == []


async def test_unknown_workbook_id_fails_closed(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    suite = _suite([{"id": "c", "steps": [], "assertions": []}])
    suite.workbook_id = "wb-does-not-exist"

    with pytest.raises(ExcelVbaMcpError):
        await rig.runner.run_suite(suite)


async def test_run_case_without_context_is_rejected(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    case = TestSuiteSpec.model_validate(
        {
            "suite": "s",
            "workbook_id": "w",
            "cases": [{"id": "c", "steps": [], "assertions": []}],
        }
    ).cases[0]

    with pytest.raises(ExcelVbaMcpError):
        await rig.runner.run_case(case)


# ---------------------------------------------------------------------------
# Unsupported actions fail closed
# ---------------------------------------------------------------------------
async def test_vba_source_action_fails_closed(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "source-write",
                "steps": [
                    {
                        "action": "vba_insert_code",
                        "component": "modTest",
                        "source": "Sub P()\nEnd Sub",
                        "insert_at": "declarations",
                    }
                ],
                "assertions": [],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    step = result.cases[0].steps[0]

    assert step.status is StepStatus.FAILED
    assert step.error_code == ERROR_UNSUPPORTED_ACTION
    assert "expected_source_hash" in (step.error_text or "")


async def test_shape_action_without_shapes_service_fails_closed(
    rig_tmp: tuple[Rig, Path],
) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "shape",
                "steps": [
                    {"action": "shape_set_on_action", "name": "btn", "macro": "modTest.NoArgs"}
                ],
                "assertions": [],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)
    step = result.cases[0].steps[0]

    assert step.status is StepStatus.FAILED
    assert step.error_code == ERROR_UNSUPPORTED_ACTION
    assert "shapes service" in (step.error_text or "")


# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------
async def test_run_is_persisted_with_evidence(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp
    info = await rig.open_original()
    baseline = await rig.create_baseline(info.workbook_id)

    suite = _suite(
        [
            {
                "id": "case/with:odd-name",
                "steps": [{"action": "sheet_add", "sheet": "A"}],
                "assertions": [{"type": "sheet_exists", "name": "A"}],
            }
        ],
        baseline="use_existing",
    )
    suite.workbook_id = info.workbook_id

    result = await rig.runner.run_suite(suite)

    # Run document round-trips through the repository.
    stored = rig.runs.get_run(result.test_run_id)
    assert stored.test_run_id == result.test_run_id
    assert stored.status is result.status
    assert len(stored.cases) == 1
    assert Path(rig.runs.path_for(result.test_run_id)).is_file()
    assert [summary.test_run_id for summary in rig.runs.list_runs()] == [result.test_run_id]
    assert rig.runs.list_runs()[0].baseline_snapshot_id == baseline.snapshot_id

    # One manifest + one per-case diagnostic artifact, both on disk.
    assert len(result.artifact_paths) == 2
    for path in result.artifact_paths:
        assert Path(path).is_file()
    manifest_path = next(
        path for path in result.artifact_paths if Path(path).name.endswith("manifest.json")
    )
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    assert manifest["external_side_effects_not_undone"] is True
    assert manifest["status"] == result.status.value
    # Artifacts are recorded as manifest / diagnostic types.
    kinds = {artifact.name for artifact in rig.artifacts.list_artifacts()}
    assert any("manifest" in name for name in kinds)
    assert any("case-with-odd-name" in name for name in kinds)

    # The runner's read API returns the same document.
    assert rig.runner.get_run(result.test_run_id).test_run_id == result.test_run_id


async def test_get_unknown_run_raises(rig_tmp: tuple[Rig, Path]) -> None:
    rig, _tmp = rig_tmp

    with pytest.raises(RunNotFoundError):
        rig.runner.get_run("run-does-not-exist")


def test_run_repository_rejects_traversal(tmp_path: Path) -> None:
    repository = RunRepository(tmp_path / "runs")

    for bad in ["../escape", "..\\escape", "", "   ", "run id"]:
        with pytest.raises(PathNotAllowedError):
            repository.path_for(bad)


def test_run_repository_write_is_atomic(tmp_path: Path) -> None:
    repository = RunRepository(tmp_path / "runs")
    run = _sample_run()

    target = repository.save_run(run)

    assert target.is_file()
    # No leftover temp files from the atomic write.
    assert [path.name for path in target.parent.iterdir()] == [target.name]
    assert repository.get_run(run.test_run_id).model_dump() == run.model_dump()


def test_every_result_declares_external_side_effects() -> None:
    run = _sample_run()

    assert EXTERNAL_SIDE_EFFECTS_NOTE in run.cases[0].notes
    # Assertion results are typed DTOs, never free-form payloads.
    assertion = run.cases[0].assertions[0]
    assert isinstance(assertion, AssertionResult)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _rebuild_runner(rig: Rig) -> Runner:
    """Rebuild the runner with the rig's current (possibly swapped) stubs."""

    return Runner(
        rig.worker,  # type: ignore[arg-type]
        rig.session,
        rig.snapshot_service,
        rig.compile,  # type: ignore[arg-type]
        rig.macro,  # type: ignore[arg-type]
        rig.sheets,  # type: ignore[arg-type]
        rig.tables,  # type: ignore[arg-type]
        rig.names,  # type: ignore[arg-type]
        rig.artifacts,
        rig.runs,
        rig.path_policy,
        rig.audit,
    )


def _sha256(path: str | Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _sample_run() -> Any:
    """Build a minimal persisted :class:`TestRunResult` for repository tests."""

    from excel_vba_mcp.domain.models import TestCaseResult, TestRunResult

    now = datetime.now(UTC)
    return TestRunResult(
        test_run_id="run-sample",
        suite="sample-suite",
        status=RunStatus.PASSED,
        workbook_id="wb-sample",
        baseline_snapshot_id="snap-sample",
        cases=[
            TestCaseResult(
                case_id="case-1",
                status=CaseStatus.PASSED,
                steps=[],
                assertions=[
                    AssertionResult(type="sheet_exists", passed=True, expected=True, actual=True)
                ],
                duration_ms=1,
                working_copy_path=None,
                notes=[EXTERNAL_SIDE_EFFECTS_NOTE],
                created_at=now,
                updated_at=now,
            )
        ],
        started_at=now,
        finished_at=now,
        duration_ms=1,
        artifact_paths=[str(Path("artifact.json"))],
        created_at=now,
        updated_at=now,
    )


@pytest.fixture
def rig_tmp(tmp_path: Path) -> tuple[Rig, Path]:
    instance = Rig(tmp_path)
    yield instance, tmp_path
    instance.audit.close()
