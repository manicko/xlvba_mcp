"""Unit tests for the policy-controlled macro service (task_022).

These tests never open Excel and never import COM. They drive the real
:class:`~excel_vba_mcp.application.macro_service.MacroService`,
:class:`~excel_vba_mcp.adapters.excel.macro_runner.MacroRunner` and the real
approval / background-safety / audit collaborators against a pure-Python fake
``Application`` / ``Workbook`` / ``VBProject`` object graph (plus a fake STA
worker that runs the submitted callable inline).

They prove the binding task_022 contract without Excel:

* ``list_macros`` enumerates public procedures with their ``background_safe``
  classification (deny-by-default -> ``False`` until classified);
* ``run_macro`` is deny-by-default: no approval -> ``BLOCKED``;
* the background_safe gate (BG06): unclassified -> ``BACKGROUND_SAFETY_NOT_VERIFIED``,
  classified unsafe -> ``INTERACTIVE_UI_NOT_SUPPORTED``;
* the interactive-UI preflight (R13): ``ShowMsg`` / ``ShowForm`` are blocked even
  when (wrongly) classified safe, because their body references UI tokens;
* a successful/approved/background-safe macro runs and its return value is
  serialized; a missing macro -> ``NOT_FOUND``; a runtime error -> ``RUNTIME_ERROR``;
* an unsupported argument is rejected (pre-dispatch) as ``UNSUPPORTED_ARGUMENT``;
* a dispatch timeout -> ``MACRO_TIMEOUT`` (no retry); a modal hang -> owned-PID
  terminate with ``dismissed=false`` and ``MACRO_RUNTIME_ERROR``;
* ``get_last_result`` returns the cache without re-running.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelHungError, WorkerStatus
from excel_vba_mcp.adapters.excel.macro_runner import (
    ERROR_NON_SERIALIZABLE_RETURN,
    ERROR_UNSUPPORTED_ARGUMENT,
    MacroRunner,
    coerce_args,
    serialize_return,
)
from excel_vba_mcp.application import macro_service as macro_service_module
from excel_vba_mcp.application.macro_service import (
    ERROR_BACKGROUND_SAFETY_NOT_VERIFIED,
    ERROR_INTERACTIVE_UI_NOT_SUPPORTED,
    BackgroundSafetyRegistry,
    MacroInfo,
    MacroService,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import MacroRunResult
from excel_vba_mcp.domain.statuses import (
    MacroRunStatus,
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

# Outer DISP_E_EXCEPTION wrapper (signed) carried by a pywintypes.com_error.
_DISP_E_EXCEPTION_SIGNED = 0x80020009 - 0x100000000

MOD_TEST_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Public Function NoArgs() As String
    NoArgs = "no-args"
End Function

Public Sub Echo(ByVal s As String)
    Debug.Print s
End Sub

Public Function MakeArray() As Variant
    MakeArray = Array(1, 2, 3)
End Function

Public Sub RaiseError()
    Err.Raise 1234, "modTest", "boom"
End Sub

Public Sub ShowMsg()
    MsgBox "interactive", vbInformation, "probe"
End Sub

Public Sub ShowForm()
    frmTest.Show
End Sub
"""


# ---------------------------------------------------------------------------
# Fake COM graph (no Excel / no pywin32)
# ---------------------------------------------------------------------------
class FakeComError(Exception):
    """Stand-in for ``pywintypes.com_error`` with an EXCEPINFO-style shape."""

    def __init__(self, description: str | None, inner_hresult: int) -> None:
        excepinfo = (0, "Microsoft Excel", description, "xlmain11.chm", 0, inner_hresult)
        super().__init__("Exception occurred.", "Exception occurred.", excepinfo)
        self.hresult = _DISP_E_EXCEPTION_SIGNED


def _parse_procedures(text: str) -> list[dict[str, Any]]:
    """Parse ``{name, start (1-based), count, kind_code}`` for a module text."""

    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    procs: list[dict[str, Any]] = []
    index = 0
    header_re = re.compile(
        r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
        r"([A-Za-z_][A-Za-z0-9_]*)",
        re.IGNORECASE,
    )
    end_re = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
    while index < len(lines):
        match = header_re.match(lines[index])
        if not match:
            index += 1
            continue
        kind_word = match.group(1).lower()
        kind_code = 0 if kind_word == "sub" else 1 if kind_word == "function" else 2
        start = index + 1
        index += 1
        while index < len(lines):
            if end_re.match(lines[index]):
                index += 1
                break
            index += 1
        procs.append(
            {
                "name": match.group(2),
                "start": start,
                "count": index - start + 1,
                "kind": kind_code,
            }
        )
    return procs


class FakeCodeModule:
    """A ``CodeModule`` supporting the VBIDE procedure APIs the adapter uses."""

    def __init__(self, text: str) -> None:
        self._lines = text.replace("\r\n", "\n").split("\n")
        self._procs = _parse_procedures(text)
        self._by_name = {p["name"].lower(): p for p in self._procs}

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - COM casing
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:  # noqa: N802 - COM casing
        return "\n".join(self._lines[start - 1 : start - 1 + count])

    def ProcOfLine(self, line: int, _kind: int = 0) -> tuple[str, int]:  # noqa: N802 - COM casing
        for proc in self._procs:
            if proc["start"] == line:
                return (proc["name"], proc["kind"])
        raise KeyError(line)

    def ProcStartLine(self, name: str, _kind: int = 0) -> int:  # noqa: N802 - COM casing
        return self._by_name[name.lower()]["start"]

    def ProcCountLines(self, name: str, _kind: int = 0) -> int:  # noqa: N802 - COM casing
        return self._by_name[name.lower()]["count"]


class FakeComponent:
    def __init__(self, name: str, type_code: int, code: str) -> None:
        self.Name = name
        self.Type = type_code
        self.CodeModule = FakeCodeModule(code)
        self.Designer = None

    def Export(self, path: str) -> None:  # noqa: N802 - COM casing
        Path(path).write_text(self.CodeModule.Lines(1, self.CodeModule.CountOfLines))


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
    def __init__(self, components: list[FakeComponent]) -> None:
        self.Name = "FakeProject"
        self.Protection = 0
        self.VBComponents = FakeComponents(components)
        self.References = []


class FakeWorkbooks:
    def __init__(self, project: FakeProject, app: FakeApp) -> None:
        self._project = project
        self._app = app
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **_kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(self._project, self._app)
        wb.FullName = path  # type: ignore[attr-defined]
        wb.Name = Path(path).name  # type: ignore[attr-defined]
        self.items.append(wb)
        return wb

    def __iter__(self) -> Any:
        return iter(self.items)


class FakeWorkbook:
    def __init__(self, project: FakeProject, app: FakeApp) -> None:
        self._project = project
        self.Application = app
        self.FullName = ""
        self.Name = ""
        self.ReadOnly = False

    @property
    def VBProject(self) -> FakeProject:  # noqa: N802 - COM casing
        return self._project


class FakeApp:
    """A fake ``Excel.Application`` whose ``Run`` simulates ``modTest`` macros."""

    def __init__(self, project: FakeProject) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(project, self)

    def Run(self, macro: str, *args: Any) -> Any:  # noqa: N802 - COM casing
        procedure = macro.split("!", 1)[-1]
        name = procedure.split(".")[-1]
        if name == "AddNumbers":
            return args[0] + args[1]
        if name == "NoArgs":
            return "no-args"
        if name == "MakeArray":
            return (1, 2, 3)
        if name == "Echo":
            return None
        if name == "ReturnObject":
            return object()
        if name == "RaiseError":
            raise FakeComError(None, 0x800A03EC - 0x100000000)
        raise FakeComError(
            f"Cannot run the macro '{macro}'. The macro may not be available in this "
            "workbook or all macros may be disabled.",
            0x800A03ED - 0x100000000,
        )


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, records submissions."""

    def __init__(self, *, owned_pid: int = 4242) -> None:
        self.submits: list[dict[str, Any]] = []
        self.hang_execute = False
        self._owned_pid = owned_pid

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
        if self.hang_execute and section == "execute":
            raise ExcelHungError("simulated dispatch timeout")
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(
            state=WorkbookSessionState.RUNNING, accepting=True, owned_pid=self._owned_pid
        )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _default_project() -> FakeProject:
    return FakeProject(
        [
            FakeComponent("modTest", 1, MOD_TEST_SOURCE),
            FakeComponent("ThisWorkbook", 100, "Option Explicit\n"),
        ]
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class Rig:
    def __init__(self, tmp_path: Path, *, project: FakeProject | None = None) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.background = BackgroundSafetyRegistry()
        self.app = FakeApp(project or _default_project())
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
        self.service = MacroService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.approval,
            self.background,
            self.audit,
        )
        self.workbook_hash = _sha256(self.original)

    def grant_open(self) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(
                    canonical_path=self.canonical,
                    sha256=self.workbook_hash,
                    action="excel_open",
                ),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_macro_run(self, procedure: str | None) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(
                    canonical_path=self.canonical,
                    sha256=self.workbook_hash,
                    action="macro_run",
                    procedure=procedure,
                ),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def open(
        self, *, mode: SessionMode = SessionMode.EXECUTE, read_only: bool = False
    ) -> Any:
        if mode == SessionMode.EXECUTE:
            self.grant_open()
        return await self.session.open_workbook(str(self.original), mode=mode, read_only=read_only)


# ---------------------------------------------------------------------------
# MacroRunner: argument coercion / return serialization (pure)
# ---------------------------------------------------------------------------
def test_coerce_args_accepts_supported_scalars_and_arrays() -> None:
    coerced = coerce_args([1, 2.5, True, None, "x", [1, 2], [[1], [2]]])
    assert coerced == [1, 2.5, True, None, "x", [1, 2], [[1], [2]]]


def test_coerce_args_rejects_date_and_objects() -> None:
    with pytest.raises(ExcelVbaMcpError):
        coerce_args([_dt.date(2026, 11, 19)])
    with pytest.raises(ExcelVbaMcpError):
        coerce_args([object()])


def test_coerce_args_accepts_datetime() -> None:
    value = _dt.datetime(2026, 11, 19, 12, 0, tzinfo=UTC)
    assert coerce_args([value]) == [value]


def test_serialize_return_rejects_object_and_dict() -> None:
    assert serialize_return((1, 2, 3)) == [1, 2, 3]
    assert serialize_return("x") == "x"
    with pytest.raises(ExcelVbaMcpError):
        serialize_return(object())
    with pytest.raises(ExcelVbaMcpError):
        serialize_return({"a": 1})


def test_runner_success_and_classification() -> None:
    app = FakeApp(_default_project())

    ok = MacroRunner(app).run("'Book.xlsm'!modTest.AddNumbers", [2, 3])
    assert ok.status is MacroRunStatus.SUCCESS
    assert ok.return_value == 5

    missing = MacroRunner(app).run("'Book.xlsm'!modTest.Nope")
    assert missing.status is MacroRunStatus.NOT_FOUND
    assert missing.error_code == ErrorCode.MACRO_NOT_FOUND.value

    runtime = MacroRunner(app).run("'Book.xlsm'!modTest.RaiseError")
    assert runtime.status is MacroRunStatus.RUNTIME_ERROR
    assert runtime.error_code == ErrorCode.MACRO_RUNTIME_ERROR.value


def test_runner_blocks_unsupported_argument_before_dispatch() -> None:
    app = FakeApp(_default_project())
    result = MacroRunner(app).run("'Book.xlsm'!modTest.Echo", [object()])
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_UNSUPPORTED_ARGUMENT


def test_runner_blocks_non_serializable_return() -> None:
    app = FakeApp(_default_project())
    result = MacroRunner(app).run("'Book.xlsm'!modTest.ReturnObject")
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_NON_SERIALIZABLE_RETURN


# ---------------------------------------------------------------------------
# list_macros
# ---------------------------------------------------------------------------
async def test_list_macros_enumerates_public_procedures(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    macros = await rig.service.list_macros(info.workbook_id)

    names = {macro.name for macro in macros}
    assert {"modTest.AddNumbers", "modTest.NoArgs", "modTest.MakeArray"} <= names
    add = next(m for m in macros if m.name == "modTest.AddNumbers")
    assert isinstance(add, MacroInfo)
    assert add.module == "modTest"
    assert add.procedure == "AddNumbers"
    # Deny-by-default: nothing is classified yet.
    assert add.background_safe is False
    assert add.background_safe_classified is False


async def test_list_macros_reflects_background_classification(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.background.classify(rig.workbook_hash, "modTest.AddNumbers", background_safe=True)

    macros = await rig.service.list_macros(info.workbook_id)

    add = next(m for m in macros if m.name == "modTest.AddNumbers")
    assert add.background_safe is True
    assert add.background_safe_classified is True


# ---------------------------------------------------------------------------
# run_macro: deny-by-default approval
# ---------------------------------------------------------------------------
async def test_run_macro_denied_without_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    # No macro_run grant, no background classification.
    rig.background.classify(rig.workbook_hash, "modTest.AddNumbers", background_safe=True)

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ErrorCode.MACRO_BLOCKED.value
    # Never dispatched.
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


# ---------------------------------------------------------------------------
# run_macro: background_safe gate (BG06)
# ---------------------------------------------------------------------------
async def test_run_macro_unclassified_is_blocked(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.AddNumbers")
    # Approved but NOT classified background_safe.

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_BACKGROUND_SAFETY_NOT_VERIFIED
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


async def test_run_macro_classified_unsafe_is_blocked(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.AddNumbers")
    rig.background.classify(rig.workbook_hash, "modTest.AddNumbers", background_safe=False)

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_INTERACTIVE_UI_NOT_SUPPORTED
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


# ---------------------------------------------------------------------------
# run_macro: interactive-UI preflight (R13) blocks ShowMsg / ShowForm
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("macro", ["modTest.ShowMsg", "modTest.ShowForm"])
async def test_run_macro_ui_preflight_blocks(tmp_path: Path, macro: str) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    # Even if wrongly classified background_safe, the preflight must still block.
    rig.grant_macro_run(macro)
    rig.background.classify(rig.workbook_hash, macro, background_safe=True)

    result = await rig.service.run_macro(info.workbook_id, macro)

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_INTERACTIVE_UI_NOT_SUPPORTED
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


# ---------------------------------------------------------------------------
# run_macro: success / not-found / runtime-error
# ---------------------------------------------------------------------------
async def test_run_macro_success_returns_serialized_value(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.AddNumbers")
    rig.background.classify(rig.workbook_hash, "modTest.AddNumbers", background_safe=True)

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == 5
    execute = [s for s in rig.worker.submits if s["section"] == "execute"]
    assert len(execute) == 1
    assert execute[0]["op_class"] is OperationClass.EXECUTING


async def test_run_macro_missing_is_not_found(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    result = await rig.service.run_macro(info.workbook_id, "modTest.Nope")

    assert result.status is MacroRunStatus.NOT_FOUND
    assert result.error_code == ErrorCode.MACRO_NOT_FOUND.value
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


async def test_run_macro_runtime_error(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.RaiseError")
    rig.background.classify(rig.workbook_hash, "modTest.RaiseError", background_safe=True)

    result = await rig.service.run_macro(info.workbook_id, "modTest.RaiseError")

    assert result.status is MacroRunStatus.RUNTIME_ERROR
    assert result.error_code == ErrorCode.MACRO_RUNTIME_ERROR.value


async def test_run_macro_unsupported_argument_is_blocked(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.Echo")
    rig.background.classify(rig.workbook_hash, "modTest.Echo", background_safe=True)

    result = await rig.service.run_macro(info.workbook_id, "modTest.Echo", [object()])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_UNSUPPORTED_ARGUMENT


# ---------------------------------------------------------------------------
# run_macro: timeout (no retry) and modal-hang (owned-PID terminate)
# ---------------------------------------------------------------------------
async def test_run_macro_timeout_no_retry(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.Echo")
    rig.background.classify(rig.workbook_hash, "modTest.Echo", background_safe=True)
    rig.worker.hang_execute = True

    terminated: list[int] = []
    monkeypatch.setattr(macro_service_module, "_owned_modal_present", lambda _pid: False)
    monkeypatch.setattr(
        macro_service_module, "_terminate_owned_pid", lambda pid: terminated.append(pid) or True
    )

    result = await rig.service.run_macro(info.workbook_id, "modTest.Echo")

    assert result.status is MacroRunStatus.TIMEOUT
    assert result.error_code == ErrorCode.MACRO_TIMEOUT.value
    # Exactly one execute submission; no blind retry.
    assert len([s for s in rig.worker.submits if s["section"] == "execute"]) == 1
    # No modal -> the owned PID was not terminated.
    assert terminated == []


async def test_run_macro_modal_hang_terminates_owned_pid(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.RaiseError")
    rig.background.classify(rig.workbook_hash, "modTest.RaiseError", background_safe=True)
    rig.worker.hang_execute = True

    terminated: list[int] = []
    monkeypatch.setattr(macro_service_module, "_owned_modal_present", lambda _pid: True)
    monkeypatch.setattr(
        macro_service_module, "_terminate_owned_pid", lambda pid: terminated.append(pid) or True
    )

    result = await rig.service.run_macro(info.workbook_id, "modTest.RaiseError")

    assert result.status is MacroRunStatus.RUNTIME_ERROR
    assert result.error_code == ErrorCode.MACRO_RUNTIME_ERROR.value
    assert "NOT dismissed" in (result.error_text or "")
    # ONLY the owned PID was terminated (no wildcard, no dialog dismissal).
    assert terminated == [4242]
    assert len([s for s in rig.worker.submits if s["section"] == "execute"]) == 1


# ---------------------------------------------------------------------------
# Execute-mode requirement
# ---------------------------------------------------------------------------
async def test_run_macro_requires_execute_mode(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    # Open in inspect (macros disabled) mode.
    info = await rig.open(mode=SessionMode.INSPECT, read_only=True)

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    assert excinfo.value.code is ErrorCode.MACRO_BLOCKED


# ---------------------------------------------------------------------------
# get_last_result (cache, no re-run)
# ---------------------------------------------------------------------------
async def test_get_last_result_uses_cache_without_rerun(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.AddNumbers")
    rig.background.classify(rig.workbook_hash, "modTest.AddNumbers", background_safe=True)

    first = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    execute_count = len([s for s in rig.worker.submits if s["section"] == "execute"])

    cached = await rig.service.get_last_result(info.workbook_id)

    assert isinstance(cached, MacroRunResult)
    assert cached.run_id == first.run_id
    assert len([s for s in rig.worker.submits if s["section"] == "execute"]) == execute_count


async def test_get_last_result_none_before_any_run(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    assert await rig.service.get_last_result(info.workbook_id) is None


# ---------------------------------------------------------------------------
# Unknown workbook
# ---------------------------------------------------------------------------
async def test_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.run_macro("ghost", "modTest.AddNumbers")
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.list_macros("ghost")
