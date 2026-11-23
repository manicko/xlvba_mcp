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
import json
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
    _UNREADABLE_MODULE_SENTINEL,
    ERROR_BACKGROUND_SAFETY_NOT_VERIFIED,
    ERROR_INTERACTIVE_UI_NOT_SUPPORTED,
    BackgroundSafetyClassification,
    BackgroundSafetyRegistry,
    MacroInfo,
    MacroService,
    _compute_project_source_hash,
    _find_interactive_tokens,
    _hash_module_lines,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, StaleTargetError
from excel_vba_mcp.domain.models import MacroRunResult
from excel_vba_mcp.domain.statuses import (
    AccessMode,
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
    """A ``CodeModule`` supporting the VBIDE procedure APIs the adapter uses.

    When ``heuristic=True`` the procedure-enumeration APIs raise, so ``VBIDEAdapter``
    falls back to the textual scan and flags every procedure ``is_heuristic=True``
    (used to prove the run/classify fail-closed behavior on heuristic procedures).
    """

    def __init__(self, text: str, *, heuristic: bool = False) -> None:
        self._lines = text.replace("\r\n", "\n").split("\n")
        self._procs = _parse_procedures(text)
        self._by_name = {p["name"].lower(): p for p in self._procs}
        self._heuristic = heuristic

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - COM casing
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:  # noqa: N802 - COM casing
        return "\n".join(self._lines[start - 1 : start - 1 + count])

    def ProcOfLine(self, line: int, _kind: int = 0) -> tuple[str, int]:  # noqa: N802 - COM casing
        if self._heuristic:
            raise KeyError(line)
        for proc in self._procs:
            if proc["start"] == line:
                return (proc["name"], proc["kind"])
        raise KeyError(line)

    def ProcStartLine(self, name: str, _kind: int = 0) -> int:  # noqa: N802 - COM casing
        if self._heuristic:
            raise KeyError(name)
        return self._by_name[name.lower()]["start"]

    def ProcCountLines(self, name: str, _kind: int = 0) -> int:  # noqa: N802 - COM casing
        if self._heuristic:
            raise KeyError(name)
        return self._by_name[name.lower()]["count"]


class FakeComponent:
    def __init__(self, name: str, type_code: int, code: str, *, heuristic: bool = False) -> None:
        self.Name = name
        self.Type = type_code
        self.CodeModule = FakeCodeModule(code, heuristic=heuristic)
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
    """A fake ``Excel.Application`` whose ``Run`` simulates ``modTest`` macros.

    ``Interactive`` / ``DisplayAlerts`` are properties that record every write in
    :attr:`writes` (so the F2a dialog-suppression guard can be asserted) and
    snapshot both values at the start of each :meth:`Run` (so "were they ``False``
    during the dispatch?" can be checked). Both start ``True`` (their Excel
    default) so a guard that failed to set/restore them is observable.
    """

    def __init__(self, project: FakeProject) -> None:
        self.Visible = False
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(project, self)
        # Guard-tracked properties (F2a). Private backing fields keep the init
        # writes out of the log; writes made after construction are recorded.
        self._interactive = True
        self._display_alerts = True
        self.writes: list[tuple[str, bool]] = []
        self.run_snapshots: list[dict[str, bool]] = []

    @property
    def Interactive(self) -> bool:  # noqa: N802 - COM casing
        return self._interactive

    @Interactive.setter
    def Interactive(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._interactive = value
        self.writes.append(("Interactive", value))

    @property
    def DisplayAlerts(self) -> bool:  # noqa: N802 - COM casing
        return self._display_alerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._display_alerts = value
        self.writes.append(("DisplayAlerts", value))

    def Run(self, macro: str, *args: Any) -> Any:  # noqa: N802 - COM casing
        # Snapshot both guard properties at dispatch time: the guard must have set
        # them False for the duration of the macro.
        self.run_snapshots.append(
            {"Interactive": self._interactive, "DisplayAlerts": self._display_alerts}
        )
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


class GuardApp:
    """Configurable fake ``Excel.Application`` for the F2a guard failure tests.

    Each flag injects a failure at a specific point of the guard's defensive
    read/set/restore cycle so the guard's skip / swallow behavior can be asserted:

    * ``*_read_raises`` -> ``getattr`` raises (property skipped, not restored);
    * ``*_set_false_raises`` -> the guard's ``setattr(False)`` raises (skipped);
    * ``*_restore_raises`` -> the guard's restore write (a non-``False`` value)
      raises (swallowed; prior value left as-is).
    """

    def __init__(
        self,
        *,
        interactive_read_raises: bool = False,
        alerts_read_raises: bool = False,
        interactive_set_false_raises: bool = False,
        alerts_set_false_raises: bool = False,
        interactive_restore_raises: bool = False,
        alerts_restore_raises: bool = False,
        run_result: Any = "ok",
    ) -> None:
        self._interactive = True
        self._display_alerts = True
        self._interactive_read_raises = interactive_read_raises
        self._alerts_read_raises = alerts_read_raises
        self._interactive_set_false_raises = interactive_set_false_raises
        self._alerts_set_false_raises = alerts_set_false_raises
        self._interactive_restore_raises = interactive_restore_raises
        self._alerts_restore_raises = alerts_restore_raises
        self._run_result = run_result
        self.writes: list[tuple[str, bool]] = []
        self.run_count = 0

    @property
    def Interactive(self) -> bool:  # noqa: N802 - COM casing
        if self._interactive_read_raises:
            raise RuntimeError("Interactive read failed")
        return self._interactive

    @Interactive.setter
    def Interactive(self, value: bool) -> None:  # noqa: N802 - COM casing
        if value is False and self._interactive_set_false_raises:
            raise RuntimeError("Interactive set(False) failed")
        if value is not False and self._interactive_restore_raises:
            raise RuntimeError("Interactive restore failed")
        self._interactive = value
        self.writes.append(("Interactive", value))

    @property
    def DisplayAlerts(self) -> bool:  # noqa: N802 - COM casing
        if self._alerts_read_raises:
            raise RuntimeError("DisplayAlerts read failed")
        return self._display_alerts

    @DisplayAlerts.setter
    def DisplayAlerts(self, value: bool) -> None:  # noqa: N802 - COM casing
        if value is False and self._alerts_set_false_raises:
            raise RuntimeError("DisplayAlerts set(False) failed")
        if value is not False and self._alerts_restore_raises:
            raise RuntimeError("DisplayAlerts restore failed")
        self._display_alerts = value
        self.writes.append(("DisplayAlerts", value))

    def Run(self, macro: str, *args: Any) -> Any:  # noqa: N802 - COM casing
        self.run_count += 1
        return self._run_result


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
    def __init__(
        self,
        tmp_path: Path,
        *,
        project: FakeProject | None = None,
        mode: AccessMode = AccessMode.RESTRICTED,
    ) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy(mode=mode)
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

    def _record(self, workbook_id: str) -> Any:
        for info in self.session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise AssertionError(f"unknown workbook_id {workbook_id!r}")

    async def compute_project_source_hash(self, workbook_id: str) -> str:
        """Compute the live ``project_source_hash`` via the service's single scan."""

        record = self._record(workbook_id)
        _inventory, project_source_hash = await self.service._collect_project_state(
            workbook_id, record
        )
        return project_source_hash

    async def classify_safe(
        self, workbook_id: str, procedure: str, *, background_safe: bool = True
    ) -> None:
        """Classify via the raw registry, bound to the live hash + source revision.

        Bypasses :meth:`MacroService.classify_background` on purpose: the R13
        preflight tests must be able to (wrongly) mark an interactive procedure
        ``background_safe`` to prove the deterministic veto still blocks it.
        """

        record = self._record(workbook_id)
        workbook_hash = _sha256(Path(record.path))
        project_source_hash = await self.compute_project_source_hash(workbook_id)
        self.background.classify(
            workbook_hash, project_source_hash, procedure, background_safe=background_safe
        )

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
# MacroRunner: F2a dialog-suppression guard (Interactive / DisplayAlerts)
# ---------------------------------------------------------------------------
def test_guard_sets_false_during_run_and_restores_in_reverse() -> None:
    app = FakeApp(_default_project())

    result = MacroRunner(app).run("'Book.xlsm'!modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == 5
    # During the dispatch both properties were False.
    assert app.run_snapshots == [{"Interactive": False, "DisplayAlerts": False}]
    # After the run both are restored to their original True.
    assert app.Interactive is True
    assert app.DisplayAlerts is True
    # Set order Interactive -> DisplayAlerts; restore order DisplayAlerts -> Interactive.
    assert app.writes == [
        ("Interactive", False),
        ("DisplayAlerts", False),
        ("DisplayAlerts", True),
        ("Interactive", True),
    ]


def test_guard_restores_when_run_raises() -> None:
    app = FakeApp(_default_project())

    result = MacroRunner(app).run("'Book.xlsm'!modTest.RaiseError")

    assert result.status is MacroRunStatus.RUNTIME_ERROR
    # Both properties restored despite the dispatch raising (restore never raises).
    assert app.Interactive is True
    assert app.DisplayAlerts is True
    assert app.writes == [
        ("Interactive", False),
        ("DisplayAlerts", False),
        ("DisplayAlerts", True),
        ("Interactive", True),
    ]


def test_guard_skips_property_whose_set_fails() -> None:
    # Interactive setter raises on set(False): it is skipped and NOT restored, the
    # dispatch still runs, and DisplayAlerts is still set + restored.
    app = GuardApp(interactive_set_false_raises=True)

    result = MacroRunner(app).run("'Book.xlsm'!modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == "ok"
    assert app.run_count == 1
    # Interactive was never written (skipped); its value is untouched (still True).
    assert app.Interactive is True
    assert ("Interactive", False) not in app.writes
    assert ("Interactive", True) not in app.writes
    # DisplayAlerts was set False and restored to True.
    assert app.DisplayAlerts is True
    assert app.writes == [("DisplayAlerts", False), ("DisplayAlerts", True)]


def test_guard_skips_property_whose_read_fails() -> None:
    # Interactive getter raises: it is skipped entirely; DisplayAlerts is still
    # set + restored; the dispatch runs.
    app = GuardApp(interactive_read_raises=True)

    result = MacroRunner(app).run("'Book.xlsm'!modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == "ok"
    assert app.run_count == 1
    assert app.DisplayAlerts is True
    assert app.writes == [("DisplayAlerts", False), ("DisplayAlerts", True)]


def test_guard_swallows_restore_failure_and_returns_result() -> None:
    # Restore order is DisplayAlerts then Interactive (reverse of the applied set).
    # Interactive (the second restore) raises: it is swallowed and the original
    # result is returned unchanged; DisplayAlerts is restored first.
    app = GuardApp(interactive_restore_raises=True)

    result = MacroRunner(app).run("'Book.xlsm'!modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == "ok"
    # DisplayAlerts restore succeeded; Interactive restore failed (left False).
    assert app.DisplayAlerts is True
    assert app.Interactive is False


def test_guard_not_entered_on_coercion_early_return() -> None:
    # An unsupported argument is rejected before the guard: ZERO property writes.
    app = FakeApp(_default_project())

    result = MacroRunner(app).run("'Book.xlsm'!modTest.Echo", [object()])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_UNSUPPORTED_ARGUMENT
    assert app.writes == []
    assert app.run_snapshots == []


def test_guard_sets_each_property_false_exactly_once() -> None:
    app = FakeApp(_default_project())

    MacroRunner(app).run("'Book.xlsm'!modTest.NoArgs")

    assert app.writes.count(("Interactive", False)) == 1
    assert app.writes.count(("DisplayAlerts", False)) == 1


def test_guard_restores_before_non_serializable_return_rejection() -> None:
    app = FakeApp(_default_project())

    result = MacroRunner(app).run("'Book.xlsm'!modTest.ReturnObject")

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_NON_SERIALIZABLE_RETURN
    # The guard ran and restored both properties before serialize_return rejected.
    assert app.run_snapshots == [{"Interactive": False, "DisplayAlerts": False}]
    assert app.Interactive is True
    assert app.DisplayAlerts is True
    assert app.writes == [
        ("Interactive", False),
        ("DisplayAlerts", False),
        ("DisplayAlerts", True),
        ("Interactive", True),
    ]


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
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

    macros = await rig.service.list_macros(info.workbook_id)

    add = next(m for m in macros if m.name == "modTest.AddNumbers")
    assert add.background_safe is True
    assert add.background_safe_classified is True


# ---------------------------------------------------------------------------
# run_macro: deny-by-default approval (restricted mode)
# ---------------------------------------------------------------------------
async def test_run_macro_denied_without_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)  # default RESTRICTED
    info = await rig.open()
    # No macro_run grant, but classified background_safe so the approval gate is
    # what denies the run.
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ErrorCode.APPROVAL_DENIED.value
    assert result.error_text is not None and "approval policy" in result.error_text
    # Never dispatched.
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []


# ---------------------------------------------------------------------------
# run_macro: full-access mode allows EXECUTING with no per-run grant
# ---------------------------------------------------------------------------
async def test_run_macro_full_mode_runs_without_grant(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()  # FULL mode: open allowed without a grant
    # Classified background_safe, but NO macro_run grant.
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == 5
    execute = [s for s in rig.worker.submits if s["section"] == "execute"]
    assert len(execute) == 1
    assert execute[0]["op_class"] is OperationClass.EXECUTING


# ---------------------------------------------------------------------------
# run_macro: restricted mode runs with a matching grant
# ---------------------------------------------------------------------------
async def test_run_macro_restricted_with_grant_runs(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.RESTRICTED)
    info = await rig.open()
    rig.grant_macro_run("modTest.AddNumbers")
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])

    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == 5


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
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers", background_safe=False)

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
    await rig.classify_safe(info.workbook_id, macro)

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
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

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
    await rig.classify_safe(info.workbook_id, "modTest.RaiseError")

    result = await rig.service.run_macro(info.workbook_id, "modTest.RaiseError")

    assert result.status is MacroRunStatus.RUNTIME_ERROR
    assert result.error_code == ErrorCode.MACRO_RUNTIME_ERROR.value


async def test_run_macro_unsupported_argument_is_blocked(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant_macro_run("modTest.Echo")
    await rig.classify_safe(info.workbook_id, "modTest.Echo")

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
    await rig.classify_safe(info.workbook_id, "modTest.Echo")
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
    await rig.classify_safe(info.workbook_id, "modTest.RaiseError")
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
    await rig.classify_safe(info.workbook_id, "modTest.AddNumbers")

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


# ---------------------------------------------------------------------------
# ErrorCode surface (additive Phase-2 result codes) + literal aliases
# ---------------------------------------------------------------------------
def test_background_error_codes_exist() -> None:
    assert ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED.value == "BACKGROUND_SAFETY_NOT_VERIFIED"
    assert ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED.value == "INTERACTIVE_UI_NOT_SUPPORTED"
    # The macro_service aliases are the enum's single source of truth.
    assert ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED.value == ERROR_BACKGROUND_SAFETY_NOT_VERIFIED
    assert ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED.value == ERROR_INTERACTIVE_UI_NOT_SUPPORTED


# ---------------------------------------------------------------------------
# _find_interactive_tokens: case-insensitive + expanded token set
# ---------------------------------------------------------------------------
def test_find_interactive_tokens_is_case_insensitive() -> None:
    # Lowercased tokens (VBA is case-insensitive) must still be caught.
    assert "MsgBox" in _find_interactive_tokens('x = msgbox("hi")')
    assert ".Show" in _find_interactive_tokens("frmTest.show")
    assert "Shell" in _find_interactive_tokens("shell \"calc\"")
    assert _find_interactive_tokens("Public Sub AddNumbers()\nEnd Sub\n") == []


def test_find_interactive_tokens_covers_expanded_set() -> None:
    for token in ("SendKeys", "DoEvents", "CallByName", "ExecuteGlobal", "Eval", "OnTime"):
        assert token in _find_interactive_tokens(f"Before\n    {token} something\nAfter")


def test_find_interactive_tokens_covers_window_visibility_set() -> None:
    # Focus / visibility tokens (FE-04, D-3) that a MsgBox/.Show-only scan misses.
    assert "AppActivate" in _find_interactive_tokens("AppActivate \"Book1\"")
    assert "Application.Visible" in _find_interactive_tokens("Application.Visible = True")
    assert ".Activate" in _find_interactive_tokens("Sheet1.Activate")
    assert "WindowState" in _find_interactive_tokens("ActiveWindow.WindowState = xlMaximized")


def test_find_interactive_tokens_new_tokens_case_insensitive() -> None:
    # VBA is case-insensitive: a lowercased focus/visibility token is still caught.
    assert "AppActivate" in _find_interactive_tokens('x = appactivate "w"')
    assert "Application.Visible" in _find_interactive_tokens("application.visible = true")
    assert ".Activate" in _find_interactive_tokens("sheet1.activate")
    assert "WindowState" in _find_interactive_tokens("activewindow.windowstate = xlminimized")


def test_find_interactive_tokens_clean_deterministic_macro_not_flagged() -> None:
    # The Recalculate-style body (reads a table cell, writes a result) must not be
    # flagged by any token, including the four new focus/visibility tokens.
    source = (
        "Option Explicit\n"
        "\n"
        "Public Sub Recalculate()\n"
        "    Dim ws As Worksheet, tbl As ListObject, out As Worksheet\n"
        "    Dim factor As Double\n"
        '    Set ws = ThisWorkbook.Worksheets("Sheet1")\n'
        '    Set tbl = ws.ListObjects("tblSettings")\n'
        "    factor = CDbl(tbl.DataBodyRange.Cells(1, 2).Value)\n"
        '    Set out = ThisWorkbook.Worksheets("Result")\n'
        '    out.Range("A1").Value = factor * 2\n'
        '    out.Range("A2").Value = "factor=" & CStr(factor)\n'
        "End Sub\n"
    )
    assert _find_interactive_tokens(source) == []


# ---------------------------------------------------------------------------
# project_source_hash: empty vs unreadable sentinel + determinism
# ---------------------------------------------------------------------------
def test_hash_module_lines_empty_vs_unreadable_sentinel() -> None:
    empty_hash = _hash_module_lines([])
    unreadable_hash = _hash_module_lines(None)
    # Empty (readable, zero lines) is a real hash; unreadable is a distinct sentinel.
    assert empty_hash != _UNREADABLE_MODULE_SENTINEL
    assert unreadable_hash == _UNREADABLE_MODULE_SENTINEL
    assert empty_hash != unreadable_hash
    # Deterministic for the same input.
    assert _hash_module_lines(["a", "b"]) == _hash_module_lines(["a", "b"])


def test_compute_project_source_hash_is_order_independent() -> None:
    first = _compute_project_source_hash({"modA": "h1", "modB": "h2", "ThisWorkbook": "h3"})
    second = _compute_project_source_hash({"ThisWorkbook": "h3", "modB": "h2", "modA": "h1"})
    assert first == second
    # A different per-component hash changes the digest.
    changed = _compute_project_source_hash({"modA": "h1", "modB": "CHANGED", "ThisWorkbook": "h3"})
    assert changed != first
    # The unreadable sentinel participates (invalidates), never the empty hash.
    unreadable = _compute_project_source_hash({"modA": _UNREADABLE_MODULE_SENTINEL})
    empty = _compute_project_source_hash({"modA": _hash_module_lines([])})
    assert unreadable != empty


# ---------------------------------------------------------------------------
# BackgroundSafetyRegistry: 3-tuple key lifecycle
# ---------------------------------------------------------------------------
def test_registry_classify_and_invalidate() -> None:
    registry = BackgroundSafetyRegistry()
    registry.classify("wbhash", "srchash", "modTest.AddNumbers", background_safe=True)
    assert registry.is_safe("wbhash", "srchash", "modTest.AddNumbers") is True
    # Case-insensitive procedure key.
    assert registry.is_safe("wbhash", "srchash", "  modtest.addnumbers ") is True
    # Source revision change -> unclassified (None).
    assert registry.is_safe("wbhash", "OTHER_SRCHASH", "modTest.AddNumbers") is None
    # Workbook hash change -> unclassified (None).
    assert registry.is_safe("OTHER_WBHASH", "srchash", "modTest.AddNumbers") is None
    # Explicit unsafe classification.
    registry.classify("wbhash", "srchash", "modTest.ShowMsg", background_safe=False)
    assert registry.is_safe("wbhash", "srchash", "modTest.ShowMsg") is False
    registry.clear()
    assert registry.is_safe("wbhash", "srchash", "modTest.AddNumbers") is None


# ---------------------------------------------------------------------------
# classify_background service method
# ---------------------------------------------------------------------------
def _audit_events(rig: Rig) -> list[dict[str, Any]]:
    path = rig.tmp_path / "audit.jsonl"
    if not path.is_file():
        return []
    events: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        events.append(json.loads(line))
    return events


async def test_classify_background_writes_and_is_idempotent(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    dto = await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
    assert isinstance(dto, BackgroundSafetyClassification)
    assert dto.background_safe is True
    assert dto.procedure == "modTest.AddNumbers"
    assert dto.project_source_hash == await rig.compute_project_source_hash(info.workbook_id)
    assert dto.workbook_sha256

    # Idempotent second write; the run now succeeds (no grant needed in FULL mode).
    again = await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
    assert again.project_source_hash == dto.project_source_hash

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    assert result.status is MacroRunStatus.SUCCESS
    assert result.return_value == 5

    # Audit recorded a config/macro_classify_background event.
    events = [e for e in _audit_events(rig) if e.get("tool") == "macro_classify_background"]
    assert events and events[-1]["operation_class"] == "config"


async def test_classify_background_unknown_is_macro_not_found(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await rig.service.classify_background(info.workbook_id, "modTest.Nope", True)
    assert excinfo.value.code is ErrorCode.MACRO_NOT_FOUND


async def test_classify_background_stale_source_hash_no_write(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    with pytest.raises(StaleTargetError) as excinfo:
        await rig.service.classify_background(
            info.workbook_id, "modTest.AddNumbers", True,
            expected_project_source_hash="deadbeef",
        )
    assert excinfo.value.code is ErrorCode.STALE_TARGET
    # No write occurred: a later run is still BACKGROUND_SAFETY_NOT_VERIFIED.
    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_BACKGROUND_SAFETY_NOT_VERIFIED


async def test_classify_background_interactive_token_no_write(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await rig.service.classify_background(info.workbook_id, "modTest.ShowMsg", True)
    assert excinfo.value.code is ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED

    # No write: ShowMsg remains unclassified.
    live_hash = await rig.compute_project_source_hash(info.workbook_id)
    workbook_hash = _sha256(Path(rig._record(info.workbook_id).path))
    assert rig.background.is_safe(workbook_hash, live_hash, "modTest.ShowMsg") is None


async def test_classify_background_heuristic_fail_closed(tmp_path: Path) -> None:
    project = FakeProject(
        [
            FakeComponent(
                "modH",
                1,
                "Option Explicit\n\nPublic Sub RunIt()\n    x = 1\nEnd Sub\n",
                heuristic=True,
            ),
            FakeComponent("ThisWorkbook", 100, "Option Explicit\n"),
        ]
    )
    rig = Rig(tmp_path, mode=AccessMode.FULL, project=project)
    info = await rig.open()

    # A heuristic procedure cannot be classified background_safe.
    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await rig.service.classify_background(info.workbook_id, "modH.RunIt", True)
    assert excinfo.value.code is ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED

    # ... and even if the registry is (wrongly) set safe, run_macro vetoes it.
    await rig.classify_safe(info.workbook_id, "modH.RunIt")
    result = await rig.service.run_macro(info.workbook_id, "modH.RunIt")
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_INTERACTIVE_UI_NOT_SUPPORTED


# ---------------------------------------------------------------------------
# list_macros reflects classification + project_source_hash annotation
# ---------------------------------------------------------------------------
async def test_list_macros_annotates_project_source_hash(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()
    live_hash = await rig.compute_project_source_hash(info.workbook_id)

    await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)

    macros = await rig.service.list_macros(info.workbook_id)
    add = next(m for m in macros if m.name == "modTest.AddNumbers")
    assert add.background_safe is True
    assert add.background_safe_classified is True
    assert add.project_source_hash == live_hash
    # An unclassified sibling is still deny-by-default but carries the same hash.
    other = next(m for m in macros if m.name == "modTest.NoArgs")
    assert other.background_safe is False
    assert other.project_source_hash == live_hash


# ---------------------------------------------------------------------------
# single-source hash consistency + source-edit invalidation at run time
# ---------------------------------------------------------------------------
async def test_run_after_classify_same_source_succeeds(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
    # No edits between classify and run: the live project_source_hash is identical.
    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    assert result.status is MacroRunStatus.SUCCESS


def _mutate_module(rig: Rig, name: str, new_source: str) -> None:
    components = rig.app.Workbooks._project.VBComponents  # noqa: SLF001 - test reach-in
    for component in components:
        if component.Name == name:
            component.CodeModule = FakeCodeModule(new_source)
            return
    raise AssertionError(f"component {name!r} not found")


async def test_unsaved_source_edit_invalidates_classification(tmp_path: Path) -> None:
    rig = Rig(tmp_path, mode=AccessMode.FULL)
    info = await rig.open()

    await rig.service.classify_background(info.workbook_id, "modTest.AddNumbers", True)
    # Mutate the in-memory VBA source WITHOUT saving: the file hash is unchanged,
    # but project_source_hash changes -> the classification no longer matches.
    _mutate_module(rig, "modTest", MOD_TEST_SOURCE + "\n' edited in memory\n")

    result = await rig.service.run_macro(info.workbook_id, "modTest.AddNumbers", [2, 3])
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_BACKGROUND_SAFETY_NOT_VERIFIED


# ---------------------------------------------------------------------------
# R13 preflight still blocks a lowercased interactive token even if wrongly safe
# ---------------------------------------------------------------------------
async def test_run_macro_lowercase_msgbox_blocked(tmp_path: Path) -> None:
    source = "Option Explicit\n\nPublic Sub Sneaky()\n    msgbox \"hi\"\nEnd Sub\n"
    project = FakeProject(
        [
            FakeComponent("modLow", 1, source),
            FakeComponent("ThisWorkbook", 100, "Option Explicit\n"),
        ]
    )
    rig = Rig(tmp_path, mode=AccessMode.FULL, project=project)
    info = await rig.open()

    # Wrongly classify safe (raw registry), yet R13 must still veto the lowercase token.
    await rig.classify_safe(info.workbook_id, "modLow.Sneaky")
    result = await rig.service.run_macro(info.workbook_id, "modLow.Sneaky")
    assert result.status is MacroRunStatus.BLOCKED
    assert result.error_code == ERROR_INTERACTIVE_UI_NOT_SUPPORTED
    assert [s for s in rig.worker.submits if s["section"] == "execute"] == []
