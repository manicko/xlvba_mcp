"""Bounded declarative workbook test runner (task_035; spec §7.5, §8.5, §8.6, A09).

:class:`TestRunner` executes a :class:`~excel_vba_mcp.domain.test_spec.TestSuiteSpec`
— a *closed*, declarative description of a test scenario — against disposable
copies of a workbook. It is the only place where the declarative vocabulary is
turned into service calls, and it is deliberately narrow:

* **Closed vocabulary.** Steps and assertions are validated by
  :mod:`excel_vba_mcp.domain.test_spec` **before** any Excel call; an unknown
  action or assertion kind never reaches the worker. The runner itself never
  evaluates Python, a shell command, a COM expression or VBA text from a suite.
* **Snapshot isolation.** :meth:`run_suite` creates one immutable baseline
  snapshot, then :meth:`run_case` opens a **fresh disposable working copy** of
  that baseline per case and closes it without saving. A case therefore cannot
  inherit another case's mutations, and the baseline file is never modified.
* **Delegation, not reimplementation.** Every step is dispatched to the existing
  application service (compile / macro / sheets / tables / names / shapes); the
  runner owns orchestration and status mapping only.
* **Compile gates macros.** A failed ``vba_compile`` step blocks every later
  ``macro_run`` step of the same case (``skipped`` + ``details["blocked"]``; the
  macro is **not** dispatched).
* **Honest statuses.** ``passed`` / ``failed`` / ``blocked`` / ``skipped`` /
  ``indeterminate`` are reused from :mod:`excel_vba_mcp.domain.statuses`. A case
  whose compile outcome is ``indeterminate`` (the non-activating build) is **not**
  treated as a compile failure: no compile error was proven, so macro steps run
  and the case records a note.
* **Evidence.** The run manifest, one diagnostic document per case and the run
  document are persisted through
  :class:`~excel_vba_mcp.storage.artifact_store.ArtifactStore` and
  :class:`~excel_vba_mcp.storage.run_repository.RunRepository`.

External side effects (database writes, email, network calls, files touched on
disk, other workbooks, printing) are **not** undone by a snapshot: a snapshot
restores the workbook *file* only and is not a sandbox (spec §9.4). Every run
result therefore states ``external_side_effects_not_undone``.

Safety (addendum BG03/BG04/BG05): the runner only ever drives the automation-owned
hidden Excel instance through the single-STA worker, never activates or selects
anything, writes only under the configured work / snapshot / artifact roots and
never enumerates or terminates a process it does not own.
"""

from __future__ import annotations

import contextlib
import dataclasses
import hashlib
import json
import math
import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.compile_service import CompileResult, CompileService
from excel_vba_mcp.application.macro_service import MacroService
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    AssertionResult,
    MacroRunResult,
    TestCaseResult,
    TestRunResult,
    TestStepResult,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import (
    BaselinePolicy,
    CaseStatus,
    CompileStatus,
    MacroRunStatus,
    NameScope,
    OperationClass,
    SessionMode,
    StepStatus,
    TestRunStatus,
)
from excel_vba_mcp.domain.test_spec import (
    Assertion,
    AssertionType,
    TestAction,
    TestCaseSpec,
    TestStep,
    TestSuiteSpec,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.artifact_store import ArtifactStore, ArtifactType
from excel_vba_mcp.storage.run_repository import RunRepository
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest

__all__ = [
    "ERROR_CASE_BLOCKED",
    "ERROR_COMPILE_FAILED",
    "ERROR_INVALID_STEP",
    "ERROR_MACROS_BLOCKED_BY_COMPILE",
    "ERROR_UNSUPPORTED_ACTION",
    "EXTERNAL_SIDE_EFFECTS_NOTE",
    "CaseRun",
    "RunContext",
    "TestRunner",
]

# Default per-call soft ceilings (seconds).
_CALL_TIMEOUT_S = 60.0
_MACRO_TIMEOUT_S = 120.0

# Structured step error codes (kept alongside ``ErrorCode`` values in
# ``TestStepResult.error_code`` so a caller can tell a policy/compile block apart
# from a malformed step or an unsupported action).
ERROR_INVALID_STEP = "INVALID_STEP"
ERROR_UNSUPPORTED_ACTION = "UNSUPPORTED_ACTION"
ERROR_COMPILE_FAILED = "COMPILE_FAILED"
ERROR_MACROS_BLOCKED_BY_COMPILE = "MACROS_BLOCKED_BY_COMPILE"
ERROR_CASE_BLOCKED = "CASE_BLOCKED"

#: Stated on every run/case result: a snapshot reverts the workbook file only.
EXTERNAL_SIDE_EFFECTS_NOTE = (
    "A snapshot restores the workbook FILE only. External macro side effects "
    "(database writes, email/network calls, files touched on disk, other "
    "workbooks, printing) are NOT rolled back; a snapshot is not a sandbox "
    "(spec §9.4)."
)

#: Actions of the closed vocabulary that this runner cannot dispatch, mapped to
#: the reason surfaced in the step result. ``vba_replace_component_source`` /
#: ``vba_insert_code`` require an ``expected_source_hash`` revision token that a
#: declarative step cannot carry; they stay in the allowlist because they are
#: bounded service operations, but the runner fails them closed.
_UNSUPPORTED_ACTIONS: dict[TestAction, str] = {
    TestAction.VBA_REPLACE_COMPONENT_SOURCE: (
        "vba_replace_component_source needs an expected_source_hash revision token that "
        "a declarative step cannot carry; use the guarded MCP tools instead."
    ),
    TestAction.VBA_INSERT_CODE: (
        "vba_insert_code needs an expected_source_hash revision token that a declarative "
        "step cannot carry; use the guarded MCP tools instead."
    ),
}

#: Macro outcomes that mean "the macro was never dispatched" (policy /
#: background-safety / preflight block). These block the step instead of failing it.
_BLOCKING_MACRO_STATUSES: frozenset[MacroRunStatus] = frozenset({MacroRunStatus.BLOCKED})

#: Characters that are not allowed in a persisted artifact / case name component.
_UNSAFE_NAME_CHARS = re.compile(r"[^A-Za-z0-9._-]+")

#: Closed action families (exhaustive over :class:`TestAction`).
_SHEET_ACTIONS: frozenset[TestAction] = frozenset(
    {
        TestAction.SHEET_ADD,
        TestAction.SHEET_RENAME,
        TestAction.SHEET_COPY,
        TestAction.SHEET_DELETE,
        TestAction.SHEET_SET_VISIBILITY,
    }
)

_TABLE_ACTIONS: frozenset[TestAction] = frozenset(
    {
        TestAction.TABLE_CREATE,
        TestAction.TABLE_APPEND_ROWS,
        TestAction.TABLE_INSERT_ROW,
        TestAction.TABLE_UPDATE_ROWS,
        TestAction.TABLE_DELETE_ROWS,
        TestAction.TABLE_ADD_COLUMN,
        TestAction.TABLE_DELETE_COLUMN,
        TestAction.TABLE_RESIZE,
    }
)

_NAME_ACTIONS: frozenset[TestAction] = frozenset(
    {
        TestAction.NAME_CREATE,
        TestAction.NAME_UPDATE,
        TestAction.NAME_DELETE,
    }
)


# ---------------------------------------------------------------------------
# Orchestration state (no COM)
# ---------------------------------------------------------------------------
@dataclasses.dataclass
class RunContext:
    """Immutable per-suite context shared by every case of one run."""

    suite: str
    workbook_id: str
    baseline_snapshot_id: str
    baseline_path: str
    baseline_sha256: str


@dataclasses.dataclass
class CaseRun:
    """Mutable per-case execution state (no COM proxies, no Excel objects)."""

    case_id: str
    baseline_sha256: str
    workbook_id: str | None = None
    working_copy_path: str | None = None
    working_copy_sha256: str | None = None
    snapshot_id: str | None = None
    compile_result: CompileResult | None = None
    macro_result: MacroRunResult | None = None
    compile_failed: bool = False
    blocked_reason: str | None = None
    step_results: list[TestStepResult] = dataclasses.field(default_factory=list)
    notes: list[str] = dataclasses.field(default_factory=list)
    started_at: datetime = dataclasses.field(default_factory=lambda: datetime.now(UTC))

    @property
    def blocked(self) -> bool:
        """Whether the case never ran because its environment was unusable."""

        return self.blocked_reason is not None

    def note(self, message: str) -> None:
        """Append one diagnostic note (never a COM object)."""

        self.notes.append(message)


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------
class TestRunner:
    """Execute bounded declarative suites with snapshot isolation.

    Args:
        worker: The single-STA COM worker (all Excel access runs there).
        session_service: Workbook lifecycle service (opens/closes the disposable
            per-case working copies).
        snapshot_service: Immutable snapshot service (baseline + per-case
            snapshot, which also satisfies the pre-mutation gate).
        compile_service: Non-activating compile service.
        macro_service: Policy-controlled macro execution service.
        sheets_service: Sheet inventory + worksheet mutations.
        tables_service: ListObject inventory + row/column mutations.
        names_service: Defined-name inventory + scoped CRUD.
        artifact_store: Evidence store for the run manifest and case diagnostics.
        run_repository: Persistence of the run result documents.
        path_policy: Allowlist / canonicalisation gate; the runner re-validates
            the baseline path before handing it to Excel.
        audit: Append-only JSONL audit trail.
        shapes_service: Optional shape service; when omitted
            ``shape_set_on_action`` fails closed.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        snapshot_service: SnapshotService,
        compile_service: CompileService,
        macro_service: MacroService,
        sheets_service: SheetsService,
        tables_service: TablesService,
        names_service: NamesService,
        artifact_store: ArtifactStore,
        run_repository: RunRepository,
        path_policy: PathPolicy,
        audit: AuditLogger,
        *,
        shapes_service: ShapesService | None = None,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._snapshot = snapshot_service
        self._compile = compile_service
        self._macro = macro_service
        self._sheets = sheets_service
        self._tables = tables_service
        self._names = names_service
        self._artifacts = artifact_store
        self._runs = run_repository
        self._path_policy = path_policy
        self._audit = audit
        self._shapes = shapes_service
        # Suite-level context (installed by run_suite, consumed by run_case).
        self._context: RunContext | None = None
        # The case currently executing (consumed by execute_step /
        # evaluate_assertion when no explicit case is passed).
        self._active_case: CaseRun | None = None

    # ------------------------------------------------------------------
    # Public API (async; no COM here)
    # ------------------------------------------------------------------
    async def run_suite(self, suite: TestSuiteSpec) -> TestRunResult:
        """Run every case of ``suite``, each on a fresh copy of one baseline.

        The suite's ``workbook_id`` must be an open workbook. One immutable
        baseline snapshot is created first (or reused when the suite asks for
        ``baseline=use_existing`` and a snapshot is already known); every case
        then opens its own disposable working copy of that baseline and closes it
        without saving. One case's failure never contaminates the next.

        External macro side effects are **not** rolled back by the snapshot
        (spec §9.4); the result records ``external_side_effects_not_undone``.

        Args:
            suite: The validated declarative suite.

        Returns:
            A :class:`TestRunResult` with per-case and per-step statuses; the run
            is persisted before it is returned.

        Raises:
            ExcelVbaMcpError: If the workbook is unknown or the baseline snapshot
                cannot be produced.
        """

        started_at = datetime.now(UTC)
        run_id = f"run-{uuid4().hex}"
        info = self._resolve_workbook(suite.workbook_id)
        baseline = await self._resolve_baseline(suite)

        context = RunContext(
            suite=suite.suite,
            workbook_id=suite.workbook_id,
            baseline_snapshot_id=baseline.snapshot_id,
            baseline_path=baseline.stored_path,
            baseline_sha256=baseline.sha256,
        )
        workbook = {"canonical_path": info.source_path or info.path}
        self._record_audit(
            tool="test_run_start",
            params={
                "test_run_id": run_id,
                "suite": suite.suite,
                "workbook_id": suite.workbook_id,
                "case_count": len(suite.cases),
                "baseline_snapshot_id": baseline.snapshot_id,
                "baseline_sha256": baseline.sha256,
            },
            workbook=workbook,
            snapshot_id=baseline.snapshot_id,
            result={"test_run_id": run_id, "status": TestRunStatus.RUNNING.value},
        )

        previous = self._context
        self._context = context
        try:
            cases = [await self._run_case_isolated(case) for case in suite.cases]
        finally:
            self._context = previous

        finished_at = datetime.now(UTC)
        result = TestRunResult(
            test_run_id=run_id,
            suite=suite.suite,
            status=_run_status(cases),
            workbook_id=suite.workbook_id,
            baseline_snapshot_id=baseline.snapshot_id,
            cases=cases,
            started_at=started_at,
            finished_at=finished_at,
            duration_ms=_elapsed_ms(started_at, finished_at),
            artifact_paths=[],
            created_at=started_at,
            updated_at=finished_at,
        )
        result = self._attach_artifacts(run_id, suite, result)
        self._runs.save_run(result)
        self._record_audit(
            tool="test_run_finish",
            params={
                "test_run_id": run_id,
                "suite": suite.suite,
                "status": result.status.value,
                "case_count": len(result.cases),
                "artifact_count": len(result.artifact_paths),
            },
            workbook=workbook,
            snapshot_id=baseline.snapshot_id,
            result={
                "test_run_id": run_id,
                "status": result.status.value,
                "duration_ms": result.duration_ms,
                "external_side_effects_not_undone": True,
            },
        )
        return result

    def get_run(self, run_id: str) -> TestRunResult:
        """Return a previously persisted run (read-only; never touches Excel).

        Args:
            run_id: The ``test_run_id`` reported by :meth:`run_suite`.

        Raises:
            excel_vba_mcp.storage.run_repository.TestRunNotFoundError: If no run
                document exists for ``run_id``.
        """

        return self._runs.get_run(run_id)

    async def run_case(
        self,
        case: TestCaseSpec,
        *,
        context: RunContext | None = None,
    ) -> TestCaseResult:
        """Run one case on its own disposable copy of the baseline.

        Sequence: open a fresh working copy of the baseline snapshot → snapshot
        that copy (pre-mutation gate + per-case evidence) → execute the steps in
        order → evaluate the assertions → close the working copy **without
        saving**.

        Args:
            case: The declarative case to execute.
            context: The suite context to use; defaults to the context installed
                by :meth:`run_suite`.

        Returns:
            A :class:`TestCaseResult` with step and assertion detail.

        Raises:
            ExcelVbaMcpError: If no suite context is available.
        """

        run_context = context if context is not None else self._require_context()
        run = CaseRun(case_id=case.id, baseline_sha256=run_context.baseline_sha256)
        previous = self._active_case
        self._active_case = run
        try:
            if await self._open_case_copy(run_context, run):
                await self._snapshot_case(run)
                for step in case.steps:
                    run.step_results.append(await self.execute_step(step))
                assertions = [await self.evaluate_assertion(item) for item in case.assertions]
            else:
                assertions = []
        finally:
            self._active_case = previous
            await self._close_case_copy(run)

        finished_at = datetime.now(UTC)
        result = TestCaseResult(
            case_id=case.id,
            status=_case_status(run, assertions),
            steps=list(run.step_results),
            assertions=assertions,
            duration_ms=_elapsed_ms(run.started_at, finished_at),
            working_copy_path=run.working_copy_path,
            notes=_case_notes(run, run_context),
            created_at=run.started_at,
            updated_at=finished_at,
        )
        self._record_audit(
            tool="test_case_finish",
            params={
                "case_id": case.id,
                "suite": run_context.suite,
                "status": result.status.value,
                "step_count": len(result.steps),
                "assertion_count": len(result.assertions),
            },
            workbook={"canonical_path": run_context.baseline_path},
            snapshot_id=run.snapshot_id,
            result={
                "case_id": case.id,
                "status": result.status.value,
                "duration_ms": result.duration_ms,
                "working_copy_sha256": run.working_copy_sha256,
                "baseline_sha256": run_context.baseline_sha256,
                "external_side_effects_not_undone": True,
            },
        )
        return result

    async def execute_step(
        self,
        step: TestStep,
        *,
        case: CaseRun | None = None,
    ) -> TestStepResult:
        """Dispatch one bounded step to the owning application service.

        Args:
            step: The validated step (its ``action`` is already a member of the
                closed :class:`~excel_vba_mcp.domain.test_spec.TestAction` set).
            case: The case state to execute against; defaults to the case
                installed by :meth:`run_case`.

        Returns:
            A :class:`TestStepResult` (``passed`` / ``failed`` / ``skipped``).
            A step that was **not** executed because an upstream condition blocks
            it (a failed compile before a macro) is ``skipped`` with
            ``details["blocked"]=True`` and a ``details["reason"]``; the case
            status then becomes ``blocked``.

        Raises:
            ExcelVbaMcpError: If no case is active.
        """

        run = case if case is not None else self._require_active_case()
        action = TestAction(step.action)
        unsupported = _UNSUPPORTED_ACTIONS.get(action)
        if unsupported is not None:
            return _failed_step(step, ERROR_UNSUPPORTED_ACTION, unsupported)
        if run.blocked:
            return _blocked_step(step, ERROR_CASE_BLOCKED, run.blocked_reason or "case blocked")
        if run.workbook_id is None:
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")

        if action is TestAction.VBA_COMPILE:
            return await self._exec_compile(step, run)
        if action is TestAction.MACRO_RUN:
            return await self._exec_macro(step, run)
        if action in _SHEET_ACTIONS:
            return await self._exec_sheet(step, run, action)
        if action in _TABLE_ACTIONS:
            return await self._exec_table(step, run, action)
        if action in _NAME_ACTIONS:
            return await self._exec_name(step, run, action)
        if action is TestAction.SHAPE_SET_ON_ACTION:
            return await self._exec_shape(step, run)
        # Every TestAction member is covered above; reaching this line means the
        # closed vocabulary grew without a dispatch branch.
        return _failed_step(step, ERROR_UNSUPPORTED_ACTION, f"no dispatch for {action.value}")

    async def evaluate_assertion(
        self,
        assertion: Assertion,
        *,
        case: CaseRun | None = None,
    ) -> AssertionResult:
        """Evaluate one bounded, read-only assertion against the case workbook.

        Args:
            assertion: The validated assertion.
            case: The case state to evaluate against; defaults to the case
                installed by :meth:`run_case`.

        Returns:
            An :class:`AssertionResult`. An assertion that cannot be evaluated
            (no macro ran, no compile ran, unreadable object) is ``passed=False``
            with an explanatory ``message`` — never a silent pass.

        Raises:
            ExcelVbaMcpError: If no case is active.
        """

        run = case if case is not None else self._require_active_case()
        kind = AssertionType(assertion.type)
        if run.blocked or run.workbook_id is None:
            return _assertion_result(
                assertion,
                passed=False,
                message=run.blocked_reason or "the case has no open working copy",
            )
        try:
            return await self._dispatch_assertion(assertion, kind, run, run.workbook_id)
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, message=f"{exc.code.value}: {exc.message}"
            )
        except Exception as exc:  # noqa: BLE001 - an assertion must never abort a case
            return _assertion_result(
                assertion, passed=False, message=f"{exc.__class__.__name__}: {exc}"
            )

    # ------------------------------------------------------------------
    # Case lifecycle (open / snapshot / close)
    # ------------------------------------------------------------------
    async def _run_case_isolated(self, case: TestCaseSpec) -> TestCaseResult:
        """Run one case, converting an unexpected crash into a failed result.

        A single case must never abort the suite: an unexpected exception is
        recorded as a ``failed`` case carrying the error text.
        """

        try:
            return await self.run_case(case)
        except ExcelVbaMcpError as exc:
            return _crashed_case(case, f"{exc.code.value}: {exc.message}")
        except Exception as exc:  # noqa: BLE001 - isolation: one case must not abort the suite
            return _crashed_case(case, f"{exc.__class__.__name__}: {exc}")

    async def _open_case_copy(self, run_context: RunContext, run: CaseRun) -> bool:
        """Open a fresh disposable working copy of the baseline for one case."""

        try:
            canonical = self._path_policy.validate_path(
                run_context.baseline_path, OperationClass.MUTATING
            )
            info = await self._session.open_workbook(
                str(canonical),
                mode=SessionMode.EXECUTE,
                read_only=False,
                timeout_s=_CALL_TIMEOUT_S,
            )
        except ExcelVbaMcpError as exc:
            reason = f"could not open a disposable copy: {exc.code.value}: {exc.message}"
            if exc.code is ErrorCode.MACRO_BLOCKED:
                # Deny-by-default: the baseline path is server-generated, so make
                # the required grant explicit and actionable instead of opaque.
                reason += (
                    " (grant action='excel_open' for canonical_path="
                    f"{run_context.baseline_path!r}, or pre-create the baseline with "
                    "snapshot_create and run the suite with baseline='use_existing')"
                )
            run.blocked_reason = reason
            run.note(reason)
            return False

        run.workbook_id = info.workbook_id
        run.working_copy_path = info.path
        run.working_copy_sha256 = _sha256_file(info.path)
        if run.working_copy_sha256 != run_context.baseline_sha256:
            run.blocked_reason = (
                "the disposable copy does not match the baseline hash "
                f"({run.working_copy_sha256} != {run_context.baseline_sha256})"
            )
            run.note(run.blocked_reason)
        return True

    async def _snapshot_case(self, run: CaseRun) -> None:
        """Snapshot the case working copy (pre-mutation gate + per-case evidence)."""

        if run.workbook_id is None:
            return
        try:
            manifest = await self._snapshot.create_snapshot(run.workbook_id)
        except ExcelVbaMcpError as exc:
            run.blocked_reason = (
                f"could not snapshot the disposable copy: {exc.code.value}: {exc.message}"
            )
            run.note(run.blocked_reason)
            return
        run.snapshot_id = manifest.snapshot_id

    async def _close_case_copy(self, run: CaseRun) -> None:
        """Close the case working copy **without saving** (mutations are discarded)."""

        if run.workbook_id is None:
            return
        with contextlib.suppress(ExcelVbaMcpError):
            await self._session.close_workbook(run.workbook_id, save_changes=False)

    # ------------------------------------------------------------------
    # Step dispatch
    # ------------------------------------------------------------------
    async def _exec_compile(self, step: TestStep, run: CaseRun) -> TestStepResult:
        """Run the non-activating compile probe and gate later macro steps."""

        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        try:
            result = await self._compile.compile_project(workbook_id)
        except ExcelVbaMcpError as exc:
            run.compile_failed = True
            return _failed_step(
                step,
                exc.code.value,
                f"compile probe failed and macro steps are blocked: {exc.message}",
            )
        run.compile_result = result
        if result.status is CompileStatus.COMPILE_ERROR:
            run.compile_failed = True
            return _failed_step(
                step,
                ERROR_COMPILE_FAILED,
                f"compile reported an error ({result.reason}); macro steps are blocked",
                details={"compile_status": result.status.value, "reason": result.reason},
            )
        note = f"compile status={result.status.value} reason={result.reason}"
        if result.status is CompileStatus.INDETERMINATE:
            note += " (no compile error proven; macro steps are not blocked)"
        run.note(note)
        return TestStepResult(
            action=TestAction.VBA_COMPILE.value,
            status=StepStatus.PASSED,
            details={
                "compile_status": result.status.value,
                "reason": result.reason,
                "is_compile": False,
                "macros_blocked": False,
            },
        )

    async def _exec_macro(self, step: TestStep, run: CaseRun) -> TestStepResult:
        """Run one policy-controlled macro, honouring the compile gate."""

        if run.compile_failed:
            return _blocked_step(
                step,
                ERROR_MACROS_BLOCKED_BY_COMPILE,
                "a failed compile step blocks macro execution in this case",
            )
        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        macro = _require_str(step.macro)
        if macro is None:
            return _failed_step(step, ERROR_INVALID_STEP, "macro_run requires 'macro'")

        result = await self._macro.run_macro(
            workbook_id, macro, list(step.args), timeout_s=_MACRO_TIMEOUT_S
        )
        run.macro_result = result
        details: dict[str, Any] = {
            "macro": result.macro_name,
            "macro_status": result.status.value,
            "duration_ms": result.duration_ms,
            "return_value": result.return_value,
            "error_code": result.error_code,
        }
        if result.status is MacroRunStatus.SUCCESS:
            return TestStepResult(
                action=TestAction.MACRO_RUN.value,
                status=StepStatus.PASSED,
                details=details,
            )
        if result.status in _BLOCKING_MACRO_STATUSES:
            # The macro did not run (policy / background-safety / preflight block).
            return _blocked_step(
                step,
                result.error_code or result.status.value,
                result.error_text or f"macro {macro!r} did not run",
                details=details,
            )
        # NOT_FOUND / RUNTIME_ERROR / TIMEOUT / NEEDS_RECOVERY are real failures.
        return _failed_step(
            step,
            result.error_code or result.status.value,
            result.error_text or f"macro {macro!r} ended with {result.status.value}",
            details=details,
        )

    async def _exec_sheet(self, step: TestStep, run: CaseRun, action: TestAction) -> TestStepResult:
        """Dispatch a sheet action to :class:`SheetsService`."""

        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        if action is TestAction.SHEET_ADD:
            name = _first_str(step.sheet, step.new_name)
            if name is None:
                return _failed_step(step, ERROR_INVALID_STEP, "sheet_add requires 'sheet'")
            sheet = await self._sheets.add_sheet(workbook_id, name)
            return _passed_step(step, {"sheet": sheet.name, "index": sheet.index})
        if action is TestAction.SHEET_RENAME:
            name = _require_str(step.sheet)
            new_name = _require_str(step.new_name)
            if name is None or new_name is None:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "sheet_rename requires 'sheet' and 'new_name'"
                )
            sheet = await self._sheets.rename_sheet(workbook_id, name, new_name)
            return _passed_step(step, {"sheet": sheet.name, "code_name": sheet.code_name})
        if action is TestAction.SHEET_COPY:
            name = _require_str(step.sheet)
            if name is None:
                return _failed_step(step, ERROR_INVALID_STEP, "sheet_copy requires 'sheet'")
            sheet = await self._sheets.copy_sheet(workbook_id, name)
            return _passed_step(step, {"sheet": sheet.name, "index": sheet.index})
        if action is TestAction.SHEET_DELETE:
            name = _require_str(step.sheet)
            if name is None:
                return _failed_step(step, ERROR_INVALID_STEP, "sheet_delete requires 'sheet'")
            sheet = await self._sheets.delete_sheet(workbook_id, name)
            return _passed_step(step, {"sheet": sheet.name})
        # SHEET_SET_VISIBILITY
        name = _require_str(step.sheet)
        if name is None or step.visible is None:
            return _failed_step(
                step, ERROR_INVALID_STEP, "sheet_set_visibility requires 'sheet' and 'visible'"
            )
        sheet = await self._sheets.set_visibility(workbook_id, name, bool(step.visible))
        return _passed_step(step, {"sheet": sheet.name, "visibility": sheet.visibility.value})

    async def _exec_table(self, step: TestStep, run: CaseRun, action: TestAction) -> TestStepResult:
        """Dispatch a table action to :class:`TablesService`."""

        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        if action is TestAction.TABLE_CREATE:
            sheet = _require_str(step.sheet)
            range_or_name = _require_str(step.range_address)
            if sheet is None or range_or_name is None:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "table_create requires 'sheet' and 'range_address'"
                )
            result = await self._tables.create_table(
                workbook_id, sheet, range_or_name, style=step.style
            )
            return _passed_step(step, {"table": result.after.name, "rows": result.after.row_count})
        if action is TestAction.TABLE_RESIZE:
            name = _require_str(step.table)
            new_range = _require_str(step.range_address)
            if name is None or new_range is None:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "table_resize requires 'table' and 'range_address'"
                )
            result = await self._tables.resize_table(workbook_id, name, new_range)
            return _passed_step(
                step, {"table": result.after.name, "range": result.after.range_address}
            )
        if action is TestAction.TABLE_APPEND_ROWS:
            name = _require_str(step.table)
            rows = step.rows if step.rows is not None else ([step.values] if step.values else None)
            if name is None or not rows:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "table_append_rows requires 'table' and 'rows'"
                )
            result = await self._tables.append_table_rows(workbook_id, name, rows)
            return _passed_step(step, {"table": result.after.name, "rows": result.after.row_count})
        if action is TestAction.TABLE_INSERT_ROW:
            name = _require_str(step.table)
            if name is None or step.row_index is None:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "table_insert_row requires 'table' and 'row_index'"
                )
            result = await self._tables.insert_table_row(
                workbook_id, name, step.row_index, step.values
            )
            return _passed_step(step, {"table": result.after.name, "rows": result.after.row_count})
        if action is TestAction.TABLE_UPDATE_ROWS:
            name = _require_str(step.table)
            updates = step.updates if step.updates is not None else _updates_from_step(step)
            if name is None or not updates:
                return _failed_step(
                    step,
                    ERROR_INVALID_STEP,
                    "table_update_rows requires 'table' plus 'updates' or 'key'/'values'",
                )
            result = await self._tables.update_table_rows(workbook_id, name, updates)
            return _passed_step(
                step, {"table": result.after.name, "affected": result.affected_rows}
            )
        if action is TestAction.TABLE_DELETE_ROWS:
            name = _require_str(step.table)
            indices = step.indices if step.indices is not None else _indices_from_step(step)
            if name is None or (not indices and step.key is None):
                return _failed_step(
                    step,
                    ERROR_INVALID_STEP,
                    "table_delete_rows requires 'table' plus 'indices' or 'key'",
                )
            result = await self._tables.delete_table_rows(
                workbook_id, name, indices or None, step.key
            )
            return _passed_step(
                step, {"table": result.after.name, "affected": result.affected_rows}
            )
        if action is TestAction.TABLE_ADD_COLUMN:
            name = _require_str(step.table)
            header = _require_str(step.column_name)
            if name is None or header is None:
                return _failed_step(
                    step, ERROR_INVALID_STEP, "table_add_column requires 'table' and 'column_name'"
                )
            result = await self._tables.add_table_column(workbook_id, name, header, step.position)
            return _passed_step(
                step, {"table": result.after.name, "columns": len(result.after.columns)}
            )
        # TABLE_DELETE_COLUMN
        name = _require_str(step.table)
        header = _require_str(step.column_name)
        if name is None or header is None:
            return _failed_step(
                step, ERROR_INVALID_STEP, "table_delete_column requires 'table' and 'column_name'"
            )
        result = await self._tables.delete_table_column(workbook_id, name, header)
        return _passed_step(
            step, {"table": result.after.name, "columns": len(result.after.columns)}
        )

    async def _exec_name(self, step: TestStep, run: CaseRun, action: TestAction) -> TestStepResult:
        """Dispatch a defined-name action to :class:`NamesService`."""

        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        name = _require_str(step.name)
        if name is None:
            return _failed_step(step, ERROR_INVALID_STEP, f"{action.value} requires 'name'")
        scope = _scope(step.scope)
        if scope is None:
            return _failed_step(
                step,
                ERROR_INVALID_STEP,
                f"{action.value} requires scope 'workbook' or 'worksheet'",
            )
        sheet = step.sheet
        if scope is NameScope.WORKSHEET and sheet is None:
            return _failed_step(
                step, ERROR_INVALID_STEP, "a worksheet-scoped name requires 'sheet'"
            )
        if action is TestAction.NAME_DELETE:
            result = await self._names.delete_name(workbook_id, name, scope=scope, sheet=sheet)
            return _passed_step(step, {"name": result.name, "scope": result.scope.value})
        refers_to = _require_str(step.refers_to)
        if refers_to is None:
            return _failed_step(step, ERROR_INVALID_STEP, f"{action.value} requires 'refers_to'")
        if action is TestAction.NAME_CREATE:
            info = await self._names.create_name(
                workbook_id, name, refers_to, scope=scope, sheet=sheet
            )
        else:
            info = await self._names.update_name(
                workbook_id, name, refers_to, scope=scope, sheet=sheet
            )
        return _passed_step(step, {"name": info.name, "refers_to": info.refers_to})

    async def _exec_shape(self, step: TestStep, run: CaseRun) -> TestStepResult:
        """Assign ``Shape.OnAction`` (the macro is never executed)."""

        if self._shapes is None:
            return _failed_step(
                step,
                ERROR_UNSUPPORTED_ACTION,
                "shape_set_on_action requires the shapes service, which this runner was not given",
            )
        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        name = _require_str(step.name)
        macro = _require_str(step.macro)
        if name is None or macro is None:
            return _failed_step(
                step, ERROR_INVALID_STEP, "shape_set_on_action requires 'name' and 'macro'"
            )
        shape = await self._shapes.set_on_action(workbook_id, name, macro)
        return _passed_step(step, {"shape": shape.name, "on_action": shape.on_action})

    # ------------------------------------------------------------------
    # Assertion dispatch
    # ------------------------------------------------------------------
    async def _dispatch_assertion(
        self,
        assertion: Assertion,
        kind: AssertionType,
        run: CaseRun,
        workbook_id: str,
    ) -> AssertionResult:
        """Evaluate one assertion kind (always read-only)."""

        if kind is AssertionType.MACRO_STATUS:
            return _macro_status_assertion(assertion, run)
        if kind is AssertionType.MACRO_RETURN_VALUE:
            return _macro_return_assertion(assertion, run)
        if kind is AssertionType.MACRO_DURATION_LTE:
            return _macro_duration_assertion(assertion, run)
        if kind is AssertionType.COMPILE_STATUS:
            return _compile_status_assertion(assertion, run)
        if kind is AssertionType.SHEET_EXISTS or kind is AssertionType.SHEET_VISIBILITY:
            return await self._sheet_assertion(assertion, kind, workbook_id)
        if kind in _TABLE_ASSERTION_KINDS:
            return await self._table_assertion(assertion, kind, workbook_id)
        if kind is AssertionType.NAME_EXISTS or kind is AssertionType.NAME_REFERS_TO:
            return await self._name_assertion(assertion, kind, workbook_id)
        if kind is AssertionType.CELL_VALUE_EQUALS or kind is AssertionType.READ_CELL:
            return await self._cell_assertion(assertion, run, workbook_id)
        if kind is AssertionType.SNAPSHOT_EXISTS:
            return _snapshot_assertion(assertion, run)
        if kind is AssertionType.FILE_HASH_EXISTS:
            return _file_hash_assertion(assertion, run)
        # ARTIFACT_EXISTS
        return _artifact_assertion(assertion, self._artifacts)

    async def _sheet_assertion(
        self, assertion: Assertion, kind: AssertionType, workbook_id: str
    ) -> AssertionResult:
        """``sheet_exists`` / ``sheet_visibility`` (read-only)."""

        name = _require_str(assertion.name) or _require_str(assertion.sheet)
        if name is None:
            return _assertion_result(assertion, passed=False, message="needs 'name'")
        try:
            sheet = await self._sheets.get_sheet(workbook_id, name)
        except ExcelVbaMcpError:
            return _assertion_result(
                assertion,
                passed=False,
                expected=True,
                actual=False,
                message=f"sheet {name!r} not found",
            )
        if kind is AssertionType.SHEET_EXISTS:
            return _assertion_result(assertion, passed=True, expected=True, actual=True)
        expected = str(assertion.expected or assertion.visibility or "visible")
        actual = sheet.visibility.value
        return _assertion_result(
            assertion,
            passed=actual == expected,
            expected=expected,
            actual=actual,
            message=None if actual == expected else f"visibility {actual!r} != {expected!r}",
        )

    async def _table_assertion(
        self, assertion: Assertion, kind: AssertionType, workbook_id: str
    ) -> AssertionResult:
        """Table existence / row-count / column assertions (read-only)."""

        name = _require_str(assertion.table)
        if name is None:
            return _assertion_result(assertion, passed=False, message="needs 'table'")
        try:
            table = await self._tables.get_table(workbook_id, name)
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion,
                passed=False,
                actual=False,
                message=f"table {name!r} not readable: {exc.message}",
            )
        if kind is AssertionType.TABLE_EXISTS:
            return _assertion_result(assertion, passed=True, expected=True, actual=True)
        if kind is AssertionType.TABLE_ROW_COUNT:
            return _compare_bound(assertion, actual=table.row_count, mode="eq")
        if kind is AssertionType.TABLE_ROW_COUNT_GTE:
            return _compare_bound(assertion, actual=table.row_count, mode="gte")
        if kind is AssertionType.TABLE_COLUMN_COUNT:
            return _compare_bound(assertion, actual=len(table.columns), mode="eq")
        expected_columns = _string_list(assertion.expected)
        actual_columns = [column.name for column in table.columns]
        return _assertion_result(
            assertion,
            passed=actual_columns == expected_columns,
            expected=expected_columns,
            actual=actual_columns,
            message=None if actual_columns == expected_columns else "column names differ",
        )

    async def _name_assertion(
        self, assertion: Assertion, kind: AssertionType, workbook_id: str
    ) -> AssertionResult:
        """``name_exists`` / ``name_refers_to`` (read-only)."""

        name = _require_str(assertion.name)
        if name is None:
            return _assertion_result(assertion, passed=False, message="needs 'name'")
        scope = _scope(assertion.expected) if isinstance(assertion.expected, str) else None
        if kind is AssertionType.NAME_REFERS_TO and scope is not None:
            try:
                info = await self._names.get_name(workbook_id, name, scope, sheet=assertion.sheet)
            except ExcelVbaMcpError as exc:
                return _assertion_result(
                    assertion, passed=False, message=f"name {name!r} not readable: {exc.message}"
                )
            expected = (
                assertion.refers_to if assertion.refers_to is not None else assertion.expected
            )
            return _assertion_result(
                assertion,
                passed=info.refers_to == expected,
                expected=expected,
                actual=info.refers_to,
                message=None if info.refers_to == expected else "RefersTo differs",
            )
        try:
            names = await self._names.list_names(workbook_id)
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, message=f"defined names unreadable: {exc.message}"
            )
        matches = [item for item in names if item.name.casefold() == name.casefold()]
        if not matches:
            return _assertion_result(
                assertion,
                passed=False,
                expected=True,
                actual=False,
                message=f"name {name!r} not found",
            )
        if kind is AssertionType.NAME_EXISTS:
            return _assertion_result(assertion, passed=True, expected=True, actual=True)
        expected = assertion.refers_to if assertion.refers_to is not None else assertion.expected
        actual = matches[0].refers_to
        return _assertion_result(
            assertion,
            passed=actual == expected,
            expected=expected,
            actual=actual,
            message=None if actual == expected else "RefersTo differs",
        )

    async def _cell_assertion(
        self, assertion: Assertion, run: CaseRun, workbook_id: str
    ) -> AssertionResult:
        """``read_cell`` / ``cell_value_equals`` (targeted, read-only)."""

        sheet = _require_str(assertion.sheet)
        cell = _require_str(assertion.cell)
        if sheet is None or cell is None or run.working_copy_path is None:
            return _assertion_result(assertion, passed=False, message="needs 'sheet' and 'cell'")
        actual = await self._read_cell(workbook_id, run.working_copy_path, sheet, cell)
        return _assertion_result(
            assertion,
            passed=_values_equal(actual, assertion.expected),
            expected=assertion.expected,
            actual=actual,
            message=None if _values_equal(actual, assertion.expected) else "cell value differs",
        )

    async def _read_cell(self, workbook_id: str, working_path: str, sheet: str, cell: str) -> Any:
        """Read one cell value on the STA worker (read-only, never activates)."""

        app_adapter = self._app_adapter()

        def read_fn() -> Any:
            wb = _find_open_workbook(app_adapter, working_path)
            if wb is None:
                raise ExcelVbaMcpError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": working_path},
                )
            worksheet = wb.Worksheets(sheet)
            return _coerce_cell(worksheet.Range(cell).Value)

        read_fn.__name__ = "read_cell"
        return await self._worker.submit(
            read_fn,
            timeout=_CALL_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )

    # ------------------------------------------------------------------
    # Internal helpers (non-COM)
    # ------------------------------------------------------------------
    async def _resolve_baseline(self, suite: TestSuiteSpec) -> SnapshotManifest:
        """Return the immutable baseline manifest for the suite."""

        known = self._snapshot.has_snapshot(suite.workbook_id)
        if suite.baseline is BaselinePolicy.USE_EXISTING and known:
            return self._snapshot.require_snapshot(suite.workbook_id)
        return await self._snapshot.create_snapshot(suite.workbook_id)

    def _resolve_workbook(self, workbook_id: str) -> WorkbookInfo:
        """Resolve an open ``workbook_id`` (fail closed when it is not open)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise ExcelVbaMcpError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            code=ErrorCode.SNAPSHOT_FAILED,
            context={"workbook_id": workbook_id, "tool": "test_run_suite"},
        )

    def _require_context(self) -> RunContext:
        if self._context is None:
            raise ExcelVbaMcpError(
                "run_case requires a suite context; call run_suite first.",
                context={"tool": "test_runner"},
            )
        return self._context

    def _require_active_case(self) -> CaseRun:
        if self._active_case is None:
            raise ExcelVbaMcpError(
                "No test case is active; run_case installs the case context.",
                context={"tool": "test_runner"},
            )
        return self._active_case

    def _app_adapter(self) -> ExcelAppAdapter:
        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "test_runner"},
            )
        return adapter

    def _attach_artifacts(
        self, run_id: str, suite: TestSuiteSpec, result: TestRunResult
    ) -> TestRunResult:
        """Persist the run manifest + per-case diagnostics and reference them.

        ``TestRunResult`` carries no free-form note field, so the run manifest
        artifact additionally records ``external_side_effects_not_undone`` and the
        persistence diagnostics; the per-case documents carry the case notes.
        """

        paths: list[str] = list(result.artifact_paths)
        notes: list[str] = []
        payload: dict[str, Any] = {
            **result.model_dump(mode="json"),
            "external_side_effects_not_undone": True,
            "notes": notes,
        }
        manifest = _save_json_artifact(
            self._artifacts,
            name=f"{run_id}--manifest",
            payload=payload,
            artifact_type=ArtifactType.MANIFEST,
            workbook_id=suite.workbook_id,
            notes=notes,
        )
        if manifest is not None:
            paths.append(manifest)
        for case in result.cases:
            stored = _save_json_artifact(
                self._artifacts,
                name=f"{run_id}--{_safe_name(case.case_id)}",
                payload=case.model_dump(mode="json"),
                artifact_type=ArtifactType.DIAGNOSTIC,
                workbook_id=suite.workbook_id,
                notes=notes,
            )
            if stored is not None:
                paths.append(stored)
        return result.model_copy(update={"artifact_paths": paths})

    def _record_audit(
        self,
        *,
        tool: str,
        params: dict[str, Any],
        result: dict[str, Any] | None = None,
        workbook: dict[str, Any] | None = None,
        snapshot_id: str | None = None,
        operation_class: str = "executing",
    ) -> None:
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=getattr(self._session, "session_id", None),
            workbook=workbook,
            snapshot_id=snapshot_id,
            result=result,
        )


#: Assertion kinds evaluated against a ListObject.
_TABLE_ASSERTION_KINDS: frozenset[AssertionType] = frozenset(
    {
        AssertionType.TABLE_EXISTS,
        AssertionType.TABLE_ROW_COUNT,
        AssertionType.TABLE_ROW_COUNT_GTE,
        AssertionType.TABLE_COLUMN_COUNT,
        AssertionType.TABLE_COLUMN_NAMES,
    }
)


# ---------------------------------------------------------------------------
# Step / assertion result builders (no COM)
# ---------------------------------------------------------------------------
def _passed_step(step: TestStep, details: dict[str, Any]) -> TestStepResult:
    """Build a ``passed`` step result."""

    return TestStepResult(
        action=TestAction(step.action).value,
        status=StepStatus.PASSED,
        details=details,
    )


def _failed_step(
    step: TestStep,
    error_code: str,
    message: str,
    *,
    details: dict[str, Any] | None = None,
) -> TestStepResult:
    """Build a ``failed`` step result."""

    return TestStepResult(
        action=TestAction(step.action).value,
        status=StepStatus.FAILED,
        error_code=error_code,
        error_text=message,
        details={"reason": message, **(details or {})},
    )


def _blocked_step(
    step: TestStep,
    reason_code: str,
    message: str,
    *,
    details: dict[str, Any] | None = None,
) -> TestStepResult:
    """Build a step result for a step that was **not** executed.

    :class:`~excel_vba_mcp.domain.statuses.StepStatus` has no ``blocked`` member,
    so a blocked step is recorded as ``skipped`` with ``details["blocked"]=True``
    and a ``details["reason"]``; the case-level status then becomes ``blocked``.
    """

    return TestStepResult(
        action=TestAction(step.action).value,
        status=StepStatus.SKIPPED,
        error_code=reason_code,
        error_text=message,
        details={"blocked": True, "reason": message, "reason_code": reason_code, **(details or {})},
    )


def _assertion_result(
    assertion: Assertion,
    *,
    passed: bool,
    expected: Any | None = None,
    actual: Any | None = None,
    message: str | None = None,
) -> AssertionResult:
    """Build an :class:`AssertionResult` tagged with the assertion kind."""

    return AssertionResult(
        type=AssertionType(assertion.type).value,
        passed=passed,
        expected=expected,
        actual=actual,
        message=message,
    )


def _compare_bound(assertion: Assertion, *, actual: int, mode: str) -> AssertionResult:
    """Compare an integer actual value against ``assertion.value``."""

    bound = assertion.value
    if bound is None:
        return _assertion_result(
            assertion,
            passed=False,
            actual=actual,
            message=f"{AssertionType(assertion.type).value} requires a numeric 'value'",
        )
    passed = actual == bound if mode == "eq" else actual >= bound
    operator = "==" if mode == "eq" else ">="
    return _assertion_result(
        assertion,
        passed=passed,
        expected={"operator": operator, "value": bound},
        actual=actual,
        message=None if passed else f"{actual} {operator} {bound} is false",
    )


def _macro_status_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``macro_status`` against the case's last macro run."""

    if run.macro_result is None:
        return _assertion_result(assertion, passed=False, message="no macro step ran in this case")
    actual = run.macro_result.status.value
    expected = (
        assertion.expected if assertion.expected is not None else MacroRunStatus.SUCCESS.value
    )
    return _assertion_result(
        assertion,
        passed=str(actual) == str(expected),
        expected=expected,
        actual=actual,
        message=None
        if str(actual) == str(expected)
        else f"macro status {actual!r} != {expected!r}",
    )


def _macro_return_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``macro_return_value`` against the case's last macro run."""

    if run.macro_result is None:
        return _assertion_result(assertion, passed=False, message="no macro step ran in this case")
    actual = run.macro_result.return_value
    return _assertion_result(
        assertion,
        passed=_values_equal(actual, assertion.expected),
        expected=assertion.expected,
        actual=actual,
        message=None if _values_equal(actual, assertion.expected) else "return value differs",
    )


def _macro_duration_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``macro_duration_lte`` against the case's last macro run."""

    if run.macro_result is None:
        return _assertion_result(assertion, passed=False, message="no macro step ran in this case")
    if assertion.value is None:
        return _assertion_result(
            assertion, passed=False, message="macro_duration_lte requires a numeric 'value'"
        )
    actual = run.macro_result.duration_ms
    return _assertion_result(
        assertion,
        passed=actual <= assertion.value,
        expected={"operator": "<=", "value": assertion.value},
        actual=actual,
        message=None if actual <= assertion.value else f"{actual}ms exceeds {assertion.value}ms",
    )


def _compile_status_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``compile_status`` against the case's compile probe.

    By default the assertion means *no compile error was proven*, so it accepts
    ``compiled`` / ``already_compiled`` / ``indeterminate`` (the non-activating
    build reports ``indeterminate``). Pass an explicit ``expected`` to require a
    specific status.
    """

    if run.compile_result is None:
        return _assertion_result(
            assertion, passed=False, message="no compile step ran in this case"
        )
    actual = run.compile_result.status.value
    expected_list = _string_list(
        assertion.expected
        if assertion.expected is not None
        else [
            CompileStatus.COMPILED.value,
            CompileStatus.ALREADY_COMPILED.value,
            CompileStatus.INDETERMINATE.value,
        ]
    )
    return _assertion_result(
        assertion,
        passed=actual in expected_list,
        expected=expected_list,
        actual=actual,
        message=None
        if actual in expected_list
        else f"compile status {actual!r} not in {expected_list}",
    )


def _snapshot_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``snapshot_exists`` — the case must own a snapshot."""

    stored = run.snapshot_id is not None
    return _assertion_result(
        assertion,
        passed=stored,
        expected=True,
        actual=stored,
        message=None if stored else "the case has no snapshot",
    )


def _file_hash_assertion(assertion: Assertion, run: CaseRun) -> AssertionResult:
    """``file_hash_exists`` — the working copy hash equals the baseline hash."""

    if run.working_copy_path is None:
        return _assertion_result(
            assertion, passed=False, message="the case has no working copy to hash"
        )
    actual = _sha256_file(run.working_copy_path)
    expected = assertion.hash if assertion.hash is not None else run.baseline_sha256
    return _assertion_result(
        assertion,
        passed=actual == expected,
        expected=expected,
        actual=actual,
        message=None if actual == expected else "working copy hash differs from the baseline",
    )


def _artifact_assertion(assertion: Assertion, store: ArtifactStore) -> AssertionResult:
    """``artifact_exists`` — a stored artifact name contains the expectation."""

    needle = assertion.expected if assertion.expected is not None else assertion.name
    if needle is None:
        return _assertion_result(
            assertion, passed=False, message="artifact_exists requires 'expected' or 'name'"
        )
    try:
        artifacts = store.list_artifacts()
    except ExcelVbaMcpError as exc:
        return _assertion_result(
            assertion, passed=False, message=f"artifact index unreadable: {exc.message}"
        )
    matches = [item.name for item in artifacts if str(needle).lower() in item.name.lower()]
    return _assertion_result(
        assertion,
        passed=bool(matches),
        expected=str(needle),
        actual=matches,
        message=None if matches else f"no artifact matches {str(needle)!r}",
    )


# ---------------------------------------------------------------------------
# Status mapping
# ---------------------------------------------------------------------------
def _case_status(run: CaseRun, assertions: list[AssertionResult]) -> CaseStatus:
    """Map one case's steps + assertions onto :class:`CaseStatus`.

    Precedence: ``failed`` > ``blocked`` > ``indeterminate`` > ``passed``. A case
    with no assertions verifies nothing, so it is ``indeterminate``.
    """

    if any(step.status is StepStatus.FAILED for step in run.step_results):
        return CaseStatus.FAILED
    if run.blocked or any(step.details.get("blocked") for step in run.step_results):
        return CaseStatus.BLOCKED
    if any(not assertion.passed for assertion in assertions):
        return CaseStatus.FAILED
    if not assertions:
        return CaseStatus.INDETERMINATE
    return CaseStatus.PASSED


def _run_status(cases: list[TestCaseResult]) -> TestRunStatus:
    """Map case statuses onto :class:`TestRunStatus` (worst case wins)."""

    statuses = {case.status for case in cases}
    if CaseStatus.FAILED in statuses:
        return TestRunStatus.FAILED
    if CaseStatus.BLOCKED in statuses:
        return TestRunStatus.BLOCKED
    if CaseStatus.INDETERMINATE in statuses or not statuses:
        return TestRunStatus.INDETERMINATE
    return TestRunStatus.PASSED


def _crashed_case(case: TestCaseSpec, message: str) -> TestCaseResult:
    """Build a ``failed`` case result for a case that raised unexpectedly."""

    now = datetime.now(UTC)
    return TestCaseResult(
        case_id=case.id,
        status=CaseStatus.FAILED,
        steps=[],
        assertions=[],
        duration_ms=0,
        working_copy_path=None,
        notes=[f"case aborted before completion: {message}", EXTERNAL_SIDE_EFFECTS_NOTE],
        created_at=now,
        updated_at=now,
    )


def _case_notes(run: CaseRun, run_context: RunContext) -> list[str]:
    """Build the per-case note list (baseline identity + external-effects notice)."""

    notes: list[str] = []
    if run.working_copy_sha256 is not None:
        notes.append(
            f"baseline_sha256={run.baseline_sha256} "
            f"working_copy_sha256={run.working_copy_sha256} "
            f"baseline_snapshot_id={run_context.baseline_snapshot_id}"
        )
    if run.snapshot_id is not None:
        notes.append(f"case_snapshot_id={run.snapshot_id}")
    notes.extend(run.notes)
    notes.append(EXTERNAL_SIDE_EFFECTS_NOTE)
    return notes


# ---------------------------------------------------------------------------
# Parameter / value helpers (no COM)
# ---------------------------------------------------------------------------
def _require_workbook(run: CaseRun) -> str | None:
    """Return the case workbook id when the case owns an open working copy."""

    return run.workbook_id


def _require_str(value: Any) -> str | None:
    """Return a non-empty stripped string, or ``None``."""

    if isinstance(value, str) and value.strip():
        return value.strip()
    return None


def _first_str(*values: Any) -> str | None:
    """Return the first usable string among ``values``."""

    for value in values:
        resolved = _require_str(value)
        if resolved is not None:
            return resolved
    return None


def _scope(value: Any) -> NameScope | None:
    """Resolve a scope string to :class:`NameScope` (``None`` when absent/invalid)."""

    if value is None:
        return None
    try:
        return NameScope(str(value))
    except ValueError:
        return None


def _string_list(value: Any) -> list[str]:
    """Coerce a scalar / list expectation into a list of strings."""

    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [str(item) for item in value]
    return [str(value)]


def _updates_from_step(step: TestStep) -> list[dict[str, Any]]:
    """Derive a ``table_update_rows`` payload from ``key`` / ``values`` / ``row_index``."""

    if step.values is None:
        return []
    if step.key is not None:
        return [{"where": dict(step.key), "values": dict(step.values)}]
    if step.row_index is not None:
        return [{"index": step.row_index, "values": dict(step.values)}]
    return []


def _indices_from_step(step: TestStep) -> list[int]:
    """Derive a ``table_delete_rows`` index list from ``row_index``."""

    return [step.row_index] if step.row_index is not None else []


def _coerce_cell(value: Any) -> Any:
    """Convert a raw COM cell value into a JSON-friendly scalar."""

    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, str)):
        return value
    return str(value)


def _values_equal(actual: Any, expected: Any) -> bool:
    """Compare an actual value with an expectation (numeric-aware)."""

    if actual is None or expected is None:
        return actual is None and expected is None
    if isinstance(actual, bool) or isinstance(expected, bool):
        return bool(actual) is bool(expected) and str(actual).lower() == str(expected).lower()
    if isinstance(actual, (int, float)) and isinstance(expected, (int, float)):
        return math.isclose(float(actual), float(expected), rel_tol=1e-9, abs_tol=1e-9)
    if isinstance(actual, (int, float)) and isinstance(expected, str):
        try:
            return math.isclose(float(actual), float(expected), rel_tol=1e-9, abs_tol=1e-9)
        except ValueError:
            return False
    return str(actual).strip() == str(expected).strip()


def _sha256_file(path: str | Path) -> str:
    """Return the SHA-256 hex digest of ``path`` (empty string when absent)."""

    target = Path(path)
    if not target.is_file():
        return ""
    digest = hashlib.sha256()
    with target.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _elapsed_ms(started_at: datetime, finished_at: datetime) -> int:
    """Whole milliseconds between two timestamps."""

    return int(round((finished_at - started_at).total_seconds() * 1000.0))


def _safe_name(value: str) -> str:
    """Sanitize ``value`` into a safe artifact-name component."""

    cleaned = _UNSAFE_NAME_CHARS.sub("-", value).strip("-.")
    return (cleaned or "case")[:80]


def _save_json_artifact(
    store: ArtifactStore,
    *,
    name: str,
    payload: dict[str, Any],
    artifact_type: ArtifactType,
    workbook_id: str,
    notes: list[str],
) -> str | None:
    """Persist one JSON evidence document; degrade to a note on failure."""

    safe = _safe_name(name)
    try:
        info = store.save_artifact(
            safe,
            json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8"),
            artifact_type=artifact_type,
            workbook_id=workbook_id,
        )
    except ExcelVbaMcpError as exc:
        notes.append(f"artifact {safe!r} not stored: {exc.code.value}: {exc.message}")
        return None
    return info.stored_path


# ---------------------------------------------------------------------------
# STA-thread helpers (run on the worker thread only)
# ---------------------------------------------------------------------------
def _find_open_workbook(app_adapter: ExcelAppAdapter, working_path: str) -> Any | None:
    """Return the open ``Workbook`` proxy matching ``working_path`` (STA thread)."""

    app = app_adapter.app
    if app is None:
        return None
    target = working_path.lower()
    try:
        for wb in app.Workbooks:
            try:
                if str(wb.FullName).lower() == target:
                    return wb
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return None
