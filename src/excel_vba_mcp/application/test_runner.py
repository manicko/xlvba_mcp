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
  application service (compile / macro / sheets / tables / names / shapes, plus the
  Phase 2 ranges and pivot services); the runner owns orchestration and status
  mapping only. The Range/Pivot actions added by task_053 call the **same**
  :class:`~excel_vba_mcp.application.ranges_service.RangesService` and
  :class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
  the direct MCP tools use, so the guard chain (writable working copy →
  snapshot → approval) and every postcondition check apply unchanged. The runner
  contains no COM logic for them.
* **Compile gates macros.** A failed ``vba_compile`` step blocks every later
  ``macro_run`` step of the same case (``skipped`` + ``details["blocked"]``; the
  macro is **not** dispatched). Compile-before-macro is **explicit**: a suite that
  gates macros must list a ``vba_compile`` step first; the runner never
  auto-compiles.
* **Background-safe declaration.** A ``macro_run`` step may carry the data-only
  ``background_safe`` flag: the suite author's declaration that the procedure is
  reviewed non-interactive for the current VBA source revision. When true, the
  runner records it through the guarded
  :meth:`~excel_vba_mcp.application.macro_service.MacroService.classify_background`
  **before** dispatch, so the classify-time heuristic/UI veto and the run-time
  R13 veto still apply; the runner never writes the classification registry
  directly. The declaration is a claim, not proof: the residual boundary R-15
  (dynamic/indirect or transitive UI is not fully covered) still applies, and a
  false declaration is caught only by the deterministic veto or runtime
  containment.
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
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.ranges_service import RangesService
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
    PivotRefreshStatus,
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

#: Phase 2 Range/Cells actions (task_053), dispatched to :class:`RangesService`.
_RANGE_ACTIONS: frozenset[TestAction] = frozenset(
    {
        TestAction.RANGE_GET,
        TestAction.RANGE_READ_VALUES,
        TestAction.RANGE_READ_FORMULAS,
        TestAction.RANGE_GET_FORMAT,
        TestAction.RANGE_WRITE_VALUES,
        TestAction.RANGE_WRITE_FORMULAS,
        TestAction.RANGE_SET_FORMAT,
        TestAction.RANGE_COPY,
        TestAction.RANGE_INSERT,
        TestAction.RANGE_DELETE,
        TestAction.RANGE_MERGE,
        TestAction.RANGE_UNMERGE,
    }
)

#: Phase 2 PivotTables actions (task_053), dispatched to :class:`PivotTablesService`.
_PIVOT_ACTIONS: frozenset[TestAction] = frozenset(
    {
        TestAction.PIVOT_TABLE_REFRESH,
        TestAction.PIVOT_FIELD_SET_LAYOUT,
        TestAction.PIVOT_FIELD_SET_AGGREGATION,
        TestAction.PIVOT_FIELD_APPLY_FILTER,
    }
)

#: Pivot refresh outcomes that prove the refresh did not happen or was refused.
#: They block the case (no false success) instead of failing it.
_BLOCKING_REFRESH_STATUSES: frozenset[PivotRefreshStatus] = frozenset(
    {
        PivotRefreshStatus.UNSUPPORTED_SOURCE,
        PivotRefreshStatus.BLOCKED,
    }
)

#: The default expected refresh status of a ``pivot_refresh_status`` assertion.
_DEFAULT_REFRESH_STATUS = PivotRefreshStatus.REFRESHED.value

#: Sentinel returned by the dotted-property reader when a path does not exist.
_MISSING_PROPERTY = object()


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
        ranges_service: Optional Phase 2 ranges service; when omitted every
            ``range_*`` step fails closed with ``UNSUPPORTED_ACTION``.
        pivot_tables_service: Optional Phase 2 pivot service; when omitted every
            ``pivot_*`` step fails closed with ``UNSUPPORTED_ACTION``.
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
        ranges_service: RangesService | None = None,
        pivot_tables_service: PivotTablesService | None = None,
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
        self._ranges = ranges_service
        self._pivots = pivot_tables_service
        # Suite-level context (installed by run_suite, consumed by run_case).
        self._context: RunContext | None = None
        # The case currently executing (consumed by execute_step /
        # evaluate_assertion when no explicit case is passed).
        self._active_case: CaseRun | None = None
        # Outcomes of the pivot steps of the active case, in execution order
        # (consumed by the pivot assertions). Each entry is
        # ``(action, model_dump(mode="json"))``.
        self._pivot_outcomes: list[tuple[str, dict[str, Any]]] = []

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
        # Per-case execution evidence (no COM): pivot step outcomes feed the
        # pivot assertions. Reset here so one case never inherits another's.
        previous_pivots = self._pivot_outcomes
        self._pivot_outcomes = []
        self._active_case = run
        try:
            if await self._open_case_copy(run_context, run):
                await self._snapshot_case(run)
                recovery_required = False
                for step in case.steps:
                    step_result = await self.execute_step(step)
                    run.step_results.append(step_result)
                    if step_result.details.get("blocked"):
                        # A blocked step (failed compile, refused mutation) means
                        # the case cannot continue truthfully: record a note and
                        # stop issuing further steps.
                        run.note(f"case stopped after a blocked {step_result.action} step")
                        break
                    if step_result.action == TestAction.MACRO_RUN.value and _needs_recovery(run):
                        # B6-004: a NEEDS_RECOVERY macro means the owned worker is
                        # no longer trustworthy. Stop dispatching further steps and
                        # skip assertions rather than failing them against a dead
                        # worker; the failed macro step already makes the case fail.
                        recovery_required = True
                        run.note(
                            "case stopped after a NEEDS_RECOVERY macro step: the worker "
                            "requires recovery, so no later step or assertion was dispatched"
                        )
                        break
                assertions = (
                    []
                    if recovery_required
                    else [await self.evaluate_assertion(item) for item in case.assertions]
                )
            else:
                assertions = []
        finally:
            self._active_case = previous
            self._pivot_outcomes = previous_pivots
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
        if action in _RANGE_ACTIONS:
            return await self._exec_range(step, run, action)
        if action in _PIVOT_ACTIONS:
            return await self._exec_pivot(step, run, action)
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
            if exc.code in (ErrorCode.MACRO_BLOCKED, ErrorCode.APPROVAL_DENIED):
                # Deny-by-default: the baseline path is server-generated, so make
                # the required grant explicit and actionable instead of opaque. After
                # B4 the common restricted-mode denial is APPROVAL_DENIED; accept both
                # codes so the hint is never silently dropped (B6-003).
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

        # B6 (task_b6): a suite author may declare a macro background-safe. Record
        # that declaration through the GUARDED MacroService.classify_background so
        # the classify-time heuristic/UI veto and the run-time R13 veto still apply.
        # Never write the BackgroundSafetyRegistry directly here: that would drop the
        # classify-time guard and become a false-safe path. A refusal (interactive UI,
        # heuristic, unknown procedure) blocks the step and the macro is never
        # dispatched.
        if getattr(step, "background_safe", None) is True:
            try:
                await self._macro.classify_background(workbook_id, macro, True)
            except ExcelVbaMcpError as exc:
                return _blocked_step(step, exc.code.value, f"{macro}: {exc.message}")

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

    async def _exec_range(self, step: TestStep, run: CaseRun, action: TestAction) -> TestStepResult:
        """Dispatch a Range action to :class:`RangesService` (task_053).

        Every branch calls the same service method as the matching
        ``range_*`` MCP tool, so the writable → snapshot → approval guard chain
        and the postcondition re-read apply unchanged. A service refusal (deny-by-
        default approval, missing snapshot, read-only copy) is reported as a
        **blocked** step; a malformed step never reaches the worker.
        """

        if self._ranges is None:
            return _failed_step(
                step,
                ERROR_UNSUPPORTED_ACTION,
                "range steps require the ranges service, which this runner was not given",
            )
        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        sheet = _require_str(step.sheet)
        if sheet is None:
            return _failed_step(step, ERROR_INVALID_STEP, f"{action.value} requires 'sheet'")
        try:
            outcome = await self._call_ranges(step, action, workbook_id, sheet)
        except ExcelVbaMcpError as exc:
            return _blocked_or_failed_service_step(
                step,
                exc,
                blocked_codes=frozenset({ErrorCode.COM_BUSY, ErrorCode.APPROVAL_DENIED}),
            )
        return _passed_step(step, _range_step_details(outcome))

    async def _call_ranges(
        self, step: TestStep, action: TestAction, workbook_id: str, sheet: str
    ) -> Any:
        """Call the single :class:`RangesService` method of one Range action."""

        service = self._ranges
        assert service is not None  # noqa: S101 - guarded by _exec_range
        address = _require_str(step.range_address)
        if action is TestAction.RANGE_GET:
            if address is None:
                raise ExcelVbaMcpError(
                    "range_get requires 'range_address'", code=ErrorCode.INVALID_RANGE
                )
            return await service.get_range(workbook_id, sheet, address)
        if action is TestAction.RANGE_READ_VALUES:
            if address is None:
                raise ExcelVbaMcpError(
                    "range_read_values requires 'range_address'", code=ErrorCode.INVALID_RANGE
                )
            return await service.read_values(
                workbook_id, sheet, address, offset=step.offset or 0, limit=step.limit
            )
        if action is TestAction.RANGE_READ_FORMULAS:
            if address is None:
                raise ExcelVbaMcpError(
                    "range_read_formulas requires 'range_address'", code=ErrorCode.INVALID_RANGE
                )
            return await service.read_formulas(
                workbook_id, sheet, address, offset=step.offset or 0, limit=step.limit
            )
        if action is TestAction.RANGE_GET_FORMAT:
            if address is None:
                raise ExcelVbaMcpError(
                    "range_get_format requires 'range_address'", code=ErrorCode.INVALID_RANGE
                )
            return await service.get_format(workbook_id, sheet, address)
        if action is TestAction.RANGE_WRITE_VALUES:
            if address is None or step.values_grid is None:
                raise ExcelVbaMcpError(
                    "range_write_values requires 'range_address' and 'values_grid'",
                    code=ErrorCode.INVALID_RANGE,
                )
            return await service.write_values(workbook_id, sheet, address, step.values_grid)
        if action is TestAction.RANGE_WRITE_FORMULAS:
            if address is None or step.formulas_grid is None:
                raise ExcelVbaMcpError(
                    "range_write_formulas requires 'range_address' and 'formulas_grid'",
                    code=ErrorCode.INVALID_RANGE,
                )
            return await service.write_formulas(workbook_id, sheet, address, step.formulas_grid)
        if action is TestAction.RANGE_SET_FORMAT:
            if address is None or step.format_patch is None:
                raise ExcelVbaMcpError(
                    "range_set_format requires 'range_address' and 'format_patch'",
                    code=ErrorCode.INVALID_RANGE,
                )
            return await service.set_format(workbook_id, sheet, address, step.format_patch)
        if action is TestAction.RANGE_COPY:
            source = _require_str(step.source_range)
            destination_sheet = _require_str(step.destination_sheet)
            destination = _require_str(step.destination)
            if source is None or destination_sheet is None or destination is None:
                raise ExcelVbaMcpError(
                    "range_copy requires 'source_range', 'destination_sheet' and 'destination'",
                    code=ErrorCode.INVALID_RANGE,
                )
            return await service.copy_range(
                workbook_id, sheet, source, destination_sheet, destination
            )
        if action is TestAction.RANGE_INSERT or action is TestAction.RANGE_DELETE:
            if address is None or step.shift is None:
                raise ExcelVbaMcpError(
                    f"{action.value} requires 'range_address' and an explicit 'shift'",
                    code=ErrorCode.INVALID_RANGE,
                )
            if action is TestAction.RANGE_INSERT:
                return await service.insert_range(
                    workbook_id, sheet, address, step.shift, allow_merged=bool(step.allow_merged)
                )
            return await service.delete_range(
                workbook_id, sheet, address, step.shift, allow_merged=bool(step.allow_merged)
            )
        if action is TestAction.RANGE_MERGE:
            if address is None:
                raise ExcelVbaMcpError(
                    "range_merge requires 'range_address'", code=ErrorCode.INVALID_RANGE
                )
            return await service.merge_range(
                workbook_id,
                sheet,
                address,
                across=bool(step.across),
                allow_data_loss=bool(step.allow_data_loss),
            )
        # RANGE_UNMERGE
        if address is None:
            raise ExcelVbaMcpError(
                "range_unmerge requires 'range_address'", code=ErrorCode.INVALID_RANGE
            )
        return await service.unmerge_range(workbook_id, sheet, address)

    async def _exec_pivot(self, step: TestStep, run: CaseRun, action: TestAction) -> TestStepResult:
        """Dispatch a Pivot action to :class:`PivotTablesService` (task_053).

        The refresh step maps Excel's truthful ``PivotRefreshStatus`` onto the
        step status: ``refreshed`` passes, ``failed`` fails, and
        ``unsupported_source`` / ``blocked`` block the case (the operation never
        ran, so no false success is reported).
        """

        if self._pivots is None:
            return _failed_step(
                step,
                ERROR_UNSUPPORTED_ACTION,
                "pivot steps require the pivot tables service, which this runner was not given",
            )
        workbook_id = _require_workbook(run)
        if workbook_id is None:  # pragma: no cover - guarded by execute_step
            return _failed_step(step, ERROR_CASE_BLOCKED, "the case has no open working copy")
        pivot_table = _require_str(step.pivot_table)
        if pivot_table is None:
            return _failed_step(step, ERROR_INVALID_STEP, f"{action.value} requires 'pivot_table'")
        try:
            outcome = await self._call_pivots(step, action, workbook_id, pivot_table)
        except ExcelVbaMcpError as exc:
            return _blocked_or_failed_service_step(
                step,
                exc,
                blocked_codes=frozenset({ErrorCode.COM_BUSY, ErrorCode.APPROVAL_DENIED}),
            )
        self._pivot_outcomes.append((TestAction(action).value, _jsonable(outcome)))
        return _pivot_step_result(step, action, outcome)

    async def _call_pivots(
        self, step: TestStep, action: TestAction, workbook_id: str, pivot_table: str
    ) -> Any:
        """Call the single :class:`PivotTablesService` method of one Pivot action."""

        service = self._pivots
        assert service is not None  # noqa: S101 - guarded by _exec_pivot
        if action is TestAction.PIVOT_TABLE_REFRESH:
            return await service.refresh_pivot_table(workbook_id, pivot_table)
        field = _require_str(step.field)
        if field is None:
            raise ExcelVbaMcpError(f"{action.value} requires 'field'", code=ErrorCode.INVALID_RANGE)
        if action is TestAction.PIVOT_FIELD_SET_LAYOUT:
            if step.orientation is None:
                raise ExcelVbaMcpError(
                    "pivot_field_set_layout requires 'orientation'", code=ErrorCode.INVALID_RANGE
                )
            return await service.set_pivot_field_layout(
                workbook_id, pivot_table, field, step.orientation, delete=bool(step.delete)
            )
        if action is TestAction.PIVOT_FIELD_SET_AGGREGATION:
            if step.function is None:
                raise ExcelVbaMcpError(
                    "pivot_field_set_aggregation requires 'function'",
                    code=ErrorCode.INVALID_RANGE,
                )
            return await service.set_pivot_field_aggregation(
                workbook_id, pivot_table, field, step.function, caption=step.caption
            )
        # PIVOT_FIELD_APPLY_FILTER
        if step.filter_spec is None:
            raise ExcelVbaMcpError(
                "pivot_field_apply_filter requires 'filter_spec'", code=ErrorCode.INVALID_RANGE
            )
        return await service.apply_pivot_field_filter(
            workbook_id, pivot_table, field, step.filter_spec
        )

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
        if kind in _RANGE_ASSERTION_KINDS:
            return await self._range_assertion(assertion, kind, workbook_id)
        if kind in _PIVOT_ASSERTION_KINDS:
            return await self._pivot_assertion(assertion, kind, run, workbook_id)
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

    async def _range_assertion(
        self, assertion: Assertion, kind: AssertionType, workbook_id: str
    ) -> AssertionResult:
        """Evaluate one Range assertion through :class:`RangesService` (task_053).

        Every branch is read-only: it calls the same read service the direct MCP
        tools use, so the assertion reflects actual Excel state. A read failure
        (unknown sheet, invalid range) is ``passed=False`` with the reason, never
        a silent pass.
        """

        if self._ranges is None:
            return _assertion_result(
                assertion,
                passed=False,
                message=(
                    "range assertions require the ranges service, which this runner was not given"
                ),
            )
        sheet = _require_str(assertion.sheet)
        address = _require_str(assertion.address)
        if sheet is None or address is None:
            return _assertion_result(assertion, passed=False, message="needs 'sheet' and 'address'")
        try:
            return await self._eval_range_assertion(assertion, kind, workbook_id, sheet, address)
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, message=f"{exc.code.value}: {exc.message}"
            )

    async def _eval_range_assertion(
        self,
        assertion: Assertion,
        kind: AssertionType,
        workbook_id: str,
        sheet: str,
        address: str,
    ) -> AssertionResult:
        """Dispatch one Range assertion kind to its read service (STA-backed)."""

        service = self._ranges
        assert service is not None  # noqa: S101 - guarded by _range_assertion
        if kind is AssertionType.RANGE_FORMULA_EQUALS:
            formulas = await service.read_formulas(workbook_id, sheet, address)
            actual = [row[0].formula2 for row in formulas.formulas]
            expected = assertion.expected if isinstance(assertion.expected, list) else [address]
            expected_formulas = [str(item) for item in expected]
            passed = actual == expected_formulas
            return _assertion_result(
                assertion,
                passed=passed,
                expected=expected_formulas,
                actual=actual,
                message=None if passed else f"formulas {actual!r} != {expected_formulas!r}",
            )
        if kind is AssertionType.RANGE_FORMAT_EQUALS:
            info = await service.get_format(workbook_id, sheet, address)
            prop = _require_str(assertion.property)
            if prop is None:
                return _assertion_result(
                    assertion, passed=False, message="range_format_equals requires 'property'"
                )
            actual = _format_property(info, prop)
            passed = actual is not _MISSING_PROPERTY and _values_equal(actual, assertion.expected)
            return _assertion_result(
                assertion,
                passed=passed,
                expected=assertion.expected,
                actual=None if actual is _MISSING_PROPERTY else actual,
                message=(
                    None
                    if passed
                    else f"format property {prop!r} "
                    f"{'<missing>' if actual is _MISSING_PROPERTY else repr(actual)}"
                    f" != {assertion.expected!r}"
                ),
            )
        if kind is AssertionType.RANGE_STRUCTURE:
            resolved = await service.get_range(workbook_id, sheet, address)
            problems = _structure_problems(resolved, assertion)
            return _assertion_result(
                assertion,
                passed=not problems,
                expected=_structure_expectation(assertion),
                actual={
                    "rows": resolved.rows,
                    "cols": resolved.cols,
                    "is_merged": resolved.is_merged,
                    "merged_area": resolved.merged_area,
                },
                message=None if not problems else "; ".join(problems),
            )
        if kind is AssertionType.VALIDATION_EXISTS:
            rule = await service.get_validation(workbook_id, sheet, address)
            expected_present = _expected_present(assertion)
            return _assertion_result(
                assertion,
                passed=rule.present is expected_present,
                expected=expected_present,
                actual=rule.present,
                message=None
                if rule.present is expected_present
                else f"validation present={rule.present}, expected {expected_present}",
            )
        # CONDITIONAL_FORMAT_EXISTS
        rules = await service.get_conditional_formatting(workbook_id, sheet, address)
        supported = [rule for rule in rules if rule.supported]
        expected_count = assertion.value if assertion.value is not None else 1
        passed = len(supported) >= int(expected_count)
        return _assertion_result(
            assertion,
            passed=passed,
            expected={"min_supported_rules": int(expected_count)},
            actual={"supported_rules": len(supported), "rules": len(rules)},
            message=None
            if passed
            else f"supported conditional-format rules {len(supported)} < {int(expected_count)}",
        )

    async def _pivot_assertion(
        self,
        assertion: Assertion,
        kind: AssertionType,
        run: CaseRun,
        workbook_id: str,
    ) -> AssertionResult:
        """Evaluate one Pivot assertion (task_053), preferring real Excel reads.

        Layout / aggregation / filter assertions are read from Excel through the
        pivot service; ``pivot_table_exists`` prefers the case's own refresh
        outcome (which proves the table exists) and otherwise reads the inventory.
        """

        if self._pivots is None:
            return _assertion_result(
                assertion,
                passed=False,
                message=(
                    "pivot assertions require the pivot tables service, which this "
                    "runner was not given"
                ),
            )
        pivot_table = _require_str(assertion.pivot_table)
        if pivot_table is None:
            return _assertion_result(assertion, passed=False, message="needs 'pivot_table'")
        refresh_outcome = _last_pivot_outcome(
            self._pivot_outcomes, "pivot_table_refresh", pivot_table
        )
        if kind is AssertionType.PIVOT_REFRESH_STATUS:
            return _pivot_refresh_assertion(assertion, run, pivot_table, refresh_outcome)
        if kind is AssertionType.PIVOT_TABLE_EXISTS:
            return await self._pivot_table_exists_assertion(
                assertion, workbook_id, pivot_table, refresh_outcome
            )
        if kind is AssertionType.PIVOT_OUTPUT_RANGE:
            return await self._pivot_output_range_assertion(
                assertion, workbook_id, pivot_table, refresh_outcome
            )
        try:
            return await self._pivot_field_assertion(assertion, kind, workbook_id, pivot_table)
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, message=f"{exc.code.value}: {exc.message}"
            )

    async def _pivot_table_exists_assertion(
        self,
        assertion: Assertion,
        workbook_id: str,
        pivot_table: str,
        refresh_outcome: dict[str, Any] | None,
    ) -> AssertionResult:
        """``pivot_table_exists`` from the case's refresh outcome or the inventory."""

        if refresh_outcome is not None:
            return _assertion_result(
                assertion, passed=True, expected=True, actual=True, message=None
            )
        try:
            await self._pivots.get_pivot_table(workbook_id, pivot_table)  # type: ignore[union-attr]
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, expected=True, actual=False, message=exc.message
            )
        return _assertion_result(
            assertion,
            passed=True,
            expected=True,
            actual=True,
            message=None,
        )

    async def _pivot_output_range_assertion(
        self,
        assertion: Assertion,
        workbook_id: str,
        pivot_table: str,
        refresh_outcome: dict[str, Any] | None,
    ) -> AssertionResult:
        """``pivot_output_range`` from the case's refresh outcome or the inventory."""

        expected = (
            _require_str(assertion.output_range)
            or _require_str(assertion.expected)
            or _require_str(assertion.address)
        )
        if expected is None:
            return _assertion_result(
                assertion,
                passed=False,
                message="pivot_output_range requires 'output_range' (or 'expected')",
            )
        if refresh_outcome is not None:
            actual = _normalize_range(refresh_outcome.get("table_range1"))
            if actual is None:
                return _assertion_result(
                    assertion,
                    passed=False,
                    expected=expected,
                    actual=None,
                    message="the refresh outcome carried no output range",
                )
            passed = _same_range(actual, expected)
            return _assertion_result(
                assertion,
                passed=passed,
                expected=expected,
                actual=actual,
                message=None if passed else f"output range {actual!r} != {expected!r}",
            )
        try:
            info = await self._pivots.get_pivot_table(workbook_id, pivot_table)  # type: ignore[union-attr]
        except ExcelVbaMcpError as exc:
            return _assertion_result(
                assertion, passed=False, expected=expected, actual=None, message=exc.message
            )
        actual = _normalize_range(info.table_range1)
        passed = actual is not None and _same_range(actual, expected)
        return _assertion_result(
            assertion,
            passed=passed,
            expected=expected,
            actual=actual,
            message=None if passed else f"output range {actual!r} != {expected!r}",
        )

    async def _pivot_field_assertion(
        self,
        assertion: Assertion,
        kind: AssertionType,
        workbook_id: str,
        pivot_table: str,
    ) -> AssertionResult:
        """Field layout / aggregation / filter-state assertions (read-only)."""

        field_name = _require_str(assertion.field)
        if field_name is None:
            return _assertion_result(assertion, passed=False, message="needs 'field'")
        if kind is AssertionType.PIVOT_FILTER_STATE:
            filters = await self._pivots.get_pivot_field_filters(  # type: ignore[union-attr]
                workbook_id, pivot_table, field_name
            )
            return _pivot_filter_state_result(assertion, filters)
        fields = await self._pivots.list_pivot_fields(workbook_id, pivot_table)  # type: ignore[union-attr]
        match = _find_field(fields, field_name)
        if match is None:
            return _assertion_result(
                assertion,
                passed=False,
                expected=field_name,
                actual=[item.source_name for item in fields],
                message=f"field {field_name!r} not found in {pivot_table!r}",
            )
        if kind is AssertionType.PIVOT_FIELD_LAYOUT:
            return _pivot_field_layout_result(assertion, match)
        return _pivot_field_aggregation_result(assertion, match)

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

#: Phase 2 Range/Cells assertion kinds (task_053), evaluated read-only through
#: :class:`RangesService`.
_RANGE_ASSERTION_KINDS: frozenset[AssertionType] = frozenset(
    {
        AssertionType.RANGE_FORMULA_EQUALS,
        AssertionType.RANGE_FORMAT_EQUALS,
        AssertionType.RANGE_STRUCTURE,
        AssertionType.VALIDATION_EXISTS,
        AssertionType.CONDITIONAL_FORMAT_EXISTS,
    }
)

#: Phase 2 PivotTables assertion kinds (task_053).
_PIVOT_ASSERTION_KINDS: frozenset[AssertionType] = frozenset(
    {
        AssertionType.PIVOT_TABLE_EXISTS,
        AssertionType.PIVOT_FIELD_LAYOUT,
        AssertionType.PIVOT_FIELD_AGGREGATION,
        AssertionType.PIVOT_FILTER_STATE,
        AssertionType.PIVOT_REFRESH_STATUS,
        AssertionType.PIVOT_OUTPUT_RANGE,
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


# ---------------------------------------------------------------------------
# Phase 2 step / assertion helpers (task_053; no COM)
# ---------------------------------------------------------------------------
def _jsonable(value: Any) -> dict[str, Any]:
    """Serialize a service DTO into a JSON-friendly dict (no COM objects)."""

    dump = getattr(value, "model_dump", None)
    if callable(dump):
        return dict(dump(mode="json"))
    if isinstance(value, dict):
        return {str(key): item for key, item in value.items()}
    return {"value": value}


def _range_step_details(outcome: Any) -> dict[str, Any]:
    """Summarize a Range service result for a step's ``details``."""

    if hasattr(outcome, "model_dump"):
        data = outcome.model_dump(mode="json")
    elif isinstance(outcome, list):
        data = {"rules": [_jsonable(item) for item in outcome]}
    else:  # pragma: no cover - defensive
        data = {"value": outcome}
    summary: dict[str, Any] = {"service_result": data}
    if isinstance(outcome, list):
        summary["rule_count"] = len(outcome)
    elif hasattr(outcome, "verified"):
        summary["verified"] = bool(outcome.verified)
    if hasattr(outcome, "range_after") and outcome.range_after is not None:
        summary["address"] = outcome.range_after.address
    elif hasattr(outcome, "range") and outcome.range is not None:
        summary["address"] = outcome.range.address
    return summary


def _pivot_step_result(step: TestStep, action: TestAction, outcome: Any) -> TestStepResult:
    """Map one Pivot step's service result onto a step status (no COM).

    ``refresh`` truthfully reports ``unsupported_source`` / ``blocked`` as a
    blocked step and ``failed`` as a failed step; the other pivot mutations
    pass when Excel accepted them (the service already re-read the state as a
    postcondition and raises on a mismatch).
    """

    details = _range_step_details(outcome)
    if action is TestAction.PIVOT_TABLE_REFRESH:
        details = {"service_result": _jsonable(outcome)}
        status = getattr(outcome, "status", None)
        raw_status = getattr(status, "value", status)
        value = str(raw_status) if raw_status is not None else ""
        details["pivot_refresh_status"] = value
        details["refreshed"] = bool(getattr(outcome, "refreshed", False))
        if value == PivotRefreshStatus.REFRESHED.value:
            return _passed_step(step, details)
        if value == PivotRefreshStatus.FAILED.value:
            return _failed_step(
                step,
                value,
                getattr(outcome, "error", None) or "the pivot refresh did not complete",
                details=details,
            )
        return _blocked_step(
            step,
            value,
            f"pivot {getattr(outcome, 'pivot_table', '')!r} refresh {value}",
            details=details,
        )
    return _passed_step(step, details)


def _blocked_or_failed_service_step(
    step: TestStep, exc: ExcelVbaMcpError, *, blocked_codes: frozenset[ErrorCode]
) -> TestStepResult:
    """Map a service refusal onto a blocked (policy/guard) or failed step.

    The ranges / pivot services guard mutations deny-by-default, so a refused
    mutation is reported as a **blocked** step (the operation never ran) rather
    than a failure. Every other service error is a real failure.
    """

    if exc.code in blocked_codes:
        return _blocked_step(step, exc.code.value, exc.message)
    return _failed_step(step, exc.code.value, exc.message)


def _format_property(info: Any, path: str) -> Any:
    """Read a dotted property of a :class:`FormatInfo` (never a COM expression)."""

    current: Any = info
    for part in path.split("."):
        if isinstance(current, dict):
            if part not in current:
                return _MISSING_PROPERTY
            current = current[part]
            continue
        if not hasattr(current, part):
            return _MISSING_PROPERTY
        current = getattr(current, part)
    return current


def _structure_expectation(assertion: Assertion) -> dict[str, Any]:
    """Build the ``expected`` payload of a ``range_structure`` assertion."""

    expected: dict[str, Any] = {}
    if assertion.rows is not None:
        expected["rows"] = assertion.rows
    if assertion.cols is not None:
        expected["cols"] = assertion.cols
    if assertion.merged is not None:
        expected["merged"] = assertion.merged
    if assertion.expected is not None:
        expected["merged_area"] = assertion.expected
    return expected


def _structure_problems(info: Any, assertion: Assertion) -> list[str]:
    """Return the list of structure mismatches of a Range target (empty = match)."""

    problems: list[str] = []
    if assertion.rows is not None and info.rows != assertion.rows:
        problems.append(f"rows {info.rows} != {assertion.rows}")
    if assertion.cols is not None and info.cols != assertion.cols:
        problems.append(f"cols {info.cols} != {assertion.cols}")
    if assertion.merged is not None and bool(info.is_merged) is not bool(assertion.merged):
        problems.append(f"merged {info.is_merged} != {assertion.merged}")
    if assertion.expected is not None:
        expected_area = _normalize_range(assertion.expected)
        actual_area = _normalize_range(info.merged_area)
        if (
            expected_area is None
            or actual_area is None
            or not _same_range(actual_area, expected_area)
        ):
            problems.append(f"merged_area {actual_area!r} != {expected_area!r}")
    if not problems:
        return []
    if (
        assertion.rows is None
        and assertion.cols is None
        and assertion.merged is None
        and assertion.expected is None
    ):
        return ["range_structure requires 'rows', 'cols', 'merged' or 'expected'"]
    return problems


def _expected_present(assertion: Assertion) -> bool:
    """Resolve the expected ``present`` flag of a ``validation_exists`` assertion."""

    if isinstance(assertion.expected, bool):
        return assertion.expected
    if assertion.expected is not None:
        text = str(assertion.expected).strip().lower()
        if text in {"true", "present", "1", "yes"}:
            return True
        if text in {"false", "absent", "0", "no"}:
            return False
    return True


def _normalize_range(value: Any) -> str | None:
    """Normalize a range/address into a comparable ``A1`` form (``$`` removed)."""

    text = _require_str(value)
    if text is None:
        return None
    return text.replace("$", "").upper()


def _same_range(left: str, right: str) -> bool:
    """Compare two addresses after normalization (case / ``$`` insensitive)."""

    normalized_left = _normalize_range(left)
    normalized_right = _normalize_range(right)
    return (
        normalized_left is not None
        and normalized_right is not None
        and normalized_left == normalized_right
    )


def _last_pivot_outcome(
    outcomes: list[tuple[str, dict[str, Any]]], action: str, pivot_table: str
) -> dict[str, Any] | None:
    """Return the last outcome of one pivot action for ``pivot_table`` (if any)."""

    target = pivot_table.strip().lower()
    for name, payload in reversed(outcomes):
        if name != action:
            continue
        if str(payload.get("pivot_table", "")).strip().lower() == target:
            return dict(payload)
    return None


def _pivot_refresh_assertion(
    assertion: Assertion, run: CaseRun, pivot_table: str, outcome: dict[str, Any] | None
) -> AssertionResult:
    """``pivot_refresh_status`` against the case's refresh step (read-only)."""

    if outcome is None:
        return _assertion_result(
            assertion,
            passed=False,
            message=f"no pivot_table_refresh step ran for {pivot_table!r} in this case",
        )
    expected = _require_str(assertion.refresh_status) or _require_str(assertion.expected)
    if expected is None:
        expected = _DEFAULT_REFRESH_STATUS
    expected = expected.strip().lower()
    actual = str(outcome.get("status", ""))
    refreshed = outcome.get("refreshed")
    if assertion.refreshed is not None:
        passed = actual == expected and bool(refreshed) is bool(assertion.refreshed)
    else:
        passed = actual == expected
    return _assertion_result(
        assertion,
        passed=passed,
        expected=assertion.refreshed if assertion.refreshed is not None else expected,
        actual={"status": actual, "refreshed": refreshed},
        message=None
        if passed
        else f"refresh status {actual!r} (refreshed={refreshed!r}) != {expected!r}",
    )


def _pivot_filter_state_result(assertion: Assertion, filters: Any) -> AssertionResult:
    """``pivot_filter_state`` against a field's current filter state (read-only)."""

    actual: dict[str, Any] = {}
    if assertion.visible_items is not None:
        actual["visible_items"] = list(filters.visible_items)
    if assertion.hidden_items is not None:
        actual["hidden_items"] = list(filters.hidden_items)
    if assertion.expected is not None:
        actual["all_items_visible"] = filters.all_items_visible
    problems: list[str] = []
    visible = _string_list_or_none(assertion.visible_items)
    if visible is not None and sorted(filters.visible_items) != sorted(visible):
        problems.append(f"visible_items {list(filters.visible_items)!r} != {visible!r}")
    hidden = _string_list_or_none(assertion.hidden_items)
    if hidden is not None and sorted(filters.hidden_items) != sorted(hidden):
        problems.append(f"hidden_items {list(filters.hidden_items)!r} != {hidden!r}")
    if assertion.expected is not None and filters.all_items_visible is not _expected_present(
        assertion
    ):
        problems.append(
            f"all_items_visible {filters.all_items_visible!r} != {_expected_present(assertion)!r}"
        )
    if assertion.property is not None:
        actual_prop = _format_property(filters, assertion.property)
        if actual_prop is _MISSING_PROPERTY or not _values_equal(actual_prop, assertion.expected):
            problems.append(
                f"{assertion.property} "
                f"{'<missing>' if actual_prop is _MISSING_PROPERTY else repr(actual_prop)}"
                f" != {assertion.expected!r}"
            )
    if not problems and not actual:
        problems = [
            "pivot_filter_state requires 'visible_items', 'hidden_items', 'expected' or 'property'"
        ]
    return _assertion_result(
        assertion,
        passed=not problems,
        expected=(
            assertion.expected
            if assertion.expected is not None
            else {
                "visible_items": visible,
                "hidden_items": hidden,
            }
        ),
        actual=actual,
        message=None if not problems else "; ".join(problems),
    )


def _pivot_field_layout_result(assertion: Assertion, field: Any) -> AssertionResult:
    """``pivot_field_layout`` against a field's orientation/position (read-only)."""

    actual = {
        "orientation": field.orientation.value,
        "position": field.position,
    }
    problems: list[str] = []
    if assertion.orientation is not None and field.orientation is not assertion.orientation:
        problems.append(f"orientation {actual['orientation']!r} != {assertion.orientation.value!r}")
    if assertion.position is not None and field.position != assertion.position:
        problems.append(f"position {field.position!r} != {assertion.position!r}")
    if assertion.expected is not None and field.orientation.value != str(assertion.expected):
        problems.append(f"orientation {actual['orientation']!r} != {str(assertion.expected)!r}")
    if (
        not problems
        and assertion.orientation is None
        and assertion.position is None
        and assertion.expected is None
    ):
        problems = ["pivot_field_layout requires 'orientation', 'position' or 'expected'"]
    return _assertion_result(
        assertion,
        passed=not problems,
        expected=assertion.orientation.value
        if assertion.orientation is not None
        else assertion.expected,
        actual=actual,
        message=None if not problems else "; ".join(problems),
    )


def _pivot_field_aggregation_result(assertion: Assertion, field: Any) -> AssertionResult:
    """``pivot_field_aggregation`` against a data field's function (read-only)."""

    actual = {
        "function": field.function.value if field.function is not None else None,
        "caption": field.data_field_name,
    }
    problems: list[str] = []
    if assertion.function is not None and field.function is not assertion.function:
        problems.append(f"function {actual['function']!r} != {assertion.function.value!r}")
    if assertion.caption is not None and field.data_field_name != assertion.caption:
        problems.append(f"caption {field.data_field_name!r} != {assertion.caption!r}")
    if assertion.expected is not None and actual["function"] != str(assertion.expected):
        problems.append(f"function {actual['function']!r} != {str(assertion.expected)!r}")
    if (
        not problems
        and assertion.function is None
        and assertion.caption is None
        and assertion.expected is None
    ):
        problems = ["pivot_field_aggregation requires 'function', 'caption' or 'expected'"]
    return _assertion_result(
        assertion,
        passed=not problems,
        expected=assertion.function.value if assertion.function is not None else assertion.expected,
        actual=actual,
        message=None if not problems else "; ".join(problems),
    )


def _find_field(fields: list[Any], name: str) -> Any | None:
    """Resolve a PivotField by ``SourceName`` or data-field caption (casefolded)."""

    target = name.strip().casefold()
    for item in fields:
        if item.source_name.casefold() == target:
            return item
    for item in fields:
        if (item.data_field_name or "").casefold() == target:
            return item
    return None


def _string_list_or_none(value: Any) -> list[str] | None:
    """Coerce an expectation into a string list (``None`` when absent)."""

    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        return [str(item) for item in value]
    return [str(value)]


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
def _needs_recovery(run: CaseRun) -> bool:
    """Whether the case's last macro result requires worker recovery (B6-004)."""

    return run.macro_result is not None and run.macro_result.status is MacroRunStatus.NEEDS_RECOVERY


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
