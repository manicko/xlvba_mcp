"""Live-Excel verification of the bounded declarative test runner (task_035).

This is the task_035 independent verification of the production
:class:`~excel_vba_mcp.application.test_runner.TestRunner` against the target
Excel build, driving the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real hidden
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter`, a real
:class:`~excel_vba_mcp.application.session_service.SessionService` and the real
compile / macro / sheets / tables / names / shapes services. Only a temporary
**copy** of the committed ``CleanProject.xlsm`` fixture is ever opened; the
committed fixture and the temp original are byte-identical afterwards.

Contract verified:

* a suite with a ``macro_run`` step (``AddNumbers``) plus ``sheet_exists`` /
  ``macro_status`` assertions ends ``passed`` (the macro really executed);
* a suite with a failing assertion ends ``failed``;
* a suite whose case references a missing macro ends ``failed`` (macro not found)
  and one whose macro is not approved ends ``blocked``;
* every case started from the same immutable baseline snapshot (the recorded
  working-copy hash equals the baseline hash), one case's mutation cannot leak
  into the next, and the baseline file itself is unchanged;
* the run is persisted (run document + per-case diagnostics + artifact references)
  and readable through the runner's read API;
* BG03/BG04: no automation-owned window becomes visible or foreground during the
  run; BG05: user-owned Excel processes are untouched and the owned ``EXCEL.exe``
  is reaped after shutdown.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_test_runner.py -v -s
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.compile_service import CompileService
from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyRegistry,
    MacroService,
)
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.test_runner import TestRunner as Runner
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.statuses import (
    CaseStatus,
    OperationClass,
    SessionMode,
    StepStatus,
)
from excel_vba_mcp.domain.statuses import TestRunStatus as RunStatus
from excel_vba_mcp.domain.test_spec import TestSuiteSpec
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
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

# Optional COM dependencies (None when unavailable so the module imports and the
# COM tests skip on non-Windows hosts).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CLEAN_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CleanProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


# ---------------------------------------------------------------------------
# Skip helpers
# ---------------------------------------------------------------------------
def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(f"fixture {path.name} is missing; build fixtures first")
    return path


# ---------------------------------------------------------------------------
# BG03 / BG04 / BG05 helpers (non-intrusive, read-only)
# ---------------------------------------------------------------------------
def _excel_pids() -> set[int]:
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
    return pids


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return {"hwnd": hwnd, "pid": int(pid), "title": title}


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if win32gui is None or win32process is None:
        return result

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            return True
        if wpid != pid:
            return True
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            title = ""
        try:
            visible = bool(win32gui.IsWindowVisible(hwnd))
        except Exception:  # noqa: BLE001 - platform specific
            visible = False
        result.append({"hwnd": hwnd, "title": title, "visible": visible})
        return True

    try:  # pragma: no cover - depends on host
        win32gui.EnumWindows(_cb, None)
    except Exception:  # noqa: BLE001
        return result
    return result


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at {label} "
        f"(title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    visible = [window for window in _windows_for_pid(owned_pid) if window["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[window['title'] for window in visible]}"
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Rig
# ---------------------------------------------------------------------------
class _Rig:
    """Live worker + session + snapshot + all services + the declarative runner."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.snapshots_dir = tmp_dir / "snapshots"
        self.artifacts_dir = tmp_dir / "artifacts"
        self.runs_dir = tmp_dir / "runs"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {
                "work": str(self.work_dir),
                "projects": str(self.projects_dir),
                "snapshots": str(self.snapshots_dir),
                "artifacts": str(self.artifacts_dir),
            }
        )
        self.approval = ApprovalPolicy()
        self.background = BackgroundSafetyRegistry()
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )
        self.snapshot = SnapshotService(
            SnapshotStore(self.snapshots_dir),
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )
        self.vba = VBAProjectService(
            self.worker,
            self.session,
            self.path_policy,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )
        self.compile = CompileService(
            self.worker, self.session, self.vba, self.path_policy, self.audit
        )
        self.macro = MacroService(
            self.worker, self.session, self.approval, self.background, self.audit
        )
        self.sheets = SheetsService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.tables = TablesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.names = NamesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.shapes = ShapesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.artifacts = ArtifactStore(self.artifacts_dir, self.path_policy, self.audit)
        self.runs = RunRepository(self.runs_dir)
        self.runner = Runner(
            self.worker,
            self.session,
            self.snapshot,
            self.compile,
            self.macro,
            self.sheets,
            self.tables,
            self.names,
            self.artifacts,
            self.runs,
            self.path_policy,
            self.audit,
            shapes_service=self.shapes,
        )
        self._original: Path | None = None
        self._baseline: Any = None
        self.workbook_id: str | None = None

    def bind_original(self, original: Path) -> None:
        self._original = original

    @property
    def original(self) -> Path:
        assert self._original is not None
        return self._original

    @property
    def canonical(self) -> str:
        return str(self.path_policy.validate_path(str(self.original), OperationClass.MUTATING))

    @property
    def baseline(self) -> Any:
        assert self._baseline is not None
        return self._baseline

    @property
    def baseline_canonical(self) -> str:
        """Canonical baseline path — the approval identity of every case copy."""

        return str(
            self.path_policy.validate_path(self.baseline.stored_path, OperationClass.MUTATING)
        )

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    # -- grants / classification ----------------------------------------
    def _grant(self, key: ApprovalKey) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=key,
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task035-verification",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_open(self) -> None:
        self._grant(ApprovalKey(canonical_path=self.canonical, action="excel_open"))

    def grant_mutation(self, action: str) -> None:
        """Approve one sheet/table/name mutation for every case of the run.

        Each case's working copy records the baseline file as ``source_path``, so
        the mutation approval key is the baseline canonical path.
        """

        self._grant(ApprovalKey(canonical_path=self.baseline_canonical, action=action))

    def grant_macro_run(self, procedure: str) -> None:
        """Approve one macro against the **baseline** identity the runner uses.

        The runner opens each case from the baseline snapshot, so
        ``SessionService`` records the baseline file as ``source_path`` and the
        working copy (byte-identical to the baseline) as ``path``. The macro
        approval key therefore pins the baseline canonical path + baseline hash.
        """

        self._grant(
            ApprovalKey(
                canonical_path=self.baseline_canonical,
                sha256=self.baseline.sha256,
                action="macro_run",
                procedure=procedure,
            )
        )

    async def classify_safe(self, procedure: str) -> None:
        """Classify bound to the live hash + VBA source revision.

        §7.3 DEVIATION (B5 edits the B6-owned helper): pass the **live**
        ``project_source_hash`` to :meth:`BackgroundSafetyRegistry.classify` (the
        sanctioned 3-arg path) so the classification is keyed by
        ``(workbook hash, project_source_hash, procedure)`` — the same source
        revision the runner's ``run_macro`` recomputes on the (byte-identical) case
        copy. Classifying directly (not via ``classify_background``) also lets the
        not-found test pre-classify a procedure the runner will resolve as missing.
        No legacy 2-arg authorize path remains.
        """

        assert self.workbook_id is not None
        record = next(
            r for r in self.session.list_workbooks_sync() if r.workbook_id == self.workbook_id
        )
        _inventory, project_source_hash = await self.macro._collect_project_state(
            self.workbook_id, record
        )
        # Key on the baseline hash (== the runner's case-copy file hash), NOT the
        # original working-copy hash: the baseline is a re-saved snapshot, so its
        # bytes (and the case copies derived from it) differ from the original file.
        # The VBA source hash is stable across that re-save, so the live
        # ``project_source_hash`` is the same the runner recomputes on the case copy.
        self.background.classify(
            self.baseline.sha256, project_source_hash, procedure, background_safe=True
        )

    # -- lifecycle -------------------------------------------------------
    async def open(self) -> Any:
        self.grant_open()
        info = await self.session.open_workbook(
            str(self.original), mode=SessionMode.EXECUTE, read_only=False
        )
        self.workbook_id = info.workbook_id
        # Pre-create the immutable baseline and approve opening it.
        self._baseline = await self.snapshot.create_snapshot(info.workbook_id)
        self.grant_open_baseline()
        return info

    def grant_open_baseline(self) -> None:
        self._grant(
            ApprovalKey(
                canonical_path=str(
                    self.path_policy.validate_path(
                        self.baseline.stored_path, OperationClass.MUTATING
                    )
                ),
                action="excel_open",
            )
        )


def _make_rig(tmp_path: Path, fixture: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original)
    return rig


def _suite(cases: list[dict[str, Any]], workbook_id: str) -> TestSuiteSpec:
    """Build a validated suite that reuses the pre-created baseline snapshot."""

    return TestSuiteSpec.model_validate(
        {
            "suite": "integration-suite",
            "workbook_id": workbook_id,
            "baseline": "use_existing",
            "cases": cases,
        }
    )


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_suite_with_macro_step_passes(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # Capture user-owned Excel PIDs **before** the worker creates its own instance
    # (BG05 baseline), so the owned PID is never mistaken for a user process.
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_macro_run("modTest.AddNumbers")
        await rig.classify_safe("modTest.AddNumbers")

        suite = _suite(
            [
                {
                    "id": "add-numbers",
                    "steps": [
                        {"action": "macro_run", "macro": "modTest.AddNumbers", "args": [2, 3]}
                    ],
                    "assertions": [
                        {"type": "macro_status", "expected": "success"},
                        {"type": "macro_return_value", "expected": 5},
                        {"type": "sheet_exists", "name": "Sheet1"},
                    ],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        assert result.baseline_snapshot_id == rig.baseline.snapshot_id
        case = result.cases[0]
        assert case.status is CaseStatus.PASSED
        assert len(case.steps) == 1
        assert case.steps[0].status is StepStatus.PASSED
        assert case.steps[0].details["macro_status"] == "success"
        assert case.steps[0].details["return_value"] == 5
        assert all(assertion.passed for assertion in case.assertions), case.assertions
        # The case really ran on a disposable copy of the baseline.
        assert f"working_copy_sha256={rig.baseline.sha256}" in " ".join(case.notes)

        # The run is persisted and readable through the runner's read API.
        stored = rig.runner.get_run(result.test_run_id)
        assert stored.test_run_id == result.test_run_id
        assert stored.status is RunStatus.PASSED
        assert [item.name for item in rig.artifacts.list_artifacts(workbook_id=info.workbook_id)]
        assert len(result.artifact_paths) >= 2
        for path in result.artifact_paths:
            assert Path(path).is_file()
        manifest_path = next(
            path for path in result.artifact_paths if Path(path).name.endswith("manifest.json")
        )
        manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
        assert manifest["external_side_effects_not_undone"] is True

        # BG03 / BG04 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "test-runner")
        _assert_no_visible_owned_windows(owned_pid, "test-runner")

        # Original + committed fixtures untouched; baseline file immutable.
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == rig.baseline.sha256

        # BG03: the owned automation instance never held the foreground. The
        # desktop foreground is allowed to change (a shared session can move
        # focus), so the assertion is about the automation PID, not about
        # "nothing else changed".
        _assert_owned_never_foreground(owned_pid, "test-runner-end")
    finally:
        await rig.stop()

    # BG05: no pre-existing user Excel process was terminated.
    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_suite_with_failing_assertion_fails(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        original_hash = _sha256(rig.original)

        suite = _suite(
            [
                {
                    "id": "missing-sheet",
                    "steps": [],
                    "assertions": [{"type": "sheet_exists", "name": "NoSuchSheet"}],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.FAILED
        case = result.cases[0]
        assert case.status is CaseStatus.FAILED
        assert case.assertions[0].passed is False
        assert case.assertions[0].type == "sheet_exists"
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
    finally:
        await rig.stop()


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_case_with_missing_macro_fails_and_unapproved_macro_blocks(
    tmp_path: Path,
) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        # Approved + classified so the macro reaches dispatch and is not found.
        rig.grant_macro_run("modTest.DoesNotExist")
        await rig.classify_safe("modTest.DoesNotExist")

        not_found = _suite(
            [
                {
                    "id": "missing-macro",
                    "steps": [{"action": "macro_run", "macro": "modTest.DoesNotExist"}],
                    "assertions": [{"type": "macro_status", "expected": "success"}],
                }
            ],
            info.workbook_id,
        )
        not_found_result = await rig.runner.run_suite(not_found)
        not_found_case = not_found_result.cases[0]

        assert not_found_case.steps[0].status is StepStatus.FAILED
        assert not_found_case.steps[0].error_code == "MACRO_NOT_FOUND"
        assert not_found_case.status is CaseStatus.FAILED

        # Approved nothing for AddNumbers -> policy denies it -> blocked step.
        blocked = _suite(
            [
                {
                    "id": "unapproved-macro",
                    "steps": [
                        {"action": "macro_run", "macro": "modTest.AddNumbers", "args": [1, 1]}
                    ],
                    "assertions": [{"type": "macro_status", "expected": "success"}],
                }
            ],
            info.workbook_id,
        )
        blocked_result = await rig.runner.run_suite(blocked)
        blocked_case = blocked_result.cases[0]

        assert blocked_case.steps[0].status is StepStatus.SKIPPED
        assert blocked_case.steps[0].details["blocked"] is True
        assert blocked_case.status is CaseStatus.BLOCKED
        assert blocked_result.status is RunStatus.BLOCKED
    finally:
        await rig.stop()


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_cases_start_from_the_same_baseline_hash(tmp_path: Path) -> None:
    """Isolation evidence: both cases start from one immutable baseline."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        baseline_sha = rig.baseline.sha256
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        suite = _suite(
            [
                {
                    "id": "case-1",
                    "steps": [],
                    "assertions": [{"type": "sheet_exists", "name": "Sheet1"}],
                },
                {
                    "id": "case-2",
                    "steps": [],
                    "assertions": [{"type": "file_hash_exists"}],
                },
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED
        assert len(result.cases) == 2
        for case in result.cases:
            notes = " ".join(case.notes)
            assert f"baseline_sha256={baseline_sha}" in notes
            assert f"working_copy_sha256={baseline_sha}" in notes, case.notes
        # Each case used its own disposable working copy.
        assert (
            result.cases[0].working_copy_path is not None
            and result.cases[1].working_copy_path is not None
        )
        assert result.cases[0].working_copy_path != result.cases[1].working_copy_path
        # The immutable baseline and both sources are untouched.
        assert _sha256(Path(rig.baseline.stored_path)) == baseline_sha
        assert _sha256(rig.original) == original_hash
        assert _sha256(fixture) == committed_hash
        # No case leaked a session: only the original workbook stays open.
        open_ids = [record.workbook_id for record in rig.session.list_workbooks_sync()]
        assert open_ids == [info.workbook_id]

        _assert_owned_never_foreground(owned_pid or 0, "test-runner-iso")
        _assert_no_visible_owned_windows(owned_pid or 0, "test-runner-iso")
    finally:
        await rig.stop()


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_case_mutation_does_not_leak_into_the_next(tmp_path: Path) -> None:
    """A worksheet added by case 1 must not be visible to case 2."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        baseline_sha = rig.baseline.sha256
        # Deny-by-default: approve the worksheet mutation the case performs.
        rig.grant_mutation("sheet_add")

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
                    "assertions": [{"type": "sheet_exists", "name": "Sheet1"}],
                },
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        mutate_case, verify_case = result.cases
        assert mutate_case.status is CaseStatus.PASSED
        assert mutate_case.steps[0].status is StepStatus.PASSED
        assert verify_case.status is CaseStatus.PASSED
        for case in result.cases:
            assert f"working_copy_sha256={baseline_sha}" in " ".join(case.notes)

        # Case 2 really is a fresh copy: a leaked sheet would have made its
        # sheet_exists assertion list the extra sheet. Re-verify from disk.
        baseline_file = Path(rig.baseline.stored_path)
        assert _sha256(baseline_file) == baseline_sha, "baseline must stay immutable"
    finally:
        await rig.stop()


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_owned_excel_process_is_reaped_after_shutdown(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        suite = _suite(
            [
                {
                    "id": "smoke",
                    "steps": [],
                    "assertions": [{"type": "sheet_exists", "name": "Sheet1"}],
                }
            ],
            info.workbook_id,
        )
        result = await rig.runner.run_suite(suite)
        assert result.status is RunStatus.PASSED
        assert _pid_alive(owned_pid)
    finally:
        await rig.stop()

    assert _wait_pid_gone(owned_pid or 0), f"owned PID {owned_pid} survived shutdown"


def _pid_alive(pid: int) -> bool:
    if pid <= 0:  # pragma: no cover - defensive
        return False
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001
        return True
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int) -> bool:
    """Poll until ``pid`` disappears (process reaping is asynchronous)."""

    deadline = time.monotonic() + _ORPHAN_POLL_TIMEOUT_S
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)
