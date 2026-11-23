"""Real-Excel verification of the Phase 2 declarative TestRunner (task_053).

This drives the production
:class:`~excel_vba_mcp.application.test_runner.TestRunner` against a real
automation-created **hidden** Excel instance (never the user's Excel) and a
disposable copy of the committed ``RangesCellsProject.xlsm`` /
``PivotTablesProject.xlsm`` fixtures. The runner executes the new Range/Pivot
actions through the **same** :class:`~excel_vba_mcp.application.ranges_service.
RangesService` / :class:`~excel_vba_mcp.application.pivot_tables_service.
PivotTablesService` the direct MCP tools use.

Covered:

* a suite with a ``range_write_values`` step plus a ``cell_value_equals``
  assertion ends ``passed`` (the value really landed in the disposable copy);
* a pivot suite with a ``pivot_table_refresh`` step plus a
  ``pivot_refresh_status`` assertion ends ``passed``;
* a case whose Range step is **not** approved (deny-by-default) is ``blocked``
  and never reaches Excel; a case whose pivot cache is unsupported would likewise
  block, never report success;
* every case starts from the same immutable baseline snapshot, one case's
  mutation cannot leak into the next, and the baseline / originals stay
  byte-identical;
* BG03/BG04/BG05 hold (the owned Excel is hidden + non-activating, user-owned
  Excel processes are untouched, the owned ``EXCEL.exe`` is reaped afterwards).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_test_runner_ranges_pivots.py -v -s
"""

from __future__ import annotations

import hashlib
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
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.ranges_service import RangesService
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

try:  # pragma: no cover - exercised implicitly by skip logic
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_RANGES_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "RangesCellsProject.xlsm"
_PIVOTS_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "PivotTablesProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5

#: The approval action a ``range_write_values`` step needs (RangesService guard).
_ACTION_WRITE_VALUES = "range_write_values"
#: The approval action a ``range_merge`` step needs (RangesService guard).
_ACTION_MERGE = "range_merge"
#: The approval action a ``pivot_table_refresh`` step needs (PivotTablesService).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"


# ---------------------------------------------------------------------------
# Skip helpers
# ---------------------------------------------------------------------------
def _pywin32_available() -> bool:
    return sys.platform == "win32" and win32gui is not None and win32process is not None


requires_excel = pytest.mark.skipif(
    not _pywin32_available(),
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


def _settled_excel_pids(stable_for_s: float = 1.0, timeout_s: float = 20.0) -> set[int]:
    """Return the user-owned Excel PID set once it stops churning.

    Other integration tests create and reap automation-owned Excel instances, so
    a naive ``baseline - current`` check races with those exits and misreports an
    automation PID as a terminated user process (BG05 false positive). Polling
    until the set is stable makes the baseline deterministic without weakening
    the assertion.
    """

    deadline = time.monotonic() + timeout_s
    previous = _excel_pids()
    stable_since = time.monotonic()
    while time.monotonic() < deadline:
        time.sleep(0.25)
        current = _excel_pids()
        if current != previous:
            previous = current
            stable_since = time.monotonic()
        elif time.monotonic() - stable_since >= stable_for_s:
            return current
    return previous


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
        self.ranges = RangesService(
            self.worker, self.session, self.snapshot, self.approval, self.path_policy, self.audit
        )
        self.pivots = PivotTablesService(
            self.worker,
            self.session,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
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
            ranges_service=self.ranges,
            pivot_tables_service=self.pivots,
        )
        self._original: Path | None = None
        self._baseline: Any = None

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
                granted_by="task053-verification",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_open(self) -> None:
        self._grant(ApprovalKey(canonical_path=self.canonical, action="excel_open"))

    def grant_mutation(self, action: str) -> None:
        """Approve one Range/Pivot mutation for every case of the run.

        Each case's working copy records the baseline file as ``source_path``, so
        the mutation approval key is the baseline canonical path (the same
        identity the direct Range/Pivot tools use).
        """

        self._grant(ApprovalKey(canonical_path=self.baseline_canonical, action=action))

    # -- lifecycle -------------------------------------------------------
    async def open(self) -> Any:
        self.grant_open()
        info = await self.session.open_workbook(
            str(self.original), mode=SessionMode.EXECUTE, read_only=False
        )
        # Pre-create the immutable baseline and approve opening it.
        self._baseline = await self.snapshot.create_snapshot(info.workbook_id)
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
        return info


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
            "suite": "phase2-runner-suite",
            "workbook_id": workbook_id,
            "baseline": "use_existing",
            "cases": cases,
        }
    )


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Range write -> cell_value_equals passes
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_range_write_then_cell_value_equals_passes(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_RANGES_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _settled_excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_mutation(_ACTION_WRITE_VALUES)

        suite = _suite(
            [
                {
                    "id": "write-then-assert",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "Cells",
                            "range_address": "B2",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "string", "value": "task053"}]],
                            },
                        }
                    ],
                    "assertions": [
                        {
                            "type": "cell_value_equals",
                            "sheet": "Cells",
                            "cell": "B2",
                            "expected": "task053",
                        },
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
        assert case.steps[0].status is StepStatus.PASSED
        assert case.steps[0].details.get("verified") is True
        assert all(assertion.passed for assertion in case.assertions), case.assertions

        # BG03 / BG04 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "runner-range")
        _assert_no_visible_owned_windows(owned_pid, "runner-range")

        # Originals + committed fixtures untouched; baseline file immutable.
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == rig.baseline.sha256

        _assert_owned_never_foreground(owned_pid, "runner-range-end")
    finally:
        await rig.stop()

    missing_user_pids = baseline_pids - _settled_excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


# ---------------------------------------------------------------------------
# Pivot refresh -> pivot_refresh_status passes
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_pivot_refresh_then_refresh_status_passes(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_PIVOTS_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _settled_excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_mutation(_ACTION_REFRESH_TABLE)

        suite = _suite(
            [
                {
                    "id": "refresh-then-assert",
                    "steps": [
                        {"action": "pivot_table_refresh", "pivot_table": "ptA"},
                    ],
                    "assertions": [
                        {
                            "type": "pivot_refresh_status",
                            "pivot_table": "ptA",
                            "expected": "refreshed",
                        },
                        {"type": "pivot_table_exists", "pivot_table": "ptA"},
                    ],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        case = result.cases[0]
        assert case.status is CaseStatus.PASSED
        assert case.steps[0].status is StepStatus.PASSED
        assert case.steps[0].details.get("pivot_refresh_status") == "refreshed"
        assert all(assertion.passed for assertion in case.assertions), case.assertions

        _assert_owned_never_foreground(owned_pid, "runner-pivot")
        _assert_no_visible_owned_windows(owned_pid, "runner-pivot")

        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == rig.baseline.sha256
    finally:
        await rig.stop()

    missing_user_pids = baseline_pids - _settled_excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


# ---------------------------------------------------------------------------
# A refused Range step blocks (no approval) and never reaches Excel
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_unapproved_range_step_blocks(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_RANGES_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        original_hash = _sha256(rig.original)
        # Deliberately no grant for range_write_values: deny-by-default.

        suite = _suite(
            [
                {
                    "id": "blocked-write",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "Cells",
                            "range_address": "B2",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "string", "value": "nope"}]],
                            },
                        },
                        {"action": "range_get", "sheet": "Cells", "range_address": "B2"},
                    ],
                    "assertions": [],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)
        case = result.cases[0]

        assert case.steps[0].status is StepStatus.SKIPPED
        assert case.steps[0].details["blocked"] is True
        assert case.status is CaseStatus.BLOCKED
        assert result.status is RunStatus.BLOCKED
        # The case stopped after the blocked step: the follow-up read never ran.
        assert len(case.steps) == 1
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# Isolation: each case starts from the same baseline, one case cannot leak
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_cases_start_from_the_same_baseline(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_RANGES_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        baseline_sha = rig.baseline.sha256
        rig.grant_mutation(_ACTION_WRITE_VALUES)

        suite = _suite(
            [
                {
                    "id": "mutate",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "Cells",
                            "range_address": "B2",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "string", "value": "changed"}]],
                            },
                        }
                    ],
                    "assertions": [
                        {
                            "type": "cell_value_equals",
                            "sheet": "Cells",
                            "cell": "B2",
                            "expected": "changed",
                        },
                    ],
                },
                {
                    "id": "verify-pristine",
                    "steps": [],
                    "assertions": [
                        {
                            "type": "cell_value_equals",
                            "sheet": "Cells",
                            "cell": "B2",
                            "expected": "hello world",
                        },
                    ],
                },
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        mutate_case, verify_case = result.cases
        assert mutate_case.status is CaseStatus.PASSED
        assert verify_case.status is CaseStatus.PASSED
        for case in result.cases:
            assert f"baseline_sha256={baseline_sha}" in " ".join(case.notes)
            assert f"working_copy_sha256={baseline_sha}" in " ".join(case.notes), case.notes
        assert mutate_case.working_copy_path != verify_case.working_copy_path

        # Both cases really started from the pristine baseline copy.
        assert _sha256(Path(rig.baseline.stored_path)) == baseline_sha, (
            "baseline must stay immutable"
        )
        open_ids = [record.workbook_id for record in rig.session.list_workbooks_sync()]
        assert open_ids == [info.workbook_id]
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# range_structure {merged: true/false} observes real merge geometry (F1)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_range_structure_merged_flag_matches_geometry(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_RANGES_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    try:
        info = await rig.open()
        rig.grant_mutation(_ACTION_MERGE)

        suite = _suite(
            [
                {
                    "id": "merge-then-structure",
                    "steps": [
                        {
                            "action": "range_merge",
                            "sheet": "StructOps",
                            "range_address": "P20:Q20",
                            "allow_data_loss": True,
                        }
                    ],
                    "assertions": [
                        {
                            "type": "range_structure",
                            "sheet": "StructOps",
                            "address": "P20:Q20",
                            "merged": True,
                        },
                        {
                            "type": "range_structure",
                            "sheet": "StructOps",
                            "address": "A1:B2",
                            "merged": False,
                        },
                    ],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        case = result.cases[0]
        assert case.status is CaseStatus.PASSED
        assert case.steps[0].status is StepStatus.PASSED
        assert all(assertion.passed for assertion in case.assertions), case.assertions
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# Orphan check: the owned EXCEL.exe is reaped after shutdown (BG05)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_owned_excel_process_is_reaped_after_shutdown(tmp_path: Path) -> None:
    fixture = _fixture_or_skip(_RANGES_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_mutation(_ACTION_WRITE_VALUES)

        suite = _suite(
            [
                {
                    "id": "smoke",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "Cells",
                            "range_address": "B2",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "number", "value": 5}]],
                            },
                        }
                    ],
                    "assertions": [
                        {
                            "type": "cell_value_equals",
                            "sheet": "Cells",
                            "cell": "B2",
                            "expected": 5,
                        },
                    ],
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
