"""End-to-end live-Excel verification of the declarative test runner (task_036).

This is the task_036 independent, end-to-end verification of the production
:class:`~excel_vba_mcp.application.test_runner.TestRunner` against the target
Excel build. It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real hidden
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter`, the real
:class:`~excel_vba_mcp.application.session_service.SessionService` and the real
compile / macro / sheets / tables / names / shapes services. Only temporary
**copies** of the committed fixtures are ever opened; the committed fixtures and
the temp originals are byte-identical afterwards.

Unlike ``test_test_runner.py`` (task_035), the assertions here are written
independently and cover the end-to-end workflows end to end:

* ``test_successful_suite_returns_passed_with_artifacts`` — a ``macro_run``
  (``AddNumbers``) plus ``macro_status`` / ``macro_return_value`` assertions ends
  ``passed``; the run manifest + per-case diagnostic are persisted and readable
  back through ``get_run``;
* ``test_compile_error_blocks_macro_step`` — the real compile probe on a
  ``CompileErrorProject.xlsm`` copy fails closed (``compile_error`` /
  ``indeterminate``, never a false ``compiled``) and captures diagnostics; a
  deterministic ``compile_error`` result then blocks the following ``macro_run``
  (the macro service is **not** called);
* ``test_runtime_error_represented_with_evidence`` — ``RaiseError`` ends
  ``MACRO_RUNTIME_ERROR`` with persisted diagnostic evidence and no false
  success;
* ``test_cases_are_isolated_from_each_other`` — a case that mutates a sheet and a
  case that asserts the baseline value; each case's working copy hash equals the
  baseline hash, the immutable baseline itself stays clean and the original
  fixtures are untouched;
* ``test_original_fixtures_unchanged`` — the SHA-256 of ``CleanProject.xlsm`` /
  ``CompileErrorProject.xlsm`` is unchanged after a suite;
* ``test_run_resource_retrieval`` — the ``excel://test-runs/{run_id}`` resource
  path (``test_get_run`` -> ``RunRepository``) returns the persisted run.

Compile-build note (BG03/BG04/BG07/BG09): on the validated build the production
``CompileService`` never compiles and never returns ``compile_error``; it fails
closed to ``indeterminate``. The compile *gate* is therefore verified twice: (1)
against the real service on ``CompileErrorProject`` (fail-closed + diagnostics),
and (2) against a deterministic ``compile_error`` test double that proves the
runner blocks a later ``macro_run`` and never dispatches the macro.

Table note: the plan's ``TableNamesProject.xlsm`` fixture does not exist in the
repository. Rather than fabricate a table inside a macro-enabled fixture (which
would mutate a copy in a way unrelated to the runner), the table-specific case is
intentionally not part of this end-to-end suite; the runner's table dispatch is
covered by the task_035 runner tests. This is a deliberate, documented omission.

Safety: hidden Excel, disposable copies, per-case snapshots, no ``.Activate`` /
``.Select`` from the tests, temp dirs only, no arbitrary execution, and never
killing a process the rig does not own (BG03/BG04/BG05). Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_end_to_end_test_runner.py -v -s
"""

from __future__ import annotations

import asyncio
import contextlib
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
from excel_vba_mcp.adapters.windows_ui.vbe_compile_ui import CommandDiscovery
from excel_vba_mcp.api import test_tools
from excel_vba_mcp.application import test_runner as runner_module
from excel_vba_mcp.application.compile_service import (
    CompileDiagnostics,
    CompileResult,
    CompileService,
)
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
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.models import MacroRunResult
from excel_vba_mcp.domain.statuses import (
    CaseStatus,
    CompileStatus,
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
_COMPILE_ERROR_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CompileErrorProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
#: Short macro soft ceiling so the RaiseError modal watchdog fires quickly instead
#: of waiting the runner's 120s default (monkeypatched only for that test).
_MODAL_MACRO_TIMEOUT_S = 15.0

_ADD_NUMBERS = "modTest.AddNumbers"
_RAISE_ERROR = "modTest.RaiseError"


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


def _terminate_owned_pid(pid: int | None) -> None:
    """Defensive last-resort cleanup for a rig-owned PID (never a user PID)."""

    if pid is None or pid <= 0:
        return
    # pragma: no cover - depends on host
    with contextlib.suppress(Exception):
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, timeout=60)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


# ---------------------------------------------------------------------------
# Test doubles (compile gate + macro spy) — never used for the success path
# ---------------------------------------------------------------------------
class _CompileErrorStub:
    """Deterministic compile service that always reports ``compile_error``.

    The production :class:`CompileService` fails closed to ``indeterminate`` on
    this build and therefore never sets the runner's compile gate. This double
    exists solely to exercise the runner's ``compile_error`` branch (a failed
    ``vba_compile`` must block a later ``macro_run``); the real service is
    verified separately against ``CompileErrorProject`` in the same test.
    """

    def __init__(self) -> None:
        self.calls: list[str] = []

    async def compile_project(self, workbook_id: str, *, timeout_s: float = 60.0) -> CompileResult:
        self.calls.append(workbook_id)
        now = datetime.now(UTC)
        return CompileResult(
            workbook_id=workbook_id,
            status=CompileStatus.COMPILE_ERROR,
            reason="STUB_COMPILE_ERROR",
            diagnostics=CompileDiagnostics(
                is_compile=False,
                component_count=1,
                notes=["stub: deterministic compile error for compile-gate verification"],
            ),
            command_discovery=CommandDiscovery(attempted=False, notes=["stub compile service"]),
            notes=["stub compile result"],
            created_at=now,
            updated_at=now,
        )


class _SpyMacroService(MacroService):
    """Macro service that records ``run_macro`` invocations (delegates to super)."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self.run_calls: list[tuple[str, str]] = []

    async def run_macro(
        self,
        workbook_id: str,
        macro_name: str,
        args: list[Any] | None = None,
        *,
        timeout_s: float = 120.0,
    ) -> MacroRunResult:
        self.run_calls.append((workbook_id, macro_name))
        return await super().run_macro(workbook_id, macro_name, args, timeout_s=timeout_s)


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
        self.runner = self.make_runner(self.compile, self.macro)
        self._original: Path | None = None
        self._baseline: Any = None
        self.workbook_id: str | None = None

    # -- construction helpers -------------------------------------------
    def make_runner(self, compile_service: Any, macro_service: Any) -> Any:
        """Build a runner over the rig's services, swapping compile / macro."""

        return runner_module.TestRunner(
            self.worker,
            self.session,
            self.snapshot,
            compile_service,
            macro_service,
            self.sheets,
            self.tables,
            self.names,
            self.artifacts,
            self.runs,
            self.path_policy,
            self.audit,
            shapes_service=self.shapes,
        )

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

    async def shutdown_and_reap(self, owned_pid: int | None) -> None:
        """Stop the rig and reap its owned Excel so it cannot leak into the next test.

        ``Excel.App.Quit`` returns before the OS process has fully exited, so a
        short poll is required; otherwise a lingering owned ``EXCEL.exe`` would be
        captured as a "user" PID by the next test's BG05 baseline and reported as
        a disappeared user process.
        """

        await self.stop()
        if owned_pid is not None:
            _wait_pid_gone(owned_pid)

    # -- grants / classification ----------------------------------------
    def _grant(self, key: ApprovalKey) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=key,
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="task036-verification",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_open(self) -> None:
        self._grant(ApprovalKey(canonical_path=self.canonical, action="excel_open"))

    def grant_mutation(self, action: str) -> None:
        """Approve one sheet/table/name mutation for every case of the run."""

        self._grant(ApprovalKey(canonical_path=self.baseline_canonical, action=action))

    def grant_macro_run(self, procedure: str) -> None:
        """Approve one macro against the **baseline** identity the runner uses."""

        self._grant(
            ApprovalKey(
                canonical_path=self.baseline_canonical,
                sha256=self.baseline.sha256,
                action="macro_run",
                procedure=procedure,
            )
        )

    def grant_macro_run_unbound(self, procedure: str) -> None:
        """Approve one macro for any file hash under the baseline identity.

        A mutation before the macro changes the case copy's hash, so the
        hash-bound grant would deny it; the B6 ``background_safe`` declaration is
        re-recorded by the runner against the live case-copy hash instead.
        """

        self._grant(
            ApprovalKey(
                canonical_path=self.baseline_canonical,
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
        copy. No legacy 2-arg authorize path remains.
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

    async def open_copy(self, fixture: Path) -> Any:
        """Copy ``fixture`` into the temp projects root, approve and open it."""

        dest = self.projects_dir / f"extra-{fixture.name}"
        shutil.copy2(fixture, dest)
        canonical = str(self.path_policy.validate_path(str(dest), OperationClass.MUTATING))
        self._grant(ApprovalKey(canonical_path=canonical, action="excel_open"))
        return await self.session.open_workbook(
            str(dest), mode=SessionMode.EXECUTE, read_only=False
        )


def _make_rig(tmp_path: Path, fixture: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original)
    return rig


def _suite(cases: list[dict[str, Any]], workbook_id: str, name: str = "e2e-suite") -> TestSuiteSpec:
    """Build a validated suite that reuses the pre-created baseline snapshot."""

    return TestSuiteSpec.model_validate(
        {
            "suite": name,
            "workbook_id": workbook_id,
            "baseline": "use_existing",
            "cases": cases,
        }
    )


async def _run_suite_sampling_foreground(
    rig: _Rig, suite: TestSuiteSpec, owned_pid: int | None
) -> tuple[Any, list[dict[str, Any]]]:
    """Run a suite while sampling the foreground window (BG03 evidence)."""

    task = asyncio.ensure_future(rig.runner.run_suite(suite))
    owned_samples: list[dict[str, Any]] = []
    while not task.done():
        fg = _foreground()
        if fg is not None and owned_pid is not None and fg["pid"] == owned_pid:
            owned_samples.append(fg)
        await asyncio.sleep(0.1)
    return await task, owned_samples


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_successful_suite_returns_passed_with_artifacts(tmp_path: Path) -> None:
    """A macro suite passes end to end and persists its evidence."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    # BG05 baseline: user-owned Excel PIDs captured before the rig starts.
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_macro_run(_ADD_NUMBERS)
        await rig.classify_safe(_ADD_NUMBERS)

        suite = _suite(
            [
                {
                    "id": "add-numbers",
                    "steps": [{"action": "macro_run", "macro": _ADD_NUMBERS, "args": [2, 3]}],
                    "assertions": [
                        {"type": "macro_status", "expected": "success"},
                        {"type": "macro_return_value", "expected": 5},
                    ],
                }
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        assert result.workbook_id == info.workbook_id
        assert result.baseline_snapshot_id == rig.baseline.snapshot_id
        case = result.cases[0]
        assert case.status is CaseStatus.PASSED
        assert len(case.steps) == 1
        assert case.steps[0].status is StepStatus.PASSED
        assert case.steps[0].details["macro_status"] == "success"
        assert case.steps[0].details["return_value"] == 5
        assert all(assertion.passed for assertion in case.assertions), case.assertions
        # The macro really executed on a disposable copy of the baseline.
        assert f"working_copy_sha256={rig.baseline.sha256}" in " ".join(case.notes)

        # Evidence persisted: manifest + per-case diagnostic + run document.
        assert len(result.artifact_paths) >= 2
        for path in result.artifact_paths:
            assert Path(path).is_file(), path
        stored_artifacts = rig.artifacts.list_artifacts(workbook_id=info.workbook_id)
        assert stored_artifacts, "no artifacts indexed for the run's workbook"
        manifest_path = next(
            path for path in result.artifact_paths if Path(path).name.endswith("manifest.json")
        )
        manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
        assert manifest["external_side_effects_not_undone"] is True
        assert manifest["test_run_id"] == result.test_run_id

        # Readable back through the runner read API (same identity + status).
        stored = rig.runner.get_run(result.test_run_id)
        assert stored.test_run_id == result.test_run_id
        assert stored.status is RunStatus.PASSED
        assert [c.case_id for c in stored.cases] == ["add-numbers"]

        # BG03 / BG04 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "e2e-success")
        _assert_no_visible_owned_windows(owned_pid, "e2e-success")

        # Original + committed fixture untouched; baseline file immutable.
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == rig.baseline.sha256
    finally:
        await rig.shutdown_and_reap(owned_pid)

    # BG05: no pre-existing user Excel process was terminated.
    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_sequential_mutation_macro_workflow_uses_background_safe(
    tmp_path: Path,
) -> None:
    """End to end: edit -> macro -> edit -> macro via the background_safe field.

    The declarative ``background_safe`` declaration lets the runner classify each
    macro against the live case-copy hash, so the workflow survives the file hash
    change caused by the mutation between macros. The original stays byte-identical.
    """

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_mutation("sheet_add")
        rig.grant_macro_run_unbound(_ADD_NUMBERS)

        suite = _suite(
            [
                {
                    "id": "edit-macro-edit-macro",
                    "steps": [
                        {"action": "sheet_add", "sheet": "B6First"},
                        {
                            "action": "macro_run",
                            "macro": _ADD_NUMBERS,
                            "args": [2, 3],
                            "background_safe": True,
                        },
                        {"action": "sheet_add", "sheet": "B6Second"},
                        {
                            "action": "macro_run",
                            "macro": _ADD_NUMBERS,
                            "args": [4, 5],
                            "background_safe": True,
                        },
                    ],
                    "assertions": [
                        {"type": "sheet_exists", "name": "B6First"},
                        {"type": "sheet_exists", "name": "B6Second"},
                        {"type": "macro_status", "expected": "success"},
                        {"type": "macro_return_value", "expected": 9},
                    ],
                }
            ],
            info.workbook_id,
            name="e2e-background-safe",
        )

        result, owned_samples = await _run_suite_sampling_foreground(rig, suite, owned_pid)

        assert result.status is RunStatus.PASSED, f"run failed: {result}"
        case = result.cases[0]
        assert case.status is CaseStatus.PASSED
        assert [step.status for step in case.steps] == [StepStatus.PASSED] * 4
        assert all(assertion.passed for assertion in case.assertions), case.assertions
        assert not owned_samples, f"owned automation held foreground: {owned_samples}"

        # Evidence persisted and readable; the original + baseline are untouched.
        assert len(result.artifact_paths) >= 2
        for path in result.artifact_paths:
            assert Path(path).is_file(), path
        assert rig.runner.get_run(result.test_run_id).status is RunStatus.PASSED
        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == rig.baseline.sha256

        _assert_owned_never_foreground(owned_pid, "e2e-background-safe")
        _assert_no_visible_owned_windows(owned_pid, "e2e-background-safe")
    finally:
        await rig.shutdown_and_reap(owned_pid)

    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_compile_error_blocks_macro_step(tmp_path: Path) -> None:
    """A failed compile blocks the macro step; the macro service is not called."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    compile_error_fixture = _fixture_or_skip(_COMPILE_ERROR_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_clean = _sha256(fixture)
        committed_bad = _sha256(compile_error_fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # -- Part A: the REAL compile probe on CompileErrorProject fails closed.
        ce_info = await rig.open_copy(compile_error_fixture)
        try:
            real = await rig.compile.compile_project(ce_info.workbook_id)
            # Fail-closed: never a false "compiled"; indeterminate is expected on
            # this build, compile_error is accepted as the strongest proof.
            assert real.status in (CompileStatus.COMPILE_ERROR, CompileStatus.INDETERMINATE)
            assert real.status is not CompileStatus.COMPILED
            assert real.diagnostics.is_compile is False
            assert real.diagnostics.notes, "compile diagnostics must carry notes"
            assert real.reason, "a structured reason must be recorded"
            if real.diagnostics.vbom_available:
                assert real.diagnostics.component_count >= 1, real.diagnostics
        finally:
            await rig.session.close_workbook(ce_info.workbook_id, save_changes=False)

        _assert_owned_never_foreground(owned_pid, "e2e-compile-a")
        _assert_no_visible_owned_windows(owned_pid, "e2e-compile-a")

        # -- Part B: a deterministic compile_error must block the macro step.
        stub_compile = _CompileErrorStub()
        spy_macro = _SpyMacroService(
            rig.worker, rig.session, rig.approval, rig.background, rig.audit
        )
        stub_runner = rig.make_runner(stub_compile, spy_macro)
        # Approve + classify so the block is attributable to the compile gate
        # only, never to a missing macro approval.
        rig.grant_macro_run(_ADD_NUMBERS)
        await rig.classify_safe(_ADD_NUMBERS)

        suite = _suite(
            [
                {
                    "id": "compile-gate",
                    "steps": [
                        {"action": "vba_compile"},
                        {"action": "macro_run", "macro": _ADD_NUMBERS, "args": [2, 3]},
                    ],
                    "assertions": [],
                }
            ],
            info.workbook_id,
        )

        result = await stub_runner.run_suite(suite)

        assert len(stub_compile.calls) == 1, stub_compile.calls
        compile_step, macro_step = result.cases[0].steps
        assert compile_step.action == "vba_compile"
        assert compile_step.status is StepStatus.FAILED
        assert compile_step.error_code == "COMPILE_FAILED"
        assert compile_step.details["compile_status"] == "compile_error"
        assert macro_step.action == "macro_run"
        assert macro_step.status is StepStatus.SKIPPED
        assert macro_step.details["blocked"] is True
        assert macro_step.error_code == "MACROS_BLOCKED_BY_COMPILE"
        # The decisive evidence: the macro service was never invoked.
        assert spy_macro.run_calls == [], spy_macro.run_calls
        assert result.cases[0].status is CaseStatus.FAILED
        assert result.status is RunStatus.FAILED

        _assert_owned_never_foreground(owned_pid, "e2e-compile-b")
        _assert_no_visible_owned_windows(owned_pid, "e2e-compile-b")

        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_clean, "committed CleanProject must never change"
        assert _sha256(compile_error_fixture) == committed_bad, "committed CompileErrorProject"
    finally:
        await rig.shutdown_and_reap(owned_pid)

    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_runtime_error_represented_with_evidence(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """RaiseError wedges the STA past the soft deadline -> SUSPECTED_STALLED evidence."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    # Short macro ceiling so the modal watchdog fires quickly (test-only patch of
    # the runner's module constant; the production value is left untouched).
    monkeypatch.setattr(runner_module, "_MACRO_TIMEOUT_S", _MODAL_MACRO_TIMEOUT_S)

    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_macro_run(_RAISE_ERROR)
        await rig.classify_safe(_RAISE_ERROR)

        suite = _suite(
            [
                {
                    "id": "raise-error",
                    "steps": [{"action": "macro_run", "macro": _RAISE_ERROR}],
                    # Expecting success here is deliberately wrong: the assertion
                    # must fail, proving the runtime error is not a false success.
                    "assertions": [{"type": "macro_status", "expected": "success"}],
                }
            ],
            info.workbook_id,
        )

        result, owned_samples = await _run_suite_sampling_foreground(rig, suite, owned_pid)

        # No false success: the run failed, the macro step failed with evidence.
        assert result.status is RunStatus.FAILED
        case = result.cases[0]
        assert case.status is CaseStatus.FAILED
        macro_step = case.steps[0]
        assert macro_step.action == "macro_run"
        assert macro_step.status is StepStatus.FAILED
        assert macro_step.error_code == "MACRO_TIMEOUT"
        assert macro_step.details["macro_status"] == "suspected_stalled"
        assert case.assertions[0].passed is False
        assert case.assertions[0].actual == "suspected_stalled"

        # Diagnostic evidence persisted: the per-case document records the failure.
        assert result.artifact_paths
        case_doc = next(
            path
            for path in result.artifact_paths
            if Path(path).name.startswith(result.test_run_id) and "raise-error" in Path(path).name
        )
        diagnostic = json.loads(Path(case_doc).read_text(encoding="utf-8"))
        assert diagnostic["status"] == "failed"
        assert diagnostic["steps"][0]["error_code"] == "MACRO_TIMEOUT"
        assert diagnostic["steps"][0]["details"]["macro_status"] == "suspected_stalled"
        manifest_path = next(
            path for path in result.artifact_paths if Path(path).name.endswith("manifest.json")
        )
        manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
        assert manifest["external_side_effects_not_undone"] is True

        # BG03: the owned automation instance never held the foreground, even
        # across the runtime-error modal.
        assert not owned_samples, f"BG03 violation: owned PID held foreground: {owned_samples}"

        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
    finally:
        # Block E leaves the owned PID alive (the modal wedges the STA); terminate it
        # defensively so a failed assertion can never leave an orphaned EXCEL.exe.
        if owned_pid is not None and _pid_alive(owned_pid):
            _terminate_owned_pid(owned_pid)
        await rig.shutdown_and_reap(owned_pid)

    assert owned_pid is not None and not _pid_alive(owned_pid), (
        f"owned PID {owned_pid} survived shutdown"
    )
    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_cases_are_isolated_from_each_other(tmp_path: Path) -> None:
    """A mutating case cannot leak into a later case; each copy matches the baseline."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(fixture)
        original_hash = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        baseline_sha = rig.baseline.sha256
        rig.grant_mutation("sheet_add")

        suite = _suite(
            [
                {
                    "id": "mutate",
                    "steps": [{"action": "sheet_add", "sheet": "Mutation"}],
                    "assertions": [{"type": "sheet_exists", "name": "Mutation"}],
                },
                {
                    "id": "baseline",
                    "steps": [],
                    "assertions": [
                        {"type": "sheet_exists", "name": "Sheet1"},
                        {"type": "file_hash_exists"},
                    ],
                },
            ],
            info.workbook_id,
        )

        result = await rig.runner.run_suite(suite)

        assert result.status is RunStatus.PASSED
        mutate_case, baseline_case = result.cases
        assert mutate_case.status is CaseStatus.PASSED
        assert mutate_case.steps[0].status is StepStatus.PASSED
        assert baseline_case.status is CaseStatus.PASSED
        # Both cases ran on their own disposable copy of the one baseline.
        assert mutate_case.working_copy_path is not None
        assert baseline_case.working_copy_path is not None
        assert mutate_case.working_copy_path != baseline_case.working_copy_path
        for case in result.cases:
            notes = " ".join(case.notes)
            assert f"baseline_sha256={baseline_sha}" in notes
            assert f"working_copy_sha256={baseline_sha}" in notes, case.notes

        # Direct non-leak proof: the immutable baseline (the source every case is
        # copied from) does not contain the mutation, so the later case cannot see
        # it. Open the baseline file read-only and inspect its sheet inventory.
        base_info = await rig.session.open_workbook(
            rig.baseline.stored_path, mode=SessionMode.INSPECT, read_only=True
        )
        try:
            names = {sheet.name for sheet in await rig.sheets.list_sheets(base_info.workbook_id)}
        finally:
            await rig.session.close_workbook(base_info.workbook_id, save_changes=False)
        assert "Mutation" not in names, f"baseline leaked the mutation: {sorted(names)}"
        assert "Sheet1" in names

        _assert_owned_never_foreground(owned_pid, "e2e-isolation")
        _assert_no_visible_owned_windows(owned_pid, "e2e-isolation")

        assert _sha256(rig.original) == original_hash, "original must be unchanged"
        assert _sha256(fixture) == committed_hash, "committed fixture must never change"
        assert _sha256(Path(rig.baseline.stored_path)) == baseline_sha
    finally:
        await rig.shutdown_and_reap(owned_pid)

    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_original_fixtures_unchanged(tmp_path: Path) -> None:
    """A full suite leaves both committed fixtures byte-identical."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    compile_error_fixture = _fixture_or_skip(_COMPILE_ERROR_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        clean_before = _sha256(fixture)
        bad_before = _sha256(compile_error_fixture)
        original_before = _sha256(rig.original)

        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_macro_run(_ADD_NUMBERS)
        await rig.classify_safe(_ADD_NUMBERS)

        # A passing macro suite exercises snapshot / copy / close on CleanProject.
        suite = _suite(
            [
                {
                    "id": "add",
                    "steps": [{"action": "macro_run", "macro": _ADD_NUMBERS, "args": [4, 6]}],
                    "assertions": [
                        {"type": "macro_status", "expected": "success"},
                        {"type": "macro_return_value", "expected": 10},
                    ],
                }
            ],
            info.workbook_id,
        )
        result = await rig.runner.run_suite(suite)
        assert result.status is RunStatus.PASSED

        # Exercise CompileErrorProject too (open + real compile probe).
        ce_info = await rig.open_copy(compile_error_fixture)
        try:
            probe = await rig.compile.compile_project(ce_info.workbook_id)
            assert probe.status in (CompileStatus.COMPILE_ERROR, CompileStatus.INDETERMINATE)
        finally:
            await rig.session.close_workbook(ce_info.workbook_id, save_changes=False)

        _assert_owned_never_foreground(owned_pid, "e2e-originals")
        _assert_no_visible_owned_windows(owned_pid, "e2e-originals")

        # The committed fixtures and the temp original are byte-identical.
        assert _sha256(fixture) == clean_before, "committed CleanProject must never change"
        assert _sha256(compile_error_fixture) == bad_before, "CompileErrorProject must never change"
        assert _sha256(rig.original) == original_before, "temp original must be unchanged"
    finally:
        await rig.shutdown_and_reap(owned_pid)

    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"


@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_run_resource_retrieval(tmp_path: Path) -> None:
    """The excel://test-runs/{run_id} resource path returns the persisted run."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    rig = _make_rig(tmp_path, fixture)
    baseline_pids = _excel_pids()
    await rig.start()
    owned_pid: int | None = None
    try:
        info = await rig.open()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None
        rig.grant_macro_run(_ADD_NUMBERS)
        await rig.classify_safe(_ADD_NUMBERS)

        suite = _suite(
            [
                {
                    "id": "add",
                    "steps": [{"action": "macro_run", "macro": _ADD_NUMBERS, "args": [10, 5]}],
                    "assertions": [
                        {"type": "macro_status", "expected": "success"},
                        {"type": "macro_return_value", "expected": 15},
                    ],
                }
            ],
            info.workbook_id,
            name="resource-suite",
        )
        result = await rig.runner.run_suite(suite)
        run_id = result.test_run_id

        # The resource handler delegates to test_tools.get_test_runner().get_run();
        # publish the rig runner and drive the same read path.
        test_tools.set_test_runner(rig.runner)
        try:
            assert test_tools.get_test_runner() is rig.runner
            retrieved = await test_tools.test_get_run(run_id)
        finally:
            test_tools.set_test_runner(None)

        assert retrieved.test_run_id == run_id
        assert retrieved.suite == "resource-suite"
        assert retrieved.status is RunStatus.PASSED
        assert [c.case_id for c in retrieved.cases] == ["add"]
        # The resource returns run.model_dump_json(indent=2); verify that payload.
        payload = json.loads(retrieved.model_dump_json(indent=2))
        assert payload["test_run_id"] == run_id
        assert payload["status"] == "passed"
        assert payload["baseline_snapshot_id"] == rig.baseline.snapshot_id

        _assert_owned_never_foreground(owned_pid, "e2e-resource")
        _assert_no_visible_owned_windows(owned_pid, "e2e-resource")
    finally:
        await rig.shutdown_and_reap(owned_pid)

    missing_user_pids = baseline_pids - _excel_pids()
    assert not missing_user_pids, f"user-owned Excel PIDs disappeared: {missing_user_pids}"
