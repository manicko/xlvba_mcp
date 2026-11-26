"""Live-Excel verification of guarded VBA source mutations (task_019).

Independent verification of the mutation-safety contract enforced by
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService`
(``replace_component_source`` / ``insert_code``), the pre-mutation snapshot gate
(:class:`~excel_vba_mcp.application.snapshot_service.SnapshotService`) and the
deny-by-default approval gate
(:class:`~excel_vba_mcp.security.approval.ApprovalPolicy`).

Every test drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` on a hidden,
non-activating Excel instance and only ever mutates a disposable working copy of
the committed ``CleanProject.xlsm`` fixture. The committed fixture and the
per-test original copy are never modified.

Coverage:

* ``test_replace_and_insert_on_disposable_module`` — full replace plus two
  semantic inserts; before/after hashes, export files and read-back source.
* ``test_stale_revision_rejected`` — a mismatched ``expected_source_hash`` is
  rejected and the source is unchanged.
* ``test_mutation_blocked_without_snapshot`` — ``SNAPSHOT_FAILED`` blocks the
  mutation when no snapshot exists.
* ``test_mutation_blocked_without_approval`` — an empty approval policy denies
  ``authorize_vba_write`` and the mutation; the source is unchanged.
* ``test_failed_mutation_keeps_snapshot_and_diagnostics`` — a rejected mutation
  keeps the snapshot and returns a structured, coded error.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vba_project_write.py -v -s
"""

from __future__ import annotations

import contextlib
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
from excel_vba_mcp.adapters.excel.vbide_adapter import hash_code
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, SnapshotFailedError
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
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

# A full replacement for ``modTest`` that keeps ``Option Explicit`` and adds a
# verifiable marker procedure (task_019).
_REPLACE_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Private Sub Helper()
    ' helper body
End Sub

Public Sub Task019Marker()
    ' task_019 replace marker
End Sub
"""

# Distinct single-line markers so each insert can be located unambiguously.
_DECL_MARKER = "' task_019 declarations insert"
_BEFORE_MARKER = "' task_019 before-procedure insert"


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


# ---------------------------------------------------------------------------
# Read-only desktop / process helpers (never activate / show / kill)
# ---------------------------------------------------------------------------
def _pid_from_hwnd(hwnd: int) -> int | None:
    if win32process is None:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


def _pid_alive(pid: int) -> bool:
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive
        return False
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        pid = _pid_from_hwnd(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    if pid is None:
        return None
    return {"hwnd": hwnd, "pid": pid, "title": title}


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
        result.append({"hwnd": hwnd, "pid": int(pid), "title": title, "visible": visible})
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_cb, None)
    return result


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    """BG03: the owned automation instance never holds the foreground."""

    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at {label} "
        f"(title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    """BG04: the owned automation instance has no visible windows."""

    visible = [w for w in _windows_for_pid(owned_pid) if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[w['title'] for w in visible]}"
    )


def _assert_user_session_preserved(
    owned_pid: int, fg_before: dict[str, Any] | None, label: str
) -> None:
    """BG05: the interactive foreground is not stolen by the owned instance."""

    fg_after = _foreground()
    if fg_after is not None:
        assert fg_after["pid"] != owned_pid, (
            f"BG05 violation: owned automation PID {owned_pid} became foreground at {label} "
            f"(title={fg_after['title']!r})"
        )
    if fg_before is not None and fg_before["pid"] != owned_pid and fg_after is not None:
        assert fg_after["hwnd"] == fg_before["hwnd"], (
            f"BG05 violation: interactive foreground changed across {label} "
            f"(before={fg_before['title']!r} after={fg_after['title']!r})"
        )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _grant(
    approval: ApprovalPolicy,
    canonical: str,
    action: str,
    *,
    sha256: str | None = None,
    scope: ApprovalScope = ApprovalScope.SESSION,
) -> None:
    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, sha256=sha256, action=action),
            scope=scope,
            granted_at=datetime.now(UTC),
            granted_by="task019-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session + snapshot + VBA project service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.snapshots_dir = tmp_dir / "snapshots"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.snapshots_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {
                "work": str(self.work_dir),
                "projects": str(self.projects_dir),
                "snapshots": str(self.snapshots_dir),
            }
        )
        self.approval = ApprovalPolicy()
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
        self.service = VBAProjectService(
            self.worker,
            self.session,
            self.path_policy,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )
        self._original: Path | None = None
        self._writable = False

    @property
    def original(self) -> Path:
        assert self._original is not None, "original path not bound"
        return self._original

    def bind_original(self, original: Path, *, writable: bool = False) -> None:
        self._original = original
        self._writable = writable

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    async def open(self) -> Any:
        canonical = str(self.path_policy.validate_path(str(self.original), OperationClass.MUTATING))
        open_op = OperationClass.MUTATING if self._writable else OperationClass.READ_ONLY
        if open_op == OperationClass.MUTATING:
            _grant(self.approval, canonical, "excel_open", sha256=_sha256(self.original))
        return await self.session.open_workbook(
            str(self.original),
            mode=SessionMode.INSPECT,
            read_only=not self._writable,
        )

    def grant_mutation(self, canonical: str, expected_hash: str, action: str) -> None:
        _grant(
            self.approval,
            canonical,
            action,
            sha256=expected_hash,
            scope=ApprovalScope.ONE_TIME,
        )

    def canonical_working(self, workbook_id: str) -> str:
        return str(
            self.path_policy.validate_path(self.working_path(workbook_id), OperationClass.MUTATING)
        )

    def working_path(self, workbook_id: str) -> str:
        return self.session._records[workbook_id].working_path


def _make_rig(tmp_path: Path, *, writable: bool = True) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    rig.bind_original(original, writable=writable)
    return rig


# ---------------------------------------------------------------------------
# Test 1: replace + insert on a disposable module
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_replace_and_insert_on_disposable_module(tmp_path: Path) -> None:
    """Replace ``modTest`` then insert at two semantic positions."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig.original)
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert info.has_vba is True
        assert info.read_only is False, "writable session expected"

        snapshot = await rig.snapshot.create_snapshot(info.workbook_id)
        assert snapshot.snapshot_id
        assert rig.snapshot.has_snapshot(info.workbook_id)

        # --- replace_component_source -------------------------------------
        before = await rig.service.read_component(info.workbook_id, "modTest")
        assert before.source_hash == hash_code(before.source)
        assert "Task019Marker" not in before.source

        canonical = rig.canonical_working(info.workbook_id)
        rig.grant_mutation(canonical, before.source_hash, "vba_replace_component_source")

        replace_result = await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            _REPLACE_SOURCE,
            expected_source_hash=before.source_hash,
        )

        assert replace_result.before_hash == before.source_hash
        assert replace_result.after_hash == hash_code(_REPLACE_SOURCE)
        assert replace_result.before_hash != replace_result.after_hash

        before_export = Path(replace_result.before_export_path)
        after_export = Path(replace_result.after_export_path)
        assert before_export.is_file(), f"before export missing: {before_export}"
        assert after_export.is_file(), f"after export missing: {after_export}"

        after_replace = await rig.service.read_component(info.workbook_id, "modTest")
        assert "Task019Marker" in after_replace.source
        assert after_replace.source_hash == replace_result.after_hash

        # --- insert_code at "declarations" --------------------------------
        decl_hash = after_replace.source_hash
        rig.grant_mutation(canonical, decl_hash, "vba_insert_code")

        decl_result = await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            _DECL_MARKER,
            position="declarations",
            expected_source_hash=decl_hash,
        )

        assert decl_result.before_hash == decl_hash
        assert decl_result.after_hash != decl_hash
        assert Path(decl_result.after_export_path).is_file()

        after_decl = await rig.service.read_component(info.workbook_id, "modTest")
        assert _DECL_MARKER in after_decl.source
        option_pos = after_decl.source.find("Option Explicit")
        decl_pos = after_decl.source.find(_DECL_MARKER)
        add_pos = after_decl.source.find("AddNumbers")
        assert option_pos != -1, "Option Explicit missing after declarations insert"
        assert decl_pos != -1, "declarations insert marker missing"
        assert add_pos != -1, "AddNumbers missing after declarations insert"
        assert option_pos < decl_pos < add_pos, (
            "declarations insert must land after Option Explicit and before the first procedure"
        )
        assert after_decl.source_hash == decl_result.after_hash

        # --- insert_code at "before_procedure:AddNumbers" -----------------
        before_proc_hash = after_decl.source_hash
        rig.grant_mutation(canonical, before_proc_hash, "vba_insert_code")

        before_result = await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            _BEFORE_MARKER,
            position="before_procedure:AddNumbers",
            expected_source_hash=before_proc_hash,
        )

        assert before_result.before_hash == before_proc_hash
        assert before_result.after_hash != before_proc_hash
        assert before_result.procedure == "AddNumbers"
        assert Path(before_result.after_export_path).is_file()

        after_before = await rig.service.read_component(info.workbook_id, "modTest")
        assert _BEFORE_MARKER in after_before.source
        before_pos = after_before.source.find(_BEFORE_MARKER)
        add_pos_2 = after_before.source.find("AddNumbers")
        assert before_pos != -1, "before-procedure insert marker missing"
        assert add_pos_2 != -1, "AddNumbers missing after before-procedure insert"
        assert before_pos < add_pos_2, "before-procedure insert must precede AddNumbers"
        assert after_before.source.find("Option Explicit") < before_pos
        assert after_before.source_hash == before_result.after_hash

        # Originals are never touched.
        assert _sha256(rig.original) == original_hash, "per-test original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        _assert_owned_never_foreground(owned_pid, "vba-project-write")
        _assert_no_visible_owned_windows(owned_pid, "vba-project-write")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-project-write")

        print(
            "[replace] "
            f"before={replace_result.before_hash[:16]}… after={replace_result.after_hash[:16]}…"
        )
        print(
            "[insert:declarations] "
            f"before={decl_result.before_hash[:16]}… after={decl_result.after_hash[:16]}…"
        )
        print(
            "[insert:before_procedure:AddNumbers] "
            f"before={before_result.before_hash[:16]}… after={before_result.after_hash[:16]}…"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: stale revision rejected
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_stale_revision_rejected(tmp_path: Path) -> None:
    """A mismatched ``expected_source_hash`` is rejected; the source is unchanged."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        current = await rig.service.read_component(info.workbook_id, "modTest")
        actual_hash = current.source_hash

        stale_hash = "f" * 64
        assert stale_hash != actual_hash

        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=stale_hash,
            )

        err = exc_info.value
        # FE-22 (F-5): a stale expected_source_hash is an optimistic-concurrency
        # failure (REVISION_CONFLICT), not generic COM contention.
        assert err.code is ErrorCode.REVISION_CONFLICT
        assert err.context.get("reason") == "revision_conflict"
        assert err.context.get("expected_source_hash") == stale_hash
        assert err.context.get("current_source_hash") == actual_hash
        diagnostics = err.to_dict()
        assert diagnostics["errorCode"] == ErrorCode.REVISION_CONFLICT.value
        assert diagnostics["details"]["reason"] == "revision_conflict"

        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == actual_hash, "stale mutation must not overwrite the source"
        assert "Task019Marker" not in after.source

        # The snapshot is still present (nothing was mutated).
        assert rig.snapshot.has_snapshot(info.workbook_id)

        _assert_owned_never_foreground(owned_pid, "vba-stale-revision")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-stale-revision")

        print(
            f"[stale-revision] rejected: expected={stale_hash[:12]}… "
            f"current={actual_hash[:16]}… source unchanged"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: mutation blocked without a snapshot
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_mutation_blocked_without_snapshot(tmp_path: Path) -> None:
    """Without a snapshot the mutation raises ``SNAPSHOT_FAILED`` and changes nothing."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert rig.snapshot.has_snapshot(info.workbook_id) is False

        current = await rig.service.read_component(info.workbook_id, "modTest")

        with pytest.raises(SnapshotFailedError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=current.source_hash,
            )

        err = exc_info.value
        assert err.code is ErrorCode.SNAPSHOT_FAILED
        assert err.to_dict()["errorCode"] == ErrorCode.SNAPSHOT_FAILED.value

        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == current.source_hash, "blocked mutation must not change source"
        assert "Task019Marker" not in after.source
        assert rig.snapshot.has_snapshot(info.workbook_id) is False

        _assert_owned_never_foreground(owned_pid, "vba-no-snapshot")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-no-snapshot")

        print(f"[no-snapshot] blocked with code={err.code.value}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4: mutation blocked without approval
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_mutation_blocked_without_approval(tmp_path: Path) -> None:
    """An empty approval policy denies ``authorize_vba_write`` and the mutation."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)
        current = await rig.service.read_component(info.workbook_id, "modTest")
        canonical = rig.canonical_working(info.workbook_id)

        # A brand-new, empty policy wired into a service that shares the same
        # worker / session / snapshot but has no grants at all.
        deny_policy = ApprovalPolicy()
        assert deny_policy.allowlist == []
        deny_service = VBAProjectService(
            rig.worker,
            rig.session,
            rig.path_policy,
            rig.audit,
            snapshot_service=rig.snapshot,
            approval=deny_policy,
        )

        key = ApprovalKey(
            canonical_path=canonical,
            sha256=current.source_hash,
            action="vba_replace_component_source",
            procedure="modTest",
        )
        decision = deny_policy.authorize_vba_write(key)
        assert decision.allowed is False, "deny-by-default policy must refuse the VBA write"
        assert decision.reason == "denied by default: no matching allowlist entry or live grant"

        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await deny_service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=current.source_hash,
            )

        err = exc_info.value
        assert err.code is ErrorCode.APPROVAL_DENIED
        assert err.context.get("reason") == decision.reason
        assert err.to_dict()["details"]["reason"] == decision.reason

        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == current.source_hash, "denied mutation must not change source"
        assert "Task019Marker" not in after.source

        _assert_owned_never_foreground(owned_pid, "vba-no-approval")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-no-approval")

        print(f"[no-approval] allowed={decision.allowed} reason={decision.reason!r}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 5: failed mutation keeps the snapshot and returns diagnostics
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_failed_mutation_keeps_snapshot_and_diagnostics(tmp_path: Path) -> None:
    """A rejected Document-module mutation keeps the snapshot and is structured."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        snapshot = await rig.snapshot.create_snapshot(info.workbook_id)
        snapshot_id = snapshot.snapshot_id
        snapshot_dir = rig.snapshot.store.path_for(snapshot_id)
        assert snapshot_dir.is_dir(), "snapshot directory must exist before the failed mutation"
        assert rig.snapshot.has_snapshot(info.workbook_id)

        doc = await rig.service.read_component(info.workbook_id, "ThisWorkbook")
        doc_hash = doc.source_hash

        # Grant approval so the failure is specifically the non-replaceable guard.
        canonical = rig.canonical_working(info.workbook_id)
        rig.grant_mutation(canonical, doc_hash, "vba_replace_component_source")

        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "ThisWorkbook",
                "Option Explicit\n",
                expected_source_hash=doc_hash,
            )

        err = exc_info.value
        # A Document-module (non-replaceable) mutation is a PROTECTED_TARGET
        # failure, not generic COM contention.
        assert err.code is ErrorCode.PROTECTED_TARGET
        assert err.context.get("reason") == "non_replaceable_component"
        assert err.context.get("component") == "ThisWorkbook"
        assert "component_type" in err.context

        diagnostics = err.to_dict()
        assert diagnostics["errorCode"] == ErrorCode.PROTECTED_TARGET.value
        assert diagnostics["message"]
        assert diagnostics["details"]["reason"] == "non_replaceable_component"
        assert diagnostics["details"]["component"] == "ThisWorkbook"

        # The snapshot survives the failed mutation.
        assert rig.snapshot.has_snapshot(info.workbook_id)
        manifest = rig.snapshot.store.load(snapshot_id)
        assert manifest.snapshot_id == snapshot_id
        assert Path(manifest.stored_path).is_file()
        assert snapshot_id in {m.snapshot_id for m in rig.snapshot.store.list()}

        # The Document module is untouched.
        after = await rig.service.read_component(info.workbook_id, "ThisWorkbook")
        assert after.source_hash == doc_hash

        _assert_owned_never_foreground(owned_pid, "vba-failed-mutation")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-failed-mutation")

        print(
            f"[failed-mutation] code={err.code.value} "
            f"reason={err.context.get('reason')} snapshot={snapshot_id} preserved"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
