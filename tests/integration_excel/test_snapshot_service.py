"""Live-Excel integration verification for the snapshot service (task_014).

Drives the *production* :class:`SnapshotService` (with the real single-STA
:class:`ExcelComWorker` + :class:`ExcelAppAdapter`, a real
:class:`SessionService` and a real :class:`SnapshotStore`) against the target
Excel build, always hidden / non-activating. It uses only a temporary **copy** of
the ``CleanProject.xlsm`` fixture as the "original"; the committed fixture is
never opened or modified.

Coverage:

* ``test_snapshot_create_open_and_restore`` — open the working copy, snapshot it
  via ``SaveCopyAs`` (open branch), verify the manifest + SHA-256, mutate the
  working copy in memory (unsaved), restore (close-without-save + reopen the
  snapshot baseline), and assert the baseline is restored, the snapshot file is
  unchanged and the original fixture is unchanged. Also asserts
  ``external_side_effects_not_undone=true`` and BG03/BG04/BG05.
* ``test_snapshot_closed_branch_and_diff`` — after closing the working copy, take
  a closed-branch (file-copy) snapshot of the same workbook id and diff it against
  the earlier SaveCopyAs snapshot.
* ``test_each_case_starts_from_same_baseline`` — create one immutable baseline
  snapshot, then run two independent cases that each open the baseline, mutate the
  working copy differently and close without saving. Asserts the baseline snapshot
  file SHA-256 is identical before/after both cases (one case cannot contaminate
  another), each case's working copy starts byte-identical to the baseline, and the
  original fixture is unchanged (task_015).
* ``test_restore_after_mutation`` — mutate a working copy and ``restore_snapshot``;
  assert the restored working copy matches the baseline, the snapshot file is
  unchanged and ``external_side_effects_not_undone=true`` is recorded in the
  manifest and audit trail. Then simulate an abnormal Excel end by terminating ONLY
  the owned ``EXCEL.exe`` PID, prove a restore on the dead rig returns a structured
  error (no false success), and recover from the same immutable snapshot into a
  fresh session (task_015).

Every test asserts the owned instance stays hidden / non-activating
(BG03/BG04) and that the owned ``EXCEL.exe`` is reaped on shutdown (no orphan).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_snapshot_service.py -v -s
"""

from __future__ import annotations

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
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import WorkbookInfo
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
_CALL_TIMEOUT_S = 30.0

_EXTERNAL_SIDE_EFFECTS_NOTE = (
    "A snapshot/restore reverts the workbook FILE only. It does NOT revert "
    "external side effects (database writes, email/network calls, files touched "
    "on disk, etc.). A snapshot is not a sandbox (spec §9.4)."
)


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
        result.append({"hwnd": hwnd, "pid": int(pid), "title": title, "visible": visible})
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_cb, None)
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
    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[w['title'] for w in visible]}"
    )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _terminate_owned_pid(pid: int) -> None:
    """Force-terminate ONLY the automation-owned ``EXCEL.exe`` (abnormal end).

    Safety: this is only ever called with the owned PID returned by
    ``worker.get_status().owned_pid`` (a ``DispatchEx`` instance owned by the
    test). It never touches a non-owned / user-session Excel process.

    Raises:
        RuntimeError: If ``taskkill`` reports a non-zero exit code.
    """

    result = subprocess.run(
        ["taskkill", "/F", "/PID", str(pid)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"taskkill /F /PID {pid} failed (rc={result.returncode}): "
            f"{result.stdout.strip()} {result.stderr.strip()}"
        )


def _read_audit_events(path: Path) -> list[dict[str, Any]]:
    """Parse an append-only JSONL audit file into a list of event dicts."""

    events: list[dict[str, Any]] = []
    if not path.is_file():
        return events
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        try:
            parsed = json.loads(stripped)
        except (ValueError, TypeError):
            continue
        if isinstance(parsed, dict):
            events.append(parsed)
    return events


# ---------------------------------------------------------------------------
# STA-thread cell helpers (must run on the worker that owns ``app_adapter``)
# ---------------------------------------------------------------------------
def _read_a1(app_adapter: ExcelAppAdapter, path: str) -> str:
    """Read ``A1`` as a string on the owning STA thread."""

    wb = _find_wb(app_adapter, path)
    return str(wb.Worksheets(1).Range("A1").Value)


def _mutate_a1(app_adapter: ExcelAppAdapter, path: str, value: str) -> str:
    """Write ``A1`` on the owning STA thread and echo the written value."""

    wb = _find_wb(app_adapter, path)
    wb.Worksheets(1).Range("A1").Value = value
    return value


# ---------------------------------------------------------------------------
# Fixtures / service builder
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _SnapshotRig:
    """Bundle of a live worker + session + snapshot service + temp dirs."""

    def __init__(
        self,
        tmp_dir: Path,
        original: Path,
        *,
        store_root: Path | None = None,
        audit_path: Path | None = None,
    ) -> None:
        self.tmp_dir = tmp_dir
        self.original = original
        self.path_policy = PathPolicy({str(tmp_dir): str(tmp_dir)})
        self.canonical = str(self.path_policy.validate_path(str(original), OperationClass.MUTATING))
        self.approval = ApprovalPolicy()
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(audit_path or (tmp_dir / "audit.jsonl"))
        self.session = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_dir / "work",
        )
        self.store = SnapshotStore(store_root or (tmp_dir / "snapshots"))
        self.snapshot = SnapshotService(
            self.store,
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )

    def grant(self, action: str, canonical: str | None = None) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=canonical or self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="integration-test",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_open_for(self, path: Path) -> None:
        canonical = str(self.path_policy.validate_path(str(path), OperationClass.MUTATING))
        self.grant("excel_open", canonical=canonical)

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


def _make_rig(tmp_path: Path) -> _SnapshotRig:
    original = tmp_path / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    return _SnapshotRig(tmp_path, original)


# ---------------------------------------------------------------------------
# STA-thread cell helpers
# ---------------------------------------------------------------------------
def _find_wb(app_adapter: ExcelAppAdapter, working_path: str) -> Any:
    app = app_adapter.app
    target = working_path.lower()
    for wb in app.Workbooks:  # type: ignore[union-attr]
        if str(wb.FullName).lower() == target:
            return wb
    raise RuntimeError(f"working copy not found: {working_path}")


# ---------------------------------------------------------------------------
# Test 1: create (open branch) + mutate + restore
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_snapshot_create_open_and_restore(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")  # writable open of the original working copy
    await rig.start()
    owned_pid: int | None = None
    old_workbook_id: str | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.session.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        old_workbook_id = info.workbook_id
        working_path = info.path
        assert Path(working_path) != rig.original.resolve(), "must open a working copy"

        # Read the baseline A1 BEFORE any mutation.
        def read_a1() -> str:
            wb = _find_wb(rig.app_adapter, working_path)
            return str(wb.Worksheets(1).Range("A1").Value)

        baseline_a1 = await rig.worker.submit(
            read_a1, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.READ_ONLY
        )

        # --- Create a snapshot (open branch: SaveCopyAs) ---
        manifest = await rig.snapshot.create_snapshot(old_workbook_id)

        assert manifest.snapshot_method == "SaveCopyAs"
        assert manifest.external_side_effects_not_undone is True
        snapshot_file = rig.store.path_for(manifest.snapshot_id) / manifest.file_name
        assert snapshot_file.is_file(), "snapshot file must exist"
        # Manifest SHA-256 matches the stored snapshot file.
        assert rig.store.sha256(snapshot_file) == manifest.sha256
        assert manifest.sha256
        # The original fixture is unchanged by the snapshot.
        assert _sha256(rig.original) == original_hash, "snapshot must not modify the original"
        # The working copy is still open (SaveCopyAs does not replace it).
        assert await rig.session.get_status(old_workbook_id) is not None

        _assert_owned_never_foreground(owned_pid, "after-create")
        _assert_no_visible_owned_windows(owned_pid, "after-create")

        # --- Mutate the working copy in memory (do NOT save) ---
        def mutate_a1() -> str:
            wb = _find_wb(rig.app_adapter, working_path)
            wb.Worksheets(1).Range("A1").Value = "MUTATED-TASK014"
            return "mutated"

        await rig.worker.submit(
            mutate_a1, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.MUTATING
        )

        def read_a1_after_mutation() -> str:
            wb = _find_wb(rig.app_adapter, working_path)
            return str(wb.Worksheets(1).Range("A1").Value)

        mutated_a1 = await rig.worker.submit(
            read_a1_after_mutation, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.READ_ONLY
        )
        assert mutated_a1 == "MUTATED-TASK014", "mutation should be visible in memory"

        # --- Restore (close-without-save + reopen the snapshot baseline) ---
        rig.grant_open_for(snapshot_file)  # restore opens a writable working copy
        restored = await rig.snapshot.restore_snapshot(
            manifest.snapshot_id, workbook_id=old_workbook_id
        )

        assert restored.workbook.workbook_id != old_workbook_id, (
            "restore opens a fresh working copy"
        )
        restored_working_path = restored.workbook.path
        assert Path(restored_working_path) != snapshot_file.resolve(), (
            "restore must open a working copy, not the snapshot file itself"
        )
        # R1: a clean previous-copy close is reported as "closed".
        assert restored.previous_close == "closed"
        # R3: the restored identity is rebound to the ORIGINAL source path.
        assert restored.workbook.source_path == manifest.source_path

        # The restored working copy shows the BASELINE A1 (mutation discarded).
        def read_restored_a1() -> str:
            wb = _find_wb(rig.app_adapter, restored_working_path)
            return str(wb.Worksheets(1).Range("A1").Value)

        restored_a1 = await rig.worker.submit(
            read_restored_a1, timeout=_CALL_TIMEOUT_S, op_class=OperationClass.READ_ONLY
        )
        assert restored_a1 == baseline_a1, (
            f"restore must restore the baseline A1 ({baseline_a1!r}), got {restored_a1!r}"
        )
        assert restored_a1 != "MUTATED-TASK014", "mutation must not survive restore"

        # The immutable snapshot file is byte-identical after restore.
        assert rig.store.sha256(snapshot_file) == manifest.sha256
        # The original fixture is still unchanged (BG05).
        assert _sha256(rig.original) == original_hash, "restore must not modify the original"

        # external_side_effects_not_undone is stated in the manifest.
        assert rig.store.load(manifest.snapshot_id).external_side_effects_not_undone is True

        _assert_owned_never_foreground(owned_pid, "after-restore")
        _assert_no_visible_owned_windows(owned_pid, "after-restore")

        # Cleanup the restored working copy.
        await rig.session.close_workbook(restored.workbook.workbook_id)

        print(
            f"[snapshot] snapshot_id={manifest.snapshot_id} method={manifest.snapshot_method} "
            f"sha256={manifest.sha256[:12]}… baseline_a1={baseline_a1!r} "
            f"restored_a1={restored_a1!r} owned_pid={owned_pid} original_unchanged=True"
        )
        print(f"[side-effects] {_EXTERNAL_SIDE_EFFECTS_NOTE}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: closed branch (file copy) + diff against the SaveCopyAs snapshot
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_snapshot_closed_branch_and_diff(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")
    await rig.start()
    owned_pid: int | None = None
    old_workbook_id: str | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.session.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        old_workbook_id = info.workbook_id

        # Open-branch snapshot (SaveCopyAs).
        open_manifest = await rig.snapshot.create_snapshot(old_workbook_id)
        assert open_manifest.snapshot_method == "SaveCopyAs"

        # Close the working copy; the service remembers the source path.
        await rig.session.close_workbook(old_workbook_id)
        assert await rig.session.list_workbooks() == []

        # Closed-branch snapshot of the same workbook id -> file copy.
        closed_manifest = await rig.snapshot.create_snapshot(old_workbook_id)
        assert closed_manifest.snapshot_method == "file_copy"
        assert closed_manifest.sha256 == _sha256(rig.original), (
            "file-copy snapshot must hash to the source file"
        )

        # Diff the two snapshots: same source, different method + re-serialization.
        diff = await rig.snapshot.diff_snapshots(
            open_manifest.snapshot_id, closed_manifest.snapshot_id
        )
        assert diff.sha256_a == open_manifest.sha256
        assert diff.sha256_b == closed_manifest.sha256
        assert diff.file_identical is (open_manifest.sha256 == closed_manifest.sha256)
        assert "snapshot_method" in diff.changed
        assert diff.changed["snapshot_method"]["a"] == "SaveCopyAs"
        assert diff.changed["snapshot_method"]["b"] == "file_copy"

        # List returns both snapshots.
        listed = await rig.snapshot.list_snapshots()
        listed_ids = {m.snapshot_id for m in listed}
        assert open_manifest.snapshot_id in listed_ids
        assert closed_manifest.snapshot_id in listed_ids

        _assert_owned_never_foreground(owned_pid, "closed-branch")
        _assert_no_visible_owned_windows(owned_pid, "closed-branch")

        print(
            f"[closed] open_sha={open_manifest.sha256[:12]}… "
            f"closed_sha={closed_manifest.sha256[:12]}… identical={diff.file_identical}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3 (task_015): every case starts from the same immutable baseline
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_each_case_starts_from_same_baseline(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")  # writable open of the original working copy
    await rig.start()
    owned_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- Create the single immutable baseline snapshot from a working copy ---
        info0 = await rig.session.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        baseline_manifest = await rig.snapshot.create_snapshot(info0.workbook_id)
        assert baseline_manifest.snapshot_method == "SaveCopyAs"
        baseline_file = (
            rig.store.path_for(baseline_manifest.snapshot_id) / baseline_manifest.file_name
        )
        baseline_sha = baseline_manifest.sha256
        assert _sha256(baseline_file) == baseline_sha, "manifest SHA must match the baseline file"

        # The baseline A1 value (empty cell -> "None").
        baseline_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, info0.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        await rig.session.close_workbook(info0.workbook_id)

        # Approval for opening the immutable baseline snapshot as a working copy.
        rig.grant_open_for(baseline_file)

        # --- Case A: open baseline, mutate A1, close WITHOUT save ---
        info_a = await rig.session.open_workbook(
            str(baseline_file), mode=SessionMode.INSPECT, read_only=False
        )
        assert info_a.workbook_id != info0.workbook_id
        working_sha_a = _sha256(Path(info_a.path))
        a1_start_a = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, info_a.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert working_sha_a == baseline_sha, (
            "case A working copy must start byte-identical to the baseline snapshot"
        )
        assert a1_start_a == baseline_a1, "case A must start from the baseline A1"
        await rig.worker.submit(
            lambda: _mutate_a1(rig.app_adapter, info_a.path, "MUT-CASE-A"),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.MUTATING,
        )
        await rig.session.close_workbook(info_a.workbook_id)  # default: no save
        assert _sha256(baseline_file) == baseline_sha, (
            "case A must not mutate the baseline snapshot"
        )

        # --- Case B: open the SAME baseline again, mutate differently, close ---
        info_b = await rig.session.open_workbook(
            str(baseline_file), mode=SessionMode.INSPECT, read_only=False
        )
        assert info_b.workbook_id not in {info0.workbook_id, info_a.workbook_id}
        working_sha_b = _sha256(Path(info_b.path))
        a1_start_b = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, info_b.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert working_sha_b == baseline_sha, (
            "case B working copy must start byte-identical to the baseline snapshot"
        )
        assert a1_start_b == baseline_a1, "case B must start from the same baseline value as case A"
        await rig.worker.submit(
            lambda: _mutate_a1(rig.app_adapter, info_b.path, "MUT-CASE-B"),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.MUTATING,
        )
        await rig.session.close_workbook(info_b.workbook_id)

        # --- Isolation / integrity assertions ---
        # The immutable baseline file is byte-identical before/after both cases.
        assert _sha256(baseline_file) == baseline_sha, "baseline snapshot changed across cases"
        assert rig.store.sha256(baseline_file) == baseline_manifest.sha256
        # Both cases started from the exact same baseline hash (no cross-contamination).
        assert working_sha_a == working_sha_b == baseline_sha
        # The original fixture is untouched by every case.
        assert _sha256(rig.original) == original_hash, "original fixture must be unchanged"

        _assert_owned_never_foreground(owned_pid, "cases")
        _assert_no_visible_owned_windows(owned_pid, "cases")

        print(
            f"[isolation] baseline_sha={baseline_sha[:12]}… "
            f"caseA_start_sha={working_sha_a[:12]}… caseB_start_sha={working_sha_b[:12]}… "
            f"baseline_a1={baseline_a1!r} original_unchanged=True"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4 (task_015): restore after mutation + crash recovery into a fresh session
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_after_mutation(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    rig.grant("excel_open")
    await rig.start()
    owned_pid: int | None = None
    fresh: _SnapshotRig | None = None
    fresh_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        # --- Create the immutable baseline snapshot ---
        info0 = await rig.session.open_workbook(
            str(rig.original), mode=SessionMode.EXECUTE, read_only=False
        )
        manifest = await rig.snapshot.create_snapshot(info0.workbook_id)
        snapshot_file = rig.store.path_for(manifest.snapshot_id) / manifest.file_name
        baseline_sha = manifest.sha256
        baseline_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, info0.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        await rig.session.close_workbook(info0.workbook_id)

        # --- Part A: normal restore after an unsaved mutation ---
        rig.grant_open_for(snapshot_file)
        info = await rig.session.open_workbook(
            str(snapshot_file), mode=SessionMode.INSPECT, read_only=False
        )
        await rig.worker.submit(
            lambda: _mutate_a1(rig.app_adapter, info.path, "MUT-RESTORE"),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.MUTATING,
        )
        restored = await rig.snapshot.restore_snapshot(
            manifest.snapshot_id, workbook_id=info.workbook_id
        )
        assert isinstance(restored.workbook, WorkbookInfo)
        assert restored.workbook.workbook_id != info.workbook_id, (
            "restore opens a fresh working copy"
        )
        assert Path(restored.workbook.path) != snapshot_file.resolve(), (
            "restore must open a working copy, not the snapshot file itself"
        )
        restored_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, restored.workbook.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert restored_a1 == baseline_a1, (
            f"restore must restore the baseline A1 ({baseline_a1!r}), got {restored_a1!r}"
        )
        # The immutable snapshot file is byte-identical after restore.
        assert rig.store.sha256(snapshot_file) == baseline_sha
        # external_side_effects_not_undone is stated on the manifest.
        assert rig.store.load(manifest.snapshot_id).external_side_effects_not_undone is True
        # ... and recorded in the audit trail for the restore.
        rig.audit.flush()
        restore_events = [
            event
            for event in _read_audit_events(rig.audit.path)
            if event.get("tool") == "snapshot_restore"
        ]
        assert restore_events, "restore must be recorded in the audit trail"
        assert any(
            (event.get("result") or {}).get("external_side_effects_not_undone") is True
            for event in restore_events
        ), "audit must record external_side_effects_not_undone=true"
        # The original fixture is unchanged (BG05).
        assert _sha256(rig.original) == original_hash

        _assert_owned_never_foreground(owned_pid, "after-restore")
        _assert_no_visible_owned_windows(owned_pid, "after-restore")
        await rig.session.close_workbook(restored.workbook.workbook_id)

        # --- Part B: crash recovery (abnormal end of ONLY the owned EXCEL.exe) ---
        info2 = await rig.session.open_workbook(
            str(snapshot_file), mode=SessionMode.INSPECT, read_only=False
        )
        await rig.worker.submit(
            lambda: _mutate_a1(rig.app_adapter, info2.path, "MUT-BEFORE-CRASH"),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.MUTATING,
        )
        crash_pid = rig.worker.get_status().owned_pid
        assert crash_pid is not None and crash_pid == owned_pid, (
            "must terminate ONLY the owned automation PID"
        )
        _terminate_owned_pid(crash_pid)
        assert _wait_pid_gone(crash_pid), f"owned PID {crash_pid} did not terminate"

        # A restore attempt on the crashed rig must NOT falsely succeed; it must
        # return a structured error (stale session / dead instance -> ExcelVbaMcpError).
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await rig.snapshot.restore_snapshot(manifest.snapshot_id)
        structured = excinfo.value.to_dict()
        assert structured.get("code"), "structured error must carry a stable code"
        assert _sha256(snapshot_file) == baseline_sha, "a crash must not alter the snapshot file"

        # Recovery: a FRESH session restores from the same immutable snapshot.
        fresh = _SnapshotRig(
            tmp_path,
            rig.original,
            store_root=rig.store.root,  # share the immutable snapshot store
            audit_path=tmp_path / "audit_fresh.jsonl",
        )
        await fresh.start()
        fresh_pid = fresh.worker.get_status().owned_pid
        assert fresh_pid is not None and fresh_pid != crash_pid, (
            "recovery must run on a fresh, distinct Excel instance"
        )
        fresh.grant_open_for(snapshot_file)
        recovered = await fresh.snapshot.restore_snapshot(manifest.snapshot_id)
        assert isinstance(recovered.workbook, WorkbookInfo)
        # No previous workbook_id was supplied -> previous_close is not_requested (R1).
        assert recovered.previous_close == "not_requested"
        recovered_a1 = await fresh.worker.submit(
            lambda: _read_a1(fresh.app_adapter, recovered.workbook.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert recovered_a1 == baseline_a1, (
            f"fresh-session recovery must restore the baseline A1 ({baseline_a1!r})"
        )
        assert fresh.store.sha256(snapshot_file) == baseline_sha
        assert _sha256(rig.original) == original_hash, "original fixture must be unchanged"

        _assert_owned_never_foreground(fresh_pid, "fresh-recovery")
        _assert_no_visible_owned_windows(fresh_pid, "fresh-recovery")
        await fresh.session.close_workbook(recovered.workbook.workbook_id)

        print(
            f"[restore] baseline_sha={baseline_sha[:12]}… restored_a1={restored_a1!r} "
            f"external_side_effects_not_undone=True crash_pid={crash_pid}(dead) "
            f"fresh_pid={fresh_pid} recovered_a1={recovered_a1!r}"
        )
        print(f"[side-effects] {_EXTERNAL_SIDE_EFFECTS_NOTE}")
    finally:
        if fresh is not None:
            await fresh.stop()
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
    if fresh_pid is not None:
        assert _wait_pid_gone(fresh_pid), f"fresh owned PID {fresh_pid} still alive after shutdown"
