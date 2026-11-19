"""Live-Excel verification of guarded VBA source mutations (task_018).

This is the task_018 independent verification of the production
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService`
mutation methods (``replace_component_source`` / ``insert_code``) and the
:class:`~excel_vba_mcp.security.approval.ApprovalPolicy.authorize_vba_write`
gate against the target Excel build. It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter`, a real
:class:`~excel_vba_mcp.application.session_service.SessionService` and a real
:class:`~excel_vba_mcp.application.snapshot_service.SnapshotService`, always
hidden / non-activating, and only ever touches a temporary **copy** of the
committed ``CleanProject.xlsm`` fixture. The committed fixture is never opened
or modified.

Coverage (spec §7.2; research ``g1-poc-research-2026-11-19.md`` §2;
``task_009-research-2026-11-19.md`` §3):

* ``test_replace_component_source_success`` — replace ``modTest`` source with a
  new revision; assert before/after hashes and export files.
* ``test_stale_revision_rejected`` — stale ``expected_source_hash`` is rejected
  and the source is unchanged.
* ``test_insert_declarations`` — insert at declarations position.
* ``test_insert_before_procedure`` — insert before a named procedure.
* ``test_insert_after_procedure`` — insert after a named procedure.
* ``test_document_modules_survive_mutation`` — Document modules are still
  present after a mutation.
* ``test_original_fixture_unchanged`` — the committed fixture is byte-identical
  after all mutations.
* ``test_read_only_rejected`` — a read-only session is rejected.
* ``test_snapshot_required`` — a mutation without a snapshot is blocked.
* ``test_approval_deny_by_default`` — a mutation without an approval grant is
  denied.
* BG03/BG04/BG05 background non-interference + orphan check after shutdown.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vba_mutation.py -v -s
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
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    VbaComponentType,
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
_CALL_TIMEOUT_S = 60.0


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
            granted_by="task018-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# New source for modTest that adds a new procedure (verifiable change).
_REPLACE_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Private Sub Helper()
    ' helper body
End Sub

Public Sub Task018Marker()
    ' task_018 mutation marker
End Sub
"""

# A simple comment line to insert.
_INSERT_COMMENT = "' task_018 insert marker"


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

    def bind_original(self, original: Path, *, writable: bool = False) -> None:
        self._original = original
        self._writable = writable

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    async def open(self) -> Any:
        assert self._original is not None
        canonical = str(
            self.path_policy.validate_path(str(self._original), OperationClass.MUTATING)
        )
        # Grant the open approval. A writable inspect open is MUTATING; a
        # read-only inspect open is READ_ONLY (auto-allowed).
        open_op = OperationClass.MUTATING if self._writable else OperationClass.READ_ONLY
        if open_op == OperationClass.MUTATING:
            _grant(self.approval, canonical, "excel_open", sha256=_sha256(self._original))
        return await self.session.open_workbook(
            str(self._original),
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


def _make_rig(tmp_path: Path, *, writable: bool = True) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    rig.bind_original(original, writable=writable)
    return rig


def _working_copy_path(rig: _Rig, workbook_id: str) -> str:
    """Return the working copy path for a workbook id (from session records)."""

    record = rig.session._records[workbook_id]
    return record.working_path


# ---------------------------------------------------------------------------
# Test 1: replace_component_source success
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_replace_component_source_success(tmp_path: Path) -> None:
    """Replace modTest source; assert before/after hashes and export files."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert info.has_vba is True
        assert info.read_only is False, "writable session expected"

        # Create a snapshot (required before mutation).
        snapshot = await rig.snapshot.create_snapshot(info.workbook_id)
        assert snapshot.snapshot_id

        # Read the current source to get the revision hash.
        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash
        assert len(expected_hash) == 64

        # Grant approval for the mutation.
        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_replace_component_source")

        # Perform the mutation.
        result = await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            _REPLACE_SOURCE,
            expected_source_hash=expected_hash,
        )

        # Before hash matches the original.
        assert result.before_hash == expected_hash
        # After hash is different (new source was written).
        assert result.after_hash != expected_hash
        assert result.after_hash == hash_code(_REPLACE_SOURCE)

        # Export paths exist.
        before_path = Path(result.before_export_path)
        after_path = Path(result.after_export_path)
        assert before_path.is_file(), f"before export missing: {before_path}"
        assert after_path.is_file(), f"after export missing: {after_path}"

        # The source actually changed (read back via the service).
        after_source = await rig.service.read_component(info.workbook_id, "modTest")
        assert "Task018Marker" in after_source.source
        assert after_source.source_hash == result.after_hash

        # Original and committed fixture are untouched.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        # BG03/BG04/BG05 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "vba-replace")
        _assert_no_visible_owned_windows(owned_pid, "vba-replace")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-replace")

        print(
            f"[vba-replace] before={result.before_hash[:16]}… "
            f"after={result.after_hash[:16]}… "
            f"before_export={before_path.name} after_export={after_path.name}"
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
    """A stale expected_source_hash is rejected and the source is unchanged."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        actual_hash = source.source_hash

        # Use a wrong expected hash.
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash="0" * 64,
            )
        assert exc_info.value.context.get("reason") == "revision_conflict"

        # The source must be unchanged.
        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == actual_hash
        assert "Task018Marker" not in after.source

        _assert_owned_never_foreground(owned_pid, "vba-stale-rev")
        print(f"[vba-stale-rev] rejected without write, hash={actual_hash[:16]}…")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: insert at declarations
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_insert_declarations(tmp_path: Path) -> None:
    """Insert at declarations; the marker appears before the first procedure."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_insert_code")

        result = await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            _INSERT_COMMENT,
            position="declarations",
            expected_source_hash=expected_hash,
        )

        after = await rig.service.read_component(info.workbook_id, "modTest")
        insert_pos = after.source.find(_INSERT_COMMENT)
        add_pos = after.source.find("AddNumbers")
        assert insert_pos != -1, "inserted comment not found"
        assert add_pos != -1, "AddNumbers not found"
        assert insert_pos < add_pos, "inserted text must appear before the first procedure"

        assert result.before_hash == expected_hash
        assert result.after_hash != expected_hash

        _assert_owned_never_foreground(owned_pid, "vba-insert-decl")
        print(
            f"[vba-insert-decl] before={result.before_hash[:16]}… after={result.after_hash[:16]}…"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4: insert before a procedure
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_insert_before_procedure(tmp_path: Path) -> None:
    """Insert before a named procedure; the marker appears before its start."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_insert_code")

        result = await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            _INSERT_COMMENT,
            position="before_procedure:AddNumbers",
            expected_source_hash=expected_hash,
        )

        after = await rig.service.read_component(info.workbook_id, "modTest")
        insert_pos = after.source.find(_INSERT_COMMENT)
        add_pos = after.source.find("AddNumbers")
        assert insert_pos != -1, "inserted comment not found"
        assert add_pos != -1, "AddNumbers not found"
        assert insert_pos < add_pos, "inserted text must appear before AddNumbers"

        assert result.before_hash == expected_hash
        assert result.after_hash != expected_hash

        _assert_owned_never_foreground(owned_pid, "vba-insert-before")
        print(
            f"[vba-insert-before] before={result.before_hash[:16]}… after={result.after_hash[:16]}…"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 5: insert after a procedure
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_insert_after_procedure(tmp_path: Path) -> None:
    """Insert after a named procedure; the marker appears after its end."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_insert_code")

        result = await rig.service.insert_code(
            info.workbook_id,
            "modTest",
            _INSERT_COMMENT,
            position="after_procedure:AddNumbers",
            expected_source_hash=expected_hash,
        )

        after = await rig.service.read_component(info.workbook_id, "modTest")
        insert_pos = after.source.find(_INSERT_COMMENT)
        add_pos = after.source.find("AddNumbers")
        assert insert_pos != -1, "inserted comment not found"
        assert add_pos != -1, "AddNumbers not found"
        # The insert must appear after AddNumbers.
        assert insert_pos > add_pos, "inserted text must appear after AddNumbers"

        assert result.before_hash == expected_hash
        assert result.after_hash != expected_hash

        _assert_owned_never_foreground(owned_pid, "vba-insert-after")
        print(
            f"[vba-insert-after] before={result.before_hash[:16]}… after={result.after_hash[:16]}…"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 6: Document modules survive mutation
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_document_modules_survive_mutation(tmp_path: Path) -> None:
    """Document modules (ThisWorkbook / Sheet1) are still present after mutation."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_replace_component_source")

        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            _REPLACE_SOURCE,
            expected_source_hash=expected_hash,
        )

        # Document modules must still be present.
        components = await rig.service.list_components(info.workbook_id)
        by_name = {c.component_name: c for c in components}
        assert "ThisWorkbook" in by_name, "ThisWorkbook removed after mutation"
        assert "Sheet1" in by_name, "Sheet1 removed after mutation"
        assert by_name["ThisWorkbook"].component_type is VbaComponentType.DOCUMENT_MODULE
        assert by_name["Sheet1"].component_type is VbaComponentType.DOCUMENT_MODULE

        # UserForm must still be present too.
        assert "frmTest" in by_name, "frmTest removed after mutation"
        assert by_name["frmTest"].component_type is VbaComponentType.USER_FORM

        _assert_owned_never_foreground(owned_pid, "vba-doc-modules")
        print(
            f"[vba-doc-modules] ThisWorkbook + Sheet1 + frmTest present "
            f"after mutation (components={len(components)})"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 7: original fixture unchanged
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_original_fixture_unchanged(tmp_path: Path) -> None:
    """The committed fixture is byte-identical after all mutations."""

    committed_hash = _sha256(_CLEAN_PROJECT)

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)

        source = await rig.service.read_component(info.workbook_id, "modTest")
        expected_hash = source.source_hash

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, expected_hash, "vba_replace_component_source")

        await rig.service.replace_component_source(
            info.workbook_id,
            "modTest",
            _REPLACE_SOURCE,
            expected_source_hash=expected_hash,
        )

        # The committed fixture is untouched.
        assert _sha256(_CLEAN_PROJECT) == committed_hash

        _assert_owned_never_foreground(owned_pid, "vba-fixture-unchanged")
        print(f"[vba-fixture-unchanged] committed fixture sha256={committed_hash[:16]}…")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 8: read-only session rejected
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_read_only_session_rejected(tmp_path: Path) -> None:
    """A read-only session is rejected with WORKBOOK_READ_ONLY."""

    rig = _make_rig(tmp_path, writable=False)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert info.read_only is True

        await rig.snapshot.create_snapshot(info.workbook_id)
        source = await rig.service.read_component(info.workbook_id, "modTest")

        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=source.source_hash,
            )
        assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY

        # The source must be unchanged.
        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == source.source_hash

        _assert_owned_never_foreground(owned_pid, "vba-readonly")
        print(f"[vba-readonly] rejected with code={exc_info.value.code.value}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 9: snapshot required
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_snapshot_required(tmp_path: Path) -> None:
    """A mutation without a snapshot is blocked with SNAPSHOT_FAILED."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        # Do NOT create a snapshot.
        source = await rig.service.read_component(info.workbook_id, "modTest")

        with pytest.raises(SnapshotFailedError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=source.source_hash,
            )
        assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

        _assert_owned_never_foreground(owned_pid, "vba-no-snapshot")
        print(f"[vba-no-snapshot] blocked with code={exc_info.value.code.value}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 10: approval deny-by-default
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_approval_deny_by_default(tmp_path: Path) -> None:
    """A mutation without an approval grant is denied."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)
        source = await rig.service.read_component(info.workbook_id, "modTest")

        # Do NOT grant a mutation approval.
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "modTest",
                _REPLACE_SOURCE,
                expected_source_hash=source.source_hash,
            )
        assert (
            exc_info.value.context.get("reason")
            == "denied by default: no matching allowlist entry or live grant"
        )

        # The source must be unchanged.
        after = await rig.service.read_component(info.workbook_id, "modTest")
        assert after.source_hash == source.source_hash

        _assert_owned_never_foreground(owned_pid, "vba-no-approval")
        print(f"[vba-no-approval] denied: {exc_info.value.context.get('reason')}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 11: Document / UserForm rejection
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_document_module_rejected(tmp_path: Path) -> None:
    """Replacing a Document module is rejected with a clear error."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        await rig.snapshot.create_snapshot(info.workbook_id)
        source = await rig.service.read_component(info.workbook_id, "ThisWorkbook")

        working_path = _working_copy_path(rig, info.workbook_id)
        canonical_working = str(
            rig.path_policy.validate_path(working_path, OperationClass.MUTATING)
        )
        rig.grant_mutation(canonical_working, source.source_hash, "vba_replace_component_source")

        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.replace_component_source(
                info.workbook_id,
                "ThisWorkbook",
                "Option Explicit\n",
                expected_source_hash=source.source_hash,
            )
        assert exc_info.value.context.get("reason") == "non_replaceable_component"

        _assert_owned_never_foreground(owned_pid, "vba-doc-reject")
        print(f"[vba-doc-reject] reason={exc_info.value.context.get('reason')}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------
