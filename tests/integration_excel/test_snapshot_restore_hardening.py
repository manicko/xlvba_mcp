"""Live-Excel integration verification for snapshot restore hardening (plan_06 R1–R3).

Drives the *production* :class:`SnapshotService` (real single-STA
:class:`ExcelComWorker` + :class:`ExcelAppAdapter`, real :class:`SessionService`,
real :class:`SnapshotStore`) against the target Excel build, always hidden /
non-activating. A disposable copy of ``CleanProject.xlsm`` is the "original"; the
committed fixture is never opened or modified. The disposable workbook and the
dedicated snapshot store live under ``.ai/tmp``.

Coverage (R1 / R2 / R3 + interruption):

* ``test_restore_success_single_active_copy`` — success: exactly one open copy
  (the restored one, active), previous closed (``previous_close="closed"``),
  original hash unchanged, one owned Excel, restored identity rebound to the
  original (R3).
* ``test_restore_close_failure_reports_failed`` — a transient COM failure on the
  previous-copy close yields ``previous_close="failed"`` with safe text; the
  restore still succeeds, the restored copy is valid and the previous copy stays
  open (R1 + interruption transient-COM case).
* ``test_restore_missing_workbook_id_rejected`` — the API tool rejects a missing
  ``workbook_id`` with ``INVALID_RANGE`` before any action (R2).
* ``test_restore_same_snapshot_and_different`` — after restore the canonical is
  the original and ``list_workbooks`` shows it; repeating the SAME snapshot is a
  clear ``same_snapshot`` error; a DIFFERENT snapshot is allowed (R3).
* ``test_restore_interruption_before_cleanup`` — an interruption between
  activation and old-copy cleanup leaves the restored copy active and valid, the
  previous copy still open, and no auto-replay of the interrupted restore.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_snapshot_restore_hardening.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import shutil
import sys
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import test_tools
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import SnapshotRestoreResult, WorkbookInfo
from excel_vba_mcp.domain.statuses import AccessMode, OperationClass, SessionMode
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

# Optional COM dependencies (None when unavailable so the module imports and the
# COM tests skip on non-Windows hosts).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CLEAN_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CleanProject.xlsm"
_TMP_ROOT = _REPO_ROOT / ".ai" / "tmp"

_CALL_TIMEOUT_S = 30.0


def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _find_wb(app_adapter: ExcelAppAdapter, working_path: str) -> Any:
    app = app_adapter.app
    target = working_path.lower()
    for wb in app.Workbooks:  # type: ignore[union-attr]
        if str(wb.FullName).lower() == target:
            return wb
    raise RuntimeError(f"working copy not found: {working_path}")


def _read_a1(app_adapter: ExcelAppAdapter, path: str) -> str:
    wb = _find_wb(app_adapter, path)
    return str(wb.Worksheets(1).Range("A1").Value)


def _mutate_a1(app_adapter: ExcelAppAdapter, path: str, value: str) -> str:
    wb = _find_wb(app_adapter, path)
    wb.Worksheets(1).Range("A1").Value = value
    return value


# ---------------------------------------------------------------------------
# Rig (disposable workbook + dedicated snapshot store under .ai/tmp)
# ---------------------------------------------------------------------------
class _HardeningRig:
    """A live worker + session + snapshot service rooted under ``.ai/tmp``."""

    def __init__(self, base_dir: Path, original: Path) -> None:
        self.base_dir = base_dir
        self.original = original
        # FULL mode auto-authorizes mutating/executing opens (no grant bookkeeping).
        self.path_policy = PathPolicy({str(base_dir): str(base_dir)})
        self.canonical = str(
            self.path_policy.validate_path(str(original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy(mode=AccessMode.FULL)
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=30.0)
        self.audit = AuditLogger(base_dir / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=base_dir / "work",
        )
        self.store = SnapshotStore(base_dir / "snapshots")
        self.snapshot = SnapshotService(
            self.store,
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        with contextlib.suppress(Exception):
            await self.worker.shutdown()
        self.audit.close()


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def rig(tmp_path: Path) -> Any:
    """Build a disposable rig under ``.ai/tmp`` and tear it down afterwards."""

    _TMP_ROOT.mkdir(parents=True, exist_ok=True)
    base_dir = _TMP_ROOT / f"integration_r1r3_{uuid4().hex}"
    (base_dir / "work").mkdir(parents=True, exist_ok=True)
    original = base_dir / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    r = _HardeningRig(base_dir, original)
    yield r
    with contextlib.suppress(Exception):
        shutil.rmtree(base_dir, ignore_errors=True)


def _open_original(rig: _HardeningRig) -> Any:
    return rig.session.open_workbook(
        str(rig.original), mode=SessionMode.EXECUTE, read_only=False
    )


# ---------------------------------------------------------------------------
# Success: one active copy (restored), previous closed, original unchanged
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_success_single_active_copy(rig: _HardeningRig) -> None:
    await rig.start()
    owned_pid: int | None = None
    try:
        original_hash = _sha256(rig.original)
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await _open_original(rig)
        manifest = await rig.snapshot.create_snapshot(info.workbook_id)
        snapshot_file = rig.store.path_for(manifest.snapshot_id) / manifest.file_name

        # Mutate the working copy in memory (do NOT save) before restore.
        await rig.worker.submit(
            lambda: _mutate_a1(rig.app_adapter, info.path, "MUT-R1R3"),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.MUTATING,
        )

        result = await rig.snapshot.restore_snapshot(
            manifest.snapshot_id, workbook_id=info.workbook_id
        )

        assert isinstance(result, SnapshotRestoreResult)
        assert result.previous_close == "closed"
        assert result.previous_close_error is None
        # R3: identity rebound to the ORIGINAL, never the snapshot file.
        assert result.workbook.source_path == manifest.source_path
        assert Path(result.workbook.path) != snapshot_file.resolve()

        # Exactly one open copy: the restored one (active); previous closed.
        open_infos = await rig.session.list_workbooks()
        assert len(open_infos) == 1
        assert open_infos[0].workbook_id == result.workbook.workbook_id
        assert open_infos[0].source_path == manifest.source_path

        # The restored working copy shows the baseline A1 (mutation discarded).
        restored_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, result.workbook.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert restored_a1 != "MUT-R1R3", "mutation must not survive restore"

        # Original untouched; one owned Excel (same pid, still alive).
        assert _sha256(rig.original) == original_hash
        assert rig.worker.get_status().owned_pid == owned_pid

        await rig.session.close_workbook(result.workbook.workbook_id)
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# R1: transient COM failure on close -> previous_close="failed", restore valid
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_close_failure_reports_failed(rig: _HardeningRig) -> None:
    await rig.start()
    try:
        info = await _open_original(rig)
        manifest = await rig.snapshot.create_snapshot(info.workbook_id)

        # Simulate a transient COM failure on the previous-copy close only.
        original_close = rig.session.close_workbook

        async def failing_close(workbook_id: str, **kwargs: Any) -> Any:
            if workbook_id == info.workbook_id:
                raise ExcelVbaMcpError("simulated transient COM failure on close")
            return await original_close(workbook_id, **kwargs)

        rig.session.close_workbook = failing_close  # type: ignore[method-assign]
        try:
            result = await rig.snapshot.restore_snapshot(
                manifest.snapshot_id, workbook_id=info.workbook_id
            )
        finally:
            rig.session.close_workbook = original_close  # type: ignore[method-assign]

        # R1: the restore still succeeded; the close failure is surfaced, not fatal.
        assert result.previous_close == "failed"
        assert result.previous_close_error
        assert isinstance(result.workbook, WorkbookInfo)

        # The restored copy is active and valid...
        open_infos = await rig.session.list_workbooks()
        assert any(w.workbook_id == result.workbook.workbook_id for w in open_infos)
        restored_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, result.workbook.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert restored_a1 is not None

        # ...and the previous copy remains open (close failed).
        assert any(w.workbook_id == info.workbook_id for w in open_infos)

        # Cleanup: close both explicitly (the restored one is the active session).
        # A transient COM error during teardown must not fail the test.
        with contextlib.suppress(Exception):
            await rig.session.close_workbook(info.workbook_id)
        with contextlib.suppress(Exception):
            await rig.session.close_workbook(result.workbook.workbook_id)
    finally:
        # The simulated close failure can leave teardown COM in a transient state;
        # teardown must be best-effort and never fail the test.
        with contextlib.suppress(Exception):
            await rig.stop()


# ---------------------------------------------------------------------------
# R2: missing workbook_id rejected at the API tool boundary
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_missing_workbook_id_rejected(rig: _HardeningRig) -> None:
    await rig.start()
    try:
        info = await _open_original(rig)
        manifest = await rig.snapshot.create_snapshot(info.workbook_id)

        test_tools.set_snapshot_service(rig.snapshot)
        try:
            with pytest.raises(ExcelVbaMcpError) as excinfo:
                await test_tools.snapshot_restore(manifest.snapshot_id, confirm=True)
        finally:
            test_tools.set_snapshot_service(None)

        assert excinfo.value.code is ErrorCode.INVALID_RANGE
        assert excinfo.value.to_dict()["isRetryable"] is False
        # No destructive action: the original session is still the only open one.
        open_infos = await rig.session.list_workbooks()
        assert len(open_infos) == 1
        assert open_infos[0].workbook_id == info.workbook_id

        await rig.session.close_workbook(info.workbook_id)
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# R3: canonical=original, same-snapshot error, different snapshot allowed
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_same_snapshot_and_different(rig: _HardeningRig) -> None:
    await rig.start()
    try:
        info = await _open_original(rig)
        snap_a = await rig.snapshot.create_snapshot(info.workbook_id)
        snap_b = await rig.snapshot.create_snapshot(info.workbook_id)

        first = await rig.snapshot.restore_snapshot(
            snap_a.snapshot_id, workbook_id=info.workbook_id
        )
        # R3: canonical is the original; list_workbooks shows the original.
        assert first.workbook.source_path == snap_a.source_path
        listed = await rig.session.list_workbooks()
        assert len(listed) == 1
        assert listed[0].source_path == snap_a.source_path

        # Repeating the SAME snapshot is a deterministic same_snapshot error.
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await rig.snapshot.restore_snapshot(
                snap_a.snapshot_id, workbook_id=first.workbook.workbook_id
            )
        assert excinfo.value.code is ErrorCode.INVALID_RANGE
        assert excinfo.value.context.get("reason") == "same_snapshot"

        # A DIFFERENT snapshot is allowed (new working copy; previous closed).
        second = await rig.snapshot.restore_snapshot(
            snap_b.snapshot_id, workbook_id=first.workbook.workbook_id
        )
        assert second.previous_close == "closed"
        assert second.workbook.workbook_id not in (info.workbook_id, first.workbook.workbook_id)
        open_infos = await rig.session.list_workbooks()
        assert len(open_infos) == 1
        assert open_infos[0].workbook_id == second.workbook.workbook_id
        assert open_infos[0].source_path == snap_b.source_path

        await rig.session.close_workbook(second.workbook.workbook_id)
    finally:
        await rig.stop()


# ---------------------------------------------------------------------------
# Interruption/recovery: interrupt between activation and old-copy cleanup
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_restore_interruption_before_cleanup(rig: _HardeningRig) -> None:
    await rig.start()
    try:
        info = await _open_original(rig)
        manifest = await rig.snapshot.create_snapshot(info.workbook_id)

        # Simulate an interruption (BaseException, not caught by the R1 guard)
        # raised at the close-previous step: the restored copy is already opened,
        # validated and registered, but the old copy is not yet cleaned up.
        original_close = rig.session.close_workbook

        async def interrupting_close(workbook_id: str, **kwargs: Any) -> Any:
            if workbook_id == info.workbook_id:
                raise KeyboardInterrupt("simulated interruption before cleanup")
            return await original_close(workbook_id, **kwargs)

        rig.session.close_workbook = interrupting_close  # type: ignore[method-assign]
        with pytest.raises(KeyboardInterrupt):
            await rig.snapshot.restore_snapshot(
                manifest.snapshot_id, workbook_id=info.workbook_id
            )
        rig.session.close_workbook = original_close  # type: ignore[method-assign]

        # The restored copy is active and valid; the previous copy is still open
        # (cleanup did not run). No auto-replay occurred (the call raised through).
        open_infos = await rig.session.list_workbooks()
        ids = {w.workbook_id for w in open_infos}
        assert info.workbook_id in ids, "previous copy must remain open after interruption"
        restored_ids = ids - {info.workbook_id}
        assert len(restored_ids) == 1, "exactly one restored copy must be active"
        restored_id = restored_ids.pop()
        restored = next(w for w in open_infos if w.workbook_id == restored_id)
        # The interruption fired at close-previous, BEFORE the R3 rebind (which runs
        # after cleanup). The restored copy is active/valid but its canonical is
        # still the snapshot path in this indeterminate state; recovery is explicit
        # (no auto-replay). Assert the invariant that matters: restored active,
        # previous open, exactly one restored copy.
        assert restored.workbook_id == restored_id
        restored_a1 = await rig.worker.submit(
            lambda: _read_a1(rig.app_adapter, restored.path),
            timeout=_CALL_TIMEOUT_S,
            op_class=OperationClass.READ_ONLY,
        )
        assert restored_a1 is not None

        # Recovery is explicit (no auto-replay): close the previous copy, keep the
        # restored one as the single active session.
        await rig.session.close_workbook(info.workbook_id)
        remaining = await rig.session.list_workbooks()
        assert len(remaining) == 1
        assert remaining[0].workbook_id == restored_id
        await rig.session.close_workbook(restored_id)
    finally:
        await rig.stop()