"""Unit tests for Save As (``excel_save_as`` / ``SessionService.save_as_workbook``).

Work Package U7 (plan_05 §5.5, §8, §9). These tests never open Excel: they
inject a fake single-STA worker (runs the submitted callable inline) and a real
:class:`ExcelAppAdapter` whose ``_app`` is a fake ``Application`` — the same
pattern as ``test_session_service.py``. A real :class:`PathPolicy`,
:class:`ApprovalPolicy` (FULL, so mutating opens/saves auto-allow) and
:class:`AuditLogger` run against a real temp directory, so the pure-filesystem
atomic copy + single ``.bak`` + final size/SHA-256 verify are exercised for
real.

The 14-case decision matrix (plan_05 §9 / §8):

1. new path (target absent) -> write, outcome ok, no backup
2. save-as new path -> target bytes == working copy bytes
3. overwrite existing, unauthorized -> PROTECTED_TARGET/target_exists
4. overwrite existing, authorized -> ok + ``.bak`` of old target
5. target == source, no overwrite_source -> PROTECTED_TARGET/source_protected
6. target == source, overwrite_source + hash match -> source overwritten + ``.bak``
7. equivalent Windows path spelling/case -> recognised as the same target
8. external source modification -> STALE_TARGET/source_modified_since_open
9. path outside the allowlisted roots -> PATH_NOT_ALLOWED
10. permission/disk failure (pre-backup) -> outcome error, destination untouched
11. interruption (post-backup) -> outcome unknown + backup_path, no success
12. second save with an existing ``.bak`` -> deterministic PROTECTED_TARGET
    (backup_exists), ``.bak`` preserved (no auto-resume)
13. session identity unchanged + a later excel_save writes the working copy
    (target hash unchanged)
14. source replaced (re-canonical differs) -> STALE_TARGET/source_replaced
15. pre-existing ``.bak`` (prior interrupted save) -> deterministic
    PROTECTED_TARGET (backup_exists), not overwritten, dest untouched
"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.application import session_service as session_service_module
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import (
    ErrorCode,
    ExcelVbaMcpError,
    PathNotAllowedError,
    ProtectedTargetError,
    StaleTargetError,
)
from excel_vba_mcp.domain.models import SaveAsResult
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no Excel / no COM) — mirror test_session_service.py
# ---------------------------------------------------------------------------
class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeWorkbook:
    def __init__(self, app: FakeApp, path: str, *, read_only: bool) -> None:
        self._app = app
        self.FullName = path
        self.Name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.ReadOnly = read_only
        self.Saved = True
        self.save_count = 0
        self.Worksheets = [FakeSheet("Sheet1")]

    def Save(self) -> None:  # noqa: N802 - COM casing
        self.save_count += 1
        self.Saved = True

    @property
    def Application(self) -> FakeApp:  # noqa: N802 - COM casing
        return self._app


class FakeWorkbooks:
    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(self._app, path, read_only=bool(kwargs.get("ReadOnly", False)))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)

    def __len__(self) -> int:
        return len(self.items)


class FakeApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Workbooks = FakeWorkbooks(self)


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self) -> None:
        self.submits: list[dict[str, Any]] = []
        self._status = WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append({"name": getattr(fn, "__name__", "?"), "section": section})
        return fn()

    def get_status(self) -> WorkerStatus:
        return self._status


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_service(tmp_path: Path) -> tuple[SessionService, FakeApp, Path, Path]:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"original-working-bytes")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    app = FakeApp()
    adapter = ExcelAppAdapter()
    adapter._app = app  # inject without DispatchEx
    service = SessionService(
        FakeWorker(),
        adapter,
        path_policy,
        ApprovalPolicy(mode=AccessMode.FULL),
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )
    return service, app, original, path_policy


async def _open(service: SessionService, original: Path) -> str:
    info = await service.open_workbook(str(original), mode=SessionMode.EXECUTE, read_only=False)
    return info.workbook_id


# ---------------------------------------------------------------------------
# 1 + 2. New path
# ---------------------------------------------------------------------------
async def test_save_as_new_path_writes_file(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Copy.xlsm"
    original_hash = _sha256(original)

    result = await service.save_as_workbook(wid, str(target))

    assert isinstance(result, SaveAsResult)
    assert result.outcome == "ok"
    assert result.target_path == str(target.resolve())
    assert result.source_overwritten is False
    assert result.overwrote_existing is False
    assert result.backup_path is None
    assert target.is_file()
    # Target bytes equal the working copy (== the untouched original here).
    assert _sha256(target) == result.sha256
    assert target.stat().st_size == result.size
    # The original is never modified by Save As.
    assert _sha256(original) == original_hash


async def test_save_as_new_path_copies_working_copy_bytes(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    # Simulate an edit by mutating the on-disk working copy the server opened.
    working = Path(service.list_workbooks_sync()[0].path)
    working.write_bytes(b"edited-working-copy")
    submits_before = len(service._worker.submits)  # type: ignore[attr-defined]

    target = tmp_path / "Edited.xlsm"
    result = await service.save_as_workbook(wid, str(target))

    assert result.outcome == "ok"
    assert target.read_bytes() == b"edited-working-copy"
    assert result.sha256 == _sha256(target)
    # The flush ran on the STA worker ("save" section) before the copy.
    new_submits = service._worker.submits[submits_before:]  # type: ignore[attr-defined]
    assert any(s["section"] == "save" for s in new_submits)


# ---------------------------------------------------------------------------
# 3. Overwrite existing, unauthorized
# ---------------------------------------------------------------------------
async def test_overwrite_existing_unauthorized_rejected(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"
    target.write_bytes(b"pre-existing")

    with pytest.raises(ProtectedTargetError) as exc_info:
        await service.save_as_workbook(wid, str(target))
    assert exc_info.value.context.get("reason") == "target_exists"
    # The existing target is untouched.
    assert target.read_bytes() == b"pre-existing"


# ---------------------------------------------------------------------------
# 4. Overwrite existing, authorized (+ .bak)
# ---------------------------------------------------------------------------
async def test_overwrite_existing_authorized_creates_backup(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"
    target.write_bytes(b"old-target-contents")

    result = await service.save_as_workbook(wid, str(target), overwrite_existing=True)

    assert result.outcome == "ok"
    assert result.overwrote_existing is True
    assert result.backup_path == f"{target}.bak"
    assert Path(result.backup_path).read_bytes() == b"old-target-contents"
    assert target.read_bytes() == b"original-working-bytes"


# ---------------------------------------------------------------------------
# 5. Source without overwrite_source
# ---------------------------------------------------------------------------
async def test_source_without_overwrite_source_rejected(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    original_hash = _sha256(original)

    with pytest.raises(ProtectedTargetError) as exc_info:
        await service.save_as_workbook(wid, str(original))
    assert exc_info.value.context.get("reason") == "source_protected"
    assert _sha256(original) == original_hash


# ---------------------------------------------------------------------------
# 6. Source with overwrite_source + hash match
# ---------------------------------------------------------------------------
async def test_source_with_overwrite_source_hash_match(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)

    result = await service.save_as_workbook(wid, str(original), overwrite_source=True)

    assert result.outcome == "ok"
    assert result.source_overwritten is True
    assert result.backup_path == f"{original}.bak"
    # A single .bak of the source was written.
    assert Path(result.backup_path).read_bytes() == b"original-working-bytes"
    assert _sha256(original) == result.sha256


# ---------------------------------------------------------------------------
# 7. Equivalent Windows path spelling / case
# ---------------------------------------------------------------------------
async def test_equivalent_windows_path_case_recognised(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "CaseFile.xlsm"
    target.write_bytes(b"first-version")

    # First save with one spelling.
    first = await service.save_as_workbook(wid, str(target), overwrite_existing=True)
    assert first.outcome == "ok"

    # Drop the .bak so the second save is not blocked by the interrupted-save
    # guard (plan_05 §9); this test isolates path-spelling recognition.
    Path(first.backup_path).unlink()

    # Second save with a different-case spelling of the SAME path: normcase must
    # recognise it as the existing target (overwrite required, .bak refreshed).
    alt_spelling = str(tmp_path / "CASEFILE.xlsm")
    second = await service.save_as_workbook(wid, alt_spelling, overwrite_existing=True)
    assert second.outcome == "ok"
    assert second.overwrote_existing is True
    assert second.backup_path is not None


# ---------------------------------------------------------------------------
# 8. External source modification -> STALE_TARGET
# ---------------------------------------------------------------------------
async def test_external_source_modification_is_stale(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    # Modify the original on disk after open (external change).
    original.write_bytes(b"tampered-source")

    with pytest.raises(StaleTargetError) as exc_info:
        await service.save_as_workbook(wid, str(tmp_path / "Out.xlsm"))
    assert exc_info.value.context.get("reason") == "source_modified_since_open"
    assert exc_info.value.code is ErrorCode.STALE_TARGET
    # No output file was written.
    assert not (tmp_path / "Out.xlsm").exists()


# ---------------------------------------------------------------------------
# 9. Path outside roots -> PATH_NOT_ALLOWED
# ---------------------------------------------------------------------------
async def test_path_outside_roots_rejected(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    outside = tmp_path.parent / "outside-roots" / "Escape.xlsm"

    with pytest.raises(PathNotAllowedError):
        await service.save_as_workbook(wid, str(outside))


# ---------------------------------------------------------------------------
# 10. Permission / disk failure (pre-backup) -> outcome error, dest untouched
# ---------------------------------------------------------------------------
async def test_disk_failure_before_backup_is_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"
    target.write_bytes(b"keep-me")

    real_mkstemp = session_service_module.tempfile.mkstemp

    def boom(*args: Any, **kwargs: Any) -> Any:
        raise OSError("simulated disk failure")

    monkeypatch.setattr(session_service_module.tempfile, "mkstemp", boom)
    try:
        result = await service.save_as_workbook(wid, str(target), overwrite_existing=True)
    finally:
        monkeypatch.setattr(session_service_module.tempfile, "mkstemp", real_mkstemp)

    assert result.outcome == "error"
    assert result.backup_path is None
    # Destination untouched, no .bak written.
    assert target.read_bytes() == b"keep-me"
    assert not Path(f"{target}.bak").exists()


# ---------------------------------------------------------------------------
# 11. Interruption (post-backup) -> outcome unknown + backup_path
# ---------------------------------------------------------------------------
async def test_interruption_after_backup_is_unknown(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"
    target.write_bytes(b"old-target")

    real_replace = session_service_module.os.replace
    state = {"raised": False}

    def flaky_replace(src: Any, dst: Any) -> Any:
        # Raise only on the publish step (temp -> canonical); let the audit
        # atomic-write replace that follows succeed.
        if not state["raised"] and str(src).endswith(".tmp"):
            state["raised"] = True
            raise OSError("simulated interruption at publish")
        return real_replace(src, dst)

    monkeypatch.setattr(session_service_module.os, "replace", flaky_replace)
    result = await service.save_as_workbook(wid, str(target), overwrite_existing=True)

    assert result.outcome == "unknown"
    assert result.backup_path == f"{target}.bak"
    assert Path(result.backup_path).read_bytes() == b"old-target"
    # The destination was not replaced (still the old content).
    assert target.read_bytes() == b"old-target"


# ---------------------------------------------------------------------------
# 12. Second save with an existing .bak -> deterministic PROTECTED_TARGET
# ---------------------------------------------------------------------------
async def test_second_save_with_existing_backup_fails_closed(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"

    target.write_bytes(b"version-1")
    first = await service.save_as_workbook(wid, str(target), overwrite_existing=True)
    assert Path(first.backup_path).read_bytes() == b"version-1"
    assert target.read_bytes() == b"original-working-bytes"

    # Second save: a .bak already exists, so the save is rejected pre-mutation
    # with a deterministic PROTECTED_TARGET (backup_exists) error without
    # overwriting the .bak or replacing the destination (no auto-resume).
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await service.save_as_workbook(wid, str(target), overwrite_existing=True)
    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET
    assert exc_info.value.context.get("reason") == "backup_exists"
    assert exc_info.value.context.get("backup_path") == f"{target}.bak"
    # The pre-existing .bak and the destination are both untouched.
    assert Path(first.backup_path).read_bytes() == b"version-1"
    assert target.read_bytes() == b"original-working-bytes"


# ---------------------------------------------------------------------------
# 13. Session identity unchanged + later excel_save writes the working copy
# ---------------------------------------------------------------------------
async def test_session_identity_unchanged_and_later_save(tmp_path: Path) -> None:
    service, app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    info_before = service.list_workbooks_sync()[0]
    target = tmp_path / "Snapshot.xlsm"

    result = await service.save_as_workbook(wid, str(target))
    assert result.outcome == "ok"
    target_hash = _sha256(target)

    # Session identity (workbook_id / working copy / canonical) is unchanged.
    info_after = service.list_workbooks_sync()[0]
    assert info_after.workbook_id == info_before.workbook_id
    assert info_after.path == info_before.path
    assert info_after.source_path == info_before.source_path

    # A later excel_save writes the working copy; the Save As target is unchanged.
    refreshed = await service.save_workbook(wid)
    assert refreshed.workbook_id == wid
    assert _sha256(target) == target_hash
    # The flush (wb.Save) was invoked for the later save.
    wb = app.Workbooks.items[-1]
    assert wb.save_count >= 1


# ---------------------------------------------------------------------------
# 14. Source replaced (re-canonical differs) -> STALE_TARGET/source_replaced
# ---------------------------------------------------------------------------
async def test_source_replaced_is_stale(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    service, _app, original, path_policy = _make_service(tmp_path)
    wid = await _open(service, original)
    record = service._records[wid]
    real_validate = path_policy.validate_path

    def fake_validate(raw: Any, operation: Any, **kwargs: Any) -> Path:
        resolved = real_validate(raw, operation, **kwargs)
        is_source_revalidation = (
            OperationClass(operation) == OperationClass.READ_ONLY
            and os.path.normcase(str(resolved)) == os.path.normcase(record.canonical)
        )
        if is_source_revalidation:
            # Simulate the original having been replaced by a different path.
            return Path(str(resolved) + ".replaced")
        return resolved

    monkeypatch.setattr(path_policy, "validate_path", fake_validate)
    with pytest.raises(StaleTargetError) as exc_info:
        await service.save_as_workbook(wid, str(tmp_path / "Out.xlsm"))
    assert exc_info.value.context.get("reason") == "source_replaced"
    assert exc_info.value.code is ErrorCode.STALE_TARGET


# ---------------------------------------------------------------------------
# 15. Pre-existing .bak (prior interrupted save) -> deterministic PROTECTED_TARGET
# ---------------------------------------------------------------------------
async def test_preexisting_backup_is_not_overwritten(tmp_path: Path) -> None:
    service, _app, original, _pp = _make_service(tmp_path)
    wid = await _open(service, original)
    target = tmp_path / "Existing.xlsm"
    target.write_bytes(b"current-target")

    # A prior interrupted save left a .bak behind.
    backup = Path(f"{target}.bak")
    backup.write_bytes(b"prior-interrupted-backup")

    # Deterministic, actionable error (isError), not an indeterminate unknown:
    # the pre-existing .bak is a conflict detected before the destination is
    # replaced. The .bak is preserved verbatim and the destination is not replaced.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await service.save_as_workbook(wid, str(target), overwrite_existing=True)
    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET
    assert exc_info.value.context.get("reason") == "backup_exists"
    assert exc_info.value.context.get("backup_path") == f"{target}.bak"
    assert exc_info.value.context.get("target_path") == str(target.resolve())
    assert backup.read_bytes() == b"prior-interrupted-backup"
    assert target.read_bytes() == b"current-target"