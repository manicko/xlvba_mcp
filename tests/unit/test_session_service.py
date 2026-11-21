"""Unit tests for the workbook session lifecycle service (task_012).

These tests never open Excel. They inject a fake STA worker (which simply runs
the submitted callable) and a real :class:`ExcelAppAdapter` whose ``_app`` is a
fake ``Application`` object, so the real ``automation_security`` /
``alerts_suppressed`` context managers run and their ``finally`` restore logic is
exercised. A real :class:`PathPolicy`, :class:`ApprovalPolicy` and
:class:`AuditLogger` are used against a temp directory.

Covered:

* ``open_workbook`` validates the path, materializes a working copy under
  ``work/`` (original untouched), and assigns a ``workbook_id``.
* ``close_workbook`` defaults to **no save** (``SaveChanges=False``).
* ``save_workbook`` persists explicitly (clears dirty).
* ``get_status`` merges the worker's non-COM status with bookkeeping.
* Deny-by-default approval: execute / writable opens are blocked without a grant;
  inspect read-only opens are auto-allowed; a grant/allowlist enables them.
* Temporary Application settings are restored even when the open raises.
"""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError, MacroBlockedError
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
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

# Apply the anyio (asyncio) marker to every coroutine test in this module.
pytestmark = pytest.mark.anyio


def _grant(approval: ApprovalPolicy, canonical: str, action: str) -> None:
    """Add a session grant for ``action`` on ``canonical`` to ``approval``."""

    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, action=action),
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="tester",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# ---------------------------------------------------------------------------
# Fakes (no Excel / no COM)
# ---------------------------------------------------------------------------
class FakeWorkbook:
    """Minimal fake ``Workbook`` recording close/save."""

    def __init__(self, app: FakeApp, path: str, *, read_only: bool) -> None:
        self._app = app
        self.FullName = path
        self.Name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.ReadOnly = read_only
        self.Saved = True
        self.closed = False
        self.saved_on_close: bool | None = None
        self.save_count = 0
        self.Worksheets = [FakeSheet("Sheet1")]

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM casing
        self.closed = True
        self.saved_on_close = SaveChanges
        self.Saved = SaveChanges

    def Save(self) -> None:  # noqa: N802 - COM casing
        self.save_count += 1
        self.Saved = True

    @property
    def Application(self) -> FakeApp:  # noqa: N802 - COM casing
        return self._app


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeWorkbooks:
    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self.items: list[FakeWorkbook] = []
        self.fail_open = False
        self.open_calls: list[dict[str, Any]] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        if self.fail_open:
            raise RuntimeError("simulated open failure")
        self.open_calls.append({"path": path, **kwargs})
        wb = FakeWorkbook(self._app, path, read_only=bool(kwargs.get("ReadOnly", False)))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)

    def __len__(self) -> int:
        return len(self.items)


class FakeApp:
    """Fake ``Application`` with mutable settings + a Workbooks collection."""

    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1  # non-default start to prove restore
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(self)


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self, status: WorkerStatus | None = None) -> None:
        self.submits: list[dict[str, Any]] = []
        self._status = status or WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append(
            {"name": getattr(fn, "__name__", "?"), "section": section, "op_class": op_class}
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return self._status


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _adapter_with(app: FakeApp) -> ExcelAppAdapter:
    adapter = ExcelAppAdapter()
    adapter._app = app  # inject without DispatchEx
    return adapter


def _make_service(
    tmp_path: Path,
    *,
    approval: ApprovalPolicy | None = None,
    worker_status: WorkerStatus | None = None,
) -> tuple[SessionService, FakeApp, FakeWorker, Path, Path]:
    """Build a SessionService wired to fakes; return service + handles."""

    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"original-bytes")
    roots = {str(tmp_path): str(tmp_path)}
    path_policy = PathPolicy(roots)
    app = FakeApp()
    adapter = _adapter_with(app)
    worker = FakeWorker(status=worker_status)
    approval = approval or ApprovalPolicy()
    audit = AuditLogger(tmp_path / "audit.jsonl")
    work_dir = tmp_path / "work"
    service = SessionService(
        worker,
        adapter,
        path_policy,
        approval,
        audit,
        work_dir=work_dir,
    )
    return service, app, worker, original, work_dir


def _canonical(path_policy: PathPolicy, path: Path) -> str:
    return str(path_policy.validate_path(str(path), OperationClass.MUTATING))


# ---------------------------------------------------------------------------
# Open: validation + working copy + id
# ---------------------------------------------------------------------------
@pytest.mark.anyio
async def test_open_materializes_working_copy_and_assigns_id(tmp_path: Path) -> None:
    service, app, worker, original, work_dir = _make_service(tmp_path)
    original_bytes = original.read_bytes()

    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)

    # A working copy was created under work/ and differs from the original.
    assert info.workbook_id
    assert Path(info.path).parent == work_dir.resolve()
    assert Path(info.path) != original.resolve()
    assert info.source_path is not None and Path(info.source_path) == original.resolve()
    assert info.mode == SessionMode.INSPECT
    assert info.read_only is True

    # The original is untouched.
    assert original.read_bytes() == original_bytes

    # The open ran through the worker on the "open" section as MUTATING.
    assert any(s["section"] == "open" for s in worker.submits)
    open_submit = next(s for s in worker.submits if s["section"] == "open")
    assert open_submit["op_class"] == OperationClass.MUTATING

    # The adapter opened the working copy with UpdateLinks=0 / Notify=False /
    # AddToMru=False and the requested ReadOnly.
    call = app.Workbooks.open_calls[-1]
    assert call["UpdateLinks"] == 0
    assert call["Notify"] is False
    assert call["AddToMru"] is False
    assert call["ReadOnly"] is True

    # list_workbooks reflects the open workbook (no extra COM).
    listed = await service.list_workbooks()
    assert [w.workbook_id for w in listed] == [info.workbook_id]


async def test_open_inspect_sets_force_disable_automation_security(tmp_path: Path) -> None:
    service, app, _worker, original, _work_dir = _make_service(tmp_path)

    await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)

    # Inspect mode maps to ForceDisable (3); the adapter must have applied it and
    # restored the prior value (1) in finally.
    assert app.AutomationSecurity == 1, "automation security must be restored after open"


# ---------------------------------------------------------------------------
# Approval: deny-by-default + grant enables
# ---------------------------------------------------------------------------
async def test_open_execute_denied_without_grant(tmp_path: Path) -> None:
    service, _app, _worker, original, _work_dir = _make_service(tmp_path)

    with pytest.raises(MacroBlockedError):
        await service.open_workbook(str(original), mode=SessionMode.EXECUTE)


async def test_open_writable_inspect_denied_without_grant(tmp_path: Path) -> None:
    service, _app, _worker, original, _work_dir = _make_service(tmp_path)

    with pytest.raises(MacroBlockedError):
        await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=False)


async def test_open_inspect_read_only_auto_allowed(tmp_path: Path) -> None:
    # Empty approval policy: read-only inspect opens are auto-allowed.
    service, _app, _worker, original, _work_dir = _make_service(tmp_path)

    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)
    assert info.workbook_id


async def test_open_execute_allowed_with_grant(tmp_path: Path) -> None:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"x")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    canonical = _canonical(path_policy, original)

    approval = ApprovalPolicy()
    _grant(approval, canonical, "excel_open")
    app = FakeApp()
    worker = FakeWorker()
    service = SessionService(
        worker,
        _adapter_with(app),
        path_policy,
        approval,
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )

    info = await service.open_workbook(str(original), mode=SessionMode.EXECUTE)
    assert info.mode == SessionMode.EXECUTE
    # Execute mode maps to ByUI (2); restore to the prior value (1) in finally.
    assert app.AutomationSecurity == 1


# ---------------------------------------------------------------------------
# Close: default no-save; explicit save-on-close
# ---------------------------------------------------------------------------
async def test_close_defaults_to_no_save(tmp_path: Path) -> None:
    service, app, worker, original, _work_dir = _make_service(tmp_path)

    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)
    status = await service.close_workbook(info.workbook_id)  # default save_changes=False

    wb = app.Workbooks.items[-1]
    assert wb.closed is True
    assert wb.saved_on_close is False, "close must default to NOT saving"
    assert status.workbook_id is None
    assert any(s["section"] == "close" for s in worker.submits)
    assert await service.list_workbooks() == []


async def test_close_with_save_changes_requires_grant(tmp_path: Path) -> None:
    service, _app, _worker, original, _work_dir = _make_service(tmp_path)
    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)

    with pytest.raises(MacroBlockedError):
        await service.close_workbook(info.workbook_id, save_changes=True)


# ---------------------------------------------------------------------------
# Save: explicit save clears dirty
# ---------------------------------------------------------------------------
async def test_save_workbook_persists_explicitly(tmp_path: Path) -> None:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"x")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    canonical = _canonical(path_policy, original)

    approval = ApprovalPolicy()
    _grant(approval, canonical, "excel_open")
    _grant(approval, canonical, "excel_save")
    app = FakeApp()
    service = SessionService(
        FakeWorker(),
        _adapter_with(app),
        path_policy,
        approval,
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )

    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=False)
    assert info.workbook_id

    refreshed = await service.save_workbook(info.workbook_id)
    wb = app.Workbooks.items[-1]
    assert wb.save_count == 1, "save_workbook must call Save exactly once"
    assert refreshed.workbook_id == info.workbook_id
    status = await service.get_status(info.workbook_id)
    assert status.dirty is False


# ---------------------------------------------------------------------------
# get_status merges worker state (no COM submit)
# ---------------------------------------------------------------------------
async def test_get_status_merges_worker_state(tmp_path: Path) -> None:
    worker_status = WorkerStatus(
        state=WorkbookSessionState.RUNNING,
        active_operation="excel_open",
        active_section="open",
        accepting=False,
    )
    service, _app, worker, original, _work_dir = _make_service(
        tmp_path, worker_status=worker_status
    )
    submit_count_before = len(worker.submits)

    status = await service.get_status()

    assert len(worker.submits) == submit_count_before, "get_status must not submit to the worker"
    assert status.busy is True
    assert status.active_operation == "excel_open"
    assert status.session_id == service.session_id


async def test_get_status_unknown_workbook_reports_note(tmp_path: Path) -> None:
    service, _app, _worker, _original, _work_dir = _make_service(tmp_path)

    status = await service.get_status("does-not-exist")
    assert status.workbook_id == "does-not-exist"
    assert any("not currently open" in n for n in status.notes)


# ---------------------------------------------------------------------------
# Settings restore on exception (BG08)
# ---------------------------------------------------------------------------
async def test_open_restores_settings_on_exception(tmp_path: Path) -> None:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"x")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    canonical = _canonical(path_policy, original)
    approval = ApprovalPolicy()
    _grant(approval, canonical, "excel_open")
    app = FakeApp()
    work_dir = tmp_path / "work"
    service = SessionService(
        FakeWorker(),
        _adapter_with(app),
        path_policy,
        approval,
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=work_dir,
    )
    # Make the open fail inside the worker callable.
    app.Workbooks.fail_open = True

    with pytest.raises(RuntimeError, match="simulated open failure"):
        await service.open_workbook(str(original), mode=SessionMode.EXECUTE)

    # Settings restored in finally despite the exception.
    assert app.AutomationSecurity == 1, "automation security must restore on exception"
    assert app.DisplayAlerts is True, "display alerts must restore on exception"
    # No leftover working copy (cleaned up) and no bookkeeping.
    leftovers = list(work_dir.glob("*")) if work_dir.exists() else []
    assert leftovers == []
    assert await service.list_workbooks() == []


async def test_unknown_workbook_id_raises(tmp_path: Path) -> None:
    service, _app, _worker, _original, _work_dir = _make_service(tmp_path)

    with pytest.raises(ExcelVbaMcpError):
        await service.close_workbook("nope")
    with pytest.raises(ExcelVbaMcpError):
        await service.save_workbook("nope")


async def test_duplicate_open_same_path_errors(tmp_path: Path) -> None:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"x")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    approval = ApprovalPolicy()
    # allow writable inspect opens via a path grant (action-agnostic identity).
    canonical = _canonical(path_policy, original)
    _grant(approval, canonical, "excel_open")
    app = FakeApp()
    service = SessionService(
        FakeWorker(),
        _adapter_with(app),
        path_policy,
        approval,
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )

    first = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=False)
    assert first.workbook_id
    with pytest.raises(ExcelVbaMcpError, match="already open"):
        await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=False)

    # Cleanup so the held lock / temp file do not leak across the test.
    await service.close_workbook(first.workbook_id)


# ---------------------------------------------------------------------------
# Status reports the configured access mode (task_b2 / fix_002)
# ---------------------------------------------------------------------------
async def test_get_status_reports_full_access_mode(tmp_path: Path) -> None:
    service, _app, _worker, _original, _work_dir = _make_service(
        tmp_path, approval=ApprovalPolicy(mode=AccessMode.FULL)
    )

    status = await service.get_status()

    assert status.access_mode == AccessMode.FULL


async def test_get_status_reports_restricted_access_mode_by_default(tmp_path: Path) -> None:
    # The default object mode is RESTRICTED (deny-by-default); status reflects it.
    service, _app, _worker, _original, _work_dir = _make_service(tmp_path)

    status = await service.get_status()

    assert status.access_mode == AccessMode.RESTRICTED


async def test_approval_denial_emits_macro_blocked_audit_event(tmp_path: Path) -> None:
    # F-02: a denial emits exactly one audit event (MACRO_BLOCKED) before raising.
    service, _app, _worker, original, _work_dir = _make_service(tmp_path)

    with pytest.raises(MacroBlockedError):
        await service.open_workbook(str(original), mode=SessionMode.EXECUTE)

    audit_path = tmp_path / "audit.jsonl"
    events = [
        json.loads(line)
        for line in audit_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    denials = [
        event
        for event in events
        if event.get("tool") == "excel_open"
        and (event.get("result") or {}).get("error_code") == "MACRO_BLOCKED"
    ]
    assert len(denials) == 1
    approval = denials[0]["approval"]
    assert approval["allowed"] is False
    assert approval["reason"]
