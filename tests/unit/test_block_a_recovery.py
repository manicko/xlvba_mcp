"""Unit tests for Block A — Recovery (held handle, excel_recover, fail-closed).

No Excel / COM is launched. The handle layer is faked (``_terminate_and_confirm``
is patched) and the Win32 identity/liveness probes are injected, so the suite runs
on any host. Coverage (plan_02 Work item A; PoC P2/P7):

* same-handle terminate + confirm (``TerminateProcess`` -> ``WAIT_OBJECT_0``);
* identity-failure (Hwnd=0 / OpenProcess fail / image mismatch) => refuse, no kill;
* the ``PyHANDLE`` is retained as a **live object** (never reduced to ``int``);
* ``excel_recover`` is ALLOWED by default in RESTRICTED, a MUTATING op is DENIED;
  recovery is audited and pid-scoped (never a wildcard);
* startup reconcile: a live prior owned pid => no second launch; a dead one =>
  proceed.
"""

from __future__ import annotations

import json
import sys
import types
from collections.abc import Iterator
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel import app_adapter as app_adapter_module
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.recovery_tools import RecoveryResult, excel_recover
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.server import _reconcile_prior_owned_pid, _start_worker_guarded
from excel_vba_mcp.storage.owned_pid_store import (
    STATE_CREATING,
    OwnedPidStore,
    OwnedProcessRecord,
    build_creating_record,
    build_owned_record,
)

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Handle terminate + confirm via the SAME handle (P2)
# ---------------------------------------------------------------------------
class _FakeHandle:
    """A live, non-int handle object (stands in for a ``PyHANDLE``)."""


def test_terminate_owned_handle_confirms_on_wait_object_0(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    adapter = ExcelAppAdapter()
    handle = _FakeHandle()
    adapter._owned_handle = handle
    seen: list[tuple[object, int]] = []

    def fake_terminate(h: object, timeout_ms: int) -> int:
        seen.append((h, timeout_ms))
        return 0  # WAIT_OBJECT_0

    monkeypatch.setattr(app_adapter_module, "_terminate_and_confirm", fake_terminate)

    assert adapter.terminate_owned_handle() is True
    # The SAME held handle object was terminated (never an int / re-resolved pid).
    assert seen and seen[0][0] is handle


def test_terminate_owned_handle_not_confirmed_when_wait_times_out(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    adapter = ExcelAppAdapter()
    adapter._owned_handle = _FakeHandle()
    # A non-zero wait result (e.g. WAIT_TIMEOUT 0x102) is NOT a confirmed death.
    monkeypatch.setattr(app_adapter_module, "_terminate_and_confirm", lambda _h, _t: 0x102)

    assert adapter.terminate_owned_handle() is False


# ---------------------------------------------------------------------------
# Fail-closed: identity / handle failure => refuse, no terminate
# ---------------------------------------------------------------------------
def test_terminate_refuses_when_no_handle(monkeypatch: pytest.MonkeyPatch) -> None:
    adapter = ExcelAppAdapter()
    adapter._owned_handle = None
    called = False

    def fail(_h: object, _t: int) -> int:
        nonlocal called
        called = True
        return 0

    monkeypatch.setattr(app_adapter_module, "_terminate_and_confirm", fail)

    assert adapter.terminate_owned_handle() is False
    assert adapter.recovery_available is False
    assert called is False, "no handle => refuse; the Win32 terminate must not run"


def test_terminate_refuses_when_handle_reduced_to_int(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """P7: reducing the handle to ``int`` breaks liveness — refuse (fail closed)."""

    adapter = ExcelAppAdapter()
    adapter._owned_handle = 123  # a reduced int, not a live object
    called = False

    def fail(_h: object, _t: int) -> int:
        nonlocal called
        called = True
        return 0

    monkeypatch.setattr(app_adapter_module, "_terminate_and_confirm", fail)

    assert adapter.terminate_owned_handle() is False
    assert adapter.recovery_available is False
    assert called is False


def test_open_owned_handle_fails_closed_on_identity_mismatch(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # Image name is not EXCEL.EXE => no handle, recovery unavailable.
    monkeypatch.setattr(
        ExcelAppAdapter, "_image_name_of_pid", staticmethod(lambda _pid: r"C:\Windows\NOTEPAD.EXE")
    )
    assert ExcelAppAdapter._open_owned_handle(4242) is None


def test_open_owned_handle_fails_closed_when_openprocess_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(ExcelAppAdapter, "_pid_is_excel", staticmethod(lambda _pid: True))
    import win32api  # pywin32 is present in the venv

    def deny(*_args: Any, **_kwargs: Any) -> Any:
        raise OSError("access denied")

    monkeypatch.setattr(win32api, "OpenProcess", deny)
    assert ExcelAppAdapter._open_owned_handle(4242) is None


def test_open_owned_handle_fails_closed_without_pid() -> None:
    assert ExcelAppAdapter._open_owned_handle(None) is None
    assert ExcelAppAdapter._open_owned_handle(0) is None


def test_pid_is_excel_true_only_for_excel(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        ExcelAppAdapter,
        "_image_name_of_pid",
        staticmethod(lambda _pid: r"C:\Program Files\Microsoft Office\EXCEL.EXE"),
    )
    assert ExcelAppAdapter._pid_is_excel(1) is True
    monkeypatch.setattr(
        ExcelAppAdapter, "_image_name_of_pid", staticmethod(lambda _pid: r"C:\x\winword.exe")
    )
    assert ExcelAppAdapter._pid_is_excel(1) is False


# ---------------------------------------------------------------------------
# PyHANDLE retained as a live object (P7)
# ---------------------------------------------------------------------------
def test_create_hidden_retains_handle_object(monkeypatch: pytest.MonkeyPatch) -> None:
    """``create_hidden`` stores the live handle object (not an int)."""

    adapter = ExcelAppAdapter()
    handle = _FakeHandle()

    class _FakeDispatch:
        Visible = False
        Caption = "Excel"
        Hwnd = 4242

        def __setattr__(self, name: str, value: Any) -> None:
            object.__setattr__(self, name, value)

    import win32com.client  # noqa: F401 - ensure the lazy import target exists

    monkeypatch.setattr("win32com.client.DispatchEx", lambda _prog: _FakeDispatch(), raising=False)
    monkeypatch.setattr(ExcelAppAdapter, "_pid_from_hwnd", staticmethod(lambda _h: 4242))
    monkeypatch.setattr(ExcelAppAdapter, "_open_owned_handle", staticmethod(lambda _pid: handle))
    # No job layer in this unit path (job behaviour is covered separately).
    monkeypatch.setattr(ExcelAppAdapter, "_create_job_object", staticmethod(lambda: None))

    adapter.create_hidden()

    assert adapter.owned_handle is handle
    assert not isinstance(adapter.owned_handle, int)
    assert adapter.recovery_available is True
    assert adapter.job_handle is None
    assert adapter.job_protected is False


# ---------------------------------------------------------------------------
# Job object (kernel KILL_ON_JOB_CLOSE guarantee): created before DispatchEx,
# assign+verify called, fail-closed on assign failure, handle held (not int)
# ---------------------------------------------------------------------------
class _FakeJob:
    """A live, non-int job handle recording Close()."""

    def __init__(self) -> None:
        self.closed = False

    def Close(self) -> None:  # noqa: N802 - PyHANDLE casing
        self.closed = True


def _patch_create_hidden(
    monkeypatch: pytest.MonkeyPatch,
    *,
    job: Any,
    assign_result: bool,
    handle: Any = None,
) -> list[tuple[Any, int | None]]:
    """Patch create_hidden's collaborators; return the recorded assign calls."""

    if handle is None:
        handle = _FakeHandle()

    class _FakeDispatch:
        Visible = False
        Caption = "Excel"
        Hwnd = 4242

        def __setattr__(self, name: str, value: Any) -> None:
            object.__setattr__(self, name, value)

    import win32com.client  # noqa: F401

    assign_calls: list[tuple[Any, int | None]] = []

    def _assign(j: Any, pid: int | None) -> bool:
        assign_calls.append((j, pid))
        return assign_result

    monkeypatch.setattr("win32com.client.DispatchEx", lambda _p: _FakeDispatch(), raising=False)
    monkeypatch.setattr(ExcelAppAdapter, "_pid_from_hwnd", staticmethod(lambda _h: 4242))
    monkeypatch.setattr(ExcelAppAdapter, "_open_owned_handle", staticmethod(lambda _p: handle))
    monkeypatch.setattr(ExcelAppAdapter, "_create_job_object", staticmethod(lambda: job))
    monkeypatch.setattr(ExcelAppAdapter, "_assign_process_to_job", staticmethod(_assign))
    return assign_calls


def test_create_hidden_creates_job_before_dispatch_and_assigns(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The job is created, the owned pid is assigned into it, and it is held (not int)."""

    adapter = ExcelAppAdapter()
    job = _FakeJob()
    order: list[str] = []

    def _create_job() -> Any:
        order.append("create_job")
        return job

    class _FakeDispatch:
        Visible = False
        Caption = "Excel"
        Hwnd = 4242

        def __setattr__(self, name: str, value: Any) -> None:
            object.__setattr__(self, name, value)

    import win32com.client  # noqa: F401

    def _dispatch(_p: str) -> Any:
        order.append("dispatch")
        return _FakeDispatch()

    def _assign(j: Any, pid: int | None) -> bool:
        order.append("assign")
        return True

    monkeypatch.setattr("win32com.client.DispatchEx", _dispatch, raising=False)
    monkeypatch.setattr(ExcelAppAdapter, "_pid_from_hwnd", staticmethod(lambda _h: 4242))
    monkeypatch.setattr(
        ExcelAppAdapter, "_open_owned_handle", staticmethod(lambda _p: _FakeHandle())
    )
    monkeypatch.setattr(ExcelAppAdapter, "_create_job_object", staticmethod(_create_job))
    monkeypatch.setattr(ExcelAppAdapter, "_assign_process_to_job", staticmethod(_assign))

    adapter.create_hidden()

    # Job created BEFORE DispatchEx (shrink the create->assign window).
    assert order == ["create_job", "dispatch", "assign"], order
    assert adapter.job_handle is job
    assert adapter.job_protected is True
    assert not isinstance(adapter.job_handle, int)


def test_create_hidden_assigns_the_owned_pid_into_the_job(monkeypatch: pytest.MonkeyPatch) -> None:
    adapter = ExcelAppAdapter()
    job = _FakeJob()
    calls = _patch_create_hidden(monkeypatch, job=job, assign_result=True)

    adapter.create_hidden()

    assert calls == [(job, 4242)], "the owned pid is assigned into the held job"


def test_create_hidden_fails_closed_when_assign_fails(monkeypatch: pytest.MonkeyPatch) -> None:
    """Assign failure => the Excel is terminated (image-gated) and NOT handed over."""

    adapter = ExcelAppAdapter()
    job = _FakeJob()
    handle = _FakeHandle()
    terminated: list[object] = []
    _patch_create_hidden(monkeypatch, job=job, assign_result=False, handle=handle)
    monkeypatch.setattr(
        app_adapter_module,
        "_terminate_and_confirm",
        lambda h, _t: terminated.append(h) or 0,  # WAIT_OBJECT_0
    )

    with pytest.raises(RuntimeError, match="fail-closed"):
        adapter.create_hidden()

    assert terminated == [handle], "the owned Excel is terminated via the held handle"
    assert job.closed is True, "the job handle is closed on the fail-closed path"
    assert adapter.app is None, "the Excel is NOT handed over"
    assert adapter.owned_pid is None
    assert adapter.job_handle is None
    assert adapter.recovery_available is False


def test_assign_process_to_job_fails_closed_on_bad_input() -> None:
    """No job / no pid => refuse (False), never a partial assign."""

    assert ExcelAppAdapter._assign_process_to_job(None, 4242) is False
    assert ExcelAppAdapter._assign_process_to_job(_FakeJob(), None) is False
    assert ExcelAppAdapter._assign_process_to_job(_FakeJob(), 0) is False
    assert ExcelAppAdapter._assign_process_to_job(123, 4242) is False  # int handle rejected


def test_create_job_object_is_non_inheritable_on_windows() -> None:
    """On a host with the Win32 job layer the created job handle is non-inheritable."""

    try:
        import win32api
        import win32job  # noqa: F401
    except Exception:  # noqa: BLE001 - non-Windows / no pywin32
        pytest.skip("win32job unavailable")

    job = ExcelAppAdapter._create_job_object()
    if job is None:
        pytest.skip("job objects unavailable on this host")
    try:
        assert not isinstance(job, int)
        assert (win32api.GetHandleInformation(job) & 1) == 0, "job handle must be non-inheritable"
    finally:
        ExcelAppAdapter._close_job_handle(job)


# ---------------------------------------------------------------------------
# excel_recover policy: RECOVERY allowed in RESTRICTED, MUTATING denied
# ---------------------------------------------------------------------------
def test_recovery_allowed_in_restricted_but_mutating_denied() -> None:
    policy = ApprovalPolicy(mode=AccessMode.RESTRICTED)
    recovery = policy.check_recovery(ApprovalKey(canonical_path="pid:4242", action="excel_recover"))
    assert recovery.allowed is True
    assert recovery.source is not None

    mutating = policy.check(
        ApprovalKey(canonical_path="C:/work/Book.xlsm", action="sheet_delete"),
        operation_class=OperationClass.MUTATING,
    )
    assert mutating.allowed is False


def test_recovery_allowed_in_full() -> None:
    policy = ApprovalPolicy(mode=AccessMode.FULL)
    recovery = policy.check_recovery(ApprovalKey(canonical_path="pid:1", action="excel_recover"))
    assert recovery.allowed is True


# ---------------------------------------------------------------------------
# excel_recover tool: fail-closed + audited + pid-scoped, on a fake session
# ---------------------------------------------------------------------------
class _FakeWorker:
    def __init__(self, adapter: ExcelAppAdapter) -> None:
        self.adapter = adapter
        self.state = WorkbookSessionState.RUNNING
        self.accepting = True
        self.owned_pid = 4242
        self.notes: list[str] = []
        self.active_operation: str | None = None
        self.reset_count = 0

    def get_status(self) -> Any:
        from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus

        return WorkerStatus(
            state=self.state,
            accepting=self.accepting,
            owned_pid=self.owned_pid,
            notes=list(self.notes),
            active_operation=self.active_operation,
        )

    async def reset(self) -> bool:
        self.reset_count += 1
        self.state = WorkbookSessionState.RUNNING
        self.accepting = True
        self.owned_pid = 9999
        self.adapter._owned_pid = 9999
        self.adapter._owned_handle = _FakeHandle()
        return True


class _Rig:
    def __init__(self, tmp_path: Path) -> None:
        self.adapter = ExcelAppAdapter()
        self.adapter._owned_pid = 4242
        self.adapter._owned_handle = _FakeHandle()
        self.worker = _FakeWorker(self.adapter)
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            PathPolicy({str(tmp_path): str(tmp_path)}),
            ApprovalPolicy(mode=AccessMode.RESTRICTED),
            self.audit,
            work_dir=tmp_path / "work",
        )


@pytest.fixture
def rig(tmp_path: Path) -> Iterator[_Rig]:
    instance = _Rig(tmp_path)
    yield instance
    instance.audit.close()
    workbook_tools.set_session_service(None)


async def test_excel_recover_terminates_and_rebuilds(rig: _Rig) -> None:
    terminated: list[object] = []
    monkey_handle = rig.adapter._owned_handle
    rig.adapter.terminate_owned_handle = lambda: terminated.append(monkey_handle) or True  # type: ignore[method-assign]

    workbook_tools.set_session_service(rig.session)  # type: ignore[arg-type]
    result = await excel_recover()

    assert isinstance(result, RecoveryResult)
    assert result.ok is True
    assert result.terminated is True
    assert result.confirmed is True
    assert result.recovery_possible is True
    assert result.process_restarted is True
    assert terminated == [monkey_handle]
    assert rig.worker.reset_count == 1
    # Originals-safe wording is always present.
    assert "originals are never modified" in result.what_lost


async def test_excel_recover_fails_closed_without_handle(rig: _Rig) -> None:
    rig.adapter._owned_handle = None  # identity/handle unavailable

    workbook_tools.set_session_service(rig.session)  # type: ignore[arg-type]
    result = await excel_recover()

    assert result.ok is False
    assert result.recovery_possible is False
    assert result.terminated is False
    assert result.confirmed is False
    assert result.process_restarted is False
    assert rig.worker.reset_count == 0, "fail closed: no reset / no second launch"
    assert "server restart" in (result.reason or "")


async def test_excel_recover_marks_prior_job_unknown_when_active(rig: _Rig) -> None:
    rig.worker.active_operation = "macro_run"
    workbook_tools.set_session_service(rig.session)  # type: ignore[arg-type]

    result = await excel_recover()

    assert result.prior_job == "unknown"


async def test_excel_recover_is_audited(rig: _Rig) -> None:
    workbook_tools.set_session_service(rig.session)  # type: ignore[arg-type]
    await excel_recover()

    events = [
        json.loads(line)
        for line in (rig.audit.path).read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    recover_events = [e for e in events if e.get("tool") == "excel_recover"]
    assert recover_events, "excel_recover must be audited"
    event = recover_events[-1]
    # pid-scoped: the event records the owned pid, never a workbook wildcard/path.
    assert event["operation_class"] == "recovery"
    assert event["params"]["owned_pid"] == 4242
    params_blob = json.dumps(event["params"])
    assert "*" not in params_blob, "recovery must be pid-scoped, never a wildcard"
    assert ".xlsm" not in params_blob.lower()
    assert "\\" not in params_blob and "/" not in params_blob


# ---------------------------------------------------------------------------
# Startup reconcile: identity-based (owner alive+match => refuse; owner dead +
# verified orphan => reap; pid reuse => never kill on pid-alone)
# ---------------------------------------------------------------------------
_EXCEL_EXE = r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
_PY_EXE = r"C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe"


def _write_record(
    store: OwnedPidStore,
    *,
    owner_pid: int,
    owner_ct: int | None,
    excel_pid: int,
    excel_ct: int | None,
    excel_exe: str | None = _EXCEL_EXE,
    owner_exe: str | None = _PY_EXE,
    owner_sess: int | None = 1,
    excel_sess: int | None = 1,
) -> None:
    store.write_record(
        build_owned_record(
            owner_pid=owner_pid,
            excel_pid=excel_pid,
            owner_create_time=owner_ct,
            owner_exe=owner_exe,
            owner_session_id=owner_sess,
            excel_create_time=excel_ct,
            excel_exe=excel_exe,
            excel_session_id=excel_sess,
        )
    )


def _reconcile(
    store: OwnedPidStore,
    *,
    alive: dict[int, bool],
    create: dict[int, int | None],
    image: dict[int, str | None],
    terminated: list[int] | None = None,
    terminate_ok: bool = True,
    current_session_id: Any = None,
) -> str | None:
    def _terminate(pid: int) -> bool:
        if terminated is not None:
            terminated.append(pid)
        if terminate_ok:
            alive[pid] = False
        return terminate_ok

    return _reconcile_prior_owned_pid(
        store,
        pid_alive=lambda pid: alive.get(pid, False),
        create_time=lambda pid: create.get(pid),
        image_name=lambda pid: image.get(pid),
        terminate_verified=_terminate,
        current_session_id=current_session_id,
    )


def test_reconcile_refuses_when_owner_alive_and_identity_matches(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)

    reason = _reconcile(
        store,
        alive={100: True, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
    )

    assert reason is not None
    assert "100" in reason
    # The marker is left in place so the operator can act on it.
    assert store.read() == 200


def test_reconcile_reaps_verified_orphan_when_owner_dead(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        terminate_ok=True,
    )

    assert reason is None, "a verified orphan is reaped and startup proceeds"
    assert terminated == [200], "exactly the verified orphan pid is terminated"
    assert store.read() is None, "the marker is cleared after the confirmed reap"


def test_reconcile_refuses_when_orphan_kill_fails(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        terminate_ok=False,  # kill failed, orphan still alive
    )

    assert reason is not None
    assert terminated == [200]
    assert store.read() == 200, "a still-alive orphan keeps the marker (no double launch)"


def test_reconcile_pid_reuse_create_time_mismatch_does_not_kill(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 999},  # excel create-time mismatch => reused pid
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
    )

    assert reason is not None, "an unverifiable EXCEL.EXE is refused, not killed"
    assert terminated == [], "never kill on pid-alone / create-time mismatch"
    assert store.read() == 200


def test_reconcile_pid_reused_by_non_excel_clears_and_proceeds(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 999},
        image={100: _PY_EXE, 200: r"C:\Windows\NOTEPAD.EXE"},
        terminated=terminated,
    )

    assert reason is None, "a non-Excel reuse is not our orphan; proceed"
    assert terminated == [], "never kill a non-Excel pid"
    assert store.read() is None, "the stale marker is cleared"


def test_reconcile_proceeds_when_prior_excel_dead(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)

    reason = _reconcile(
        store,
        alive={100: False, 200: False},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
    )

    assert reason is None
    assert store.read() is None, "a dead prior pid marker is cleared so startup proceeds"


def test_reconcile_proceeds_when_no_prior_record(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)

    reason = _reconcile(store, alive={}, create={}, image={})

    assert reason is None


def test_owned_pid_store_round_trips(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    assert store.read_record() is None
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=1234, excel_ct=222)
    record = store.read_record()
    assert record is not None
    assert isinstance(record, OwnedProcessRecord)
    assert record.excel.pid == 1234
    assert record.owner.pid == 100
    assert record.excel.create_time_filetime == 222
    # Full owner + excel identity survives the round trip (G2a): exe, session id and
    # the owner create-time must persist so the reconcile can verify identity.
    assert record.owner.create_time_filetime == 111
    assert record.owner.exe == _PY_EXE
    assert record.excel.exe == _EXCEL_EXE
    assert record.owner.session_id == 1
    assert record.excel.session_id == 1
    assert record.excel.image_verified is True
    assert store.read() == 1234
    store.clear()
    assert store.read() is None


def test_owned_pid_store_reads_legacy_pid_only(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.path.write_text('{"owned_pid": 1234, "updated_at": "x"}', encoding="utf-8")
    record = store.read_record()
    assert record is not None
    assert record.excel.pid == 1234
    assert record.owner.pid == 0, "a legacy record has no owner identity"
    assert record.excel.create_time_filetime is None
    assert store.read() == 1234


def test_owned_pid_store_ignores_malformed(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.path.write_text("{not json", encoding="utf-8")
    assert store.read() is None
    assert store.read_record() is None


# ---------------------------------------------------------------------------
# Regression guard: the server never auto-kills a live pid on deadline (U2 / E)
# ---------------------------------------------------------------------------
def test_no_taskkill_in_app_adapter_source() -> None:
    """Block A removed the pid-only ``taskkill`` fallback (same-handle only)."""

    import io
    import tokenize

    source = (
        Path(__file__).resolve().parents[2]
        / "src"
        / "excel_vba_mcp"
        / "adapters"
        / "excel"
        / "app_adapter.py"
    ).read_text(encoding="utf-8")
    # Strip comments / string literals (docstrings mention taskkill to explain its
    # removal) so the check is on actual code only.
    code_tokens: list[str] = []
    for tok in tokenize.generate_tokens(io.StringIO(source).readline):
        if tok.type in (tokenize.COMMENT, tokenize.STRING):
            continue
        code_tokens.append(tok.string)
    code = " ".join(code_tokens)
    assert "taskkill" not in code
    assert "subprocess" not in code
    # The same-handle terminate is present.
    assert "TerminateProcess" in code
    assert "WaitForSingleObject" in code


# ---------------------------------------------------------------------------
# Material 1: reconcile second-launch gate drives worker.start (not just a string)
# ---------------------------------------------------------------------------
class _StartRecorder:
    """Fake worker recording whether ``start`` was awaited."""

    def __init__(self, *, owned_pid: int | None = 7777) -> None:
        self.owned_pid = owned_pid
        self.started = False

    async def start(self) -> None:
        self.started = True

    def get_status(self) -> Any:
        return SimpleNamespace(owned_pid=self.owned_pid)


def _patch_adapter_probes(
    monkeypatch: pytest.MonkeyPatch,
    *,
    alive: dict[int, bool],
    create: dict[int, int | None],
    image: dict[int, str | None],
    terminate_ok: bool = True,
) -> None:
    monkeypatch.setattr(
        ExcelAppAdapter, "_pid_alive", staticmethod(lambda pid: alive.get(pid, False))
    )
    monkeypatch.setattr(
        ExcelAppAdapter,
        "_process_create_time_filetime",
        staticmethod(lambda pid: create.get(pid)),
    )
    monkeypatch.setattr(
        ExcelAppAdapter, "_image_name_of_pid", staticmethod(lambda pid: image.get(pid))
    )
    monkeypatch.setattr(
        ExcelAppAdapter, "_terminate_pid_verified", staticmethod(lambda _pid: terminate_ok)
    )
    monkeypatch.setattr(
        ExcelAppAdapter,
        "_capture_process_identity",
        staticmethod(lambda pid: (create.get(pid), image.get(pid), 1)),
    )


async def test_start_worker_guarded_skips_start_when_prior_owner_alive(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    _patch_adapter_probes(
        monkeypatch,
        alive={100: True, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
    )
    worker = _StartRecorder()
    runtime = SimpleNamespace(worker=worker)

    refusal = await _start_worker_guarded(runtime, store)  # type: ignore[arg-type]

    assert refusal is not None, "a live prior owner must surface a refusal"
    assert worker.started is False, "second launch prevented: worker.start() not awaited"


async def test_start_worker_guarded_starts_and_persists_when_prior_dead(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    _patch_adapter_probes(monkeypatch, alive={}, create={}, image={})
    worker = _StartRecorder(owned_pid=7777)
    runtime = SimpleNamespace(worker=worker)

    refusal = await _start_worker_guarded(runtime, store)  # type: ignore[arg-type]

    assert refusal is None
    assert worker.started is True, "a dead prior owner must allow the (single) launch"
    record = store.read_record()
    assert record is not None and record.excel.pid == 7777, (
        "the freshly launched owned pid is persisted with owner identity"
    )
    assert record.owner.pid != 0, "the owner (this server) identity is recorded"


# ---------------------------------------------------------------------------
# Material 3: runtime refuse is taskkill-resistant (no kill path runs)
# ---------------------------------------------------------------------------
def test_terminate_refuses_without_any_kill_path_running(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import os
    import subprocess

    import win32api

    trips: list[str] = []
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: trips.append("subprocess.run"))
    monkeypatch.setattr(os, "system", lambda *a, **k: trips.append("os.system"))
    monkeypatch.setattr(
        win32api, "TerminateProcess", lambda *a, **k: trips.append("TerminateProcess")
    )
    monkeypatch.setattr(
        app_adapter_module,
        "_terminate_and_confirm",
        lambda *a, **k: trips.append("_terminate_and_confirm"),
    )

    for bad_handle in (None, 123):  # no handle / handle reduced to int
        adapter = ExcelAppAdapter()
        adapter._owned_handle = bad_handle
        assert adapter.terminate_owned_handle() is False

    assert trips == [], f"no kill path may run on a fail-closed refuse: {trips}"


# ---------------------------------------------------------------------------
# Material 4: _terminate_and_confirm -> None => not confirmed; bounded budget passed
# ---------------------------------------------------------------------------
def test_terminate_not_confirmed_when_helper_returns_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    adapter = ExcelAppAdapter()
    handle = _FakeHandle()
    adapter._owned_handle = handle
    seen: list[int] = []

    def none_helper(_h: object, timeout_ms: int) -> None:
        seen.append(timeout_ms)
        return None  # pywin32 absent / a call failed

    monkeypatch.setattr(app_adapter_module, "_terminate_and_confirm", none_helper)

    assert adapter.terminate_owned_handle() is False
    # The bounded re-verification budget (ms) is the timeout passed.
    assert seen == [int(app_adapter_module._REAP_VERIFY_TIMEOUT_S * 1000)]


# ---------------------------------------------------------------------------
# Material 5: OpenProcess returns 0 (falsy, no exception) => fail closed (None)
# ---------------------------------------------------------------------------
def test_open_owned_handle_fail_closed_when_openprocess_returns_zero(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(ExcelAppAdapter, "_pid_is_excel", staticmethod(lambda _pid: True))
    import win32api

    monkeypatch.setattr(win32api, "OpenProcess", lambda *a, **k: 0)  # falsy, no raise
    assert ExcelAppAdapter._open_owned_handle(4242) is None


# ---------------------------------------------------------------------------
# Material 6: create_hidden with Hwnd=0 => no handle, recovery unavailable
# ---------------------------------------------------------------------------
def test_create_hidden_with_hwnd_zero_has_no_handle(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    adapter = ExcelAppAdapter()

    class _ZeroHwndApp:
        Visible = False
        Caption = "Excel"
        Hwnd = 0

        def __setattr__(self, name: str, value: Any) -> None:
            object.__setattr__(self, name, value)

    monkeypatch.setattr("win32com.client.DispatchEx", lambda _prog: _ZeroHwndApp(), raising=False)
    # No job layer in this unit path (job behaviour is covered separately).
    monkeypatch.setattr(ExcelAppAdapter, "_create_job_object", staticmethod(lambda: None))

    adapter.create_hidden()

    assert adapter.owned_pid is None
    assert adapter.owned_handle is None
    assert adapter.recovery_available is False
    assert adapter.job_protected is False


# ---------------------------------------------------------------------------
# G1: the REAL _assign_process_to_job (win32job/win32api faked via sys.modules)
# ---------------------------------------------------------------------------
def _install_fake_win32(
    monkeypatch: pytest.MonkeyPatch,
    *,
    in_job: bool,
    assign_raises: bool = False,
) -> dict[str, list[Any]]:
    """Install fake ``win32api`` / ``win32job`` modules; return a call log.

    ``_assign_process_to_job`` imports both modules at call time, so injecting them
    into ``sys.modules`` exercises the real production control flow (open assign
    handle -> assign -> close -> open query handle -> IsProcessInJob) without any
    real process or job.
    """

    calls: dict[str, list[Any]] = {"assign": [], "open": [], "close": []}

    fake_api = types.ModuleType("win32api")

    def _open_process(access: int, _inherit: bool, pid: int) -> object:
        calls["open"].append((access, pid))
        return object()

    def _close_handle(handle: object) -> None:
        calls["close"].append(handle)

    fake_api.OpenProcess = _open_process  # type: ignore[attr-defined]
    fake_api.CloseHandle = _close_handle  # type: ignore[attr-defined]

    fake_job = types.ModuleType("win32job")

    def _assign(job: object, handle: object) -> None:
        calls["assign"].append((job, handle))
        if assign_raises:
            raise OSError("assign denied")

    def _is_in_job(_handle: object, _job: object) -> bool:
        return in_job

    fake_job.AssignProcessToJobObject = _assign  # type: ignore[attr-defined]
    fake_job.IsProcessInJob = _is_in_job  # type: ignore[attr-defined]

    monkeypatch.setitem(sys.modules, "win32api", fake_api)
    monkeypatch.setitem(sys.modules, "win32job", fake_job)
    return calls


def test_assign_process_to_job_true_when_isprocessinjob_true(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = _install_fake_win32(monkeypatch, in_job=True)
    job = object()

    assert ExcelAppAdapter._assign_process_to_job(job, 4242) is True
    # The owned pid was assigned into the given job, and both handles were closed.
    assert calls["assign"] and calls["assign"][0][0] is job
    assert len(calls["close"]) == 2, "assign + query handles are both closed"


def test_assign_process_to_job_false_when_isprocessinjob_false(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _install_fake_win32(monkeypatch, in_job=False)

    assert ExcelAppAdapter._assign_process_to_job(object(), 4242) is False


def test_assign_process_to_job_fails_closed_when_assign_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls = _install_fake_win32(monkeypatch, in_job=True, assign_raises=True)

    assert ExcelAppAdapter._assign_process_to_job(object(), 4242) is False
    assert calls["assign"], "AssignProcessToJobObject was attempted before failing"


def test_assign_process_to_job_false_when_openprocess_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake_api = types.ModuleType("win32api")

    def _deny(*_args: Any, **_kwargs: Any) -> Any:
        raise OSError("access denied")

    fake_api.OpenProcess = _deny  # type: ignore[attr-defined]
    fake_api.CloseHandle = lambda _h: None  # type: ignore[attr-defined]
    fake_job = types.ModuleType("win32job")
    monkeypatch.setitem(sys.modules, "win32api", fake_api)
    monkeypatch.setitem(sys.modules, "win32job", fake_job)

    assert ExcelAppAdapter._assign_process_to_job(object(), 4242) is False


# ---------------------------------------------------------------------------
# G2b: session_id is enforced in the reconcile identity (cross-session guard)
# ---------------------------------------------------------------------------
def test_reconcile_refuses_cross_session_orphan(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(
        store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222, excel_sess=5
    )
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        current_session_id=lambda: 7,  # this server is in a different session
    )

    assert reason is not None
    assert "session" in reason.lower()
    assert terminated == [], "never kill a cross-session process"
    assert store.read() == 200, "the marker is kept for an operator decision"


def test_reconcile_reaps_same_session_orphan(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(
        store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222, excel_sess=5
    )
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        current_session_id=lambda: 5,  # same session => safe to reap
    )

    assert reason is None
    assert terminated == [200]
    assert store.read() is None


def test_reconcile_unknown_session_falls_back_to_create_time(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_record(
        store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222, excel_sess=5
    )
    terminated: list[int] = []

    # current_session_id None => the session check is skipped; the create-time /
    # image gate still governs, so a verified orphan is reaped.
    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        current_session_id=None,
    )

    assert reason is None
    assert terminated == [200]


# ---------------------------------------------------------------------------
# G3: _resolve_owned_pid_dir env + default_owned_pid_dir %LOCALAPPDATA%
# ---------------------------------------------------------------------------
def test_resolve_owned_pid_dir_honors_env(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from excel_vba_mcp.server import _resolve_owned_pid_dir

    target = tmp_path / "custom_owned"
    monkeypatch.setenv("EXCEL_VBA_MCP_OWNED_PID_DIR", str(target))
    assert _resolve_owned_pid_dir() == target


def test_resolve_owned_pid_dir_defaults_to_localappdata(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from excel_vba_mcp.server import _resolve_owned_pid_dir
    from excel_vba_mcp.storage.owned_pid_store import default_owned_pid_dir

    monkeypatch.delenv("EXCEL_VBA_MCP_OWNED_PID_DIR", raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert default_owned_pid_dir() == tmp_path / "excel-vba-mcp"
    assert _resolve_owned_pid_dir() == tmp_path / "excel-vba-mcp"


def test_default_owned_pid_dir_falls_back_to_temp(monkeypatch: pytest.MonkeyPatch) -> None:
    import tempfile

    from excel_vba_mcp.storage.owned_pid_store import default_owned_pid_dir

    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    assert default_owned_pid_dir() == Path(tempfile.gettempdir()) / "excel-vba-mcp"


# ---------------------------------------------------------------------------
# G4: quit() closes the job AFTER reap_owned_pid; release() closes the job
# ---------------------------------------------------------------------------
def test_quit_closes_job_after_reap(monkeypatch: pytest.MonkeyPatch) -> None:
    adapter = ExcelAppAdapter()
    order: list[str] = []

    class _App:
        DisplayAlerts = True

        def Quit(self) -> None:  # noqa: N802 - COM casing
            order.append("app_quit")

    class _Job:
        def Close(self) -> None:  # noqa: N802 - PyHANDLE casing
            order.append("close_job")

    adapter._app = _App()
    adapter._owned_pid = 4242
    adapter._owned_handle = _FakeHandle()
    adapter._job_handle = _Job()
    monkeypatch.setattr(adapter, "reap_owned_pid", lambda _pid: order.append("reap") or True)

    adapter.quit()

    assert "reap" in order and "close_job" in order, order
    assert order.index("reap") < order.index("close_job"), (
        f"the job must close AFTER the reap attempt: {order}"
    )
    assert adapter.job_handle is None


def test_release_closes_job() -> None:
    adapter = ExcelAppAdapter()
    closed: list[bool] = []

    class _Job:
        def Close(self) -> None:  # noqa: N802 - PyHANDLE casing
            closed.append(True)

    adapter._job_handle = _Job()
    adapter.release()

    assert closed == [True]
    assert adapter.job_handle is None


# ---------------------------------------------------------------------------
# G5: job_protected is False when the job handle is reduced to int
# ---------------------------------------------------------------------------
def test_job_protected_false_when_handle_reduced_to_int() -> None:
    adapter = ExcelAppAdapter()

    adapter._job_handle = 123  # a reduced int, not a live object
    assert adapter.job_protected is False

    adapter._job_handle = None
    assert adapter.job_protected is False

    adapter._job_handle = _FakeHandle()
    assert adapter.job_protected is True


# ---------------------------------------------------------------------------
# Verify A mitigation: state="creating" record + reconcile (create->assign crash)
# ---------------------------------------------------------------------------
def _write_creating(
    store: OwnedPidStore,
    *,
    owner_pid: int,
    owner_ct: int | None,
    excel_pid: int = 0,
    excel_ct: int | None = None,
    excel_sess: int | None = 1,
) -> None:
    if excel_pid <= 0:
        store.write_record(
            build_creating_record(
                owner_pid=owner_pid,
                owner_create_time=owner_ct,
                owner_exe=_PY_EXE,
                owner_session_id=1,
            )
        )
    else:
        store.write_record(
            build_owned_record(
                owner_pid=owner_pid,
                excel_pid=excel_pid,
                owner_create_time=owner_ct,
                owner_exe=_PY_EXE,
                owner_session_id=1,
                excel_create_time=excel_ct,
                excel_exe=_EXCEL_EXE,
                excel_session_id=excel_sess,
                state=STATE_CREATING,
            )
        )


def test_reconcile_creating_refuses_when_owner_dead_and_pid_unknown(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_creating(store, owner_pid=100, owner_ct=111)  # excel pid unknown (0)
    terminated: list[int] = []

    reason = _reconcile(
        store, alive={}, create={}, image={}, terminated=terminated
    )

    assert reason is not None, "an unidentifiable orphan may exist => refuse"
    assert terminated == [], "never pid-kill on an unknown excel pid"
    assert store.read() == 0, "the creating marker is kept for an operator decision"


def test_reconcile_creating_reaps_verified_orphan(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_creating(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
        current_session_id=lambda: 1,
    )

    assert reason is None
    assert terminated == [200], "a verified orphan is reaped exactly"
    assert store.read() is None, "the marker is cleared after the confirmed reap"


def test_reconcile_creating_refuses_unverified_orphan(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_creating(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)
    terminated: list[int] = []

    reason = _reconcile(
        store,
        alive={100: False, 200: True},
        create={100: 111, 200: 999},  # create-time mismatch => unverified
        image={100: _PY_EXE, 200: _EXCEL_EXE},
        terminated=terminated,
    )

    assert reason is not None, "an unverified EXCEL.EXE is refused, not killed"
    assert terminated == [], "never kill on create-time mismatch"
    assert store.read() == 200


def test_reconcile_creating_refuses_when_owner_alive(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_creating(store, owner_pid=100, owner_ct=111)

    reason = _reconcile(
        store,
        alive={100: True},
        create={100: 111},
        image={100: _PY_EXE},
    )

    assert reason is not None, "a live owner mid-launch => refuse a second launch"


def test_reconcile_creating_clears_when_known_excel_dead(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    _write_creating(store, owner_pid=100, owner_ct=111, excel_pid=200, excel_ct=222)

    reason = _reconcile(
        store,
        alive={100: False, 200: False},  # known excel did not survive
        create={100: 111, 200: 222},
        image={100: _PY_EXE, 200: _EXCEL_EXE},
    )

    assert reason is None
    assert store.read() is None, "a known-dead excel leaves no orphan => clear+proceed"


def test_owned_pid_store_round_trips_creating_unknown_pid(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.write_record(
        build_creating_record(
            owner_pid=100, owner_create_time=111, owner_exe=_PY_EXE, owner_session_id=1
        )
    )
    record = store.read_record()
    assert record is not None
    assert record.state == STATE_CREATING
    assert record.excel.pid == 0, "a creating record has an unknown (0) excel pid"
    assert record.owner.pid == 100


class _CreatingCheckWorker:
    """Worker that records the store state at ``start`` time (creating marker)."""

    def __init__(self, store: OwnedPidStore, *, owned_pid: int | None = 7777) -> None:
        self.store = store
        self.owned_pid = owned_pid
        self.state_at_start: str | None = None
        self.started = False

    async def start(self) -> None:
        rec = self.store.read_record()
        self.state_at_start = rec.state if rec is not None else None
        self.started = True

    def get_status(self) -> Any:
        return SimpleNamespace(owned_pid=self.owned_pid)


async def test_start_worker_guarded_writes_creating_before_start(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    store = OwnedPidStore(tmp_path)
    _patch_adapter_probes(monkeypatch, alive={}, create={}, image={})
    worker = _CreatingCheckWorker(store)
    runtime = SimpleNamespace(worker=worker)

    refusal = await _start_worker_guarded(runtime, store)  # type: ignore[arg-type]

    assert refusal is None
    assert worker.started is True
    assert worker.state_at_start == STATE_CREATING, "creating marker precedes start"
    final = store.read_record()
    assert final is not None and final.state != STATE_CREATING
    assert final.excel.pid == 7777, "the running record replaces the creating marker"


class _FailingStartWorker:
    def __init__(self) -> None:
        self.started = False

    async def start(self) -> None:
        self.started = True
        raise RuntimeError("boom")

    def get_status(self) -> Any:
        return SimpleNamespace(owned_pid=None)


async def test_start_worker_guarded_clears_creating_on_start_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    store = OwnedPidStore(tmp_path)
    _patch_adapter_probes(monkeypatch, alive={}, create={}, image={})
    worker = _FailingStartWorker()
    runtime = SimpleNamespace(worker=worker)

    with pytest.raises(RuntimeError, match="boom"):
        await _start_worker_guarded(runtime, store)  # type: ignore[arg-type]

    assert worker.started is True
    assert store.read() is None, "a fail-closed start clears the creating marker"
