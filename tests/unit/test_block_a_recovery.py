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
from excel_vba_mcp.storage.owned_pid_store import OwnedPidStore

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

    adapter.create_hidden()

    assert adapter.owned_handle is handle
    assert not isinstance(adapter.owned_handle, int)
    assert adapter.recovery_available is True


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
# Startup reconcile: live prior pid => no second launch; dead => proceed
# ---------------------------------------------------------------------------
def test_reconcile_refuses_when_prior_pid_alive(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.write(4242)

    reason = _reconcile_prior_owned_pid(store, pid_alive=lambda _pid: True)

    assert reason is not None
    assert "4242" in reason
    # The marker is left in place so the operator can act on it.
    assert store.read() == 4242


def test_reconcile_proceeds_when_prior_pid_dead(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.write(4242)

    reason = _reconcile_prior_owned_pid(store, pid_alive=lambda _pid: False)

    assert reason is None
    assert store.read() is None, "a dead prior pid marker is cleared so startup proceeds"


def test_reconcile_proceeds_when_no_prior_pid(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)

    reason = _reconcile_prior_owned_pid(store, pid_alive=lambda _pid: True)

    assert reason is None


def test_owned_pid_store_round_trips(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    assert store.read() is None
    store.write(1234)
    assert store.read() == 1234
    store.clear()
    assert store.read() is None


def test_owned_pid_store_ignores_malformed(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.path.write_text("{not json", encoding="utf-8")
    assert store.read() is None


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


async def test_start_worker_guarded_skips_start_when_prior_pid_alive(tmp_path: Path) -> None:
    store = OwnedPidStore(tmp_path)
    store.write(4242)  # a prior owned pid
    worker = _StartRecorder()
    runtime = SimpleNamespace(worker=worker)

    refusal = await _start_worker_guarded(
        runtime,  # type: ignore[arg-type]
        store,
        pid_alive=lambda _pid: True,  # prior pid is still alive
    )

    assert refusal is not None, "a live prior pid must surface a refusal"
    assert worker.started is False, "second launch prevented: worker.start() not awaited"


async def test_start_worker_guarded_starts_and_persists_when_prior_pid_dead(
    tmp_path: Path,
) -> None:
    store = OwnedPidStore(tmp_path)
    store.write(4242)
    worker = _StartRecorder(owned_pid=7777)
    runtime = SimpleNamespace(worker=worker)

    refusal = await _start_worker_guarded(
        runtime,  # type: ignore[arg-type]
        store,
        pid_alive=lambda _pid: False,  # prior pid is dead
    )

    assert refusal is None
    assert worker.started is True, "a dead prior pid must allow the (single) launch"
    assert store.read() == 7777, "the freshly launched owned pid is persisted"


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

    adapter.create_hidden()

    assert adapter.owned_pid is None
    assert adapter.owned_handle is None
    assert adapter.recovery_available is False
