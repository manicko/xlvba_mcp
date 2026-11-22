"""Live-Excel integration verification for the production STA COM worker (task_011).

This module verifies the *production* worker
(:class:`excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`) against the
real target Excel build (Windows 10 LTSC 2019 17763 x64, Excel 16.0.x). It does
not re-implement any COM plumbing; it drives the public worker API and observes
the desktop.

Coverage (exact test names required by task_011):

* ``test_requests_are_serialized`` — concurrent ``READ_ONLY`` submits complete one
  at a time in FIFO order, every call runs on the same STA thread id, the owned
  process id is a distinct automation instance, and the owned instance stays
  hidden / non-intrusive (addendum BG03/BG04/BG05).
* ``test_owned_excel_instance_shutdown`` — ``shutdown()`` closes the owned
  workbook and reaps the owned ``EXCEL.exe`` while a *bystander* Excel instance
  (representing the user session) is left alive and untouched; no focus steal.
  The owned pid is gone the moment ``shutdown()`` returns (FE-17: no poll
  barrier).
* ``test_com_busy_does_not_replay_mutation`` — ``RPC_E_CALL_REJECTED`` is never
  replayed for a ``MUTATING`` op (exactly one attempt) but is retried a bounded
  number of times for a ``READ_ONLY`` op; a timed-out call moves the worker to
  ``NEEDS_RECOVERY`` and a subsequent non-``close`` submit fails fast without
  starting a new COM call.

Safety / addendum compliance:

* Real Excel only in tests 1–2, always ``Visible=False``; the worker never calls
  ``.Activate``/``.Select`` and these tests never do either.
* All desktop checks are read-only Win32 observations (``GetForegroundWindow``,
  ``EnumWindows`` + ``IsWindowVisible``); nothing is shown, activated, clicked or
  typed. Only temp copies of fixtures are opened.
* A bystander Excel instance is created and closed only by the test that owns it;
  no non-owned Excel process is ever killed.
* Tests run serially (no xdist); each test cleans up its own worker + bystander.

``test_com_busy_does_not_replay_mutation`` uses a small ``FakeComError`` stub that
exposes ``.hresult`` (the worker's ``_hresult_of`` reads ``.hresult`` first and
then ``.args[0]``), so ``pythoncom.com_error`` need not be constructed directly.
It also injects a duck-typed fake ``ExcelAppAdapter`` so no real Excel is needed
for that case; this approach is documented inline.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_com_worker.py -v -s
"""

from __future__ import annotations

import asyncio
import contextlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.domain.errors import ErrorCode, ExcelHungError, RpcCallRejectedError
from excel_vba_mcp.domain.statuses import OperationClass, WorkbookSessionState

# HRESULT exercised by the busy/rejected retry policy.
RPC_E_CALL_REJECTED = 0x80010001

# Optional COM dependencies. ``None`` when unavailable so the module still
# imports (and the COM tests skip) on non-Windows / pywin32-less hosts.
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

# Excel shutdown on the target host is variable; poll long enough to observe a
# slow-but-legitimate reaping (matches the snapshot PoC experience).
ORPHAN_POLL_TIMEOUT_S = 30.0
ORPHAN_POLL_INTERVAL_S = 0.5

# Generous start/join budget for real-Excel tests (cold launch can be slow).
_WORKER_TIMEOUT_S = 30.0


# ---------------------------------------------------------------------------
# Availability / skip helpers
# ---------------------------------------------------------------------------
def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_com = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required for the STA worker tests (platform={sys.platform})",
)
requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason="pywin32/Windows required to drive a live Excel instance",
)


# ---------------------------------------------------------------------------
# Process / window helpers (all READ-ONLY; never activate / show / kill)
# ---------------------------------------------------------------------------
def _pid_from_hwnd(hwnd: int) -> int | None:
    """Best-effort process id for a window handle (read-only)."""

    if win32process is None:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


def _pid_alive(pid: int) -> bool:
    """Return whether an OS process with ``pid`` currently exists."""

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


def _wait_pid_gone(pid: int, timeout: float = ORPHAN_POLL_TIMEOUT_S) -> bool:
    """Poll until ``pid`` is gone (bounded); return True if it disappeared."""

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _foreground() -> dict[str, Any] | None:
    """Read the current foreground window (never changes it)."""

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
    """Enumerate top-level windows for ``pid`` with visibility (read-only)."""

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
    """BG03: the owned automation window must never hold foreground."""

    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground "
        f"at {label} (hwnd={fg['hwnd']}, title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    """BG04: the owned instance must expose no visible top-level window."""

    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows "
        f"at {label}: {[w['title'] for w in visible]}"
    )


# ---------------------------------------------------------------------------
# Bystander Excel (represents the user session; created/closed only by a test)
# ---------------------------------------------------------------------------
@contextlib.contextmanager
def _bystander_excel() -> Iterator[tuple[Any, int]]:
    """Create a *separate* Excel instance to represent the user's session.

    Yields ``(app, pid)``. The bystander is always ``Quit``-ed by this context
    manager; it is never activated or made visible.

    ``DispatchEx`` here runs on the asyncio loop (main) thread. In a full-suite
    run that thread's COM apartment may already have been torn down by the
    synchronous PoC tests, which would make ``CoCreateInstanceEx`` fail with
    ``CO_E_NOTINITIALIZED``. We therefore initialize COM on the calling thread
    for the lifetime of the bystander and undo it on exit. The owned worker runs
    on its own dedicated STA thread with its own independent COM init, so this
    per-test init only scopes the loop-thread bystander.
    """

    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app: Any = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        pid = _pid_from_hwnd(int(app.Hwnd))
        assert pid is not None, "could not resolve the bystander Excel PID"
        yield app, pid
    finally:
        if app is not None:
            with contextlib.suppress(Exception):
                app.Quit()
            with contextlib.suppress(Exception):
                del app
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# Fakes for the busy/rejected test (no real Excel needed)
# ---------------------------------------------------------------------------
class FakeComError(Exception):
    """Minimal stand-in for ``pythoncom.com_error`` carrying an ``hresult``.

    The worker's ``_hresult_of`` reads ``.hresult`` first (and falls back to
    ``.args[0]``), so this cleanly drives the busy/rejected classification
    without constructing a real ``pythoncom.com_error``.
    """

    def __init__(self, hresult: int) -> None:
        super().__init__(f"com_error 0x{hresult & 0xFFFFFFFF:08X}")
        self.hresult = hresult


class _FakeAdapter(ExcelAppAdapter):
    """Duck-typed adapter that records lifecycle calls without touching COM."""

    def __init__(self) -> None:
        super().__init__()
        self.create_count = 0
        self.quit_count = 0

    def create_hidden(self) -> None:
        self.create_count += 1
        self._owned_pid = 4242
        self._owned_caption = "fake"

    def quit(self) -> None:
        self.quit_count += 1
        self._app = None
        self._owned_pid = None
        self._owned_caption = None

    def release(self) -> None:
        self._app = None
        self._owned_pid = None
        self._owned_caption = None


# ---------------------------------------------------------------------------
# anyio backend (asyncio only: the worker bridges via asyncio.wrap_future)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Test 1: serialization on one STA thread + non-intrusive hidden instance
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_requests_are_serialized() -> None:
    adapter = ExcelAppAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=_WORKER_TIMEOUT_S)
    await worker.start()

    owned_pid: int | None = None
    try:
        main_tid = threading.get_ident()

        # A bystander Excel stands in for the user session so we can prove the
        # owned PID is a distinct process (BG05 ownership separation). The helper
        # context manager always Quits the bystander on exit.
        with _bystander_excel() as (_bystander_app, bystander_pid):
            order: list[str] = []
            thread_ids: list[int] = []

            def make(label: str, reader: Any) -> Any:
                def fn() -> Any:
                    thread_ids.append(threading.get_ident())
                    order.append(label)
                    app = adapter.app
                    return reader(app)

                return fn

            readers = {
                "version": lambda app: str(app.Version),
                "caption": lambda app: str(app.Caption),
                "build": lambda app: int(app.Build),
            }
            labels = ["version", "caption", "build"]

            # Submit concurrently; the single-consumer worker must serialize.
            results = await asyncio.gather(
                *(
                    worker.submit(
                        make(label, readers[label]),
                        timeout=_WORKER_TIMEOUT_S,
                        op_class=OperationClass.READ_ONLY,
                    )
                    for label in labels
                )
            )

            # All ops returned correct values (gather preserves submission order).
            version, caption, build = results
            assert re.match(r"^\d+\.\d+", version), f"unexpected Version: {version!r}"
            assert isinstance(caption, str), f"unexpected Caption type: {caption!r}"
            assert isinstance(build, int) and build > 0, f"unexpected Build: {build!r}"

            # FIFO: completion order equals submission order.
            assert order == labels, f"completion order {order} != submission order {labels}"

            # Every op observed the same STA thread id, distinct from the loop.
            assert len(set(thread_ids)) == 1, f"calls ran on multiple threads: {thread_ids}"
            sta_tid = thread_ids[0]
            assert sta_tid != main_tid, "worker ran on the event-loop thread, not a dedicated STA"

            # Hidden/non-intrusive (BG03/04/05).
            visible = await worker.submit(
                lambda: bool(adapter.app.Visible),
                timeout=_WORKER_TIMEOUT_S,
                op_class=OperationClass.READ_ONLY,
            )
            assert visible is False, "owned Excel instance must stay Visible=False"

            status = worker.get_status()
            owned_pid = status.owned_pid
            assert owned_pid is not None, "worker did not record an owned PID"
            assert owned_pid != os.getpid(), "owned PID must be a separate EXCEL.exe process"
            assert owned_pid != bystander_pid, (
                f"owned PID {owned_pid} must differ from bystander PID {bystander_pid}"
            )

            # Read-only desktop evidence.
            _assert_owned_never_foreground(owned_pid, "serialized-ops")
            _assert_no_visible_owned_windows(owned_pid, "serialized-ops")

            print(
                f"[serialized] order={order} sta_tid={sta_tid} main_tid={main_tid} "
                f"owned_pid={owned_pid} bystander_pid={bystander_pid} "
                f"version={version} build={build} Visible={visible}"
            )
    finally:
        await worker.shutdown()

    # The owned process must be gone after shutdown (no orphan).
    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: shutdown closes the owned instance, leaves a bystander untouched
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_owned_excel_instance_shutdown() -> None:
    adapter = ExcelAppAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=_WORKER_TIMEOUT_S)
    await worker.start()

    bystander_pid: int | None = None
    owned_pid: int | None = None
    tmp_dir: tempfile.TemporaryDirectory[str] | None = None
    try:
        with _bystander_excel() as (bystander, bystander_pid):
            owned_pid = worker.get_status().owned_pid
            assert owned_pid is not None

            # Record the bystander's state so we can prove it is untouched.
            bystander_visible_before = bool(bystander.Visible)
            bystander_wb_count_before = int(bystander.Workbooks.Count)
            bystander_windows_before = _windows_for_pid(bystander_pid)
            fg_before = _foreground()
            _assert_owned_never_foreground(owned_pid, "after-start")

            # Open a workbook in the owned (worker) instance from a temp copy.
            tmp_dir = tempfile.TemporaryDirectory(prefix="task011_")
            wb_copy = Path(tmp_dir.name) / "CleanProject.xlsm"
            shutil.copy2(_CLEAN_PROJECT, wb_copy)

            def open_owned() -> str:
                wb = adapter.open_workbook(str(wb_copy), read_only=True)
                return str(wb.Name)

            opened_name = await worker.submit(
                open_owned,
                timeout=_WORKER_TIMEOUT_S,
                section="open",
                op_class=OperationClass.MUTATING,
            )
            assert opened_name, "owned workbook did not report a name"

            def open_count() -> int:
                return int(adapter.app.Workbooks.Count)

            owned_wb_count = await worker.submit(
                open_count,
                timeout=_WORKER_TIMEOUT_S,
                op_class=OperationClass.READ_ONLY,
            )
            assert owned_wb_count >= 1, "owned workbook was not open before shutdown"

            _assert_owned_never_foreground(owned_pid, "after-open")
            _assert_no_visible_owned_windows(owned_pid, "after-open")

            # Shutdown must close the owned instance and leave the bystander alone.
            await worker.shutdown()

            # FE-17: shutdown enforces the owned process death, so it is already
            # gone when shutdown() returns (no poll barrier) and the worker's
            # status - the snapshot ``excel_health`` reports - no longer holds it.
            assert not _pid_alive(owned_pid), (
                f"owned PID {owned_pid} survived worker.shutdown() (FE-17 regression)"
            )
            assert worker.get_status().owned_pid is None, (
                "worker status must drop the owned pid once the process is reaped"
            )

            # Owned workbook is gone with the owned process.
            assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
            assert worker.get_status().state == WorkbookSessionState.CLOSED

            # Bystander still alive and untouched (BG05 user-session preservation).
            assert _pid_alive(bystander_pid), f"bystander PID {bystander_pid} was terminated"
            assert bool(bystander.Visible) == bystander_visible_before, (
                "bystander Visible state changed"
            )
            assert int(bystander.Workbooks.Count) == bystander_wb_count_before, (
                "bystander workbook count changed"
            )
            bystander_windows_after = _windows_for_pid(bystander_pid)
            assert len(bystander_windows_after) == len(bystander_windows_before), (
                "bystander window set changed"
            )

            # No focus steal at any point.
            _assert_owned_never_foreground(owned_pid, "after-shutdown")
            fg_after = _foreground()
            if fg_before is not None and fg_after is not None:
                assert fg_after["hwnd"] == fg_before["hwnd"], (
                    "foreground window changed across owned shutdown "
                    f"(before={fg_before['title']!r}, after={fg_after['title']!r})"
                )

            print(
                f"[shutdown] owned_pid={owned_pid} (gone) bystander_pid={bystander_pid} (alive) "
                f"owned_wb={owned_wb_count}->{0} "
                f"bystander_wb={bystander_wb_count_before}->"
                f"{int(bystander.Workbooks.Count)} "
                f"bystander_visible={bystander_visible_before} state=CLOSED"
            )
    finally:
        await worker.shutdown()
        if tmp_dir is not None:
            tmp_dir.cleanup()


# ---------------------------------------------------------------------------
# Test 3: busy/rejected never replays a mutation; timeout -> recovery
# ---------------------------------------------------------------------------
@requires_com
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_com_busy_does_not_replay_mutation() -> None:
    adapter = _FakeAdapter()
    worker = ExcelComWorker(adapter, orphan_timeout_s=5.0)
    await worker.start()
    try:
        # --- MUTATING: exactly one attempt, no replay, RPC_E_CALL_REJECTED ---
        mutating_calls = {"n": 0}

        def mutating_reject() -> str:
            mutating_calls["n"] += 1
            raise FakeComError(RPC_E_CALL_REJECTED)

        with pytest.raises(RpcCallRejectedError) as excinfo:
            await worker.submit(
                mutating_reject,
                timeout=10.0,
                op_class=OperationClass.MUTATING,
            )
        assert excinfo.value.code == ErrorCode.RPC_E_CALL_REJECTED
        assert mutating_calls["n"] == 1, "MUTATING op must be attempted exactly once (no replay)"

        # Rejection is not a hang: the worker stays RUNNING and accepting.
        assert worker.get_status().state == WorkbookSessionState.RUNNING

        # --- READ_ONLY: bounded retry (1 + 3 retries) then surface rejection ---
        read_calls = {"n": 0}

        def readonly_reject() -> str:
            read_calls["n"] += 1
            raise FakeComError(RPC_E_CALL_REJECTED)

        with pytest.raises(RpcCallRejectedError) as ro_excinfo:
            await worker.submit(
                readonly_reject,
                timeout=30.0,
                op_class=OperationClass.READ_ONLY,
            )
        assert ro_excinfo.value.code == ErrorCode.RPC_E_CALL_REJECTED
        assert read_calls["n"] == 4, (
            f"READ_ONLY must retry a bounded number of times (expected 4 attempts), "
            f"got {read_calls['n']}"
        )

        # --- Timeout -> NEEDS_RECOVERY; subsequent non-close submit fails fast ---
        recovery_started = threading.Event()

        def slow() -> str:
            recovery_started.set()
            time.sleep(0.8)
            return "done"

        with pytest.raises(ExcelHungError) as hung_excinfo:
            await worker.submit(slow, timeout=0.1)
        assert hung_excinfo.value.code == ErrorCode.EXCEL_HUNG
        assert worker.get_status().state == WorkbookSessionState.NEEDS_RECOVERY

        # A new non-close submit must fail fast and NOT start a new COM call.
        failfast_started = threading.Event()

        def should_not_run() -> str:
            failfast_started.set()
            return "x"

        with pytest.raises(ExcelHungError) as ff_excinfo:
            await worker.submit(should_not_run, timeout=5.0, op_class=OperationClass.MUTATING)
        assert ff_excinfo.value.code == ErrorCode.EXCEL_HUNG
        assert recovery_started.is_set(), "the slow fn never started"
        assert not failfast_started.is_set(), "fail-fast submit must not start a new COM call"

        # reset() is the documented recovery exit; it waits for the hung fn.
        await worker.reset()
        assert worker.get_status().state == WorkbookSessionState.RUNNING

        print(
            f"[busy] mutating_calls={mutating_calls['n']} read_calls={read_calls['n']} "
            f"recovery_state=NEEDS_RECOVERY->RUNNING failfast_started={failfast_started.is_set()}"
        )
    finally:
        await worker.shutdown()
