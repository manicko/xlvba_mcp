"""Hidden, non-activating Excel Application wrapper (task_010).

This module centralizes every mutation of the owned ``Excel.Application``
state. All methods here are intended to run **on the dedicated COM STA worker
thread** (see :mod:`excel_vba_mcp.adapters.excel.com_worker`); they are plain
synchronous functions so the worker loop can invoke them directly inside a
submitted operation.

Design invariants (ARCH-013 as amended by the background-execution addendum):

* The owned instance is created via ``win32com.client.DispatchEx`` and is kept
  ``Visible = False`` for its entire lifetime. There is **no** code path that
  ever sets ``Visible = True``.
* The adapter never calls ``.Activate`` / ``.Select`` on any object, and never
  uses ``ShowWindow`` / ``SetForegroundWindow`` / ``SendKeys``. Direct
  object-model references are used instead.
* ``DisplayAlerts`` / ``ScreenUpdating`` / ``AutomationSecurity`` are changed
  only inside narrowly scoped context managers that restore the prior value in
  ``finally``.
* Teardown is an **enforced** invariant, not a best-effort side effect:
  :meth:`ExcelAppAdapter.quit` verifies that the owned process is dead and, if
  ``Application.Quit`` left it resident, terminates exactly that pid (never a
  wildcard image kill, never a user instance).

``DisplayAlerts = False`` is **not** a universal dialog blocker: it does not
suppress ``MsgBox``, ``UserForm.Show``, ``FileDialog``, add-in UI, or arbitrary
Win32 dialogs. It only suppresses Excel's own built-in alert prompts (e.g. the
"save changes?" prompt during ``Close``/``Quit``).

COM is imported lazily inside :meth:`ExcelAppAdapter.create_hidden` and the
helpers that need it, so importing this module never requires ``pywin32`` and
stays safe on non-Windows unit-test environments.
"""

from __future__ import annotations

import contextlib
import subprocess
import time
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover - typing only, no runtime COM import
    # ``Any`` stands in for an ``excel._cxobj.Application`` proxy; the adapter
    # never holds a typed COM reference at import time.
    pass

# Sentinel used by ``automation_security`` to restore a property whose prior
# read failed (e.g. a partially initialized app object).
_UNSET: Any = object()

# Grace period (seconds) given to ``Application.Quit`` to end the owned process
# before the pid-scoped fallback runs, and the poll interval used while waiting
# for that death. ``Quit`` is best-effort: it can raise and it can return while
# ``EXCEL.EXE`` is still resident (measured on the target build: ``Quit()``
# returns in ~0.2 s and the process survives indefinitely), so teardown verifies
# the death explicitly.
#
# The budget is deliberately short and bounded: an MCP STDIO server is only
# allowed ~2 s to shut down gracefully once its stdin closes, so the whole reap
# (grace + kill + verification) must fit inside that window for the guarantee to
# hold on the server shutdown path. A still-alive process after this short grace
# is terminated immediately; that is safe because owned workbooks are already
# closed without saving before ``quit`` runs.
_REAP_TIMEOUT_S = 0.5
_REAP_INTERVAL_S = 0.05

# Bounded re-verification window (seconds) after the pid-scoped fallback kill.
_REAP_VERIFY_TIMEOUT_S = 0.5


class ExcelAppAdapter:
    """Own and mutate a single hidden Excel ``Application`` instance.

    All methods are synchronous and must be invoked from the worker's STA
    thread. The adapter holds exactly one owned proxy (:attr:`app`) created by
    :meth:`create_hidden`; it never attaches to a user instance and never makes
    the owned instance visible or foreground.
    """

    def __init__(self) -> None:
        self._app: Any = None
        self._owned_pid: int | None = None
        self._owned_caption: str | None = None

    # ------------------------------------------------------------------
    # Read-only accessors (non-COM bookkeeping)
    # ------------------------------------------------------------------

    @property
    def app(self) -> Any:
        """The owned ``Application`` proxy (``None`` before :meth:`create_hidden`)."""

        return self._app

    @property
    def owned_pid(self) -> int | None:
        """Process id of the owned instance, or ``None`` if not yet created."""

        return self._owned_pid

    @property
    def owned_caption(self) -> str | None:
        """Window caption recorded at creation, or ``None`` if not yet created."""

        return self._owned_caption

    # ------------------------------------------------------------------
    # Creation / teardown
    # ------------------------------------------------------------------

    def create_hidden(self) -> None:
        """Create a dedicated hidden Excel instance and record its identity.

        Uses ``win32com.client.DispatchEx`` so a *new* ``EXCEL.exe`` process is
        spawned (rather than reusing a running instance through the ROT). The
        instance is set ``Visible = False`` immediately, before any workbook
        operation, and is kept hidden for its whole lifetime.

        The owned process id is derived from the application window handle
        (``Hwnd``) via a best-effort ``GetWindowThreadProcessId`` lookup; if the
        handle is unavailable the pid stays ``None`` (the instance is still
        owned and tracked by proxy reference).

        Raises:
            Exception: Any COM error from ``DispatchEx`` or the initial property
                reads propagates unchanged; the worker records the failure.
        """

        # Imported lazily so module import never requires pywin32 / Windows.
        import win32com.client  # type: ignore[import-untyped]  # noqa: PLC0415

        app = win32com.client.DispatchEx("Excel.Application")
        # Set Visible=False before touching anything else.
        try:
            app.Visible = False
        except Exception:
            # If we cannot hide it, do not proceed: a visible owned instance
            # would violate the non-intrusion invariant.
            with contextlib.suppress(Exception):
                app.Quit()
            raise

        self._app = app
        self._owned_caption = self._read_caption(app)
        self._owned_pid = self._pid_from_hwnd(self._read_hwnd(app))

    def quit(self) -> None:
        """Quit the owned instance and enforce the death of its process.

        ``Application.Quit`` is best-effort: a ``com_error`` is swallowed (the
        instance may already be gone), and a successful return does **not** mean
        the process died (an open workbook, a pending macro, a modal dialog or a
        live COM reference keeps ``EXCEL.EXE`` resident). The owned pid is
        therefore captured *before* quitting and its death is enforced by
        :meth:`reap_owned_pid`; the pid / caption are cleared only after that
        verification (or after the fallback attempt).

        Never raises: teardown must not mask the failure that triggered it.
        """

        app = self._app
        self._app = None
        owned_pid = self._owned_pid
        if app is not None:
            try:
                with self.alerts_suppressed(app):
                    app.Quit()
            except Exception:
                # The instance may already be gone; the pid check below is the
                # authoritative teardown signal.
                pass
        if owned_pid is not None:
            self.reap_owned_pid(owned_pid)
        self._owned_pid = None
        self._owned_caption = None

    def reap_owned_pid(self, pid: int) -> bool:
        """Ensure the owned ``pid`` is dead; return whether it is verified gone.

        Called by :meth:`quit` (on the STA worker thread) and by the worker's
        ``shutdown`` last resort (when the daemon STA thread may never have run
        the drain). ``Application.Quit`` gets a bounded grace period
        (:data:`_REAP_TIMEOUT_S`); a process still alive afterwards is
        terminated with ``taskkill /F /PID <pid>`` and re-verified.

        Only the adapter's own pid is ever passed here, so the blast radius is
        exactly one process: no user instance is touched and no wildcard image
        kill is ever used.

        Best-effort: an unavailable Win32 / ``taskkill`` layer degrades to "not
        verified" (``False``) and never raises.
        """

        if not self._pid_alive(pid):
            return True

        deadline = time.monotonic() + _REAP_TIMEOUT_S
        while True:
            time.sleep(_REAP_INTERVAL_S)
            if not self._pid_alive(pid):
                return True
            if time.monotonic() >= deadline:
                break

        # Grace period expired: terminate only this pid, then re-verify.
        self._terminate_pid(pid)
        deadline = time.monotonic() + _REAP_VERIFY_TIMEOUT_S
        while True:
            if not self._pid_alive(pid):
                return True
            if time.monotonic() >= deadline:
                return False
            time.sleep(_REAP_INTERVAL_S)

    def release(self) -> None:
        """Drop all references to the owned proxy without quitting.

        Used by the shutdown sequence after ``quit()``; also callable on its own
        to break cycles before ``gc.collect()``.
        """

        self._app = None
        self._owned_pid = None
        self._owned_caption = None

    # ------------------------------------------------------------------
    # Workbook operations
    # ------------------------------------------------------------------

    def open_workbook(
        self,
        path: str,
        *,
        read_only: bool,
        update_links: int = 0,
        notify: bool = False,
        add_to_mru: bool = False,
    ) -> Any:
        """Open a workbook on the owned instance and return its proxy.

        The returned object is a STA-local COM proxy; it must never escape the
        worker thread (only DTOs cross the worker boundary).

        Args:
            path: Canonical absolute path passed straight to COM (ARCH-007).
            read_only: Open the workbook read-only.
            update_links: ``Workbooks.Open`` ``UpdateLinks`` argument
                (``0`` = never update links).
            notify: ``Workbooks.Open`` ``Notify`` argument.
            add_to_mru: ``Workbooks.Open`` ``AddToMru`` argument.

        Returns:
            The opened ``Workbook`` proxy.
        """

        app = self._require_app()
        return app.Workbooks.Open(
            path,
            UpdateLinks=update_links,
            ReadOnly=read_only,
            Notify=notify,
            AddToMru=add_to_mru,
        )

    def close_workbook(self, wb: Any, *, save_changes: bool = False) -> None:
        """Close an owned workbook.

        Args:
            wb: The workbook proxy to close.
            save_changes: ``Close`` ``SaveChanges`` argument (default ``False``
                so an owned working copy is never saved implicitly).
        """

        if wb is None:
            return
        with self.alerts_suppressed(self._require_app()):
            wb.Close(SaveChanges=save_changes)

    # ------------------------------------------------------------------
    # Narrowly scoped settings context managers (restore in finally)
    # ------------------------------------------------------------------

    @contextlib.contextmanager
    def alerts_suppressed(self, app: Any = None) -> Iterator[None]:
        """Temporarily set ``DisplayAlerts = False`` and restore the prior value.

        This only suppresses Excel's own built-in alert prompts; it is **not** a
        universal dialog blocker (see the module docstring).
        """

        target = app if app is not None else self._require_app()
        prior = target.DisplayAlerts
        target.DisplayAlerts = False
        try:
            yield
        finally:
            target.DisplayAlerts = prior

    @contextlib.contextmanager
    def screenupdating_off(self, app: Any = None) -> Iterator[None]:
        """Temporarily set ``ScreenUpdating = False`` (repaint optimization only).

        This is a performance tweak, not a visibility control and not a dialog
        blocker. The prior value is restored in ``finally``.
        """

        target = app if app is not None else self._require_app()
        prior = target.ScreenUpdating
        target.ScreenUpdating = False
        try:
            yield
        finally:
            target.ScreenUpdating = prior

    @contextlib.contextmanager
    def automation_security(self, mso: int, app: Any = None) -> Iterator[None]:
        """Temporarily set ``AutomationSecurity`` to ``mso`` and restore the prior value.

        Used per-open (e.g. ``msoAutomationSecurityForceDisable`` for inspect,
        ``msoAutomationSecurityByUI`` immediately before ``Open`` for execute).
        The prior value is restored in ``finally``. If the prior value cannot be
        read, it is restored to the documented default
        ``msoAutomationSecurityLow`` (1).
        """

        target = app if app is not None else self._require_app()
        prior = self._read_with_default(target, "AutomationSecurity", default=1)
        target.AutomationSecurity = mso
        try:
            yield
        finally:
            target.AutomationSecurity = prior

    # ------------------------------------------------------------------
    # Non-intrusive runtime guard
    # ------------------------------------------------------------------

    def assert_non_intrusive(self) -> None:
        """Assert the owned instance is hidden and no owned window is foreground.

        Best-effort, read-only check (no COM mutation, no focus change). Raises
        ``RuntimeError`` when the invariant is violated so the worker can enter
        ``NEEDS_RECOVERY``.

        Raises:
            RuntimeError: If ``Visible`` is not ``False``.
        """

        app = self._app
        if app is None:
            return
        visible = bool(app.Visible)
        if visible:
            raise RuntimeError(
                "Non-intrusion invariant violated: owned Excel instance is Visible=True"
            )
        # Best-effort foreground check: if the owned window handle is the current
        # foreground window, record a violation. Read-only; never steals focus.
        if self._owned_window_is_foreground():
            raise RuntimeError("Non-intrusion invariant violated: owned Excel window is foreground")

    def _owned_window_is_foreground(self) -> bool:
        """Best-effort: is the owned window handle the current foreground window?

        Returns ``False`` when the check cannot be performed (non-Windows, no
        hwnd) — absence of evidence is not treated as a violation.
        """

        try:
            import ctypes
            from ctypes import wintypes
        except Exception:  # pragma: no cover - non-Windows fallback
            return False

        try:
            hwnd = wintypes.HWND(self._read_hwnd(self._require_app()))
            user32 = ctypes.windll.user32
            foreground = user32.GetForegroundWindow()
        except Exception:  # pragma: no cover - platform specific
            return False
        if not hwnd or not foreground:
            return False
        return int(hwnd) == int(foreground)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _require_app(self) -> Any:
        if self._app is None:
            raise RuntimeError("ExcelAppAdapter.create_hidden() must be called first")
        return self._app

    @staticmethod
    def _read_caption(app: Any) -> str | None:
        try:
            value = app.Caption
        except Exception:
            return None
        return value if isinstance(value, str) else str(value)

    @staticmethod
    def _read_hwnd(app: Any) -> int | None:
        try:
            hwnd = app.Hwnd
        except Exception:
            return None
        try:
            return int(hwnd)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _pid_from_hwnd(hwnd: int | None) -> int | None:
        """Best-effort process id for a window handle (non-COM, read-only)."""

        if hwnd is None or hwnd == 0:
            return None
        try:
            import ctypes
            from ctypes import wintypes
        except Exception:  # pragma: no cover - non-Windows fallback
            return None

        pid = wintypes.DWORD()
        try:
            user32 = ctypes.windll.user32
            user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
        except Exception:  # pragma: no cover - platform specific
            return None
        return int(pid.value) or None

    @staticmethod
    def _pid_alive(pid: int) -> bool:
        """Best-effort: is ``pid`` currently a live process? Never raises.

        Uses ``OpenProcess`` + ``GetExitCodeProcess`` (no window and no COM, so
        it works for a hidden instance). A defunct pid, an inaccessible handle
        or an unavailable Win32 layer reports ``False`` ("dead"): the caller then
        has nothing to terminate and never escalates to a kill.
        """

        if pid <= 0:
            return False
        try:
            import ctypes
            from ctypes import wintypes
        except Exception:  # pragma: no cover - non-Windows fallback
            return False

        process_query_limited_information = 0x1000
        still_active = 259
        try:
            kernel32 = ctypes.windll.kernel32
            handle = kernel32.OpenProcess(process_query_limited_information, False, int(pid))
            if not handle:
                return False
            try:
                exit_code = wintypes.DWORD()
                if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                    return False
                return exit_code.value == still_active
            finally:
                kernel32.CloseHandle(handle)
        except Exception:  # pragma: no cover - platform specific
            return False

    @staticmethod
    def _terminate_pid(pid: int) -> bool:
        """Terminate ONLY ``pid`` via ``taskkill /F /PID``; return success.

        Last resort once the grace period expired: the pid is the adapter's own
        ``DispatchEx`` instance, so the blast radius is exactly one process. The
        invocation mirrors ``macro_service._terminate_owned_pid``. Best-effort:
        never raises and never falls back to a wildcard image kill.
        """

        if pid <= 0:
            return False
        try:
            completed = subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
        except Exception:  # noqa: BLE001 - best-effort teardown
            return False
        return completed.returncode == 0

    @staticmethod
    def _read_with_default(app: Any, name: str, *, default: Any) -> Any:
        try:
            return getattr(app, name)
        except Exception:
            return default
