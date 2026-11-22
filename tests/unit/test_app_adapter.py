"""Unit tests for the hidden, non-activating Excel application adapter (task_010).

These tests never touch COM / Excel. They inject a fake ``Application`` object
into :class:`ExcelAppAdapter` and assert the settings context managers save and
restore prior values (on success and on exception), that ``assert_non_intrusive``
rejects a visible instance, that ``quit`` enforces the death of the owned
process (bounded poll + pid-scoped fallback, FE-17), and that the adapter source
contains none of the forbidden focus-changing APIs (static guard).
"""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel import app_adapter as app_adapter_module
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter

_ADAPTER_SOURCE = (
    Path(__file__).resolve().parents[2]
    / "src"
    / "excel_vba_mcp"
    / "adapters"
    / "excel"
    / "app_adapter.py"
)


class FakeApp:
    """Minimal fake Excel ``Application`` recording property mutations."""

    def __init__(self, *, visible: bool = False) -> None:
        self.Visible = visible
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 3
        self.Caption = "Excel"
        self.Hwnd = 0
        self.QuitCalled = False
        self.Opened: list[dict[str, Any]] = []
        self.Closed: list[dict[str, Any]] = []

    def Quit(self) -> None:  # noqa: N802 - COM casing
        self.QuitCalled = True

    @property
    def Workbooks(self) -> FakeWorkbooks:  # noqa: N802 - COM casing
        return FakeWorkbooks(self)


class FakeWorkbook:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.closed = False
        self.saved: bool | None = None

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM casing
        self.closed = True
        self.saved = SaveChanges


class FakeWorkbooks:
    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self._items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(path.rsplit("/", 1)[-1].rsplit("\\", 1)[-1] or "book")
        self._app.Opened.append({"path": path, **kwargs})
        self._items.append(wb)
        return wb

    def __call__(self, name: str) -> FakeWorkbook:
        for wb in self._items:
            if wb.Name == name:
                return wb
        raise KeyError(name)

    def __iter__(self):
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)


def _adapter_with(app: FakeApp) -> ExcelAppAdapter:
    adapter = ExcelAppAdapter()
    adapter._app = app  # inject without DispatchEx
    adapter._owned_pid = 123
    adapter._owned_caption = app.Caption
    return adapter


# ---------------------------------------------------------------------------
# Fakes for the enforced owned-process death (FE-17)
# ---------------------------------------------------------------------------
class FakeReaper:
    """Fake PID-liveness + termination source (no Excel, no Win32, no process).

    ``alive`` is the fake world's answer to "is the owned pid running?"; a
    successful ``terminate_pid`` flips it to ``False`` so the post-kill
    verification succeeds. ``on_terminate`` lets a test observe adapter state at
    the moment the fallback runs.
    """

    def __init__(self, *, alive: bool, kill_succeeds: bool = True) -> None:
        self.alive = alive
        self.kill_succeeds = kill_succeeds
        self.alive_checks: list[int] = []
        self.terminate_calls: list[int] = []
        self.on_terminate: Callable[[], None] | None = None

    def pid_alive(self, pid: int) -> bool:
        self.alive_checks.append(pid)
        return self.alive

    def terminate_pid(self, pid: int) -> bool:
        self.terminate_calls.append(pid)
        if self.on_terminate is not None:
            self.on_terminate()
        if self.kill_succeeds:
            self.alive = False
        return self.kill_succeeds


def _install_reaper(monkeypatch: pytest.MonkeyPatch, reaper: FakeReaper) -> None:
    """Point the adapter's PID probes at ``reaper`` (class-level patch)."""

    monkeypatch.setattr(ExcelAppAdapter, "_pid_alive", staticmethod(reaper.pid_alive))
    monkeypatch.setattr(ExcelAppAdapter, "_terminate_pid", staticmethod(reaper.terminate_pid))


def _shrink_reap(monkeypatch: pytest.MonkeyPatch) -> None:
    """Shorten the reap budgets so the fallback path runs in milliseconds."""

    monkeypatch.setattr(app_adapter_module, "_REAP_TIMEOUT_S", 0.2)
    monkeypatch.setattr(app_adapter_module, "_REAP_INTERVAL_S", 0.01)
    monkeypatch.setattr(app_adapter_module, "_REAP_VERIFY_TIMEOUT_S", 0.2)


# ---------------------------------------------------------------------------
# Settings context managers restore prior value
# ---------------------------------------------------------------------------


def test_alerts_suppressed_restores_on_success() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with adapter.alerts_suppressed():
        assert app.DisplayAlerts is False

    assert app.DisplayAlerts is True


def test_alerts_suppressed_restores_on_exception() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with pytest.raises(RuntimeError), adapter.alerts_suppressed():
        assert app.DisplayAlerts is False
        raise RuntimeError("boom")

    assert app.DisplayAlerts is True


def test_screenupdating_off_restores_on_success() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with adapter.screenupdating_off():
        assert app.ScreenUpdating is False

    assert app.ScreenUpdating is True


def test_screenupdating_off_restores_on_exception() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with pytest.raises(ValueError), adapter.screenupdating_off():
        assert app.ScreenUpdating is False
        raise ValueError("boom")

    assert app.ScreenUpdating is True


def test_automation_security_restores_on_success() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with adapter.automation_security(1):
        assert app.AutomationSecurity == 1

    assert app.AutomationSecurity == 3


def test_automation_security_restores_on_exception() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    with pytest.raises(KeyError), adapter.automation_security(2):
        assert app.AutomationSecurity == 2
        raise KeyError("boom")

    assert app.AutomationSecurity == 3


# ---------------------------------------------------------------------------
# assert_non_intrusive
# ---------------------------------------------------------------------------


def test_assert_non_intrusive_passes_when_hidden() -> None:
    app = FakeApp(visible=False)
    adapter = _adapter_with(app)

    adapter.assert_non_intrusive()  # must not raise


def test_assert_non_intrusive_raises_when_visible() -> None:
    app = FakeApp(visible=True)
    adapter = _adapter_with(app)

    with pytest.raises(RuntimeError, match="Visible=True"):
        adapter.assert_non_intrusive()


def test_assert_non_intrusive_noop_without_app() -> None:
    adapter = ExcelAppAdapter()

    adapter.assert_non_intrusive()  # no app yet -> no-op, must not raise


# ---------------------------------------------------------------------------
# Workbook open / close / quit
# ---------------------------------------------------------------------------


def test_open_workbook_passes_read_only_and_link_args() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    adapter.open_workbook("/tmp/Book.xlsm", read_only=True, update_links=0)

    assert app.Opened[0]["path"] == "/tmp/Book.xlsm"
    assert app.Opened[0]["ReadOnly"] is True
    assert app.Opened[0]["UpdateLinks"] == 0


def test_close_workbook_defaults_to_not_saving() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)
    wb = adapter.open_workbook("/tmp/Book.xlsm", read_only=False)

    adapter.close_workbook(wb, save_changes=False)

    assert wb.closed is True
    assert wb.saved is False


def test_quit_calls_quit_and_clears_refs(monkeypatch: pytest.MonkeyPatch) -> None:
    app = FakeApp()
    adapter = _adapter_with(app)
    reaper = FakeReaper(alive=False)
    _install_reaper(monkeypatch, reaper)

    adapter.quit()

    assert app.QuitCalled is True
    assert adapter.app is None
    assert adapter.owned_pid is None
    assert reaper.terminate_calls == [], "a dead process must never be terminated"


# ---------------------------------------------------------------------------
# quit(): enforced owned-process death (FE-17)
# ---------------------------------------------------------------------------


def test_quit_skips_the_fallback_when_quit_ends_the_process(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A Quit that works short-circuits the pid-scoped fallback."""

    app = FakeApp()
    adapter = _adapter_with(app)
    reaper = FakeReaper(alive=False)
    _install_reaper(monkeypatch, reaper)

    adapter.quit()

    assert reaper.terminate_calls == [], "the fallback must not run when Quit ended the process"
    assert reaper.alive_checks, "the death must be verified, not assumed"
    assert set(reaper.alive_checks) == {123}, (
        f"only the owned pid may be probed: {reaper.alive_checks}"
    )
    assert adapter.owned_pid is None


def test_quit_falls_back_to_a_pid_scoped_kill_when_quit_is_a_noop(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A no-op Quit leaves the process resident: the fallback terminates the pid."""

    app = FakeApp()
    adapter = _adapter_with(app)
    reaper = FakeReaper(alive=True)
    _install_reaper(monkeypatch, reaper)
    _shrink_reap(monkeypatch)

    adapter.quit()

    assert app.QuitCalled is True
    assert reaper.terminate_calls == [123], "the fallback must terminate only the owned pid"
    assert adapter.owned_pid is None, "the pid is cleared only after the fallback attempt"


def test_quit_clears_the_owned_pid_only_after_verification(monkeypatch: pytest.MonkeyPatch) -> None:
    """``_owned_pid`` survives until the verified shutdown has run."""

    app = FakeApp()
    adapter = _adapter_with(app)
    observed: list[int | None] = []
    reaper = FakeReaper(alive=True)
    reaper.on_terminate = lambda: observed.append(adapter.owned_pid)
    _install_reaper(monkeypatch, reaper)
    _shrink_reap(monkeypatch)

    adapter.quit()

    assert observed == [123], "_owned_pid must still be set while the fallback runs"
    assert adapter.owned_pid is None


def test_quit_clears_the_owned_pid_after_a_failed_fallback(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A failed kill still clears the refs (best effort) and never raises."""

    app = FakeApp()
    adapter = _adapter_with(app)
    reaper = FakeReaper(alive=True, kill_succeeds=False)
    _install_reaper(monkeypatch, reaper)
    _shrink_reap(monkeypatch)

    adapter.quit()

    assert reaper.terminate_calls == [123]
    assert adapter.owned_pid is None
    assert adapter.owned_caption is None
    assert adapter.app is None


# ---------------------------------------------------------------------------
# Static invariant: no focus-changing APIs in the adapter source
# ---------------------------------------------------------------------------

_FORBIDDEN = [
    ".Activate",
    ".Select",
    "SetForegroundWindow",
    "ShowWindow",
    "SendKeys",
]


def _code_only(source: str) -> str:
    """Strip docstrings/comments so prose mentions do not trip the static check."""

    import ast
    import io
    import tokenize

    # Remove comments and string-literal tokens (docstrings included).
    result: list[str] = []
    tokens = tokenize.generate_tokens(io.StringIO(source).readline)
    for tok in tokens:
        if tok.type in (tokenize.COMMENT, tokenize.STRING):
            continue
        result.append(tok.string)
    # Validate it still parses (sanity); ignore return.
    ast.parse(source)
    return " ".join(result)


def test_adapter_source_has_no_forbidden_focus_apis() -> None:
    text = _code_only(_ADAPTER_SOURCE.read_text(encoding="utf-8"))
    for token in _FORBIDDEN:
        assert token not in text, f"app_adapter.py must not contain {token!r}"


def test_adapter_source_forbids_setting_visible_true() -> None:
    text = _code_only(_ADAPTER_SOURCE.read_text(encoding="utf-8"))
    # No assignment of Visible to True in actual code (docstrings may mention it).
    assert re.search(r"\.Visible\s*=\s*True", text) is None
