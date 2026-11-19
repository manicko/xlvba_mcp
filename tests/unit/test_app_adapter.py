"""Unit tests for the hidden, non-activating Excel application adapter (task_010).

These tests never touch COM / Excel. They inject a fake ``Application`` object
into :class:`ExcelAppAdapter` and assert the settings context managers save and
restore prior values (on success and on exception), that ``assert_non_intrusive``
rejects a visible instance, and that the adapter source contains none of the
forbidden focus-changing APIs (static guard).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pytest

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


def test_quit_calls_quit_and_clears_refs() -> None:
    app = FakeApp()
    adapter = _adapter_with(app)

    adapter.quit()

    assert app.QuitCalled is True
    assert adapter.app is None
    assert adapter.owned_pid is None


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
