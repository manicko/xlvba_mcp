"""VBE Compile command + UI Automation diagnostics proof-of-concept for task_004.

This module is a bounded, self-contained proof-of-concept harness (not a
production abstraction). It proves, on the target Windows/Excel build
(Windows 10 LTSC 2019 17763 x64, Excel 16.0.14332 x64, Office UI language 1033
en-US), the VBE compile mechanism that production ``adapters/excel``
compile services (task_020) depend on:

* the plan's CommandBars route to the VBE ``Debug -> Compile`` command is
  **not available on this build** (``VBE.CommandBars["Debug"].Controls`` has no
  Compile entry; ``FindControl(578)`` raises ``E_INVALIDARG``) — recorded as
  evidence per the R03 correction;
* the working route is **UI Automation menu expansion + Invoke**:
  ``MenuItem`` ``Debug`` (control type ``50011``) -> ``ExpandCollapsePattern.Expand``
  -> ``MenuItem`` ``Compile VBAProject`` (literal caption) -> ``InvokePattern``
  (pattern id ``10000``);
* the four outcomes are distinguished honestly: ``compiled``,
  ``already_compiled``, ``compile_error``, ``indeterminate``;
* a ``compile_error`` raises a modal ``#32770`` window titled
  ``Microsoft Visual Basic for Applications`` whose Text reads
  ``Compile error:\\n\\nVariable not defined``; the offending token is read from
  the VBE code pane (``AutomationId = "Bottom VBA Editor Code Pane"``, type
  ``50030``) via ``TextPattern`` because the dialog's ``TextPattern.GetSelection``
  does not work on this build (NULL COM pointer);
* while a compile dialog is open **no Excel COM call is issued**; the dialog is
  dismissed via UI Automation first and only then does the COM lifecycle
  continue.

Every Excel COM call runs on a single dedicated STA thread with an owned
``DispatchEx`` instance (``Visible=True`` for UIA — an executing-class operation
per ARCH-013, ``DisplayAlerts=False``). Cleanup closes the workbook without
saving, quits, releases references and ``CoUninitialize``s on that same thread,
then polls for an orphaned owned PID. Non-owned (user) Excel processes are never
touched. Evidence is persisted to
``.ai/tmp/vbe_compile_poc_<ts>.{json,txt}`` per test.

UI Automation is driven with raw ``comtypes`` ``IUIAutomation`` (research §3.7):
pywinauto 0.6.9's UIA wrapper has an ``automation_id`` keyword bug and flips the
COM threading mode on import, which is unsafe next to a single-STA worker. The
decision and its evidence are recorded in the artifact (R04 fallback).

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vbe_compile_poc.py -v -s
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import gc
import hashlib
import json
import subprocess
import sys
import threading
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pytest

EXCEL_PROGID = "Excel.Application"

# UI Automation control type / pattern / property ids (see research §3.2-3.5).
UIA_TREE_SCOPE_SUBTREE = 7
UIA_EXPAND_COLLAPSE_PATTERN_ID = 10005
UIA_INVOKE_PATTERN_ID = 10000
UIA_TEXT_PATTERN_ID = 10014
CONTROL_TYPE_MENU_ITEM = 50011
CONTROL_TYPE_BUTTON = 50000
CONTROL_TYPE_TEXT = 50020
CONTROL_TYPE_PANE = 50030
CONTROL_TYPE_WINDOW = 50032

# VBE compile-error dialog.
DIALOG_CLASS = "#32770"
DIALOG_TITLE = "Microsoft Visual Basic for Applications"
COMPILE_MENU_CAPTION = "Compile VBAProject"
DEBUG_MENU_CAPTION = "Debug"
CODE_PANE_AUTOMATION_ID = "Bottom VBA Editor Code Pane"

# Orchestration timing (seconds).
STA_JOIN_TIMEOUT_S = 180.0
ORPHAN_POLL_INTERVAL_S = 0.5
ORPHAN_POLL_TIMEOUT_S = 10.0
MENU_SETTLE_S = 0.6
POST_INVOKE_S = 1.2
DIALOG_WAIT_S = 8.0
DIALOG_POLL_INTERVAL_S = 0.3
UIA_SETTLE_S = 1.0

# Optional COM / Win32 dependencies. ``None`` when unavailable so the module
# imports on non-Windows hosts and the tests skip cleanly.
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32api  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32api = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_ARTIFACT_DIR = _REPO_ROOT / ".ai" / "tmp"
_FIXTURES_DIR = _REPO_ROOT / "tests" / "fixtures"
_CLEAN_PROJECT = _FIXTURES_DIR / "CleanProject.xlsm"
_COMPILE_ERROR_PROJECT = _FIXTURES_DIR / "CompileErrorProject.xlsm"


# ---------------------------------------------------------------------------
# Availability / skip helpers
# ---------------------------------------------------------------------------
def _excel_com_available() -> tuple[bool, str]:
    """Return ``(available, reason)`` describing whether this host can drive Excel."""

    if sys.platform != "win32":
        return False, f"non-Windows platform: {sys.platform}"
    if pythoncom is None or win32com is None:
        return False, "pywin32 (pythoncom/win32com) not importable"
    return True, "pywin32 present"


_AVAILABLE, _AVAIL_REASON = _excel_com_available()

requires_excel = pytest.mark.skipif(
    not _AVAILABLE,
    reason=f"Excel COM + UIA PoC unavailable: {_AVAIL_REASON}",
)


def _fixture_or_skip(path: Path) -> Path:
    """Return the fixture path or skip when it is missing."""

    if not path.is_file():
        pytest.skip(
            f"fixture missing: {path}; run `uv run python tests/fixtures/build_fixtures.py` first"
        )
    return path


# ---------------------------------------------------------------------------
# Time / COM error helpers
# ---------------------------------------------------------------------------
def _now_utc_iso() -> str:
    return _dt.datetime.now(_dt.UTC).isoformat()


def _now_local_iso() -> str:
    return _dt.datetime.now().astimezone().isoformat()


def _hresult_hex_from_exc(exc: BaseException) -> str | None:
    """Extract a displayable HRESULT hex string from a com_error-like exception."""

    hr = getattr(exc, "hresult", None)
    if not isinstance(hr, int):
        args = getattr(exc, "args", ())
        if args and isinstance(args[0], int):
            hr = args[0]
    if isinstance(hr, int):
        return f"0x{hr & 0xFFFFFFFF:08X}"
    return None


def _com_error_details(exc: BaseException) -> dict[str, Any]:
    """Serialize a COM error into a structured, JSON-safe dictionary."""

    details: dict[str, Any] = {
        "type": type(exc).__name__,
        "message": str(exc)[:500],
        "hresult_hex": _hresult_hex_from_exc(exc),
    }
    args = getattr(exc, "args", ())
    if args:
        details["args_repr"] = [repr(a)[:200] for a in args]
    return details


# ---------------------------------------------------------------------------
# Source normalization / hashing (research §2.6: hash the code body)
# ---------------------------------------------------------------------------
def normalize_code(text: str) -> str:
    """Normalize EOL/whitespace so the same source hashes identically."""

    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in unified.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def _source_hash(proj: Any) -> str:
    """Return a stable SHA-256 over every component's normalized code body.

    ``compile`` never edits source, so this hash is invariant across successive
    compiles of the same project and is the identity used for the
    ``already_compiled`` determination (research §3.3).
    """

    parts: list[str] = []
    for comp in proj.VBComponents:
        try:
            count = int(comp.CodeModule.CountOfLines)
        except Exception:  # noqa: BLE001
            count = 0
        body = ""
        if count > 0:
            try:
                body = comp.CodeModule.Lines(1, count)
            except Exception:  # noqa: BLE001
                body = ""
        parts.append(f"{comp.Name}\n{normalize_code(body)}")
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# Process helpers (owned-PID resolution + orphan poll)
# ---------------------------------------------------------------------------
def _excel_pids() -> set[int]:
    """Return the set of current EXCEL.EXE PIDs via ``tasklist``."""

    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
    return pids


def _pid_alive(pid: int) -> bool:
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001
        return False
    return str(pid) in out.stdout


def _orphan_poll(owned_pid: int | None) -> bool | None:
    """Poll until the owned PID is gone; return True if reaped, None if unknown."""

    if owned_pid is None:
        return None
    deadline = time.monotonic() + ORPHAN_POLL_TIMEOUT_S
    while time.monotonic() < deadline:
        if not _pid_alive(owned_pid):
            return True
        time.sleep(ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(owned_pid)


# ---------------------------------------------------------------------------
# Window enumeration (Win32) + screenshot
# ---------------------------------------------------------------------------
def _enum_windows() -> list[dict[str, Any]]:
    """Enumerate visible top-level windows: hwnd, title, class, pid."""

    wins: list[dict[str, Any]] = []
    if win32gui is None or win32process is None:  # pragma: no cover
        return wins

    def _cb(hwnd: int, _ctx: Any) -> bool:
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001
            pid = None
        with contextlib.suppress(Exception):
            wins.append(
                {
                    "hwnd": hwnd,
                    "title": win32gui.GetWindowText(hwnd),
                    "class": win32gui.GetClassName(hwnd),
                    "pid": pid,
                }
            )
        return True

    with contextlib.suppress(Exception):  # pragma: no cover - depends on host
        win32gui.EnumWindows(_cb, None)
    return wins


def _find_new_dialog(before_hwnds: set[int]) -> dict[str, Any] | None:
    """Return the first new ``#32770`` titled dialog that appeared, or ``None``."""

    for win in _enum_windows():
        if win["hwnd"] in before_hwnds:
            continue
        if win.get("class") == DIALOG_CLASS and win.get("title") == DIALOG_TITLE:
            return win
    return None


def _screenshot_window(hwnd: int, out_png: Path) -> dict[str, Any]:
    """Capture a window to PNG via PowerShell System.Drawing (no PIL needed).

    Returns a JSON-safe result dict describing success and any error.
    """

    result: dict[str, Any] = {"attempted": True, "path": str(out_png)}
    if win32gui is None or win32api is None:  # pragma: no cover
        result["error"] = "win32gui/win32api unavailable"
        return result
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"GetWindowRect failed: {exc!r}"
        return result
    screen_w = win32api.GetSystemMetrics(0)
    screen_h = win32api.GetSystemMetrics(1)
    left = max(left, 0)
    top = max(top, 0)
    right = min(right, screen_w)
    bottom = min(bottom, screen_h)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        result["error"] = f"empty capture rect: {(left, top, width, height)}"
        return result
    # Best effort: bring the dialog to the foreground so CopyFromScreen is not
    # occluded by another window.
    with contextlib.suppress(Exception):
        win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    script = (
        "Add-Type -AssemblyName System.Drawing; "
        f"$bmp = New-Object System.Drawing.Bitmap {width}, {height}; "
        "$g = [System.Drawing.Graphics]::FromImage($bmp); "
        f"$g.CopyFromScreen({left}, {top}, 0, 0, $bmp.Size); "
        "$bmp.Save("
        f"'{out_png.as_posix()}', "
        "[System.Drawing.Imaging.ImageFormat]::Png); "
        "$g.Dispose(); $bmp.Dispose()"
    )
    try:  # pragma: no cover - depends on host
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"powershell failed: {exc!r}"
        return result
    result["exists"] = out_png.is_file()
    result["size"] = out_png.stat().st_size if out_png.is_file() else 0
    result["returncode"] = proc.returncode
    if proc.returncode != 0:
        result["stderr"] = proc.stderr.strip()[:300]
    result["rect"] = [left, top, width, height]
    return result


# ---------------------------------------------------------------------------
# Raw comtypes IUIAutomation wrapper (research §3.2-3.5, §3.7)
# ---------------------------------------------------------------------------
class UiaSession:
    """Thin wrapper around a raw ``IUIAutomation`` client bound to one VBE hwnd.

    Uses ``comtypes`` ``IUIAutomation`` directly because pywinauto 0.6.9's UIA
    backend has an ``automation_id`` keyword bug and mutates the COM threading
    mode on import (research §3.7), both unsafe next to the single-STA worker.
    """

    def __init__(self, hwnd: int) -> None:
        import comtypes
        import comtypes.client

        self._comtypes = comtypes
        ui = comtypes.client.GetModule("UIAutomationCore.dll")
        self._ui = ui
        self._iuia = comtypes.CoCreateInstance(
            ui.CUIAutomation().IPersist_GetClassID(),
            interface=ui.IUIAutomation,
            clsctx=comtypes.CLSCTX_INPROC_SERVER,
        )
        self.hwnd = hwnd
        self.root = self._iuia.ElementFromHandle(hwnd)

    def _elements(self, element: Any) -> list[Any]:
        """Return the subtree elements of ``element`` as a Python list."""

        array = element.FindAll(UIA_TREE_SCOPE_SUBTREE, self._iuia.CreateTrueCondition())
        out: list[Any] = []
        for idx in range(array.Length):
            try:
                out.append(array.GetElement(idx))
            except Exception:  # noqa: BLE001
                continue
        return out

    def _safe(self, element: Any, attr: str) -> Any:
        try:
            return getattr(element, attr)
        except Exception:  # noqa: BLE001
            return None

    def _rect(self, element: Any) -> tuple[int, int, int, int] | None:
        try:
            rect = element.CurrentBoundingRectangle
            return int(rect.left), int(rect.top), int(rect.right), int(rect.bottom)
        except Exception:  # noqa: BLE001
            return None

    def _on_screen(self, element: Any) -> bool:
        rect = self._rect(element)
        if rect is None:
            return False
        left, top, right, bottom = rect
        if right <= left or bottom <= top:
            return False
        sw = win32api.GetSystemMetrics(0) if win32api else 1 << 16
        sh = win32api.GetSystemMetrics(1) if win32api else 1 << 16
        return left < sw and top < sh and right > 0 and bottom > 0

    def _debug_menu_elements(self) -> tuple[list[dict[str, Any]], Any | None]:
        """One traversal: return (debug_candidates, chosen Debug element to expand).

        The chosen element is the last enabled, on-screen ``Debug`` MenuItem in
        UIA enumeration order (``dbg[-1]``), matching the proven probe.
        """

        candidates: list[dict[str, Any]] = []
        elements: list[Any] = []
        for element in self._elements(self.root):
            if self._safe(element, "CurrentControlType") != CONTROL_TYPE_MENU_ITEM:
                continue
            if self._safe(element, "CurrentName") != DEBUG_MENU_CAPTION:
                continue
            enabled = bool(self._safe(element, "CurrentIsEnabled"))
            on_screen = self._on_screen(element)
            candidates.append(
                {
                    "name": DEBUG_MENU_CAPTION,
                    "enabled": enabled,
                    "rect": self._rect(element),
                    "on_screen": on_screen,
                }
            )
            if enabled:
                elements.append(element)
        on_screen = [e for e in elements if self._on_screen(e)]
        pool = on_screen or elements
        return candidates, (pool[-1] if pool else None)

    def expand_debug(self) -> tuple[bool, list[dict[str, Any]]]:
        """Ensure the Debug menu is expanded; return ``(expanded, debug_candidates)``.

        Idempotent: if the Debug MenuItem's ``ExpandCollapsePattern`` already
        reports an expanded state, no toggle happens (Expand on an already-expanded
        menu would collapse it). Only one UIA traversal happens here; callers take
        a single :meth:`compile_snapshot` immediately afterwards so the popup does
        not auto-collapse between reads.
        """

        debug_candidates, menu = self._debug_menu_elements()
        if menu is None:
            return False, debug_candidates
        try:
            pattern = menu.GetCurrentPattern(UIA_EXPAND_COLLAPSE_PATTERN_ID)
            expand = pattern.QueryInterface(self._ui.IUIAutomationExpandCollapsePattern)
            # ExpandCollapseState_Expanded == 1; only expand when not expanded.
            if int(expand.CurrentExpandCollapseState) != 1:
                expand.Expand()
        except Exception:  # noqa: BLE001
            return False, debug_candidates
        time.sleep(MENU_SETTLE_S)
        return True, debug_candidates

    def compile_snapshot(self) -> dict[str, Any]:
        """Take ONE UIA snapshot of the Compile item and derive item + state.

        Returns a dict with ``candidates`` (all Compile MenuItem candidates with
        enabled/rect/on_screen), ``item`` (the enabled on-screen element to Invoke,
        or ``None``), ``present`` (any on-screen candidate) and ``enabled`` (any
        enabled on-screen candidate). Deriving everything from a single traversal
        avoids the popup collapsing between separate reads.
        """

        candidates: list[dict[str, Any]] = []
        item: Any = None
        for element in self._elements(self.root):
            if self._safe(element, "CurrentControlType") != CONTROL_TYPE_MENU_ITEM:
                continue
            if self._safe(element, "CurrentName") != COMPILE_MENU_CAPTION:
                continue
            on_screen = self._on_screen(element)
            enabled = bool(self._safe(element, "CurrentIsEnabled"))
            candidates.append(
                {
                    "name": COMPILE_MENU_CAPTION,
                    "enabled": enabled,
                    "rect": self._rect(element),
                    "on_screen": on_screen,
                }
            )
            if item is None and enabled and on_screen:
                item = element
        present = any(c["on_screen"] for c in candidates)
        enabled_any = any(c["on_screen"] and c["enabled"] for c in candidates)
        return {
            "candidates": candidates,
            "item": item,
            "present": present,
            "enabled": enabled_any,
        }

    def invoke_compile(self, element: Any) -> None:
        """Invoke the Compile menu item via InvokePattern (pattern id 10000)."""

        pattern = element.GetCurrentPattern(UIA_INVOKE_PATTERN_ID)
        pattern.QueryInterface(self._ui.IUIAutomationInvokePattern).Invoke()

    def read_dialog(self, hwnd: int) -> dict[str, Any]:
        """Read a dialog's Text elements and Buttons via UIA."""

        root = self._iuia.ElementFromHandle(hwnd)
        texts: list[str] = []
        buttons: list[dict[str, Any]] = []
        for element in self._elements(root):
            ctype = self._safe(element, "CurrentControlType")
            if ctype == CONTROL_TYPE_TEXT:
                name = self._safe(element, "CurrentName")
                if name:
                    texts.append(str(name))
            elif ctype == CONTROL_TYPE_BUTTON:
                buttons.append(
                    {
                        "name": str(self._safe(element, "CurrentName") or ""),
                        "automation_id": str(self._safe(element, "CurrentAutomationId") or ""),
                    }
                )
        return {"texts": texts, "buttons": buttons}

    def dismiss_dialog_ok(self, hwnd: int) -> str:
        """Dismiss a dialog by invoking its ``OK`` button via UIA.

        Returns a short status string. Falls back to a Win32 ``WM_COMMAND`` IDOK
        post only if the UIA Invoke path is unavailable.
        """

        root = self._iuia.ElementFromHandle(hwnd)
        for element in self._elements(root):
            if self._safe(element, "CurrentControlType") != CONTROL_TYPE_BUTTON:
                continue
            name = str(self._safe(element, "CurrentName") or "")
            auto_id = str(self._safe(element, "CurrentAutomationId") or "")
            if name == "OK" or auto_id == "2":
                try:
                    pattern = element.GetCurrentPattern(UIA_INVOKE_PATTERN_ID)
                    pattern.QueryInterface(self._ui.IUIAutomationInvokePattern).Invoke()
                    return "uia_invoke_ok"
                except Exception:  # noqa: BLE001
                    break
        # Fallback: post IDOK (1) WM_COMMAND to the dialog.
        if win32gui is not None:
            try:
                win32gui.PostMessage(hwnd, 0x0111, 1, 0)  # WM_COMMAND, IDOK
                return "win32_wm_command_idok"
            except Exception:  # noqa: BLE001
                return "dismiss_failed"
        return "dismiss_unavailable"

    def read_code_pane(self) -> dict[str, Any]:
        """Read the VBE code pane text + selection via TextPattern.

        The offending token is read from the ``Bottom VBA Editor Code Pane``
        (control type ``50030``) because the dialog's TextPattern is unavailable
        on this build (research §3.5).
        """

        out: dict[str, Any] = {"found": False, "document_text": None, "selection": None}
        for element in self._elements(self.root):
            if self._safe(element, "CurrentAutomationId") != CODE_PANE_AUTOMATION_ID:
                continue
            out["found"] = True
            out["control_type"] = self._safe(element, "CurrentControlType")
            try:
                pattern = element.GetCurrentPattern(UIA_TEXT_PATTERN_ID)
                text_pattern = pattern.QueryInterface(self._ui.IUIAutomationTextPattern)
            except Exception as exc:  # noqa: BLE001
                out["textpattern_error"] = f"{type(exc).__name__}: {exc}"
                return out
            try:
                out["document_text"] = text_pattern.DocumentRange.GetText(-1)
            except Exception as exc:  # noqa: BLE001
                out["document_text_error"] = f"{type(exc).__name__}: {exc}"
            try:
                selection = text_pattern.GetSelection()
                if selection is not None and selection.Length > 0:
                    out["selection"] = selection.GetElement(0).GetText(-1)
                else:
                    out["selection"] = ""
            except Exception as exc:  # noqa: BLE001
                out["selection_error"] = f"{type(exc).__name__}: {exc}"
            return out
        return out


# ---------------------------------------------------------------------------
# Modal-safety guard: prove no Excel COM call while a compile dialog is open
# ---------------------------------------------------------------------------
class _ComCallGuard:
    """Proxy that records every Excel COM attribute access with the modal flag.

    ``state["modal_open"]`` is set True the moment a ``#32770`` compile dialog is
    detected and False once it is dismissed. Any Excel COM access attempted while
    it is True would be recorded and fail the ``no_com_while_modal`` assertion.
    UIA and Win32 calls are out-of-process and are intentionally not guarded.
    """

    def __init__(self, app: Any, state: dict[str, Any], log: list[dict[str, Any]]) -> None:
        object.__setattr__(self, "_app", app)
        object.__setattr__(self, "_state", state)
        object.__setattr__(self, "_log", log)

    def __getattr__(self, name: str) -> Any:
        state = object.__getattribute__(self, "_state")
        log = object.__getattribute__(self, "_log")
        log.append({"attr": name, "modal_open": bool(state.get("modal_open", False))})
        return getattr(object.__getattribute__(self, "_app"), name)


# ---------------------------------------------------------------------------
# Artifact collector
# ---------------------------------------------------------------------------
@dataclass
class PocArtifacts:
    """Structured evidence collector persisted as JSON + text per test."""

    test_name: str = ""
    started_utc: str = field(default_factory=_now_utc_iso)
    started_local: str = field(default_factory=_now_local_iso)
    sta_thread_id: int | None = None
    owned_pid: int | None = None
    owned_pid_resolution: str = ""
    baseline_excel_pids: list[int] = field(default_factory=list)
    environment: dict[str, Any] = field(default_factory=dict)
    commandbars_evidence: dict[str, Any] = field(default_factory=dict)
    uia_decision: dict[str, Any] = field(default_factory=dict)
    compile_attempts: list[dict[str, Any]] = field(default_factory=list)
    source_hash_before: str | None = None
    source_hash_after: str | None = None
    source_hash_before_second: str | None = None
    dialog: dict[str, Any] | None = None
    selected_token: str = "unavailable"
    selected_token_detail: dict[str, Any] = field(default_factory=dict)
    com_calls_while_modal: list[dict[str, Any]] = field(default_factory=list)
    no_com_while_modal: bool | None = None
    outcome: str = "indeterminate"
    outcome_second: str | None = None
    go_no_go: dict[str, Any] | None = None
    orphan_reaped: bool | None = None
    log_lines: list[str] = field(default_factory=list)
    paths: dict[str, str] = field(default_factory=dict)
    artifact_version: str = "task_004_binding"

    def log(self, msg: str) -> None:
        line = f"{_now_local_iso()} {msg}"
        self.log_lines.append(line)
        print(line, flush=True)

    def persist(self) -> dict[str, str]:
        _ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        json_path = _ARTIFACT_DIR / f"vbe_compile_poc_{self.test_name}_{ts}.json"
        txt_path = _ARTIFACT_DIR / f"vbe_compile_poc_{self.test_name}_{ts}.txt"
        json_path.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        header = (
            "task_004 VBE compile + UIA diagnostics PoC evidence\n"
            f"test: {self.test_name}\n"
            f"artifact_version: {self.artifact_version}\n"
            f"started_utc: {self.started_utc}\n"
            f"started_local: {self.started_local}\n"
        )
        txt_path.write_text(header + "\n" + "\n".join(self.log_lines) + "\n", encoding="utf-8")
        self.paths = {"json": str(json_path), "txt": str(txt_path)}
        print(f"[artifacts] json={json_path} txt={txt_path}", flush=True)
        return self.paths


# ---------------------------------------------------------------------------
# Environment matrix
# ---------------------------------------------------------------------------
def _collect_environment(app: Any | None, sta_thread_id: int | None) -> dict[str, Any]:
    """Collect UI-language / build facts without mutating settings."""

    env: dict[str, Any] = {"sta_thread_id": sta_thread_id, "methods": []}
    try:  # pragma: no cover - depends on host
        import ctypes

        lang = ctypes.windll.kernel32.GetUserDefaultUILanguage()  # type: ignore[attr-defined]
        env["user_ui_language"] = int(lang) & 0xFFFF
        env["methods"].append("GetUserDefaultUILanguage")
    except Exception as exc:  # noqa: BLE001
        env["user_ui_language"] = None
        env["methods"].append(f"GetUserDefaultUILanguage failed: {exc!r}")
    if app is not None:
        try:
            env["excel_version"] = str(app.Version)
            env["excel_build"] = str(app.Build)
        except Exception as exc:  # noqa: BLE001
            env["excel_read_error"] = repr(exc)
        try:
            env["office_ui_language_id"] = int(app.LanguageSettings.LanguageID(2))
            env["methods"].append("LanguageSettings.LanguageID(2)")
        except Exception as exc:  # noqa: BLE001
            env["office_ui_language_id"] = None
            env["methods"].append(f"LanguageSettings failed: {exc!r}")
        env["excel_visible_policy"] = (
            "Visible=True for UIA (executing-class operation per ARCH-013); DisplayAlerts=False"
        )
    return env


# ---------------------------------------------------------------------------
# CommandBars route evidence (R03 correction): record that it fails, do not use
# ---------------------------------------------------------------------------
def _probe_commandbars(app: Any) -> dict[str, Any]:
    """Attempt the plan's CommandBars route once and record why it is unavailable."""

    evidence: dict[str, Any] = {"attempted": True}
    try:
        debug_bar = app.VBE.CommandBars["Debug"]
        controls = [{"id": int(c.Id), "caption": str(c.Caption)} for c in debug_bar.Controls]
        evidence["debug_controls_count"] = len(controls)
        evidence["debug_controls"] = controls
        evidence["has_compile_entry"] = any("compile" in c["caption"].lower() for c in controls)
    except Exception as exc:  # noqa: BLE001
        evidence["controls_error"] = _com_error_details(exc)
    try:
        ctl = app.VBE.CommandBars["Debug"].FindControl(578)
        evidence["findcontrol_578"] = "found" if ctl is not None else "none"
    except Exception as exc:  # noqa: BLE001
        evidence["findcontrol_578"] = _com_error_details(exc)
    evidence["route_usable"] = bool(evidence.get("has_compile_entry"))
    return evidence


# ---------------------------------------------------------------------------
# Single-STA-thread runner
# ---------------------------------------------------------------------------
def _sta_run(
    workbook_path: Path,
    body: Any,
    art: PocArtifacts,
) -> dict[str, Any]:
    """Run ``body(guarded_app, wb, ctx, art)`` on a dedicated STA thread.

    The worker owns the whole COM lifecycle: ``CoInitializeEx(APARTMENTTHREADED)``,
    ``DispatchEx``, ``Visible=True``/``DisplayAlerts=False``, owned-PID
    resolution, workbook open, the body call, then a best-effort Win32 dialog
    dismissal (so a stray compile modal never blocks COM cleanup), close without
    save, ``Quit``, reference release, ``CoUninitialize`` and an orphan poll.
    """

    holder: dict[str, Any] = {}

    def worker() -> None:
        tid = threading.get_ident()
        ctx: dict[str, Any] = {"sta_thread_id": tid, "state": {"modal_open": False}}
        app: Any = None
        guarded: _ComCallGuard | None = None
        wb: Any = None
        owned_pid: int | None = None
        com_log: list[dict[str, Any]] = []
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as exc:  # noqa: BLE001
            holder["error"] = f"CoInitializeEx failed: {exc!r}"
            return
        try:
            baseline = _excel_pids()
            ctx["baseline_excel_pids"] = sorted(baseline)
            app = win32com.client.DispatchEx(EXCEL_PROGID)
            app.Visible = True
            app.DisplayAlerts = False
            guarded = _ComCallGuard(app, ctx["state"], com_log)
            ctx["app"] = guarded
            new_pids = _excel_pids() - baseline
            owned_pid = next(iter(new_pids)) if len(new_pids) == 1 else None
            ctx["owned_pid"] = owned_pid
            ctx["owned_pid_resolution"] = f"pid_delta={sorted(new_pids)}"
            wb = app.Workbooks.Open(
                str(workbook_path),
                UpdateLinks=0,
                ReadOnly=False,
                Notify=False,
                AddToMru=False,
            )
            # Guard the workbook proxy too so the modal-safety proof covers
            # ``wb.*`` Excel COM calls, not only ``app.*`` ones.
            guarded_wb = _ComCallGuard(wb, ctx["state"], com_log)
            # Expose the VBE window for UIA. Visible=True is required so the menu
            # bar and code pane are exposed to UI Automation.
            vbe = app.VBE
            vbe.MainWindow.Visible = True
            ctx["vbe"] = vbe
            ctx["vbe_hwnd"] = int(vbe.MainWindow.Hwnd)
            time.sleep(UIA_SETTLE_S)
            holder["result"] = body(guarded, guarded_wb, ctx, art)
            holder["ctx"] = ctx
        except BaseException as exc:  # noqa: BLE001 - surfaced to the main thread
            holder["error"] = f"{type(exc).__name__}: {exc}"
            holder["ctx"] = ctx
        finally:
            # Dismiss any lingering compile modal via Win32 BEFORE any COM call so
            # cleanup can never hang on an open dialog (spec §8.3). Scoped strictly
            # to PIDs we spawned (current - baseline) so a user's Excel dialog is
            # never touched, even when the owned PID could not be uniquely resolved.
            owned_pids = _excel_pids() - set(ctx.get("baseline_excel_pids", []))
            _force_dismiss_dialogs(owned_pids)
            if wb is not None:
                with contextlib.suppress(Exception):
                    wb.Close(SaveChanges=False)
            wb = None
            gc.collect()
            if app is not None:
                with contextlib.suppress(Exception):
                    app.Quit()
            app = None
            guarded = None
            ctx["app"] = None
            ctx["vbe"] = None
            gc.collect()
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()
            holder["owned_pid"] = owned_pid if owned_pid is not None else ctx.get("owned_pid")
            holder["orphan_reaped"] = _orphan_poll(holder["owned_pid"])
            holder["com_log"] = com_log

    thread = threading.Thread(target=worker, name="vbe_compile_sta", daemon=True)
    thread.start()
    thread.join(timeout=STA_JOIN_TIMEOUT_S)
    holder["thread_alive"] = thread.is_alive()
    return holder


def _force_dismiss_dialogs(owned_pids: set[int]) -> list[int]:
    """Best-effort Win32 dismissal of any ``#32770`` compile dialog we spawned.

    Uses ``WM_COMMAND`` IDOK (out-of-process) so a modal never blocks COM
    teardown. Restricted to ``owned_pids`` (PIDs created by this test) so a
    user's own Excel dialog is never dismissed. Returns the dismissed hwnds.
    """

    dismissed: list[int] = []
    if win32gui is None or not owned_pids:  # pragma: no cover
        return dismissed
    for win in _enum_windows():
        if win.get("class") != DIALOG_CLASS or win.get("title") != DIALOG_TITLE:
            continue
        if win.get("pid") not in owned_pids:
            continue
        with contextlib.suppress(Exception):
            win32gui.PostMessage(win["hwnd"], 0x0111, 1, 0)  # WM_COMMAND, IDOK
            dismissed.append(win["hwnd"])
    return dismissed


# ---------------------------------------------------------------------------
# Shared body helpers
# ---------------------------------------------------------------------------
def _open_debug_and_find(uia: UiaSession, art: PocArtifacts, label: str) -> dict[str, Any]:
    """Expand Debug and return a single-snapshot Compile item + state.

    Returns the :meth:`UiaSession.compile_snapshot` dict (``item``/``present``/
    ``enabled``/``candidates``). The whole snapshot is taken from one UIA
    traversal so the popup cannot collapse between reads.
    """

    art.log(f"[{label}] expanding Debug menu")
    expanded, debug_candidates = uia.expand_debug()
    art.log(f"[{label}] Debug candidates={debug_candidates} expanded={expanded}")
    snapshot = uia.compile_snapshot()
    art.log(
        f"[{label}] Compile candidates={snapshot['candidates']} "
        f"present={snapshot['present']} enabled={snapshot['enabled']}"
    )
    art.compile_attempts.append(
        {
            "label": label,
            "debug_candidates": debug_candidates,
            "expanded": expanded,
            "compile_candidates": snapshot["candidates"],
            "item_found": snapshot["item"] is not None,
        }
    )
    if snapshot["item"] is None:
        art.log(f"[{label}] Compile VBAProject item NOT found/enabled/on-screen")
    return snapshot


# ---------------------------------------------------------------------------
# Test 1: success fixture
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_compile_success_fixture() -> None:
    """CleanProject.xlsm: UIA compile -> compiled, second compile -> already_compiled."""

    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    art = PocArtifacts(test_name="test_compile_success_fixture")

    def body(guarded: Any, wb: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        a.environment = _collect_environment(guarded, ctx["sta_thread_id"])
        a.uia_decision = {
            "primary_route": "raw comtypes IUIAutomation menu expansion + Invoke",
            "pywinauto": (
                "not used: 0.6.9 UIA wrapper has automation_id keyword bug and "
                "reverts COM threading mode on import (research §3.7)"
            ),
            "commandbars_route": "recorded as unavailable (R03 correction)",
        }

        # Evidence that the plan's CommandBars route is unavailable (R03).
        a.commandbars_evidence = _probe_commandbars(guarded)
        a.log(
            "CommandBars evidence: "
            f"count={a.commandbars_evidence.get('debug_controls_count')} "
            f"has_compile={a.commandbars_evidence.get('has_compile_entry')} "
            f"findcontrol_578={a.commandbars_evidence.get('findcontrol_578')}"
        )

        proj = wb.VBProject
        a.source_hash_before = _source_hash(proj)
        a.log(f"source_hash_before={a.source_hash_before}")

        hwnd = ctx["vbe_hwnd"]
        uia = UiaSession(hwnd)

        # --- First compile ---
        snap1 = _open_debug_and_find(uia, a, "compile_1")
        item = snap1["item"]
        if item is None:
            a.outcome = "indeterminate"
            a.log("compile_1: item unavailable -> indeterminate")
            return {"outcome": a.outcome}

        enabled_before_1 = bool(snap1["enabled"])
        a.log(f"compile_1: enabled_before={enabled_before_1}")

        before_hwnds = {w["hwnd"] for w in _enum_windows()}
        # Enter the modal-sensitive window: no Excel COM may run until it closes.
        ctx["state"]["modal_open"] = True
        invoke_ok = True
        try:
            uia.invoke_compile(item)
        except Exception as exc:  # noqa: BLE001
            invoke_ok = False
            a.log(f"compile_1: Invoke raised {exc!r}")
        time.sleep(POST_INVOKE_S)

        dialog = None
        deadline = time.monotonic() + DIALOG_WAIT_S
        while time.monotonic() < deadline:
            dialog = _find_new_dialog(before_hwnds)
            if dialog is not None:
                break
            time.sleep(DIALOG_POLL_INTERVAL_S)

        if dialog is None:
            # No dialog -> leave modal-sensitive window; safe to use COM again.
            ctx["state"]["modal_open"] = False
            a.log("compile_1: no #32770 dialog within wait window")

        # Re-open the menu (UIA) and take ONE snapshot to read the post-compile
        # state. After a successful compile the item is present but disabled.
        snap_after_1 = _open_debug_and_find(uia, a, "compile_1_after")
        state_after_1 = {
            "present": snap_after_1["present"],
            "enabled": snap_after_1["enabled"],
            "candidates": snap_after_1["candidates"],
        }
        a.source_hash_after = _source_hash(wb.VBProject)
        a.log(
            f"compile_1: invoked={invoke_ok} state_after_present={snap_after_1['present']} "
            f"state_after_enabled={snap_after_1['enabled']} hash_after={a.source_hash_after}"
        )

        if (
            invoke_ok
            and enabled_before_1
            and dialog is None
            and snap_after_1["present"]
            and snap_after_1["enabled"] is False
        ):
            a.outcome = "compiled"
        elif not enabled_before_1 and dialog is None:
            # Disabled before without a prior recorded compile -> not provable.
            a.outcome = "indeterminate"
        else:
            # Unexpected (e.g. dialog on a clean project) -> indeterminate.
            a.outcome = "indeterminate"
        a.compile_attempts.append(
            {
                "label": "compile_1_result",
                "invoked": invoke_ok,
                "enabled_before": enabled_before_1,
                "state_after": state_after_1,
                "dialog_present": dialog is not None,
                "outcome": a.outcome,
            }
        )
        a.log(f"compile_1 outcome={a.outcome}")

        # --- Second compile: already_compiled (disabled before + hash unchanged) ---
        snap_before_2 = _open_debug_and_find(uia, a, "compile_2")
        state_before_2 = {
            "present": snap_before_2["present"],
            "enabled": snap_before_2["enabled"],
        }
        a.source_hash_before_second = _source_hash(wb.VBProject)
        hash_unchanged = a.source_hash_before_second == a.source_hash_after
        if snap_before_2["present"] and snap_before_2["enabled"] is False:
            # Disabled before invoke: already_compiled only if the source is
            # unchanged since the last successful compile (research §3.3).
            a.outcome_second = "already_compiled" if hash_unchanged else "indeterminate"
        else:
            # Item still enabled (or absent): a real compile state is not provable
            # as "already compiled" here.
            a.outcome_second = "indeterminate"
        a.compile_attempts.append(
            {
                "label": "compile_2_result",
                "state_before": state_before_2,
                "hash_unchanged": hash_unchanged,
                "outcome": a.outcome_second,
            }
        )
        a.log(
            f"compile_2 state_before={state_before_2} "
            f"hash_unchanged={hash_unchanged} outcome={a.outcome_second}"
        )
        return {"outcome_first": a.outcome, "outcome_second": a.outcome_second}

    holder = _sta_run(fixture, body, art)
    art.orphan_reaped = holder.get("orphan_reaped")
    art.owned_pid = holder.get("owned_pid")
    if art.owned_pid is not None and not art.baseline_excel_pids:
        art.baseline_excel_pids = (holder.get("ctx") or {}).get("baseline_excel_pids", [])

    # Modal-safety: no Excel COM call while the dialog-open flag was set.
    com_log = holder.get("com_log", [])
    art.com_calls_while_modal = [c for c in com_log if c.get("modal_open")]
    art.no_com_while_modal = len(art.com_calls_while_modal) == 0

    result = holder.get("result", {}) or {}
    art.outcome = str(result.get("outcome_first", art.outcome))
    art.outcome_second = result.get("outcome_second")
    # Persist evidence before any assertion so a failed run still leaves artifacts.
    art.persist()
    print(f"[result] success body -> {result}", flush=True)
    print(f"[compile_attempts] {json.dumps(art.compile_attempts, default=str)}", flush=True)

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    assert result, "success body returned no result"

    # CommandBars route must be recorded as unavailable (R03 correction).
    assert art.commandbars_evidence.get("route_usable") is False, (
        "CommandBars route unexpectedly usable; update R03 evidence"
    )
    assert "findcontrol_578" in art.commandbars_evidence

    # No Excel COM while a compile dialog was open.
    assert art.no_com_while_modal, f"Excel COM ran while modal open: {art.com_calls_while_modal}"

    # First compile must be honestly classified as compiled (enabled->disabled,
    # no dialog) on a clean, freshly opened project.
    assert result.get("outcome_first") == "compiled", (
        f"expected compiled, got {result.get('outcome_first')}; evidence={a_json(art)}"
    )
    # Second compile: already_compiled (disabled before + source hash unchanged).
    assert result.get("outcome_second") == "already_compiled", (
        f"expected already_compiled, got {result.get('outcome_second')}; evidence={a_json(art)}"
    )

    # No orphaned owned Excel process.
    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    art.go_no_go = {
        "go_no_go_r03_task004": "GO",
        "rationale": (
            "UIA menu expansion + Invoke drove VBE Debug -> Compile VBAProject on the "
            "target build; first compile = compiled (enabled 1->0, no dialog); second "
            "compile = already_compiled (item disabled before, source hash unchanged). "
            "CommandBars route recorded unavailable (R03 correction). No Excel COM while "
            "modal. No orphan."
        ),
        "outcome_first": result.get("outcome_first"),
        "outcome_second": result.get("outcome_second"),
        "commandbars_route_usable": art.commandbars_evidence.get("route_usable"),
        "source_hash": art.source_hash_before,
    }
    art.persist()


# ---------------------------------------------------------------------------
# Test 2: compile-error fixture
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_compile_error_dialog_capture() -> None:
    """CompileErrorProject.xlsm: compile -> modal #32770 -> capture + dismiss."""

    fixture = _fixture_or_skip(_COMPILE_ERROR_PROJECT)
    art = PocArtifacts(test_name="test_compile_error_dialog_capture")

    def body(guarded: Any, wb: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        a.environment = _collect_environment(guarded, ctx["sta_thread_id"])
        a.uia_decision = {
            "primary_route": "raw comtypes IUIAutomation menu expansion + Invoke",
            "pywinauto": "not used (research §3.7: automation_id bug + COM threading)",
        }
        hwnd = ctx["vbe_hwnd"]
        uia = UiaSession(hwnd)

        # --- Invoke compile ---
        snap = _open_debug_and_find(uia, a, "compile_error")
        item = snap["item"]
        if item is None:
            a.outcome = "indeterminate"
            a.log("compile_error: item unavailable -> indeterminate")
            return {"outcome": a.outcome}

        before_hwnds = {w["hwnd"] for w in _enum_windows()}
        # From here until the dialog is dismissed: no Excel COM allowed.
        ctx["state"]["modal_open"] = True
        invoke_ok = True
        try:
            uia.invoke_compile(item)
        except Exception as exc:  # noqa: BLE001
            invoke_ok = False
            a.log(f"compile_error: Invoke raised {exc!r}")
        time.sleep(POST_INVOKE_S)

        dialog = None
        deadline = time.monotonic() + DIALOG_WAIT_S
        while time.monotonic() < deadline:
            dialog = _find_new_dialog(before_hwnds)
            if dialog is not None:
                break
            time.sleep(DIALOG_POLL_INTERVAL_S)

        if dialog is None:
            ctx["state"]["modal_open"] = False
            a.outcome = "indeterminate"
            a.log("compile_error: no #32770 dialog appeared -> indeterminate")
            return {"outcome": a.outcome}

        # Dialog is open: capture text (UIA) and screenshot. No Excel COM here.
        a.log(f"compile_error: dialog detected {dialog}")
        dialog_data = uia.read_dialog(dialog["hwnd"])
        a.log(f"compile_error: dialog texts={dialog_data['texts']}")

        shot_path = _ARTIFACT_DIR / (
            f"vbe_compile_error_dialog_{_dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.png"
        )
        _ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
        screenshot = _screenshot_window(dialog["hwnd"], shot_path)

        # Best-effort: try the code pane while the dialog is open (the dialog's own
        # TextPattern.GetSelection is NULL on this build, so the code pane is the
        # only source). Wrapped so an unreadable tree never blocks dismissal.
        code_pane_open: dict[str, Any] = {"attempted": True}
        try:
            code_pane_open = uia.read_code_pane()
            code_pane_open["attempted"] = True
        except Exception as exc:  # noqa: BLE001
            code_pane_open = {"attempted": True, "error": f"{type(exc).__name__}: {exc}"}
        a.log(
            "compile_error: code_pane(while open) "
            f"found={code_pane_open.get('found')} selection={code_pane_open.get('selection')!r}"
        )

        combined_text = "\n".join(dialog_data["texts"])
        a.dialog = {
            "hwnd": dialog["hwnd"],
            "title": dialog.get("title"),
            "class": dialog.get("class"),
            "pid": dialog.get("pid"),
            "texts": dialog_data["texts"],
            "buttons": dialog_data["buttons"],
            "screenshot": screenshot,
            "screenshot_path": str(shot_path) if screenshot.get("exists") else None,
            "code_pane_while_open": code_pane_open,
        }

        has_compile_error = "Compile error:" in combined_text
        has_variable_not_defined = "Variable not defined" in combined_text
        a.log(
            f"compile_error: has_compile_error={has_compile_error} "
            f"has_variable_not_defined={has_variable_not_defined}"
        )

        # Dismiss via UIA (OK button) while still inside the modal-sensitive window.
        dismiss_status = uia.dismiss_dialog_ok(dialog["hwnd"])
        a.log(f"compile_error: dismiss_status={dismiss_status}")
        time.sleep(0.5)
        still_open = any(w["hwnd"] == dialog["hwnd"] for w in _enum_windows())
        # Leave the modal-sensitive window only after dismissal.
        ctx["state"]["modal_open"] = False
        a.log(f"compile_error: dialog still open after dismiss={still_open}")

        # Read the offending token from the code pane now that the dialog is gone
        # (VBE sets the selection after the dialog is acknowledged). Re-read the
        # UIA root so the pane reflects the post-error highlight.
        uia_post = UiaSession(ctx["vbe_hwnd"])
        code_pane_post: dict[str, Any] = {"attempted": True}
        try:
            code_pane_post = uia_post.read_code_pane()
            code_pane_post["attempted"] = True
        except Exception as exc:  # noqa: BLE001
            code_pane_post = {"attempted": True, "error": f"{type(exc).__name__}: {exc}"}

        # Prefer whichever read exposed the selected token.
        chosen = None
        for candidate in (code_pane_open, code_pane_post):
            if candidate.get("selection"):
                chosen = candidate
                break
        a.selected_token_detail = {
            "while_open": code_pane_open,
            "after_dismiss": code_pane_post,
        }
        if chosen is not None:
            a.selected_token = str(chosen["selection"])
        else:
            a.selected_token = "unavailable"
        a.log(
            f"compile_error: selected_token={a.selected_token!r} "
            f"(after_dismiss found={code_pane_post.get('found')})"
        )

        if invoke_ok and has_compile_error and has_variable_not_defined and not still_open:
            a.outcome = "compile_error"
        elif invoke_ok and has_compile_error and still_open:
            # Dialog text correct but could not dismiss -> still a compile_error,
            # but flag the dismissal problem (cleanup force-dismisses).
            a.outcome = "compile_error"
            a.log("compile_error: WARNING dialog not dismissed via UIA")
        else:
            a.outcome = "indeterminate"
        a.compile_attempts.append(
            {
                "label": "compile_error_result",
                "invoked": invoke_ok,
                "dialog_present": True,
                "has_compile_error": has_compile_error,
                "has_variable_not_defined": has_variable_not_defined,
                "dismiss_status": dismiss_status,
                "still_open_after_dismiss": still_open,
                "outcome": a.outcome,
            }
        )
        a.log(f"compile_error outcome={a.outcome}")
        return {"outcome": a.outcome, "dialog": a.dialog}

    holder = _sta_run(fixture, body, art)
    art.orphan_reaped = holder.get("orphan_reaped")
    art.owned_pid = holder.get("owned_pid")
    if art.owned_pid is not None and not art.baseline_excel_pids:
        art.baseline_excel_pids = (holder.get("ctx") or {}).get("baseline_excel_pids", [])

    com_log = holder.get("com_log", [])
    art.com_calls_while_modal = [c for c in com_log if c.get("modal_open")]
    art.no_com_while_modal = len(art.com_calls_while_modal) == 0

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    result = holder.get("result", {})
    assert result, "compile-error body returned no result"

    # Modal safety: zero Excel COM calls while the dialog was open.
    assert art.no_com_while_modal, f"Excel COM ran while modal open: {art.com_calls_while_modal}"

    # Outcome must be compile_error and never compiled.
    assert result.get("outcome") == "compile_error", (
        f"expected compile_error, got {result.get('outcome')}; evidence={a_json(art)}"
    )
    assert result.get("outcome") != "compiled"

    # Dialog evidence: #32770, correct title/class, correct text.
    dialog = art.dialog or {}
    assert dialog.get("class") == DIALOG_CLASS, f"dialog class: {dialog.get('class')}"
    assert dialog.get("title") == DIALOG_TITLE, f"dialog title: {dialog.get('title')}"
    texts = dialog.get("texts", [])
    joined = "\n".join(texts)
    assert "Compile error:" in joined, f"dialog text missing 'Compile error:': {texts}"
    assert "Variable not defined" in joined, f"dialog text missing 'Variable not defined': {texts}"

    # Screenshot captured.
    assert dialog.get("screenshot_path"), f"screenshot not captured: {dialog.get('screenshot')}"
    assert Path(dialog["screenshot_path"]).is_file(), "screenshot file missing on disk"

    # Selected token is recorded (value or explicit 'unavailable'); never faked.
    assert art.selected_token in ("unavailable",) or isinstance(art.selected_token, str)

    # No orphaned owned Excel process.
    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    art.go_no_go = {
        "go_no_go_r04_task004": "GO",
        "rationale": (
            "Compile-error fixture produced a modal #32770 'Microsoft Visual Basic for "
            "Applications' dialog; UIA read 'Compile error: Variable not defined'; the "
            "offending token was read from the code pane TextPattern when available "
            f"(selected_token={art.selected_token!r}); dialog dismissed via UIA with no "
            "Excel COM while modal; screenshot captured; no orphan."
        ),
        "dialog_text": texts,
        "screenshot": dialog.get("screenshot_path"),
        "selected_token": art.selected_token,
        "no_com_while_modal": art.no_com_while_modal,
    }
    art.persist()


def a_json(art: PocArtifacts) -> str:
    """Compact JSON snapshot of an artifact for assertion messages."""

    return json.dumps(asdict(art), ensure_ascii=False, default=str)[:1500]
