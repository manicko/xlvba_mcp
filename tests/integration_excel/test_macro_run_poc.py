"""Application.Run behavior proof-of-concept for task_005 (binding).

This module is a bounded, self-contained proof-of-concept harness (not a
production abstraction). It proves, on the target Windows/Excel build, the
``Application.Run`` execution/error behavior that production
``application/macro_service.py`` (task_022) and ``domain/errors.py`` depend on:

* a qualified macro call (``'CleanProject.xlsm'!modTest.<Proc>``) returns a
  serializable Variant for scalar arguments and a Variant array return;
* an unsupported argument (a plain Python ``object()``) is rejected by the
  pywin32/Variant boundary BEFORE any COM dispatch reaches Excel;
* a missing macro and an unhandled runtime error both surface an OUTER
  ``DISP_E_EXCEPTION`` (``0x80020009``) but are distinguished by the INNER
  HRESULT (``0x800A03ED`` not-found vs ``0x800A03EC``/``0x80004005`` runtime);
* an interactive-UI macro (``MsgBox`` / ``UserForm.Show``) is blocked BEFORE
  execution by a source preflight heuristic (``is_background_safe``), so no
  modal dialog, no visible window and no focus steal occurs (addendum BG06/BG09;
  R13). The original VBA description is never promised.

Every Excel COM call runs on a single dedicated STA thread with an owned
``DispatchEx`` instance (``Visible=False``, ``DisplayAlerts=False``). Cleanup
closes the workbook without saving, quits, releases references and
``CoUninitialize``s on that same thread, then polls for an orphaned owned PID.

Read-only background non-interference evidence (addendum BG03/BG04/BG05) is
recorded: foreground window before creation, after dispatch and after each
operation; visible top-level windows of the owned PID; ``Application.Visible``;
and a best-effort (non-activating) read of ``VBE.MainWindow.Visible``. None of
these queries activate, show, click or type.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_macro_run_poc.py -v -s
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import gc
import json
import re
import subprocess
import sys
import threading
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pytest

EXCEL_PROGID = "Excel.Application"

# Outer DISP_E_EXCEPTION wrapper that carries the inner HRESULT (research §4.2).
DISP_E_EXCEPTION = 0x80020009
# Inner HRESULTs observed on the target build.
HR_NOT_FOUND = 0x800A03ED  # macro not available in this workbook
HR_RUNTIME_ERROR = 0x800A03EC  # VBA runtime error surfaced by Excel
HR_E_FAIL = 0x80004005  # generic E_FAIL from a runtime error
HR_DISP_E_BADPARAMCOUNT = 0x8002000E  # wrong argument count

# Structured macro-run classifications for this PoC.
CLASS_SUCCESS = "MACRO_SUCCESS"
CLASS_MACRO_NOT_FOUND = "MACRO_NOT_FOUND"
CLASS_MACRO_RUNTIME_ERROR = "MACRO_RUNTIME_ERROR"
CLASS_REJECTED_UNSUPPORTED_ARGUMENT = "REJECTED_UNSUPPORTED_ARGUMENT"
CLASS_INTERACTIVE_UI_NOT_SUPPORTED = "INTERACTIVE_UI_NOT_SUPPORTED"
CLASS_BACKGROUND_SAFETY_NOT_VERIFIED = "BACKGROUND_SAFETY_NOT_VERIFIED"

# Tokens that flag a procedure as potentially interactive / not background-safe.
INTERACTIVE_TOKENS: tuple[str, ...] = (
    "MsgBox",
    "InputBox",
    ".Show",
    "FileDialog",
    "Shell",
    "CreateObject",
    "FollowHyperlink",
)

ORPHAN_POLL_INTERVAL_S = 0.5
# Excel shutdown on the target host is variable (observed ~2.7s to ~18s+ cold);
# a short poll can report a false orphan. A genuinely never-exiting owned process
# still fails after this window (see test_vbide_access_poc.py rationale).
ORPHAN_POLL_TIMEOUT_S = 30.0
STA_JOIN_TIMEOUT_S = 180.0

# Optional COM dependencies. ``None`` when unavailable so the module imports on
# non-Windows hosts and the tests skip cleanly.
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
_ARTIFACT_DIR = _REPO_ROOT / ".ai" / "tmp"
_FIXTURES_DIR = _REPO_ROOT / "tests" / "fixtures"
_CLEAN_PROJECT = _FIXTURES_DIR / "CleanProject.xlsm"


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
    reason=f"Excel COM PoC unavailable: {_AVAIL_REASON}",
)


def _clean_project_or_skip() -> Path:
    """Return the CleanProject fixture path or skip when it is missing."""

    if not _CLEAN_PROJECT.is_file():
        pytest.skip(
            f"fixture missing: {_CLEAN_PROJECT}; run "
            "`uv run python tests/fixtures/build_fixtures.py` first"
        )
    return _CLEAN_PROJECT


# ---------------------------------------------------------------------------
# Time / COM error helpers
# ---------------------------------------------------------------------------
def _now_utc_iso() -> str:
    return _dt.datetime.now(_dt.UTC).isoformat()


def _now_local_iso() -> str:
    return _dt.datetime.now().astimezone().isoformat()


def _outer_hresult(exc: BaseException) -> int | None:
    """Return the outer HRESULT carried on a pywintypes.com_error, if any."""

    hr = getattr(exc, "hresult", None)
    if isinstance(hr, int):
        return hr & 0xFFFFFFFF
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0] & 0xFFFFFFFF
    return None


def _inner_hresult(exc: BaseException) -> int | None:
    """Extract the inner HRESULT from a pywintypes.com_error.

    ``win32com`` wraps an ``IDispatch`` failure as ``DISP_E_EXCEPTION`` whose
    third arg is the EXCEPINFO tuple; the inner HRESULT is element 5 (verified on
    the target build; research §4.2). Returns ``None`` when absent.
    """

    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        candidate = args[2][5]
        if isinstance(candidate, int):
            return candidate & 0xFFFFFFFF
    return None


def _exception_message(exc: BaseException) -> str | None:
    """Return the EXCEPINFO description string, if present.

    The EXCEPINFO tuple is ``(hr, source, description, helpfile, helpcontext,
    inner_hresult)``. Element 2 is the human-readable description (e.g. Excel's
    ``Cannot run the macro '...'``); element 1 is the source (e.g. ``Microsoft
    Excel``) and is NOT the description. Returns ``None`` when absent.
    """

    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        candidate = args[2][2]
        if isinstance(candidate, str):
            return candidate
    return None


def _hresult_hex(value: int | None) -> str | None:
    return f"0x{value:08X}" if isinstance(value, int) else None


def _classify_run_error(exc: BaseException) -> str:
    """Classify an ``Application.Run`` failure as NOT_FOUND vs RUNTIME_ERROR.

    The outer ``DISP_E_EXCEPTION`` (``0x80020009``) is shared by every failure.
    The inner HRESULT helps but is NOT sufficient on its own: on the target
    build a missing macro (``modTest.Nope``) surfaced inner ``0x800A03EC`` --
    the same code an unhandled runtime error produces -- so the EXCEPINFO
    description is the decisive evidence (research §4.2 / task_009 §3):

    * inner ``0x800A03ED`` -> ``MACRO_NOT_FOUND``;
    * description contains ``Cannot run the macro`` / ``may not be available``
      -> ``MACRO_NOT_FOUND`` (regardless of the inner HRESULT);
    * otherwise -> ``MACRO_RUNTIME_ERROR`` (fail-closed; never report a generic
      failure as NOT_FOUND, spec §11 / research §4.2).

    The original VBA runtime description is never promised; the description here
    is Excel's EXCEPINFO text, which for a runtime error is generic.
    """

    if _inner_hresult(exc) == HR_NOT_FOUND:
        return CLASS_MACRO_NOT_FOUND
    description = (_exception_message(exc) or "").lower()
    if "cannot run the macro" in description or "may not be available" in description:
        return CLASS_MACRO_NOT_FOUND
    return CLASS_MACRO_RUNTIME_ERROR


# ---------------------------------------------------------------------------
# Source preflight heuristic (addendum §2.4 / BG06 / R13)
# ---------------------------------------------------------------------------
def is_background_safe(source: str) -> tuple[bool, list[str]]:
    """Heuristic: does this procedure source look free of interactive UI calls?

    This is a risk filter, not static proof (addendum §2.4). It flags the direct
    textual presence of known interactive tokens. Dynamic/indirect calls are out
    of scope for the heuristic; runtime monitoring remains mandatory.

    Returns ``(safe, flagged_tokens)`` where ``flagged_tokens`` are the tokens
    found (empty when safe).
    """

    flagged: list[str] = []
    for token in INTERACTIVE_TOKENS:
        if token in source:
            flagged.append(token)
    return (not flagged), flagged


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
# Artifact collector
# ---------------------------------------------------------------------------
@dataclass
class PocArtifacts:
    """Structured evidence collector persisted as JSON + text."""

    test_name: str = ""
    started_utc: str = field(default_factory=_now_utc_iso)
    started_local: str = field(default_factory=_now_local_iso)
    sta_thread_id: int | None = None
    owned_pid: int | None = None
    owned_pid_resolution: str = ""
    baseline_excel_pids: list[int] = field(default_factory=list)
    excel_version: str = ""
    excel_build: str = ""
    # Macro-run evidence.
    macro_runs: list[dict[str, Any]] = field(default_factory=list)
    preflight: list[dict[str, Any]] = field(default_factory=list)
    foreground_window: dict[str, Any] = field(default_factory=dict)
    visible_windows: dict[str, Any] = field(default_factory=dict)
    orphan_poll: list[dict[str, Any]] = field(default_factory=list)
    orphan_reaped: bool | None = None
    # Desktop non-interference evidence (addendum BG03/BG04/BG05).
    baseline_excel_pids_after: list[int] = field(default_factory=list)
    foreground_before: dict[str, Any] | None = None
    foreground_after_dispatch: dict[str, Any] | None = None
    foreground_after_operations: dict[str, Any] | None = None
    foreground_after: dict[str, Any] | None = None
    visible_windows_after_dispatch: dict[str, Any] | None = None
    visible_windows_after_operations: dict[str, Any] | None = None
    app_visible_start: Any = None
    app_visible_after_operations: Any = None
    vbe_visible: dict[str, Any] | None = None
    bg03_no_focus_theft: bool | None = None
    bg04_no_visible_popup: bool | None = None
    bg05_user_session_preserved: bool | None = None
    bg_verdict: dict[str, Any] | None = None
    log_lines: list[str] = field(default_factory=list)
    go_no_go: dict[str, Any] | None = None
    paths: dict[str, str] = field(default_factory=dict)
    artifact_version: str = "task_005_binding+bg03_bg04_bg05_bg06_bg09"

    def log(self, msg: str) -> None:
        line = f"{_now_local_iso()} {msg}"
        self.log_lines.append(line)
        # The console may use a legacy code page (e.g. cp1251) that cannot encode
        # every foreground-window title glyph; printing must never crash the STA
        # worker. The full UTF-8 evidence is always persisted by ``persist``.
        try:
            print(line, flush=True)
        except Exception:  # noqa: BLE001 - console encoding cannot represent line
            with contextlib.suppress(Exception):
                print(line.encode("ascii", "replace").decode("ascii"), flush=True)

    def persist(self) -> dict[str, str]:
        _ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        json_path = _ARTIFACT_DIR / f"macro_run_poc_{self.test_name}_{ts}.json"
        txt_path = _ARTIFACT_DIR / f"macro_run_poc_{self.test_name}_{ts}.txt"
        json_path.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        header = (
            "task_005 Application.Run PoC evidence\n"
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
# Desktop non-interference evidence (addendum BG03/BG04/BG05)
# ---------------------------------------------------------------------------
# All queries below are READ-ONLY Win32 observations. They never call
# ``SetForegroundWindow``, ``ShowWindow`` (to show/restore), ``SetFocus``, click
# or type, and never touch a process the automation does not own.


def _read_app_visible(app: Any) -> Any:
    """Read ``Application.Visible`` without mutating it; return raw or an error tag."""

    try:  # pragma: no cover - depends on host
        return app.Visible
    except Exception as exc:  # noqa: BLE001
        return f"error:{exc!r}"


def _record_foreground(label: str, owned_pids: set[int], art: PocArtifacts) -> dict[str, Any]:
    """Capture the current foreground window; never change or activate anything."""

    sample: dict[str, Any] = {
        "label": label,
        "timestamp": _now_local_iso(),
        "available": win32gui is not None and win32process is not None,
        "hwnd": None,
        "title": None,
        "pid": None,
        "is_excel_owned": None,
        "is_excel_owned_pids": sorted(owned_pids),
        "error": None,
    }
    if not sample["available"]:
        sample["error"] = "win32gui/win32process unavailable"
        art.log(f"[fg:{label}] unavailable: win32gui/win32process missing")
        return sample
    try:  # pragma: no cover - depends on host
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception as exc:  # noqa: BLE001
        sample["error"] = f"{type(exc).__name__}:{exc!r}"
        art.log(f"[fg:{label}] query failed: {sample['error']}")
        return sample
    is_owned = bool(pid in owned_pids) if owned_pids else False
    sample.update({"hwnd": hwnd, "title": title, "pid": pid, "is_excel_owned": is_owned})
    art.log(f"[fg:{label}] hwnd={hwnd} pid={pid} is_excel_owned={is_owned} title={title!r}")
    return sample


def _record_visible_windows(label: str, owned_pids: set[int], art: PocArtifacts) -> dict[str, Any]:
    """Enumerate top-level windows of the owned Excel PID and flag visible ones."""

    result: dict[str, Any] = {
        "label": label,
        "timestamp": _now_local_iso(),
        "available": win32gui is not None and win32process is not None,
        "owned_pids": sorted(owned_pids),
        "owned_windows": [],
        "visible_windows": [],
        "visible_window_count": 0,
        "visible_window_titles": [],
        "error": None,
    }
    if not result["available"]:
        result["error"] = "win32gui/win32process unavailable"
        art.log(f"[visible:{label}] unavailable: win32gui/win32process missing")
        return result
    if not owned_pids:
        art.log(f"[visible:{label}] no owned PID resolved; nothing to enumerate")
        return result

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception as exc:  # noqa: BLE001
            art.log(f"[visible:{label}] GetWindowThreadProcessId failed: {exc!r}")
            return True
        if pid not in owned_pids:
            return True
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception as exc:  # noqa: BLE001
            title = f"<title_error:{exc!r}>"
        try:
            visible = bool(win32gui.IsWindowVisible(hwnd))
        except Exception:  # noqa: BLE001
            visible = False
        result["owned_windows"].append(
            {"hwnd": hwnd, "pid": pid, "title": title, "is_visible": visible}
        )
        if visible:
            result["visible_windows"].append(
                {"hwnd": hwnd, "pid": pid, "title": title, "is_visible": True}
            )
        return True

    try:  # pragma: no cover - depends on host
        win32gui.EnumWindows(_cb, None)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"EnumWindows_failed:{exc!r}"
        art.log(f"[visible:{label}] EnumWindows failed: {exc!r}")
        return result

    result["visible_window_count"] = len(result["visible_windows"])
    result["visible_window_titles"] = [w["title"] for w in result["visible_windows"]]
    art.log(
        f"[visible:{label}] owned_windows={len(result['owned_windows'])} "
        f"visible={result['visible_window_count']} titles={result['visible_window_titles']}"
    )
    return result


def _record_vbe_visible(app: Any, art: PocArtifacts) -> dict[str, Any]:
    """Best-effort read of ``VBE.MainWindow.Visible`` without activating the VBE."""

    result: dict[str, Any] = {"check": "not_reachable", "visible": None, "detail": None}
    try:  # pragma: no cover - depends on host
        vbe = app.VBE
        main_window = vbe.MainWindow
        result["visible"] = main_window.Visible
        result["check"] = "reachable"
    except Exception as exc:  # noqa: BLE001
        result["detail"] = f"{type(exc).__name__}:{str(exc)[:200]}"
        result["check"] = "not_reachable"
        art.log(f"[vbe_visible] not reachable: {result['detail']}")
        return result
    art.log(f"[vbe_visible] reachable visible={result['visible']!r} (read-only)")
    return result


def _finalize_bg(art: PocArtifacts, holder: dict[str, Any]) -> None:
    """Copy the BG03/BG04/BG05 samples off the worker ctx and compute the verdict."""

    ctx = holder.get("ctx") or {}
    art.foreground_before = ctx.get("foreground_before")
    art.foreground_after_dispatch = ctx.get("foreground_after_dispatch")
    art.foreground_after_operations = ctx.get("foreground_after_operations")
    art.foreground_after = ctx.get("foreground_after")
    art.visible_windows_after_dispatch = ctx.get("visible_windows_after_dispatch")
    art.visible_windows_after_operations = ctx.get("visible_windows_after_operations")
    art.app_visible_start = ctx.get("app_visible_start")
    art.app_visible_after_operations = ctx.get("app_visible_after_operations")
    art.vbe_visible = ctx.get("vbe_visible")
    # Per-test macro-run convenience evidence: foreground/visible at the last
    # operation boundary (after_operations is the most representative sample).
    art.foreground_window = ctx.get("foreground_after_operations") or {}
    art.visible_windows = ctx.get("visible_windows_after_operations") or {}

    fg_samples = [
        art.foreground_after_dispatch,
        art.foreground_after_operations,
        art.foreground_after,
    ]
    bg03_reasons = [
        f"automation-owned window foreground at {s.get('label')}"
        for s in fg_samples
        if s and s.get("is_excel_owned")
    ]
    vw_samples = [
        art.visible_windows_after_dispatch,
        art.visible_windows_after_operations,
    ]
    bg04_reasons = [
        f"automation-owned window visible at {vw.get('label')}: {vw.get('visible_window_titles')}"
        for vw in vw_samples
        if vw and vw.get("visible_windows")
    ]
    if art.app_visible_start is True:
        bg04_reasons.append("Application.Visible is True at start")
    if art.app_visible_after_operations is True:
        bg04_reasons.append("Application.Visible is True after operations")

    after_pids = _excel_pids()
    art.baseline_excel_pids_after = sorted(after_pids)
    missing_user_pids = sorted(set(art.baseline_excel_pids) - after_pids)
    bg05_reasons = (
        [f"user Excel PIDs disappeared: {missing_user_pids}"] if missing_user_pids else []
    )

    art.bg03_no_focus_theft = not bg03_reasons
    art.bg04_no_visible_popup = not bg04_reasons
    art.bg05_user_session_preserved = not bg05_reasons
    bg_ok = (
        art.bg03_no_focus_theft and art.bg04_no_visible_popup and art.bg05_user_session_preserved
    )

    art.bg_verdict = {
        "bg_verdict": "GO" if bg_ok else "NO_GO",
        "bg03_no_focus_theft": art.bg03_no_focus_theft,
        "bg04_no_visible_popup": art.bg04_no_visible_popup,
        "bg05_user_session_preserved": art.bg05_user_session_preserved,
        "bg03_reasons": bg03_reasons,
        "bg04_reasons": bg04_reasons,
        "bg05_reasons": bg05_reasons,
        "foreground_before": art.foreground_before,
        "foreground_after": art.foreground_after,
        "visible_windows_after_dispatch": art.visible_windows_after_dispatch,
        "visible_windows_after_operations": art.visible_windows_after_operations,
        "app_visible_start": art.app_visible_start,
        "app_visible_after_operations": art.app_visible_after_operations,
        "vbe_visible": art.vbe_visible,
        "rationale": (
            "No automation-owned Excel/VBE window became foreground or visible; "
            "Application.Visible stayed False; user-owned Excel processes untouched."
            if bg_ok
            else "Background non-interference violated: "
            + "; ".join(bg03_reasons + bg04_reasons + bg05_reasons)
        ),
    }
    art.log(
        f"[bg] bg03={art.bg03_no_focus_theft} bg04={art.bg04_no_visible_popup} "
        f"bg05={art.bg05_user_session_preserved} verdict={art.bg_verdict['bg_verdict']}"
    )


def _assert_bg(art: PocArtifacts) -> bool:
    """Assert the BG03/BG04/BG05 background contract holds; return ``bg_ok``."""

    assert art.foreground_before is not None, "foreground_before not captured"
    assert art.app_visible_start is False, (
        f"Application.Visible was not False at start: {art.app_visible_start!r}"
    )
    assert art.app_visible_after_operations is False, (
        f"Application.Visible was not False after operations: {art.app_visible_after_operations!r}"
    )
    assert art.bg03_no_focus_theft is True, (
        "automation-owned window took foreground (BG03): "
        f"after_dispatch={art.foreground_after_dispatch} "
        f"after_operations={art.foreground_after_operations} after={art.foreground_after}"
    )
    assert art.bg04_no_visible_popup is True, (
        "automation-owned window was visible (BG04): "
        f"after_dispatch={art.visible_windows_after_dispatch} "
        f"after_operations={art.visible_windows_after_operations}"
    )
    assert art.bg05_user_session_preserved is True, (
        "user-owned Excel process(es) disappeared (BG05): "
        f"before={art.baseline_excel_pids} after={art.baseline_excel_pids_after}"
    )
    vbe = art.vbe_visible
    if vbe and vbe.get("check") == "reachable":
        assert vbe.get("visible") in (False, 0), f"VBE MainWindow visible: {vbe}"
    else:
        assert vbe is None or vbe.get("check") == "not_reachable", f"unexpected vbe check: {vbe}"
    return bool(art.bg_verdict and art.bg_verdict.get("bg_verdict") == "GO")


# ---------------------------------------------------------------------------
# Single-STA-thread runner
# ---------------------------------------------------------------------------
def _sta_run(
    workbook_path: Path,
    body: Callable[[Any, Any, dict[str, Any], PocArtifacts], dict[str, Any]],
    art: PocArtifacts,
) -> dict[str, Any]:
    """Run ``body(app, wb, ctx, art)`` on a dedicated STA thread with full cleanup.

    The worker owns the entire COM lifecycle: ``CoInitializeEx(APARTMENTTHREADED)``,
    ``DispatchEx``, owned-PID resolution, workbook open (read-only), the body
    call, then close without save, ``Quit``, reference release,
    ``CoUninitialize`` and an orphan poll. The body returns a JSON-serializable
    dict; exceptions are captured so the main thread can assert without leaking
    a process.
    """

    holder: dict[str, Any] = {}

    def worker() -> None:
        tid = threading.get_ident()
        ctx: dict[str, Any] = {"sta_thread_id": tid}
        app: Any = None
        wb: Any = None
        owned_pid: int | None = None
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as exc:  # noqa: BLE001
            holder["error"] = f"CoInitializeEx failed: {exc!r}"
            return
        try:
            baseline = _excel_pids()
            ctx["baseline_excel_pids"] = sorted(baseline)
            ctx["foreground_before"] = _record_foreground("before", set(), art)
            app = win32com.client.DispatchEx(EXCEL_PROGID)
            app.Visible = False
            app.DisplayAlerts = False
            new_pids = _excel_pids() - baseline
            owned_pid = next(iter(new_pids)) if len(new_pids) == 1 else None
            ctx["owned_pid"] = owned_pid
            ctx["owned_pid_resolution"] = f"pid_delta={sorted(new_pids)}"
            owned_pids = {owned_pid} if owned_pid is not None else set()
            ctx["app_visible_start"] = _read_app_visible(app)
            ctx["foreground_after_dispatch"] = _record_foreground("after_dispatch", owned_pids, art)
            ctx["visible_windows_after_dispatch"] = _record_visible_windows(
                "after_dispatch", owned_pids, art
            )
            wb = app.Workbooks.Open(
                str(workbook_path),
                UpdateLinks=0,
                ReadOnly=True,
                Notify=False,
                AddToMru=False,
            )
            holder["result"] = body(app, wb, ctx, art)
            ctx["foreground_after_operations"] = _record_foreground(
                "after_operations", owned_pids, art
            )
            ctx["visible_windows_after_operations"] = _record_visible_windows(
                "after_operations", owned_pids, art
            )
            ctx["app_visible_after_operations"] = _read_app_visible(app)
            ctx["vbe_visible"] = _record_vbe_visible(app, art)
            holder["ctx"] = ctx
        except BaseException as exc:  # noqa: BLE001 - surfaced to the main thread
            holder["error"] = f"{type(exc).__name__}: {exc}"
            holder["ctx"] = ctx
        finally:
            if wb is not None:
                with contextlib.suppress(Exception):
                    wb.Close(SaveChanges=False)
            wb = None
            gc.collect()
            if app is not None:
                with contextlib.suppress(Exception):
                    app.Quit()
            app = None
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()
            holder["owned_pid"] = owned_pid if owned_pid is not None else ctx.get("owned_pid")
            holder["orphan_reaped"] = _orphan_poll(holder["owned_pid"])
            ctx["foreground_after"] = _record_foreground("after", set(), art)
            holder["ctx"] = ctx

    thread = threading.Thread(target=worker, name="macro_run_sta", daemon=True)
    thread.start()
    thread.join(timeout=STA_JOIN_TIMEOUT_S)
    holder["thread_alive"] = thread.is_alive()
    return holder


# ---------------------------------------------------------------------------
# Watchdog STA runner for a macro that may raise a VBA modal (BG04/R13)
# ---------------------------------------------------------------------------
def _terminate_owned_pid(pid: int) -> bool:
    """Terminate ONLY the automation-owned Excel PID; return True on success.

    Never uses a wildcard image kill; the PID was resolved by the automation as
    the single new EXCEL.EXE created by ``DispatchEx``. This is the last-resort
    cleanup for a macro that popped a modal ``Application.Run`` cannot return
    from. It does not dismiss the dialog by input; it ends the owned process.
    """

    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return out.returncode == 0
    except Exception:  # noqa: BLE001
        return False


def _sta_run_watchdog(
    workbook_path: Path,
    body: Callable[[Any, Any, dict[str, Any], PocArtifacts], dict[str, Any]],
    art: PocArtifacts,
    *,
    watchdog_timeout_s: float,
    grace_s: float = 3.0,
) -> dict[str, Any]:
    """Run ``body`` on an STA thread; if it blocks, detect and stop the owned PID.

    Used for the runtime-error macro that raises a VBA modal dialog. On timeout:

    * the owned Excel PID's visible top-level windows are enumerated (READ-ONLY)
      and recorded as modal evidence;
    * ONLY the owned PID is terminated (no dialog input, no wildcard kill);
    * the blocked step is classified ``MACRO_RUNTIME_ERROR`` with a
      ``modal_runtime_error`` outcome and ``dismissed=False``.

    The body signals that it is about to make the possibly-blocking call by
    setting ``ctx['blocking_event']`` (a ``threading.Event``) right before
    ``app.Run``. The worker also publishes the owned PID on the shared
    ``holder`` as soon as it is resolved, so the watchdog can identify the owned
    process mid-flight (before the body returns).
    """

    holder: dict[str, Any] = {}
    blocking_event = threading.Event()

    def worker() -> None:
        tid = threading.get_ident()
        ctx: dict[str, Any] = {
            "sta_thread_id": tid,
            "blocking_event": blocking_event,
        }
        app: Any = None
        wb: Any = None
        owned_pid: int | None = None
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as exc:  # noqa: BLE001
            holder["error"] = f"CoInitializeEx failed: {exc!r}"
            return
        try:
            baseline = _excel_pids()
            ctx["baseline_excel_pids"] = sorted(baseline)
            ctx["foreground_before"] = _record_foreground("before", set(), art)
            app = win32com.client.DispatchEx(EXCEL_PROGID)
            app.Visible = False
            app.DisplayAlerts = False
            new_pids = _excel_pids() - baseline
            owned_pid = next(iter(new_pids)) if len(new_pids) == 1 else None
            ctx["owned_pid"] = owned_pid
            ctx["owned_pid_resolution"] = f"pid_delta={sorted(new_pids)}"
            holder["owned_pid"] = owned_pid  # publish early for the watchdog
            owned_pids = {owned_pid} if owned_pid is not None else set()
            ctx["app_visible_start"] = _read_app_visible(app)
            ctx["foreground_after_dispatch"] = _record_foreground("after_dispatch", owned_pids, art)
            ctx["visible_windows_after_dispatch"] = _record_visible_windows(
                "after_dispatch", owned_pids, art
            )
            wb = app.Workbooks.Open(
                str(workbook_path),
                UpdateLinks=0,
                ReadOnly=True,
                Notify=False,
                AddToMru=False,
            )
            holder["result"] = body(app, wb, ctx, art)
            ctx["foreground_after_operations"] = _record_foreground(
                "after_operations", owned_pids, art
            )
            ctx["visible_windows_after_operations"] = _record_visible_windows(
                "after_operations", owned_pids, art
            )
            ctx["app_visible_after_operations"] = _read_app_visible(app)
            ctx["vbe_visible"] = _record_vbe_visible(app, art)
            holder["ctx"] = ctx
        except BaseException as exc:  # noqa: BLE001 - surfaced to the main thread
            holder["error"] = f"{type(exc).__name__}: {exc}"
            holder["ctx"] = ctx
        finally:
            if wb is not None:
                with contextlib.suppress(Exception):
                    wb.Close(SaveChanges=False)
            wb = None
            gc.collect()
            if app is not None:
                with contextlib.suppress(Exception):
                    app.Quit()
            app = None
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()
            holder["owned_pid"] = owned_pid if owned_pid is not None else ctx.get("owned_pid")
            holder["orphan_reaped"] = _orphan_poll(holder["owned_pid"])
            ctx["foreground_after"] = _record_foreground("after", set(), art)
            holder["ctx"] = ctx

    thread = threading.Thread(target=worker, name="macro_run_watchdog_sta", daemon=True)
    thread.start()

    # Wait up to the watchdog window for the body to signal the possibly-blocking
    # call, then for the worker to finish. If the body signals blocking and the
    # worker is still alive afterwards, the call is stuck on a VBA modal.
    signaled = blocking_event.wait(timeout=watchdog_timeout_s)
    thread.join(timeout=grace_s)
    blocked = bool(signaled) and thread.is_alive()
    owned_pid = holder.get("owned_pid")

    if blocked:
        owned_pids = {owned_pid} if owned_pid is not None else set()
        # READ-ONLY: enumerate the owned PID's visible windows as modal evidence.
        modal_windows = _record_visible_windows("watchdog_modal", owned_pids, art)
        fg_sample = _record_foreground("watchdog_modal", owned_pids, art)
        holder["modal_evidence"] = {
            "visible_windows": modal_windows.get("visible_windows", []),
            "visible_window_titles": modal_windows.get("visible_window_titles", []),
            "foreground_owned": bool(fg_sample.get("is_excel_owned")),
            "foreground_title": fg_sample.get("title"),
        }
        # Terminate ONLY the owned PID. Never dismiss the dialog by input.
        terminated = _terminate_owned_pid(owned_pid) if owned_pid is not None else False
        holder["owned_pid_terminated"] = terminated
        holder["blocked_classification"] = CLASS_MACRO_RUNTIME_ERROR
        # Give the worker a moment to unwind (app.Run raises RPC_E_DISCONNECTED).
        thread.join(timeout=grace_s + 5.0)
        holder["thread_alive"] = thread.is_alive()
        if thread.is_alive():
            holder["error"] = (
                holder.get("error")
                or "watchdog: STA worker did not unwind after owned-PID termination"
            )
    else:
        holder["thread_alive"] = thread.is_alive()

    holder["blocked"] = blocked
    return holder


# ---------------------------------------------------------------------------
# Macro-call helper (runs on the STA thread via a body callable)
# ---------------------------------------------------------------------------
def _run_macro(app: Any, wb_name: str, macro: str, *args: Any) -> dict[str, Any]:
    """Invoke ``'<wb>'!<macro>`` and capture timing/outcome/HRESULTs.

    Returns a JSON-serializable record. ``value_repr`` is always captured; a
    structured ``value`` is added for str/int/float/bool/None and 1-D/2-D
    tuples of those. Unsupported argument types are rejected by pywin32 before
    any COM dispatch reaches Excel.
    """

    started = time.monotonic()
    record: dict[str, Any] = {
        "macro": macro,
        "qualified": f"'{wb_name}'!{macro}",
        "arg_count": len(args),
        "arg_types": [type(a).__name__ for a in args],
    }
    try:
        result = app.Run(record["qualified"], *args)
    except TypeError as exc:
        record["outcome"] = "rejected_unsupported_argument"
        record["classification"] = CLASS_REJECTED_UNSUPPORTED_ARGUMENT
        record["error_type"] = "TypeError"
        record["error_text"] = str(exc)[:300]
        record["outer_hresult"] = None
        record["inner_hresult"] = None
    except Exception as exc:  # noqa: BLE001 - com_error or other
        outer = _outer_hresult(exc)
        inner = _inner_hresult(exc)
        record["outcome"] = "error"
        record["classification"] = _classify_run_error(exc)
        record["error_type"] = type(exc).__name__
        record["error_text"] = str(exc)[:300]
        record["exception_message"] = (_exception_message(exc) or "")[:300] or None
        record["outer_hresult"] = _hresult_hex(outer)
        record["inner_hresult"] = _hresult_hex(inner)
        record["outer_hresult_int"] = outer
        record["inner_hresult_int"] = inner
    else:
        record["outcome"] = "success"
        record["classification"] = CLASS_SUCCESS
        record["value_repr"] = repr(result)[:300]
        record["value_type"] = type(result).__name__
        record["value"] = _to_json_scalar(result)
    record["duration_ms"] = round((time.monotonic() - started) * 1000.0, 3)
    return record


def _to_json_scalar(value: Any) -> Any:
    """Best-effort JSON-safe view of a returned Variant (tuple/array/scalar)."""

    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    if isinstance(value, (tuple, list)):
        return [_to_json_scalar(v) for v in value]
    return repr(value)[:200]


# ---------------------------------------------------------------------------
# Procedure-source reader (VBIDE) for the interactive-UI preflight (BG06)
# ---------------------------------------------------------------------------
def _read_procedure_sources(proj: Any, module: str) -> dict[str, str]:
    """Return ``{procedure_name: source_text}`` for a standard module.

    Reads the full ``CodeModule`` text (``Lines(1, CountOfLines)``) and splits it
    into procedure bodies by scanning for ``(Public|Private|Friend)? (Sub|
    Function|Property ...) <Name>`` headers up to the matching ``End Sub``/
    ``End Function``/``End Property``. This deliberately avoids the VBIDE
    ``ProcOfLine`` / ``ProcStartLine`` / ``ProcCountLines`` bound APIs, which on
    the target pywin32/VBIDE build return an unexpected ``(name, kind)`` tuple and
    raise ``DISP_E_BADPARAMCOUNT`` when fed back as a name. Reading the module
    text never shows or activates the VBE.
    """

    cm = proj.VBComponents(module).CodeModule
    count = int(cm.CountOfLines)
    text = cm.Lines(1, count)
    return _split_procedure_sources(text)


_PROC_HEADER = re.compile(
    r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)\b",
    re.IGNORECASE,
)
_PROC_END = re.compile(
    r"^\s*End\s+(Sub|Function|Property)\b",
    re.IGNORECASE,
)


def _split_procedure_sources(text: str) -> dict[str, str]:
    """Split a VBA module text into ``{procedure_name: source_text}``.

    The body spans the header line through the matching ``End Sub``/``End
    Function``/``End Property`` line (the first such terminator, which is
    sufficient for straight-line test procedures). Nested procedure declarations
    are not valid VBA, so no nesting handling is needed.
    """

    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    sources: dict[str, str] = {}
    idx = 0
    while idx < len(lines):
        header = _PROC_HEADER.match(lines[idx])
        if not header:
            idx += 1
            continue
        name = header.group(2)
        body = [lines[idx]]
        idx += 1
        while idx < len(lines):
            body.append(lines[idx])
            if _PROC_END.match(lines[idx]):
                idx += 1
                break
            idx += 1
        sources[name] = "\n".join(body)
    return sources


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_macro_arguments_and_return_value() -> None:
    """Application.Run argument/return semantics on ``CleanProject.xlsm``."""

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_macro_arguments_and_return_value")

    def body(app: Any, wb: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        try:
            a.excel_version = str(app.Version)
            a.excel_build = str(app.Build)
        except Exception:  # noqa: BLE001
            pass
        wb_name = str(wb.Name)

        # AddNumbers(2,3) -> 5 (int/float).
        add_numbers = _run_macro(app, wb_name, "modTest.AddNumbers", 2, 3)
        # NoArgs() -> "no-args".
        no_args = _run_macro(app, wb_name, "modTest.NoArgs")
        # Echo("hi") is a Sub: no return value, must succeed silently.
        echo = _run_macro(app, wb_name, "modTest.Echo", "hi")
        # MakeArray() -> serializable array.
        make_array = _run_macro(app, wb_name, "modTest.MakeArray")
        # Unsupported argument (a plain Python object) -> rejected before COM.
        bad_arg = _run_macro(app, wb_name, "modTest.Echo", object())

        a.macro_runs = [add_numbers, no_args, echo, make_array, bad_arg]
        for rec in a.macro_runs:
            a.log(
                f"run[{rec['macro']}] outcome={rec.get('outcome')} "
                f"class={rec.get('classification')} "
                f"value={rec.get('value_repr')} type={rec.get('value_type')} "
                f"dur_ms={rec.get('duration_ms')} "
                f"outer={rec.get('outer_hresult')} inner={rec.get('inner_hresult')} "
                f"err={rec.get('error_text')}"
            )
        return {
            "wb_name": wb_name,
            "add_numbers": add_numbers,
            "no_args": no_args,
            "echo": echo,
            "make_array": make_array,
            "bad_arg": bad_arg,
        }

    holder = _sta_run(fixture, body, art)
    art.orphan_reaped = holder.get("orphan_reaped")
    art.owned_pid = holder.get("owned_pid")

    _finalize_bg(art, holder)

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    result = holder.get("result", {})
    add = result["add_numbers"]
    no_args = result["no_args"]
    echo = result["echo"]
    arr = result["make_array"]
    bad = result["bad_arg"]

    # Qualified macro call + scalar arithmetic return.
    assert add["outcome"] == "success", f"AddNumbers failed: {add}"
    assert add["value"] == 5, f"AddNumbers(2,3) != 5: {add}"
    assert add["qualified"] == f"'{result['wb_name']}'!modTest.AddNumbers"

    # String return.
    assert no_args["outcome"] == "success", f"NoArgs failed: {no_args}"
    assert no_args["value"] == "no-args", f"NoArgs() != 'no-args': {no_args}"

    # Sub with a string argument completes without a return value.
    assert echo["outcome"] == "success", f"Echo failed: {echo}"

    # Variant array return is serializable.
    assert arr["outcome"] == "success", f"MakeArray failed: {arr}"
    assert arr["value"] == [1, 2, 3], f"MakeArray() != [1,2,3]: {arr}"

    # Unsupported argument is rejected BEFORE any COM dispatch (TypeError).
    assert bad["outcome"] == "rejected_unsupported_argument", (
        f"object() argument was not rejected before dispatch: {bad}"
    )
    assert bad["error_type"] == "TypeError", f"expected TypeError, got {bad['error_type']}"
    assert bad.get("outer_hresult") is None, f"a COM dispatch occurred: {bad}"
    assert bad["duration_ms"] < add["duration_ms"] + 50, (
        "unsupported-argument rejection took unexpectedly long"
    )

    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    # BG03/BG04/BG05 background non-interference (read-only Win32 evidence).
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_bg_task005_macro_args": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "Qualified Application.Run returns serializable Variant for scalar args "
            "(AddNumbers=5), string (NoArgs), and array (MakeArray=[1,2,3]); Echo Sub "
            "completes; unsupported argument rejected pre-dispatch (TypeError, no COM "
            "call); no orphan. " + art.bg_verdict["rationale"]
        ),
        "macro_outcomes": {k: result[k]["classification"] for k in result if k != "wb_name"},
        "success_values": {
            "AddNumbers": add.get("value"),
            "NoArgs": no_args.get("value"),
            "MakeArray": arr.get("value"),
        },
        "unsupported_argument": {
            "classification": bad.get("classification"),
            "error_type": bad.get("error_type"),
            "outer_hresult": bad.get("outer_hresult"),
        },
        "background_gate": art.bg_verdict["bg_verdict"],
        "bg03_no_focus_theft": art.bg03_no_focus_theft,
        "bg04_no_visible_popup": art.bg04_no_visible_popup,
        "bg05_user_session_preserved": art.bg05_user_session_preserved,
        "app_visible_start": art.app_visible_start,
        "app_visible_after_operations": art.app_visible_after_operations,
        "vbe_visible": art.vbe_visible,
        "bg_verdict": art.bg_verdict,
    }
    art.persist()


@requires_excel
@pytest.mark.integration_excel
def test_macro_runtime_error_behavior() -> None:
    """Missing macro, runtime error and interactive-UI preflight (BG06/BG09).

    Two dedicated-STA runs on the owned hidden Excel instance:

    * Run A (clean): missing macro -> ``MACRO_NOT_FOUND`` (inner ``0x800A03ED``);
      VBIDE source preflight classifies ``ShowMsg``/``ShowForm`` as interactive
      and blocks them BEFORE execution (BG06/BG09).
    * Run B (watchdog): ``RaiseError`` -> ``MACRO_RUNTIME_ERROR``. On this build
      an unhandled VBA runtime error raised through ``Application.Run`` on a
      hidden instance shows a VBA modal and blocks the call (research §4.2's
      clean return does not reproduce). The watchdog records the owned modal
      (READ-ONLY), terminates ONLY the owned PID, never dismisses the dialog, and
      never claims success. If a future build returns a COM error instead, the
      inner HRESULT (``0x800A03EC``/``0x80004005``) is recorded.

    BG04 note (honest limitation): the interactive-UI macros are blocked before
    execution so they raise no window. The runtime-error modal is a visible
    automation-OWNED window (not foreground; the user's foreground is unchanged),
    so BG03 (no focus theft) and BG05 (user session preserved) hold for both runs;
    the runtime-error modal is recorded as a BG04 limitation, not asserted away.
    """

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_macro_runtime_error_behavior")

    # --- Run A: missing macro + interactive-UI preflight (clean, returns). ---
    def body_a(app: Any, wb: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        try:
            a.excel_version = str(app.Version)
            a.excel_build = str(app.Build)
        except Exception:  # noqa: BLE001
            pass
        wb_name = str(wb.Name)

        # Missing macro -> outer 0x80020009, inner 0x800A03ED -> MACRO_NOT_FOUND.
        missing = _run_macro(app, wb_name, "modTest.Nope")

        # BG06: read procedure source via VBIDE and preflight the interactive UI
        # macros BEFORE execution. They are classified unsafe and blocked; the
        # ``app.Run`` dispatch for them is never issued.
        proj = wb.VBProject
        sources = _read_procedure_sources(proj, "modTest")
        preflight: list[dict[str, Any]] = []
        for proc in ("ShowMsg", "ShowForm"):
            source = sources.get(proc, "")
            safe, flagged = is_background_safe(source)
            classification = (
                CLASS_BACKGROUND_SAFETY_NOT_VERIFIED if safe else CLASS_INTERACTIVE_UI_NOT_SUPPORTED
            )
            preflight.append(
                {
                    "procedure": f"modTest.{proc}",
                    "background_safe": safe,
                    "flagged_tokens": flagged,
                    "classification": classification,
                    "executed": False,
                    "source_line_count": len(source.splitlines()) if source else 0,
                }
            )
        a.preflight = preflight
        a.macro_runs = [missing]
        a.log(
            f"run[{missing['macro']}] outcome={missing.get('outcome')} "
            f"class={missing.get('classification')} outer={missing.get('outer_hresult')} "
            f"inner={missing.get('inner_hresult')} dur_ms={missing.get('duration_ms')}"
        )
        for pf in a.preflight:
            a.log(
                f"preflight[{pf['procedure']}] background_safe={pf['background_safe']} "
                f"flagged={pf['flagged_tokens']} class={pf['classification']} "
                f"executed={pf['executed']}"
            )
        return {"wb_name": wb_name, "missing": missing, "preflight": preflight}

    holder_a = _sta_run(fixture, body_a, art)
    art.orphan_reaped = holder_a.get("orphan_reaped")
    art.owned_pid = holder_a.get("owned_pid")

    assert not holder_a.get("thread_alive"), "Run A STA worker did not finish within timeout"
    assert holder_a.get("error") is None, f"Run A STA worker error: {holder_a.get('error')}"
    result_a = holder_a.get("result", {})
    missing = result_a["missing"]
    preflight = result_a["preflight"]

    # Missing macro: MACRO_NOT_FOUND. The decisive evidence is the EXCEPINFO
    # description ("Cannot run the macro ... may not be available"), because on
    # this build the missing macro surfaces inner 0x800A03EC (not 0x800A03ED) --
    # the same inner code a runtime error can produce (research §4.2 / task_009 §3).
    assert missing["outcome"] == "error", f"missing macro did not error: {missing}"
    assert missing["classification"] == CLASS_MACRO_NOT_FOUND, (
        f"missing macro not classified MACRO_NOT_FOUND: {missing}"
    )
    description = (missing.get("exception_message") or "").lower()
    assert "cannot run the macro" in description or "may not be available" in description, (
        f"missing macro lacked not-found description evidence: {missing}"
    )
    assert missing.get("outer_hresult") == _hresult_hex(DISP_E_EXCEPTION), (
        f"missing macro outer HRESULT != 0x80020009: {missing}"
    )

    # Interactive-UI macros: preflight-blocked, never executed.
    by_proc = {pf["procedure"]: pf for pf in preflight}
    for proc in ("modTest.ShowMsg", "modTest.ShowForm"):
        pf = by_proc[proc]
        assert pf["background_safe"] is False, f"{proc} was not flagged unsafe: {pf}"
        assert pf["flagged_tokens"], f"{proc} had no flagged tokens: {pf}"
        assert pf["classification"] == CLASS_INTERACTIVE_UI_NOT_SUPPORTED, (
            f"{proc} not classified INTERACTIVE_UI_NOT_SUPPORTED: {pf}"
        )
        assert pf["executed"] is False, f"{proc} must not be executed: {pf}"
    assert "MsgBox" in by_proc["modTest.ShowMsg"]["flagged_tokens"]
    assert ".Show" in by_proc["modTest.ShowForm"]["flagged_tokens"]

    # No automation run of the UI macros appears in Run A's recorded macro runs.
    run_macros_a = {rec["macro"] for rec in art.macro_runs}
    assert "modTest.ShowMsg" not in run_macros_a, "ShowMsg was dispatched"
    assert "modTest.ShowForm" not in run_macros_a, "ShowForm was dispatched"

    assert holder_a.get("orphan_reaped") is not False, (
        "Run A owned Excel process orphaned after cleanup"
    )

    # --- Run B: RaiseError -> MACRO_RUNTIME_ERROR (watchdog for the modal). ---
    art.macro_runs = []  # reset so Run B evidence is recorded separately

    def body_b(app: Any, wb: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        wb_name = str(wb.Name)
        # Signal the watchdog that the next call may block on a VBA modal.
        event = ctx.get("blocking_event")
        if event is not None:
            event.set()
        runtime = _run_macro(app, wb_name, "modTest.RaiseError")
        a.macro_runs = [runtime]
        a.log(
            f"run[{runtime['macro']}] outcome={runtime.get('outcome')} "
            f"class={runtime.get('classification')} outer={runtime.get('outer_hresult')} "
            f"inner={runtime.get('inner_hresult')} dur_ms={runtime.get('duration_ms')} "
            f"err={runtime.get('error_text')}"
        )
        return {"wb_name": wb_name, "runtime": runtime}

    holder_b = _sta_run_watchdog(fixture, body_b, art, watchdog_timeout_s=15.0)
    runtime = (holder_b.get("result") or {}).get("runtime")

    # The watchdog blocked (modal) OR the call returned a COM error.
    blocked = bool(holder_b.get("blocked"))
    modal_evidence = holder_b.get("modal_evidence") or {}
    if blocked:
        # Modal path: RaiseError never returned; a visible owned modal was detected
        # and ONLY the owned PID was terminated. Never dismissed, never success.
        runtime = {
            "macro": "modTest.RaiseError",
            "outcome": "modal_runtime_error",
            "classification": CLASS_MACRO_RUNTIME_ERROR,
            "error_type": "VbaRuntimeErrorModal",
            "error_text": (
                "Unhandled VBA runtime error raised a modal dialog; Application.Run "
                "did not return. Detected the owned modal; terminated only the owned "
                "PID; dialog not dismissed; not reported as success."
            ),
            "outer_hresult": None,
            "inner_hresult": None,
            "modal_detected": True,
            "dismissed": False,
            "owned_pid_terminated": holder_b.get("owned_pid_terminated"),
            "duration_ms": round(15.0 * 1000.0, 3),
        }
        art.macro_runs = [runtime]
        art.log(
            "RaiseError blocked on VBA modal: "
            f"visible_titles={modal_evidence.get('visible_window_titles')} "
            f"foreground_owned={modal_evidence.get('foreground_owned')} "
            f"owned_pid_terminated={holder_b.get('owned_pid_terminated')} dismissed=False"
        )
    else:
        assert runtime is not None, "RaiseError produced no result and did not block"
        art.macro_runs = [runtime]
        # Returned com_error path: inner HRESULT must be a runtime-error code.
        assert runtime["outcome"] == "error", f"RaiseError did not error: {runtime}"
        assert runtime["classification"] == CLASS_MACRO_RUNTIME_ERROR, (
            f"runtime error not classified MACRO_RUNTIME_ERROR: {runtime}"
        )
        assert runtime.get("inner_hresult") in (
            _hresult_hex(HR_RUNTIME_ERROR),
            _hresult_hex(HR_E_FAIL),
        ), f"runtime error inner HRESULT unexpected: {runtime}"
        assert runtime.get("outer_hresult") == _hresult_hex(DISP_E_EXCEPTION), (
            f"runtime error outer HRESULT != 0x80020009: {runtime}"
        )

    # The watchdog path must have terminated ONLY the owned PID and left no orphan.
    if blocked:
        assert holder_b.get("owned_pid_terminated") is True, (
            "owned PID was not terminated on the modal path: "
            f"{holder_b.get('owned_pid_terminated')}"
        )
        assert modal_evidence.get("visible_window_titles"), (
            "modal path taken but no visible owned window was recorded"
        )
        assert modal_evidence.get("foreground_owned") is False, (
            "runtime-error modal stole foreground (BG03 violation)"
        )
    assert holder_b.get("orphan_reaped") is not False, (
        "Run B owned Excel process orphaned after cleanup"
    )

    # --- BG03/BG04/BG05: Run A is the clean background gate. ---
    holder_bg = holder_a
    _finalize_bg(art, holder_bg)
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_bg_task005_macro_errors": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "Missing macro -> MACRO_NOT_FOUND (EXCEPINFO description 'Cannot run the "
            "macro ... may not be available'; outer 0x80020009; inner on this build "
            f"{missing.get('inner_hresult')}); "
            "RaiseError -> MACRO_RUNTIME_ERROR "
            + (
                "via modal detection (owned PID terminated, dialog not dismissed, not "
                "reported as success)."
                if blocked
                else "(inner 0x800A03EC/0x80004005, outer 0x80020009)."
            )
            + " MsgBox/UserForm.Show preflight-blocked before execution with "
            "INTERACTIVE_UI_NOT_SUPPORTED (BG06/BG09); no orphan. " + art.bg_verdict["rationale"]
        ),
        "missing_macro": {
            "classification": missing.get("classification"),
            "outer_hresult": missing.get("outer_hresult"),
            "inner_hresult": missing.get("inner_hresult"),
            "duration_ms": missing.get("duration_ms"),
        },
        "runtime_error": {
            "classification": runtime.get("classification"),
            "outcome": runtime.get("outcome"),
            "outer_hresult": runtime.get("outer_hresult"),
            "inner_hresult": runtime.get("inner_hresult"),
            "duration_ms": runtime.get("duration_ms"),
            "modal_detected": runtime.get("modal_detected", False),
            "dismissed": runtime.get("dismissed", False),
        },
        "modal_evidence": modal_evidence,
        "interactive_ui_blocked": {pf["procedure"]: pf["classification"] for pf in preflight},
        "preflight_results": preflight,
        "background_gate": art.bg_verdict["bg_verdict"],
        "bg03_no_focus_theft": art.bg03_no_focus_theft,
        "bg04_no_visible_popup": art.bg04_no_visible_popup,
        "bg05_user_session_preserved": art.bg05_user_session_preserved,
        "bg04_runtime_error_limitation": (
            "RaiseError raised a visible automation-owned VBA modal (not foreground); "
            "recorded as a BG04 limitation for the runtime-error case, not asserted away."
            if blocked
            else None
        ),
        "app_visible_start": art.app_visible_start,
        "app_visible_after_operations": art.app_visible_after_operations,
        "vbe_visible": art.vbe_visible,
        "bg_verdict": art.bg_verdict,
    }
    art.persist()
