"""VBIDE access and component-export proof-of-concept for task_003 (binding).

This module is a bounded, self-contained proof-of-concept harness (not a
production abstraction). It proves, on the target Windows/Excel build, the VBIDE
round-trip that production ``adapters/excel`` VBA services (task_016/task_018)
depend on:

* ``Workbook.VBProject`` is reachable when VBOM trust is enabled, and a disabled
  trust is diagnosed as ``ACCESS_VBOM_DISABLED`` (HRESULT ``0x800A17B4``) without
  touching the registry;
* ``VBComponents`` exposes Standard (1), Class (2), UserForm (3, ``Designer`` not
  ``None``) and Document (100) modules, with ``ThisWorkbook``/``Sheet1`` present
  and never removed;
* ``CodeModule`` lines are readable for standard/class modules;
* ``References`` expose Name/GUID and ``IsBroken``;
* a UserForm exports as a ``.frm`` + ``.frx`` pair and the normalized code body
  hash is stable across re-export (``frx`` bytes are not reproducible);
* Document modules survive export and standard/class modules export as
  ``.bas``/``.cls``.

Every Excel COM call runs on a single dedicated STA thread with an owned
``DispatchEx`` instance (``Visible=False``, ``DisplayAlerts=False``). Cleanup
closes the workbook without saving, quits, releases references and
``CoUninitialize``s on that same thread, then polls for an orphaned owned PID.

Each test also records read-only background non-interference evidence
(addendum BG03/BG04/BG05): the foreground window (``GetForegroundWindow`` +
``GetWindowText`` + ``GetWindowThreadProcessId``) before creation, after
creation, after operations and after cleanup; the visible top-level windows of
the owned PID (``EnumWindows`` + ``IsWindowVisible``); ``Application.Visible``
at start and after operations; and a best-effort (non-activating) read of
``VBE.MainWindow.Visible`` when reachable. None of these queries activate, show,
click or type. Evidence is persisted to ``.ai/tmp/vbide_access_poc_<ts>.{json,txt}``.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vbide_access_poc.py -v -s
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import gc
import hashlib
import json
import subprocess
import sys
import tempfile
import threading
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.domain.errors import AccessVbomDisabledError, ErrorCode

EXCEL_PROGID = "Excel.Application"

# HRESULT for "Programmatic access to Visual Basic Project is not trusted".
ACCESS_VBOM_DISABLED_HRESULT = 0x800A17B4
# Outer DISP_E_EXCEPTION wrapper that carries the inner HRESULT above.
DISP_E_EXCEPTION = 0x80020009

ORPHAN_POLL_INTERVAL_S = 0.5
# Excel shutdown on the target host is variable (observed ~2.7s to ~18s+ cold);
# a 10s poll flakily reported a false orphan even though the owned process did
# exit. The orphan assertion is unchanged -- this only waits long enough to
# observe a slow-but-legitimate reaping. A genuinely never-exiting owned process
# still fails after this window.
ORPHAN_POLL_TIMEOUT_S = 30.0
STA_JOIN_TIMEOUT_S = 180.0

# VBIDE ``VBComponent.Type`` codes -> spec §7.2 names.
COMPONENT_TYPE_NAMES: dict[int, str] = {
    1: "standard_module",
    2: "class_module",
    3: "user_form",
    100: "document_module",
}

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


def _inner_hresult(exc: BaseException) -> int | None:
    """Extract the inner HRESULT from a pywintypes.com_error.

    ``win32com`` wraps an ``IDispatch`` failure as ``DISP_E_EXCEPTION`` whose
    third arg is the EXCEPINFO tuple; the inner HRESULT is element 5. Verified on
    the target build (see research §4.2). Returns ``None`` when absent.
    """

    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        candidate = args[2][5]
        if isinstance(candidate, int):
            return candidate & 0xFFFFFFFF
    return None


def _outer_hresult(exc: BaseException) -> int | None:
    """Return the outer HRESULT carried on the exception, if any."""

    hr = getattr(exc, "hresult", None)
    if isinstance(hr, int):
        return hr & 0xFFFFFFFF
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0] & 0xFFFFFFFF
    return None


def _is_vbom_disabled(exc: BaseException) -> bool:
    """True when the exception is the VBOM-disabled HRESULT (outer or inner)."""

    return (
        _outer_hresult(exc) == ACCESS_VBOM_DISABLED_HRESULT
        or _inner_hresult(exc) == ACCESS_VBOM_DISABLED_HRESULT
    )


def get_vbproject(
    accessor: Callable[[], Any],
) -> tuple[Any | None, AccessVbomDisabledError | None]:
    """Resolve ``Workbook.VBProject`` via ``accessor``, mapping VBOM-disabled access.

    On success returns ``(project, None)``. When ``accessor`` raises the
    VBOM-disabled ``com_error`` (``0x800A17B4`` as outer or inner HRESULT), returns
    ``(None, AccessVbomDisabledError)`` without ever touching the registry. Any
    other ``com_error`` is re-raised unchanged.
    """

    try:
        return accessor(), None
    except Exception as exc:  # noqa: BLE001 - re-raised below unless VBOM-disabled
        if _is_vbom_disabled(exc):
            return None, AccessVbomDisabledError(
                "Programmatic access to the VBA project is not trusted.",
                hresult=ACCESS_VBOM_DISABLED_HRESULT,
                context={
                    "outer_hresult": _outer_hresult(exc),
                    "inner_hresult": _inner_hresult(exc),
                },
            )
        raise


# ---------------------------------------------------------------------------
# Code normalization / hashing (research §2.6: hash the code body, not the .frx)
# ---------------------------------------------------------------------------
def normalize_code(text: str) -> str:
    """Normalize EOL/whitespace so the same source hashes identically.

    Strips ``\\r``, trims trailing whitespace per line and drops leading/trailing
    blank lines.
    """

    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in unified.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def hash_code(text: str) -> str:
    """Return the SHA-256 hex digest of the normalized code body."""

    return hashlib.sha256(normalize_code(text).encode("utf-8")).hexdigest()


def extract_frm_code_body(frm_text: str) -> str:
    """Return the VBA code body of an exported ``.frm``.

    A ``.frm`` is a designer header (``VERSION`` / ``Begin ... End``) followed by
    an ``Attribute VB_*`` block and then the code. The code body is everything
    after the last ``Attribute VB_`` line (research §2.2).
    """

    unified = frm_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = unified.split("\n")
    last_attr = -1
    for idx, line in enumerate(lines):
        if line.startswith("Attribute VB_"):
            last_attr = idx
    body_lines = lines[last_attr + 1 :] if last_attr >= 0 else lines
    return "\n".join(body_lines)


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
    protection: int | None = None
    has_password: bool | None = None
    vbom_status: str = ""
    components: list[dict[str, Any]] = field(default_factory=list)
    component_types_observed: list[int] = field(default_factory=list)
    references: list[dict[str, Any]] = field(default_factory=list)
    code_module_hashes: dict[str, str] = field(default_factory=dict)
    code_module_markers: dict[str, list[str]] = field(default_factory=dict)
    exports: dict[str, Any] = field(default_factory=dict)
    orphan_reaped: bool | None = None
    # Desktop non-interference evidence (addendum BG03/BG04/BG05). All values are
    # produced by read-only Win32 observation; see ``_record_foreground`` /
    # ``_record_visible_windows`` / ``_record_vbe_visible`` / ``_finalize_bg``.
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
    artifact_version: str = "task_003_binding"

    def log(self, msg: str) -> None:
        line = f"{_now_local_iso()} {msg}"
        self.log_lines.append(line)
        print(line, flush=True)

    def persist(self) -> dict[str, str]:
        _ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        json_path = _ARTIFACT_DIR / f"vbide_access_poc_{ts}.json"
        txt_path = _ARTIFACT_DIR / f"vbide_access_poc_{ts}.txt"
        json_path.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        header = (
            "task_003 VBIDE access and component-export PoC evidence\n"
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
    """Best-effort read of ``VBE.MainWindow.Visible`` without activating the VBE.

    Returns ``{"check": "reachable", "visible": <bool>}`` when the VBE object
    model is reachable, otherwise ``{"check": "not_reachable", ...}``. Reaching
    ``app.VBE`` must never show or activate the VBE window.
    """

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
    """Copy the BG03/BG04/BG05 samples off the worker ctx and compute the verdict.

    Runs on the main thread after the owned Excel instance has been reaped. It only
    reads Win32/process state; it never activates, shows, clicks or types. A
    disqualifying result sets ``bg_verdict['bg_verdict']`` to ``NO_GO`` when any
    automation-owned window is visible, any automation-owned window becomes
    foreground, or ``Application.Visible`` becomes ``True``.
    """

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

    # BG05: user-owned Excel processes present before the run must still be alive.
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
    """Assert the BG03/BG04/BG05 background contract holds; return ``bg_ok``.

    These assertions never weaken the existing task assertions; they add the
    addendum non-interference checks on top of them.
    """

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
    ``DispatchEx``, owned-PID resolution, workbook open, the body call, then close
    without save, ``Quit``, reference release, ``CoUninitialize`` and an orphan
    poll. The body returns a JSON-serializable dict; exceptions are captured so the
    main thread can assert/skip without leaking a process.
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
            # BG03/BG04: capture the user's foreground BEFORE creating Excel. The
            # owned PID set is empty here, so ``is_excel_owned`` is always False.
            ctx["foreground_before"] = _record_foreground("before", set(), art)
            app = win32com.client.DispatchEx(EXCEL_PROGID)
            app.Visible = False
            app.DisplayAlerts = False
            new_pids = _excel_pids() - baseline
            owned_pid = next(iter(new_pids)) if len(new_pids) == 1 else None
            ctx["owned_pid"] = owned_pid
            ctx["owned_pid_resolution"] = f"pid_delta={sorted(new_pids)}"
            owned_pids = {owned_pid} if owned_pid is not None else set()
            # BG03/BG04: immediately after creation the owned instance must not own
            # the foreground and must not expose a visible top-level window.
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
            # BG03/BG04: after all operations, re-check visibility, foreground and
            # the best-effort (non-activating) VBE visibility while Excel is alive.
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
            # Release child proxies (workbook + VBProject/RCWs created in ``body``)
            # while Excel is still alive; this is the only gc.collect() needed and it
            # avoids the benign RPC_E_DISCONNECTED (0x80010108) first-chance exception
            # that fires when the app proxy is finalized after ``Quit`` has torn the
            # server down. Do not gc.collect() again after ``Quit`` for that reason.
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
            # BG03/BG04: after cleanup the owned instance is gone, so the foreground
            # must not be an automation-owned window (owned PID set is empty here).
            ctx["foreground_after"] = _record_foreground("after", set(), art)
            holder["ctx"] = ctx

    thread = threading.Thread(target=worker, name="vbide_sta", daemon=True)
    thread.start()
    thread.join(timeout=STA_JOIN_TIMEOUT_S)
    holder["thread_alive"] = thread.is_alive()
    return holder


# ---------------------------------------------------------------------------
# Component / reference inventory (runs on the STA thread via the body callable)
# ---------------------------------------------------------------------------
def _enum_components(proj: Any) -> list[dict[str, Any]]:
    """Enumerate VBComponents into JSON-serializable dicts."""

    components: list[dict[str, Any]] = []
    for comp in proj.VBComponents:
        comp_type = int(comp.Type)
        has_designer = False
        if comp_type == 3:
            try:
                has_designer = comp.Designer is not None
            except Exception:  # noqa: BLE001
                has_designer = False
        line_count = 0
        try:
            line_count = int(comp.CodeModule.CountOfLines)
        except Exception:  # noqa: BLE001
            line_count = 0
        components.append(
            {
                "name": str(comp.Name),
                "type": comp_type,
                "type_name": COMPONENT_TYPE_NAMES.get(comp_type, "unknown"),
                "has_designer": has_designer,
                "line_count": line_count,
                "is_document_module": comp_type == 100,
            }
        )
    return components


def _read_references(proj: Any) -> list[dict[str, Any]]:
    """Read VBProject references (Name/GUID/major/minor/FullPath/IsBroken)."""

    refs: list[dict[str, Any]] = []
    for ref in proj.References:
        try:
            refs.append(
                {
                    "name": str(ref.Name),
                    "guid": str(ref.GUID),
                    "major": int(ref.Major),
                    "minor": int(ref.Minor),
                    "full_path": str(ref.FullPath),
                    "is_broken": bool(ref.IsBroken),
                }
            )
        except Exception as exc:  # noqa: BLE001
            refs.append({"error": f"{type(exc).__name__}: {exc}"})
    return refs


def _component_names(components: list[dict[str, Any]]) -> set[str]:
    return {str(c["name"]) for c in components if "name" in c}


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_vbproject_access_diagnostic() -> None:
    """Diagnose VBProject access, component types, code, references and protection."""

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_vbproject_access_diagnostic")

    # Unit-verify the ACCESS_VBOM_DISABLED detection with a SIMULATED exception.
    # This exercises the real mapping without needing a live disabled host.
    simulated_inner = ACCESS_VBOM_DISABLED_HRESULT - 0x100000000  # signed int32
    excepinfo = (
        0,
        "Microsoft Excel",
        "Programmatic access to Visual Basic Project is not trusted.",
        "xlmain11.chm",
        0,
        simulated_inner,
    )
    simulated = pythoncom.com_error(
        DISP_E_EXCEPTION - 0x100000000, "Exception occurred.", excepinfo
    )

    def _raise_simulated() -> Any:
        raise simulated

    _, mapped = get_vbproject(_raise_simulated)
    assert mapped is not None, "simulated VBOM-disabled exception was not mapped"
    assert mapped.code is ErrorCode.ACCESS_VBOM_DISABLED
    assert mapped.hresult == ACCESS_VBOM_DISABLED_HRESULT
    assert _is_vbom_disabled(simulated)
    art.log(
        "unit ACCESS_VBOM_DISABLED detection: PASS "
        f"(code={mapped.code.value}, hresult={mapped.hresult_hex})"
    )

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

        proj, vbom_err = get_vbproject(lambda: wb.VBProject)
        if vbom_err is not None:
            a.vbom_status = "vbom_disabled_on_this_host"
            a.log(f"VBProject access disabled on this host: {vbom_err}")
            return {"vbom_disabled": True, "error": str(vbom_err)}

        a.vbom_status = "enabled"
        a.log("VBProject access: enabled")

        # Protection (expect 0 = unprotected) and password diagnostic.
        try:
            a.protection = int(proj.Protection)
        except Exception as exc:  # noqa: BLE001
            a.log(f"Protection read failed: {exc!r}")
        try:
            a.has_password = bool(wb.HasPassword)
        except Exception as exc:  # noqa: BLE001
            a.log(f"HasPassword read failed: {exc!r}")
        a.log(f"Protection={a.protection} HasPassword={a.has_password}")

        # Component inventory.
        a.components = _enum_components(proj)
        a.component_types_observed = sorted({int(c["type"]) for c in a.components})
        a.log(f"Components: {[(c['name'], c['type']) for c in a.components]}")

        # CodeModule read + hash for modTest and clsTest.
        code_hashes: dict[str, str] = {}
        expected_markers = {
            "modTest": ("AddNumbers", "NoArgs", "Echo"),
            "clsTest": ("clsTest", "FullName"),
        }
        markers: dict[str, list[str]] = {}
        for name in ("modTest", "clsTest"):
            try:
                cm = proj.VBComponents(name).CodeModule
                text = cm.Lines(1, cm.CountOfLines)
                code_hashes[name] = hash_code(text)
                markers[name] = [token for token in expected_markers[name] if token in text]
                a.log(f"CodeModule {name}: {cm.CountOfLines} lines, hash={code_hashes[name][:16]}…")
            except Exception as exc:  # noqa: BLE001
                code_hashes[name] = f"error:{exc!r}"
                markers[name] = []
                a.log(f"CodeModule {name} read failed: {exc!r}")
        a.code_module_hashes = code_hashes
        a.code_module_markers = markers

        # References.
        a.references = _read_references(proj)
        a.log(f"References: {[(r.get('name'), r.get('is_broken')) for r in a.references]}")

        return {
            "vbom_disabled": False,
            "protection": a.protection,
            "has_password": a.has_password,
        }

    holder = _sta_run(fixture, body, art)
    art.orphan_reaped = holder.get("orphan_reaped")
    art.owned_pid = holder.get("owned_pid")
    if art.owned_pid is not None and not art.baseline_excel_pids:
        art.baseline_excel_pids = (holder.get("ctx") or {}).get("baseline_excel_pids", [])

    # BG03/BG04/BG05: copy the read-only Win32 samples and compute the verdict.
    _finalize_bg(art, holder)

    # Common lifecycle invariants.
    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    result = holder.get("result", {})

    # On this host VBOM is enabled; a disabled host must skip, not fail.
    if result.get("vbom_disabled"):
        pytest.skip("vbom disabled on this host (ACCESS_VBOM_DISABLED); registry not modified")

    # Component type coverage.
    names = _component_names(art.components)
    types_observed = set(art.component_types_observed)
    assert {1, 2, 3, 100}.issubset(types_observed), (
        f"missing component types; observed {sorted(types_observed)}"
    )
    by_type = {int(c["type"]): c for c in art.components if "type" in c}
    assert by_type[3]["has_designer"] is True, "UserForm Designer is None"
    assert {"ThisWorkbook", "Sheet1"}.issubset(names), "document modules missing"
    assert by_type[100]["name"] in names

    # Document modules were not removed.
    doc_names = {c["name"] for c in art.components if c.get("type") == 100}
    assert "ThisWorkbook" in doc_names and "Sheet1" in doc_names

    # CodeModule content non-empty and carries expected procedures.
    mod_hash = art.code_module_hashes.get("modTest", "")
    cls_hash = art.code_module_hashes.get("clsTest", "")
    assert len(mod_hash) == 64 and "error:" not in mod_hash, f"modTest hash missing: {mod_hash!r}"
    assert len(cls_hash) == 64 and "error:" not in cls_hash, f"clsTest hash missing: {cls_hash!r}"
    assert set(art.code_module_markers.get("modTest", [])) == {"AddNumbers", "NoArgs", "Echo"}, (
        f"modTest procedures missing: {art.code_module_markers.get('modTest')}"
    )
    assert "FullName" in art.code_module_markers.get("clsTest", []), (
        f"clsTest marker missing: {art.code_module_markers.get('clsTest')}"
    )

    # References: every reference has Name/GUID; none broken on a clean fixture.
    assert art.references, "no references enumerated"
    for ref in art.references:
        assert "name" in ref and "guid" in ref, f"reference missing Name/GUID: {ref}"
        assert ref.get("is_broken") is False, f"unexpected broken reference: {ref}"

    # Protection diagnostic (expect unprotected on CleanProject).
    assert art.protection == 0, f"expected Protection=0, got {art.protection}"

    # Orphan invariant.
    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    # BG03/BG04/BG05 background non-interference (read-only Win32 evidence).
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_r02_task003": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "VBProject reachable; component types 1/2/3/100 present; UserForm Designer "
            "present; Document modules preserved; CodeModule readable; references valid; "
            "Protection=0; ACCESS_VBOM_DISABLED detection verified via simulation; no orphan. "
            + art.bg_verdict["rationale"]
        ),
        "component_types": sorted(types_observed),
        "references_ok": all(r.get("is_broken") is False for r in art.references),
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
def test_userform_frm_frx_export_pair() -> None:
    """Export the UserForm as ``.frm`` + ``.frx``; verify normalized hash stability."""

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_userform_frm_frx_export_pair")

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

        proj, vbom_err = get_vbproject(lambda: wb.VBProject)
        if vbom_err is not None:
            a.vbom_status = "vbom_disabled_on_this_host"
            return {"vbom_disabled": True, "error": str(vbom_err)}
        a.vbom_status = "enabled"

        before_names = _component_names(_enum_components(proj))
        exports: dict[str, Any] = {}

        dir1 = Path(tempfile.mkdtemp(prefix="vbide_frm_"))
        dir2 = Path(tempfile.mkdtemp(prefix="vbide_frm_"))
        try:
            frm_comp = proj.VBComponents("frmTest")

            # First export.
            frm1 = dir1 / "frmTest.frm"
            frm_comp.Export(str(frm1))
            frx1 = dir1 / "frmTest.frx"
            exports["first"] = {
                "frm_exists": frm1.is_file(),
                "frx_exists": frx1.is_file(),
                "frm_size": frm1.stat().st_size if frm1.is_file() else None,
                "frx_size": frx1.stat().st_size if frx1.is_file() else None,
            }
            frm_text_1 = (
                frm1.read_text(encoding="utf-8", errors="replace") if frm1.is_file() else ""
            )
            hash1 = hash_code(extract_frm_code_body(frm_text_1))

            # Second export (independent dir) -> normalized code hash must match.
            frm2 = dir2 / "frmTest.frm"
            proj.VBComponents("frmTest").Export(str(frm2))
            frx2 = dir2 / "frmTest.frx"
            frm_text_2 = (
                frm2.read_text(encoding="utf-8", errors="replace") if frm2.is_file() else ""
            )
            hash2 = hash_code(extract_frm_code_body(frm_text_2))
            exports["second"] = {
                "frm_exists": frm2.is_file(),
                "frx_exists": frx2.is_file(),
                "code_hash": hash2,
                "frx_sha256": _sha256_file(frx2),
            }
            exports["first"]["code_hash"] = hash1
            exports["first"]["frx_sha256"] = _sha256_file(frx1)
            exports["code_hash_match"] = hash1 == hash2
            exports["first"]["code_body"] = extract_frm_code_body(frm_text_1)

            # Standard module -> .bas, document module -> .cls.
            bas = dir1 / "modTest.bas"
            proj.VBComponents("modTest").Export(str(bas))
            cls = dir1 / "ThisWorkbook.cls"
            proj.VBComponents("ThisWorkbook").Export(str(cls))
            exports["modTest_bas"] = {
                "exists": bas.is_file(),
                "size": bas.stat().st_size if bas.is_file() else None,
            }
            exports["thisworkbook_cls"] = {
                "exists": cls.is_file(),
                "size": cls.stat().st_size if cls.is_file() else None,
            }

            # Document modules must survive the export.
            after_names = _component_names(_enum_components(proj))
            exports["before_names"] = sorted(before_names)
            exports["after_names"] = sorted(after_names)
            exports["document_modules_intact"] = {"ThisWorkbook", "Sheet1"}.issubset(after_names)

            a.exports = exports
            for key in ("first", "second", "modTest_bas", "thisworkbook_cls"):
                a.log(f"export[{key}] = {exports[key]}")
            a.log(f"code_hash_match={exports['code_hash_match']}")
            a.log(f"document_modules_intact={exports['document_modules_intact']}")
            return {"vbom_disabled": False, "exports": exports}
        finally:
            for tmp in (dir1, dir2):
                for child in tmp.glob("*"):
                    with contextlib.suppress(OSError):
                        child.unlink()
                with contextlib.suppress(OSError):
                    tmp.rmdir()

    holder = _sta_run(fixture, body, art)
    art.orphan_reaped = holder.get("orphan_reaped")
    art.owned_pid = holder.get("owned_pid")

    # BG03/BG04/BG05: copy the read-only Win32 samples and compute the verdict.
    _finalize_bg(art, holder)

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    result = holder.get("result", {})

    if result.get("vbom_disabled"):
        pytest.skip("vbom disabled on this host (ACCESS_VBOM_DISABLED); registry not modified")

    exports = art.exports
    assert exports["first"]["frm_exists"] and exports["first"]["frx_exists"], (
        "UserForm export did not produce both .frm and .frx"
    )
    assert exports["second"]["frm_exists"] and exports["second"]["frx_exists"]
    assert exports["code_hash_match"] is True, "normalized code hash differs across re-export"
    assert exports["first"]["code_body"].strip(), "extracted .frm code body is empty"
    assert exports["modTest_bas"]["exists"] is True, "standard module .bas export missing"
    assert exports["thisworkbook_cls"]["exists"] is True, "document module .cls export missing"
    assert exports["document_modules_intact"] is True, "document modules removed after export"
    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    # BG03/BG04/BG05 background non-interference (read-only Win32 evidence).
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_r02_task003_export": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "UserForm exported as .frm + .frx; normalized code-body hash stable across "
            "re-export (frx bytes differ); modTest.bas and ThisWorkbook.cls exported and "
            "non-empty; Document modules preserved; no orphan. " + art.bg_verdict["rationale"]
        ),
        "code_hash": exports["first"].get("code_hash"),
        "code_hash_match": exports["code_hash_match"],
        "frx_sha256_first": exports["first"].get("frx_sha256"),
        "frx_sha256_second": exports["second"].get("frx_sha256"),
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


def _sha256_file(path: Path) -> str | None:
    """Return the SHA-256 hex digest of a file, or ``None`` when absent."""

    if not path.is_file():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()
