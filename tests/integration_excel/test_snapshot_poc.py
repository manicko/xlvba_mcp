"""Workbook SaveCopyAs snapshot + restore proof-of-concept for task_006 (binding).

This module is a bounded, self-contained proof-of-concept harness (not a
production abstraction). It proves, on the target Windows/Excel build
(Windows 10 LTSC 2019 17763 x64, Excel 16.0.14332 x64, Office UI language 1033
en-US), the snapshot/restore semantics that production
``application/snapshot_service.py`` (task_014) depends on:

* ``Workbook.SaveCopyAs`` on an OPEN macro-enabled workbook captures the current
  in-memory state into a SEPARATE ``.xlsm`` file and leaves the source file on
  disk byte-identical (SHA-256 and mtime unchanged);
* the snapshot opens as a separate workbook (``Workbooks.Count`` grows), exposes
  the same workbook model (sheet names) and keeps a stable file hash across a
  read-only reopen; its VBA project is preserved (macro-enabled round-trip);
* restore means closing the mutated working copy WITHOUT saving and reopening the
  baseline: the reopened file's SHA-256 equals the baseline (identity restored);
* ``SaveCopyAs`` re-serializes the OOXML container, so the snapshot's own SHA-256
  differs from the working file's SHA-256 even without mutation. The manifest
  therefore records BOTH ``source_sha256`` (working file) and ``copy_sha256``
  (snapshot), and the snapshot is its own immutable restore baseline
  (research §5.1 / §5.2).

Every Excel COM call runs on a single dedicated STA thread with an owned
``DispatchEx`` instance (``Visible=False``, ``DisplayAlerts=False``). Work happens
only inside a ``tempfile.TemporaryDirectory``; the committed fixture is never
opened for writing. Cleanup closes every owned workbook without saving, quits,
releases references and ``CoUninitialize``s on that same thread, then polls for an
orphaned owned PID. Non-owned (user) Excel processes are never touched.

Each test also records read-only background non-interference evidence
(addendum BG03/BG04/BG05): the foreground window (``GetForegroundWindow`` +
``GetWindowText`` + ``GetWindowThreadProcessId``) before creation, after
creation, after operations and after cleanup; the visible top-level windows of
the owned PID (``EnumWindows`` + ``IsWindowVisible``); ``Application.Visible``
at start and after operations; and a best-effort (non-activating) read of
``VBE.MainWindow.Visible`` when reachable. None of these queries activate, show,
click or type. Evidence is persisted to
``.ai/tmp/snapshot_poc_<test>_<ts>.{json,txt}`` per test.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_snapshot_poc.py -v -s
"""

from __future__ import annotations

import contextlib
import datetime as _dt
import gc
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pytest

EXCEL_PROGID = "Excel.Application"

# Excel ``FileFormat`` for a macro-enabled workbook (``.xlsm``). ``SaveCopyAs`` on
# an already-open ``.xlsm`` keeps the macro-enabled container; the constant is
# recorded as evidence of the intended format.
XLSM_FILE_FORMAT = 52

ORPHAN_POLL_INTERVAL_S = 0.5
# Excel shutdown on the target host is variable (observed ~2.7s to ~18s+ cold);
# a 10s poll flakily reported a false orphan even though the owned process did
# exit. The orphan assertion is unchanged -- this only waits long enough to
# observe a slow-but-legitimate reaping. A genuinely never-exiting owned process
# still fails after this window.
ORPHAN_POLL_TIMEOUT_S = 30.0
STA_JOIN_TIMEOUT_S = 180.0

# The explicit acceptance-criteria note both tests must emit: a snapshot restores
# the workbook FILE only and never reverts external side effects.
EXTERNAL_SIDE_EFFECTS_NOTE = (
    "A snapshot/restore reverts the workbook FILE only. It does NOT revert "
    "external side effects (database writes, email/network calls, files touched "
    "on disk, etc.). A snapshot is not a sandbox (spec §9.4)."
)

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
    reason=f"Excel COM snapshot PoC unavailable: {_AVAIL_REASON}",
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
# Time / hashing / file helpers
# ---------------------------------------------------------------------------
def _now_utc_iso() -> str:
    return _dt.datetime.now(_dt.UTC).isoformat()


def _now_local_iso() -> str:
    return _dt.datetime.now().astimezone().isoformat()


def _sha256_file(path: Path) -> str | None:
    """Return the SHA-256 hex digest of a file, or ``None`` when it is absent."""

    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _file_mtime(path: Path) -> float | None:
    """Return the file mtime (seconds since epoch), or ``None`` when absent."""

    if not path.is_file():
        return None
    return path.stat().st_mtime


def _file_size(path: Path) -> int | None:
    """Return the file size in bytes, or ``None`` when absent."""

    if not path.is_file():
        return None
    return path.stat().st_size


def _file_fingerprint(path: Path) -> dict[str, Any]:
    """Return a JSON-safe fingerprint (sha256 / mtime / size) of a file."""

    return {
        "path": str(path),
        "exists": path.is_file(),
        "sha256": _sha256_file(path),
        "mtime": _file_mtime(path),
        "size": _file_size(path),
    }


# ---------------------------------------------------------------------------
# COM error helper
# ---------------------------------------------------------------------------
def _com_error_details(exc: BaseException) -> dict[str, Any]:
    """Serialize a COM error into a structured, JSON-safe dictionary."""

    hr = getattr(exc, "hresult", None)
    if not isinstance(hr, int):
        args = getattr(exc, "args", ())
        if args and isinstance(args[0], int):
            hr = args[0]
    details: dict[str, Any] = {
        "type": type(exc).__name__,
        "message": str(exc)[:500],
        "hresult_hex": f"0x{hr & 0xFFFFFFFF:08X}" if isinstance(hr, int) else None,
    }
    args = getattr(exc, "args", ())
    if args:
        details["args_repr"] = [repr(a)[:200] for a in args]
    return details


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
# Workbook-model helper (runs on the STA thread via the body callable)
# ---------------------------------------------------------------------------
def _workbook_model(wb: Any) -> dict[str, Any]:
    """Return a JSON-safe model of an open workbook (name + sheet names)."""

    sheet_count = int(wb.Worksheets.Count)
    sheet_names = [str(wb.Worksheets(idx).Name) for idx in range(1, sheet_count + 1)]
    model: dict[str, Any] = {
        "name": str(wb.Name),
        "full_name": str(wb.FullName),
        "sheet_count": sheet_count,
        "sheet_names": sheet_names,
    }
    with contextlib.suppress(Exception):
        model["saved"] = bool(wb.Saved)
    return model


# ---------------------------------------------------------------------------
# Manifest builder (spec §9.1)
# ---------------------------------------------------------------------------
def _build_manifest(
    snapshot_path: Path,
    snapshot_sha256: str | None,
    source_path: Path,
    excel_version: str,
    excel_build: str,
    sheet_names: list[str],
    timestamp_utc: str,
) -> dict[str, Any]:
    """Build the snapshot ``manifest.json`` payload (spec §9.1)."""

    return {
        "snapshot_file": snapshot_path.name,
        "copy_sha256": snapshot_sha256,
        "source_path": str(source_path),
        "timestamp_utc": timestamp_utc,
        "excel_version": excel_version,
        "excel_build": excel_build,
        "workbook_model": {"sheet_names": sheet_names},
        "snapshot_method": "Workbook.SaveCopyAs",
        "external_side_effects_not_undone": True,
        "note": (
            "Created via Workbook.SaveCopyAs on an open macro-enabled workbook. "
            "SaveCopyAs re-serializes the OOXML container, so copy_sha256 differs "
            "from the working file's sha256; this snapshot is its own immutable "
            "restore baseline. " + EXTERNAL_SIDE_EFFECTS_NOTE
        ),
    }


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
    excel_version: str = ""
    excel_build: str = ""
    external_side_effects_not_undone: bool = True
    external_side_effects_note: str = EXTERNAL_SIDE_EFFECTS_NOTE
    baseline: dict[str, Any] = field(default_factory=dict)
    save_copy_as_params: dict[str, Any] = field(default_factory=dict)
    snapshot: dict[str, Any] = field(default_factory=dict)
    reopen: dict[str, Any] = field(default_factory=dict)
    manifest: dict[str, Any] = field(default_factory=dict)
    manifest_path: str = ""
    restore: dict[str, Any] = field(default_factory=dict)
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
    go_no_go: dict[str, Any] | None = None
    log_lines: list[str] = field(default_factory=list)
    paths: dict[str, str] = field(default_factory=dict)
    artifact_version: str = "task_006_binding"

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
        json_path = _ARTIFACT_DIR / f"snapshot_poc_{self.test_name}_{ts}.json"
        txt_path = _ARTIFACT_DIR / f"snapshot_poc_{self.test_name}_{ts}.txt"
        json_path.write_text(
            json.dumps(asdict(self), indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        header = (
            "task_006 SaveCopyAs snapshot + restore PoC evidence\n"
            f"test: {self.test_name}\n"
            f"artifact_version: {self.artifact_version}\n"
            f"started_utc: {self.started_utc}\n"
            f"started_local: {self.started_local}\n"
            f"external_side_effects_not_undone: {self.external_side_effects_not_undone}\n"
            f"note: {self.external_side_effects_note}\n"
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
    body: Any,
    art: PocArtifacts,
    work_path: Path,
    snap_path: Path,
    tmp: Path,
) -> dict[str, Any]:
    """Run ``body(app, ctx, art)`` on a dedicated STA thread with full cleanup.

    The temp paths are created on the main thread (fixture already copied in) and
    handed to the body through ``ctx`` so all COM work stays on the owned STA
    thread while the files live only inside the caller's ``TemporaryDirectory``.

    The worker owns the entire COM lifecycle: ``CoInitializeEx(APARTMENTTHREADED)``,
    ``DispatchEx``, ``Visible=False`` / ``DisplayAlerts=False``, owned-PID
    resolution, the body call, then a close of every owned workbook without saving,
    ``Quit``, reference release, ``CoUninitialize`` and an orphan poll. The body
    manages its own open/close cycles on the owned instance and returns a
    JSON-serializable dict; exceptions are captured so the main thread can
    assert/skip without leaking a process. Only workbooks on the owned ``DispatchEx``
    instance are closed; non-owned (user) Excel is never touched.
    """

    holder: dict[str, Any] = {}

    def worker() -> None:
        tid = threading.get_ident()
        ctx: dict[str, Any] = {
            "sta_thread_id": tid,
            "tmp_dir": str(tmp),
            "work_path": str(work_path),
            "snap_path": str(snap_path),
        }
        app: Any = None
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
            with contextlib.suppress(Exception):
                ctx["excel_version"] = str(app.Version)
                ctx["excel_build"] = str(app.Build)
            holder["result"] = body(app, ctx, art)
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
            # Close EVERY workbook on the owned instance without saving (reverse
            # order), then collect child RCWs while Excel is STILL alive, then Quit,
            # release, CoUninitialize and orphan-poll. Only the owned DispatchEx
            # instance is touched; user Excel is never affected.
            #
            # Ordering matters (verified on the target build, see task_003 PoC): the
            # single ``gc.collect()`` runs BEFORE ``Quit`` so that workbook / range /
            # worksheet RCWs created in ``body`` are released while the Excel server
            # is still up. Collecting AFTER ``Quit`` finalizes those proxies against a
            # torn-down server and raises the benign RPC_E_DISCONNECTED (0x80010108)
            # first-chance exception, so there is intentionally no second collect.
            if app is not None:
                try:
                    for idx in range(int(app.Workbooks.Count), 0, -1):
                        with contextlib.suppress(Exception):
                            app.Workbooks(idx).Close(SaveChanges=False)
                except Exception:  # noqa: BLE001
                    pass
                gc.collect()
                with contextlib.suppress(Exception):
                    app.Quit()
            # Drop the owned app RCW with the COM apartment still initialized.
            # Pumping once first lets any pending release marshal cleanly and
            # suppresses the benign RPC_E_DISCONNECTED (0x80010108) first-chance
            # exception that otherwise surfaces when the app proxy is finalized
            # after ``Quit`` has torn the server down (documented in task_003).
            with contextlib.suppress(Exception):
                pythoncom.PumpWaitingMessages()
            app = None
            ctx["app"] = None
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()
            holder["owned_pid"] = owned_pid if owned_pid is not None else ctx.get("owned_pid")
            holder["orphan_reaped"] = _orphan_poll(holder["owned_pid"])
            # BG03/BG04: after cleanup the owned instance is gone, so the foreground
            # must not be an automation-owned window (owned PID set is empty here).
            ctx["foreground_after"] = _record_foreground("after", set(), art)
            holder["ctx"] = ctx

    thread = threading.Thread(target=worker, name="snapshot_poc_sta", daemon=True)
    thread.start()
    thread.join(timeout=STA_JOIN_TIMEOUT_S)
    holder["thread_alive"] = thread.is_alive()
    return holder


# ---------------------------------------------------------------------------
# Test 1: SaveCopyAs + reopen as a separate workbook
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_save_copy_as_reopen() -> None:
    """SaveCopyAs an open .xlsm -> source unchanged -> snapshot reopens separately."""

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_save_copy_as_reopen")

    def body(app: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        a.excel_version = ctx.get("excel_version", "")
        a.excel_build = ctx.get("excel_build", "")

        tmp = Path(ctx["tmp_dir"])
        work_path = Path(ctx["work_path"])
        snap_path = Path(ctx["snap_path"])

        # Step 1 already copied fixture -> working file on the main thread; record
        # the working baseline fingerprint BEFORE opening.
        a.baseline = _file_fingerprint(work_path)
        baseline_sha = a.baseline["sha256"]
        a.log(f"working baseline sha256={baseline_bytes(baseline_sha)} mtime={a.baseline['mtime']}")

        # Step 2: open working, record the workbook model.
        wb = app.Workbooks.Open(
            str(work_path), UpdateLinks=0, ReadOnly=False, Notify=False, AddToMru=False
        )
        model_working = _workbook_model(wb)
        a.baseline["workbook_model_open"] = model_working
        a.baseline["workbooks_count_after_open"] = int(app.Workbooks.Count)
        a.log(
            f"opened working: count={a.baseline['workbooks_count_after_open']} "
            f"sheets={model_working['sheet_names']}"
        )

        # Step 3: SaveCopyAs (pywin32 exposes the filename arg only; the open .xlsm
        # keeps its macro-enabled container).
        save_params: dict[str, Any] = {
            "method": "Workbook.SaveCopyAs",
            "filename": str(snap_path),
            "intended_file_format": XLSM_FILE_FORMAT,
            "file_format_note": (
                "pywin32 SaveCopyAs accepts the filename only; an open .xlsm keeps "
                f"its macro-enabled container (FileFormat {XLSM_FILE_FORMAT})"
            ),
        }
        savecopyas_ok = True
        t0 = time.monotonic()
        try:
            wb.SaveCopyAs(str(snap_path))
        except Exception as exc:  # noqa: BLE001
            savecopyas_ok = False
            save_params["error"] = _com_error_details(exc)
            a.log(f"SaveCopyAs failed: {save_params['error']}")
        save_params["elapsed_s"] = round(time.monotonic() - t0, 4)
        save_params["ok"] = savecopyas_ok
        a.save_copy_as_params = save_params
        a.log(f"SaveCopyAs ok={savecopyas_ok} elapsed={save_params['elapsed_s']}s")

        # Step 4: snapshot exists; record its hash; assert the source is UNCHANGED.
        working_after = _file_fingerprint(work_path)
        snap_fp = _file_fingerprint(snap_path)
        a.snapshot = {
            "file": snap_fp,
            "source_after_savecopyas": working_after,
            "source_sha_unchanged": working_after["sha256"] == baseline_sha,
            "source_mtime_unchanged": working_after["mtime"] == a.baseline["mtime"],
            "workbook_still_open": True,
        }
        with contextlib.suppress(Exception):
            a.snapshot["workbook_saved_flag"] = bool(wb.Saved)
        a.log(
            f"snapshot sha256={baseline_bytes(snap_fp['sha256'])} "
            f"source_unchanged={a.snapshot['source_sha_unchanged']} "
            f"(snapshot!=source bytes: {snap_fp['sha256'] != baseline_sha})"
        )

        # Step 5: open the snapshot as a SEPARATE workbook; verify model + hash.
        snap_wb = app.Workbooks.Open(
            str(snap_path), UpdateLinks=0, ReadOnly=True, Notify=False, AddToMru=False
        )
        model_snapshot = _workbook_model(snap_wb)
        snapshot_sha_after_reopen = _sha256_file(snap_path)
        a.reopen = {
            "workbooks_count_with_snapshot_open": int(app.Workbooks.Count),
            "snapshot_is_separate_workbook": (
                model_snapshot["full_name"] != model_working["full_name"]
            ),
            "snapshot_model": model_snapshot,
            "working_model": model_working,
            "sheet_model_matches": (model_snapshot["sheet_names"] == model_working["sheet_names"]),
            "snapshot_sha_at_savecopyas": snap_fp["sha256"],
            "snapshot_sha_after_reopen": snapshot_sha_after_reopen,
            "snapshot_hash_stable": snapshot_sha_after_reopen == snap_fp["sha256"],
        }
        with contextlib.suppress(Exception):
            a.reopen["snapshot_vba_components"] = [
                str(c.Name) for c in snap_wb.VBProject.VBComponents
            ]
        a.log(
            f"reopen: count={a.reopen['workbooks_count_with_snapshot_open']} "
            f"separate={a.reopen['snapshot_is_separate_workbook']} "
            f"sheets_match={a.reopen['sheet_model_matches']} "
            f"hash_stable={a.reopen['snapshot_hash_stable']}"
        )
        snap_wb.Close(SaveChanges=False)

        # Step 6: write manifest.json in the temp dir and record it.
        manifest = _build_manifest(
            snapshot_path=snap_path,
            snapshot_sha256=snap_fp["sha256"],
            source_path=work_path,
            excel_version=a.excel_version,
            excel_build=a.excel_build,
            sheet_names=model_working["sheet_names"],
            timestamp_utc=_now_utc_iso(),
        )
        manifest["source_sha256"] = baseline_sha
        manifest_path = tmp / "manifest.json"
        manifest_path.write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        a.manifest = manifest
        a.manifest_path = str(manifest_path)
        a.log(f"manifest written: {manifest_path}")

        # Leave only the working workbook open; the worker closes it in teardown.
        return {
            "baseline_sha256": baseline_sha,
            "snapshot_sha256": snap_fp["sha256"],
            "source_sha_unchanged": a.snapshot["source_sha_unchanged"],
            "source_mtime_unchanged": a.snapshot["source_mtime_unchanged"],
            "snapshot_is_separate_workbook": a.reopen["snapshot_is_separate_workbook"],
            "sheet_model_matches": a.reopen["sheet_model_matches"],
            "snapshot_hash_stable": a.reopen["snapshot_hash_stable"],
            "workbooks_count_with_snapshot_open": a.reopen["workbooks_count_with_snapshot_open"],
        }

    with tempfile.TemporaryDirectory(prefix="snapshot_poc_reopen_") as tmp_dir:
        tmp = Path(tmp_dir)
        work_path = tmp / "working.xlsm"
        snap_path = tmp / "snapshot.xlsm"
        # Copy the fixture into the temp dir; the fixture itself is never opened
        # for writing.
        shutil.copy2(fixture, work_path)
        holder = _sta_run(body, art, work_path, snap_path, tmp)

    _finalize_common(art, holder)
    # BG03/BG04/BG05: copy the read-only Win32 samples and compute the verdict.
    _finalize_bg(art, holder)
    result = holder.get("result", {}) or {}

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    assert result, "save_copy_as body returned no result"

    # Snapshot was produced by SaveCopyAs.
    assert art.save_copy_as_params.get("ok") is True, (
        f"SaveCopyAs failed: {art.save_copy_as_params.get('error')}"
    )
    # Source file is byte-identical after SaveCopyAs (sha + mtime unchanged).
    assert result["source_sha_unchanged"] is True, "SaveCopyAs altered the source file SHA-256"
    assert result["source_mtime_unchanged"] is True, "SaveCopyAs altered the source file mtime"
    # Snapshot exists with a recorded hash (and re-serializes -> differs from source).
    assert result["snapshot_sha256"], "snapshot SHA-256 not recorded"
    # Snapshot opens as a separate workbook with a matching sheet model.
    assert art.reopen["workbooks_count_with_snapshot_open"] >= 2, (
        "snapshot did not open as a separate workbook (Workbooks.Count did not grow)"
    )
    assert result["snapshot_is_separate_workbook"] is True, (
        "snapshot FullName equals the working file FullName (not a separate workbook)"
    )
    assert result["sheet_model_matches"] is True, "snapshot sheet model differs from the source"
    assert result["snapshot_hash_stable"] is True, (
        "snapshot file hash changed across read-only reopen"
    )
    # Manifest present with required fields.
    assert art.manifest.get("snapshot_file") == snap_path.name
    assert art.manifest.get("copy_sha256") == result["snapshot_sha256"]
    assert art.manifest.get("source_sha256") == result["baseline_sha256"]
    assert art.manifest.get("snapshot_method") == "Workbook.SaveCopyAs"
    assert art.manifest.get("external_side_effects_not_undone") is True

    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    # BG03/BG04/BG05 background non-interference (read-only Win32 evidence).
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_r06_task006_savecopyas": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "SaveCopyAs on the open .xlsm produced a snapshot that reopens as a separate "
            "workbook with a matching sheet model and a stable file hash; the source file "
            "stayed byte-identical (SHA-256 and mtime unchanged). SaveCopyAs re-serializes "
            "so copy_sha256 differs from source_sha256; the snapshot is its own immutable "
            "baseline. External side effects are NOT undone by a snapshot. No orphan. "
            + art.bg_verdict["rationale"]
        ),
        "baseline_sha256": result["baseline_sha256"],
        "snapshot_sha256": result["snapshot_sha256"],
        "sheet_model_matches": result["sheet_model_matches"],
        "snapshot_hash_stable": result["snapshot_hash_stable"],
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


# ---------------------------------------------------------------------------
# Test 2: restore baseline hash (close-without-save identity + snapshot path)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_restore_baseline_hash() -> None:
    """Mutate in memory, close w/o save -> baseline hash restored; snapshot path too."""

    fixture = _clean_project_or_skip()
    art = PocArtifacts(test_name="test_restore_baseline_hash")

    def body(app: Any, ctx: dict[str, Any], a: PocArtifacts) -> dict[str, Any]:
        a.sta_thread_id = ctx["sta_thread_id"]
        a.owned_pid = ctx.get("owned_pid")
        a.owned_pid_resolution = ctx.get("owned_pid_resolution", "")
        a.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
        a.excel_version = ctx.get("excel_version", "")
        a.excel_build = ctx.get("excel_build", "")

        tmp = Path(ctx["tmp_dir"])
        work_path = Path(ctx["work_path"])
        snap_path = Path(ctx["snap_path"])

        # Step 1: record the working baseline fingerprint BEFORE opening.
        a.baseline = _file_fingerprint(work_path)
        baseline_sha = a.baseline["sha256"]
        baseline_mtime = a.baseline["mtime"]
        a.log(f"working baseline sha256={baseline_bytes(baseline_sha)} mtime={baseline_mtime}")

        restore: dict[str, Any] = {"baseline_sha256": baseline_sha}

        # --- Scenario A: close-without-save identity on the working file ---
        # Step 2: open working, mutate the working copy in memory (do NOT save).
        wb = app.Workbooks.Open(
            str(work_path), UpdateLinks=0, ReadOnly=False, Notify=False, AddToMru=False
        )
        baseline_model = _workbook_model(wb)
        wb.Worksheets(1).Range("A1").Value = "MUTATED-A"
        restore["mutation_a_value"] = "MUTATED-A"
        with contextlib.suppress(Exception):
            restore["saved_flag_after_mutation"] = bool(wb.Saved)
        a.log(f"mutated A1='MUTATED-A' saved_flag={restore.get('saved_flag_after_mutation')}")

        # Step 3: close WITHOUT saving; the on-disk file must still equal baseline.
        wb.Close(SaveChanges=False)
        after_close = _file_fingerprint(work_path)
        restore["after_close_without_save"] = {
            "file": after_close,
            "sha_equals_baseline": after_close["sha256"] == baseline_sha,
            "mtime_equals_baseline": after_close["mtime"] == baseline_mtime,
        }
        a.log(
            f"after close(SaveChanges=False): sha_equals_baseline="
            f"{restore['after_close_without_save']['sha_equals_baseline']}"
        )

        # Step 4: reopen working; its SHA-256 must equal the baseline (identity).
        wb = app.Workbooks.Open(
            str(work_path), UpdateLinks=0, ReadOnly=False, Notify=False, AddToMru=False
        )
        reopened_sha = _sha256_file(work_path)
        a1_after_reopen: str | None = None
        with contextlib.suppress(Exception):
            a1_after_reopen = str(wb.Worksheets(1).Range("A1").Value)
        restore["reopen_working"] = {
            "sha256": reopened_sha,
            "sha_equals_baseline": reopened_sha == baseline_sha,
            "a1_value": a1_after_reopen,
        }
        a.log(
            f"reopen working: sha_equals_baseline="
            f"{restore['reopen_working']['sha_equals_baseline']} a1={a1_after_reopen!r}"
        )

        # --- Scenario B: restore-from-snapshot path ---
        # Step 5a: SaveCopyAs a snapshot of the (baseline) open workbook.
        wb.SaveCopyAs(str(snap_path))
        snap_fp = _file_fingerprint(snap_path)
        snap_sha = snap_fp["sha256"]
        a.save_copy_as_params = {
            "method": "Workbook.SaveCopyAs",
            "filename": str(snap_path),
            "intended_file_format": XLSM_FILE_FORMAT,
            "ok": True,
        }
        a.snapshot = {"file": snap_fp}
        a.log(f"snapshot sha256={baseline_bytes(snap_sha)}")

        # Step 5b: mutate the working copy in memory, then close WITHOUT saving.
        wb.Worksheets(1).Range("A1").Value = "MUTATED-B"
        with contextlib.suppress(Exception):
            restore["saved_flag_after_mutation_b"] = bool(wb.Saved)
        wb.Close(SaveChanges=False)
        a.log("mutated A1='MUTATED-B' then close(SaveChanges=False)")

        # Step 5c: reopen the SNAPSHOT; it is the immutable restore baseline.
        snap_wb = app.Workbooks.Open(
            str(snap_path), UpdateLinks=0, ReadOnly=True, Notify=False, AddToMru=False
        )
        snapshot_sha_after_reopen = _sha256_file(snap_path)
        snap_a1: str | None = None
        with contextlib.suppress(Exception):
            snap_a1 = str(snap_wb.Worksheets(1).Range("A1").Value)
        restore["restore_from_snapshot"] = {
            "snapshot_sha_at_savecopyas": snap_sha,
            "snapshot_sha_after_reopen": snapshot_sha_after_reopen,
            "snapshot_hash_matches": snapshot_sha_after_reopen == snap_sha,
            "snapshot_a1_value": snap_a1,
            "snapshot_a1_is_baseline": snap_a1 not in ("MUTATED-A", "MUTATED-B"),
        }
        a.log(
            "reopen snapshot: "
            f"hash_matches={restore['restore_from_snapshot']['snapshot_hash_matches']} "
            f"a1={snap_a1!r}"
        )
        snap_wb.Close(SaveChanges=False)

        # Write a manifest for the snapshot restore baseline (uses the temp dir).
        manifest = _build_manifest(
            snapshot_path=snap_path,
            snapshot_sha256=snap_sha,
            source_path=work_path,
            excel_version=a.excel_version,
            excel_build=a.excel_build,
            sheet_names=baseline_model["sheet_names"],
            timestamp_utc=_now_utc_iso(),
        )
        manifest["source_sha256"] = baseline_sha
        manifest_path = tmp / "manifest.json"
        manifest_path.write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        a.manifest = manifest
        a.manifest_path = str(manifest_path)
        a.log(f"manifest written: {manifest_path}")

        # Both scenarios share the working baseline hash for reference.
        restore["baseline_model"] = baseline_model
        a.restore = restore
        return {
            "baseline_sha256": baseline_sha,
            "after_close_sha_equals_baseline": restore["after_close_without_save"][
                "sha_equals_baseline"
            ],
            "reopen_sha_equals_baseline": restore["reopen_working"]["sha_equals_baseline"],
            "snapshot_hash_matches": restore["restore_from_snapshot"]["snapshot_hash_matches"],
            "snapshot_a1_is_baseline": restore["restore_from_snapshot"]["snapshot_a1_is_baseline"],
        }

    with tempfile.TemporaryDirectory(prefix="snapshot_poc_restore_") as tmp_dir:
        tmp = Path(tmp_dir)
        work_path = tmp / "working.xlsm"
        snap_path = tmp / "snapshot.xlsm"
        shutil.copy2(fixture, work_path)
        holder = _sta_run(body, art, work_path, snap_path, tmp)

    _finalize_common(art, holder)
    # BG03/BG04/BG05: copy the read-only Win32 samples and compute the verdict.
    _finalize_bg(art, holder)
    result = holder.get("result", {}) or {}

    assert not holder.get("thread_alive"), "STA worker did not finish within timeout"
    assert holder.get("error") is None, f"STA worker error: {holder.get('error')}"
    assert result, "restore body returned no result"

    # Scenario A: close-without-save leaves the file byte-identical; reopen matches.
    assert result["after_close_sha_equals_baseline"] is True, (
        "close(SaveChanges=False) changed the on-disk working file (mutation leaked to disk)"
    )
    assert art.restore["after_close_without_save"]["mtime_equals_baseline"] is True, (
        "close(SaveChanges=False) changed the working file mtime"
    )
    assert result["reopen_sha_equals_baseline"] is True, (
        "reopened working file SHA-256 does not equal the baseline (restore identity broken)"
    )
    # Scenario B: the snapshot is an immutable restore baseline.
    assert result["snapshot_hash_matches"] is True, (
        "snapshot file hash changed across reopen (not an immutable baseline)"
    )
    assert result["snapshot_a1_is_baseline"] is True, (
        "snapshot leaked a mutation: "
        f"A1={art.restore['restore_from_snapshot']['snapshot_a1_value']!r}"
    )

    assert art.orphan_reaped is not False, "owned Excel process orphaned after cleanup"

    # BG03/BG04/BG05 background non-interference (read-only Win32 evidence).
    bg_ok = _assert_bg(art)

    art.go_no_go = {
        "go_no_go_r06_task006_restore": "GO" if bg_ok else "NO_GO",
        "rationale": (
            "Closing the mutated working copy without saving left the on-disk file "
            "byte-identical (SHA-256 and mtime unchanged); reopening it restored the "
            "baseline SHA-256 exactly. The SaveCopyAs snapshot is an immutable restore "
            "baseline: its file hash is stable across reopen and it does not carry the "
            "in-memory mutation. A snapshot restores the workbook FILE only and does NOT "
            "revert external side effects (DB/email/network/files). No orphan. "
            + art.bg_verdict["rationale"]
        ),
        "baseline_sha256": result["baseline_sha256"],
        "after_close_sha_equals_baseline": result["after_close_sha_equals_baseline"],
        "reopen_sha_equals_baseline": result["reopen_sha_equals_baseline"],
        "snapshot_hash_matches": result["snapshot_hash_matches"],
        "external_side_effects_not_undone": True,
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


# ---------------------------------------------------------------------------
# Shared helpers (keep the STA helper local to this module)
# ---------------------------------------------------------------------------
def _finalize_common(art: PocArtifacts, holder: dict[str, Any]) -> None:
    """Copy worker-owned facts onto the artifact and emit the side-effects note."""

    art.owned_pid = holder.get("owned_pid")
    art.orphan_reaped = holder.get("orphan_reaped")
    ctx = holder.get("ctx") or {}
    if not art.baseline_excel_pids:
        art.baseline_excel_pids = ctx.get("baseline_excel_pids", [])
    if not art.excel_version:
        art.excel_version = ctx.get("excel_version", "")
    if not art.excel_build:
        art.excel_build = ctx.get("excel_build", "")
    # Acceptance criterion: both tests explicitly emit the side-effects note.
    print(f"[external_side_effects_not_undone] {art.external_side_effects_not_undone}", flush=True)
    print(f"[note] {art.external_side_effects_note}", flush=True)


def baseline_bytes(sha: str | None) -> str:
    """Short display form of a SHA-256 for log lines."""

    return f"{sha[:16]}…" if sha else "<none>"
