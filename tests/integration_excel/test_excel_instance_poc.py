"""Isolated Excel COM / STA proof-of-concept for task_002 (planner v3, binding).

This module is a bounded, self-contained proof-of-concept harness. It is not a
production abstraction. Its purpose is to prove, on the target Windows/Excel
build, the single-STA-worker COM model that production
``adapters/excel/com_worker.py`` (task_010) is expected to use:

* a dedicated STA thread owns every Excel COM proxy created/used/released;
* ``win32com.client.DispatchEx("Excel.Application")`` creates a process that is
  distinct from any pre-existing user Excel session;
* a mandatory cross-thread affinity probe demonstrates that the owned proxy is
  confined to its creating STA thread (``RPC_E_WRONG_THREAD`` = ``0x8001010E``);
* ``Workbooks.Add`` / ``SaveAs`` / ``Open`` / ``Close`` and ``Application.Quit``
  complete on that STA thread without leaving an orphaned owned process;
* user Excel processes present at baseline are not touched.

The run persists machine- and human-readable evidence to
``.ai/tmp/excel_instance_poc_<YYYYMMDD_HHMMSS>.{json,txt}`` and computes a
``go_no_go_r01_task002`` verdict per planner v3 section 8.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_excel_instance_poc.py -v -s
"""

from __future__ import annotations

import base64
import contextlib
import ctypes
import datetime as _dt
import gc
import json
import platform
import struct
import subprocess
import sys
import tempfile
import threading
import time
import winreg
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import pytest

EXCEL_PROGID = "Excel.Application"

# HRESULT constants (COM).
RPC_E_WRONG_THREAD = 0x8001010E
CO_E_NOTINITIALIZED = 0x800401F0
MK_E_UNAVAILABLE = 0x800401E3
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A
XL_OPEN_WORKBOOK_XLSX = 51
ORPHAN_POLL_INTERVAL_S = 0.5
ORPHAN_POLL_TIMEOUT_S = 10.0

# Optional COM dependencies. ``None`` when unavailable so the module imports on
# non-Windows machines and the tests skip cleanly.
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


# ---------------------------------------------------------------------------
# Availability / skip helpers
# ---------------------------------------------------------------------------
def _excel_com_available() -> tuple[bool, str]:
    """Return (available, reason) describing whether this host can drive Excel."""

    if sys.platform != "win32":
        return False, f"non-Windows platform: {sys.platform}"
    if pythoncom is None or win32com is None:
        return False, "pywin32 (pythoncom/win32com) not importable"
    # Best-effort: confirm the ProgID resolves before declaring availability.
    try:  # pragma: no cover - depends on host
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:  # noqa: BLE001
        return True, "pywin32 present; COM init probe inconclusive"
    else:
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()
        return True, "pywin32 present"


_AVAILABLE, _AVAIL_REASON = _excel_com_available()

requires_excel = pytest.mark.skipif(
    not _AVAILABLE,
    reason=f"Excel COM PoC unavailable: {_AVAIL_REASON}",
)


# ---------------------------------------------------------------------------
# Time helpers
# ---------------------------------------------------------------------------
def _now_utc_iso() -> str:
    return _dt.datetime.now(_dt.UTC).isoformat()


def _now_local_iso() -> str:
    return _dt.datetime.now().astimezone().isoformat()


# ---------------------------------------------------------------------------
# COM error helpers
# ---------------------------------------------------------------------------
def _hresult_hex_from_exc(exc: BaseException) -> str | None:
    """Extract an HRESULT from a pywintypes.com_error-like exception."""

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


def _hresult_int(hex_str: str | None) -> int | None:
    if hex_str is None:
        return None
    try:
        return int(hex_str, 16)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Artifact logger
# ---------------------------------------------------------------------------
@dataclass
class PocArtifacts:
    """Structured evidence collector persisted as JSON + text."""

    started_utc: str = field(default_factory=_now_utc_iso)
    started_local: str = field(default_factory=_now_local_iso)
    matrix: dict[str, Any] = field(default_factory=dict)
    steps: list[dict[str, Any]] = field(default_factory=list)
    log_lines: list[str] = field(default_factory=list)
    baselines: dict[str, Any] = field(default_factory=dict)
    serialized_multiop_evidence: dict[str, Any] | None = None
    affinity_probe: dict[str, Any] | None = None
    message_pump_strategy: str = ""
    rpc_observations: list[dict[str, Any]] = field(default_factory=list)
    modal_notes: str = ""
    orphan_poll: list[dict[str, Any]] = field(default_factory=list)
    go_no_go: dict[str, Any] | None = None
    paths: dict[str, str] = field(default_factory=dict)
    artifact_version: str = "task_002_planner_v3"

    def add_step(
        self,
        name: str,
        thread_id: int | None,
        status: str,
        hresult: str | None = None,
        notes: str = "",
    ) -> None:
        step: dict[str, Any] = {
            "name": name,
            "thread_id": thread_id,
            "timestamp": _now_local_iso(),
            "status": status,
            "notes": notes,
        }
        if hresult is not None:
            step["hresult"] = hresult
        self.steps.append(step)
        self.append_log(f"[step] {name} thread={thread_id} status={status} hr={hresult} {notes}")

    def append_log(self, msg: str) -> None:
        line = f"{_now_local_iso()} {msg}"
        self.log_lines.append(line)
        print(line, flush=True)

    def persist(self) -> dict[str, str]:
        """Write JSON and text artifacts; return their paths."""

        _ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        json_path = _ARTIFACT_DIR / f"excel_instance_poc_{ts}.json"
        txt_path = _ARTIFACT_DIR / f"excel_instance_poc_{ts}.txt"
        payload = asdict(self)
        json_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
        header = (
            "task_002 isolated Excel COM/STA PoC evidence\n"
            f"artifact_version: {self.artifact_version}\n"
            f"started_utc: {self.started_utc}\n"
            f"started_local: {self.started_local}\n"
        )
        txt_path.write_text(
            header + "\n" + "\n".join(self.log_lines) + "\n",
            encoding="utf-8",
        )
        self.paths = {"json": str(json_path), "txt": str(txt_path)}
        print(f"[artifacts] json={json_path} txt={txt_path}", flush=True)
        return self.paths


# ---------------------------------------------------------------------------
# Environment matrix
# ---------------------------------------------------------------------------
def _read_registry_str(root: int, path: str, value: str) -> str | None:
    try:  # pragma: no cover - depends on host
        with winreg.OpenKey(root, path) as key:
            data, _ = winreg.QueryValueEx(key, value)
            return str(data) if data is not None else None
    except Exception:  # noqa: BLE001
        return None


def _collect_office_bitness() -> dict[str, Any]:
    result: dict[str, Any] = {"value": "unknown", "methods": []}
    try:  # pragma: no cover - depends on host
        platform_val = _read_registry_str(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            "Platform",
        )
        result["methods"].append("HKLM ClickToRun Configuration Platform")
        if platform_val:
            result["value"] = platform_val
            return result
    except Exception as exc:  # noqa: BLE001
        result["methods"].append(f"C2R query failed: {exc!r}")
    result["value"] = "unknown"
    return result


def _collect_windows_display_version() -> dict[str, Any]:
    out: dict[str, Any] = {
        "display_version": "unknown",
        "current_build": "unknown",
        "ubr": "unknown",
    }
    try:  # pragma: no cover - depends on host
        base = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        dv = _read_registry_str(winreg.HKEY_LOCAL_MACHINE, base, "DisplayVersion")
        cb = _read_registry_str(winreg.HKEY_LOCAL_MACHINE, base, "CurrentBuildNumber")
        ubr = _read_registry_str(winreg.HKEY_LOCAL_MACHINE, base, "UBR")
        out["display_version"] = dv or "unknown"
        out["current_build"] = cb or "unknown"
        out["ubr"] = ubr or "unknown"
    except Exception as exc:  # noqa: BLE001
        out["error"] = repr(exc)
    return out


def _collect_user_ui_language() -> dict[str, Any]:
    out: dict[str, Any] = {"value": "unknown", "methods": []}
    try:  # pragma: no cover - depends on host
        lang = ctypes.windll.kernel32.GetUserDefaultUILanguage()  # type: ignore[attr-defined]
        out["value"] = int(lang) & 0xFFFF
        out["methods"].append("GetUserDefaultUILanguage (kernel32)")
    except Exception as exc:  # noqa: BLE001
        out["methods"].append(f"GetUserDefaultUILanguage failed: {exc!r}")
    return out


def _collect_pywin32_version() -> str:
    try:  # pragma: no cover - depends on host
        import importlib.metadata as md

        return md.version("pywin32")
    except Exception:  # noqa: BLE001
        return "unknown"


def _collect_python_matrix() -> dict[str, Any]:
    return {
        "python_version": sys.version.split()[0],
        "python_full": sys.version,
        "python_executable": sys.executable,
        "python_bitness": struct.calcsize("P") * 8,
        "pywin32_version": _collect_pywin32_version(),
        "platform": platform.platform(),
        "platform_version": platform.version(),
        "windows_version_tuple": list(sys.getwindowsversion())[:4],
        "office_bitness": _collect_office_bitness(),
        "windows_display": _collect_windows_display_version(),
        "user_ui_language": _collect_user_ui_language(),
    }


def _collect_excel_matrix(app: Any) -> dict[str, Any]:
    """Read owned-instance Excel properties on the STA thread (best effort)."""

    out: dict[str, Any] = {}
    reads = {
        "excel_version": "Version",
        "excel_build": "Build",
        "excel_operating_system": "OperatingSystem",
        "excel_hwnd": "Hwnd",
        "excel_caption": "Caption",
        "excel_name": "Name",
    }
    for field_name, prop in reads.items():
        try:
            out[field_name] = str(getattr(app, prop))
        except Exception as exc:  # noqa: BLE001
            out[field_name] = f"error:{exc!r}"
    out["vbe_ui_language"] = _collect_vbe_ui_language(app)
    return out


def _collect_vbe_ui_language(app: Any) -> dict[str, Any]:
    """Best-effort VBE / Office UI language discovery (never mutates settings)."""

    result: dict[str, Any] = {
        "result": "not_discoverable",
        "confidence": "not_discoverable",
        "methods": [],
    }
    # (a) Office UI language via LanguageSettings.LanguageID(2) (msoLanguageIDUI).
    try:
        ui_lang = app.LanguageSettings.LanguageID(2)
        result["methods"].append(f"LanguageSettings.LanguageID(2)={ui_lang} (0x{int(ui_lang):X})")
        if isinstance(result.get("ui_language_id"), int) is False:
            result["ui_language_id"] = int(ui_lang)
        result["result"] = f"office_ui_language_id={int(ui_lang)}"
        result["confidence"] = "confirmed_via_LanguageSettings"
    except Exception as exc:  # noqa: BLE001
        result["methods"].append(f"LanguageSettings.LanguageID failed: {exc!r}")
    # (b) VBE MainWindow caption as a weak hint, read-only.
    with contextlib.suppress(Exception):
        vbe = app.VBE
        caption = vbe.MainWindow.Caption
        result["methods"].append(f"VBE.MainWindow.Caption={caption!r}")
        if result["result"] == "not_discoverable":
            result["result"] = f"vbe_caption_hint={caption!r}"
            result["confidence"] = "heuristic_vbe_caption"
    # (c) Windows user UI language context.
    ui = _collect_user_ui_language()
    result["methods"].append(f"GetUserDefaultUILanguage={ui['value']}")
    return result


# ---------------------------------------------------------------------------
# Process / window snapshots (N4)
# ---------------------------------------------------------------------------
def _run_powershell_json(script: str) -> tuple[list[dict[str, Any]] | None, str]:
    """Run a PowerShell script via EncodedCommand; return (rows, error)."""

    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    try:  # pragma: no cover - depends on host
        proc = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-OutputFormat",
                "Text",
                "-EncodedCommand",
                encoded,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception as exc:  # noqa: BLE001
        return None, f"powershell_launch_failed:{exc!r}"
    out = proc.stdout.strip()
    if proc.returncode != 0 and not out:
        return None, f"powershell_rc={proc.returncode}:{proc.stderr.strip()[:300]}"
    if not out or out in {"[]", "null"}:
        return [], ""
    try:
        data = json.loads(out)
    except Exception as exc:  # noqa: BLE001
        return None, f"json_parse_failed:{exc!r}:out={out[:300]}"
    if isinstance(data, dict):
        return [data], ""
    if isinstance(data, list):
        return data, ""
    return None, f"unexpected_json_type:{type(data).__name__}"


def _list_excel_processes_cim(
    art: PocArtifacts,
) -> tuple[list[dict[str, Any]], str, list[str]]:
    """Enumerate EXCEL.EXE via Get-CimInstance; fall back to tasklist."""

    partial: list[str] = []
    script = (
        "$ProgressPreference='SilentlyContinue'; "
        "$p = Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" "
        "-ErrorAction SilentlyContinue; "
        "if (-not $p) { '[]' } else { "
        "$p | Select-Object ProcessId,CreationDate,ExecutablePath,CommandLine "
        "| ConvertTo-Json -Compress }"
    )
    rows, err = _run_powershell_json(script)
    if rows is not None:
        procs = [
            {
                "pid": r.get("ProcessId"),
                "creation_date": str(r.get("CreationDate")),
                "executable_path": r.get("ExecutablePath"),
                "command_line": r.get("CommandLine"),
            }
            for r in rows
        ]
        return procs, "Get-CimInstance Win32_Process", partial
    partial.append(f"cim_failed:{err}")
    art.append_log(f"WARN CIM snapshot failed: {err}; falling back to tasklist")
    # Fallback: tasklist CSV.
    try:  # pragma: no cover - depends on host
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception as exc:  # noqa: BLE001
        partial.append(f"tasklist_failed:{exc!r}")
        return [], "tasklist (failed)", partial
    procs = []
    for line in proc.stdout.splitlines():
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            procs.append({"pid": int(parts[1]) if len(parts) > 1 else None})
    return procs, "tasklist CSV fallback", partial


def _list_excel_windows(
    art: PocArtifacts,
) -> tuple[list[dict[str, Any]], str, list[str]]:
    """Enumerate top-level windows for Excel processes (read-only)."""

    partial: list[str] = []
    if win32gui is None or win32process is None:
        return [], "win32gui (unavailable)", ["win32gui_or_win32process_missing"]
    try:  # pragma: no cover - depends on host
        target_pids = {int(p["pid"]) for p in _last_procs if p.get("pid") is not None}
    except Exception:  # noqa: BLE001
        target_pids = set()
    windows: list[dict[str, Any]] = []

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception as exc:  # noqa: BLE001
            partial.append(f"GetWindowThreadProcessId_failed:{exc!r}")
            return True
        if pid in target_pids:
            try:
                title = win32gui.GetWindowText(hwnd)
            except Exception as exc:  # noqa: BLE001
                title = f"<title_error:{exc!r}>"
                partial.append(f"GetWindowText_failed:{exc!r}")
            windows.append({"hwnd": hwnd, "pid": pid, "title": title})
        return True

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception as exc:  # noqa: BLE001
        partial.append(f"EnumWindows_failed:{exc!r}")
    return windows, "win32gui.EnumWindows + win32process.GetWindowThreadProcessId", partial


# ``_list_excel_windows`` needs the process list captured by the caller; keep a
# module-level pointer that snapshots set immediately before window enumeration.
_last_procs: list[dict[str, Any]] = []


def _snapshot_processes_and_windows(
    art: PocArtifacts,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    global _last_procs
    procs, proc_method, proc_partial = _list_excel_processes_cim(art)
    _last_procs = procs
    wins, win_method, win_partial = _list_excel_windows(art)
    meta = {
        "process_snapshot_method": proc_method,
        "window_snapshot_method": win_method,
        "snapshot_partial_failures": proc_partial + win_partial,
    }
    return procs, wins, meta


def _readonly_rot_snapshot(art: PocArtifacts) -> dict[str, Any]:
    """Read-only GetActiveObject snapshot (informational; never mutates)."""

    rot: dict[str, Any] = {"attempted": True, "result": None, "error": None}
    if win32com is None:  # pragma: no cover
        rot["error"] = "win32com unavailable"
        return rot
    app = None
    try:
        app = win32com.client.GetActiveObject(EXCEL_PROGID)
        wbs = app.Workbooks
        count = int(wbs.Count)
        names: list[str] = []
        for i in range(1, min(count, 50) + 1):
            with contextlib.suppress(Exception):
                names.append(str(wbs.Item(i).Name))
        rot["result"] = {"workbook_count": count, "workbook_names_sample": names}
    except Exception as exc:  # noqa: BLE001
        rot["error"] = f"{type(exc).__name__}:{_hresult_hex_from_exc(exc)}:{str(exc)[:200]}"
    finally:
        with contextlib.suppress(Exception):
            del app
    return rot


def snapshot_isolation_baseline(label: str, art: PocArtifacts) -> dict[str, Any]:
    """Capture user Excel PIDs/windows + read-only ROT under a named label."""

    art.append_log(f"Isolation baseline '{label}': capturing processes")
    procs, wins, meta = _snapshot_processes_and_windows(art)
    rot = _readonly_rot_snapshot(art)
    baseline = {
        "label": label,
        "timestamp": _now_local_iso(),
        "excel_processes": procs,
        "excel_windows": wins,
        "rot_snapshot": rot,
        "process_snapshot_method_used": meta["process_snapshot_method"],
        "window_snapshot_method_used": meta["window_snapshot_method"],
        "snapshot_partial_failures": meta["snapshot_partial_failures"],
    }
    art.baselines[label] = baseline
    art.add_step(
        f"baseline_{label}",
        None,
        "ok",
        notes=f"procs={len(procs)} windows={len(wins)} method={meta['process_snapshot_method']}",
    )
    return baseline


# ---------------------------------------------------------------------------
# Owned identity resolution
# ---------------------------------------------------------------------------
def _pids_from(procs: list[dict[str, Any]]) -> set[int]:
    return {int(p["pid"]) for p in procs if p.get("pid") is not None}


def _resolve_owned_pid(
    procs_before: list[dict[str, Any]],
    procs_after: list[dict[str, Any]],
    art: PocArtifacts,
) -> tuple[int | None, str]:
    before = _pids_from(procs_before)
    after = _pids_from(procs_after)
    new = sorted(after - before)
    art.append_log(f"PID delta (after - before): {new}")
    if len(new) == 1:
        return new[0], "pid_delta_single_new"
    if not new:
        return None, "pid_delta_none_new"
    return None, f"pid_delta_ambiguous:{new}"


# ---------------------------------------------------------------------------
# STA context and owned session
# ---------------------------------------------------------------------------
@dataclass
class _StaExcelContext:
    art: PocArtifacts | None = None
    sta_thread_id: int | None = None
    com_init_result: str = "not_started"
    baseline_before: dict[str, Any] | None = None
    baseline_after: dict[str, Any] | None = None
    owned_pid: int | None = None
    owned_pid_resolution: str = "not_attempted"
    owned_hwnd: Any = None
    owned_caption: Any = None
    app_visible_actual: Any = None
    app_display_alerts_actual: Any = None
    workbook_open_params: dict[str, Any] = field(default_factory=dict)
    temp_dir: tempfile.TemporaryDirectory[str] | None = None
    temp_xlsx_path: Path | None = None
    tracked_workbooks: list[Any] = field(default_factory=list)
    # Strong references to every COM proxy the fixture created. These MUST be
    # released before ``CoUninitialize`` so the owned process is reaped; keeping
    # them alive across uninit leaves an orphaned Excel process.
    app_ref: Any = None
    excel_matrix: dict[str, Any] = field(default_factory=dict)
    go_no_go: dict[str, Any] | None = None
    cleanup_complete: bool = False


def _pump_messages_best_effort() -> None:
    if pythoncom is None:  # pragma: no cover
        return
    with contextlib.suppress(Exception):
        pythoncom.PumpWaitingMessages()  # type: ignore[attr-defined]


def _init_sta(art: PocArtifacts) -> int:
    """CoInitializeEx(APARTMENTTHREADED) on the current thread; return thread id."""

    tid = threading.get_ident()
    try:
        hr = pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        result = _coinit_result_text(hr)
    except Exception as exc:  # noqa: BLE001
        result = f"error:{exc!r}"
    art.add_step("sta_init", tid, "ok", notes=f"CoInitializeEx(APARTMENTTHREADED)={result}")
    return tid


def _coinit_result_text(hr: Any) -> str:
    # pywin32 ``CoInitializeEx`` returns ``None`` on success and raises on failure;
    # some builds surface an HRESULT instead. Normalize both into readable text.
    if hr is None:
        return "S_OK (returned None on success)"
    if isinstance(hr, int):
        return f"0x{hr & 0xFFFFFFFF:08X}"
    return repr(hr)


def _read_app_props(app: Any, props: list[str]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for p in props:
        try:
            out[p] = str(getattr(app, p))
        except Exception as exc:  # noqa: BLE001
            out[p] = f"error:{exc!r}"
    return out


# ---------------------------------------------------------------------------
# Serialized multi-op evidence (N2)
# ---------------------------------------------------------------------------
def _run_serialized_multiop_evidence(
    app: Any, sta_thread_id: int, art: PocArtifacts
) -> dict[str, Any]:
    """Execute >=2 distinct COM ops on the STA thread in a defined order (N2)."""

    evidence: dict[str, Any] = {
        "sta_thread_id": sta_thread_id,
        "operations": [],
        "submission_completion_order": [],
        "same_boundary": False,
        "scope_note": (
            "Demonstrates same-thread serialized execution on the STA worker boundary; "
            "full MCP request queue remains task_010/task_011."
        ),
    }
    # Operation 1: read Application.Version.
    t0 = _now_local_iso()
    try:
        version = app.Version
        evidence["operations"].append(
            {
                "index": 1,
                "name": "read_application_version",
                "submitted_at": t0,
                "ok": True,
                "value": str(version),
            }
        )
        evidence["submission_completion_order"].append("op1_read_application_version")
        art.add_step(
            "serialized_op_1_read_version", sta_thread_id, "ok", notes=f"Version={version!r}"
        )
    except Exception as exc:  # noqa: BLE001
        evidence["operations"].append(
            {
                "index": 1,
                "name": "read_application_version",
                "submitted_at": t0,
                "ok": False,
                "error": _com_error_details(exc),
            }
        )
        evidence["submission_completion_order"].append("op1_read_application_version_failed")
        art.add_step(
            "serialized_op_1_read_version",
            sta_thread_id,
            "error",
            _hresult_hex_from_exc(exc),
            str(exc)[:200],
        )
        raise
    _pump_messages_best_effort()
    # Operation 2: read Application.Caption.
    t1 = _now_local_iso()
    try:
        caption = app.Caption
        evidence["operations"].append(
            {
                "index": 2,
                "name": "read_application_caption",
                "submitted_at": t1,
                "ok": True,
                "value": str(caption),
            }
        )
        evidence["submission_completion_order"].append("op2_read_application_caption")
        art.add_step(
            "serialized_op_2_read_caption", sta_thread_id, "ok", notes=f"Caption={caption!r}"
        )
    except Exception as exc:  # noqa: BLE001
        evidence["operations"].append(
            {
                "index": 2,
                "name": "read_application_caption",
                "submitted_at": t1,
                "ok": False,
                "error": _com_error_details(exc),
            }
        )
        evidence["submission_completion_order"].append("op2_read_application_caption_failed")
        art.add_step(
            "serialized_op_2_read_caption",
            sta_thread_id,
            "error",
            _hresult_hex_from_exc(exc),
            str(exc)[:200],
        )
        raise
    evidence["same_boundary"] = evidence["submission_completion_order"] == [
        "op1_read_application_version",
        "op2_read_application_caption",
    ]
    return evidence


# ---------------------------------------------------------------------------
# Mandatory affinity negative probe (N1)
# ---------------------------------------------------------------------------
def _affinity_probe(app: Any, sta_thread_id: int, art: PocArtifacts) -> dict[str, Any]:
    """Cross-thread read of the owned proxy must fail with RPC_E_WRONG_THREAD.

    The probe thread initializes COM first so that a wrong-thread marshalling
    failure (``0x8001010E``) is observed rather than ``CO_E_NOTINITIALIZED``.
    The underlying ``IDispatch`` is used to surface the raw HRESULT that the
    dynamic ``win32com`` wrapper would otherwise mask as ``AttributeError``.
    """

    outcome: dict[str, Any] = {
        "performed": True,
        "outcome": "unknown",
        "expected_hresult": f"0x{RPC_E_WRONG_THREAD:08X}",
        "detail": None,
        "probe_thread_id": None,
        "sta_thread_id": sta_thread_id,
    }
    try:
        raw = app._oleobj_  # underlying PyIDispatch
    except Exception as exc:  # noqa: BLE001
        outcome["outcome"] = "inconclusive_no_raw_dispatch"
        outcome["detail"] = f"_oleobj_ unavailable: {exc!r}"
        art.add_step("affinity_probe", sta_thread_id, "inconclusive", notes=outcome["detail"])
        return outcome

    probe_result: dict[str, Any] = {}

    def _probe() -> None:
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        except Exception as exc:  # noqa: BLE001
            probe_result["outcome"] = "probe_thread_com_init_failed"
            probe_result["detail"] = repr(exc)
            return
        try:
            dispid = raw.GetIDsOfNames("Version")[0]
            raw.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, 0)
            probe_result["outcome"] = "unexpected_success"
            probe_result["detail"] = "cross-thread read succeeded"
        except Exception as exc:  # noqa: BLE001
            probe_result["outcome"] = "failed"
            probe_result["details"] = _com_error_details(exc)
        finally:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    th = threading.Thread(target=_probe, name="poc_affinity_probe", daemon=True)
    started = time.monotonic()
    th.start()
    outcome["probe_thread_id"] = th.ident
    th.join(timeout=15.0)
    joined = not th.is_alive()

    if not joined:
        outcome["outcome"] = "inconclusive_probe_timeout"
        outcome["detail"] = "probe thread did not finish within 15s"
        art.add_step("affinity_probe", sta_thread_id, "inconclusive", notes=outcome["detail"])
        return outcome

    if probe_result.get("outcome") == "probe_thread_com_init_failed":
        outcome["outcome"] = "inconclusive_probe_thread_no_com"
        outcome["detail"] = probe_result.get("detail")
        art.add_step("affinity_probe", sta_thread_id, "inconclusive", notes=outcome["detail"])
        return outcome

    details = probe_result.get("details")
    hr = details.get("hresult_hex") if isinstance(details, dict) else None
    outcome["hresult_hex"] = hr
    outcome["detail"] = details

    if probe_result.get("outcome") == "unexpected_success":
        outcome["outcome"] = "unexpected_success"
        art.add_step("affinity_probe", sta_thread_id, "unexpected_success", notes=outcome["detail"])
    elif hr == f"0x{RPC_E_WRONG_THREAD:08X}":
        outcome["outcome"] = "failed_as_expected_rpc_e_wrong_thread"
        art.add_step(
            "affinity_probe", sta_thread_id, "expected_failure", hr, notes="RPC_E_WRONG_THREAD"
        )
    elif hr == f"0x{CO_E_NOTINITIALIZED:08X}":
        outcome["outcome"] = "inconclusive_co_e_notinitialized"
        art.add_step(
            "affinity_probe", sta_thread_id, "inconclusive", hr, notes="CO_E_NOTINITIALIZED"
        )
    else:
        outcome["outcome"] = f"failed_unexpected_hresult:{hr}"
        art.add_step(
            "affinity_probe", sta_thread_id, "unexpected_hresult", hr, notes=str(details)[:200]
        )

    outcome["elapsed_s"] = round(time.monotonic() - started, 3)
    return outcome


# ---------------------------------------------------------------------------
# Lifecycle on a temp workbook (N3)
# ---------------------------------------------------------------------------
def _lifecycle_temp_xlsx(app: Any, sta_thread_id: int, art: PocArtifacts) -> dict[str, Any]:
    """Add / SaveAs / Open / Close on a temp .xlsx (planner v3 section 5.2)."""

    params = {"UpdateLinks": 0, "ReadOnly": False, "Notify": False, "AddToMru": False}
    result: dict[str, Any] = {"workbook_open_params": params, "steps": []}
    tmp = tempfile.TemporaryDirectory(prefix="poc_excel_")
    try:
        path = str(Path(tmp.name) / "poc_lifecycle.xlsx")
        # 1. Add
        wb = app.Workbooks.Add()
        art.add_step("wb_add", sta_thread_id, "ok", notes=f"new workbook name={wb.Name!r}")
        result["added_name"] = str(wb.Name)
        # 2. SaveAs
        wb.SaveAs(path, FileFormat=XL_OPEN_WORKBOOK_XLSX)
        art.add_step(
            "wb_saveas",
            sta_thread_id,
            "ok",
            notes=f"path={path} FileFormat={XL_OPEN_WORKBOOK_XLSX}",
        )
        result["saved_path"] = path
        # 3. Close after save
        wb.Close(SaveChanges=False)
        art.add_step("wb_close_after_save", sta_thread_id, "ok", notes="SaveChanges=False")
        with contextlib.suppress(Exception):
            del wb
        # 4. Open with pinned params
        wb2 = app.Workbooks.Open(
            path,
            UpdateLinks=params["UpdateLinks"],
            ReadOnly=params["ReadOnly"],
            Notify=params["Notify"],
            AddToMru=params["AddToMru"],
        )
        art.add_step(
            "wb_open",
            sta_thread_id,
            "ok",
            notes=f"params={params}",
        )
        # 5. Verify
        result["opened_name"] = str(wb2.Name)
        result["opened_fullname"] = str(wb2.FullName)
        art.add_step(
            "wb_verify_open",
            sta_thread_id,
            "ok",
            notes=f"Name={wb2.Name!r} FullName={wb2.FullName!r}",
        )
        # 6. Close after open
        wb2.Close(SaveChanges=False)
        art.add_step("wb_close_after_open", sta_thread_id, "ok", notes="SaveChanges=False")
        with contextlib.suppress(Exception):
            del wb2
        result["ok"] = True
    except Exception as exc:  # noqa: BLE001
        result["ok"] = False
        result["error"] = _com_error_details(exc)
        art.add_step(
            "lifecycle_error", sta_thread_id, "error", _hresult_hex_from_exc(exc), str(exc)[:300]
        )
    finally:
        tmp.cleanup()
    return result


# ---------------------------------------------------------------------------
# Cleanup + orphan poll
# ---------------------------------------------------------------------------
def _pid_alive(pid: int) -> bool:
    out = subprocess.run(
        ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    return str(pid) in out.stdout


def _orphan_poll(
    owned_pid: int | None, sta_thread_id: int | None, art: PocArtifacts
) -> dict[str, Any]:
    result: dict[str, Any] = {"owned_pid": owned_pid, "orphan": None, "checkpoints": []}
    if owned_pid is None:
        result["orphan"] = "unknown_owned_pid"
        return result
    deadline = time.monotonic() + ORPHAN_POLL_TIMEOUT_S
    idx = 0
    while time.monotonic() < deadline:
        idx += 1
        alive = _pid_alive(owned_pid)
        art.orphan_poll.append(
            {"index": idx, "owned_pid": owned_pid, "alive": alive, "timestamp": _now_local_iso()}
        )
        art.add_step(
            f"orphan_poll_{idx}",
            sta_thread_id,
            "alive" if alive else "gone",
            notes=f"pid={owned_pid} alive={alive}",
        )
        if not alive:
            result["orphan"] = False
            result["poll_count"] = idx
            return result
        time.sleep(ORPHAN_POLL_INTERVAL_S)
    result["orphan"] = True
    result["poll_count"] = idx
    art.add_step(
        "orphan_poll_timeout",
        sta_thread_id,
        "orphan",
        notes=f"pid={owned_pid} still alive after {ORPHAN_POLL_TIMEOUT_S}s",
    )
    return result


def _quit_and_cleanup(
    ctx: _StaExcelContext,
    sta_thread_id: int,
    art: PocArtifacts,
) -> dict[str, Any]:
    """Quit owned instance, release every COM reference, orphan-poll.

    Critical ordering (verified on the target build): the owned Excel process is
    reaped only after *all* strong references to its proxies are dropped. Calling
    ``Quit`` and ``CoUninitialize`` while a proxy (e.g. ``app`` or a tracked
    workbook) is still referenced leaves an orphaned ``EXCEL.EXE`` process.
    """

    result: dict[str, Any] = {}
    app = ctx.app_ref
    # 1. Close any tracked owned workbooks (SaveChanges=False), best effort.
    for idx, wb in enumerate(list(ctx.tracked_workbooks)):
        try:
            wb.Close(SaveChanges=False)
            art.add_step(
                f"owned_wb_cleanup_{idx}",
                sta_thread_id,
                "ok",
                notes="Close(SaveChanges=False)",
            )
        except Exception as exc:  # noqa: BLE001
            art.add_step(
                f"owned_wb_cleanup_{idx}",
                sta_thread_id,
                "error",
                _hresult_hex_from_exc(exc),
                str(exc)[:200],
            )
    ctx.tracked_workbooks.clear()
    # 2. Quit.
    if app is not None:
        try:
            app.Quit()
            art.add_step("app_quit", sta_thread_id, "ok", notes="owned instance Quit issued")
            result["quit"] = "ok"
        except Exception as exc:  # noqa: BLE001
            art.add_step(
                "app_quit", sta_thread_id, "error", _hresult_hex_from_exc(exc), str(exc)[:200]
            )
            result["quit"] = f"error:{exc!r}"
    else:
        art.add_step("app_quit", sta_thread_id, "skipped", notes="no app ref")
        result["quit"] = "skipped"
    # 3. Release every reference, then collect so the RPC channel tears down.
    #    Drop the context's strong ref first; the caller drops its own after this
    #    returns (see fixture teardown), then a second gc.collect() runs.
    ctx.app_ref = None
    with contextlib.suppress(Exception):
        del app
    gc.collect()
    art.add_step(
        "release_refs_gc",
        sta_thread_id,
        "ok",
        notes="dropped ctx.app_ref; gc.collect()",
    )
    # 4. CoUninitialize on the same thread.
    try:
        pythoncom.CoUninitialize()
        art.add_step("couninitialize", sta_thread_id, "ok", notes="CoUninitialize on STA thread")
        result["coinitialize"] = "ok"
    except Exception as exc:  # noqa: BLE001
        art.add_step(
            "couninitialize", sta_thread_id, "error", _hresult_hex_from_exc(exc), str(exc)[:200]
        )
        result["coinitialize"] = f"error:{exc!r}"
    # 5. Orphan poll (after uninit so the proxy is fully released).
    result["orphan"] = _orphan_poll(ctx.owned_pid, sta_thread_id, art)
    return result


# ---------------------------------------------------------------------------
# Go/No-Go rubric (planner v3 section 8)
# ---------------------------------------------------------------------------
def _compute_go_no_go(ctx: _StaExcelContext) -> dict[str, Any]:
    art = ctx.art
    assert art is not None
    before = ctx.baseline_before or {}
    after = ctx.baseline_after or {}
    before_pids = _pids_from(before.get("excel_processes", []))
    after_pids = _pids_from(after.get("excel_processes", []))

    affinity = ctx.affinity_probe if isinstance(ctx.affinity_probe, dict) else {}
    affinity_performed = bool(affinity.get("performed"))

    multiop = art.serialized_multiop_evidence or {}
    multiop_ok = bool(multiop.get("same_boundary")) and len(multiop.get("operations", [])) >= 2

    orphan = art.orphan_poll  # populated by cleanup
    owned_pid = ctx.owned_pid
    # Orphan status is the FINAL checkpoint state: the process may still be alive
    # at poll_1 while shutting down but reaped later. Only the last checkpoint
    # (or the cleanup result) decides whether an orphan remains.
    if owned_pid is None:
        orphan_alive: bool | None = None
    elif art.matrix.get("cleanup", {}).get("orphan", {}).get("orphan") is not None:
        orphan_alive = bool(art.matrix["cleanup"]["orphan"]["orphan"])
    elif orphan:
        orphan_alive = bool(orphan[-1].get("alive", False))
    else:
        orphan_alive = None

    user_pids_missing = sorted(before_pids - after_pids) if before_pids else []
    user_harmed = len(user_pids_missing) > 0
    owned_present_after = (owned_pid in after_pids) if owned_pid is not None else None
    dispatch_ex = "DispatchEx" in art.matrix.get("dispatch_method", "")
    matrix_complete = all(
        k in art.matrix
        for k in (
            "timestamp_utc",
            "python_version",
            "python_bitness",
            "excel_version",
            "excel_build",
            "sta_thread_id_init",
            "dispatch_method",
            "process_snapshot_method",
            "window_snapshot_method",
            "app_visible",
            "app_display_alerts",
        )
    )

    reasons: list[str] = []

    # Disqualifying -> NO_GO.
    if owned_present_after or orphan_alive:
        reasons.append("owned Excel process still present after cleanup (orphan)")
    if user_harmed:
        reasons.append(f"user Excel PIDs disappeared: {user_pids_missing}")
    if not affinity_performed:
        reasons.append("affinity probe not performed (N1 mandatory)")
    elif affinity.get("outcome") == "unexpected_success":
        reasons.append("affinity probe unexpectedly succeeded (confinement broken)")
    elif affinity.get("outcome", "").startswith("inconclusive"):
        reasons.append(f"affinity probe inconclusive: {affinity.get('outcome')}")
    if not dispatch_ex:
        reasons.append("DispatchEx not used as creation method")
    if not matrix_complete:
        reasons.append("environment matrix incomplete for required fields")
    if not multiop_ok:
        reasons.append("serialized multi-op evidence incomplete")

    no_go = bool(reasons)

    if no_go:
        verdict = "NO_GO"
        rationale = "Disqualifying condition(s): " + "; ".join(reasons)
        return {
            "go_no_go_r01_task002": verdict,
            "rationale": rationale,
            "reasons": reasons,
            "follow_ups": [],
        }

    # Conditional vs Full Go: modal/RPC exercised?
    rpc_observed = any(
        o.get("hresult") in (f"0x{RPC_E_CALL_REJECTED:08X}", f"0x{RPC_E_SERVERCALL_RETRYLATER:08X}")
        for o in art.rpc_observations
    )
    modal_exercised = bool(art.rpc_observations) or bool(
        art.modal_notes and "not exercised" not in art.modal_notes.lower()
    )

    follow_ups = [
        "task_004: validate VBE compile + UIA modal/busy handling on target build",
        "task_005: validate Application.Run modal/RPC behavior on target build",
        "task_010: validate full MCP request-queue serialization (beyond STA boundary)",
    ]

    if not modal_exercised:
        verdict = "CONDITIONAL_GO"
        rationale = (
            "Core R01 proven on target build: DispatchEx isolation, single-STA ownership, "
            "affinity confinement (RPC_E_WRONG_THREAD), serialized multi-op evidence, clean "
            "lifecycle with no orphan, user session untouched. Modal/RPC busy paths not "
            "exercised this run; full R01 closure requires task_004/task_005/task_010 integration."
        )
    else:
        verdict = "GO"
        rationale = (
            "Core R01 and modal/RPC observation paths exercised; all ACs pass on target build."
        )
    return {
        "go_no_go_r01_task002": verdict,
        "rationale": rationale,
        "reasons": reasons,
        "modal_rpcs_exercised": modal_exercised,
        "rpc_observed": rpc_observed,
        "follow_ups": follow_ups,
    }


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------
@pytest.fixture
def poc_logger() -> PocArtifacts:
    art = PocArtifacts()
    art.message_pump_strategy = (
        "pythoncom.PumpWaitingMessages() opportunistically between bounded waits "
        "(orphan poll uses time.sleep; pumping applied at COM boundaries)."
    )
    # Environment matrix (planner v3 section 3). Static facts are collected here;
    # owned-instance Excel facts are merged in ``owned_excel_session``.
    art.matrix["timestamp_utc"] = art.started_utc
    art.matrix["timestamp_local"] = art.started_local
    art.matrix.update(_collect_python_matrix())
    yield art
    art.persist()


@pytest.fixture
def sta_context(poc_logger: PocArtifacts) -> Any:
    ctx = _StaExcelContext(art=poc_logger)
    ctx.sta_thread_id = _init_sta(poc_logger)
    poc_logger.matrix["sta_thread_id_init"] = ctx.sta_thread_id
    try:
        yield ctx
    finally:
        # Safety net: ensure COM is uninitialized if a test body left it initialized
        # on this thread (cleanup normally does this).
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()


@pytest.fixture
def owned_excel_session(sta_context: Any) -> Any:
    ctx: _StaExcelContext = sta_context
    art: PocArtifacts = ctx.art
    assert art is not None

    # Precondition baseline.
    ctx.baseline_before = snapshot_isolation_baseline("baseline_before", art)

    # DispatchEx (owned instance).
    app = win32com.client.DispatchEx(EXCEL_PROGID)
    ctx.app_ref = app
    art.add_step("dispatch_ex", ctx.sta_thread_id, "ok", notes=f"DispatchEx('{EXCEL_PROGID}')")
    art.matrix["dispatch_method"] = f"DispatchEx('{EXCEL_PROGID}')"

    # Owned settings.
    app.Visible = False
    app.DisplayAlerts = False
    ctx.app_visible_actual = app.Visible
    ctx.app_display_alerts_actual = app.DisplayAlerts
    art.add_step(
        "set_app_flags",
        ctx.sta_thread_id,
        "ok",
        notes=f"Visible={ctx.app_visible_actual!r} DisplayAlerts={ctx.app_display_alerts_actual!r}",
    )

    # Owned identity via PID delta.
    procs_after, wins_after, meta_after = _snapshot_processes_and_windows(art)
    owned_pid, resolution = _resolve_owned_pid(
        ctx.baseline_before.get("excel_processes", []), procs_after, art
    )
    ctx.owned_pid = owned_pid
    ctx.owned_pid_resolution = resolution
    ctx.owned_hwnd = app.Hwnd
    ctx.owned_caption = app.Caption
    art.add_step(
        "resolve_owned_identity",
        ctx.sta_thread_id,
        "ok" if owned_pid is not None else "unknown",
        notes=(
            f"pid={owned_pid} resolution={resolution} "
            f"hwnd={ctx.owned_hwnd} caption={ctx.owned_caption!r}"
        ),
    )

    # Excel matrix.
    ctx.excel_matrix = _collect_excel_matrix(app)
    art.matrix.update(ctx.excel_matrix)
    art.matrix["app_visible"] = ctx.app_visible_actual
    art.matrix["app_display_alerts"] = ctx.app_display_alerts_actual
    # N4 snapshot methods (required fields).
    art.matrix["process_snapshot_method"] = (ctx.baseline_before or {}).get(
        "process_snapshot_method_used", "unknown"
    )
    art.matrix["window_snapshot_method"] = (ctx.baseline_before or {}).get(
        "window_snapshot_method_used", "unknown"
    )
    art.matrix["snapshot_partial_failures"] = (ctx.baseline_before or {}).get(
        "snapshot_partial_failures", []
    )
    art.matrix["owned_pid"] = ctx.owned_pid
    art.matrix["owned_pid_resolution"] = ctx.owned_pid_resolution

    # Drop the fixture's local ``app`` reference so that ``ctx.app_ref`` is the
    # sole strong owner. If this local survives in the suspended generator frame
    # it keeps the proxy alive across ``CoUninitialize`` and orphans the owned
    # Excel process (verified on the target build).
    del app

    try:
        yield ctx
    finally:
        # Cleanup on the STA thread (planner v3 section 5.5).
        cleanup = _quit_and_cleanup(ctx, ctx.sta_thread_id or threading.get_ident(), art)
        ctx.cleanup_complete = True
        art.matrix["cleanup"] = cleanup
        # Post-cleanup baseline.
        ctx.baseline_after = snapshot_isolation_baseline("baseline_after", art)
        art.matrix["cleanup_orphan"] = cleanup.get("orphan")
        # Final Go/No-Go verdict computed AFTER cleanup so the orphan outcome and
        # baseline_after are included (planner v3 section 8). Persisted in the
        # artifact because ``poc_logger`` finalizes after this fixture.
        if art.modal_notes == "":
            art.modal_notes = (
                "No modal dialogs or RPC_E_CALL_REJECTED observed during this run; "
                "busy/modal handling not exercised. "
                "Pump strategy: PumpWaitingMessages at COM boundaries."
            )
        verdict = _compute_go_no_go(ctx)
        ctx.go_no_go = verdict
        art.go_no_go = verdict
        art.append_log(
            "FINAL go_no_go_r01_task002="
            f"{verdict['go_no_go_r01_task002']} :: {verdict['rationale']}"
        )
        # Cleanup temp dir if still present.
        if ctx.temp_dir is not None:
            with contextlib.suppress(Exception):
                ctx.temp_dir.cleanup()


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.integration_excel
def test_dispatch_ex_creates_isolated_excel_instance(owned_excel_session: Any) -> None:
    ctx: _StaExcelContext = owned_excel_session
    app = ctx.app_ref
    art = ctx.art
    assert art is not None

    # DispatchEx path + owned identity.
    assert art.matrix["dispatch_method"] == f"DispatchEx('{EXCEL_PROGID}')"
    assert ctx.owned_pid is not None, f"owned PID unresolved ({ctx.owned_pid_resolution})"
    assert ctx.owned_pid not in _pids_from(ctx.baseline_before.get("excel_processes", []))

    # STA thread consistency.
    assert ctx.sta_thread_id is not None
    com_steps = [s for s in art.steps if s.get("thread_id") is not None]
    assert com_steps, "no COM-touching steps recorded"
    assert all(s["thread_id"] == ctx.sta_thread_id for s in com_steps)

    # Serialized multi-op evidence (N2).
    multiop = _run_serialized_multiop_evidence(app, ctx.sta_thread_id, art)
    art.serialized_multiop_evidence = multiop
    assert multiop["same_boundary"], "serialized multi-op evidence missing boundary"
    assert len(multiop["operations"]) >= 2

    # Mandatory affinity probe (N1).
    affinity = _affinity_probe(app, ctx.sta_thread_id, art)
    ctx.affinity_probe = affinity
    art.affinity_probe = affinity
    assert affinity["performed"] is True
    assert affinity["outcome"] == "failed_as_expected_rpc_e_wrong_thread", (
        f"affinity probe did not fail as expected: {affinity}"
    )
    assert affinity.get("hresult_hex") == f"0x{RPC_E_WRONG_THREAD:08X}"

    # Baseline user-session checks (no user PID disappeared).
    ctx.baseline_after = snapshot_isolation_baseline("baseline_after_manual", art)
    before = _pids_from(ctx.baseline_before.get("excel_processes", []))
    after = _pids_from(ctx.baseline_after.get("excel_processes", []))
    assert not (before - after), f"user Excel PIDs disappeared: {sorted(before - after)}"

    # Record rubric basis (orphan re-evaluated after fixture cleanup; final
    # go/no-go is persisted in the fixture teardown artifact).
    art.modal_notes = (
        "No modal dialogs or RPC_E_CALL_REJECTED observed during this run; "
        "busy/modal handling not exercised. Pump strategy: PumpWaitingMessages at COM boundaries."
    )
    art.rpc_observations = []
    ctx.go_no_go = _compute_go_no_go(ctx)
    art.append_log(
        f"dispatch_ex test body complete; owned_pid={ctx.owned_pid} "
        f"affinity={affinity['outcome']} ({affinity.get('hresult_hex')})"
    )


@requires_excel
@pytest.mark.integration_excel
def test_excel_workbook_open_close_lifecycle(owned_excel_session: Any) -> None:
    ctx: _StaExcelContext = owned_excel_session
    app = ctx.app_ref
    art = ctx.art
    assert art is not None

    # Pinned app settings.
    assert ctx.app_visible_actual is False
    assert ctx.app_display_alerts_actual is False

    # Lifecycle on a temp workbook with pinned Open params (N3).
    lifecycle = _lifecycle_temp_xlsx(app, ctx.sta_thread_id, art)
    assert lifecycle.get("ok") is True, lifecycle
    params = lifecycle["workbook_open_params"]
    assert params == {"UpdateLinks": 0, "ReadOnly": False, "Notify": False, "AddToMru": False}
    # The reopened workbook's ``Name`` includes the saved file's base name
    # (with extension), not the transient ``Book1`` name of the added workbook.
    saved_name = Path(lifecycle["saved_path"]).name
    assert lifecycle.get("opened_name") == saved_name
    assert lifecycle.get("added_name") == "Book1"

    # Steps coverage.
    step_names = {s["name"] for s in art.steps}
    for required in (
        "dispatch_ex",
        "set_app_flags",
        "wb_add",
        "wb_saveas",
        "wb_close_after_save",
        "wb_open",
        "wb_close_after_open",
    ):
        assert required in step_names, f"missing step {required}"

    # Same STA thread for all COM steps.
    com_steps = [s for s in art.steps if s.get("thread_id") is not None]
    assert all(s["thread_id"] == ctx.sta_thread_id for s in com_steps)

    # Serialized multi-op + affinity evidence.
    multiop = _run_serialized_multiop_evidence(app, ctx.sta_thread_id, art)
    art.serialized_multiop_evidence = multiop
    assert multiop["same_boundary"]
    affinity = _affinity_probe(app, ctx.sta_thread_id, art)
    ctx.affinity_probe = affinity
    art.affinity_probe = affinity
    assert affinity.get("hresult_hex") == f"0x{RPC_E_WRONG_THREAD:08X}"

    # Pumping / RPC / modal notes.
    art.modal_notes = (
        "No modal dialogs observed during lifecycle run; pumping via PumpWaitingMessages at COM "
        "boundaries. Busy/modal sub-risk not exercised on target in this run."
    )
    art.rpc_observations = []

    art.append_log(f"lifecycle complete; open params={params}; steps={len(art.steps)}")
