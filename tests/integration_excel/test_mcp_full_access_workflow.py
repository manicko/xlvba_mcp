"""Real MCP-protocol end-to-end workflow over STDIO (task_b7, fix_007; PO-1..PO-6).

This is the first test that drives the built server **through the actual Model
Context Protocol boundary**: it launches ``python -m excel_vba_mcp`` as a child
process, speaks JSON-RPC to it with the official ``mcp`` stdio client
(``stdio_client`` + ``ClientSession``), and asserts the full default-access user
workflow end to end against a live Excel instance.

The workflow is **table-driven** and runs **two consecutive
settings-update / macro-execution cycles on one ``workbook_id``**:

1. ``excel_open`` (default => writable execute session);
2. ``snapshot_create`` first (PO-3: the pre-mutation snapshot precondition, also
   required by ``table_update_rows``);
3. ``macro_classify_background`` vetoes the interactive ``ShowMessage``
   (``INTERACTIVE_UI_NOT_SUPPORTED``); an unclassified ``Recalculate`` is blocked
   (``BACKGROUND_SAFETY_NOT_VERIFIED``);
4. ``macro_classify_background`` classifies ``modWorkflow.Recalculate`` safe ONCE;
5. **cycle 1** — ``table_update_rows`` sets ``SettingValue = 10`` on ``tblSettings``
   -> ``macro_run Recalculate`` -> ``Result!A1 == 20``;
6. **cycle 2** — ``table_update_rows`` sets ``SettingValue = 20`` -> ``macro_run``
   -> ``Result!A1 == 40`` (and ``!= 20``), proving the macro reads the LIVE table
   state on the same session. No ``excel_save`` / macro ``.Save`` happens between
   the classification and cycle 2 (the B7-F3 invariant): the on-disk working copy
   is untouched, so the hash-bound classification stays valid across both cycles.
7. ``excel_save`` -> ``excel_close(save_changes=False)``.

The main two-cycle test additionally closes the Phase-4 coverage gaps flagged in
``.ai/plans/full-edit-mode-orchestration-plan.md`` §7: it reads the settings table
**explicitly** via ``table_read_rows`` before any mutation (step 4), reads the
updated ``Sheet1!B2`` cell back after **each** ``table_update_rows`` (steps 6 and
11), and performs a final **non-VBA** ``range_write_values`` of a sentinel into
``Result!A3`` followed by a read-back (step 14) to prove a mutation surface other
than ``table_update_rows`` lands on the live session.

Background safety (PO-5) is verified per owned PID: a read-only Win32 window
enumeration (``EnumWindows`` + ``GetWindowThreadProcessId``) proves the owned
Excel instance has **no visible window** at stable points (so it cannot steal
focus; the server never calls ``SetForegroundWindow`` / ``ShowWindow`` /
``SendKeys``), and the owned PID is reaped after the server stops while every
pre-existing user Excel PID stays alive. There is no VM / container / isolated
session — this runs in the host's interactive Windows session.

The original fixture is never opened: each test copies ``FullAccessWorkflow.xlsm``
into a temp root and points ``EXCEL_VBA_MCP_ROOTS`` (and work / snapshot /
artifact / run / audit) at that root, then asserts both the repo fixture and the
temp copy stay byte-identical (PO-6).

A **fresh server subprocess** is started per test (never a reused worker), which
avoids the stale-worker / dead-owned-Excel condition (R-17). ``USERPROFILE`` /
``HOMEDRIVE`` / ``HOMEPATH`` / ``LOCALAPPDATA`` / ``APPDATA`` are intentionally
NOT overridden (VBOM trust is per profile).

Run serially (no xdist), against a host with Excel + VBOM trust::

    uv run pytest tests/integration_excel/test_mcp_full_access_workflow.py -q -p no:xdist -s
"""

from __future__ import annotations

import asyncio
import csv
import hashlib
import json
import subprocess
import sys
import time
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

import pytest

pytest.importorskip("win32com", reason="pywin32 required for the live Excel e2e test")
pytest.importorskip("pythoncom", reason="pywin32 required for the live Excel e2e test")
pytest.importorskip("win32gui", reason="pywin32 required for the window-enumeration probe")
pytest.importorskip("win32process", reason="pywin32 required for the window-enumeration probe")

import win32gui  # noqa: E402
import win32process  # noqa: E402
from mcp import ClientSession, StdioServerParameters, stdio_client  # noqa: E402
from mcp.types import CallToolResult  # noqa: E402

_PYWIN32 = sys.platform == "win32"

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_WORKFLOW_FIXTURE = _REPO_ROOT / "tests" / "fixtures" / "FullAccessWorkflow.xlsm"

# Per-request ceiling for a single MCP call. Generous enough for cold COM work,
# bounded so a genuinely hung server fails the test instead of blocking forever.
_CALL_TIMEOUT_S = 180.0

# Bounded poll for the owned Excel PID to come up after server start.
_HEALTH_POLL_TIMEOUT_S = 30.0
_HEALTH_POLL_INTERVAL_S = 0.5

# Bounded poll for the owned Excel PID to be reaped after server shutdown.
_REAP_POLL_TIMEOUT_S = 30.0
_REAP_POLL_INTERVAL_S = 0.5


# ---------------------------------------------------------------------------
# Fixtures / skip helpers
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(
            f"fixture missing: {path}; run "
            "`uv run python tests/fixtures/build_full_access_fixture.py` first"
        )
    return path


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _excel_pids() -> set[int]:
    """Return the PIDs of every running ``EXCEL.EXE`` (read-only ``tasklist``)."""

    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive: a probe failure means "unknown"
        return set()
    pids: set[int] = set()
    for row in csv.reader(out.stdout.splitlines()):
        if len(row) >= 2 and row[0].strip().upper() == "EXCEL.EXE":
            try:
                pids.add(int(row[1]))
            except ValueError:
                continue
    return pids


def _pid_alive(pid: int) -> bool:
    """Return whether ``pid`` is currently alive (generic ``tasklist`` probe)."""

    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive: a probe failure means "unknown"
        return False
    return str(pid) in out.stdout


def _wait_pid_dead(pid: int, *, timeout_s: float = _REAP_POLL_TIMEOUT_S) -> bool:
    """Poll until ``pid`` is gone (bounded); return whether it was reaped."""

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if pid not in _excel_pids():
            return True
        time.sleep(_REAP_POLL_INTERVAL_S)
    return pid not in _excel_pids()


def _visible_windows_for_pid(pid: int) -> list[int]:
    """Return the visible top-level window handles owned by ``pid`` (READ-ONLY).

    Enumerates top-level windows with ``EnumWindows`` and filters by process id +
    visibility. It never activates, focuses, moves, shows or dismisses anything:
    an owned Excel that stays ``Visible = False`` must report no visible window
    (so it cannot steal focus).
    """

    handles: list[int] = []

    def _callback(hwnd: int, _ctx: Any) -> bool:
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001 - defensive per-window probe
            return True
        if wpid == pid and win32gui.IsWindowVisible(hwnd):
            handles.append(hwnd)
        return True

    win32gui.EnumWindows(_callback, None)
    return handles


def _assert_owned_hidden(pid: int, context: str) -> None:
    """Assert the owned Excel PID has no visible window (no focus-steal surface)."""

    windows = _visible_windows_for_pid(pid)
    assert windows == [], (
        f"owned Excel PID {pid} has visible window(s) {windows} after {context}; "
        "a visible owned window could steal focus (BG03/BG04/BG05)."
    )


# ---------------------------------------------------------------------------
# Server harness
# ---------------------------------------------------------------------------
def server_env(base: Path, *, access_mode: str | None) -> dict[str, str]:
    """Environment for a fresh MCP server subprocess.

    Roots / audit / work / snapshot / artifact / run dirs are scoped to ``base``
    so the test never touches the repo tree. ``USERPROFILE`` / ``HOMEDRIVE`` /
    ``HOMEPATH`` / ``LOCALAPPDATA`` / ``APPDATA`` are intentionally **not**
    overridden: VBOM trust is per profile and those are inherited from the standard
    environment by the MCP stdio client. ``EXCEL_VBA_MCP_ACCESS_MODE`` is left
    unset (unless a restricted run is requested), so the server defaults to
    ``full``.
    """

    env = {
        "EXCEL_VBA_MCP_ROOTS": str(base),
        "EXCEL_VBA_MCP_WORK_DIR": str(base / "work"),
        "EXCEL_VBA_MCP_SNAPSHOT_DIR": str(base / "snapshots"),
        "EXCEL_VBA_MCP_ARTIFACT_DIR": str(base / "artifacts"),
        "EXCEL_VBA_MCP_RUN_DIR": str(base / "runs"),
        "EXCEL_VBA_MCP_AUDIT_LOG": str(base / "audit.jsonl"),
        "PYTHONUNBUFFERED": "1",
    }
    if access_mode is not None:
        env["EXCEL_VBA_MCP_ACCESS_MODE"] = access_mode
    return env


@asynccontextmanager
async def _running_server(
    base: Path, *, access_mode: str | None = None
) -> AsyncIterator[ClientSession]:
    """Start a fresh ``excel_vba_mcp`` STDIO server and yield an MCP session."""

    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_vba_mcp"],
        cwd=str(base),
        env=server_env(base, access_mode=access_mode),
    )
    async with (
        stdio_client(params, errlog=sys.stderr) as (read_stream, write_stream),
        ClientSession(read_stream, write_stream, read_timeout_seconds=_CALL_TIMEOUT_S) as session,
    ):
        await session.initialize()
        yield session


# ---------------------------------------------------------------------------
# Result helpers (assert the protocol-level shape and unwrap it)
# ---------------------------------------------------------------------------
async def _call(
    session: ClientSession, name: str, arguments: dict[str, Any] | None = None
) -> CallToolResult:
    result = await session.call_tool(name, arguments or {})
    assert isinstance(result, CallToolResult), f"unexpected result type: {type(result)!r}"
    return result


def structured(result: CallToolResult) -> Any:
    """Return the structured payload of a successful tool result.

    Tools whose return annotation is a union (``excel_open`` / ``macro_run`` /
    ``pivot_tables_list`` may return a ``RunningJob`` in background mode) are
    wrapped by the MCP SDK as ``{"result": <payload>}``; unwrap that transport
    envelope so callers see the domain DTO directly.
    """

    assert not result.is_error, f"expected success, got error result: {result.content!r}"
    if result.structured_content is not None:
        payload = result.structured_content
    else:
        assert result.content, "successful result carried no content"
        payload = json.loads(result.content[0].text)
    if isinstance(payload, dict) and set(payload.keys()) == {"result"}:
        return payload["result"]
    return payload


def _as_list(result: CallToolResult) -> list[Any]:
    """Unwrap a list-returning tool result (``{"result": [...]}`` or a bare list)."""

    data = structured(result)
    if isinstance(data, dict) and set(data.keys()) == {"result"}:
        inner = data["result"]
        assert isinstance(inner, list)
        return inner
    assert isinstance(data, list)
    return data


def error_code(result: CallToolResult) -> str:
    """Return the stable ``code`` of a structured (is_error) tool result."""

    assert result.is_error, f"expected an error result, got: {result.content!r}"
    assert result.content, "error result carried no content"
    payload = json.loads(result.content[0].text)
    code = payload["errorCode"]
    assert isinstance(code, str)
    return code


def macro_result(result: CallToolResult) -> dict[str, Any]:
    """Return the ``MacroRunResult`` payload of a successful ``macro_run`` call."""

    assert not result.is_error, f"macro_run returned an error result: {result.content!r}"
    data = structured(result)
    assert isinstance(data, dict)
    return data


def _single_cell(result: CallToolResult) -> Any:
    """Return the scalar value of a 1x1 ``range_read_values`` result."""

    data = structured(result)
    assert isinstance(data, dict)
    rows = data["values"]
    assert len(rows) == 1 and len(rows[0]) == 1, (
        f"expected a 1x1 read, got {data['rows']}x{data['cols']}"
    )
    return rows[0][0]["value"]


def _write_payload(value: Any) -> dict[str, Any]:
    """Build a 1x1 ``RangeWriteRequest`` payload for a single cell."""

    kind = "number" if isinstance(value, (int, float)) and not isinstance(value, bool) else "string"
    return {"rows": 1, "cols": 1, "values": [[{"kind": kind, "value": value}]]}


async def _read_result_a1(session: ClientSession, workbook_id: str) -> Any:
    """Read ``Result!A1`` (the observable output of ``Recalculate``)."""

    return _single_cell(
        await _call(
            session,
            "range_read_values",
            {"workbook_id": workbook_id, "sheet": "Result", "address": "A1"},
        )
    )


async def _read_cell(
    session: ClientSession, workbook_id: str, sheet: str, address: str
) -> Any:
    """Read a single cell's scalar value via ``range_read_values``."""

    return _single_cell(
        await _call(
            session,
            "range_read_values",
            {"workbook_id": workbook_id, "sheet": sheet, "address": address},
        )
    )


async def _write_cell(
    session: ClientSession, workbook_id: str, sheet: str, address: str, value: Any
) -> Any:
    """Write a single scalar cell via ``range_write_values``; return the structured result."""

    return structured(
        await _call(
            session,
            "range_write_values",
            {
                "workbook_id": workbook_id,
                "sheet": sheet,
                "address": address,
                "payload": _write_payload(value),
            },
        )
    )


async def _read_settings_rows(session: ClientSession, workbook_id: str) -> Any:
    """Read every data row of ``tblSettings`` via ``table_read_rows`` (Phase-4 step 4)."""

    return structured(
        await _call(
            session,
            "table_read_rows",
            {"workbook_id": workbook_id, "name": "tblSettings"},
        )
    )


def _table_rows_as_dicts(table_rows: dict[str, Any]) -> list[dict[str, Any]]:
    """Map ``TableRows.rows`` (positional per ``columns``) into ``{column: value}`` dicts."""

    columns = table_rows["columns"]
    return [
        {column: cell["value"] for column, cell in zip(columns, row, strict=True)}
        for row in table_rows["rows"]
    ]


async def _update_factor(session: ClientSession, workbook_id: str, value: int) -> None:
    """Set ``tblSettings`` row 1 ``SettingValue`` to ``value`` via ``table_update_rows``."""

    structured(
        await _call(
            session,
            "table_update_rows",
            {
                "workbook_id": workbook_id,
                "name": "tblSettings",
                "updates": [{"index": 1, "values": {"SettingValue": value}}],
            },
        )
    )


async def _run_recalculate(session: ClientSession, workbook_id: str) -> dict[str, Any]:
    """Run ``modWorkflow.Recalculate`` and return its ``MacroRunResult`` payload."""

    return macro_result(
        await _call(
            session,
            "macro_run",
            {"workbook_id": workbook_id, "macro_name": "modWorkflow.Recalculate"},
        )
    )


async def _wait_healthy(
    session: ClientSession, *, timeout_s: float = _HEALTH_POLL_TIMEOUT_S
) -> dict[str, Any]:
    """Poll ``excel_health`` until the worker is running with a live owned PID."""

    deadline = time.monotonic() + timeout_s
    health: dict[str, Any] = {}
    while time.monotonic() < deadline:
        health = structured(await _call(session, "excel_health"))
        if (
            health.get("status") == "ok"
            and health.get("worker_state") == "running"
            and health.get("accepting") is True
            and health.get("needs_recovery") is False
            and health.get("owned_pid") is not None
        ):
            return health
        await asyncio.sleep(_HEALTH_POLL_INTERVAL_S)
    return health


def _assert_healthy_and_live(health: dict[str, Any]) -> int:
    assert health.get("status") == "ok", f"excel_health not ok: {health}"
    assert health.get("worker_state") == "running", f"worker not running: {health}"
    assert health.get("accepting") is True, f"worker not accepting: {health}"
    assert health.get("needs_recovery") is False, f"worker needs recovery: {health}"
    owned_pid = health.get("owned_pid")
    assert isinstance(owned_pid, int) and owned_pid > 0, f"owned_pid missing: {health}"
    assert _pid_alive(owned_pid), f"owned PID {owned_pid} is not alive: {health}"
    return owned_pid


def _copy_fixture(base: Path) -> Path:
    """Copy the committed fixture into ``base/input`` (the original stays untouched)."""

    fixture = _fixture_or_skip(_WORKFLOW_FIXTURE)
    input_dir = base / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    target = input_dir / fixture.name
    target.write_bytes(fixture.read_bytes())
    return target


# ---------------------------------------------------------------------------
# Happy path: two-cycle table-driven workflow end to end (PO-1..PO-6)
# ---------------------------------------------------------------------------
@requires_excel
async def test_full_access_workflow_e2e(tmp_path: Path) -> None:
    """Drive the two-cycle, table-driven full-access workflow via the MCP boundary."""

    repo_hash_before = _sha256(_fixture_or_skip(_WORKFLOW_FIXTURE))
    # PID baseline BEFORE the server starts: every pre-existing user Excel PID.
    baseline_pids = _excel_pids()
    work_copy = _copy_fixture(tmp_path)
    temp_hash_before = _sha256(work_copy)

    async with _running_server(tmp_path) as session:
        # --- Tool catalog (no Excel state required). ---
        tools = (await session.list_tools()).tools
        tool_names = {tool.name for tool in tools}
        assert len(tool_names) >= 87, f"expected >= 87 tools, got {len(tool_names)}"
        assert "macro_classify_background" in tool_names

        # --- Fresh worker: healthy, running, live owned PID. ---
        owned_pid = _assert_healthy_and_live(await _wait_healthy(session))
        assert owned_pid not in baseline_pids, (
            f"owned PID {owned_pid} collided with a pre-existing user Excel PID"
        )
        _assert_owned_hidden(owned_pid, "server start")

        # --- Open default (no mode/read_only args): writable execute session. ---
        opened = structured(await _call(session, "excel_open", {"path": str(work_copy)}))
        assert opened["read_only"] is False, f"default open must be writable: {opened}"
        assert opened["mode"] == "execute", f"default open must be execute mode: {opened}"
        workbook_id = opened["workbook_id"]
        assert workbook_id

        status = structured(await _call(session, "excel_get_status", {"workbook_id": workbook_id}))
        assert status["mode"] == "execute", f"session mode: {status}"
        assert status["read_only"] is False, f"session read_only: {status}"
        _assert_owned_hidden(owned_pid, "excel_open")

        # --- PO-3: snapshot FIRST, before any content-modifying operation. ---
        # table_update_rows requires a snapshot; macro_run does not (CT-B7-macronapshot-
        # precondition is a SEPARATE task), so the e2e snapshots up front regardless.
        snap = structured(await _call(session, "snapshot_create", {"workbook_id": workbook_id}))
        assert snap.get("snapshot_id"), f"snapshot_create returned no id: {snap}"

        # --- The interactive procedure cannot be classified background_safe (R13 veto). ---
        vetoed = await _call(
            session,
            "macro_classify_background",
            {
                "workbook_id": workbook_id,
                "procedure": "modWorkflow.ShowMessage",
                "background_safe": True,
            },
        )
        assert error_code(vetoed) == "INTERACTIVE_UI_NOT_SUPPORTED"

        # --- An unclassified macro is blocked, never run (BG06 deny-by-default). ---
        unclassified = await _run_recalculate(session, workbook_id)
        assert unclassified["status"] == "blocked", f"unclassified macro must block: {unclassified}"
        assert unclassified["error_code"] == "BACKGROUND_SAFETY_NOT_VERIFIED", unclassified
        _assert_owned_hidden(owned_pid, "unclassified macro_run")

        # --- Classify Recalculate ONCE; the hash-bound classification then covers
        #     both cycles (no save between here and cycle 2 -> the on-disk hash is
        #     unchanged -> the classification stays valid: B7-F3 invariant). ---
        classified = structured(
            await _call(
                session,
                "macro_classify_background",
                {
                    "workbook_id": workbook_id,
                    "procedure": "modWorkflow.Recalculate",
                    "background_safe": True,
                },
            )
        )
        assert classified["background_safe"] is True, classified

        # --- Step 4: read the settings table EXPLICITLY (table_read_rows) and the
        #     initial Result!A1 BEFORE any mutation (Phase-4 step 4). ---
        settings = await _read_settings_rows(session, workbook_id)
        assert settings["name"] == "tblSettings", f"unexpected table: {settings}"
        assert settings["columns"] == ["SettingKey", "SettingValue"], (
            f"unexpected columns: {settings['columns']}"
        )
        settings_rows = _table_rows_as_dicts(settings)
        factor_rows = [r for r in settings_rows if r["SettingKey"] == "factor"]
        assert factor_rows == [{"SettingKey": "factor", "SettingValue": 5}], (
            f"initial factor row missing/incorrect: {settings_rows}"
        )
        assert {"SettingKey": "label", "SettingValue": "alpha"} in settings_rows, (
            f"initial label row missing: {settings_rows}"
        )
        assert await _read_result_a1(session, workbook_id) == 0, (
            "initial Result!A1 must be 0 before the first cycle"
        )

        # --- CYCLE 1: SettingValue 5 -> 10; Recalculate -> Result!A1 == 20. ---
        await _update_factor(session, workbook_id, 10)
        # Step 6: read the updated cell back and assert it equals the written value.
        assert await _read_cell(session, workbook_id, "Sheet1", "B2") == 10, (
            "cycle 1 Sheet1!B2 must read back the written value 10"
        )
        run1 = await _run_recalculate(session, workbook_id)
        assert run1["status"] == "success", f"cycle 1 Recalculate did not succeed: {run1}"
        cycle1_output = await _read_result_a1(session, workbook_id)
        assert cycle1_output == 20, f"cycle 1 expected Result!A1 == 20, got {cycle1_output!r}"
        _assert_owned_hidden(owned_pid, "cycle 1 macro_run")

        # --- CYCLE 2: SettingValue 10 -> 20; Recalculate -> Result!A1 == 40 (!= 20).
        #     NO excel_save / macro .Save between the classification and here. ---
        await _update_factor(session, workbook_id, 20)
        # Step 11: read the updated cell back and assert it equals the written value.
        assert await _read_cell(session, workbook_id, "Sheet1", "B2") == 20, (
            "cycle 2 Sheet1!B2 must read back the written value 20"
        )
        run2 = await _run_recalculate(session, workbook_id)
        assert run2["status"] == "success", f"cycle 2 Recalculate did not succeed: {run2}"
        cycle2_output = await _read_result_a1(session, workbook_id)
        assert cycle2_output == 40, f"cycle 2 expected Result!A1 == 40, got {cycle2_output!r}"
        assert cycle2_output != cycle1_output, (
            "cycle 2 output must differ from cycle 1 (proves the macro read the LIVE "
            f"table state, not a cached value): {cycle1_output!r} vs {cycle2_output!r}"
        )
        _assert_owned_hidden(owned_pid, "cycle 2 macro_run")

        # --- Step 14: at least one NON-VBA mutation that is NOT table_update_rows.
        #     Write a sentinel into Result!A3 via range_write_values and read it back,
        #     proving a second mutation surface lands on the live session. ---
        probe = "e2e-probe"
        write_result = await _write_cell(session, workbook_id, "Result", "A3", probe)
        assert write_result.get("verified") is True, (
            f"range_write_values postcondition did not verify: {write_result}"
        )
        assert await _read_cell(session, workbook_id, "Result", "A3") == probe, (
            "step 14 Result!A3 must read back the written sentinel value"
        )

        # --- Save then close WITHOUT saving again (close must not re-persist). ---
        saved = structured(await _call(session, "excel_save", {"workbook_id": workbook_id}))
        assert saved["workbook_id"] == workbook_id
        await _call(session, "excel_close", {"workbook_id": workbook_id, "save_changes": False})

    # --- Server stopped: owned PID reaped, no user Excel PID touched (PO-5). ---
    # FE-17: shutdown reaps the owned process itself, so the pid is already gone
    # when the server process exits (checked immediately, no reap barrier).
    assert not _pid_alive(owned_pid), (
        f"owned PID {owned_pid} was still alive when the server exited: "
        "shutdown must reap the owned Excel (FE-17)"
    )
    assert _wait_pid_dead(owned_pid), f"owned PID {owned_pid} was not reaped after shutdown"
    assert owned_pid not in _excel_pids(), f"owned PID {owned_pid} still present after shutdown"
    surviving_user_pids = baseline_pids - _excel_pids()
    assert not surviving_user_pids, (
        f"pre-existing user Excel PID(s) {sorted(surviving_user_pids)} were terminated by the test"
    )

    # --- Originals never modified: repo fixture and temp copy byte-identical (PO-6). ---
    assert _sha256(_WORKFLOW_FIXTURE) == repo_hash_before, "committed fixture must never change"
    assert _sha256(work_copy) == temp_hash_before, "temp copy must never change"


# ---------------------------------------------------------------------------
# FE-17: shutdown reaps the owned Excel with no test-level reap barrier
# ---------------------------------------------------------------------------
@requires_excel
async def test_owned_pid_reaped_on_shutdown_without_barrier(tmp_path: Path) -> None:
    """The owned hidden Excel is dead the moment the server exits (FE-17).

    No reap barrier and no kill helper runs here: the only actor is the server's
    own shutdown path (``ExcelComWorker.shutdown`` -> verified adapter quit ->
    pid-scoped fallback). ``excel_health`` reports the owned pid while the
    server runs; that exact process must be gone as soon as the STDIO server
    subprocess has exited.
    """

    async with _running_server(tmp_path) as session:
        owned_pid = _assert_healthy_and_live(await _wait_healthy(session))
        _assert_owned_hidden(owned_pid, "server start")

    # The server subprocess has exited, so its lifespan shutdown has completed.
    assert not _pid_alive(owned_pid), (
        f"owned PID {owned_pid} survived the server shutdown (FE-17 regression)"
    )
    assert owned_pid not in _excel_pids(), (
        f"owned PID {owned_pid} is still running after shutdown (FE-17 regression)"
    )


# ---------------------------------------------------------------------------
# FE-18 (R-17): excel_health reports a dead owned Excel truthfully
# ---------------------------------------------------------------------------
def _terminate_owned_pid(pid: int) -> bool:
    """Terminate ONLY ``pid`` via ``taskkill /F /PID`` (mirrors macro_service)."""

    try:
        completed = subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return completed.returncode == 0
    except Exception:  # noqa: BLE001 - defensive: a probe failure means "not killed"
        return False


@requires_excel
async def test_health_reports_dead_owned_pid(tmp_path: Path) -> None:
    """A killed owned Excel is reported as ``needs_recovery`` (FE-18 / R-17).

    After ``excel_open`` the worker bookkeeping says ``running`` / ``accepting``.
    Killing the *owned* hidden ``EXCEL.EXE`` (and only that pid) leaves the
    worker bookkeeping unchanged, so without the FE-18 pid-liveness probe
    ``excel_health`` would still report ``ok``. The probe must flip the verdict
    to ``needs_recovery`` truthfully. Auto-recovery is in-process (F2b): the next
    ``excel_open`` self-heals a NEEDS_RECOVERY session with a dead owned pid via an
    in-process ``recover_session`` (no server restart, no ``excel_recover_session``
    tool). This test kills the pid externally (not via a hung macro), so it verifies
    the truthful ``needs_recovery`` verdict; after the server exits the owned pid is
    gone and no orphaned Excel remains beyond the pre-existing baseline.
    """

    baseline_pids = _excel_pids()
    work_copy = _copy_fixture(tmp_path)

    async with _running_server(tmp_path) as session:
        owned_pid = _assert_healthy_and_live(await _wait_healthy(session))
        assert owned_pid not in baseline_pids, (
            f"owned PID {owned_pid} collided with a pre-existing user Excel PID"
        )

        await _call(session, "excel_open", {"path": str(work_copy)})
        # Confirm still healthy before the kill (the owned pid is alive).
        _assert_healthy_and_live(await _wait_healthy(session))

        # Kill ONLY the owned hidden Excel pid (same pattern as macro_service).
        assert _terminate_owned_pid(owned_pid), f"failed to terminate owned pid {owned_pid}"
        assert _wait_pid_dead(owned_pid), f"owned pid {owned_pid} did not die after taskkill"

        # The worker bookkeeping still says running/accepting; the read-only probe
        # must report the stale worker truthfully.
        health = structured(await _call(session, "excel_health"))
        assert health.get("needs_recovery") is True, f"dead owned pid not reported: {health}"
        assert health.get("status") == "needs_recovery", f"wrong status: {health}"
        assert health.get("owned_pid") == owned_pid, f"owned_pid changed: {health}"
        assert any("not alive" in note for note in health.get("notes", [])), (
            f"missing dead-pid note: {health}"
        )

    # Server lifespan shutdown completed: the owned pid is gone (it was killed
    # above and the worker's shutdown reap is a no-op on an already-dead pid).
    assert not _pid_alive(owned_pid), (
        f"owned PID {owned_pid} is unexpectedly alive after shutdown"
    )
    # No orphaned hidden Excel beyond the pre-existing baseline.
    leaked = _excel_pids() - baseline_pids
    assert not leaked, f"orphaned Excel process(es) after shutdown: {sorted(leaked)}"


# ---------------------------------------------------------------------------
# F1 persistence (recommended): snapshot -> close(save) -> restore -> verify
# ---------------------------------------------------------------------------
@requires_excel
async def test_persistence_roundtrip(tmp_path: Path) -> None:
    """Prove edits persist the MCP-native way: snapshot -> close(save) -> restore.

    The plan's literal "close(save) then reopen the copy" is impossible
    (``close_workbook`` always removes the working copy and ``excel_open`` always
    re-copies the untouched original), so persistence is proved through
    ``snapshot_create`` -> ``excel_close(save_changes=True)`` -> ``snapshot_restore``
    -> a fresh ``workbook_id`` whose ``Result!A1`` and table value reflect the edit.
    """

    work_copy = _copy_fixture(tmp_path)

    async with _running_server(tmp_path) as session:
        await _wait_healthy(session)

        opened = structured(await _call(session, "excel_open", {"path": str(work_copy)}))
        workbook_id = opened["workbook_id"]

        await _call(session, "snapshot_create", {"workbook_id": workbook_id})
        classified = structured(
            await _call(
                session,
                "macro_classify_background",
                {
                    "workbook_id": workbook_id,
                    "procedure": "modWorkflow.Recalculate",
                    "background_safe": True,
                },
            )
        )
        assert classified["background_safe"] is True, classified

        await _update_factor(session, workbook_id, 20)
        run = await _run_recalculate(session, workbook_id)
        assert run["status"] == "success", f"Recalculate did not succeed: {run}"
        assert await _read_result_a1(session, workbook_id) == 40

        edited_snapshot = structured(
            await _call(session, "snapshot_create", {"workbook_id": workbook_id})
        )
        edited_snapshot_id = edited_snapshot["snapshot_id"]
        await _call(session, "excel_close", {"workbook_id": workbook_id, "save_changes": True})

        restored = structured(
            await _call(
                session,
                "snapshot_restore",
                {
                    "snapshot_id": edited_snapshot_id,
                    "workbook_id": workbook_id,
                    "confirm": True,
                },
            )
        )
        # R1/R3: the result wraps the restored WorkbookInfo plus previous_close.
        workbook_id2 = restored["workbook"]["workbook_id"]
        assert workbook_id2 and workbook_id2 != workbook_id, f"expected a fresh id: {restored}"

        # FE-16: the restored session is execute-mode and writable, so macros keep
        # working after a restore instead of failing with MACRO_BLOCKED.
        restored_status = structured(
            await _call(session, "excel_get_status", {"workbook_id": workbook_id2})
        )
        assert restored_status["mode"] == "execute", (
            f"restored session must be execute: {restored_status}"
        )
        assert restored_status["read_only"] is False, (
            f"restored session must be writable: {restored_status}"
        )

        # A classify + run cycle succeeds on the restored workbook (the registry is
        # keyed to the new workbook identity/hash, so classify once on the new id).
        classified2 = structured(
            await _call(
                session,
                "macro_classify_background",
                {
                    "workbook_id": workbook_id2,
                    "procedure": "modWorkflow.Recalculate",
                    "background_safe": True,
                },
            )
        )
        assert classified2["background_safe"] is True, classified2
        run2 = await _run_recalculate(session, workbook_id2)
        assert run2["status"] == "success", f"post-restore Recalculate must succeed: {run2}"

        # Observable macro output persisted.
        assert await _read_result_a1(session, workbook_id2) == 40
        # Table cell persisted (SettingValue == 20 lives in Sheet1!B2).
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id2, "sheet": "Sheet1", "address": "B2"},
            )
        ) == 20
        # Macros / components / names persisted.
        components2 = _as_list(
            await _call(session, "vba_list_components", {"workbook_id": workbook_id2})
        )
        assert "modWorkflow" in {c["component_name"] for c in components2}
        macros = _as_list(await _call(session, "macro_list", {"workbook_id": workbook_id2}))
        macro_names = {m["name"] for m in macros}
        assert "modWorkflow.Recalculate" in macro_names, f"Recalculate missing: {macro_names}"
        assert "modWorkflow.ShowMessage" in macro_names, f"ShowMessage missing: {macro_names}"
        names = _as_list(await _call(session, "names_list", {"workbook_id": workbook_id2}))
        assert "SettingsKey" in {n["name"] for n in names}, f"SettingsKey missing: {names}"

        await _call(session, "excel_close", {"workbook_id": workbook_id2, "save_changes": False})


# ---------------------------------------------------------------------------
# Structured denial: read-only session rejects a mutation
# ---------------------------------------------------------------------------
@requires_excel
async def test_read_only_session_denies_mutation(tmp_path: Path) -> None:
    """A read-only session rejects mutations with ``WORKBOOK_READ_ONLY``.

    Covers ``range_write_values``, ``table_update_rows`` and ``macro_run`` (FE-05):
    each is refused by the writable/execute guard before any snapshot or
    background-safety precondition.
    """

    work_copy = _copy_fixture(tmp_path)

    async with _running_server(tmp_path) as session:
        await _wait_healthy(session)

        opened = structured(
            await _call(session, "excel_open", {"path": str(work_copy), "read_only": True})
        )
        assert opened["read_only"] is True, f"expected a read-only session: {opened}"
        workbook_id = opened["workbook_id"]

        denied = await _call(
            session,
            "range_write_values",
            {
                "workbook_id": workbook_id,
                "sheet": "Sheet1",
                "address": "B1",
                "payload": _write_payload("nope"),
            },
        )
        assert error_code(denied) == "WORKBOOK_READ_ONLY"

        # (2) table (ListObject) mutation — denied by the writable guard before the
        # snapshot precondition (FE-05 regression coverage).
        denied_table = await _call(
            session,
            "table_update_rows",
            {
                "workbook_id": workbook_id,
                "name": "tblSettings",
                "updates": [{"index": 1, "values": {"SettingValue": 1}}],
            },
        )
        assert error_code(denied_table) == "WORKBOOK_READ_ONLY"

        # (3) macro execution — denied by _require_execute_context before the
        # background-safety classification (FE-05 regression coverage).
        denied_macro = await _call(
            session,
            "macro_run",
            {"workbook_id": workbook_id, "macro_name": "modWorkflow.Recalculate"},
        )
        assert error_code(denied_macro) == "WORKBOOK_READ_ONLY"


# ---------------------------------------------------------------------------
# Structured denial: restricted-mode server rejects the default execute open
# ---------------------------------------------------------------------------
@requires_excel
async def test_restricted_mode_denies_execute_open(tmp_path: Path) -> None:
    """A restricted-mode server denies the default (execute) ``excel_open``.

    Restricted mode exposes no grant tool over MCP, so the deny-by-default approval
    gate on ``excel_open`` (EXECUTING) is the reachable ``APPROVAL_DENIED`` surface.

    F-4: the filesystem-write tools ``snapshot_create`` and ``vba_export_project``
    are approval-gated too. A read-only inspect open is the only session restricted
    mode admits, and both tools must return ``APPROVAL_DENIED`` on it (previously
    they succeeded ungated).
    """

    work_copy = _copy_fixture(tmp_path)

    async with _running_server(tmp_path, access_mode="restricted") as session:
        await _wait_healthy(session)

        denied = await _call(session, "excel_open", {"path": str(work_copy)})
        assert error_code(denied) == "APPROVAL_DENIED"

        # The only session restricted mode admits: inspect + read-only.
        opened = structured(
            await _call(
                session,
                "excel_open",
                {"path": str(work_copy), "mode": "inspect", "read_only": True},
            )
        )
        assert opened["read_only"] is True, f"expected a read-only session: {opened}"
        workbook_id = opened["workbook_id"]

        denied_snapshot = await _call(
            session, "snapshot_create", {"workbook_id": workbook_id}
        )
        assert error_code(denied_snapshot) == "APPROVAL_DENIED"

        denied_export = await _call(
            session, "vba_export_project", {"workbook_id": workbook_id}
        )
        assert error_code(denied_export) == "APPROVAL_DENIED"
