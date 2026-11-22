"""Real MCP-protocol end-to-end workflow over STDIO (task_b7, fix_007).

This is the first test that drives the built server **through the actual Model
Context Protocol boundary**: it launches ``python -m excel_vba_mcp`` as a child
process, speaks JSON-RPC to it with the official ``mcp`` stdio client
(``stdio_client`` + ``ClientSession``), and asserts the full default-access user
workflow end to end against a live Excel instance:

open a working copy -> mutate settings -> run a macro -> mutate -> run a macro ->
save -> snapshot -> close -> restore -> verify persisted cells / macros / names.

It exercises the production server, services and adapters exactly as a real MCP
client would; nothing is called in-process. The original fixture is never opened:
each test copies ``FullAccessWorkflow.xlsm`` into a temp root and points
``EXCEL_VBA_MCP_ROOTS`` (and work/snapshot/artifact/run/audit) at that root, then
asserts both the repo fixture and the temp copy stay byte-identical.

Coverage:

* **Default full-access workflow** (access mode unset => ``full``): ``excel_open``
  defaults to a writable execute session; ``macro_classify_background`` +
  ``macro_run`` write ``A1``/``A2``; ``range_write_values`` writes cells; the
  interactive ``ShowMessage`` is vetoed (``INTERACTIVE_UI_NOT_SUPPORTED``); an
  unclassified macro is blocked (``BACKGROUND_SAFETY_NOT_VERIFIED``).
* **Persistence (F1 round-trip):** the plan's literal "close(save) then reopen the
  copy" is impossible (``close_workbook`` always removes the working copy and
  ``excel_open`` always re-copies the untouched original), so persistence is proved
  the MCP-native way: ``snapshot_create`` -> ``excel_close(save_changes=True)`` ->
  ``snapshot_restore`` -> a fresh ``workbook_id`` whose cells / components /
  macros / names reflect the edits.
* **Structured denials:** a read-only session rejects a mutation
  (``WORKBOOK_READ_ONLY``); a restricted-mode server rejects the default execute
  open (``APPROVAL_DENIED`` — restricted exposes no grant tool, so ``excel_open``
  is the reachable ``APPROVAL_DENIED`` surface).
* **Tool catalog:** ``list_tools`` advertises the full catalog (>= 87 tools,
  including ``macro_classify_background``).

A **fresh server subprocess** is started per test (never a reused worker), which
avoids the stale-worker / dead-owned-Excel condition (R-17): every test gets a new
owned Excel instance. The server is always torn down through the ``stdio_client``
context manager (stdin close -> bounded wait -> job-object kill).

Run serially (no xdist), against a host with Excel + VBOM trust::

    uv run pytest tests/integration_excel/test_mcp_full_access_workflow.py -q -p no:xdist -s
"""

from __future__ import annotations

import asyncio
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

from mcp import ClientSession, StdioServerParameters, stdio_client
from mcp.types import CallToolResult

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


def _pid_alive(pid: int) -> bool:
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
    """Return the structured payload of a successful tool result."""

    assert not result.is_error, f"expected success, got error result: {result.content!r}"
    if result.structured_content is not None:
        return result.structured_content
    assert result.content, "successful result carried no content"
    return json.loads(result.content[0].text)


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
    code = payload["code"]
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
# Happy path: full default-access workflow end to end
# ---------------------------------------------------------------------------
@requires_excel
async def test_full_access_workflow_e2e(tmp_path: Path) -> None:
    """Drive the full default-access workflow through the real MCP boundary."""

    repo_hash_before = _sha256(_fixture_or_skip(_WORKFLOW_FIXTURE))
    work_copy = _copy_fixture(tmp_path)
    temp_hash_before = _sha256(work_copy)

    async with _running_server(tmp_path) as session:
        # --- Tool catalog (no Excel state required). ---
        tools = (await session.list_tools()).tools
        tool_names = {tool.name for tool in tools}
        assert len(tool_names) >= 87, f"expected >= 87 tools, got {len(tool_names)}"
        assert "macro_classify_background" in tool_names

        # --- Fresh worker: healthy, running, live owned PID. ---
        _assert_healthy_and_live(await _wait_healthy(session))

        # --- Open default (no mode/read_only args): writable execute session. ---
        opened = structured(await _call(session, "excel_open", {"path": str(work_copy)}))
        assert opened["read_only"] is False, f"default open must be writable: {opened}"
        assert opened["mode"] == "execute", f"default open must be execute mode: {opened}"
        workbook_id = opened["workbook_id"]
        assert workbook_id

        status = structured(await _call(session, "excel_get_status", {"workbook_id": workbook_id}))
        assert status["mode"] == "execute", f"session mode: {status}"
        assert status["read_only"] is False, f"session read_only: {status}"

        # --- VBA is readable through the protocol. ---
        components = _as_list(
            await _call(session, "vba_list_components", {"workbook_id": workbook_id})
        )
        component_names = {c["component_name"] for c in components}
        assert "modWorkflow" in component_names, f"modWorkflow missing: {component_names}"

        source = structured(
            await _call(
                session,
                "vba_read_component",
                {"workbook_id": workbook_id, "component": "modWorkflow"},
            )
        )
        assert "WriteCell" in source["source"], "modWorkflow source missing WriteCell"
        assert "ShowMessage" in source["source"], "modWorkflow source missing ShowMessage"

        # --- Interactive procedure cannot be classified background_safe (R13 veto). ---
        vetoed = await _call(
            session,
            "macro_classify_background",
            {"workbook_id": workbook_id, "procedure": "ShowMessage", "background_safe": True},
        )
        assert error_code(vetoed) == "INTERACTIVE_UI_NOT_SUPPORTED"

        # --- Unclassified macro is blocked, never run (BG06 deny-by-default). ---
        unclassified = macro_result(
            await _call(
                session,
                "macro_run",
                {"workbook_id": workbook_id, "macro_name": "WriteCell"},
            )
        )
        assert unclassified["status"] == "blocked", f"unclassified macro must block: {unclassified}"
        assert unclassified["error_code"] == "BACKGROUND_SAFETY_NOT_VERIFIED", unclassified

        # --- Classify WriteCell, then run it -> A1 == 42. ---
        classified = structured(
            await _call(
                session,
                "macro_classify_background",
                {"workbook_id": workbook_id, "procedure": "WriteCell", "background_safe": True},
            )
        )
        assert classified["background_safe"] is True, classified
        run1 = macro_result(
            await _call(
                session,
                "macro_run",
                {"workbook_id": workbook_id, "macro_name": "WriteCell"},
            )
        )
        assert run1["status"] == "success", f"WriteCell did not succeed: {run1}"
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id, "sheet": "Sheet1", "address": "A1"},
            )
        ) == 42

        # --- Tool write B1 (guarded mutation: needs a snapshot first). ---
        structured(await _call(session, "snapshot_create", {"workbook_id": workbook_id}))
        write1 = structured(
            await _call(
                session,
                "range_write_values",
                {
                    "workbook_id": workbook_id,
                    "sheet": "Sheet1",
                    "address": "B1",
                    "payload": _write_payload("tool-write"),
                },
            )
        )
        assert write1["verified"] is True, f"range_write_values not verified: {write1}"
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id, "sheet": "Sheet1", "address": "B1"},
            )
        ) == "tool-write"

        # --- Classify + run WriteCell2 -> A2 == 7. ---
        classified2 = structured(
            await _call(
                session,
                "macro_classify_background",
                {"workbook_id": workbook_id, "procedure": "WriteCell2", "background_safe": True},
            )
        )
        assert classified2["background_safe"] is True, classified2
        run2 = macro_result(
            await _call(
                session,
                "macro_run",
                {"workbook_id": workbook_id, "macro_name": "WriteCell2"},
            )
        )
        assert run2["status"] == "success", f"WriteCell2 did not succeed: {run2}"
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id, "sheet": "Sheet1", "address": "A2"},
            )
        ) == 7

        # --- Another guarded tool write (C1) then explicit save. ---
        write2 = structured(
            await _call(
                session,
                "range_write_values",
                {
                    "workbook_id": workbook_id,
                    "sheet": "Sheet1",
                    "address": "C1",
                    "payload": _write_payload("c-value"),
                },
            )
        )
        assert write2["verified"] is True, f"C1 write not verified: {write2}"
        saved = structured(await _call(session, "excel_save", {"workbook_id": workbook_id}))
        assert saved["workbook_id"] == workbook_id

        # --- Persistence via the F1 round-trip: snapshot -> close -> restore. ---
        edited_snapshot = structured(
            await _call(session, "snapshot_create", {"workbook_id": workbook_id})
        )
        edited_snapshot_id = edited_snapshot["snapshot_id"]
        await _call(session, "excel_close", {"workbook_id": workbook_id, "save_changes": True})

        restored = structured(
            await _call(session, "snapshot_restore", {"snapshot_id": edited_snapshot_id})
        )
        workbook_id2 = restored["workbook_id"]
        assert workbook_id2 and workbook_id2 != workbook_id, f"expected a fresh id: {restored}"

        # Cells persisted.
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id2, "sheet": "Sheet1", "address": "A1"},
            )
        ) == 42
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id2, "sheet": "Sheet1", "address": "B1"},
            )
        ) == "tool-write"
        assert _single_cell(
            await _call(
                session,
                "range_read_values",
                {"workbook_id": workbook_id2, "sheet": "Sheet1", "address": "A2"},
            )
        ) == 7

        # Macros / components / names persisted.
        components2 = _as_list(
            await _call(session, "vba_list_components", {"workbook_id": workbook_id2})
        )
        assert "modWorkflow" in {c["component_name"] for c in components2}

        macros = _as_list(await _call(session, "macro_list", {"workbook_id": workbook_id2}))
        macro_names = {m["name"] for m in macros}
        assert "modWorkflow.WriteCell" in macro_names, f"WriteCell missing: {macro_names}"
        assert "modWorkflow.WriteCell2" in macro_names, f"WriteCell2 missing: {macro_names}"

        names = _as_list(await _call(session, "names_list", {"workbook_id": workbook_id2}))
        assert "SettingsKey" in {n["name"] for n in names}, f"SettingsKey missing: {names}"

        await _call(session, "excel_close", {"workbook_id": workbook_id2, "save_changes": False})

    # --- Originals never modified: repo fixture and temp copy byte-identical. ---
    assert _sha256(_WORKFLOW_FIXTURE) == repo_hash_before, "committed fixture must never change"
    assert _sha256(work_copy) == temp_hash_before, "temp copy must never change"


# ---------------------------------------------------------------------------
# Structured denial: read-only session rejects a mutation
# ---------------------------------------------------------------------------
@requires_excel
async def test_read_only_session_denies_mutation(tmp_path: Path) -> None:
    """A read-only session rejects ``range_write_values`` with ``WORKBOOK_READ_ONLY``."""

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


# ---------------------------------------------------------------------------
# Structured denial: restricted-mode server rejects the default execute open
# ---------------------------------------------------------------------------
@requires_excel
async def test_restricted_mode_denies_execute_open(tmp_path: Path) -> None:
    """A restricted-mode server denies the default (execute) ``excel_open``.

    Restricted mode exposes no grant tool over MCP, so the deny-by-default approval
    gate on ``excel_open`` (EXECUTING) is the reachable ``APPROVAL_DENIED`` surface.
    """

    work_copy = _copy_fixture(tmp_path)

    async with _running_server(tmp_path, access_mode="restricted") as session:
        await _wait_healthy(session)

        denied = await _call(session, "excel_open", {"path": str(work_copy)})
        assert error_code(denied) == "APPROVAL_DENIED"