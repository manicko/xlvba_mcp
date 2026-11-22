"""Structured-error surfacing matrix for the non-read-only MCP tools (FE-02).

The auditor finding (G-9 / F-5) is that the MCP server maps only
:class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError` (found in the
``__cause__`` chain of an SDK ``ToolError``) into a structured
``ToolError(code+context)`` payload at :meth:`ExcelVbaMcpServer.call_tool`.
Any **non** domain exception — a pydantic ``ValidationError``, ``KeyError``,
bare ``ValueError`` / ``TypeError`` — reaches the client as an SDK
``UnexpectedToolError("Error executing tool <name>")`` with **no** ``code`` and
no ``context``, so a client cannot distinguish a caller error from Excel
contention.

This test drives a real STDIO server (the same harness pattern as
``test_mcp_full_access_workflow.py``: ``stdio_client`` + ``ClientSession``, env
roots scoped to a temp directory, the fixture copied into that temp root so the
committed fixture is never touched) and calls a representative set of **mutating
/ executing / admin** tools with:

* a **denial** argument — e.g. a read-only session where a writable is needed,
  an unknown ``workbook_id``, an unclassified macro, a missing snapshot / id;
* a **malformed** argument — a missing required field or a wrong-typed field.

For every call the contract is the same: the result MUST be a structured error
(``result.is_error is True`` and the JSON body carries a non-empty string
``code`` and a ``context`` object). The key assertion is that **no** tool returns
an unstructured ``UnexpectedToolError`` / "Error executing tool" body without a
``code``.

Run serially (no xdist), against a host with Excel + VBOM trust::

    .venv\\Scripts\\python.exe -m pytest \\
        tests/integration_excel/test_mcp_error_surfacing_matrix.py -q -p no:xdist
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import sys
import time
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

import pytest

pytest.importorskip("win32com", reason="pywin32 required for the live Excel error-surfacing test")
pytest.importorskip("pythoncom", reason="pywin32 required for the live Excel error-surfacing test")

from mcp import ClientSession, StdioServerParameters, stdio_client  # noqa: E402
from mcp.shared.exceptions import MCPError  # noqa: E402
from mcp.types import CallToolResult  # noqa: E402

_PYWIN32 = sys.platform == "win32"

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_WORKFLOW_FIXTURE = _REPO_ROOT / "tests" / "fixtures" / "FullAccessWorkflow.xlsm"

# Per-request ceiling for a single MCP call.
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


# ---------------------------------------------------------------------------
# Server harness (mirrors test_mcp_full_access_workflow.py)
# ---------------------------------------------------------------------------
# Bounded poll for the owned Excel PID to be reaped after server shutdown.
_REAP_POLL_TIMEOUT_S = 30.0
_REAP_POLL_INTERVAL_S = 0.5


def _excel_pids() -> set[int]:
    """Return the PIDs of currently running ``EXCEL.EXE`` processes."""

    if not _PYWIN32:
        return set()
    import subprocess

    completed = subprocess.run(  # noqa: S603 - fixed argv, no shell
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        check=False,
    )
    pids: set[int] = set()
    for line in completed.stdout.splitlines():
        parts = [part.strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE" and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _wait_pid_dead(pid: int, *, timeout_s: float = _REAP_POLL_TIMEOUT_S) -> bool:
    """Poll until ``pid`` is gone (bounded); return whether it was reaped."""

    deadline = time.monotonic() + timeout_s
    while time.monotonic() < deadline:
        if pid not in _excel_pids():
            return True
        time.sleep(_REAP_POLL_INTERVAL_S)
    return pid not in _excel_pids()


def server_env(base: Path, *, access_mode: str | None = None) -> dict[str, str]:
    """Environment for a fresh MCP server subprocess, scoped to ``base``."""

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
async def _running_server(base: Path) -> AsyncIterator[ClientSession]:
    """Start a fresh ``excel_vba_mcp`` STDIO server and yield an MCP session.

    On exit, after the server subprocess stops, wait (bounded) for the owned
    hidden Excel PID to be reaped. Without this barrier the worker's Excel can
    outlive the server briefly, and a later test that asserts "pre-existing user
    Excel PIDs survive" would misclassify that lingering owned PID as a user
    instance (test-isolation defect observed when this file runs before
    ``test_mcp_full_access_workflow.py``).
    """

    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "excel_vba_mcp"],
        cwd=str(base),
        env=server_env(base),
    )
    baseline_pids = _excel_pids()
    async with (
        stdio_client(params, errlog=sys.stderr) as (read_stream, write_stream),
        ClientSession(read_stream, write_stream, read_timeout_seconds=_CALL_TIMEOUT_S) as session,
    ):
        await session.initialize()
        yield session
    # Server stopped: reap any Excel this test's worker owned, and do not touch
    # PIDs that pre-existed this test (those belong to the user / other tests).
    for pid in _excel_pids() - baseline_pids:
        _wait_pid_dead(pid)


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


def _single_cell(result: CallToolResult) -> Any:
    data = structured(result)
    rows = data["values"]
    assert len(rows) == 1 and len(rows[0]) == 1, (
        f"expected a 1x1 read, got {data['rows']}x{data['cols']}"
    )
    return rows[0][0]["value"]


def _copy_fixture(base: Path, *, name: str | None = None) -> Path:
    """Copy the committed fixture into ``base/input`` (the original stays untouched).

    A distinct ``name`` yields a distinct input path so a second session can open
    the same content (the server rejects two sessions on one canonical path; the
    server always works on its own copy under ``work/``, so the input file is only
    a seed).
    """

    fixture = _fixture_or_skip(_WORKFLOW_FIXTURE)
    input_dir = base / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    target = input_dir / (name or fixture.name)
    target.write_bytes(fixture.read_bytes())
    return target


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


# ---------------------------------------------------------------------------
# Structured-error assertions (the contract under test)
# ---------------------------------------------------------------------------
def assert_structured_error(result: CallToolResult, tool: str, kind: str) -> dict[str, Any]:
    """Assert the result is a structured error carrying a stable ``code``.

    The payload MUST be a JSON object with a non-empty string ``code`` (one of the
    domain ``ErrorCode`` values) and a ``context`` object. A body whose text is the
    generic SDK "Error executing tool ..." (no JSON / no ``code``) is the
    unstructured leak this task guards against.
    """

    assert result.is_error, (
        f"[{tool}/{kind}] expected a structured error, got a success result: {result.content!r}"
    )
    assert result.content, f"[{tool}/{kind}] error result carried no content"

    text = result.content[0].text
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:  # the SDK's generic "Error executing tool" body
        pytest.fail(
            f"[{tool}/{kind}] UNSTRUCTURED error (no JSON payload, no code): {text!r} ({exc})"
        )
    assert isinstance(payload, dict), f"[{tool}/{kind}] payload is not a JSON object: {payload!r}"

    code = payload.get("code")
    assert isinstance(code, str) and code.strip(), (
        f"[{tool}/{kind}] structured error is missing a non-empty 'code': {payload!r}"
    )
    # A real ExcelVbaMcpError.to_dict() always includes a context dict.
    assert isinstance(payload.get("context"), dict), (
        f"[{tool}/{kind}] structured error is missing a 'context' object: {payload!r}"
    )
    return payload


# ---------------------------------------------------------------------------
# Argument matrix
# ---------------------------------------------------------------------------
def _write_payload(value: Any) -> dict[str, Any]:
    kind = "number" if isinstance(value, (int, float)) and not isinstance(value, bool) else "string"
    return {"rows": 1, "cols": 1, "values": [[{"kind": kind, "value": value}]]}


def _denial_args(tool: str, wid: str) -> dict[str, Any]:
    """Arguments that exercise a real domain denial for ``tool`` on a read-only session.

    ``wid`` is a read-only ``workbook_id``; the guard chain (writable check) rejects
    every mutation with a domain error before any snapshot/table/object lookup.
    """

    table = {"workbook_id": wid, "name": "tblSettings"}
    return {
        "range_write_values": {
            "workbook_id": wid,
            "sheet": "Sheet1",
            "address": "A1",
            "payload": _write_payload(1),
        },
        "table_update_rows": {**table, "updates": [{"index": 1, "values": {"SettingValue": 1}}]},
        "table_create": {"workbook_id": wid, "sheet": "Sheet1", "range_or_name": "A1:B2"},
        "name_create": {
            "workbook_id": wid,
            "name": "fe02Name",
            "refers_to": "=1",
            "scope": "workbook",
        },
        "sheet_add": {"workbook_id": wid, "name": "fe02Sheet"},
        "sheet_delete": {"workbook_id": wid, "name": "Sheet1"},
        "shape_set_on_action": {
            "workbook_id": wid,
            "name": "fe02Shape",
            "macro_name": "modWorkflow.Recalculate",
        },
        "control_set_on_action": {
            "workbook_id": wid,
            "name": "fe02Control",
            "macro_name": "modWorkflow.Recalculate",
        },
        "vba_replace_component_source": {
            "workbook_id": wid,
            "component": "modWorkflow",
            "new_source": 'Attribute VB_Name = "modWorkflow"\r\n',
            "expected_source_hash": "deadbeef",
        },
        "vba_insert_code": {
            "workbook_id": wid,
            "component": "modWorkflow",
            "code": "Sub X()\r\nEnd Sub\r\n",
            "position": "declarations",
            "expected_source_hash": "deadbeef",
        },
        "macro_run": {"workbook_id": wid, "macro_name": "modWorkflow.Recalculate"},
        "snapshot_create": {"workbook_id": wid},
        "snapshot_restore": {"snapshot_id": "no-such-snapshot"},
        "snapshot_diff": {"snapshot_id_a": "no-such-a", "snapshot_id_b": "no-such-b"},
        "test_run_suite": {
            "suite": {
                "suite": "fe02",
                "workbook_id": wid,
                "cases": [
                    {
                        "id": "c1",
                        "steps": [
                            {
                                "action": "range_read_values",
                                "sheet": "Sheet1",
                                "range_address": "A1",
                            }
                        ],
                        "assertions": [
                            {
                                "type": "range_structure",
                                "sheet": "Sheet1",
                                "address": "A1",
                                "rows": 1,
                                "cols": 1,
                            }
                        ],
                    }
                ],
            }
        },
        "vba_export_project": {"workbook_id": wid},
        "excel_save": {"workbook_id": wid},
        "excel_close": {"workbook_id": wid, "save_changes": False},
    }


def _malformed_args(tool: str, wid: str) -> dict[str, Any]:
    """A malformed argument set for ``tool`` (missing required field / wrong type).

    The malformed shape must still be rejected with a structured ``code``. Most
    tools validate their own arguments (raising a domain error); a wrong-typed
    payload also exercises the pydantic ``ValidationError`` boundary that this
    task specifically guards.
    """

    table = {"workbook_id": wid, "name": "tblSettings"}
    return {
        # No 'sheet'/'address'/'payload': exercises pydantic required-field validation.
        "range_write_values": {"workbook_id": wid},
        "table_update_rows": {**table, "updates": "not-a-list"},
        "table_create": {**table},  # missing 'sheet' and 'range_or_name'
        "name_create": {**table},  # missing 'name'/'refers_to'/'scope'
        "sheet_add": {"workbook_id": wid},  # missing 'name'
        "sheet_delete": {"workbook_id": wid},  # missing 'name'
        "shape_set_on_action": {"workbook_id": wid},  # missing 'name'/'macro_name'
        "control_set_on_action": {"workbook_id": wid},  # missing 'name'/'macro_name'
        "vba_replace_component_source": {"workbook_id": wid},  # missing the rest
        "vba_insert_code": {"workbook_id": wid},  # missing the rest
        "macro_run": {"workbook_id": wid},  # missing 'macro_name'
        "snapshot_create": {"workbook_id": 12345},  # wrong type for 'workbook_id'
        "snapshot_restore": {},  # missing 'snapshot_id'
        "snapshot_diff": {"snapshot_id_a": "a"},  # missing 'snapshot_id_b'
        "test_run_suite": {},  # missing 'suite'
        "vba_export_project": {"workbook_id": 42},  # wrong type for 'workbook_id'
        "excel_save": {},  # missing 'workbook_id'
        "excel_close": {},  # missing 'workbook_id'
    }


_MUTATING_TOOLS: tuple[str, ...] = (
    "range_write_values",
    "table_update_rows",
    "table_create",
    "name_create",
    "sheet_add",
    "sheet_delete",
    "shape_set_on_action",
    "control_set_on_action",
    "vba_replace_component_source",
    "vba_insert_code",
    "macro_run",
    "snapshot_create",
    "snapshot_restore",
    "snapshot_diff",
    "test_run_suite",
    "vba_export_project",
    "excel_save",
    "excel_close",
)


# ---------------------------------------------------------------------------
# The matrix test
# ---------------------------------------------------------------------------
@requires_excel
async def test_error_surfacing_matrix(tmp_path: Path) -> None:
    """Every representative non-read-only tool surfaces denial + malformed as structured errors."""

    repo_hash_before = _sha256(_fixture_or_skip(_WORKFLOW_FIXTURE))
    work_copy = _copy_fixture(tmp_path)
    ro_work_copy = _copy_fixture(tmp_path, name="FullAccessWorkflow_ro.xlsm")
    temp_hash_before = _sha256(work_copy)
    ro_hash_before = _sha256(ro_work_copy)

    report: list[str] = []

    async with _running_server(tmp_path) as session:
        await _wait_healthy(session)

        # A writable execute session (default open).
        opened = structured(await _call(session, "excel_open", {"path": str(work_copy)}))
        assert opened["read_only"] is False and opened["mode"] == "execute", (
            f"default open must be writable execute: {opened}"
        )
        write_wid = opened["workbook_id"]

        # A read-only session for the denial (writable-required) arguments.
        ro_opened = structured(
            await _call(session, "excel_open", {"path": str(ro_work_copy), "read_only": True})
        )
        assert ro_opened["read_only"] is True, f"expected a read-only session: {ro_opened}"
        read_only_wid = ro_opened["workbook_id"]

        # The denial argument must hit the REAL guard for each tool, not merely a
        # schema rejection. ``_denial_expectations`` maps each mutating tool to the
        # structured code its denial must produce; tools with no reachable denial on
        # a read-only session (excel_save / excel_close are read-only-safe) map to
        # None and are asserted to SUCCEED instead, so a silent behavior change
        # still fails. ``snapshot_create`` / ``vba_export_project`` are approval-
        # gated filesystem writes (F-4) but auto-allowed in this full-mode server,
        # so they still succeed here; their restricted-mode denial is covered by
        # ``test_restricted_mode_denies_execute_open``.
        _denial_expectations: dict[str, str | None] = {
            "range_write_values": "WORKBOOK_READ_ONLY",
            "table_update_rows": "WORKBOOK_READ_ONLY",
            "table_create": "WORKBOOK_READ_ONLY",
            "name_create": "WORKBOOK_READ_ONLY",
            "sheet_add": "WORKBOOK_READ_ONLY",
            "sheet_delete": "WORKBOOK_READ_ONLY",
            "shape_set_on_action": "WORKBOOK_READ_ONLY",
            "control_set_on_action": "WORKBOOK_READ_ONLY",
            "vba_replace_component_source": "WORKBOOK_READ_ONLY",
            "vba_insert_code": "WORKBOOK_READ_ONLY",
            "macro_run": "WORKBOOK_READ_ONLY",
            # snapshot_restore/diff fail on the unknown snapshot id, not read-only.
            "snapshot_restore": "SNAPSHOT_FAILED",
            "snapshot_diff": "SNAPSHOT_FAILED",
            # test_run_suite is read-only-safe (runs against disposable copies and
            # never mutates the session workbook); it succeeds on a read-only open.
            "test_run_suite": None,
            # Approval-gated filesystem writes (F-4): auto-allowed in full mode, so
            # they succeed on a read-only session here (restricted-mode denial is
            # asserted elsewhere).
            "snapshot_create": None,
            "vba_export_project": None,
            "excel_save": None,
            "excel_close": None,
        }

        for tool in _MUTATING_TOOLS:
            expected = _denial_expectations[tool]
            denial = await _call(session, tool, _denial_args(tool, read_only_wid)[tool])
            if expected is None:
                # Ungated / read-only-safe tool: the denial argument must succeed.
                payload = structured(denial)
                report.append(f"{tool:32s} denial   OK (ungated, no denial expected)")
                assert payload is not None, f"[{tool}] expected success, got {denial.content!r}"
            else:
                payload = assert_structured_error(denial, tool, f"denial[{expected}]")
                assert payload["code"] == expected, (
                    f"[{tool}] denial hit the wrong guard: expected {expected}, "
                    f"got {payload['code']} (payload={payload!r})"
                )
                report.append(f"{tool:32s} denial   code={payload['code']} (expected {expected})")

            # (b) malformed argument (missing required field / wrong type)
            malformed = await _call(session, tool, _malformed_args(tool, write_wid)[tool])
            malformed_payload = assert_structured_error(malformed, tool, "malformed")
            report.append(f"{tool:32s} malformed code={malformed_payload['code']}")

        # (c) model_validator-guarded argument (audit finding C-1 regression).
        # ``RangeWriteRequest`` raises a plain ``ValueError`` from its
        # ``model_validator`` when the grid exceeds ``MAX_RANGE_ROWS``. pydantic
        # embeds that live exception in ``errors()[i]['ctx']['error']``, which is
        # NOT JSON-serializable; the boundary must sanitize it so the client still
        # receives a structured ``code`` instead of a ``TypeError`` leak.
        oversized = await _call(
            session,
            "range_write_values",
            {
                "workbook_id": write_wid,
                "sheet": "Sheet1",
                "address": "A1",
                "payload": {
                    "rows": 100_001,
                    "cols": 1,
                    "values": [[{"kind": "number", "value": 1}]],
                },
            },
        )
        oversized_payload = assert_structured_error(
            oversized, "range_write_values", "model-validator"
        )
        assert oversized_payload["code"] == "INVALID_RANGE", (
            f"oversized grid must be a structured INVALID_RANGE, got {oversized_payload['code']!r}"
        )
        report.append(
            f"{'range_write_values':32s} validator code={oversized_payload['code']} "
            "(json-serializable ctx)"
        )

        await _call(session, "excel_close", {"workbook_id": write_wid, "save_changes": False})
        await _call(session, "excel_close", {"workbook_id": read_only_wid, "save_changes": False})

        # (d) Unknown workbook_id is a caller error, not COM contention (FE-22 / F-5).
        # Read-only tools that resolve the workbook record first must now report
        # WORKBOOK_NOT_FOUND instead of the former generic COM_BUSY.
        bogus = "__no_such_workbook__"
        unknown_id_cases: dict[str, dict[str, Any]] = {
            "vba_read_component": {"workbook_id": bogus, "component": "modWorkflow"},
            "names_list": {"workbook_id": bogus},
            "pivot_tables_list": {"workbook_id": bogus},
            "sheet_get": {"workbook_id": bogus, "name": "Sheet1"},
            "table_get": {"workbook_id": bogus, "name": "tblSettings"},
        }
        for tool, args in unknown_id_cases.items():
            unknown = await _call(session, tool, args)
            payload = assert_structured_error(unknown, tool, f"unknown-id[{bogus}]")
            assert payload["code"] == "WORKBOOK_NOT_FOUND", (
                f"[{tool}] unknown workbook_id must be WORKBOOK_NOT_FOUND, "
                f"got {payload['code']} (payload={payload!r})"
            )
            report.append(
                f"{tool:32s} unknown-id code={payload['code']} (expected WORKBOOK_NOT_FOUND)"
            )

        # (e) Resource errors are structured too (FE-22 / F-5 §6.4): reading a VBA
        # component source for an unknown workbook surfaces the JSON payload
        # (code + context) over the wire, not a plain-text message. The server
        # wraps the ResourceError in a JSON-RPC MCPError whose message is that JSON.
        with pytest.raises(MCPError) as res_exc:
            await session.read_resource(f"excel://workbooks/{bogus}/vba/components/modWorkflow")
        assert "WORKBOOK_NOT_FOUND" in str(res_exc.value), (
            "resource error must carry the structured WORKBOOK_NOT_FOUND code, "
            f"got {res_exc.value!r}"
        )
        report.append("resource(component)           unknown-id structured code=WORKBOOK_NOT_FOUND")

    # Originals never modified: repo fixture and temp copies byte-identical.
    assert _sha256(_WORKFLOW_FIXTURE) == repo_hash_before, "committed fixture must never change"
    assert _sha256(work_copy) == temp_hash_before, "temp copy must never change"
    assert _sha256(ro_work_copy) == ro_hash_before, "read-only temp copy must never change"

    # Surface the per-tool result table on failure / verbose runs.
    print("\nFE-02 structured-error surfacing matrix:")
    for line in report:
        print("  " + line)
