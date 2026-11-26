---
id: xlvba-troubleshooting
domain: 10-troubleshooting
tags:
  - troubleshooting
  - errors
  - diagnostics
  - excel
  - com
  - vba
related:
  - xlvba-overview
  - xlvba-installation-and-setup
  - xlvba-configuration
  - xlvba-architecture
  - xlvba-mcp-tools-reference
  - xlvba-excel-capabilities
  - xlvba-workflows-and-examples
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-development-and-maintenance
---

## Purpose

This document maps the problems you are most likely to hit with `xlvba` to
their causes, diagnostics and fixes. It focuses on the errors a client actually
sees: structured tool errors with a stable `code`, startup failures, and Excel
state problems.

For the error contract see [Architecture](04-architecture.md) and
[MCP tools reference](05-mcp-tools-reference.md). For configuration see
[Configuration](03-configuration.md).

## Problem / error reference

| Problem / Error | Possible cause | How to diagnose | Solution |
|-----------------|----------------|-----------------|----------|
| MCP client cannot connect | The server is not spawned, the command path is wrong, or `uv` is not on `PATH` | Check the client's MCP configuration; run `uv run excel-vba-mcp` manually in a terminal; look for a non-zero exit | Fix the `command` / `args` (use an absolute `--directory`); confirm `uv sync --locked` succeeded; restart the client |
| Server exits at startup | Invalid `EXCEL_VBA_MCP_ACCESS_MODE`, `EXCEL_VBA_MCP_LOG_LEVEL` or `EXCEL_VBA_MCP_MACRO_TIMEOUT_S`; a `ValidationError` is raised in `lifespan` | Read the `stderr` line `Invalid server configuration: ...`; check the three variables | Set `EXCEL_VBA_MCP_ACCESS_MODE` to `full` or `restricted`; set `EXCEL_VBA_MCP_LOG_LEVEL` to `DEBUG`/`INFO`/`WARNING`/`ERROR`/`CRITICAL` (or unset it); set `EXCEL_VBA_MCP_MACRO_TIMEOUT_S` to a positive number (or unset it) |
| Missing Python deps / `uv sync` failure | `.venv` is stale, the lock is out of date, or a wheel is unavailable for the Python version | Run `uv sync --locked`; if it fails because `pyproject.toml` changed, run `uv lock` then `uv sync` | Rebuild the environment from the lock; use a supported Python (3.11–3.14 x64) |
| Excel not starting / not responding | Desktop Excel is not installed, not activated, or the UWP/Store version is in use | `excel_health` shows `unavailable` / `needs_recovery`; the worker startup warning is on `stderr`; `server_health` still returns `ok` | Install and activate desktop Excel; launch it manually once; close other Excel instances if they interfere; retry `excel_open` |
| `EXCEL.exe` process leak | The owned hidden instance did not die on shutdown, or a macro/modal kept it resident | Check `excel_health.owned_pid`; inspect Task Manager for an orphaned `EXCEL.EXE` | The Job Object (`KILL_ON_JOB_CLOSE`) auto-reaps the owned Excel when the server dies, and a normal shutdown quits/reaps the owned pid. If a process survives (legacy server without the Job Object), verify the pid's identity (create-time + `EXCEL.EXE` image) and only then `Stop-Process -Id <pid>`, then restart the server. Do not kill user Excel processes; never kill on pid-alone |
| COM errors (`COM_BUSY`, `RPC_E_CALL_REJECTED`) | Excel/VBE is busy; a call was rejected while the STA was occupied | The error `code` is `COM_BUSY` or `RPC_E_CALL_REJECTED`; `context.class` shows the operation class | Retry a **read** after a short pause (reads are retried up to 3 times). Do not blindly retry a mutation; wait and check `excel_health` |
| `PATH_NOT_ALLOWED` | Path is outside `EXCEL_VBA_MCP_ROOTS`, or a device/UNC/traversal/symlink path, or a disallowed extension | Read `context.reason`: `empty_path`, `device_path_denied`, `unc_denied`, `traversal`, `outside_roots`, `symlink_escape`, `root_mismatch`, `extension_not_allowed`; `context.allowed_roots` lists the roots the running server loaded. For `vba_export_project` an out-of-root `out_dir` reports the **absolute** artifacts root in `context.allowed_root` / `context.expected_root` (F4) | Put the workbook inside an allowed root or extend `EXCEL_VBA_MCP_ROOTS`; use a plain path with an allowed extension. `EXCEL_VBA_MCP_ROOTS` is read once at startup — **restart/reload the MCP server** after changing it (see [Configuration](03-configuration.md)) |
| `APPROVAL_DENIED` | The server runs in `restricted` mode and no allowlist entry or live grant matches | The error `code` is `APPROVAL_DENIED`; `excel_get_status.access_mode` reports `restricted` | Unset `EXCEL_VBA_MCP_ACCESS_MODE` (or set `full`) for full access; a restricted server exposes only read-only work over MCP |
| `WORKBOOK_READ_ONLY` | The session was opened with `read_only=true` | Any mutating tool returns `WORKBOOK_READ_ONLY`; `excel_get_status.read_only` is `true` | Reopen writable (`excel_open` without `read_only`) or do read-only work only |
| `ACCESS_VBOM_DISABLED` | "Trust access to the VBA project object model" is off, or corporate policy blocks it | VBA/VBIDE tools and `vba_compile_diagnostics` return `ACCESS_VBOM_DISABLED`; `workbook_get_model` reports `vba_access = "vbom_disabled"` | Enable VBOM in Excel Trust Center → Macro Settings, restart Excel, retry. If policy blocks it, VBA features stay unavailable |
| `VBA_PROJECT_LOCKED` | The VBA project is password-protected | VBA tools return `VBA_PROJECT_LOCKED`; `vba_project_info` reports the protection state | Unlock the project manually in the VBE (the owner must supply the password); protection is never bypassed |
| `ACCESS_VBOM_DISABLED` vs locked project | Two different causes: missing trust vs a protected project | `ACCESS_VBOM_DISABLED` means the VBIDE is not trusted; `VBA_PROJECT_LOCKED` means it is trusted but the project is protected | Enable VBOM for the first; unlock the project for the second. Do not confuse the two |
| Macro failures (`MACRO_NOT_FOUND`, `MACRO_BLOCKED`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`) | Wrong name; Office macro policy; a runtime error in the macro; the macro exceeded its soft deadline (default 90 s, `EXCEL_VBA_MCP_MACRO_TIMEOUT_S`) | `macro_get_last_result` returns the status and structured error; `context` carries the HRESULT/section; `session_recovery` reports the in-process recovery outcome | `MACRO_NOT_FOUND`: fix `Module.Procedure`; `MACRO_BLOCKED`: check Office macro settings / execute mode; `MACRO_RUNTIME_ERROR`: inspect the macro (a modal was terminated, `dismissed=false`); `MACRO_TIMEOUT`: simplify the macro or raise no expectation of completion. After either timeout form the owned Excel PID was terminated and the session recovered automatically (F2b) — **re-open the workbook** (the old `workbook_id` is invalid) and check `session_recovery` |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | The procedure is not classified `background_safe` for the current workbook hash and VBA source revision | `macro_run` returns `blocked` with `BACKGROUND_SAFETY_NOT_VERIFIED`; the macro was not started | Classify it with `macro_classify_background` after reviewing the code; re-classify after any VBA source change; do not retry blindly |
| `INTERACTIVE_UI_NOT_SUPPORTED` | The procedure hits the R13 token scan, is heuristic, or was classified safe while unsafe | `macro_run` / `macro_classify_background` returns `INTERACTIVE_UI_NOT_SUPPORTED`; `context.flagged_tokens` lists the hits | Remove the interactive UI or leave the procedure unclassified; do not bypass the preflight |
| `EXCEL_HUNG` / `needs_recovery` | A call exceeded its soft deadline, Excel is hung, or a modal dialog is blocking COM | `excel_health` shows `status = "needs_recovery"` (or `worker_state = "running/deadline_exceeded"`) and `needs_recovery = true`; the error `code` is `EXCEL_HUNG` with `context.verdict = "SUSPECTED_STALLED"` | A `macro_run` hang/timeout recovers the session **automatically** (F2b) — re-open the workbook. For any other hung/deadline-exceeded call, call `excel_recover` (agent-directed; the server never auto-kills a live pid on a deadline). If `excel_recover` is refused (`recovery_possible=false`), restart the server process. Only the owned Excel pid is ever terminated |
| `DEADLINE_EXCEEDED` / `SUSPECTED_STALLED` | A single COM call started but exceeded its execution deadline | `excel_health.worker_state = "running/deadline_exceeded"`; the error is `EXCEL_HUNG` with `context.deadline_exceeded=true`, `context.verdict="SUSPECTED_STALLED"`. The call is **not** cancelled and the pid is **not** terminated | Agent-directed: call `excel_recover` when the state persists. The worker resumes `running` on its own when the slow call returns, so a transient deadline-exceeded is not a poison state — re-check `excel_health` before recovering |
| `interrupted_unknown` job | A background job was left non-terminal by a prior process (restart / crash), or finalized after a successful `excel_recover` | `job_get` returns `status="interrupted_unknown"`, `outcome="unknown"`, `retryable=false` | **Do not replay.** The job is never completed and never auto-retried; re-issue the operation as a new job if still needed |
| Orphaned `EXCEL.EXE` | The server was tree-killed / crashed and left an owned Excel resident | `excel_health.owned_pid` names the pid; Task Manager shows a lingering `EXCEL.EXE` | The Job Object (`KILL_ON_JOB_CLOSE`) auto-reaps the owned Excel when the server dies, so a new start is normally clean. For a legacy marker (a server without the Job Object), verify the pid's identity (create-time + `EXCEL.EXE` image) and only then `Stop-Process -Id <pid>`; the startup reconcile never kills on pid-alone and never kills a cross-session process |
| Pivot inventory `too_large` / `capped` | A worksheet's measured field / `PivotItems` read cost reached the per-sheet budget during the single-sweep inventory | The `pivot_tables_list` audit record reports the sheet status `too_large` (no partial, sheet named) or `capped` (partial results kept); the overall status is the most severe across sheets | Read the specific PivotTable with `pivot_table_get` / `pivot_fields_list`, or narrow the sheet. The budget is measured between COM ops (never file size / row count / `UsedRange`); output cells are never materialized |
| `UNSUPPORTED_PIVOT_SOURCE` / `source_status="unsupported_source"` | The PivotCache source is external, OLAP or Data Model | `pivot_table_get` / `pivot_cache_get` shows the source type and `source_status="unsupported_source"`; a refresh returns `unsupported_source` | Read-only inspection still works; only database/Table-backed caches can be refreshed. The source type is gated first, so an unsupported cache is reported with identity only |
| `COM_BUSY` (1004) on external `SourceData` | An external PivotCache's `SourceData` getter raises Excel 1004 (`DISP_E_EXCEPTION`) | Previously this aborted the sweep; it is now handled | No action needed: `PivotCache.SourceData` is optional metadata read defensively — the cache is reported with `source_data_kind="unavailable"` and the single-sweep inventory continues |
| `SNAPSHOT_FAILED` | No successful snapshot exists, or the snapshot write failed | Any mutating tool returns `SNAPSHOT_FAILED`; `snapshot_list` is empty for the workbook | Call `snapshot_create` first; check write access to the snapshot directory |
| `INVALID_RANGE` / `LISTOBJECT_INTERSECTION` / `PROTECTED_TARGET` | Malformed/oversized address or payload; a range operation intersects an Excel Table; the target is protected | Read `context` (payload shape, `tables`, reason); `INVALID_RANGE` is also used for argument-schema failures (`reason=invalid_arguments`) | Fix the address/payload; use the Tables tools instead of structural range operations; unprotect the sheet or choose another target |
| `UNSUPPORTED_PIVOT_SOURCE` / `REFRESH_FAILED` | The PivotCache source is external, OLAP or Data Model; or the refresh returned no success signal | `pivot_table_get` / `pivot_cache_get` shows the source type; the refresh result reports `unsupported_source` / `failed` | Read-only inspection still works; only database/Table-backed caches can be refreshed. Check the source data |
| `WORKBOOK_NOT_FOUND` | Unknown or stale `workbook_id`, or the owned instance was lost | The error `code` is `WORKBOOK_NOT_FOUND`; `excel_list_workbooks` no longer lists the id | Re-open with `excel_open` and use the new `workbook_id`; restart the server if the instance was lost |
| `NOT_FOUND` | A shape or control addressed by `name` does not exist (F6) | `shape_get` / `shape_set_on_action` / `control_get` / `control_set_on_action` return `code="NOT_FOUND"` with `context.name` | List the objects first (`shapes_list` / `controls_list`) and use an exact name. This is a caller error, not contention — do **not** retry. Genuine COM contention on the same tools still reports `COM_BUSY` |
| `REVISION_CONFLICT` / `STALE_TARGET` | A stale `expected_source_hash` (VBA edit) or a stale PivotTable identity | `vba_replace_component_source` / `vba_insert_code` returns `REVISION_CONFLICT`; `macro_classify_background` returns `STALE_TARGET` | Re-read the component (`vba_read_component`) and pass the current `expected_source_hash`; re-read the target and retry |
| `AMBIGUOUS_FILTER` | A `where` filter matched zero or multiple rows | `table_update_rows` / `table_delete_rows` returns `AMBIGUOUS_FILTER` with a preview | Use explicit `indices`, or refine the filter until it matches exactly one row |
| Unexpected or stale values | Excel has not recalculated; `range_read_values` returns cached values | Re-read the cell; compare with `range_read_formulas` | The server does not force `Application.Calculate`; trigger a recalculation in Excel or account for cached values |
| Timeouts | A large workbook, a slow COM call, or a hung operation | `excel_health` shows `needs_recovery`; the error is `EXCEL_HUNG` | Read large ranges in chunks; close/reopen; restart the server if needed |
| Workbook locking | The file is locked by another process, or the same path is already open in the server | `excel_open` reports the path is already open, or a file-copy error occurs | Close the other process; `excel_close` the existing `workbook_id` before reopening |
| Wrong paths | A relative path, a path outside the roots, or a non-workbook file | `PATH_NOT_ALLOWED` with a `reason` | Use an absolute path inside an allowed root with an allowed extension; see [Configuration](03-configuration.md) |

## Reading a structured error

Every tool failure reaches the client as a structured JSON payload. The shape is:

```json
{
  "code": "PATH_NOT_ALLOWED",
  "message": "The path is not allowed by policy.",
  "hresult": null,
  "hresult_hex": null,
  "context": {
    "reason": "outside_roots",
    "operation": "mutating",
    "name": "report.xlsm"
  }
}
```

- `code` — a stable `ErrorCode` string (for example `APPROVAL_DENIED`,
  `PATH_NOT_ALLOWED`, `WORKBOOK_READ_ONLY`).
- `message` — a human-readable message.
- `hresult` / `hresult_hex` — the raw COM HRESULT and its `0x%08X` form, or `null`
  when the error did not come from COM.
- `context` — a structured dictionary (path, `workbook_id`, `reason`, `section`,
  `token`, `class`, and so on).

`ExcelVbaMcpServer.call_tool` catches the MCP SDK's generic wrapper, finds the
underlying `ExcelVbaMcpError` in the `__cause__` chain and re-raises a
`ToolError` whose message is the JSON of `to_dict()`. Argument-schema failures
(`pydantic.ValidationError`) are translated into `INVALID_RANGE` with
`reason=invalid_arguments`, so **every** tool failure reaches the client
structured.

## Health checks

Use these before and after a failure:

| Tool | What it tells you | Touches Excel? |
|------|-------------------|----------------|
| `server_health` | The server process is alive: `{"status": "ok", "service": "excel-vba-mcp", "version": ...}` | No |
| `excel_health` | Worker/session state, `accepting`, `owned_pid`, `active_operation`, `needs_recovery`, `notes`. Performs a read-only owned-pid liveness probe. | No (no COM, no forced re-run) |
| `excel_list_workbooks` | The open workbooks from session bookkeeping | No |

`excel_health` is the fastest way to distinguish "the server is fine but the
workbook call failed" from "the worker needs recovery".

## Collecting diagnostics

| Source | Location / tool | Use |
|--------|-----------------|-----|
| Server logs | `stderr` of the server process | Startup warnings, worker failures, tracebacks (fixed `INFO` level) |
| Audit trail | `EXCEL_VBA_MCP_AUDIT_LOG` (default `<temp>/excel_vba_mcp/audit.jsonl`) | One JSON record per operation: tool, redacted params, approval, result |
| Compile diagnostics | `vba_get_compile_diagnostics` | Cached compile evidence (always `is_compile=false` on the validated build) |
| Last macro result | `macro_get_last_result` | Status, duration, return value and structured error of the last run |
| Persisted test run | `test_get_run` (with `run_id`) | Per-case and per-step statuses without touching Excel |

When reporting an issue, include the tool name, the structured error `code` and
`context`, the `excel_health` payload, and the relevant `audit.jsonl` line
(with secrets redacted).

## When to restart the server / recover

- **A hung or timed-out macro needs no restart.** `macro_run` terminates the owned Excel PID and
  rebuilds the hidden instance in-process (F2b); `MacroRunResult.session_recovery` reports
  `"recovered"` or `"needs_recovery"`. Open workbooks are discarded, so re-open the workbook with
  `excel_open` and continue. The server process itself is never restarted by this path.
- **Close and reopen the workbook** first for a stale session or a per-path lock:
  `excel_close(workbook_id, save_changes=false)` then `excel_open` again. This resolves most
  stale-session and per-path lock issues.
- **Restart the server process** when `excel_recover` is refused (`recovery_possible=false`, no live
  identity-verified handle), the owned Excel process is gone and cannot be rebuilt, or `excel_health`
  reports `unavailable`. Otherwise prefer `excel_recover` (agent-directed) over a restart.
- **Do not kill user Excel processes.** The server only ever terminates its own
  `DispatchEx` pid; use the specific `owned_pid` from `excel_health` if you must
  clean up a legacy orphan, and verify its identity (create-time + `EXCEL.EXE` image) first — never
  kill on pid-alone.

## Startup vs runtime failures

| Symptom | Likely layer | First check |
|---------|--------------|-------------|
| The client reports the server failed to start | Process / launch | Run `uv run excel-vba-mcp` manually; check the client `command` / `args` |
| The server starts but no tools respond | Transport / config | `server_health`; check `stderr` for `Invalid server configuration` |
| Tools respond but every workbook call fails | Excel / COM | `excel_health`; confirm desktop Excel is running and VBOM is trusted |
| Read-only works but mutations fail | Session / policy | `excel_get_status`; check `read_only`, `access_mode` and the snapshot |
| One workbook works but another fails | Path / file | `PATH_NOT_ALLOWED` reason; file lock; extension |

## Diagnostic recipes

### A path is rejected

1. Read `context.reason` from the `PATH_NOT_ALLOWED` error.
2. Confirm the path is absolute and inside `EXCEL_VBA_MCP_ROOTS`.
3. If it is a UNC or device path, copy the file into an allowed root.
4. Re-run `excel_open`.

### A mutation or macro is denied

1. Check `excel_get_status.access_mode`.
2. If it is `restricted`, switch to `full` (or accept read-only work).
3. If it is `full` and `APPROVAL_DENIED` persists, re-check the operation class
   and confirm no other precondition is being misreported.

### A macro is blocked

1. Call `macro_list` to confirm the `Module.Procedure` and its classification.
2. Read the procedure with `vba_read_component`.
3. If `BACKGROUND_SAFETY_NOT_VERIFIED`, call `macro_classify_background` after
   review; re-classify after any VBA source change.
4. If `INTERACTIVE_UI_NOT_SUPPORTED`, remove the flagged token or leave the
   procedure unclassified.

### A session is hung

1. Call `excel_health`.
2. If the hang was a `macro_run` timeout/hang, the server already recovered in-process (F2b):
   read `session_recovery` from the `MacroRunResult` and **re-open the workbook** (the prior
   `workbook_id` is invalid). If `session_recovery` is `"needs_recovery"`, call `excel_recover` (or
   restart the server if recovery is refused).
3. For any other `needs_recovery` / `running/deadline_exceeded` state, call `excel_recover`
   (agent-directed). If it is refused, restart the server process.
4. If the worker is `unavailable`, restart the server process.
5. Do not kill user Excel processes; only the owned pid may be terminated (through the held handle
   by `excel_recover`, or a verified legacy `Stop-Process -Id` after an identity check).

## Before reporting an issue

Include: the tool name, the structured error `code` and `context`, the
`excel_health` payload, the Excel build (`excel_version` in the audit record or
`excel_open` result), and the relevant `audit.jsonl` line with secrets redacted.
This is usually enough to distinguish a configuration problem from an Excel or
COM problem.

## Related

- [Overview](01-overview.md) — scope.
- [Installation and setup](02-installation-and-setup.md) — first-run checks.
- [Configuration](03-configuration.md) — variables and paths.
- [Architecture](04-architecture.md) — error handling and recovery.
- [MCP tools reference](05-mcp-tools-reference.md) — per-tool errors.
- [Excel capabilities](06-excel-capabilities.md) — capability limits.
- [Workflows and examples](07-workflows-and-examples.md) — error-handling workflow.
- [Security](08-security.md) — security boundaries.
- [Limitations and compatibility](09-limitations-and-compatibility.md) — known limits.
- [Development and maintenance](11-development-and-maintenance.md) — contributor guide.