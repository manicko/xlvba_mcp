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

This document starts from **what you see** — the symptom — and maps it to the cause, the
diagnostic and the fix. It then lists the structured error codes for when you need the exact
`code`/`context`.

For the error contract see [Architecture](04-architecture.md) and
[MCP tools reference](05-mcp-tools-reference.md). For configuration see
[Configuration](03-configuration.md).

## Symptoms first

Find the symptom that matches what you are seeing. Each row names the likely cause and the first
thing to do.

| Symptom | Likely cause | First thing to do |
|---------|--------------|-------------------|
| **Excel is not responding / a call hangs** | A long read, export or macro is still running on the single Excel instance; or a modal dialog is blocking it | Check `excel_health`. A `macro_run` hang/timeout self-recovers — re-open the workbook. For any other `needs_recovery` / `running/deadline_exceeded` state, call `excel_recover` (`confirm=true`). See [Excel not responding](#excel-not-responding). |
| **The workbook is already open** | The same file path is already open in the server, or another process locks the file | `workbook` (`close`) the existing `workbook_id` before reopening, or close the other program. See [Workbook already open](#workbook-already-open). |
| **The PivotTable source is unsupported** | The cache source is external, OLAP or Data Model | Read-only inspection and `pivot_read` (`result-read`) still work; only database/Table-backed caches can be refreshed or reconfigured. See [Unsupported pivot source](#unsupported-pivot-source). |
| **The range is too large** | A well-formed range exceeds the 100,000-cell cap | Split the work into smaller ranges with `offset` / `limit`, or stream it with `data_export`. See [Range too large](#range-too-large). |
| **The macro needs interactive UI** | The procedure shows a dialog or form (`MsgBox`, `InputBox`, `UserForm.Show`, `FileDialog`, …) | Remove the interactive UI or run it in an interactive Excel session; the headless server refuses it (`INTERACTIVE_UI_NOT_SUPPORTED`). See [Macro needs interactive UI](#macro-needs-interactive-ui). |
| **A mutation or macro is denied** | The server runs in `restricted` mode | Unset `EXCEL_VBA_MCP_ACCESS_MODE` (or set `full`) for full access; a restricted server exposes only read-only work over MCP. |
| **A path is rejected** | The path is outside `EXCEL_VBA_MCP_ROOTS`, or is a device/UNC/traversal/symlink path, or has a disallowed extension | Read `context.reason`; put the file inside an allowed root or extend the roots and restart the server. |
| **The VBA tools return `ACCESS_VBOM_DISABLED`** | Trust access to the VBA project object model is off | Enable it in Excel Trust Center → Macro Settings, restart Excel, retry. Corporate policy may block it. |
| **A macro is blocked (`BACKGROUND_SAFETY_NOT_VERIFIED`)** | The procedure is not classified `background_safe` for the current workbook hash and VBA source revision | Review the code, then classify it with `macro` (`classify-background`); re-classify after any VBA source change. |
| **A VBA edit returns `REVISION_CONFLICT`** | The `expected_source_hash` is stale | Re-read the component with `vba_read` (`read-component`) and pass the current hash. |
| **Values look stale** | Excel has not recalculated; `range_read` returns cached values | Re-read the cell; the server does not force `Application.Calculate`. |
| **The client cannot connect / the server exits at startup** | Wrong `command`/`args`, `uv` not on `PATH`, or an invalid environment value | Run `uv run excel-vba-mcp` manually; check `stderr` for `Invalid server configuration`. See [Installation and setup](02-installation-and-setup.md). |

### Excel not responding

The server owns **one** hidden Excel instance, so a single long operation (a big read, a large
export, a slow macro) delays everything else. This is expected, not a deadlock.

1. Call `excel_health`.
   - `worker_state="running/deadline_exceeded"` means a call exceeded its soft deadline. The call is
     **not** cancelled and the process is **not** terminated; the worker resumes `running` when the
     slow call returns, so re-check before recovering.
   - `needs_recovery=true` means the owned process is wedged.
2. If the hang was a `macro_run` timeout/hang, the server already recovered **in-process** — read
   `session_recovery` from the `MacroRunResult` and **re-open the workbook** (the old `workbook_id`
   is invalid).
3. For any other `needs_recovery` state, call `excel_recover` with `confirm=true` (agent-directed,
   pid-scoped, originals-safe). If it is refused (`recovery_possible=false`), restart the server
   process.
4. Do not kill user Excel processes; only the owned pid may be terminated.

### Workbook already open

A per-path lock serializes opens, so the same canonical path cannot be open twice.

1. Call `workbook_read` (`list`) to see the open workbooks and their ids.
2. Call `workbook` (`close`) on the existing `workbook_id` (with `save_changes=false` to discard),
   then `workbook` (`open`) again.
3. If another program (or a manual Excel window) holds the file, close it first.

### Unsupported pivot source

External / OLAP / Data-Model PivotTables cannot be refreshed or reconfigured, but their **displayed
result is readable**.

1. Call `pivot_read` (`table-get` / `cache-get`) to confirm the source type
   (`source_status="unsupported_source"`).
2. Use `pivot_read` (`result-read`) to read the rendered output, or `data_export`
   (`source_kind="pivot-result"`) to stream it to a file.
3. Only database/Table-backed caches can be refreshed (`pivot_refresh`) or reconfigured
   (`pivot_field`).

### Range too large

A well-formed range whose cell count exceeds 100,000 is rejected **before execution** with
`RANGE_TOO_LARGE` (never silently truncated).

1. Split the range into smaller windows and loop with `offset` / `limit` on `range_read`
   (`read-values` / `read-formulas`).
2. For a one-shot file, use `data_export` (`source_kind="range"`) to stream it with bounded memory.
3. A page clipped by `offset` / `limit` is reported with `truncated=true`, which is different from
   an over-cap rejection.

### Macro needs interactive UI

`macro_run` and the declarative test runner cannot run a procedure that shows UI. The R13 preflight
vetoes it **before** dispatch with `INTERACTIVE_UI_NOT_SUPPORTED` and the macro is not started.

1. Read the procedure with `vba_read` (`read-component`).
2. Remove the interactive tokens (`MsgBox`, `InputBox`, `UserForm.Show`, `FileDialog`, …), or run
   the procedure in an interactive Excel session instead.
3. Do not try to bypass the preflight; the hidden, non-activating instance has no GUI, so a modal
   would block the single worker indefinitely.

## Health checks

Use these before and after a failure:

| Tool | What it tells you | Touches Excel? |
|------|-------------------|----------------|
| `server_health` | The server process is alive: `{"status": "ok", "service": "excel-vba-mcp", "version": ...}` | No |
| `excel_health` | Worker/session state, `accepting`, `owned_pid`, `active_operation`, `needs_recovery`, `notes`; performs a read-only owned-pid liveness probe | No (no COM, no forced re-run) |
| `workbook_read` (`list`) | The open workbooks from session bookkeeping | No |
| `xlvba_environment` | Access mode, policy, allowed roots, capabilities (no COM) | No |

`excel_health` is the fastest way to distinguish "the server is fine but the workbook call failed"
from "the worker needs recovery".

## Reading a structured error

Every tool failure reaches the client as a structured JSON payload:

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

- `code` — a stable `ErrorCode` string (for example `APPROVAL_DENIED`, `PATH_NOT_ALLOWED`,
  `WORKBOOK_READ_ONLY`).
- `message` — a human-readable message.
- `hresult` / `hresult_hex` — the raw COM HRESULT and its `0x%08X` form, or `null` when the error
  did not come from COM.
- `context` — a structured dictionary (path, `workbook_id`, `reason`, `section`, `token`, `class`,
  and so on).

Argument-schema failures (`pydantic.ValidationError`) are translated into `INVALID_RANGE` with
`reason=invalid_arguments`, so **every** tool failure reaches the client structured.

## Error code reference

| Code | Where it surfaces | Typical cause | What to do |
|------|-------------------|---------------|------------|
| `PATH_NOT_ALLOWED` | `workbook` (`open`), `vba_export_project` | Path outside `EXCEL_VBA_MCP_ROOTS`, a device/UNC/traversal/symlink path, or a disallowed extension. `context.reason` names the category; `context.allowed_roots` lists the loaded roots. | Put the workbook inside an allowed root or extend `EXCEL_VBA_MCP_ROOTS` and restart the server. |
| `APPROVAL_DENIED` | `workbook` (writable/execute), any mutating tool, `macro_run`, `test_suite` | The server runs in `restricted` mode and no allowlist entry matches. | Unset `EXCEL_VBA_MCP_ACCESS_MODE` (or set `full`). |
| `WORKBOOK_READ_ONLY` | Any mutating tool | The session was opened with `read_only=true`. | Reopen writable (`workbook` `open` without `read_only`). |
| `ACCESS_VBOM_DISABLED` | `vba_read`, `vba_*`, `macro`, `vba_compile_diagnostics` | VBOM trust is off, or corporate policy blocks it. | Enable **Trust access to the VBA project object model**, restart Excel, retry. |
| `VBA_PROJECT_LOCKED` | VBA tools | The VBA project is password-protected. | Unlock the project in the VBE; protection is never bypassed. |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_run` (status `blocked`) | The procedure is not classified `background_safe` for the current workbook hash and VBA source revision. | The macro was **not** started. Classify it with `macro` (`classify-background`) and retry. |
| `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_run`, `macro` (`classify-background`), `test_suite` | The procedure shows UI, is heuristic, or was classified safe while unsafe. `context.flagged_tokens` lists the hits. | Remove the interactive UI or leave it unclassified. |
| `MACRO_NOT_FOUND` / `MACRO_BLOCKED` / `MACRO_RUNTIME_ERROR` / `MACRO_TIMEOUT` | `macro_run` | Wrong name; Office macro policy; a runtime error; the 90 s soft deadline was exceeded. | Fix `Module.Procedure`; check macro settings; inspect the macro. After a timeout the owned Excel was terminated and the session recovered — re-open the workbook. |
| `REVISION_CONFLICT` | `vba_replace_component_source`, `vba_insert_code` | A stale `expected_source_hash`. | Re-read the component with `vba_read` (`read-component`) and pass the current hash. |
| `STALE_TARGET` | `macro` (`classify-background`) | A stale `expected_project_source_hash`. | Re-read the target and retry. |
| `NOT_FOUND` | `shape`, `control` | A shape or control addressed by `name` does not exist. `context.name` carries the requested name. | List the objects (`shape` / `control` `list`) and use an exact name. Do not retry. |
| `COM_BUSY` | Any mutating tool | Genuine COM contention (the single worker is busy). | Read `context.reason`; retry a read after the operation completes. Do not blindly retry a mutation. |
| `SNAPSHOT_FAILED` | Any mutating tool | No successful snapshot exists, or the snapshot write failed. | Call `snapshot_create` first; check write access to the snapshot directory. |
| `UNSUPPORTED_PIVOT_SOURCE` | `pivot_refresh`, `pivot_field` | The cache source is external, OLAP or Data Model. | Read-only inspection and `pivot_read` (`result-read`) still work. |
| `AMBIGUOUS_FILTER` | `table` (`update-rows`), `table_delete_rows` | A `where` filter matched zero or multiple rows. | Use explicit `indices`, or refine the filter. |
| `INVALID_RANGE` / `LISTOBJECT_INTERSECTION` / `PROTECTED_TARGET` | Range and table tools | A malformed/oversized address or payload; a range operation intersects an Excel Table; the target is protected. | Fix the address/payload; use the table tools; unprotect or choose another target. |
| `STALE_CURSOR` | `pivot_read` (`result-read`) | The paged cursor version no longer matches the live result. | Restart the read at `offset` 0. |
| `WORKBOOK_NOT_FOUND` | Any tool with `workbook_id` | Unknown or stale `workbook_id`, or the owned instance was lost. | Re-open with `workbook` (`open`) and use the new `workbook_id`. |
| `EXCEL_HUNG` | Any tool | A call exceeded its soft deadline; `context.verdict="SUSPECTED_STALLED"`. | A macro hang self-recovers; otherwise call `excel_recover`. See [Excel not responding](#excel-not-responding). |

Additional codes declared by the server include `MISSING_REFERENCE`, `RPC_E_CALL_REJECTED`,
`COMPILE_ERROR`, `COMPILE_INDETERMINATE`, `MACRO_BLOCKED`, `REFRESH_FAILED`, `MODAL_UI_BLOCKED` and
`EXCEL_HUNG`. The full contract is in [MCP tools reference](05-mcp-tools-reference.md).

## When to restart the server / recover

- **A hung or timed-out macro needs no restart.** `macro_run` rebuilds the hidden instance in-process
  and reports `session_recovery`; re-open the workbook and continue.
- **Close and reopen the workbook** first for a stale session or a per-path lock:
  `workbook` (`close`, `save_changes=false`) then `workbook` (`open`).
- **Restart the server process** when `excel_recover` is refused (`recovery_possible=false`), the
  owned Excel is gone and cannot be rebuilt, or `excel_health` reports `unavailable`.
- **Do not kill user Excel processes.** The server only ever terminates its own `DispatchEx` pid; use
  the specific `owned_pid` from `excel_health` if you must clean up a legacy orphan, and verify its
  identity first — never kill on pid-alone.

## Diagnostic recipes

### A path is rejected

1. Read `context.reason` from the `PATH_NOT_ALLOWED` error.
2. Confirm the path is absolute and inside `EXCEL_VBA_MCP_ROOTS`.
3. If it is a UNC or device path, copy the file into an allowed root.
4. Re-run `workbook` (`open`).

### A mutation or macro is denied

1. Check `xlvba_environment` (or `workbook_read` `status`) for the access mode.
2. If it is `restricted`, switch to `full` (or accept read-only work).
3. If it is `full` and `APPROVAL_DENIED` persists, re-check the operation class.

### A macro is blocked

1. Call `macro` (`list`) to confirm the `Module.Procedure` and its classification.
2. Read the procedure with `vba_read` (`read-component`).
3. If `BACKGROUND_SAFETY_NOT_VERIFIED`, call `macro` (`classify-background`) after review; re-classify
   after any VBA source change.
4. If `INTERACTIVE_UI_NOT_SUPPORTED`, remove the flagged token or leave the procedure unclassified.

### A session is hung

1. Call `excel_health`.
2. If the hang was a `macro_run` timeout/hang, the server already recovered in-process — read
   `session_recovery` and **re-open the workbook**.
3. For any other `needs_recovery` / `running/deadline_exceeded` state, call `excel_recover`
   (agent-directed). If it is refused, restart the server process.
4. Do not kill user Excel processes.

## Before reporting an issue

Include: the tool name, the structured error `code` and `context`, the `excel_health` payload, the
Excel build (`excel_version` in the audit record or the `workbook` `open` result), and the relevant
`audit.jsonl` line with secrets redacted. This is usually enough to distinguish a configuration
problem from an Excel or COM problem.

## Related

- [Overview](01-overview.md) — scope.
- [Installation and setup](02-installation-and-setup.md) — first-run checks.
- [Configuration](03-configuration.md) — variables and paths.
- [Architecture](04-architecture.md) — error handling and recovery.
- [MCP tools reference](05-mcp-tools-reference.md) — per-tool errors.
- [Excel capabilities](06-excel-capabilities.md) — capability limits.
- [Workflows and examples](07-workflows-and-examples.md) — safe workflow recipes.
- [Security](08-security.md) — security boundaries.
- [Limitations and compatibility](09-limitations-and-compatibility.md) — known limits.
- [Development and maintenance](11-development-and-maintenance.md) — contributor guide.