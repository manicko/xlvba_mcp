---
id: xlvba-architecture
domain: 04-architecture
tags:
  - architecture
  - mcp
  - excel
  - com
  - sta
  - vba
related:
  - xlvba-overview
  - xlvba-configuration
  - xlvba-mcp-tools-reference
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This document explains how `xlvba` **behaves from your point of view** — the consequences that
matter when you automate Excel: that there is one managed Excel session, that a long operation
delays later ones, that your original file is never touched, and how status, health and recovery
work. It deliberately avoids internal design detail; the tool contract is in
[MCP tools reference](05-mcp-tools-reference.md), environment settings in
[Configuration](03-configuration.md), the security model in [Security](08-security.md), and limits
in [Limitations and compatibility](09-limitations-and-compatibility.md).

## One managed Excel session

- `xlvba` owns **one hidden Excel instance** for its whole lifetime. You never see a window, and the
  agent never activates a window or sends keystrokes.
- All work is **serialized**: the server handles one Excel operation at a time. A long read, export
  or macro therefore **delays** later operations. This is the single most important operational fact
  to plan around.
- A **per-path lock** means the same workbook file cannot be opened twice at once. Reopening the same
  path is rejected until you close the first `workbook_id`.

**What this means for you:** treat the server as a single worker. Run very large reads or exports
with `data_export` (streams to a file) or `workbook_snapshot_read` (reads a snapshot without
occupying the worker), and expect a big macro to hold the queue until it finishes or hits its
deadline.

## Your original file is never modified

- `workbook` (`open`) copies the source into a working copy and opens **only the copy**. Every read,
  write and mutation targets the working copy; the original file is never written.
- A workbook carries **two paths**: `path` (the working copy actually opened) and `source_path`
  (the original). There is no `original_name`.
- **Edits live only in the working copy.** `workbook` (`close`) deletes it, and a later
  `workbook` (`open`) re-copies the untouched original — so "close, then reopen" does **not** show
  prior edits.
- **Persist deliberately.** `workbook` (`save`) writes the working copy; `workbook` (`close` with
  `save_changes=true`) saves then closes; `workbook_save_copy` writes a verified copy to a new (or
  the original) path with a `.bak` and a size + SHA-256 check.

## Modes you choose at open

| Choice | Effect |
|--------|--------|
| `mode="execute"` (default) | Macros are enabled per Office policy; the session is writable and execution-capable. |
| `mode="inspect"` | Opens with macros disabled (`AutomationSecurity = ForceDisable`). |
| `read_only=true` | Opens read-only; every mutating tool returns `WORKBOOK_READ_ONLY`. |

`inspect` does **not** imply read-only — pass `mode="inspect"` **and** `read_only=true` for a
strictly read-only session. Neither mode changes Trust Center or registry settings.

## Snapshots and rollback

- `snapshot_create` captures an **immutable** baseline with a SHA-256 manifest. Open workbooks are
  captured with `SaveCopyAs`; closed workbooks are copied.
- `snapshot_list` and `snapshot_diff` compare snapshots by manifest metadata and file hash (never a
  binary diff).
- `snapshot_restore` discards the current working copy (unsaved changes) and reopens from the
  baseline. It requires `confirm=true` and the **old** `workbook_id`; the returned `workbook.id` is
  the **new** active id — use it for every later call.
- **Snapshots restore files, not external effects.** A macro's database writes, emails, network calls
  and printed pages are **not** undone.

## The mutation guard chain (what you can rely on)

Every mutating or executing tool runs the same order, and you can depend on it:

1. **Writable session** — a read-only session is rejected with `WORKBOOK_READ_ONLY`.
2. **Baseline snapshot** — without a successful `snapshot_create`, the mutation fails with
   `SNAPSHOT_FAILED`.
3. **Approval** — in the default `full` mode the operation is authorised; in `restricted` mode it is
   refused with `APPROVAL_DENIED` unless an in-process allowlist entry matches.
4. **Operation-specific checks** — payload shape, explicit `shift`, unambiguous `where`, no
   ListObject intersection, supported PivotTable source.
5. **Postcondition re-read + audit** — the result is re-read and `verified` reflects the real state,
   not merely the absence of an error; every call is written to the audit log.

The one documented exception: `macro_run` does **not** require a prior snapshot (snapshot first
anyway so a macro's external effects have a rollback baseline).

## Macros are gated, and interactive macros are refused

- `macro_run` runs a procedure only when it is classified `background_safe` for the current workbook
  hash **and** VBA source revision, and it passes the interactive-UI preflight.
- A procedure that shows UI (`MsgBox`, `InputBox`, `UserForm.Show`, `FileDialog`, …) is vetoed
  **before** dispatch with `INTERACTIVE_UI_NOT_SUPPORTED` — the macro is not started. The hidden,
  non-activating instance has no GUI, so a modal would block the single worker indefinitely.
- The R13 scan is a **heuristic veto, not proof**. Review VBA before classifying it.
- A macro that hangs or exceeds its soft deadline (default 90 s) terminates only the owned Excel
  process and recovers the session **in-process**; open workbooks are discarded and the macro is
  never auto-replayed.

## Long operations, jobs and cancellation

- Four tools accept `background=true`: `workbook` (`open`), `macro_run`, `test_suite` (`run`) and
  `pivot_read` (`tables-list`). They return a `RunningJob` handle promptly.
- Poll with `job_get(job_id)` no sooner than `poll_after_s`. The server **owns the result**, so a
  client timeout or disconnect never drops it.
- `job_request_cancel` records **cancel intent only** — it never interrupts the running call, which
  still reaches a terminal state.
- A job left non-terminal by a restart is reported `interrupted_unknown` (`outcome="unknown"`,
  `retryable=false`) and is **never** auto-retried; re-issue it as a new job if still needed.

## Status, health and recovery

| Tool | What it tells you | Touches Excel? |
|------|-------------------|----------------|
| `server_health` | The server process is alive. | No |
| `excel_health` | Worker/session state, `accepting`, `owned_pid`, `active_operation`, `needs_recovery`. | No |
| `workbook_read` (`status` / `list`) | Session status and the open-workbook list. | No |
| `xlvba_environment` | Access mode, policy, allowed roots, capabilities (build is always `null`). | No |

- **Excel unavailable at startup is non-fatal.** The server still starts; the tool list and
  `server_health` answer, and workbook tools return a truthful unavailability error instead of a
  false success.
- **Recovery is agent-directed.** When `excel_health` shows `needs_recovery` (or
  `running/deadline_exceeded`), call `excel_recover` with `confirm=true`. It is pid-scoped,
  originals-safe (it discards working copies only) and fail-closed. If it is refused, restart the
  server process.
- **The one automatic path** is a `macro_run` hang/timeout, which rebuilds the instance itself and
  reports `session_recovery` in the result.

## Errors are structured

Every tool failure reaches the client as JSON with a stable `code` and a structured `context`
(for example `{"code": "PATH_NOT_ALLOWED", "context": {"reason": "outside_roots", ...}}`). The same
error is written to the audit log. Argument-schema failures are normalised to `INVALID_RANGE` with
`reason=invalid_arguments`. A symptom-first guide is in [Troubleshooting](10-troubleshooting.md).

## Compile is diagnostic-only

`vba_compile_diagnostics` never claims a successful compile. On the validated build it always
returns `status="indeterminate"` (`reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"`,
`is_compile=false`) and never shows UI. To actually compile, open the workbook in the Excel VBA editor
and use **Debug → Compile VBAProject**.

## Where to read more

- [MCP tools reference](05-mcp-tools-reference.md) — the full per-tool contract.
- [Security](08-security.md) — trust boundaries, access model and the risk table.
- [Limitations and compatibility](09-limitations-and-compatibility.md) — validated platform and limits.
- [Troubleshooting](10-troubleshooting.md) — symptoms, diagnostics and recovery.
- [Development and maintenance](11-development-and-maintenance.md) — the internal layer design for
  contributors.