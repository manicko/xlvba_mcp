---
id: xlvba-workflows-and-examples
domain: 07-workflows-and-examples
tags:
  - workflows
  - examples
  - mcp
  - excel
  - vba
related:
  - xlvba-overview
  - xlvba-excel-capabilities
  - xlvba-mcp-tools-reference
  - xlvba-configuration
  - xlvba-security
  - xlvba-troubleshooting
---

## Purpose

This document shows **real, usable workflows** for the `xlvba` MCP server. Each workflow states its
goal, the required setup, the exact tool calls with JSON arguments, the expected result, how to
verify it, and its limitations and safety notes.

Every example uses the canonical MCP tool names and the parameter names from
[MCP tools reference](05-mcp-tools-reference.md). Each workflow is labelled:

- **Tested (integration)** — a real-workbook integration test composes the full chain (verified on a
  real `.xlsm`).
- **Illustrative** — every link exists and is unit-tested, but no single end-to-end test drives the
  whole chain; treat it as a recipe, not a transcript.

The capabilities behind these workflows are classified in
[Excel capabilities](06-excel-capabilities.md). Settings are in
[Configuration](03-configuration.md); the security and approval model is in
[Architecture](04-architecture.md).

## Conventions

| Convention | Meaning |
|------------|---------|
| `<workbook_id>` | Server-assigned identifier of an open working copy, returned by `workbook` (`open`). It is **not** a file path. Replace it in every subsequent call. |
| JSON blocks | Show the tool **arguments** object exactly as it is sent over MCP. A block is not the full JSON-RPC envelope. |
| `xlvba_` prefix | Clients expose published names with a server prefix. In Kilo, for example, `workbook` is called as `xlvba_workbook`. This document uses the published names without the prefix; add your client's prefix. |
| **Approval-gated (restricted mode)** | The call is a mutating or executing operation. In the default `full` access mode it is authorised (a baseline snapshot is still required before a mutation); in `restricted` mode it is refused with `APPROVAL_DENIED`. The implemented `xlvba_environment` tool reports the active access mode, policy and build without touching Excel. |
| Read-only | No workbook change; auto-allowed by the approval policy. |

Paths are Windows paths. A workbook path must be inside `EXCEL_VBA_MCP_ROOTS`
(by default the current working directory and the system temp directory); otherwise the call
fails with `PATH_NOT_ALLOWED`.

## Workflow 1: Audit an unfamiliar workbook (read-only) — Tested (integration)

**Goal.** Understand a workbook's structure before touching it, with zero mutation risk.

**Setup.** The server is running and the workbook is inside an allowed root.

**Step 1 — open a read-only inspect session.** Call `workbook` with `action="open"`:

```json
{ "action": "open", "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
```

The result contains `workbook_id`, `name`, `source_path` (the original), `path` (the working copy),
`read_only: true` and `mode: "inspect"`.

**Step 2 — read the aggregate model (or one inventory).** Call `workbook_read` with
`action="model"`:

```json
{ "action": "model", "workbook_id": "<workbook_id>", "include_vba": false }
```

Alternatively call `sheet_read` with `action="list"`, `table_read` with `action="list"`, or
`pivot_read` with `action="tables-list"` for one inventory.

**Step 3 — close without saving.** Call `workbook` with `action="close"`:

```json
{ "action": "close", "workbook_id": "<workbook_id>", "save_changes": false }
```

**Expected result.** A `WorkbookModel` (or one inventory list) followed by a session status with
`state="closed"`.

**How to verify.** The model reports the sheet/table/name/pivot counts; the close result reports the
session closed and the working copy deleted.

**Limitations / safety.** `mode="inspect"` disables macros; `read_only=true` makes every mutation
unavailable (`WORKBOOK_READ_ONLY`). The model never reads all cells and never executes VBA. The
original is never modified.

## Workflow 2: Add a macro to an existing `.xlsm` — Tested (integration)

**Goal.** Insert or replace VBA code in an existing macro-enabled workbook with revision safety.

**Setup.** A workbook opened read-only is enough to read; use a writable open
(`workbook` `open` without `read_only`) to edit. VBOM trust must be on and the project unlocked.

**Step 1 — read the component and keep its revision token.** Call `vba_read` with
`action="read-component"`:

```json
{ "action": "read-component", "workbook_id": "<workbook_id>", "component": "Module1" }
```

Keep the returned `source_hash` — it is the revision token the edit tools require.

**Step 2 — insert code at a semantic position (or replace the whole module).** Call
`vba_insert_code`:

```json
{
  "workbook_id": "<workbook_id>",
  "component": "Module1",
  "code": "Option Explicit\n\nPublic Sub Hello()\n  MsgBox \"hi\"\nEnd Sub",
  "position": "after_procedure:Hello",
  "expected_source_hash": "<source_hash from step 1>"
}
```

Use `vba_replace_component_source` with `new_source` instead to replace the whole module.

**Step 3 — confirm the new hash.** Call `vba_read` `read-component` again and compare `source_hash`.

**Expected result.** A `VbaMutationResult` and a new, different `source_hash`.

**How to verify.** Re-read the component; the source contains the new code and the hash changed.

**Limitations / safety.** Standard/Class modules only (Document modules and UserForms are rejected).
A stale `expected_source_hash` raises `REVISION_CONFLICT` — re-read and retry. Edits live only in the
working copy; save or copy to persist.

## Workflow 3: Run a reviewed macro and capture the result — Tested (integration)

**Goal.** Run a non-interactive macro safely and get a typed outcome.

**Setup.** An execute-mode, writable session (`workbook` `open` without `read_only`) and a baseline
`snapshot_create`. Review the code before classifying it.

**Step 1 — list macros and classify the target as background-safe.** Call `macro` with
`action="list"`, then `action="classify-background"`:

```json
{ "action": "list", "workbook_id": "<workbook_id>" }
```

```json
{ "action": "classify-background", "workbook_id": "<workbook_id>", "procedure": "Module1.Stamp", "background_safe": true }
```

**Step 2 — run the macro.** Call `macro_run`:

```json
{ "workbook_id": "<workbook_id>", "macro_name": "Module1.Stamp", "args": [] }
```

**Step 3 — read the last result (or a cell the macro writes).** Call `macro` with
`action="get-last-result"`, or `range_read` `read-values` on the expected cell:

```json
{ "action": "get-last-result", "workbook_id": "<workbook_id>" }
```

**Expected result.** A `MacroRunResult` with `status="success"` (or a typed error), and the expected
cell value.

**How to verify.** `get-last-result` returns the typed outcome; a follow-up `range_read` confirms the
macro's effect.

**Limitations / safety.** The macro must be `background_safe` for the current workbook hash **and**
VBA source revision; otherwise it is `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`) and **not started**.
A procedure that shows UI is refused (`INTERACTIVE_UI_NOT_SUPPORTED`). A hang/timeout terminates only
the owned Excel process and recovers the session; open workbooks are discarded and the macro is never
auto-replayed. External side effects are not rolled back. (Reviewed real-data runs were R13-blocked or
not applicable, so a successful end-to-end run is verified at the chain level, not as a specific
macro transcript.)

## Workflow 4: Regression-test a change with disposable copies — Tested (integration)

**Goal.** Verify a change with repeatable, isolated edit→macro→assert cases; originals untouched.

**Setup.** An open workbook. The suite opens its own disposable copies in execute mode; in `full`
mode this is authorised, while `restricted` mode returns the run as `blocked` (`APPROVAL_DENIED`).

**Step 1 — submit a closed-vocabulary suite.** Call `test_suite` with `action="run"`:

```json
{
  "action": "run",
  "suite": {
    "suite": "smoke",
    "workbook_id": "<workbook_id>",
    "baseline": "create",
    "cases": [
      {
        "id": "sheet-and-range",
        "steps": [
          { "action": "sheet_add", "sheet": "Scratch" },
          {
            "action": "range_write_values",
            "sheet": "Scratch",
            "range_address": "A1:B2",
            "values_grid": {
              "rows": 2,
              "cols": 2,
              "values": [
                [ { "kind": "string", "value": "Name" }, { "kind": "string", "value": "Qty" } ],
                [ { "kind": "string", "value": "Widget" }, { "kind": "number", "value": 7 } ]
              ]
            }
          }
        ],
        "assertions": [
          { "type": "sheet_exists", "sheet": "Scratch" },
          { "type": "cell_value_equals", "sheet": "Scratch", "cell": "B2", "expected": 7 }
        ]
      },
      {
        "id": "compile-and-macro",
        "steps": [
          { "action": "vba_compile" },
          { "action": "macro_run", "macro": "Module1.Hello", "args": [], "background_safe": true }
        ],
        "assertions": [
          { "type": "compile_status", "expected": "indeterminate" },
          { "type": "macro_status", "expected": "success" }
        ]
      }
    ]
  }
}
```

**Step 2 — read the persisted run.** Call `test_suite` with `action="get-run"` and the `run_id`:

```json
{ "action": "get-run", "run_id": "<test_run_id>" }
```

**Expected result.** A `TestRunResult` with one entry per case and step, plus a persisted run.

**How to verify.** `get-run` returns per-case and per-step statuses without touching Excel; a failed
`vba_compile` step blocks later `macro_run` steps in the same case (an `indeterminate` compile is not
a failure).

**Limitations / safety.** The action/assertion vocabularies are closed (37 actions, 29 assertions);
`vba_replace_component_source` / `vba_insert_code` steps are rejected (a declarative step cannot carry
the revision hash). External side effects are not rolled back.

## Workflow 5: Roll back a bad change — Tested (integration)

**Goal.** Return a workbook to a known-good baseline after a bad edit.

**Setup.** A workbook open writable, with a snapshot taken before the change.

**Step 1 — create a baseline snapshot.** Call `snapshot_create`:

```json
{ "workbook_id": "<workbook_id>" }
```

**Step 2 — make the change** (for example `range_write` `write-values`), then **restore**. Call
`snapshot_restore` with the old `workbook_id` and `confirm=true`:

```json
{ "snapshot_id": "snap-…", "workbook_id": "<old-workbook_id>", "confirm": true }
```

**Step 3 — use the returned new id.** The result's `workbook.id` is the new active id; re-read the
restored cell to confirm.

**Expected result.** A `SnapshotRestoreResult` with the restored `WorkbookInfo`, a new `workbook.id`,
and `previous_close` (`"closed"` / `"failed"` / `"not_requested"`).

**How to verify.** Re-read the target with `range_read`; the value matches the baseline.

**Limitations / safety.** Restore is destructive to the working copy (unsaved changes are discarded)
and needs `confirm=true` plus the old `workbook_id`. It does **not** undo external macro side effects.
Re-restoring the already-active baseline is rejected (`reason="same_snapshot"`).

## Workflow 6: Bulk-edit a large table safely — Tested (integration)

**Goal.** Update thousands of rows in an Excel Table without editing cells manually.

**Setup.** A workbook open writable, with a baseline snapshot.

**Step 1 — read a bounded slice to confirm structure.** Call `table_read` with `action="read-rows"`:

```json
{ "action": "read-rows", "workbook_id": "<workbook_id>", "name": "Sales", "offset": 0, "limit": 50 }
```

**Step 2 — update rows by explicit index or unambiguous `where`.** Call `table` with
`action="update-rows"`:

```json
{
  "action": "update-rows",
  "workbook_id": "<workbook_id>",
  "name": "Sales",
  "updates": [
    { "index": 2, "values": { "Amount": 120 } },
    { "where": { "Region": "North" }, "values": { "Flag": "Y" } }
  ]
}
```

Use `table` `append-rows` to add rows, `table_delete_rows` to remove them.

**Step 3 — verify.** Call `table_read` `read-rows` again.

**Expected result.** A `TableMutationResult` with `verified=true` and the affected row count.

**How to verify.** Re-read the rows; the values match the update and the row count is as expected.

**Limitations / safety.** Append/insert go through `ListRows.Add`, so table structure is preserved. An
ambiguous `where` (zero or multiple rows) is rejected with `AMBIGUOUS_FILTER` — use explicit indices
or refine the filter.

## Workflow 7: Page through a huge range — Tested (integration)

**Goal.** Read a range too large to load at once, with bounded memory and stale-cursor protection.

**Setup.** A workbook open read-only.

**Step 1 — read the first page.** Call `range_read` with `action="read-values"`:

```json
{ "action": "read-values", "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:T200000", "offset": 0, "limit": 2000 }
```

**Step 2 — request the next page** with the returned `cursor_version` (and a higher `offset`):

```json
{ "action": "read-values", "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:T200000", "offset": 2000, "limit": 2000, "cursor_version": "<page.version>" }
```

Repeat until `has_more` is `false` / `complete` is `true`.

**Expected result.** A sequence of `RangeValues` pages, each with `rows`, `cols`, `truncated` and a
`page` cursor.

**How to verify.** The last page reports completeness; the row count across pages matches the range.

**Limitations / safety.** Reads are bounded (100,000 cells per target, 50,000 cells per response). If
the workbook changes between pages, the cursor is stale and the next page returns `STALE_CURSOR` —
restart at offset 0.

## Workflow 8: Refresh and read a PivotTable — Tested (integration)

**Goal.** Get the current PivotTable output without `RefreshAll`.

**Setup.** A workbook open read-only. Only database/Table-backed PivotTables are refreshable;
external, OLAP and Data Model sources are read-only for refresh but their **displayed result** can be
read.

**Step 1 — inventory PivotTables.** Call `pivot_read` with `action="tables-list"`.

**Step 2 — refresh one table.** Call `pivot_refresh` with `action="refresh-table"`:

```json
{ "action": "refresh-table", "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1" }
```

**Step 3 — read the rendered result.** Call `pivot_read` with `action="result-read"`:

```json
{ "action": "result-read", "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1", "offset": 0, "limit": 100 }
```

**Expected result.** A `PivotRefreshResult` (with a truth-signal `refreshed` status) and a paged
`PivotResultPage` of the rendered output.

**How to verify.** `result-read` returns the current values; the refresh result reports the
before/after `RefreshDate`.

**Limitations / safety.** External/OLAP/Data-Model sources return `UNSUPPORTED_PIVOT_SOURCE` for
refresh; their displayed result is still readable via `result-read`. `Workbook.RefreshAll` is never
used. `result-read` is paged; a stale `cursor_version` returns `STALE_CURSOR`.

## Workflow 9: Export a large dataset for downstream analysis — Illustrative

**Goal.** Stream a large range, table or PivotTable result to a CSV/JSONL file with bounded memory.

**Setup.** A workbook open read-only.

**Step 1 — export a range, table or pivot result.** Call `data_export`:

```json
{ "workbook_id": "<workbook_id>", "source_kind": "table", "table": "Sales", "fmt": "csv" }
```

Use `source_kind="range"` with `sheet` + `address`, or `source_kind="pivot-result"` with
`pivot_table`.

**Step 2 — read the returned metadata.** The result reports `path`, `rows`, `bytes`, `sha256` and
`truncated`/`complete`.

**Expected result.** A `DataExportResult` pointing at a new artifact file; the source workbook is
unchanged.

**How to verify.** Open the artifact and compare the row count with the source; check the SHA-256.

**Limitations / safety.** Returns metadata only (never the data inline); `limit_rows` may truncate.
The MCP-transport run is validated in-process only. The export is non-destructive and idempotent (a
new artifact per call).

## Workflow 10: Produce an independent copy (Save As) — Tested (integration)

**Goal.** Write a verified copy to a new (or the original) path without rebinding the session.

**Setup.** A workbook open writable.

**Step 1 — save a copy.** Call `workbook_save_copy`:

```json
{ "workbook_id": "<workbook_id>", "path": "C:/data/report_copy.xlsm", "overwrite_existing": false, "overwrite_source": false }
```

**Expected result.** A `SaveAsResult` with `outcome="ok"` only after a final size + SHA-256 verify,
plus a single `<target>.bak` if a previous target existed.

**How to verify.** The reported size and SHA-256 match the written file; the session is still bound to
its working copy.

**Limitations / safety.** An existing target needs `overwrite_existing`; overwriting the original needs
`overwrite_source` plus a matching source hash (`STALE_TARGET` otherwise). The open working copy as
target is always refused.

## Workflow 11: Recover a wedged Excel — Tested (integration)

**Goal.** Recover from a hung macro or call without restarting the server.

**Setup.** None beyond a running server.

**Step 1 — check health.** Call `excel_health`. If it shows `needs_recovery` (or
`worker_state="running/deadline_exceeded"`), continue.

**Step 2 — recover the owned process.** Call `excel_recover` with `confirm=true`:

```json
{ "reason": "agent_directed", "confirm": true }
```

**Step 3 — reopen the workbook.** The old `workbook_id` is invalid; call `workbook` `open` again.

**Expected result.** A recovery report; a fresh hidden Excel instance; the originals untouched.

**How to verify.** `excel_health` returns to a healthy state; a new `workbook` `open` succeeds.

**Limitations / safety.** A `macro_run` hang/timeout recovers the session **automatically** — just
reopen the workbook. `excel_recover` is pid-scoped and originals-safe but **discards open working
copies**; it never terminates a live pid on a deadline. If recovery is refused
(`recovery_possible=false`), restart the server process.

## Workflow 12: Discover policy and capabilities before automating — Illustrative

**Goal.** Adapt behaviour to the server's access mode and capabilities before automating.

**Setup.** None.

**Step 1 — read the environment.** Call `xlvba_environment` (no parameters). It reports the access
mode, policy, build (always `null`) and capabilities without touching Excel.

**Step 2 — confirm liveness.** Call `server_health` for a pure liveness probe.

**Expected result.** An `EnvironmentReport` (access mode, allowed roots, macro policy, capabilities)
and a `{"status": "ok", ...}` health payload.

**How to verify.** The reported access mode matches `EXCEL_VBA_MCP_ACCESS_MODE`; the capabilities list
matches what you plan to use.

**Limitations / safety.** `xlvba_environment` and `server_health` never touch Excel, so they work even
when Excel is unavailable. `excel.build` is always `null` (never fabricated). This end-to-end discovery
flow is illustrative; the tools themselves are implemented and unit-tested.

## Combining tools safely

### Ordering rules

1. **Open before anything else.** Every tool that takes a `workbook_id` needs an open working copy.
   Start with `workbook` (`open`) and keep the returned `workbook_id`.
2. **Inspect with `mode="inspect", read_only=True`.** Use this for every read-only workflow.
3. **Snapshot before a mutation.** Call `snapshot_create` after opening writable and before the first
   mutating call.
4. **Read before you write.** Use `vba_read` (`read-component`) for the `source_hash`,
   `table_read` (`get`) for the exact table name, `pivot_read` (`table-get` / `fields-list`) for
   PivotTable and field names, and `defined_name_read` before touching a defined name.
5. **One object per mutation, then re-read** the returned postcondition (`verified`,
   `postcondition_values`, `result_format`, `affected_rows`).
6. **Close deliberately.** `workbook` `close` without `save_changes` is the safe default.
7. **Never retry a blocked call blindly.** Read `context.reason` and change the precondition instead.

### Guard chain to keep in mind

| Order | Gate | Failure code |
|-------|------|--------------|
| 1 | Writable session (mutating tools) | `WORKBOOK_READ_ONLY` |
| 2 | Baseline snapshot | `SNAPSHOT_FAILED` |
| 3 | Approval (`full` default / `restricted`) | `APPROVAL_DENIED` (restricted mode) |
| 4 | Operation-specific validation (payload shape, explicit `shift`, unambiguous `where`, no ListObject intersection, supported pivot source) | `INVALID_RANGE`, `AMBIGUOUS_FILTER`, `LISTOBJECT_INTERSECTION`, `UNSUPPORTED_PIVOT_SOURCE` |
| 5 | Postcondition re-read + audit | `verified=false` in the result |

### Do and do not

| Do | Do not |
|----|--------|
| Use read-only sessions for inspection and reporting. | Assume every session must be read-only. |
| Open writable (`mode="execute"`, the default) for mutations in full mode. | Assume mutations need a grant in full mode. |
| Classify with `macro` (`classify-background`) before `macro_run`. | Run a macro without a `background_safe` classification. |
| Snapshot before any mutating workflow. | Skip the snapshot before a mutation. |
| Use the table tools for table structure. | Use `range_structural` `insert` / `range_delete` over a ListObject. |
| Keep `expected_source_hash` from `vba_read` (`read-component`). | Edit VBA without a revision token. |
| Treat `indeterminate` compile as "unknown", not "passed". | Treat a clean `vba_compile_diagnostics` as proof of compilation. |
| Read `context.reason` on every refusal. | Blind-retry a blocked macro or a denied mutation. |

## Error codes you are most likely to see

Tool failures surface as a structured error whose message is the JSON of `to_dict()`: a stable `code`
and a `context` object. The same error is logged to `stderr` and the audit trail.

| Code | Where it surfaces | Typical cause | What to do |
|------|-------------------|---------------|------------|
| `PATH_NOT_ALLOWED` | `workbook` (`open`), `vba_export_project` | Path outside `EXCEL_VBA_MCP_ROOTS`, a device/UNC path, traversal, symlink escape, or a disallowed extension. | Put the workbook inside an allowed root or extend `EXCEL_VBA_MCP_ROOTS` and restart the server. |
| `LISTOBJECT_INTERSECTION` | `range_write`, `range_structural`, `range_delete`, `range_merge` | The target intersects an Excel Table (ListObject). | Use the table tools (`table`, `table_delete_rows`, `table_delete_column`) instead. |
| `ACCESS_VBOM_DISABLED` | `vba_read`, `vba_*`, `macro`, `vba_compile_diagnostics` | VBOM trust is off. | Enable **Trust access to the VBA project object model** in Excel Trust Center → Macro Settings, restart Excel, retry. |
| `APPROVAL_DENIED` | `workbook` (writable/execute), any mutating tool, `macro_run`, `test_suite` | The server runs in `restricted` mode and no allowlist entry matches. | Unset `EXCEL_VBA_MCP_ACCESS_MODE` (or set `full`) for full access. |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_run` (status `blocked`) | The procedure is not classified `background_safe` for the current workbook hash and VBA source revision. | The macro was **not** started. Classify it with `macro` (`classify-background`) and retry. |
| `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_run`, `macro` (`classify-background`), `test_suite` | The procedure shows UI, is heuristic, or was classified safe while unsafe. | Remove the interactive UI or leave it unclassified; do not bypass the preflight. |
| `REVISION_CONFLICT` | `vba_replace_component_source`, `vba_insert_code` | A stale `expected_source_hash`. | Re-read the component with `vba_read` (`read-component`) and pass the current hash. |
| `NOT_FOUND` | `shape`, `control` | A shape or control addressed by `name` does not exist. | List the objects (`shape` / `control` `list`) and use an exact name. Do not retry. |
| `WORKBOOK_READ_ONLY` | Any mutating tool | The session was opened with `read_only=true`. | Reopen writable (`workbook` `open` without `read_only`). |
| `SNAPSHOT_FAILED` | Any mutating tool | No successful snapshot exists for the workbook. | Call `snapshot_create` first. |
| `UNSUPPORTED_PIVOT_SOURCE` | `pivot_refresh`, `pivot_field` | The cache source is external, OLAP or Data Model. | Read-only inspection and `pivot_read` `result-read` still work. |
| `AMBIGUOUS_FILTER` | `table` (`update-rows`), `table_delete_rows` | A `where` filter matched zero or multiple rows. | Use explicit `indices`, or refine the filter. |
| `INVALID_RANGE` | Range tools | A malformed or oversized address, a wrong payload shape, or a merge blocked by data loss. | Fix the address/payload; pass `allow_data_loss=true` only after reviewing the merge preview. |
| `STALE_CURSOR` | `pivot_read` (`result-read`) | The paged cursor version no longer matches the live result. | Restart the read at `offset` 0. |

The full contract is in [MCP tools reference](05-mcp-tools-reference.md); a symptom-first guide is in
[Troubleshooting](10-troubleshooting.md).

Related documents: [Overview](01-overview.md), [Excel capabilities](06-excel-capabilities.md),
[MCP tools reference](05-mcp-tools-reference.md), [Configuration](03-configuration.md),
[Architecture](04-architecture.md), [Security](08-security.md),
[Troubleshooting](10-troubleshooting.md) and the [root README](../README.md).