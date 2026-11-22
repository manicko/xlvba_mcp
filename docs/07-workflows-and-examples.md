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

This document shows **real, usable workflows** for the `xlvba` MCP server. Each workflow lists
its prerequisites, the exact tool calls with JSON arguments, the expected outcome and its
limitations.

Every example uses the published MCP tool names and the parameter names from
[MCP tools reference](05-mcp-tools-reference.md). None of the examples were executed while
writing this document: they are recipes, not transcripts.

The capabilities behind these workflows are classified in
[Excel capabilities](06-excel-capabilities.md). Settings are in
[Configuration](03-configuration.md); the security and approval model is in
[Security](08-security.md).

## Conventions

| Convention | Meaning |
|------------|---------|
| `<workbook_id>` | Server-assigned identifier of an open working copy, returned by `excel_open`. It is **not** a file path. Replace it in every subsequent call. |
| JSON blocks | Show the tool **arguments** object exactly as it is sent over MCP. A block is not the full JSON-RPC envelope. |
| `xlvba_` prefix | Clients expose published names with a server prefix. In Kilo, for example, `excel_open` is called as `xlvba_excel_open`. This document uses the published names without the prefix; add your client's prefix. |
| **Approval-gated (restricted mode)** | The call is a mutating or executing operation. In the default `full` access mode it is authorized (a baseline snapshot is still required before a mutation); in `restricted` mode it is refused with `APPROVAL_DENIED` unless an in-process allowlist entry or live grant matches. See [Security](08-security.md). |
| Read-only | No workbook change; auto-allowed by the approval policy. |

Paths are Windows paths. A workbook path must be inside `EXCEL_VBA_MCP_ROOTS`
(by default the current working directory and the system temp directory); otherwise the call
fails with `PATH_NOT_ALLOWED`.

## Workflow 1: Open and inspect a workbook (read-only)

**Prerequisites.** The server is running and the workbook is inside an allowed root.

**Step 1 — open a read-only inspect session.**

```json
{
  "path": "C:/data/report.xlsm",
  "mode": "inspect",
  "read_only": true
}
```

Call `excel_open`. The expected result contains:

```json
{ "workbook_id": "1f0c…", "name": "report.xlsm", "format": "xlsm", "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm", "source_path": "C:/data/report.xlsm", "read_only": true, "has_vba": true, "excel_version": "16.0.14332", "session_id": "9b2a…", "mode": "inspect" }
```

**Step 2 — read the aggregate model (or one inventory).**

```json
{ "workbook_id": "<workbook_id>", "include_vba": false }
```

Call `workbook_get_model`. Alternatively, call `sheets_list` with the same `workbook_id`.

**Step 3 — close without saving.**

```json
{ "workbook_id": "<workbook_id>", "save_changes": false }
```

Call `excel_close`. The working copy is deleted; the original is untouched.

**Expected outcome.** A `WorkbookModel` (or `list[SheetInfo]`) followed by a `SessionStatus`
with `state="closed"`.

**Limitations.**

- `mode="inspect"` disables macros for the open; `read_only=true` makes every mutation
  unavailable (`WORKBOOK_READ_ONLY`).
- `excel_open` defaults to a writable, execution-capable session (`mode="execute"`,
  `read_only=False`); pass `mode="inspect", read_only=true` (as in Step 1) to open read-only.
- `workbook_get_model` never reads all cells and never executes VBA.

## Workflow 2: Read a range and its formulas

**Prerequisites.** A workbook opened as in Workflow 1.

**Step 1 — read typed values.**

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
```

Call `range_read_values`. Each cell is typed (`string`, `number`, `boolean`, `date`, `error`,
`empty`), and the response carries `rows`, `cols` and `truncated`.

**Step 2 — read the formulas in the same window.** Use the same arguments object as Step 1
and call `range_read_formulas`:

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
```

Each cell has `formula2` and, when different, `formula`. Prefixes such as `_xlfn`, `_xlws` and
`_xlpm` are preserved and listed in `notes`.

**Expected outcome.** A `RangeValues` grid and a `RangeFormulas` grid for the same window.

**Limitations.**

- Reads are bounded: 100,000 cells per target and 50,000 cells per response. Read a larger
  area in chunks with `offset`/`limit`.
- `range_read_values` returns the values Excel currently holds. It does not force a
  recalculation and does not evaluate formulas.

## Workflow 3: Inspect VBA and export the project

**Prerequisites.** A workbook opened read-only; Excel Trust Center must trust access to the VBA
project object model (VBOM), and the project must not be locked.

**Step 1 — project summary.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `vba_project_info`. The result reports the project name, protection/lock state, component
inventory and references.

**Step 2 — component inventory.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `vba_list_components`. Each entry has a type (`standard_module`, `class_module`,
`user_form`, `document_module`), a line count and a normalized source hash.

**Step 3 — read one component.**

```json
{ "workbook_id": "<workbook_id>", "component": "Module1" }
```

Call `vba_read_component`. Keep the returned `source_hash` — it is the revision token required
by the guarded VBA editing tools.

**Step 4 — export the whole project.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `vba_export_project`. Without `out_dir` the export goes to a temporary directory; a custom
`out_dir` must live inside the `snapshots` or `artifacts` root. The result is a manifest with
per-file hashes.

**Expected outcome.** A `VbaProjectInfo`, a component list, one `VbaComponentSource` and a
`VbaExportManifest`.

**Limitations.**

- Without VBOM trust the tools return `ACCESS_VBOM_DISABLED`; a locked project returns
  `VBA_PROJECT_LOCKED`.
- Export never modifies the workbook and never removes Document modules.

## Workflow 4: Snapshot before a change

**Prerequisites.** A workbook is open (read-only is enough for `snapshot_create`).

**Step 1 — create an immutable snapshot.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `snapshot_create`. Open workbooks are captured with `Workbook.SaveCopyAs`; closed workbooks
are copied. The result is a manifest containing `snapshot_id`, `created_at`, `source_path`,
`stored_path`, `sha256`, `excel_version`, `has_vba` and `model_version`.

**Step 2 — list snapshots.**

```json
{}
```

Call `snapshot_list` (no parameters).

**Step 3 — compare two snapshots.**

```json
{
  "snapshot_id_a": "snap-…-a",
  "snapshot_id_b": "snap-…-b"
}
```

Call `snapshot_diff`. The result reports `file_identical`, `added`, `removed`, `changed` and
`notes`. It compares manifest metadata and file hashes only, never a binary diff.

**Expected outcome.** A snapshot manifest, a list of manifests and a `SnapshotDiff`.

**Limitations.**

- `snapshot_restore` is a **mutating** operation (authorized in `full` mode; `APPROVAL_DENIED` in
  `restricted` mode). It closes the current copy without saving and reopens the baseline; external
  macro side effects are not undone.
- Snapshots are immutable; there is no delete tool.

## Workflow 5: Inspect tables and read rows

**Prerequisites.** A workbook opened as in Workflow 1.

**Step 1 — list every Excel Table.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `tables_list`. Each `TableInfo` reports the sheet, range, header/body/total ranges, style,
row count and columns.

**Step 2 — read one table.**

```json
{ "workbook_id": "<workbook_id>", "name": "Sales" }
```

Call `table_get` with the table name or display name.

**Step 3 — read a bounded slice of data rows.**

```json
{ "workbook_id": "<workbook_id>", "name": "Sales", "offset": 0, "limit": 50 }
```

Call `table_read_rows`. `offset` is a 0-based data-row index; `limit` caps the returned rows.
Values are typed and the response has a `truncated` flag.

**Expected outcome.** A table inventory, one `TableInfo`, and a `TableRows` payload.

**Limitations.**

- Lists are full and bounded; there is no pagination.
- `table_read_rows` is additionally capped at 2,000 cells per response.
- Row mutations are authorized in `full` mode (`APPROVAL_DENIED` in `restricted` mode); see
  Workflow 7.

## Workflow 6: Inspect PivotTables

**Prerequisites.** A workbook opened read-only. Only database/Table-backed PivotTables are
refreshable; external, OLAP and Data Model sources are read-only.

**Step 1 — inventory PivotTables and caches.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `pivot_tables_list`. The inventory is a worksheet sweep; shared caches are reported once
with their affected tables.

**Step 2 — read one PivotTable.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1" }
```

Call `pivot_table_get`. The result includes `TableRange1`/`TableRange2`, the data-body range, the
row/column/page/hidden/data field layout and the source type.

**Step 3 — list its fields.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1" }
```

Call `pivot_fields_list`. Each field reports orientation, 1-based position within the
orientation, aggregation and manual filter state.

**Step 4 — read one field's filter state.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1", "field": "Region" }
```

Call `pivot_field_get_filters`. The result reports `all_items_visible`, visible and hidden
`PivotItems`, the classified filter kind and any detected `PivotFilters`.

**Expected outcome.** A `PivotInventory`, a `PivotTableInfo`, a field list and a
`PivotFieldFilters`.

**Limitations.**

- Creating or deleting PivotTables/PivotCaches is not supported.
- `PivotFilters.Add` (label/value/date/top-bottom) is unsupported on the validated build; only
  manual visible-items filters can be applied.
- Refresh, layout, aggregation, properties and filters are authorized in `full` mode
  (`APPROVAL_DENIED` in `restricted` mode).

## Workflow 7: Mutating workflow (full access mode)

In the default `full` access mode, a mutation sequence runs without per-action grants: open a
writable working copy (the default), create a baseline snapshot, then one or more guarded
mutations. Each guarded mutation still re-reads a postcondition and writes an audit record. In
`restricted` mode the approval gate refuses every mutating call with `APPROVAL_DENIED`.

**Step 1 — open a writable working copy (the default).**

```json
{ "path": "C:/data/report.xlsm" }
```

Call `excel_open` (omitting `mode`/`read_only` opens `mode="execute", read_only=False`). In `full`
mode this is authorized and a working copy is created; in `restricted` mode it is `EXECUTING` and
returns `APPROVAL_DENIED`.

**Step 2 — create a baseline snapshot.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `snapshot_create`. This is read-only and allowed once a writable session exists.

**Step 3 — write values.**

```json
{
  "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "B2:C3",
  "payload": {
    "rows": 2, "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "North" } ],
      [ { "kind": "number", "value": 2 }, { "kind": "string", "value": "South" } ]
    ]
  }
}
```

Call `range_write_values`. In `full` mode it succeeds with `verified=true` and a re-read
`postcondition_values`. A string beginning with `=` is stored as text, not a formula.

**Step 4 — add a worksheet.**

```json
{ "workbook_id": "<workbook_id>", "name": "Scratch", "sheet_type": "worksheet" }
```

Call `sheet_add`. Succeeds in `full` mode.

**Step 5 — append table rows.**

```json
{
  "workbook_id": "<workbook_id>",
  "name": "Sales",
  "rows": [
    { "Region": "North", "Amount": 120 },
    { "Region": "South", "Amount": 80 }
  ]
}
```

Call `table_append_rows`. Succeeds in `full` mode. The table must already exist; the table
row/column structure is preserved.

**Expected outcome (full access mode).** Each mutation returns a result with `verified=true` and a
postcondition re-read (`postcondition_values`, `result_format`, `affected_rows`, and so on).

**Persistence.** Edits live only in the working copy; the original is never written. To persist,
call `excel_close` with `save_changes=true` (writes the working copy) or `excel_save`. Because
`excel_close` deletes the working copy and a later `excel_open` re-copies the untouched original, a
literal "close, then reopen" does **not** show prior edits. To return to an earlier state, use
`snapshot_create` → `excel_close` → `snapshot_restore`.

**Restricted mode / embedder configuration.** In `restricted` mode the server builds
`ApprovalPolicy(mode="restricted")` with an empty allowlist and exposes no grant or allowlist tool,
so every `MUTATING`/`EXECUTING` call is refused with `APPROVAL_DENIED`. An in-process embedder that
needs mutations in restricted mode constructs the services with an `ApprovalPolicy` allowlist entry
or a live grant matching the canonical workbook path (or SHA-256), the action name and, for
`macro_run`, the procedure. The policy implementation is `src/excel_vba_mcp/security/approval.py`.

**Limitations.**

- Every mutation requires a successful snapshot first (`SNAPSHOT_FAILED` otherwise).
- `range_write_values` on a range that changes a ListObject's row/column structure is rejected
  with `LISTOBJECT_INTERSECTION`.
- `range_merge` can discard non-anchor data unless `allow_data_loss=true` is passed after
  reviewing the preview.

## Workflow 7b: Run a macro in full access mode (classify → run → verify)

`macro_run` requires an execute-mode, writable session and a `background_safe` classification bound
to the current workbook hash and VBA source revision. In the default `full` mode no per-run approval
grant is needed.

**Step 1 — open (default execute/writable) and snapshot.**

```json
{ "path": "C:/data/report.xlsm" }
```

Call `excel_open`, then `snapshot_create` with the returned `workbook_id`.

**Step 2 — list macros and classify the target as background-safe.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `macro_list` to confirm the exact `Module.Procedure` name and current `background_safe` state.
Then classify:

```json
{ "workbook_id": "<workbook_id>", "procedure": "Module1.Stamp", "background_safe": true }
```

Call `macro_classify_background`. An interactive/heuristic procedure is rejected with
`INTERACTIVE_UI_NOT_SUPPORTED`; a stale `expected_project_source_hash` is rejected with `STALE_TARGET`.

**Step 3 — run the macro, then verify a cell it writes.**

```json
{ "workbook_id": "<workbook_id>", "macro_name": "Module1.Stamp", "args": [] }
```

Call `macro_run`. Then read the cell the macro is expected to change:

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1", "offset": 0, "limit": 1 }
```

Call `range_read_values` and confirm the expected value.

**Step 4 — edit again, then run a second macro (edit→macro→edit→macro).**

Write a new value with `range_write_values` (or `table_update_rows`), then call `macro_run` again
for a second procedure. Each guarded step is independent; the sequence is exactly what a
declarative `test_run_suite` expresses with `macro_run` steps carrying an optional `background_safe`
field.

**Expected outcome.** Each `macro_run` returns `success` (or `blocked` with
`BACKGROUND_SAFETY_NOT_VERIFIED` if the classification was invalidated by a VBA source change), and
each edit returns a verified postcondition.

**Limitations.**

- If the VBA source changes between classification and run, the classification is invalidated and
  the macro is `blocked` with `BACKGROUND_SAFETY_NOT_VERIFIED`; re-classify after re-reading.
- The R13 token scan is a heuristic veto, not static proof: dynamic `CallByName`/`ExecuteGlobal`/
  `Eval` indirection and transitive UI are not fully covered. The veto list does include
  `AppActivate`, `Application.Visible`, `.Activate` and `WindowState`, so direct use is caught;
  UI reached only indirectly can still evade the substring scan.
- External macro side effects are never rolled back by a snapshot.

## Workflow 8: Declarative test suite

**Prerequisites.** An open workbook (read-only inspection is enough to open it). The suite opens
its own disposable copies in execute mode; in the default `full` access mode this is authorized,
while `restricted` mode returns the run as `blocked` (`APPROVAL_DENIED`). A `macro_run` step on an
unclassified procedure is `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`) in either mode.

**Step 1 — submit a closed-vocabulary suite.**

```json
{
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
          { "type": "sheet_visibility", "sheet": "Scratch", "expected": "visible" },
          { "type": "cell_value_equals", "sheet": "Scratch", "cell": "B2", "expected": 7 }
        ]
      },
      {
        "id": "table-rows",
        "steps": [
          { "action": "table_append_rows", "table": "Sales", "rows": [ { "Region": "North", "Amount": 120 } ] }
        ],
        "assertions": [
          { "type": "table_exists", "table": "Sales" },
          { "type": "table_row_count", "table": "Sales", "value": 4 }
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
          { "type": "macro_status", "expected": "success" },
          { "type": "read_cell", "sheet": "Sheet1", "cell": "A1", "expected": "ok" }
        ]
      }
    ]
  }
}
```

Call `test_suite_run` with this object as the `suite` argument (`test_run_suite` is equivalent
and takes the same shape). The vocabularies are closed:

The closed vocabularies contain **37 actions** and **29 assertions**, defined in
`src/excel_vba_mcp/domain/test_spec.py` and summarized in
[MCP tools reference](05-mcp-tools-reference.md). The examples above use a subset:
`sheet_add`, `range_write_values`, `table_append_rows`, `vba_compile`, `macro_run`,
`sheet_exists`, `sheet_visibility`, `cell_value_equals`, `table_exists`, `table_row_count`,
`compile_status`, `macro_status`, `read_cell`.

**Step 2 — read the persisted run.**

```json
{ "run_id": "<test_run_id>" }
```

Call `test_get_run` with the `test_run_id` from the `TestRunResult`. It returns per-case and
per-step statuses without touching Excel.

**Expected outcome.** A `TestRunResult` with one entry per case and step, plus a persisted run
under `runs/`.

**Limitations.**

- `vba_replace_component_source` and `vba_insert_code` exist in the action enum but are
  **rejected by the runner**: a declarative step cannot carry the `expected_source_hash` revision
  token. Use the guarded MCP tools for those edits.
- `macro_status` and `compile_status` assert against the case's last macro run and compile probe.
- A failed `vba_compile` step blocks later `macro_run` steps in the same case; an
  `indeterminate` compile is not a failure.
- A `macro_run` step may carry an optional data-only `background_safe` field; the runner classifies
  the procedure through the guarded `classify_background` path before dispatch.
- In `restricted` mode every case is `blocked` (`APPROVAL_DENIED`); an unclassified macro step is
  `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`) in either mode.

## Workflow 9: Diagnose compile state

**Prerequisites.** A workbook opened read-only; VBOM trust is required for any VBA evidence.

**Step 1 — request a compile result.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `vba_compile`. On the validated build the result is fail-closed:

```json
{
  "workbook_id": "1f0c…", "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": { "is_compile": false, "vbom_available": true, "protection": 0, "locked": false, "component_count": 3, "broken_references": [] },
  "command_discovery": { "route_usable": false },
  "notes": ["NOT a compile: non-activating VBA compilation is unavailable on this build.", "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"]
}
```

**Step 2 — re-read the cached diagnostics.**

```json
{ "workbook_id": "<workbook_id>" }
```

Call `vba_get_compile_diagnostics`. It returns the same diagnostic data without re-collecting.

**Expected outcome.** An `indeterminate` status with a structured reason and partial diagnostics.

**Limitations.**

- `vba_compile` never returns `compiled` or `compile_error`; `diagnostics.is_compile` is always
  `false`. A clean `indeterminate` result is **not** proof that the project compiles.
- Possible reasons: `NON_ACTIVATING_COMPILE_UNAVAILABLE`, `ACCESS_VBOM_DISABLED`,
  `VBA_PROJECT_LOCKED`.
- No UI is shown or activated.

## Workflow 10: Handle errors

Tool failures surface as a structured MCP tool error whose message is the JSON of `to_dict()`: a
stable `code` and a `context` object (for example `{"code": "APPROVAL_DENIED", "context": {…}}`).
The same error is logged to `stderr` and recorded in the audit trail. The table below shows the
codes most likely to appear and the correct response.

| Code | Where it surfaces | Typical cause | What to do |
|------|-------------------|---------------|------------|
| `PATH_NOT_ALLOWED` | `excel_open`, `vba_export_project` | Path outside `EXCEL_VBA_MCP_ROOTS`, a device/UNC path, traversal, symlink escape, or a disallowed extension. `context.reason` names the category (`outside_roots`, `extension_not_allowed`, `unc_denied`, `symlink_escape`, `root_mismatch`, `traversal`, `empty_path`, `device_path_denied`). | Put the workbook inside an allowed root or extend `EXCEL_VBA_MCP_ROOTS`; use a plain path and an allowed extension. See [Configuration](03-configuration.md). |
| `LISTOBJECT_INTERSECTION` | `range_write_values`, `range_write_formulas`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge` | The target intersects an Excel Table (ListObject). `context.tables` lists the intersecting tables. | Use the Tables tools (`table_append_rows`, `table_update_rows`, `table_add_column`, and so on) instead of structural range operations. |
| `ACCESS_VBOM_DISABLED` | `vba_*`, `macro_list`, `vba_compile` | Excel cannot read the VBA project object model; Trust Center VBOM trust is off. | Enable **Trust access to the VBA project object model** in Excel Trust Center → Macro Settings, close Excel and retry. Corporate policy may block this. |
| `APPROVAL_DENIED` | `excel_open` (writable/execute), any mutating tool, `macro_run`, `test_run_suite` | The server runs in `restricted` mode (`EXCEL_VBA_MCP_ACCESS_MODE=restricted`) and no allowlist entry or live grant matches. | Unset `EXCEL_VBA_MCP_ACCESS_MODE` (or set `full`) for the default full-access behavior; a restricted server exposes only read-only work over MCP. |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_run` (in `MacroRunResult.error_code`, status `blocked`) | The procedure is not classified `background_safe` for the current workbook hash and VBA source revision. | The macro was **not** started. Classify it with `macro_classify_background` (re-classify after any VBA source change) and retry. Do not retry blindly. |
| `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_run`, `macro_classify_background`, `test_run_suite` | The procedure references interactive UI, is heuristic, or was classified `background_safe` while unsafe. | Inspect the procedure and remove the interactive UI, or leave it unclassified. Do not bypass the preflight. |
| `COM_BUSY` | Any mutating tool; `vba_replace_component_source`/`vba_insert_code` | Genuine COM contention (server busy) or a VBA `source_hash` revision conflict (`context.reason="revision_conflict"`). Approval denial is a separate code (`APPROVAL_DENIED`). | Read `context.reason`. Retry after the operation completes; for a revision conflict, re-read and pass the current `expected_source_hash`. |
| `WORKBOOK_READ_ONLY` | Any mutating tool | The session was opened with `read_only=true`. | Reopen writable (`excel_open` without `read_only`) or perform read-only work only. |
| `SNAPSHOT_FAILED` | Any mutating tool | No successful snapshot exists for the workbook. | Call `snapshot_create` first. |
| `UNSUPPORTED_PIVOT_SOURCE` | `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_*` | The cache source is external, OLAP or Data Model. | Refresh or reconfigure is impossible for that source; read-only inspection still works. |
| `AMBIGUOUS_FILTER` | `table_update_rows`, `table_delete_rows` | A `where` filter matched zero or multiple rows. | Use explicit `indices`, or refine the filter until it matches exactly one row. |
| `INVALID_RANGE` | Range tools | A malformed or oversized address, a wrong payload shape, or a merge blocked by data loss. | Fix the address/payload; pass `allow_data_loss=true` only after reviewing the merge preview. |

Additional codes declared by the server include `VBA_PROJECT_LOCKED`, `MISSING_REFERENCE`,
`RPC_E_CALL_REJECTED`, `COMPILE_ERROR`, `COMPILE_INDETERMINATE`, `MACRO_BLOCKED`,
`MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `EXCEL_HUNG`, `PROTECTED_TARGET`,
`STALE_TARGET`, `REFRESH_FAILED` and `MODAL_UI_BLOCKED`. The full contract is in
[MCP tools reference](05-mcp-tools-reference.md).

## Combining tools safely

### Ordering rules

1. **Open before anything else.** Every tool that takes a `workbook_id` needs an open working
   copy. Start with `excel_open` and keep the returned `workbook_id`.
2. **Inspect with `mode="inspect", read_only=true`.** Use this for every read-only workflow. It
   disables macros and makes accidental mutation impossible.
3. **Snapshot before a mutation.** Call `snapshot_create` after opening writable and before the
   first mutating call. Without it, every mutation fails with `SNAPSHOT_FAILED`.
4. **Read before you write.** Use `vba_read_component` for the `source_hash`, `table_get` for
   the exact table name, `pivot_table_get`/`pivot_fields_list` for the PivotTable and field
   names, and `names_list` before touching a defined name.
5. **One object per mutation.** Change one table, range, sheet or field at a time, then inspect
   the returned postcondition (`verified`, `postcondition_values`, `result_format`,
   `affected_rows`).
6. **Re-read after a mutation.** Do not assume success from the absence of an error; use the
   returned postcondition or a fresh read.
7. **Close deliberately.** `excel_close` without `save_changes` is the safe default. Saving is a
   separate, guarded decision.
8. **Never retry a blocked call blindly.** A blocked macro or a denied mutation will be blocked
   again. Read the `context.reason` and change the precondition instead.

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
| Classify with `macro_classify_background` before `macro_run`. | Run a macro without a `background_safe` classification. |
| Snapshot before any mutating workflow. | Skip the snapshot before a mutation. |
| Use the Tables tools for table structure. | Use `range_insert`/`range_delete` over a ListObject. |
| Keep `expected_source_hash` from `vba_read_component`. | Edit VBA without a revision token. |
| Treat `indeterminate` compile as "unknown", not "passed". | Treat a clean `vba_compile` as proof of compilation. |
| Read `context.reason` on every refusal. | Blind-retry a blocked macro or a denied mutation. |

Related documents: [Overview](01-overview.md), [Excel capabilities](06-excel-capabilities.md),
[MCP tools reference](05-mcp-tools-reference.md), [Configuration](03-configuration.md),
[Security](08-security.md), [Troubleshooting](10-troubleshooting.md) and the
[root README](../README.md).