---
id: xlvba-excel-capabilities
domain: 06-excel-capabilities
tags:
  - excel
  - capabilities
  - coverage
  - mcp
  - vba
  - pivottables
related:
  - xlvba-overview
  - xlvba-mcp-tools-reference
  - xlvba-workflows-and-examples
  - xlvba-configuration
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This document answers one question precisely: **what can an AI agent actually do in
Excel through `xlvba`?**

It is a capability map, not a tool manual. For each Excel object-model area it states which
operations exist, which MCP tools implement them, how the operation is performed internally,
and what is deliberately absent. The exhaustive per-tool contract (parameters, DTOs, errors)
is in [MCP tools reference](05-mcp-tools-reference.md); this document classifies the surface
and its limits.

Two facts frame everything below:

1. `xlvba` drives a **single hidden, non-activating Excel instance** through COM on a single
   STA worker thread. There is no GUI, no window activation and no keystroke input.
2. The server runs in **full access mode** by default: mutating and executing operations are
   authorized without an in-process grant. Setting `EXCEL_VBA_MCP_ACCESS_MODE=restricted` opts
   into deny-by-default approval, where those operations are refused with `APPROVAL_DENIED`.
   "Implemented" and "permitted over MCP" still differ (a baseline snapshot is required before a
   mutation, and `macro_run` needs a `background_safe` classification), so every row below keeps
   them separate.

## How to read this document

### Coverage statuses

| Status | Meaning |
|--------|---------|
| **Fully supported** | `xlvba` has an implemented, bounded path for the operation, with validation and a postcondition read where the operation mutates. For mutating rows this means "implemented and guarded", not "freely callable over MCP". |
| **Partially supported** | Some sub-operations or variants work while others are rejected or out of scope, or the result is inherently limited (for example an always-indeterminate compile, or manual PivotTable filters only). |
| **Not supported** | No tool exists for the capability. Nothing in the server exposes it. |
| **Not verified** | The path is allowed or inferred, but the behaviour is not validated on the target Excel build (validated build: Excel 16.0.14332, `.xlsm`). |

Coverage describes the **Excel object-model surface**. It does not mean the operation is
unconditionally permitted through MCP. In the default `full` access mode mutating and executing
operations are authorized (a snapshot is still required before a mutation); in `restricted` mode
they are denied by default. The Limitations column of each row says so explicitly. The security
model is in [Architecture](04-architecture.md).

### Operation kinds

| Kind | Meaning | MCP approval class |
|------|---------|--------------------|
| READ | Returns information without changing workbook state. | `read-only`, auto-allowed. |
| CREATE | Adds a new object, row or rule. | `mutating`, guarded. |
| UPDATE | Changes an existing object, value or property. | `mutating`, guarded. |
| DELETE | Removes an object, row or cell content. | `mutating`, guarded. |
| EXECUTE | Runs code (a macro or a declarative suite). | `executing`, guarded and additionally gated. |

### Guard chain in one line

Every CREATE/UPDATE/DELETE/EXECUTE tool runs the same order: **writable session → snapshot →
approval → operation-specific checks → postcondition re-read + audit**. The full chain and
its rationale are in [Architecture](04-architecture.md).

## Capability matrix

| Excel capability | Supported operations | MCP tools | Coverage | Limitations |
|------------------|----------------------|-----------|----------|-------------|
| Application management (hidden instance, health, recovery) | READ | `server_health`, `excel_health`, `excel_list_workbooks`, `excel_get_status` | Fully supported | `server_health` never touches COM; `excel_health` reports `ok` / `needs_recovery` / `unavailable` without forcing a restart. No `.Activate`, `.Select`, `ShowWindow` or `SendKeys` exists. Recovery is implicit: a later call rebuilds the owned instance. |
| Workbook open (existing file, inspect/execute, read-only) | READ / EXECUTE | `excel_open` | Fully supported | The original is copied into `work/` and only the copy is opened. The default open is writable and execution-capable (`mode="execute"`, `read_only=False`), authorized in `full` mode and refused with `APPROVAL_DENIED` in `restricted` mode. Path must be inside `EXCEL_VBA_MCP_ROOTS` (`PATH_NOT_ALLOWED`). |
| Workbook close and save | UPDATE | `excel_close`, `excel_save` | Fully supported | `excel_close` without saving is allowed; `save_changes=true` and `excel_save` are mutating (authorized in `full` mode; `APPROVAL_DENIED` in `restricted` mode). Close deletes the working copy; edits live only in the working copy, so use snapshot-restore to persist and return to an earlier state. Changes are never saved implicitly. |
| Workbook creation (new workbook) | CREATE | — | Not supported | There is no new-workbook tool. `excel_open` only opens an existing file; a new workbook cannot be created through MCP. |
| Workbook model (aggregate read) | READ | `workbook_get_model` | Fully supported | Aggregates sheets, tables, names, shapes, controls, optional VBA and optional pivots with a per-section status. It never reads all cells and never executes VBA. When VBOM is unavailable the other sections are still returned. |
| Worksheets: list and get | READ | `sheets_list`, `sheet_get` | Fully supported | Enumerates the `Sheets` collection (worksheets, chart, macro, dialog). Reports `code_name` read-only, `protected`, `table_count` and `used_range` for worksheets. |
| Worksheets: add, rename, copy, delete, visibility | CREATE / UPDATE / DELETE | `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` | Partially supported | All mutating (authorized in `full` mode; `APPROVAL_DENIED` in `restricted` mode). `sheet_add` supports `worksheet` only. `sheet_copy` is an in-workbook copy placed after the last sheet; no cross-workbook copy and no arbitrary position. `sheet_delete` refuses to leave zero visible sheets; `sheet_set_visibility` refuses to hide the last visible sheet. No sheet move/reorder tool. |
| Cells and ranges: read values | READ | `range_get`, `range_read_values` | Fully supported | Bounded windows (`offset` 0-based, `limit`); hard ceilings: 100,000 cells per target and 50,000 cells per response. Oversized requests are rejected, not silently clipped; `truncated=true` marks a clipped window. `range_get` reports merged-cell geometry (`is_merged`, `merged_area`): a target is a single merged block only when every cell belongs to the same `MergeArea`, sub-range/interior addresses resolve to that anchor area, and `is_merged` is `false` again after `range_unmerge` (F1). |
| Cells and ranges: write values | CREATE / UPDATE | `range_write_values` | Fully supported | Guarded mutation. Payload shape must match the target exactly. A string beginning with `=` is stored as **text**, not a formula. Cell writes inside a table are allowed only when the table row/column structure is unchanged. |
| Cells and ranges: read formulas | READ | `range_read_formulas` | Fully supported | `Formula2` primary, `Formula` fallback. `_xlfn` / `_xlws` / `_xlpm` prefixes are returned verbatim and flagged in `notes`. |
| Cells and ranges: write formulas | UPDATE | `range_write_formulas` | Fully supported | Guarded mutation. Every entry must be a formula string beginning with `=`; constants belong to `range_write_values`. Same ListObject structure rule as value writes. |
| Cells and ranges: copy, insert, delete | CREATE / UPDATE / DELETE | `range_copy`, `range_insert`, `range_delete` | Partially supported | Same workbook only, no clipboard. `range_insert`/`range_delete` require an explicit `shift` (`right`/`down`, `left`/`up`); never guessed. Protected sheets, ListObject intersections (`LISTOBJECT_INTERSECTION`), dynamic-array spill ranges and merged cells (unless `allow_merged`) are rejected. `range_copy` requires a shape-compatible destination and rejects merged destinations. |
| Cells and ranges: merge and unmerge | UPDATE | `range_merge`, `range_unmerge` | Partially supported | Merge previews non-anchor cells holding data and blocks with `INVALID_RANGE` unless `allow_data_loss=true`; `across=true` merges one cell per row. Unmerge acts only when a merged area actually intersects the target. Both reject ListObject and spill targets. |
| Formatting and styles | READ / UPDATE | `range_get_format`, `range_set_format` | Partially supported | Read is tri-state: a value, `mixed` (cells disagree) or `unsupported`; a mixed range is never flattened. The sparse patch exposes 11 fields: `number_format`, `font_name`, `font_bold`, `font_italic`, `font_size`, `font_color`, `fill_color`, `fill_pattern`, `alignment_horizontal`, `alignment_vertical`, `alignment_wrap_text`. No borders, row/column size, or style-name copy. |
| Data validation | READ / UPDATE | `range_get_validation`, `range_set_validation` | Partially supported | Types: list, whole number, decimal, date, time, text length, custom (`input_only` is reported but never set). `Formula1`/`Formula2` are opaque strings capped at 255 characters and never evaluated. A non-uniform target is reported (`non_uniform_validation`) instead of being flattened. |
| Conditional formatting | READ / CREATE / UPDATE / DELETE | `range_get_conditional_formatting`, `range_set_conditional_formatting` | Partially supported | Only plain `FormatCondition` types: cell value, expression, text string, blanks, no blanks, errors, no errors, time period, top 10, above average, unique values. `ColorScale`, `DataBar` and `IconSet` are reported with `supported=false` and rejected on write. `priority` is the only stable identity; priorities are sheet-wide and renumbered. |
| Named ranges: list and get | READ | `names_list`, `name_get` | Fully supported | Workbook and worksheet scopes; `RefersTo` / `RefersToLocal`, `Visible`, and a `target_kind` classification. Non-range names are classified heuristically from the string; there is no `Application.Evaluate` and no Python `eval`. |
| Named ranges: create, update, delete | CREATE / UPDATE / DELETE | `name_create`, `name_update`, `name_delete` | Fully supported | Guarded mutation. Scope is explicit (`workbook` or `worksheet` + owning sheet); duplicates and cross-scope conflicts are rejected. `name_update` preserves scope. |
| Excel tables: list, get, read rows | READ | `tables_list`, `table_get`, `table_read_rows` | Fully supported | Full bounded lists, no pagination. `table_read_rows` uses a 0-based `offset` and a `limit`; the response is additionally capped at 2,000 cells with a `truncated` flag. |
| Excel tables: append, insert, update, delete rows | CREATE / UPDATE / DELETE | `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows` | Fully supported | Guarded mutations. Append/insert go through `ListRows.Add` so table structure is preserved. `table_insert_row` uses a 1-based `at`. Update/delete accept explicit indices or an **unambiguous** `where`; ambiguity is rejected with a preview, never guessed. |
| Excel tables: add and delete columns | CREATE / DELETE | `table_add_column`, `table_delete_column` | Fully supported | Guarded. Add uses `ListColumns.Add(Position)` with a unique header; delete resolves one unique header. |
| Excel tables: create and resize | CREATE / UPDATE | `table_create`, `table_resize` | Fully supported | Guarded. Create uses `ListObjects.Add` over an A1 range or a defined name; resize uses `ListObject.Resize`. |
| PivotTables: list, get, cache, fields | READ | `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list` | Fully supported | Enumeration is a worksheet-scoped sweep because `Workbook.PivotTables` is unusable on the target build. Caches are identified by 1-based `CacheIndex`; shared caches are reported once with the affected tables. Output cells are never materialized. |
| PivotTables: refresh | UPDATE | `pivot_table_refresh`, `pivot_cache_refresh` | Partially supported | Table refresh uses `PivotTable.RefreshTable`; cache refresh uses `PivotCache.Refresh`. Only database/Table-backed caches are refreshed; external, OLAP and Data Model sources return `UNSUPPORTED_PIVOT_SOURCE`. `Workbook.RefreshAll` is never used. |
| PivotTables: layout, position, aggregation, properties | UPDATE | `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties` | Partially supported | Layout accepts hidden/row/column/page; `orientation="data"` is rejected and aggregation must be set through `pivot_field_set_aggregation`. Only eleven aggregation codes are accepted; legacy codes are rejected. `pivot_field_set_properties` is sparse (`caption`, `number_format`, `show_all_items`) and requires a unique caption. |
| PivotTables: filters | READ / UPDATE | `pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters` | Partially supported | Only manual visible-items filters (`PivotItems[i].Visible`). Label, value, date and top/bottom filters (`PivotFilters.Add`) are unsupported on the target build and return an explicit unsupported result instead of being converted. Clear scope is `field` or `table`. Page (report-filter) fields are classified from `CurrentPage` / `AllItemsVisible` (F3): an untouched page field reports `kind:"none"`, `hidden_items:[]`, matching `pivot_table_get.filter_state`; clearing an untouched page field no longer raises a false `AMBIGUOUS_FILTER`; `pivot_field_apply_filter` on a page field is rejected (`supported:false`) rather than downgraded. |
| PivotTables and PivotCaches: create or delete | CREATE / DELETE | — | Not supported | No tool creates or deletes a PivotTable or a PivotCache. Only existing objects can be inspected, refreshed and reconfigured. |
| Shapes: list, get, set on-action | READ / UPDATE | `shapes_list`, `shape_get`, `shape_set_on_action` | Partially supported | Inventory and `OnAction` assignment only. No shape creation, deletion, formatting, grouping, z-order or text editing. `on_action` is read verbatim and never executed; an empty `macro_name` clears the assignment. |
| Controls (Form Controls and ActiveX): list, get, set on-action | READ / UPDATE | `controls_list`, `control_get`, `control_set_on_action` | Partially supported | Form Controls report type, caption, linked cell and `on_action`; ActiveX reports `prog_id` and an **inferred** event procedure. `control_set_on_action` writes `OnAction` on Form Controls only; ActiveX controls are rejected because they have no `OnAction` semantics. No control creation, deletion or property editing. |
| VBA project: info, components, procedures, read, export, references, search | READ | `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code` | Fully supported | Requires VBOM trust (`ACCESS_VBOM_DISABLED` otherwise) and an unlocked project (`VBA_PROJECT_LOCKED`). Procedures are read through the VBIDE CodeModule API with heuristic fallbacks flagged. Export writes `.bas` / `.cls` / `.frm` + `.frx` under the artifact roots and returns per-file hashes; Document modules are never removed. |
| VBA editing: replace source, insert code | UPDATE | `vba_replace_component_source`, `vba_insert_code` | Partially supported | Standard and Class modules only (Document modules and UserForms are rejected). Requires a matching `expected_source_hash` revision token — the **normalized `source_hash`** from `vba_read_component`, not the snapshot export file hash (`snapshot_create.vba_exports[].export_file_sha256`) (F7) — checked before and inside the STA worker; a mismatch raises `REVISION_CONFLICT` with `reason="revision_conflict"` (F-5 taxonomy; `COM_BUSY` is reserved for genuine contention). Guarded mutation (authorized in `full` mode; `APPROVAL_DENIED` in `restricted` mode). |
| Compile diagnostics | READ | `vba_compile`, `vba_get_compile_diagnostics` | Partially supported | Fail-closed by design: on the validated build the status is always `indeterminate` with `reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"` and `diagnostics.is_compile=false`. It never returns `compiled` or `compile_error` and never shows UI. Reasons: `NON_ACTIVATING_COMPILE_UNAVAILABLE`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`. |
| Macro execution | EXECUTE | `macro_list`, `macro_run`, `macro_get_last_result`, `macro_classify_background` | Partially supported | Gates: execute mode + writable session → procedure resolution → approval (auto in `full` mode) → `background_safe` for the current workbook hash **and** VBA source revision → interactive-UI preflight (R13). `BackgroundSafetyRegistry` starts empty (fail-closed) but is populated via `macro_classify_background`; an unclassified procedure returns `BACKGROUND_SAFETY_NOT_VERIFIED`, an interactive one `INTERACTIVE_UI_NOT_SUPPORTED`. The dispatch has a 90 s soft deadline (`EXCEL_VBA_MCP_MACRO_TIMEOUT_S`) and runs with `Application.Interactive=False` / `DisplayAlerts=False` restored in `finally` (F2a, defense-in-depth only). On timeout the owned PID is inspected read-only for a modal (`MACRO_RUNTIME_ERROR`, `dismissed=false`) or not (`MACRO_TIMEOUT`); in both cases only the owned PID is terminated and the session is recovered in-process, reported as `MacroRunResult.session_recovery` (`recovered` / `needs_recovery`), with open workbooks discarded (F2b). The R13 token scan is a heuristic veto, not static proof (dynamic `CallByName`/`ExecuteGlobal`/`Eval` indirection and transitive UI not fully covered; the veto list does include `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`). No blind retry. |
| Snapshots | READ / CREATE / UPDATE | `snapshot_create`, `snapshot_list`, `snapshot_diff`, `snapshot_restore` | Fully supported | Open workbooks use `Workbook.SaveCopyAs`, closed workbooks a file copy. Snapshots are immutable, carry a manifest (`model_version: 2`) and SHA-256. Each `vba_exports` entry records `export_file_sha256` (the exported **file** byte hash, distinct from the normalized `source_hash`) (F7). `snapshot_restore` is a guarded mutation and does **not** undo external macro side effects. `snapshot_diff` compares manifest metadata and the top-level snapshot file hashes, never a binary diff. |
| Declarative test runner | EXECUTE | `test_run_suite`, `test_suite_run`, `test_get_run` | Partially supported | Closed vocabularies: 37 bounded actions and 29 read-only assertions, `extra="forbid"`; unknown actions or unbounded payloads are rejected before any Excel call. One immutable baseline and one disposable copy per case. Sequential edit→macro→edit→macro is supported via an optional data-only `background_safe` step field (the runner classifies through the guarded path before dispatch); compile-before-macro is explicit (`vba_compile` step). `vba_replace_component_source` and `vba_insert_code` steps are rejected by the runner because a declarative step cannot carry the revision hash. In `restricted` mode the run is `blocked` (`APPROVAL_DENIED`); a macro step on an unclassified procedure is `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`). |
| Charts, PivotCharts and slicers | — | — | Not supported | No tool lists, creates, formats or deletes charts, PivotCharts or slicers. A chart sheet is visible only as a `chart` entry in `sheets_list`. |
| Row/column sizing and worksheet layout (width, height, freeze panes, page setup) | — | — | Not supported | No tool exposes row height, column width, freeze panes, print/page setup or zoom. `range_get_format` covers cells only. |
| Worksheet protection | — | — | Not supported | `sheets_list` / `sheet_get` report a read-only `protected` flag. There is no protect/unprotect tool, and mutating tools reject a protected target with `PROTECTED_TARGET`. |
| Workbook calculation | — | — | Not supported | There is no tool for `Application.Calculate` or calculation properties. `range_read_values` returns the cached/calculated values Excel currently holds; nothing forces a recalculation. |
| External links, connections and queries | — | — | Not supported | No tool lists, refreshes or edits external links, data connections or Power Query definitions. |
| Import, export, print, PDF and email | — | — | Not supported | No CSV import/export, PDF export, print, email or file-conversion tool. `vba_export_project` exports VBA source only. |
| Events and event handlers | — | — | Not supported | No tool registers, lists or edits workbook/worksheet events. ActiveX event procedures are only inferred for reading. |
| Non-`.xlsm` Excel formats | READ / CREATE / UPDATE | `excel_open` and all workbook tools | Not verified | The path allowlist accepts `.xlsm .xlsb .xlsa .xlam .xltm .xltx .xlsx .xls`, but only `.xlsm` is validated on the target build. Behaviour with other extensions is not verified. |
| ActiveX event mapping | READ | `controls_list`, `control_get` | Not verified | The probable event procedure is inferred from `<sheet_code_name>_<control>_<Event>` naming. It is never click-tested and may be absent or wrong. |

## Formula support (separate from the Excel object model)

Excel formulas such as `SUM`, `VLOOKUP`, `XLOOKUP` or `FILTER` are **not** an `xlvba`
capability in the object-model sense. They are worksheet content that `xlvba` can read and
write as strings, but `xlvba` is not a calculation engine.

Three distinct operations must not be confused:

| Operation | Tool | What it touches | What it returns |
|-----------|------|-----------------|-----------------|
| Read formula text | `range_read_formulas` | `Range.Formula2` (primary) and `Range.Formula` (fallback) | A bounded grid of `FormulaCell` values. `formula2` is always present; `formula` is included only when it differs (for example an implicit-intersection `@` added by Excel). Prefixes `_xlfn`, `_xlws` and `_xlpm` are preserved verbatim and flagged in `notes`. |
| Write formula text | `range_write_formulas` | `Range.Formula2` (fallback `Range.Formula`) | `RangeMutationResult` with the re-read formula grid and a `verified` flag. Every entry must begin with `=`; a bare constant is rejected. |
| Read calculated values | `range_read_values` | `Range.Value2` | Typed cells: `string`, `number`, `boolean`, `date`, `error`, `empty`. Dates are typed only from the cell number format. |

Explicit boundaries:

1. `xlvba` does **not** provide a calculation engine. It never evaluates a formula in Python
   and never calls `Application.Evaluate`.
2. `xlvba` does **not** expose `Application.Calculate` (or `CalculateFull`, `CalculateRow` or
   any calculation property). There is no tool to force a recalculation.
3. `xlvba` does **not** guarantee recalculated values after a write. `range_write_formulas`
   writes formula strings; the values returned by a later `range_read_values` are whatever
   Excel currently holds in the working copy. If Excel has not recalculated, the cached value
   may be stale or empty.
4. Dynamic-array formulas are read in their `Formula2` form. A spill range is a protected
   target for structural operations: `range_insert`, `range_delete` and `range_merge` reject it.

Worked examples are in [Workflows and examples](07-workflows-and-examples.md), Workflows 2 and 7.

## Per-area detail

### Session and workbook lifecycle

An agent first calls `excel_open` with a path inside `EXCEL_VBA_MCP_ROOTS`. The default open is
`mode="execute", read_only=False` (writable, execution-capable). The server copies the file into
`work/`, hashes it, opens the **copy** in a hidden instance and returns a `workbook_id`.
`mode="inspect"` disables macros for the open (`AutomationSecurity = ForceDisable`); `mode="execute"`
applies `ByUI` immediately before `Open`. Neither mode changes Trust Center or registry settings.
`excel_close` closes the working copy and deletes it; `excel_save` writes it explicitly. The
original file is never modified by any tool. Edits live only in the working copy, so a literal
"close, then reopen" does not show prior edits; persist and roll back with
`snapshot_create` → `excel_close` → `snapshot_restore`.

Not exposed: workbook properties, custom document properties, workbook-level protection,
`Workbook.Close` variants, add-in management, or a new-workbook operation.

### Worksheets

Implemented against the Excel `Sheets` / `Worksheets` collections. Read tools return `SheetInfo`
(name, index, object type, visibility, `code_name`, protection, table count, used range).
Mutating tools call `Sheets.Add`, `Worksheet.Name`, `Worksheet.Copy`, `Worksheet.Delete` and
`Worksheet.Visible`. `sheet_copy` anchors the copy with `Before`/`After` to guarantee an
in-workbook copy rather than a new workbook. Not exposed: sheet move/reorder, sheet-level
protection, print settings, tab colour, and chart-sheet creation.

### Cells and ranges

Implemented against `Worksheet.Range`, `Range.Value2` and `Range.Formula2`. Reads are bounded
windows over an explicit A1 address. Writes are bulk 2-D payloads whose shape must match the
target. Structural operations use `Range.Copy`, `Range.Insert`, `Range.Delete`, `Range.Merge`
and `Range.UnMerge` with an explicit shift. Guards: protected sheet (`PROTECTED_TARGET`),
ListObject intersection (`LISTOBJECT_INTERSECTION`), dynamic-array spill range, merged cells
without `allow_merged`, and merge data loss without `allow_data_loss`. Postconditions re-read the
target and set `verified` from the actual state. Not exposed: clipboard copy/paste,
cross-workbook copy, whole-row/column operations, find/replace, autofill, or random/volatile
functions.

### Formatting and styles

Implemented against `Range.Interior`, `Range.Font`, `Range.NumberFormat` and
`Range.Alignment`. `range_get_format` returns each property as a value, `mixed` or
`unsupported` and lists the mixed property names; it never flattens a mixed range.
`range_set_format` writes only the fields explicitly present in the patch. Colour values are
integer RGB/BGR values, font size is bounded, and format strings are capped at 255 characters.
Not exposed: borders, row/column dimensions, cell styles, theme colours, number-format
localisation or conditional formats (those have their own tools).

### Data validation

Implemented against `Range.Validation`. `range_set_validation` supports the rule-bearing types
and modifies an existing rule so unspecified properties are preserved. `Formula1`/`Formula2` are
opaque strings, capped at 255 characters and never evaluated. A target whose cells do not share
one rule is reported as non-uniform rather than flattened. Not exposed: validation on tables as a
whole, input-message-only editing as a distinct operation, or cross-sheet list sources beyond an
opaque formula string.

### Conditional formatting

Implemented against `Range.FormatConditions`. Reads return rules in ascending priority with a
bounded fill/font subset. Writes add, modify or delete one rule identified by `priority`;
priorities are sheet-wide and other rules are renumbered. `ColorScale`, `DataBar` and `IconSet`
are distinct COM objects outside the bounded scope and are rejected on write. Not exposed:
managing rules through the conditional-formatting manager UI, icon sets, data bars, colour
scales, or stop-if-true chains beyond the single-rule patch.

### Named ranges

Implemented against `Workbook.Names` and `Worksheet.Names`. Scope is explicit; `RefersTo` is
stored and returned as an opaque string; `target_kind` is classified heuristically into range,
formula, constant, external, error or unknown. Not exposed: name evaluation, name-local
constants with type coercion, or name ordering.

### Excel tables (ListObject)

Implemented against `ListObjects`, `ListRows` and `ListColumns`. Reads return range, header,
data-body and total-row ranges, style, row count and columns. Row mutations use `ListRows.Add`
and `ListRows.Item(i).Delete`; column mutations use `ListColumns.Add(Position)`. Create uses
`ListObjects.Add` over an A1 range or defined name; resize uses `ListObject.Resize`. Not exposed:
table styles/formatting, total-row functions, structured-reference formulas, slicer connections
or table merging.

### PivotTables

Implemented against `Worksheet.PivotTables`, `PivotCache`, `PivotFields`, `PivotItems` and
`PivotTable.RefreshTable`. Inventory is a worksheet sweep. Refresh is table-scoped or cache-scoped
and only for database/Table sources. Field layout, position, aggregation and properties are
patched and re-read. Filters are manual visible-items only. Not exposed: creating or deleting
PivotTables/PivotCaches, changing the PivotTable source range, calculated fields/items, grouping,
slicers, timelines, `Workbook.RefreshAll`, or label/value/date/top-bottom filters.

### Shapes and controls

Implemented against `Worksheet.Shapes` and `Worksheet.Shapes OLEObjects` / Form Controls. Read
tools report kind, `on_action` verbatim, visibility, position/size and best-effort text. The only
mutation is assigning `OnAction` (Form Controls and shapes); ActiveX controls are rejected for
`OnAction` and only report an inferred event procedure. Not exposed: creating, deleting, grouping,
formatting or renaming shapes and controls.

### VBA project and source

Implemented against the VBIDE object model (`VBProject`, `VBComponents`, `CodeModule`) on the
STA worker. Read tools require VBOM trust and an unlocked project. Procedures are listed through
the CodeModule API, with heuristic fallbacks flagged `is_heuristic`. Export writes each component
to `.bas` / `.cls` / `.frm` (+ `.frx`) and returns a manifest with per-file hashes. The two
mutations replace a whole module or insert at a semantic position; both require a matching
`expected_source_hash`, and both are limited to Standard/Class modules. Not exposed: adding or
removing components, form designer edits, reference changes, or project protection changes.

### Compile diagnostics

Implemented without launching UI. `vba_compile` collects read-only evidence (VBOM availability,
protection, lock state, component inventory, broken references, a read-only VBE command discovery
probe) and returns `status="indeterminate"` with `diagnostics.is_compile=false`. On the validated
build the reason is always `NON_ACTIVATING_COMPILE_UNAVAILABLE`. `vba_get_compile_diagnostics`
returns the cached result without re-collecting. Not exposed: a real compile, a build log, or an
error list beyond broken references.

### Macros

Implemented on the STA worker through a single `Application.Run` dispatch. The gate order is
execute mode + writable session, procedure resolution from `Module.Procedure`, the approval gate
(auto in `full` mode), `background_safe` classification bound to the workbook hash and VBA source
revision, and the interactive-UI preflight. `macro_classify_background` records the classification
(ungated, bound to the source hash). Blocked macros are never started.

The dispatch is bounded by a soft deadline (`macro_timeout_s`, default **90 s**,
`EXCEL_VBA_MCP_MACRO_TIMEOUT_S`) and wrapped in a defense-in-depth guard that sets
`Application.Interactive=False` / `DisplayAlerts=False` and restores the prior values in `finally`
(F2a). This suppresses only Excel's own built-in alerts — it does **not** block `MsgBox`,
`UserForm.Show`, `FileDialog`, add-in UI or Win32 dialogs, so the BG06/R13 veto and owned-PID
termination remain the real protections. On deadline the owned PID is inspected read-only for a
modal: if present the run is `MACRO_RUNTIME_ERROR` and only the owned PID is terminated
(`dismissed=false`); otherwise it is `MACRO_TIMEOUT`. In both branches the session is then
recovered **in-process** (no server restart, no new tool) and `MacroRunResult.session_recovery`
reports `"recovered"` / `"needs_recovery"` (`null` when no recovery was attempted). Recovery
discards open workbooks, so the prior `workbook_id` is invalid and the macro is never
auto-replayed (F2b). The R13 token scan
is a heuristic veto, not static proof (dynamic `CallByName`/`ExecuteGlobal`/`Eval` indirection and
transitive UI not fully covered; the veto list does include `AppActivate`, `Application.Visible`,
`.Activate`, `WindowState`). Not exposed: running macros in
inspect or read-only sessions, or dismissing dialogs.

### Snapshots

Implemented through `storage/snapshot_store.py`. Open workbooks are captured with
`Workbook.SaveCopyAs`; closed workbooks are copied. The store is immutable (hidden staging plus
atomic replace), and each snapshot carries a manifest with SHA-256. Restore closes the current
working copy without saving and reopens from the baseline. Diff compares manifest metadata and
file hashes only. Not exposed: binary diffs, partial restore, or rollback of external macro side
effects.

### Declarative test runner

Implemented in `application/test_runner.py` on top of the closed schema in
`domain/test_spec.py`. Every case starts from its own disposable working copy of one immutable
baseline snapshot. Actions and assertions come from closed enums, so an unknown action, an
arbitrary code payload or an unbounded matrix is rejected before any Excel call. A failed
`vba_compile` step blocks later `macro_run` steps in the same case; an `indeterminate` compile is
not treated as a failure. Runs are persisted under `runs/` and read by `test_get_run`. Not
exposed: arbitrary Python, shell or COM in a step, or a declarative replacement for the guarded
`vba_replace_component_source` / `vba_insert_code` tools (those steps are rejected by design).

## Not supported and why

| Missing capability | Reason |
|--------------------|--------|
| Create or delete PivotTables / PivotCaches | Out of scope by design. The Phase 2 contract manages **existing** PivotTables and caches only; external/OLAP/Data-Model sources are rejected. |
| Charts, PivotCharts, slicers | Not implemented. No adapter or tool exists; a chart sheet is only a `chart` entry in `sheets_list`. |
| `PivotFilters.Add` (label/value/date/top-bottom filters) | Unsupported on the validated build. `pivot_field_apply_filter` returns an explicit unsupported result with `filter.supported=false` and a reason; it never silently converts to a manual filter. |
| New workbook creation | No tool. `excel_open` opens an existing file only, so a blank workbook cannot be created over MCP. |
| Range structural operations over a ListObject | Rejected with `LISTOBJECT_INTERSECTION` and routed to the Tables tools. This is a deliberate ownership boundary, not a missing feature. |
| Worksheet or workbook protection changes | No tool. Protection is read-only diagnostics; a protected target is rejected with `PROTECTED_TARGET`. |
| Forced recalculation | `Application.Calculate` is not exposed. Formula writes do not guarantee a recalculated cached value. |
| Clipboard and cross-workbook copy | `range_copy` is same-workbook and never uses the clipboard. |
| Rows/columns sizing, freeze panes, page setup | No tool. |
| External links, connections, queries, import/export, print, PDF, email | No tool. |
| Events | No tool. |
| Approval grant / allowlist management | Approval grants and allowlist entries are deliberately not exposed over MCP. In the default `full` access mode no grant is needed (mutations/execute are authorized); `restricted` mode re-enables deny-by-default. An in-process embedder configures the policy in code. `background_safe` classification **is** exposed through `macro_classify_background`. |
| GUI automation, dialogs, keystrokes | Forbidden by design. The instance is hidden and non-activating; `DisplayAlerts=False` is not treated as a universal dialog blocker. |
| Server-side log-level control | `EXCEL_VBA_MCP_LOG_LEVEL` is parsed but not applied; logging is fixed at `INFO` on `stderr`. |

## Coverage summary

The capability matrix contains **44 capability rows**.

| Coverage | Count | Rows |
|----------|-------|------|
| Fully supported | 18 | Application management; workbook open; close/save; workbook model; sheet list/get; range values read; range values write; range formulas read; range formulas write; names list/get; names create/update/delete; tables list/get/read; tables rows; tables columns; tables create/resize; pivot inventory; VBA read; snapshots |
| Partially supported | 15 | Sheet mutations; range copy/insert/delete; merge/unmerge; formatting; data validation; conditional formatting; pivot refresh; pivot field configuration; pivot filters; shapes; controls; VBA editing; compile diagnostics; macros; test runner |
| Not supported | 9 | Workbook creation; PivotTable/PivotCache create-delete; charts/PivotCharts/slicers; row/column sizing and layout; worksheet protection; workbook calculation; external links/connections; import/export/print/PDF/email; events |
| Not verified | 2 | Non-`.xlsm` formats; ActiveX event mapping |
| **Total** | **44** | — |

Read-only coverage is broad: every inventory and read surface above is available over MCP
without additional configuration, provided the path is allowed and the workbook can be opened.
In the default `full` access mode mutating and executing coverage is also reachable over MCP (a
baseline snapshot is required before a mutation); `EXCEL_VBA_MCP_ACCESS_MODE=restricted` reverts to
deny-by-default. Compilation remains `indeterminate`, and macro execution requires a
`background_safe` classification (via `macro_classify_background`) and passes the interactive-UI
veto.

Related documents: [Overview](01-overview.md), [Configuration](03-configuration.md),
[Architecture](04-architecture.md), [MCP tools reference](05-mcp-tools-reference.md),
[Workflows and examples](07-workflows-and-examples.md),
[Security](08-security.md), [Limitations and compatibility](09-limitations-and-compatibility.md)
and the [root README](../README.md).