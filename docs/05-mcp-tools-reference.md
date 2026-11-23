---
id: xlvba-mcp-tools-reference
domain: 05-mcp-tools-reference
tags:
  - mcp
  - tools
  - reference
  - excel
  - vba
related:
  - xlvba-overview
  - xlvba-architecture
  - xlvba-excel-capabilities
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This reference describes **all 88 tools** of the `xlvba` MCP server: purpose, parameters,
return value, side effects, typical errors and access class. The document is based on the
actual tool schemas obtained from a live `MCPServer` and on the source modules `api/*`,
`application/*` and `adapters/*`.

This document does not restate the architecture — that is covered in
[Architecture](04-architecture.md). Settings are in [Configuration](03-configuration.md),
the security model is in [Architecture](04-architecture.md), and limits are in
[Excel capabilities](06-excel-capabilities.md).

## How to read this reference

| Element | Meaning |
|---------|---------|
| `Tool` | The exact MCP name. Clients display it with a prefix (for example, Kilo uses `xlvba_excel_open`). Tables show it without the prefix. |
| `Parameters` | `name:type`; `?` means optional; `=value` is the default. `—` means no parameters. |
| `Returns` | The DTO type or response shape. All DTOs serialize to JSON and contain no COM references. |
| `Side effects` | What changes outside the response: files, snapshots, workbook state, artifacts, audit. |
| `Errors` | Stable `ErrorCode` values or other codes the tool may return. |
| `Access` | `read-only` (auto-allowed), `mutating` (guarded), `executing` (gated), `mixed` (depends on parameters). |

Shared conventions:

1. **Access mode.** In the default `full` mode (`EXCEL_VBA_MCP_ACCESS_MODE=full`), `ApprovalPolicy`
   authorizes `mutating` and `executing` operations without a per-action grant. In `restricted` mode
   the deny-by-default gate applies: those operations are refused with `APPROVAL_DENIED` unless an
   in-process allowlist entry or live grant matches. There is no MCP tool to grant approval or edit
   the allowlist, so a restricted server exposes only `read-only` work over MCP. Per the B4 error
   taxonomy: an approval refusal is `APPROVAL_DENIED`; `WORKBOOK_READ_ONLY` is raised only for a
   read-only session; an unknown or no-longer-open `workbook_id` is `WORKBOOK_NOT_FOUND` (a caller
   error, not contention); a stale `expected_source_hash` is `REVISION_CONFLICT`; and
   `COM_BUSY`/`RPC_E_CALL_REJECTED` are reserved for genuine COM contention only. `MACRO_BLOCKED`
   is only for Office macro-security / non-execute contexts. A lost owned Excel instance is reported
   as `WORKBOOK_NOT_FOUND` on the plain read/mutation paths and as `EXCEL_HUNG` (`needs_recovery`)
   from `macro_run` / `vba_compile`, which require recovery. In the per-tool tables, `COM_BUSY` on a
   mutating row denotes genuine contention, not approval denial, a bad id, or a revision conflict.
2. **`excel_open` first.** Any tool that needs a `workbook_id` requires the workbook to be
   opened first. `workbook_id` is the identifier of the working copy issued by the server, not a
   file path.
3. **Snapshot before mutation.** Mutating tools require a successful snapshot
   (`snapshot_create`); otherwise — `SNAPSHOT_FAILED`.
4. **Bounds.** Range and table reads are bounded (`offset`/`limit`, the `truncated` flag);
   requests beyond the limit are rejected, not silently truncated.
5. **No GUI.** Tools never activate windows, never show dialogs and never send keystrokes.

## Access classes

| Class | Count | Behavior |
|-------|-------|----------|
| `read-only` | 41 | Allowed automatically; does not change the workbook. May write snapshot, export, artifact and in-memory classification state. |
| `mutating` | 41 | Guarded: writable session → snapshot → approval → checks → postcondition + audit. |
| `executing` | 3 | Runs code (`macro_run`, `test_run_suite`, `test_suite_run`); additionally gated by approval (in `restricted` mode) and `background_safe`. |
| `mixed` | 2 | `excel_open` and `excel_close`: the class depends on parameters (`read_only`, `mode`, `save_changes`). |
| `config` (bookkeeping) | 1 | `macro_classify_background`: changes in-memory classification state only (no workbook mutation); ungated bookkeeping, `macro_run` remains the security boundary. |

It is important to distinguish "the tool is implemented" from "the tool is available through
MCP". All 41 `mutating` and 3 `executing` tools exist and are described in this reference. In the
default `full` access mode the approval gate authorizes them (a baseline snapshot is still required
before a mutation); only in `restricted` mode are they refused with `APPROVAL_DENIED`. The
`background_safe` requirement on `macro_run` applies in both modes.

**`vba_compile` is `read-only`, not `executing`.** `vba_compile` submits its STA diagnostic probe
with `op_class=EXECUTING`, but that tag only governs **retry eligibility** (the probe is never
blindly replayed after a busy/timeout) and the **audit `operation_class`** label — it does not drive
the approval gate. `CompileService` is constructed without an `ApprovalPolicy` and never calls the
approval gate, so `vba_compile` is **not** approval-gated: it runs in both `full` and `restricted`
modes, is non-mutating/non-activating, and never changes the workbook or shows UI.
`vba_get_compile_diagnostics` only reads the cached result. `vba_compile` is therefore counted as
`read-only` and is **not** one of the three `executing` tools (`macro_run`, `test_run_suite`,
`test_suite_run`), whose `EXECUTING` classification does gate approval and execute user code.

## Abbreviations

| Abbreviation | Meaning |
|--------------|---------|
| `STA` | Single-Threaded Apartment — the COM apartment model in which Excel lives. |
| `DTO` | Data Transfer Object — a typed structure without COM references. |
| `VBOM` | VBA project object model — the VBA project object model (requires trust in Trust Center). |
| `ListObject` | Excel Table — a structured worksheet table. |
| `BG06` | The `background_safe` classification rule for background macro execution. |
| `R13` | Interactive-UI preflight (case-insensitive): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`, `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`, `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`, `DisplayAlerts`, `.Run`, `CommandBars`, `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`. |
| `VBE` | VBA Editor — the VBA development environment. |
| `UIA` | UI Automation — read-only access to the Windows element tree (no activation). |

## Group summary

| Group | Count | Tools | Access class |
|-------|-------|-------|--------------|
| Workbook/session | 6 | `excel_open`, `excel_close`, `excel_save`, `excel_get_status`, `excel_list_workbooks`, `excel_health` | mixed / mutating / read-only |
| Snapshots | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` | read-only + 2 mutating |
| Test runner | 3 | `test_run_suite`, `test_suite_run`, `test_get_run` | executing + read-only |
| VBA | 9 | `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code`, `vba_replace_component_source`, `vba_insert_code` | read-only + 2 mutating |
| Compile | 2 | `vba_compile`, `vba_get_compile_diagnostics` | read-only (see access-class note) |
| Macros | 4 | `macro_list`, `macro_run`, `macro_get_last_result`, `macro_classify_background` | read-only + 1 executing + 1 config (bookkeeping) |
| Sheets | 7 | `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` | read-only + 5 mutating |
| Tables | 11 | `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`, `table_delete_column`, `table_create`, `table_resize` | read-only + 8 mutating |
| PivotTables | 13 | `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters` | read-only + 8 mutating |
| Ranges | 16 | `range_get`, `range_read_values`, `range_read_formulas`, `range_write_values`, `range_write_formulas`, `range_get_format`, `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`, `range_get_validation`, `range_set_validation`, `range_get_conditional_formatting`, `range_set_conditional_formatting` | read-only + 10 mutating |
| Names | 5 | `names_list`, `name_get`, `name_create`, `name_update`, `name_delete` | read-only + 3 mutating |
| Shapes | 3 | `shapes_list`, `shape_get`, `shape_set_on_action` | read-only + 1 mutating |
| Controls | 3 | `controls_list`, `control_get`, `control_set_on_action` | read-only + 1 mutating |
| Workbook model | 1 | `workbook_get_model` | read-only |
| Health | 1 | `server_health` | read-only |
| **Total** | **88** | — | 41 read-only, 41 mutating, 3 executing, 2 mixed, 1 config (bookkeeping) |

Sum across groups: 6 + 4 + 3 + 9 + 2 + 4 + 7 + 11 + 13 + 16 + 5 + 3 + 3 + 1 + 1 = **88**.

## Workbook/session

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `excel_open` | Opens an allowed workbook as a hidden working copy. | `path:string, mode:SessionMode=execute?, read_only:boolean=false?` | `WorkbookInfo` | File copy in `work/`; audit record | `PATH_NOT_ALLOWED`, `APPROVAL_DENIED` (restricted mode), `EXCEL_HUNG` | mixed |
| `excel_close` | Closes an open workbook; does not save by default. | `workbook_id:string, save_changes:boolean=false?` | `SessionStatus` | Closes the working copy; deletes the copy; audit record | `APPROVAL_DENIED` (restricted mode, when `save_changes=true`), `EXCEL_HUNG` | mixed |
| `excel_save` | Explicitly saves the working copy. | `workbook_id:string` | `WorkbookInfo` | Writes the working-copy file; audit record | `APPROVAL_DENIED` (restricted mode), `EXCEL_HUNG`, `COM_BUSY` | mutating |
| `excel_get_status` | Returns session/workbook state without touching COM. | `workbook_id:string?` | `SessionStatus` | — | — | read-only |
| `excel_list_workbooks` | Lists open workbooks from bookkeeping. | — | `list[WorkbookInfo]` | — | — | read-only |
| `excel_health` | Worker/session state without forcing a restart and without COM. | — | `ExcelHealth` | — | — (always a payload) | read-only |

`excel_open` is `mixed`: `mode="inspect"` + `read_only=true` passes as `READ_ONLY`;
`mode="inspect"` + `read_only=false` is `MUTATING`; `mode="execute"` is `EXECUTING`.
`excel_close` is `read-only` with `save_changes=false` and `mutating` with `save_changes=true`.

## Snapshots

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `snapshot_create` | Creates an immutable snapshot of an open or known workbook. | `workbook_id:string` | `SnapshotManifest` | Snapshot files + `manifest.json`; audit record | `SNAPSHOT_FAILED`, `APPROVAL_DENIED` (restricted mode), `EXCEL_HUNG` | mutating (approval-gated) |
| `snapshot_list` | Lists snapshots and their metadata. | — | `list[SnapshotManifest]` | — | `SNAPSHOT_FAILED` | read-only |
| `snapshot_restore` | Restores a working session from a snapshot. | `snapshot_id:string, workbook_id:string?` | `WorkbookInfo` | Closes the current copy without saving; opens a new working copy; audit | `SNAPSHOT_FAILED`, `APPROVAL_DENIED` (restricted mode), `EXCEL_HUNG` | mutating |
| `snapshot_diff` | Compares two snapshots by metadata and hash. | `snapshot_id_a:string, snapshot_id_b:string` | `SnapshotDiff` | — | `SNAPSHOT_FAILED` | read-only |

## Test runner

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `test_run_suite` | Runs a declarative test suite against disposable copies of the workbook. | `suite:TestSuiteSpec` | `TestRunResult` | Baseline snapshot, a working copy per case, artifacts, `runs/`, audit | `INVALID_RANGE`, `APPROVAL_DENIED` (restricted mode), `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `EXCEL_HUNG` | executing |
| `test_suite_run` | Explicit entry point: accepts a suite object or a JSON-like mapping. | `suite:TestSuiteSpec` or `object` | `TestRunResult` | Same as `test_run_suite` | `INVALID_RANGE`, `APPROVAL_DENIED` (restricted mode), `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED` | executing |
| `test_get_run` | Returns a persisted run by `run_id`. | `run_id:string` | `TestRunResult` | — | `ExcelVbaMcpError` (not found) | read-only |

Both runners open the working copy in `execute` mode. In the default `full` mode this is
authorized; in `restricted` mode every case gets `blocked` (`APPROVAL_DENIED`) and the suite is
`blocked`. A `macro_run` step on an unclassified procedure is blocked with
`BACKGROUND_SAFETY_NOT_VERIFIED` in either mode.

## VBA

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `vba_project_info` | VBA project summary: name, protection, components, references. | `workbook_id:string` | `VbaProjectInfo` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | read-only |
| `vba_list_components` | Lists components with type, line count and hash. | `workbook_id:string` | `list[VbaComponentInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_list_procedures` | Procedures of one component via the VBIDE CodeModule. | `workbook_id:string, component:string` | `list[VbaProcedureInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_read_component` | Reads a component module (code + normalized hash). | `workbook_id:string, component:string` | `VbaComponentSource` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_export_project` | Exports every component and returns a manifest. | `workbook_id:string, out_dir:string?` | `VbaExportManifest` | `.bas`/`.cls`/`.frm`/`.frx` files; artifact root; audit | `ACCESS_VBOM_DISABLED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED` | read-only |
| `vba_get_references` | Project references with Name/GUID/version/path. | `workbook_id:string` | `list[VbaReferenceInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_search_code` | Searches all components' code (substring or regex). | `workbook_id:string, pattern:string, regex:boolean=false?` | `list[VbaSearchMatch]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_replace_component_source` | Fully replaces a Standard/Class module. | `workbook_id:string, component:string, new_source:string, expected_source_hash:string` | `VbaMutationResult` | Changes the VBA project; before/after export; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED` (restricted mode), `REVISION_CONFLICT` (revision conflict), `PROTECTED_TARGET` (non-replaceable component), `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | mutating |
| `vba_insert_code` | Inserts code at a semantic position in a Standard/Class module. | `workbook_id:string, component:string, code:string, position:string, expected_source_hash:string` | `VbaMutationResult` | Changes the VBA project; audit | same; an invalid `position` — `COM_BUSY` | mutating |

`vba_replace_component_source` and `vba_insert_code` work only with Standard/Class modules;
Document modules and UserForms are rejected with `PROTECTED_TARGET`
(`reason=non_replaceable_component`). A stale `expected_source_hash` is rejected with
`REVISION_CONFLICT` (`reason=revision_conflict`).

`vba_export_project` writes files and is approval-gated (F-4). With `out_dir` omitted it exports to
a fresh temporary directory (the manifest returns absolute paths). A supplied `out_dir` must live
under the resolved `snapshots` / `artifacts` artifact root (default
`<temp>/excel_vba_mcp/artifacts`, or `EXCEL_VBA_MCP_ARTIFACT_DIR`); an out-of-root `out_dir` fails
closed with `PATH_NOT_ALLOWED` whose `context.allowed_root` / `context.expected_root` carry the
**absolute** allowed artifacts-root path (not the symbolic `"artifacts"` name).

## Compile

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `vba_compile` | Fail-closed compile result; always `indeterminate` on the validated build. | `workbook_id:string` | `CompileResult` | — (does not launch UI) | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |
| `vba_get_compile_diagnostics` | Last diagnostics without re-running. | `workbook_id:string` | `CompileDiagnostics` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | read-only |

`vba_compile`'s worker probe is tagged `op_class=EXECUTING` for retry/audit only; the tool itself is
`read-only` and not approval-gated (see the access-class note).

## Macros

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `macro_list` | Lists callable macros with their `background_safe` classification. | `workbook_id:string` | `list[MacroInfo]` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | read-only |
| `macro_run` | Runs one approved `background_safe` macro. | `workbook_id:string, macro_name:string, args:array?` | `MacroRunResult` | VBA execution; possible external effects; audit | `APPROVAL_DENIED` (restricted mode), `MACRO_BLOCKED` (Office macro-security), `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED` | executing |
| `macro_classify_background` | Records a procedure's `background_safe` for the current workbook hash and VBA source revision (ungated bookkeeping). | `workbook_id:string, procedure:string, background_safe:boolean, expected_project_source_hash:string?` | `BackgroundSafetyClassification` | Mutates in-memory classification state (bookkeeping); audit | `MACRO_NOT_FOUND`, `STALE_TARGET` (source hash mismatch), `INTERACTIVE_UI_NOT_SUPPORTED` (interactive/heuristic + `background_safe=true`), `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | config (bookkeeping, ungated) |
| `macro_get_last_result` | Returns the cached last run without re-executing. | `workbook_id:string` | `MacroRunResult` | — | — | read-only |

`macro_classify_background` is **ungated**: it only records classification state and is not the
security boundary — `macro_run` keeps the approval gate and the deterministic R13 veto. The
classification is bound to `(workbook_file_sha256, project_source_hash, procedure)` and is
invalidated when the VBA source changes (a later `macro_run`/`macro_list` recomputes the live
`project_source_hash`; a mismatch means the procedure is unclassified again and returns
`BACKGROUND_SAFETY_NOT_VERIFIED`). Passing a stale `expected_project_source_hash` rejects the
write with `STALE_TARGET`.

## Sheets

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `sheets_list` | All sheets with type, visibility, `code_name` and protection. | `workbook_id:string` | `list[SheetInfo]` | — | — | read-only |
| `sheet_get` | One sheet by name (honours `Sheets`/`Worksheets`). | `workbook_id:string, name:string` | `SheetInfo` | — | `COM_BUSY` (sheet not found) | read-only |
| `sheet_add` | Adds a worksheet (MVP: only `worksheet`). | `workbook_id:string, name:string, sheet_type:string=worksheet?` | `SheetInfo` | New sheet in the workbook; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_rename` | Renames a sheet; `CodeName` is unchanged. | `workbook_id:string, name:string, new_name:string` | `SheetInfo` | Sheet name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_copy` | Copies a sheet and returns the new sheet's DTO. | `workbook_id:string, name:string` | `SheetInfo` | New sheet; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_delete` | Deletes a sheet (destructive). | `workbook_id:string, name:string` | `SheetInfo` | Sheet deletion; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `sheet_set_visibility` | Sets sheet visibility. | `workbook_id:string, name:string, visible:boolean` | `SheetInfo` | Sheet visibility; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

`sheet_delete` is rejected if no visible sheet would remain after the deletion;
`sheet_set_visibility` rejects hiding the last visible sheet.

## Tables

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `tables_list` | All tables (ListObject) with ranges and columns. | `workbook_id:string` | `list[TableInfo]` | — | — | read-only |
| `table_get` | One table by name/`DisplayName`. | `workbook_id:string, name:string` | `TableInfo` | — | `COM_BUSY` (not found) | read-only |
| `table_read_rows` | Typed, bounded slice of table rows. | `workbook_id:string, name:string, offset:integer=0?, limit:integer?` | `TableRows` | — | `COM_BUSY` | read-only |
| `table_append_rows` | Appends rows via `ListRows.Add` (structure preserved). | `workbook_id:string, name:string, rows:array<object>` | `TableMutationResult` | Table rows; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_insert_row` | Inserts one row at the 1-based position `at`. | `workbook_id:string, name:string, at:integer, values:object?` | `TableMutationResult` | Table row; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_update_rows` | Updates rows by index or an unambiguous `where`. | `workbook_id:string, name:string, updates:array<object>` | `TableMutationResult` | Row values; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_delete_rows` | Deletes rows by indices or an unambiguous `where`. | `workbook_id:string, name:string, indices:array?, where:object?` | `TableMutationResult` | Table rows; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_add_column` | Adds a column via `ListColumns.Add(Position)`. | `workbook_id:string, name:string, header:string, position:integer?` | `TableMutationResult` | Table column; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_delete_column` | Deletes a column by its unique header. | `workbook_id:string, name:string, header:string` | `TableMutationResult` | Table column; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_create` | Creates a ListObject on a sheet over an A1 range or name. | `workbook_id:string, sheet:string, range_or_name:string, style:string?` | `TableMutationResult` | New table; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `table_resize` | Resizes the table via `ListObject.Resize`. | `workbook_id:string, name:string, new_range:string` | `TableMutationResult` | Table range; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

## PivotTables

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `pivot_tables_list` | All PivotTables and their caches across all sheets. | `workbook_id:string` | `PivotInventory` | — | `COM_BUSY` | read-only |
| `pivot_table_get` | Metadata for one PivotTable by `Name`. | `workbook_id:string, pivot_table:string` | `PivotTableInfo` | — | `COM_BUSY` | read-only |
| `pivot_cache_get` | PivotCache metadata by 1-based `CacheIndex`. | `workbook_id:string, cache_index:integer` | `PivotCacheInfo` | — | `COM_BUSY` | read-only |
| `pivot_fields_list` | PivotTable fields with orientation and aggregation. | `workbook_id:string, pivot_table:string` | `list[PivotFieldInfo]` | — | `COM_BUSY` | read-only |
| `pivot_table_refresh` | Refreshes one PivotTable (`RefreshTable`). | `workbook_id:string, pivot_table:string` | `PivotRefreshResult` | PivotTable data; audit | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_cache_refresh` | Refreshes one PivotCache (`PivotCache.Refresh`). | `workbook_id:string, cache_index:integer` | `PivotRefreshResult` | Data of all tables in the cache; audit | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_layout` | Sets a field's orientation or removes it. | `workbook_id:string, pivot_table:string, field:string, orientation:PivotFieldOrientation, delete:boolean=false?` | `PivotFieldInfo` | PivotTable structure; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_position` | Sets the 1-based position of a field in its orientation. | `workbook_id:string, pivot_table:string, field:string, position:integer` | `PivotFieldInfo` | Field order; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_aggregation` | Adds or updates a data-field aggregation. | `workbook_id:string, pivot_table:string, field:string, function:PivotAggregation, caption:string?` | `PivotFieldInfo` | Data field; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_set_properties` | Sets `caption` / `number_format` / `show_all_items`. | `workbook_id:string, pivot_table:string, field:string, properties:PivotFieldProperties` | `PivotFieldInfo` | Field properties; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_get_filters` | Current filter state of a field. | `workbook_id:string, pivot_table:string, field:string` | `PivotFieldFilters` | — | `AMBIGUOUS_FILTER`, `COM_BUSY` | read-only |
| `pivot_field_apply_filter` | Applies a manual filter over visible items. | `workbook_id:string, pivot_table:string, field:string, spec:PivotFilterSpec` | `PivotFieldFilters` | Item visibility; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `pivot_field_clear_filters` | Clears a field's filters or the whole table's filters. | `workbook_id:string, pivot_table:string, field:string?, scope:string=field?` | `PivotFieldFilters` | Filters; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

Creating and deleting PivotTables/PivotCaches is not supported. `PivotFilters.Add`
(label/value/date/top-bottom) does not work on the validated build: it returns an explicit
`unsupported` result with `filter.supported=false` and a reason, rather than silently converting
to a manual filter.

## Ranges

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `range_get` | Metadata for an A1 range (address, size, shape). | `workbook_id:string, sheet:string, address:string` | `RangeAddress` | — | `INVALID_RANGE` | read-only |
| `range_read_values` | Typed values of a bounded window. | `workbook_id:string, sheet:string, address:string, offset:integer=0?, limit:integer?` | `RangeValues` | — | `INVALID_RANGE` | read-only |
| `range_read_formulas` | Typed formulas of a bounded window. | `workbook_id:string, sheet:string, address:string, offset:integer=0?, limit:integer?` | `RangeFormulas` | — | `INVALID_RANGE` | read-only |
| `range_write_values` | Writes typed values to a bounded range. | `workbook_id:string, sheet:string, address:string, payload:RangeWriteRequest` | `RangeMutationResult` | Cell values; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_write_formulas` | Writes formulas to a bounded range. | `workbook_id:string, sheet:string, address:string, payload:FormulaWriteRequest` | `RangeMutationResult` | Cell formulas; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_get_format` | Tri-state formatting of a range. | `workbook_id:string, sheet:string, address:string` | `FormatInfo` | — | `INVALID_RANGE` | read-only |
| `range_set_format` | Partial formatting update. | `workbook_id:string, sheet:string, address:string, patch:RangeFormatPatch` | `FormatMutationResult` | Formatting; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE` | mutating |
| `range_copy` | Copies a range to another sheet/address (no clipboard). | `workbook_id:string, sheet:string, source:string, destination_sheet:string, destination:string` | `StructuralMutationResult` | Destination data; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION` | mutating |
| `range_insert` | Inserts cells with a `right`/`down` shift. | `workbook_id:string, sheet:string, address:string, shift:string, allow_merged:boolean=false?` | `StructuralMutationResult` | Cell shift; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_delete` | Deletes cells with a `left`/`up` shift. | `workbook_id:string, sheet:string, address:string, shift:string, allow_merged:boolean=false?` | `StructuralMutationResult` | Cell deletion; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_merge` | Merges a range into a single cell. | `workbook_id:string, sheet:string, address:string, across:boolean=false?, allow_data_loss:boolean=false?` | `StructuralMutationResult` | Merge; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_unmerge` | Unmerges every merged area intersecting the target. | `workbook_id:string, sheet:string, address:string` | `StructuralMutationResult` | Unmerge; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | mutating |
| `range_get_validation` | Data-validation rule of a range. | `workbook_id:string, sheet:string, address:string` | `ValidationRule` | — | `INVALID_RANGE` | read-only |
| `range_set_validation` | Adds or changes a data-validation rule. | `workbook_id:string, sheet:string, address:string, spec:ValidationSpec` | `ValidationRule` | Validation rule; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `PROTECTED_TARGET` | mutating |
| `range_get_conditional_formatting` | Conditional-formatting rules by priority. | `workbook_id:string, sheet:string, address:string` | `list[ConditionalFormatRule]` | — | `INVALID_RANGE` | read-only |
| `range_set_conditional_formatting` | Adds, changes or deletes a rule. | `workbook_id:string, sheet:string, address:string, spec:ConditionalFormatSpec, mode:string=add?` | `ConditionalFormatMutationResult` | Rules; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `PROTECTED_TARGET` | mutating |

Structural operations (`range_insert`, `range_delete`, `range_copy`, `range_merge`) that
intersect an Excel Table (ListObject) are rejected with `LISTOBJECT_INTERSECTION` and must be
performed through the Tables tools.

## Names

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `names_list` | All workbook and worksheet-scoped defined names. | `workbook_id:string` | `list[DefinedNameInfo]` | — | — | read-only |
| `name_get` | One name by name and explicit scope. | `workbook_id:string, name:string, scope:string, sheet:string?` | `DefinedNameInfo` | — | `COM_BUSY` | read-only |
| `name_create` | Creates a name in an explicit scope. | `workbook_id:string, name:string, refers_to:string, scope:string, sheet:string?` | `DefinedNameInfo` | New name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `name_update` | Updates `RefersTo` (scope is preserved). | `workbook_id:string, name:string, refers_to:string, scope:string, sheet:string?` | `DefinedNameInfo` | Name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |
| `name_delete` | Deletes a name and returns its previous value. | `workbook_id:string, name:string, scope:string, sheet:string?` | `DefinedNameInfo` | Name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

## Shapes

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `shapes_list` | All shapes with type, `on_action`, position and text. | `workbook_id:string` | `list[ShapeInfo]` | — | — | read-only |
| `shape_get` | One shape by name (case-insensitive). | `workbook_id:string, name:string` | `ShapeInfo` | — | `COM_BUSY` | read-only |
| `shape_set_on_action` | Assigns a macro to a shape (the macro is not executed). | `workbook_id:string, name:string, macro_name:string` | `ShapeInfo` | `OnAction`; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

An empty `macro_name` clears the assignment.

## Controls

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `controls_list` | All Form Controls and ActiveX controls in the workbook. | `workbook_id:string` | `list[ControlEntry]` | — | — | read-only |
| `control_get` | One control by name (case-insensitive). | `workbook_id:string, name:string` | `ControlEntry` | — | `COM_BUSY` | read-only |
| `control_set_on_action` | Assigns a macro to a Form Control (not ActiveX). | `workbook_id:string, name:string, macro_name:string` | `ControlEntry` | `OnAction`; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | mutating |

ActiveX controls are rejected: they do not have Form Control `OnAction` semantics.

## Workbook model

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `workbook_get_model` | Aggregated workbook model: sheets, tables, names, shapes, controls, VBA. | `workbook_id:string, include_vba:boolean=true?` | `WorkbookModel` | — | — (VBA state is in `vba_access`) | read-only |

The model does not read "all cells" and does not execute VBA; when VBOM is unavailable the
other sections are still returned and `vba_access` reports the state.

## Health

| Tool | Purpose | Parameters | Returns | Side effects | Errors | Access |
|------|---------|------------|---------|--------------|--------|--------|
| `server_health` | Read-only liveness probe; never touches Excel/COM. | — | `{"status", "service", "version"}` | — | — | read-only |

`excel_health` is grouped with Workbook/session because it describes session and worker state;
the Health group contains only the server liveness probe.

## Complex tools in detail

### excel_open

`excel_open` defaults to `mode="execute", read_only=False` — a writable, execution-capable
session. Pass `mode="inspect"` and/or `read_only=True` to opt into a macros-disabled or read-only
session. Read-only inspect mode:

```json
{ "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
```

Expected response shape:

```json
{
  "workbook_id": "1f0c…",
  "name": "report.xlsm",
  "format": "xlsm",
  "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm",
  "source_path": "C:/data/report.xlsm",
  "read_only": true,
  "has_vba": true,
  "excel_version": "16.0.14332",
  "session_id": "9b2a…",
  "mode": "inspect"
}
```

A writable/execute open (the default when `mode`/`read_only` are omitted):

```json
{ "path": "C:/data/report.xlsm", "mode": "execute", "read_only": false }
```

In the default `full` access mode this open is authorized and a working copy is created. In
`restricted` mode it is classified `MUTATING`/`EXECUTING` and **refused** with
`code="APPROVAL_DENIED"`; the workbook is not opened and no working copy is created, because no
grant tool exists through MCP.

### excel_close / excel_save

```json
{ "workbook_id": "1f0c…", "save_changes": false }
```

`excel_close` without saving is read-only and allowed; the response is a `SessionStatus` with
`state="closed"`. Saving on close is a mutation:

```json
{ "workbook_id": "1f0c…", "save_changes": true }
```

Authorized in the default `full` mode; in `restricted` mode it is refused with `APPROVAL_DENIED`.
Explicit save:

```json
{ "workbook_id": "1f0c…" }
```

Requires `MUTATING` approval (authorized in `full` mode; `APPROVAL_DENIED` in `restricted` mode)
and returns an updated `WorkbookInfo` with `dirty=false`. Changes are never saved implicitly.

### range_read_values / range_write_values

Reading a bounded window:

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "offset": 0,
  "limit": 10
}
```

Response — `RangeValues`:

```json
{
  "range": { "sheet": "Sheet1", "address": "A1:C10", "rows": 10, "cols": 3 },
  "shape": "block",
  "rows": 10,
  "cols": 3,
  "values": [
    [ { "kind": "string", "value": "Header" }, { "kind": "empty" }, { "kind": "empty" } ]
  ],
  "truncated": false,
  "notes": []
}
```

Values are typed: `string`, `number`, `boolean`, `date`, `error`, `empty`; precision loss from
coercion to string is excluded. A request beyond the cell limit is rejected, not truncated;
`truncated=true` appears when the `offset`/`limit` window is clipped.

Writing requires an exact rectangular `payload`:

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "B2:C3",
  "payload": {
    "rows": 2,
    "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "=1+1" } ],
      [ { "kind": "boolean", "value": true }, { "kind": "empty" } ]
    ]
  }
}
```

A string beginning with `=` is stored as **text**, not as a formula (a separate tool exists for
formulas). A wrong payload shape is rejected with `INVALID_RANGE` and no changes. The response
is a `RangeMutationResult` with `range_before`, `range_after`, `verified` and
`postcondition_values` (re-read values). The mutation requires a writable session and a snapshot;
the approval gate authorizes it in the default `full` mode and refuses it with `APPROVAL_DENIED` in
`restricted` mode.

### range_set_format

Sparse write: only the fields explicitly set in `patch` are applied.

```json
{
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C1",
  "patch": {
    "font_bold": true,
    "font_size": 11,
    "fill_color": 15773696,
    "number_format": "0.00",
    "alignment_horizontal": "center"
  }
}
```

Missing fields stay unchanged; `null` resets a value to automatic only where Excel allows it.
Response:

```json
{
  "action": "set_format",
  "range": { "sheet": "Sheet1", "address": "A1:C1", "rows": 1, "cols": 3 },
  "applied": ["font.bold", "number_format"],
  "ignored": [],
  "verified": true,
  "result_format": { "font": { "bold": true }, "number_format": "0.00" },
  "notes": []
}
```

### table_update_rows / table_delete_rows

An update accepts either an **explicit index** or an **unambiguous `where`**:

```json
{
  "workbook_id": "1f0c…",
  "name": "Sales",
  "updates": [
    { "index": 2, "values": { "Amount": 120 } },
    { "where": { "Region": "North" }, "values": { "Flag": "Y" } }
  ]
}
```

If `where` does not resolve to exactly one row, the tool returns `COM_BUSY` with a message about
the ambiguity and a preview of the matched rows — there is no automatic "guessing". (Table
ambiguity deliberately reuses the base `COM_BUSY` code with `context.reason="ambiguous"`; it is
distinct from the PivotTable `AMBIGUOUS_FILTER` code used by the pivot tools.) The response
is a `TableMutationResult` with `before`/`after` (`TableInfo`) and `affected_rows`.

Deleting rows:

```json
{ "workbook_id": "1f0c…", "name": "Sales", "indices": [3, 5] }
```

or via `where`; omitting both arguments is a validation error. Deletion and update require a
writable session and a snapshot; the approval gate authorizes them in the default `full` mode and
refuses them with `APPROVAL_DENIED` in `restricted` mode.

### snapshot_create / snapshot_restore

```json
{ "workbook_id": "1f0c…" }
```

Response — `SnapshotManifest`:

```json
{
  "snapshot_id": "snap-…",
  "created_at": "2026-11-21T18:41:56Z",
  "source_path": "C:/data/report.xlsm",
  "file_name": "report.xlsm",
  "stored_path": "<snapshots>/snap-…/report.xlsm",
  "sha256": "…",
  "size": 48213,
  "snapshot_method": "SaveCopyAs",
  "vba_exports": [
    {
      "component_name": "Module1",
      "type": 1,
      "path": "exports/Module1.bas",
      "export_file_sha256": "…"
    }
  ],
  "model_version": 2,
  "external_side_effects_not_undone": true
}
```

The top-level `sha256` is the snapshot **file** hash. Each `vba_exports` entry's
`export_file_sha256` is the SHA-256 of the **exported component file bytes** (F7 renamed this
from `sha256` at `model_version: 2` so it is not mistaken for the normalized `source_hash`). It
is **not** a valid `expected_source_hash` for the VBA mutation tools — those require the
normalized `source_hash` from `vba_read_component`. `snapshot_list` returns the same manifests;
`snapshot_diff` compares the top-level snapshot file hashes (not `export_file_sha256`).

Restore:

```json
{ "snapshot_id": "snap-…", "workbook_id": "1f0c…" }
```

The current working copy is closed **without saving**, then a new copy is opened from the
baseline; a `WorkbookInfo` is returned. External macro effects are not rolled back.

Note that `snapshot_create` and `vba_export_project` write files: snapshots under the snapshot
temp directory and exported components under the artifact/temp directory. Both are
**approval-gated** (F-4): they are auto-allowed in the default `full` mode and refused with
`APPROVAL_DENIED` in `restricted` mode. `snapshot_restore` adds no second gate — its approval
point is the inherited writable `excel_open` (EXECUTING) it performs on the snapshot baseline.

### vba_replace_component_source

Before replacing, two calls are needed: `vba_list_components` (confirm the component is
Standard/Class) and `vba_read_component` (obtain `source_hash`). Then:

```json
{
  "workbook_id": "1f0c…",
  "component": "Module1",
  "new_source": "Option Explicit\n\nPublic Sub Hello()\nEnd Sub",
  "expected_source_hash": "<source_hash from vba_read_component>"
}
```

`expected_source_hash` is the stale-revision guard: the server compares it with the current
normalized hash before and **inside** the STA worker. A mismatch raises `REVISION_CONFLICT` with
`context.reason="revision_conflict"` and the mutation is not performed. The guard token is the
**normalized `source_hash`** from `vba_read_component` — never the snapshot export **file** hash
(`snapshot_create.vba_exports[].export_file_sha256`, the SHA-256 of the exported `.bas`/`.cls`/`.frm`
bytes). The two hashes differ (UserForm exports carry a designer header and `.frx` is not
byte-reproducible), so passing the snapshot file hash yields `REVISION_CONFLICT` (F7). The response
is a `VbaMutationResult` with `before_hash`, `after_hash` and export paths. Document modules and
UserForms are rejected.

### macro_run

```json
{ "workbook_id": "1f0c…", "macro_name": "Module1.Hello", "args": [] }
```

Gate order:

1. execute mode + writable session (`_require_execute_context`);
2. procedure resolution from the inventory (`Module.Procedure`);
3. approval (key: canonical path + SHA-256 + action + procedure) — authorized in the default
   `full` mode, `APPROVAL_DENIED` in `restricted` mode;
4. `background_safe` for the current workbook hash **and** `project_source_hash` (BG06);
5. interactive-UI preflight (R13): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`,
   `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`,
   `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`,
   `DisplayAlerts`, `.Run`, `CommandBars`, `AppActivate`, `Application.Visible`, `.Activate`,
   `WindowState` (case-insensitive);
6. one `Application.Run` on the STA worker.

Possible `MacroRunResult` statuses: `success`, `not_found`, `runtime_error`, `blocked`,
`timeout`, `needs_recovery`. When blocked the macro **is not started**: `blocked` with
`APPROVAL_DENIED` (restricted-mode approval), `BACKGROUND_SAFETY_NOT_VERIFIED` (unclassified for
the current workbook hash and source revision) or `INTERACTIVE_UI_NOT_SUPPORTED` (dangerous token,
an "unsafe" classification, or a heuristic procedure). A timeout with a modal window owned by the
owned PID yields `MACRO_RUNTIME_ERROR` and `dismissed=false`; without a modal window —
`MACRO_TIMEOUT`/`needs_recovery`. There is no blind re-run.

**Residual false-safe boundary.** The R13 token scan is a heuristic veto, not static proof.
Dynamic/indirect UI (`CallByName` / `ExecuteGlobal` / `Eval` indirection) and UI reached
transitively through other procedures the scan does not see are not fully covered. The veto list
now also covers window/focus and visibility tokens (`AppActivate`, `Application.Visible`,
`.Activate`, `WindowState`), but a procedure that reaches the UI only indirectly can still evade a
substring scan. A `background_safe` classification is a human review outcome, and the deterministic
layer can only veto, never certify.

### macro_classify_background

Classify a procedure's `background_safe` for the current workbook hash and VBA source revision:

```json
{
  "workbook_id": "1f0c…",
  "procedure": "Module1.Hello",
  "background_safe": true
}
```

The optional `expected_project_source_hash` is an optimistic-concurrency guard: when it no longer
matches the live source hash the write is rejected with `STALE_TARGET`. Classifying a heuristic or
interactive-token procedure as `background_safe` is rejected with `INTERACTIVE_UI_NOT_SUPPORTED`
(the R13 veto). The result is a `BackgroundSafetyClassification` carrying `workbook_sha256`,
`project_source_hash` and `classified_at`. The tool is **ungated** (no approval) and is not the
security boundary — `macro_run` still applies the approval gate and the R13 veto. The registry is
in-memory and server-lifetime; a later VBA source change invalidates the classification.

### test_suite_run

```json
{
  "suite": {
    "suite": "smoke",
    "workbook_id": "1f0c…",
    "baseline": "create",
    "cases": [
      {
        "id": "reads-range",
        "steps": [
          {
            "action": "range_read_values",
            "sheet": "Sheet1",
            "range_address": "A1:B2",
            "limit": 2
          }
        ],
        "assertions": [
          {
            "type": "range_structure",
            "sheet": "Sheet1",
            "address": "A1:B2",
            "rows": 2,
            "cols": 2
          }
        ]
      }
    ]
  }
}
```

`action` and `type` come from **closed** vocabularies: `TestAction` (37 actions) and
`AssertionType` (29 checks); the schema forbids extra fields (`extra="forbid"`). An unknown
action, arbitrary code, a shell command or an unbounded payload is rejected before any Excel
access (`INVALID_RANGE`).

Isolation: one immutable baseline snapshot, then **a disposable working copy per case**; a case
cannot "dirty" the next one. Sequential edit→macro→edit→macro flows are supported: a `macro_run`
step may carry an optional data-only `background_safe` field, and the runner classifies the
procedure through the guarded `MacroService.classify_background` path before dispatch (never
touching the registry directly). Compile-before-macro is explicit — include a `vba_compile` step;
a failed `vba_compile` step blocks later `macro_run` steps in the same case, and a `macro_run` step
on an unclassified procedure is blocked with `BACKGROUND_SAFETY_NOT_VERIFIED`. An `indeterminate`
compile is not treated as a compile failure. The return value is a `TestRunResult` with
per-case/per-step statuses; the run is persisted (`runs/`) and read via `test_get_run`.

### vba_compile

```json
{ "workbook_id": "1f0c…" }
```

Fail-closed response:

```json
{
  "workbook_id": "1f0c…",
  "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": {
    "is_compile": false,
    "vbom_available": true,
    "protection": 0,
    "locked": false,
    "component_count": 3,
    "broken_references": []
  },
  "command_discovery": { "route_usable": false },
  "notes": [
    "NOT a compile: non-activating VBA compilation is unavailable on this build.",
    "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"
  ]
}
```

The tool never returns `compiled` or `compile_error`, and never shows or activates UI. Reasons:
`NON_ACTIVATING_COMPILE_UNAVAILABLE`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`.
`vba_get_compile_diagnostics` returns the same data from the cache without re-collecting.

## Tool dependencies

1. **`excel_open` before any workbook tool.** Every tool with a `workbook_id` requires an open
   working copy; otherwise — `ExcelVbaMcpError` "unknown workbook_id".
2. **`snapshot_create` before any guarded mutation.** Without a successful snapshot the mutation
   is rejected with `SNAPSHOT_FAILED`.
3. **`vba_list_components` / `vba_read_component` before `vba_replace_component_source` and
   `vba_insert_code`.** They provide the component name and the current `source_hash` for
   `expected_source_hash`.
4. **`vba_compile` before `vba_get_compile_diagnostics` — only to seed the cache.** Without the
   first call the diagnostics are collected again; this does not change the result
   (`is_compile=false`).
5. **`macro_list` before `macro_run`.** It reveals the exact `Module.Procedure` name and the
   `background_safe` status before execution.
6. **`test_get_run` after `test_run_suite` / `test_suite_run`.** The `run_id` comes from
   `TestRunResult.test_run_id`; the tool only reads the persisted run.

Additionally, `excel_close` (without saving) and `excel_save` are incompatible with subsequent
calls using the same `workbook_id` — after closing, the identifier is no longer valid.

## Mutating and potentially dangerous operations

| Tool | What changes | Why it is gated | How to use it safely |
|------|--------------|-----------------|----------------------|
| `excel_open` (writable/execute) | Creates a working copy and opens it with macros (`execute`) | Default open; authorized in `full` mode, `APPROVAL_DENIED` in `restricted` | Use `mode="inspect", read_only=true` for read-only work; apply `execute` only to verified copies |
| `excel_close(save_changes=true)` | Saves the working copy to disk | `MUTATING` | Save only after `snapshot_create` and a deliberate check |
| `excel_save` | Writes the working-copy file | `MUTATING` | Combine with a snapshot; do not save "just in case" |
| `snapshot_restore` | Discards the current copy without saving and opens the baseline | `mutating` (including writable-open) | Take a fresh `snapshot_create` before restoring |
| `sheet_delete` | Deletes a sheet with its data | Destructive, gated | Check `sheets_list`; rely on a snapshot |
| `table_delete_rows` | Deletes table rows | Destructive, gated | Use explicit `indices`; use `where` only when unambiguous |
| `table_delete_column` | Deletes a table column | Destructive, gated | Check `table_get`; do not pass an ambiguous header |
| `range_write_values` | Overwrites cell values | `MUTATING`, gated | Check the payload size; take a snapshot |
| `range_write_formulas` | Overwrites formulas | `MUTATING`, gated | Do not use it for constants (use values) |
| `range_delete` | Deletes cells with a shift | Destructive, gated | Always pass `shift` explicitly; check table intersections |
| `range_merge` | Merges cells, discarding data outside the anchor | Can lose data | Review `preconditions`; pass `allow_data_loss=true` only after review |
| `name_delete` | Deletes a defined name | Can break formulas | Check `names_list` and dependencies; rely on a snapshot |
| `vba_replace_component_source` | Fully replaces a module's code | `MUTATING`, gated | Pass the exact `expected_source_hash`; check `vba_compile` afterwards |
| `vba_insert_code` | Inserts code into a module | `MUTATING`, gated | Specify an exact `position`; do not insert code blindly |
| `shape_set_on_action` | Assigns a macro to a shape | `MUTATING`, gated | Check `shapes_list`; an empty `macro_name` clears the assignment |
| `control_set_on_action` | Assigns a macro to a Form Control | `MUTATING`, gated | Work only with Form Controls; ActiveX is unsupported |
| `macro_run` | Executes VBA code | `EXECUTING`, gated by approval (restricted mode) + `background_safe` + R13 | Only after `macro_list` / `macro_classify_background` and preflight; do not run macros with external effects |

## Read-only resources

| URI template | Name | MIME | Content |
|--------------|------|------|---------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` | Stage 1 workbook model (sheets, tables, names, shapes, controls, optional VBA) |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` | VBA component inventory + `vba_access` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` | Source text of one VBA component |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` | Persisted declarative test run |

Resources are read-only: they do not mutate the workbook and do not execute VBA.

## Known tool limitations

1. **Compilation is always `indeterminate`.** `vba_compile` cannot confirm compilation on the
   validated build; `is_compile=false`.
2. **Macros require background-safety classification.** `BackgroundSafetyRegistry` starts empty
   (fail-closed); classify a procedure with `macro_classify_background`. An unclassified procedure
   returns `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`); an interactive one is vetoed
   (`INTERACTIVE_UI_NOT_SUPPORTED`). The R13 token scan is a heuristic veto, not static proof.
3. **No approval grant tool.** Approval/allowlist management cannot be granted through MCP. In the
   default `full` access mode mutations and execute mode are authorized without a grant; in
   `restricted` mode they are refused with `APPROVAL_DENIED`. `background_safe` classification is
   available through `macro_classify_background`. Consequently, in `restricted` mode there is no
   in-band path to grant approvals or manage the allowlist, so the default execute `excel_open` is
   denied and only `inspect`+`read_only` sessions are usable; the remedy is to run in `full` mode or
   pre-authorize via the library.
4. **No PivotTable/PivotCache creation or deletion.** Only existing objects can be managed;
   external/OLAP/Data-Model sources are rejected.
5. **Range and ListObject intersection.** Structural range operations over tables are rejected
   with `LISTOBJECT_INTERSECTION`; use the Tables tools instead.
6. **Bounded reads.** `range_read_values`, `range_read_formulas` and `table_read_rows` return
   bounded windows; read large ranges in chunks via `offset`/`limit`.
7. **No charts, PivotCharts or slicers.** No tools are provided for them.
8. **No workbook creation.** `excel_open` opens only existing files; a new workbook cannot be
   created through MCP.
9. **No full recalculation or cross-workbook copy.** `Application.Calculate` is not exposed;
   `range_copy` works only within one workbook and without the clipboard.
10. **No list pagination.** `tables_list` and `pivot_tables_list` return the full bounded list
    without paging.
11. **Dead owned Excel is detected as `needs_recovery`, but recovery is close/reopen only.**
    `excel_health` reads worker/session bookkeeping without touching COM and additionally runs a
    read-only owned-pid liveness probe (`OpenProcess` + `GetExitCodeProcess`, no COM, no kill): if
    the hidden `EXCEL.EXE` the worker owns has died while the bookkeeping still says
    `running`/`accepting`, `excel_health` reports `needs_recovery` instead of a false `ok`. There
    is still no MCP-callable recovery tool. Remedy: `excel_close` the affected `workbook_id` and
    re-open, or restart the server process.
12. **`macro_run` has no snapshot precondition (spec deviation, tracked).**
    Unlike content-mutating tools, `macro_run` does not require a prior `snapshot_create`; it runs
    a classified, approved macro without one. The recommended workflow still snapshots first so a
    macro's external side effects have a rollback baseline. This asymmetry is a documented
    limitation, not a guarantee that macros are snapshot-gated.
13. **Session recovery is close/reopen only; there is no `excel_recover_session` tool.**
    `SessionService.recover_session`/`close_session` exist internally but are not exposed as MCP
    tools. `excel_health` surfaces `needs_recovery`, but a client cannot trigger recovery in-band;
    the supported remedy is to close and re-open the workbook (or restart the server).
14. **`macro_policy.classify_operation` is non-authoritative (dead/illustrative table).**
    The `classify_operation` function in `security/macro_policy.py` is retained for illustration
    and is **not** wired into the runtime enforcement path; the live operation class of each tool
    is defined per-service. Its table is stale (it predates the current 88-tool catalogue) and
    must not be read as the authoritative tool classification — use this reference or the service
    code for that.