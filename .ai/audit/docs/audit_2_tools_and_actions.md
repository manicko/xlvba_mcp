# Independent Audit 2 — MCP Tools & Actions Catalogue

- **Auditor:** Independent Auditor 2 (MCP tools & actions catalogue)
- **Repo:** `xlvba_mcp` @ `C:\py_exp\xlvba_mcp`
- **Mode:** READ-ONLY (no code/docs modified). Only read-only commands + `pytest` run.
- **Date:** 2026-11-28
- **Baseline verified:** HEAD `f79faf2b8797a897429754d71e89f8fb9d031317`, branch `master`, `git status --porcelain` clean.

## Evidence grading legend (mandatory)

| Grade | Meaning |
|-------|---------|
| (1) Source-verified | Read directly in source (`server.py`, `api/*.py`). |
| (2) Unit-tested | Covered by a passing `tests/unit/*` test. |
| (3) Integration-tested | Covered by `tests/integration_excel/*` (real Excel). |
| (4) Real-workbook-tested | Exercised against a real `.xlsm` fixture. |
| (5) MCP-runtime-tested | Observed via live `mcp.list_tools()` / `list_resources()` / `list_resource_templates()`. |
| (6) Unverified | Not confirmed by any of the above. |

Primary runtime evidence command (grade 5):

```
& .\.venv\Scripts\python.exe -c "import asyncio,json,excel_vba_mcp.server as s; t=asyncio.run(s.mcp.list_tools()); print(len(t)); [print(x.name, sorted((x.input_schema or {}).get('properties',{}).get('action',{}).get('enum',[]))) for x in t]"
```

Contract tests run (grade 2), all 30 passed:

```
& .\.venv\Scripts\python.exe -m pytest tests\unit\test_tool_rename_contract.py tests\unit\test_description_contract.py tests\unit\test_annotation_matrix.py -q
# -> 30 passed
```

---

## 1. Baseline verification (independently reproduced)

| Check | Result | Evidence |
|-------|--------|----------|
| HEAD | `f79faf2b8797a897429754d71e89f8fb9d031317` | `git rev-parse HEAD` (grade 5) |
| Working tree | clean | `git status --porcelain` empty (grade 5) |
| Tool count | **43** | `mcp.list_tools()` len == 43 (grade 5); `test_description_contract.py::test_tool_count_is_final_exposed` (grade 2) |
| Dispatch tools (with `action`) | **20** | runtime action enums non-empty (grade 5) |
| Standalone tools (no `action`) | **23** | 43 − 20 (grade 5) |
| Total distinct actions | **74** | sum of action enums (grade 5) |
| 6 canonical renamed names present | **yes** | `test_tool_rename_contract.py::test_all_six_canonical_names_registered` (grade 2) + runtime (grade 5) |
| 6 old names absent | **yes** | `test_all_six_old_names_absent` (grade 2) + runtime (grade 5) |
| `test_suite` action enum | **`["get-run","run"]`** | runtime (grade 5); `test_test_suite_action_enum_is_run_get_run` (grade 2) |
| Annotation totals (R/D/I) | **14 / 16 / 23** | runtime dump (grade 5); `test_annotation_totals_are_14_16_23` (grade 2) |
| Resources | 1 static (`xlvba://glossary`) | `mcp.list_resources()` (grade 5) |
| Resource templates | 4 | `mcp.list_resource_templates()` (grade 5) |

**The 6 renamed canonical names (runtime-exact, old name → canonical):**
`excel_save_as`→`workbook_save_copy`, `test`→`test_suite`, `name`→`defined_name`, `name_read`→`defined_name_read`, `name_delete`→`defined_name_delete`, `workbook_file_read`→`workbook_snapshot_read`. (grade 2: `test_tool_rename_contract.py::_RENAMED`; grade 5: runtime.)

> Note: the internal Python handler functions keep the OLD names (`excel_save_as`, `test`, `name`, `name_read`, `name_delete`, `workbook_file_read`); only the **public MCP name** is the canonical one. This is by design and asserted by `test_renamed_tools_map_to_expected_handlers`.

---

## 2. README summary table (43 rows)

> One row per tool. "Actions" lists the runtime `action` enum values; standalone tools show `—` (single implicit action = tool name). R/D/I = runtime annotations.

| Domain | Tool | Available actions | What it lets you do |
|--------|------|-------------------|---------------------|
| Lifecycle | `workbook` | `open`, `close`, `save` | Open an allowlisted workbook as a hidden working copy; close it (refuse-by-default when dirty); save in place. |
| Lifecycle | `workbook_read` | `status`, `list`, `model` | Read session/workbook status, list open workbooks, or get the aggregated Stage 1 model. |
| Lifecycle | `workbook_save_copy` | — | Save the working copy to a (possibly new) target path (fail-closed Save As) without rebinding the session. |
| Lifecycle | `excel_health` | — | Worker/session health (state, owned pid, active op, needs_recovery) without touching COM. |
| Lifecycle | `excel_recover` | — | Agent-directed recovery of the owned Excel process (fail-closed; originals safe; `confirm=true`). |
| Snapshots | `snapshot_create` | — | Create an immutable snapshot (manifest + SHA-256) of an open/known workbook. |
| Snapshots | `snapshot_list` | — | List snapshots and their metadata. |
| Snapshots | `snapshot_restore` | — | Restore a working session from a snapshot (discards current copy; `confirm=true`). |
| Snapshots | `snapshot_diff` | — | Compare two snapshots by manifest metadata and file hash. |
| Snapshots | `workbook_snapshot_read` | `metadata`, `sheet-rows`, `pivot-cache-records` | Non-COM streaming OOXML read of an immutable snapshot file (metadata / sheet rows / saved pivot-cache records). |
| Jobs | `job_get` | — | Poll one background job's status/result by `job_id` (never touches COM). |
| Jobs | `job_request_cancel` | — | Record a job's cancel **intent** (intent-only; never interrupts the COM call). |
| VBA | `vba_read` | `project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics` | Read the VBA project surface (summary, components, procedures, source, references, search, cached compile diagnostics). |
| VBA | `vba_export_project` | — | Export every VBA component (`.bas/.cls/.frm+.frx`) and return a manifest with per-file hashes. |
| VBA | `vba_replace_component_source` | — | Fully replace a Standard/Class module's source (revision-guarded). |
| VBA | `vba_insert_code` | — | Insert code at a semantic position in a Standard/Class module (revision-guarded). |
| VBA | `vba_compile_diagnostics` | — | Fail-closed compile diagnostic; always `indeterminate` (does NOT compile). |
| Macro | `macro` | `list`, `classify-background`, `get-last-result` | List callable macros; record a procedure's `background_safe`; re-read the last run result. |
| Macro | `macro_run` | — | Run one approved, `background_safe` macro of an execute-mode workbook (gated). |
| Macro/Test | `test_suite` | `run`, `get-run` | Run a declarative test suite on disposable copies, or read a persisted run. |
| Sheets | `sheet` | `add`, `rename`, `copy`, `set-visibility` | Add/rename/copy a worksheet; set visibility. |
| Sheets | `sheet_read` | `list`, `get` | List all sheets (all types) or get one sheet's metadata. |
| Sheets | `sheet_delete` | — | Delete a sheet (destructive; rejected if it would leave no visible sheet). |
| Tables | `table` | `create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column` | Create/resize a ListObject; append/insert/update rows; add a column. |
| Tables | `table_read` | `list`, `get`, `read-rows` | List tables, get one table, or read a bounded typed slice of rows. |
| Tables | `table_delete_rows` | — | Delete rows by explicit indices or an unambiguous `where` (XOR guard). |
| Tables | `table_delete_column` | — | Delete a column by its unique header. |
| Names | `defined_name` | `create`, `update` | Create a name in an explicit scope; update `RefersTo` (scope preserved). |
| Names | `defined_name_read` | `list`, `get` | List all defined names, or get one by name + explicit scope. |
| Names | `defined_name_delete` | — | Delete a defined name and return its previous value. |
| Ranges | `range_read` | `get`, `read-values`, `read-formulas`, `get-format`, `get-validation`, `get-conditional-formatting` | Resolve an address; read typed values/formulas/format/validation/conditional-format. |
| Ranges | `range_write` | `write-values`, `write-formulas`, `set-format`, `set-validation`, `set-conditional-formatting` | Write typed values/formulas; set format/validation/conditional-format. |
| Ranges | `range_structural` | `copy`, `insert`, `unmerge` | Copy a range; insert cells (`right`/`down`); unmerge. |
| Ranges | `range_delete` | — | Delete cells with a `left`/`up` shift. |
| Ranges | `range_merge` | — | Merge a range into one cell (data-loss preflight). |
| Pivot | `pivot_read` | `tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters`, `result-read` | PivotTable/cache inventory & metadata; field list & filter state; paged rendered result. |
| Pivot | `pivot_refresh` | `refresh-table`, `refresh-cache` | Refresh one PivotTable (table-scoped) or one PivotCache (cache-scoped). |
| Pivot | `pivot_field` | `field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters` | Configure a field's layout/position/aggregation/properties; apply/clear a manual filter. |
| Objects | `shape` | `list`, `get`, `set-on-action` | List/read drawing shapes; assign/clear a shape's `OnAction` macro (not executed). |
| Objects | `control` | `list`, `get`, `set-on-action` | List/read Form/ActiveX controls; assign/clear a Form Control's `OnAction` (not executed). |
| Capability | `xlvba_environment` | — | Report access mode, policy, build (always null), capabilities — no COM. |
| Capability | `server_health` | — | Pure liveness probe (`status/service/version`); never touches Excel/COM. |
| Data export | `data_export` | — | Stream a large range/table/pivot result to a CSV/JSONL file artifact (bounded memory). |

---

## 3. MCP resources & templates (grade 5)

### Static resource (1)

| URI | Name | MIME | Content |
|-----|------|------|---------|
| `xlvba://glossary` | `glossary` | `text/markdown` | Static jargon glossary (R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, SUSPECTED_STALLED, interrupted_unknown, indeterminate, working copy, snapshot gate, guarded mutation, approval gate). Never touches Excel. |

### Resource templates (4)

| URI template | Name | MIME | Content |
|--------------|------|------|---------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` | Read-only Stage 1 workbook model (sheets, tables, names, shapes, controls, optional VBA) + section statuses + content hash. |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` | Read-only VBA component inventory + `vba_access` (reports state instead of failing when VBA unavailable/locked). |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` | Read-only VBA source text of one component. |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` | Read-only summary of one persisted declarative test run. |

All resources are read-only: no workbook mutation, no VBA execution. (grade 1: `server.py:2495-2611`; grade 5: runtime list.)

---

## 4. Comprehensive reference catalogue (per tool + per action)

Conventions: **R/D/I** = runtime annotations. "Guarded mutation" = writable session + snapshot + approval gate + postcondition re-read (see `xlvba://glossary`). Schemas are available over MCP (`input_schema`); not pasted here.

### 4.1 Lifecycle / session

#### `workbook` — R=F D=T I=F (grade 5)
Purpose: open/close/save an open workbook through one lifecycle tool. Original file never modified.
- **`open`** — Copy an allowlisted file into `work/`, open hidden. `mode="execute"` (default; macros by UI) or `"inspect"` (macros disabled); `read_only=false` default; `background=false`. Returns `WorkbookInfo` (or `RunningJob` if background). Errors: `PATH_NOT_ALLOWED`, `APPROVAL_DENIED` (restricted + writable/execute), `EXCEL_HUNG`. (grade 1 `server.py:180-226`; grade 2 `test_workbook_dispatch.py`.)
- **`close`** — Refuse-by-default (D1): clean workbook closes; dirty refuses unless `discard_changes=true`; `save_changes=true` saves then closes; contradictory flags rejected. Returns `SessionStatus`. (grade 1; grade 2 `test_confirm_flags.py`.)
- **`save`** — Persist working copy in place; returns `WorkbookInfo` with `dirty=false`. Never implicit. (grade 1; grade 2.)

#### `workbook_read` — R=T D=F I=T
Purpose: read workbook/session state without mutating.
- **`status`** — `SessionStatus` (busy, read-only, dirty, active op, policy block). No COM. (grade 1.)
- **`list`** — `list[WorkbookInfo]` of open workbooks from session bookkeeping. (grade 1.)
- **`model`** — Aggregated Stage 1 `WorkbookModel` (sheets/tables/names/shapes/controls + optional VBA). Never executes VBA; `vba_access` reports state. `include_vba=true` default. (grade 1 `server.py:995-1001`; grade 3 `test_workbook_model_service.py`.)

#### `workbook_save_copy` — R=F D=T I=F (canonical; handler `excel_save_as`)
Purpose: Save As to a (possibly new) path without rebinding the session. Flushes with `wb.Save` on STA, then atomic temp→`os.replace` (never `wb.SaveAs`); single `<target>.bak`.
- Params: `workbook_id`, `path`, `overwrite_existing=false`, `overwrite_source=false`.
- New path → `outcome="ok"`; existing target → `PROTECTED_TARGET/target_exists` unless `overwrite_existing=true`; original source target → `PROTECTED_TARGET/source_protected` unless `overwrite_source=true` (+ source hash re-verify → `STALE_TARGET`); open working copy target always refused. `outcome="ok"` only after size+SHA-256 verify; else `"unknown"/"error"` fail-closed with `backup_path`. Skips snapshot gate. Returns `SaveAsResult`. (grade 1 `server.py:255-275`; grade 2 `test_save_as.py`.)

#### `excel_health` — R=T D=F I=T
Purpose: worker/session health (`state`, `accepting`, owned pid, active op, `needs_recovery`) without COM. Always returns a payload. (grade 1 `server.py:287-291`; grade 2 `test_health_tools.py`.)

#### `excel_recover` — R=F D=T I=F
Purpose: agent-directed recovery of the owned Excel process. Terminates wedged `EXCEL.EXE` via held handle, discards working copies, rebuilds hidden instance; auto-allowed both modes, audited, pid-scoped, fail-closed.
- Params: `reason="agent_directed"`, `confirm=false`. Missing/false `confirm` → `CONFIRMATION_REQUIRED` before any termination. Returns recovery report; interrupted op → `prior_job="unknown"`, never auto-retried. (grade 1 `server.py:292-310`; grade 2 `test_recovery.py`/`test_block_a_recovery.py`; grade 3 `test_recovery.py`.)

### 4.2 Snapshots

#### `snapshot_create` — R=F D=F I=F
Create immutable snapshot (`SaveCopyAs` for open books, file copy for closed) + `manifest.json` + SHA-256. Approval-gated. Returns `SnapshotManifest`. Errors `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `EXCEL_HUNG`. (grade 1; grade 2 `test_snapshot_service.py`; grade 3 `test_snapshot_service.py`.)

#### `snapshot_list` — R=T D=F I=T
List snapshots + metadata. Returns `list[SnapshotManifest]`. (grade 1; grade 2.)

#### `snapshot_restore` — R=F D=T I=F
Restore a working session from a snapshot. `workbook_id` **required** (the OLD id to close; missing → `INVALID_RANGE`, `isRetryable=false`); `confirm` required (false → `CONFIRMATION_REQUIRED` before close). Opens fresh copy from baseline, validates, closes previous without saving. Returns `SnapshotRestoreResult` (restored `WorkbookInfo` with new id + `previous_close` ∈ closed/failed/not_requested + `previous_close_error`). Re-restoring active baseline → `reason="same_snapshot"`. Does NOT undo external macro side effects. (grade 1 `server.py:316-336`; grade 2 `test_snapshot_service.py`; grade 3 `test_snapshot_restore_hardening.py`.)

#### `snapshot_diff` — R=T D=F I=T
Compare two snapshots by manifest metadata + file hash. Params `snapshot_id_a`, `snapshot_id_b`. Returns `SnapshotDiff`. (grade 1; grade 2.)

#### `workbook_snapshot_read` — R=T D=F I=T (canonical; handler `workbook_file_read`)
Purpose: non-COM streaming OOXML read of an immutable snapshot file (never occupies the STA worker). When `snapshot_id` omitted a fresh snapshot is created.
- **`metadata`** — `OoxmlMetadata` (workbook structure). (grade 1; grade 2 `test_ooxml_dispatch.py`.)
- **`sheet-rows`** — `OoxmlSheetPage`; params `sheet`, `offset` (0-based), `limit`. (grade 1; grade 2 `test_ooxml_worksheet.py`.)
- **`pivot-cache-records`** — `OoxmlPivotCachePage`; param `cache_index` (1-based). Empty page `reason='unsupported_source'` when no saved local records. (grade 1; grade 2 `test_ooxml_pivotcache.py`.)
Errors `INVALID_RANGE`, `SNAPSHOT_FAILED`.

### 4.3 Jobs

#### `job_get` — R=T D=F I=T
Poll one job by `job_id` (never COM). `status` ∈ queued/running/succeeded/failed/expired/interrupted_unknown/unknown. `interrupted_unknown` (outcome unknown, retryable false) never auto-retried; `SUSPECTED_STALLED` = soft deadline exceeded. Returns `JobStatus`. Errors `NOT_FOUND`. (grade 1 `server.py:493-508`; grade 2 `test_job_tools.py`/`test_job_registry.py`.)

#### `job_request_cancel` — R=F D=F I=T
Record cancel **intent** only (`cancel_requested=true` if non-terminal); never interrupts the single-STA COM call. Returns `JobStatus`. (grade 1 `server.py:509-520`; grade 2 `test_job_tools.py`.)

### 4.4 VBA

#### `vba_read` — R=T D=F I=T
Read-only VBA surface dispatcher. `workbook_id` required for all.
- **`project-info`** → `VbaProjectInfo`. (grade 1; grade 3 `test_vba_project_read.py`.)
- **`list-components`** → `list[VbaComponentInfo]` (name, type, lines, source hash, designer/export flags). (grade 1; grade 3.)
- **`list-procedures`** → `list[VbaProcedureInfo]`; needs `component`. (grade 1; grade 3.)
- **`read-component`** → `VbaComponentSource` (source + normalized `source_hash`); needs `component`. (grade 1; grade 3.)
- **`get-references`** → `list[VbaReferenceInfo]`. (grade 1; grade 3.)
- **`search-code`** → `list[VbaSearchMatch]`; params `pattern`, `regex=false`. (grade 1; grade 3.)
- **`get-compile-diagnostics`** → cached `CompileDiagnostics` (re-reads cache; does not re-collect). (grade 1; grade 2 `test_vba_dispatch.py`.)
Errors `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG`.

#### `vba_export_project` — R=F D=F I=F
Export every component (`.bas/.cls/.frm+.frx`) + manifest with per-file hashes; Document modules never removed. Param `out_dir?`. Writes files under artifact root; approval-gated. Returns `VbaExportManifest`. Errors `ACCESS_VBOM_DISABLED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED`. (grade 1; grade 2 `test_export_dispatch.py`.)

#### `vba_replace_component_source` — R=F D=T I=F
Fully replace a Standard/Class module (revision-guarded). Params `component`, `new_source`, `expected_source_hash` (normalized source hash from `vba_read read-component`, NOT snapshot export file hash). Revision check before + inside STA; mismatch → `REVISION_CONFLICT`. Document modules/UserForms → `PROTECTED_TARGET` (`reason=non_replaceable_component`). Not idempotent (re-call same hash fails guard). Returns `VbaMutationResult`. (grade 1 `server.py:403-417`; grade 2 `test_vba_mutation.py`; grade 3 `test_vba_mutation.py`.)

#### `vba_insert_code` — R=F D=T I=F
Insert code at `position` ∈ `declarations` / `before_procedure:<name>` / `after_procedure:<name>`; same revision guard. Returns `VbaMutationResult`. Invalid position → `COM_BUSY`. (grade 1 `server.py:418-428`; grade 2.)

#### `vba_compile_diagnostics` — R=F D=F I=F
Fail-closed compile probe; **always** `status="indeterminate"`, `reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"`, `is_compile=false` on the validated build. Never returns compiled/compile_error; never shows/activates UI. Returns `CompileResult`. (grade 1 `server.py:429-443`; grade 2 `test_compile_service.py`; grade 3 `test_vba_compile.py`/`test_vbe_compile_poc.py`.)

### 4.5 Macro / test

#### `macro` — R=F D=F I=T
- **`list`** → `list[MacroInfo]` (callable macros + `background_safe`). Read. (grade 1; grade 3 `test_macro_execution.py`.)
- **`classify-background`** → `BackgroundSafetyClassification`. UNGATED bookkeeping bound to (workbook hash, source hash, procedure) (BG06); invalidated on VBA source change. Params `procedure`, `background_safe`, optional `expected_project_source_hash` (mismatch → `STALE_TARGET`). Heuristic/interactive-UI procedure cannot be classified safe → `INTERACTIVE_UI_NOT_SUPPORTED`. Never grants run permission. (grade 1 `server.py:465-492`; grade 2 `test_macro_service.py`.)
- **`get-last-result`** → cached `MacroRunResult` / `None`. Read. (grade 1.)
Errors `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `MACRO_NOT_FOUND`, `INTERACTIVE_UI_NOT_SUPPORTED`, `EXCEL_HUNG`.

#### `macro_run` — R=F D=T I=F
Run one approved, `background_safe` macro of an execute-mode workbook. Gate order: (1) execute+writable; (2) procedure resolution; (3) approval; (4) BG06 background_safe for current workbook hash; (5) R13 interactive-UI preflight; (6) `Application.Run`. Params `macro_name`, `args?` (≤30; bare `datetime.date` rejected), `background=false`. Blocked → truthful blocked result, macro NOT dispatched. `background=true` → `RunningJob`. Returns `MacroRunResult`. Errors `APPROVAL_DENIED`, `MACRO_BLOCKED`, `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`. Hang/timeout recovers session in-process (F2b), discards open workbooks, never auto-replays. (grade 1 `server.py:444-464`; grade 2 `test_macro_service.py`; grade 3 `test_macro_run_poc.py`/`test_macro_execution.py`.)

#### `test_suite` — R=F D=T I=F (canonical; handler `test`) — **actions exactly `["run","get-run"]`** (grade 5 + grade 2)
- **`run`** — Validate + run a declarative suite on disposable working copies. Param `suite` (TestSuiteSpec/mapping; `suite.workbook_id` must be open; baseline snapshot auto-created unless `baseline="use_existing"`), `background=false`. Closed `TestAction`/`AssertionType` vocabularies validated before any Excel call; one disposable copy per case; external macro side effects NOT rolled back. Returns `TestRunResult` (or `RunningJob`). Missing `suite` → `INVALID_RANGE`. (grade 1 `server.py:338-372`; grade 2 `test_test_spec.py`/`test_test_runner.py`; grade 3 `test_end_to_end_test_runner.py`.)
- **`get-run`** — Read a persisted run by `run_id`; never touches Excel, never re-runs. Returns `TestRunResult`. Missing `run_id` → structured error. (grade 1; grade 2.)
`"run-suite"` is rejected with `INVALID_RANGE` (grade 2 `test_run_suite_action_rejected_with_invalid_range`).

### 4.6 Sheets

#### `sheet` — R=F D=F I=F
- **`add`** — Add worksheet (`sheet_type="worksheet"` default; others rejected). Returns `SheetInfo`. (grade 1; grade 2 `test_sheet_dispatch.py`; grade 3 `test_sheets_service.py`.)
- **`rename`** — Rename; CodeName never changed. Param `new_name`. (grade 1; grade 3.)
- **`copy`** — In-workbook duplicate; Excel-generated name (`<name> (2)`) unless `target_name` (validated ≤31 chars, unique). Returns `SheetInfo`. (grade 1; grade 3 `test_sheets_copy_target_name_live.py`.)
- **`set-visibility`** — `visible` bool; hiding last visible sheet rejected. (grade 1; grade 3.)
Guarded mutation. Errors `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `NOT_FOUND`, `INVALID_RANGE`, `COM_BUSY`.

#### `sheet_read` — R=T D=F I=T
- **`list`** → `list[SheetInfo]` (all types; name, index, object_type, visibility, code_name, protected, table_count, used_range). (grade 1; grade 3.)
- **`get`** → `SheetInfo`; needs `name`. Errors `NOT_FOUND`. (grade 1.)

#### `sheet_delete` — R=F D=T I=F
Delete a sheet (destructive). Rejected if it would leave no visible worksheet. Returns `SheetInfo`. (grade 1; grade 2; grade 3.)

### 4.7 Tables

#### `table` — R=F D=F I=F
- **`create`** — ListObject over A1 range or defined name; params `sheet`, `range_or_name`, `style?`. (grade 1; grade 2 `test_table_dispatch.py`; grade 3 `test_tables_service.py`.)
- **`resize`** — `ListObject.Resize`; params `name`, `new_range` (header row required). (grade 1; grade 3.)
- **`append-rows`** — `ListRows.Add` (structure preserved); `rows` non-empty `[{header:value}]` matching columns (case-insensitive). Returns new row count. (grade 1; grade 3.)
- **`insert-row`** — Insert one data row at 1-based `at` + `values`. (grade 1; grade 3.)
- **`update-rows`** — `updates=[{index|where, values}]`; ambiguous `where` (0 or >1 rows) → `AMBIGUOUS_FILTER` (no guessing). (grade 1; grade 3.)
- **`add-column`** — `ListColumns.Add(Position)`; unique `header` required; `position?`. (grade 1; grade 3.)
Guarded mutation. Returns `TableMutationResult`. Errors `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`.

#### `table_read` — R=T D=F I=T
- **`list`** → `list[TableInfo]`. (grade 1; grade 3.)
- **`get`** → `TableInfo`; needs `name`. (grade 1.)
- **`read-rows`** → `TableRows` (typed: str/int/float/bool/null/ISO date/Excel error; `truncated` flag); `offset` 0-based, `limit`. (grade 1; grade 3.)

#### `table_delete_rows` — R=F D=T I=F
Delete rows by `indices` (1-based) XOR `where` (exactly one; ambiguous → `AMBIGUOUS_FILTER`). Returns `TableMutationResult`. (grade 1; grade 2; grade 3.)

#### `table_delete_column` — R=F D=T I=F
Delete column by unique `header`. Returns `TableMutationResult`. (grade 1; grade 2; grade 3.)

### 4.8 Names

#### `defined_name` — R=F D=F I=T (canonical; handler `name`)
- **`create`** — Create name/constant in explicit `scope` (`workbook` or `worksheet`+`sheet`); `refers_to` normalized (leading `=` optional). Rejects duplicate/cross-scope conflict. Returns `DefinedNameInfo`. (grade 1; grade 2 `test_name_dispatch.py`; grade 3 `test_names_service.py`.)
- **`update`** — Repoint `RefersTo`; scope preserved (never re-scopes). (grade 1; grade 3.)
Guarded mutation.

#### `defined_name_read` — R=T D=F I=T (canonical; handler `name_read`)
- **`list`** → `list[DefinedNameInfo]` (scope, NameLocal, Visible, RefersTo/RefersToLocal, target_kind). (grade 1; grade 3.)
- **`get`** → `DefinedNameInfo`; needs `name` + `scope` (+ `sheet` for worksheet scope). (grade 1.)

#### `defined_name_delete` — R=F D=T I=F (canonical; handler `name_delete`)
Delete a name in explicit scope; returns the pre-delete `DefinedNameInfo`. (grade 1; grade 2; grade 3.)

### 4.9 Ranges

#### `range_read` — R=T D=F I=T
`sheet` + `address` required for all.
- **`get`** → `RangeAddress` (address, rows, cols, shape, merged/table flags). (grade 1; grade 2 `test_range_dispatch.py`.)
- **`read-values`** → `RangeValues` (typed empty/string/number/boolean/date/error; `truncated`). Caps: request >100000 cells rejected; response ≤50000 cells / 8 MiB. `offset`/`limit`. (grade 1; grade 2 `test_ranges_service_read.py`.)
- **`read-formulas`** → `RangeFormulas` (Formula2 primary, Formula fallback; `_xlfn/_xlws/_xlpm` verbatim + flagged). (grade 1; grade 2.)
- **`get-format`** → `FormatInfo` (tri-state; `mixed_properties`/`unsupported`; never flattened). (grade 1; grade 2 `test_ranges_service_format.py`.)
- **`get-validation`** → `ValidationRule` (`present=false` if none; non-uniform reported). (grade 1; grade 2 `test_ranges_service_validation_cf.py`.)
- **`get-conditional-formatting`** → `list[ConditionalFormatRule]` (priority order; ColorScale/DataBar/IconSet `supported=false` + reason). (grade 1; grade 2.)
Errors `INVALID_RANGE`.

#### `range_write` — R=F D=F I=F (does NOT snapshot — call `snapshot_create` first)
- **`write-values`** — Typed `{kind,value}` cells or raw JSON scalars (literal by real JSON type; leading `=` stays text). Payload `{rows,cols,values}`; rows≤100000, cols≤512, cells≤100000; over-cap → `RANGE_TOO_LARGE`. Table writes only if row/col structure unchanged. Returns `RangeMutationResult`. (grade 1; grade 2 `test_ranges_service_write.py`.)
- **`write-formulas`** — Entries must begin with `=`; else use write-values. Returns `RangeMutationResult`. (grade 1; grade 2.)
- **`set-format`** — Sparse `patch` (only set fields written; null clears where Excel allows). Returns `FormatMutationResult`. (grade 1; grade 2.)
- **`set-validation`** — Add/modify rule (list/whole/decimal/date/time/text-length/custom); Formula1/2 opaque ≤255 chars; existing rule modified (unspecified preserved). Returns `ValidationRule`. (grade 1; grade 2.)
- **`set-conditional-formatting`** — `mode` add/modify/delete; `spec.priority` is identity (required for modify/delete); `spec.type` ∈ plain FormatCondition set only (color_scale/data_bar/icon_set/unsupported rejected); `mode="add"` renumbers sheet-wide priorities → NOT idempotent. Returns `ConditionalFormatMutationResult`. (grade 1; grade 2.)
Guarded mutation. Errors `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET`, `RANGE_TOO_LARGE`.

#### `range_structural` — R=F D=T I=F (NOT idempotent)
- **`copy`** — Explicit `source`+`destination_sheet`+`destination`; disjoint, shape-matched; ListObject intersection rejected (routed to table tools). Returns `StructuralMutationResult`. (grade 1; grade 2 `test_ranges_service_structural.py`.)
- **`insert`** — `address` + required `shift` (`right`/`down`, never guessed) + `allow_merged?`. Rejected on protected sheet / ListObject / spill / merged without allow. (grade 1; grade 2.)
- **`unmerge`** — Discard merges intersecting `address`; rejected if none / protected / ListObject. (grade 1; grade 2.)
Guarded mutation.

#### `range_delete` — R=F D=T I=F
Delete cells at `address` with required `shift` (`left`/`up`). Same rejections as insert. Returns `StructuralMutationResult`. (grade 1; grade 2.)

#### `range_merge` — R=F D=T I=T (dual-hint)
Merge into one cell (or per-row with `across=true`). Data-loss preflight; blocked unless `allow_data_loss=true`. Rejected on protected/ListObject/spill. Returns `StructuralMutationResult`. (grade 1; grade 2.)

### 4.10 Pivot

#### `pivot_read` — R=T D=F I=T
`workbook_id` required for all.
- **`tables-list`** → `PivotInventory`; `background=true` → `RunningJob`. (grade 1; grade 2 `test_pivot_dispatch.py`; grade 3 `test_pivot_tables_inventory.py`.)
- **`table-get`** → `PivotTableInfo`; needs `pivot_table`. (grade 1; grade 3.)
- **`cache-get`** → `PivotCacheInfo`; needs `cache_index`. (grade 1; grade 3.)
- **`fields-list`** → `list[PivotFieldInfo]`; needs `pivot_table`. (grade 1; grade 3.)
- **`field-get-filters`** → `PivotFieldFilters`; needs `pivot_table`+`field`. (grade 1; grade 3 `test_pivot_filters.py`.)
- **`result-read`** → `PivotResultPage` (paged **rendered** `TableRange1` output; works for any source type incl. external/OLAP/Data Model). Params `pivot_table`, `offset` 0-based, `limit?`, `cursor_version?` (prior `page.version`; mismatch → `STALE_CURSOR`, restart at offset 0). Read-only. **(grade 1 `pivot_tools.py:314-352,607-692`; grade 2 `test_pivot_result_read.py`; grade 3 `test_test_runner_ranges_pivots.py`.)**
Errors `COM_BUSY`, `AMBIGUOUS_FILTER`, `STALE_CURSOR` (result-read), `INVALID_RANGE`.

> **DISCREPANCY:** `result-read` is present at runtime but **absent from `docs/05-mcp-tools-reference.md`** (and `docs/ru/05`). See §6, D-01.

#### `pivot_refresh` — R=F D=F I=T (NOT idempotent — re-refresh may read new data)
- **`refresh-table`** — `PivotTable.RefreshTable`; needs `pivot_table`; source-type checked first (external/OLAP/Data Model → `UNSUPPORTED_PIVOT_SOURCE`). Returns `PivotRefreshResult` (affected tables, RefreshDate before/after, refreshed status via truth signal). (grade 1; grade 2 `test_pivot_refresh.py`; grade 3 `test_pivot_refresh.py`.)
- **`refresh-cache`** — `PivotCache.Refresh`; needs `cache_index`; affects every table on the cache. (grade 1; grade 3.)
Guarded mutation.

#### `pivot_field` — R=F D=T I=T (dual-hint)
`pivot_table` + `field` required for every action except `field-clear-filters scope="table"`.
- **`field-set-layout`** — `orientation` ∈ hidden/row/column/page (`data` rejected → use aggregation); `delete=true` destructive (`PivotField.Delete`). Returns `PivotFieldInfo`. (grade 1; grade 2 `test_pivot_field_config.py`; grade 3 `test_pivot_field_config.py`.)
- **`field-set-position`** — `position` (1-based). Returns `PivotFieldInfo`. (grade 1; grade 3.)
- **`field-set-aggregation`** — `function` (sum/average/count/countNums/max/min/product/stDev/stDevP/var/varP) + `caption`. Returns `PivotFieldInfo`. (grade 1; grade 3.)
- **`field-set-properties`** — `properties` (caption/number_format/show_all_items). Returns `PivotFieldInfo`. (grade 1; grade 3.)
- **`field-apply-filter`** — `spec` (manual visible-items only; label/value/date/top-bottom unsupported). Page field rejected before mutation. Returns `PivotFieldFilters`. (grade 1; grade 2 `test_pivot_filters.py`.)
- **`field-clear-filters`** — `scope` field/table. Returns `PivotFieldFilters`. (grade 1; grade 2.)
Guarded mutation. Errors `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`.

### 4.11 Objects

#### `shape` — R=F D=F I=T
- **`list`** → `list[ShapeInfo]` (name, kind, on_action verbatim, visibility, position/size, best-effort text). (grade 1; grade 2 `test_shape_control_dispatch.py`; grade 3 `test_shapes_controls.py`.)
- **`get`** → `ShapeInfo`; needs `name`. (grade 1.)
- **`set-on-action`** — Write `Shape.OnAction` only (macro NOT executed); `macro_name=""` clears. Needs `name`+`macro_name`. Guarded mutation. (grade 1; grade 3.)
Errors `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`; ambiguous duplicate name → `COM_BUSY` (`reason="ambiguous"`).

#### `control` — R=F D=F I=T
- **`list`** → `list[ControlEntry]` (Form Controls: control_type, caption, linked_cell, on_action; ActiveX: prog_id, inferred event proc). (grade 1; grade 3.)
- **`get`** → `ControlEntry`; needs `name`. (grade 1.)
- **`set-on-action`** — Write `OnAction` on a **Form Control only** (ActiveX rejected). `macro_name=""` clears. Guarded mutation. (grade 1; grade 3.)
Errors `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`.

### 4.12 Capability / health

#### `xlvba_environment` — R=T D=F I=T (inline annotation)
Report access mode, policy (`classification_is_authorization=false`), build (always null), capabilities (VBA, macros, pivot database-source/manual filters, background jobs, recovery, non-activating compile diagnostics — NOT compilation), `glossary_uri`, workbooks. Non-COM, side-effect-free. Returns `EnvironmentReport`. (grade 1 `server.py:1103-1135`; grade 2 `test_environment_tools.py`.)

#### `server_health` — R=T D=F I=T (inline annotation)
Pure liveness probe → `{"status":"ok","service":"excel-vba-mcp","version":...}`. Never touches Excel/COM. (grade 1 `server.py:2452-2464`; grade 2 `test_health_tools.py`.)

### 4.13 Data export

#### `data_export` — R=F D=F I=T
Stream a large dataset to a CSV/JSONL file artifact (bounded memory; data never inline; non-destructive, idempotent — new artifact each call). Params `source_kind` ∈ range/table/pivot-result; `sheet`+`address` (range); `table` (table); `pivot_table` (pivot-result); `fmt="csv"`; `limit_rows?`. Returns `DataExportResult` (path, rows, bytes, sha256, truncated/complete). Errors `INVALID_RANGE`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED`. (grade 1 `server.py:1160-1181`; grade 2 `test_data_export_service.py`; grade 3 `test_test_runner_ranges_pivots.py`.)

---

## 5. Representative operations with real examples

### 5.1 `workbook` open (grade 1 `server.py:225`)
```json
{ "action": "open", "path": "C:/data/report.xlsm", "mode": "execute", "read_only": false }
```

### 5.2 `workbook_save_copy` (grade 1 `server.py:273`)
```json
{ "workbook_id": "<id>", "path": "C:/data/report_copy.xlsm", "overwrite_existing": false, "overwrite_source": false }
```

### 5.3 `macro_run` (grade 1 `server.py:622` docs; handler `macro_tools.py`)
```json
{ "workbook_id": "<id>", "macro_name": "Module1.Hello", "args": [] }
```
Preconditions: `macro classify-background` first; execute-mode writable session; approval; R13 preflight.

### 5.4 `test_suite` run (grade 1 `server.py:366`)
```json
{ "action": "run", "suite": { "suite": "smoke", "workbook_id": "<id>",
  "cases": [ { "id": "c1", "steps": [ { "action": "range_write_values", "sheet": "Sheet1", "range_address": "A1", "values_grid": { "rows": 1, "cols": 1, "values": [[1]] } } ] } ] }, "background": false }
```

### 5.5 `table` append-rows (grade 1 `server.py:678`)
```json
{ "action": "append-rows", "workbook_id": "<id>", "name": "Sales", "rows": [ { "Amount": 100 } ] }
```

### 5.6 `range_read` read-values (grade 1 `server.py:851`)
```json
{ "action": "read-values", "workbook_id": "<id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 50 }
```

### 5.7 `range_write` write-values (grade 1 `server.py:883`)
```json
{ "action": "write-values", "workbook_id": "<id>", "sheet": "Sheet1", "address": "A1", "payload": { "rows": 1, "cols": 2, "values": [[1, 2]] } }
```

### 5.8 `pivot_read` result-read (grade 1 `pivot_tools.py:314`) — **undocumented in docs/05**
```json
{ "action": "result-read", "workbook_id": "<id>", "pivot_table": "PivotTable1", "offset": 0, "limit": 100 }
```

### 5.9 `pivot_refresh` refresh-table (grade 1 `pivot_tools.py:358`)
```json
{ "action": "refresh-table", "workbook_id": "<id>", "pivot_table": "PivotTable1" }
```

### 5.10 `pivot_field` field-set-aggregation (grade 1 `pivot_tools.py:739`)
```json
{ "action": "field-set-aggregation", "workbook_id": "<id>", "pivot_table": "PivotTable1", "field": "Amount", "function": "sum", "caption": "Sum of Amount" }
```

### 5.11 `workbook_snapshot_read` sheet-rows (grade 1 `server.py:1157`)
```json
{ "action": "sheet-rows", "workbook_id": "<id>", "sheet": "Sheet1", "offset": 0, "limit": 1000 }
```

### 5.12 `data_export` (grade 1 `server.py:1179`)
```json
{ "workbook_id": "<id>", "source_kind": "range", "sheet": "Sheet1", "address": "A1:C100000", "fmt": "csv" }
```

### 5.13 `snapshot_restore` (grade 1 `server.py:316`; docs example)
```json
{ "snapshot_id": "snap-…", "workbook_id": "<old-id>", "confirm": true }
```
Returns new active `workbook.id`; old id is closed. `previous_close` reports disposition.

---

## 6. Discrepancies vs `docs/05-mcp-tools-reference.md`

The English `docs/05` is otherwise **highly accurate**: all 43 tool names, all standalone tools, the 6 canonical names, `test_suite` enum, annotation totals (14/16/23), resource list, and per-tool action enums match runtime (grades 1+2+5). One action-level gap and one description-level gap found.

| ID | Severity | Type | Detail | Evidence |
|----|----------|------|--------|----------|
| **D-01** | **MEDIUM** | Runtime action missing from docs | `pivot_read` has a 6th action **`result-read`** (paged rendered PivotTable output, `PivotResultPage`, params `offset`/`limit`/`cursor_version`, `STALE_CURSOR` on mismatch) that is **not listed** in `docs/05` §Pivot (line 341 lists only 5 actions; returns omit `PivotResultPage`; params omit `offset`/`limit`/`cursor_version`). Same gap in `docs/ru/05` line 354. | grade 5 runtime enum; grade 1 `pivot_tools.py:314-352,607-692`; grade 2 `tests/unit/test_pivot_result_read.py` |
| **D-02** | **LOW** | Tool description drift (source, not docs) | `server.py::_DESC_PIVOT_READ` (lines 1002-1027) also omits `result-read` from its `Parameters`/`Behavior`/`Result` text even though the schema enum includes it. This is the MCP-facing description agents read; it should mention `result-read`. (Not a docs/05 issue per se, but the same underlying gap.) | grade 1 `server.py:1002-1027` vs `pivot_tools.py:608-615` |
| **D-03** | **INFO** | No tool in docs absent at runtime; no runtime tool absent from docs | All 43 documented tools exist at runtime and vice-versa. No stale/old names (`excel_save_as`, `test`, `name`, `name_read`, `name_delete`, `workbook_file_read`) are documented as public tools. | grade 5 + grade 2 `test_tool_rename_contract.py` |

**Net:** No CRITICAL/HIGH doc-vs-runtime mismatches. The single actionable gap is **D-01** (document `pivot_read` `result-read`), with **D-02** as its source-side twin. Recommend updating `docs/05` (EN + RU) and `_DESC_PIVOT_READ` to include `result-read` — a `[DOC-UPDATE]` (docs are behind code; the code is correct and tested).

---

## 7. Evidence appendix

- Runtime tool/action/annotation dump: `mcp.list_tools()` (grade 5), 43 tools, annotations R/D/I = 14/16/23.
- Runtime resources: `mcp.list_resources()` → 1 (`xlvba://glossary`); `mcp.list_resource_templates()` → 4 (grade 5).
- Contract tests (grade 2, 30 passed): `tests/unit/test_tool_rename_contract.py`, `tests/unit/test_description_contract.py`, `tests/unit/test_annotation_matrix.py`.
- Action enums source-verified (grade 1) in `src/excel_vba_mcp/api/*_tools.py` (`Literal[...]` + dispatch funcs).
- Descriptions/registrations/annotations source-verified (grade 1) in `src/excel_vba_mcp/server.py` (`_DESC_*`, `_register_tools`, `_annotations`, `@mcp.resource`).
- Integration coverage exists (grade 3) for every domain under `tests/integration_excel/*` (e.g. `test_ranges_service_*`, `test_tables_service.py`, `test_pivot_*`, `test_vba_*`, `test_macro_*`, `test_snapshot_*`, `test_shapes_controls.py`, `test_workbook_*`).
- Artifacts: `.ai/audit/artifacts/tools_list_after.json`, `tools_list_before.json`.

_End of report._