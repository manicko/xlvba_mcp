# Independent Auditor 1 — Product Capabilities & User Workflows

**Repository:** `xlvba_mcp` (C:\py_exp\xlvba_mcp)
**Auditor role:** Independent Auditor 1 (product capabilities & user workflows)
**Mode:** READ-ONLY. No code or docs modified. Only read/grep and read-only commands were run.
**Date:** 2026-11-28

## 1. Baseline (independently verified)

| Item | Expected | Verified | Evidence |
|------|----------|----------|----------|
| HEAD | f79faf2 on master | ✅ `f79faf2b8797a897429754d71e89f8fb9d031317`, branch `master` | `git rev-parse HEAD` / `git branch --show-current` |
| Working tree | clean | ✅ clean | `git status --porcelain` returned nothing |
| Tool count | 43 | ✅ 43 | `python -c "import asyncio,excel_vba_mcp.server as s; t=asyncio.run(s.mcp.list_tools()); print(len(t))"` → `43` |
| Unit tests | pass | ✅ all pass, no failures | `pytest tests\unit -q` → all dots, 0 failures |

The 43 runtime tool names (MCP-runtime-tested, level 5):
`server_health, workbook, workbook_read, workbook_save_copy, excel_health, excel_recover, snapshot_create, snapshot_list, snapshot_restore, snapshot_diff, test_suite, vba_read, vba_export_project, vba_replace_component_source, vba_insert_code, vba_compile_diagnostics, macro, macro_run, job_get, job_request_cancel, sheet, sheet_read, sheet_delete, table, table_read, table_delete_rows, table_delete_column, pivot_read, pivot_refresh, pivot_field, range_read, range_write, range_structural, range_delete, range_merge, defined_name, defined_name_read, defined_name_delete, shape, control, xlvba_environment, workbook_snapshot_read, data_export`.

## 2. Scope / files inspected

- **Tool registration:** `src/excel_vba_mcp/server.py` (`_register_tools`, L2030–2282; resources L2495–2608).
- **API handlers (all 19 modules):** `api/workbook_tools.py`, `vba_tools.py`, `macro_tools.py`, `compile_tools.py`, `job_tools.py`, `test_tools.py`, `sheet_tools.py`, `table_tools.py`, `range_tools.py`, `pivot_tools.py`, `name_tools.py`, `shape_tools.py`, `control_tools.py`, `environment_tools.py`, `export_tools.py`, `ooxml_tools.py`, `health_tools.py`, `recovery_tools.py`, `__init__.py`.
- **Application services:** `application/*` (session, vba_project, macro, compile, snapshot, test_runner, sheets, tables, ranges, pivot_tables, names, shapes, controls, data_export, ooxml, job_registry, workbook_model).
- **Adapters:** `adapters/excel/*`, `adapters/ooxml/*`, `adapters/windows_ui/*` (existence + role only).
- **Tests:** `tests/unit/*` (executed), `tests/integration_excel/*` (enumerated; real `.xlsm` fixtures present; not re-executed this session), `tests/fixtures/*.xlsm`.
- **Docs:** `README.md`, `docs/05-mcp-tools-reference.md`, `docs/06-excel-capabilities.md`, `docs/07-workflows-and-examples.md`.

### Evidence-grading legend (used in every table)
1. **Source-verified** — read in code, no test.
2. **Unit-tested** — passing `tests/unit` test (executed this session).
3. **Integration-tested** — `tests/integration_excel` test present (real Excel; not re-run by auditor).
4. **Real-workbook-tested** — fixture `.xlsm` exercised by an integration test.
5. **MCP-runtime-tested** — observed via live `mcp.list_tools()`.
6. **Unverified/unsupported** — no implemented path, or not validated on target build.

> Grading rule applied: a domain is graded at the **highest** level with concrete evidence. "Integration-tested (3/4)" means a real-workbook test exists in the repo; the auditor did not re-execute the Excel-dependent suite, so it is **not** claimed as re-verified this session.

## 3. Per-domain capability table

| # | Domain | User problem solved | Implemented operations (handler → service) | Verification | Key limitations |
|---|--------|--------------------|--------------------------------------------|--------------|-----------------|
| 1 | **Workbook / session lifecycle** | Open an existing workbook as a safe hidden working copy; close/save without touching the original | `workbook` (open/close/save) → `SessionService.open_workbook/close_workbook/save_workbook`; `workbook_read` (status/list/model) → `get_status/list_workbooks`; `workbook_get_model` → `WorkbookModelService.get_workbook_model` | 2 (unit: `test_workbook_dispatch`, `test_session_service`, `test_workbook_model_service`), 3 (`test_workbook_session`, `test_session_service`, `test_workbook_model_service`), 5 (tool names) | Original never modified; edits live only in working copy; no new-workbook creation; non-`.xlsm` unverified (6) |
| 2 | **Save-as / independent copy** | Produce a new/again-saved copy (Save As) without rebinding the session | `workbook_save_copy` → `excel_save_as` → `SessionService.save_as_workbook` (atomic temp→replace, `.bak`, size+SHA-256 verify, fail-closed `outcome`) | 2 (`test_save_as` 15 tests, `test_export_dispatch`), 3/4 (`test_full_access_workflow::test_persistence_roundtrip`) | Existing target needs `overwrite_existing`; overwriting source needs `overwrite_source` + matching hash |
| 3 | **VBA project read** | Inspect components, procedures, references, search code | `vba_read` (project-info/list-components/list-procedures/read-component/get-references/search-code) → `VBAProjectService.*`; `vba_export_project` → `export_project` | 2 (`test_vba_project_service`, `test_vba_dispatch`), 3/4 (`test_vba_project_read`) | Requires VBOM trust (`ACCESS_VBOM_DISABLED`) + unlocked project (`VBA_PROJECT_LOCKED`) |
| 4 | **VBA editing** | Replace a module / insert code at a semantic position | `vba_replace_component_source`, `vba_insert_code` → `replace_component_source/insert_code` (guarded; `expected_source_hash` revision check) | 2 (`test_vba_mutation`, `test_vba_dispatch`), 3/4 (`test_vba_mutation`, `test_vba_project_write`) | Standard/Class modules only; Document/UserForm rejected; `REVISION_CONFLICT` on stale hash |
| 5 | **Compile diagnostics** | Get a compile verdict / diagnostics | `vba_compile_diagnostics` → `CompileService.compile_project`; `vba_read` action `get-compile-diagnostics` / `vba_get_compile_diagnostics` → `get_last_diagnostics` | 2 (`test_compile_service`, `test_vba_compile`), 3 (`test_compile_service`, `test_vba_compile`, `test_vbe_compile_poc`) | **Always `indeterminate`** / `NON_ACTIVATING_COMPILE_UNAVAILABLE`; never `compiled`/`compile_error` (by design) |
| 6 | **Macro execution** | Run an approved, non-interactive macro and get a typed result | `macro_run` → `MacroService.run_macro`; `macro` (list/classify-background/get-last-result); `macro_get_last_result` | 2 (`test_macro_service` 55, `test_macro_policy`, `test_macro_test_dispatch`), 3/4 (`test_macro_execution` 11, `test_macro_run_poc`) | Deny-by-default: needs approval + `background_safe` (BG06) + R13 veto pass; 90 s soft deadline; hang ⇒ self-recovery; no dialog dismissal |
| 7 | **Jobs / cancellation** | Run long ops asynchronously; poll; record cancel intent | `job_get` → `JobRegistry.status`; `job_request_cancel` → `cancel` (intent-only); `start_long_op` producer used by open/macro/test/pivot | 2 (`test_job_registry`, `test_job_tools`, `test_job_store`, `test_job_recovery_interrupt`) | Cancel is intent-only; never interrupts a running COM call; in-memory registry |
| 8 | **Snapshots** | Immutable baseline; list/diff/restore to roll back a change | `snapshot_create/list/restore/diff` → `SnapshotService.*` | 2 (`test_snapshot_service`, `test_snapshot_store`), 3/4 (`test_snapshot_service`, `test_snapshot_poc`, `test_snapshot_restore_hardening`) | Restore is destructive to working copy, needs `confirm=true` + `workbook_id` (R2); does not undo external macro side effects; diff is metadata+hash only |
| 9 | **Declarative test runner** | Run bounded edit→macro→assert suites on disposable copies | `test_suite` (run/get-run) → `TestRunner.run_suite/get_run`; closed vocab (37 actions, 29 assertions) | 2 (`test_test_runner`, `test_test_spec`), 3/4 (`test_test_runner`, `test_end_to_end_test_runner`, `test_test_runner_ranges_pivots`) | `extra=forbid`; VBA replace/insert steps rejected (no hash); restricted mode blocked; external side effects not rolled back |
| 10 | **Sheets** | List/get/add/rename/copy/delete/visibility | `sheet_read` (list/get); `sheet` (add/rename/copy/set-visibility); `sheet_delete` → `SheetsService.*` | 2 (`test_sheets_service`, `test_sheet_dispatch`), 3/4 (`test_sheets_service`, `test_sheets_copy_target_name_live`) | `add` worksheet-only; copy is in-workbook; delete refuses zero visible; no move/reorder |
| 11 | **Tables (ListObject)** | List/get/read rows; append/insert/update/delete rows; add/delete column; create/resize | `table_read` (list/get/read-rows); `table` (create/resize/append-rows/insert-row/update-rows/add-column); `table_delete_rows`, `table_delete_column` → `TablesService.*` | 2 (`test_tables_service` 42, `test_table_dispatch` 37), 3/4 (`test_tables_service` 16) | Bounded reads (≤2000 cells, `truncated`); ambiguous `where` rejected; delete-column standalone |
| 12 | **Ranges / cells** | Read/write values & formulas, format, validate, CF, copy/insert/delete/merge/unmerge | `range_read` (6 verbs); `range_write` (5 verbs); `range_structural` (copy/insert/unmerge); `range_delete`; `range_merge` → `RangesService.*` | 2 (`test_ranges_service_*` ~190, `test_range_dispatch` 45), 3/4 (`test_ranges_service` 42 + 5 more, `test_phase2_ranges_pivots_regression` 19) | 100k cells/target, 50k/response; `=`-string stored as text; ListObject intersection rejected; explicit shift required; 11 format fields; CF plain only |
| 13 | **Defined names** | List/get/create/update/delete workbook & worksheet names | `defined_name_read` (list/get); `defined_name` (create/update); `defined_name_delete` → `NamesService.*` | 2 (`test_names_service` 33, `test_name_dispatch` 24), 3/4 (`test_names_service` 10) | Explicit scope; duplicates/cross-scope rejected; `RefersTo` opaque (no eval) |
| 14 | **PivotTables** | Inventory; refresh table/cache; field layout/position/aggregation/properties; manual filters; read rendered result | `pivot_read` (6 verbs incl. `result-read`); `pivot_refresh` (refresh-table/refresh-cache); `pivot_field` (6 verbs) → `PivotTablesService.*` | 2 (`test_pivot_*` ~230), 3/4 (`test_pivot_*` ~90, `test_phase2_ranges_pivots_regression`) | Database/Table sources only; external/OLAP/Data Model `unsupported_source`; manual filters only; no create/delete; `RefreshAll` never used |
| 15 | **Shapes** | List/get shapes; assign/clear `OnAction` | `shape` (list/get/set-on-action) → `ShapesService.*` | 2 (`test_shapes_controls` 23, `test_shape_control_dispatch` 19), 3/4 (`test_shapes_controls` 12) | Inventory + OnAction only; macro never executed; no create/delete/format |
| 16 | **Controls (Form/ActiveX)** | List/get; assign `OnAction` to Form Controls | `control` (list/get/set-on-action) → `ControlsService.*` | 2 (`test_shapes_controls`, `test_shape_control_dispatch`), 3/4 (`test_shapes_controls`) | ActiveX rejected for OnAction; ActiveX event proc is inferred (6); no create/delete |
| 17 | **Large-data export** | Stream a range/table/pivot result to CSV/JSONL artifact (bounded memory) | `data_export` → `DataExportService.export` | 2 (`test_data_export_service` 8, `test_export_dispatch` 4, `test_large_data_*`) | Returns metadata only (path/rows/bytes/sha256); `limit_rows` may truncate |
| 18 | **Non-COM snapshot read (OOXML)** | Read metadata / sheet rows / saved PivotCache records from an immutable snapshot without the STA worker | `workbook_snapshot_read` → `OoxmlService.read_metadata/read_sheet_rows/read_pivot_cache_records` | 2 (`test_ooxml_service` 4, `test_ooxml_dispatch` 8, `test_ooxml_worksheet`, `test_ooxml_pivotcache`, `test_ooxml_service_pivotcache`) | Reads saved cache records, not the displayed result; `unsupported_source` for no local records |
| 19 | **Environment / health / recovery** | Discover access mode/policy/capabilities; check worker health; recover a wedged owned process | `xlvba_environment` (non-COM); `excel_health` (non-COM + pid liveness); `excel_recover` → `SessionService.recover_owned_process`; `server_health` | 2 (`test_environment_tools`, `test_health_tools`, `test_recovery` 11, `test_block_a_recovery` 58), 3/4 (`test_recovery` 4) | `excel.build` always null; recovery needs `confirm=true`, pid-scoped, originals safe; no auto-terminate |

### Domains NOT implemented (verified absent — no handler, no service)
- New-workbook creation; PivotTable/PivotCache create/delete (6, by design).
- Charts / PivotCharts / slicers; row/column sizing, freeze panes, page setup; worksheet/workbook protection changes; forced recalculation (`Application.Calculate`); external links/connections/queries; import/export/print/PDF/email; event registration; GUI automation/dialogs/keystrokes. (Source-verified absent: no handler in `api/*`, no service method in `application/*`.)

## 4. Verified end-to-end workflows

Legend: **Verified** = a real-workbook integration test composes the full chain (level 3/4). "Source-composed" = every link exists in code and unit tests cover the links, but no single integration test drives the whole chain.

| Candidate workflow | Status | Evidence |
|--------------------|--------|----------|
| Open → inspect model → close (read-only) | **Verified** | `test_mcp_full_access_workflow::test_full_access_workflow_e2e`; `test_workbook_session` |
| Develop/insert VBA into an existing workbook (read hash → replace/insert) | **Verified** | `test_vba_mutation`, `test_vba_project_write` (replace + insert with revision guard) |
| Run a macro and check result (classify → run → get-last-result) | **Verified** | `test_macro_execution`, `test_macro_run_poc`; `test_end_to_end_test_runner::test_sequential_mutation_macro_workflow_uses_background_safe` |
| Debug via repeated declarative test runs (edit→macro→assert, isolated cases) | **Verified** | `test_end_to_end_test_runner` (7 tests incl. isolation, compile-error-blocks-macro, runtime-error evidence) |
| Restore after a failed change (snapshot → mutate → restore) | **Verified** | `test_snapshot_restore_hardening` (5), `test_snapshot_service` |
| Edit a large Excel Table (read rows → append/update/delete) | **Verified** | `test_tables_service` (16) |
| Read large ranges in pages (offset/limit/cursor) | **Verified** | `test_ranges_service_read`, `test_stale_cursor_error` |
| Read/export PivotTable results (result-read / data_export) | **Verified (read)** / source-composed (export) | `test_pivot_result_read`, `test_pivot_tables_inventory`; `data_export` unit-tested (`test_data_export_service`) but no single e2e pivot→file test |
| Inspect workbook structure before modifying (model/inventory) | **Verified** | `test_workbook_model_service`, `test_mcp_full_access_workflow` |
| Work with several workbook sessions | **Source-composed** | `SessionService` keyed by `workbook_id`; `test_session_service` covers multi-open bookkeeping; no dedicated multi-session e2e |
| Safely create an independent workbook copy (save-as) | **Verified** | `test_save_as` (15), `test_full_access_workflow::test_persistence_roundtrip` |
| Recover from a hung macro/call | **Verified** | `test_recovery::test_hung_macro_drives_auto_recovery`, `test_excel_health_truthful_while_hung` |
| Large-data export to artifact (range/table/pivot → CSV/JSONL) | **Source-composed** | `test_data_export_service`, `test_export_dispatch` (unit); no real-workbook e2e |
| Non-COM snapshot read while worker busy | **Source-composed** | `test_ooxml_*` unit; no Excel e2e |

## 5. Practical use cases (verified-first)

1. **Audit an unfamiliar workbook before editing.** Open read-only → `workbook_read(model)` / `sheet_read(list)` / `table_read(list)` / `pivot_read(tables-list)` → close. *Verified.* Value: safe reconnaissance with zero mutation risk.
2. **Add a macro to an existing `.xlsm`.** `vba_read(read-component)` to get `source_hash` → `vba_insert_code` (or `vba_replace_component_source`) → `vba_read` to confirm new hash. *Verified.* Value: controlled VBA edits with revision safety.
3. **Run a reviewed macro and capture the result.** `macro(classify-background=true)` → `macro_run` → `macro(get-last-result)`. *Verified.* Value: safe, non-interactive execution with a typed outcome.
4. **Regression-test a change with disposable copies.** `snapshot_create` → `test_suite(run)` with edit→macro→assert cases → `test_suite(get-run)`. *Verified.* Value: repeatable, isolated verification; originals untouched.
5. **Roll back a bad change.** `snapshot_create` → mutate → `snapshot_restore(confirm=true, workbook_id=…)` → re-read. *Verified.* Value: deterministic recovery to a known baseline.
6. **Bulk-edit a large table safely.** `table_read(read-rows)` → `table(update-rows / append-rows)` with unambiguous selectors → `table_read` verify. *Verified.* Value: structure-preserving bulk edits.
7. **Page through a huge range without loading it all.** `range_read(read-values, offset, limit, cursor_version)` loop. *Verified.* Value: bounded memory, stale-cursor protection.
8. **Refresh and read a PivotTable.** `pivot_read(tables-list)` → `pivot_refresh(refresh-table)` → `pivot_read(result-read)`. *Verified (read/refresh).* Value: current pivot output without `RefreshAll`.
9. **Export a large dataset for downstream analysis.** `data_export(source_kind=table/range/pivot-result, fmt=csv|jsonl)` → returns artifact path + sha256. *Source-composed.* Value: bounded-memory file export.
10. **Produce an independent copy (Save As).** `workbook_save_copy(path, overwrite_existing=…)` → verified size+hash + `.bak`. *Verified.* Value: safe externalization without rebinding the session.
11. **Recover a wedged Excel without a server restart.** `excel_health` shows `needs_recovery` → `excel_recover(confirm=true)` → reopen. *Verified.* Value: in-process recovery, originals safe.
12. **Discover policy/capabilities before automating.** `xlvba_environment` (access mode, allowed roots, macro policy, capabilities) + `server_health`. *Unit-tested (2).* Value: adapt behavior to full vs restricted mode.

## 6. Discrepancies vs current docs (severity)

| # | Discrepancy | Severity | Evidence |
|---|-------------|----------|----------|
| D1 | **`docs/06-excel-capabilities.md` uses pre-consolidation per-verb tool names** (`excel_open`, `workbook_get_model`, `sheets_list`, `names_list`, `tables_list`, `pivot_tables_list`, `range_get`, `shapes_list`, `vba_project_info`, `macro_list`, `test_suite_run`, `workbook_file_read`, …). 66 old-name occurrences vs 9 consolidated. The live surface is the 43 consolidated tools. | **HIGH** | name scan: doc06 old=66/new=9; `server.py` L2030–2282; `mcp.list_tools()`=43 |
| D2 | **`docs/07-workflows-and-examples.md` uses old tool names** (`excel_open`, `workbook_get_model`, `sheets_list`, `excel_close`) in every recipe; 37 old vs 8 new. Also calls `xlvba_environment` "planned" though it is implemented. Examples are explicitly "not executed". | **HIGH** | name scan doc07 old=37/new=8; L25–27, L41, L62–74 |
| D3 | **Doc 06 capability matrix omits whole implemented domains:** large-data export (`data_export`), non-COM snapshot read (`workbook_snapshot_read`), jobs/cancellation (`job_get`/`job_request_cancel`), environment (`xlvba_environment`), save-as copy (`workbook_save_copy`). Matrix has 44 rows but none for these. | **MEDIUM** | doc06 grep for those names = 0 hits; tools registered in `server.py` |
| D4 | **Doc 06 "Application management" row lists `excel_list_workbooks`/`excel_get_status`** which no longer exist as tools (now `workbook_read` actions). | **MEDIUM** | doc06 L80; `server.py` L2038–2043 |
| D5 | **README says "4 read-only MCP resources"** but `server.py` declares 5 `@mcp.resource` (workbook model, vba components, vba component source, test run, glossary). Likely the glossary is excluded from the count; the number is ambiguous. | **LOW** | README L215/256; `server.py` L2495–2608 |
| D6 | Doc 06 "Coverage summary" counts (18/15/9/2=44) are internally consistent but describe the pre-consolidation surface; counts should be restated against the 43-tool surface. | **LOW** | doc06 L346–356 |

**Docs that are CURRENT (no action):** `README.md` (43 tools, consolidated names, resource note) and `docs/05-mcp-tools-reference.md` (explicitly "43 tools", consolidated names, truthful annotations). These are the authoritative references.

## 7. Recommended doc changes (advisory)

1. **[DOC-UPDATE] Rewrite `docs/06` tool references to the 43 consolidated names** (map each per-verb name to its dispatcher + `action`, e.g. `sheets_list`→`sheet_read(list)`, `range_get`→`range_read(get)`, `pivot_tables_list`→`pivot_read(tables-list)`, `vba_project_info`→`vba_read(project-info)`, `macro_list`→`macro(list)`). *Effort: medium. Priority: recommended.* Benefit: docs match the live surface agents actually call.
2. **[DOC-UPDATE] Add capability rows to `docs/06`** for `data_export`, `workbook_snapshot_read`, `job_get`/`job_request_cancel`, `xlvba_environment`, `workbook_save_copy`. *Effort: small.* Benefit: the capability map becomes complete.
3. **[DOC-UPDATE] Rewrite `docs/07` recipes with consolidated names** and change "planned `xlvba_environment`" to "implemented". *Effort: medium.* Benefit: copy-pasteable, correct examples.
4. **[DOC-UPDATE] Reconcile README resource count** (state 4 data resources + 1 glossary, or 5 total). *Effort: trivial.*
5. **[BEST-PRACTICE] Add an integration test that drives pivot-result → `data_export`** and a multi-session e2e, to upgrade those workflows from "source-composed" to "verified". *Effort: small.* Benefit: closes the two remaining evidence gaps.
6. **[BEST-PRACTICE] Version-stamp the docs' tool list** (e.g., "surface: 43 tools @ f79faf2") so future drift is detectable. *Effort: trivial.*

## 8. Auditor verdict

- **Implemented domains:** 19 of 19 audited domains have real handlers + substantive services (no stubs found; `NotImplementedError`/TODO grep = 0). All 43 tools are MCP-runtime-visible.
- **Not implemented (by design, verified absent):** new-workbook creation, PivotTable/PivotCache create-delete, charts/slicers, layout/protection, recalculation, external links, import/export/print/email, events, GUI automation.
- **Verified end-to-end:** 11 of 14 candidate workflows are backed by real-workbook integration tests; 3 (pivot→export, multi-session, non-COM snapshot read e2e) are source-composed/unit-tested only.
- **Top doc risk:** docs 06 & 07 are stale (old tool names, missing domains) while README + doc 05 are current — HIGH severity for 06/07.
- **No files modified.**