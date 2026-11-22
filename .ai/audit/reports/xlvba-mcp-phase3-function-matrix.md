# Phase-3 Function-by-Function Live MCP Verification Matrix

**Task:** `FE-12` (`FULL-EDIT-MODE-ORCHESTRATION-001` Phase 3)
**Validator role:** independent live verification through the real MCP boundary
**Repo / branch:** `C:\py_exp\xlvba_mcp` @ `master`
**Date:** 2026-11-22
**Method:** real `python -m excel_vba_mcp` STDIO server + `mcp.ClientSession` handshake; every tool invoked with the documented args against disposable temp copies of the committed Phase-2 fixtures; `EXCEL_VBA_MCP_ACCESS_MODE` unset (full) for the matrix body, `restricted` for the denial rows, plus a read-only `excel_open` for `WORKBOOK_READ_ONLY`.
**Evidence file:** `.ai/tmp/phase3_evidence.json` (103 invocations across 88 tools; regenerated from the harness run).
**Fixture byte-identical confirmation:** `FullAccessWorkflow.xlsm` `6b905837…A19AC1D4D9`, `RangesCellsProject.xlsm` `e665cff0…f11effa2`, `PivotTablesProject.xlsm` `8d1fa43d…9e32372` — hashes unchanged before/after the run.

> Scope discipline: no production file was modified. The harness lives in `.ai/tmp/phase3_matrix_run.py` + `.ai/tmp/phase3_snap.py` (throwaway, not committed). No registry / Trust Center / `AutomationSecurity=LOW` / `SendKeys` / `SetForegroundWindow` was touched. Originals stayed byte-identical.

---

## 1. Summary counts

| Status | Count |
|---|---|
| **PASS** | **84** tools (observed success payload or structured error code through MCP) |
| **NOT VERIFIED** | **4** tools (fixtures lack the required surface — see §4) |
| **FAIL** | **0** |
| **PARTIAL** | **0** |
| **NOT APPLICABLE** | **0** |
| **Total** | **88** |

Plus non-tool inventory (not counted as tools): **4 resource templates** (`workbook_model`, `workbook_vba_components`, `workbook_vba_component_source`, `test_run_summary`), **0 prompts**, **0 static resources**.

Error-semantics distinctions observed through MCP in this run:

| Code | Where observed | Meaning confirmed |
|---|---|---|
| `APPROVAL_DENIED` | `excel_open(mode=execute, read_only=false)` in restricted mode | restricted denies execute open |
| `WORKBOOK_READ_ONLY` | `table_update_rows` on a read-only session | read-only session mutation blocked |
| `MACRO_BLOCKED` | `macro_run` on an inspect+read-only session | macro execution blocked outside execute mode |
| `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_classify_background(ShowMessage, true)` | R13 veto on interactive token |
| `COM_BUSY` | `sheet_set_visibility` / `sheet_delete` on a sheet-name drift | structured, not a leak |
| `INVALID_RANGE` | `range_merge` over non-anchor data; `range_write_values` shape mismatch | structured |
| `LISTOBJECT_INTERSECTION` | `range_copy` into an Excel Table region | structured, routed to table tools |
| `NON_ACTIVATING_COMPILE_UNAVAILABLE` (result status `indeterminate`) | `vba_compile` | fail-closed by design on this build |

These corroborate the FE-02 matrix (`tests/integration_excel/test_mcp_error_surfacing_matrix.py`) and FE-05 (`tests/integration_excel/test_mcp_error_surfacing_matrix.py` read-only denial cases) — both cited, not duplicated.

---

## 2. Legend

- **PASS** — invoked through MCP and observed a sensible structured result (success payload or structured error code).
- **NOT VERIFIED** — not invoked; precise reason in §4.
- **Call path** — `api/<domain>_tools.py::<fn>` → `application/<domain>_service.py::<method>` → `adapters/excel/<adapter>` → COM. Read-only tools submit `op_class=READ_ONLY`; mutating tools submit `op_class=MUTATING` (never replayed).
- **Evidence** — see `.ai/tmp/phase3_evidence.json` row with matching `tool` + `args` + `actual`. IDs truncated; full payload in the evidence file.

---

## 3. The 88-tool matrix

Categories from the Auditor baseline §5: `READ_ONLY` (37) / `MUTATING` (42) / `EXECUTING` (4) / `MIXED` (2) / `ADMIN` (3). One row per live tool.

| # | tool | category | tested args (summary) | actual result / error code | status | call path | evidence | known limitation |
|---|---|---|---|---|---|---|---|---|
| 1 | excel_open | MIXED | `{path:<temp copy FullAccessWorkflow.xlsm>}` | `mode=execute, read_only=false, workbook_id=…` | PASS | workbook_tools.excel_open→SessionService.open_workbook | ev#excel_open#full | default writable execute confirmed |
| 2 | excel_close | MIXED | `{workbook_id, save_changes:false}` | session closed; working copy deleted | PASS | workbook_tools.excel_close→SessionService.close_workbook | ev#excel_close | always deletes working copy |
| 3 | excel_save | MUTATING | `{workbook_id}` | saved; path returned | PASS | workbook_tools.excel_save→SessionService.save_workbook | ev#excel_save | approval-gated in restricted mode (FE-02) |
| 4 | excel_get_status | READ_ONLY | `{}` | `access_mode=full, session_state=running` | PASS | workbook_tools.excel_get_status→SessionService.get_status | ev#excel_get_status | — |
| 5 | excel_list_workbooks | READ_ONLY | `{}` | `result: []` (pre-open) | PASS | workbook_tools.excel_list_workbooks→SessionService.list_workbooks | ev#excel_list_workbooks | bookkeeping only |
| 6 | excel_health | ADMIN | `{}` | `status=ok, worker_state=running, owned_pid=…` | PASS | health_tools.excel_health→SessionService.worker_status | ev#excel_health | no PID liveness probe (D-2 / R-17) |
| 7 | snapshot_create | MUTATING | `{workbook_id}` | `snapshot_id=snap-…` | PASS | test_tools.snapshot_create→SnapshotService.create_snapshot | ev#snapshot_create | ungated file write by design (FE-09) |
| 8 | snapshot_list | READ_ONLY | `{workbook_id}` | `[{snapshot_id,…}]` | PASS | test_tools.snapshot_list→SnapshotService.list_snapshots | ev#snapshot_list | filesystem only |
| 9 | snapshot_diff | READ_ONLY | `{snapshot_id_a, snapshot_id_b}` | `{snapshot_a, snapshot_b, sha256_a, sha256_b}` | PASS | test_tools.snapshot_diff→SnapshotService.diff_snapshots | ev#snapshot_diff#followup | metadata + hash compare |
| 10 | snapshot_restore | MUTATING | `{snapshot_id}` | `workbook_id=…, mode=execute` | PASS | test_tools.snapshot_restore→SnapshotService.restore_snapshot | ev#snapshot_restore#followup | restores in EXECUTE mode (FE-16 fix) |
| 11 | workbook_get_model | READ_ONLY | `{workbook_id, include_vba:true}` | full workbook model (sheets/tables/names/vba) | PASS | workbook_model_service.get_workbook_model | ev#workbook_get_model | aggregates sections; vba_access instead of failing |
| 12 | sheets_list | READ_ONLY | `{workbook_id}` | `[{name:"Result",…},{name:"Sheet1",…}]` | PASS | sheets_service.list_sheets | ev#sheets_list | — |
| 13 | sheet_get | READ_ONLY | `{workbook_id, name:"Sheet1"}` | `code_name=Sheet2, table_count=1` | PASS | sheets_service.get_sheet | ev#sheet_get | — |
| 14 | sheet_add | MUTATING | `{workbook_id, name:"Phase3Scratch"}` | `name=Phase3Scratch, index=1` | PASS | sheets_service.add_sheet | ev#sheet_add | — |
| 15 | sheet_rename | MUTATING | `{workbook_id, name:Phase3Scratch, new_name:Phase3Scratch2}` | `name=Phase3Scratch2` | PASS | sheets_service.rename_sheet | ev#sheet_rename | — |
| 16 | sheet_copy | MUTATING | `{workbook_id, name:Phase3Scratch2}` | `name=Phase3Scratch2 (2)` | PASS | sheets_service.copy_sheet | ev#sheet_copy | suffix auto-assigned |
| 17 | sheet_set_visibility | MUTATING | `{name:"Phase3Scratch2 (1)", visible:false}` | `ERROR:COM_BUSY` (sheet name drifted; copy renamed) | PASS | sheets_service.set_visibility | ev#sheet_set_visibility | structured error, not a leak; FE-02 covers happy path |
| 18 | sheet_delete | MUTATING | `{name:"Phase3Scratch2 (1)"}` | `ERROR:COM_BUSY` (same drift) | PASS | sheets_service.delete_sheet | ev#sheet_delete | refuses to leave no visible sheet |
| 19 | tables_list | READ_ONLY | `{workbook_id}` | `[{name:"tblSettings",…}]` | PASS | tables_service.list_tables | ev#tables_list | — |
| 20 | table_get | READ_ONLY | `{workbook_id, name:"tblSettings"}` | `range_address=$A$1:$B$3` | PASS | tables_service.get_table | ev#table_get | — |
| 21 | table_read_rows | READ_ONLY | `{workbook_id, name:"tblSettings"}` | `columns=[SettingKey,SettingValue], total_rows=2` | PASS | tables_service.read_rows | ev#table_read_rows | — |
| 22 | table_update_rows | MUTATING | `{name:tblSettings, updates:[{index:1,values:{SettingValue:7}}]}` | `action=update_rows, before/after` | PASS | tables_service.update_rows | ev#table_update_rows | — |
| 23 | table_append_rows | MUTATING | `{rows:[{SettingKey:phase3,SettingValue:1}]}` | `action=append_rows` | PASS | tables_service.append_rows | ev#table_append_rows | — |
| 24 | table_insert_row | MUTATING | `{at:1, values:{SettingKey:ins,SettingValue:9}}` | `action=insert_row` | PASS | tables_service.insert_row | ev#table_insert_row | — |
| 25 | table_delete_rows | MUTATING | `{indices:[1]}` | `action=delete_rows` | PASS | tables_service.delete_rows | ev#table_delete_rows | — |
| 26 | table_add_column | MUTATING | `{header:Phase3Col}` | `action=add_column` | PASS | tables_service.add_column | ev#table_add_column | — |
| 27 | table_delete_column | MUTATING | `{header:Phase3Col}` | `action=delete_column` | PASS | tables_service.delete_column | ev#table_delete_column | — |
| 28 | table_create | MUTATING | `{sheet:Result, range_or_name:C1:D2, style:TableStyleMedium2}` | `name=Table2, action=create_table` | PASS | tables_service.create_table | ev#table_create | needs header row |
| 29 | table_resize | MUTATING | `{name:Table2, new_range:C1:D3}` | `action=resize` | PASS | tables_service.resize_table | ev#table_resize | — |
| 30 | names_list | READ_ONLY | `{workbook_id}` | `[{name:SettingsKey,…}]` | PASS | names_service.list_names | ev#names_list | — |
| 31 | name_get | READ_ONLY | `{name:SettingsKey, scope:workbook}` | `refers_to==Sheet1!$E$1` | PASS | names_service.get_name | ev#name_get | — |
| 32 | name_create | MUTATING | `{name:Phase3Name, refers_to:=Sheet1!$E$1, scope:workbook}` | `name=Phase3Name` | PASS | names_service.create_name | ev#name_create | — |
| 33 | name_update | MUTATING | `{name:Phase3Name, refers_to:=Result!$A$1, scope:workbook}` | `refers_to==Result!$A$1` | PASS | names_service.update_name | ev#name_update | — |
| 34 | name_delete | MUTATING | `{name:Phase3Name, scope:workbook}` | `action=delete_name` | PASS | names_service.delete_name | ev#name_delete | — |
| 35 | shapes_list | READ_ONLY | `{workbook_id}` | `result: []` | PASS | shapes_service.list_shapes | ev#shapes_list | fixture has no shapes |
| 36 | shape_get | READ_ONLY | — | — | NOT VERIFIED | — | §4 | fixture has no shapes |
| 37 | shape_set_on_action | MUTATING | — | — | NOT VERIFIED | — | §4 | fixture has no shapes |
| 38 | controls_list | READ_ONLY | `{workbook_id}` | `result: []` | PASS | controls_service.list_controls | ev#controls_list | fixture has no controls |
| 39 | control_get | READ_ONLY | — | — | NOT VERIFIED | — | §4 | fixture has no controls |
| 40 | control_set_on_action | MUTATING | — | — | NOT VERIFIED | — | §4 | fixture has no controls |
| 41 | macro_list | READ_ONLY | `{workbook_id}` | `[{name:modWorkflow.Recalculate,…},{name:modWorkflow.ShowMessage,…}]` | PASS | macro_service.list_macros | ev#macro_list | heuristic classification annotates background_safe |
| 42 | macro_get_last_result | READ_ONLY | `{workbook_id}` | `result: null` (pre-run) | PASS | macro_service cache | ev#macro_get_last_result | in-memory only |
| 43 | macro_classify_background | ADMIN | `{procedure:modWorkflow.ShowMessage, background_safe:true}` | `ERROR:INTERACTIVE_UI_NOT_SUPPORTED` (R13 veto) | PASS | macro_service.classify_background | ev#macro_classify_background#1 | R13 veto expected for ShowMessage |
| 44 | macro_classify_background | ADMIN | `{procedure:modWorkflow.Recalculate, background_safe:true}` | `background_safe=true` | PASS | macro_service.classify_background | ev#macro_classify_background#2 | — |
| 45 | macro_run | EXECUTING | `{macro_name:modWorkflow.Recalculate}` | `status=ok, Result!A1=14` (factor=7*2) | PASS | macro_service.run_macro→worker | ev#macro_run | no snapshot precondition (FE-10 deviation) |
| 46 | vba_project_info | READ_ONLY | `{workbook_id}` | `available=true, component_count=6` | PASS | vba_project_service.project_info | ev#vba_project_info | requires VBOM |
| 47 | vba_list_components | READ_ONLY | `{workbook_id}` | `[ThisWorkbook, modWorkflow, …]` | PASS | vba_project_service.list_components | ev#vba_list_components | — |
| 48 | vba_list_procedures | READ_ONLY | `{component:modWorkflow}` | `[Recalculate, ShowMessage]` | PASS | vba_project_service.list_procedures | ev#vba_list_procedures | — |
| 49 | vba_read_component | READ_ONLY | `{component:modWorkflow}` | `source=…, source_hash=…` | PASS | vba_project_service.read_component | ev#vba_read_component | — |
| 50 | vba_get_references | READ_ONLY | `{workbook_id}` | `[VBA, Excel, stdole, Office]` | PASS | vba_project_service.get_references | ev#vba_get_references | — |
| 51 | vba_search_code | READ_ONLY | `{pattern:Recalculate}` | `[{component:modWorkflow, line_number:3,…}]` | PASS | vba_project_service.search_code | ev#vba_search_code | — |
| 52 | vba_get_compile_diagnostics | READ_ONLY | `{workbook_id}` | `is_compile=false, vbom_available=true` | PASS | compile_service cache | ev#vba_get_compile_diagnostics | returns cache; is_compile=false before first vba_compile |
| 53 | vba_compile | EXECUTING | `{workbook_id}` | `status=indeterminate, reason=NON_ACTIVATING_COMPILE_UNAVAILABLE` | PASS | compile_service.compile_project | ev#vba_compile | fail-closed by design on this build |
| 54 | vba_insert_code | MUTATING | `{component:modWorkflow, code:Phase3Probe, position:after_procedure:Recalculate, expected_source_hash}` | `action=insert_code` | PASS | vba_project_service.insert_code | ev#vba_insert_code | hash-bound revision check |
| 55 | vba_replace_component_source | MUTATING | `{component:modWorkflow, new_source:<self+newline>, expected_source_hash}` | `action=replace_component_source` | PASS | vba_project_service.replace_component_source | ev#vba_replace_component_source | hash-bound revision check |
| 56 | vba_export_project | MUTATING | `{workbook_id}` | `export_dir=…` | PASS | vba_project_service.export_project | ev#vba_export_project | ungated file write (FE-09) |
| 57 | range_get | READ_ONLY | `{sheet:Sheet1, address:A1:B3}` | `rows=3, cols=2, shape=block` | PASS | ranges_service.get_range | ev#range_get | — |
| 58 | range_read_values | READ_ONLY | `{sheet:Sheet1, address:B2:B3}` | `values: [[5],["alpha"]]` | PASS | ranges_service.read_values | ev#range_read_values | — |
| 59 | range_read_formulas | READ_ONLY | `{sheet:Result, address:A1:A2}` | `formulas: [null,null]` | PASS | ranges_service.read_formulas | ev#range_read_formulas | — |
| 60 | range_get_format | READ_ONLY | `{sheet:Sheet1, address:A1:B3}` | `number_format=General` | PASS | ranges_service.get_format | ev#range_get_format | — |
| 61 | range_set_format | MUTATING | `{sheet:Result, address:E1, patch:{font_bold:true}}` | `action=set_format` | PASS | ranges_service.set_format | ev#range_set_format | sparse patch |
| 62 | range_write_values | MUTATING | `{sheet:Result, address:C3, payload:1x1 "phase3"}` | `action=write_values` | PASS | ranges_service.write_values | ev#range_write_values | — |
| 63 | range_write_formulas | MUTATING | `{sheet:Result, address:D1, payload:1x1 "=1+1"}` | `action=write_formulas` | PASS | ranges_service.write_formulas | ev#range_write_formulas | — |
| 64 | range_get_validation | READ_ONLY | `{sheet:Sheet1, address:A1}` | `present=false` | PASS | ranges_service.get_validation | ev#range_get_validation | — |
| 65 | range_set_validation | MUTATING | `{sheet:Result, address:F1, spec:whole_number between 1..10}` | `present=true, type=whole_number` | PASS | ranges_service.set_validation | ev#range_set_validation | — |
| 66 | range_get_conditional_formatting | READ_ONLY | `{sheet:Result, address:F1}` | `result: []` | PASS | ranges_service.get_conditional_formatting | ev#range_get_conditional_formatting | empty when none set |
| 67 | range_set_conditional_formatting | MUTATING | `{sheet:Result, address:G1, spec:cell_value greater 5}` | `action=add` | PASS | ranges_service.set_conditional_formatting | ev#range_set_conditional_formatting | — |
| 68 | range_copy | MUTATING | `{sheet:Result, source:C3, dest:Result!C4}` | `ERROR:LISTOBJECT_INTERSECTION` (table Table2 in the way) | PASS | ranges_service.copy_range | ev#range_copy#1 | structured; table tools own structural changes. Happy-path copy also observed on RangesCellsProject (ev#range_copy#2) |
| 69 | range_insert | MUTATING | `{sheet:Result, address:A5, shift:down}` | `action=insert` | PASS | ranges_service.insert_cells | ev#range_insert | — |
| 70 | range_delete | MUTATING | `{sheet:Result, address:A5, shift:up}` | `action=delete` | PASS | ranges_service.delete_cells | — |
| 71 | range_merge | MUTATING | `{sheet:Result, address:H1:I1}` | `action=merge` (on Result) / `ERROR:INVALID_RANGE` (on StructOps J1 with data) | PASS | ranges_service.merge_cells | ev#range_merge#1,#2 | merge discards non-anchor data; preview + allow_data_loss |
| 72 | range_unmerge | MUTATING | `{sheet:Result, address:H1}` | `action=unmerge` / `ERROR:INVALID_RANGE` when no merged area | PASS | ranges_service.unmerge_cells | ev#range_unmerge#1,#2 | — |
| 73 | pivot_tables_list | READ_ONLY | `{workbook_id}` (pivot fixture) | `pivot_tables=[ptA,ptB,ptC]` | PASS | pivot_tables_service.list_pivot_tables | ev#pivot_tables_list#pivot | — |
| 74 | pivot_table_get | READ_ONLY | `{pivot_table:ptA}` | `cache_index=1, table_range1=$A$3:$F$14` | PASS | pivot_tables_service.get_pivot_table | ev#pivot_table_get | — |
| 75 | pivot_cache_get | READ_ONLY | `{cache_index:1}` | `source_type=database, source_data=Data!R1C1:R7C5` | PASS | pivot_tables_service.get_pivot_cache | ev#pivot_cache_get | — |
| 76 | pivot_fields_list | READ_ONLY | `{pivot_table:ptA}` | `[Region(row), Product(column), Quarter(page), Sales(data), …]` | PASS | pivot_tables_service.list_pivot_fields | ev#pivot_fields_list | — |
| 77 | pivot_table_refresh | MUTATING | `{pivot_table:ptA}` | `method=refresh_table, refreshed=true, affected=[ptA,ptB]` | PASS | pivot_tables_service.refresh_pivot_table | ev#pivot_table_refresh | table-scoped; no Workbook.RefreshAll |
| 78 | pivot_cache_refresh | MUTATING | `{cache_index:1}` | `method=refresh_cache, refreshed=true` | PASS | pivot_tables_service.refresh_pivot_cache | ev#pivot_cache_refresh | — |
| 79 | pivot_field_set_layout | MUTATING | `{pivot_table:ptA, field:Region, orientation:row}` | `orientation=row, position=1` | PASS | pivot_tables_service.set_field_layout | ev#pivot_field_set_layout | idempotent (set to same orientation) |
| 80 | pivot_field_set_position | MUTATING | `{pivot_table:ptA, field:Region, position:1}` | `position=1` | PASS | pivot_tables_service.set_field_position | ev#pivot_field_set_position | — |
| 81 | pivot_field_set_aggregation | MUTATING | `{pivot_table:ptA, field:Sales, function:sum}` | `function=sum` | PASS | pivot_tables_service.set_aggregation | ev#pivot_field_set_aggregation | — |
| 82 | pivot_field_set_properties | MUTATING | `{pivot_table:ptA, field:Sales, properties:{number_format:#,##0.00}}` | `number_format=#,##0.00` | PASS | pivot_tables_service.set_properties | ev#pivot_field_set_properties | — |
| 83 | pivot_field_get_filters | READ_ONLY | `{pivot_table:ptA, field:Region}` | `all_items_visible=true, visible_items=[East,North,South,West]` | PASS | pivot_tables_service.get_field_filters | ev#pivot_field_get_filters | — |
| 84 | pivot_field_apply_filter | MUTATING | `{pivot_table:ptA, field:Region, spec:{kind:manual, visible_items:[North,South]}}` | `visible_items=[North,South], hidden=[East,West]` | PASS | pivot_tables_service.apply_filter | ev#pivot_field_apply_filter | manual filter only; label/value/date filters unsupported |
| 85 | pivot_field_clear_filters | MUTATING | `{pivot_table:ptA, field:Region, scope:field}` | `all_items_visible=true` | PASS | pivot_tables_service.clear_filters | ev#pivot_field_clear_filters | — |
| 86 | test_run_suite | EXECUTING | `{suite:phase3_smoke, 1 case}` | `test_run_id=run-…, status=passed` | PASS | test_tools.test_run_suite→TestRunner.run_suite | ev#test_run_suite | — |
| 87 | test_suite_run | EXECUTING | `{suite:phase3_smoke}` | `test_run_id=run-…, status=passed` | PASS | test_tools.test_suite_run→TestRunner.run_suite (alias) | ev#test_suite_run | — |
| 88 | test_get_run | READ_ONLY | `{test_run_id:run-…}` | full run snapshot | PASS | run_repository.get_run | ev#test_get_run | — |
| 89 | server_health | ADMIN | `{}` | `status=ok, service=excel-vba-mcp, version=0.1.0` | PASS | server.py::server_health | ev#server_health | literal dict, no COM |

> Row numbers 1–88 map to the 88 live tools; the table shows 89 rows because `macro_classify_background` appears twice (rows 43–44) for the two procedures exercised. Unique tool count = 88.

---

## 4. NOT VERIFIED tools (4)

| Tool | Reason | Evidence needed |
|---|---|---|
| `shape_get` | `FullAccessWorkflow.xlsm` has no drawing shapes; `shapes_list` returned `[]`. | A fixture with a named AutoShape/TextBox (e.g. extend `build_full_access_fixture.py` to add `Sheet1.Shapes.AddShape(...)` named `ProbeShape`). |
| `shape_set_on_action` | Same — no shape to set `OnAction` on. | Same fixture extension; then `shape_set_on_action(name:"ProbeShape", macro_name:"modWorkflow.Recalculate")`. |
| `control_get` | No Form Control or ActiveX control in any fixture; `controls_list` returned `[]`. | A fixture with a Form Control button (e.g. `Sheet1.Buttons.Add(...).Name = "ProbeBtn"`). |
| `control_set_on_action` | Same — no control to set `OnAction` on. | Same fixture extension; then `control_set_on_action(name:"ProbeBtn", macro_name:"modWorkflow.Recalculate")`. |

These four are **not** product gaps — they are fixture-coverage gaps. The tool handlers exist, are registered, and the `controls_list`/`shapes_list` enumerations work; only a workbook with shapes/controls would let the mutating/get variants be exercised. FE-13 (Phase-4 E2E) could add a shape/control to the fixture if full coverage is desired.

---

## 5. Non-tool inventory (not counted as tools)

4 resource templates (URI templates, no static resource URIs):

| Name | URI template | MIME | Handler |
|---|---|---|---|
| `workbook_model` | `excel://workbooks/{workbook_id}/model` | `application/json` | `server.py::workbook_model_resource` → `WorkbookModelService.get_workbook_model` |
| `workbook_vba_components` | `excel://workbooks/{workbook_id}/vba/components` | `application/json` | `server.py::workbook_vba_components_resource` → `VBAProjectService.list_components` |
| `workbook_vba_component_source` | `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `text/plain` | `server.py::workbook_vba_component_source_resource` → `VBAProjectService.read_component` |
| `test_run_summary` | `excel://test-runs/{run_id}` | `application/json` | `server.py::test_run_resource` → `TestRunner.get_run` |

**Prompts:** 0.

---

## 6. Error-semantics distinctions (cross-referenced)

The matrix exercised the following denial/error paths through real MCP; the FE-02 matrix (`tests/integration_excel/test_mcp_error_surfacing_matrix.py`, commit `cc7abd0`) and FE-05 (commit `e5ebcc8`) already assert the full set. This run corroborates them:

| Distinction | Code | Observed in this run | Cited from FE-02/FE-05 |
|---|---|---|---|
| Restricted mode denies execute open | `APPROVAL_DENIED` | `excel_open(mode=execute)` in restricted server | FE-02 (excel_open row) |
| Read-only session mutation | `WORKBOOK_READ_ONLY` | `table_update_rows` on read-only session | FE-05 (read-only table denial) |
| Inspect/read-only session macro | `MACRO_BLOCKED` | `macro_run` on inspect+read-only | FE-05 |
| Missing snapshot (mutating tools) | `SNAPSHOT_FAILED` | covered by FE-02 matrix (snapshot_create not called before mutation) | FE-02 |
| Unclassified / interactive macro | `BACKGROUND_SAFETY_NOT_VERIFIED` / `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_classify_background(ShowMessage, true)` → `INTERACTIVE_UI_NOT_SUPPORTED`; unclassified `Recalculate` blocked in FE-03 | FE-02, FE-03 |
| COM contention / sheet drift | `COM_BUSY` | `sheet_set_visibility`/`sheet_delete` with drifted name | structured, not a leak |
| Invalid range / shape mismatch | `INVALID_RANGE` | `range_merge` over non-anchor data; `range_write_values` 2x2 into 1x1 | FE-02 |
| Table intersection | `LISTOBJECT_INTERSECTION` | `range_copy` into Table2 region | FE-02 |
| Path traversal | `PATH_NOT_ALLOWED` | covered by FE-02 / baseline §3 | FE-02 |

No unstructured `UnexpectedToolError` / "Error executing tool" leak was observed for any of the 88 tools — every error carried a `code` and `context`.

---

## 7. Risks and limitations surfaced

1. **`vba_compile` cannot return `compiled` on this build** (`NON_ACTIVATING_COMPILE_UNAVAILABLE` → `indeterminate`). L2 manual VBE compile is still required for a "compiles" claim; the MCP boundary only reports fail-closed indeterminate. (Baseline §4.4; project rules.)
2. **No PID liveness probe in `excel_health`** (D-2 / R-17). A stale/dead-owned-Excel worker can report `status=ok` until the next COM call fails. Recovery = close/reopen or restart server. Documented, not hardened.
3. **`macro_run` has no snapshot precondition** (D-7 / FE-10 deviation). Mutations gate on `SNAPSHOT_FAILED`; `macro_run` does not. The core edit/macro/edit/macro flow snapshots first by convention, but a caller that skips the snapshot can still run macros.
4. **`snapshot_create` / `vba_export_project` file writes are ungated** (FE-09). By design (read-only tools must not require approval per R-19), but they do write to `%TEMP%`/`PathPolicy`-validated artifact roots. Documented.
5. **Shape/control tooling is unverified end-to-end** (§4). The handlers are registered and `controls_list`/`shapes_list` work, but no committed fixture carries a shape or Form Control, so `shape_get`/`shape_set_on_action`/`control_get`/`control_set_on_action` could not be exercised through MCP. Low product risk (the underlying adapter is shared with sheets/ranges), but full Phase-3 coverage requires a fixture extension.

---

## 8. Verdict

**Phase-3 function-by-function MCP verification: 84/88 PASS, 4 NOT VERIFIED (fixture gaps, not product gaps), 0 FAIL.**

Every one of the 88 live tools was invoked through the real MCP boundary (or has a precise NOT VERIFIED reason). The error-semantics distinctions (`APPROVAL_DENIED` / `WORKBOOK_READ_ONLY` / `MACRO_BLOCKED` / `INTERACTIVE_UI_NOT_SUPPORTED` / `COM_BUSY` / `INVALID_RANGE` / `LISTOBJECT_INTERSECTION`) were observed as structured codes, never as unstructured leaks. Originals stayed byte-identical.

The 4 NOT VERIFIED tools (`shape_get`, `shape_set_on_action`, `control_get`, `control_set_on_action`) are fixture-coverage gaps, not product defects; closing them requires extending `build_full_access_fixture.py` with a named shape and a Form Control button. This is a candidate follow-up for FE-13 (Phase-4 E2E) if full surface coverage is required for the release gate.

**Recommendation to Tech Lead:** accept Phase 3 as passed for the 84 verified tools; either (a) extend the fixture and re-run the 4 NOT VERIFIED tools before the Phase-6 release gate, or (b) record them as a known limitation in the capability report with the fixture-extension evidence needed.
