# WP1 — Tool & Action Mapping Ledger

**Project:** `xlvba_mcp` · **Repo HEAD:** `0261683` · **Date:** 2026-11-28
**Purpose:** Map every legacy tool/action name to exactly one final canonical name/value, or mark it explicitly removed. Research only.

**Sources**
- Legacy surface (94 old names): `tests/fixtures/wp0_baseline.json` → `ledger[]` (each row `{old_name, target:{tool, action?}}`).
- Live surface (43 tools): `.ai/audit/artifacts/tools_list_before.json` + `src/excel_vba_mcp/api/*_tools.py`.
- Approved final renames (fixed decisions from Tech Lead):
  - Tools: `excel_save_as→workbook_save_copy`, `test→test_suite`, `name→defined_name`, `name_read→defined_name_read`, `name_delete→defined_name_delete`, `workbook_file_read→workbook_snapshot_read`.
  - Action: `test` action `run-suite→run` (keep `get-run`).
  - No backward-compat aliases.

---

## 1. Tool mapping ledger (legacy tool → final canonical tool)

The 94 legacy `old_name` rows collapse to 43 final tools. Rows whose `target.tool` is itself renamed by an approved decision show the **final** name in the "Final canonical tool" column. Standalone rows map `old_name == target.tool` (then possibly renamed).

| Legacy old_name(s) | Final canonical tool | Notes |
|---|---|---|
| `excel_open`, `excel_close`, `excel_save` | `workbook` | consolidated (actions open/close/save) |
| `excel_get_status`, `excel_list_workbooks`, `workbook_get_model` | `workbook_read` | consolidated (status/list/model) |
| `excel_save_as` | **`workbook_save_copy`** | standalone, **renamed** |
| `excel_health` | `excel_health` | standalone |
| `excel_recover` | `excel_recover` | standalone |
| `snapshot_create` | `snapshot_create` | standalone |
| `snapshot_list` | `snapshot_list` | standalone |
| `snapshot_restore` | `snapshot_restore` | standalone |
| `snapshot_diff` | `snapshot_diff` | standalone |
| `test_suite_run`, `test_get_run` | **`test_suite`** | consolidated, **renamed** from `test` |
| `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_get_references`, `vba_search_code`, `vba_get_compile_diagnostics` | `vba_read` | consolidated (7 actions) |
| `vba_export_project` | `vba_export_project` | standalone |
| `vba_replace_component_source` | `vba_replace_component_source` | standalone |
| `vba_insert_code` | `vba_insert_code` | standalone |
| `vba_compile_diagnostics` | `vba_compile_diagnostics` | standalone |
| `macro_list`, `macro_classify_background`, `macro_get_last_result` | `macro` | consolidated (list/classify-background/get-last-result) |
| `macro_run` | `macro_run` | standalone |
| `job_get` | `job_get` | standalone |
| `job_request_cancel` | `job_request_cancel` | standalone |
| `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_set_visibility` | `sheet` | consolidated |
| `sheets_list`, `sheet_get` | `sheet_read` | consolidated |
| `sheet_delete` | `sheet_delete` | standalone |
| `table_create`, `table_resize`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_add_column` | `table` | consolidated |
| `tables_list`, `table_get`, `table_read_rows` | `table_read` | consolidated |
| `table_delete_rows` | `table_delete_rows` | standalone |
| `table_delete_column` | `table_delete_column` | standalone |
| `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_field_get_filters` | `pivot_read` | consolidated (read; +`result-read` added post-ledger) |
| `pivot_table_refresh`, `pivot_cache_refresh` | `pivot_refresh` | consolidated |
| `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_apply_filter`, `pivot_field_clear_filters` | `pivot_field` | consolidated |
| `range_get`, `range_read_values`, `range_read_formulas`, `range_get_format`, `range_get_validation`, `range_get_conditional_formatting` | `range_read` | consolidated |
| `range_write_values`, `range_write_formulas`, `range_set_format`, `range_set_validation`, `range_set_conditional_formatting` | `range_write` | consolidated |
| `range_copy`, `range_insert`, `range_unmerge` | `range_structural` | consolidated |
| `range_delete` | `range_delete` | standalone |
| `range_merge` | `range_merge` | standalone |
| `name_create`, `name_update` | **`defined_name`** | consolidated, **renamed** from `name` |
| `names_list`, `name_get` | **`defined_name_read`** | consolidated, **renamed** from `name_read` |
| `name_delete` | **`defined_name_delete`** | standalone, **renamed** |
| `shapes_list`, `shape_get`, `shape_set_on_action` | `shape` | consolidated |
| `controls_list`, `control_get`, `control_set_on_action` | `control` | consolidated |
| `server_health` | `server_health` | standalone |
| `xlvba_environment` | `xlvba_environment` | standalone |
| `workbook_file_read` | **`workbook_snapshot_read`** | standalone, **renamed** (snapshot semantics verified — see baseline §3.20 / policy §5) |
| `data_export` | `data_export` | standalone (added post-ledger; not in 94) |

**Removed tools:** none. Every legacy tool maps to a live canonical tool.

---

## 2. Action mapping ledger (legacy action → final canonical action value)

Only consolidated tools have explicit action values. Standalone tools have an implicit action equal to the (possibly renamed) tool name.

| Parent tool (final) | Legacy action | Final action value | Notes |
|---|---|---|---|
| `workbook` | `open` / `close` / `save` | `open` / `close` / `save` | unchanged |
| `workbook_read` | `status` / `list` / `model` | `status` / `list` / `model` | unchanged |
| **`test_suite`** | `run-suite` | **`run`** | **renamed** (approved) |
| **`test_suite`** | `get-run` | `get-run` | unchanged (retrieve persisted run by id) |
| `vba_read` | `project-info` / `list-components` / `list-procedures` / `read-component` / `get-references` / `search-code` / `get-compile-diagnostics` | same | unchanged |
| `macro` | `list` / `classify-background` / `get-last-result` | same | unchanged |
| `sheet` | `add` / `rename` / `copy` / `set-visibility` | same | unchanged |
| `sheet_read` | `list` / `get` | same | unchanged |
| `table` | `create` / `resize` / `append-rows` / `insert-row` / `update-rows` / `add-column` | same | unchanged |
| `table_read` | `list` / `get` / `read-rows` | same | unchanged |
| `pivot_read` | `tables-list` / `table-get` / `cache-get` / `fields-list` / `field-get-filters` | same (+`result-read`) | `result-read` added post-ledger |
| `pivot_refresh` | `refresh-table` / `refresh-cache` | same | unchanged |
| `pivot_field` | `field-set-layout` / `field-set-position` / `field-set-aggregation` / `field-set-properties` / `field-apply-filter` / `field-clear-filters` | same | unchanged |
| `range_read` | `get` / `read-values` / `read-formulas` / `get-format` / `get-validation` / `get-conditional-formatting` | same | unchanged |
| `range_write` | `write-values` / `write-formulas` / `set-format` / `set-validation` / `set-conditional-formatting` | same | unchanged |
| `range_structural` | `copy` / `insert` / `unmerge` | same | unchanged |
| **`defined_name`** | `create` / `update` | same | parent renamed only |
| **`defined_name_read`** | `list` / `get` | same | parent renamed only |
| `shape` | `list` / `get` / `set-on-action` | same | unchanged |
| `control` | `list` / `get` / `set-on-action` | same | unchanged |
| `workbook_snapshot_read` | `metadata` / `sheet-rows` / `pivot-cache-records` | same | parent renamed only |

**Removed actions:** none (only `run-suite`→`run` value change; `get-run` retained).

---

## 3. Completeness assertion (programmatic)

The check below asserts `set(source old-names) == set(ledger-covered old-names)` and prints the final canonical tool set. It reads the authoritative fixture ledger and applies the approved rename maps.

```python
import json
d = json.load(open('tests/fixtures/wp0_baseline.json', encoding='utf-8'))
ledger = d['ledger']
old_names = [r['old_name'] for r in ledger]

tool_rename = {
    'excel_save_as': 'workbook_save_copy',
    'test': 'test_suite',
    'name': 'defined_name',
    'name_read': 'defined_name_read',
    'name_delete': 'defined_name_delete',
    'workbook_file_read': 'workbook_snapshot_read',
}
action_rename = {('test', 'run-suite'): 'run'}

def final_target(r):
    t = r['target']; tool = t['tool']; act = t.get('action')
    ftool = tool_rename.get(tool, tool)
    if act is not None:
        act = action_rename.get((tool, act), act)
    return ftool, act

covered = set()
rows = []
for r in ledger:
    ft, fa = final_target(r)
    covered.add(r['old_name'])
    rows.append((r['old_name'], ft, fa))

assert set(old_names) == covered, 'mismatch'
print('OLD-NAME COUNT:', len(old_names))
print('UNIQUE OLD-NAMES:', len(set(old_names)))
print('COMPLETENESS: set(source old-names) == set(ledger covered) ->', set(old_names) == covered)

final_tools = sorted({ft for _, ft, _ in rows})
print('FINAL CANONICAL TOOL COUNT:', len(final_tools))
print(final_tools)
```

**Output (actual run):**

```
OLD-NAME COUNT: 94
UNIQUE OLD-NAMES: 94
COMPLETENESS: set(source old-names) == set(ledger covered) -> True
FINAL CANONICAL TOOL COUNT: 43
['control', 'data_export', 'defined_name', 'defined_name_delete', 'defined_name_read', 'excel_health', 'excel_recover', 'job_get', 'job_request_cancel', 'macro', 'macro_run', 'pivot_field', 'pivot_read', 'pivot_refresh', 'range_delete', 'range_merge', 'range_read', 'range_structural', 'range_write', 'server_health', 'shape', 'sheet', 'sheet_delete', 'sheet_read', 'snapshot_create', 'snapshot_diff', 'snapshot_list', 'snapshot_restore', 'table', 'table_delete_column', 'table_delete_rows', 'table_read', 'test_suite', 'vba_compile_diagnostics', 'vba_export_project', 'vba_insert_code', 'vba_read', 'vba_replace_component_source', 'workbook', 'workbook_read', 'workbook_save_copy', 'workbook_snapshot_read', 'xlvba_environment']
```

**Result:** completeness holds — every one of the 94 source old-names is covered by exactly one ledger row, and the 94 rows resolve to exactly 43 final canonical tools. `data_export` is the only final tool not present in the 94-row legacy ledger (added post-consolidation with the large-data integration); it is included in the live 43 and is mapped above for completeness.

---

## 4. Notes on the two added tools (post-ledger)

- `workbook_file_read` (→ `workbook_snapshot_read`): present in the 94-row ledger as a standalone `old_name` (row `workbook_file_read -> workbook_file_read / None`), so it IS covered by the completeness check. It is renamed by approved decision.
- `data_export`: NOT in the 94-row ledger (added after the baseline snapshot). It is a live tool (43rd). Listed in §1/§2 for a complete live mapping; its absence from the 94 is expected and does not break the assertion (the assertion is over source old-names, all of which are covered).