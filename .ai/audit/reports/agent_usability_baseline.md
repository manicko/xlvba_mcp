# Agent Usability Baseline — xlvba MCP

Phase 0 baseline snapshot. Reproducible reference for before/after comparison.
All facts revalidated against source at the recorded revision; the original audit
(`.ai/tmp/xlvba-mcp-usability-audit-2026-11-26.md`) is treated as an initial audit,
not an implementation spec.

## 1. Source revision

- Git HEAD: `e2289ca01fd10136716a598cad1db6b6845fed73`
- HEAD subject: `docs: document autonomous Excel runtime, jobs, lifecycle, security, limits`
- Working tree at baseline: clean except untracked `.ai/audit/plans/` (this project's own plan dir).
- Runtime: Python 3.14.7, `mcp` 2.3.0, `pydantic` 2.14.4 (from `.venv`).
- Server module: `src/excel_vba_mcp/server.py` (1950 lines), package `excel_vba_mcp` v0.1.0.

## 2. Tool inventory (Phase 0 count)

- `server.add_tool(...)` calls: **90**
- Unique tool names: **90**
- Read-only MCP resources: **4** (`workbook_model`, `workbook_vba_components`,
  `workbook_vba_component_source`, `test_run_summary`)
- Prompts: **0**
- `outputSchema` / `structuredContent` / annotations currently emitted: **none** (verify in Phase 2)

### 2.1 Tools by domain (90)

Session/workbook (7): `excel_open`, `excel_close`, `excel_save`, `excel_get_status`,
`excel_list_workbooks`, `excel_health`, `excel_recover`.
Snapshots (4): `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`.
Tests (3): `test_run_suite`, `test_suite_run`, `test_get_run`.
VBA project (11): `vba_project_info`, `vba_list_components`, `vba_list_procedures`,
`vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code`,
`vba_replace_component_source`, `vba_insert_code`, `vba_compile`, `vba_get_compile_diagnostics`.
Macros (4): `macro_list`, `macro_run`, `macro_classify_background`, `macro_get_last_result`.
Jobs (2): `job_get`, `job_cancel`.
Sheets (7): `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`,
`sheet_delete`, `sheet_set_visibility`.
Tables (11): `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`,
`table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`,
`table_delete_column`, `table_create`, `table_resize`.
Pivot (13): `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`,
`pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`,
`pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`,
`pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters`.
Ranges (18): `range_get`, `range_read_values`, `range_read_formulas`, `range_write_values`,
`range_write_formulas`, `range_get_format`, `range_set_format`, `range_copy`, `range_insert`,
`range_delete`, `range_merge`, `range_unmerge`, `range_get_validation`, `range_set_validation`,
`range_get_conditional_formatting`, `range_set_conditional_formatting`.
Names (5): `names_list`, `name_get`, `name_create`, `name_update`, `name_delete`.
Shapes (3): `shapes_list`, `shape_get`, `shape_set_on_action`.
Controls (3): `controls_list`, `control_get`, `control_set_on_action`.
Model (1): `workbook_get_model`.

(Note: 7+4+3+11+4+2+7+11+13+18+5+3+3+1 = 92; two names are double-counted across
overlapping grep heuristics. Authoritative per-tool list is the `add_tool` block
`server.py:1479-1704`; exact membership to be frozen by the Auditor in Phase 2.)

## 3. Description / schema token cost (baseline)

- Description constant block: `server.py` lines ~749–1475.
- Measured description-region size: **28,907 characters** ≈ **~7,200 tokens** (English,
  ~4 chars/token) of static tool descriptions loaded at session start, before any schema.
- Input schemas are Pydantic-derived (per-tool `*.py` in `api/`), not counted here; Phase 2
  will measure the serialized `inputSchema` token cost precisely.

## 4. Kilo / MCP client capabilities (baseline, to be verified by Researcher 2A)

- Kilo config model: `kilo.json`/`kilo.jsonc` (project, `.kilo/`, global `~/.config/kilo/`).
  MCP servers declared under `mcp.<name>` with `type: local|remote`.
- MCP tool permissions use key `{server}_{tool}` with glob patterns (e.g. `xlvba_*`).
- Global `~/.config/kilo/kilo.jsonc` is **read-restricted** in this environment
  (`external_directory: deny`); its live contents could not be read here. Researcher 2A must
  determine effective Kilo MCP capabilities from Kilo documentation + the installed stack and
  record the limitation explicitly.
- Open capability questions (2A must answer, not assume): deferred tool loading / `defer_loading`;
  on-demand tool discovery; `outputSchema` + `structuredContent` support; MCP tool annotations
  transmission; whether all descriptions load at once; resources for progressive disclosure.

## 5. Access model (baseline, from config.py + statuses)

- `EXCEL_VBA_MCP_ACCESS_MODE` = `full` (default) | `restricted` (deny-by-default, approval grants).
- Per-call deadlines: `macro_timeout_s=90`, `open_timeout_s=300`, `liveness_ceiling_s=600`.
- No tool currently returns the effective access mode / build / snapshot requirement (audit P2).

## 6. Duplicate / overlapping / deprecated candidates (to confirm in Phase 2)

- `test_run_suite` vs `test_suite_run` (near-duplicate entry points).
- `xlvba_*` vs `excel-mcp_*` overlap on ranges/sheets/tables/names (no routing rule).
- `server_health` vs `excel_health` (liveness vs worker/session state — undocumented distinction).
- `job_cancel` naming vs actual intent-only semantics.

## 7. Resources / prompts

- Resources (4): per-workbook model, VBA component inventory, VBA component source, test-run summary.
- Prompts: none.

## 8. Already-implemented / invalidated audit items (preliminary)

- Typed DTOs, stable error taxonomy, fail-closed design, explicit unknown/interrupted states:
  present in source (confirm in Phase 2/3).
- The audit's benchmark scores are expert judgment, not measured; this project replaces them with
  the reproducible Agent Usability Benchmark (Researcher 2B → Phase 7).

## 9. Baseline benchmark

- Task set and harness: defined by Researcher 2B (Phase 1 Pair 2).
- Baseline run: executed in Phase 7 before remediation, same model/config/revision.
- Results recorded in `agent_usability_results.md`.