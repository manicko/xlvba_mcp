# WP0 — Tool & Action Naming Baseline

**Project:** `xlvba_mcp` (Excel VBA MCP server)
**Repo HEAD:** `0261683` on `master`
**Date:** 2026-11-28
**Scope:** Research only. No production code, tests, or docs were modified.
**Ground truth:** `.ai/audit/artifacts/tools_list_before.json` (serialized `tools/list`) + source in `src/excel_vba_mcp/api/*_tools.py` + `src/excel_vba_mcp/server.py`.

---

## 1. Authoritative tool count

**The actual runtime `tools/list` exposes 43 tools.**

Reproduced with the authoritative command (does not need Excel):

```
& .\.venv\Scripts\python.exe -c "import asyncio,json,excel_vba_mcp.server as s; t=asyncio.run(s.mcp.list_tools()); data=[x.model_dump(by_alias=True) for x in t]; print(len(t)); [print(x.name) for x in t]"
```

Output (count + names):

```
43
server_health
workbook
workbook_read
excel_save_as
excel_health
excel_recover
snapshot_create
snapshot_list
snapshot_restore
snapshot_diff
test
vba_read
vba_export_project
vba_replace_component_source
vba_insert_code
vba_compile_diagnostics
macro
macro_run
job_get
job_request_cancel
sheet
sheet_read
sheet_delete
table
table_read
table_delete_rows
table_delete_column
pivot_read
pivot_refresh
pivot_field
range_read
range_write
range_structural
range_delete
range_merge
name
name_read
name_delete
shape
control
xlvba_environment
workbook_file_read
data_export
```

The same 43 names are confirmed by `tests/unit/test_description_contract.py::test_tool_count_is_final_exposed` (`assert len(tools) == 43`).

### 1.1 Stale-fixture discrepancy (94 vs 43) — CONFIRMED

`tests/fixtures/wp0_baseline.json` contains a `baseline` block that is **stale**:

```json
"baseline": {
  "tool_count": 94,
  "serialized_bytes": 352155,
  "measured_at": "2026-11-27T19:13:53.775526+00:00",
  ...
}
```

- The `tool_count: 94` is **NOT** the runtime tool count. The runtime count is **43**.
- The `94` is the number of **ledger `old_name` rows** (the pre-consolidation surface: 94 legacy verb-level names that were consolidated down to 43 tools). The fixture author recorded the ledger row count in the `tool_count` field, which is misleading.
- The fixture's own `ledger` array has exactly **94 rows**, each mapping one `old_name` to a `target` (`tool` + optional `action`). The 94 old names resolve to exactly **43** distinct final tools (see `tool_action_mapping.md`).
- **Conclusion:** treat `tools_list_before.json` (43 tools) and the live `tools/list` as ground truth. The fixture `baseline.tool_count` value of 94 is stale/mislabeled and must not be cited as the tool count.

---

## 2. The 43 tools — names, annotations, action enums, required params

Annotations are the truthful MCP hints emitted by `server._annotations()` (frozensets `_ANN_READ_ONLY` / `_ANN_DESTRUCTIVE` / `_ANN_IDEMPOTENT` at `server.py:1942-2002`), plus the two inline-annotated tools (`server_health`, `xlvba_environment`). `R`=readOnlyHint, `D`=destructiveHint, `I`=idempotentHint.

| # | Tool | R | D | I | action enum | required params |
|---|------|---|---|---|-------------|-----------------|
| 1 | `server_health` | T | F | T | — (standalone) | — |
| 2 | `workbook` | F | T | F | `open`, `close`, `save` | `action` |
| 3 | `workbook_read` | T | F | T | `status`, `list`, `model` | `action` |
| 4 | `excel_save_as` | F | T | F | — (standalone) | `workbook_id`, `path` |
| 5 | `excel_health` | T | F | T | — (standalone) | — |
| 6 | `excel_recover` | F | T | F | — (standalone) | — (`confirm` gated) |
| 7 | `snapshot_create` | F | F | F | — (standalone) | `workbook_id` |
| 8 | `snapshot_list` | T | F | T | — (standalone) | — |
| 9 | `snapshot_restore` | F | T | F | — (standalone) | `snapshot_id` (+`workbook_id` enforced, +`confirm`) |
| 10 | `snapshot_diff` | T | F | T | — (standalone) | `snapshot_id_a`, `snapshot_id_b` |
| 11 | `test` | F | T | F | `run-suite`, `get-run` | `action` |
| 12 | `vba_read` | T | F | T | `project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics` | `action` |
| 13 | `vba_export_project` | F | F | F | — (standalone) | `workbook_id` |
| 14 | `vba_replace_component_source` | F | T | F | — (standalone) | `workbook_id`, `component`, `new_source`, `expected_source_hash` |
| 15 | `vba_insert_code` | F | T | F | — (standalone) | `workbook_id`, `component`, `code`, `position`, `expected_source_hash` |
| 16 | `vba_compile_diagnostics` | F | F | F | — (standalone) | `workbook_id` |
| 17 | `macro` | F | F | T | `list`, `classify-background`, `get-last-result` | `action` |
| 18 | `macro_run` | F | T | F | — (standalone) | `workbook_id`, `macro_name` |
| 19 | `job_get` | T | F | T | — (standalone) | `job_id` |
| 20 | `job_request_cancel` | F | F | T | — (standalone) | `job_id` |
| 21 | `sheet` | F | F | F | `add`, `rename`, `copy`, `set-visibility` | `action` |
| 22 | `sheet_read` | T | F | T | `list`, `get` | `action` |
| 23 | `sheet_delete` | F | T | F | — (standalone) | `workbook_id`, `name` |
| 24 | `table` | F | F | F | `create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column` | `action` |
| 25 | `table_read` | T | F | T | `list`, `get`, `read-rows` | `action` |
| 26 | `table_delete_rows` | F | T | F | — (standalone) | `workbook_id`, `name` |
| 27 | `table_delete_column` | F | T | F | — (standalone) | `workbook_id`, `name`, `header` |
| 28 | `pivot_read` | T | F | T | `tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters`, `result-read` | `action` |
| 29 | `pivot_refresh` | F | F | T | `refresh-table`, `refresh-cache` | `action` |
| 30 | `pivot_field` | F | T | T | `field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters` | `action` |
| 31 | `range_read` | T | F | T | `get`, `read-values`, `read-formulas`, `get-format`, `get-validation`, `get-conditional-formatting` | `action` |
| 32 | `range_write` | F | F | F | `write-values`, `write-formulas`, `set-format`, `set-validation`, `set-conditional-formatting` | `action` |
| 33 | `range_structural` | F | T | F | `copy`, `insert`, `unmerge` | `action` |
| 34 | `range_delete` | F | T | F | — (standalone) | `workbook_id`, `sheet`, `address`, `shift` |
| 35 | `range_merge` | F | T | T | — (standalone) | `workbook_id`, `sheet`, `address` |
| 36 | `name` | F | F | T | `create`, `update` | `action` |
| 37 | `name_read` | T | F | T | `list`, `get` | `action` |
| 38 | `name_delete` | F | T | F | — (standalone) | `workbook_id`, `name`, `scope` |
| 39 | `shape` | F | F | T | `list`, `get`, `set-on-action` | `action` |
| 40 | `control` | F | F | T | `list`, `get`, `set-on-action` | `action` |
| 41 | `xlvba_environment` | T | F | T | — (standalone) | — |
| 42 | `workbook_file_read` | T | F | T | `metadata`, `sheet-rows`, `pivot-cache-records` | `action`, `workbook_id` |
| 43 | `data_export` | F | F | T | — (standalone) | `workbook_id`, `source_kind` |

**Annotation distribution (live, from `tools_list_before.json`):** readOnly=**14**, destructive=**16**, idempotent=**23**. (The docs page `05-mcp-tools-reference.md` says "41 tools" with 13/16/21 — stale: it predates the two large-data tools `workbook_file_read` + `data_export`, which add one readOnly and two idempotent. Live truth is 14/16/23.)

### 2.1 Annotation membership (from `server.py` frozensets)

- **readOnly (14):** `workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `sheet_read`, `table_read`, `pivot_read`, `range_read`, `name_read`, `workbook_file_read` (12 frozenset members) + inline `server_health`, `xlvba_environment` = **14**.

  > Note: `_ANN_READ_ONLY` has 12 members; `server_health` + `xlvba_environment` carry inline readOnly. The docs table lists 13 and omits `workbook_file_read` (added later). Live truth: **14 tools report readOnly=true** (12 frozenset + 2 inline).
- **destructive (16):** `workbook`, `excel_save_as`, `sheet_delete`, `range_merge`, `range_structural`, `excel_recover`, `snapshot_restore`, `table_delete_rows`, `table_delete_column`, `name_delete`, `range_delete`, `vba_replace_component_source`, `vba_insert_code`, `macro_run`, `test`, `pivot_field`.
- **idempotent (23):** the 21 `_ANN_IDEMPOTENT` members (`workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `job_request_cancel`, `sheet_read`, `table_read`, `pivot_read`, `pivot_refresh`, `pivot_field`, `range_read`, `range_merge`, `name`, `name_read`, `shape`, `control`, `macro`, `workbook_file_read`, `data_export`) + inline `server_health`, `xlvba_environment` = **23**.

---

## 3. Consolidated tools — full action-enum reference

Every consolidated tool routes through a flat-signature dispatcher (`option (a)`): it validates flat args and forwards to the existing per-verb handler unchanged; the target handler enforces per-action required params. An unknown `action` raises a structured `ExcelVbaMcpError` (`INVALID_RANGE`).

### 3.1 `workbook` — `open`, `close`, `save`
- `open`: path required; `mode` (execute/inspect), `read_only`, `background`. Creates hidden working copy. Mutating (copies file into `work/`). Non-destructive.
- `close`: `workbook_id` required; `save_changes`, `discard_changes`. Destructive path via `discard_changes=true`; refuse-by-default when dirty.
- `save`: `workbook_id` required. Writes working copy in place. Non-destructive.

### 3.2 `workbook_read` — `status`, `list`, `model`
- `status`: optional `workbook_id`. Read-only, no COM.
- `list`: no extra params. Read-only.
- `model`: `workbook_id` required; `include_vba`. Read-only (never executes VBA).

### 3.3 `test` — `run-suite`, `get-run`
- `run-suite`: `suite` required; `background`. **Executing** (runs declarative suite on disposable copies; may dispatch `macro_run` steps). Destructive hint. Not read-only.
- `get-run`: `run_id` required. Read of a persisted run. Read-only.

### 3.4 `vba_read` — `project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics`
- `project-info`: `workbook_id`. Read-only.
- `list-components`: `workbook_id`. Read-only.
- `list-procedures`: `workbook_id`, `component`. Read-only.
- `read-component`: `workbook_id`, `component`. Read-only.
- `get-references`: `workbook_id`. Read-only.
- `search-code`: `workbook_id`, `pattern`, `regex`. Read-only.
- `get-compile-diagnostics`: `workbook_id`. Read-only (cached diagnostics; does not compile).

### 3.5 `macro` — `list`, `classify-background`, `get-last-result`
- `list`: `workbook_id`. Read-only.
- `classify-background`: `workbook_id`, `procedure`, `background_safe`, optional `expected_project_source_hash`. **Ungated bookkeeping** (records classification; never grants run permission). Idempotent.
- `get-last-result`: `workbook_id`. Read-only (cached).

### 3.6 `sheet` — `add`, `rename`, `copy`, `set-visibility`
- `add`: `workbook_id`, `name`; `sheet_type` (worksheet only). Guarded mutation.
- `rename`: `workbook_id`, `name`, `new_name`. Guarded mutation.
- `copy`: `workbook_id`, `name`; optional `target_name`. Guarded mutation.
- `set-visibility`: `workbook_id`, `name`, `visible`. Guarded mutation.

### 3.7 `sheet_read` — `list`, `get`
- `list`: `workbook_id`. Read-only.
- `get`: `workbook_id`, `name`. Read-only.

### 3.8 `table` — `create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column`
- `create`: `workbook_id`, `sheet`, `range_or_name`; optional `style`. Guarded mutation.
- `resize`: `workbook_id`, `name`, `new_range`. Guarded mutation.
- `append-rows`: `workbook_id`, `name`, `rows`. Guarded mutation.
- `insert-row`: `workbook_id`, `name`, `at`; optional `values`. Guarded mutation.
- `update-rows`: `workbook_id`, `name`, `updates`. Guarded mutation.
- `add-column`: `workbook_id`, `name`, `header`; optional `position`. Guarded mutation.
- (Note: `delete-column` is intentionally NOT an action — standalone `table_delete_column`.)

### 3.9 `table_read` — `list`, `get`, `read-rows`
- `list`: `workbook_id`. Read-only.
- `get`: `workbook_id`, `name`. Read-only.
- `read-rows`: `workbook_id`, `name`; `offset`, `limit`, `cursor_version`. Read-only.

### 3.10 `pivot_read` — `tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters`, `result-read`
- `tables-list`: `workbook_id`; `background`. Read-only.
- `table-get`: `workbook_id`, `pivot_table`. Read-only.
- `cache-get`: `workbook_id`, `cache_index`. Read-only.
- `fields-list`: `workbook_id`, `pivot_table`. Read-only.
- `field-get-filters`: `workbook_id`, `pivot_table`, `field`. Read-only.
- `result-read`: `workbook_id`, `pivot_table`; `offset`, `limit`, `cursor_version`. Read-only.

### 3.11 `pivot_refresh` — `refresh-table`, `refresh-cache`
- `refresh-table`: `workbook_id`, `pivot_table`. Guarded mutation (not idempotent — re-refresh may read changed source).
- `refresh-cache`: `workbook_id`, `cache_index`. Guarded mutation.

### 3.12 `pivot_field` — `field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters`
- `field-set-layout`: `workbook_id`, `pivot_table`, `field`, `orientation`; `delete`. Guarded; destructive when `delete=true`.
- `field-set-position`: `workbook_id`, `pivot_table`, `field`, `position`. Guarded.
- `field-set-aggregation`: `workbook_id`, `pivot_table`, `field`, `function`; optional `caption`. Guarded.
- `field-set-properties`: `workbook_id`, `pivot_table`, `field`, `properties`. Guarded.
- `field-apply-filter`: `workbook_id`, `pivot_table`, `field`, `spec`. Guarded.
- `field-clear-filters`: `workbook_id`, `pivot_table`; optional `field`; `scope` (field/table). Guarded.

### 3.13 `range_read` — `get`, `read-values`, `read-formulas`, `get-format`, `get-validation`, `get-conditional-formatting`
- `get`: `workbook_id`, `sheet`, `address`. Read-only.
- `read-values`: + `offset`, `limit`, `cursor_version`. Read-only.
- `read-formulas`: + `offset`, `limit`, `cursor_version`. Read-only.
- `get-format`: `workbook_id`, `sheet`, `address`. Read-only.
- `get-validation`: `workbook_id`, `sheet`, `address`. Read-only.
- `get-conditional-formatting`: `workbook_id`, `sheet`, `address`. Read-only.

### 3.14 `range_write` — `write-values`, `write-formulas`, `set-format`, `set-validation`, `set-conditional-formatting`
- `write-values`: `workbook_id`, `sheet`, `address`, `payload`. Guarded mutation.
- `write-formulas`: `workbook_id`, `sheet`, `address`, `payload`. Guarded mutation.
- `set-format`: `workbook_id`, `sheet`, `address`, `patch`. Guarded mutation.
- `set-validation`: `workbook_id`, `sheet`, `address`, `spec`. Guarded mutation.
- `set-conditional-formatting`: `workbook_id`, `sheet`, `address`, `spec`; `mode` (add/modify/delete). Guarded; not idempotent (`mode=add` appends + renumbers priorities).

### 3.15 `range_structural` — `copy`, `insert`, `unmerge`
- `copy`: `workbook_id`, `sheet`, `source`, `destination_sheet`, `destination`. Guarded; destructive hint; not idempotent.
- `insert`: `workbook_id`, `sheet`, `address`, `shift` (right/down); `allow_merged`. Guarded; not idempotent.
- `unmerge`: `workbook_id`, `sheet`, `address`. Guarded; not idempotent.

### 3.16 `name` — `create`, `update`
- `create`: `workbook_id`, `name`, `refers_to`, `scope`; optional `sheet`. Guarded mutation.
- `update`: `workbook_id`, `name`, `refers_to`, `scope`; optional `sheet`. Guarded mutation.
- (`delete` intentionally NOT an action — standalone `name_delete`.)

### 3.17 `name_read` — `list`, `get`
- `list`: `workbook_id`. Read-only.
- `get`: `workbook_id`, `name`, `scope`; optional `sheet`. Read-only.

### 3.18 `shape` — `list`, `get`, `set-on-action`
- `list`: `workbook_id`. Read-only.
- `get`: `workbook_id`, `name`. Read-only.
- `set-on-action`: `workbook_id`, `name`, `macro_name` (empty clears). Guarded mutation (writes `Shape.OnAction`; macro never executed).

### 3.19 `control` — `list`, `get`, `set-on-action`
- `list`: `workbook_id`. Read-only.
- `get`: `workbook_id`, `name`. Read-only.
- `set-on-action`: `workbook_id`, `name`, `macro_name` (empty clears). Guarded mutation (Form Control only; ActiveX rejected).

### 3.20 `workbook_file_read` — `metadata`, `sheet-rows`, `pivot-cache-records`
- `metadata`: `workbook_id`; optional `snapshot_id`. Read-only (non-COM snapshot read).
- `sheet-rows`: `workbook_id`, `sheet`; `offset`, `limit`, `snapshot_id`. Read-only.
- `pivot-cache-records`: `workbook_id`, `cache_index`; `offset`, `limit`, `snapshot_id`. Read-only.

---

## 4. Standalone tools (single implicit action = tool name)

`server_health`, `excel_save_as`, `excel_health`, `excel_recover`, `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`, `vba_export_project`, `vba_replace_component_source`, `vba_insert_code`, `vba_compile_diagnostics`, `macro_run`, `job_get`, `job_request_cancel`, `sheet_delete`, `table_delete_rows`, `table_delete_column`, `range_delete`, `range_merge`, `name_delete`, `xlvba_environment`, `data_export`.

---

## 5. Kilo permission pattern

The Kilo approval/permission pattern is a **server-name glob**: `xlvba_*` (the client exposes every published tool with the `xlvba_` prefix, e.g. `xlvba_workbook`). The documented client config sets:

```json
"permission": { "xlvba_*": "ask" }
```

Source: `docs/02-installation-and-setup.md` (lines 108-137) and `docs/05-mcp-tools-reference.md` (line 81: "Kilo approval remains the `xlvba_*` name-glob"). Individual tools can be allowlisted later (e.g. `"xlvba_server_health": "allow"`).

**Where the live Kilo config lives on this machine:** see `tool_action_naming_policy.md` §6 — the global config path `C:\Users\MatveichevN\.config\kilo\kilo.json` exists per the environment, but this research session is blocked by a global `external_directory` deny rule from reading outside the workspace, so the live glob value could not be read directly. The `xlvba_*` glob is documented in-repo and is the authoritative pattern for this project.

---

## 6. Examples & descriptions (source of truth)

Full tool descriptions and the single example strings live in `server.py` (`_DESC_*` constants) and the per-tool docstrings in `src/excel_vba_mcp/api/*_tools.py`. The serialized descriptions + input/output schemas for all 43 tools are in `.ai/audit/artifacts/tools_list_before.json`. Example-bearing tools (per `test_description_contract.py`): `range_write`, `table`, `vba_replace_component_source`, `test`.