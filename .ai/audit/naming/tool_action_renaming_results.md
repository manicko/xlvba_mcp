# Tool & Action Renaming — Final Results

**Project:** `xlvba_mcp` · **Date:** 2026-11-28 · **Owner:** Tech Lead
**Scope:** Complete MCP tool + action naming audit, approved renames, tests, docs, runtime + independent validation.
**Status:** COMPLETE — accepted with documented non-blocking limitations.

---

## 1. Tool counts

| | Count |
|---|---|
| Tools before | **43** |
| Tools after | **43** (renames are 1:1; no tools added/removed) |
| Consolidated tools with `action` enums | 20 |
| Standalone tools | 23 |

> The pre-existing `tests/fixtures/wp0_baseline.json` carried a **stale** `tool_count: 94` (that number was the legacy *ledger row* count, not tools). Corrected to **43**; the 94 legacy names resolve to exactly 43 canonical tools (programmatic completeness check passes).

## 2. Complete old → new mapping

### Tools (6)
| Old | New canonical | Kind |
|---|---|---|
| `excel_save_as` | `workbook_save_copy` | standalone |
| `test` | `test_suite` | consolidated |
| `name` | `defined_name` | consolidated (create/update) |
| `name_read` | `defined_name_read` | consolidated (list/get) |
| `name_delete` | `defined_name_delete` | standalone |
| `workbook_file_read` | `workbook_snapshot_read` | standalone (snapshot semantics verified) |

### Action (1)
| Parent | Old action | New action |
|---|---|---|
| `test` → `test_suite` | `run-suite` | `run` (keep `get-run`) |

No aliases, no dual-run, no automatic translation. Full ledger: `.ai/audit/naming/tool_action_mapping.md`.

## 3. Complete action enum inventory (after)

`workbook`: open, close, save · `workbook_read`: status, list, model ·
`test_suite`: **run**, get-run · `vba_read`: project-info, list-components, list-procedures, read-component, get-references, search-code, get-compile-diagnostics ·
`macro`: list, classify-background, get-last-result · `sheet`: add, rename, copy, set-visibility · `sheet_read`: list, get ·
`table`: create, resize, append-rows, insert-row, update-rows, add-column · `table_read`: list, get, read-rows ·
`pivot_read`: tables-list, table-get, cache-get, fields-list, field-get-filters, result-read · `pivot_refresh`: refresh-table, refresh-cache ·
`pivot_field`: field-set-layout, field-set-position, field-set-aggregation, field-set-properties, field-apply-filter, field-clear-filters ·
`range_read`: get, read-values, read-formulas, get-format, get-validation, get-conditional-formatting ·
`range_write`: write-values, write-formulas, set-format, set-validation, set-conditional-formatting · `range_structural`: copy, insert, unmerge ·
`defined_name`: create, update · `defined_name_read`: list, get ·
`shape`: list, get, set-on-action · `control`: list, get, set-on-action ·
`workbook_snapshot_read`: metadata, sheet-rows, pivot-cache-records.

## 4. Old names / values absent — confirmation

- `tools/list` after: 43 tools, **0** of the 6 old names present (artifact diff: removed set == the 6 old, added set == the 6 new, other 31 identical).
- `test_suite` action enum == `["run","get-run"]`; `"run-suite"` rejected with `INVALID_RANGE`.
- No `run-suite` in any active description, source, or artifact.
- Client-facing dispatcher `context["tool"]` strings canonical for all renamed tools.
- **Accepted cosmetic exception (non-blocking):** each renamed tool's auto-derived `inputSchema.title` still reflects the unchanged *internal* Python function name (e.g. `excel_save_asArguments`, `testArguments`, `nameArguments`). This is generated metadata, not an addressable tool/action name; agents select on `name`/`description`. Internal symbols were intentionally left unchanged per the approved scope. See §9.

## 5. Test results (exact)

Environment: `.venv` Python; commands run from repo root.

| Check | Command | Result |
|---|---|---|
| Targeted contract/dispatch/baseline | `pytest tests/unit/test_wp0_baseline.py test_description_contract.py test_macro_test_dispatch.py test_macro_policy.py test_workbook_dispatch.py test_name_dispatch.py test_mapping_and_approval.py test_annotation_matrix.py test_tool_rename_contract.py` | **193 passed**, 0 failed |
| Full unit suite | `pytest tests/unit` | **1615 passed**, 0 failed, 0 skipped |
| Ruff | `ruff check src tests` | **All checks passed** (exit 0) |
| mypy (changed files) | `mypy src/excel_vba_mcp/server.py api/test_tools.py api/name_tools.py api/ooxml_tools.py` | **Success: no issues** (exit 0) |
| mypy (full src) | `mypy src` | 6 errors, **all pre-existing** in `adapters/excel/ranges_adapter.py` (large-data WIP, untouched) |

New contract coverage added in `tests/unit/test_tool_rename_contract.py`: canonical-present / old-absent / count-43 / handler-mapping / description+schema / annotation-triples / `test_suite` action enum / `run-suite` rejection / `run` & `get-run` dispatch / required-arg structured errors. `test_description_contract.py` gained annotation-total (14/16/23) and canonical/old-name assertions.

## 6. MCP runtime results

Authoritative dumps (via `mcp.list_tools()`, no Excel needed): `.ai/audit/artifacts/tools_list_before.json`, `tools_list_after.json`.
- 43 tools; annotations **14 readOnly / 16 destructive / 23 idempotent** (unchanged meaning; membership preserved 1:1).
- Renamed triples: workbook_save_copy F/T/F, test_suite F/T/F, defined_name F/F/T, defined_name_read T/F/T, defined_name_delete F/T/F, workbook_snapshot_read T/F/T.
- Resources intact: templates `excel://test-runs/{run_id}`, `excel://workbooks/{workbook_id}/model`, `.../vba/components`, `.../vba/components/{component_name}`; static `xlvba://glossary`.

## 7. Kilo permission + agent usability

- The Kilo client namespaces this server as **`xlvba_*`** (server name `excel-vba-mcp`; live tools observed as `xlvba_*`). The documented permission glob `xlvba_*` therefore matches every renamed tool — **no per-tool glob change needed** (0 explicit old-name globs in docs). Approval/confirmation policy unchanged.
- **Limitation (stated, not simulated):** the live Kilo-connected server process predates the rename (it exposes `xlvba_excel_save_as`, `xlvba_test` `run-suite`, …). Live new-name agent validation (WP8/WP9) requires **reconnecting the MCP server** so Kilo launches the updated build. No simulated calls were claimed. Static/artifact validation (WP7) and the independent Validator confirm the new surface is correct.

## 8. Independent review

- **Auditor** (`.ai/audit/reports/naming_auditor_report.md`): **ACCEPT-WITH-NONBLOCKING** — 0 blocking, 3 non-blocking (internal-only symbols / cosmetic schema titles). Re-ran 104 hermetic tests green.
- **Validator** (`.ai/audit/reports/naming_validator_report.md`): **VALIDATED-WITH-LIMITATION** — checks 1–6 PASS from artifacts; check 7 (live new-name) is the reconnect limitation. No simulated execution.
- Tech Lead personally reviewed the full source diff; the only non-naming hunk is a one-line `_build_runtime` re-wire matching the inherited `DataExportService` signature (no business logic entangled).

## 9. Remaining limitations (non-blocking)

1. **Cosmetic schema titles** retain internal function names (e.g. `excel_save_asArguments`) — generated metadata, not addressable names; internal symbols intentionally unchanged.
2. **Live Kilo new-name validation** pending an MCP server reconnect (old build still connected).
3. **Pre-existing large-data WIP** (`data_export_service.py`, `ranges_adapter.py`, etc.) is out of scope; its 6 mypy errors pre-date this task.
4. **Pre-existing 88-era prose debt**: some docs narrative still uses older consolidated-era per-verb names (e.g. `excel_open`, `sheets_list`) that predate the consolidation and are outside the approved 6 renames; recorded in the decision log for a separate follow-up.

## 10. Deliverables

1. `.ai/audit/naming/tool_action_naming_baseline.md`
2. `.ai/audit/naming/tool_action_mapping.md`
3. `.ai/audit/naming/tool_action_naming_policy.md`
4. Source: `server.py`, `api/test_tools.py`, `api/name_tools.py`, `api/ooxml_tools.py`
5. Tests: `tests/fixtures/wp0_baseline.json` + 6 updated unit tests + new `tests/unit/test_tool_rename_contract.py`
6. Docs: `README.md`, `docs/**` (EN + `docs/ru/**`), reconciled to 43 / 14/16/23
7. Kilo config: `xlvba_*` glob confirmed sufficient (no change required)
8. Artifacts: `.ai/audit/artifacts/tools_list_before.json`, `tools_list_after.json`
9. This report
10. Naming benchmark: not run (live Kilo new-name build unavailable pre-reconnect) — see limitation 2
11. `.ai/audit/reports/naming_auditor_report.md`
12. `.ai/audit/reports/naming_validator_report.md`
13. `.ai/audit/naming/tool_action_decision_log.md`