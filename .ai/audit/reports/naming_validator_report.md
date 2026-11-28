# Naming Refactor — Independent Validator Report

**Project:** `xlvba_mcp` · **Reviewer:** Independent Validator (read-only + artifact inspection) · **Date:** 2026-11-28
**Scope:** Tech-Lead-approved MCP tool+action rename only. No code/test/doc edits performed. No live/Excel/COM calls executed.

## Verdict: VALIDATED-WITH-LIMITATION

All 7 checks PASS against the authoritative artifacts and source. The single limitation is that the live Kilo-connected server process predates the rename (exposes old names), so live new-name agent validation requires an MCP reconnect. This is a stated limitation, not a pass, and does not affect the artifact-based verdict.

---

## Check results

| # | Check | Result | Evidence |
|---|-------|--------|----------|
| 1 | after = 43 tools; 6 canonical present; 6 old absent; no dupes | **PASS** | `tools_list_after.json` parsed: `Count=43`, `unique names=43`. Canonical present: `workbook_save_copy, test_suite, defined_name, defined_name_read, defined_name_delete, workbook_snapshot_read`. Old absent: `excel_save_as, test, name, name_read, name_delete, workbook_file_read` (none in after-name list). |
| 2 | before↔after diff = exactly the 6 removed + 6 added; others identical | **PASS** | Programmatic set diff. REMOVED = `excel_save_as, test, name, name_read, name_delete, workbook_file_read` (matches expected). ADDED = `workbook_save_copy, test_suite, defined_name, defined_name_read, defined_name_delete, workbook_snapshot_read` (matches expected). Common 31 names identical and in the same order. before Count=43, after Count=43. |
| 3 | action enums canonical | **PASS** | `test_suite.action.enum = [run, get-run]` (no `run-suite`); `defined_name = [create, update]`; `defined_name_read = [list, get]`; `workbook_snapshot_read = [metadata, sheet-rows, pivot-cache-records]`. |
| 4 | annotation totals 14/16/23 + per-tool triples | **PASS** | Totals: `readOnly=14, destructive=16, idempotent=23` (43 tools). Triples (R/D/I): workbook_save_copy F/T/F; test_suite F/T/F; defined_name F/F/T; defined_name_read T/F/T; defined_name_delete F/T/F; workbook_snapshot_read T/F/T — all match. |
| 5 | descriptions reference canonical name, no "run-suite" | **PASS** | All 6 renamed-tool descriptions contain their canonical name (`referencesCanonical=True`) and none contain `run-suite`. Full after-artifact scan: `run-suite` absent from all 43 descriptions. |
| 6 | Kilo `xlvba_*` glob still matches; any explicit per-tool glob to update? | **PASS** | `README.md:158-164`, `docs/02-installation-and-setup.md:127-137,264-268` (and `docs/ru/...`) document `"xlvba_*": "ask"` as a **server-name** glob that matches every published tool name; README explicitly states renamed tools remain covered "without any configuration change." Exact whole-word search of `docs/**` + `README.md` for old names (`excel_save_as`, `workbook_file_read`, `name_read`, `name_delete`) in permission context: **0 matches**. No explicit per-tool permission glob references any old tool name → **no permission config update required**. (Only explicit per-tool example is `"xlvba_server_health": "allow"`, unrelated to the rename.) |
| 7 | Live server predates rename → reconnect required (limitation) | **LIMITATION (not a pass)** | The live Kilo-connected session exposes OLD tool names (`xlvba_excel_save_as`, `xlvba_test`, `xlvba_name`, `xlvba_name_read`, `xlvba_name_delete`, `xlvba_workbook_file_read`) and `xlvba_server_health` reports `service="excel-vba-mcp" version="0.1.0"`. Live new-name agent validation therefore requires an MCP reconnect. Not executed here by design. |

---

## Notes

- **No simulation:** No live/Excel/COM/MCP tool call was performed. `xlvba_server_health` was the only server touch and it is a pure liveness probe (no COM, no workbook). All other evidence is from saved artifacts (`.ai/audit/artifacts/tools_list_before.json`, `tools_list_after.json`), the mapping ledger, the auditor report, and static source/doc reads.
- **Consistency with auditor report:** The independent auditor report (`naming_auditor_report.md`, verdict ACCEPT-WITH-NONBLOCKING, 0 blocking) is consistent with these artifact findings (43/43, 14/16/23, correct triples, `run-suite` absent, ledger 94→43). The auditor's non-blocking residuals are internal-only symbols / cosmetic schema titles (e.g. `excel_save_asArguments` inputSchema title, `SaveAsResult` docstring) — none are a published tool name, action value, or `context["tool"]`, and none affect this validation.
- **Mapping ledger:** `tool_action_mapping.md` §3 asserts 94 legacy old-names → 43 canonical tools with completeness True; the 6 approved tool renames + `run-suite→run` (get-run kept) match the artifacts exactly. No aliases.
- **Deliverable path:** `.ai/audit/reports/naming_validator_report.md` (this file).