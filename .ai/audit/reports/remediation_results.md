# Remediation Results — Error Taxonomy & Final Usability Corrections

- **Baseline HEAD:** `7961038`. No production rollout. No commit (not authorized).
- **Constraint compliance:** one owned Excel; single-STA FIFO; original protection; snapshot/approval; structured envelope; no blind retries; no new worker/architecture; no external-Pivot change; no DTO/schema churn; `SCHEMA_VERSION=1`.

## WP results
- **WP1 (error taxonomy) — DONE.** All deterministic raise sites in `names_adapter.py`/`tables_adapter.py` now non-retryable with correct codes (NOT_FOUND / INVALID_RANGE / NAME_CONFLICT); genuine transient COM re-raised retryable (`com_transient.py` guard). Additive `NAME_CONFLICT`. See `error_taxonomy_remediation_report.md`.
- **WP2 (RANGE_TOO_LARGE doc) — DONE.** docs/09 (+RU) "pending" replaced with delivered behavior; consistent with 05/08; reject-before-execution, non-retryable, split/paginate, no silent truncation.
- **WP3 (snapshot R1–R3) — VERIFIED COMPLETE** (implemented in `2a90e7a`; Auditor `report_13` ACCEPT). Caller contract documented (docs/05 +RU). See `snapshot_restore_final_contract.md`.
- **WP4 (sheet_copy target_name) — DONE.** Source-verified + live regression (`test_sheets_copy_target_name_live.py`): rename, default `… (2)`, duplicate/invalid rejected (INVALID_RANGE), hashes unchanged.

## Test commands & results (reported separately)
- **Unit:** `.venv\Scripts\python.exe -m pytest tests/unit -q` → **green** (incl. new taxonomy + transient-COM tests). Ruff: **All checks passed**.
- **Live Excel (real hidden single-STA Excel, disposable workbooks):**
  - `pytest tests/integration_excel/test_error_taxonomy_live.py …` → **4 passed** (duplicate→NAME_CONFLICT, unknown header→NOT_FOUND, missing table/sheet→NOT_FOUND; no partial mutation; original hash unchanged).
  - `pytest tests/integration_excel/test_sheets_copy_target_name_live.py …` → **1 passed**.
  - `pytest tests/integration_excel/test_snapshot_restore_hardening.py …` → **5 passed** (R1/R2/R3 + interruption).
- **Source-verified:** raise-site codes, NAME_CONFLICT, transient guard, R1–R3, sheet_copy forwarding/validation.
- **Live MCP-server-process note:** the running MCP server still holds pre-WP1 code; in-process MCP validation of the new envelopes requires a server restart. The live integration tests exercise the same service/adapter code path against real Excel.

## Independent review
- Auditor `report_14` → **ACCEPT** (WP1–WP4 PASS; unit 1483, live 5).
- Validator `report_14a` → **APPROVE** (all items, unit + live).
- MEDIUM advisory (broad wrappers masking transient COM) → **closed** by `com_transient.py` guard + 13 unit tests.

## Remaining limitations
- In-process MCP-server restart needed to validate the new envelopes through the live server process (code path identical).
- No L2 VBA compile / no macros run (out of scope).
- excel-mcp `table append` opacity is out of scope (other server).

## Final decision: **ACCEPTED** (source-verified + unit-tested + live-Excel-verified; independent Auditor + Validator pass).