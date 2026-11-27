# Pivot Hang Remediation — Results (unit vs live reported separately)

- **Revision:** `55b9398` + WP1/WP2/WP3/WP-Range/WP6 (uncommitted).
- **Architecture:** unchanged (single hidden Excel, single-STA FIFO, Job Object, working-copy, SCHEMA_VERSION=1). No timeout/background changes (deferred). No new workers.

## 1. Root cause (recap)
`pivot_read(fields-list)` / `field-get-filters` enumerated `PivotFields` on external/OLAP pivots with no source-type gate → Excel resolves external data → single-STA blocks past the 60s soft deadline → client `-32001` (~70s) → `excel_recover`. `pivot_tables_list`/refresh already gated; these two did not.

## 2. WP1 — source-type gate (A, approved)
- `_gate_supported_source_read(pivot, pivot_table)` (adapter:941) calls `PivotCache`→`_classify_source_type`; if not `DATABASE` (covers `_UNSUPPORTED_SOURCE_TYPES` + `UNKNOWN`, fail-closed) → `UnsupportedPivotSourceError` (`UNSUPPORTED_PIVOT_SOURCE`, context `{pivot_table, source_type, reason:"unsupported_source"}`) **before** any `PivotFields`.
- Call sites: `enumerate_pivot_fields` (adapter:1926, before `_data_field_index`/`_describe_all_fields`), `get_field_filters` (adapter:2610, before `_resolve_field_target`). DATABASE behavior unchanged.
- **Unit (8 cases):** EXTERNAL/OLAP/UNKNOWN → raises; fake `PivotFields`/`DataFields` accessors are tripwires (never touched); DATABASE returns fields. **PASS.**
- **Live integration (real Excel, `test_pivot_read_source_gate.py`):** external/OLAP workbook `C:\py_exp\xlvba_mcp\.ai\tmp\theiss\weekly\report_w.xlsx` → `UNSUPPORTED_PIVOT_SOURCE`, elapsed < 10s, worker still accepting (**no recovery**); supported DATABASE pivot (`PivotTablesProject.xlsm`, `ptA`) returns fields. **2 passed in 5.91s, exit 0.**
- **Before/after:** before = hang >70s → `-32001` → recover; after = prompt structured `unsupported_source` (~seconds), no recovery.

## 3. Other approved fixes (unit-verified)
- **WP2 not_found (D):** missing PivotTable/cache/sheet/name now `NOT_FOUND` (category `not_found`, isRetryable=false) instead of `COM_BUSY`; genuine conflicts/transient (already-open, RPC) unchanged. Tests added.
- **WP2 sheet_copy target_name (A):** optional `target_name` renames the copy (validated ≤31 chars, unique); None → unchanged. Tool schema/desc + docs updated. Tests added.
- **WP-Range raw scalars:** `write-values` accepts raw JSON scalars as literals (string→string incl leading `=`; number→number; bool→boolean; null→empty; no ambiguous string coercion); `CellValue` dicts still accepted; formulas exclusive to `write-formulas`. Schema `anyOf[CellValue,string,integer,number,boolean,null]`. Tests added.
- **WP-Range 100k (7):** additive `ErrorCode.RANGE_TOO_LARGE` (category validation, isRetryable=false) for a valid range over `MAX_RANGE_CELLS=100_000`; `INVALID_RANGE` kept for malformed. Tests added.
- **WP3 snapshot_restore (5):** lifecycle = open restored → validate restored → close obsolete previous (discard); success leaves exactly 1 open (restored); validation failure keeps previous open; never closes both. `confirm` gating unchanged. 3 tests added.
- **WP6 docs:** R13 MsgBox/headless limitation + 100k limit documented (EN+RU); veto not weakened.

## 4. Verification summary
- **Unit:** `pytest tests/unit -q` → **1454 passed**, ruff **clean**.
- **Live integration:** WP1 external+supported → **2 passed (5.91s)**, no recovery. (Other WPs covered by unit; live Excel integration for them not separately run.)
- **Deferred (per decision 3):** timeout/background-job path — reconsider only if benchmarks show supported-source enumeration is slow.

## 5. Notes
- Transient `0x80010108` recover-once behavior unchanged (decision 9).
- mypy strict shows 6 pre-existing errors in the untouched conditional-formatting section (outside scope).
- Changes uncommitted (no commit requested).