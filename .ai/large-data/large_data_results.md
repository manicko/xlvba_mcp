# Large-Data Results

## Verdict: **LARGE-DATA CAPABILITY VERIFIED** (core workflows)

The critical large-data workflows pass on real Excel with a real 30.8 MB worksheet, bounded memory,
version-consistent paging, and unchanged originals. Remaining limitations are explicit below.

## What was delivered (all merged, unit-tested, live-validated)
- **Unified limits** (`LargeDataConfig`, env `EXCEL_VBA_MCP_LARGE_DATA_*`) — single source; no per-tool limits.
- **Response contract** (`PageInfo`): offset/limit/returned/total/has_more/complete/truncated/version/source/reason.
- **Stable pagination + version binding**: per-workbook content `version` bumped on every successful
  mutation; paged reads capture the version atomically inside the STA fn; mismatch → `STALE_CURSOR`.
- **Large range/table reads**: removed the >100k full-range pre-reject for reads; a 200 000-row sheet
  pages at ≤50 000 cells/response; table cap raised to 50 000 (config-driven).
- **Pivot rendered-result reader** (`pivot_read action="result-read"`): paged `TableRange1` values for
  **all** source types (proven on an **external**-source pivot); header-row count + fingerprint.
- **Streaming OOXML reader** (`workbook_file_read`, non-COM, `asyncio.to_thread`, immutable snapshot
  input): metadata, worksheet rows (shared strings, inline/typed), PivotCache records; ZIP/XML safety
  (member count, uncompressed size, ratio, traversal, depth).
- **Streamed export** (`data_export`): range/table/pivot-result → CSV/JSONL artifact, one page in
  memory, atomic finalize, returns only path/rows/bytes/sha/completeness.
- Server wiring: 2 justified new tools; WP0 baseline coherently extended (92→94 rows / 41→43 targets).

## Data sizes tested
- Real: `report_m.xlsx` 5.61 MB; `dashboard_data` sheet **30.8 MB uncompressed**, 200 000 rows × 20 cols
  (read fully in 105 pages); external-source PivotTable (50×20 result); PivotCache with records.
- Unit: synthetic tiny OOXML (worksheet/sharedStrings/pivotCache) + paging/stale/export/hermetic suites.

## Streaming vs materializing
- **Streaming / bounded (one page in memory):** range read, table read, pivot result, OOXML worksheet,
  OOXML pivot-cache records, CSV/JSONL export. Memory does **not** grow with total dataset.
- **Still materializes (bounded):** one response page (≤50 000 cells) as Python DTOs before JSON;
  the worker return-boundary DTO. Bounded per page, not per dataset.
- **Non-COM:** OOXML read + export run off the STA worker (export uses the async services; OOXML uses
  `asyncio.to_thread`). No second Excel/STA.

## Pivot sources supported
- **DATABASE:** full (metadata, fields, members, filters, refresh, **displayed result**, cache records).
- **EXTERNAL / OLAP / DATA_MODEL:** identity, source metadata, dimensions, and **displayed result (paged)**
  — proven on an external-source pivot. Field/member/filter enumeration and refresh remain **fail-closed**
  (structured `unsupported_source`) because single-STA hang safety for member enumeration was not
  executed on our runtime (reference evidence is inferred). Capability matrix: `large_data_pivot_capability_matrix.md`.

## Cache freshness
- OOXML reader reports `saveData`, `refreshOnLoad`, `refreshDate`, `recordCount`, and
  `has_records_part`; when records are absent/unsaved it returns `reason="unsupported_source"`,
  `complete=false` — never fabricated. Displayed results vs saved cache records vs live source are
  distinct DTOs (`source` = live/snapshot/cache).

## Max tested dimensions / memory / time
- Max sheet read fully: 200 000 rows × 20 cols (4 M cells) via paging.
- Peak Python (export): 33.9 MB (one page). Full 200k extraction ≈36 s (105 pages, single STA).
- Original SHA-256 unchanged after all ops.

## Remaining limitations (explicit)
1. **OLAP/Data Model member/CubeField enumeration** is deferred (fail-closed) pending live single-STA
   safety validation; displayed results ARE supported.
2. **MCP-transport validation** of the 2 new tools was done in-process (real runtime + service layer +
   dispatch unit tests); the pre-connected MCP build predates these changes, so a JSON-transport run of
   `workbook_file_read`/`data_export` was not executed here.
3. **Import** of large CSV and server-side filtering/aggregation are not implemented (no justified use
   case yet); export + selective column/row export via range/table is.
4. **Parquet** export deferred (dependency growth); CSV/JSONL delivered.
5. Full 200k×20 extraction is ~36 s single-STA (bounded, not hung); page size trades latency vs response size.
6. `.xlsb` not supported (no research path); `.xlsx`/`.xlsm` only.

## Verification levels
1. Source review — Tech Lead + independent Auditor (PROCEED WITH CHANGES; fixes applied).
2. Unit — 1584 pass (config/models/errors/session/range/table/pivot/ooxml/export/dispatch/baseline).
3. Real Excel — live harness on the 30.8 MB workbook (all PASS).
4. MCP tool-call — dispatch unit tests + in-process service calls (see limitation 2).
5. Benchmark — `large_data_benchmark_report.md`.
6. Independent review — Auditor report `.ai/large-data/large_data_auditor_review.md`.