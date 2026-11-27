# Large-Data Test Fixtures (provided + synthetic)

> **Policy (per Tech Lead instruction):** the provided `.ai/tmp/theiss/` dataset is **development
> data only**. It is **never** referenced by unit or any automated test, is **not** committed to git
> (too large), and must be **deleted when this task completes**. It is used solely for manual/agent
> validation of the large-data tools. All automated tests use **synthetic** fixtures generated
> ephemerally under `.ai/tmp` (or `tmp_path`) at run time — never the provided files.
>
> **Deletion status:** automated deletion was attempted but blocked by a `Remove-Item -Recurse -Force`
> permission rule. **Manual removal required:** delete `C:\py_exp\xlvba_mcp\.ai\tmp\theiss\` (and the
> `.ai/tmp/bench/` working copies). Confirmed: no file under `tests/` references this dataset except
> the pre-existing `tests/integration_excel/test_pivot_read_source_gate.py` (Excel-gated, not run in CI).

## Provided real workbooks — DEV/AGENT VALIDATION ONLY (not for automated tests; delete at task end)
Root: `.ai/tmp/theiss/`. Verified read-only via non-COM zip inspection.

| File | Size | Sheets | Notable |
|---|---|---|---|
| `monthly/report_m.xlsx` | 5.61 MB | 7 (`clean_data_full`, `agregated_dict`, `dashboard_data`, `FilePath`, `merged_dictionary`, `column_settings`, `dashboard_sortings`) | `dashboard_data` sheet XML = **30.8 MB uncompressed**; 1 PivotTable; 1 `pivotCacheRecords`; 2 `pivotCacheDefinition`; sharedStrings present; **cache `saveData=False`** (records not saved locally) + `refreshOnLoad=False` |
| `weekly/report_w.xlsx` | 5.51 MB | 7 (same names) | same shape; `dashboard_data` = **30.8 MB uncompressed**; 1 PivotTable; `saveData=False` |

Also present: per-week/per-month `raw_data/*.csv.gz` (small), `clean_data/*.csv.gz`
(1.29–2.90 MB gz — large when decompressed), `output/*.xlsx` (small brand outputs).

### Why these are ideal
- **Large worksheet paging (WP-B):** a 30.8 MB sheet cannot be read in one response; exercises the
  removed >100k pre-reject + window cap + version-bound paging.
- **Streaming OOXML (WP-E):** real sharedStrings + a 30.8 MB `sheet3.xml` exercises bounded
  `iterparse` + shared-string streaming; the 5.6x zip ratio is a realistic (non-bomb) archive.
- **Pivot result reader (WP-C):** a real rendered PivotTable on `dashboard_data`.
- **Cache freshness (WP-E):** `saveData=False` is a genuine "PivotCache records not saved / stale vs
  source" case — the reader must report records unavailable rather than fabricate them.
- **Export (WP-F):** export the 30.8 MB sheet to CSV/JSONL and verify streamed memory + exact counts.

### Safety
- Originals are read-only inputs; all tests copy to `.ai/tmp` working copies first (existing
  working-copy model). Original SHA-256 must be unchanged after every test (asserted).
- `.xlsm`/VBA fixtures: none in this dataset; the synthetic `.xlsm` fixture (WP-G) covers the
  VBA-package-untouched requirement.

## Synthetic fixtures to generate (WP-G)
Small/medium (fast unit): 100k-row single sheet; 120k-row (just over the old 100k cap) for the
paging test; wide 200-col sheet; large Excel Table; standard DB PivotCache + large rendered pivot;
`.xlsm` with a VBA module; corrupt OOXML (truncated zip); ZIP-bomb (high ratio) for safety tests.
Large (benchmark, generated on demand into `.ai/tmp`): 500k-row sheet; multi-million-cell sheet
subject to resources. Generator script + exact parameters recorded in the benchmark report.

## Benchmark inputs
Primary: `monthly/report_m.xlsx` and `weekly/report_w.xlsx` (copy to `.ai/tmp/bench/` per run).
Secondary: generated 500k-row synthetic for scaling curves. ≥3 reps where practical.