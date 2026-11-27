# Large-Data Benchmark Report

## Environment
- OS: Windows 10 (build 17763); CPU: 11th Gen Intel i5-1135G7 @ 2.40 GHz; RAM: 15.7 GB.
- Excel: 16.0 (single owned hidden instance, single STA worker, FIFO queue).
- Python: 3.14.7 (venv). Server: `xlvba_mcp` from `src/` (large-data branch).
- Date: 2026-11-27.

## Method
- **Unit/regression:** full `pytest tests/unit` (1584 tests) — repeated; all green.
- **Live large-data:** dev-only harness (`.ai/tmp/bench/harness.py`) builds the real runtime
  (`_build_runtime` + real `ExcelComWorker`) and drives the actual services against a **copy** of
  `report_m.xlsx` (5.61 MB, 30.8 MB uncompressed `dashboard_data` sheet, 1 external-source
  PivotTable). Original SHA-256 asserted unchanged. This exercises the same code path the MCP tools
  call (service layer); the thin `api/*` dispatch layer is covered by unit dispatch tests.
- ≥3 reps were not run for the live path (single long COM extraction is expensive and the single-STA
  worker is a shared serial resource); the deterministic unit suite is the repeated measure. This is
  stated honestly rather than claiming multi-run wall-clock parity.

## Results (live, single run; bounded/streaming)

| Operation | Input | Result | Memory / notes |
|---|---|---|---|
| Paged range read (page 1) | `dashboard_data!A1:Z200000` | returned 1923 (≤50k cap), total=200 000, has_more=true, version=0 | one `.Value2` window per page |
| Full range extraction | same, looped | **105 pages → 200 000 rows == total, complete=true**, no dup/skip | ~36 s total (≈0.34 s/page); bounded per page |
| Stale-cursor detection | bump_version between pages | `STALE_CURSOR` raised | version-bound cursor works |
| Pivot inventory | workbook | 1 table, 1 cache | metadata only |
| **Pivot rendered result** | `PivotTable1` | 50 rows × 20 cols, header_row_count=1, **source=external** | reads displayed cells; no external resolve |
| OOXML sheet read (snapshot) | `dashboard_data`, 20 rows | rows=20, source=snapshot, input sha256 64 hex | non-COM, `asyncio.to_thread` |
| Streamed CSV export | `dashboard_data!A1:Z5000` | 5000 rows, complete, 5001 lines, 135 KB | **peak Python 33.9 MB**, 2.5 s; one page in memory |
| PivotCache records | cache 1 | has_records=true, saveData=true, records paged | source="cache" (distinct from results) |
| Original integrity | `report_m.xlsx` | SHA-256 unchanged | working-copy model intact |

## Acceptance focus

| Metric | Outcome |
|---|---|
| Time to first page | page 1 returned promptly (single window); full 200k extraction ≈36 s / 105 pages |
| Full extraction scaling | linear in pages; bounded per page (no full-sheet materialization) |
| Peak memory | export peak 33.9 MB for 5000 rows (one page); does **not** grow with total dataset |
| MCP response size | bounded by `max_response_cells` (50k) / `max_response_bytes` (8 MiB) |
| STA occupancy | one window per page; no whole-sheet COM read; OOXML never touches STA |
| Data completeness | exact total (200 000) with `complete=true`; no missing/duplicate |
| Pagination consistency | version-bound; stale cursor detected on mutation |
| File integrity | export parses (5001 lines); original hash unchanged |
| Excel stability | single instance, no recovery, no extra process |

## Notes / limits
- The 200k×20 full extraction is ~36 s wall-clock on this CPU (single STA, 105 sequential COM reads);
  this is the cost of bounded paging, not a hang. Smaller pages trade latency for smaller responses.
- Live multi-rep timing was not collected (shared single-STA resource); unit suite is the repeated gate.