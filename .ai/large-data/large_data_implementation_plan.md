# Large-Data Implementation Plan

Atomic, independently testable work packages in priority order. Each WP has a single owner
(Implementor), explicit acceptance, and out-of-scope. The Tech Lead reviews and accepts each before
the next depends on it. No WP mixes unrelated tool-consolidation or doc work.

---

## Priority 0 — Core scalability

### WP-A  Unified limits + PageInfo + version primitive
- **Goal:** one `LargeDataConfig`; `PageInfo` DTO; `STALE_CURSOR` error; `SessionRecord.version` +
  `get_version`/`bump_version`.
- **Accept:** unit tests for config env load + validation; `PageInfo` serialization; version bump on a
  simulated mutation; `STALE_CURSOR` category/retryable.
- **Out:** no read behaviour change yet.

### WP-B  Paged reads over large ranges + tables + version binding
- **Goal:** remove full-range read pre-reject; wire `cursor_version` + `PageInfo` into
  `range_read`/`table_read`; raise table response cap to config.
- **Accept:** unit tests — a 200k-row logical range pages with each response ≤ cap and
  `has_more/complete/total` correct; a version bump between pages raises `STALE_CURSOR`; write cap
  unchanged; existing range/table tests still pass.
- **Depends:** WP-A.

## Priority 1 — Pivot result + metadata

### WP-C  Pivot rendered-result reader (all source types)
- **Goal:** `read_pivot_result` over `TableRange1` reusing the shared bounded range reader; new
  `pivot_read action="result-read"` with `offset/limit/cursor_version`; `PivotResultPage` DTO.
- **Accept:** unit tests with a fake pivot (range + fingerprint) for paging + stale detection; live
  integration test on a real large PivotTable (Tech Lead / Validator) retrieving pages without
  ambiguity; OLAP pivot *displayed result* read proven (metadata-only for members).
- **Depends:** WP-A, WP-B (shared reader).

### WP-D  OLAP/Data Model capability matrix + safe reads
- **Goal:** document + implement only research-proven safe reads; keep fail-closed elsewhere;
  structured unsupported outcomes.
- **Accept:** `large_data_pivot_capability_matrix.md`; tests assert unsupported ops return
  `supported=false` and never touch Excel; one safe OLAP read (displayed result) demonstrated.
- **Depends:** WP-C + research gate.

## Priority 1 — File / OOXML path

### WP-E  Streaming OOXML reader (non-COM) + ZIP/XML safety
- **Goal:** `adapters/ooxml/*` bounded reader; `OoxmlService` (immutable snapshot input, async job);
  `workbook_file_read` tool (`metadata`/`sheet-rows`/`table-rows`/`pivot-cache-records`).
- **Accept:** unit tests for ZIP-bomb/ratio/traversal rejection; streaming worksheet + shared-strings
  extraction equals COM values on a fixture; pivot cache records paged; corrupt archive fails safely;
  original + VBA parts unchanged (sha256); runs off the STA (no worker submit); cancellation leaves no
  partial file.
- **Depends:** WP-A (limits), snapshot store (existing).

## Priority 2 — File workflows

### WP-F  Streamed export (CSV/JSONL)
- **Goal:** `DataExportService` streams a source to an artifact temp file → atomic finalize; `data_export`
  tool; returns artifact reference + count + sha + completeness.
- **Accept:** export of a >100k-row fixture streams with bounded memory (tracemalloc), output parses,
  row count exact, original sha unchanged, no full data in response.
- **Depends:** WP-B/C/E.

### WP-G  Metrics + benchmark fixtures + report
- **Goal:** instrumentation + synthetic fixtures (100k/500k rows, wide, large pivot, xlsm, corrupt zip,
  zip-bomb) + benchmark script; `large_data_benchmark_report.md`.
- **Accept:** ≥3 reps where practical; records time-to-first-page, full extraction, peak memory, STA
  occupancy, response size, completeness, original-hash-unchanged.

## Priority 3 — Independent verification
- Auditor: source-level review of WP-A..F against the contract + safety.
- Validator: live Excel/MCP verification on real large workbooks, confirming actual workbook contents.
- Tech Lead: final verdict + `large_data_results.md` + updated tool/limitations docs.

---

## Acceptance matrix (WP → criterion)

| Criterion | A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|---|
| Bounded open/metadata | ✓ | | | | | | |
| Large sheet/table paged | | ✓ | | | | | |
| Version-consistent paging | ✓ | ✓ | ✓ | | ✓ | | |
| No silent truncation | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | |
| Pivot result paged | | | ✓ | | | | |
| OLAP safe read / unsupported explicit | | | ✓ | ✓ | | | |
| OOXML streaming + safety | | | | | ✓ | | |
| Export as artifact | | | | | | ✓ | |
| Original unchanged / no 2nd STA | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | |
| Benchmarks | | | | | | | ✓ |

## Risk register
- R1 Removing read pre-reject could let a caller request a huge window → mitigated by window cap
  (response ≤ max_response_cells) + `RANGE_TOO_LARGE` on the *window*.
- R2 Version bump wiring missed in one mutating service → stale cursor not detected there → mitigated
  by a unit test that mutates via each service and asserts version increments; auditor check.
- R3 OOXML snapshot staleness (unsaved changes) → mitigated by SaveCopyAs (captures in-memory state),
  never the disk working copy.
- R4 Pivot `TableRange1` semantics vary by layout → mitigated by reading the rendered range + totals,
  fingerprint + version, and documenting omitted hierarchy in the DTO.
- R5 Single-STA OLAP hang → mitigated by reading only displayed output cells (no external resolve),
  fail-closed for member enumeration, agent-directed recovery preserved.