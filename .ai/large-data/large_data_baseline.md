# Large-Data Baseline (WP0)

Status: WP0 inventory of the **current** `xlvba_mcp` large-data behaviour, before any
large-data implementation. Source references are `path:line` against the working tree at
commit `7961038` (plus uncommitted working-tree changes noted where relevant).

Scope: every layer that touches large data — tool surface, application services, Excel COM
adapters, session/worker, jobs, snapshots, path/artifact policy, models. For each capability
the data flow is classified (streaming vs full materialization, COM vs non-COM, bounded vs
unbounded, sync vs async, max rows/cells/bytes, memory growth, staleness risk, partial-vs-complete
ambiguity).

---

## 1. Architecture recap (constraints that shape everything)

- **One** MCP server, **one** owned hidden Excel instance, **one** STA COM worker, **one** FIFO
  queue (`adapters/excel/com_worker.py`). All Excel access is serialized through
  `ExcelComWorker.submit` (`com_worker.py:376`). The worker enforces a per-call deadline clamped
  to `liveness_ceiling_s` (default 600 s, `config.py:67`), never auto-kills a live pid, never
  replays a mutating call.
- Return-boundary proxy-leak guard: a worker `fn` may only return Pydantic models / primitives /
  containers (`com_worker.py:197` `_ensure_dto`). This bounds what can cross the boundary but also
  means a "large" read is materialized as a Python DTO **before** it crosses.
- Working-copy model: `excel_open` copies the source file to `work/<id>.ext`
  (`session_service.py:379` `_materialize_working_copy`) — a **file copy**, not cell materialization.
  The original is never opened for write; `source_sha256`/`source_size` are captured at open and
  re-verified at Save-As (`session_service.py:378`, `:779`).
- Path allowlist + atomic writes: `security/path_policy.py`, `storage/artifact_store.py`
  (temp file → `os.replace`, SHA-256 + size recorded).

---

## 2. Per-function data-flow classification

Legend — **Mat**: full materialization / streaming / bounded-window. **Bus**: COM / non-COM.
**Bnd**: bounded / unbounded. **Max**: documented ceiling. **Stale**: multi-page staleness risk.
**Partial**: can a partial result be mistaken for complete?

### 2.1 `excel_open` + workbook init — `session_service.py:291 open_workbook`
- Mat: file copy of the workbook (no cell read). Bus: non-COM copy + one STA `Open`.
- Bnd: bounded by file size. Max: `open_timeout_s` 300 s (`config.py:61`).
- Stale: n/a. Partial: n/a. **Verdict: already large-safe** (no full materialization).

### 2.2 `workbook_read model` — `application/workbook_model_service.py:142`
- Mat: aggregates inventory services only; **no Range/Cell reads** (`workbook_model_service.py:19`
  design rule). VBA procedures enumerated per component (bounded by component count).
- Bus: non-COM orchestration over read services. Bnd: bounded (object inventory).
- Stale: model_hash is content hash of metadata. Partial: per-section status
  (`available/unavailable/partial`) — explicit. **Verdict: bounded metadata; large-safe.**
  - Caveat: `get_sheet_model` calls `list_tables/list_names/list_shapes/list_controls/list_pivots`
  for the **whole workbook** then filters in Python (`:251-270`) — O(workbook inventory) per sheet
  model, not O(cells). Acceptable but not free for a workbook with thousands of objects.

### 2.3 `range_read` (values/formulas) — `application/ranges_service.py:164` / `:203`, adapter `ranges_adapter.py:859` / `:918`
- Mat: **bounded window** via a single bulk `.Value2` / `.Formula2` read of the sub-range
  (`ranges_adapter.py:900-904`, `_sub_range` `:782`). Good: one COM call per page, not per cell.
- Bus: COM (STA). Bnd: **bounded by `MAX_RESPONSE_CELLS=50_000`** and `MAX_RESPONSE_BYTES=8 MiB`
  (`domain/models.py:913-914`, enforced `models.py:988`, adapter `_window` `ranges_adapter.py:3740`).
- **CRITICAL LIMITATION:** `_window` (`ranges_adapter.py:3754`) rejects the request with
  `RANGE_TOO_LARGE` when the **full** range `total_rows*total_cols > MAX_RANGE_CELLS (100_000)` —
  **before** any paging. A 200 000-row sheet therefore **cannot be paged at all** through
  `range_read`; the only way to read it is to issue many smaller *explicit* A1 addresses that each
  individually cover ≤100 000 cells. This blocks the core "large worksheet in pages" workflow.
- Memory: `_normalize_grid` (`:668`) duplicates the COM tuple into nested Python tuples, then
  `_classify_grid` (`:728`) builds `list[list[CellValue]]` — ~2–3× the window in transient Python
  objects. Bounded by 50 000 cells, so acceptable per page.
- Stale: **offset/limit only, no version binding** — pages can mix workbook versions if a mutation
  lands between pages. Partial: `truncated` flag + `rows/cols` in `RangeAddress`; a truncated page
  is explicit, but there is no `has_more`/total/completeness beyond `truncated`.
- Mixed number formats: `_classify_grid` falls back to **per-cell** `NumberFormat` reads
  (`:762-778`) bounded to returned cells — a correctness/perf tradeoff, bounded.

### 2.4 Table rows — `adapters/excel/tables_adapter.py:420 read_rows`
- Mat: bounded window, single `.Value` read (`:494`). Bus: COM. Bnd: **`_MAX_RESPONSE_CELLS=2000`**
  (`tables_adapter.py:67`) — very small; a 50-column table returns ≤40 rows/page.
- Uses `.Value` (not `.Value2`) and re-classifies per cell (`:499-508`).
- Stale: offset/limit only, no version binding. Partial: `total_rows` + `truncated` present
  (`TableRows`), but `truncated` conflates "limit cut it" and "budget cut it".
- **Limitation:** 2 000-cell response cap makes large-table paging chatty; no total-page hint.

### 2.5 Worksheet inventory / used range — `adapters/excel/sheets_adapter.py`, `SheetInfo` (`models.py:641`)
- Mat: `UsedRange` address read once per sheet (metadata), not cell values. Bnd: bounded.
- Used-range scan for references is skipped when `row*col > MAX_RANGE_CELLS`
  (`ranges_adapter.py:3585`) — explicit note, not silent. **Verdict: bounded.**

### 2.6 PivotTables — `application/pivot_tables_service.py`, adapter `pivot_tables_adapter.py`
- **Metadata** (`list_pivot_tables`/`get_pivot_table`): single worksheet sweep
  (`inventory_single_sweep` `pivot_tables_adapter.py:1737`), source-type gated FIRST
  (`:567`), per-sheet cost meter budget `_DEFAULT_SHEET_FIELD_BUDGET=2000` (`:634`) → statuses
  `ok/capped/too_large` (`:637-639`). Reads `TableRange1/2`, `DataBodyRange` **addresses only**
  (`:557-559`), `RowGrand/ColumnGrand/RefreshDate`. Bounded, explicit partial (`capped`/`too_large`).
- **Fields/members/filters** (`list_pivot_fields`, `get_pivot_field_filters`): gated to `DATABASE`
  source (`_gate_supported_source_read` `:946`); bounded by `_MAX_PIVOT_ITEMS=10000`,
  `_MAX_PIVOT_FILTERS=100`, `_MAX_FILTER_ITEMS=1000` (`:1308-1312`).
- **CRITICAL GAP — no result-data reader.** There is **no tool that reads the rendered PivotTable
  output** (the values in `TableRange1`/`DataBodyRange`). Metadata reads the *addresses* but never
  the *values*. An agent cannot retrieve the actual PivotTable result rows at all, let alone paged.
- **OLAP / Data Model:** `SourceType`/`OLAP` detected (`_classify_source_type` `:327`); all
  non-`DATABASE` sources are fail-closed (`UNSUPPORTED_PIVOT_SOURCE`) for refresh/field/filter.
  No safe read path (displayed result, `RowItems`/`ColumnItems`, `CubeFields`, member unique names)
  is implemented. Capability matrix is therefore almost entirely "unsupported" for OLAP/Data Model.
- Stale: n/a (no result paging yet). Partial: source-type + sheet status explicit.

### 2.7 PivotCache access — `read_pivot_cache` (`pivot_tables_service.py:235`)
- Metadata only (`PivotCacheInfo` `models.py:1448`): source type, `record_count`, refresh metadata,
  shared-cache N:1. **No record reading** of any kind (no `PivotCache.Records`, no OOXML cache
  records). Large caches cannot be read at all.

### 2.8 Formula/value retrieval + array conversion — `ranges_adapter.py:668-986`, `cell_values.py`
- Covered in 2.3. `Formula2` primary with `Formula` fallback (`:793`). Prefix notes surfaced, never
  rewritten. Bounded by response cap.

### 2.9 File open / Save-As / snapshot / temp copy
- `excel_save_as` (`session_service.py:699`): pure-filesystem atomic temp→replace with SHA verify
  (`:913-983`), source re-hash guard (`:779`). Bounded by file size, non-COM. **Large-safe.**
- `snapshot_create` (`snapshot_service.py:139`): `SaveCopyAs` on STA (captures in-memory state to a
  separate file) or file-copy for closed books; manifest carries `stored_path` + `sha256`
  (`models.py:523-547`). **This is the immutable input mechanism available for a future OOXML reader.**
- Artifact store (`artifact_store.py`): atomic, path-policy enforced, SHA-256 + size, index.json.
  `save_artifact` takes the **whole payload as `bytes`** (`:208`) — i.e. **not streaming**; an
  export of millions of rows would have to be fully materialized in memory before saving.
  `ArtifactType.EXPORT` maps to `.zip` only (`:78`).

### 2.10 Operations that materialize full worksheets / large collections
- No operation reads an entire `UsedRange.Value` unbounded. The structural merge/protection scans
  are explicitly bounded (`_MAX_MERGE_SCAN_CELLS=4096` `:303`). The only full-materialization risk
  is the worker return-boundary DTO for a single page (≤50 000 cells) and `artifact_store` bytes.
- `test_runner.py` (2 435 lines) runs declarative suites on disposable copies; not a large-data read
  path.

### 2.11 JSON serialization / MCP response size
- Worker `_ensure_dto` forces `model_dump(mode="json")` once at the boundary (`com_worker.py:212`).
  The MCP layer serializes the DTO again. Response size is bounded only indirectly by
  `MAX_RESPONSE_CELLS`/`MAX_RESPONSE_BYTES`. No explicit serialized-byte budget on the MCP side.

### 2.12 Memory allocation / duplicate copies
- Per page: COM tuple → `_normalize_grid` tuples → `CellValue` DTOs → JSON. ~2–3× transient. Bounded
  per page; **does not grow with total dataset** for range/table reads (good). Would grow with total
  dataset for any future "read whole sheet" path (none exists) and for `artifact_store` bytes export.

### 2.13 Job execution / polling — `application/job_registry.py`
- Loop-side, owns the result; supports **sync** (wrapped by single-STA `submit`) and **async**
  callables (`start` `:241`, `_run_async` `:469`). `job_get` polls by `job_id`. `poll_after_s`
  clamped [1,30]. **This is the correct hook for a non-COM background OOXML job**: an async callable
  that does file parsing off the STA thread, owned by the registry, no second STA.
- Persistence via `JobStore`; interrupted jobs finalize to `interrupted_unknown`, never replayed.

### 2.14 Import / export size limits
- No large-data import/export exists. `artifact_store` export is bytes-in-memory (2.9). No CSV/JSONL
  streaming. No server-side filtering/aggregation.

---

## 3. Implemented capabilities vs known limitations (summary table)

| # | Capability | State | Key limit / gap | Source |
|---|---|---|---|---|
| 1 | Open large workbook without cell materialization | **OK** | file copy only | `session_service.py:379` |
| 2 | Bounded workbook metadata model | **OK** | O(objects), not O(cells) | `workbook_model_service.py:142` |
| 3 | Paged range read (bulk `.Value2`) | **PARTIAL** | full-range 100k pre-reject blocks paging >100k cells | `ranges_adapter.py:3754` |
| 4 | Paged table read | **PARTIAL** | 2 000-cell response cap; no version binding | `tables_adapter.py:67` |
| 5 | Stable pagination / version binding | **MISSING** | offset/limit only; no stale detection | `ranges_adapter.py:3740` |
| 6 | Pivot metadata without item enumeration | **OK** | source-gated + cost meter | `pivot_tables_adapter.py:1737` |
| 7 | Bounded field/member/filter enumeration | **OK** | 10k items / 100 filters | `pivot_tables_adapter.py:1308` |
| 8 | **Pivot result-data reader** | **MISSING** | no tool reads rendered output | — |
| 9 | OLAP / Data Model safe read path | **MISSING** | fail-closed only | `pivot_tables_adapter.py:946` |
| 10 | PivotCache record reading | **MISSING** | metadata only | `pivot_tables_service.py:235` |
| 11 | Streaming OOXML (non-COM) reader | **MISSING** | none | — |
| 12 | Large file export (CSV/JSONL, streamed) | **MISSING** | artifact store is bytes-in-memory | `artifact_store.py:208` |
| 13 | Large import / server-side filter/aggregate | **MISSING** | none | — |
| 14 | Immutable snapshot input for file reads | **OK (primitive)** | `SaveCopyAs` + sha manifest | `snapshot_service.py:197` |
| 15 | Non-COM background job execution | **OK (primitive)** | async job supported | `job_registry.py:469` |
| 16 | Path allowlist + atomic artifact writes | **OK** | extension set is workbook/zip | `path_policy.py`, `artifact_store.py` |
| 17 | ZIP/XML safety (bomb, ratio, traversal) | **N/A** | no OOXML reader yet | — |
| 18 | Large-data metrics/observability | **MISSING** | only audit + worker notes | — |

## 4. Reproduction evidence (baseline)

- 100 000-cell cap: `domain/models.py:909` `MAX_RANGE_CELLS = 100_000`; enforced pre-read at
  `ranges_adapter.py:3754`; `RANGE_TOO_LARGE` mapping `domain/errors.py:49,246`.
- Response cap: `domain/models.py:913-914`; enforced `models.py:988` and `ranges_adapter.py:3770`.
- Table cap: `tables_adapter.py:67` `_MAX_RESPONSE_CELLS = 2000`.
- Pivot source gate: `pivot_tables_adapter.py:327,946,1737`; `_UNSUPPORTED_SOURCE_TYPES` `:191`.
- No pivot result reader: `grep -R "DataBodyRange"` → only address reads in
  `pivot_tables_adapter.py:559`; no `.Value2` of a pivot range anywhere.
- Artifact store is bytes-in-memory: `artifact_store.py:208` `data: bytes`.
- Job registry async path: `job_registry.py:289,469`.

## 5. Consequences for the large-data design

1. **Remove the full-range pre-reject for reads** so a >100 000-cell range can be paged; keep the
   per-response cap and the write cap.
2. **Add version binding** (per-workbook monotonic content version) + an explicit completeness
   contract (`total`, `offset`/`cursor`, `has_more`, `complete`, `version`, `stale`).
3. **Add a bounded Pivot result reader** over `TableRange1`/`DataBodyRange` reusing the range reader.
4. **Add a non-COM streaming OOXML reader** fed by an immutable snapshot copy, with ZIP/XML safety,
   run as an async registry job (no second STA).
5. **Add streamed file export** (CSV/JSONL) writing incrementally to an artifact (not bytes-in-memory).
6. Keep OLAP/Data Model fail-closed until a safe read path is proven; never weaken the gate blindly.