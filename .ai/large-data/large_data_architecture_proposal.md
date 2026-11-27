# Large-Data Architecture Proposal

Proposed architecture to give `xlvba_mcp` genuine large-data capability while preserving every
accepted constraint (single owned Excel instance, single STA worker, FIFO queue, working-copy
protection, snapshot/approval gates, job lifecycle, honest errors, consolidated tool surface).

The design is **additive and layered**: new capabilities reuse existing primitives (range reader,
job registry, artifact store, snapshot store, path policy) and add a small number of justified new
modules/tools. No re-architecture of the COM/session model.

---

## 1. Component map (what changes, what is reused)

```
                         ┌────────────────────────── MCP tool surface (api/*.py) ──────────────────────────┐
                         │  range_read  table_read  pivot_read(+result-read)  data_export  workbook_file_read │
                         └───────────────┬───────────────────┬───────────────┬───────────────┬──────────────┘
                                         │                   │               │               │
        ┌────────────────────────────────▼─────┐   ┌─────────▼──────┐  ┌─────▼──────┐  ┌─────▼───────────────┐
        │ RangesService / TablesService /       │   │ PivotTables   │  │ DataExport │  │ OoxmlService (NEW)  │
        │ PivotTablesService (async, COM-free)  │   │ Service       │  │ Service(NEW)│  │ non-COM (NEW)       │
        │ + NEW: version bump on mutation       │   │ +result reader│  │ streams    │  │ streaming ZIP/XML   │
        └───────────────┬───────────────────────┘   └───────┬───────┘  └─────┬──────┘  └─────┬───────────────┘
                        │ single-STA submit                │                │ artifact      │ async job (no STA)
                        ▼                                  ▼                ▼ store         ▼
                 ┌──────────────────── ExcelComWorker (single STA, FIFO) ────────┐   SnapshotStore (immutable input)
                 │  RangesAdapter / TablesAdapter / PivotTablesAdapter (COM)    │   ArtifactStore (atomic, streamed)
                 └──────────────────────────────────────────────────────────────┘
```

Reused unchanged: `ExcelComWorker`, `SessionService` working-copy model, `SnapshotService`
(`SaveCopyAs`), `ArtifactStore` atomic write, `PathPolicy`, `JobRegistry` (async path), `AuditLogger`.

New modules: `adapters/ooxml/*` (non-COM), `application/ooxml_service.py`,
`application/data_export_service.py`, `LargeDataConfig` (config), `PageInfo`/`STALE_CURSOR`.

---

## 2. Work package designs

### WP-A — Unified limits + response contract (Priority 0)
- `config.py`: add `LargeDataConfig` + env loading (single source of limits).
- `domain/models.py`: add `PageInfo` (additive) and `version` to read DTOs via an embedded optional
  `page: PageInfo | None`. Add `STALE_CURSOR` to `domain/errors.py` (+ category + recovery).
- `session_service.py`: `version` on `_SessionRecord`; `get_version` / `bump_version`.
- Mutating services call `bump_version` in `_submit_mutation` success (one line each).

### WP-B — Paged reads over large ranges + version binding (Priority 0)
- `ranges_adapter._window`: for **reads**, drop the full-range `RANGE_TOO_LARGE` pre-reject; bound by
  the window only. Writes keep the full cap (`_validate_target_cells` unchanged).
- `ranges_service.read_values/read_formulas`: accept optional `cursor_version`; resolve record,
  compare to current version → `STALE_CURSOR` on mismatch; emit `PageInfo`.
- `tables_adapter.read_rows`: raise `_MAX_RESPONSE_CELLS` to `LargeDataConfig.max_table_cells`;
  add version binding + `PageInfo` in `TablesService.read_rows`.
- `api/range_tools.py` + `api/table_tools.py`: add `cursor_version` param + surface `page`.

### WP-C — Pivot result reader (Priority 1)
- New adapter method `read_pivot_result(wb, pivot_table, offset, limit)`:
  - Resolve the `PivotTable` by name (reuse `_find_pivot_table`).
  - Read `TableRange1` (full rendered range incl. headers/totals) and `DataBodyRange` addresses.
  - Page over `TableRange1` rows with the **shared** bounded range reader (not a parallel impl):
    build the sub-range and reuse `_normalize_grid`/`_classify_grid`.
  - Capture structural fingerprint (`table_range1`, `refresh_date`) + `version`.
  - Works for **any** source type that exposes a rendered range (incl. OLAP/Data Model) because it
    reads the *displayed output cells*, not the external data — this is the safe OLAP read path.
- `PivotTablesService.read_pivot_result(...)` + `pivot_read` new action `result-read`
  (`pivot_table`, `offset`, `limit`, `cursor_version`).
- DTO `PivotResultPage` (headers + typed values + `PageInfo` + `source` + fingerprint).

### WP-D — OLAP / Data Model capability matrix (Priority 1)
- Research-proven safe reads implemented: **displayed result** (WP-C, safe for all source types),
  **metadata already read** (`SourceType`, `OLAP`, `RefreshDate`, `TableRange*`).
- Member/CubeField enumeration stays fail-closed unless proven safe under single-STA (research gate).
- Capability matrix documented in `large_data_pivot_capability_matrix.md`; unsupported ops return
  structured `supported=false` outcomes, never disabling safe paths.

### WP-E — Streaming OOXML reader (Priority 1, non-COM)
- `adapters/ooxml/zip_safety.py`: bounded `ZipFile` open — member count, per-member ratio,
  total uncompressed size, path-traversal rejection, no `extractall`.
- `adapters/ooxml/streaming_reader.py`: `iterparse`-based worksheet row/cell extraction with
  shared-strings resolution (streaming `sharedStrings.xml`), inline strings, styles/number formats,
  defined names, tables, hidden rows/cols, formulas + cached values; workbook/rels metadata.
- `adapters/ooxml/pivotcache_reader.py`: streaming `pivotCacheRecords` with shared/inline values,
  disk-backed paging, `saveData`/refresh metadata, explicit staleness.
- `application/ooxml_service.py`: resolves an **immutable snapshot** input (SaveCopyAs) + sha256,
  schedules parsing as an **async JobRegistry** job bounded by `max_file_jobs`, returns paged DTOs.
- `api/ooxml_tools.py`: one new tool `workbook_file_read` (actions: `metadata`, `sheet-rows`,
  `table-rows`, `pivot-cache-records`). Justified: reading the saved package (worksheet XML, shared
  strings, pivot cache records) cannot be done via COM tools without full materialization and must
  not occupy the STA.

### WP-F — Streamed file export (Priority 2)
- `application/data_export_service.py`: streams rows from a range/table/pivot-result/OOXML source to
  a temp file in the artifact dir (CSV or JSONL), `fsync` + `os.replace`, returns
  `{path, rows, bytes, sha256, complete, source_version, format}` — never the data in the response.
- `api/export_tools.py`: new tool `data_export` (justified: no existing tool streams to a file;
  artifact store is bytes-in-memory).
- Path policy enforced; overwrite requires explicit authorization (`.bak`).

### WP-G — Metrics/observability (Priority 2)
- Each large op records: file size, rows/cells, page size, bytes returned, extraction ms, COM ms,
  peak Python memory (best-effort `tracemalloc` around the op), temp bytes, job state. No cell
  contents logged.

---

## 3. Why this satisfies the acceptance criteria (traceability)

| Acceptance | WP |
|---|---|
| Open/metadata bounded; large sheets/tables paged; no silent truncation | A, B |
| Paging consistent across a stable version | A, B, C |
| Pivot results paged; metadata without item enumeration; bounded members | C (existing B/D) |
| OLAP/Data Model supported reads where proven safe; unsupported explicit | C, D |
| `.xlsx`/`.xlsm` streaming, bounded memory, package untouched, corrupt archives fail safely | E |
| Large exports as artifacts, not giant responses | F |
| Pages from same verified version; originals unchanged; no 2nd Excel/STA | A,B,E,F + existing |

## 4. Rejected / deferred (with reason)

- **Second Excel instance / STA worker** for OOXML — rejected (constraint); non-COM path used instead.
- **Weakening original protection / auto-replay** — rejected (constraint).
- **Generic mega `call_tool` / toolsets / lazy loading** — rejected (constraint).
- **Parquet export** — deferred (dependency growth without ecosystem justification); CSV/JSONL first.
- **General import framework / arbitrary code execution** — deferred/rejected (no justified use case;
  security). A narrow CSV import may be added only if a concrete need is demonstrated.
- **Blind OLAP member enumeration** — deferred until a single-STA-safe path is proven.

## 5. Auditor-driven revisions (independent review: PROCEED WITH CHANGES)

Accepted into the design (see `.ai/large-data/large_data_auditor_review.md`):

1. **HIGH — OOXML parse must not block the event loop.** The non-COM parse runs via
   `asyncio.to_thread` (or a bounded executor), **never** inline in the `JobRegistry` coroutine;
   cooperative cancel is checked between chunks. A loop-responsiveness test is added (WP-E).
2. **HIGH — version capture is atomic with the read (no TOCTOU).** `get_version` is captured **inside
   the submitted STA `fn`** (ranges/tables/pivot), not outside the FIFO, so a mutation queued after
   capture cannot mislabel a single page. `PageInfo.version` reflects the version at read time (WP-B/C).
3. **MED — `bump_version` is wired at every mutating choke point.** All `_submit_mutation` helpers
   (ranges/tables/pivots/sheets/names/shapes/controls + save) call `bump_version` on success; a
   per-service test asserts the increment; snapshot restore documents its identity/version change (WP-B).
4. **MED — streamed export reuses an extended ArtifactStore.** Add a streaming artifact writer and
   `.csv`/`.jsonl` artifact types (not bytes-in-memory, not `.zip`-only) (WP-F).
5. **MED — new page DTOs enforce caps.** `PivotResultPage` and OOXML page DTOs carry explicit
   validators bounding cells/bytes; the async job-result path (which bypasses the worker `_ensure_dto`)
   validates the DTO before returning (WP-C/E).

Also accepted: deterministic temp-file cleanup ownership on `interrupted_unknown`; ZIP ratio floor
documented (1000, configurable, vs reference 500); snapshot-GC-vs-job-lifetime note (a job holds its
immutable input until completion).