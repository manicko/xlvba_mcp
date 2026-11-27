# PivotTable / PivotCache Capability Matrix

Per-source-type capabilities for `xlvba_mcp`. "Supported" = implemented and safe on our single-STA
worker; "Unsupported (fail-closed)" = returns a structured `unsupported_source` outcome and never
touches Excel; "Planned (WP)" = added by the large-data work packages. Member/CubeField enumeration
for OLAP/Data Model stays fail-closed this pass (hang evidence is inferred, not executed on our
runtime — see reference review).

Source types (Excel `PivotCache.SourceType`): `DATABASE` (worksheet/Table), `EXTERNAL` (external
connection), `OLAP` (OLAP/SSAS), `DATA_MODEL` (embedded Data Model), `UNKNOWN`.

| Capability | DATABASE | EXTERNAL | OLAP | DATA_MODEL | UNKNOWN |
|---|---|---|---|---|---|
| Identify PivotTable (name, sheet, index) | Supported | Supported | Supported | Supported | Supported |
| Read source metadata (SourceType, OLAP, RecordCount, RefreshDate, shared-cache N:1) | Supported | Supported¹ | Supported¹ | Supported¹ | Supported¹ |
| Read dimensions (TableRange1/2, DataBodyRange addresses) | Supported | Supported | Supported | Supported | Supported |
| Read field definitions (orientation/position/aggregation) | Supported | Unsupported (fail-closed) | Unsupported (fail-closed) | Unsupported (fail-closed) | Unsupported (fail-closed) |
| Read selected filters (manual visible/hidden) | Supported | Unsupported | Unsupported | Unsupported | Unsupported |
| Enumerate members (bounded) | Supported (≤10k, truncated flag) | Unsupported | Unsupported² | Unsupported² | Unsupported |
| **Read displayed results (paged)** | Planned WP-C | Planned WP-C | Planned WP-C³ | Planned WP-C³ | Planned WP-C |
| Read cached source records (paged) | Planned WP-E (OOXML `pivotCacheRecords`) | Unsupported (no local cache) | Unsupported (no local records) | Unsupported (no local records) | Unsupported |
| Refresh (table/cache) | Supported | Unsupported | Unsupported | Unsupported | Unsupported |
| Field config (layout/position/aggregation/filter) | Supported | Unsupported | Unsupported | Unsupported | Unsupported |

¹ Metadata is read from the cache object only (`SourceType`/`OLAP`/`RecordCount`/`RefreshDate`); no
external resolve is performed. `source_data` text is unavailable for external/OLAP.
² Proven scalar/array-safe in the Sbroenne reference (CubeFields/PivotFields/VisibleItemsList/
CurrentPageName), but **not executed** on our single-STA worker; kept fail-closed pending live
validation. Tracked as a future safe-read addition, not a permanent limitation.
³ The displayed-result reader reads the **rendered output cells** (`TableRange1`/`DataBodyRange`) via
the shared bounded range reader — it does **not** resolve the external/OLAP source, so it is safe for
OLAP/Data Model. This is the one safe OLAP/Data Model read path.

## Key distinctions (never conflate)
- **Displayed results** = the values Excel renders in the PivotTable range now (live workbook state).
- **PivotCache records** = the saved source rows in the package (`saveData`); may be stale vs source.
- **Live external source** = the database/OLAP cube; never read directly by the record path.

## Safety notes
- The source-type gate runs FIRST for field/member/filter/config/refresh; unsupported sources get
  identity-only metadata and never enumerate items (prevents single-STA hangs).
- A failure reading one optional metadata property never disables the safe paths (displayed result,
  identity, dimensions).
- Unsupported operations return `supported=false` + `reason` + the capability, never a silent empty.