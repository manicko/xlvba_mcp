# Large-Data Execution Contract (WP1)

Normative contract for every large-data operation in `xlvba_mcp`. Additive: existing DTO fields
are preserved; new fields are optional and defaulted so current callers keep working. This document
defines (1) the unified configurable limits, (2) the response/completeness contract, (3) stable
pagination + version binding, and (4) the structured outcome taxonomy.

---

## 1. Unified configurable limits

All limits live in one typed model `LargeDataConfig` (new, added to `ServerConfig`) and are
overridable via `EXCEL_VBA_MCP_LARGE_DATA_*` env vars. Defaults are conservative and documented;
they are **not** raised to make tests pass — capability comes from bounded/streaming processing.

| Limit | Env suffix | Default | Applies to | Notes |
|---|---|---|---|---|
| `max_range_cells` | `MAX_RANGE_CELLS` | 100 000 | per **request window** (read) / per write payload | full-range pre-reject removed for reads (see §3.1); the **write** cap is the module constant `MAX_RANGE_CELLS` (default 100 000), not a separately-wired config value — `max_range_cells` is the read-window default |
| `max_response_cells` | `MAX_RESPONSE_CELLS` | 50 000 | cells returned in one MCP response | hard response cap |
| `max_response_bytes` | `MAX_RESPONSE_BYTES` | 8 MiB | serialized response budget | enforced at the value-serialization boundary (range + OOXML): trailing rows are dropped so the estimated UTF-8 value payload fits, page marked `truncated` |
| `max_table_cells` | `MAX_TABLE_CELLS` | 50 000 | table `read-rows` response | was 2 000; aligned + configurable |
| `max_pivot_result_cells` | `MAX_PIVOT_RESULT_CELLS` | 50 000 | pivot result page | one page of rendered output |
| `max_pivot_members` | `MAX_PIVOT_MEMBERS` | 10 000 | field/member enumeration | existing |
| `max_text_len` | `MAX_TEXT_LEN` | 32 767 | per returned text field | enforced at the value-serialization boundary (range + OOXML): a longer text field is truncated to the cap and the page is marked `truncated` (never silent) |
| `max_file_mb` | `MAX_FILE_MB` | 100 | OOXML input file size | reject larger before open |
| `max_uncompressed_mb` | `MAX_UNCOMPRESSED_MB` | 1024 | total uncompressed OOXML size | ZIP-bomb guard |
| `max_compression_ratio` | `MAX_COMPRESSION_RATIO` | 1000 | per-member ratio | ZIP-bomb guard |
| `max_zip_members` | `MAX_ZIP_MEMBERS` | 10 000 | archive member count | |
| `max_shared_strings` | `MAX_SHARED_STRINGS` | 1 000 000 | sharedStrings count | |
| `max_xml_depth` | `MAX_XML_DEPTH` | 64 | XML nesting depth | parser guard |
| `max_export_rows` | `MAX_EXPORT_ROWS` | 10 000 000 | streamed export rows | |
| `max_file_jobs` | `MAX_FILE_JOBS` | 2 | concurrent non-COM file jobs | bounded, no thread/process pool |

Derivation rules (avoid duplicate settings): the response **byte** budget is independent of the
**cell** budget because text cells vary; everything else is 1:1. No per-tool limit is introduced;
tools reference these.

Enforcement scope of the text/byte bounds: `max_text_len` and `max_response_bytes` are enforced
at the **value-serialization boundary** of the range reader (`RangesAdapter.read_values`, also used
by the PivotTable rendered-result reader) and the OOXML worksheet reader
(`iter_sheet_rows`). The byte budget is a conservative estimate (sum of the UTF-8 length of each
non-empty cell value, not the full JSON envelope); a page that would exceed it drops trailing rows
and is marked `truncated`, and the next page resumes at `offset + returned` so no row is skipped.
Table reads are bounded by `max_table_cells` (cell count); PivotTable results inherit the range
reader's bounds.

## 2. Response / completeness contract

Every paged/large read returns a `PageInfo` block (additive; existing fields unchanged):

```jsonc
"page": {
  "offset": 0,            // 0-based row/item offset of the first returned item
  "limit": 1000,          // requested cap (null/absent = server default)
  "returned": 1000,       // items actually returned
  "total": 250000,        // total available when safely known, else null
  "has_more": true,       // more items exist beyond this page
  "complete": false,      // true iff returned == total (all data delivered)
  "truncated": false,     // true iff a single response was cut by a limit
  "version": 7,           // workbook content version this page belongs to
  "source": "live"        // "live" | "snapshot" | "cache"
}
```

Rules:
- **No empty result is ambiguous.** An empty page always carries `total` (0 when known) and a
  distinct `reason` when the emptiness is not "no data": one of
  `no_data` | `unsupported_source` | `stale_cursor` | `invalid_page` | `read_failed`.
- `complete=true` only when the agent has seen every item (single page with `has_more=false`, or the
  final page). Partial data is never presented as complete.
- The applicable limit is echoed in the rejection context when a request is refused
  (`{"limit": <value>, "cells": <n>, "reason": "..."}`).
- Source identity (`workbook_id`, `version`, and for OOXML the input `sha256`) is always present.

## 3. Stable pagination + version binding

### 3.1 Read paging over arbitrarily large ranges
The full-range `RANGE_TOO_LARGE` pre-reject is **removed for reads**. A read request is bounded by
the **window** (`returned rows × cols ≤ max_response_cells`) not by the full range. The full-range
cap remains for **writes** (unchanged) and as the default per-request window. This lets a 566k-row
sheet be paged with `offset` while each response stays ≤ 50 000 cells.

### 3.2 Workbook content version
- `SessionRecord.version: int` (new), starts at `0` on open.
- `SessionService.bump_version(workbook_id)` increments it; called by every mutating service on a
  **successful** mutation (range write, table write, pivot refresh/config, sheet/name/shape/control
  mutation, save). Reads never bump it.
- A cursor captures `(workbook_id, version, kind, object_key, offset)`.
- On any paged call, if `cursor.version != current version` → raise `STALE_CURSOR` (new error code,
  category `validation`, not retryable) with context `{expected_version, actual_version, kind}`.
  The agent must restart paging from offset 0. This prevents mixing workbook versions across pages.
- `object_key` is the stable object identity (range address for ranges, table name for tables,
  pivot name for pivot results). For pivot results the key also includes a structural fingerprint
  (`table_range1` + `refresh_date`) so a refresh that does not bump version is still caught.

### 3.3 Opaque vs explicit cursor
The offset is small and human-debuggable, so the cursor is **explicit** (`offset` + `version` +
`object_key`) rather than an opaque token — it hides no internal state and is easy to log. (Opaque
tokens are reserved for a future case where continuation state is non-trivial; not needed now.)

## 4. Structured outcome taxonomy (new error code)

`STALE_CURSOR` — a continuation cursor refers to a different workbook version than the current one.
Category `validation`; `retryable=false`; recovery = restart paging. Distinct from `INVALID_RANGE`
(bad address) and `RANGE_TOO_LARGE` (oversized window).

Unsupported OOXML/OLAP operations return a **structured unsupported outcome** (not an exception
where a result DTO is natural), carrying `supported=false` + `reason` + the exact capability, so one
unsafe property never disables the safe read paths (capability matrix per source type).

## 5. Non-COM execution contract (OOXML / export)

- Pure file parsing runs off the STA worker. Large ops are scheduled through the existing
  `JobRegistry` as **async** callables (`job_registry._run_async`), bounded by `max_file_jobs` and
  the limits above; the registry owns the result (client disconnect never drops it).
- Input is always an **immutable snapshot file** (`SaveCopyAs` output, `stored_path` + `sha256`),
  never the live working copy and never the original. Results/cursors bind to that `sha256`.
- Cooperative cancellation via `cancel_requested` (intent-only; parsing checks it between chunks).
- Deterministic temp-file cleanup; a parsing job that is interrupted leaves no partial artifact and
  reports `interrupted_unknown` (never completed, never replayed).
- The reader is read-only; VBA/bin parts are never rewritten.

## 6. What this contract forbids

- Silent truncation (every cut sets `truncated`/`has_more` + `complete=false`).
- Raising a limit to pass a test (limits are config, defaults justified by benchmarks).
- Mixing workbook versions across pages (version-bound cursors).
- Reading the live working copy or original through OOXML (immutable snapshot only).
- Putting a full large dataset in an MCP response (file export returns an artifact reference).