# Large-Data Reference Implementations

Two reference servers were inspected at pinned commits. For each technique: **ADOPT** (use as-is),
**ADAPT** (usable after change for our single-COM-instance model), **REJECT** (incompatible with our
constraints). "Documented" vs "verified in code" are separated; anything not verified in code is
marked as such.

Our constraints (the filter): one MCP server, **one** owned hidden Excel instance, **one** STA COM
worker, one FIFO queue, working-copy protection, snapshot/approval gates, agent-directed recovery,
no second Excel process/STA, consolidated tool surface.

---

## Reference A — Haris Musa `excel-mcp-server`

- **Commit checked:** `aa9c33d65d3967dd31723d165aa0526ccc952feb` (release 2.3.7, 2026-11-26, `main`).
- **Architecture:** `openpyxl` + direct XLSX/ZIP manipulation, **no live Excel COM**. Its large-data
  machinery largely compensates for Python-side zip/XML parsing — a layer our COM server does not need.

| Technique | Source (file:function) | Verdict | Rationale |
|---|---|---|---|
| Paged range reads + `next_range` cursor, per-cell `TextBudget` (256 chars) | `server/data_tools.py:read_range`, `operations/cells.py:read_range`, `cell_budget.py` | **ADAPT** | cursor ergonomics + per-cell text cap fit our `PageInfo`; wire into our COM range reader, not their file reader |
| Streaming worksheet XML (expat, 1 MiB chunks; 128-row lazy blocks; iterparse+`root.clear()`) | `cell_budget.py:spend_sheet`, `lazy_workbook.py:LazyCells`, `package/lazy_records.py:_records` | **ADAPT (OOXML track only)** | needed only for our non-COM OOXML reader; not for COM reads |
| Limit taxonomy: `max_file_bytes`=100 MiB, `max_read_cells`=10 000, `max_cells`=100 000, `max_pivot_source_cells`/`max_loaded_cells`=3 000 000, `max_unpack_factor`=10, `max_compression_ratio`=500 | `config.py:Limits`, `workspace.py`, `pivot_source.py:read_source` | **ADAPT** | directly informs our `LargeDataConfig` defaults (we chose 100k range / 50k response / 100 MiB file / ratio 1000) |
| PivotCache records streamed ~1 MiB batches with `SharedIndex` dedup; raw-bytes records | `operations/pivot_records.py:PivotRecords._write_xml`, `package/lazy_records.py` | **ADAPT (OOXML track)** | model for our streaming `pivotCacheRecords` reader |
| ZIP safety: ratio 500:1 (1 MiB floor) + total-unpacked 1000 MiB; path-traversal in `opc.py:resolve`; **no member-count limit** | `zip_limits.py`, `opc.py` | **ADAPT** | we add a member-count limit (they lack one) |
| Bounded edits for ≥1 MB sheets (byte-level row moves) | `lazy_edit.py:has_large_sheet`, `lazy_lines.py:move_lines` | **REJECT** | we use COM bulk `.Value2` writes, not Python file edits |
| Shared strings streamed | — | **REJECT the claim** | **not verified in code**; they bound via char budgets only. Our OOXML reader will stream sharedStrings itself |

**Net (A):** adopt their *guard discipline* (per-call cell/text budgets, limit taxonomy, atomic writes,
cursor ergonomics); reject the entire openpyxl-compensation layer (XML iterparse for COM, zip-bomb
checks for COM, PivotCache XML *generation*, shared-string streaming claims).

---

## Reference B — Sbroenne `mcp-server-excel`

- **Commit checked:** `e42059685810775ee5f8277a56972c44546f40b0` (2026-11-24, `main` HEAD).
- **Architecture:** multi-process — `ExcelController` spawns up to **32 private worker processes**,
  each one STA + one hidden Excel, over named-pipe JSON-RPC. The **opposite** of our single-instance
  model; the worker fabric is rejected wholesale, but the per-operation discipline is reusable.

| Technique | Source (file:function) | Verdict | Rationale |
|---|---|---|---|
| Metadata vs data split (config/snapshot read vs `TableRange2.Value2` data read) | `PivotTableCommands.Lifecycle.cs:Read` vs `PivotTableCommands.Analysis.cs:GetData` | **ADOPT split; ADAPT GetData** | our WP-C mirrors this; we add a row cap (their GetData is uncapped) |
| Bounded member enumeration (`DefaultMaxSelectedMembers=100`, max 1000) + `MembersTruncated`/`SnapshotComplete` flags | `PivotTableSnapshotHelpers.cs`, `SemanticSnapshot.cs:BoundedMemberNames` | **ADOPT DIRECTLY** | matches our `_MAX_PIVOT_ITEMS` + completeness flags |
| OLAP property reads (scalars/arrays) | `PivotTableHelpers.cs:HasOlapCache/IsOlapPivotTable`, `SemanticSnapshot.cs:ReadOlapFieldSelection/ReadOlapMemberState` | **ADOPT (deferred in our pass)** | proven scalar/array reads; we keep the gate until live-validated (see §OLAP) |
| `pivottable_calc` get-data + calc fields/items/members; per-item cancellation, no count cap | `IPivotTableCalcCommands`, `BatchCommand.cs` | **ADAPT** | add a count cap; not a priority for our pass |
| Outcome-unknown / no-repeat on worker exit; reveal-window | `ExcelController.cs:CompleteWorkerRequestAsync/PollAsync`, `ExcelWorkerClient.cs:CallAsync`, `ExcelWorkerHost.cs` | **ADOPT concept; REJECT fabric** | we already have agent-directed recovery + `interrupted_unknown`; reinforce "do not repeat" |
| CLI + MCP share one Core | `ExcelBatchTool.cs:ExecuteBatchAsync`, `ServiceRegistry` | **ADOPT DIRECTLY** | our api/* + services already share Core |
| No PivotCache record iteration; scalar `RecordCount` + bounded ADO schema rowsets (≤250 page, ≤1000 scanned, 2-min CTS) | `ExternalOlapSchema.cs:OpenSchemaRowset` | **ADOPT DIRECTLY** | model for bounded external metadata; we do not iterate records via COM |
| Same-thread COM cancellation via `OleMessageFilter` (`PENDINGMSG_CANCELCALL`) without a 2nd process | `OleMessageFilter.cs:MessagePending`, `SetPendingCancellationToken` | **ADAPT (future)** | pairs with our agent-directed recovery; not required for this pass |

### OLAP / Data Model property safety (single-STA)
**Safe (scalar/array, no external resolve):** `PivotCache.OLAP/SourceType/RecordCount/RefreshDate`,
`CubeFields.Count`, `CubeField.Name/Caption/Orientation/Position/CubeFieldType/PivotFields/
EnableMultiplePageItems/AllItemsVisible`, `PivotField.VisibleItemsList/CurrentPageName/Orientation/
Name/SourceName/Position/Calculation`, `TableRange2.Address`, Model tables/measures/relationships
counts/names. **Risky (bound + cancel):** `RefreshTable()`, `PivotCache.ADOConnection`+schema rowsets,
full `TableRange2.Value2`, DAX `Execute`. **No OLAP/Data Model property requires a second Excel
process.** Confidence: MEDIUM — the repo wraps every read in `TryReadSnapshotProperty` defensively; it
does **not** assert specific hangs, and runtime hang behaviour was **not executed**.

**Decision for our pass:** the *displayed result* read (WP-C) is safe for all source types and is
implemented. Member/CubeField enumeration stays **fail-closed** this pass because the hang evidence is
inferred, not executed on our single-STA worker; the capability matrix records these as
"safe-in-reference, deferred pending live validation". This honors "do not weaken the gate until a
safe alternative has been demonstrated" on *our* runtime.

---

## Adopt / Adapt / Reject — summary

- **ADOPT DIRECTLY:** bounded-member flags + completeness; metadata/data split; CLI+MCP shared Core;
  bounded external schema rowsets (≤250/≤1000/2-min); outcome-unknown/no-repeat.
- **ADAPT:** paging-cursor ergonomics + per-cell text budget; limit taxonomy/defaults; streaming
  worksheet + PivotCache records (OOXML track only); `pivottable_calc` count cap; same-thread COM
  cancel (future).
- **REJECT:** 32-worker/pipe fabric; openpyxl-compensation XML/zip layers for COM; PivotCache XML
  generation; Python byte-level sheet edits; the unverified "shared strings streamed" claim.

## Source citations
Full per-file citations and verbatim notes: `.ai/research/researcher_a_haris_musa.md`,
`.ai/research/researcher_b_sbroenne.md` (each records the checked commit SHA, repo-relative paths and
function names, and documented-vs-verified separation).