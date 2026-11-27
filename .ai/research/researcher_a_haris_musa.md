# Researcher A — Large-Data Handling in `haris-musa/excel-mcp-server`

**Repository:** https://github.com/haris-musa/excel-mcp-server
**Commit checked:** `aa9c33d65d3967dd31723d165aa0526ccc952feb` ("chore: release 2.3.7 (#286)")
**Commit date:** 2026-11-26T09:23:11Z
**Default branch:** `main`
**Method:** Read actual source via `raw.githubusercontent.com` at the pinned commit. Every claim below is tagged **[VERIFIED IN CODE]** (I read the function body) or **[DOCUMENTED]** (README/docs claim, not independently traced to enforcement) or **[NOT FOUND]** (searched, absent).

---

## 0. Architectural distinction (read this first)

Their server is **pure file manipulation with `openpyxl`** — no Excel application, no COM. Almost every large-data technique exists *because* openpyxl builds a full in-memory object graph per cell and re-parses whole sheets. Core abstractions:

- `Workspace` (`src/excel_mcp/workspace.py`) opens/saves `.xlsx` via `openpyxl.load_workbook` and `save_atomically`. **[VERIFIED]**
- `CellBudget` / `TextBudget` cap the number of cells/text openpyxl is allowed to materialize. **[VERIFIED]**
- `LazyCells` / `LazyEdit` keep a sheet's raw XML and parse only 128-row blocks on demand. **[VERIFIED]**
- Streaming reads use openpyxl `read_only=True` `ReadOnlyWorksheet`. **[VERIFIED]**

Our server (`xlvba_mcp`) drives **one owned hidden Excel COM instance through a single STA worker**. The OS/Excel process owns the spreadsheet model; we never parse XLSX. This inverts almost every assumption: memory is bounded by Excel's own process, not by Python; "streaming" means COM range reads, not XML iterparse; "paging" means `Range.Value2` over a sub-range, not slicing XML. Verdicts below reflect that.

---

## 1. Streaming worksheet reads & paged responses

**What they do [VERIFIED]:**
- `read_range` tool (`src/excel_mcp/server/data_tools.py:read_range`) takes `max_cells` (default **2000**, `Field(ge=1)`), clamps with `max_read = min(max_cells, limits.max_read_cells)` where `max_read_cells = 10_000` (`config.py:Limits`).
- Two modes:
  - `mode="values"` → `operations/calculated.py:read_calculated` (lazy workbook, computes formulas).
  - `mode="formulas"` → `workspace.stream(path)` → openpyxl `read_only` `ReadOnlyWorksheet` → `operations/cells.py:read_range`.
- `operations/cells.py:read_range` computes `last_row = min(target.max_row, target.min_row + max_cells // target.cols - 1)`, iterates only that window with `sheet.iter_rows(...)`, and returns a `RangeData` Pydantic model with `next_range` set to the remainder (`CellRange(last_row+1, ...)`) or `None`. The client re-calls with `next_range` to page. **[VERIFIED]**
- A second cap: page text is bounded by `max_cells * CHARS_PER_CELL` (256 chars/cell, `text_budget.py`); if a single row exceeds it the call raises `LimitExceededError`; otherwise it truncates and sets `next_range`. **[VERIFIED]**
- `find_cells` (`operations/cells.py:find_cells`) is a streaming full scan over `ReadOnlyWorksheet` with a `TextBudget(max_loaded_cells, ...)` and a result page of `max_read_cells * CHARS_PER_CELL`, returning `truncated=True` when over. **[VERIFIED]**

**Limits summary:** page size default 2000, hard ceiling 10000 cells/read; 256 chars/cell text budget; `next_range` cursor for continuation.

**Verdict: ADAPT.** The *paging contract* (small default page + explicit `next_range`/continuation cursor + per-page char budget) is directly applicable to COM: read a sub-range via `Range.Value2`, return a cursor (next row), cap chars. Reject the openpyxl `ReadOnlyWorksheet` mechanics.

---

## 2. Partial processing of large worksheet XML (iterparse / streaming)

**What they do [VERIFIED]:**
- They avoid a DOM parse of big sheets. Three distinct streaming paths:
  1. **Cell counting pre-pass** with `xml.parsers.expat` (`cell_budget.py:spend_sheet`) — reads worksheet XML in 1 MiB chunks (`_CHUNK = 1<<20`), fires `StartElementHandler` for `<c>`, `<row>`, `<mergeCell>`, `<hyperlink>` and *spends* budget without building objects. **[VERIFIED]**
  2. **Lazy block parsing** (`lazy_workbook.py:LazyCells`) — sheet held as raw `bytes` + row spans (`lazy_source.py:SheetSource`); `load(row)` parses only the 128-row `BLOCK_ROWS` block via regex fast-path (`lazy_rows.py:read`) and falls back to openpyxl `WorkSheetParser` for "unusual" rows. **[VERIFIED]**
  3. **Pivot records iterparse** (`package/lazy_records.py:_records`) — uses `openpyxl.xml.functions.iterparse` over `BytesIO`, yields each `<r>` then `root.clear()` to free memory. **[VERIFIED]**
- Shared-formula text expansion is pre-counted in chunks that end on `</c>` boundaries (`shared_text.py:_pieces`, `spend_sheet`) so no formula is split. **[VERIFIED]**

**Verdict: REJECT (mechanics) / ADOPT (principle).** XML iterparse/expat is meaningless when Excel owns the data — there is no XML for us to stream. The transferable principle is "pre-flight a cheap count/size check before doing the expensive materialization," which for COM maps to reading `UsedRange`/cell counts cheaply before a bulk `Value2` read.

---

## 3. Separate limits — exact defaults & enforcement

All defaults from `src/excel_mcp/config.py:Limits` **[VERIFIED]**, CLI flags in `cli.py:_parser` **[VERIFIED]**:

| Limit (config field) | Default | CLI flag / env | Enforced at |
|---|---|---|---|
| `max_file_bytes` | **100 MiB** (`100*1024*1024`) | `--max-file-mb` / `EXCEL_MCP_MAX_FILE_MB` | `workspace.py:resolve_existing` (size check) + `_check_expansion` |
| `max_read_cells` | **10 000** | (not a CLI flag; fixed) | `data_tools.py:read_range` via `min(max_cells, max_read_cells)` |
| `max_cells` | **100 000** | `--max-cells` / `EXCEL_MCP_MAX_CELLS` | `operations/cells.py:write_range/clear_range`, `parse_range(...).within(max_cells)`, `pivot.py` |
| `max_pivot_source_cells` | **3 000 000** | `--max-pivot-cells` / `EXCEL_MCP_MAX_PIVOT_CELLS` | `operations/pivot_source.py:read_source` (`area.size > max_cells` → `LimitExceededError`), called from `pivot.py` |
| `max_copy_cells` | **1 000 000** | (fixed) | copy/paste paths |
| `max_loaded_cells` | **3 000 000** | `--max-loaded-cells` / `EXCEL_MCP_MAX_LOADED_CELLS` | `cell_budget.py:CellBudget.spend` (per workbook, shared across sheets) |
| `max_unpack_factor` | **10** | (derived) | `zip_limits.check_expansion` |
| `max_compression_ratio` | **500** | (fixed) | `zip_limits.check_ratio` |

Notes:
- There is **no `max_pivot_cells` field**; the name is `max_pivot_source_cells` (CLI flag is `--max-pivot-cells`). **[VERIFIED]**
- `max_unpacked_bytes` is a computed property = `max_file_bytes * max_unpack_factor` = **1 000 MiB** default. **[VERIFIED]**
- `max_loaded_cells` is the "cells+rows openpyxl may create in memory" budget; `CellBudget.spend` raises when exceeded, and `spend_sheet` pre-counts before parse. Rows count as 1 cell; merged/hyperlink ranges cost `range.size` (hyperlinks ×3, `_LINK_COST`). **[VERIFIED]**

**Verdict: ADAPT.** The *taxonomy* (file-size / per-call-cell / per-workbook-loaded-cell / pivot-source-cell / unpacked-byte) is worth copying conceptually. For COM, "loaded cells" is not a Python concern — Excel's memory is the process's — so `max_loaded_cells`/`max_cells` become **per-call range-size caps** (cells read/written per `Range.Value2`), and `max_file_bytes`/unpacked limits become largely irrelevant (we open files in Excel, not parse zips).

---

## 4. PivotCache XML generation without an in-memory object per value

**What they do [VERIFIED]:**
- `operations/pivot_records.py:PivotRecords` subclasses openpyxl `RecordList` but stores rows as plain Python tuples (`rows: Sequence[Row]`), NOT one openpyxl object per item. Its `_write` opens the zip member in streaming write mode (`archive.open(path, "w", force_zip64=True)`) and calls `_write_xml`.
- `_write_xml` writes the `<pivotCacheRecords>` envelope, then batches rows into ~1 MiB (`BATCH_CHARACTERS = 1<<20`) string chunks and writes each chunk to the stream — never materializing the full XML. **[VERIFIED]**
- `_items` yields per-item XML strings with manual escaping (`ESCAPES` translate table) and a `match` on type: `<m/>` missing, `<x v=>` shared index, `<s>`, `<b>`, `<e>`, `<d>`, `<n>`. **[VERIFIED]**
- `operations/pivot_cache.py:build_cache` builds the `CacheDefinition` (fields, shared items flags) but attaches `cache.records = _records(setup)` (the streaming `PivotRecords`). Shared items are deduplicated via `SharedIndex` so repeated strings become index references. **[VERIFIED]**
- On load, `package/lazy_records.py:LazyRecords` keeps the records part as raw `bytes` and only iterparses a single column when a slicer needs shared items (`column`, `share`). `install()` monkeypatches `WorkbookParser.pivot_caches`. **[VERIFIED]**

**Verdict: ADAPT (conceptually) / mostly N/A.** In COM, `PivotCache` is created/refreshed by Excel itself (`PivotCaches.Create` / `PivotTable.PivotCache.Refresh`); we never generate `pivotCacheRecords.xml`. The transferable idea is *stream/batch serialization with dedup (shared-item indexing) instead of per-value objects* — relevant to us only if we ever emit XML/OPC parts ourselves (we generally don't). For our model: **REJECT** the XML generation; **ADOPT** the "batch + dedup" mindset if any bulk payload is ever serialized.

---

## 5. ZIP archive safety / decompression-bomb limits

**What they do [VERIFIED] (`src/excel_mcp/zip_limits.py`):**
- `check_expansion(infos, limits)`:
  - Per-member ratio: `check_ratio(info.file_size, info.compress_size, limits)` — raises if `size > max(_RATIO_FLOOR, compressed * max_compression_ratio)`. `_RATIO_FLOOR = 1 MiB` exempts tiny parts. Default ratio **500:1**.
  - Total: if `sum(info.file_size) > limits.max_unpacked_bytes` (= `max_file_bytes * max_unpack_factor` = 100 MiB × 10 = **1000 MiB**) → `too_large`.
- Called from `workspace.py:_check_expansion` (on `resolve_existing`) and `_check_zip`/`_load` paths; `Reader.read` (`package/reader.py`) also enforces a per-part `budget` (non-spreadsheet parts). **[VERIFIED]**
- **Path traversal:** handled in `package/opc.py:resolve` — normalizes relationship targets and raises `WorkbookError` if a target escapes the package (`../`). **[VERIFIED]** (Note: this is OPC relationship resolution, not zip member-name sanitization; they rely on openpyxl + their own `resolve`.)
- **Member count:** **[NOT FOUND]** — no explicit cap on number of zip members anywhere in `zip_limits.py` or `package/reader.py`. Only ratio + total-unpacked + per-part budget.
- Writes are atomic via `tempfile.mkstemp` + `replace` (`workspace.py:write_atomically`) to avoid symlink/clobber issues. **[VERIFIED]**

**Verdict: REJECT.** All of this protects against hostile XLSX because *they* parse untrusted zips in Python. In the COM model Excel opens the file; we never run `zipfile`/`iterparse` on workbook bytes, so ratio/bomb/member checks are not our attack surface. (If `xlvba_mcp` ever accepts uploaded `.xlsm` bytes and writes them to disk for Excel to open, a *file-size* cap is still worth keeping, but not ratio bombs.)

---

## 6. How large sheet operations stay bounded

**What they do [VERIFIED]:**
- **Edits** (`workspace.py:edit`) take a `CellBudget(max_loaded_cells)` and route through `lazy_edit.load` when `lazy_edit.has_large_sheet(path)` (any worksheet part ≥ `LAZY_SHEET_BYTES = 1_000_000` bytes, and no array/dataTable/`cm`/`vm` markers). `_split` refuses sheets with cross-cell features. **[VERIFIED]**
- `LazyCells` parses 128-row blocks on demand; operations that touch all cells (`__iter__`, `len`, `keys`) force `load_all`, but row-scoped ops (`stored_cells_in`, `in_row`, `stored_in`) parse only the requested rows. **[VERIFIED]**
- Row insert/delete on unparsed sheets is done as **byte-level row renumbering + formula reference rewriting** on raw XML (`lazy_lines.py:move_lines`, `_row_mover`, `_column_mover`) without parsing cells. **[VERIFIED]**
- Every write/clear/copy/sort/transform path passes `limits.max_cells` and calls `parse_range(...).within(max_cells)` or checks `count > max_cells` → `LimitExceededError` (`operations/cells.py`). **[VERIFIED]**
- Save re-checks total text size (`workspace.py:_check_text_size`) against `max_unpacked_bytes` before writing, because shared text expands per-cell on save. **[VERIFIED]**

**Verdict: ADAPT (the discipline) / REJECT (the mechanism).** The transferable discipline: *every mutating tool validates an explicit cell budget up front and refuses over-budget work with a clear, actionable error.* For COM, the equivalent is a per-call `Range` cell-count guard before touching `Value2`/formatting, plus relying on Excel to hold the data. Reject the XML byte-surgery entirely.

---

## 7. Other large-data techniques

- **Shared-strings handling [VERIFIED]:** they do NOT stream sharedStrings.xml itself; openpyxl loads `reader.shared_strings` into a list, and `lazy_rows._item` indexes it (`reading.strings[int(value)]`). The *text budget* (`TextBudget`) is what bounds the cost of many cells pointing at one long string (`text_budget.py`, `cell_budget.py` `_FORMULA_CHARS_PER_CELL=128`). So "shared strings streaming" is **NOT** implemented — they bound it indirectly via character budgets. **[VERIFIED — claim of streaming shared strings would be false]**
- **Chunked writes [VERIFIED]:** pivot records (1 MiB batches, §4); worksheet XML assembled from row spans (`lazy_source._assembled`); atomic file write via temp file.
- **`max_copy_cells` = 1 000 000** separate ceiling for copy/paste. **[VERIFIED]**
- **Formula-text budget [VERIFIED]:** shared-formula expansion pre-counted (`shared_text.spend_sheet`) and charged at load (`LazyCells.__init__`), `_FORMULA_CHARS_PER_CELL = 128`.
- **Calc on large sheets [VERIFIED]:** `workspace.read_lazy` loads two lazy workbooks (data_only + formulas) so only rows a formula needs are parsed — relevant to their formula engine, not to COM.
- **No compression of in-memory workbook; no mmap; no external sort.** **[NOT FOUND]**

**Verdict:** Character-budget-as-a-proxy-for-memory (ADOPT as a guard concept), chunked/atomic write (ADOPT discipline), shared-string streaming (REJECT — not actually implemented; and N/A for COM).

---

## Adopt / Adapt / Reject — consolidated

| Technique | Verdict | One-line rationale |
|---|---|---|
| Paged reads with small default + `next_range` cursor + per-page char cap (§1) | **ADAPT** | Same contract works over COM `Range.Value2` sub-range reads; drop openpyxl. |
| XML iterparse / expat pre-count (§2) | **REJECT** | No XML in COM model; keep only the "cheap pre-count before materialize" principle. |
| Limit taxonomy: file / per-call-cell / loaded-cell / pivot-source / unpacked (§3) | **ADAPT** | Keep per-call cell caps; loaded-cell & unpacked/file-zip limits are openpyxl-specific. |
| PivotCache streaming XML + shared-item dedup (§4) | **REJECT** (XML) / ADOPT (batch+dedup mindset) | Excel builds PivotCaches in COM; we never emit records XML. |
| Zip ratio/bomb/member/path safety (§5) | **REJECT** | We don't parse zips; Excel opens files. Keep only a file-size cap if accepting uploads. |
| Up-front per-call cell budget + clear refusal on large edits (§6) | **ADAPT** | Enforce a `Range` cell-count guard before bulk COM writes. |
| Char-budget proxy for memory; chunked/atomic writes (§7) | **ADOPT** (as guards) | Cheap, model-agnostic guardrails worth copying. |
| Shared-strings streaming | **REJECT** | Not implemented by them; N/A for COM. |

**Net for `xlvba_mcp`:** Adopt the *ergonomics and guard discipline* (paging cursors, per-call cell/text budgets, explicit actionable `LimitExceeded`-style errors, atomic/locked writes). Reject everything that exists to compensate for openpyxl's in-memory object graph and Python-side zip parsing — that entire layer is invisible to a single-COM-instance server.

---

## Source citations (repo-relative paths + functions, commit `aa9c33d65d3967dd31723d165aa0526ccc952feb`)

- `src/excel_mcp/config.py` — `Limits` (all default values), `Settings`.
- `src/excel_mcp/cli.py` — `_parser`, `_settings` (CLI flags → `Limits`, incl. `--max-pivot-cells` → `max_pivot_source_cells`).
- `src/excel_mcp/zip_limits.py` — `check_expansion`, `check_ratio`, `too_large`, `_RATIO_FLOOR`.
- `src/excel_mcp/cell_budget.py` — `CellBudget.spend`, `spend_sheet` (expat), `check_workbook`, `_LINK_COST`, `_FORMULA_CHARS_PER_CELL`.
- `src/excel_mcp/text_budget.py` — `TextBudget`, `CHARS_PER_CELL`.
- `src/excel_mcp/shared_text.py` — `spend_sheet`, `_pieces`, `master_lengths`, `users`.
- `src/excel_mcp/lazy_workbook.py` — `LazyCells` (`load`, `_parse`, `stored_rows`, `count`), `BLOCK_ROWS`, `_SheetArchive`, `open_lazily`, `load_lazy`.
- `src/excel_mcp/lazy_source.py` — `SheetSource`, `split_sheet`.
- `src/excel_mcp/lazy_rows.py` — `read`, `render`, `_item` (regex row/cell fast path).
- `src/excel_mcp/lazy_shared.py` — `SharedMaster`, `index`, `Translators`, `settle`, `restore`.
- `src/excel_mcp/lazy_lines.py` — `move_lines`, `_row_mover`, `_column_mover`, `rewrite_formulas`.
- `src/excel_mcp/lazy_edit.py` — `has_large_sheet`, `_split`, `LAZY_SHEET_BYTES`, `RowDimensions`.
- `src/excel_mcp/workspace.py` — `Workspace.resolve_existing`, `_check_expansion`, `stream`, `edit`, `read_lazy`, `save_atomically`, `_check_text_size`, `write_atomically`.
- `src/excel_mcp/operations/cells.py` — `read_range` (paging/`next_range`), `find_cells`, `write_range`, `clear_range`, `streamed_cells`.
- `src/excel_mcp/server/data_tools.py` — `read_range` tool (`max_cells` default 2000, `min(..., max_read_cells)`), `find_cells`.
- `src/excel_mcp/operations/pivot_records.py` — `PivotRecords`, `_write`, `_write_xml`, `_items`, `BATCH_CHARACTERS`, `SharedIndex`.
- `src/excel_mcp/operations/pivot_cache.py` — `build_cache`, `_records`, `_shared_items`.
- `src/excel_mcp/operations/pivot_source.py` — `read_source` (`area.size > max_cells` enforcement), `record_count`.
- `src/excel_mcp/package/lazy_records.py` — `LazyRecords`, `install`, `_records` (iterparse), `column`, `share`.
- `src/excel_mcp/package/reader.py` — `Reader.read` (per-part budget).
- `src/excel_mcp/package/opc.py` — `resolve` (path-traversal guard), `parse_rels`.
- `README.md` / `docs/index.html` — documented "Large files" claims and limit table (used only as **[DOCUMENTED]** cross-check; all numbers above confirmed in code).