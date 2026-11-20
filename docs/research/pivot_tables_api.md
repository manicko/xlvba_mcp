# PivotTables / PivotCaches / PivotFields API — Decision Record

**Task:** `task_039` (Phase 2 research gate)
**Repo:** `C:\py_exp\xlvba_mcp`
**Status:** Research only — no production code was modified.
**Target build (empirically verified):** Windows 10 LTSC 2019 x64; Excel 16.0 **14332.0** (en-US); Python 3.14.7; pywin32 early + late binding; hidden automation-owned `Excel.Application` via the existing single-STA worker.
**Companion plan:** `.ai/plans/plan_phase2_ranges_cells_pivottables.md` (`task_039`), `.ai/plans/plan_addendum_background_execution.md`.

---

## 1. Scope

This record covers **inventory and management of existing PivotTables** in an open workbook:

- Discover every PivotTable on every worksheet and identify its cache.
- Read PivotTable / PivotCache / PivotField metadata.
- Targeted refresh of local (worksheet/Table-backed) PivotTables and caches.
- Bounded PivotField layout: orientation, position, data-field aggregation, captions, and a small set of safe properties.
- Bounded filters: manual (`PivotItems.Visible`) and a limited `PivotFilters.Add` surface where it works on the target build.
- Explicit unsupported/read-only outcomes for external, OLAP, and Data Model sources.

### Explicitly out of scope (hard boundaries)

- **No PivotTable or PivotCache creation/deletion.** No `PivotTables.Add`, `PivotCaches.Add`, `PivotTable.Delete`, `PivotCache.Delete`.
- No Workbook Connections, QueryTables, Power Query, or Data Model/Power Pivot management.
- No `Workbook.RefreshAll`.
- No Slicers/SlicerCaches, no Charts.
- No arbitrary COM property editor; only the bounded matrix in §5–§6.
- No UI activation, `.Activate`, `.Select`, foreground focus, or dialog handling (background contract unchanged).

### Confidence legend

| Level | Meaning |
|---|---|
| **HIGH** | Confirmed by execution against Excel 16.0 14332 **and** Microsoft Learn documentation. |
| **MEDIUM** | Confirmed by execution but documentation is silent/ambiguous, or documentation-only. |
| **LOW** | Inferred; not verified on the target build. |

---

## 2. Q1 — Inventory: worksheet enumeration vs workbook-level

### Finding (HIGH, empirical)

`Workbook.PivotTables` **exists as an object on this build but is not a usable inventory**:

```
wb.PivotTables              -> <PivotTables object>
wb.PivotTables.Count        -> 0          # even with 3 PivotTables present
```

`Worksheet.PivotTables()` **is complete and correct**:

```
ps.PivotTables().Count      -> 3
ps.PivotTables(1).Name      -> ptC  CacheIndex=2
ps.PivotTables(2).Name      -> ptB  CacheIndex=1
ps.PivotTables(3).Name      -> ptA  CacheIndex=1
```

Every worksheet reports its own count (e.g. `Sheet2 -> 3`, `Data -> 0`), so a full-workbook sweep over `wb.Worksheets` is required.

### Recommendation — deterministic full inventory

Enumerate **all worksheets** (not `Worksheets` only; include any sheet type that exposes `PivotTables`, guarded) and read `ws.PivotTables()` on each:

1. For each `ws` in `wb.Worksheets`: `count = ws.PivotTables().Count`.
2. For `i in 1..count`: `pt = ws.PivotTables(i)`; read identity + cache metadata.
3. Deduplicate caches by `PivotTable.CacheIndex` (HIGH, see §3).
4. Sort deterministically by `(sheet_index, pivot_table_name)`.

Do **not** use `wb.PivotTables`. Do **not** rely on the active sheet or selection.

### Identity fields

| Object | Identity fields (DTO) | Notes |
|---|---|---|
| **PivotTable** | `name` (unique per workbook), `sheet_name`, `sheet_index`, `cache_index` | `Name` is the stable key; `CacheIndex` links to a cache (HIGH). |
| **PivotCache** | `cache_index` (unique per workbook, 1-based, assigned by Excel), `source_type`, `source_data` | `PivotCaches` collection is **not usable** for enumeration on this build (see §3); resolve caches only through `PivotTable.PivotCache()`. |
| **PivotField** | `(pivot_table_name, source_name)` for source fields; `(pivot_table_name, data_field_name)` for data fields | `SourceName` is stable; a data field's `Name` is the caption (may duplicate the source name). Use composite identity, never index alone (HIGH). |

**Pitfall (HIGH):** `PivotField.Position` is only meaningful **within one orientation**. A hidden field returns the sentinel `-2146826246` (`#N/A`). Field indices shift when fields are added/removed, so identity must be by name/source name, not ordinal.

---

## 3. Q2 — Metadata

All properties below were read successfully on 16.0 14332 unless noted.

### PivotTable (HIGH)

| Property | Result / behavior |
|---|---|
| `Name` | stable report name |
| `Index` | 1-based position on the sheet |
| `CacheIndex` | 1-based cache id; `PivotTable.PivotCache().Index` matches it |
| `TableRange1` | full PivotTable range **including** page fields (`$A$1:$B$7` when a page field exists) |
| `TableRange2` | PivotTable range **excluding** page fields (`$A$3:$B$7`) |
| `DataBodyRange` | data area only (`$C$5:$E$10`, `$B$5:$B$7`) — `None` when no data field |
| `RowGrand` / `ColumnGrand` | grand-total flags |
| `PreserveFormatting`, `ShowPageMultipleItemLabel`, `DisplayErrorString`, `PrintTitles`, `PageFieldOrder`, `SubtotalHiddenPageItems`, `RepeatAllLabels`, `CompactRowIndent`, `TableStyle2`, `EnableDrilldown`, `MissingItemsLimit` | readable |
| `RefreshDate` | last refresh timestamp |
| `RefreshTable` | method; returns **Boolean** (`True`) |
| `AllowRefresh`, `PivotTableType`, `OLAP`, `MDX`, `CubeFields`, `EnableDataModel`, `DataModel` | **not exposed** on this build for a regular cache (`AttributeError`); treat as absent, not `False`. |

### PivotCache (HIGH)

| Property | Result |
|---|---|
| `SourceType` | `1 = xlDatabase`, `2 = xlExternal`, `3 = xlConsolidation`, `4 = xlScenario`, `-4148 = xlPivotTable` |
| `SourceData` | `"Data!R1C1:R6C5"` (A1/R1C1 text) for database/Table sources; **array** for external; 2-D array for consolidation |
| `RefreshDate` | last refresh timestamp |
| `RefreshOnFileOpen` | Boolean |
| `EnableRefresh` | Boolean |
| `Index` | 1-based id |
| `RecordCount` | source row count |
| `BackgroundQuery`, `OptimizeCache`, `MissingItemsLimit` | readable |
| `ManualUpdate`, `RefreshError`, `SaveData` | **not exposed** (absent) |
| `Refresh` | method; returns **empty/`None`** (not a success Boolean) |

**Shared caches (HIGH, verified):** two PivotTables built from the same `PivotCache` both report `CacheIndex = 1`, `PivotCaches().Count = 1`; refreshing the cache updates both tables' `RefreshDate` identically. **One cache N:1 PivotTables.**

**Collection-access pitfall (HIGH):** `wb.PivotCaches` behaves as a **method** under early binding and returns an empty collection (`Count == 0`) even when caches exist; `PivotCaches[i]` raises. **Do not use `Workbook.PivotCaches` for inventory.** Resolve a cache via `PivotTable.PivotCache()`.

### PivotField (HIGH)

Readable on regular (non-OLAP) fields: `Name`, `SourceName`, `Caption`, `Orientation`, `Position`, `ShowAllItems`, `AllItemsVisible`, `SubtotalName`, `NumberFormat`, `Value`, `IsCalculated`, `DataType`, `LabelRange`, `DataRange` (data/row fields), `CurrentPage`, `EnableMultiplePageItems`, `ServerBased`, `LayoutForm`, `Subtotals`, `TotalLevels`.

**Not readable on a plain `PivotFields(...)` handle (HIGH):** `Function`, `NumberFormat`, `DataRange`, `LabelRange` raise `Unable to get the <prop> property of the PivotField class` for **any** field obtained from `PivotFields(...)` — including data fields. Use the `DataFields` collection for data-field properties.

### Field-collection map (HIGH, verified)

| Collection | Contains | Access |
|---|---|---|
| `PivotFields()` | **all** source fields + a synthetic `"Data"` field when ≥2 data fields (orientation = row) | method → `PivotFields()`; `len()` / `[i]` |
| `RowFields` / `ColumnFields` / `PageFields` / `DataFields` / `HiddenFields` / `VisibleFields` | orientation subsets | **property** → `pt.RowFields`, `len()`, `[i]`; `pt.RowFields(i)` raises |

Data fields live in `DataFields`, **not** in `PivotFields` as data-orientation entries. A data field's `Name` is its caption; `SourceName` is the underlying source field; `Function` is readable **only** from `DataFields[i]`.

---

## 4. Q3 — Refresh

### Comparison (HIGH)

| Option | Return | Scope | Verdict |
|---|---|---|---|
| `PivotTable.RefreshTable` | **Boolean `True`** | that table's cache + that table | **Primary — table-scoped** |
| `PivotCache.Refresh` | empty/`None` | the whole shared cache (all tables using it) | **Secondary — cache-scoped**, needs explicit justification |
| `Workbook.RefreshAll` | — | entire workbook incl. unrelated connections/queries | **FORBIDDEN** (plan + spec) |

### Recommendation

- **`pivot_table_refresh`** → `PivotTable.RefreshTable`. It is table-addressable, returns a truth signal, and is the least surprising operation for a caller naming one PivotTable.
- **`pivot_cache_refresh`** → `PivotCache.Refresh` **only** when the caller explicitly targets a cache. Because a cache is shared, this affects **every** PivotTable with the same `CacheIndex`; the tool **must** report the affected-table list before and after.
- Because `RefreshTable` operates on the table's cache, a shared-cache refresh triggered through one table can still refresh sibling tables — **always compute and report `affected_pivot_tables` from `CacheIndex`**, regardless of which entry point was used.

### Reporting contract

`PivotRefreshResult`:

- `pivot_table`, `cache_index`, `method` (`refresh_table` | `refresh_cache`)
- `affected_pivot_tables[]` (names on the same cache)
- `refresh_return` (Boolean for `RefreshTable`; `null` for `Refresh`)
- `started_at`, `finished_at`, `duration_ms`
- `before_refresh_date`, `after_refresh_date`
- `status` ∈ `refreshed | unsupported_source | blocked | failed`
- `error` (structured, never a raw COM string)

### Verification rule (HIGH)

A refresh is successful only if the method returned without error **and** `RefreshDate` advanced (or the operation is genuinely a no-op). **Absence of a COM exception is not proof** — this is called out explicitly in the plan.

### Unsupported sources (see §7)

- `SourceType` ∈ `{xlExternal (2), xlScenario (4), xlPivotTable (-4148)}` → `unsupported_source`.
- OLAP / Data Model indicators (MDX, CubeFields, `OLAP`, `ServerBased`, `SourceType == xlExternal`) → `unsupported_source`, **never** invoke a connection refresh.
- Never call `Workbook.RefreshAll`, `Connections`, `QueryTables`, or any Power Query/Data Model entry point.

### Background-safety note (MEDIUM)

`RefreshTable` on a local database cache was observed to return `True` with the hidden instance and no dialog. External/OLAP refresh is explicitly refused before invocation, so the risk of a hidden modal (e.g. a connection prompt) is avoided by policy. If an unexpected modal is ever detected, return `blocked` and enter the existing recovery path — never surface the window.

---

## 5. Q4 — Field configuration matrix

### Orientation (HIGH)

| Code | `XlPivotFieldOrientation` | Observed effect |
|---|---|---|
| 0 | `xlHidden` | removes field from layout |
| 1 | `xlRowField` | adds to `RowFields` |
| 2 | `xlColumnField` | adds to `ColumnFields` |
| 3 | `xlPageField` | adds to `PageFields` (report filter) |
| 4 | `xlDataField` | **does not** create a data field from `PivotFields`; `PivotFields(...).Orientation` reads back `0` |

**Key finding (HIGH):** setting `PivotFields("Sales").Orientation = 4` does **not** add a data field and does not stick (reads back `xlHidden`). Creating a data field **requires** `PivotTable.AddDataField(...)`. This answers the plan's open question: **moving a field into the data area is `AddDataField`, not `Orientation = xlDataField`.**

### Adding a data field — `AddDataField` (HIGH)

```text
PivotTable.AddDataField(Field, Caption?, Function?)   -> PivotField
```

- `Field` is a `PivotField` from `PivotFields(...)`.
- `Caption` optional; `Function` optional.
- Default `Function` when omitted = **`xlSum` (-4157)**.
- `AddDataField` takes **1–3** arguments on this build; a 4th positional argument raises `TypeError` (the documented 3-arg form is correct).
- The returned object **is** the data field; it also appears in `pt.DataFields`.

### Aggregation — `Function` (`XlConsolidationFunction`) (HIGH, verified accepted)

| Code | Name | Accepted |
|---|---|---|
| -4157 | `xlSum` | ✅ |
| -4106 | `xlAverage` | ✅ (note: **-4106**, not -4114) |
| -4112 | `xlCount` | ✅ |
| -4113 | `xlCountNums` | ✅ |
| -4136 | `xlMax` | ✅ |
| -4139 | `xlMin` | ✅ |
| -4149 | `xlProduct` | ✅ (note: **-4149**, not 9) |
| -4155 | `xlStDev` (sample) | ✅ |
| -4156 | `xlStDevP` (population) | ✅ |
| -4164 | `xlVar` (sample) | ✅ (note: **-4164**) |
| -4165 | `xlVarP` (population) | ✅ (note: **-4165**) |
| 11 | `xlDistinctCount` | ❌ (requires Data Model) |
| 1000 | `xlUnknown` | ❌ |

`Function` is **read/written only via `DataFields[i].Function`**, never via `PivotFields(...)`.

**Legacy-code warning (HIGH):** the commonly cited legacy values `-4114` (Average), `-4154`, `-4153`, and `9` (Product) are **not** accepted by this build. Use the `XlConsolidationFunction` values above. This is a concrete compatibility correction to typical VBA examples.

### Position (HIGH)

- `PivotField.Position` is 1-based **within its orientation** (Rows/Columns/Pages/Data).
- Setting `Position` reorders **within the current orientation** and swaps neighbours; it does not move a field to another orientation.
- Out-of-range values (`0`, `-1`, `99`) raise `Unable to set the Position property`.
- A hidden/not-placed field reads `Position = -2146826246` (`#N/A` sentinel) — treat as "no position".

### Removal (HIGH)

| Operation | Effect |
|---|---|
| `PivotField.Orientation = 0` | removes a field from the layout (row/col/page); for a data field, set `DataFields[i].Orientation = 0` — verified `DataFields` len 2→1 |
| `PivotField.Delete` | removes the field from `PivotFields` entirely — verified `PivotFields` len 5→4 |

Prefer `Orientation = 0` for layout changes (non-destructive to the field list); `Delete` is destructive to the field list and must be an explicit, separate capability or omitted.

### Captions and duplicate rules (HIGH)

- `DataFields[i].Caption` is writable and independent of `PivotFields(...).Caption`.
- `DataFields[i].NumberFormat` is writable (verified `General` → `#,##0.00`).
- **Duplicate captions are rejected**: `AddDataField(field, "<existing caption>")` raises "unable to add the data field"; the existing data field is unchanged. Callers must supply a unique caption or omit it (Excel generates `"Sum of <field>"`).
- A data field whose caption equals a source name is also rejected as a duplicate.

### Bounded field-configuration matrix (recommendation)

| Capability | API | Allowed | Notes |
|---|---|---|---|
| Set orientation | `PivotField.Orientation` | `xlHidden`, `xlRowField`, `xlColumnField`, `xlPageField` | `xlDataField` rejected → use AddDataField |
| Add data field | `PivotTable.AddDataField` | field + optional caption + optional function | 1–3 args; unique caption |
| Set aggregation | `DataFields[i].Function` | the 11 accepted codes above | via `DataFields` only |
| Remove data field | `DataFields[i].Orientation = 0` | — | non-destructive |
| Remove layout field | `PivotField.Orientation = 0` | — | non-destructive |
| Set position | `PivotField.Position` | 1..N within orientation | swap semantics; validate bounds |
| Set caption | `DataFields[i].Caption` | unique | read back via `DataFields` |
| Set number format | `DataFields[i].NumberFormat` | Excel format string | data fields only |
| Set `ShowAllItems` | `PivotField.ShowAllItems` | Boolean | see note |
| Delete field | `PivotField.Delete` | — | **optional/omit**; destructive to field list |

**`ShowAllItems` note (HIGH, subtle):** `ShowAllItems` defaults to `False`. After hiding an item manually, reading `ShowAllItems` still returns `False`, but setting `ShowAllItems = True` does **not** make `AllItemsVisible` return `True` in the same session — use `ClearAllFilters()` to restore full visibility. `ShowAllItems` controls display of items with no data, not manual-filter state.

### Data-field naming / identity (HIGH)

`DataFields[i]` and `PivotFields("Sum of Sales")` are **different COM objects** (`_oleobj_` identity check → `False`), and both differ from `PivotFields("Sales")`. The data field's `Name` is its caption; the source field's `SourceName` is the underlying column. A field identity DTO must therefore carry **both** `source_name` and `data_field_name` (caption) plus the owning `pivot_table`.

---

## 6. Q5 — Filters

### Manual / label / value filters (HIGH)

**Manual filter via `PivotItems.Visible` — works:**

```text
pt.PivotFields("Region").PivotItems()      # collection, len()/[i]
pt.PivotFields("Region").PivotItems(1).Visible = False
pt.PivotFields("Region").AllItemsVisible   # -> False
pt.PivotFields("Region").VisibleItems      # ['South']
pt.PivotFields("Region").HiddenItems       # ['North']
pt.PivotFields("Region").ClearAllFilters() # -> AllItemsVisible True
```

- `PivotItems` is a **method** → call `PivotItems()`, then index with `[i]` or `len()`.
- `PivotItem.Visible` is readable/writable; `PivotItem.Value` / `.Name` give the item identity.
- Scope: **one field**.

**`PivotFilters.Add` — fails on this build (HIGH, empirically):**

Every attempt to add a filter through `PivotField.PivotFilters.Add(...)` returned `0x80020009` (`DISP_E_EXCEPTION`) with an empty `PivotFilters` collection, for all of:
- positional and named forms,
- correct `XlPivotFilterType` values (15 `xlCaptionEquals`, 17 `xlCaptionBeginsWith`, 21 `xlCaptionContains`, 7 `xlValueEquals`, 13 `xlValueIsBetween`, 29 `xlSpecificDate`, 35 `xlDateBetween`, 1 `xlTopCount`, …),
- early-bound `PivotFilters` and `PivotFilters.Add2`,
- direct `InvokeTypes(181, …)` and `**{"Type": …}`.

The failure was independent of field orientation (row/column/page/hidden) and of table layout. The generated signature is correct (`Type, DataField, Value1, Value2, Order, Name, Description, MemberPropertyField[, WholeDayFilter]`), so this is a **runtime behavior of the target build**, not a marshalling error. `ClearAllFilters` / `ClearLabelFilters` / `ClearValueFilters` / `ClearManualFilter` **do** work.

### Correct `XlPivotFilterType` values (reference, HIGH from docs)

For completeness if a later build supports `Add`:

| Group | Values |
|---|---|
| Top/Bottom | 1 `xlTopCount`, 2 `xlBottomCount`, 3 `xlTopPercent`, 4 `xlBottomPercent`, 5 `xlTopSum`, 6 `xlBottomSum` |
| Value | 7 `xlValueEquals`, 8 `xlValueDoesNotEqual`, 9 `xlValueIsGreaterThan`, 10 `xlValueIsGreaterThanOrEqualTo`, 11 `xlValueIsLessThan`, 12 `xlValueIsLessThanOrEqualTo`, 13 `xlValueIsBetween`, 14 `xlValueIsNotBetween` |
| Caption (label) | 15 `xlCaptionEquals`, 16 `xlCaptionDoesNotEqual`, 17 `xlCaptionBeginsWith`, 18 `xlCaptionDoesNotBeginWith`, 19 `xlCaptionEndsWith`, 20 `xlCaptionDoesNotEndWith`, 21 `xlCaptionContains`, 22 `xlCaptionDoesNotContain`, 23–26 caption >/≥/</≤, 27 `xlCaptionIsBetween`, 28 `xlCaptionIsNotBetween` |
| Date | 29 `xlSpecificDate`, 30 `xlNotSpecificDate`, 31 `xlBefore`, 32 `xlBeforeOrEqualTo`, 33 `xlAfter`, 34 `xlAfterOrEqualTo`, 35 `xlDateBetween`, 37–51 relative dates, 52 `xlYearToDate`, 53–68 period quarters/months |

**Caution:** the widely circulated "0 = xlCaptionEquals" style map is **wrong**; the correct base is 15 for caption equals and 7 for value equals.

### Clear operations (HIGH)

- `PivotField.ClearAllFilters` — clears manual + label + value filters on that field (verified `AllItemsVisible` → `True`).
- `PivotTable.ClearAllFilters` — clears filters for the whole PivotTable (verified).
- `PivotField.ClearLabelFilters`, `ClearValueFilters`, `ClearManualFilter` — targeted clears (available).

### Recommendation for `task_051`

Implement filters as:

1. **Read** — for each field report `AllItemsVisible`, `VisibleItems[]`, `HiddenItems[]`, and (if present) `PivotFilters` count/types. Treat `PivotFilters.Add` as **unsupported** on this build: expose an explicit `unsupported` status rather than a partially-working path.
2. **Apply manual filter** — set `PivotItems[i].Visible` for the named item set (bounded to one field). This is the reliable, verified path.
3. **Clear** — `pivot_field_clear_filters` with explicit scope `field` (default) or `table`; both verified.
4. **Label/value/date filters** — return `unsupported` with a clear reason on the target build; do **not** silently convert to a manual filter.

### Filter matrix

| Filter kind | Mechanism | Status on target build | Scope |
|---|---|---|---|
| Manual (hide/show items) | `PivotItem.Visible` | ✅ supported | one field |
| Clear all (field) | `PivotField.ClearAllFilters` | ✅ supported | one field |
| Clear all (table) | `PivotTable.ClearAllFilters` | ✅ supported | whole table |
| Clear label / value / manual | `ClearLabelFilters` / `ClearValueFilters` / `ClearManualFilter` | ✅ supported | one field |
| Label (caption) filter | `PivotFilters.Add(15..28)` | ❌ unsupported (runtime error) | — |
| Value filter | `PivotFilters.Add(7..14)` | ❌ unsupported | — |
| Date filter | `PivotFilters.Add(29..68)` | ❌ unsupported | — |
| Top/Bottom | `PivotFilters.Add(1..6)` | ❌ unsupported | — |

---

## 7. Q6 — Source types

### Detection (HIGH)

| Source | Detection | Behavior |
|---|---|---|
| Worksheet range / ListObject | `SourceType == 1` (`xlDatabase`), `SourceData` is text `"Sheet!R1C1:RnCm"` | **supported** — inventory, refresh, fields, manual filters |
| External (ODBC/OLEDB/Query) | `SourceType == 2` (`xlExternal`), `SourceData` is an **array**, `CommandType`/`CommandText`/`Connection` present | **unsupported/read-only** |
| Consolidation | `SourceType == 3` (`xlConsolidation`) | **unsupported/read-only** |
| Scenario | `SourceType == 4` (`xlScenario`) | **unsupported/read-only** |
| Another PivotTable | `SourceType == -4148` (`xlPivotTable`) | **unsupported/read-only** |
| OLAP / Data Model | `PivotField.ServerBased == True`, `MDX`/`CubeFields`/`OLAP` present, `Function = xlDistinctCount` only | **unsupported/read-only** |

`SourceData` is a **string** for database/Table sources and a **Variant array** for external sources — the DTO must carry `source_data_kind` (`text` | `array`) plus a bounded text rendering, never a raw nested array of arbitrary size.

### Rule

- Do **not** invoke `PivotCache.Refresh` for `SourceType != 1`.
- Do **not** touch `Workbook.Connections`, `QueryTables`, `Model`, `PowerPivot`, or any deferred feature.
- Return a structured `unsupported_source` status carrying `source_type`, `source_type_name`, and a human-readable reason.

### OLAP-specific field caveats (HIGH, documented)

- For OLAP sources, setting `Orientation` on one field of a hierarchy sets it for the whole hierarchy; dimension fields are limited to row/column/page and measures to data (`PivotField.Orientation` docs).
- For OLAP, `ShowAllItems` is always `False` and there are no hidden fields (`PivotTable.PivotFields` / `ShowAllItems` docs).
- These mutations are **out of scope**; OLAP is detected and refused.

---

## 8. Q7 — Recommendation: minimal Phase-2 Pivot contract

### Layering (unchanged)

Follow the existing `TablesService` / `tables_adapter` / `table_tools` pattern exactly:

```
api/pivot_tools.py            thin async tool handlers; validate args; delegate; return DTO
  └─ application/pivot_tables_service.py   COM-free orchestration; session resolve;
       snapshot/approval/audit gates for mutations; submit single STA callable
       └─ adapters/excel/pivot_tables_adapter.py   STA-local; touches PivotTable/PivotCache/
            PivotField; returns DTOs/primitives only (com_worker._ensure_dto guard)
```

All COM stays inside `com_worker.py`'s single STA thread. No proxy leaves the worker.

### DTO sketch (domain/models.py additions — spec §5.2, DTO-only)

```text
PivotSourceType      enum: database | external | consolidation | scenario | pivot_table | olap | unknown

PivotCacheInfo
  cache_index: int
  source_type: PivotSourceType
  source_data: str | None            # text form (database/Table)
  source_data_kind: str              # "text" | "array" | "unavailable"
  refresh_date: datetime | None
  refresh_on_file_open: bool | None
  enable_refresh: bool | None
  record_count: int | None
  shared: bool                       # >1 PivotTable references this cache
  affected_pivot_tables: list[str]

PivotFieldInfo
  source_name: str
  caption: str | None
  orientation: str                   # hidden|row|column|page|data
  position: int | None               # None when hidden / #N/A sentinel
  function: str | None               # data fields only, from DataFields
  show_all_items: bool | None
  all_items_visible: bool | None
  is_data_field: bool
  data_field_name: str | None        # caption when is_data_field
  number_format: str | None          # data fields only

PivotTableInfo
  name: str
  sheet_name: str
  sheet_index: int
  cache_index: int
  table_range1: str | None
  table_range2: str | None
  data_body_range: str | None
  row_grand: bool | None
  column_grand: bool | None
  refresh_date: datetime | None
  data_fields: list[PivotFieldInfo]
  row_fields / column_fields / page_fields / hidden_fields: list[PivotFieldInfo]
  source_status: str                 # supported | unsupported_source
  source_type: PivotSourceType

PivotInventory
  workbook_id: str
  pivot_tables: list[PivotTableInfo]
  caches: list[PivotCacheInfo]

PivotRefreshResult
  pivot_table: str
  cache_index: int
  method: str                        # refresh_table | refresh_cache
  affected_pivot_tables: list[str]
  refresh_return: bool | None
  started_at / finished_at / duration_ms
  before_refresh_date / after_refresh_date: datetime | None
  status: str                        # refreshed | unsupported_source | blocked | failed
  error: dict | None
```

### Tool sketch (api/pivot_tools.py)

Read-only (no snapshot/approval):

- `pivot_tables_list(workbook_id) -> PivotInventory`
- `pivot_table_get(workbook_id, pivot_table) -> PivotTableInfo`
- `pivot_cache_get(workbook_id, cache_index) -> PivotCacheInfo`
- `pivot_fields_list(workbook_id, pivot_table) -> list[PivotFieldInfo]`
- `pivot_field_get_filters(workbook_id, pivot_table, field) -> PivotFieldFilters`

Mutating (snapshot + approval + audit + postcondition):

- `pivot_table_refresh(workbook_id, pivot_table) -> PivotRefreshResult`
- `pivot_cache_refresh(workbook_id, cache_index) -> PivotRefreshResult` (reports affected tables)
- `pivot_field_set_layout(workbook_id, pivot_table, field, orientation) -> PivotFieldInfo`
- `pivot_field_set_position(workbook_id, pivot_table, field, position) -> PivotFieldInfo`
- `pivot_field_set_aggregation(workbook_id, pivot_table, field, function, caption?) -> PivotFieldInfo` (adds/updates a data field via `AddDataField` / `DataFields[i].Function`)
- `pivot_field_remove(workbook_id, pivot_table, field, scope) -> …` (Orientation=0)
- `pivot_field_apply_manual_filter(workbook_id, pivot_table, field, visible_items) -> PivotFieldInfo`
- `pivot_field_clear_filters(workbook_id, pivot_table, field?, scope) -> …`

**Not in Phase 2 (explicit):** `pivot_table_create`, `pivot_table_delete`, `pivot_cache_create`, `pivot_cache_delete`, label/value/date `PivotFilters.Add`, `Workbook.RefreshAll`, external/OLAP refresh.

### Status/error extensions (domain)

Reuse the existing typed-error/status model. Add only if needed, keeping the closed vocabulary style:

- `unsupported_source` (status value on the result, not an exception, for a determinable-but-unsupported cache)
- `ambiguous_target` (field/caption resolution matched 0 or >1)
- `refresh_failed` (Excel returned no success signal)
- `blocked` (unexpected modal / cannot complete without UI)
- `field_not_found`, `invalid_orientation`, `invalid_function`, `duplicate_caption`, `position_out_of_range`

These map onto the existing `ExcelVbaMcpError` + `context` conventions; prefer a structured result object over raising for expected unsupported cases.

### Workbook-model integration (task_048)

`WorkbookModelService.get_sheet_model` / `WorkbookModel` gain a **`pivot_tables` section** built from `PivotTablesService.list_pivot_tables` (aggregate, never re-traverse). Add `pivot_tables: list[PivotTableInfo]` to `SheetInfo`/`SheetModel` and a `SECTION_PIVOTS` section with `ModelSectionStatus`. Do **not** materialize the PivotTable cell output — metadata only.

### Test fixtures (task_041)

`PivotTablesProject.xlsm` with:
- a worksheet-range cache and a ListObject-backed cache (both `SourceType == 1`);
- row/column/page/data fields with distinct aggregations;
- two PivotTables sharing one cache (to verify `affected_pivot_tables` and `CacheIndex`);
- a manual filter to verify `PivotItem.Visible`;
- a deterministic local refresh source (values changed between refreshes);
- optional (only if creatable without a Data Model) an external/OLAP-labelled cache to exercise `unsupported_source`.

---

## 9. Constraints / Risks

| ID | Risk | Mitigation |
|---|---|---|
| P1 | `wb.PivotTables` returns an empty inventory | Use per-worksheet `ws.PivotTables()` sweep (§2). |
| P2 | `wb.PivotCaches` is unusable for enumeration | Resolve caches via `PivotTable.PivotCache()`; never enumerate `Workbook.PivotCaches`. |
| P3 | `PivotFilters.Add` fails on 14332 | Expose manual filters only; return `unsupported` for label/value/date. Re-verify on other builds before enabling. |
| P4 | `Function` unreadable/writable via `PivotFields` | Use `DataFields[i].Function`; enforce the accepted code set. |
| P5 | Shared cache refresh affects sibling tables | Report `affected_pivot_tables` before/after; require explicit cache-targeting. |
| P6 | OLAP/Data Model accidentally mutated | Detect via `SourceType`/`ServerBased`/`MDX`; refuse with `unsupported_source`. |
| P7 | Hidden field `Position` sentinel `-2146826246` | Treat as `None` in the DTO. |
| P8 | Duplicate data-field captions rejected | Validate uniqueness pre-call; generate or require a unique caption. |
| P9 | Refresh success is not a COM-exception absence | Verify `RefreshDate` advanced and/or a truth return. |
| P10 | Legacy aggregation codes silently wrong | Use `XlConsolidationFunction` values only (§5). |
| P11 | Unexpected modal during a mutation | Return `blocked`; do not surface/dismiss; enter existing recovery. |
| P12 | Data-field identity is composite | Carry `source_name` + `data_field_name` + owning table. |

### Assumptions

- **MEDIUM:** the observed `PivotFilters.Add` failure reproduces on other 16.0 builds with the same major version; it has only been verified on 14332. The tool must degrade to `unsupported`, not fail hard.
- **MEDIUM:** `RefreshTable` returning `True` with the hidden instance and no dialog is treated as the success signal; this was observed on local database caches only.

---

## 10. References (Microsoft Learn)

- PivotTable object — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable
- PivotTable.PivotFields — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.pivotfields
- PivotTable.AddDataField — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.adddatafield
- PivotTable.RefreshTable — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.refreshtable
- PivotTable.TableRange1 / TableRange2 / DataBodyRange — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable
- PivotCache object — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache
- PivotCache.SourceData — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache.sourcedata
- PivotCache.Refresh — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache.refresh
- XlPivotTableSourceType — https://learn.microsoft.com/en-us/office/vba/api/excel.xlpivottablesourcetype
- XlPivotFieldOrientation — https://learn.microsoft.com/en-us/office/vba/api/excel.xlpivotfieldorientation
- PivotField.Orientation — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.orientation
- PivotField.Position — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.position
- PivotField.ShowAllItems — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.showallitems
- PivotField.Delete — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.delete
- PivotFilters.Add — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfilters.add
- XlPivotFilterType — https://learn.microsoft.com/en-us/office/vba/api/excel.xlpivotfiltertype
- XlConsolidationFunction — https://learn.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
- PivotItems / PivotItem — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotitems
- PivotTable.ClearAllFilters — https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.clearallfilters
- PivotField.ClearAllFilters / ClearLabelFilters / ClearValueFilters / ClearManualFilter — https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield

---

## Appendix A — Empirical evidence index

All probes were read-only, in-memory, against a hidden `DispatchEx` Excel instance; nothing was written to disk or to the tracked workbook. Probe scripts live under `.ai/tmp/pivot_probe*.py`.

| # | Probe | Key result |
|---|---|---|
| 3 | first successful build | `PivotCaches().Add(1, src)` + `PivotTables().Add(...)`; `TableRange1/2`, `CacheIndex`, `SourceType`, `RefreshDate` |
| 8 | data-field creation | `Orientation=4` does not create a data field; `AddDataField` does |
| 9 | field enumeration | `PivotFields` vs `RowFields`/`DataFields`/`PageFields` |
| 11 | orientation matrix | codes 0–4 mapped to `xlHidden`…`xlDataField` |
| 12 | filters | `PivotFilters.Add` fails; `ClearAllFilters` works |
| 13 | filter sweep | all `PivotFilters.Add` types fail (wrong map initially) |
| 14 | early binding | `PivotFilters` type is `PivotFilters`; `PivotItems` is a method |
| 15 | direct invoke | `InvokeTypes(181)` also fails |
| 17 | cache sharing | shared cache → identical `RefreshDate`; `wb.PivotTables.Count == 0` |
| 18 | metadata | cache/table/field properties; `TableRange1` includes page fields |
| 22 | removal | `Orientation=0` removes; `Delete` removes from field list |
| 24 | collections | correct `len()`/`[i]` access; data-field `Function` readable via `DataFields` |
| 25 | aggregation | accepted `XlConsolidationFunction` codes; legacy codes rejected |
| 26 | `AddDataField` arity | 1–3 args; 4th positional raises |
| 27 | ListObject cache | `SourceData` = `"Data!R1C1:R4C3"`; `SourceType == 1` |
| 28 | identity | `DataFields[1]` ≠ `PivotFields("Sum of Sales")` ≠ `PivotFields("Sales")` |
| 29 | PivotItems | `len()==2`; manual hide → `Visible`/`Hidden` sets; `ClearAllFilters` restores |