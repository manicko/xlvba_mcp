---
id: ranges-cells-api-research
domain: research
tags:
  - research
  - ranges
  - cells
  - excel-object-model
  - phase2
  - decision-record
related:
  - doc-maintenance-rules
  - stage1-acceptance-matrix
---

# Range & Cells API — Research Decision Record (task_038)

## Purpose

This document records the research findings and the recommended Phase‑2
**Range / Cells** contract for the Excel VBA MCP Server. It is the decision
record required by `plan_phase2_ranges_cells_pivottables.md` → `task_038` and
feeds `task_040` (freeze DTO/tool contracts) and the implementation tasks
`task_042`–`task_046`. It compares the documented Excel Object Model
alternatives, records compatibility constraints for the declared target build,
and proposes a minimal contract that preserves the existing
`domain → application → adapters/excel → api` layering and the **DTO-only
STA-worker boundary**.

It does **not** modify source code and does **not** expand Phase‑2 scope beyond
the approved specification (§14: read/write values & formulas, bounded
formatting, copy/insert/delete/merge/unmerge, data validation, conditional
formatting, bounded bulk I/O, typed values/errors).

### Scope

| In scope (this record) | Out of scope (owned elsewhere) |
|---|---|
| Value/Value2 typing, 1‑D vs 2‑D shape normalization | PivotTables/Caches/Fields (`task_039`+) |
| Formula vs Formula2 read/write | Charts, Slicers, Connections, Power Query, Data Model |
| Bounded bulk I/O and response-size policy | ListObject row/column structure changes (`TablesService`) |
| Bounded formatting property set + partial updates | Arbitrary COM property/Style/Font/Interior serialization |
| Copy / Insert / Delete / Merge / UnMerge + preconditions | Clipboard, UI automation, Activate/Select navigation |
| DataValidation & FormatConditions supported rule matrix | AutoFilter, hyperlinks, comments, PageSetup (deferred) |
| Multi-area / merged / spill / protected / hidden constraints | Any Python/shell/VBA/formula evaluation |

### Confidence legend

**HIGH** = directly documented by Microsoft Learn / Office VBA reference (cited).
**MEDIUM** = documented behavior + inference from the object model that must be
confirmed by an integration test on the target build.
**LOW** = behavior that is version/locale-dependent or not officially specified;
treated as a constraint to reject/guard rather than rely on.

---

## 1. Value serialization — `Range.Value` vs `Range.Value2`

### 1.1 Documented behavior

- **`Range.Value2`** differs from **`Range.Value`** in exactly one way: it does
  **not** use the `Currency` and `Date` data types; date/currency-formatted
  values are returned as **floating-point `Double`** serials.
  ([Range.Value2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2))
  — **HIGH**
- **`Range.Value`** returns a `Variant`; for date/currency-formatted cells the
  `Variant` carries `VT_DATE`/`VT_CY` so a typed consumer sees `Date`/`Currency`.
  ([Range.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)) — **HIGH**
- **Shape**: for a range whose **first area contains more than one cell**,
  `Value` returns a **2‑D `Variant` array**. A single-cell range returns a
  **scalar `Variant`**, not an array. (Same remark applies to `Value2`.)
  ([Range.Value remarks](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)) — **HIGH**
- **Assigning** a 2‑D array writes in one operation; if the target is larger than
  the array, remaining cells receive an **error value**. Assigning an array to a
  **multi-area** range is "not properly supported and should be avoided."
  ([Range.Value remarks](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value)) — **HIGH**
- **`Range.Text`** is the formatted display string; for multi-cell ranges it
  returns `Null` unless all cells share identical content+format. It is
  read-only and display-dependent (column width can alter it). Not suitable as a
  data channel. ([Range.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.range.text)) — **HIGH**
- **`Range.NumberFormat`** returns `Null` when cells disagree; this is the signal
  used to detect *mixed* formatting. ([Range.NumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat)) — **HIGH**

### 1.2 pywin32 / COM marshalling reality (target stack)

- The server reads `Range.Value` **through pywin32 on the STA worker**. A
  `VT_DATE` `Variant` is marshalled by pywin32 into a Python `datetime`; a
  `VT_CY` (Currency) `Variant` is marshalled as a numeric type. Excel **error**
  cells read via `Value`/`Value2` surface as **signed OLE error codes**
  (`xlErr*`), *not* as literal text such as `"#N/A"`. The Stage‑1
  `tables_adapter.py` already documents and maps these codes
  (`_EXCEL_ERROR_CODE_TO_LITERAL`) and is the established convention to reuse.
  — **MEDIUM** (documented `xlErr*` values; pywin32 numeric marshalling confirmed
  by existing Stage‑1 code + tests)
- Because `Value2` collapses `Date`/`Currency` to `Double`, **Value2 alone cannot
  distinguish a genuine number from a date/currency** — a raw serial like
  `45123.0` is ambiguous. The task's constraint "do not guess dates from a raw
  numeric value alone" follows directly. — **HIGH**

### 1.3 Options

| Option | How | Pros | Cons |
|---|---|---|---|
| **A. Value only** | Read `range.Value`; trust pywin32 to yield `datetime` for dates. | Typed dates for free; single read. | Date detection depends on the cell being date-formatted *and* pywin32 marshalling; a numeric-looking date is indistinguishable from a number if format is lost; Currency/Decimal handling is version-sensitive. |
| **B. Value2 only** | Read `range.Value2`; type numbers/bools/strings/empties/errors. | Stable numeric channel; never emits a `datetime` unexpectedly. | **Cannot** detect dates or currency — a date serial looks like a plain number. Violates "no guessing". |
| **C. Value + NumberFormat (recommended)** | Read `range.Value2` for the raw channel **and** read `range.NumberFormat` (or per-cell `NumberFormatLocal`) to decide date/currency typing. | Lossless, explicit, no guessing; matches Stage‑1 error typing; format tells the truth about dates. | Two properties per read (cheap); needs a bounded format-normalization helper. |
| **D. Value + Text** | Use `Text` for display. | Human-readable. | Display/column-width dependent, lossy, locale-dependent — reject for data. |

### 1.4 Recommendation (Q1)

**Adopt Option C.** Read `Range.Value2` as the raw value channel, and
classify dates **only** with an explicit, non-guessing rule based on
`Range.NumberFormat`:

1. Read `Value2` for the whole bounded range (scalar for 1 cell, 2‑D tuple for a
   block). This is the existing `_normalize_grid` behavior in
   `tables_adapter.py` — reuse it.
2. Read `Range.NumberFormat` once for the range. If it is `Null` (mixed), fall
   back to **per-cell** `NumberFormat` reads **only for the cells being
   returned** (bounded), so mixed date/non-date columns are typed correctly
   without guessing.
3. Classification order per cell (mirrors Stage‑1 `_cell_value`):
   `None → EMPTY`; `bool → BOOLEAN`; OLE error code or error literal
   `→ ERROR`; number `→ NUMBER` **unless** the cell's number format is a date/time
   format **and** the value is a plausible Excel serial **→ DATE** (ISO‑8601
   string); `str → STRING`.
4. **Never** infer a date from a bare numeric when the number format is
   `General` or numeric. If the format is ambiguous (`General`, mixed, or
   unavailable) the value is returned as `NUMBER` with the format echoed so the
   caller can decide. This is the truthful "don't guess" behavior.
5. Booleans are checked **before** numbers (`bool` is an `int` subclass) — same
   ordering Stage‑1 already uses.
6. Blanks: an **unset** cell returns `None` from `Value2` → `EMPTY` with
   `value=None`. Do not coerce to `""` or `0`.
7. Errors: map the OLE code to the literal (`#NULL!`, `#DIV/0!`, `#VALUE!`,
   `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA`) and type as
   `CellValueKind.ERROR` — never as a plain string. Reuse
   `_EXCEL_ERROR_CODE_TO_LITERAL`.

**Why not `Range.Value`?** `Value` is a legitimate alternative (Option A) and
would surface `datetime`, but the classification would then depend on pywin32
marshalling and on the cell already being date-formatted — and a wrong/General
format still yields a number. Reading `Value2` + `NumberFormat` makes the type
decision **explicit and auditable**, which is what "no lossy assumptions" needs.
The two-property read is bounded and cheap.

**1‑D vs 2‑D shape:** `Value2` returns a scalar for a single cell and a **2‑D**
tuple-of-tuples for any multi-cell range (even a single row `A1:C1` or single
column `A1:A3` comes back as a 2‑D array, per the "first area contains more than
one cell" rule). The contract **normalizes everything to 2‑D** internally and
reports explicit `rows` × `cols` plus a `shape` tag so a single row/column is
unambiguous to the client. Do **not** depend on a 1‑D array being returned.

---

## 2. Formulas — `Formula` vs `Formula2`

### 2.1 Documented behavior

- **`Range.Formula2`** "supersedes" `Range.Formula` in Dynamic-Arrays-enabled
  Excel; `Formula` remains supported for backward compatibility.
  ([Range.Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2),
  [Formula vs Formula2](https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2)) — **HIGH**
- `Formula` uses **Implicit Intersection Evaluation (IIE)** and "can never
  spill"; `Formula2` uses **Array Evaluation (AE)** and "may spill if more than
  1 result is returned." Excel auto-translates between them, inserting/removing
  the implicit-intersection operator `@`. — **HIGH**
- Both return the formula as a string **including the leading `=`**, or the
  constant for a non-formula cell, or `""` for an empty cell. Both accept a
  same-dimensioned array for multi-cell writes. — **HIGH**
- Neither is available for **OLAP** data sources. — **HIGH**
- Newer functions are stored in the file with an **`_xlfn.`** prefix (and
  worksheet-only functions with **`_xlfn._xlws.`**). When a formula is read via
  `Formula`/`Formula2` on a build that recognizes the function, the prefix is
  stripped; on an older build it is preserved and the cell shows `#NAME?` after
  recalculation. `LET`/`LAMBDA` parameters use `_xlpm`. Spill references (`#`)
  and `@` can appear as `_xlfn.ANCHORARRAY(...)` / `_xlfn.SINGLE(...)` in older
  builds. ([MS-XLSX future functions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf);
  [Peltier Tech `_xlfn` test](https://peltiertech.com/vba-test-new-excel-functions/)) — **HIGH**
- `Range.SavedAsArray` reports whether the cells would be saved as a legacy
  array formula (used to detect DA/IIE round-trip differences). — **HIGH**

### 2.2 Target build: Excel 16.0 build 14332

Build 14332 is a **Dynamic-Arrays-capable** build, so **`Formula2` is available
and is the correct primary channel**. `Formula` is still present for
compatibility and for callers that want pre-DA (IIE) semantics. — **HIGH**
(build capability) / **MEDIUM** (exact feature flags of 14332 must be confirmed
by the `task_041` fixture; do not hard-code).

### 2.3 "A string starting with `=` is not a formula"

Assigning a string that begins with `=` through **`Range.Value`/`Value2`** does
**not** create a formula — Excel stores it as **text**. Only assigning through
**`Formula`/`Formula2`** creates a formula. This is why the contract exposes
`range_write_values` and `range_write_formulas` as **separate explicit
operations** and never infers a formula from a value string. — **HIGH**
(value assignment is literal; formula assignment is the only path to a formula)

### 2.4 Recommendation (Q2)

- **Read:** return `Formula2` (DA formula bar form) as the primary formula
  string, plus `SavedAsArray` (bool/null) and `HasArray` (bool/null) so callers
  can see whether a cell participates in a (legacy) array. Optionally include
  `Formula` (IIE form) as a secondary field when it differs from `Formula2`
  (i.e. when an `@` was added/removed) — this is cheap and aids round-tripping,
  but the **primary** is `Formula2`.
- **Write:** `range_write_formulas` writes through **`Formula2`** (so dynamic
  arrays can spill as authored). Accept a formula string **with** the leading
  `=`; reject a string that does not start with `=` (a formula write must be an
  actual formula — a bare constant belongs in `range_write_values`).
- **`_xlfn`/`_xlws`:** pass formula strings **verbatim** to Excel; do **not**
  strip or rewrite prefixes. On write, if a caller supplies `_xlfn.`-prefixed
  text, pass it through unchanged (Excel strips it when recognized). On read,
  report whatever `Formula2` returns verbatim (do not "clean" it), and surface a
  `notes`/diagnostic when an `_xlfn`/`_xlws`/`_xlpm` prefix is present so the
  client knows the formula depends on a newer function.
- **No evaluation:** never call `Application.Evaluate`, `eval`, or any formula
  engine. Formulas are opaque strings passed to/from Excel only.
- **OLAP:** if `Range.Formula2` raises (OLAP source), return an explicit
  `unsupported`/read-only status, not a fabricated empty string.

---

## 3. Bounded bulk I/O

### 3.1 Problem

`Range.Value`/`Value2` can return an arbitrarily large 2‑D array, and a
single-cell read loop across `UsedRange` can be enormous. The worker boundary
must never carry an unbounded payload, and a single COM call must not block the
STA worker past its soft deadline.

### 3.2 Recommended limits (concrete defaults; configurable)

| Limit | Default | Rationale |
|---|---|---|
| `MAX_RANGE_CELLS` (per read/write request) | **100,000 cells** | ≈ 316×316; keeps a `Value2` matrix and JSON response bounded. |
| `MAX_RANGE_ROWS` | **100,000** | Guard against `A:A` / whole-column requests. |
| `MAX_RANGE_COLS` | **16,384** (XFD) hard cap; practical default **512** | Excel's own column ceiling. |
| `MAX_RESPONSE_CELLS` (returned cells, after offset/limit) | **50,000** | Response-size ceiling; truncate with a `truncated` flag, like `TableRows`. |
| `MAX_RESPONSE_BYTES` (serialized tool result) | **~8 MB** | Final guard before returning; if exceeded, return a structured `oversized` diagnostic instead of the payload. |
| `MAX_WRITE_PAYLOAD_CELLS` | **100,000** | Reject a write payload larger than this **before** touching Excel. |
| Per-call soft timeout | existing worker `_DEFAULT_TIMEOUT_S` (60 s) | Reuse the existing per-call deadline. |

These mirror the established Stage‑1 pattern (`TablesAdapter._MAX_RESPONSE_CELLS
= 2000`, `truncated` flag, `_MAX_PREVIEW_ROWS`) and simply scale it to full-range
requests.

### 3.3 Rejection & response-size behavior

1. **Resolve + validate the address first** (sheet exists, A1 parses, `Areas.Count
   = 1`). Reject multi-area with a typed error (`Areas.Count > 1`).
2. **Compute cell count** (`rows * cols`) **before** reading. If it exceeds
   `MAX_RANGE_CELLS`, **reject** with a structured `RANGE_TOO_LARGE` /
   `VALUE_ERROR` diagnostic containing `rows`, `cols`, `cells`, and the limit —
   do **not** partially read.
3. For reads, if `rows * cols` is within the request cap but the **returned**
   cells would exceed `MAX_RESPONSE_CELLS` (e.g. caller asked `A1:Z10000`), read
   a **bounded window** (`offset`/`limit` or a row cap) and set `truncated=true`.
4. For writes, validate payload shape (`rows × cols` equals target) and cell count
   **before** any COM mutation (this is already the Stage‑1 rule "wrong-sized
   payloads are rejected before any workbook mutation").
5. **Byte guard:** after building the DTO, if `model_dump_json()` length exceeds
   `MAX_RESPONSE_BYTES`, replace the cell matrix with a structured
   `oversized=true` diagnostic (address, rows, cols, cells, bytes) rather than
   emitting a huge payload.

### 3.4 Recommendation (Q3)

Expose `offset`/`limit` (or `start_row`/`max_rows`) on `range_read_values`; enforce
the cell/row/column/byte ceilings above; **reject oversized requests before
reading** with a typed error, and **truncate oversized responses** with a
`truncated` flag. Never return a whole-sheet `UsedRange` unless explicitly
requested **and** within limits.

---

## 4. Formatting contract

### 4.1 Principle

Do **not** serialize `Style`, `Font`, `Interior`, `Borders`, or any COM object.
Expose a **minimal, bounded, typed** property set. Every property is
**tri-state** on read: `value` / `mixed` (cells disagree → Excel returns `Null`)
/ `unsupported` (cannot be represented without flattening).

### 4.2 Recommended property set (minimal but sufficient for Phase 2)

| DTO field | Excel property | Type | Notes |
|---|---|---|---|
| `number_format` | `Range.NumberFormat` | `str \| null` | `null` = mixed across range (`mixed=true`). Use `NumberFormat` (en-US codes), not `NumberFormatLocal`. |
| `font.bold` | `Range.Font.Bold` | `bool \| null` | `null` = mixed. |
| `font.italic` | `Range.Font.Italic` | `bool \| null` | `null` = mixed. |
| `font.size` | `Range.Font.Size` | `float \| null` | `null` = mixed. |
| `font.color` | `Range.Font.Color` | `int (RGB) \| null` | Return the `RGB` integer (`BGR` OLE `Color` value); `null` = mixed/automatic. |
| `font.name` | `Range.Font.Name` | `str \| null` | Optional; include only if cheap to read. |
| `fill.color` | `Range.Interior.Color` | `int (RGB) \| null` | `null` = no fill / mixed. `xlColorIndexNone` → `null`. |
| `fill.pattern` | `Range.Interior.Pattern` | `str (enum) \| null` | Optional; default `xlSolid`. |
| `alignment.horizontal` | `Range.HorizontalAlignment` | `str (enum) \| null` | `xlLeft`/`xlCenter`/`xlRight`/`xlJustify`/`xlGeneral`/`xlCenterAcrossSelection`/`xlDistributed`/`xlFill`. |
| `alignment.vertical` | `Range.VerticalAlignment` | `str (enum) \| null` | `xlTop`/`xlCenter`/`xlBottom`/`xlJustify`/`xlDistributed`. |
| `alignment.wrap_text` | `Range.WrapText` | `bool \| null` | `null` = mixed. |

Plus a **read-side** companion block:

- `mixed_properties: list[str]` — names of properties that returned `Null`
  (cells disagree). The value field is `null` **and** the property is listed here,
  so a mixed range is never silently flattened to one cell's value.
- `unsupported: list[str]` — properties that exist on the range but are **not**
  represented by this DTO (e.g. `Borders`, `Gradient`, `ThemeColor`,
  `NumberFormatLocal`, `ProtectContents`). Reporting them explicitly (not
  omitting them) satisfies "keep mixed/unsupported formatting explicit".

### 4.3 Partial updates

`range_set_format` accepts a **sparse** `RangeFormat` where every field is
**optional**. Only **present** (non-`None`) fields are written; absent fields
are **left untouched** (do **not** reset unspecified properties). To
distinguish "set to null/empty" from "leave unchanged", use a Pydantic
`model_fields_set` check (`"font.bold" in request.model_fields_set`) rather than
value-truthiness. `number_format` is a string set; `font.color`/`fill.color`
accept an RGB int; `null` explicitly clears to automatic where Excel allows.

### 4.4 Recommendation (Q4)

Adopt the table above as the **only** writable/read formatting surface.
`range_get_format` returns the typed DTO + `mixed_properties` + `unsupported`.
`range_set_format` performs a **partial** update of the provided fields only,
then **re-reads** the same properties as a postcondition. Any property outside
this set is rejected with a structured `unsupported` diagnostic rather than
silently ignored.

---

## 5. Structural operations — Copy / Insert / Delete / Merge / UnMerge

### 5.1 Documented behavior

- **`Range.Copy(Destination)`** copies to `Destination` when given; **if omitted
  it copies to the Clipboard.** → The contract **always requires an explicit
  destination**; omitting it is rejected (no clipboard). ([Range.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy)) — **HIGH**
- **`Range.Insert(Shift, CopyOrigin)`** shifts cells: `xlShiftToRight` (-4161) or
  `xlShiftDown` (-4121). If `Shift` is omitted, **Excel decides from the range
  shape** (implicit). `CopyOrigin` = `xlFormatFromLeftOrAbove` (default) /
  `xlFormatFromRightOrBelow`. ([Range.Insert](https://learn.microsoft.com/en-us/office/vba/api/excel.range.insert),
  [XlInsertShiftDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.xlinsertshiftdirection)) — **HIGH**
- **`Range.Delete(Shift)`** shifts with `xlShiftToLeft` (-4159) or
  `xlShiftUp` (-4162); if omitted Excel decides from shape.
  ([Range.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.range.delete),
  [XlDeleteShiftDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.xldeleteshiftdirection)) — **HIGH**
  (Insert uses `XlInsertShiftDirection`: `xlShiftToRight` (-4161) /
  `xlShiftDown` (-4121); Delete uses `XlDeleteShiftDirection`: `xlShiftToLeft`
  (-4159) / `xlShiftUp` (-4162). Use the documented enum names; do not hard-code
  raw ints — resolve the enum in the adapter.)
- **`Range.Merge(Across)`** merges into one cell; the value lives in the
  **upper-left** cell; other cells' contents are **discarded**.
  `Across=True` merges each row separately. ([Range.Merge](https://learn.microsoft.com/en-us/office/vba/api/excel.range.merge)) — **HIGH**
- **`Range.UnMerge`** separates any merged area in the range; no return value.
  ([Range.UnMerge](https://learn.microsoft.com/en-us/office/vba/api/excel.range.unmerge)) — **HIGH**
- **`Range.MergeCells`** returns `True`/`False`/`Null` (mixed);
  **`Range.MergeArea`** is the merged range containing a cell (single-cell only).
  ([Range.MergeCells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergecells),
  [Range.MergeArea](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea)) — **HIGH**
- **`Application.Intersect(a, b, ...)`** returns the rectangular intersection or
  **`Nothing`** if they don't intersect; errors if areas are on different
  worksheets. This is the guard primitive for ListObject intersection.
  ([Application.Intersect](https://learn.microsoft.com/en-us/office/vba/api/excel.application.intersect)) — **HIGH**

### 5.2 Preconditions & guards (must be enforced before mutation)

1. **Explicit destination for Copy.** Reject if `destination` is missing. Reject
   if source and destination overlap. Reject if destination is on a **different
   worksheet than the source is allowed to be** — cross-sheet copy with an explicit
   destination is permitted, but the destination range must be resolvable and the
   target sheet writable.
2. **Explicit Shift for Insert/Delete.** Reject if `shift` is omitted
   ("ambiguous shift behavior must be rejected"). Only accept
   `right`/`down` (insert) and `left`/`up` (delete). Reject a Shift value that is
   not one of the documented enum members.
3. **Merged cells.** If the target range intersects a merged area (check
   `MergeCells`/`MergeArea`), reject unless the caller explicitly sets
   `allow_merged=true`; when allowed, report the affected `MergeArea` addresses
   in the preview. Rationale: Insert/Delete/merge over merged cells changes
   merge geometry in ways that are easy to get wrong.
4. **Non-anchor data loss (Merge).** `Merge` discards every non-anchor value.
   **Preflight** must read the range, detect **non-empty non-anchor cells**, and
   reject with a structured preview (`cells_with_data`) unless the caller passes
   explicit approval (`allow_data_loss=true` after preview). This mirrors the
   `task_045` acceptance criterion "merge with non-anchor data is blocked unless
   explicitly approved after preview."
5. **Formulas referencing the range.** Insert/Delete rewrites references across
   the workbook; if the range is referenced by formulas elsewhere, report
   `referenced=true` in the preview (best-effort via `Range` dependency scan is
   expensive — at minimum detect formulas **within the same sheet's**
   `UsedRange` that textually reference the address; do not claim completeness).
6. **ListObject intersection (see §5.3).**
7. **Protected sheet / read-only / hidden (see §7).**

### 5.3 ListObject interaction — hard guard

- **Rule:** any **structural** Range operation (`Insert`, `Delete`, `Merge`,
  and `Copy` **whose destination intersects a ListObject**) that **intersects a
  ListObject's `Range`** MUST be **rejected** and routed to `TablesService`
  (ListRows/ListColumns). Rationale (spec §7.6, plan task_045): "Structural
  operations that intersect or alter a ListObject are rejected and direct callers
  to the existing TablesService ListRows/ListColumns operations. Range writes to
  explicit cells inside a table are a separate cell-level capability."
- **Detection:** for each ListObject on the target sheet, compute
  `Application.Intersect(targetRange, listObject.Range)`; if the result is
  **not `Nothing`**, the target intersects the table → reject with
  `LISTOBJECT_INTERSECTION` and include `table`, `table_range`,
  `intersection_address` in the context.
- **Cell-level writes** (values/formulas, §1–2) **inside** a table are allowed
  as explicit cell updates **provided they do not change the table's row/column
  structure** — they must not call `ListRows.Add`/`ListColumns.Add` or resize.
  This is the line between "edit a cell in a table" (allowed) and "change table
  structure" (TablesService only).
- **`Range.ListObject`** returns the `ListObject` for a range that is (part of)
  a table; use it as a secondary check.

### 5.4 Recommendation (Q5)

Implement `range_copy`, `range_insert`, `range_delete`, `range_merge`,
`range_unmerge` with **mandatory explicit arguments** (destination; shift),
**preflight previews** for merged cells / non-anchor data / references, and a
**hard ListObject-intersection rejection** that names the owning table and
points to `TablesService`. Every structural op is snapshot-backed, approved,
audited, and postcondition-verified (resulting address/dimensions, affected
table metadata unchanged). No clipboard, no `Select`/`Activate`, no UI.

---

## 6. DataValidation & FormatConditions

### 6.1 `Range.Validation` (DataValidation)

Documented `Validation.Add(Type, AlertStyle, Operator, Formula1, Formula2)`:

- **`Type`** ∈ `XlDVType`: `xlValidateWholeNumber`(1), `xlValidateDecimal`(2),
  `xlValidateList`(3), `xlValidateDate`(4), `xlValidateTime`(5),
  `xlValidateTextLength`(6), `xlValidateCustom`(7), `xlInputOnly`(0).
  ([XlDVType](https://learn.microsoft.com/en-us/office/vba/api/excel.xldvtype)) — **HIGH**
- **`Operator`** ∈ `XlFormatConditionOperator`: `xlBetween`, `xlEqual`,
  `xlGreater`, `xlGreaterEqual`, `xlLess`, `xlLessEqual`, `xlNotBetween`,
  `xlNotEqual`. ([Validation.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.add)) — **HIGH**
- **Per-type argument rules** (documented):
  - `xlValidateCustom`: `Formula1` required (an expression → True/False);
    `Formula2` ignored.
  - `xlValidateList`: `Formula1` required (comma-delimited list **or** a
    worksheet reference); `Formula2` ignored.
  - `xlValidateWholeNumber` / `xlValidateDate` / `xlValidateDecimal` /
    `xlValidateTextLength` / `xlValidateTime`: `Formula1` and/or `Formula2`
    (both for `xlBetween`/`xlNotBetween`).
  - `xlInputOnly`: no `Formula1`/`Formula2`.
- **`Formula1` length limit: ≤ 255 characters** (documented on `Validation.Add`).
  → enforce and reject longer with a structured error.
- **Properties:** `IgnoreBlank` (bool), `InCellDropdown` (bool, **only** for
  `xlValidateList`), `AlertStyle` (`xlValidAlertInformation`/`Stop`/`Warning`),
  `ErrorMessage`/`ErrorTitle`, `InputMessage`/`InputTitle`,
  `ShowError`/`ShowInput`, `IMEMode`, `ErrorTitle`. `Validation.Modify` mirrors
  `Add`. `Validation.Delete` removes it.
  ([Validation.InCellDropdown](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.incelldropdown)) — **HIGH**

### 6.2 `Range.FormatConditions` (Conditional Formatting)

Documented `FormatConditions.Add(Type, Operator, Formula1, Formula2)` returns a
`FormatCondition`:

- **`Type`** ∈ `XlFormatConditionType`: `xlCellValue`(1), `xlExpression`(2),
  `xlColorScale`(3), `xlDataBar`(4), `xlTop10`(5), `xlIconSet`(6),
  `xlUniqueValues`(8), `xlTextString`(9), `xlBlanksCondition`(10),
  `xlTimePeriod`(11), `xlAboveAverageCondition`(12), `xlNoBlanksCondition`(13),
  `xlErrorsCondition`(16), `xlNoErrorsCondition`(17).
  ([XlFormatConditionType](https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype)) — **HIGH**
- **`Operator`** same `XlFormatConditionOperator` set; ignored when
  `Type=xlExpression`.
- **`Formula1`/`Formula2`**: value, string, cell reference, or formula (no
  evaluation).
- **`FormatCondition` properties:** `Priority` (1..N, unique across the sheet;
  setting it renumbers others — documented), `StopIfTrue` (bool; **default
  `True`** for backward compatibility, opposite of UI-created rules),
  `Interior.Color`, `Font`, `NumberFormat`, `AppliesTo` (the range the rule
  applies to; for `xlExpression`/`xlCellValue`).
  ([FormatCondition.StopIfTrue](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.stopiftrue),
  [FormatCondition.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.priority)) — **HIGH**
- `FormatConditions.Delete` / `FormatCondition.Delete` / `FormatConditions.Modify`.
- Newer rule objects (`ColorScale`, `DataBar`, `IconSet`, `Top10`,
  `AboveAverage`) are **different** COM objects, not plain `FormatCondition`.

### 6.3 Support matrix (what to support vs reject)

| Area | Support | Reject / report unsupported |
|---|---|---|
| Validation types | `xlValidateWholeNumber`, `xlValidateDecimal`, `xlValidateList`, `xlValidateDate`, `xlValidateTime`, `xlValidateTextLength`, `xlValidateCustom` | `xlInputOnly` (no rule semantics) → reject for `set`; read as `input_only` if present |
| Validation props | `Type`, `Operator`, `Formula1`, `Formula2`, `IgnoreBlank`, `InCellDropdown`, `AlertStyle`, `ErrorMessage`, `ErrorTitle`, `InputMessage`, `InputTitle`, `ShowError`, `ShowInput` | — |
| Validation string limits | `Formula1`/`Formula2` ≤ **255 chars**; message/title strings ≤ 255 chars (Excel UI limit) | over-long → structured reject |
| FormatCondition types | `xlCellValue`, `xlExpression`, `xlTextString`, `xlBlanksCondition`, `xlNoBlanksCondition`, `xlErrorsCondition`, `xlNoErrorsCondition`, `xlTimePeriod`, `xlTop10`, `xlAboveAverageCondition`, `xlUniqueValues` | `xlColorScale`, `xlDataBar`, `xlIconSet` → read-only detection + explicit `unsupported` status (different COM objects; out of minimal Phase‑2 scope) |
| FormatCondition props | `Type`, `Operator`, `Formula1`, `Formula2`, `Priority`, `StopIfTrue`, `Interior.Color`, `Font` (subset), `AppliesTo` | Full `Borders`/gradient/theme rendering → not serialized |
| Order/priority | Read all rules; set `Priority` to insert at a position (renumbers others) | — |

**Formula evaluation:** `Formula1`/`Formula2` (validation and conditional
formatting) are **opaque strings** passed to Excel; never evaluated in Python.
Unsupported rule types and mixed/invalid formulas return explicit structured
diagnostics — never silently convert one rule type into another.

### 6.4 Recommendation (Q6)

`range_get_validation` returns a typed `ValidationRule` DTO (type, operator,
formula1/2, ignore_blank, in_cell_dropdown, alert_style, messages, show flags) or
`present=false`. `range_set_validation` uses `Validation.Add`/`Modify` with the
documented per-type argument rules, enforces the 255-char limit, and
postcondition-reads the result. `range_get_conditional_formatting` returns a
list of `ConditionalFormatRule` DTOs (with `priority`, `stop_if_true`,
`interior_color`, `font` subset, `applies_to`) and flags `unsupported` types;
`range_set_conditional_formatting` supports add/modify/delete of the **supported**
types, preserves priority ordering, and rejects unsupported types explicitly.

---

## 7. Constraints — reject vs guess

| Constraint | Detection | Behavior |
|---|---|---|
| **Multi-area range** (`Areas.Count > 1`) | `Range.Areas.Count` | **Reject** — `Value`/array writes are "not properly supported" on multi-area; never guess an area. |
| **Merged cells** | `MergeCells`, `MergeArea` | Reads: return only the anchor value (`MergeArea.Cells(1,1)`), flag `merged=true`. Writes: reject non-anchor writes unless `allow_merged`. Structural merge: reject non-anchor data loss without approval. |
| **Spill ranges** | `Range.Spilling` / `Range.SpillParent` (DA), `HasArray` (legacy) | Reads: reading the anchor returns the anchor value; report `spilling`/`in_spill_range` so callers don't mistake spill cells for independent inputs. Writes into a spill range (non-anchor) → **reject** (`SPILL_RANGE_WRITE`) to avoid `#SPILL!`; writing the anchor formula is allowed via `range_write_formulas`. **MEDIUM** — confirm property names on 14332 via fixture. |
| **Protected sheet / read-only** | `Worksheet.ProtectContents`, `Range.Locked` (True/False/Null), `WorkbookInfo.read_only` | Mutating ops → **reject** `PROTECTED_TARGET`/`WORKBOOK_READ_ONLY`; reads allowed. Never unprotect automatically. |
| **Hidden rows/columns** | `Range.Hidden` (whole row/col only) | Reads: include hidden cells but flag `hidden=true` if any target row/col is hidden (so callers know values may be hidden). Writes: allowed but flagged; do not auto-unhide. |
| **ListObject intersection** | `Application.Intersect(target, listObject.Range)` | Structural ops → **reject** `LISTOBJECT_INTERSECTION`, route to `TablesService`. Cell writes inside table allowed if structure unchanged. |
| **Ambiguous shift (Insert/Delete)** | caller omitted `shift` | **Reject** — never let Excel guess from shape. |
| **Oversized request** | `rows*cols > MAX_RANGE_CELLS` | **Reject** before reading with `RANGE_TOO_LARGE`. |
| **Formula string without `=`** | caller passed to `range_write_formulas` | **Reject** — use `range_write_values` for constants. |
| **OLAP formula source** | `Range.Formula2` raises | Return `unsupported`/read-only status, not fabricated `""`. |
| **Clipboard copy** | `destination` omitted | **Reject** — clipboard is never used. |

---

## 8. Recommendation — minimal Phase‑2 Range/Cells contract

### 8.1 Layering preserved (no architecture change)

```
api/range_tools.py            ← MCP tool schemas (thin, validate + delegate)
      │
application/ranges_service.py ← RangesService: guards, snapshot, approval, audit,
      │                         bounded limits, DTO assembly; NO COM
      │
adapters/excel/ranges_adapter.py ← RangesAdapter: STA-local; touches Range only;
      │                            returns DTOs/primitives only
      ▼
adapters/excel/com_worker.py  ← single STA worker (unchanged boundary)
```

Every new type is a **Pydantic DTO** in `domain/models.py`; new statuses in
`domain/statuses.py`; new errors reuse the closed `domain/errors.py` vocabulary
where possible (add codes only when spec §11 has no fit — e.g. a
`RANGE_TOO_LARGE`/`LISTOBJECT_INTERSECTION`/`PROTECTED_TARGET` code, or map to the
existing `VALUE`-style base error with a rich `context`). **No COM object,
pywin32 proxy, or Excel enum object crosses the worker boundary** — the
`_ensure_dto` guard in `com_worker.py` already enforces this.

### 8.2 DTO sketch (names only; no code changes in this task)

**Read**
- `RangeAddress { sheet, address, a1, rows, cols, cells, areas, is_merged, merged_area?, in_table?, table_name?, spilling?, spill_parent? }`
- `RangeValues { range: RangeAddress, shape: "scalar"|"row"|"column"|"block", rows: int, cols: int, values: list[list[CellValue]], truncated: bool, notes: list[str] }` — `CellValue` reuses the Stage‑1 typed value (`kind/value/address`).
- `RangeFormulas { range: RangeAddress, shape, rows, cols, formulas: list[list[FormulaCell]], notes }`; `FormulaCell { address, formula2, formula?, is_array, saved_as_array }`
- `RangeFormat { range, number_format?, font{name?,bold?,italic?,size?,color?}, fill{color?,pattern?}, alignment{horizontal?,vertical?,wrap_text?}, mixed_properties: list[str], unsupported: list[str] }`
- `ValidationRule { present, type, operator?, formula1?, formula2?, ignore_blank?, in_cell_dropdown?, alert_style?, error_message?, error_title?, input_message?, input_title?, show_error?, show_input? }`
- `ConditionalFormatRule { priority, type, operator?, formula1?, formula2?, stop_if_true, applies_to, interior_color?, font{...}, supported: bool, unsupported_reason? }`

**Write / mutate** (all return a result DTO with before/after + postcondition)
- `range_write_values` → `RangeMutationResult { action, range_before: RangeAddress, range_after: RangeAddress, verified: bool, notes }`
- `range_write_formulas` → same shape
- `range_set_format` → `FormatMutationResult { range, applied: list[str], ignored: list[str], verified }`
- `range_copy` / `range_insert` / `range_delete` / `range_merge` / `range_unmerge` → `StructuralMutationResult { action, source, destination?, shift?, before, after, affected_addresses, table_metadata_unchanged: bool, verified, preconditions: {...} }`
- `range_set_validation` / `range_set_conditional_formatting` → rule mutation results with priority preserved

### 8.3 Tool sketch (names only)

Read-only: `range_get`, `range_read_values`, `range_read_formulas`,
`range_get_format`, `range_get_validation`, `range_get_conditional_formatting`.
Mutating (guarded: writable + snapshot + approval + audit): `range_write_values`,
`range_write_formulas`, `range_set_format`, `range_copy`, `range_insert`,
`range_delete`, `range_merge`, `range_unmerge`, `range_set_validation`,
`range_set_conditional_formatting`.

### 8.4 Cross-cutting rules

- All targets are **explicit** (`workbook_id`, `sheet`, A1 address). No
  `ActiveSheet`/`Selection`/`ActiveCell`.
- Reads are bounded (§3); writes validate shape before mutation.
- Every mutating tool follows the Stage‑1 guard chain (writable session →
  snapshot → approval `authorize_vba_write` with a per-action key → audit →
  postcondition read-back).
- `ListObject` structural intersection → reject + point to `TablesService`.
- No `Activate`/`Select`/`SendKeys`/clipboard/UI.
- Background contract (addendum): hidden, non-activating, settings restored on
  all paths; an unexpected modal → truthful blocked/`indeterminate`, no dismiss.

---

## 9. Constraints & Risks

| ID | Risk | Mitigation |
|---|---|---|
| R‑Q1 | Date detection ambiguity if relying on `Value2` alone | Use `NumberFormat` to type dates; never guess from a bare serial (§1.4). |
| R‑Q2 | `_xlfn`/`_xlws` prefixes vary by build | Pass formula strings verbatim; flag prefixes in notes; confirm on 14332 via fixture. |
| R‑Q3 | `Formula` vs `Formula2` spill semantics differ | Default writes to `Formula2`; expose `SavedAsArray`/`HasArray` on read. |
| R‑Q4 | Oversized read floods the worker boundary | Cell/row/byte ceilings + `truncated` flag + pre-read reject (§3). |
| R‑Q5 | Mixed formatting silently flattened | Tri-state DTO + `mixed_properties`/`unsupported` (§4.2). |
| R‑Q6 | Merge discards data / Insert/Delete rewrites refs | Preflight previews + explicit approval; best-effort reference report (§5.2). |
| R‑Q7 | Range structural op corrupts a ListObject | Hard `Intersect` guard → route to TablesService (§5.3). |
| R‑Q8 | Conditional-formatting type object model differs (`ColorScale`/`DataBar`/`IconSet`) | Support only plain `FormatCondition` types; report others as `unsupported` (§6.3). |
| R‑Q9 | Validation 255-char limit / message limits | Validate lengths before COM mutation (§6.1). |
| R‑Q10 | Protected / hidden / spill / multi-area edge cases | Explicit reject-vs-flag matrix (§7). |
| R‑Q11 | Target build feature flags unknown (14332) | Confirm `Formula2`/`Spilling`/`SavedAsArray` via `task_041` fixtures; do not hard-code. |
| R‑Q12 | Any dependency on Clipboard/UI/Activate | Reject missing destination; no clipboard; no Select/Activate (§5, §8.4). |

---

## 10. References (Microsoft Learn / official)

**Values & shape**
- Range.Value2 — https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2
- Range.Value — https://learn.microsoft.com/en-us/office/vba/api/excel.range.value
- Range.Text — https://learn.microsoft.com/en-us/office/vba/api/excel.range.text
- Range.NumberFormat — https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformat
- Range.NumberFormatLocal — https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformatlocal
- Range.Count / CountLarge — https://learn.microsoft.com/en-us/office/vba/api/excel.range.count
- Range.Areas — https://learn.microsoft.com/en-us/office/vba/api/excel.range.areas

**Formulas**
- Range.Formula — https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula
- Range.Formula2 — https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2
- Formula vs Formula2 — https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2
- Range.SavedAsArray — https://learn.microsoft.com/en-us/office/vba/api/excel.range.savedasarray
- Range.HasArray — https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasarray
- MS-XLSX future functions (`_xlfn`) — https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf

**Structural**
- Range.Copy — https://learn.microsoft.com/en-us/office/vba/api/excel.range.copy
- Range.Insert — https://learn.microsoft.com/en-us/office/vba/api/excel.range.insert
- XlInsertShiftDirection — https://learn.microsoft.com/en-us/office/vba/api/excel.xlinsertshiftdirection
- Range.Delete — https://learn.microsoft.com/en-us/office/vba/api/excel.range.delete
- XlDeleteShiftDirection — https://learn.microsoft.com/en-us/office/vba/api/excel.xldeleteshiftdirection
- Range.Merge — https://learn.microsoft.com/en-us/office/vba/api/excel.range.merge
- Range.UnMerge — https://learn.microsoft.com/en-us/office/vba/api/excel.range.unmerge
- Range.MergeCells — https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergecells
- Range.MergeArea — https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea
- Application.Intersect — https://learn.microsoft.com/en-us/office/vba/api/excel.application.intersect
- Range.ListObject — https://learn.microsoft.com/en-us/office/vba/api/excel.range.listobject
- ListRows.Add — https://learn.microsoft.com/en-us/office/vba/api/excel.listrows.add
- ListObject / ListRows / ListColumns — https://learn.microsoft.com/en-us/office/vba/api/excel.listobject

**Data validation**
- Validation.Add — https://learn.microsoft.com/en-us/office/vba/api/excel.validation.add
- Validation.Modify — https://learn.microsoft.com/en-us/office/vba/api/excel.validation
- Validation.InCellDropdown — https://learn.microsoft.com/en-us/office/vba/api/excel.validation.incelldropdown
- XlDVType — https://learn.microsoft.com/en-us/office/vba/api/excel.xldvtype

**Conditional formatting**
- FormatConditions.Add — https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.add
- XlFormatConditionType — https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype
- FormatCondition.StopIfTrue — https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.stopiftrue
- FormatCondition.Priority — https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.priority

**Formatting**
- Range.WrapText — https://learn.microsoft.com/en-us/office/vba/api/excel.range.wraptext
- Interior.ColorIndex — https://learn.microsoft.com/en-us/office/vba/api/excel.interior.colorindex
- Font.ColorIndex / Font.FontStyle — https://learn.microsoft.com/en-us/office/vba/api/excel.font.colorindex

**Constraints**
- Range.Locked — https://learn.microsoft.com/en-us/office/vba/api/excel.range.locked
- Range.Hidden — https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden
- Range.SpecialCells — https://learn.microsoft.com/en-us/office/vba/api/excel.range.specialcells
- AllowEditRanges.Add — https://learn.microsoft.com/en-us/office/vba/api/excel.alloweditranges

**Project documents**
- Plan: `.ai/plans/plan_phase2_ranges_cells_pivottables.md` (task_038, task_040–task_047)
- Addendum: `.ai/plans/plan_addendum_background_execution.md` (BG01–BG09)
- Spec: `docs/01-Spec/excel_vba_mcp_technical_spec.docx` §14
- Stage‑1 conventions reused: `src/excel_vba_mcp/adapters/excel/tables_adapter.py`
  (`_cell_value`, `_normalize_grid`, `_MAX_RESPONSE_CELLS`, `truncated`),
  `application/tables_service.py` (guard chain, `_bind_workbook`,
  `_submit_read`/`_submit_mutation`).

---

**Confidence statement.** Value/Value2 typing, formula (Formula2 vs Formula,
`_xlfn`), structural method signatures, validation types/properties, and
conditional-formatting types are **HIGH** (directly documented and cited).
Build-specific behavior on Excel 16.0.14332 (`Formula2`, `Spilling`/`SpillParent`,
`SavedAsArray` availability and exact enums) is **MEDIUM** and must be confirmed
by the `task_041` integration fixtures before `task_040` freezes the contract.
No unsupported requirement has been added beyond the approved Phase‑2 scope.