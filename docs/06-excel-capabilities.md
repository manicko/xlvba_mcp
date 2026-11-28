---
id: xlvba-excel-capabilities
domain: 06-excel-capabilities
tags:
  - excel
  - capabilities
  - coverage
  - mcp
  - vba
  - pivottables
related:
  - xlvba-overview
  - xlvba-mcp-tools-reference
  - xlvba-workflows-and-examples
  - xlvba-configuration
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This document answers one question: **what can you (or your AI agent) actually accomplish in
Excel through `xlvba`?**

It is a capability map organised by user goal, not a tool manual. For each area it states what you
can do, which of the **43 MCP tools** (and which `action` of a dispatch tool) performs it, the
prerequisites, and the limits that matter in practice. The exhaustive per-tool contract —
parameters, return values, errors — is in [MCP tools reference](05-mcp-tools-reference.md); this
document classifies the surface and its limits.

Two facts frame everything below:

1. `xlvba` drives a **single hidden Excel instance** on Windows. There is no visible window, no
   window activation and no keystroke input. Because there is one Excel instance, **a long-running
   operation delays later ones** — plan long reads, exports or macros accordingly.
2. The server runs in **full access mode** by default: mutating and macro-running operations are
   authorised without a per-action grant. Setting `EXCEL_VBA_MCP_ACCESS_MODE=restricted` opts into
   deny-by-default approval, where those operations are refused with `APPROVAL_DENIED`. "Implemented"
   and "permitted over MCP" still differ (a baseline snapshot is required before a mutation, and
   `macro_run` needs a `background_safe` classification), so every row keeps them separate.

## How to read this document

### Coverage statuses

| Status | Meaning |
|--------|---------|
| **Fully supported** | `xlvba` has an implemented, bounded path for the operation, with validation and a postcondition read where the operation mutates. For mutating rows this means "implemented and guarded", not "freely callable over MCP". |
| **Partially supported** | Some sub-operations or variants work while others are rejected or out of scope, or the result is inherently limited (for example an always-indeterminate compile, manual PivotTable filters only, or a fixed set of cell-format fields). |
| **Not supported** | No tool exists for the capability. Nothing in the server exposes it. |
| **Not verified** | The path is allowed or inferred, but the behaviour is not validated on the target Excel build (validated build: Excel 16.0.14332, `.xlsm`). |

Coverage describes the **Excel surface**. It does not mean the operation is unconditionally permitted
through MCP. In the default `full` access mode mutating and executing operations are authorised (a
snapshot is still required before a mutation); in `restricted` mode they are denied by default. The
Limitations column says so explicitly. The security model is in
[Security](08-security.md) and [Architecture](04-architecture.md).

### Operation kinds

| Kind | Meaning | Approval class |
|------|---------|----------------|
| READ | Returns information without changing workbook state. | `read-only`, auto-allowed. |
| CREATE | Adds a new object, row or rule. | `mutating`, guarded. |
| UPDATE | Changes an existing object, value or property. | `mutating`, guarded. |
| DELETE | Removes an object, row or cell content. | `mutating`, guarded. |
| EXECUTE | Runs code (a macro or a declarative suite). | `executing`, guarded and additionally gated. |

### Guard chain in one line

Every CREATE/UPDATE/DELETE/EXECUTE tool runs the same order: **writable session → snapshot →
approval → operation-specific checks → postcondition re-read + audit**. The full chain and its
rationale are in [Architecture](04-architecture.md).

## Capability matrix

The matrix covers the **43-tool surface** (20 dispatch tools with an `action`, plus 23 standalone
tools; 74 action values in total — see [MCP tools reference](05-mcp-tools-reference.md)). Tool names
are the canonical MCP registration names; clients may show them with a prefix (for example
`xlvba_workbook`).

| User goal (Excel capability) | What you can do | MCP tools (action) | Coverage | Prerequisites and limits |
|------------------------------|-----------------|--------------------|----------|--------------------------|
| **Workbook sessions and safe copies** | Open an existing workbook as a hidden working copy; close or save it; read session status, the open-workbook list, or an aggregated model. | `workbook` (`open`, `close`, `save`), `workbook_read` (`status`, `list`, `model`) | Fully supported | The original is copied into `work/` and only the copy is opened — the original is never modified. Default open is writable and execution-capable (`mode="execute"`, `read_only=False`); pass `mode="inspect"` and/or `read_only=True` for a macros-disabled or read-only session. Path must be inside `EXCEL_VBA_MCP_ROOTS` (`PATH_NOT_ALLOWED`). No new-workbook creation; non-`.xlsm` unverified. |
| **Save As / independent copy** | Produce a new (or again-saved) copy without rebinding the session; verified write with a backup. | `workbook_save_copy` | Fully supported | Atomic write with a single `<target>.bak` and a final size + SHA-256 verify (`outcome="ok"` only after verify). An existing target needs `overwrite_existing`; overwriting the original needs `overwrite_source` plus a matching source hash (`STALE_TARGET` otherwise). |
| **Snapshots and rollback** | Capture an immutable baseline; list, compare, and restore to it. | `snapshot_create`, `snapshot_list`, `snapshot_diff`, `snapshot_restore` | Fully supported | Snapshots are immutable and carry a SHA-256 manifest. `snapshot_restore` is destructive to the working copy and needs `confirm=true` plus the old `workbook_id`; it rebinds to the original and **does not undo external macro side effects** (database writes, emails, files). |
| **Non-COM snapshot read (OOXML)** | Read workbook metadata, a page of sheet rows, or saved PivotCache records from an immutable snapshot **without occupying the Excel worker**. | `workbook_snapshot_read` (`metadata`, `sheet-rows`, `pivot-cache-records`) | Fully supported | Reads the **saved** cache records, not the displayed PivotTable result; `unsupported_source` when a cache has no local records. The MCP-transport run of this tool is validated in-process only (not over a live client transport). |
| **Background jobs and cancellation** | Run a long operation asynchronously, poll its status, and record a cancel intent. | `job_get`, `job_request_cancel` | Fully supported | `job_get` never touches Excel. `job_request_cancel` records **intent only** — it never interrupts the running call, which still reaches a terminal state. |
| **VBA inspection and search** | Inspect the VBA project, components, procedures and references; search code; export all components with per-file hashes. | `vba_read` (`project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics`), `vba_export_project` | Fully supported | Requires **Trust access to the VBA project object model** (VBOM) and an unlocked project (`ACCESS_VBOM_DISABLED` / `VBA_PROJECT_LOCKED` otherwise). Export writes `.bas`/`.cls`/`.frm`+`.frx` under the artifact root; Document modules are never removed. |
| **VBA editing** | Replace a whole module, or insert code at a semantic position, with revision safety. | `vba_replace_component_source`, `vba_insert_code` | Partially supported | Standard and Class modules only (Document modules and UserForms are rejected). Requires a matching `expected_source_hash` (the normalized `source_hash` from `vba_read` `read-component`); a stale hash raises `REVISION_CONFLICT`. |
| **Compile diagnostics** | Get a compile verdict and partial diagnostics. | `vba_compile_diagnostics`, `vba_read` (`get-compile-diagnostics`) | Partially supported | **Fail-closed by design:** on the validated build the status is always `indeterminate` (`reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"`, `is_compile=false`). It **never** returns `compiled` or `compile_error` and never shows UI. A clean `indeterminate` result is **not** proof that the project compiles. |
| **Macro execution** | Run an approved, non-interactive macro and get a typed result. | `macro` (`list`, `classify-background`, `get-last-result`), `macro_run` | Partially supported | Gates: execute mode + writable session → procedure resolution → approval → `background_safe` for the current workbook hash **and** VBA source revision → interactive-UI preflight (R13). An unclassified procedure returns `BACKGROUND_SAFETY_NOT_VERIFIED`; a procedure that shows UI is refused (`INTERACTIVE_UI_NOT_SUPPORTED`). 90 s soft deadline; a hang self-recovers the session. |
| **Declarative testing** | Run bounded edit→macro→assert suites on disposable copies; read a persisted run. | `test_suite` (`run`, `get-run`) | Partially supported | Closed vocabularies (37 actions, 29 assertions); unknown actions or unbounded payloads are rejected before any Excel call. One disposable copy per case. VBA replace/insert steps are rejected (a declarative step cannot carry the revision hash); external side effects are not rolled back. |
| **Sheets** | List/get sheets; add, rename, copy, delete; set visibility. | `sheet_read` (`list`, `get`), `sheet` (`add`, `rename`, `copy`, `set-visibility`), `sheet_delete` | Partially supported | `sheet` `add` is worksheet-only; `copy` is in-workbook; `sheet_delete` refuses to leave zero visible sheets; `set-visibility` refuses to hide the last visible sheet. No sheet move/reorder. |
| **Excel tables (ListObject)** | List/get tables and read rows; append/insert/update/delete rows; add/delete a column; create/resize. | `table_read` (`list`, `get`, `read-rows`), `table` (`create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column`), `table_delete_rows`, `table_delete_column` | Fully supported | Bounded reads (`offset`/`limit`, `truncated` flag; response capped at 50,000 cells, config-driven `max_table_cells`). Update/delete accept explicit indices or an **unambiguous** `where`; ambiguity is rejected (`AMBIGUOUS_FILTER`), never guessed. |
| **Ranges and cells: values and formulas** | Read and write typed values and formulas. | `range_read` (`get`, `read-values`, `read-formulas`), `range_write` (`write-values`, `write-formulas`) | Fully supported | Reads are bounded windows: 100,000 cells per target, 50,000 cells per response; oversized requests are rejected (`RANGE_TOO_LARGE`), not silently clipped. A string beginning with `=` is stored as **text**, not a formula. |
| **Ranges: copy, insert, delete, merge, unmerge** | Copy a range; insert/delete cells with a shift; merge/unmerge. | `range_structural` (`copy`, `insert`, `unmerge`), `range_delete`, `range_merge` | Partially supported | Same workbook only, no clipboard. `insert`/`delete` require an explicit `shift` (`right`/`down`, `left`/`up`), never guessed. Protected sheets, ListObject intersections (`LISTOBJECT_INTERSECTION`), spill ranges and unmerged cells (unless `allow_merged`) are rejected; merge blocks data loss unless `allow_data_loss=true`. |
| **Formatting and styles** | Read and partially write cell formatting. | `range_read` (`get-format`), `range_write` (`set-format`) | Partially supported | Read is tri-state (value / `mixed` / `unsupported`) and never flattens a mixed range. The write patch exposes 11 fields (number format, font name/bold/italic/size/colour, fill colour/pattern, alignment horizontal/vertical/wrap). No borders, row/column size, or style-name copy. |
| **Data validation** | Read and write data-validation rules. | `range_read` (`get-validation`), `range_write` (`set-validation`) | Partially supported | Types: list, whole, decimal, date, time, text length, custom (`input_only` reported but never set). `Formula1`/`Formula2` are opaque strings capped at 255 characters, never evaluated. A non-uniform target is reported, not flattened. |
| **Conditional formatting** | Read and add/modify/delete conditional-format rules. | `range_read` (`get-conditional-formatting`), `range_write` (`set-conditional-formatting`) | Partially supported | Only plain `FormatCondition` types (cell value, expression, text, blanks, errors, time period, top 10, above average, unique). `ColorScale`, `DataBar`, `IconSet` are reported `supported=false` and rejected on write. `priority` is the identity; priorities are sheet-wide and renumbered on add. |
| **Defined names (named ranges)** | List/get names; create, update, delete in an explicit scope. | `defined_name_read` (`list`, `get`), `defined_name` (`create`, `update`), `defined_name_delete` | Fully supported | Scope is explicit (`workbook` or `worksheet` + owning sheet); duplicates and cross-scope conflicts are rejected; `update` preserves scope. `RefersTo` is opaque (no evaluation). |
| **PivotTables: read, refresh, read result** | Inventory tables/caches; read metadata and fields; refresh a table or cache; **read the rendered result**. | `pivot_read` (`tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters`, `result-read`), `pivot_refresh` (`refresh-table`, `refresh-cache`) | Partially supported | Field operations and refresh need a database/Table source; external / OLAP / Data-Model sources return `unsupported_source` for those. **Displayed-result read (`result-read`) works for all source types**, including external/OLAP/Data-Model. Manual filters only; no create/delete; `Workbook.RefreshAll` is never used. |
| **PivotTables: field layout and filters** | Set a field's orientation/position/aggregation/properties; apply/clear a manual filter. | `pivot_field` (`field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters`) | Partially supported | Layout accepts hidden/row/column/page (`data` rejected — use aggregation). Only eleven aggregation codes are accepted. Filters are manual visible-items only; label/value/date/top-bottom filters are unsupported on the validated build. Page fields cannot be filtered. |
| **Shapes** | List/get drawing shapes; assign or clear a shape's `OnAction` macro. | `shape` (`list`, `get`, `set-on-action`) | Partially supported | Inventory and `OnAction` assignment only; the macro is **not executed**. No create/delete/format/group/z-order. |
| **Controls (Form and ActiveX)** | List/get controls; assign or clear a Form Control's `OnAction`. | `control` (`list`, `get`, `set-on-action`) | Partially supported | `set-on-action` writes `OnAction` on **Form Controls only**; ActiveX is rejected (no `OnAction` semantics) and only reports an inferred event procedure. No create/delete. |
| **Large-data export** | Stream a large range, table body or PivotTable result to a CSV/JSONL file with bounded memory. | `data_export` | Fully supported | Returns metadata only (path, rows, bytes, SHA-256, completeness); `limit_rows` may truncate. The MCP-transport run is validated in-process only (not over a live client transport). |
| **Environment, health and recovery** | Discover access mode/policy/capabilities; check worker health; recover a wedged owned Excel process. | `xlvba_environment`, `server_health`, `excel_health`, `excel_recover` | Fully supported | `xlvba_environment` and `server_health` never touch Excel; `excel_health` reads state without COM. `excel_recover` needs `confirm=true`, is pid-scoped and originals-safe; `excel.build` is always `null`. |
| **Workbook creation (new workbook)** | — | — | Not supported | No tool creates a new workbook; `workbook` `open` opens an existing file only. |
| **PivotTable / PivotCache create or delete** | — | — | Not supported | Only existing PivotTables and caches can be inspected, refreshed and reconfigured. |
| **Charts, PivotCharts and slicers** | — | — | Not supported | No tool lists, creates, formats or deletes them. A chart sheet appears only as a `chart` entry in `sheet_read` `list`. |
| **Row/column sizing and layout** | — | — | Not supported | No tool exposes row height, column width, freeze panes, print/page setup or zoom. |
| **Worksheet protection changes** | — | — | Not supported | Protection is read-only diagnostics (`protected` flag); a protected target is rejected with `PROTECTED_TARGET`. |
| **Workbook calculation** | — | — | Not supported | No `Application.Calculate`; `range_read` `read-values` returns whatever Excel currently holds — nothing forces a recalculation. |
| **External links, connections, queries** | — | — | Not supported | No tool lists, refreshes or edits external links, data connections or Power Query definitions. |
| **Import, export, print, PDF, email** | — | — | Not supported | No CSV import, PDF export, print, email or file-conversion tool. `vba_export_project` exports VBA source only. |
| **Events and event handlers** | — | — | Not supported | No tool registers, lists or edits workbook/worksheet events. |
| **Non-`.xlsm` Excel formats** | Open and edit other Excel formats. | `workbook` `open` and all workbook tools | Not verified | The path allowlist accepts `.xlsm .xlsb .xlsa .xlam .xltm .xltx .xlsx .xls`, but only `.xlsm` is validated. Behaviour with other extensions is not verified. |
| **ActiveX event mapping** | Read a control's probable event procedure. | `control` (`list`, `get`) | Not verified | Inferred from `<sheet_code_name>_<control>_<Event>` naming; never click-tested and may be absent or wrong. |

## Formula support (separate from the Excel object model)

Excel formulas such as `SUM`, `VLOOKUP`, `XLOOKUP` or `FILTER` are **not** an `xlvba` capability in
the object-model sense. They are worksheet content that `xlvba` can read and write as strings, but
`xlvba` is not a calculation engine.

Three distinct operations must not be confused:

| Operation | Tool (action) | What it touches | What it returns |
|-----------|---------------|-----------------|-----------------|
| Read formula text | `range_read` (`read-formulas`) | `Range.Formula2` (primary) and `Range.Formula` (fallback) | A bounded grid of formula cells. `formula2` is always present; `formula` is included only when it differs (for example an implicit-intersection `@` added by Excel). Prefixes `_xlfn`, `_xlws`, `_xlpm` are preserved verbatim and flagged in `notes`. |
| Write formula text | `range_write` (`write-formulas`) | `Range.Formula2` (fallback `Range.Formula`) | A mutation result with the re-read formula grid and a `verified` flag. Every entry must begin with `=`; a bare constant is rejected. |
| Read calculated values | `range_read` (`read-values`) | `Range.Value2` | Typed cells: `string`, `number`, `boolean`, `date`, `error`, `empty`. Dates are typed only from the cell number format. |

Explicit boundaries:

1. `xlvba` does **not** provide a calculation engine. It never evaluates a formula and never calls
   `Application.Evaluate`.
2. `xlvba` does **not** expose `Application.Calculate` (or `CalculateFull`, `CalculateRow` or any
   calculation property). There is no tool to force a recalculation.
3. `xlvba` does **not** guarantee recalculated values after a write. `range_write` `write-formulas`
   writes formula strings; a later `range_read` `read-values` returns whatever Excel currently holds.
   If Excel has not recalculated, the cached value may be stale or empty.
4. Dynamic-array formulas are read in their `Formula2` form. A spill range is a protected target for
   structural operations: `range_structural` `insert`, `range_delete` and `range_merge` reject it.

Worked examples are in [Workflows and examples](07-workflows-and-examples.md).

## Prerequisites and good practice

These apply across domains and are the most common reasons a call is refused:

- **Open first.** Every tool that takes a `workbook_id` needs an open working copy. Start with
  `workbook` (`open`) and keep the returned `workbook_id`.
- **Read-only for inspection.** Use `mode="inspect", read_only=True` for any workflow that does not
  need mutations or macros. It disables macros and makes accidental mutation impossible
  (`WORKBOOK_READ_ONLY`).
- **Snapshot before a mutation.** Call `snapshot_create` after opening writable and before the first
  mutating call; otherwise every mutation fails with `SNAPSHOT_FAILED`.
- **Read before you write.** Use `vba_read` (`read-component`) for the `source_hash`,
  `table_read` (`get`) for the exact table name, `pivot_read` (`table-get` / `fields-list`) for
  PivotTable and field names, and `defined_name_read` before touching a name.
- **One object per mutation, then re-read.** Change one table, range, sheet or field at a time and
  inspect the returned postcondition (`verified`, `postcondition_values`, `affected_rows`).
- **Never retry a blocked call blindly.** A blocked macro or a denied mutation will be blocked again.
  Read `context.reason` and change the precondition instead.
- **VBA needs trust.** Enable **Trust access to the VBA project object model** in Excel Trust Center
  → Macro Settings; corporate policy may block it.

## Per-domain notes

### Macros and interactive UI

`macro_run` runs a procedure only when it is classified `background_safe` for the current workbook
hash **and** VBA source revision, and passes the interactive-UI preflight. A procedure containing
`MsgBox`, `InputBox`, `UserForm.Show`, `FileDialog` or any other interactive token is vetoed
**before** dispatch with `INTERACTIVE_UI_NOT_SUPPORTED` — the macro is not started. This is expected:
the server owns a hidden, non-activating Excel instance with no GUI, so a modal would block the single
worker indefinitely. The R13 scan is a heuristic veto, not static proof; indirect UI can still evade
it, so review VBA before classifying it. A macro hang or timeout terminates only the owned Excel
process and recovers the session in-process; open workbooks are discarded and the macro is never
auto-replayed.

### PivotTables and external sources

Inventory, field configuration, filters and refresh work for **database/Table-backed** PivotTables.
For **external / OLAP / Data-Model** sources those operations return `unsupported_source`, but
**reading the displayed result** (`pivot_read` `result-read`) works for every source type — use it to
export or inspect what the PivotTable currently shows. Only manual visible-items filters are
supported; label/value/date/top-bottom filters are rejected on the validated build. `xlvba` never
creates or deletes PivotTables/PivotCaches and never calls `Workbook.RefreshAll`.

### Compile diagnostics are honest, not a verdict

`vba_compile_diagnostics` is a **diagnostic**, not a compile. On the validated build it always
returns `status="indeterminate"` with `is_compile=false` and never claims `compiled` or
`compile_error`. Treat a clean `indeterminate` result as "unknown", not "passed". To actually compile,
open the workbook in the Excel VBA editor and use **Debug → Compile VBAProject**.

### Large data and the single worker

Large reads are paged (`offset`/`limit`) and bounded (100,000 cells per range target, 50,000 cells
per response, 50,000 cells per table response). A 200,000 × 20 range, for example, is read in about
105 pages. Because there is one Excel instance, a very large read, export or macro **delays** other
operations; prefer `data_export` to stream a large range/table/pivot result to a file, and
`workbook_snapshot_read` to read saved snapshot data without occupying the worker. The MCP-transport
run of `workbook_snapshot_read` and `data_export` is validated in-process only.

## Not supported and why

| Missing capability | Reason |
|--------------------|--------|
| Create or delete PivotTables / PivotCaches | Out of scope by design. Only **existing** PivotTables and caches are managed; external/OLAP/Data-Model sources are rejected for field/refresh. |
| Charts, PivotCharts, slicers | Not implemented. A chart sheet is only a `chart` entry in `sheet_read` `list`. |
| Label/value/date/top-bottom PivotTable filters | Unsupported on the validated build; `pivot_field` `field-apply-filter` returns an explicit unsupported result rather than silently converting to a manual filter. |
| New workbook creation | No tool. `workbook` `open` opens an existing file only. |
| Range structural operations over a ListObject | Rejected with `LISTOBJECT_INTERSECTION` and routed to the table tools — a deliberate ownership boundary, not a missing feature. |
| Worksheet or workbook protection changes | No tool. A protected target is rejected with `PROTECTED_TARGET`. |
| Forced recalculation | `Application.Calculate` is not exposed; formula writes do not guarantee a recalculated cached value. |
| Clipboard and cross-workbook copy | `range_structural` `copy` is same-workbook and never uses the clipboard. |
| Rows/columns sizing, freeze panes, page setup | No tool. |
| External links, connections, queries, import/export, print, PDF, email | No tool. |
| Events | No tool. |
| Approval grant / allowlist management | Deliberately not exposed over MCP. In the default `full` mode no grant is needed; `restricted` mode re-enables deny-by-default. `background_safe` classification **is** exposed through `macro` (`classify-background`). |
| GUI automation, dialogs, keystrokes | Forbidden by design. The instance is hidden and non-activating. |
| Server-side log-level control | `EXCEL_VBA_MCP_LOG_LEVEL` is parsed but not applied; logging is fixed at `INFO` on `stderr`. |

## Coverage summary

The capability matrix contains **35 capability rows** across the **43-tool surface**.

| Coverage | Rows | What they cover |
|----------|------|-----------------|
| Fully supported | 11 | Workbook sessions and safe copies; Save As copy; snapshots; non-COM snapshot read; jobs/cancel; VBA read/export; tables; range values/formulas; defined names; large-data export; environment/health/recovery |
| Partially supported | 13 | VBA editing; compile diagnostics; macros; declarative testing; sheets; range structure; formatting; data validation; conditional formatting; PivotTable read/refresh/result; Pivot field configuration/filters; shapes; controls |
| Not supported | 9 | Workbook creation; PivotTable/PivotCache create-delete; charts/PivotCharts/slicers; row/column sizing and layout; worksheet protection; workbook calculation; external links/connections; import/export/print/PDF/email; events |
| Not verified | 2 | Non-`.xlsm` formats; ActiveX event mapping |
| **Total** | **35** | — |

Read-only coverage is broad: every inventory and read surface above is available over MCP without
additional configuration, provided the path is allowed and the workbook can be opened. In the default
`full` access mode mutating and executing coverage is also reachable over MCP (a baseline snapshot is
required before a mutation); `EXCEL_VBA_MCP_ACCESS_MODE=restricted` reverts to deny-by-default.
Compilation remains `indeterminate`, and macro execution requires a `background_safe` classification
(via `macro` `classify-background`) and passes the interactive-UI veto.

Related documents: [Overview](01-overview.md), [Configuration](03-configuration.md),
[Architecture](04-architecture.md), [MCP tools reference](05-mcp-tools-reference.md),
[Workflows and examples](07-workflows-and-examples.md),
[Security](08-security.md), [Limitations and compatibility](09-limitations-and-compatibility.md)
and the [root README](../README.md).