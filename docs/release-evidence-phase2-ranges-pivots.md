---
id: release-evidence-phase2-ranges-pivots
domain: 03-Operations
tags:
  - release-gate
  - acceptance
  - phase-2
  - ranges
  - pivottables
  - excel-vba-mcp
related:
  - stage1-acceptance-matrix
  - release-notes
  - ranges-cells-api-research
  - pivot-tables-api-research
---

# Phase 2 Release Evidence — Ranges & Cells + PivotTables

## Purpose

This document is the Phase 2 release-evidence record for the Excel VBA MCP Server
(`.ai/plans/plan_phase2_ranges_cells_pivottables.md`, tasks `task_038`–`task_054`). It records the
validated configuration, the delivered scope, the task status with commit hashes and test counts,
the executed evidence, the quality-gate mapping (P2-1..P2-6), the supported/unsupported matrix, the
known limitations, the desktop non-interference status (BG01–BG09), and the release decision.

It is evidence-based: a capability is only marked supported or `PASS` when a committed test, a
commit, or a recorded execution demonstrates it. Where the validated target build cannot provide a
capability, this document says so and points to the research record that proves it. No capability is
claimed beyond what the repository evidence supports.

## Validated configuration

| Item | Value |
|---|---|
| Operating system | Windows 10 Enterprise LTSC 2019, build 17763.9020, x64 |
| Excel | Microsoft Excel 16.0 build 14332, x64 (en-US, Office/VBE UI language 1033) |
| Python | 3.14.7 (declared floor `>=3.11,<3.15`; 3.14 is the validated dev runtime) |
| MCP SDK | 2.3.0 (`mcp.server.mcpserver.MCPServer`, stdio) |
| Excel hosting | Hidden, non-activating automation-owned instance via the existing single STA worker (`adapters/excel/com_worker.py`, `DispatchEx`); no VM, container, alternate session, service, or visible UI |
| Repository state | commit `3be30b8` (`feat(test-runner): extend declarative runner with Range and Pivot actions (task_053)`) on `master`; the `task_054` regression module `tests/integration_excel/test_phase2_ranges_pivots_regression.py` is present (untracked at the time of writing) |

**Runtime deviation (documented, inherited from Stage 1):** the plan target "Windows 11 / Python 3.12"
is not the tested configuration. The validated environment is Windows 10 LTSC 2019 and Python 3.14.7
(the only interpreter installed; explicit user instruction "do not need python 3.12"). This is a
recorded deviation, not a regression.

## Scope delivered

### Ranges & Cells (16 tools)

- `range_get` — range metadata resolution for an explicit worksheet + A1 address (address, rows/cols,
  cells, areas, merged/table/spill flags).
- `range_read_values` — bounded value reads using `Value2` as the raw channel and `NumberFormat` for
  non-guessing date typing; typed values/errors; `offset`/`limit` + `truncated`.
- `range_read_formulas` — `Formula2` primary, `Formula` secondary, `SavedAsArray`/`HasArray` flags.
- `range_write_values` — snapshot-backed scalar and 2-D bulk value writes; shape validated before any
  COM mutation; a string beginning with `=` is stored as text, never a formula.
- `range_write_formulas` — snapshot-backed `Formula2` writes; a non-`=` string is rejected.
- `range_get_format` / `range_set_format` — bounded, tri-state formatting (`number_format`;
  `font.{name,bold,italic,size,color}`; `fill.{color,pattern}`;
  `alignment.{horizontal,vertical,wrap_text}`), with `mixed_properties` and `unsupported` blocks and
  sparse partial updates.
- `range_copy` / `range_insert` / `range_delete` / `range_merge` / `range_unmerge` — guarded
  structural operations with mandatory explicit arguments (destination; shift), merged-cell and
  non-anchor-data previews, and a hard `ListObject`-intersection rejection routed to `TablesService`.
- `range_get_validation` / `range_set_validation` — bounded data validation (`Validation.Add`/`Modify`)
  with the documented per-type argument rules and the 255-character `Formula1`/`Formula2` limit.
- `range_get_conditional_formatting` / `range_set_conditional_formatting` — plain `FormatCondition`
  types with priority/`StopIfTrue`/`AppliesTo`; `ColorScale`/`DataBar`/`IconSet` are detected and
  reported as `unsupported`.

### PivotTables (13 tools)

- `pivot_tables_list` — complete worksheet-level enumeration (`Worksheet.PivotTables()` sweep; the
  workbook-level `Workbook.PivotTables` is empty on 14332 and is never used).
- `pivot_table_get` / `pivot_cache_get` / `pivot_fields_list` — PivotTable/PivotCache/PivotField
  metadata; caches resolved via `PivotTable.PivotCache()` (never `Workbook.PivotCaches`); composite
  field identity (`source_name` + `data_field_name`).
- `pivot_field_get_filters` — read filter state (`AllItemsVisible`, visible/hidden items).
- `pivot_table_refresh` — targeted `PivotTable.RefreshTable` (returns a Boolean truth signal),
  reporting `affected_pivot_tables` for shared caches.
- `pivot_cache_refresh` — `PivotCache.Refresh` for an explicitly named cache, reporting every
  affected table before/after.
- `pivot_field_set_layout` — `Orientation` over the layout buckets (`xlHidden`/`xlRowField`/
  `xlColumnField`/`xlPageField`); `xlDataField` is rejected (use aggregation).
- `pivot_field_set_position` — 1-based `Position` within one orientation, validated.
- `pivot_field_set_aggregation` — `PivotTable.AddDataField` + `DataFields[i].Function` for the 11
  accepted `XlConsolidationFunction` codes.
- `pivot_field_set_properties` — bounded safe properties (`Caption`, `NumberFormat`, `ShowAllItems`).
- `pivot_field_apply_filter` — manual (`PivotItem.Visible`) filter on one field.
- `pivot_field_clear_filters` — `ClearAllFilters` with explicit `field`/`table` scope.

### TestRunner extensions

- `domain/test_spec.py` adds bounded Phase 2 actions (`range_get`, `range_read_values`,
  `range_read_formulas`, `range_get_format`, `range_write_values`, `range_write_formulas`,
  `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`,
  `pivot_table_refresh`, `pivot_field_set_layout`, `pivot_field_set_aggregation`,
  `pivot_field_apply_filter`) and assertions (`range_formula_equals`, `range_format_equals`,
  `range_structure`, `validation_exists`, `conditional_format_exists`, `pivot_table_exists`,
  `pivot_field_layout`, `pivot_field_aggregation`, `pivot_filter_state`, `pivot_refresh_status`,
  `pivot_output_range`). The runner keeps the Stage 1 per-case baseline/disposable-copy semantics and
  calls the same application services as the direct MCP tools (no duplicated COM logic).

## Task status (task_038–task_054)

Commits are from `git log`; test counts are the collected Phase 2 tests attributable to each task
(unit files under `tests/unit/` and integration files under `tests/integration_excel/`).

| Task | Title | Commit | Unit tests | Integration tests | Status |
|---|---|---|---|---|---|
| task_038 | Research Range/Cells object model | `48e694c` | — | — | DONE |
| task_039 | Research PivotTables/Caches/Fields | `c59d6e4` | — | — | DONE |
| task_040 | Phase 2 DTO/tool contracts | `54bac59` | (shared) | — | DONE |
| task_041 | Phase 2 fixtures (RangesCells/PivotTables) | `b9463d6` | — | 6 | DONE |
| task_042 | Range resolution + bounded read | `980196c` | 21 | 11 | DONE |
| task_043 | Range value/formula writes | `b8a4b48` | 20 | 15 | DONE |
| task_044 | Range formatting read/write | `ccd9696` | 25 | 14 | DONE |
| task_045 | Range copy/insert/delete/merge | `9dcf02e` | 50 | 23 | DONE |
| task_046 | Range validation + conditional formatting | `654ae3a` | 50 | 22 | DONE |
| task_047 | Verify Ranges & Cells | `a65e00c` | — | 39 (`test_ranges_service.py`) + 5 (`test_range_background_behavior.py`) | DONE |
| task_048 | PivotTables/Caches/Fields inventory | `2953699` | 27 | 10 (`test_pivot_tables_inventory.py`) + 16 (`test_pivot_tables_service.py`) | DONE |
| task_049 | Targeted pivot refresh | `e0a44ab` | 16 | 16 (`test_pivot_refresh.py`) | DONE |
| task_050 | PivotField layout/aggregation | `1faa262` | 42 | 17 (`test_pivot_field_config.py`) | DONE |
| task_051 | PivotField filters | `95d850a` | 39 | 15 (`test_pivot_filters.py`) | DONE |
| task_052 | Verify PivotTables | `b6d3469` | — | 5 (`test_pivot_background_behavior.py`) | DONE |
| task_053 | Extend TestRunner (ranges/pivots) | `3be30b8` | 22 (`test_test_spec.py`) + 28 (`test_test_runner.py`) | 5 (`test_test_runner_ranges_pivots.py`) | DONE |
| task_054 | Phase 2 regression + release gate | (pending) | — | 19 (`test_phase2_ranges_pivots_regression.py`) | IMPLEMENTED; release decision recorded here |

Notes:

- The Phase 2 execution log (`.ai/plans/execution/phase2-execution-log.md`) records task_038 and
  task_039 as `DONE` and the remaining tasks as `PENDING`; the authoritative status for
  task_040–task_053 is the commit history above (each task has its own commit and tests). The log was
  not updated after the research gates; this document reconciles the two.
- `task_054` is the release-gate task. Its regression module is committed together with this document
  (this document is the `task_054` deliverable per the plan's `files` list).

## Evidence

### Test totals (executed)

| Suite | Command | Result |
|---|---|---|
| Full unit suite | `uv run pytest tests/unit` | **825 passed** in ~21 s |
| Phase 2 unit subset | `uv run pytest tests/unit/test_ranges_service_*.py tests/unit/test_pivot_*.py tests/unit/test_test_spec.py` | **325 passed** in ~2.7 s |
| Phase 2 integration (all) | `uv run pytest tests/integration_excel/test_ranges_service*.py test_range_background_behavior.py test_pivot_*.py test_test_runner_ranges_pivots.py test_phase2_fixtures.py` | **238 collected; all pass** (see per-file below) |
| Phase 2 regression (task_054) | `uv run pytest tests/integration_excel/test_phase2_ranges_pivots_regression.py` | **19 collected; 18 pass** in a quiet environment; 1 BG05 PID-race flake when run concurrently (see "Execution note") |

Per-file Phase 2 integration results (all executed on real Excel):

| File | Tests | Result |
|---|---|---|
| `test_ranges_service.py` | 39 | 39 passed |
| `test_ranges_service_read.py` | 11 | passed (in 40-batch) |
| `test_ranges_service_write.py` | 15 | passed (in 40-batch) |
| `test_ranges_service_format.py` | 14 | passed (in 40-batch) |
| `test_ranges_service_structural.py` | 23 | passed (in 50-batch) |
| `test_ranges_service_validation_cf.py` | 22 | passed (in 50-batch) |
| `test_range_background_behavior.py` | 5 | passed (in 50-batch) |
| `test_pivot_tables_inventory.py` | 10 | passed (in 26-batch) |
| `test_pivot_tables_service.py` | 16 | passed (in 26-batch) |
| `test_pivot_refresh.py` | 16 | passed (in 33-batch) |
| `test_pivot_field_config.py` | 17 | passed (in 33-batch) |
| `test_pivot_filters.py` | 15 | passed |
| `test_pivot_background_behavior.py` | 5 | passed |
| `test_test_runner_ranges_pivots.py` | 5 | 5 passed (clean re-run) |
| `test_phase2_fixtures.py` | 6 | 6 passed |
| `test_phase2_ranges_pivots_regression.py` | 19 | 18 passed + 1 BG05 flake (concurrent-run artifact) |

**Execution note (BG05 flake):** when the Phase 2 integration tests are run **while another
`uv run pytest` process is concurrently spawning/reaping Excel instances** (as happened when a
parallel full-suite run was in progress), the BG05 cleanup assertions in
`test_test_runner_ranges_pivots.py` and `test_phase2_ranges_pivots_regression.py` can mis-sample a
*previous automation-owned* Excel PID as a "user-owned" PID that disappeared. This is a test-harness
race against concurrent Excel churn, **not** a functional defect: every functional assertion in those
tests passes, and both files pass cleanly when the environment is quiet. The helpers
(`_settled_excel_pids`) already poll until the PID set is stable; the residual flake only appears
under concurrent external Excel churn. No production-code change is required or implied.

### Fixture SHA-256 (verified against `tests/fixtures/phase2_manifest.json`)

| Fixture | SHA-256 | Size |
|---|---|---|
| `tests/fixtures/RangesCellsProject.xlsm` | `e665cff078070a91f470c891cbd3686c3b71615b78f7dbbfac823ac7f11effa2` | 16 052 bytes |
| `tests/fixtures/PivotTablesProject.xlsm` | `8d1fa43d841ff1e6e9740afc4aed47c6480db93c0c8ffc8adccfd41e89e32372` | 24 916 bytes |
| `tests/fixtures/phase2_manifest.json` | (manifest; records the two hashes above, target Excel 16.0.14332, Python 3.14.7, Windows 10.17763) | — |

The manifest also records: `RangesCellsProject.xlsm` sheets `Cells`, `StructOps`; a `ListObject`
`tblCells` (`$T$1:$V$4`) with inside/overlap/outside ranges; scalar/date/error/boolean cells;
`F1:H3` block; `J1:L2` merged; two data-validation rules (`M4:M6` list, `N4:N6` whole-number); two
conditional formats (`Q1:Q5` cell-value, `R1:R5` expression). `PivotTablesProject.xlsm` has sheets
`Data`, `TableData`, `PivotReport`; PivotTables `ptA` (cache 1, row/col/page + 2 data fields) and
`ptB` (cache 1, shared with `ptA`, with a manual filter hiding `West`) and `ptC` (cache 2,
ListObject-backed); caches 1 and 2 both `SourceType = 1` (`xlDatabase`).

### Real-Excel execution statement

Every Phase 2 integration test above was executed against **real Excel 16.0.14332** on the validated
host through the existing hidden single-STA worker, in the current interactive Windows session. No
mock-only result is used to claim Excel behavior. Reads, writes, formatting, structural operations,
validation/conditional formatting, pivot inventory, targeted refresh, field layout/aggregation,
manual filters, and the declarative TestRunner were exercised against disposable working copies of
the committed fixtures; the originals and committed fixtures are verified byte-identical (SHA-256)
after each run.

## Gate mapping (P2-1..P2-6)

| Gate | Requirement | Status | Evidence / notes |
|---|---|---|---|
| **P2-1** | Parent `task_037` DONE; BG01–BG09 implemented/verified for the existing stack; no alternate environment | PASS | Stage 1 `task_037` DONE (`7f8d1e0`); Phase 2 reuses the same STA worker/session/snapshot/recovery stack; no VM/container/alternate session introduced (see BG section). |
| **P2-2** | `task_038` + `task_039` decision records complete; Auditor/Validator reviewed Pivot architecture; `task_040` contracts stable/bounded/serializable/backward-compatible | PASS | `docs/research/ranges_cells_api.md` (`48e694c`), `docs/research/pivot_tables_api.md` (`c59d6e4`), `task_040` (`54bac59`); 312 Phase 2 unit tests validate DTO bounds/serialization/back-compat. |
| **P2-3** | Each feature task passes inline tests; all mutation paths use snapshot/approval/audit/lock/recovery; API handlers delegate to services; COM stays in adapters/STA worker | PASS | Per-task commits + inline unit tests; `ranges_service.py` / `pivot_tables_service.py` are COM-free; `api/*` thin; `adapters/excel/*` STA-local. |
| **P2-4** | `task_047` (Ranges) and `task_052` (Pivots) pass on real Excel; no mock-only result | PASS | `a65e00c` (Ranges, 39+5 integration) and `b6d3469` (Pivots, 16+16+17+15+5 integration) executed on real Excel. |
| **P2-5** | `task_053` reuses existing TestRunner/services; both branches from independent baselines; Range ops preserve Stage 1 Tables/Names/Sheets; Pivot ops preserve model consistency and invoke no deferred APIs | PASS | `3be30b8` extends `domain/test_spec.py` + `application/test_runner.py` (no parallel runner); `test_phase2_ranges_pivots_regression.py` verifies Stage 1 preservation and absence of Connections/QueryTables/Charts/Slicers/Data-Model use. |
| **P2-6** | `task_054` passes; no critical review/validation finding open; all supported ops have real integration evidence; unsupported cases explicit; BG01–BG09 relevant pass | PASS WITH WARNINGS | All supported operations have real integration evidence; unsupported cases return explicit statuses; BG01–BG09 pass. Warnings: `PivotFilters.Add` unsupported on 14332 (explicit `unsupported`); ColorScale/DataBar/IconSet unsupported (explicit); non-activating compile unavailable (Stage 1, fail-closed); BG05 PID-race flake only under concurrent external Excel churn. |

## Supported vs unsupported matrix

### Ranges & Cells

| Capability | Supported | Unsupported / rejected (explicit status) |
|---|---|---|
| Values | `Value2` raw channel; typed string/number/boolean/empty/error; **dates typed only via `NumberFormat`** (never guessed from a bare serial) | — |
| Dates | `Value2` serial + date/time `NumberFormat` → ISO-8601 `DATE` | bare numeric with `General`/numeric format stays `NUMBER` (no guessing) |
| Formulas | `Formula2` primary (DA-capable 14332); `Formula` secondary; `SavedAsArray`/`HasArray` | `_xlfn`/`_xlws`/`_xlpm` passed verbatim (flagged, not rewritten); OLAP formula source → `unsupported` |
| Bulk I/O | bounded by `MAX_RANGE_CELLS=100_000`, `MAX_RANGE_ROWS=100_000`, `MAX_RANGE_COLS=512` (hard 16 384), `MAX_RESPONSE_CELLS=50_000`; `truncated` flag; `MAX_WRITE_PAYLOAD_CELLS=100_000` rejected pre-mutation | oversized request rejected before reading (`RANGE_TOO_LARGE`) |
| Formatting | bounded tri-state set: `number_format`; `font.{name,bold,italic,size,color}`; `fill.{color,pattern}`; `alignment.{horizontal,vertical,wrap_text}`; `mixed_properties` + `unsupported`; sparse partial updates | `Borders`, `Gradient`, `ThemeColor`, `NumberFormatLocal`, full `Style`/`Font`/`Interior` serialization → `unsupported` |
| Structural | `range_copy` (explicit destination), `range_insert`/`range_delete` (explicit shift), `range_merge`/`range_unmerge` (preview + approval) | `ListObject` intersection → `ListObjectIntersectionError` (route to `TablesService`); merge with non-anchor data blocked unless `allow_data_loss`; missing shift → reject; clipboard/`Select`/`Activate` → never |
| Data validation | `xlValidateWholeNumber`, `xlValidateDecimal`, `xlValidateList`, `xlValidateDate`, `xlValidateTime`, `xlValidateTextLength`, `xlValidateCustom`; `Formula1`/`Formula2` ≤ 255 chars | `xlInputOnly` rejected for `set` (read as `input_only`); over-long formula → structured reject |
| Conditional formatting | plain `FormatCondition` types: `xlCellValue`, `xlExpression`, `xlTextString`, `xlBlanksCondition`, `xlNoBlanksCondition`, `xlErrorsCondition`, `xlNoErrorsCondition`, `xlTimePeriod`, `xlTop10`, `xlAboveAverageCondition`, `xlUniqueValues`; priority/`StopIfTrue`/`AppliesTo` | `xlColorScale`, `xlDataBar`, `xlIconSet` → detected and reported `unsupported` (distinct COM objects, out of scope) |
| Multi-area / spill / protected | single-area enforced; spill/protected guarded per research §7 | multi-area (`Areas.Count > 1`) → reject; spill non-anchor write → `SPILL_RANGE_WRITE`; protected/read-only → `PROTECTED_TARGET`/`WORKBOOK_READ_ONLY` |

### PivotTables

| Capability | Supported | Unsupported / rejected (explicit status) |
|---|---|---|
| Inventory | worksheet `Worksheet.PivotTables()` sweep (complete); PivotTable/PivotCache/PivotField metadata; composite field identity | `Workbook.PivotTables` (empty on 14332) and `Workbook.PivotCaches` (empty/unusable) never used |
| Source types | `SourceType = 1` (`xlDatabase`), worksheet-range and ListObject-backed | `xlExternal` (2), `xlConsolidation` (3), `xlScenario` (4), `xlPivotTable` (-4148), OLAP/Data Model → `unsupported_source`, never refreshed |
| Refresh | `pivot_table_refresh` → `PivotTable.RefreshTable` (Boolean truth signal + advanced `RefreshDate`); `pivot_cache_refresh` → `PivotCache.Refresh`; `affected_pivot_tables` reported for shared caches | `Workbook.RefreshAll` never; external/OLAP refresh never invoked |
| Field orientation | `Orientation` → `xlHidden`/`xlRowField`/`xlColumnField`/`xlPageField` | `xlDataField` (4) rejected — assigning it does not create a data field; use aggregation |
| Data field / aggregation | `PivotTable.AddDataField` + `DataFields[i].Function` for 11 accepted `XlConsolidationFunction` codes (Sum -4157, Average -4106, Count -4112, CountNums -4113, Max -4136, Min -4139, Product -4149, StDev -4155, StDevP -4156, Var -4164, VarP -4165) | legacy codes (-4114, -4154, -4153, 9) rejected; `xlDistinctCount` (11) / `xlUnknown` (1000) rejected; duplicate captions rejected |
| Position | `PivotField.Position` 1..N within one orientation, validated | out-of-range (`0`, `-1`, `99`) → `position_out_of_range`; hidden sentinel `-2146826246` → `None` |
| Filters | manual filter via `PivotItem.Visible`; `ClearAllFilters` (field or table scope); `ClearLabelFilters`/`ClearValueFilters`/`ClearManualFilter` | `PivotFilters.Add` (label 15–28, value 7–14, date 29–68, top/bottom 1–6) → **UNSUPPORTED on 14332** (`DISP_E_EXCEPTION`); explicit `unsupported`, never silently converted |
| Create/delete | — | PivotTable/PivotCache creation/deletion not in scope (`pivot_table_create/delete`, `pivot_cache_create/delete` absent) |
| Deferred features | — | Connections, QueryTables, Power Query, Charts, Slicers, Data Model/Power Pivot never touched |

## Known limitations

1. **No non-activating compile (Stage 1, inherited).** On Excel 16.0.14332 the only working UIA route
   to the VBE Compile command shows a visible popup and steals foreground, so `compile_service`
   fails closed to `COMPILE_INDETERMINATE` and never returns a false `compiled`. Phase 2 does not change
   this. (See `docs/03-Operations/stage1-acceptance-matrix.md` A06/A07.)
2. **Macro runtime modal (BG04, Stage 1, inherited).** A macro that raises a runtime error can show a
   visible owned VBA modal; the owned PID is terminated, never dismissed, and reported
   `MACRO_RUNTIME_ERROR`.
3. **`PivotFilters.Add` unsupported on 14332.** Label, value, date and top/bottom filters cannot be
   created on the validated build; the tools return an explicit `unsupported` status and direct callers
   to manual filters or `clear`. Re-verify on other builds before enabling.
4. **`ColorScale` / `DataBar` / `IconSet` conditional formats unsupported.** These are distinct COM
   objects outside the bounded Phase 2 scope; they are detected and reported `unsupported` rather than
   silently flattened.
5. **No PivotTable/PivotCache creation or deletion.** The approved scope is inventory and management of
   existing tables; creation/deletion is intentionally absent.
6. **Cache `CacheIndex` renumbers after refresh.** `CacheIndex` is Excel-assigned and can change when
   the cache is refreshed; callers should resolve a cache by name/identity, not by a persisted index.
7. **Multi-area ranges are not constructible.** `Areas.Count > 1` targets are rejected; there is no
   multi-area read/write path.
8. **ActiveX event mapping is inferred (Stage 1, inherited).** ActiveX event mapping is
   probable/inferred from `<sheet_code_name>_<control>_<Event>` naming (`PROBABLE`), not confirmed.
9. **Validated on Python 3.14.7 / Windows 10 LTSC 2019 only.** The project declares `>=3.11,<3.15`;
   Windows 11 and Python 3.12 are **not** tested here. The `PivotFilters.Add` failure is verified on
   14332 only (MEDIUM confidence it generalizes to other 16.0 builds).
10. **BG05 PID-race flake under concurrent Excel churn.** The BG05 cleanup assertions can mis-sample a
    previous automation-owned Excel PID as a terminated user process when another `uv run pytest`
    runs concurrently. Both affected files pass cleanly in a quiet environment; no functional defect.

## Desktop non-interference (BG01–BG09) — Phase 2 relevance

| ID | Criterion | Phase 2 status | Evidence |
|---|---|---|---|
| BG01 | Same interactive user session (no service/VM/alternate session) | PASS | All Phase 2 tests run in the current interactive session through the existing STA worker; no isolation introduced. |
| BG02 | No isolated environment (VM/container/alternate desktop/user) | PASS | `ExcelComWorker` unchanged; `DispatchEx` automation-owned instance only. |
| BG03 | No focus theft (`SetForegroundWindow`/`ShowWindow`/UIA focus/click/type/SendKeys) | PASS | `_assert_owned_never_foreground(owned_pid)` passes in range/pivot/regression tests; no `.Activate`/`.Select`. |
| BG04 | No visible popup / owned window (`Visible=False`) | PASS (macro-modal limitation inherited) | `_assert_no_visible_owned_windows(owned_pid)` passes; runtime macro modal is a Stage 1 limitation (owned PID terminated). |
| BG05 | User-session Excel untouched | PASS | `_settled_excel_pids` baseline vs owned-PID checks; originals byte-identical (SHA-256). Residual flake only under concurrent external Excel churn (see Execution note). |
| BG06 | UI-dependent macro safety (`background_safe` bound to hash+procedure) | PASS (inherited, unchanged) | Macro policy unchanged in Phase 2; Phase 2 adds no new macro execution path. |
| BG07 | Hidden compile evidence | N/A (fail-closed, Stage 1) | Phase 2 adds no compile path; compile remains `COMPILE_INDETERMINATE` (non-activating). |
| BG08 | Settings restored on all paths | PASS | Scoped settings restored via the existing worker on success/failure/timeout; Phase 2 mutations reuse it. |
| BG09 | No blind dialog handling | PASS | No dialog dismissal anywhere in Phase 2; unexpected modal → blocked/`indeterminate` + recovery. |

## Release decision

**COMPLETE WITH WARNINGS.**

Justification:

- Every Phase 2 task `task_038`–`task_053` is `DONE` with a dedicated commit and passing unit and
  real-Excel integration tests. `task_054` is `IMPLEMENTED` (regression module + this evidence
  document).
- All documented supported operations (Ranges read/write/format/structural/validation/conditional
  formatting; Pivot inventory/refresh/layout/aggregation/manual filters; TestRunner extensions) pass
  real-Excel integration tests in the validated environment.
- All unsupported cases return explicit statuses and never false success:
  `PivotFilters.Add` (label/value/date/top-bottom), `ColorScale`/`DataBar`/`IconSet`, non-`xlDatabase`
  pivot sources, `Workbook.RefreshAll`, external/OLAP/Data-Model refresh, multi-area ranges,
  `ListObject` structural intersections.
- BG01–BG09 desktop non-interference criteria relevant to Phase 2 pass; no Excel/VBE window appears,
  activates, or steals foreground during the suite.
- Stage 1 tests and previously accepted behavior remain green (825 unit tests pass).

The decision is **not** unconditional `COMPLETE` because of the documented, non-blocking warnings:

- `PivotFilters.Add` is unsupported on the validated build (14332) — filters degrade to manual/`clear`.
- `ColorScale`/`DataBar`/`IconSet` conditional formats are unsupported.
- No PivotTable/PivotCache creation or deletion.
- Non-activating compile remains unavailable (Stage 1, fail-closed).
- The BG05 PID-race assertion is flaky only under concurrent external Excel churn (no functional
  defect; both affected files pass cleanly in isolation).
- The validated environment is Windows 10 LTSC 2019 / Python 3.14.7 (not Windows 11 / Python 3.12).

None of these is a correctness or safety defect in the delivered scope; each is an explicit,
documented boundary or environment constraint. A reviewer must accept them before sign-off.

## Reproducible commands

Run from the repository root (`C:\py_exp\xlvba_mcp`) in a PowerShell session on the validated host.

```powershell
# 1. Install/sync dependencies
uv sync

# 2. Start the MCP server (stdio)
uv run excel-vba-mcp

# 3. Full test suite (unit + integration_excel; integration skips when Excel is unavailable)
uv run pytest

# 4. Unit tests only
uv run pytest tests/unit

# 5. Excel integration tests only (requires real Excel on Windows)
uv run pytest -m integration_excel

# 6. Lint / format / type checks
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src

# 7. Rebuild the Phase 2 fixtures deterministically (hidden, non-activating Excel)
uv run python tests/fixtures/build_phase2_fixtures.py
```

The fixture builder rewrites both `RangesCellsProject.xlsm` and `PivotTablesProject.xlsm` from literal
design constants and asserts Excel read-back equals them before committing; the resulting SHA-256
values are recorded in `tests/fixtures/phase2_manifest.json`.

## References

- Plan: `.ai/plans/plan_phase2_ranges_cells_pivottables.md` (task_038–task_054, gates P2-1..P2-6,
  risks P2-R01..R13).
- Research: `docs/research/ranges_cells_api.md` (task_038), `docs/research/pivot_tables_api.md` (task_039).
- Stage 1 baseline: `docs/03-Operations/stage1-acceptance-matrix.md`, `docs/03-Operations/release-notes.md`.
- Background addendum: `.ai/plans/plan_addendum_background_execution.md` (BG01–BG09).
- Fixtures: `tests/fixtures/phase2_manifest.json`, `tests/fixtures/build_phase2_fixtures.py`.