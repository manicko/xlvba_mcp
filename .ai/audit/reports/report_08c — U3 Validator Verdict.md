# Report 08c — U3 Validator Verdict (Schema/Enum Consistency + Truthful Annotations)

- **Role:** Independent Validator for Work Package **U3** (plan_03 §11, §12.2, §14 U3, §17).
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca` (working tree = U1 + U3 changes, uncommitted).
- **Scope:** VALIDATION ONLY. No code/test/config modified. U3 was not implemented by this validator.
- **Method:** source inspection (file:line), `git diff HEAD`, unit suite run, implementor `verify_u3.py`, integration-test static inspection. Integration tests requiring live Excel were **not executed** (no Excel session; marker `requires_excel`).
- **Authority:** plan_03 §11/§12.2/§14-U3/§17; decision_log D-U* > report_08a carve-outs > source at HEAD.

---

## 1. Per-item verdict table

| # | Item | Verdict | Primary evidence |
|---|------|---------|------------------|
| 1 | Three request enums added; full read/classify enums unchanged | **APPROVE** | `statuses.py:443-464, 497-509, 566-576` (request); `418-440, 487-494, 549-563` (full, unchanged) |
| 2 | Request fields retyped; `PivotFilterState` split; read fidelity preserved; identity→value fixes | **APPROVE** | `models.py:981,1203,1224,1230-1242,1253`; `ranges_adapter.py:2743,2886`; adapter/service read constructors migrated |
| 3 | Truthful `ToolAnnotations` on all tools + `server_health`; no `title=` | **APPROVE** | `server.py:1733-1827,2489`; 91/91 annotated; `macro_run`/`vba_compile` not read-only |
| 4 | 3 enum-tool descriptions state pruned values not accepted, match schema | **APPROVE** | `server.py:763-774, 899-911, 954-966` |
| 5 | Pruned values rejected at runtime (pydantic); read tools still report `supported=false` | **APPROVE** | unit tests `test_ranges_service_validation_cf.py:780-790`, `test_pivot_filters.py:844-880`; read `ConditionalFormatRule.type` full `models.py:960` |
| 6 | Scope: no new tools / no DTO slimming / no save_as/job_list / SCHEMA_VERSION=1; out-of-list adapter edits | **APPROVE** | `models.py:65`; tool_count=91; adapter/service edits necessary & correct (see §4) |
| 7 | Unit suite green; `verify_u3.py` PASS; integration pruned-value tests | **APPROVE (with tracked test-debt)** | 1081 unit passed / 0 failed; `RESULT=PASS`; 6 integration sites still construct pruned writes (deferred) |

**Final U3 verdict: ACCEPT** — production code, unit tests and `verify_u3.py` are correct and green. One **non-blocking, tracked** follow-up: 6 integration tests still construct pruned write DTOs and must be aligned (see §6). This is a deferred test update, **not** a production incompatibility.

---

## 2. Item 1 — Request enums added; full enums unchanged — APPROVE

`domain/statuses.py`:

- **`ConditionalFormatRequestType`** (`443-464`): exactly the 11 plain `FormatCondition` values — `cell_value, expression, text_string, blanks, no_blanks, errors, no_errors, time_period, top10, above_average, unique_values`. **No** `color_scale/data_bar/icon_set/unsupported`. Matches plan §12.2 verbatim.
- **`PivotFieldLayoutOrientation`** (`497-509`): `hidden, row, column, page`. **No** `data`.
- **`PivotFilterRequestKind`** (`566-576`): `manual` only. (Plan §12.2 permits `none` "where the request needs it"; the request does not — a manual filter with an empty set is rejected in `normalize_filter_spec`, and "show all" is `pivot_field_clear_filters`. Manual-only is correct.)
- **Full read/classify enums unchanged:** `ConditionalFormatType` (`418-440`, 15 values incl. `color_scale/data_bar/icon_set/unsupported`), `PivotFieldOrientation` (`487-494`, incl. `data`), `PivotFilterKind` (`549-563`, incl. `label/value/date/top_bottom/none/unsupported`) are byte-identical to baseline (no diff in those class bodies).

Runtime confirmation (`verify_u3.py`): `cf_request_types` = the 11 plain values; `layout_request_orientations=['column','hidden','page','row']`; `filter_request_kinds=['manual']`; `read_enums_full=True`.

---

## 3. Item 2 — Request retypes + `PivotFilterState` split + read fidelity — APPROVE

### 3.1 Retypes (`domain/models.py`)
- `ConditionalFormatSpec.type: ConditionalFormatRequestType` (`981`) — was `ConditionalFormatType`.
- `PivotFieldLayout.orientation: PivotFieldLayoutOrientation` (`1203`) — was `PivotFieldOrientation`.
- `PivotFilterSpec.kind: PivotFilterRequestKind = MANUAL` (`1224`) — was `PivotFilterKind`.
- Read DTOs retained full enums: `ConditionalFormatRule.type: ConditionalFormatType` (`960`), `PivotFieldInfo.orientation: PivotFieldOrientation` (`1125`) / `filter_state: PivotFilterKind` (`1130`).

### 3.2 The dual-use split (the core risk) — handled correctly
`PivotFilterSpec` was previously **dual-use** (write request *and* read result). The implementor introduced **`PivotFilterState`** (`models.py:1230-1242`) carrying the **full** `PivotFilterKind`, and repointed `PivotFieldFilters.filter` to it (`1253`). Every **read-path** constructor was migrated off the now-pruned `PivotFilterSpec`:
- `pivot_tables_adapter.py:1640` (`_field_filters`), `:1438` (`_unsupported_filter_result`), `:2788` (inventory).
- `pivot_tables_service.py:1054` (`_unsupported_filter_dto`).

A grep for `PivotFilterSpec(` across `src/` returns **only the class definition** (`models.py:1215`) — no read path constructs the pruned write DTO. Therefore the split cannot reject a non-manual kind on the read path.

### 3.3 Read fidelity is NOT degraded
The read path's `filter.kind` is computed by `_classify_filter_state`/`_page_filter_state` and is **only ever `NONE` or `MANUAL`** (`pivot_tables_adapter.py:1540,1548,1557,1586`); `label/value/date/top_bottom/unsupported` are reported as **strings in the `unsupported` list** via `_detected_filter_types`/`_filter_kind_for_code` (`1294-1312,1590-1621`). This is **identical to baseline** — the only change was the DTO type swapped (`PivotFilterSpec`→`PivotFilterState`), with the same `kind` value and `supported` flag. `PivotFilterState(kind=PivotFilterKind.LABEL, supported=False)` is representable (asserted in `test_phase2_contracts.py:456-457`), so the full kind set with `supported=false` remains expressible on read. **No read-fidelity regression.**

### 3.4 Identity-vs-value comparisons — correct and necessary
Because `ConditionalFormatSpec.type` and `ConditionalFormatRule.type` are now **different enum classes**, the cross-class `is` identity tests would always evaluate "mismatch". The implementor fixed exactly the two cross-class sites:
- `ranges_adapter.py:2743` `match.type.value != spec.type.value` (modify/delete type guard).
- `ranges_adapter.py:2886` `rule.type.value != spec.type.value` (postcondition re-read matcher).

Same-class identity comparisons were correctly **left untouched** (e.g. `ranges_adapter.py:2538` `_validation_matches_spec` compares `ValidationRule.type`/`ValidationSpec.type`, both `DataValidationType`; `:2547` `ValidationOperator`). Cross-class **value** lookups (`spec.type in _SUPPORTED_CF_TYPES` `:2247`, `_CF_TYPE_VALUES[spec.type]` `:2794`) remain correct because both enums are `StrEnum` (hash+eq by string value). This is the minimal, correct fix set.

---

## 4. Item 6 (scope) — out-of-list adapter/service edits are NECESSARY, not overreach — APPROVE

Plan §14 U3 "Changed files" listed `statuses.py`, `models.py`, `server.py`, and `api/pivot_tools.py`/`range_tools.py` ("only if a type import changes"). The implementor additionally touched `adapters/excel/ranges_adapter.py`, `adapters/excel/pivot_tables_adapter.py`, `application/pivot_tables_service.py`. Judgment: **necessary and correct**, a direct and unavoidable consequence of the `PivotFilterSpec` dual-use split and the CF enum retype:
- Splitting `PivotFilterSpec` forces every **read** constructor to move to `PivotFilterState` (adapter ×3, service ×1). Omitting this would make the read path construct the pruned write DTO and raise `ValidationError` on any non-manual kind — a real read-fidelity regression. (Exactly the scenario the task anticipated.)
- The CF `is`→`.value` fix (`ranges_adapter.py:2743,2886`) is required for the modify/delete guard and postcondition matcher to function across the two enum classes.
- `pivot_tools.py:107-140,351` `_require_orientation` + `pivot_field_set_layout` signature retype is required by the pruned tool param; it still returns the **full** `PivotFieldOrientation` to the unchanged service.

No scope violations: **no new tool** (tool_count=91; `xlvba_environment` is referenced only in descriptions/comments, not registered — that is U2), **no DTO slimming**, **no `save_as`/`job_list`**, **`SCHEMA_VERSION = 1`** (`models.py:65`). `api/range_tools.py` correctly needed no edit (it passes `ConditionalFormatSpec` through; the pruned type flows from `models.py`).

---

## 5. Items 3, 4, 5 — annotations, descriptions, runtime rejection — APPROVE

### 5.1 Annotations (§11) — `server.py:1733-1827, 2489`
- `_ANN_READ_ONLY` (`1733-1775`) matches the §11 read list (reads/lists/`server_health`/`excel_health`/`job_get`/`snapshot_list`/`snapshot_diff`/`test_get_run`/`macro_get_last_result`/`vba_read_component`/`vba_list_*`/`pivot_*_list|get`/`range_read_*`/`range_get_*`/`table_read_rows`/`sheets_list`/`sheet_get`/`names_list`/`name_get`/`shapes_list`/`shape_get`/`controls_list`/`control_get`/`workbook_get_model`/`excel_get_status`/`excel_list_workbooks`). `server_health` is read-only via `@mcp.tool(annotations=ToolAnnotations(read_only_hint=True))` (`2489`).
- `_ANN_DESTRUCTIVE` (`1776-1788`) = exactly the 9 §11 destructive tools.
- `_ANN_IDEMPOTENT` (`1789-1817`) covers the guarded mutations with postcondition re-read (§11 "most guarded mutations").
- **`macro_run` and `vba_compile` are in none of the sets** → not read-only, `destructiveHint` unset for `macro_run` (per §11). Confirmed by `verify_u3.py` `macro_run_vba_compile_not_read_only=True`.
- **Completeness:** 90 `server.add_tool(` + 1 `@mcp.tool` `server_health` = **91** `annotations=`; `missing_annotations=[]`; **`title=` count = 0** (no churn, per §11 non-goal).

### 5.2 Descriptions match pruned schema (§12.1 "D in U3", R-D2) 
- CF `server.py:763-774`: lists the 11 plain types; "`color_scale, data_bar, icon_set and unsupported` are **not accepted**".
- Layout `899-911`: "`orientation` is limited to hidden/row/column/page; 'data' is **not accepted**".
- Filter `954-966`: "`spec.kind` is **manual only**; label, value, date and top_bottom are **not accepted**".

All three are consistent with the pruned request schemas (verified live).

### 5.3 Runtime rejection now pydantic; read still `supported=false`
- Request DTOs reject pruned values at construction (pydantic `ValidationError`) — unit tests updated to assert this: `test_ranges_service_validation_cf.py:780-790,1295-1305,1520-1530`; `test_pivot_filters.py:844-880,1082,1223-1240`; `test_phase2_contracts.py:444-462`. This is the §17 **versioned request-contract narrowing** (domain rejection → pydantic `INVALID_RANGE`; both reject).
- Read still reports `supported=false`: `ConditionalFormatRule.type` stays full `ConditionalFormatType` with `supported=False` for color/data/icon (`models.py:951-969`, `_classify_cf_type` `ranges_adapter.py:2572-2589`); pivot reads report non-manual kinds via the `unsupported` string list and `PivotFilterState.supported=false` on the unsupported-result paths (`pivot_tables_adapter.py:1438-1445`, `pivot_tables_service.py:1051-1061`).

---

## 6. Item 7 — tests — APPROVE with tracked test-debt

- **Unit:** `pytest tests/unit -q` → **1081 passed, 0 failed / 0 error / 0 skipped** (progress output is all `.`; summary line suppressed by config but char census confirms 1081 `.` and zero `F/E/s/x`). The U3-relevant unit tests were correctly migrated to assert pydantic rejection **and** read-state fidelity via `PivotFilterState`.
- **Implementor `verify_u3.py`:** `RESULT=PASS` — `tool_count=91`, `missing_annotations=[]`, `pruned_values_absent=True`, `read_enums_full=True`, `macro_run_vba_compile_not_read_only=True`.
- **Integration test-debt (deferred, non-blocking):** 6 integration sites still construct pruned **write** DTOs directly and will now raise `ValidationError` at construction (they expect the old domain `INVALID_RANGE`/unsupported result):
  - `test_phase2_ranges_pivots_regression.py:1431` (`PivotFilterSpec(kind=LABEL)`), `:1467` (`ConditionalFormatSpec(type=COLOR_SCALE)`)
  - `test_pivot_filters.py:545` (`PivotFilterSpec(kind=LABEL)`)
  - `test_pivot_tables_service.py:773` (`PivotFilterSpec(kind=LABEL)`)
  - `test_ranges_service.py:1151` (`ConditionalFormatSpec(type=COLOR_SCALE)`)
  - `test_ranges_service_validation_cf.py:560` (`ConditionalFormatSpec(type=COLOR_SCALE)`)

  **Assessment:** this is a **deferred test update, not a production incompatibility**. The production narrowing is intentional and documented (§17); the tests bypass the wire and construct the DTO with a pruned value, so the fix is mechanical (assert `ValidationError`, or send the value over the wire and assert the validation error per §15.2). It does **not** block U3 production acceptance, but it **does** block the plan §15.3 "§15.2 integration green" gate for the overall remediation until aligned. These tests require live Excel and were not runnable in this validation.

---

## 7. Warnings / non-blocking observations

- **W1 (test alignment, required follow-up):** update the 6 integration sites in §6 to the narrowed contract. Until then the integration marker is red for enum-rejection cases. Owner: U3 implementer / integration owner.
- **W2 (vestigial fields, harmless):** `PivotFilterSpec` retains `supported`/`unsupported_reason` (`models.py:1226-1227`). With `kind` pruned to `manual`, these now only serve the "manual request explicitly marked unsupported" path (used by `test_service_unsupported_filter_returns_before_dispatch`). Not a defect; do not remove without checking that path.
- **W3 (dead-but-safe branches):** `_unsupported_filter_result` (`pivot_tables_adapter.py:1431`) and `_unsupported_filter_dto` (`pivot_tables_service.py:1046-1050`) retain an `else PivotFilterKind.UNSUPPORTED` branch that is now unreachable from the wire (pydantic rejects non-manual first). Defensive; no action needed.
- **W4 (annotation hints are best-effort):** `idempotentHint` is applied broadly to guarded mutations per §11; a few (e.g. `pivot_table_refresh`) rely on truth-signal re-read rather than a strict postcondition. Acceptable as hints; enforcement is name-glob based and unaffected.

---

## 8. Required fixes (blocking)

- **None blocking U3 production acceptance.**

## 9. Advisory (non-blocking)

- Align the 6 integration tests (W1) before the Phase-7 integration gate.
- Optionally document the `PivotFilterSpec` `supported`/`unsupported_reason` vestige (W2) in `docs/09` alongside the §17 narrowing note.

---

## 10. Final verdict

**U3 = ACCEPT.** All seven validation items pass on the production code. Request enums are correctly pruned (11 CF / 4 orientation / 1 filter), full read/classify enums are untouched, the `PivotFilterSpec`→`PivotFilterState` split preserves read fidelity with all read constructors migrated, the two cross-class identity comparisons were correctly converted to value comparisons, annotations are truthful and complete on 91/91 tools with no `title=` and `macro_run`/`vba_compile` not read-only, the three enum descriptions match the pruned schemas, pruned values are rejected at the pydantic layer while reads still report `supported=false`, scope is respected (SCHEMA_VERSION=1, no new tool, no slimming), the unit suite is green (1081 passed) and `verify_u3.py` reports PASS. The only outstanding item is **deferred integration-test alignment** (6 sites), which is non-blocking for production and must be closed before the overall remediation's integration-green gate.