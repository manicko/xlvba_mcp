# report_08e — Final Auditor Spot-Check (source-to-contract consistency)

- **Role:** Auditor (final consistency spot-check). AUDIT ONLY — no code/test/config modified.
- **Repo / working tree:** `C:\py_exp\xlvba_mcp` @ baseline `e2289ca` + accepted U1+U2+U3, U4 documented no-op, U5 docs, + aligned integration tests.
- **Date:** 2026-11-26.
- **Method:** read-only inspection + `capture_baseline.py` (no Excel) + `pytest tests/unit` + targeted `git diff`. Scratch scripts confined to `.ai/tmp` (gitignored); no tracked file touched.

## PASS/FAIL table

| # | Check | Verdict | Primary evidence |
|---|-------|---------|------------------|
| 1 | Aggregate claims (92 tools, 92/92 annotations, glossary resource, error to_dict 8 keys, request enums pruned / read enums full) | **PASS** | capture output; `statuses.py`; `errors.py`; `server.py:2653` |
| 2 | No unrelated changes (only expected files; no stray production edit) | **PASS** (with note) | `git status`; `git diff --stat`; pivot_tools/pivot_tables_service are legitimate U3 |
| 3 | Description↔schema consistency (3 pruned-enum tools + vba_compile / macro_classify_background / job_cancel / xlvba_environment) | **PASS** | baseline JSON descriptions vs pruned enums; `environment_tools.py` |
| 4 | Non-negotiables intact | **PASS** | `git diff` (pyproject/uv.lock untouched); `server.py:2502`; `statuses.py`; names check |
| 5 | 6 aligned integration tests assert pydantic-layer rejection (not weakened) + PivotFilterState read fidelity | **PASS** | 6 test functions across 5 files |
| 6 | Compatibility: `>=87` assert valid; additive only | **PASS** | `test_mcp_full_access_workflow.py:499`; models.py retypes only |

**Blocking inconsistencies: NONE.**

## Evidence

### 1 — Aggregate claims — PASS
- `capture_baseline.py`: `TOOL_COUNT=92`, `TOOLS_WITH_ANNOTATIONS=92`, `TOOLS_WITH_OUTPUT_SCHEMA=92`.
- **Annotations nuance (truthful, not a defect):** 92/92 tools emit an `annotations` object (key present); **73** carry ≥1 hint and **19** carry an empty `{}` (truthfully neither read-only/destructive/idempotent — e.g. `macro_run`, `vba_compile`, `job_cancel`, `excel_open`). `macro_run`/`vba_compile` are correctly **not** `readOnlyHint` (verified `annotations={}`); `xlvba_environment`/`server_health` are `readOnlyHint:true`; `sheet_delete`/`excel_recover` are `destructiveHint:true`.
- **Glossary resource present:** `server.py:2653` `@mcp.resource("xlvba://glossary")` → `glossary_resource()` returns `_GLOSSARY_MD` (`server.py:2626-2661`); resources list = `['xlvba://glossary']` + 4 URI templates.
- **Error `to_dict` = 8 keys:** `{code, context, errorCategory, hresult, hresult_hex, isRetryable, message, suggestedNextActions}` (runtime check on `SNAPSHOT_FAILED` → `errorCategory=missing_prerequisite`, `isRetryable=False`, `suggestedNextActions=['snapshot_create(workbook_id)','retry mutation']`).
- **Request enums pruned (actual enum arrays, not description text):**
  - `ConditionalFormatRequestType` = `[cell_value, expression, text_string, blanks, no_blanks, errors, no_errors, time_period, top10, above_average, unique_values]` (`statuses.py:443-464`).
  - `PivotFieldLayoutOrientation` = `[hidden, row, column, page]` (`statuses.py:497-509`).
  - `PivotFilterRequestKind` = `[manual]` (`statuses.py:566-576`).
- **Read enums full (unchanged):** `ConditionalFormatType` incl. `color_scale/data_bar/icon_set/unsupported` (`statuses.py:418-440`); `PivotFieldOrientation` incl. `data` (`statuses.py:487-494`); `PivotFilterKind` incl. `label/value/date/top_bottom/unsupported` (`statuses.py:549-563`).
- `pytest tests/unit` → **1127 passed** (incl. new `test_description_contract.py`, `test_environment_tools.py`, `test_error_envelope.py`).

### 2 — No unrelated changes — PASS (note on 2 files)
`git status` / `git diff --stat` show only: `server.py`, `domain/{models,statuses,errors}.py`, `api/environment_tools.py` (new), `application/session_service.py`, `adapters/excel/{pivot_tables_adapter,ranges_adapter}.py`, `api/pivot_tools.py`, `application/pivot_tables_service.py`, docs (EN+RU), and tests (5 integration + 4 unit modified, 3 unit new). `pyproject.toml`/`uv.lock` **not** in diff.
- **Note (not a finding):** the task's enumerated list ("adapters pivot/ranges") did not literally name `api/pivot_tools.py` or `application/pivot_tables_service.py`. Both are **legitimate U3 work**, not stray edits:
  - `pivot_tools.py:110-140,351` retypes the `orientation` request param to `PivotFieldLayoutOrientation` and rejects `data` (plan_03 §12.2/§14 U3 explicitly anticipates `api/pivot_tools.py` "only if a type import changes").
  - `pivot_tables_service.py:1043-1058` converts `PivotFilterRequestKind`→full `PivotFilterKind` for the unsupported-result path and emits `PivotFilterState` (read fidelity). No originals-safe / open-close / hashing logic touched.

### 3 — Description↔schema consistency (validator Q15) — PASS
Descriptions (from `tools/list`) match the pruned request schemas verbatim:
- `range_set_conditional_formatting`: "…plain FormatCondition set only (cell_value…unique_values); color_scale, data_bar, icon_set and unsupported are not accepted" ↔ 11-value request enum. ✔
- `pivot_field_set_layout`: "orientation is limited to hidden/row/column/page; 'data' is not accepted" ↔ pruned orientation enum. ✔
- `pivot_field_apply_filter`: "spec.kind is manual only; label, value, date and top_bottom are not accepted (rejected at the request layer)" ↔ `[manual]`. ✔
- `vba_compile`: "diagnostic only, does NOT compile … never returns compiled or compile_error" — no false capability. ✔
- `macro_classify_background`: "classification is bookkeeping only and never grants run permission; macro_run still enforces approval + BG06 + R13". ✔
- `job_cancel`: "intent-only; never interrupts; observe cancel_requested via job_get; use excel_recover". ✔
- `xlvba_environment`: "build … always null", "classification_is_authorization=false", "excel.available=false when the worker is down" — matches `environment_tools.py:151,166,178,216`. ✔

### 4 — Non-negotiables — PASS
- **No SDK change:** `pyproject.toml`/`uv.lock` absent from diff.
- **No error `structuredContent`:** `server.py:2502` `raise ToolError(json.dumps(cause.to_dict(), ensure_ascii=False))` — text-only envelope.
- **No new `ErrorCode` members:** 29 at HEAD == 29 baseline (`git show HEAD` count).
- **`SCHEMA_VERSION=1`:** runtime check = 1; `EnvironmentReport(schema_version=1)` (`environment_tools.py:227`).
- **No `save_as`/`job_list` tool:** names check → `save_as`, `job_list`, `xlvba_job_list`, `xlvba_save_as`, `xlvba_search_tools` all ABSENT.
- **No renames:** `job_cancel`, `vba_compile`, `test_run_suite`, `test_suite_run` all present (total 92).
- **Originals-safe untouched:** `session_service.py` diff is additive only (`_snapshot_gate` + `set_snapshot_gate` + `_build_policy`; `has_snapshot` reads in-memory gate, not store); no open/close/hash changes.

### 5 — 6 aligned integration tests — PASS
Six test functions across 5 files now assert rejection at the **pydantic** layer (not weakened) and preserve read fidelity:
1. `test_ranges_service_validation_cf.py::test_set_conditional_formatting_rejects_unsupported_type` → `pytest.raises(ValidationError)` on `ConditionalFormatSpec(type=COLOR_SCALE)`.
2. `test_ranges_service.py::test_set_conditional_formatting_rejects_unsupported_type` → same pydantic assertion.
3. `test_pivot_filters.py::test_apply_unsupported_label_filter_is_explicit` → `pytest.raises(ValidationError)` on `PivotFilterSpec(kind=LABEL)` + `PivotFilterState(kind=LABEL, supported=False)` fidelity.
4. `test_pivot_tables_service.py::test_unsupported_label_filter_is_explicit` → same pydantic + `PivotFilterState` fidelity.
5. `test_phase2_ranges_pivots_regression.py::test_unsupported_pivot_filter_kind_is_explicit` → pydantic + `PivotFilterState` (`kind is LABEL`, `supported is False`).
6. `test_phase2_ranges_pivots_regression.py::test_unsupported_color_scale_cf_is_explicit` → pydantic rejection.
Rejection is still asserted (moved from domain/COM to request layer per U3); `@requires_excel` dropped because these are now pure-model checks — alignment, not weakening.

### 6 — Compatibility — PASS
- `test_mcp_full_access_workflow.py:499` `assert len(tool_names) >= 87` still valid (92 ≥ 87); no "all-90" claim.
- Additive only: new `xlvba_environment` (91→92); new DTOs (`SessionPolicy`, `Environment*`, `PivotFilterState`); additive `SessionStatus.policy`; additive `to_dict` keys. `models.py` deletions (10) are exclusively the intended U3 request-enum **retypes** (`type`/`orientation`/`kind`) + `PivotFilterSpec`→`PivotFilterState` read swap + docstrings — **no field removed**.

## Advisories (non-blocking)
- **A-1 (annotation empty-set):** 19 tools carry an empty `ToolAnnotations{}`. This is truthful per plan §11, but a reviewer skimming "92/92 annotated" should note 19 are intentionally empty (no hint applies). No action required.
- **A-2 (task-list shorthand):** the audit brief's file list omitted `api/pivot_tools.py` and `application/pivot_tables_service.py`; both are in-scope U3. Recommend the brief's expected-file list be read as "pivot/ranges surface" inclusive of these two. No code action.
- **A-3 (excel_close hint):** `excel_close` defaults to discarding unsaved edits yet has empty annotations (plan §11 did not list it as destructive). Advisory only — annotations are display hints, not enforcement.

## Final recommendation

**ACCEPT** the remediation as a whole.

All six spot-check items PASS with source-cited evidence; no blocking inconsistency. The working tree is `e2289ca` + the accepted U1/U2/U3 (+ U4 no-op, U5 docs) + 6 aligned integration tests, with no stray production edits, no SDK/contract violations, and unit suite green (1127 passed). The only deviations from the brief's literal file list are two legitimate U3 files, documented above.

*End of report_08e. AUDIT ONLY — nothing modified.*