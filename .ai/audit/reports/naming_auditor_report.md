# Naming Refactor — Independent Auditor Report

**Project:** `xlvba_mcp` · **Reviewer:** Independent Auditor (read-only) · **Date:** 2026-11-28
**Scope:** Tech-Lead-approved MCP tool+action rename only. Working-tree large-data WIP (`data_export_service.py`, `ranges_adapter.py`, `cell_values.py`, `_build_runtime` re-wire, etc.) is OUT OF SCOPE and was not treated as a naming defect; it was checked only to confirm the naming diff did not entangle business logic.

## Verdict: ACCEPT-WITH-NONBLOCKING

- **Blocking issues: 0**
- **Non-blocking issues: 3** (all internal-only / cosmetic residuals; no client-facing behavior defect)

The 6 approved tool renames and the single action rename are implemented correctly and completely. Registration, descriptions, action enum, dispatcher `context["tool"]` strings, annotations (membership + totals 14/16/23), and the fixture ledger are all consistent with the approved decisions. No permission, approval, or annotation-meaning regression. All 104 affected hermetic tests pass (25 rename/description-contract + 79 dispatch/mapping/baseline).

---

## Checklist results

| # | Check | Result | Evidence |
|---|-------|--------|----------|
| 1 | 6 renames in `name=`; old names absent from registration | **PASS** | `server.py:2046,2088,2235,2241,2247,2273` set `name=` to canonical; before/after artifact diff: removed `{excel_save_as,name,name_delete,name_read,test,workbook_file_read}`, added `{workbook_save_copy,test_suite,defined_name,defined_name_delete,defined_name_read,workbook_snapshot_read}`; 43 tools before and after |
| 2 | `_ANN_*` membership preserved 1:1; totals 14/16/23 | **PASS** | `server.py:1951-2011` — only strings renamed; artifact totals readOnly=14, destructive=16, idempotent=23 |
| 3 | test action Literal `["run","get-run"]`; dispatch `if action=="run"`; `run-suite`→INVALID_RANGE; get-run intact | **PASS** | `api/test_tools.py:311,343,357`; `test_tool_rename_contract.py::test_run_suite_action_rejected_with_invalid_range` + `test_get_run_dispatches_to_get_run_path` pass |
| 4 | Descriptions canonical + valid actions + examples; test_suite run+get-run examples + relationship; no `run-suite` | **PASS** | `server.py:337-375` (_DESC_TEST has Relationship + run/get-run examples); artifact scan: `run-suite` absent from all 43; `_DESC_SAVE_AS` example added `server.py:272-274` |
| 5 | Client-facing `context["tool"]` canonical | **PASS** | `api/name_tools.py:284,322`; `api/ooxml_tools.py:45,163`; `api/test_tools.py:349,359` all canonical |
| 6 | No business logic/control-flow/validation change beyond names+strings; internal symbols unchanged | **PASS** | api diffs are string-only + the required `if action=="run"` token; only non-naming server.py hunk is `_build_runtime` (line 1345, noted WIP). Internal function/constant symbols unchanged |
| 7 | Annotation triples | **PASS** | artifact: workbook_save_copy F/T/F, test_suite F/T/F, defined_name F/F/T, defined_name_read T/F/T, defined_name_delete F/T/F, workbook_snapshot_read T/F/T |
| 8 | Mapping ledger completeness + policy consistency | **PASS** | `tool_action_mapping.md` §3 asserts 94 old-names → 43 canonical, completeness True; policy §5 verdicts match approved set; no aliases (policy §7) |
| 9 | Residual old PUBLIC name reachable by client | **NON-BLOCKING** | Registration/descriptions/action-enum/ctx are clean; residuals are internal-only + cosmetic schema (see table below) |
| 10 | No permission/approval/annotation-meaning regression | **PASS** | Annotation frozensets 1:1; approval-key label `session_service.py:770,855` unchanged; `macro_policy` op-taxonomy unchanged; fixture `permission_profile` values unchanged (only `new_name` updated) |

---

## Findings table

| Severity | File:line | Evidence | Recommendation |
|----------|-----------|----------|----------------|
| NON-BLOCKING | `api/workbook_tools.py:189,379`; `api/__init__.py:37,51` | Internal function `excel_save_as` + `__all__` export; not a registered name | None — approved to keep internal symbols unchanged |
| NON-BLOCKING | `api/ooxml_tools.py:3,80,170` | Internal function/docstring/`__all__` `workbook_file_read` | None — internal symbol; registered description comes from `_DESC_WORKBOOK_FILE_READ` (canonical) |
| NON-BLOCKING | `api/name_tools.py:4,6,14,210,288,328-332` (+ message text `:320`) | Internal `name`/`name_read`/`name_delete` symbols, docstrings, `__all__`; error **message** still says "Unknown name_read action" (ctx field is canonical `defined_name_read`) | Optional: update the human-readable message string to `defined_name_read` for polish; not client-behavior-affecting |
| NON-BLOCKING | `api/test_tools.py:271,373` | `section="test"` job label + `__all__ "test"` | None — internal job-section label |
| NON-BLOCKING | `domain/test_spec.py:69`; `security/macro_policy.py:154` | `NAME_DELETE="name_delete"` TestAction enum; macro_policy op-taxonomy | None — explicitly acceptable per task (test_spec TestAction enum / macro_policy taxonomy) |
| NON-BLOCKING | `application/names_service.py:5,54,213`; `application/session_service.py:770,855` | Audit `tool="name_delete"` / approval-key + audit `tool="excel_save_as"` | None — internal audit/approval-key labels; explicitly acceptable |
| NON-BLOCKING | `domain/models.py:1690` | `SaveAsResult` docstring "Result of ``excel_save_as``" → surfaces in `workbook_save_copy` **outputSchema** description | Optional: reword docstring to canonical name; cosmetic schema text only |
| NON-BLOCKING | (schema, auto-derived) | `workbook_save_copy` **inputSchema** `title: "excel_save_asArguments"` (auto-generated from unchanged internal function name) | None required — unavoidable consequence of keeping internal symbols; cosmetic schema title, not name/action/ctx |
| NON-BLOCKING | (schema, `test_suite` inputSchema `$defs`) | TestAction enum values `name_create`/`name_update`/`name_delete` (suite step vocabulary) | None — explicitly acceptable (test_spec TestAction enum) |

---

## Residual internal-only old-name list (NON-BLOCKING, with justification)

These are NOT reachable as an MCP tool name, action value, or dispatcher `context["tool"]`; they are internal Python symbols, audit/approval labels, model docstrings, or the test-suite step-action vocabulary — all of which the approved scope intentionally leaves unchanged.

- `excel_save_as`: `api/workbook_tools.py:189,379`, `api/__init__.py:37,51`, `server.py:122,2045` (import + positional handler ref; `name=` is canonical), `application/session_service.py:770,855` (approval-key/audit label), `domain/models.py:1690` (docstring).
- `workbook_file_read`: `api/ooxml_tools.py:3,80,170`, `server.py:79,2272` (import + positional handler ref).
- `name` / `name_read` / `name_delete`: `api/name_tools.py` (symbols/docstrings/`__all__`), `server.py:75-76,2240,2246` (imports + positional refs), `application/names_service.py:5,54,213` (audit label), `security/macro_policy.py:154`, `domain/test_spec.py:69` (TestAction enum).
- `test`: `api/test_tools.py:271` (job `section` label), `:373` (`__all__`).
- Schema-cosmetic only: `workbook_save_copy` inputSchema title `excel_save_asArguments`; outputSchema description from `SaveAsResult` docstring; `test_suite` step-action enum `name_*`.

Justification: the task's approved scope explicitly keeps internal Python function/constant symbols unchanged and lists `macro_policy` op-taxonomy, `test_spec` TestAction enum, and `session_service` approval-key label as ACCEPTABLE. The schema `title`/docstring residuals are a direct, unavoidable byproduct of that decision and are cosmetic (not the published name, action enum, or error `context`).

---

## Runtime verification performed (read-only, no edits)

- `.venv\Scripts\python.exe -m pytest tests/unit/test_tool_rename_contract.py tests/unit/test_description_contract.py -q` → 25 passed.
- `.venv\Scripts\python.exe -m pytest tests/unit/test_macro_test_dispatch.py tests/unit/test_name_dispatch.py tests/unit/test_workbook_dispatch.py tests/unit/test_mapping_and_approval.py tests/unit/test_wp0_baseline.py -q` → 79 passed.
- Artifact diff of `tools_list_before.json` vs `tools_list_after.json`: exactly the 6 approved renames, 43 tools each, totals 14/16/23, `run-suite` absent.

## Notes / out-of-scope confirmations

- The naming diff did NOT entangle business logic: every `server.py` hunk is a `_DESC_*` string, an `_ANN_*` frozenset string, a `_register_tools` `name=`/`annotations=` string, or the pre-existing `_build_runtime` DataExportService re-wire (line 1345) that the task flagged as inherited WIP.
- Fixture `wp0_baseline.json` `baseline.tool_count` was corrected 94→43 (the baseline doc had already flagged 94 as a stale/mislabeled ledger-row count); `serialized_bytes`/`measured_at` remain stale metadata but are not naming and out of scope.