# Agent Usability Evaluation — Phase 2 Change Map

**Auditor:** Architecture Auditor (read-only)
**Baseline:** `e2289ca` (91 tools, 4 URI-templated resources, 0 annotations, 5-key error envelope)
**Remediated:** `81bcf2c` (92 tools, 5 resources incl. `xlvba://glossary`, 92/92 annotations, additive error envelope)
**HEAD:** `40c8b4d` (docs-only; empty src diff vs `81bcf2c`)
**Repo:** `C:\py_exp\xlvba_mcp` · **Method:** isolated subprocess import of each revision's `src` tree + `mcp.list_tools()` / `list_resources()` / `list_resource_templates()` + `git show`/`git diff` source verification + live Kilo MCP read-only probes.

> **Scope discipline:** No production, test, or config files were modified. The only artifact written is this report. All Python helpers were written to `.ai/tmp/` (throwaway) and the baseline tree was extracted to the OS temp dir via `git archive` (no repo mutation).

---

## 1. Tool NAME set — PASS

| Metric | Baseline `e2289ca` | Remediated `81bcf2c` |
|---|---|---|
| Tool count | 91 | 92 |
| Added | — | `xlvba_environment` |
| Removed | — | none |
| Renamed | — | none |

**Verified by:** programmatic `mcp.list_tools()` on each isolated revision. The 91-name baseline set is a strict subset of the 92-name remediated set; the sole delta is `+xlvba_environment`. No removals, no renames. **PASS.**

---

## 2. Tool DESCRIPTIONS — PASS

53 of the 91 shared tools changed description (58%). All changes follow one of five categories; none is a regression.

| Category | Nature | Example tools |
|---|---|---|
| **Structured template** | Rewritten into `Purpose / Use when / Do not use when / Preconditions / Parameters / Behavior / Result / Next actions` | excel_open, range_write_values, sheet_add, table_create (most of the 53) |
| **Neighbor routing** | "Do not use when … use `<neighbor>` instead" | excel_open→excel_list_workbooks, vba_compile→vba_get_compile_diagnostics, pivot_field_set_layout→pivot_field_set_aggregation |
| **Examples** | Added `Example: {…}` JSON arg block | excel_open, range_write_values, table_update_rows |
| **Guard sentence** | Uniform R-D3 mutation guard ("snapshot + approval … postcondition re-read") | range_set_format, pivot_field_set_aggregation, all guarded mutations |
| **Truthful-claim fixes** | Disclaimer that the tool does NOT do what its name implies | vba_compile, macro_classify_background, job_cancel (see below) |

**Truthful-claim spot-checks (all confirmed against source + description):**

| Tool | Baseline claim (risk) | Remediated claim | Source truth |
|---|---|---|---|
| `vba_compile` | "fail-closed compile result" (implied it compiles) | "diagnostic only, **does NOT compile** … none is available; use `vba_get_compile_diagnostics`" | `vba_compile` returns `status='indeterminate'`, `is_compile=false` — matches |
| `macro_classify_background` | "UNGATED bookkeeping" (soft) | "classification is bookkeeping only and **never grants run permission**; macro_run still enforces approval + BG06 + R13" | `MacroService.classify_background` is UNGATED bookkeeping; macro_run enforces gates — matches |
| `job_cancel` | "cancel intent (intent-only; never interrupts …)" | "it is **intent-only; never interrupts**; observe cancel_requested via job_get and use excel_recover to actually stop a wedged process" | `job_cancel` sets `cancel_requested` only; single-STA call runs to completion — matches |

**No description claims a capability the code lacks** — the three highest-risk names now carry explicit disclaimers that match the implementation. The description-contract test (`tests/unit/test_description_contract.py`) enforces the template, truthful phrases, examples, uniform guard sentence, and the 92-count. **PASS.**

---

## 3. INPUT schemas — PASS

Only 5 tools changed input schema; all are the intended request-enum pruning (plus test-suite propagation of the same prune). **No top-level `required` field changed for any tool** (verified programmatically: 0 changes). No defaults changed on any input field.

| Tool | Change | Classification |
|---|---|---|
| `range_set_conditional_formatting` | Request enum `ConditionalFormatType` (15 vals incl. color_scale/data_bar/icon_set/unsupported) → pruned `ConditionalFormatRequestType` (11 plain FormatCondition vals) | **Pruned request enum** (intended) |
| `pivot_field_set_layout` | `PivotFieldOrientation` (incl. `data`) → `PivotFieldLayoutOrientation` (hidden/row/column/page; `data` rejected) | **Pruned request enum** (intended) |
| `pivot_field_apply_filter` | `PivotFilterKind` (7 vals) → `PivotFilterRequestKind` (`manual` only) | **Pruned request enum** (intended) |
| `test_run_suite` | Embedded `PivotFilterSpec.kind` enum pruned (same `PivotFilterRequestKind`) | Propagation of the prune (not independent) |
| `test_suite_run` | Same embedded prune | Propagation of the prune (not independent) |

**Read/classify enums unchanged:** the full `ConditionalFormatType` (15), `PivotFieldOrientation` (5), and `PivotFilterKind` (7) are retained in the **output** schemas of the read/classify tools (`range_get_conditional_formatting`, `pivot_fields_list`, `pivot_field_get_filters`, `pivot_field_clear_filters`, `pivot_field_apply_filter` output). Verified: read tools still report all 7 filter kinds with `supported=false`. **PASS.**

---

## 4. OUTPUT schemas — PASS (additive only)

U4 was a documented no-op; the only material output changes are **additive** (new optional DTOs), not restructures.

| Tool(s) | Change | Classification |
|---|---|---|
| `excel_close`, `excel_get_status` | `SessionStatus` gains optional `policy: SessionPolicy \| None = None`; new `SessionPolicy` DTO (approval_mode, mutating/executing_requires_grant, recovery_auto_allowed, snapshot_required_for_mutations, has_snapshot, classification_is_authorization=false, r13_preflight="veto") | **Additive** (optional field, default None) |
| `pivot_field_apply_filter`, `pivot_field_clear_filters`, `pivot_field_get_filters` | `PivotFilterSpec` → `PivotFilterState` in the `filter` field of `PivotFieldFilters` (read/classify DTO); full `PivotFilterKind` retained | **Additive/rename of read DTO** (same 7-kind enum) |
| (all tools) | `EnvironmentReport` output only on the new `xlvba_environment` | New tool |

No output schema removed a field or changed a required field on an existing tool. **PASS.**

### tools/list payload size (chars, full serialized)

| Metric | Baseline | Remediated | Δ |
|---|---|---|---|
| **Full tools/list** | 458,854 | 495,792 | **+36,938 (+8.0%)** |
| Descriptions | 24,272 | 37,125 | +12,853 (+53%) |
| inputSchema | 80,656 | 82,373 | +1,717 (+2.1%) |
| outputSchema | 338,547 | 350,758 | +12,211 (+3.6%) |

*Method: `json.dumps({"tools":[t.model_dump(by_alias=True)…]})` on each isolated import. The +8% total is dominated by the description template expansion (+53% on descriptions alone), which is the intended usability investment; schema growth is marginal (+2–4%).* **PASS.**

---

## 5. ANNOTATIONS — PASS

| Metric | Baseline | Remediated |
|---|---|---|
| Tools with annotations | 0 / 91 | **92 / 92** |
| `readOnlyHint=true` | 0 | **41** |
| `destructiveHint=true` | 0 | **9** |
| `idempotentHint=true` | 0 | **25** |
| No annotations | 91 | 0 |

- **`macro_run` and `vba_compile` are NOT readOnly** — confirmed absent from `_ANN_READ_ONLY` and confirmed `readOnlyHint=None` in the live/imported annotation dump. Source comment: *"macro_run and vba_compile are intentionally NOT read-only."*
- Distribution is internally consistent: every readOnly tool is a read/get/list/model tool; destructive = {sheet_delete, range_merge, excel_recover, snapshot_restore, table_delete_rows, name_delete, range_delete, vba_replace_component_source, vba_insert_code}; idempotent = the 25 settable/idempotent mutations. **PASS.**

---

## 6. GLOSSARY resource — PASS

| Metric | Baseline | Remediated |
|---|---|---|
| Static resources | 0 | **1** (`xlvba://glossary`, name=glossary, mime=text/markdown) |
| URI-templated resources | 4 | 4 (unchanged) |

- `xlvba://glossary` is **absent at `e2289ca`**, **present at `81bcf2c`** (verified by `list_resources()` on each isolated import and by live Kilo `list_mcp_resources`).
- Live `read_mcp_resource(xlvba://glossary)` returns **exactly 14 terms**: R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, SUSPECTED_STALLED, interrupted_unknown, indeterminate, working copy, snapshot gate. **PASS.**

---

## 7. ENVIRONMENT tool + ERROR envelope — PASS

### `xlvba_environment` contract (source: `api/environment_tools.py`, 195 lines)

| Contract clause | Verified |
|---|---|
| Non-COM / side-effect-free | Handler reads `_config`, `_session`, `_path_policy` bookkeeping only; docstring + code confirm no worker start, no COM, no snapshot/grant/classification |
| `readOnlyHint` | `readOnlyHint=true` (in `_ANN_READ_ONLY`, confirmed in dump) |
| `build` always null | `EnvironmentExcelInfo(..., build=None, target_build="16.0.14332")` — `build` is always `None`; `_TARGET_BUILD` is a design-time label only |
| `has_snapshot` from in-memory gate | `_workbooks()` reads `status.policy.has_snapshot` ← `SessionService._build_policy` ← `_snapshot_gate` = `SnapshotService.has_snapshot` → `_latest_snapshot` (in-memory, **not** the on-disk store) |
| `policy.classification_is_authorization=false` | `_policy()` sets `classification_is_authorization=False`; DTO field is `const: false` |
| `excel.available=false` when worker down | `_excel_info` returns `available=False, build=None` when no worker |

Live `xlvba_environment` output matches: `build=null`, `classification_is_authorization=false`, `r13_preflight="veto"`, `glossary_uri="xlvba://glossary"`, `pivot_filter_kinds=["manual"]`, `non_activating_compile=false`. **PASS.**

### Additive error envelope (`domain/errors.py`)

`to_dict()` preserves the original 5 keys (`code, message, hresult, hresult_hex, context`) and **adds** `errorCategory, isRetryable, suggestedNextActions` via pure functions `category_for` / `retryable_for` / `actions_for`. No schema bump, no new `ErrorCode` members.

| D-U10 case | Expected | Verified in source + test |
|---|---|---|
| `STALE_TARGET` | retryable=false + "may have partially applied" caveat | `_RETRYABLE_CODES` excludes it; `_ACTIONS_BY_CODE[STALE_TARGET]` contains "…MAY have partially applied — re-read state; do NOT blind-replay"; test asserts `retryable_for is False` and `"partially applied" in joined` |
| `REVISION_CONFLICT` | retryable=false + caveat | same as above |
| `interrupted_unknown` (job) | retryable=false | `_JOB_CATEGORY` → `interrupted_unknown`; `retryable_for` returns False for any job key; actions = `["never resubmit; report status"]`; test asserts all three |
| Only COM_BUSY / RPC_E_CALL_REJECTED retryable | true | `_RETRYABLE_CODES = {COM_BUSY, RPC_E_CALL_REJECTED}` |

### Additive `SessionStatus.policy`

`SessionService._build_policy` assembles `SessionPolicy` with `classification_is_authorization=False`, `r13_preflight="veto"`, `has_snapshot` from the in-memory gate; `SessionStatus.policy` defaults to `None`. Purely additive; changes no existing behavior. **PASS.**

---

## 8. Confounding runtime changes — PASS (none behavioral beyond enum pruning)

All 10 changed source files reviewed. Non-server behavioral deltas:

| File | Change | Judgment |
|---|---|---|
| `adapters/excel/ranges_adapter.py` | `match.type is not spec.type` → `match.type.value != spec.type.value` (2 sites) | **Behavioral but a latent-bug fix**: identity comparison of StrEnums could misbehave; value comparison is strictly more correct. No usability confound; improves correctness. |
| `adapters/excel/pivot_tables_adapter.py` | `PivotFilterKind`→`PivotFilterRequestKind` on request path; `PivotFilterSpec`→`PivotFilterState` on read path | **Enum-pruning consequence only** (request pruned, read retained). No behavior change beyond rejecting previously-accepted-but-unsupported kinds. |
| `api/pivot_tools.py` | `_require_orientation` accepts pruned enum; `pivot_field_set_layout` param type pruned | **Enum-pruning consequence only**; `data` now rejected at request (was silently ineffective). |
| `application/pivot_tables_service.py` | `_unsupported_filter_dto` builds `PivotFilterState` | **Enum-pruning consequence only**. |
| `application/session_service.py` | +`_snapshot_gate`, +`_build_policy`, +`policy=` on status | **Additive**; no existing behavior changed (gate is read-only, policy defaults None). |
| `domain/errors.py` | +3 pure functions, +additive keys in `to_dict` | **Additive**. |
| `domain/models.py` | +`SessionPolicy`, +`Environment*` DTOs, +`PivotFilterState`; `filter` field type swap | **Additive** (207 add / 11 remove; removes are the type swap + policy field). |
| `domain/statuses.py` | +3 request enums | **Additive**. |
| `api/environment_tools.py` | new file | New tool. |
| `server.py` | +env tool, +glossary, +annotations, +descriptions | Registration/metadata only. |

**Conclusion:** the only non-additive runtime change is the `ranges_adapter` `is`→`.value` comparison fix (a correctness improvement, not a confound). Everything else is additive or enum-pruning. A usability comparison between the two revisions is **not confounded** by unrelated behavioral changes. **PASS.**

---

## 9. Live Kilo forwarding — PASS

Confirmed from inside this session (read-only MCP calls only; no Excel start/stop, no mutation):

| Probe | Result |
|---|---|
| `xlvba_xlvba_environment` | Returns `schema_version=1`, `build=null`, `classification_is_authorization=false`, `r13_preflight="veto"`, `glossary_uri="xlvba://glossary"`, `pivot_filter_kinds=["manual"]`, `non_activating_compile=false` — matches `81bcf2c` |
| `xlvba_server_health` | `{status: ok, service: excel-vba-mcp, version: 0.1.0}` |
| `list_mcp_resources` | `xlvba://glossary` present (server=xlvba) alongside context7/memory |
| `read_mcp_resource(xlvba://glossary)` | 14 terms, matches source `_GLOSSARY_MD` |

The live server is consistent with `81bcf2c`. **PASS.**

---

## Verification summary

| # | Item | Classification |
|---|---|---|
| 1 | Tool name set (+xlvba_environment, 0 removals/renames) | **PASS** |
| 2 | Description changes (template/routing/examples/guard/truthful) | **PASS** |
| 3 | Input schema (3 pruned request enums + 2 propagation; read enums unchanged; 0 required/default changes) | **PASS** |
| 4 | Output schema (additive only; U4 no-op) + payload size | **PASS** |
| 5 | Annotations 92/92; macro_run/vba_compile not readOnly; 41/9/25 | **PASS** |
| 6 | Glossary resource (absent→present, 14 terms) | **PASS** |
| 7 | xlvba_environment contract + additive error envelope (D-U10) + SessionStatus.policy | **PASS** |
| 8 | Confounding runtime changes (none behavioral beyond enum prune + 1 correctness fix) | **PASS** |
| 9 | Live Kilo forwarding matches 81bcf2c | **PASS** |

**Overall: PASS (9/9).**

---

## Verified vs unverifiable

**Verified (with evidence):**
- Exact tool name sets at both revisions (programmatic import of each isolated `src` tree).
- Description, inputSchema, outputSchema, annotations for all tools at both revisions (programmatic `model_dump(by_alias=True)` diff).
- Full tools/list payload char counts at both revisions.
- Resource/template sets at both revisions; live glossary content (14 terms).
- `xlvba_environment` source contract and live output.
- Error-envelope pure functions and D-U10 cases (source + unit test `test_error_envelope.py`).
- Annotation frozensets (source) and resulting distribution.
- Absence of behavioral confounds (full source diff of all 10 changed files).
- Live Kilo forwarding of env/health/resources/glossary.

**Unverifiable in this audit (stated limitations):**
- **Runtime COM behavior** of the pruned enums (e.g., that Excel actually rejects `color_scale` on write) was not exercised — no Excel instance was started (read-only mandate). Source + tests are the evidence; live COM rejection is asserted by code/tests, not re-run here.
- **`vba_compile` "diagnostic only"** was verified at the contract/source level (returns `indeterminate`, `is_compile=false`) and via the description; the actual VBE compile path was not invoked (would require Excel + VBOM trust).
- **MCP client rendering** of annotations (how Kilo displays readOnly/destructive hints) is client-side and out of scope; only the emitted `ToolAnnotations` were verified.
- **Performance/latency** of the larger tools/list payload (+8%) was not benchmarked; only char counts are reported.

**No production changes were made.** No `.py`, test, or config file under version control was modified. The only file written is this report (`.ai/audit/reports/agent_usability_eval_change_map.md`); throwaway helpers live in `.ai/tmp/` and the baseline tree was extracted to the OS temp dir via `git archive` (no repo mutation).