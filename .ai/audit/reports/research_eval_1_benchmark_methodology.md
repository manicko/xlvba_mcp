# Research 1 (Benchmark Methodology) — Rigor Review of `benchmark_tasks.yaml`

- **Author:** Researcher 1 (Benchmark Methodology) — xlvba MCP Agent Usability Evaluation
- **Repo root:** `C:\py_exp\xlvba_mcp`
- **Read-only scope:** No production code, tests, config, or existing benchmark files were modified. This is the only artifact written.
- **Primary sources inspected (current revision `40c8b4d`, docs-only atop `81bcf2c`):**
  - `.ai/tmp/benchmark_tasks.yaml` (28 tasks + meta)
  - `.ai/audit/reports/research_2b_usability_benchmark.md` (original design)
  - `.ai/audit/reports/agent_usability_baseline.md` (Phase 0 grounding)
  - `src/excel_vba_mcp/server.py`, `src/excel_vba_mcp/api/*.py`, `src/excel_vba_mcp/domain/models.py`, `src/excel_vba_mcp/domain/errors.py`, `src/excel_vba_mcp/application/{ranges_service,macro_service,job_registry,job_store,session_service}.py`, `src/excel_vba_mcp/adapters/excel/tables_adapter.py`
  - Ground truth: `tests/fixtures/phase2_manifest.json`, `tests/fixtures/build_full_access_fixture.py`

---

## 1. Executive summary (feasibility verdict)

**Verdict: FEASIBLE WITH REQUIRED FIXES.** The 28-task set is well-grounded in the real tool contracts and fixtures, spans the required categories, and is executable against a fresh server per run. However, it is **not yet internally consistent** and would produce **false failures** on at least three tasks if run as authored:

1. **T05 and T06 are defective (HIGH).** Both call `range_write_values` with **no `snapshot_create`** in `expected_sequence` or `setup`, but `RangesService.write_values` enforces the snapshot gate *before* any write (`ranges_service.py:285-286` → `_require_snapshot` → `SNAPSHOT_FAILED`). The gold path is blocked; `success_criteria: result.verified==True` is unreachable. Sibling mutating tasks (T08/T09/T12/T13/T14/T25/T26) correctly include `snapshot_create`, so T05/T06 are inconsistent. **Fix: add `snapshot_create` to both.**
2. **T22 is not deterministically achievable (HIGH).** `job_cancel` only sets `cancel_requested=True` when the job is **non-terminal** (`job_registry.py:339-345`); a background `excel_open` completes in well under a second, so by cancel time the job is terminal and `cancel_requested` stays `False`. **Fix: inject a genuinely long-running job** (see §6).
3. **T14 has an order-sensitivity false-failure risk (MEDIUM).** The criterion `visible_items==[East,North,South]` assumes alphabetical order, but pivot item order follows source order (`North,South,East` in `phase2_manifest.json:693-736`). **Fix: compare as a set/sorted.**

Two tasks are **over-strict / ambiguous** and need fair-scoring rules (T11 `pivot_table_get` is a valid alternative; T03/T15/T08 have valid precursor/selector alternatives). One required category — **cross-server routing (`xlvba_*` vs `excel-mcp_*`)** — is **not covered** (the benchmark connects only the single `xlvba` server). The remediated-only `xlvba_environment` tool and `xlvba://glossary` resource are **not exercised** by any task, so the benchmark does not directly reward the remediation. Pagination/limits and unsupported-pivot-source handling are also uncovered (the latter needs a new fixture and is deferred).

None of these are fatal; all are fixable without weakening the benchmark. The suite is **not** artificially easy — most tasks require exact tool selection, correct argument shapes, snapshot/classify preconditions, and deterministic postconditions.

---

## 2. Per-category coverage

| Category (required) | Task IDs | Adequate? | Gaps / notes |
|---|---|---|---|
| Workbook lifecycle (open/close/save) | T01, T02, T20, T28 | Adequate | `excel_close` not directly asserted; `excel_save` only via T28 gap task. |
| Read (values/formulas) | T03, T04 | Adequate | `range_get` precursor not modeled (see fair-scoring). |
| Write (values) | T05, T06 | **Defective** | Missing `snapshot_create` (T05/T06). |
| Tables (read/update/append) | T07, T08, T09 | Adequate | No `table_create`/`table_delete_rows`/column ops; sufficient for usability signal. |
| Pivots (read/mutate) | T10, T11, T12, T13, T14 | Adequate | T14 order risk; T11 over-strict; **no unsupported-source case** (needs new fixture). |
| VBA (read/replace/stale) | T25, T26 | Adequate | Strong (hash-guard + stale rejection). |
| Macros (run/classify/policy) | T15, T16, T17 | Adequate | Covers success, interactive veto, unclassified block. |
| Permissions / denials | T18, T19 | Adequate | T18 restricted-open, T19 read-only write. |
| Recovery / jobs | T20, T21, T22, T23, T24 | Partial | T22 not deterministic; T21/T22 need long-running job. |
| Testing (declarative) | T27 | Adequate (thin) | Single test task; no negative (invalid-suite) case. |
| **Cross-server routing (`xlvba_*` vs `excel-mcp_*`)** | **—** | **MISSING** | Only the `xlvba` server is connected; cannot be tested without a second server. See T29 (conditional). |
| Env/policy discovery (`xlvba_environment`, glossary) | — | Missing | Remediation not rewarded. See T30. |
| Pagination / limits | — | Missing | See T31 (uses existing fixture). |
| Unsupported pivot source | — | Missing (deferred) | Needs external/OLAP fixture; offline build infeasible (`phase2_manifest.json:16`). |

---

## 3. Per-task correctness review

Legend: **first** = `expected_first_tool` matches the real handler; **args** = `expected_args` valid against the contract; **obj** = success criteria objectively/independently verifiable (not model self-report); **amb** = ambiguity that needs fair scoring.

| ID | first | args | obj | amb | Notes (evidence) |
|---|---|---|---|---|---|
| T01 | ✓ | ✓ | ✓ | no | `excel_open` `workbook_tools.py:88`; defaults `mode=execute,read_only=False` (L90-91); `WorkbookInfo.mode/read_only/workbook_id` `models.py:86-99`. |
| T02 | ✓ | ✓ | ✓ | no | inspect+read_only auto-allowed READ_ONLY `session_service.py:270-271`. |
| T03 | ✓ | ✓ | ✓ | mild | `range_read_values` `range_tools.py:174`; `RangeValues.values` typed `models.py:894-908`; B8 text / B10 error per `phase2_manifest.json:78-98`. Allow `range_get` precursor. |
| T04 | ✓ | ✓ | ✓ | no | `range_read_formulas` `range_tools.py:206`; formulas B9-B12 `phase2_manifest.json:367-372`. |
| T05 | ✓ | ✓ | ✗ | no | **DEFECT**: `write_values` requires snapshot `ranges_service.py:285-286`; no `snapshot_create` in sequence/setup → `SNAPSHOT_FAILED`. Add snapshot. |
| T06 | ✓ | ✓ | ✗ | no | **DEFECT**: same as T05. Add snapshot. |
| T07 | ✓ | ✓ | ✓ | no | `table_read_rows` `table_tools.py:135`; `TableRows.columns`+`rows` positional `tables_adapter.py:93-111`; harness must zip to dicts. |
| T08 | ✓ | ✓ | ✓ | yes | `table_update_rows` index/where `table_tools.py:220-242`; factor value at `Sheet1!B2` `build_full_access_fixture.py:147-149`. Accept index **or** where. |
| T09 | ✓ | ✓ | ✓ | no | `table_append_rows` `table_tools.py:166`; headers id/name/qty `phase2_manifest.json:448-475`. |
| T10 | ✓ | ✓ | ✓ | no | `pivot_tables_list` `pivot_tools.py:211`; caches ptA/ptB=1, ptC=2 `phase2_manifest.json:562-669`. |
| T11 | ✓ | ✓ | ✓ | **yes** | `pivot_fields_list` `pivot_tools.py:270`; `pivot_table_get` also returns layout → must_not over-strict. Accept either. |
| T12 | ✓ | ✓ | ✓ | no | `pivot_table_refresh` `pivot_tools.py:309`; shared cache affects ptA+ptB `phase2_manifest.json:822-828`; source_type=1 (database) refreshable. |
| T13 | ✓ | ✓ | ✓ | no | `pivot_field_set_aggregation` `pivot_tools.py:419-454`; caption unique. |
| T14 | ✓ | ✓ | ✓ | **yes** | `pivot_field_apply_filter` `pivot_tools.py:491`; **order risk**: source order North,South,East `phase2_manifest.json:693-736`. Compare as set. West already hidden (idempotent). |
| T15 | ✓ | ✓ | ✓ | mild | `macro_classify_background` `macro_tools.py:132`, `macro_run` L87; Recalculate writes `Result!A1=factor*2` `build_full_access_fixture.py:102-111`. Allow `macro_list` precursor. |
| T16 | ✓ | ✓ | ✓ | no | classify raises `INTERACTIVE_UI_NOT_SUPPORTED` `macro_service.py:786-793` (is_error). |
| T17 | ✓ | ✓ | ✓ | no | `macro_run` returns truthful `status=blocked`, `error_code=BACKGROUND_SAFETY_NOT_VERIFIED` `macro_service.py:507-524`, `models.py:457-467` (NOT is_error — criteria correctly checks fields). |
| T18 | ✓ | ✓ | ✓ | no | restricted execute open → `APPROVAL_DENIED` `session_service.py:272-279,1168-1197`; fallback inspect+read_only allowed. Harness must start restricted server. |
| T19 | ✓ | ✓ | ✓ | no | read-only check precedes snapshot `ranges_service.py:285-286` → `WORKBOOK_READ_ONLY` wins. Correct. |
| T20 | ✓ | ✓ | ✓ | no | `excel_open(background=true)` → `RunningJob{status:running,job_id,poll_after_s∈[1,30]}` `job_registry.py:192-207`, `workbook_tools.py:120-126`. |
| T21 | ✓ | ✓ | ✓ | no | `job_get` `job_tools.py:96`; terminal status; job may already be terminal (fine). |
| T22 | ✓ | ✓ | ✗ | no | **RISK**: `cancel_requested=True` only if non-terminal `job_registry.py:339-345`; fast open job is terminal → false failure. Needs long-running job. |
| T23 | ✓ | ✓ | ✓ | no | `excel_health` dead-pid → `needs_recovery` `health_tools.py:141-151`; `excel_recover` ok/terminated/process_restarted/prior_job=unknown `recovery_tools.py:29-44`. Harness kills owned pid; healthy control = no recover. |
| T24 | ✓ | ✓ | ✓ | no | `JobState.INTERRUPTED_UNKNOWN` `job_registry.py:81`; outcome='unknown'/retryable=False `job_store.py:59`, `errors.py:267-281`. Harness seeds + reconciles at startup. |
| T25 | ✓ | ✓ | ✓ | no | `vba_read_component` `vba_tools.py:115`; `vba_replace_component_source` requires `expected_source_hash` L191-224; hash from read. |
| T26 | ✓ | ✓ | ✓ | no | wrong hash → `REVISION_CONFLICT` `errors.py:585-593`; source unchanged. |
| T27 | ✓ | ✓ | ✓ | no | `test_suite_run` `test_tools.py:189-234`; TestStep `table_update_rows`/`macro_run(background_safe)` + `cell_value_equals`; factor→10 then Recalculate→20. |
| T28 | ✓ | ✓ | ✓ | **yes** | no `save_as` (grep of `server.py` returns none; persistence = `excel_save`/`snapshot`/`vba_export_project`); inherently ambiguous — use gap-scoring (§5). |

**Totals:** first_tool correct 28/28; args valid 28/28; objective criteria 25/28 (T05, T06, T22 fail as authored); ambiguity flagged on T03, T08, T11, T14, T15, T28.

---

## 4. Recommended additions (NEW tasks only; original 28 untouched)

Additions are provided in the **same YAML field shape** as the existing tasks. They are justified only where a required category or a remediation-critical behavior is otherwise unmeasured.

### T29 — Cross-server routing (CONDITIONAL: only if a second `excel-mcp_*` server is connected)
Justification: the assignment lists cross-server routing as required; the baseline report flags `xlvba_*` vs `excel-mcp_*` overlap with "no routing rule" (`agent_usability_baseline.md:88`). Not testable with the single `xlvba` server; add **only** if the harness connects a second overlapping server. Scored on correct target selection, not on a fixed sequence.

```yaml
  - id: T29-cross-server-routing
    category: cross-server-routing
    prompt: "Read Cells!B2:B5 from RangesCellsProject.xlsm. Two MCP servers are connected that both expose range-read tools; use the one that owns this workbook."
    fixture: RangesCellsProject.xlsm
    setup: "two servers connected (xlvba + an excel-mcp_* server); workbook opened on the xlvba server; fresh full-access"
    access_mode: full
    expected_first_tool: excel_open            # only the xlvba server can open the allowlisted copy
    expected_sequence: [excel_open, range_read_values]
    expected_args:
      excel_open: {path: "<allowlisted temp copy>"}
      range_read_values: {workbook_id: "<id from xlvba server>", sheet: Cells, address: "B2:B5"}
    success_criteria: "the range read is issued to the SAME server that holds the workbook_id (xlvba); result.values is 4x1 and matches B2..B5; the agent does not send an xlvba workbook_id to the excel-mcp_* server"
    must_not: ["route an xlvba workbook_id to a different server", "open the same file on both servers concurrently (single-STA violation)", "modify the repo fixture"]
    token_measure: tool_description_tokens(range_read_values)+context_tokens
    failure_taxonomy: [WRONG_TOOL, HALLUCINATED_TOOL, INCOMPLETE]
```

### T30 — Environment / policy discovery (exercises the remediated `xlvba_environment` + `xlvba://glossary`)
Justification: the 92nd tool and 5th resource are the remediation's headline additions but no task rewards them. Scored on **answer correctness**, not on mandating the tool (baseline has no such tool, so requiring it would be unfair to baseline).

```yaml
  - id: T30-environment-policy-discovery
    category: env-discovery
    prompt: "Before I edit anything: tell me this server's access mode, whether a snapshot is required before a mutation, and whether macro classification alone authorizes a run. Do not modify the workbook."
    fixture: FullAccessWorkflow.xlsm
    setup: "fresh server, full access mode; no workbook open required"
    access_mode: full
    expected_first_tool: xlvba_environment
    expected_sequence: [xlvba_environment]      # baseline may substitute excel_get_status + inference; scored on answer
    expected_args:
      xlvba_environment: {}                      # non-COM, read-only, side-effect-free
    success_criteria: "agent reports access_mode=='full' (or the effective mode), snapshot_required_for_mutations==True, and classification_is_authorization==False; no mutation tool is called; if xlvba_environment is unavailable (baseline) the agent may answer from excel_get_status/policy and still pass"
    must_not: ["call any mutating tool", "claim classification authorizes a run", "modify the repo fixture"]
    token_measure: tool_description_tokens(xlvba_environment)+context_tokens
    failure_taxonomy: [HALLUCINATED_TOOL, WRONG_ARG_SHAPE, UNSAFE_ACTION]
```

### T31 — Pagination / limits (uses existing fixture; no new fixture)
Justification: `range_read_values`/`table_read_rows` support `offset`/`limit`/`truncated` (`range_tools.py:174-203`, `models.py:894-909`) but no task exercises windowing.

```yaml
  - id: T31-read-range-limit-window
    category: read-range
    prompt: "Read only the first 3 rows of Cells!B2:B15 in RangesCellsProject.xlsm and tell me whether the response was truncated."
    fixture: RangesCellsProject.xlsm
    setup: "workbook open"
    access_mode: full
    expected_first_tool: excel_open
    expected_sequence: [excel_open, range_read_values]
    expected_args:
      range_read_values: {workbook_id: "<id>", sheet: Cells, address: "B2:B15", limit: 3}
    success_criteria: "result.rows==3 AND result.truncated==True AND the 3 returned rows equal B2,B3,B4 (string,number,number)"
    must_not: ["omit limit and return all 14 rows", "write anything"]
    token_measure: tool_description_tokens(range_read_values)+context_tokens
    failure_taxonomy: [WRONG_ARG_SHAPE, WRONG_TOOL, INCOMPLETE]
```

### Deferred (NOT added as a task) — Unsupported pivot source
`UNSUPPORTED_PIVOT_SOURCE` (`errors.py:508-512`) is referenced in T12 `must_not` but never exercised. All fixture pivots are `source_type=1` database/Table-backed (`phase2_manifest.json:562-669`) and the manifest states fixtures are offline with "no workbook connections, Power Query or Data Model" (`phase2_manifest.json:16`). A faithful unsupported-source fixture cannot be built offline, so **no task is added**; this gap is recorded rather than faked.

---

## 5. Fair-scoring rules for ambiguous tasks

Apply these **in addition to** the deterministic predicates; they prevent penalizing valid alternative workflows without making the suite easy.

1. **Neutral read-only precursors (T03, T04, T07, T10, T11, T15).** A read-only, goal-advancing call issued *before* the expected call — `range_get`, `table_get`, `pivot_table_get`, `macro_list`, `sheets_list`, `workbook_get_model` — is **neutral**: it does not reduce CTS and does not count as UC, provided it is not an error and precedes the expected call. (Rationale: these are legitimate discovery steps; the expected call must still occur and the postcondition must hold.)
2. **Alternative valid selectors (T08).** Accept `updates:[{index:1,...}]` **or** `updates:[{where:{SettingKey:factor},...}]`; both are valid per `table_tools.py:220-242`. Reject only a `where` that matches >1 row (`AMBIGUOUS_FILTER`).
3. **Alternative valid tool (T11).** Accept `pivot_fields_list` **or** `pivot_table_get` for the fields+orientation question; both return orientation. Relax the T11 `must_not` accordingly (do not fail a run solely for choosing `pivot_table_get`).
4. **Order-insensitive list equality (T14, T10).** Compare `visible_items`/`hidden_items`/pivot-name lists as **sets** (or sorted), not exact order, because pivot item order follows source order (`phase2_manifest.json:693-736`).
5. **Gap task (T28).** Success = (a) no `save_as`/`save_copy_as`/`export_fixed_format` call, (b) persistence via a real surface (`excel_save` and/or `snapshot_create`), (c) VBA exported via `vba_export_project`, (d) the agent **explicitly states** the save-as-new-file gap. Any of the real persistence paths is acceptable.
6. **Snapshot placement (T05/T06 after fix).** Accept `snapshot_create` anywhere before the first mutation; do not require a specific position.
7. **Precursor for T20/T21/T22.** `excel_get_status`/`excel_health` before `job_get`/`job_cancel` is neutral (read-only).
8. **No double-jeopardy.** A single `must_not` violation caps WC=0 for that run but must not also be counted as a separate CTS failure for the same call.

---

## 6. Risks to validity and mitigations

| # | Risk | Affected tasks | Mitigation (deterministic) |
|---|---|---|---|
| R1 | Missing snapshot precondition → false failure | T05, T06 | Add `snapshot_create` to `expected_sequence`+`setup`; or harness injects a snapshot for all mutating tasks (then remove it from T08/T09/T12/T13/T14/T25/T26 for consistency). |
| R2 | Cancel on an already-terminal job → `cancel_requested=False` | T22 (and T21 latency) | Inject a **genuinely long-running** job: add a `SleepMacro` (e.g. `Sleep 5`) to a fixture module, or start `test_suite_run(background=true)` on a multi-case suite; cancel/poll while non-terminal. |
| R3 | Pivot item order assumed alphabetical | T14 | Set/sorted comparison (§5.4). |
| R4 | Over-strict tool mandate | T11 | Accept `pivot_table_get` (§5.3). |
| R5 | Live-Excel state not reset deterministically | T08/T09/T12/T13/T14/T15/T25/T26/T27/T28 | Fresh server + fresh fixture copy per run; assert `sha256(repo_fixture)` unchanged (design §4.1). Each mutating task starts from the pristine fixture, so `Result!A1`/`tblSettings` reset automatically. |
| R6 | Recovery task needs a dead owned pid | T23 | Harness reads `excel_health.owned_pid`, terminates it via the held handle (TerminateProcess), then re-runs; healthy control uses a separate fresh server with a live pid. |
| R7 | `interrupted_unknown` injection fidelity | T24 | Fresh base dir with a pre-seeded non-terminal `JobStore` record reconciled at startup (`job_store.py:47-59`, `server.py:1187-1193`); fresh server per run. |
| R8 | Baseline replay needs a different server build | all | The live server is a single shared instance; baseline (`e2289ca`, 90 tools) and remediated (`81bcf2c`, 92 tools) **cannot run concurrently** — restart the server per revision and per run; never mix catalogs in one run. |
| R9 | Model nondeterminism | all | N≥3 runs, fixed model/temperature, report CI; median CT. |
| R10 | Token accounting drift | all | Record method; use `bytes/4` proxy consistently; compare within the same runtime. |
| R11 | Setup calls leaking into the trace | all | Record injected setup calls separately; exclude from CTS/UC (design §5.2). |
| R12 | `test_suite_run` macro step needs classification | T27 | The `background_safe:true` declaration is recorded by the runner before dispatch; Recalculate is classifiable (no interactive tokens, `build_full_access_fixture.py:28-32`). Keep as-is. |

**Deterministic reset summary for live-Excel-state tasks:** fresh server + fresh temp fixture copy per run (SHA-256 asserted) is the universal reset; T23 additionally kills the owned pid; T24 additionally seeds the job store; T21/T22 additionally require a long-running injected job. No original is ever opened in place.

---

## 7. What I did NOT change

- No production code, tests, config, or existing benchmark files were modified.
- The original 28 tasks in `.ai/tmp/benchmark_tasks.yaml` were **not** rewritten; all recommended additions are NEW ids (T29–T31) and one deferred gap (unsupported pivot source) is documented, not faked.
- I did not read sibling researcher outputs (`research_eval_2_*`, `research_1a/1b/2a/2b`) for independence; grounding is from source + fixtures only.
- Only this report file was written: `.ai/audit/reports/research_eval_1_benchmark_methodology.md`.