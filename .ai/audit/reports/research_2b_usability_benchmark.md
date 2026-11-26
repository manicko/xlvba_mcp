# Research 2B — Reproducible Agent Usability Benchmark for `xlvba`

- **Author:** Researcher 2B (usability-remediation research)
- **Repo root:** `C:\py_exp\xlvba_mcp`
- **Server revision:** `e2289ca` (git HEAD at authoring)
- **Scope:** RESEARCH / DESIGN ONLY. No production code, tests, or config were modified.
- **Independence:** Authored without reading `research_1a_*` or `research_1b_*`.
- **Companion artifact:** `C:\py_exp\xlvba_mcp\.ai\tmp\benchmark_tasks.yaml` (`BENCHMARK_TASKS`, 28 tasks, machine-readable, validated parseable).

---

## 0. What this benchmark measures (and what it does NOT)

This benchmark measures **actual agent task execution against the live MCP server** — the tools the agent really calls, in what order, with what arguments, and whether the task's deterministic success criteria are met. It does **not** ask an LLM to rate tool *descriptions*. Every expected tool name and argument below is grounded in the source at `e2289ca`:

- Tool registration: `src/excel_vba_mcp/server.py` `_register_tools` (L1476–L1704) — **90 tools** registered.
- Handler signatures / argument contracts: `src/excel_vba_mcp/api/*.py`.
- Payload schemas: `src/excel_vba_mcp/domain/models.py`, `src/excel_vba_mcp/domain/test_spec.py`.
- Enums / error codes: `src/excel_vba_mcp/domain/statuses.py`, `src/excel_vba_mcp/domain/errors.py`.
- Fixtures + ground truth: `tests/fixtures/phase2_manifest.json`, `tests/fixtures/build_full_access_fixture.py`.
- Live-server harness pattern to reuse: `tests/integration_excel/test_mcp_full_access_workflow.py`.

### 0.1 Grounding facts the task set relies on (verified at `e2289ca`)

| Fact | Evidence |
|---|---|
| 90 tools registered; test asserts `>= 87` | `server.py` L1479–L1704; `test_mcp_full_access_workflow.py` L499 |
| **No `save_as` / `save_copy_as` / `export_fixed_format` tool exists** | grep of `server.py` returns none; persistence is `excel_save`, `excel_close(save_changes)`, `snapshot_create/restore`, `vba_export_project` |
| Default `excel_open` ⇒ `mode="execute"`, `read_only=False` | `workbook_tools.excel_open` L88–127; e2e L510–512 |
| `range_write_values` payload = `{rows, cols, values:[[CellValue]]}` | `models.RangeWriteRequest` L1000–1030; `CellValue` L390–396 |
| `table_update_rows` = `[{index,values}]` XOR `[{where,values}]` | `table_tools.table_update_rows` L220–242 |
| `macro_run` needs approved + `background_safe` (BG06) + R13 preflight, else `blocked` | `macro_tools.macro_run` L87–129; e2e L540–542 |
| `macro_classify_background` vetoes interactive ⇒ `INTERACTIVE_UI_NOT_SUPPORTED` | e2e L528–537; `errors.py` L56 |
| Background ops return `{status:"running", job_id, poll_after_s}`; `poll_after_s∈[1,30]` | `job_registry.py` L61, L192–207; `job_tools.start_long_op` |
| `excel_recover` agent-directed, fail-closed, originals safe, `prior_job="unknown"` | `recovery_tools.excel_recover` L46–69 |
| `vba_replace_component_source` needs `expected_source_hash`; stale ⇒ `REVISION_CONFLICT`/`STALE_TARGET` | `vba_tools` L191–224; `errors.py` L47,L62 |
| `test_suite_run` validates closed `TestSuiteSpec` (`extra="forbid"`) pre-Excel | `test_tools._parse_suite` L237–249; `test_spec.py` |
| Fixtures: `FullAccessWorkflow.xlsm` (tblSettings/Result/modWorkflow), `RangesCellsProject.xlsm` (Cells/StructOps/tblCells), `PivotTablesProject.xlsm` (ptA/ptB/ptC) | `phase2_manifest.json`; `build_full_access_fixture.py` |

---

## 1. The Task Set (28 tasks)

Categories span all required areas: open-mode (2), read-range (2), write-values (2), table (3), pivot (5), macro/policy (3), policy-denial (2), background-op (3), recovery (2), vba-mutate (2), test-workflow (1), save-export-gap (1).

### 1.1 Task table (summary)

| ID | Category | Prompt (abbrev.) | Expected first tool | Expected sequence | Success criteria (deterministic) | Key must-not |
|---|---|---|---|---|---|---|
| T01 | open-mode | Open FullAccessWorkflow for editing | `excel_open` | `excel_open` | mode=execute, read_only=False | inspect/read_only |
| T02 | open-mode | Inspect PivotTables read-only | `excel_open` | `excel_open` | mode=inspect, read_only=True | omit flags |
| T03 | read-range | Typed values Cells!B2:B15 | `excel_open` | `excel_open,range_read_values` | 14×1; kinds incl error; B8 text | use formulas tool |
| T04 | read-range | Formulas Cells!B9:B12 | `excel_open` | `excel_open,range_read_formulas` | =1+2,=1/0,=NA(),=B3*2 | use values tool |
| T05 | write-values | Write 2×2 block StructOps!D2:E3 | `excel_open` | `excel_open,range_write_values` | verified; re-read=[[1,2],[3,4]] | write into ListObject |
| T06 | write-values | Literal text `=1+1` into Cells!B20 | `excel_open` | `excel_open,range_write_values` | verified; stays text | use formulas tool |
| T07 | table-read | Rows of tblSettings | `excel_open` | `excel_open,table_read_rows` | columns+rows match | range tool on table |
| T08 | table-update | factor→10, read Sheet1!B2 | `excel_open` | `excel_open,snapshot_create,table_update_rows,range_read_values` | B2==10 | no snapshot; ambiguous where |
| T09 | table-update | Append row to tblCells | `excel_open` | `excel_open,snapshot_create,table_append_rows` | row count +1 | write into ListObject |
| T10 | pivot-read | PivotTables + shared caches | `excel_open` | `excel_open,pivot_tables_list` | ptA,ptB(cache1),ptC(cache2) | invent RefreshAll |
| T11 | pivot-read | ptA fields+orientation | `excel_open` | `excel_open,pivot_fields_list` | Region=row,Product=col,Quarter=page | use table_get |
| T12 | pivot-mutate | Refresh ptA (writable) | `excel_open` | `excel_open,snapshot_create,pivot_table_refresh` | refreshed; affects ptA+ptB | RefreshAll; no snapshot |
| T13 | pivot-mutate | Sales→Sum 'Total Sales' | `excel_open` | `excel_open,snapshot_create,pivot_field_set_aggregation` | data/sum/caption | legacy agg code |
| T14 | pivot-mutate | ptB hide 'West' (manual) | `excel_open` | `excel_open,snapshot_create,pivot_field_apply_filter` | hidden=[West] | label/value filter |
| T15 | macro-run | Classify+run Recalculate | `excel_open` | `excel_open,macro_classify_background,macro_run,range_read_values` | safe=True; success; A1==10 | run w/o classify |
| T16 | macro-policy | Mark ShowMessage safe | `excel_open` | `excel_open,macro_classify_background` | INTERACTIVE_UI_NOT_SUPPORTED | retry/run it |
| T17 | macro-policy | Run Recalculate (no classify) | `excel_open` | `excel_open,macro_run` | blocked/BACKGROUND_SAFETY_NOT_VERIFIED | classify to bypass |
| T18 | policy-denial | Edit on restricted server | `excel_open` | `excel_open` | APPROVAL_DENIED | loop retries |
| T19 | policy-denial | Read-only then write B2 | `excel_open` | `excel_open,range_write_values` | WORKBOOK_READ_ONLY | reopen writable |
| T20 | background-op | Open in background | `excel_open` | `excel_open` | running+job_id+poll_after_s | block synchronously |
| T21 | background-op | Poll job to terminal | `job_get` | `job_get[,job_get…]` | terminal status; no busy-spin | poll faster than hint |
| T22 | background-op | Cancel intent for job | `job_cancel` | `job_cancel` | cancel_requested=True | claim force-kill |
| T23 | recovery | Health→recover if needed | `excel_health` | `excel_health[,excel_recover]` | needs_recovery⇒recover ok; healthy⇒no recover | recover when healthy |
| T24 | recovery | interrupted_unknown job | `job_get` | `job_get` | status=interrupted_unknown,retryable=False | resubmit |
| T25 | vba-mutate | Read+replace modWorkflow | `excel_open` | `excel_open,vba_read_component,snapshot_create,vba_replace_component_source` | before/after hash; hash from read | guess hash |
| T26 | vba-mutate | Replace with wrong hash | `excel_open` | `excel_open,snapshot_create,vba_replace_component_source` | REVISION_CONFLICT/STALE_TARGET; unchanged | guess correct hash |
| T27 | test-workflow | Bounded declarative suite | `excel_open` | `excel_open,test_suite_run` | overall passed | arbitrary code |
| T28 | save-export-gap | Save-as-new + export VBA | `excel_open` | `excel_open,excel_save,vba_export_project` | recognizes no save_as; persists+exports; reports gap | hallucinate save_as |

### 1.2 Machine-readable block (authoritative full spec)

The authoritative, loadable spec is `BENCHMARK_TASKS` at `.ai/tmp/benchmark_tasks.yaml` (28 tasks, validated). A compact JSON projection for quick loading is below; the YAML carries the full per-task `expected_args`, `setup`, `access_mode`, `token_measure`, and `failure_taxonomy`.

```json
{
  "meta": {"server_revision": "e2289ca", "tool_count_registered": 90, "transport": "stdio", "entrypoint": "excel_vba_mcp"},
  "constants": {"model": "set-by-runner", "agent_config": "set-by-runner", "kilo_version": "set-by-runner",
                 "connected_tools": "all-90", "access_mode_default": "full", "excel_build": "16.0.14332"},
  "tasks": [
    {"id":"T01","category":"open-mode","first":"excel_open","seq":["excel_open"],"expect":{"mode":"execute","read_only":false}},
    {"id":"T02","category":"open-mode","first":"excel_open","seq":["excel_open"],"expect":{"mode":"inspect","read_only":true}},
    {"id":"T03","category":"read-range","first":"excel_open","seq":["excel_open","range_read_values"],"expect":{"rows":14,"cols":1}},
    {"id":"T04","category":"read-range","first":"excel_open","seq":["excel_open","range_read_formulas"],"expect":{"formulas":["=1+2","=1/0","=NA()","=B3*2"]}},
    {"id":"T05","category":"write-values","first":"excel_open","seq":["excel_open","range_write_values"],"expect":{"verified":true}},
    {"id":"T06","category":"write-values","first":"excel_open","seq":["excel_open","range_write_values"],"expect":{"stored_as_text":true}},
    {"id":"T07","category":"table-read","first":"excel_open","seq":["excel_open","table_read_rows"],"expect":{"columns":["SettingKey","SettingValue"]}},
    {"id":"T08","category":"table-update","first":"excel_open","seq":["excel_open","snapshot_create","table_update_rows","range_read_values"],"expect":{"Sheet1!B2":10}},
    {"id":"T09","category":"table-update","first":"excel_open","seq":["excel_open","snapshot_create","table_append_rows"],"expect":{"row_delta":1}},
    {"id":"T10","category":"pivot-read","first":"excel_open","seq":["excel_open","pivot_tables_list"],"expect":{"pivots":["ptA","ptB","ptC"]}},
    {"id":"T11","category":"pivot-read","first":"excel_open","seq":["excel_open","pivot_fields_list"],"expect":{"ptA":{"Region":"row","Product":"column","Quarter":"page"}}},
    {"id":"T12","category":"pivot-mutate","first":"excel_open","seq":["excel_open","snapshot_create","pivot_table_refresh"],"expect":{"status":"refreshed","affects":["ptA","ptB"]}},
    {"id":"T13","category":"pivot-mutate","first":"excel_open","seq":["excel_open","snapshot_create","pivot_field_set_aggregation"],"expect":{"function":"sum","caption":"Total Sales"}},
    {"id":"T14","category":"pivot-mutate","first":"excel_open","seq":["excel_open","snapshot_create","pivot_field_apply_filter"],"expect":{"hidden":["West"]}},
    {"id":"T15","category":"macro-run","first":"excel_open","seq":["excel_open","macro_classify_background","macro_run","range_read_values"],"expect":{"Result!A1":10}},
    {"id":"T16","category":"macro-policy","first":"excel_open","seq":["excel_open","macro_classify_background"],"expect":{"error":"INTERACTIVE_UI_NOT_SUPPORTED"}},
    {"id":"T17","category":"macro-policy","first":"excel_open","seq":["excel_open","macro_run"],"expect":{"status":"blocked","error":"BACKGROUND_SAFETY_NOT_VERIFIED"}},
    {"id":"T18","category":"policy-denial","first":"excel_open","seq":["excel_open"],"expect":{"error":"APPROVAL_DENIED"},"access_mode":"restricted"},
    {"id":"T19","category":"policy-denial","first":"excel_open","seq":["excel_open","range_write_values"],"expect":{"error":"WORKBOOK_READ_ONLY"}},
    {"id":"T20","category":"background-op","first":"excel_open","seq":["excel_open"],"expect":{"status":"running","has_job_id":true}},
    {"id":"T21","category":"background-op","first":"job_get","seq":["job_get"],"expect":{"terminal":true}},
    {"id":"T22","category":"background-op","first":"job_cancel","seq":["job_cancel"],"expect":{"cancel_requested":true}},
    {"id":"T23","category":"recovery","first":"excel_health","seq":["excel_health","excel_recover?"],"expect":{"recover_if":"needs_recovery"}},
    {"id":"T24","category":"recovery","first":"job_get","seq":["job_get"],"expect":{"status":"interrupted_unknown","retryable":false}},
    {"id":"T25","category":"vba-mutate","first":"excel_open","seq":["excel_open","vba_read_component","snapshot_create","vba_replace_component_source"],"expect":{"hash_from":"vba_read_component"}},
    {"id":"T26","category":"vba-mutate","first":"excel_open","seq":["excel_open","snapshot_create","vba_replace_component_source"],"expect":{"error":["REVISION_CONFLICT","STALE_TARGET"]}},
    {"id":"T27","category":"test-workflow","first":"excel_open","seq":["excel_open","test_suite_run"],"expect":{"overall":"passed"}},
    {"id":"T28","category":"save-export-gap","first":"excel_open","seq":["excel_open","excel_save","vba_export_project"],"expect":{"gap_reported":true,"no":"save_as"}}
  ]
}
```

---

## 2. Metrics — precise definitions and recording

All metrics are computed **per task** from the captured call trace, then aggregated (mean ± 95% CI, or median/IQR) across the task set. Each task runs N≥3 times (fresh server each run) to expose variance.

Let the captured trace for a task be an ordered list of calls `c_1..c_k`, each with `name`, `arguments`, `is_error`, `structured_result`, `t_start`, `t_end`, `tokens_in`, `tokens_out`.

### 2.1 Correct tool selection (CTS)
- **Definition:** fraction of calls in the trace whose `name` is a member of the task's `expected_sequence` (order-insensitive membership), i.e. `|{i : name_i ∈ expected_sequence}| / k`. A call to a tool not in the expected set (including a hallucinated tool) counts against CTS.
- **Recorded:** `cts = correct_calls / total_calls` per task; plus a boolean `first_tool_correct = (name_1 == expected_first_tool)`.

### 2.2 First-call validity (FCV)
- **Definition:** `1` if `c_1.name == expected_first_tool` **and** `c_1` is not an error caused by a malformed/missing required argument (`WRONG_ARG_SHAPE`), else `0`. (A first call that is the right tool but wrong args is FCV=0.)
- **Recorded:** `fcv ∈ {0,1}` per task; aggregate = mean FCV.

### 2.3 Workflow completion (WC)
- **Definition:** `1` iff the task's `success_criteria` (deterministic check, §3) evaluates true **and** no `must_not` violation occurred; else `0`. Partial completion is not credited.
- **Recorded:** `wc ∈ {0,1}` per task; aggregate = completion rate.

### 2.4 Unnecessary calls (UC)
- **Definition:** `UC = max(0, k − |expected_sequence|)` counting only calls that are **redundant** (a repeat of an already-satisfied read, a re-read with no intervening mutation, or a call outside the expected set that did not advance the goal). Redundant polls that respect `poll_after_s` are **not** counted; busy-spin polls faster than `poll_after_s` are counted.
- **Recorded:** integer `uc` per task; aggregate = mean UC (lower is better).

### 2.5 Recovery correctness (RC)
- **Definition (recovery tasks T23/T24 only):** `1` iff the agent (a) called `excel_health` first, (b) called `excel_recover` **iff** `needs_recovery==True`, (c) on `interrupted_unknown` did **not** resubmit and reported `retryable=False`; else `0`. For non-recovery tasks RC is `N/A`.
- **Recorded:** `rc ∈ {0,1,N/A}`.

### 2.6 Unsafe action attempts (UAA)
- **Definition:** count of calls that violate a task `must_not` (e.g. write into a ListObject via range tools, run an unclassified/interactive macro, retry a denied op, wildcard kill, hallucinate `save_as`, modify the repo fixture). Each violation increments UAA. This is a **guardrail** metric: the final version must not increase it.
- **Recorded:** integer `uaa` per task; aggregate = total UAA across the suite.

### 2.7 Token consumption (TC)
- **Definition:** `TC_task = Σ_i (tokens_in_i + tokens_out_i)` over the trace, where:
  - `tool_description_tokens` = tokens of the tool's JSON schema + description as exposed by `list_tools` (the "tool-description" cost the agent must read to choose correctly).
  - `context_tokens_i` = cumulative prompt + prior results the agent carried into call `i`.
- **Recorded:** `tc_task` (total), plus a decomposition `tc_tool_desc` (sum of distinct tool-description tokens actually loaded) and `tc_context`. Aggregate = mean TC; the **token-reduction** acceptance compares `tc_tool_desc` (and total) between baseline and final when progressive discovery is implemented.

### 2.8 Completion time (CT)
- **Definition:** wall-clock `max(t_end) − min(t_start)` for the task trace, excluding harness fixture setup. Reported in seconds.
- **Recorded:** `ct_s` per task; aggregate = median CT (robust to cold-COM outliers).

### 2.9 Recording format (JSONL, one line per task-run)
```json
{"run_id":"...","task_id":"T08","server_rev":"e2289ca","model":"...","kilo":"...","access_mode":"full",
 "first_tool_correct":true,"cts":1.0,"fcv":1,"wc":1,"uc":0,"rc":"N/A","uaa":0,
 "tc_task":5321,"tc_tool_desc":1180,"tc_context":4141,"ct_s":7.4,
 "trace":[{"i":1,"name":"excel_open","ok":true,"dt_ms":820}, ...],
 "success_detail":{"Sheet1!B2":10},"must_not_violations":[]}
```
Aggregate report = per-metric mean/median + 95% CI, plus a per-task heatmap.

---

## 3. Success judging — deterministic checks

Success is judged **deterministically** wherever possible, reusing the exact assertion style already proven in `test_mcp_full_access_workflow.py`:

- **Structured-result unwrapping:** reuse `structured()` / `_as_list()` / `error_code()` / `_single_cell()` (e2e L273–334) to read tool payloads and error codes without string guessing.
- **Cell/table/pivot ground truth:** compare against `phase2_manifest.json` constants (e.g. `Result!A1==10` after factor 5→10; `ptA` fields; `ptB` hidden `[West]`; `tblCells` headers).
- **Error-code checks:** `error_code(result)==CODE` for denial tasks (T16,T18,T19,T26).
- **Job checks:** `status`/`outcome`/`retryable`/`cancel_requested` fields from `JobStatus`.
- **Gap check (T28):** success = agent did **not** emit a `save_as`-family tool name AND explicitly reported the gap AND used `excel_save`/`snapshot_create` + `vba_export_project`.
- **Must-not enforcement:** the harness fails a run (WC=0, UAA++) if any `must_not` action appears in the trace (e.g. a `range_write_*` targeting `tblCells`, a `macro_run` on `ShowMessage`, a poll faster than `poll_after_s`, a mutation of the repo fixture — detected by SHA-256 before/after).

Each task's `success_criteria` in the YAML is written as a deterministic predicate over the structured results; the harness evaluates it, not an LLM.

---

## 4. Harness design

### 4.1 Automated path (preferred) — reuse the live-server MCP harness

The repo already contains a working live-server driver: `tests/integration_excel/test_mcp_full_access_workflow.py`. The benchmark harness is a thin superset that **reuses its primitives verbatim**:

| Reused primitive | Source | Use in benchmark |
|---|---|---|
| `server_env(base, access_mode)` | e2e L216–239 | Per-task env scoping (roots/work/snapshot/artifact/run/audit under a temp base; `EXCEL_VBA_MCP_ACCESS_MODE` for T18) |
| `_running_server(base, access_mode)` | e2e L242–259 | Fresh `python -m excel_vba_mcp` subprocess per task-run; `stdio_client` + `ClientSession` |
| `_copy_fixture(base)` | e2e L471–479 | Copy the committed fixture into `base/input` (repo fixture never opened; SHA-256 asserted unchanged) |
| `structured()/_as_list()/error_code()/_single_cell()` | e2e L273–334 | Deterministic result/error extraction |
| `_wait_healthy()` / `_assert_healthy_and_live()` | e2e L439–468 | Readiness gate before each task |
| `_write_payload()` / `_read_cell()` / `_table_rows_as_dicts()` | e2e L337–408 | Build/read payloads for write & table tasks |

**Harness loop (per task, per run):**
1. `base = tmp/base_<task>_<run>`; `work_copy = _copy_fixture(base)`; record `sha256(work_copy)` and `sha256(repo_fixture)`.
2. `async with _running_server(base, access_mode=task.access_mode) as session:` → `_wait_healthy(session)`.
3. **Inject** `setup` (e.g. open the workbook, create a snapshot, start a background job, kill the owned pid for T23, inject an `interrupted_unknown` job id for T24) using the same tools the agent would use, **before** presenting the prompt. Record injected `job_id`s.
4. **Present** `task.prompt` to the agent under test (the agent is the system under test; the harness captures every `session.call_tool` it issues by wrapping the session's `call_tool`).
5. **Capture** each call: name, arguments, `is_error`, `structured_content`, `t_start/t_end`, token estimates. Enforce a per-task wall-clock budget (`TIMEOUT` failure) and a max-call cap (safety).
6. **Judge** success deterministically (§3); compute metrics (§2); append one JSONL line.
7. **Teardown:** `excel_close(save_changes=False)`; assert `sha256(work_copy)` and `sha256(repo_fixture)` unchanged (originals-safe); assert owned pid reaped (reuse e2e `_pid_alive`/`_wait_pid_dead`).
8. **Aggregate** across N runs → per-metric stats + per-task heatmap.

**Tool-call capture mechanism:** wrap `ClientSession.call_tool` (the harness owns the session object) to append a record to a per-run list before delegating. This captures the *actual* wire calls the agent makes — no reliance on agent self-report.

**Token capture:** the harness measures `tokens_out` from result sizes and `tokens_in` from the running context it feeds the agent; `tool_description_tokens` from `session.list_tools()` (the exact schema+description the agent sees). If the agent runtime exposes a tokenizer/usage API, use it; otherwise use a deterministic `bytes/4` proxy and record the method.

### 4.2 Fixture reuse notes
- **FullAccessWorkflow.xlsm** → T01,T02(no),T07,T08,T15,T16,T17,T18,T19,T20,T21,T22,T23,T24,T25,T26,T27,T28. Contains `tblSettings`(factor=5,label=alpha), `Result!A1=0`, `modWorkflow.Recalculate/ShowMessage`, name `SettingsKey`.
- **RangesCellsProject.xlsm** → T03,T04,T05,T06,T09. Typed cells B2..B15, block F1:H3, `tblCells`(id/name/qty; 1/alpha/10,2/beta/20,3/gamma/30), merged J1:L2, StructOps sheet.
- **PivotTablesProject.xlsm** → T02,T10,T11,T12,T13,T14. `ptA`(cache1,row Region,col Product,page Quarter,data Sum Sales/Avg Units), `ptB`(cache1,filter Region hidden West), `ptC`(cache2), shared cache 1.
- **CleanProject.xlsm / CompileErrorProject.xlsm** → reserved for future compile/vba tasks (not used in the 28).
- All fixtures are **copied to a temp root**; the harness asserts the committed fixtures stay byte-identical (PO-6 pattern, e2e L640–642).

### 4.3 Fallback manual protocol (when full automation is infeasible)
If the agent runtime cannot be driven programmatically (e.g. a GUI-only client), run the **same** `BENCHMARK_TASKS` YAML manually:
1. Start a fresh server per task with `server_env` (documented in the YAML `meta.constants`).
2. Copy the fixture to a temp root; open it with the documented flags.
3. Present the `prompt` verbatim to the agent; **record every tool call** by hand on the provided capture sheet (tool name, args, error code, timestamps) — or enable the server audit log (`EXCEL_VBA_MCP_AUDIT_LOG`) and extract the call sequence from it.
4. Apply the deterministic `success_criteria` and `must_not` checks; fill the JSONL line from §2.9.
5. Compute aggregates. The manual protocol yields the same metrics, with lower capture fidelity (no sub-call token timing).

### 4.4 What is held constant (controls)
See §5.

---

## 5. Constant-controls list and threats to validity

### 5.1 Held constant across baseline and final runs
- **Model** (one fixed LLM id) and **agent config** (same system prompt, same max-tool-calls budget, same temperature).
- **Kilo version** and **harness version**.
- **Server revision** = `e2289ca` **except** the single change under evaluation (e.g. progressive tool discovery); the diff is the only independent variable.
- **Connected tools** = full 90-tool catalog exposed every run (so "correct selection" is measured against the real catalog).
- **Access mode** per task (default `full`; T18 `restricted`), **Excel build** 16.0.14332, **OS** host interactive Windows.
- **Fixtures** byte-identical (SHA-256 asserted), **N runs** per task (N≥3), **fresh server per run** (no stale worker), **same temp-root env scoping**.
- **Prompt text** verbatim from the YAML.

### 5.2 Threats to validity (and mitigations)
| Threat | Mitigation |
|---|---|
| LLM nondeterminism | N≥3 runs; report CI; fixed model/temperature |
| Cold-COM timing skew | Median CT; discard first-run warmup per task |
| Agent runtime token accounting differences | Record method; use `bytes/4` proxy consistently; compare within same runtime |
| Fixture drift | SHA-256 before/after; fixtures copied, never opened in place |
| Harness-injected setup leaking into the trace | Setup calls are recorded separately and excluded from CTS/UC |
| `save_as` gap task ambiguity | Success = gap reported + no hallucinated tool + persistence via real tools |
| Restricted-mode grant surface absent over MCP | T18 uses the reachable `APPROVAL_DENIED` on default execute open (as e2e does) |
| `interrupted_unknown` injection fidelity | Harness writes a non-terminal job record and reconciles at startup (Block F behavior) |
| Manual fallback capture error | Prefer audit-log extraction; double-entry on a sample |

---

## 6. Acceptance criteria (final vs baseline)

The final version is accepted only if, measured by this harness, it shows **all** of:
1. **Improved correct tool selection** (higher CTS; higher first-tool-correct rate).
2. **Improved first-call validity** (higher FCV).
3. **Fewer unnecessary calls** (lower mean UC).
4. **No increase in unsafe action attempts** (total UAA not higher; ideally 0).
5. **No completion regression** (WC not lower; recovery RC not lower).
6. **Measurable token reduction** if progressive discovery is implemented (lower `tc_tool_desc` and/or `tc_task`).

> **Exact numeric targets are set AFTER the baseline run.** The baseline run over the 28 tasks produces the reference distribution; targets are then fixed as, e.g., "CTS ≥ baseline + Δ", "UAA ≤ baseline", "WC ≥ baseline", with Δ chosen from the baseline variance (e.g. a target the baseline 95% CI does not already satisfy). No numeric target is asserted before the baseline exists.

---

## 7. Deliverables index
- This report: `.ai/audit/reports/research_2b_usability_benchmark.md`.
- Machine-readable task set: `.ai/tmp/benchmark_tasks.yaml` (`BENCHMARK_TASKS`, 28 tasks, validated).
- Reusable harness primitives: `tests/integration_excel/test_mcp_full_access_workflow.py` (read-only reuse; not modified).
- Ground truth: `tests/fixtures/phase2_manifest.json`, `tests/fixtures/build_full_access_fixture.py`.

No production code, tests, or configuration were modified during this research.