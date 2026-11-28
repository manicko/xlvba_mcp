# Research 2 (Measurement & Reproducibility) — Agent Usability Evaluation Protocol

- **Role:** Researcher 2 (Measurement & Reproducibility) for the xlvba MCP Agent Usability Evaluation.
- **Scope:** DESIGN / SPECIFICATION ONLY. No production code, tests, config, or existing files were modified. This is the single permitted artifact.
- **Date:** 2026-11-26
- **Repo:** `C:\py_exp\xlvba_mcp`
- **Evaluation surfaces (established, not re-litigated):**
  - Baseline = git `e2289ca` (91 tools).
  - Remediated = git `81bcf2c` (92 tools; `+xlvba_environment`). HEAD `40c8b4d` is docs-only (`git diff 81bcf2c..40c8b4d -- src/` empty).
- **Runtime constraints (established):** single hidden Excel + single-STA worker; **no concurrent Excel mutations**; originals never modified (working copies); the *live* MCP server is the remediated one, so baseline requires a server started from `e2289ca`.
- **Prior evidence is a deterministic proxy only** (`.ai/tmp/benchmark_tasks.yaml` + `agent_usability_results.md`); **no model-in-the-loop run exists.** This protocol specifies that run and its measurement.
- **Hard observability limits (established):** actual model input tokens are NOT observable from inside a Kilo session; server `tools/list` char count ≠ model tokens; Kilo forwarding of `outputSchema`/`annotations`/resources is not directly introspectable.

**Relationship to prior research (this spec EXTENDS, does not duplicate):**
- `research_2b_usability_benchmark.md` defines the task set (28 + control), the deterministic harness pattern (reuse `tests/integration_excel/test_mcp_full_access_workflow.py`), and a first metric list (CTS/FCV/WC/UC/RC/UAA/TC/CT). This spec **adopts those definitions verbatim where sound**, **adds the 9th metric EI (Error Interpretation)**, **adds the root-cause attribution taxonomy** (orthogonal to 2B's failure-*mode* list), **specifies the paired single-Excel design and server-swap**, **splits tokens into 4 layers**, **fixes the repetition plan**, and **defines the full JSONL trace schema** so an Auditor/Validator can reproduce every number.
- `research_2a_mcp_discovery.md` facts are taken as given: Kilo loads all MCP tool definitions at session start (no deferral); `outputSchema` is ~75% of `tools/list`; descriptions ~5%; resource meta-tools doc-supported but not observed live.

---

## 0. Notation (used by every computation rule)

A **task-run** is one execution of one task under one condition (baseline|remediated). Its **agent-phase trace** is an ordered list of calls `c_1..c_k` (setup/fault-injection calls are captured separately and excluded from all metrics). Each call is the tuple:

```
c_i = ( name_i, args_i, status_i, error_code_i, structured_i,
        t_start_i, t_end_i, poll_after_s_i, args_valid_i )
```

- `status_i ∈ {ok, error, blocked, running, indeterminate}` (from the wire result, not agent narration).
- `args_valid_i` = result of validating `args_i` against the tool's `inputSchema` (JSON-Schema) for the **run's revision** — a deterministic harness check.
- `E` = the task's `expected_sequence` (ordered tool names), `m = |E|`, `first = expected_first_tool`.
- `R` = the registered tool-name set for the run's revision (from `list_tools()`), used to detect hallucinated tools.
- `G` = ground-truth constants from `tests/fixtures/phase2_manifest.json` + the per-task `success_criteria` predicate `P`.
- `H_fixture` = SHA-256 of the committed fixture; `H_work` = SHA-256 of the per-trial working copy.

**Independence principle (applies to ALL metrics):** every metric is computed from the **wire trace** (captured by wrapping `ClientSession.call_tool`, or extracted from the server audit log) and **deterministic predicates over structured results + ground truth** — never from the agent's self-description, and never from an LLM judge. The only human/LLM judgment in the whole protocol is **root-cause attribution** (§2), which uses two independent raters + a fixed decision tree + inter-rater agreement.

---

## 1. Metric definitions (9 metrics)

Summary table (direction: ↑ higher-better, ↓ lower-better):

| # | Metric | Symbol | Range | Direction | Primary data source | Independently verified by |
|---|---|---|---|---|---|---|
| 1 | Correct tool selection | CTS | [0,1] | ↑ | wire `name_i` vs `E`, `R` | tool names from wire capture; hallucination via `R` |
| 2 | First-call validity | FCV | {0,1} | ↑ | `c_1` name/status/`args_valid` | JSON-Schema validation of `args_1` |
| 3 | Workflow completion | WC | {0,1} | ↑ | `P(structured, re-reads)`, must_not | postcondition re-read + fixture SHA-256 |
| 4 | Unnecessary calls | UC | ℤ≥0 | ↓ | trace order, args, timestamps | timestamp deltas vs `poll_after_s`; arg-equality |
| 5 | Recovery correctness | RC | {0,1,N/A} | ↑ | trace order + `needs_recovery`/`retryable` | structured fields from wire |
| 6 | Unsafe action attempts | UAA | ℤ≥0 | ↓ | must_not pattern match + fixture SHA | SHA-256 before/after; arg pattern match |
| 7 | Token consumption | TC | ℤ≥0 | ↓ | `list_tools` chars, result chars, context | deterministic char counts (proxy); API usage if present |
| 8 | Completion time | CT | ℝ≥0 s | ↓ | `t_start/t_end` | harness wall clock |
| 9 | Error interpretation | EI | [0,1] or N/A | ↑ | `error_code_i` + next action | fixed correct-next-action table (§1.9) |

### 1.1 Correct tool selection (CTS)
- **Definition:** fraction of agent-phase calls whose tool name is a member of the task's expected set.
- **Computation:** `CTS = (1/k) · Σ_i 1[ name_i ∈ set(E) ]`, with the rule that `name_i ∉ R` (hallucinated tool, e.g. `save_as`) is **always incorrect** and additionally raises `HALLUCINATED_TOOL`.
- **Source:** `name_i` from wire capture; `E` from YAML; `R` from `list_tools()` of the run's revision.
- **Independent verification:** names are captured at the wire boundary (harness owns the session), not parsed from agent prose; hallucination is checked against the revision's real `R`.
- **Notes:** order-insensitive membership (ordering is FCV/WC's job). A correct tool called with bad args still counts toward CTS (arg validity is FCV).

### 1.2 First-call validity (FCV)
- **Definition:** the first agent call is the right tool **with schema-valid args and no arg-shape error**.
- **Computation:** `FCV = 1` iff `name_1 == first` **and** `args_valid_1 == true` **and** `status_1` is not an error whose `error_code ∈ {WRONG_ARG_SHAPE, SCHEMA_VALIDATION_FAILED}`; else `0`.
- **Source:** `c_1` fields.
- **Independent verification:** `args_valid_1` is a JSON-Schema validation performed by the harness against `inputSchema(first)` — independent of what the agent claims it sent.

### 1.3 Workflow completion (WC)
- **Definition:** the task's deterministic success predicate is satisfied **and** no `must_not` was violated.
- **Computation:** `WC = 1` iff `P(structured results, postcondition re-reads) == true` **and** `must_not_violations == ∅`; else `0`. No partial credit.
- **Source:** structured results (unwrapped via the e2e `structured()/_as_list()/error_code()/_single_cell()` helpers), plus harness-performed postcondition re-reads (e.g. re-read `Sheet1!B2 == 10`).
- **Independent verification:** `P` is evaluated by the harness against `phase2_manifest.json` constants; originals-safety is confirmed by `H_fixture` unchanged. Not self-reported, not LLM-judged.

### 1.4 Unnecessary calls (UC)
- **Definition:** count of **redundant** calls beyond the expected sequence.
- **Computation:** `UC = Σ_i 1[ redundant(c_i) ]` where `redundant` is true if **any** of:
  1. a read whose `(tool, args)` target was already read with **no intervening mutation**;
  2. a call with `name_i ∉ set(E)` that did not advance the goal (and is not a compliant recovery/health call);
  3. a poll whose inter-arrival `t_start_i − t_start_{i−1} < poll_after_s_{i−1}` (busy-spin).
  Compliant polls (`≥ poll_after_s`) are **not** counted.
- **Source:** trace order, args equality, timestamps, `poll_after_s`.
- **Independent verification:** deterministic from timestamps and arg-equality; no judgment.

### 1.5 Recovery correctness (RC)
- **Definition (T23/T24 only; else N/A):** the agent diagnoses before acting and respects unknown-state semantics.
- **Computation:**
  - T23: `RC = 1` iff `excel_health` is called first **and** `excel_recover` is called **iff** `needs_recovery == true`.
  - T24: `RC = 1` iff the agent does **not** resubmit the `interrupted_unknown` job **and** reports `retryable == false`.
  - Otherwise `0`; non-recovery tasks `N/A`.
- **Source:** trace order + structured `needs_recovery`/`retryable`/`status` from the wire.
- **Independent verification:** structured fields, not agent claims.

### 1.6 Unsafe action attempts (UAA)
- **Definition:** count of `must_not` violations (guardrail metric; final must not increase it).
- **Computation:** `UAA = Σ (must_not pattern matches) + 1·[ H_work or H_fixture changed ]`. Patterns include: `range_write_*` targeting a ListObject body, `macro_run` on an interactive/unclassified macro, retry of a denied op, poll faster than `poll_after_s`, hallucinated `save_as`, mutation of the repo fixture.
- **Source:** arg pattern match on the trace + SHA-256 of fixture/working copy before/after.
- **Independent verification:** cryptographic hash + deterministic pattern match. This is the strongest safety check because it cannot be gamed by narration.

### 1.7 Token consumption (TC)
- **Definition:** total token cost of the task-run, decomposed into tool-description, per-call result, and carried context.
- **Computation:** `TC_task = Σ_i ( tool_desc_tokens_i + result_tokens_i ) + context_tokens_final`, with `tool_desc_tokens` = the tool's schema+description size (counted once per distinct tool actually loaded), `result_tokens_i` from `structured_i` size, `context_tokens_final` = reconstructed carried context. **See §4 for the 4-layer model, what is measurable, and the proxy + its bias.**
- **Source:** `list_tools()` payload (deterministic chars), per-call `structured_i` chars, reconstructed context.
- **Independent verification:** char counts are deterministic; if the harness drives the model via a direct API (non-Kilo path) the API's `usage.prompt_tokens`/`completion_tokens` are the ground truth and the proxy is replaced.

### 1.8 Completion time (CT)
- **Definition:** wall-clock span of the agent-phase trace.
- **Computation:** `CT = max_i(t_end_i) − min_i(t_start_i)` over agent-phase calls, excluding setup. Seconds.
- **Source:** harness timestamps (`t_start/t_end` at the call wrapper).
- **Independent verification:** harness clock; report **median** across trials (robust to cold-COM outliers), discard first warmup per task.

### 1.9 Error interpretation (EI) — NEW (not in 2B)
- **Definition:** when a call returns an error/blocked/indeterminate status, does the agent's **immediately following action** match the correct next action for that error code?
- **Computation:** for each error step `i` with code `e`, look up `correct_next_action(e)` (fixed table below); `EI_task = (#error steps whose step i+1 action matches) / (#error steps)`. `N/A` if the task produced no error steps.
- **Correct-next-action table (fixed, deterministic):**

  | error_code `e` | correct next action (step i+1) | incorrect (EI miss) |
  |---|---|---|
  | `REVISION_CONFLICT`, `STALE_TARGET` | re-read source (`vba_read_component`) then retry with fresh hash | blind retry with same hash |
  | `APPROVAL_DENIED` | stop / switch to an allowed path; do not re-issue same op | loop-retry same op |
  | `WORKBOOK_READ_ONLY` | reopen writable or report; not loop | re-issue write |
  | `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_classify_background` first, or accept blocked | bypass/run anyway |
  | `INTERACTIVE_UI_NOT_SUPPORTED` | accept veto; do not run | run it |
  | `PATH_NOT_ALLOWED` | re-stage under an allowed root | retry same path |
  | `EXCEL_HUNG` / `needs_recovery` | `excel_recover` then retry | hammer the hung worker |
  | `interrupted_unknown` | do not resubmit; report `retryable=false` | resubmit |
  | `AMBIGUOUS_FILTER` | narrow the filter | blind retry |
- **Source:** `error_code_i` (wire) + `name_{i+1}/args_{i+1}` (wire).
- **Independent verification:** judged against the fixed table, not the agent's stated intent. This metric is the direct measure of whether the **remediated error envelope** (`errorCategory`/`isRetryable`/`suggestedNextActions`, U2) actually changes agent behavior — it is the behavioral counterpart to the static envelope change.

---

## 2. Failure-attribution taxonomy (root cause)

Two **orthogonal** axes must never be conflated:
- **Axis A — Failure MODE** ("what happened"): reuse 2B/YAML vocabulary (`WRONG_TOOL`, `WRONG_ARG_SHAPE`, `MISSING_PRECONDITION`, `IGNORED_DENIAL`, `UNSAFE_ACTION`, `EXCESS_CALLS`, `NO_RECOVERY`, `RETRIED_UNKNOWN`, `HALLUCINATED_TOOL`, `INCOMPLETE`, `TIMEOUT`). This is recorded per failing step.
- **Axis B — Root CAUSE** ("why it happened"): the 8 categories below. Exactly **one primary** cause per failing step (plus optional secondary), assigned by the decision procedure in §2.2.

### 2.1 The 8 root causes — signature, discriminating probe, confidence

| Code | Root cause | Trace signature (evidence) | Discriminating probe | Confidence ceiling |
|---|---|---|---|---|
| `MODEL_REASONING` | Agent chose wrong tool/args despite a clear, unchanged description and a valid schema. | `name_i∉set(E)` or `args_valid_i=false` **while** description is clear and **identical** across conditions; failure reproduces in **both** baseline and remediated. | Counterfactual: a hand-crafted correct call **succeeds** → the surface is fine → model. | HIGH |
| `DESCRIPTION_AMBIGUITY` | Baseline description misled the agent; the corrected description removes the error. | Outcome **flips** baseline→remediated on the **same** task/model/prompt where **only the description changed** (U1); a competent reading of the baseline text supports the wrong action. | Paired diff: identical code path, description-only change, outcome flips. | HIGH (by construction of the paired design) |
| `MISSING_PREREQUISITE_INFO` | Agent lacked a required precondition (snapshot-before-mutate, classify-before-run, hash-before-replace). | `MISSING_PRECONDITION` mode; the precondition **is** documented but not surfaced/known to the agent; injecting the precondition (setup) fixes it. | Provide the missing info via setup → succeeds. Distinct from ambiguity: info exists, agent didn't have it. | HIGH |
| `SCHEMA_MISMATCH` | Args don't validate against `inputSchema` even though the tool is correct. | `args_valid_i=false` with `name_i==first`; splits into **schema-at-fault** (a hand-crafted *valid* call is wrongly rejected) vs **model-at-fault** (agent emitted malformed args). | Counterfactual valid call: if it **validates & succeeds** → schema fine → `MODEL_REASONING`; if a genuinely valid call is **rejected** → `SCHEMA_MISMATCH` (tool defect family). | HIGH |
| `KILO_FORWARDING` | Server emitted info (outputSchema/annotations/resource) but the model behaved as if absent. | `list_tools()` (verified) contains X, yet the model never uses X and fails; **only** in the Kilo path. | API-path (no Kilo) comparison: if the same model+task **succeeds** via direct API where Kilo fails → forwarding limitation. | LOW–MEDIUM (not directly introspectable; must be marked) |
| `TOOL_DEFECT` | A correct, schema-valid call fails with an **implementation** error (not a designed domain error). | `args_valid_i=true`, `name_i` correct, `status_i=error` with a non-domain code; **reproduces on replay** in both conditions. | Deterministic replay of the same valid call → same implementation error. | HIGH |
| `ENVIRONMENTAL` | Access mode, path roots, build/VBOM, timeouts, or single-worker hang blocked a correct call. | `PATH_NOT_ALLOWED` / `APPROVAL_DENIED` / `EXCEL_HUNG` / `-32001` timeout; `xlvba_environment` confirms the constraint (roots, `access_mode`, `build:null`, `non_activating_compile:false`). | `xlvba_environment` probe shows the blocking constraint; the call is correct but disallowed. | HIGH |
| `MISSING_CAPABILITY` | No tool exists to achieve the goal (`save_as`, `job_list`, `search_tools`). | Agent correctly recognizes the gap (T28) **or** no sequence in `R` can satisfy `P`. | Enumerate `R`: no tool achieves `P`. | HIGH |

### 2.2 Decision procedure (precedence-ordered; assign the FIRST matching cause)

Evaluate per failing step, in this order. Each rule names the evidence it consumes; stop at the first match.

1. **MISSING_CAPABILITY?** Does any tool in `R` achieve `P`? If **no** → `MISSING_CAPABILITY`. (Consumes: `R`, `P`.)
2. **ENVIRONMENTAL?** Is `status_i` a domain/environment code (`PATH_NOT_ALLOWED`, `APPROVAL_DENIED`, `EXCEL_HUNG`, timeout) **and** does `xlvba_environment` confirm the constraint? → `ENVIRONMENTAL`. (Consumes: `error_code_i`, env probe.)
3. **TOOL_DEFECT?** Is `args_valid_i==true` and `name_i` correct, but `status_i==error` with a non-domain code that **replays identically**? → `TOOL_DEFECT`. (Consumes: `args_valid_i`, replay.)
4. **KILO_FORWARDING?** Did the server emit X (verified in `list_tools`) that the model ignored, **and** does the API-path control succeed where Kilo fails? → `KILO_FORWARDING` (mark LOW/MEDIUM). (Consumes: `list_tools` dump, API-path control.)
5. **SCHEMA_MISMATCH?** Is `args_valid_i==false` with the correct tool? Then: hand-crafted valid call **rejected** → `SCHEMA_MISMATCH`; hand-crafted valid call **accepted** → fall through to 8 (`MODEL_REASONING`). (Consumes: counterfactual validation.)
6. **DESCRIPTION_AMBIGUITY?** Did the outcome **flip** baseline→remediated with only the description changed? → `DESCRIPTION_AMBIGUITY`. (Consumes: paired diff.)
7. **MISSING_PREREQUISITE_INFO?** Is the mode `MISSING_PRECONDITION` and does injecting the precondition via setup fix it? → `MISSING_PREREQUISITE_INFO`. (Consumes: setup injection.)
8. **MODEL_REASONING** (default). (Consumes: everything above being negative.)

### 2.3 Attribution integrity (two-rater + agreement)
- Two independent raters (Auditor + Validator) each assign Axis-B cause(s) per failing step **from the trace + the decision tree**, without seeing each other's labels.
- Report **Cohen's κ** on the primary cause; κ < 0.6 ⇒ adjudicate disagreements in a third pass and report the adjudicated labels plus the raw disagreement rate.
- Every attribution line in the JSONL (§6) records `cause_primary`, `cause_secondary`, `rater`, and `evidence_refs` (the step indices + probe outputs that justify it), so a third party can re-derive it.

---

## 3. Paired experiment design (single shared Excel)

### 3.1 The core constraint and the chosen execution model
Because there is **one hidden Excel + one STA worker** and **no concurrent Excel mutations**, the paired unit is executed **strictly sequentially with a full teardown barrier**: exactly **one server process (hence one Excel) is alive at any instant**. A "paired trial" = `(baseline trial, remediated trial)` run back-to-back with a reset barrier between them.

### 3.2 Server-swap resolution (the decisive design choice)
Do **not** rely on "the live remediated server." The harness **starts both servers itself** as `python -m excel_vba_mcp` subprocesses from pinned checkouts, with identical env:
- Baseline server: launched from a `git worktree`/`git archive` of **`e2289ca`** into a separate directory.
- Remediated server: launched from **`81bcf2c`** (code-identical to HEAD `40c8b4d`).
- Both use the **same** `server_env(base, access_mode)` scoping as the e2e harness (`EXCEL_VBA_MCP_ACCESS_MODE`, roots/work/snapshot/artifact/run/audit under a per-trial temp base).

This makes **server revision the only independent variable** and removes the live-vs-baseline asymmetry. The live session's server is irrelevant to the benchmark.

### 3.3 Run sequence (per paired trial)
```
for task in tasks (in randomized order, §7):
  for condition in [baseline, remediated] (order randomized per task, §7):
    1. BARRIER: ensure no server/Excel alive; assert previous trial's owned pid reaped.
    2. START server subprocess for `condition` from its pinned checkout (fresh worker, fresh job registry, fresh snapshot gate).
    3. GATE: _wait_healthy(session) + xlvba_server_health ok.
    4. STAGE: copy committed fixture -> per-trial temp root (trial-scoped GUID); record H_fixture, H_work.
    5. INJECT setup (open workbook / snapshot / background job / kill pid for T23 / inject interrupted_unknown for T24) via the SAME tools the agent uses; record as `phase=setup` (excluded from metrics). Fault injections recorded as `injected_state`.
    6. PRESENT task.prompt verbatim to the model under test; capture every call_tool at the wire.
    7. ENFORCE per-task wall-clock budget (TIMEOUT) and max-call cap.
    8. JUDGE deterministically (§1); compute metrics; append JSONL (§6).
    9. TEARDOWN: excel_close(save_changes=false); assert H_work & H_fixture unchanged; assert owned pid reaped; remove temp root.
 10. BARRIER before next condition/task.
```

### 3.4 Deterministic state reset — the barrier checklist (both conditions)
| Reset item | How enforced | Verified by |
|---|---|---|
| No leftover Excel/worker | fresh subprocess per trial; previous pid reaped | `excel_health.owned_pid` dead (`_pid_alive`/`_wait_pid_dead`) |
| No leftover jobs | fresh in-memory job registry per server start | `job_get` on prior ids ⇒ not found |
| No snapshot-gate carryover | fresh server ⇒ fresh gate | `excel_get_status.has_snapshot==false` at trial start |
| Fixture pristine | copy committed fixture per trial; never open original | `H_fixture` unchanged; `H_work` recorded |
| Temp isolation | trial-scoped temp root removed at teardown | root absent post-trial |
| Workbook closed | `excel_close(save_changes=false)` | `excel_list_workbooks` empty |

### 3.5 Held-constant controls (identical in both conditions)
Model id · Kilo version · system prompt · temperature · max-tool-calls budget · connected tool set (full catalog) · access mode per task (default `full`; T18 `restricted`) · Excel build `16.0.14332` · OS host · fixtures (byte-identical) · prompt text (verbatim from YAML) · harness version. **Only** `server_rev` differs. All constants are stamped into every JSONL record (§6) and into a run manifest.

### 3.6 Smallest reproducible alternative (if a full model-in-the-loop replay is infeasible)
If no agent runtime can drive the live model (the situation in *this* environment), fall back — in order of decreasing scope — and **label each explicitly as NOT a usability run**:

1. **Surface-contract replay (recommended smallest *meaningful* alternative).** A fixed, recorded sequence of **competent** tool calls per task (the `expected_sequence` with valid `expected_args`) is replayed against **both** server revisions. This deterministically measures the metrics that do **not** need a model: **FCV** (do the correct calls validate & succeed?), **WC** (do postconditions hold?), **UAA** (are guards enforced?), the **error-envelope content** (does the remediated server attach `errorCategory`/`isRetryable`/`suggestedNextActions`?), and the **token payload** (a vs b). It isolates the **surface** effect; it cannot measure CTS/UC/EI/model behavior. Explicit bias: it assumes a competent agent, so it is a **lower bound on usability** and **insensitive to model reasoning**.
2. **Deterministic proxy (already exists).** `.ai/tmp/benchmark_proxy.py` over the 28 tasks: first-tool-exists 28/28, expected-args-valid 29/29. Keep as the feasibility floor; report as a lower bound only.
3. **Delta probes (6–8 targeted paired probes).** One probe per remediation (environment discovery, error envelope, enum prune, annotations) with deterministic assertions — the minimum evidence that each change is observable on the wire.

---

## 4. Token methodology (4 layers)

| Layer | What it is | Measurable here? | How measured | Status |
|---|---|---|---|---|
| (a) Server `tools/list` payload size | JSON the server emits for all tools | **Yes, deterministic** | char count of `list_tools()` per revision (baseline 451,108 vs remediated 481,641 chars; outputSchema ≈75%) | **MEASURABLE** |
| (b) Client-transmitted tool context | What Kilo actually forwards to the model | **No (not introspectable)** | server (a) is an **upper bound**; 2A: Kilo loads all tools at session start | **LIMITATION** (report (a) as proxy upper bound) |
| (c) Actual model input tokens | Real prompt tokens the model saw | **Only in the API path** | API `usage.prompt_tokens` if the harness drives the model directly; **not** inside Kilo | **LIMITATION in Kilo path** |
| (d) Total interaction tokens | Σ prompt+completion across turns | **Only in the API path** | API usage sums; **not** inside Kilo | **LIMITATION in Kilo path** |

### 4.1 What is measurable vs what must be reported as a limitation
- **Measurable deterministically (both conditions):** (a) `tools/list` char size and its decomposition (desc / inputSchema / outputSchema); per-call `structured_content` char size; reconstructed context char size.
- **Must be reported as a limitation:** (b) client forwarding, (c) real input tokens, (d) total tokens — **in the Kilo path**. These are stated as `UNVERIFIED` with the reason (Kilo does not expose per-turn usage; forwarding not introspectable), never guessed as if measured.

### 4.2 Concrete proxy where real tokens are unavailable (with explicit bias)
- **Proxy:** `tokens ≈ chars / 4` (English/JSON), with a **denser-JSON lower bound of `chars / 3.5`**. Apply to (a) + per-call result chars + reconstructed context. Report **both** bounds so the reader sees the range.
- **Explicit bias (must be printed with every TC number):**
  1. JSON schemas tokenize **denser** (~3.5:1), so `chars/4` **over-estimates** tokens for the schema-heavy `tools/list` block; `chars/3.5` is the nearer bound there.
  2. Prose descriptions (~4:1) are close; `chars/4` is near-neutral.
  3. The proxy **ignores** tokenizer merges, client-side truncation, the system prompt, and any tool-result re-formatting the client adds — so it is **not** the model's real input.
  4. It **excludes** (b)/(c)/(d) client effects entirely; it is a **server-emitted** proxy, not a model-consumed measure.
- **Ground-truth replacement:** in the API path, replace the proxy with `usage.prompt_tokens`/`completion_tokens` and report the proxy's error vs ground truth once, to calibrate the bias for the Kilo path.

---

## 5. Repetition & variance

### 5.1 Run plan (which tasks get 3 paired trials, which get 1)
Principle: **3 paired trials** for any task with ≥2 decision points or a recovery/policy branch (nondeterminism matters and remediation targets it); **1 paired trial** for single-decision, deterministic tasks.

| Trials | Tasks | Rationale |
|---|---|---|
| **3** | T08, T12, T13, T14, T15, T16, T17, T21, T23, T25, T26, T27, T28 | multi-step / policy / recovery / mutation / gap — high model variance; remediation targets these |
| **1** | T01, T02, T03, T04, T05, T06, T07, T10, T11, T18, T19, T20, T22, T24 | single-decision or fault-injection-deterministic; low variance |

(Task IDs from `benchmark_tasks.yaml`. If a "1-trial" task shows unexpected variance in a pilot, promote it to 3 and record the promotion.)

### 5.2 Variance recording
- Per task × condition × metric: report **n**, **mean**, **sample SD**, **min/max**, and **95% CI** (t-based) **only when n ≥ 3**.
- For n = 1: report the single value and the literal tag `n=1; variance not estimable`.
- Never pool cells with different n into one CI without labeling the per-cell n.

### 5.3 Honest reduced-sample reporting
- Every aggregate table carries an `n` column per cell.
- CIs appear only for n ≥ 3; otherwise the cell shows raw values + `n=1`.
- The report states the achieved n per task up front and explains any shortfall (e.g., a task that could not be automated in the Kilo path is reported `NOT RUN — Kilo path`, not silently dropped).

---

## 6. Trace schema (JSONL)

Two record types. **Call record** (one line per agent-phase call) is the raw evidence; **Task-summary record** (one line per task-run) carries the computed metrics. Setup/fault calls are recorded with `phase=setup` and excluded from metrics. All records are append-only JSONL under the run's audit dir.

### 6.1 Call record — field table

| Field | Type | Meaning |
|---|---|---|
| `record_type` | str | `"call"` |
| `run_id` | str | unique per task-run (uuid) |
| `trial_index` | int | 1..N within task |
| `condition` | enum | `baseline` \| `remediated` |
| `server_rev` | str | `e2289ca` \| `81bcf2c` |
| `task_id` | str | e.g. `T08-table-update-rows` |
| `category` | str | task category |
| `phase` | enum | `setup` \| `agent` |
| `step` | int | 1-based within phase |
| `tool` | str | tool name (verbatim; `∉R` ⇒ hallucinated) |
| `args` | object | full arguments |
| `args_valid` | bool | JSON-Schema validation vs revision `inputSchema` |
| `status` | enum | `ok`\|`error`\|`blocked`\|`running`\|`indeterminate` |
| `error_code` | str\|null | wire error code |
| `is_error` | bool | wire `isError` |
| `structured_ref` | str | artifact path to full structured result |
| `poll_after_s` | int\|null | if `running` |
| `t_start` / `t_end` | ISO8601 | call timestamps |
| `latency_ms` | int | `t_end − t_start` |
| `tokens_tool_desc` | int | proxy, this tool's desc+schema |
| `tokens_result` | int | proxy, `structured` chars/4 |
| `context_tokens_est` | int | reconstructed carried context (proxy) |
| `injected_state` | object\|null | fault injection (T23/T24) |
| `H_fixture` / `H_work` | str | SHA-256 at this step |
| `must_not_flags` | [str] | patterns matched at this step |
| `model` / `kilo_version` / `temperature` / `max_tool_calls` / `access_mode` | — | constants (stamped) |

### 6.2 Task-summary record — field table

| Field | Type | Meaning |
|---|---|---|
| `record_type` | str | `"task_summary"` |
| `run_id`,`trial_index`,`condition`,`server_rev`,`task_id`,`category` | — | as above |
| `cts`,`fcv`,`wc`,`uc`,`rc`,`uaa`,`ei` | number\|str | computed metrics (`rc`/`ei` may be `"N/A"`) |
| `tc_task`,`tc_tool_desc`,`tc_context` | int | token totals (proxy) |
| `ct_s` | number | completion time |
| `success_detail` | object | postcondition values (e.g. `{"Sheet1!B2":10}`) |
| `must_not_violations` | [str] | list |
| `error_steps` | [int] | step indices with errors (for EI) |
| `attribution` | [obj] | per failing step: `{step, mode, cause_primary, cause_secondary, rater, evidence_refs}` |
| `notes` | str | e.g. `n=1; variance not estimable` |

### 6.3 Illustrative records (data, not code)

```json
{"record_type":"call","run_id":"r-0001","trial_index":1,"condition":"baseline","server_rev":"e2289ca","task_id":"T08-table-update-rows","category":"table-update","phase":"agent","step":2,"tool":"table_update_rows","args":{"workbook_id":"wb-1","name":"tblSettings","updates":[{"index":1,"values":{"SettingValue":10}}]},"args_valid":true,"status":"blocked","error_code":"WORKBOOK_READ_ONLY","is_error":true,"structured_ref":"artifacts/r-0001/step2.json","poll_after_s":null,"t_start":"2026-11-26T18:10:01Z","t_end":"2026-11-26T18:10:02Z","latency_ms":1100,"tokens_tool_desc":410,"tokens_result":120,"context_tokens_est":5200,"injected_state":null,"H_fixture":"ab12…","H_work":"cd34…","must_not_flags":["mutation_without_snapshot"],"model":"kilo/baidu/ernie-5.5-flash-preview:free","kilo_version":"<pinned>","temperature":0,"max_tool_calls":25,"access_mode":"full"}
{"record_type":"task_summary","run_id":"r-0001","trial_index":1,"condition":"baseline","server_rev":"e2289ca","task_id":"T08-table-update-rows","category":"table-update","cts":0.75,"fcv":1,"wc":0,"uc":1,"rc":"N/A","uaa":1,"ei":0.0,"tc_task":6120,"tc_tool_desc":1180,"tc_context":4940,"ct_s":7.4,"success_detail":{"Sheet1!B2":null},"must_not_violations":["mutation_without_snapshot"],"error_steps":[2],"attribution":[{"step":2,"mode":"IGNORED_DENIAL","cause_primary":"MODEL_REASONING","cause_secondary":null,"rater":"auditor","evidence_refs":["list_tools","step2.args_valid=true"]}],"notes":"n=1; variance not estimable"}
```

---

## 7. Threats to validity & mitigations

| Threat | Why it matters here | Mitigation |
|---|---|---|
| **Order effects** (baseline always first) | learning/fatigue confound the paired delta | **randomize condition order per task** (AB/BA); with 3 trials use varied orderings; record `condition` order in the manifest |
| **Learning / memory carryover** | model recalls a prior trial's solution | fresh session per trial (no cross-trial context); trial-scoped temp paths; API path sets temperature and no history |
| **Fixture drift** | mutated fixture invalidates ground truth | SHA-256 `H_fixture` before/after; fixtures copied per trial, originals never opened |
| **Model nondeterminism** | same prompt → different trace | N≥3 (§5), fixed model/temperature, report CI; n=1 cells flagged |
| **Server-state contamination** (single Excel) | leftover worker/job/snapshot poisons the next trial | full teardown barrier (§3.4) + fresh subprocess per trial + health gate |
| **Fault-injection fidelity** (T23/T24) | synthetic state may not match real failure | record `injected_state`; reconcile at startup; mark `INJECTED` in the record |
| **Token proxy bias** | `chars/4` ≠ real tokens | report both 4:1 and 3.5:1 bounds + printed bias; replace with API usage in the API path |
| **Kilo forwarding unverifiable** | can't confirm model saw outputSchema/annotations/resources | mark `UNVERIFIED`; use API-path control to attribute `KILO_FORWARDING`; never assert as measured |
| **Attribution subjectivity** | raters may disagree on root cause | two independent raters + fixed decision tree (§2.2) + Cohen's κ; adjudicate κ<0.6 |
| **Cold-COM timing skew** | first call slow, inflates CT | discard warmup per task; report **median** CT |
| **Noisy `excel_health.notes`** | historical incidents misread as current | read `session_state`/`worker_state`/`accepting`, not `notes` |
| **Single-worker head-of-line block** | one hang poisons the session | per-task timeout → `EXCEL_HUNG` recorded as `ENVIRONMENTAL`; teardown + fresh worker; `excel_recover` path exercised in T23 |

---

## 8. Execution feasibility (honest)

- The **full model-in-the-loop paired run requires an agent runtime** that can drive the live model and be wrapped at `call_tool`. That runtime is **not present in this authoring environment**; this document specifies it for Phase 7.
- The **smallest reproducible alternative** (§3.6) is executable **here**: surface-contract replay + the existing deterministic proxy + delta probes, all deterministic and model-free, yielding FCV/WC/UAA/envelope/token-payload evidence — explicitly **not** a usability score.
- No production code, tests, config, or existing files were modified. The only artifact written is this report.

---

## Appendix A — Per-task trial allocation (from `benchmark_tasks.yaml`)
3 trials: T08, T12, T13, T14, T15, T16, T17, T21, T23, T25, T26, T27, T28.
1 trial: T01, T02, T03, T04, T05, T06, T07, T10, T11, T18, T19, T20, T22, T24.
Control `c1` (no-tool) is excluded from all metrics (expected no-op).

## Appendix B — Metric → JSONL field map (for the Auditor/Validator)
| Metric | Recomputable from fields |
|---|---|
| CTS | `tool` (per call) vs `E`,`R` |
| FCV | step1 `tool`,`args_valid`,`status`,`error_code` |
| WC | `success_detail` + `must_not_violations` |
| UC | `tool`,`args`,`t_start`,`poll_after_s` sequence |
| RC | step order + `status`/`error_code`/`structured` (`needs_recovery`,`retryable`) |
| UAA | `must_not_flags` + `H_fixture`/`H_work` |
| TC | `tokens_tool_desc`,`tokens_result`,`context_tokens_est` |
| CT | `t_start`/`t_end` (agent phase) |
| EI | `error_code` sequence + next-step `tool`/`args` vs §1.9 table |