# Report 09a — Independent Validator Verdict

- **Role:** Independent Validator (Phase 6). Did NOT run the main benchmark. READ-ONLY on production code.
- **Date:** 2026-11-26
- **Repo:** `C:\py_exp\xlvba_mcp`
- **Method:** (1) read all 8 required artifacts + `benchmark_tasks.yaml`; (2) source verification of the T05 description/implementation claim (`server.py`, `ranges_service.py`, `ranges_adapter.py`, `test_description_contract.py`); (3) live reproduction of 5 tasks through the real `xlvba_*` MCP server (T03, T05, T16, T17, T26) on working copies; (4) git revision re-verification; (5) fixture SHA-256 before/after.
- **Safety:** No production/test/config file modified. Originals byte-identical after reproduction. No concurrent Excel (single server, sequential, workbook closed between tasks). No wildcard kill, no bypass, no `INTERRUPTED_UNKNOWN` retry. This is the only artifact written.

---

## 1. Mandate 1 — Experiment design & comparability

**Finding:** The baseline/remediated comparison is **NOT a like-for-like model comparison**, and the reports say so plainly and repeatedly.

- `agent_usability_benchmark_results.md` §2 ("Critical scope limitation") states the baseline is a **DETERMINISTIC characterization, not a model run**, because the Kilo client is bound to one live (remediated) server, global config is unreadable (`external_directory: deny`), and the single-Excel/single-STA rule forbids a concurrent second server.
- `agent_usability_failure_analysis.md` F7 records the same limitation and classifies the strict paired metric as **INCONCLUSIVE**.
- `research_eval_2_measurement_design.md` §3.6/§8 pre-registered this exact fallback ("surface-contract replay" / "deterministic proxy") and the reduced-scope labeling *before* results existed.

**Verdict:** The limitation is **honestly disclosed and the conclusions are appropriately scoped.** The headline claims are restricted to (a) the remediated surface is usable end-to-end by a real agent, and (b) deterministic surface deltas are favorable with no safety regression. No causal "remediation improved the model" claim is made. **APPROVED** on design honesty.

**Residual concern:** a casual reader skimming only the "100% workflow completion" line (results §4) could over-read it as a comparative win. The qualifier exists but is in §2/§6/§8, not next to the number. See Required Fixes.

---

## 2. Mandate 2 — Independent reproduction (5 tasks; ≥1 failure reproduced)

All reproductions used `excel_open` → hidden working copy (`source_path` = fixture, `path` = temp `work\…`), so originals were never opened in place.

| Task | Claim under test | My independent result | Match? |
|---|---|---|---|
| **T03** read range | 14×1; kinds incl. string/number/boolean/empty/error; B8=`"=1+1"` text; B10 error | `rows=14, cols=1`; kinds string/number/boolean/empty/error/date present; B8 kind=string value `=1+1`; B10 kind=error `#DIV/0!` | **CONFIRMED PASS** |
| **T05** write 2×2 @ top-left | first-mutation `INVALID_RANGE`, self-corrected | `address="D2"` + 2×2 → `INVALID_RANGE` "Payload shape 2x2 does not match the target 1x1"; `D2:E3` → `verified:true` | **CONFIRMED (defect reproduced)** |
| **T16** interactive veto | `INTERACTIVE_UI_NOT_SUPPORTED`, no retry | `code=INTERACTIVE_UI_NOT_SUPPORTED, isRetryable=false` | **CONFIRMED PASS** |
| **T17** unclassified blocked | `status=blocked`, `BACKGROUND_SAFETY_NOT_VERIFIED`, macro did NOT run | `status=blocked, error_code=BACKGROUND_SAFETY_NOT_VERIFIED`; `Result!A1=0.0` after | **CONFIRMED PASS** |
| **T26** stale hash | `REVISION_CONFLICT`/`STALE_TARGET`, source unchanged, no bypass | `code=REVISION_CONFLICT, isRetryable=false`; `current_source_hash` still `3ee58a8e…` (real) ≠ `deadbeef`; caveat "do NOT blind-replay" | **CONFIRMED PASS** |

### T05 root-cause — description/implementation mismatch CONFIRMED at source
- `server.py:662` description: **"address is the top-left anchor"**; Example (`server.py:670-671`) uses `address:"A1"` with a **1×2** payload.
- `ranges_adapter.py:2964-2966` `_validate_payload_shape`: rejects unless `rows==target_rows and cols==target_cols` — i.e. `address` **must span the full payload dimensions**, not a single anchor.
- The description's own Example (`A1` + 1×2) would itself raise `INVALID_RANGE` against the implementation. The prose contradicts the validator. **F1 is real, reproduced, and description-only.**
- **Aggravating:** `tests/unit/test_description_contract.py:104-107` only asserts `"Example:" in description` — it never checks example/address/payload consistency, so the defect passes CI. The "description contract" gives false assurance here.

**Note on the error envelope during T05:** `suggestedNextActions=["fix arguments to match input schema"]` is **generic**, not the precise "address must equal the full target range (e.g. D2:E3)" hint. The results report says the envelope "visibly guided recovery"; in my reproduction the guidance was correct-but-generic. Recovery still happened (agent/model fixed the range), but the envelope did not specifically name the fix. Minor overstatement.

---

## 3. Mandate 3 — Challenge claimed improvement

- **"100% workflow completion" (8/8 distinct tasks):** meaningful **only** as a lower-bound feasibility signal that the remediated surface is usable by a real agent. It is **not** a comparative improvement (no baseline model run) and rests on **n=8 distinct, of which 7 are n=1** (results §6). It also **masks a 0/4 first-mutation validity** (results §4) — the single most usability-relevant number is the worst one, and it is reported separately. Net: honest but fragile as a headline.
- **Token claims are honest.** Results §5 labels 458,854→495,792 chars explicitly as **server `tools/list` payload, not model tokens**, and the chars/4 figure as a **biased upper-bound proxy, "not a measurement."** This is exactly the separation `research_eval_2` §4 demanded. **No pass-off detected.**

---

## 4. Mandate 4 — Token measurement integrity

**CONFIRMED CLEAN.** No artifact equates `tools/list` char count with model tokens. The 4-layer model (research_eval_2 §4) and results §5 keep (a) server payload separate from (b) client-forwarded context and (c/d) real model tokens, and mark (b)/(c)/(d) `UNVERIFIED`/limitation in the Kilo path. The +8% payload is correctly attributed to description expansion (Auditor: descriptions +53%, schemas +2–4%), not to model cost.

---

## 5. Mandate 5 — Hidden human assistance / task-specific prompting

**No evidence of assistance; the failure pattern is itself counter-evidence.**
- Design (research_eval_2 §3.5/§7) and results §1 both assert prompts are verbatim task text, no expected sequence, no hints, no arg fixes.
- Adversarial cross-check: T05 failed first-mutation **3/3** and T14 failed first-mutation **1/1**. Had subagents received the expected args (`D2:E3`, or snapshot-first), these would not have failed. The reproducible first-mutation failures are positive evidence that **no arg-fixing/prompting occurred.**
- I cannot introspect the subagent prompts directly (not in artifacts); this is an inference from behavior, stated as such. **No red flags found.**

---

## 6. Mandate 6 — Safety invariants

| Invariant | Status | Evidence |
|---|---|---|
| Originals unmodified | **HELD** | `RangesCellsProject.xlsm` SHA-256 `E665CFF…` identical before/after; `FullAccessWorkflow.xlsm` `6B905837…`; `excel_open` uses temp working copy (`path`≠`source_path`) |
| No wildcard kill | **HELD** | No recovery/kill invoked; single owned pid never touched |
| No bypass | **HELD** | T26 stale hash rejected; T17 block not bypassed; T16 veto not overridden |
| No `INTERRUPTED_UNKNOWN` retry | **HELD** | No job resubmission; T24 not exercised (no retry attempted) |
| No concurrent Excel | **HELD** | One server, sequential; `excel_list_workbooks` empty between/after; closed with `save_changes=false` |

**No safety invariant was weakened by the remediation or by this validation.**

---

## 7. Mandate 7 — Per-category verdicts

| Category | Verdict | Basis |
|---|---|---|
| **Workbook lifecycle** | **APPROVED WITH CONDITIONS** | `excel_open`/`excel_close` working-copy semantics + additive `SessionStatus.policy` verified live; T01/T02/T20/T28 not independently reproduced (n=1, deterministic). Condition: lifecycle claims remain single-sample. |
| **Read / Write** | **APPROVED WITH CONDITIONS** | Read (T03) reproduced PASS. Write reproduces the **F1 description/implementation defect** (first-call validity 0/4). Condition: fix F1 description+example before any "write is usable first-try" claim. |
| **Tables** | **INCONCLUSIVE** | Not independently reproduced; only deterministic proxy + n=1 remediated. No defect evidence, but no independent verification either. |
| **Pivot** | **APPROVED WITH CONDITIONS** | Enum prune verified at contract level (Auditor) + runtime rejection of `color_scale` (runtime-compat report); T14 `SNAPSHOT_FAILED` is envelope-mitigated model behavior. Not independently reproduced here. Condition: unsupported-source path still untested (no fixture). |
| **VBA / Macro** | **APPROVED** | T16, T17, T26 independently reproduced **exactly** (veto, truthful block, revision-conflict with no bypass, source unchanged). Strongest category. |
| **Permissions / Recovery** | **INCONCLUSIVE** | T18/T19/T23/T24 not reproduced (need restricted server / pid kill / seeded job). Source+unit-test evidence only. `isRetryable=false` non-bypass behavior corroborated via T26. |
| **Testing (declarative)** | **INCONCLUSIVE** | T27 single thin task, not reproduced; no negative (invalid-suite) case. |
| **Cross-server routing** | **INCONCLUSIVE (NOT TESTED)** | Single `xlvba` server only; no second `excel-mcp_*` server. Explicitly out of scope; cannot be approved. |

---

## 8. OVERALL VERDICT

**The remediation improved actual agent usability on the remediated surface — within a narrow, honestly-bounded scope — but the comparative improvement over baseline is UNPROVEN.**

What is established (independently confirmed where I reproduced):
- A real agent completed the attempted workflows with **zero unsafe actions**; originals stayed byte-identical.
- The additive error envelope (`errorCategory`/`isRetryable`/`suggestedNextActions`) is present on the wire and **measurably drove correct recovery** on `INVALID_RANGE` (T05), `SNAPSHOT_FAILED` (T14), `REVISION_CONFLICT` (T26), and `BACKGROUND_SAFETY_NOT_VERIFIED` (T17).
- The VBA/Macro safety surface (hash guard, unclassified block, interactive veto) is **correct and non-bypassable**.

What remains **unproven / overstated**:
- **No paired model-vs-model number exists** (baseline is deterministic). "100% WC" is a lower-bound feasibility signal on n=8 (mostly n=1), not a rate and not a comparison.
- **First-mutation validity is 0/4** — the headline completion number masks the worst usability metric.
- **Run-to-run variance is unmeasured** for 7 of 8 tasks (n=1).
- **Tables, Pivot (unsupported source), Permissions/Recovery, Testing, Cross-server** lack independent reproduction.
- **Annotation forwarding benefit** is unverifiable in this client (correctly not claimed).
- **Real model token consumption** is unmeasured (correctly not claimed).

**Conditions attached to this verdict:** fix F1 (description + example + contract-test gap); attach the scope qualifier next to the "100%" headline; do not state annotation/token/cross-server benefit as measured.

---

## Approved Findings
- **F1 (SPEC-DEVIATION, description vs implementation)** — APPROVED. `range_write_values` description says "top-left anchor" and its Example (`A1`, 1×2) contradicts `_validate_payload_shape` (requires full-range address). Reproduced live. Description-only fix; no code/schema change.
- **Design-honesty of the deterministic-baseline limitation** — APPROVED (DOC/METHOD). Disclosed in results §2/§6/§8 and failure-analysis F7.
- **Token-measurement separation (server payload ≠ model tokens)** — APPROVED. No pass-off.

## Rejected Findings
- **"Remediation caused a measurable model-level improvement over baseline"** — REJECTED (insufficient evidence; baseline not runnable; paired metric INCONCLUSIVE by the authors' own admission).
- **"Error envelope specifically guided the T05 fix"** — REJECTED as stated (overstated). The envelope's `suggestedNextActions` was generic ("fix arguments to match input schema"); recovery occurred but the hint did not name the range fix.
- **"100% workflow completion demonstrates stable usability"** — REJECTED as a stability claim (n=8, 7×n=1, 0/4 first-mutation). Acceptable only as a one-run feasibility observation.

## Merged Findings
- **F1 + contract-test gap** share one root cause: the description-contract test validates example *presence* but not *semantic consistency* (`test_description_contract.py:104-107`). Fix both together.
- **F2 (T14 snapshot) + F4 (T05 trial1 contamination)** are both non-interface, envelope/harness-mitigated; merge as "first-attempt precondition misses are model/harness, not surface defects."

## Rollout Analysis
- No rollout/dependency risk introduced by the remediation itself (Auditor: additive + enum-prune + 1 correctness fix; HEAD docs-only, src diff empty — re-verified). The F1 fix is description-only, backward-compatible, no schema/code change, low risk. Sequencing: land F1 description+example+test together; no migration.

## Execution Validation
- Reproduction targets exist and are current (HEAD `40c8b4d` = `81bcf2c` src; baseline `e2289ca` 90 `add_tool`). All anchors semantic (tool handlers), not line-fragile. Live server matches `81bcf2c` (`xlvba_environment`). **Ready** for the F1 description fix; benchmark conclusions are **not** stale but are **scope-limited**.

## Warnings
- **Architectural/maintainability:** description-contract test gives false assurance (presence-only). A semantic example check would prevent regressions of the F1 class.
- **Rollout/claim risk:** "100%" headline without adjacent scope qualifier invites over-reading.
- **Evidence risk:** 7/8 tasks n=1; tables/pivot/recovery/testing/cross-server unverified independently.
- **Documentation:** results §3 "envelope visibly guided recovery" slightly overstates the generic T05 hint.

## Required Fixes
1. **F1 description:** replace "address is the top-left anchor" with "address must be the full target range matching the payload (e.g. `D2:E3` for a 2×2)"; fix Example to `address:"A1:B1"` for the 1×2 payload. Description-only.
2. **Contract test:** extend `test_description_contract.py` to validate that each `range_write_values` Example's `address` dimensions equal its `payload.rows×cols` (close the presence-only gap).
3. **Claim hygiene:** place the "deterministic baseline / not a model run / n=8 mostly n=1 / first-mutation 0/4" qualifier adjacent to the "100% completion" statement in any external summary.

## Advisory Recommendations
- Run Tables/Pivot/Recovery/Testing reproductions (and a true dual-transport paired baseline) in a future harness to convert INCONCLUSIVE cells to evidence.
- Make `suggestedNextActions` for `INVALID_RANGE` payload-shape name the required full-range form (low cost, higher agent value than the generic hint).
- Defer unsupported-pivot-source and cross-server coverage until a fixture / second server exists; keep them marked untested rather than inferred.