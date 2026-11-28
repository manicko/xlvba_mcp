# Agent Usability Failure Analysis

- **Role:** Tech Lead (Phase 5). Companion to `agent_usability_benchmark_results.md`.
- **Scope:** every failed or partially-completed task/trial in the remediated model-in-the-loop run, plus interface/client limitations observed. No production changes.

## Root-cause taxonomy used
Model reasoning · Description ambiguity · Missing prerequisite info · Schema mismatch · Kilo/client forwarding · Tool implementation · Environmental limitation · Missing product capability.

## F1 — `range_write_values` first-mutation `INVALID_RANGE` (T05 ×3, also T14 pattern)
- **Observed:** 3/3 T05 trials (and analogous first-mutation misses) pass `address="D2"` with a 2×2 payload; server rejects `INVALID_RANGE` ("payload 2x2 vs target 1x1"); agent self-corrects to `D2:E3`.
- **Root cause:** **Description ambiguity / description-implementation mismatch.** `server.py:662` (`81bcf2c`) states "address is the top-left anchor" and the description's own Example uses `address:"A1"` with a 1×2 payload — both imply a single anchor cell, but the implementation requires `address` to span the full payload dimensions. The description contradicts the validator.
- **Class:** Interface-related (description-only). Not a model defect; not a schema defect (schema is consistent, prose is not).
- **Evidence:** `agent_usability_benchmark_results.md` §3 (T05 trials); `server.py:653-676` `_DESC_RANGE_WRITE_VALUES`; `ranges_service.py:275-291` validates payload shape against the address-derived target.
- **Recommended corrective action:** Description-only: replace "top-left anchor" with "address must be the full target range matching the payload (e.g., `D2:E3` for a 2×2)"; fix the Example to `address:"A1:B1"` for the 1×2 payload. No code/schema change.
- **Priority:** High (affects first-call validity on every multi-cell write). **Complexity:** Low (description + example). **Backward-compat risk:** none.

## F2 — `pivot_field_apply_filter` first-mutation `SNAPSHOT_FAILED` (T14)
- **Observed:** agent applied the filter before `snapshot_create` → `SNAPSHOT_FAILED`; recovered after creating a snapshot (suggested by the envelope).
- **Root cause:** **Model reasoning** (precondition not followed on first attempt). The description *does* state the snapshot precondition; the agent skipped it once. The error envelope (`errorCategory=missing_prerequisite`, `isRetryable=false`, `suggestedNextActions=[snapshot_create...]`) produced correct recovery.
- **Class:** Model-specific behavior; **mitigated by the remediation's error envelope**. Not an interface defect.
- **Recommended corrective action:** None required. (Optional, low value: no change — the envelope already guides recovery; over-prompting would add tokens.)
- **Priority:** Low. **Complexity:** n/a.

## F3 — Save-as-new-file capability absent (T28)
- **Observed:** no `save_as`/`SaveCopyAs` tool exists; agent correctly reported the gap and did not hallucinate; VBA export succeeded after a `PATH_NOT_ALLOWED` recovery.
- **Root cause:** **Missing product capability** (intentionally deferred, D-U9/U5).
- **Class:** Product decision, not a defect.
- **Recommended corrective action:** Keep deferred unless separately approved; the product contract is already outlined in `plan_03a`. Revisit only with a product decision (original protection, output path policy, overwrite, result retrieval).
- **Priority:** Product. **Complexity:** Medium (new tool + path policy).

## F4 — Harness state contamination (T05 trial1 `COM_BUSY`)
- **Observed:** a prior task left a read-only workbook open on the shared server; T05 trial1 hit `COM_BUSY` and spent 3 calls recovering.
- **Root cause:** **Environmental limitation** (evaluation harness did not reset server session between tasks). Not an interface defect; the agent recovered correctly.
- **Class:** Environmental (evaluation-side).
- **Recommended corrective action:** Harness must close all workbooks (or restart the worker) before each task; this evaluation adopted close-at-end per task thereafter. No product change.
- **Priority:** Evaluation integrity. **Complexity:** Low (harness).

## F5 — Annotation forwarding not observable (Phase 3)
- **Observed:** server emits 92/92 truthful annotations (Auditor), but the Kilo client's rendering of `readOnlyHint/destructiveHint/idempotentHint` to the model is not observable in-session.
- **Root cause:** **Kilo/client limitation** (unverifiable in this client).
- **Class:** Client limitation.
- **Recommended corrective action:** None to the MCP server; verify annotation rendering only with a client that exposes them. Do not claim annotation benefit without such evidence.
- **Priority:** Low. **Complexity:** n/a.

## F6 — Upfront token/context increase (+8.0% payload)
- **Observed:** `tools/list` 458,854 → 495,792 chars; outputSchema ≈73%.
- **Root cause:** **Design trade-off** (U2 environment DTO + richer descriptions), not a defect. U4 found no safe field-preserving reduction.
- **Class:** Trade-off (not a failure).
- **Recommended corrective action:** Only a product decision (remove specific DTO fields / client-side filtering) with evidence that correctness is unaffected — deferred. Do not slim blindly.
- **Priority:** Product. **Complexity:** Medium, risk-bearing.

## F7 — Baseline model-in-the-loop run infeasible
- **Observed:** cannot run the baseline server without swapping the production Kilo MCP connection; single-Excel rule forbids a concurrent second server.
- **Root cause:** **Environmental limitation.**
- **Class:** Environmental (evaluation design).
- **Recommended corrective action:** Future evaluation with a swappable/dual-transport harness (or a scripted MCP stdio client that can launch either revision) to obtain a true paired model run. Until then the paired metric is INCONCLUSIVE.
- **Priority:** Evaluation method. **Complexity:** Medium (harness).

## Summary
- **Interface-caused failures:** 1 (F1, description-only, high value, low cost).
- **Model-specific, envelope-mitigated:** 1 (F2).
- **Missing product capability (deferred):** 1 (F3).
- **Environmental / client / trade-off:** F4, F5, F6, F7 (no production change).
- **Safety:** no unsafe action attempts in any trial; no original modified; no `INTERRUPTED_UNKNOWN` retry; no concurrent Excel.