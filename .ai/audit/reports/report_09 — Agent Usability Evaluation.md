# Report 09 — Agent Usability Evaluation

- **Role:** Tech Lead (final owner). Evaluation, runtime compatibility, diagnosis. **No production changes made.**
- **Date:** 2026-11-26
- **Surfaces:** Baseline `e2289ca` (91 tools) → Remediated `81bcf2c` (92 tools; HEAD `40c8b4d` docs-only).
- **Evidence base:** Phase 0 environment record; Auditor change map (PASS 9/9); runtime compatibility; model-in-the-loop benchmark (remediated, real); failure analysis; independent Validator; final Auditor (CONSISTENT).

## 1. Executive verdict

The Agent Usability Remediation **improves the remediated surface's real, autonomous usability within a narrow, honestly-bounded scope**: a real agent completed every attempted workflow with **zero unsafe actions**, correct tool selection, and correct interpretation of every structured error. The structured error envelope (`errorCategory`/`isRetryable`/`suggestedNextActions`) measurably drove correct recovery, and previously-advertised-but-rejected enum values are gone.

**However, the comparative claim "the remediated interface makes agents perform better than the original interface" is NOT proven by this evaluation.** A true paired baseline model run was infeasible (the Kilo client is bound to one MCP server; the global config is unreadable; the single-Excel/single-STA rule forbids a concurrent baseline server). The baseline is characterized **deterministically** (surface deltas + prior proxy), which is favorable but is not a model-in-the-loop measurement. The verdict on comparative improvement is therefore **INCONCLUSIVE**, by design and disclosure — not by omission.

One reproducible, interface-caused defect was found and is worth fixing before any further rollout: the `range_write_values` address description contradicts the implementation and caused first-mutation failure in **4/4** mutation trials (self-corrected every time).

## 2. Runtime compatibility (Phase 3)

PASS: current descriptions; input schemas; pruned-enum rejection (`color_scale` → `INVALID_RANGE`, valid values still advertised); `xlvba_environment` (access mode, policy, honest `build:null`); additive `SessionStatus.policy`; glossary resource reachable and readable (14 terms); structured error envelope visible to the model. PARTIAL: live error-category coverage (2 of 8 exercised; rest source/test-verified). UNSUPPORTED: in-session observation of Kilo annotation rendering (server emits 92/92, client forwarding unverifiable). No FAIL/BLOCKED. Full detail: `agent_usability_runtime_compatibility.md`.

## 3. Baseline comparison (deterministic)

| Dimension | Baseline `e2289ca` | Remediated `81bcf2c` | Delta |
|---|---|---|---|
| Tools / resources | 91 / 4 | 92 / 5 | +xlvba_environment, +glossary |
| Annotations | 0 | 92/92 truthful | +92 |
| Error envelope keys | 5 | 8 | +errorCategory,isRetryable,suggestedNextActions |
| Env/policy discovery | partial | full | new capability |
| Advertised-but-rejected enums | present | pruned (request) | fixed |
| `tools/list` payload | 458,854 chars | 495,792 chars | **+8.0%** |

This is a **surface** comparison, not a model-use comparison (see §1, §7).

## 4. Measured usability (remediated, real model-in-the-loop)

Fresh subagent conversations, identical prompts, no hints, no arg fixing, single shared Excel (serialized), originals never modified.

- **Workflow completion:** 8/8 attempted tasks (T28 = correct gap handling). **Fragile**: n=8, seven at n=1; not a stable rate.
- **Correct tool selection:** 8/8.
- **First-call validity (overall):** 8/8. **First-mutation validity (first try): 0/4** (T05×3, T14×1) — all recovered via the error envelope. This is the headline weakness and is interface-caused for T05 (F1).
- **Unsafe action attempts:** 0.
- **Error interpretation:** all structured errors interpreted correctly; `isRetryable`/`suggestedNextActions` guided recovery (the T05 hint was generic, so "the envelope visibly guided recovery" is mildly overstated — the agent self-corrected from the schema error text).
- **Calls per task:** 1–10 (median 6).

Per-category (Validator): VBA/Macro **APPROVED**; Workbook lifecycle, Read/Write, Pivot **APPROVED WITH CONDITIONS**; Tables, Permissions/Recovery, Testing, Cross-server **INCONCLUSIVE** (not exercised; cross-server untested at all).

## 5. Token-efficiency assessment (honest)

- The remediation increased upfront **server** `tools/list` payload by **+8.0%** (outputSchema ≈73%). This is **not** model tokens; real model input tokens are not observable in-session and are reported as a limitation, not estimated as fact.
- No safe, field-preserving outputSchema reduction exists (U4 no-op); further reduction is a **product decision** requiring evidence that correctness is unaffected. A larger surface is acceptable here because it buys discoverability/error-recovery with zero measured safety cost — but the token trade-off is real and unresolved.

## 6. Safety evaluation

All invariants held across every trial and reproduction: originals byte-identical before/after; no wildcard process termination; no policy bypass; no `INTERRUPTED_UNKNOWN` auto-retry; no concurrent Excel; no production data corruption. The benchmark did not weaken safety to improve completion.

## 7. Remaining limitations

1. **Comparative improvement unproven** — no baseline model run (environmental).
2. **Reduced sample** — 7 of 8 tasks at n=1; variance unmeasured.
3. **21 of 28 benchmark tasks not run** (T01,02,04,06–13,15,18–25,27); cross-server routing untested.
4. **Annotation forwarding** unverifiable in this Kilo client.
5. **Model tokens** unobservable; only server payload measured.
6. **Kilo version / global config** unreadable (permission + PATH).

## 8. Validator verdict (independent)

`report_09a`: per-category as in §4; **OVERALL: remediation improved actual agent usability within a narrow, honestly-bounded scope; the comparative improvement over baseline is UNPROVEN.** Required fixes: F1 description+example; close the contract-test semantic gap (it only checks `"Example:"` presence, so the defect passes CI); attach the scope qualifier to the "100%" headline (done here). No safety invariant weakened; token claims honest; no hidden assistance.

## 9. Auditor verdict (final consistency)

`report_09b`: **CONSISTENT (7/7)** — traces match metrics; task set is the approved subset; no production changes (`git diff 40c8b4d -- src/ tests/` empty; src==`81bcf2c`); all failures documented; token figures labeled server-payload; safety held.

## 10. Final Tech Lead recommendation

1. **Accept** the remediation as a real usability improvement on the remediated surface, with **no safety regression**, scoped exactly as §4/§7 state. Do **not** publish a causal "better than baseline" claim.
2. **Execute `plan_04`** — one bounded, high-value, low-risk follow-up: fix the `range_write_values` description + example (F1) and add a semantic contract test so description and validator cannot drift again. Description/test-only; no schema/behavior change.
3. **Keep deferred** (no approval): `save_as`/export, `xlvba_job_list`, major outputSchema redesign, new tool-discovery infrastructure.
4. **For a true paired verdict**, run a future evaluation with a swappable/dual-transport harness (or a scripted MCP stdio client that launches either revision) holding model/config constant. Until then, the comparative metric stays INCONCLUSIVE.

**Bottom line:** the remediation is a net positive for autonomous Excel work and is safe; its superiority over the original interface is plausible and deterministic-surface-supported but **not yet empirically proven** — and one concrete description defect should be fixed first.