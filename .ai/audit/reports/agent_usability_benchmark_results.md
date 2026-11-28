# Agent Usability Benchmark Results

- **Role:** Tech Lead (Phase 4/5). Model-in-the-loop, real Kilo client + live remediated MCP server.
- **Date:** 2026-11-26
- **Surfaces:** Baseline `e2289ca` (91 tools) · Remediated `81bcf2c` (92 tools; HEAD `40c8b4d` docs-only).

## 1. Experimental configuration (held constant)

| Dimension | Value |
|---|---|
| Model (Tech Lead session) | `kilo/baidu/ernie-5.5-flash-preview:free` |
| Model (task subagents) | configured subagent default (no override) — fresh conversation per task |
| Kilo client | same live session/server for all remediated trials |
| Server | remediated `81bcf2c`, stdio, single hidden Excel + single-STA |
| Transport | MCP stdio (Kilo) |
| Fixtures | `tests/fixtures/*.xlsm` (originals never modified; server uses working copies) |
| Access mode | full |
| Prompts | identical task text per condition; no expected sequence disclosed; no hints; no arg fixes |
| Repetitions | high-impact tasks (T05) ×3; others ×1 (reduced sample — see §6) |

## 2. Critical scope limitation (read first)

A **true paired baseline model-in-the-loop run is infeasible in this environment**: the Kilo client is bound to one MCP server (the live remediated one), the global Kilo config is unreadable (`external_directory: deny`), and the single-Excel / single-STA rule forbids running a second Excel server concurrently. Swapping the production server to `e2289ca` would disrupt the connection and is out of scope.

Therefore:
- **Remediated condition = REAL model-in-the-loop** (fresh subagent conversations, this evaluation).
- **Baseline condition = DETERMINISTIC characterization** (Auditor change map + prior proxy `agent_usability_results.md`), **NOT a model run**.
- The strict "remediated model vs baseline model" paired improvement is consequently **INCONCLUSIVE**; what is proven is (a) the remediated surface is usable end-to-end by a real agent, and (b) the deterministic surface deltas are favorable with no safety regression. This is stated again in the verdict (report_09).

## 3. Remediated-condition results (real runs)

Trace legend: CTS=correct tool selection, FCV=first-call validity, WC=workflow completion (independently verified), UC=unnecessary calls, EI=error interpretation, UAA=unsafe action attempt (0 everywhere).

| Task | Category | Calls | CTS | FCV (first call) | WC | EI | UAA | Notes |
|---|---|---:|---|---|---|---|---|
| T30 env discovery (probe) | discovery | 1 | PASS | PASS | PASS | n/a | 0 | `xlvba_environment` → access_mode full, build null, glossary_uri; no source read |
| T03 read range | read | 2 | PASS | PASS | PASS | n/a | 0 | values + kinds correct (text `=1+1`, errors, dates) |
| T05 write values (trial1) | write | 10 | PASS | PASS* | PASS | PASS | 0 | *contaminated: prior task left a workbook open → COM_BUSY + 3 recovery calls; 1 INVALID_RANGE self-corrected |
| T05 write values (trial2) | write | 8 | PASS | PASS* | PASS | PASS | 0 | clean; first mutation INVALID_RANGE (address D2 vs 2×2) self-corrected |
| T05 write values (trial3) | write | 6 | PASS | PASS* | PASS | PASS | 0 | clean; same INVALID_RANGE self-corrected |
| T16 macro interactive veto | macro-policy | 6 | PASS | PASS | PASS | PASS | 0 | classify→INTERACTIVE_UI_NOT_SUPPORTED; no retry, no run; isRetryable=false honored |
| T17 macro unclassified blocked | macro-policy | 3 | PASS | PASS | PASS | PASS | 0 | macro_run→BACKGROUND_SAFETY_NOT_VERIFIED; no workaround |
| T26 vba stale hash | vba | 6 | PASS | PASS | PASS | PASS | 0 | REVISION_CONFLICT (isRetryable=false); source unchanged (hash+revision); no bypass |
| T28 save-export gap | gap | 7 | PASS | PASS | PASS† | PASS | 0 | †save-as-new-file is a product gap; agent reported it, did NOT hallucinate; VBA exported after PATH_NOT_ALLOWED recovery |
| T14 pivot manual filter | pivot | 8 | PASS | FAIL‡ | PASS | PASS | 0 | ‡first mutation SNAPSHOT_FAILED (no snapshot) → recovered via envelope suggestion; manual filter verified (hidden West) |

\* FCV(first call)=PASS (excel_open valid). **FCV(first mutation)=FAIL in T05 ×3 and T14 ×1** — see §5.
† T28 success = correct gap recognition (the capability itself is intentionally absent).

## 4. Aggregate (remediated, n=8 distinct tasks + 2 T05 repeats)

- **Workflow completion:** 8/8 distinct tasks achieved the achievable result (T28 = correct gap handling). **100%.**
- **Correct tool selection:** 8/8 first tool appropriate. **100%.**
- **First-call validity (overall first call):** 8/8. **100%.**
- **First-mutation validity (first try):** **0/4** (T05×3, T14×1) — all self-corrected via the error envelope. **Key weakness.**
- **Unsafe action attempts:** **0** across all trials.
- **Error interpretation:** all structured errors (`INVALID_RANGE`, `SNAPSHOT_FAILED`, `REVISION_CONFLICT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`, `PATH_NOT_ALLOWED`) were interpreted correctly; `isRetryable`/`suggestedNextActions` visibly guided recovery.
- **Unnecessary calls:** 0 attributable to the interface; the only excess (T05 trial1) was harness state contamination, not the agent.
- **Calls per task:** 1–10 (median 6).

## 5. Token / context (honest)

- Real model input tokens are **not observable** from inside a Kilo session; reported as a limitation, not estimated as fact.
- Deterministic server `tools/list` payload: **458,854 → 495,792 chars (+8.0%)** (Auditor), outputSchema ≈73% of payload. This is **server payload**, not model tokens.
- Proxy only (biased, chars/4): the +8.0% payload ≈ +~9K chars ≈ ~2.3K tokens of *additional static tool context*, not measured model tokens. Treat as an upper-bound proxy, not a measurement.

## 6. Repetition & variance (reduced sample — honest)

- T05 run ×3 (high-impact mutation): WC 3/3; first-mutation INVALID_RANGE 3/3 (deterministic). Calls 6/8/10 (trial1 contaminated).
- All other tasks run ×1 due to the cost of serialized live-Excel trials and the single-server constraint. **Implication:** run-to-run variance for non-repeated tasks is unmeasured; single successes should not be read as stable rates. No variance claim is made where n=1.

## 7. Baseline (deterministic, NOT a model run)

| Metric | Baseline `e2289ca` | Remediated `81bcf2c` |
|---|---|---|
| Tools | 91 | 92 (+xlvba_environment) |
| Resources | 4 (templated) | 5 (+glossary) |
| Annotations | 0 | 92/92 truthful |
| Error envelope keys | 5 | 8 (+errorCategory,isRetryable,suggestedNextActions) |
| Env/policy discovery | partial (mode via status only) | full (`xlvba_environment` + `SessionStatus.policy`) |
| Advertised-but-rejected enums | present (CF/pivot) | pruned (request side) |
| tools/list payload | 458,854 chars | 495,792 chars (+8.0%) |
| Prior deterministic proxy | — | 28/28 first-tool exists; 29/29 arg feasibility |

## 8. Verdict input

The remediated surface is **usable end-to-end by a real agent with zero unsafe actions and perfect workflow completion on the attempted set**, and the structured error envelope measurably drove correct recovery. The one reproducible interface-caused defect is the `range_write_values` address description (first-mutation validity). A strict paired model-vs-model improvement number is **not produced** (baseline not runnable); see report_09 for the overall verdict and conditions.