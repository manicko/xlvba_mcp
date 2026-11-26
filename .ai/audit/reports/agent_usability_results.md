# Agent Usability Results — Before/After (plan_03)

- **Revision:** baseline `e2289ca` → final working tree (U1+U3+U2 accepted; U4 documented no-op; U5 docs).
- **Date:** 2026-11-26
- **Method:** deterministic capture of the live `tools/list` (no Excel) + a static benchmark-feasibility proxy over `.ai/tmp/benchmark_tasks.yaml`. The full model-in-the-loop 28-task run is specified but was NOT executed here (no agent runtime in this environment) — see §5.

## 1. Before/after tool-surface metrics

| Metric | Before (`e2289ca`) | After (final) | Delta |
|---|---|---|---|
| Tools exposed | 91 | 92 | +1 (`xlvba_environment`) |
| `tools/list` JSON chars | 451,119 | 481,641 | +30,522 (+6.8%) |
| Description chars | 24,272 | 37,125 | +12,853 (standard templates + examples) |
| Input-schema chars | 80,656 | 82,373 | +1,717 |
| Output-schema chars | 338,547 | 350,758 | +12,211 (U2 env DTOs; U4 no-op) |
| Tools with annotations | 0 | 92 | +92 (truthful readOnly/destructive/idempotent) |
| Static resources | 0 (4 templates) | 1 (glossary) + 4 templates | +1 |
| Error envelope keys | 5 | 8 | +`errorCategory`,`isRetryable`,`suggestedNextActions` |
| Discoverable access mode / policy / build | partial (mode only) | `xlvba_environment` + `excel_get_status.policy` | full |
| Advertised-but-rejected enum values | present (CF/pivot) | pruned (request side) | fixed |
| Glossary terms | 0 | 14 (`xlvba://glossary`) | +14 |

Artifacts: after-state dump `.ai/tmp/tools_list_after.json`; capture/measure scripts `.ai/tmp/capture_baseline.py`, `.ai/tmp/measure_after.py`. Baseline aggregates are authoritative in `agent_usability_baseline.md` (the baseline JSON dump was regenerated during capture; aggregates unchanged).

## 2. Token-efficiency outcome (honest)

**The token-reduction objective was NOT achieved.** Upfront context rose ~6.8%. Root cause (verified by U4): Pydantic already fully dedupes within each per-tool schema (`$defs`/`$ref`), and MCP `tools/list` emits an independent self-contained schema per tool with **no cross-tool `$ref`**, so `SheetInfo` (×7) and the test DSL (×2) cannot be factored without removing fields. `structured_output=False`, a custom search tool, and `defer_loading` are rejected (D-U1: the first removes `structuredContent`, the second adds a 92nd always-loaded tool, the third is absent from mcp 2.3.0). The increase is the deliberate cost of U2 (environment/policy DTO) and richer descriptions. **Further token reduction requires a product decision** (remove specific DTO fields, or client-side schema filtering) — deferred (D-U9/U4 follow-up).

## 3. Correctness & safety improvements (validator-accepted)

- **U1 (ACCEPT):** 25 workflow-critical descriptions follow the 8-section standard; uniform guard sentence; neighbor routing; 4 worked examples; truthful claims (`vba_compile` diagnostic-only, `macro_classify_background` never grants run, `job_cancel` intent-only); glossary resource (14 terms); docs EN+RU corrected (`test_run_suite` `background` mis-attribution removed; tool count 91).
- **U3 (ACCEPT):** request enums pruned (CF plain set; pivot orientation no `data`; filter `manual` only) while read/classify enums intact; new `PivotFilterState` preserves read fidelity (`supported=false` for non-manual); truthful annotations on 91/91 tools (`macro_run`/`vba_compile` not read-only); pruned values still rejected (now pydantic). 6 integration tests aligned to the new request layer.
- **U2 (ACCEPT):** read-only, non-COM `xlvba_environment` (`build:null`, `target_build` labelled, capabilities, policy with `classification_is_authorization=false`, `has_snapshot` from the in-memory gate); additive `SessionStatus.policy`; additive error envelope with per-code `isRetryable`/`suggestedNextActions` (D-U10: `STALE_TARGET`/`REVISION_CONFLICT` false + partial-apply caveat; `interrupted_unknown` never retried). No SDK change; no error `structuredContent`.
- **U4 (documented no-op):** no safe field-preserving schema reduction; rationale in `docs/09`.
- **U5 (docs):** `save_as`/export and `xlvba_job_list` product contracts outlined; NOT implemented.

## 4. Deterministic benchmark-feasibility proxy

`& ".venv\Scripts\python.exe" .ai\tmp\benchmark_proxy.py` over the 29-entry task set:

- **First tool exists:** 28/28 real tasks (the 29th, `c1`, is a no-tool control — expected).
- **Expected argument keys valid against final input schemas:** 29/29 (no advertised/rejected mismatch for the benchmark's expected args).

This is a lower bound (tool/arg feasibility), not a model-in-the-loop success rate.

## 5. Full Agent Usability Benchmark — status

- **Specified:** `.ai/tmp/benchmark_tasks.yaml` (28 tasks + control) and `research_2b_usability_benchmark.md` (metrics, harness reusing `test_mcp_full_access_workflow.py` primitives, constant controls, JSONL recording).
- **Executed here:** deterministic proxy (§4) + token/schema capture (§1).
- **Not executed here:** the 28-task model-in-the-loop run (requires an agent runtime driving the live server). It is ready to run: hold constant model/agent/Kilo/revision/fixtures; run baseline (`e2289ca`) and final; record CTS/FCV/WC/UC/RC/UAA/TC/CT per task.
- **Targets:** set only after the baseline model-in-the-loop run (per plan §16); none asserted here.

## 6. Remaining friction

1. Upfront token cost increased; outputSchema is ~73% of payload (product decision needed to reduce).
2. Full model-in-the-loop benchmark not yet run.
3. `excel.application_version`/`vbom_available` may be `null` at server level (non-COM constraint); 4-part Excel build unavailable (`build:null`).
4. Kilo forwarding of `outputSchema`/annotations/resources and resource meta-tools remain unverified from inside the session.

## 7. Acceptance vs plan §15.3

- Unit: **1127 passed**; ruff clean on changed files; integration **collects** cleanly (live-Excel run not performed here).
- Per-package validators: U1/U2/U3 ACCEPT; U4 documented no-op; U5 docs.
- Overall: correctness/discoverability/safety improved with no regression to the accepted Excel runtime architecture; token-reduction target not met (reported, deferred to product).