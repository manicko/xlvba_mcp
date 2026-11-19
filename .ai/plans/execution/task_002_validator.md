# task_002 — Validator output (Tech Lead session, 2026-11-19)

_Source: `task` subagent_type=validator. Review of combined Researcher/Auditor/Planner package (as summarized to validator) against plan/spec._

## Verified block
`task_002` PoC design gate (dedicated_gate). Must validate before implementing `tests/integration_excel/test_excel_instance_poc.py`.

## Checks performed
- Inspected plan `task_002`, R01, G1; spec excerpts in `.ai/tmp/spec_text.txt` (§5.3, §7.1 context, §9.4, §11/11.1, §13.1).
- Searched repo for persisted R/A/P artifacts and for `tests/integration_excel/test_excel_instance_poc.py` (none found at time of review).
- Evaluated summarized Researcher/Auditor/Planner inputs against ACs: environment capture, `DispatchEx`, `CoInitializeEx(STA)`, shutdown/finally/orphans, STA proxy confinement, message pumping + COM/UI events + RPC_E_CALL_REJECTED/modal behavior, Go/No-Go rubric.

## Findings
- ACs/goals are high-level, not execution-grade (no measurable evidence methods, no negative/edge scenarios, no decision rules tied to Go/No-Go).
- R01/spec require proving message pumping/dispatch behavior and behavior under `RPC_E_CALL_REJECTED`/modal dialogs; summarized package lacks concrete pumping strategy, instrumentation, scenario set, HRESULT/classification, and stated “not exercised” consequences.
- Isolation proof not rigorous: no required baseline process/PID/window snapshots, owned PID/HWND evidence, ROT/read-only rules, or precondition (e.g. “no Excel running” verified) spelled out as testable steps.
- Shutdown/orphans lack detection procedure (track PID, timed poll, timeout, failure → No-Go input) and explicit ownership/save policy on test workbook.
- Proxy thread-affinity not testably specified; risk of diverging from production single-STA `com_worker.py` model.
- No explicit Go/No-Go rubric (environment completeness, required evidences, disqualifying conditions). “Record HRESULT” not tied to rubric.
- VBE language capture method not specified in AC/checklist language.
- Artifacts for R/A/P not persisted in repo at review time (visibility/provenance gap) — cannot independently verify closure of D1–D9/(a)–(h).

## Final result
**FAIL** — do not proceed to implement `tests/integration_excel/test_excel_instance_poc.py` as the gate PoC now. Required before proceeding (blocking): persist full R/A/P artifacts (with D1–D9→(a)–(h) trace), make planner skeleton/checklist/evidence template concrete and persisted, require explicit Go/No-Go rubric and all evidence fields (env matrix incl. VBE language method, isolation baselines+owned identity, STA thread-id/logging for COM calls, `finally`+timed orphan poll, pumping strategy+observations+RPC/modal handling or justified not-exercised with impact), and align PoC to production single-STA model. Reassess after artifacts fixed.
