# Stage 1 Report — Integration Validation & Real Agent Benchmark

## 1. Repository & implementation state
- Candidate `ef3e6cb` (U6-A..E + U7 + U6-D) + uncommitted Save As `.bak` fix (safe-to-ship per Technical Auditor).
- Baseline `40c8b4d` (pre-plan_05), identifiable.
- Branch `master`; no git remotes (rollout = local).

## 2. Live environment
- Excel 16.0, owned PID 27704, access_mode full, single-STA worker, single owned instance (verified).
- Model-in-the-loop: a real LLM subagent drove the live MCP server (candidate).

## 3. Save As live integration matrix
9 scenarios PASS (new target, reopen, existing-no-overwrite, existing-overwrite + exact `.bak`, source-protect, source-overwrite + original `.bak`, external-mod drift, subsequent-edit target-unchanged, repeated). Scenario 6 (pre-existing `.bak`) exposed a defect (reported `unknown`); fixed to deterministic `PROTECTED_TARGET/backup_exists` (pre-mutation) + regression tests. Failure/interruption covered by 15 unit cases. Independently re-hashed by the Validator (all match). **Verdict: PASS.**

## 4. Candidate real-agent run
Autonomous LLM subagent completed environment discovery, read, and Save As with correct tool selection and typed args, no intervention. Supplementary (not a comparison).

## 5. Comparative baseline benchmark — BLOCKED
The required baseline-vs-candidate model-in-the-loop benchmark (≥3 trials/workflow, same model/Kilo/MCP) is **not executable** in this environment: only the candidate server is connected; the Kilo global config is read-restricted (`external_directory: deny`) and cannot be re-pointed; no baseline server is reachable by a subagent. No proxy was substituted (per rules).

## 6. Independent review
- Integration Validator (report_11a): Save-As integration **PASS**; fix correct; contract doc corrected.
- Technical Auditor (report_11b): **SAFE TO SHIP**; architecture intact; no new ErrorCode; SCHEMA_VERSION=1.

## 7. Changes during Stage 1
- `session_service.py` + `test_save_as.py`: `.bak` conflict reclassified to deterministic `PROTECTED_TARGET/backup_exists`.
- `save_as_export_contract.md`: §5.3 updated.

## 8. Stage 1 decision: **FAIL (BLOCKED on comparative benchmark)**
Save As live integration = PASS; candidate real-agent = PASS; fix safe. But the mandatory comparative real model-in-the-loop benchmark cannot be completed in this environment. Per the Stage 1 gate, **Stage 2 is NOT started**. Minimum remediation: run the comparative benchmark in an environment with both baseline and candidate servers routable by the agent (or an approved equivalent), then re-run Stage 1D.