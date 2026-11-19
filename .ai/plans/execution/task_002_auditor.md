# task_002 — Auditor output (Tech Lead session, 2026-11-19)

_Source: `task` subagent_type=auditor. User-reported conclusions summarized below; persisted for traceability._

## Verdict
**FAIL** (PoC plan as written insufficient as critical Go/No-Go gate).

## Identified defects / risks (from auditor)
- **D1 (HIGH):** STA message pumping named but not operationally specified/proven (no pump strategy tied to queued work, no explicit handling of blocking/wait patterns, modal/COM event behavior evidence missing). R01/spec expectations for message pumping and `RPC_E_CALL_REJECTED`/modal dialogs not concretely proven.
- **D2 (HIGH):** Isolation proof not rigorous (no required PID/HWND/process-set + pre/post snapshot of existing Excel sessions beyond “separate instance identified”; risk of misidentifying attachment).
- **D3 (HIGH):** Shutdown/order/ownership under-specified as verifiable (who closes what, `SaveChanges` policy on test workbook, ref release order on STA thread, bounded wait for process exit, diagnostics; only owned workbooks/instance quit).
- **D4 (HIGH):** Proxy/thread-affinity “remains in creating STA thread” asserted, not tested (no cross-thread misuse/marshaling probe).
- **D5 (MEDIUM-HIGH):** R01 requires evidence for `RPC_E_CALL_REJECTED` and modal dialogs; task_002 does not require reproducible scenario or defined “not reproducible”/Go-No-Go impact.
- **D6 (MEDIUM):** PoC fixture/open params/Application-state controls unspecified (dialogs, links, macro/trust prompts risk non-reproducibility).
- **D7 (MEDIUM):** “Serialized access” mentioned in description but not evidenced in goals/acceptance (harness is temporary; scope ambiguity).
- **D8 (MEDIUM):** Go/No-Go/HRESULT/environment evidence underspecified (no artifact schema/location, pass/fail rubric, minimum diagnostics).
- **D9 (LOW):** Terminology precision (“instance” vs process) should be explicit.

## Architectural / COM risks (summary)
STA message loop correctness vs deadlocks; `DispatchEx` vs attachment boundaries; RCW/ref leaks/orphans; modal/busy envelope and “no duplicate mutating COM ops”; Python bitness/pywin32/Office bitness; MCP stdout discipline; no untrusted macro execution in PoC.

## Spec mismatches / coverage gaps
No direct normative contradictions between task_002 text and cited spec sections noted by auditor, but evidence/acceptance deficiencies relative to R01, §5.3/§13.1 expectations for pumping, modal/busy, rigorous isolation, and lifecycle evidence.

## Recommended corrections (auditor)
- Define bounded PoC harness contract aligned to production single-STA intent; explicit fixture and open params + Application-state handling; capture full env (incl. VBE language method/result policy).
- Make isolation verifiable: pre/post Excel process/PID (and window/context) snapshots; owned PID/HWND proof; do not touch user session (read-only observation rules stated).
- Prove lifecycle/ownership/no-orphan on STA thread in `finally`: close only owned workbooks with explicit `SaveChanges`, release refs, `Quit`, release app refs, uninit per thread model, bounded wait + diagnostics.
- Add probe/assertion for proxy confinement (no cross-thread proxy use in practice/design evidence).
- Specify message pumping strategy and required evidence; for `RPC_E_CALL_REJECTED`/modals define bounded probe or conditional No-Go with named follow-up gate and forbid mutating retries in procedure.
- Define Go/No-Go rubric and persisted evidence artifact (schema/location) with a–h style criteria (per auditor message).

## Final result
FAIL — do not treat task_002 plan as sufficient gate without applying corrections above.
