# task_002 — Researcher output (Tech Lead session, 2026-11-19)

_Source: `task` subagent_type=researcher (invoked by Tech Lead). This file persists the report for traceability (validator noted prior lack of persisted artifacts)._

## 1) Problem summary
(as provided by researcher in this session)

`task_002` (critical, depends on `task_001`, Go/No-Go gate for `task_010`/`task_012`/`task_014` and critical path) must prove on target Windows/Excel configuration:
- Isolation: separate Excel instance via `win32com.client.DispatchEx("Excel.Application")`, no impact to existing user Excel session (A02).
- COM STA: `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)`, all COM proxies created/read/released on same STA worker, proxies never cross to MCP.
- Lifecycle: open/close test workbook, clean shutdown (close owned workbooks, Quit, release refs in finally), no orphaned owned process.
- Message pumping/COM-UI behavior: verify STA message pumping/dispatch behavior and behavior under modal UI/COM-busy conditions; document limits; capture failures/HRESULTs.
- Evidence/Go-No-Go: record Windows build, Office build, bitness, VBE UI language, observed COM behavior, Go/No-Go for `task_010`.

Repo state (bootstrap only): minimal MCP STDIO skeleton; no `tests/integration_excel/test_excel_instance_poc.py` production COM worker yet; `tests/test_smoke.py` only.

## 2) Available implementation approaches (compatible)
- **A. Minimal `pythoncom` STA thread + explicit pump loop (recommended baseline).** Dedicated thread `CoInitializeEx(STA)`, `DispatchEx`, queued ops, service message queue (`PumpWaitingMessages` / controlled loop patterns as appropriate). `finally` cleanup. Directly matches plan primitives; proves isolation/lifecycle/pump behavior documentation; later can add retry policy evidence.
- **B. STA worker with COM call-rejection handling (`IMessageFilter`/retry-aware wrapper).** Same as A plus retry/backoff evidence for `RPC_E_CALL_REJECTED`/busy; distinguishes busy/modal vs fatal. Better observability for Office; must never blindly replay mutations.
- **C. Message-pump via UI framework dispatcher.** Risk of architecture drift; not recommended for task_002 PoC driver.

Not acceptable: `Dispatch`/`GetActiveObject` for creation as primary path; proxies across threads; MTA assumptions; claims without observed evidence.

## 3) Recommended approach
Adopt **A as core, instrumented with bounded B-style evidence collection**:
1. Temporary harness `tests/integration_excel/test_excel_instance_poc.py` (helpers local to test area if needed), disposable workbook copy in allowed paths; never touch `projects/`.
2. Capture env: Windows/arch, Python bitness, Excel `Version/Build/OperatingSystem`, VBE UI language best-effort or mark unknown with method.
3. Isolation via `DispatchEx`; capture PID/HWND/process evidence pre/post; no `Dispatch`/`GetActiveObject` creation.
4. STA affinity proof: `CoInitializeEx(COINIT_APARTMENTTHREADED)` on worker; all COM on same thread; queued/serialized demo.
5. Message pumping: run pump-capable wait/idle loop on STA during ops/waits; document behavior under busy/modal (HRESULT observations if safe); no UIA remediation in task_002 beyond observation unless trivial; no SendKeys/coords as primary.
6. Lifecycle: open test workbook, minimal non-destructive op on disposable copy if desired; `finally`: close owned workbooks (no silent save unless explicit), `Quit`, release refs, verify no orphan owned process.
7. Go/No-Go artifact (allowed non-source path, e.g. `.ai/tmp`) with pass/fail, builds/bitness/languages, HRESULT/errors, pump limitations, explicit Go/No-Go for `task_010`.

## 4) Constraints and risks
(table summarized in session): STA/marshaling, `DispatchEx` vs attachment, process leaks, `RPC_E_CALL_REJECTED`/busy/modal, Office/Windows variability, bitness, MCP stdout discipline, no untrusted macro execution in task_002. Mitigations listed per area.

## 5) Testing recommendations
- Run `tests/integration_excel/test_excel_instance_poc.py` on target Windows/Excel (interactive). Assertions mapped to ACs: isolation via PID/HWND/process delta; STA affinity; lifecycle/orphans; pump/COM-UI observations including RPC_E_CALL_REJECTED if reproducible; failure capture with HRESULT/context.
- Log to stderr and/or `.ai/tmp/...json|txt` with env, Excel/VBE details, step results, limitations, go_no_go.
- Minimal fixture: disposable macro-enabled workbook copy in `work/` or temp for open/close (can be simple generated/copied). No macro execution required.
- Preconditions: interactive desktop, Excel installed, record Trust Center notes if probed, avoid interfering automation.

## 6) Relevant authoritative sources
(See researcher message in session log context: Microsoft Learn COM/STA/message loop/IMessageFilter/Excel Application properties; pywin32 guidance references; Office server-side automation considerations.)

## 7) Actionable next-step checklist (for implementer)
(as listed in researcher output: preconditions, harness steps, evidence collection, Go/No-Go handoff to task_010/011/014 and gate failures.)
