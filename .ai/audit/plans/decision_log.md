# decision_log — Autonomous Excel Observability Architecture

- **Revision:** `04437f5` (`04437f582ba4050469ecd99f298e19f64c814650`)
- **Date:** 2026-11-25
- **Scope:** One entry per decision underpinning `reports/report_07` and `plans/plan_02`. Each entry records: **Decision · Alternatives considered · Evidence (anchors/URLs) · Classification (FACT/HYPOTHESIS/RECOMMENDATION) · Remaining product decision (U1/U2).**
- **Source of truth:** report_07. Anchors re-derived @`04437f5`, static + read-only (no runtime Excel). `src/` paths are under `src/excel_vba_mcp/`.
- **Non-negotiables (apply to every entry):** single-STA; originals never modified; pid-scoped kill only; no auto-retry of interrupted mutation; no pivot output materialization; no `Workbook.RefreshAll`; `SCHEMA_VERSION=1` additive; server never auto-terminates a live pid; fail-closed recovery.

---

## D1 — Atomic, lock-guarded status snapshot

- **Decision:** `get_status`/`excel_health` read `_active_*` **only** from a single lock-guarded **immutable** snapshot; enforce **clear-before-complete**; keep `list_workbooks` non-COM; a reconciled variant (if built) goes through the single STA or returns `busy`. Gates the multi-dimensional snapshot (C3).
- **Alternatives considered:** (a) reorder `_complete_result`/`_clear_active` only — **rejected**, insufficient (data race remains); (b) reconciled-as-default listing — rejected (forces COM into a non-COM path / blocks behind a wedged STA); (c) clear notes globally — rejected (loses history); (d) make `list_workbooks` reconcile on an arbitrary thread — rejected (violates single-STA).
- **Evidence:** `com_worker.py` `get_status` `:411-426` (lock `:414-416` covers only `state`/`accepting`; `_active_*` lock-free `:420-421`); writes from two threads — worker `:625-627`, loop `:397-399`; `_complete_result` `:632` before `finally: _clear_active()` `:639-640` (`_clear_active` `:839-848`); `session_service.py` `list_workbooks`/`list_workbooks_sync` `:527-541`.
- **Classification:** **FACT** (race observed in source) + **RECOMMENDATION** (atomic snapshot).
- **Remaining product decision:** none (U4 only — additive schema acceptance).

---

## D2 — Multi-dimensional state model + derived verdict (not one enum)

- **Decision:** Publish **five orthogonal dimensions** (Server / Worker / Process / Operation-Job / Recovery) and derive an **additive** verdict `PROGRESSING | INCONCLUSIVE | BLOCKED_ON_DIALOG | EXITED | SUSPECTED_STALLED | RECOVERY_RECOMMENDED | DONE` with `evidence[]` + `confidence`.
- **Alternatives considered:** (a) single status enum (`RUNNING/DEADLINE_EXCEEDED/CONFIRMED_HUNG/...`) — rejected, conflates independent facts and invites single-signal "confirmed"; (b) boolean `hung` flag — rejected, unprovable for opaque calls; (c) progress-only — rejected, no liveness/exit/dialog coverage.
- **Evidence:** report_06 D1 single-enum vocabulary; `com_worker.py:391,395-396` (one conflated deadline); macro 90 s kill `macro_service.py:107,620,828,866,1311-1332`; single-STA `com_worker.py:267,300,534`. Signal semantics: `WaitForSingleObject` signaled=EXIT / `WAIT_TIMEOUT`=alive; `#32770`=DIALOG (Microsoft Learn: `WaitForSingleObject`, window class `#32770`).
- **Classification:** **RECOMMENDATION** (state model) grounded in **FACT** (single conflated deadline).
- **Remaining product decision:** U2 (resolved: never auto-kill live pid).

---

## D3 — VBA file heartbeat as the only macro progress source

- **Decision:** The only non-COM macro progress is a **file heartbeat** using built-in VBA `Open`/`Print`/`Close` in **instrumented** macros. **Not** `CreateObject`/`DoEvents` (R13-vetoed). Absence of heartbeat = `INCONCLUSIVE`, never fabricated `%`.
- **Alternatives considered:** (a) `DoEvents` heartbeat — rejected (R13 interactive-UI veto); (b) `CreateObject` FSO — rejected (R13 veto, not background-safe); (c) no macro progress (elapsed only) — rejected as sole signal (slow≠hung); (d) fabricated percentage — rejected (dishonest).
- **Evidence:** `macro_service.py` R13 `INTERACTIVE_TOKENS` `:128-152` (`CreateObject` `:134`, `DoEvents` `:137`); BG06 background_safe gate `:22-34`; single-STA blocks `Application.*` reads mid-call `com_worker.py:300,534`.
- **Classification:** **FACT** (R13 veto list) + **RECOMMENDATION** (file heartbeat).
- **Remaining product decision:** none.

---

## D4 — Win32 trends as low-confidence corroborators (k-consecutive only)

- **Decision:** `GetProcessTimes`/memory/IO deltas, `IsHungAppWindow`, `SendMessageTimeout(WM_NULL)` are **low-confidence corroborators requiring k-consecutive confirmation**, **never the sole trigger** for `SUSPECTED_STALLED`. `WaitForSingleObject(handle,0)` is the only **direct** exit/alive signal.
- **Alternatives considered:** (a) single-sample `IsHungAppWindow` ⇒ stalled — rejected ("not for general use"; ghosting false positives); (b) `SendMessageTimeout` as primary — rejected (can itself block; modal pumps own loop); (c) `GetExitCodeProcess` 259 as hang — rejected (ambiguous); (d) elapsed time alone — rejected (slow≠hung).
- **Evidence:** Microsoft Learn: `IsHungAppWindow` "not for general use"; `WaitForSingleObject` `WAIT_TIMEOUT`≠hang; `GetExitCodeProcess` `STILL_ACTIVE`=259 ambiguous; modal `#32770` pumps its own message loop. `app_adapter.py` transient `OpenProcess(0x1000)` `:452-456` (query-limited info, not held).
- **Classification:** **FACT** (Win32 semantics) + **RECOMMENDATION** (k-consecutive corroboration).
- **Remaining product decision:** none.

---

## D5 — Pull/poll + JobRegistry + JobStore + reconcile

- **Decision:** Robust baseline = **pull/poll** (`*_start`→`{status:running, job_id, poll_after_s}` ≪60 s; agent re-calls `job_get`). A loop-side, lock-guarded, non-COM **JobRegistry owns the Future** (timeout/disconnect never drops the result). A new **JobStore** persists a **RUNNING marker at start**, **finalizes** it, and **reconciles RUNNING→unknown** at startup. Reuse the extracted `_atomic_write_json` (temp+fsync+`os.replace`); **do NOT reuse** the `TestRunResult`-typed `RunRepository`.
- **Alternatives considered:** (a) rely on request coroutine to deliver result — rejected (dropped if no awaiter, `com_worker.py:391`); (b) reuse `RunRepository` for jobs — rejected (typed `TestRunResult`); (c) in-memory only — rejected (reconnect/restart loses results); (d) auto-retry RUNNING on restart — rejected (unsafe; unknown never replayed); (e) detached broker now — deferred (blocked by job breakaway).
- **Evidence:** `com_worker.py:391` single end-to-end `wait_for` (result dropped if no awaiter; no registry); `run_repository.py` typed `TestRunResult` `:34,166`, reusable `_atomic_write_json` `:113`; `mcp` 2.3.0 has no Tasks (types only); Kilo `resetTimeoutOnProgress:true`, 60 s/call, no `maxTotalTimeout` [HYPOTHESIS → P5].
- **Classification:** **FACT** (no registry; RunRepository typing; helper present) + **RECOMMENDATION** (JobRegistry+JobStore).
- **Remaining product decision:** **U1** (tree-kill reload survival — detached broker deferred by default).

---

## D6 — Allow-by-default `excel_recover` (RECOVERY class)

- **Decision:** Add `excel_recover` as an **allow-by-default RECOVERY class** (auto-allowed in FULL **and** RESTRICTED; audited; pid-scoped; discards only working copies; response states destructive-to-working-copy). A **mutating** op stays denied in RESTRICTED.
- **Alternatives considered:** (a) deny-by-default like admin ops — rejected (recovery must be reachable autonomously); (b) reuse MUTATING class — rejected (would be blocked in RESTRICTED); (c) tool-only recovery without a class — rejected (no policy seam).
- **Evidence:** `security/approval.py`, `security/macro_policy.py` (policy classes); originals-safe invariant (originals hashed before/after); single-STA `com_worker.py:267,300`.
- **Classification:** **RECOMMENDATION** (policy class).
- **Remaining product decision:** none (risk sign-off via U2 resolution).

---

## D7 — Fail-closed held handle; same-handle terminate; no `taskkill`

- **Decision:** Hold a process **HANDLE** at launch (`PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION`); terminate+confirm via the **same handle** (`TerminateProcess(handle)`→`WaitForSingleObject(handle)`). Handle unobtainable / `Hwnd=0` / ambiguous identity ⇒ **fail closed** (recovery unavailable; **server restart** fallback). **No `taskkill` fallback.**
- **Alternatives considered:** (a) `taskkill /F /PID` — rejected (re-resolves pid; pid-reuse risk; current `app_adapter.py:470-491`, `macro_service.py:1311-1332`); (b) `taskkill` fallback if handle fails (report_06) — rejected (C5: silent weaker path; can hit a reused pid); (c) auto-kill on `Hwnd=0` — rejected (identity unproven).
- **Evidence:** `app_adapter.py` transient `OpenProcess(0x1000)` `:452-456` (**not held**; query-limited info only), `_terminate_pid` taskkill `:470-491`, reap `:207`; pid via `Hwnd`→`GetWindowThreadProcessId` `:429` / `macro_service.py:1294`; `Hwnd=0`⇒`owned_pid=None`⇒reap no-op. Microsoft Learn: `TerminateProcess(handle)` async + `WaitForSingleObject(handle)` confirm; `OpenProcess` rights.
- **Classification:** **FACT** (no handle held; taskkill today) + **HYPOTHESIS** (handle obtainable on `DispatchEx` child → **U5/P2**) + **RECOMMENDATION** (fail-closed).
- **Remaining product decision:** **U5** verify by P2 (fail-closed ⇒ server restart).

---

## D8 — No `CONFIRMED_HUNG` from a single signal

- **Decision:** Never assert `CONFIRMED_HUNG` from `WaitForSingleObject` (proves EXIT) or `#32770` (proves DIALOG) alone. Sustained ambiguity ⇒ `SUSPECTED_STALLED`/`RECOVERY_RECOMMENDED`; the **agent** decides recovery. Missing a hang costs time, not data.
- **Alternatives considered:** (a) `WAIT_TIMEOUT`⇒hung — rejected (alive≠hung); (b) `#32770`⇒hung — rejected (dialog≠hang); (c) elapsed>deadline⇒hung — rejected (slow≠hung); (d) single `IsHungAppWindow`⇒hung — rejected (unreliable).
- **Evidence:** Microsoft Learn `WaitForSingleObject`/`GetExitCodeProcess`/`IsHungAppWindow`; report_06 D1; validator C4; single-STA opaque-call ceiling `com_worker.py:300,534`.
- **Classification:** **FACT** (Win32 semantics) + **RECOMMENDATION** (verdict discipline).
- **Remaining product decision:** none.

---

## D9 — Remove the macro 90 s auto-kill; unify under agent-directed path

- **Decision:** **Remove** `_DEFAULT_TIMEOUT_S=90.0` + `_handle_hung` + `_terminate_owned_pid` auto-kill; route `macro_run` through the **unified** per-op-class deadline (deadline ⇒ `SUSPECTED_STALLED`/`RECOVERY_RECOMMENDED`; agent calls `excel_recover`). Single invariant: **server never auto-terminates a live pid**; only agent-directed `excel_recover` or dead-pid lazy recovery.
- **Alternatives considered:** (a) keep 90 s kill — rejected (violates no-fixed-timeout-auto-kill; slow≠hung); (b) reconcile/raise the 90 s threshold — rejected (still a fixed-timeout kill); (c) narrow EXECUTING auto-reap (report_06 D2/U2) — rejected (C1: no live-pid auto-kill).
- **Evidence:** `macro_service.py` `_DEFAULT_TIMEOUT_S=90.0` `:107`; `_handle_hung` `:828` (invoked `:620`) → `_terminate_owned_pid` `:866` → `taskkill /F /PID` `:1311-1332` (taskkill `:1325`); generic path dead-pid-only `session_service.py:729-757` (guard `:754`).
- **Classification:** **FACT** (90 s live kill present) + **RECOMMENDATION** (remove; unify).
- **Remaining product decision:** U2 **resolved** (never auto-kill live pid).

---

## D10 — Detached broker deferred

- **Decision:** **Defer** a detached broker / on-disk-job-independent-of-stdio. Document that pull/poll + in-memory registry does **not** survive a client tree-kill/server restart; only JobStore-persisted results are retrievable after reconnect.
- **Alternatives considered:** (a) detached worker now — rejected (blocked by job breakaway; single-STA redesign; cannot save an in-flight call); (b) rely on MCP Tasks — rejected (absent in `mcp` 2.3.0); (c) heavyweight Operation Manager — deferred.
- **Evidence:** Windows kill = direct-child only (EXCEL orphaned) [HYPOTHESIS → P5]; `mcp` 2.3.0 no Tasks; single-STA `com_worker.py:267,300`; JobStore persistence `run_repository.py:113` (reused).
- **Classification:** **HYPOTHESIS** (transport kill behavior → P5) + **RECOMMENDATION** (defer).
- **Remaining product decision:** **U1** (if "must survive tree-kill" ⇒ detached broker becomes an F blocker).

---

## D11 — Pivot single-sweep + per-sheet measured budget (Block B)

- **Decision:** Single-sweep + `CacheIndex` memoization + **SourceType gate FIRST**; per-sheet **MEASURED** budget checked **BETWEEN completed COM ops** (never interrupts in-flight); partial ⇒ `capped`; `too_large` distinct; current op stays `running`/`long_running`. **Never** key on file/row/UsedRange. A slow single COM stays running/slow, never too_large/hung.
- **Alternatives considered:** (a) file-size/row/UsedRange trigger — rejected (false `too_large` on 5.5 MB / 566k-row); (b) interrupt in-flight COM at budget — rejected (cannot safely preempt `:602`); (c) bigger per-class budget — rejected (masks); (d) auto-kill — rejected (slow≠hung).
- **Evidence:** `pivot_tables_service.py` `list_pivot_tables` `:171-172`, `_collect_caches` `:899-904`, `_find_pivot_table` `:906-916`, `_DEFAULT_TIMEOUT_S=60.0` `:98`; adapter `_UNSUPPORTED_SOURCE_TYPES` ~`:188-199`, `_require_supported_source` ~`:833-851`; `Workbook.PivotTables` incomplete / no `PivotCache.PivotTables` (must sweep worksheets + `CacheIndex`).
- **Classification:** **FACT** (re-sweeps; 60 s; no guard) + **RECOMMENDATION** (single-sweep + measured budget).
- **Remaining product decision:** none.

---

## D12 — Paths / docs (Block C)

- **Decision:** `original_name` stays **absent** (do not add; `source_path` exists). Log resolved roots at startup (replace silent fallback). Docs: inspect-is-writable, double-I/O. `SCHEMA_VERSION=1` additive.
- **Alternatives considered:** (a) add `original_name` — rejected (redundant/absent); (b) hardlink/move originals — rejected (breaks originals-never-modified); (c) runtime root add — rejected (weakens deny-by-default); (d) silently flip `read_only=True` — deferred (contract change).
- **Evidence:** `domain/models.py` `WorkbookInfo` `:83-96` (`source_path` `:90`; `original_name` absent); `server.py` silent root fallback `:765-774`; `api/workbook_tools.py` `excel_open` defaults `:86-87`; `session_service.py` double-I/O `:867-874` (`copy2` `:873`).
- **Classification:** **FACT** (anchors) + **RECOMMENDATION** (log + docs).
- **Remaining product decision:** none (U4 additive-schema acceptance).

---

## Open decisions (summary)

- **U1 (product):** survive a client **tree-kill** on reload? **Default:** graceful reconnect + completed-job retrieval; detached broker deferred. **Branch (yes):** detached broker becomes an F blocker (blocked by job breakaway).
- **U2 (risk):** auto-kill a **live** pid at the deadline? **RESOLVED (C1): never** — agent-directed `excel_recover` or dead-pid lazy recovery only.
- **U3 (verify P5):** real Kilo timeout/progress — **HYPOTHESIS** until exercised.
- **U4:** additive `SCHEMA_VERSION=1` acceptance — confirm consumers tolerate additive v1.
- **U5 (verify P2):** `PROCESS_TERMINATE` handle on the `DispatchEx` child — **HYPOTHESIS**; fail-closed ⇒ server restart.

---

*End of decision_log. Anchors @`04437f5`, static + read-only; no runtime Excel launched. Authorizes no code change; see plan_02 for execution.*

---

# decision_log — Agent Usability Remediation (plan_03)

- **Revision:** `e2289ca` (`e2289ca01fd10136716a598cad1db6b6845fed73`)
- **Date:** 2026-11-26
- **Scope:** Decisions underpinning `reports/report_08`, `report_08a`, and `plans/plan_03 — Agent Usability Remediation.md`. Each entry: **Decision · Alternatives · Evidence · Status (confirmed/recommended/assumed/unresolved) · Compatibility impact.**
- **Non-negotiables:** single hidden Excel + single-STA FIFO worker; Job Object lifecycle; agent-directed `excel_recover`; originals never modified; working-copy model; `INTERRUPTED_UNKNOWN` never auto-retried; no pivot output materialization; no `Workbook.RefreshAll`; no mandatory human approval in normal flows; `SCHEMA_VERSION=1` additive.

## D-U1 — Token strategy = slim output schemas, not a search tool
- **Decision:** Reduce upfront tool-context by slimming auto-generated `outputSchema` (DTO/field slimming, `title` suppression) while KEEPING `structuredContent`. Do NOT add a custom `xlvba_search_tools`, do NOT rely on `defer_loading`, do NOT use `structured_output=False` (it also removes `structuredContent`).
- **Alternatives:** custom search tool (rejected: adds a 92nd always-loaded definition → net context increase; Validator U4 REJECT); `defer_loading` (rejected: absent from mcp/mcp-types 2.3.0); `structured_output=False` (rejected: kills structuredContent clients rely on).
- **Evidence:** Auditor reproduced `tools/list`=451,108 chars, outputSchema=75.05%, descriptions=5.4%, input=17.9%, 91/91 emit outputSchema, 0 annotations; `defer_loading` 0 hits in installed stack.
- **Status:** confirmed. **Compatibility:** additive/opt-in; must verify no consumer depends on removed schema fields.

## D-U2 — Structured error = additive text-JSON envelope inside isError content
- **Decision:** On `isError=true`, return a JSON envelope as text content with `errorCode, errorCategory, message, isRetryable, suggestedNextActions, details`. No SDK change; `SCHEMA_VERSION` stays 1.
- **Alternatives:** error `structuredContent` (rejected: SDK `ToolError` is text-only `isError` — `server.py:456`, `func_metadata.py:226`; infeasible without SDK change); prose-only (rejected: not machine-readable).
- **Evidence:** Validator carve-out; Auditor SDK inspection.
- **Status:** confirmed (feasible). **Compatibility:** additive; existing clients reading text still work.

## D-U3 — Environment/policy discovery = new read-only `xlvba_environment` + additive `policy` on `excel_get_status`
- **Decision:** Add read-only `xlvba_environment` (server/excel/capabilities/policy/workbooks). `has_snapshot` reads the in-memory gate (`_latest_snapshot`), not the store. Do not conflate static capabilities with live Excel state.
- **Alternatives:** extend only `server_health` (rejected: liveness ≠ policy); only `excel_get_status` (rejected: needs server/capability blocks).
- **Evidence:** 1B discoverability matrix; Auditor (access_mode already via `excel_get_status`; snapshot gate in-memory).
- **Status:** recommended. **Compatibility:** additive new tool + additive field.

## D-U4 — Cancellation: keep `job_cancel` name; fix description only
- **Decision:** Do NOT rename `job_cancel` → `job_request_cancel`. Correct the description to state intent-only / runs-to-completion. Revisit alias only if a consumer demands it.
- **Alternatives:** rename + alias (rejected now: high compat, low value; description already accurate per Auditor).
- **Evidence:** Validator U3 REJECT-unless-aliased; Auditor job_cancel Major→Minor.
- **Status:** confirmed. **Compatibility:** none (description-only).

## D-U5 — Do not collapse test tools or rename `vba_compile`
- **Decision:** Keep `test_run_suite` and `test_suite_run` as separate tools; add routing description (canonical = `test_suite_run` for new use; `test_run_suite` = legacy). Keep `vba_compile` name; fix description to "diagnostics only, never verified compilation".
- **Alternatives:** collapse/rename (rejected: HIGH compat — tests, docs, `>=87` tool-count assert at `test_mcp_full_access_workflow.py:499`).
- **Evidence:** Auditor compat risks; 1A (vba_compile description honest → naming issue).
- **Status:** confirmed. **Compatibility:** preserved.

## D-U6 — Enum pruning is a schema change and must ship WITH description fixes
- **Decision:** Prune rejected request-enum values (`range_set_conditional_formatting`: color_scale/data_bar/icon_set; `pivot_field_set_layout`: `data`) to accepted values; keep read-side enums intact. Each change is additive-safe and versioned.
- **Alternatives:** leave enums and only document (rejected: schema still advertises rejected values — Validator Q15 FAIL).
- **Evidence:** 1A contradictions CONFIRMED; ranges_adapter.py:2247-2252; pivot_tables_service.py:958-968.
- **Status:** recommended. **Compatibility:** low (request-side only); test.

## D-U7 — Truthful MCP annotations
- **Decision:** Emit `readOnlyHint`/`destructiveHint`/`idempotentHint` only where accurate, from existing access-class metadata.
- **Alternatives:** omit (rejected: omitted = dangerous defaults per audit).
- **Evidence:** SDK `ToolAnnotations` supported but unemitted (Auditor).
- **Status:** recommended. **Compatibility:** additive; client-dependent display.

## D-U8 — Glossary as resource + compact mirror
- **Decision:** Publish glossary as a static MCP resource and mirror key terms into `xlvba_environment`/a compact doc; descriptions reference it rather than repeat.
- **Alternatives:** repeat in every description (rejected: token cost); tool-only (rejected: 2A resource support partial).
- **Evidence:** 2A (resources supported server-side; Kilo forwarding unverified).
- **Status:** recommended. **Compatibility:** additive.

## D-U9 — `save_as`/export and `xlvba_job_list` are deferred product decisions (U5)
- **Decision:** Investigate only; do NOT implement in this remediation. `save_as` needs a separate product contract (original protection, output path policy, overwrite, result retrieval).
- **Alternatives:** implement now (rejected: product scope, path-policy risk).
- **Evidence:** 2B T28 gap; no persistence tool today.
- **Status:** unresolved (product). **Compatibility:** n/a.

## D-U10 — `isRetryable` must not invite blind replay
- **Decision:** `STALE_TARGET`/`REVISION_CONFLICT` are NOT `isRetryable=true` without a fresh read; `INTERRUPTED_UNKNOWN` is never auto-retried. Retryability is per-code and conservative.
- **Alternatives:** blanket retryable (rejected: partial-application risk).
- **Evidence:** Validator blocking concern #4.
- **Status:** confirmed. **Compatibility:** additive.

*End of plan_03 decision_log section.*

## Implementation outcome (2026-11-26)

- **U1 (descriptions + glossary + docs):** IMPLEMENTED + VALIDATED (ACCEPT). 25 critical descriptions on the 8-section standard; `xlvba://glossary` (14 terms); docs EN+RU corrected; tool count 91.
- **U2 (environment/policy + error envelope):** IMPLEMENTED + VALIDATED (ACCEPT). `xlvba_environment` (read-only, non-COM, `build:null`, in-memory `has_snapshot`); additive `SessionStatus.policy`; additive error envelope (`errorCategory`/`isRetryable`/`suggestedNextActions`, D-U10). Tool count 92. No SDK change.
- **U3 (enum prune + annotations):** IMPLEMENTED + VALIDATED (ACCEPT). Request enums pruned (read enums intact; `PivotFilterState` preserves read fidelity); truthful annotations 92/92; 6 integration tests aligned to pydantic rejection.
- **U4 (outputSchema slimming):** NO SAFE FIELD-PRESERVING REDUCTION — documented no-op (`docs/09`). Token-reduction objective NOT met; context +6.8% (U2 DTO + richer descriptions). Further reduction needs a product decision (D-U9 follow-up).
- **U5 (deferred product contracts):** DELIVERED as documents only (`plan_03a`); `save_as`/export and `xlvba_job_list` NOT implemented (require product approval).
- **Verification:** unit 1127 passed; ruff clean; integration collects; final Auditor spot-check ACCEPT (no blocking inconsistencies). Full model-in-the-loop 28-task benchmark specified (`.ai/tmp/benchmark_tasks.yaml`) + deterministic proxy run (28/28 first-tool, 29/29 arg feasibility); model-in-the-loop run remains (no agent runtime here).
- **Unresolved (product):** token/outputSchema reduction; `save_as`/`job_list`; Kilo forwarding of outputSchema/annotations/resources; real 4-part Excel build.

---

# decision_log — Remaining Usability & Save As (plan_05)

- **Revision:** baseline `40c8b4d`; final working tree (U6-A..E + U7 + U6-D).
- **Scope:** naming, error envelope, routing, confirm, Save As. Backward compatibility NOT required; lazy loading excluded.

## D-R1 — Additive error codes — CONFIRMED
Add `INTERNAL_ERROR` (hatch fallback, category internal, not retryable) + `CONFIRMATION_REQUIRED` (missing/false confirm, category validation, not retryable). Extend totality maps. SCHEMA_VERSION=1.

## D-R2 — Error envelope keys at the MCP boundary — CONFIRMED
`errorCode, errorCategory, message, isRetryable, suggestedNextActions, details` (+ additive `hresult`/`hresult_hex`). Text-JSON only (SDK ToolError is text-only); no error `structuredContent`; hatch closed so non-domain errors never leak "Error executing tool".

## D-R3 — Canonical names, no aliases — CONFIRMED
`vba_compile_diagnostics`, `job_request_cancel`, `test_suite_run` (sole test entry; `test_run_suite` removed). `TestAction.VBA_COMPILE` untouched.

## D-R4 — Save As = one `excel_save_as`, fail-closed — CONFIRMED
`(workbook_id, path, *, overwrite_existing=False, overwrite_source=False)`. Session stays bound (flush `wb.Save()` + atomic copy; never `wb.SaveAs`). Source overwrite only via `overwrite_source` + hash match; existing dest only via `overwrite_existing`; drift → `STALE_TARGET`; missing flags → `PROTECTED_TARGET`.

## D-R5 — Source identity + atomic + single `.bak` — CONFIRMED
Capture `source_sha256`+size at open before materialize; re-hash + normcase/realpath compare at Save As. Atomic temp→replace; single `<target>.bak` (never overwritten if pre-existing → `unknown`, no auto-resume). `outcome="ok"` only after final verify.

## D-R6 — Confirm flags — CONFIRMED
`snapshot_restore`/`excel_recover` require `confirm=true` (agent intent, not human approval); reject before action; descriptions state discard vs preserve + external-side-effect caveat.

## D-R7 — Routing + compression — CONFIRMED
Capability-based routing in `xlvba_environment` (no other-server detection). Guard/approval boilerplate → glossary tokens; safety text kept.

## Implementation outcome (2026-11-26)
U6-B/A/C/E/U7/U6-D all implemented + Validator-ACCEPT; alignment pass fixed integration envelope keys + docs/08; unit 1152 passed; integration collects; final Auditor spot-check ACCEPT. Full model-in-the-loop benchmark remains (no runtime here). Deferred: PDF export, `xlvba_job_list`, outputSchema field reduction.

## Stage 1 outcome (2026-11-26)
- Save As live integration (Excel 16.0, PID 27704): 9 scenarios PASS; scenario-6 `.bak` conflict defect fixed (deterministic `PROTECTED_TARGET/backup_exists`); candidate real-agent run PASS.
- Comparative baseline-vs-candidate model-in-the-loop benchmark: **BLOCKED** (only candidate server reachable; Kilo global config read-restricted; no baseline routing). No proxy substituted.
- **Stage 1 = FAIL (blocked on comparative benchmark). Stage 2 NOT started.** Remediation: run the comparative benchmark where both variants are routable, then repeat Stage 1D.

## Stage 2 outcome (2026-11-26, benchmark waived by user)
- User waived the comparative benchmark; Stage 1 treated as PASS for gating.
- Stage 1 fix (deterministic `PROTECTED_TARGET/backup_exists`) committed `fefc9ed`; docs `55b9398`; ruff clean; unit 1152 pass.
- **Stage 2 ACCEPTED; rollout COMPLETED** on local `master` (no remote). Post-deploy health: 92 tools, canonical names present, single owned Excel, accepting. Rollback: revert `55b9398 fefc9ed` → `ef3e6cb`. Open (non-blocking): comparative benchmark (waived); post-restart live re-test of scenario-6 fix.

## Research B & D acceptance (2026-11-27, HEAD 24201f4)
- Auditor `report_12` CONFIRMED all nine claims; two minor citation typos corrected (orphan-join production default **10 s** not 3 s; `_find_open_workbook` at **1684–1700**).
- **Research B — CLOSED.** Accept shipped source-type gate (B1). Do NOT implement shorter timeouts (B2 rejected). Background execution (B3) **deferred** as an optional future WP, kept SEPARATE from D. Original Pivot hang investigation closed.
- **Research D — ACCEPTED, no rewrite.** WP3 lifecycle stands; R1–R3 are **optional hardening**, not invariant fixes. A separate bounded plan (`plan_06 — Snapshot Restore Hardening R1-R3.md`) is prepared; **no production code change for R1–R3 until that plan is reviewed and approved.** Preserve original workbook / immutable snapshot / mutable working copy distinction.
- No production code changes authorized by this acceptance.