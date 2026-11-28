# plan_02 — Autonomous Excel Implementation Plan

- **Purpose:** Standalone, executable plan derived from `reports/report_07 — Autonomous Excel Observability Architecture.md`. An implementor with **no prior context** can execute it. It converts the observability architecture (5 state dimensions, derived verdict, agent-directed recovery, JobRegistry+JobStore) into dependency-safe, independently reviewable work items with explicit acceptance criteria, test strategy, PoC gates, and open decisions.
- **Revision:** `04437f5` (`04437f582ba4050469ecd99f298e19f64c814650`)
- **Date:** 2026-11-25
- **How to use:** Read §1 (non-negotiables) and §2 (decisions that must hold) first — binding. Execute work items in **DAG order** (§4): **D → E → A → F**, with **B ∥ E** and **C anytime**. Do not start a dependent block until its **PoC gate** (§5) passes. Record exact build/test/PoC commands and artifacts. **Plan only — no implementation code.**
- **Source of truth:** report_07. If this plan and report_07 disagree, report_07 wins; raise the discrepancy, do not silently diverge. report_02–06 and the validity file are **not** modified by planning.
- **Evidence legend (used throughout):** **[LIVE-VERIFIED 2026-11-25]** observed on a live Excel instance / running server on 2026-11-25 (copies only; originals hashed before/after); **[PoC-MODEL]** a modeled/simulated result (e.g., the P1 torn-read harness), **not** a live observation; **[STATIC]** code-verified by read-only source inspection at `04437f5`, **not** runtime-proven.

---

## 1. Non-negotiables (binding on every work item)

1. **Single-STA:** exactly one hidden Excel instance in the interactive user session. No VM, no container, no second STA, no per-op workers.
2. **Originals never modified.** All work on working copies. Originals hashed before/after every PoC; byte-identical.
3. **pid-scoped kill only.** Never a wildcard image kill (`taskkill /IM`, `Stop-Process -Name`). Blast radius = exactly one owned pid, confirmed via a **held handle**.
4. **No auto-retry of an interrupted mutation.** Interrupted = `unknown`, never replayed.
5. **No pivot output materialization.** Never expand pivot output to cells; never open `pivotCacheRecords`.
6. **No `Workbook.RefreshAll`.** Cache refresh is table-scoped/cache-scoped only.
7. **`SCHEMA_VERSION=1`, additive only.** New enum members/fields are additive; no breaking DTO change.
8. **Server never auto-terminates a live pid (C1).** Only agent-directed `excel_recover` or dead-pid lazy recovery terminates anything. **No fixed-timeout auto-kill.**
9. **Fail-closed recovery (C5).** Handle unobtainable / `Hwnd=0` / ambiguous identity ⇒ recovery unavailable ⇒ server restart fallback. **No `taskkill` fallback.**
10. **English only** in code, comments, logs, docs.
11. **Planning only:** do not modify report_02–07, the validity file, production code, or tests as part of planning — this plan authorizes the *future* implementation work items only.

### Definition of done (whole plan)

- D, E, A, F implemented and independently reviewed; B and C implemented.
- All PoC gates P1–P7 pass (or their branch explicitly accepted per §6).
- Unit + integration tests added/updated per block; originals hash-unchanged across all PoCs.
- U1/U2 resolved with recorded sign-off (default accepted or branch chosen); **U5 verified by P2 (RESOLVED-POSITIVE)**; **U3 OPEN** (P5 did not exercise progress/reset — polling must not depend on it).
- Docs updated (Block C); `SCHEMA_VERSION` still `1`; no fixed-timeout auto-kill remains.

---

## 2. Architectural decisions (must hold)

Implementations must satisfy these verbatim (from report_07).

### 2.1 Five-dimensional state + derived verdict

- Publish **five** dimensions — **Server** / **Worker** / **Process** / **Operation-Job** / **Recovery** (fields per report_07 §3.1).
- Derive an **additive** verdict: `PROGRESSING | INCONCLUSIVE | BLOCKED_ON_DIALOG | EXITED | SUSPECTED_STALLED | RECOVERY_RECOMMENDED | DONE`, with `evidence[]` + `confidence`.
- **NEVER `CONFIRMED_HUNG` from one signal.** `SUSPECTED_STALLED` requires **k-consecutive** corroboration (C4) and is **not** a kill trigger.

### 2.2 Signal discipline

- Non-COM progress = **server bookkeeping + Win32 liveness + VBA file heartbeat** (built-in `Open`/`Print`/`Close` only; **not** `CreateObject`/`DoEvents` — R13-vetoed).
- **i/N checkpoints** only for **server-controlled loops**; **file heartbeat** only for **instrumented macros**.
- **Win32 trends** (CPU/IO/mem, `IsHungAppWindow`, `SendMessageTimeout`) = **low-confidence corroborators requiring k-consecutive confirmation, never the sole trigger.**
- **Opaque calls emit no mid-call signal ⇒ absence of progress = INCONCLUSIVE** (never fabricate `%`).

### 2.3 Agent-facing contract

- `excel_health` (5 dims + verdict), `job_get(job_id)`, optional `job_events`, optional `excel_dialogs`, `excel_recover`.
- `poll_after_s` clamped **[1, 30]**, default **5 s** (≪60 s client wall-clock).
- Additive, `SCHEMA_VERSION=1`; **progress optional**, pull/poll is the robust baseline.

### 2.4 Recovery safety

- Hold a process **HANDLE** at launch (`PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION`); terminate+confirm via the **same handle**: `TerminateProcess(handle)` → `WaitForSingleObject(handle)`.
- **Fail closed** if handle unobtainable (C5); **no `taskkill` fallback**.
- `excel_recover` = **allow-by-default RECOVERY class** (FULL **and** RESTRICTED; audited; pid-scoped; discards only working copies). A **mutating** op stays denied in RESTRICTED.
- Interrupted mutation = `unknown`, **never auto-retried**; snapshot restore = file-consistent only.
- **Remove the macro 90 s auto-kill**; route `macro_run` through the unified agent-directed path (C1).

### 2.5 Job lifecycle & transport

- **JobRegistry owns the Future** (loop-side, lock-guarded, non-COM) wrapping the FIFO `submit()`; timeout/disconnect never drops a result.
- **JobStore** (new sibling of `RunRepository`; reuse extracted `_atomic_write_json`): persist **RUNNING at start**, **finalize** on completion; **startup reconcile RUNNING→unknown** (never completed/retried).
- `job_cancel` intent-only; idempotency key; pull/poll; **detached broker deferred** (blocked by job breakaway).
- **E/F boundary (RESOLVED):** the **minimal job registry + prompt `job_id` response** are implemented in **Block E**; E is **not accepted with a non-functional polling contract**. F adds the full `JobStore` + reconcile + detached-limit documentation.

### 2.6 Critical code map

All paths under `src/excel_vba_mcp/`. Anchors re-derived @`04437f5` (line numbers are hints; prefer the named symbol).

**Block D — atomic snapshot** (`adapters/excel/com_worker.py`)
- `get_status()` `:411-426` — lock `:414-416` covers only `state`/`accepting`; `_active_*` read lock-free `:420-421`.
- Active writes: worker `:625-627`, loop `:397-399`; `_complete_result` `:632` **before** `finally: _clear_active()` `:639-640`; `_clear_active` `:839-848`.
- `application/session_service.py` `list_workbooks`/`list_workbooks_sync` `:527-541` (non-COM); `api/health_tools.py` `excel_health` (non-COM).

**Block E — unified timeout/liveness + remove 90 s kill** (`adapters/excel/com_worker.py`, `application/macro_service.py`, `application/session_service.py`, `adapters/excel/app_adapter.py`, `config.py`)
- `com_worker.py` `submit()` `:391`; `TimeoutError`→`NEEDS_RECOVERY` `:395-396`; pre-start deadline `:602`.
- `macro_service.py` `_DEFAULT_TIMEOUT_S=90.0` `:107`; `_handle_hung` `:828` (invoked `:620`) → `_terminate_owned_pid` `:866` → `taskkill /F /PID` `:1311-1332` (**remove**).
- `session_service.py` `_maybe_auto_recover` `:729-757` (dead-pid guard `:754`).
- `app_adapter.py` `_terminate_pid` `:470-491`; reap `:207`; transient `OpenProcess(0x1000)` `:452-456`.
- `config.py` — add `open_timeout_s` / `liveness_ceiling_s` env knobs.

**Block A — recovery (handle, excel_recover, fail-closed)** (`adapters/excel/com_worker.py`, `adapters/excel/app_adapter.py`, `application/session_service.py`, `application/snapshot_service.py`, `security/approval.py`, `security/macro_policy.py`, `api/health_tools.py`)
- `com_worker.py` `_WorkItem` `:133-147` (add `started`); interrupted-mutation path `:388-409`.
- `app_adapter.py` `_terminate_pid` `:470-491` → same-handle `TerminateProcess`+`WaitForSingleObject`; identity `Hwnd`→`GetWindowThreadProcessId` `:429`.
- `session_service.py` `_maybe_auto_recover` `:729-757`; `snapshot_service.py` restore contract `:18-22`.
- `security/approval.py` + `security/macro_policy.py` — new allow-by-default RECOVERY class; `api/health_tools.py` `excel_recover`.

**Block F — JobRegistry + JobStore + pull/poll + reconcile** (`server.py`, `adapters/excel/com_worker.py`, `storage/run_repository.py`, new `storage/job_store.py`, new `application/job_registry.py`)
- `server.py` tool surface — add `job_get`/`job_cancel`; long ops return `{status:"running", job_id, poll_after_s}`.
- `com_worker.py` single `submit` (reuse; **no** second STA).
- `storage/run_repository.py` — **do NOT reuse** (typed `TestRunResult` `:34,166`); reuse extracted `_atomic_write_json` `:113` in the new `JobStore`.
- New `application/job_registry.py` (owns Future, lock-guarded, non-COM) + new `storage/job_store.py` (RUNNING marker + finalize + reconcile).

**Block B — pivot single-sweep + budget (∥ E)** (`application/pivot_tables_service.py`, `adapters/excel/pivot_tables_adapter.py`)
- `pivot_tables_service.py` `list_pivot_tables` `:171-172`, `_collect_caches` `:899-904`, `_find_pivot_table` `:906-916`, `_DEFAULT_TIMEOUT_S=60.0` `:98`.
- `pivot_tables_adapter.py` `_UNSUPPORTED_SOURCE_TYPES` (~`:188-199`), `_require_supported_source` (~`:833-851`).

**Block C — paths/docs (anytime)** (`server.py`, `domain/models.py`, `api/workbook_tools.py`, `application/session_service.py`, `docs/`)
- `server.py` root resolution `:765-774` (add startup log); `domain/models.py` `WorkbookInfo` `:83-96` (`source_path` `:90`; **no `original_name`**); `api/workbook_tools.py` `excel_open` `:86-87`; `session_service.py` double-I/O `:867-874`; `docs/`.

---

## 3. Work item template

Each item below states: **Goal · Scope (files) · Tasks · Acceptance criteria · Test strategy · PoC gate · Dependencies · Size (S/M/L)**. Items are atomic and independently reviewable. No implementation code is included.

---

## 4. Sequenced work items (DAG order)

> Order: **D → E → A → F**; **B ∥ E**; **C anytime**.

### Work item D — Atomic runtime state snapshot  ·  Size **S**  ·  Priority 1 (ship first)

- **Goal:** Make `get_status`/`excel_health` truthful and race-free; publish the **atomic snapshot** every later block observes (gates the multi-dimensional snapshot, C3).
- **Scope (files):** `adapters/excel/com_worker.py`, `application/session_service.py` (read-only use), `api/health_tools.py` (verify non-COM).
- **Tasks:**
  1. Introduce a single **lock-guarded immutable** status snapshot; `get_status` reads `_active_*` **only** from it (never lock-free `:420-421`).
  2. Enforce **clear-before-complete**: `_complete_result` (`:632`) publishes, then `finally: _clear_active()` (`:639-640`) clears — never expose a stale active op after completion.
  3. Keep `list_workbooks`/`list_workbooks_sync` **non-COM** (`session_service.py:527-541`).
  4. If `list_workbooks_reconciled` is built, route via the single STA (`submit`) or return `busy`; never an arbitrary thread.
- **Acceptance criteria:** no torn `_active_*` read; `excel_health` accurate while busy; `list_workbooks` non-COM; reconciled variant never spawns a thread.
- **Test strategy:** unit — concurrent reader/writer (no torn state); integration — `excel_health` during a slow op shows the correct active op and clears after completion.
- **PoC gate:** **P1** must pass before E.
- **Dependencies:** none.

### Work item E — Unified timeout / liveness + remove macro 90 s kill (keystone)  ·  Size **M**  ·  Priority 2 (foundation)

- **Goal:** One deadline mechanism for all COM ops; **remove** the macro 90 s auto-kill; split queue-wait from exec; emit the unified verdict vocabulary (no live-pid auto-kill).
- **Scope (files):** `adapters/excel/com_worker.py`, `application/macro_service.py`, `application/session_service.py`, `adapters/excel/app_adapter.py`, `config.py`.
- **Tasks:**
  1. Split queue-wait (returns `busy`/`retry_after_s`, **never** `NEEDS_RECOVERY`) from exec deadline (drives `SUSPECTED_STALLED`/`RECOVERY_RECOMMENDED`, **not** a kill).
  2. **Remove** `_DEFAULT_TIMEOUT_S=90.0` + `_handle_hung` + `_terminate_owned_pid` auto-kill (`macro_service.py:107,620,828,866,1311-1332`); route `macro_run` through the unified path.
  3. Add env knobs `EXCEL_VBA_MCP_OPEN_TIMEOUT_S=300`, `EXCEL_VBA_MCP_LIVENESS_CEILING_S=600` (`config.py`).
  4. Define per-tool `timeout_s` as a **server-side soft hint** (no client contract).
  5. Long ops return `{status:"running", job_id, poll_after_s}` ≪60 s (contract shared with F); `poll_after_s` clamp [1,30] default 5.
  6. Emit derived verdict + `evidence[]`/`confidence`; **no** code path terminates a live pid.
- **Acceptance criteria:** slow-but-alive ⇒ `INCONCLUSIVE`/`PROGRESSING` (+`busy`/`retry_after_s`), no poison; deadline ⇒ `SUSPECTED_STALLED` (k-consecutive), never a kill; env knobs honored; **no 90 s kill remains**; native call never cancelled; **minimal job registry + prompt `job_id` response present (E not accepted with a non-functional polling contract)**.
- **Test strategy:** unit — queue-wait timeout does not set `NEEDS_RECOVERY`; deadline yields `SUSPECTED_STALLED` not kill. integration — a deliberately slow op returns `INCONCLUSIVE`; a blocked op yields sustained `SUSPECTED_STALLED`; env override changes ceiling.
- **PoC gate:** **P2, P3** must pass before A and F.
- **Dependencies:** D.

### Work item A — Recovery: held handle, `excel_recover`, fail-closed  ·  Size **M**  ·  Priority 3

- **Goal:** Off-STA, pid-scoped, **agent-directed** recovery via a held handle; fail-closed when the handle is unobtainable; allow-by-default `excel_recover`.
- **Scope (files):** `adapters/excel/com_worker.py`, `adapters/excel/app_adapter.py`, `application/session_service.py`, `application/snapshot_service.py`, `security/approval.py`, `security/macro_policy.py`, `api/health_tools.py`.
- **Tasks:**
  1. Add `started` flag to `_WorkItem` (`:133-147`).
  2. Hold a process **HANDLE** at launch (`PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION`); terminate+confirm via the **same handle** (`TerminateProcess`→`WaitForSingleObject`), replacing `taskkill /F /PID` (`app_adapter.py:470-491`).
  3. **Fail closed**: handle unobtainable / `Hwnd=0` / ambiguous identity ⇒ **recovery unavailable** ⇒ server restart fallback (documented). **No `taskkill` fallback** (C5).
  4. Add allow-by-default **`excel_recover`** = originals-safe RECOVERY class (auto-allowed FULL **and** RESTRICTED; audited; pid-scoped; discards only working copies). Mutating op stays denied in RESTRICTED.
  5. Interrupted mutation = **UNKNOWN, never auto-retry** (`com_worker.py:388-409`); snapshot restore file-consistent only; external macro side effects not undone (`snapshot_service.py:18-22`).
  6. **No server auto-kill of a live pid**; only agent `excel_recover` or dead-pid lazy recovery (`session_service.py:754`).
- **Acceptance criteria:** `excel_recover` works in RESTRICTED and a mutating op does not; originals hash unchanged; no wildcard kill; handle-unobtainable ⇒ fail closed (no kill); no live-pid auto-kill anywhere; the `PyHANDLE` is held as a **live object for the entire ownership session** (never reduced to an `int` — P7 observed a false `EXITED` when reduced to `int`).
- **Test strategy:** unit — `started` transitions; same-handle terminate/confirm; fail-closed branch. integration — recovery on a working copy with a `background_safe` infinite-loop macro; policy test RECOVERY-vs-MUTATING in RESTRICTED.
- **PoC gate:** **P2, P7** must pass before F.
- **Dependencies:** E (classification) → D (atomic snapshot).

### Work item F — JobRegistry + JobStore + pull/poll + reconcile  ·  Size **M**  ·  Priority 6 (after E→A)

- **Goal:** Loop-side registry that **owns the Future**; a `JobStore` that persists RUNNING+finalizes and reconciles RUNNING→unknown; pull/poll contract; document the detached limit.
- **Scope (files):** `server.py`, `adapters/excel/com_worker.py` (reuse single `submit`), `storage/run_repository.py` (verify only; reuse `_atomic_write_json`), new `storage/job_store.py`, new `application/job_registry.py`.
- **Tasks:**
  1. `JobRegistry` (lock-guarded, non-COM) **wraps `submit()`** and **owns the Future** ⇒ timeout/disconnect never drops the result.
  2. `JobStore` (sibling of `RunRepository`; **not** TestRunResult-typed): persist **RUNNING marker at start**, **finalize** on completion; reuse extracted `_atomic_write_json` (`run_repository.py:113`) — atomic temp+fsync+`os.replace`.
  3. **Startup reconcile**: non-terminal (`RUNNING`) ⇒ **`unknown`** — never completed, never retried.
  4. `job_get(job_id)` returns status/result; `job_cancel` **marks intent only**; idempotency key on start.
  5. Long ops return `{status:"running", job_id, poll_after_s}` ≪60 s; progress best-effort.
  6. **DOCUMENT**: pull/poll + in-memory registry does **NOT** survive a client tree-kill/server restart; only JobStore-persisted results are retrievable after reconnect. **DEFER** detached broker (blocked by job breakaway).
- **Acceptance criteria:** start→job_id promptly; `job_get` transitions; timeout/reconnect/tree-kill behave per §6.4 of report_07; RUNNING reconciled to `unknown`; no second STA; `job_cancel` never interrupts COM.
- **Test strategy:** unit — registry transitions; JobStore atomic write + reconcile RUNNING→unknown. integration — start→poll→result; client timeout ⇒ result still retrievable by `job_id`; SIGKILL mid-op ⇒ job not retrievable (limit documented).
- **PoC gate:** **P5** (job lifecycle+transport; F build consumes E + A).
- **Dependencies:** E → A; B is the >idle evidence (not an ordering edge).

### Work item B — Pivot single-sweep + measured budget  ·  Size **M**  ·  Priority 4 (parallel with E)

- **Goal:** Single-sweep inventory with SourceType gate and a per-sheet **measured** budget; a slow single COM stays **running/slow**, never `too_large`/hung.
- **Scope (files):** `application/pivot_tables_service.py`, `adapters/excel/pivot_tables_adapter.py`.
- **Tasks:**
  1. Single-sweep + `CacheIndex` memoization; remove 2+N re-sweeps (`service:171-172,899-904,906-916`).
  2. **SourceType gate FIRST** → `unsupported_source` (`_UNSUPPORTED_SOURCE_TYPES` ~`:188-199`, `_require_supported_source` ~`:833-851`).
  3. Per-sheet **MEASURED** budget checked **BETWEEN completed COM ops** (never interrupts in-flight `:602`); expiry stops the **NEXT** op.
  4. Partial ⇒ `capped`; `too_large` distinct; current op stays `running`/`long_running` (coordinate F).
  5. **Never** key on file/row/UsedRange.
- **Acceptance criteria:** bounded `too_large` naming the sheet; `capped`≠`too_large`; external/OLAP `unsupported_source`; no materialization; no `RefreshAll`; slow single COM stays running.
- **Test strategy:** unit — budget between-ops; `capped` vs `too_large`. integration — 5.5 MB `report_w` (566k-row) returns bounded `too_large` naming the sheet.
- **PoC gate:** **P6** (no hard gate; informs F's >idle evidence).
- **Dependencies:** E (soft — reuse deadline vocabulary only; independent core proceeds in parallel).

### Work item C — Paths / docs  ·  Size **S**  ·  Priority 5 (anytime)

- **Goal:** Log resolved roots at startup; document footguns; keep `original_name` absent; `SCHEMA_VERSION=1`.
- **Scope (files):** `server.py`, `domain/models.py` (verify only), `api/workbook_tools.py` (verify only), `application/session_service.py` (verify only), `docs/`.
- **Tasks:**
  1. Log resolved roots at startup (replace silent fallback `server.py:765-774`).
  2. Docs: inspect-is-writable (`workbook_tools.py:86-87`), double-I/O (`session_service.py:873`).
  3. Confirm `original_name` **absent** (do **not** add; `source_path` exists `models.py:90`).
  4. Keep `SCHEMA_VERSION=1` (additive).
- **Acceptance criteria:** startup log shows resolved roots; `describe` uses `source_path`; docs state the footgun; `SCHEMA_VERSION=1`.
- **Test strategy:** unit — startup log emits roots; `WorkbookInfo` has no `original_name`. integration — none required.
- **PoC gate:** **—** (trivial; verified inline — startup log shows resolved roots, `source_path` present, no `original_name`; no dedicated PoC).
- **Dependencies:** none.

---

## 5. PoC gates (Stage 0, pass/fail)

All PoCs run **on copies**; originals hashed before/after and never modified; single-STA maintained; pid-scoped only.

| PoC | Block | Setup | Pass criteria | Fail ⇒ |
|-----|-------|-------|---------------|--------|
| **P1** | D | No Excel; unit torn-read harness (concurrent reader/writer) | `get_status` never exposes a torn `_active_*`; snapshot immutable per read; completion ≠ stale active | Block E |
| **P2** | A/E (+U5) | Copy of workbook; `background_safe` infinite-loop macro; `DispatchEx` child | Same-handle `TerminateProcess`+`WaitForSingleObject` returns; identity failure ⇒ **refuse** (fail closed); original hash unchanged | A/F blocked; handle unobtainable ⇒ server-restart fallback (documented) |
| **P3** | E | Slow-healthy vs blocked op; expired deadline | Slow-healthy **changes feedback** (`PROGRESSING`/`INCONCLUSIVE`); known progress vs inconclusive distinguished; expired deadline ≠ confirmed hang | Revisit E classification |
| **P4** | A/E | Modal `#32770` dialog blocks the STA | Non-COM `excel_health`/`excel_dialogs` reports `BLOCKED_ON_DIALOG` truthfully while COM is blocked; server stays non-COM-truthful; agent can inspect/recover | Revisit dialog detection |
| **P5** | F | Start long op; poll; timeout / cancel / graceful-disconnect / tree-kill separately | `start`→`job_id` promptly; poll works; timeout/cancel/disconnect/tree-kill distinguished; JobStore persist + reconcile RUNNING→unknown | Revisit U1 scope |
| **P6** | B | Large sheet (5.5 MB `report_w`, 566k rows) | Single-sweep; measured budget between-ops; slow single COM stays **running/slow** (not too_large/hung); **[LIVE-VERIFIED 2026-11-25]** a 5.5 MB copy wedged the single-STA >3.5 min and `pivot_tables_list` returned MCP -32001 | Revisit B guard |
| **P7** | A | Full loop on a copy (RESTRICTED) | Wait when progressing; recover when justified; verify termination + session; prior op `completed\|failed\|unknown`, no replay; `excel_recover` allowed, MUTATING denied | Fix policy/loop |

---

## 6. Open decisions (may alter F / A / E)

Record the sign-off; default lets work proceed; branch changes scope.

- **U1 — Tree-kill reload survival (product; affects F). RESOLVED (v1).**
  - **Decision (v1):** in-flight work is **not guaranteed to survive a Kilo tree-kill**; **completed results are persisted** where restart retrieval is required; interrupted jobs are reconciled to **`INTERRUPTED_UNKNOWN`** and **never auto-retried**.
  - **Default:** graceful reconnect + completed-job retrieval; **detached broker deferred**.
  - **Branch (yes):** a detached broker (execution lifecycle independent of the stdio connection, same interactive session) becomes an **F blocker** — re-plan F as **L** and add a detached-broker work item (note: blocked by job breakaway).
- **U2 — Auto-kill a LIVE owned pid at the deadline. RESOLVED (C1).**
  - **Decision:** **Never** auto-kill a live pid. Recovery is **agent-directed `excel_recover`** — **autonomous, no human approval**; the server **never auto-terminates a live pid on deadline/grace**. Dead-pid lazy recovery only. (No open branch; recorded for sign-off.)
- **U3 — Real Kilo timeout/progress. OPEN.**
  - Actual = `kilo serve` 7.8.32 / TS SDK 1.32.1, 60 s/call, `resetTimeoutOnProgress:true`, no `maxTotalTimeout`; Windows kill = direct-child ⇒ EXCEL orphaned. **P5 did NOT exercise progress/timeout-reset** ⇒ the polling contract **must not depend on progress or timeout-reset**. **OPEN — not verified; do not mark verified.**
- **U4:** confirm consumers treat added `WorkbookSessionState` members as additive v1 (keep `SCHEMA_VERSION=1`).
- **U5 — `PROCESS_TERMINATE` handle on the `DispatchEx` child. RESOLVED-POSITIVE [LIVE-VERIFIED 2026-11-25].**
  - P2 obtained `OpenProcess(PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION)` on the real `DispatchEx` child in-session; same-handle `TerminateProcess`+`WaitForSingleObject` = `WAIT_OBJECT_0`; identity-failure **refused** (fail closed).
- **E/F boundary — RESOLVED.** The **minimal job registry + prompt `job_id` response are implemented in Block E** (E is **not accepted with a non-functional polling contract**); F adds the full `JobStore` + reconcile + detached-limit documentation.
- **Live-verified premises (a)–(e) [LIVE-VERIFIED 2026-11-25]** (copies only; originals hashed before/after): (a) pivot cost NOT predicted by file/row size — a 5.5 MB / 566k-row `report_w` copy wedged the single-STA >3.5 min and `pivot_tables_list` returned MCP -32001; (b) health non-COM + truthful during a block (`server_health` ok instantly; `excel_health` `needs_recovery`/`active_operation`/`owned_pid`); (c) originals never modified (SHA-256 before/after identical; `excel_open` uses a working copy — auditor before/after hash; validator could not re-baseline); (d) single-STA (one owned pid 2856); (e) `_active_*` stale-but-consistent live, with the **[PoC-MODEL]** P1 model at 21.1M torn / 9.9M stale. **CategoryData 124.8 MB external/OLAP 1.3 GB cache = [STATIC]** code-verified `unsupported_source` only (P6 sub-case BLOCKED, not runtime).
- **Operational note [LIVE-VERIFIED 2026-11-25]:** `excel_recover` **does not exist** at rev `04437f5`; the live shared worker is currently **wedged in `needs_recovery`** — recovery unavailable ⇒ **server restart is the fallback** (U8).

---

## 7. Out of scope

- Detached broker / on-disk-job-independent-of-stdio (single-STA redesign; cannot save an in-flight call; blocked by job breakaway).
- MCP Tasks (absent in `mcp` 2.3.0; Kilo doesn't advertise).
- Heavyweight Operation Manager.
- Per-op workers / multiple STAs.
- VM / container execution.
- Wildcard kill (`taskkill /IM`, `Stop-Process -Name`) and any `taskkill` fallback.
- Any mutation of originals (originals are read-only; only working copies touched).
- Pivot output materialization; `Workbook.RefreshAll`.
- `CONFIRMED_HUNG` from a single signal.
- Silent flip of `excel_open` to `read_only=True` (contract change; deferred, tracked under C).

---

*End of plan_02. Plan only — no implementation code. Execute in DAG order D → E → A → F (B ∥ E, C anytime); pass each PoC gate before dependent blocks; record exact commands and PoC artifacts; keep originals hash-unchanged; server never auto-terminates a live pid; fail-closed recovery; no fixed-timeout auto-kill.*