# plan_01 — Rework Plan (Blocks A–F)

- **Purpose:** Standalone, executable rework plan derived from `reports/report_06 — Final Reconciled Architecture Decisions.md`. A Planner/Implementor with **no prior context** can execute it. It converts the reconciled decisions (D1–D8) into dependency-safe, independently reviewable work items with explicit acceptance criteria, test strategy, and PoC gates.
- **Revision:** `04437f5` (`04437f582ba4050469ecd99f298e19f64c814650`)
- **Date:** 2026-11-25
- **How to use:** Read §1 (non-negotiables) and §2 (decisions that must hold) first — they are binding. Then execute work items in **DAG order** (§4): **D → E → A → F**, with **B ∥ E** and **C anytime**. Do not start a dependent block until its **PoC gate** (§5) passes. Record the exact build/test commands and PoC artifacts. This is a **plan only** — it contains no implementation code.
- **Source of truth:** report_06. If this plan and report_06 ever disagree, report_06 wins; raise the discrepancy, do not silently diverge.

---

## 1. Non-negotiables (binding on every work item)

1. **Single-STA:** exactly one hidden Excel instance in the interactive user session. No VMs, no containers, no per-op workers, no second STA.
2. **Originals never modified.** All work happens on working copies. Originals are hashed before/after every PoC and must be byte-identical.
3. **pid-scoped kill only.** Never a wildcard image kill (`taskkill /IM`, `Stop-Process -Name`). Blast radius is exactly one owned pid, confirmed via a held handle (Block A).
4. **No pivot output materialization.** Never expand pivot output to cells; never open `pivotCacheRecords`.
5. **No `Workbook.RefreshAll`.** Cache refresh is table-scoped/cache-scoped only.
6. **`SCHEMA_VERSION=1`, additive only.** New enum members / fields are additive; no breaking DTO change.
7. **Deny-by-default for genuinely destructive admin ops** — **except** the new originals-safe **recovery** class (`excel_recover`), which is allow-by-default in FULL **and** RESTRICTED (audited, pid-scoped, discards only working copies). A **mutating** op stays denied in RESTRICTED.
8. **English only** in code, comments, logs, docs.
9. **Do not modify** report_02/03/04/05, the validity file, production code, or tests **as part of planning** — this plan authorizes the *future* implementation work items only.

### Definition of done (whole plan)

- D, E, A, F implemented and independently reviewed; B and C implemented.
- All PoC gates P1–P7 pass (or their branch is explicitly accepted per §6).
- Unit + integration tests added/updated for each block; originals hash-unchanged across all PoCs.
- U1/U2 resolved with recorded sign-off (default accepted or branch chosen).
- Docs updated (Block C) and `SCHEMA_VERSION` still `1`.

---

## 2. Architectural decisions (must hold)

These are the FINAL decisions from report_06. Implementations must satisfy them verbatim.

### 2.1 State set (rename per validator)

`RUNNING` / `SLOW` / `DEADLINE_EXCEEDED` / `BLOCKED_BY_DIALOG` / `PROCESS_EXITED` / `CONFIRMED_HUNG`.

- **`DEADLINE_EXCEEDED`** is the honest state at the server deadline: the call is still in-flight and **unprovable**.
- **`CONFIRMED_HUNG` is NEVER asserted from `WaitForSingleObject` alone** (that proves **EXIT**) **or `#32770` alone** (that proves **DIALOG**). It requires **post-hoc evidence** (owned pid alive AND no completion after a grace window) **or** an explicit operator/agent decision.
- `BLOCKED_BY_DIALOG` = a `#32770` modal is present (informational, not proof of hang).
- `PROCESS_EXITED` = `WaitForSingleObject(handle)` returned `WAIT_OBJECT_0` (the process is gone).

### 2.2 Timeout / liveness contract (Block E — keystone)

- **ONE mechanism for ALL COM ops:** `excel_open`, `pivot_tables_list`, `pivot_table_get`, `pivot_table_refresh`, `vba_compile`, `macro_run`.
- **Queue-wait vs exec deadline are split:** queue-wait returns `busy` / `retry_after_s` and **NEVER** sets `NEEDS_RECOVERY`; the exec deadline may set `DEADLINE_EXCEEDED`.
- **Per-op-class server deadline is the authority.** Env knobs: `EXCEL_VBA_MCP_OPEN_TIMEOUT_S=300`, `EXCEL_VBA_MCP_LIVENESS_CEILING_S=600`.
- **Per-tool `timeout_s` = SERVER-SIDE SOFT HINT ONLY** (inert under the Kilo legacy 60 s wall-clock). It is **never** a client contract.
- **Long ops MUST use PULL/POLL:** return `{status:"running", job_id, poll_after_s}` in **≪60 s**. Progress notifications are best-effort keep-alive, **never** a completion signal.
- **Remove** the separate macro 90 s auto-kill (`macro_service.py:107,866`). Replace with the unified deadline driving an **automatic narrow reap for EXECUTING** (live pid included) after a grace window — **OR** (U2 branch) agent-only with documented bricking risk surfaced via `excel_health`.

### 2.3 Block decisions (A–F)

- **D — Atomic runtime state (ship first):** single lock-guarded **immutable** status snapshot. `get_status` must not read `_active_*` lock-free while the worker and loop write them. Clear-before-complete. Keep `list_workbooks` non-COM; a `list_workbooks_reconciled`, if built, goes through the single STA (`submit`) or returns `busy` — never an arbitrary thread.
- **E — Unified timeout/liveness:** see §2.2.
- **A — Recovery (consumes D+E):** off-STA, pid-scoped. Hold a process **HANDLE** at launch (`PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION`); terminate+confirm via the **same handle**: `TerminateProcess(handle)` → `WaitForSingleObject(handle)` (replaces `taskkill /F /PID`). `Hwnd=0`/ambiguous ⇒ **NO auto-kill**. Add `started` to `_WorkItem`. Add allow-by-default `excel_recover` (D5). Interrupted mutation = **UNKNOWN, never auto-retry**. Snapshot restore = **file-consistent only**; external macro side effects not undone.
- **B — Pivot cost guard (∥ E):** single-sweep + `CacheIndex` memoization + **SourceType gate FIRST**; per-sheet **MEASURED** budget checked **BETWEEN completed COM ops** (never interrupts in-flight); expiry stops the **NEXT** op; partial ⇒ `capped`; `too_large` distinct; current op stays `running`/`long_running` (coordinate F). **Never** key on file/row/UsedRange. Remove the 2+N re-sweeps.
- **C — Paths/docs (independent):** `original_name` is **absent** (do not add; `source_path` exists). Log resolved roots at startup. Docs: inspect-is-writable, double-I/O. `SCHEMA_VERSION=1`.
- **F — Minimal op seam (after E→A):** in-process job registry (`job_id→{status,result?,error?,poll_after_s,created_at,deadline_at}`) + `job_get(job_id)` + optional on-disk result persistence + best-effort progress; runs on E's deadline + the **SAME single STA**. `job_cancel` only marks intent (cannot interrupt COM). **DEFER** detached worker / on-disk-job-independent-of-stdio; **DOCUMENT** that pull/poll does **not** survive a client tree-kill/server death (only already-persisted results are retrievable after reconnect). Heavyweight Operation Manager + MCP Tasks deferred.

---

## 3. Critical code map

All paths under `src/excel_vba_mcp/`. Anchors re-derived @`04437f5` (line numbers are hints; prefer the named symbol).

### Block D — `adapters/excel/com_worker.py`
- `_WorkItem` dataclass (`:133-147`) — add `started` is **A**'s concern; D only reads/records.
- `submit()` end-to-end `asyncio.wait_for(...)` (`:391`); `TimeoutError`→`NEEDS_RECOVERY` (`:395-396`); active fields set (`:397-399`).
- `_process_item()` — pre-start deadline (`:602`); active writes (`:625-627`); `_complete_result` (`:632`) **before** `finally: _clear_active()` (`:639-640`).
- `get_status()` (`:411-426`) — lock (`:414-416`) covers only `state`/`accepting`; `_active_*` read lock-free (`:420-421`).
- `application/session_service.py` — `list_workbooks` / `list_workbooks_sync` (`:527-541`, non-COM).
- `api/health_tools.py` — `excel_health` (non-COM).

### Block E — `adapters/excel/com_worker.py`, `application/macro_service.py`, `application/session_service.py`, `adapters/excel/app_adapter.py`, `config.py`
- `com_worker.py` `submit()` (`:391`) + `_process_item` pre-start (`:602`).
- `macro_service.py` `_DEFAULT_TIMEOUT_S=90.0` (`:107`); `_terminate_owned_pid` at `:866` (the 90 s kill to remove).
- `session_service.py` `_maybe_auto_recover` (`:729-757`, dead-pid-only guard `:754`).
- `app_adapter.py` reap loop (`:185-202`), `_terminate_pid` (`:470-491`).
- `config.py` — add `open_timeout_s` / `liveness_ceiling_s` env knobs.

### Block A — `adapters/excel/com_worker.py`, `adapters/excel/app_adapter.py`, `application/session_service.py`, `application/snapshot_service.py`, `security/approval.py`, `security/macro_policy.py`, `api/health_tools.py`
- `com_worker.py` `_WorkItem` (`:133-147`, add `started`); interrupted-mutation path (`:388-409`).
- `app_adapter.py` `_terminate_pid` (`:470-491` → same-handle `TerminateProcess`+`WaitForSingleObject`); reap (`:185-202`).
- `session_service.py` `_maybe_auto_recover` (`:729-757`).
- `snapshot_service.py` restore contract (`:18-22`).
- `security/approval.py` + `security/macro_policy.py` — new allow-by-default **recovery** class.
- `api/health_tools.py` — `excel_recover` tool surface.

### Block B — `application/pivot_tables_service.py`, `adapters/excel/pivot_tables_adapter.py`
- `pivot_tables_service.py` `list_pivot_tables` (`:171-172`), `_collect_caches` (`:899-904`), `_find_pivot_table` (`:906-916`), `_DEFAULT_TIMEOUT_S=60.0` (`:98`).
- `pivot_tables_adapter.py` `enumerate_pivot_tables`, `read_pivot_cache`, `_describe_cache`; `_UNSUPPORTED_SOURCE_TYPES` (~`:188-199`); `_require_supported_source` (~`:833-851`).

### Block C — `server.py`, `domain/models.py`, `api/workbook_tools.py`, `application/session_service.py`, docs
- `server.py` root resolution (`:765-774`) — add startup log of resolved roots.
- `domain/models.py` `WorkbookInfo` (`:83-96`, `source_path` `:90`) — **no `original_name`**.
- `api/workbook_tools.py` `excel_open` defaults (`:86-87`).
- `session_service.py` `_materialize_working_copy` double I/O (`:867-874`).
- `docs/` — inspect-is-writable + double-I/O notes.

### Block F — `server.py`, `adapters/excel/com_worker.py`, `storage/run_repository.py`, `storage/snapshot_store.py`, new in-process registry module
- `server.py` tool surface — add `job_get` / `job_cancel`; long ops return `{status:"running", job_id, poll_after_s}`.
- `com_worker.py` single `submit` (reuse; **no** second STA).
- `storage/run_repository.py`, `storage/snapshot_store.py` — reuse for optional on-disk result persistence (completed artifacts only).
- New module (e.g. `application/job_registry.py`) — in-process `job_id` map.

---

## 4. Sequenced work items (DAG order)

> Order: **D → E → A → F**; **B ∥ E**; **C anytime**. Each item is atomic and independently reviewable. Size: **S/M/L**.

### Work item D — Atomic runtime state  ·  Size **S**  ·  Priority 1 (ship first)

- **Goal:** Make `get_status` truthful and race-free; establish the atomic snapshot every later block observes.
- **Scope (files):** `adapters/excel/com_worker.py`, `application/session_service.py` (read-only use), `api/health_tools.py` (verify non-COM).
- **Tasks:**
  1. Introduce a single lock-guarded **immutable** status snapshot; `get_status` reads `_active_*` **only** from that snapshot (never lock-free).
  2. Enforce **clear-before-complete**: `_complete_result` (`:632`) publishes the result, then `finally: _clear_active()` (`:639-640`) clears active — the snapshot must never expose a stale active op after completion.
  3. Keep `list_workbooks` / `list_workbooks_sync` **non-COM** (`session_service.py:527-541`).
  4. If `list_workbooks_reconciled` is built, route it through the single STA (`submit`) or return `busy`; **never** an arbitrary thread.
- **Acceptance criteria:** no torn read of `_active_*`; `excel_health` accurate while busy; `list_workbooks` stays non-COM; reconciled variant (if present) never spawns a thread.
- **Test strategy:** unit — concurrent reader/writer on the snapshot (no torn state); integration — `excel_health` during a slow op shows the correct active op and clears after completion.
- **PoC gate:** **P1** must pass before E.
- **Dependencies:** none.

### Work item E — Unified timeout / liveness (keystone)  ·  Size **M**  ·  Priority 2 (foundation)

- **Goal:** One deadline mechanism for all COM ops; kill the separate 90 s macro kill; split queue-wait from exec; introduce pull/poll contract.
- **Scope (files):** `adapters/excel/com_worker.py`, `application/macro_service.py`, `application/session_service.py`, `adapters/excel/app_adapter.py`, `config.py`.
- **Tasks:**
  1. Split queue-wait (returns `busy`/`retry_after_s`, **never** `NEEDS_RECOVERY`) from exec deadline (may set `DEADLINE_EXCEEDED`).
  2. Remove `_DEFAULT_TIMEOUT_S=90.0` auto-kill (`macro_service.py:107,866`); route `macro_run` through the unified per-op-class deadline.
  3. Add env knobs `EXCEL_VBA_MCP_OPEN_TIMEOUT_S=300`, `EXCEL_VBA_MCP_LIVENESS_CEILING_S=600` (`config.py`).
  4. Define per-tool `timeout_s` as a **server-side soft hint** (no client contract).
  5. Long ops return `{status:"running", job_id, poll_after_s}` in ≪60 s (contract shared with F).
  6. Implement the U2 default: automatic **narrow EXECUTING reap** after grace; if U2 = agent-only, document bricking risk and surface `DEADLINE_EXCEEDED` via `excel_health`.
- **Acceptance criteria:** slow-but-alive ⇒ `busy`/`retry_after_s`, no poison; exec breach ⇒ `DEADLINE_EXCEEDED` (not hung); env knobs honored; native call never cancelled; no 90 s kill remains.
- **Test strategy:** unit — queue-wait timeout does not set `NEEDS_RECOVERY`; exec deadline sets `DEADLINE_EXCEEDED`. integration — a deliberately slow op returns `busy`; a blocked op yields `DEADLINE_EXCEEDED`; env override changes the ceiling.
- **PoC gate:** **P2, P3** must pass before A and F.
- **Dependencies:** D.

### Work item A — Recovery  ·  Size **M**  ·  Priority 3

- **Goal:** Off-STA, pid-scoped recovery using a held handle; honest state classification; allow-by-default `excel_recover`.
- **Scope (files):** `adapters/excel/com_worker.py`, `adapters/excel/app_adapter.py`, `application/session_service.py`, `application/snapshot_service.py`, `security/approval.py`, `security/macro_policy.py`, `api/health_tools.py`.
- **Tasks:**
  1. Add `started` flag to `_WorkItem` (`:133-147`) so the awaiter distinguishes "never started" from "started + slow".
  2. Hold a process **HANDLE** at launch (`PROCESS_TERMINATE|SYNCHRONIZE|QUERY_LIMITED_INFORMATION`); terminate+confirm via the **same handle** (`TerminateProcess` → `WaitForSingleObject`), replacing `taskkill /F /PID` (`app_adapter.py:470-491`). Fallback to pid+`taskkill` only if U5/P2 fails (documented, weaker).
  3. `Hwnd=0`/ambiguous identity ⇒ **NO auto-kill** (operator/agent action).
  4. Add allow-by-default **`excel_recover`** = originals-safe liveness class (auto-allowed FULL **and** RESTRICTED, audited, pid-scoped, discards only working copies). Mutating op stays denied in RESTRICTED.
  5. Interrupted mutation = **UNKNOWN, never auto-retry** (`com_worker.py:388-409`). Snapshot restore = file-consistent only; external macro side effects not undone (`snapshot_service.py:18-22`).
  6. Server auto-reap narrow (dead-pid + confirmed deadline for EXECUTING per U2).
- **Acceptance criteria:** `DEADLINE_EXCEEDED` at deadline, `CONFIRMED_HUNG` only post-hoc/operator; `excel_recover` works in RESTRICTED and a mutating op does not; originals hash unchanged; no wildcard kill; `Hwnd=0` ⇒ no auto-kill.
- **Test strategy:** unit — `started` flag transitions; handle-scoped terminate/confirm. integration — recovery on a working copy with a background_safe infinite-loop macro; policy test for allow-by-default RECOVERY vs denied MUTATING in RESTRICTED.
- **PoC gate:** **P2, P7** must pass before F.
- **Dependencies:** E (classification) → D (atomic snapshot).

### Work item F — Minimal op seam  ·  Size **S now / L later**  ·  Priority 6 (after E→A)

- **Goal:** In-process job registry + `job_get` + optional on-disk results + best-effort progress, on E's deadline and the same single STA; document the detached limit.
- **Scope (files):** `server.py`, `adapters/excel/com_worker.py` (reuse single `submit`), `storage/run_repository.py`, `storage/snapshot_store.py`, new `application/job_registry.py`.
- **Tasks:**
  1. In-process registry `job_id→{status,result?,error?,poll_after_s,created_at,deadline_at}`.
  2. `job_get(job_id)` returns status/result; `job_cancel` only **marks intent** (cannot interrupt COM).
  3. Optional on-disk result persistence (completed artifacts only) via existing stores.
  4. Long ops return `{status:"running", job_id, poll_after_s}` ≪60 s; progress is best-effort keep-alive.
  5. **DOCUMENT**: pull/poll + in-process registry does **NOT** survive a client tree-kill/server death; only already-persisted results are retrievable after reconnect.
  6. **DEFER** detached worker / on-disk-job-independent-of-stdio; heavyweight Operation Manager + MCP Tasks.
- **Acceptance criteria:** open/save/test_suite_run/macro_run succeed under post-E budget + progress; op seam start→status→result works; disconnect ⇒ `lost`/`interrupted_unknown`; no second STA.
- **Test strategy:** unit — registry transitions; integration — start→poll→result; SIGKILL mid-op ⇒ job not retrievable (documents limit).
- **PoC gate:** **P5** (documents detached limit; F build consumes E + A).
- **Dependencies:** E → A; B is the >idle evidence (not an ordering edge).

### Work item B — Pivot cost guard  ·  Size **M**  ·  Priority 4 (parallel with E)

- **Goal:** Single-sweep inventory with SourceType gate and a per-sheet measured budget; distinct `capped`/`too_large`.
- **Scope (files):** `application/pivot_tables_service.py`, `adapters/excel/pivot_tables_adapter.py`.
- **Tasks:**
  1. Single-sweep + `CacheIndex` memoization; remove 2+N re-sweeps (`service:171-172,899-904,906-916`; adapter `_describe_cache`).
  2. **SourceType gate FIRST** → `unsupported_source` for external/OLAP/Data-Model (`_UNSUPPORTED_SOURCE_TYPES` ~`:188-199`, `_require_supported_source` ~`:833-851`).
  3. Per-sheet **MEASURED** budget checked **BETWEEN completed COM ops** (never interrupts in-flight `:602`); expiry stops the **NEXT** op.
  4. Partial ⇒ `capped`; `too_large` = distinct policy result; current op stays `running`/`long_running` (coordinate F).
  5. **Never** key on file/row/UsedRange.
- **Acceptance criteria:** bounded `too_large` naming the sheet within budget; `capped`≠`too_large`; external/OLAP `unsupported_source`; no materialization; no `RefreshAll`.
- **Test strategy:** unit — budget between-ops; `capped` vs `too_large`. integration — 5.5 MB `report_w` (566k-row) returns bounded `too_large` naming the sheet; no UsedRange/row/file keying.
- **PoC gate:** **P4** (no hard gate; informs F's >idle evidence).
- **Dependencies:** E (soft — reuse deadline vocabulary only; independent core proceeds in parallel).

### Work item C — Paths / docs  ·  Size **S**  ·  Priority 5 (anytime)

- **Goal:** Log resolved roots at startup; document footguns; confirm `original_name` stays absent; keep `SCHEMA_VERSION=1`.
- **Scope (files):** `server.py`, `domain/models.py` (verify only), `api/workbook_tools.py` (verify only), `application/session_service.py` (verify only), `docs/`.
- **Tasks:**
  1. Log resolved roots at startup (replace silent fallback `server.py:765-774`).
  2. Docs: inspect-is-writable (`workbook_tools.py:86-87`), double-I/O (`session_service.py:873`).
  3. Confirm `original_name` **absent** (do **not** add; `source_path` exists `models.py:90`).
  4. Keep `SCHEMA_VERSION=1` (additive).
- **Acceptance criteria:** startup log shows resolved roots; `describe` uses `source_path`; docs state the footgun; `SCHEMA_VERSION=1`.
- **Test strategy:** unit — startup log emits roots; `WorkbookInfo` has no `original_name`. integration — none required.
- **PoC gate:** **P6** (trivial).
- **Dependencies:** none.

---

## 5. PoC gates (pass/fail)

All PoCs run **on copies**; originals hashed before/after and never modified; single-STA maintained; pid-scoped only.

| PoC | Block | Setup | Pass criteria | Fail ⇒ |
|-----|-------|-------|---------------|--------|
| **P1** | D | No Excel; unit torn-read harness (concurrent reader/writer) | `get_status` never exposes a torn `_active_*`; snapshot immutable per read | Block E |
| **P2** | A/E (+U5) | Copy of workbook; `background_safe` infinite-loop macro | State = `DEADLINE_EXCEEDED` (not CONFIRMED); owned pid alive; `excel_recover` → `TerminateProcess(handle)`+`WaitForSingleObject(handle)` returns; original hash unchanged | A/F blocked; if handle unobtainable ⇒ U5 fallback to pid+`taskkill` (documented) |
| **P3** | E (U2 branch) | No agent recover; slow/blocked op | Stays `DEADLINE_EXCEEDED`; live pid **NOT** auto-killed (gap documented) | Revisit U2 default |
| **P4** | B | Large sheet (5.5 MB `report_w`) | Budget checked between-ops; `capped` vs `too_large` distinct; no UsedRange/row/file keying; current op stays running | Revisit B guard |
| **P5** | F | Start long op; poll; SIGKILL mid-op | not-ready ≪60 s; `job_get` transitions; after SIGKILL job **NOT** retrievable (detached limit documented) | Revisit U1 scope |
| **P6** | C | No roots env | Startup log shows resolved roots; `source_path` present; no `original_name` | Fix C logging |
| **P7** | A | RESTRICTED mode | `excel_recover` (RECOVERY) allowed by default; MUTATING still denied | Fix policy class |

---

## 6. Open decisions (may alter F / A / E)

Record the sign-off; default lets work proceed; branch changes scope.

- **U1 — Tree-kill reload survival (product; affects F).**
  - **Default:** scope to graceful reconnect + completed-job retrieval; **detached worker deferred**.
  - **Branch (yes):** a detached worker (execution lifecycle independent of the stdio connection, same interactive session) becomes an **F blocker** — re-plan F as **L** and add a detached-worker work item.
- **U2 — Auto-kill a LIVE owned pid at the unified deadline for EXECUTING (risk; affects E/A).**
  - **Default:** automatic **narrow** reap for EXECUTING after grace.
  - **Branch (agent-only):** `DEADLINE_EXCEEDED` + agent `excel_recover` + documented bricking risk surfaced via `excel_health`; E task 6 and A task 6 change accordingly.
- **U3 (verify by PoC):** which Kilo backend (legacy 60 s vs `kilo serve` 12 h) and does it honor progress/reset — fixes `poll_after_s`/long-poll ≤25 s.
- **U4:** confirm consumers treat added `WorkbookSessionState` members as additive v1 (keep `SCHEMA_VERSION=1`).
- **U5 (P2):** `PROCESS_TERMINATE` handle obtainable on the `DispatchEx` child; else fallback to pid+`taskkill` (documented, weaker).

---

## 7. Out of scope

- Detached worker / on-disk-job-independent-of-stdio (single-STA redesign; cannot save an in-flight call).
- MCP Tasks (absent in `mcp` 2.3.0; Kilo doesn't advertise).
- Heavyweight Operation Manager.
- Per-op workers / multiple STAs.
- VM / container execution.
- Wildcard kill (`taskkill /IM`, `Stop-Process -Name`).
- Any mutation of originals (originals are read-only; only working copies are touched).
- Pivot output materialization; `Workbook.RefreshAll`.
- Silent flip of `excel_open` to `read_only=True` (contract change; deferred, tracked under C).

---

*End of plan_01. Plan only — no implementation code. Execute in DAG order D → E → A → F (B ∥ E, C anytime); pass each PoC gate before dependent blocks; record exact commands and PoC artifacts; keep originals hash-unchanged.*