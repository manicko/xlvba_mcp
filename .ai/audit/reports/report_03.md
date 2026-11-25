# Architecture Audit Report 03 — Long-Running Operations, Worker Recovery, and the "Slow ≠ Hung" Reframe

- **Status:** Recommended, subject to PoC (see §9). Advisory; no production code is changed by this report.
- **Date:** 2026-11-25
- **Author:** Architecture auditor (read-only, evidence-driven)
- **Scope:** The `xlvba_mcp` MCP server's single-STA COM worker, session lifecycle, pivot enumeration, path/allowlist handling, and diagnostics — evaluated against (a) the agent-usage audit `.ai/audit/problems/04_xlvba-usage-audit-agents-2026-11-25.md`, (b) a real-data verification on `.ai/tmp/theiss`, and (c) a research round on long-running-operation best practices.
- **Method:** Static code reading with `file:line` citations, plus read-only in-memory `zipfile` inspection of the real workbooks (no extraction, no Excel launch). All file sizes below were re-confirmed on disk.

---

## 2. Executive Summary

The agent-usage audit reported a "worker hang" that required an out-of-band `Stop-Process`. Real-data verification shows **this is not a deadlock**. Excel is legitimately spending a long time materialising very large pivot caches (a single `pivotCacheRecords1.xml` inflates to ~1.3 GB; `recordCount` 1,251,228). The single 60 s soft ceiling (`session_service.py:68` `_DEFAULT_TIMEOUT_S = 60.0`, applied to the `open` section at `session_service.py:329-334`) fires while Excel is **alive but slow**, the worker is moved to `NEEDS_RECOVERY` (`com_worker.py:390-409`), and the only auto-recovery self-heals a **dead** pid (`session_service.py:752-755`) — so a live-but-slow call stays poisoned until a human kills the process.

**The core rule this report recommends:**

> **An MCP request timeout must never be interpreted as proof that Excel is hung, nor as a reason to recover or terminate the worker.** A slow-but-alive operation must be reported as `SLOW`, not `NEEDS_RECOVERY`. Only a **confirmed** loss of liveness (owned pid dead, or no progress past a liveness ceiling) may trigger destructive recovery.

The fix is **not** to abandon the single-STA channel (ARCH-009 is non-negotiable and correct). The fix is to **separate two concepts the code currently conflates into one number**:

- **Deadline** — how long an operation of a given class may *legitimately* take (per-op-class budget `B`).
- **Liveness** — whether the worker is still *progressing* (a free-thread `pid`-alive / progress check, ceiling `C`).

Everything else (state machine, off-thread `pid`-scoped reap, a job-seam for genuinely long operations) follows from that one separation. The single-STA design, the copy-to-`work/<GUID>` originals-never-modified invariant, deny-by-default approval, and the "never materialise pivot output / never `Workbook.RefreshAll`" rules all remain intact.

---

## 3. Problem Statement

The usage audit grouped the pain points into four blocks. Each is restated here with verified `file:line` evidence and the audit's finding IDs (P1–P7).

### Block A — Worker hang / single-STA head-of-line blocking / no in-tool recovery (P2, P3, P4)

- **One FIFO queue + one STA thread.** `com_worker.py:267` declares a single `queue.SimpleQueue`; the worker is a single `threading.Thread(..., name="excel-com-sta")` (`com_worker.py:300`). A slow/hung COM call therefore blocks every later call (head-of-line blocking).
- **A timeout poisons the worker.** `submit()` wraps the future in `asyncio.wait_for(..., timeout=timeout)` (`com_worker.py:391`); on `TimeoutError` it sets `WorkbookSessionState.NEEDS_RECOVERY` (`com_worker.py:395-396`) and raises `ExcelHungError` (`com_worker.py:401-409`). The still-running native call is **not** cancelled (correct — an uninterruptible COM call cannot be), and mutating/executing items are never replayed.
- **`close_workbook` queues behind the hang.** `close_workbook` submits a normal `section="close"` work item through the same FIFO (`session_service.py:381-386`); it is **not** the internal close-drain (`is_close_drain`, `com_worker.py:145-147`, produced only by `_make_internal_item`, `com_worker.py:715-726`). So it waits behind the stuck call and its own `wait_for` times out (P4).
- **No in-tool recovery.** The registered tool set (`server.py:1004-1165+`) has no "restart worker" tool. Recovery was manual via OS `Stop-Process` on `owned_pid` (problem doc lines 96, 147).

### Block B — Pivot enumeration/metadata performance + no bounded reads on large workbooks (P2, P3)

- **Full worksheet sweep + heavy per-table describe, repeated.** `PivotTablesService` enumerates per worksheet via `ws.PivotTables()` because `Workbook.PivotTables` is unusable on the target build (`pivot_tables_service.py:36, 158-161`). Each table's fields/cache are described per call.
- **60 s timeout → `NEEDS_RECOVERY` poisons the worker.** The pivot service uses the same `_DEFAULT_TIMEOUT_S = 60.0` (`pivot_tables_service.py:98`); on a large workbook this fires and poisons the single worker (P2/P3).
- **No bounded "too large" result.** There is no per-sheet cost guard that returns a bounded `too_large` answer; the call fails hard and takes the session with it (problem doc line 119).

### Block C — Path allowlist ergonomics, working-copy GUID indirection, double I/O, inspect-is-writable footgun (P1; sections 5/6)

- **Allowlist is CWD-dependent.** `_configured_roots()` falls back to `[str(Path.cwd()), tempfile.gettempdir()]` when `EXCEL_VBA_MCP_ROOTS` is unset (`server.py:765-770`). The same file is "inside" when launched from `C:\py_exp\xlvba_mcp` and "outside" when launched elsewhere.
- **GUID indirection.** `open_workbook` mints `workbook_id = str(uuid4())` (`session_service.py:246`) and the working copy is `f"{workbook_id}{suffix}"` (`session_service.py:872`); `WorkbookAdapter.describe` sets `name = wb.Name` (the GUID) and keeps the original only in `source_path` (`workbook_adapter.py:97-98, 108-119`).
- **Double I/O.** `_materialize_working_copy` does `shutil.copy2(canonical, working)` (`session_service.py:867-874`) — a full second copy of the original into `work/`.
- **`inspect` is writable.** `excel_open` defaults to `read_only=False` (`workbook_tools.py:84-88`), so `mode="inspect"` opens a **writable** copy; the approval class is `MUTATING` for a writable inspect (`session_service.py:211-216`). "Inspect ≠ read-only" is a footgun (problem doc line 121).

### Block D — Diagnostics / bookkeeping hygiene (P5, P6, P7)

- **P5 noisy historical notes.** `_note` appends to a rolling 50-entry list and is never cleared between logical sessions (`com_worker.py:854-859`); `excel_health` copies the whole list verbatim (`health_tools.py:139, 166`). Historical incidents (`submit timeout ... token=9`, `drain_close complete`, `reset rebuilt owned instance`) appear from the first call.
- **P6 transient `busy:true` race.** In `_process_item`, the result future is completed (`_complete_result`, `com_worker.py:632`) **before** `finally: self._clear_active()` (`com_worker.py:640`). Awaiting callers resume and read `active_operation` while it is still set, producing a transient `busy:true` on the first close after a large workbook.
- **P7 stale `list_workbooks`.** `list_workbooks` / `list_workbooks_sync` return bookkeeping only (`session_service.py:527-541`), with no reconciliation against the owned pid, so they can list workbooks whose worker has already been killed.

---

## 4. Real-Data Verification

All facts below were re-confirmed by read-only inspection of `.ai/tmp/theiss` (in-memory `zipfile`, no extraction). File sizes on disk match the audit exactly.

### 4.1 Measured per-workbook table

| Workbook | Size (MB) | Huge sheet (dimension) | Pivots | Cache source / `recordCount` |
|---|---|---|---|---|
| `base/CategoryData.xlsx` | **124.81** | large data sheet | 1 (on small `clean_data_full`) | `external` (Power Query `Query - clean_data_full`, `background=1`, `keepAlive=1`), `connectionId=2` / **1,251,228** |
| `base/report_wclip.xlsx` | **123.59** | large data sheet | 1 | `external`, same ~1.3 GB cache / **1,251,228** |
| `weekly/report_w.xlsx` | **5.51** | `sheet3` `dashboard_data` = `A1:B566386` (566,386 rows, ~30.79 MB) | 1 (on `clean_data_full`, e.g. `A7:FI51`) | `external` / **51,760** |
| `monthly/report_m.xlsx` | **5.61** | `sheet3` `dashboard_data` = `A1:B566386` (**identical**) | 1 | `external` / **51,760** |
| `treport.xlsm` | **0.33** | none | **0** | **0 caches** |

Key structural facts:

- `pivotCacheRecords1.xml` inflates to **~1.3 GB** in the 124 MB base files; `recordCount` **1,251,228**.
- The single pivot table lives on the **small** sheet `clean_data_full`; the **566k-row `dashboard_data`** sheet carries **no pivot** and is **not a ListObject** (so `table_read_rows` cannot address it — only `range_read_values` can).
- `treport.xlsm` (the tracked working workbook) has **0 pivots / 0 caches** — the hazard is entirely in the data workbooks, not the template.

### 4.2 Corrected claims (what the audit got wrong, and the accurate model)

1. **File size is an UNRELIABLE predictor.** The 5.5 MB `report_w.xlsx` carries the same 566k-row hazard as the 124 MB `CategoryData.xlsx`. **The hazard is PER-SHEET (sheet dimension / row count), not per-file.** Any cost guard keyed on file size will both miss the small-file hazard and over-trigger on large-but-benign files.
2. **The 60 s risk is the OPEN, not describe.** The binding deadline is `_DEFAULT_TIMEOUT_S = 60.0` (`session_service.py:68`) applied to the `open` submission (`section="open"`, `session_service.py:329-334`). Excel is alive (`owned pid` alive) but slow → `submit` `TimeoutError` → `NEEDS_RECOVERY` (`com_worker.py:391-409`). `_maybe_auto_recover` only self-heals a **dead** pid (`session_service.py:752-755`), so a live-but-slow open stays poisoned.
3. **The real caches are `external`.** `cacheSource type="external"` (Power Query). On the target build, external / OLAP / Data Model sources return `unsupported_source` for refresh/config (`pivot_tables_service.py:29-31, 51-52`; `server.py:607-699`). So even a *successful* pivot describe cannot be followed by a refresh — the pivot surface is read-only for these files by design.
4. **`table_read_rows` is not usable on `dashboard_data`.** It is not a ListObject; reading it requires `range_read_values`. At the `MAX_RANGE_CELLS = 100_000` cap (`models.py:666`) with 2 columns, each window is `floor(100000/2) = 50,000` rows, so `ceil(566386/50000) = 12` windows are needed to page the whole sheet (`ranges_adapter.py:3710-3735`).
5. **Allowlist is CWD-dependent.** The `theiss` files sit **inside** the default allowed roots only because CWD = `C:\py_exp\xlvba_mcp` (`server.py:769`). Launched elsewhere they become `OUTSIDE_ROOTS` (P1).
6. **Double I/O is real and large.** `copy2` of a 124.8 MB file into `work/<GUID>` (`session_service.py:867-874`) plus ~1.3 GB of pivot cache in RAM.
7. **GUID indirection confirmed.** `workbook_adapter.describe` returns `name = wb.Name = <GUID>`; the original path survives only in `source_path` (`workbook_adapter.py:97-98`).
8. **`excel_open` default `read_only=False` makes `mode="inspect"` writable** (`workbook_tools.py:84-88` vs `session_service.open_workbook`).

---

## 5. Root-Cause Analysis

### 5.1 Deadline-vs-liveness conflation (the primary root cause)

`submit()` builds a single `deadline = time.monotonic() + timeout` (`com_worker.py:381`) and uses the **same** `timeout` for both the queue-wait (`asyncio.wait_for`, `com_worker.py:391`) and the worker-side execution deadline (`com_worker.py:602`). One number is asked to answer two different questions — "how long may this legitimately run?" and "is the worker still alive?" — and it answers both with "kill the session." A legitimately long open is therefore indistinguishable from a hang.

### 5.2 Uncancellable native call

A COM call in flight cannot be cancelled from another thread. The code is correct to *not* cancel (`com_worker.py:36-40, 392-409`); the error is treating the resulting `TimeoutError` as proof of a hang. The call may simply be making progress.

### 5.3 `reset()` cannot preempt a live call

`reset()` enqueues `_do_reset` as a `section="close"` item on the **same** worker (`com_worker.py:489-507`, body `_do_reset` at `767-784`). Because the queue is FIFO and the worker is single-threaded, a `reset` submitted while a live call is running **queues behind it** and cannot preempt it. Recovery must therefore run **off-thread** (pid-scoped), not through the worker's own queue.

### 5.4 Auto-recover is dead-pid-only

`_maybe_auto_recover` returns early unless the owned pid is **dead** (`session_service.py:752-755`). A live-but-slow (or genuinely hung but alive) call is deliberately left for "the explicit recovery path," but there is no in-tool explicit recovery — so it stays poisoned until an out-of-band kill.

---

## 6. Research-Backed Best Practices

Four researchers converged on one principle: **separate deadline from liveness; a slow-but-alive op is not hung; only confirmed loss of liveness triggers destructive recovery.** Details and sources:

### 6.1 MCP protocol facts

- **No protocol-level timeout.** MCP defines no request timeout; the 60 s ceiling is *local server policy* and is freely changeable. (MCP specification — https://modelcontextprotocol.io/specification)
- **Progress is advisory.** `notifications/progress` + `progressToken` are server→client, rate-limited, and informational only. (MCP specification, "Progress" — https://modelcontextprotocol.io/specification)
- **Cancellation is cooperative and may be ignored** for uncancellable calls; late results are ignored. (MCP specification, "Cancellation" — https://modelcontextprotocol.io/specification)
- **Tasks extension is experimental and not in the Python SDK 2.x.** `io.modelcontextprotocol/tasks` was moved out of core (SEP-2663) and has no client support yet. Recommendation: build an **internal seam** shaped like Tasks now; do **not** depend on the extension. (SEP-2663 — https://github.com/modelcontextprotocol/modelcontextprotocol/issues/2663 ; Python SDK — https://github.com/modelcontextprotocol/python-sdk)

### 6.2 Mature-server patterns (decision rule)

| Operation profile | Pattern |
|---|---|
| `< 10 s` | Block (synchronous). |
| `10–60 s` | Block + progress, **or** block-then-fallback (e.g. Databricks `wait_timeout`, then return a handle). |
| `> 60 s`, side-effects, or must-survive-disconnect | Job + poll: `start` / `status` / `cancel`, durable store, `poll_after_s`, long-poll ≤ 45 s, explicit `lost` status. |

Sources: Google AIP-151 "Long-running operations" (https://google.aip.dev/151); Databricks Jobs API `wait_timeout` (https://docs.databricks.com/api/workspace/jobs); MCP specification.

### 6.3 Documented polling failure mode (why "honest async only when warranted")

Small models, when handed a job handle, frequently **print "waiting" and end the turn** (documented 0/5 success in one study). Mitigation: a harness guard, and expose async job+poll **only** for operations that genuinely need it (open/save/test_suite/macro > 90 s) — not as a blanket pattern.

### 6.4 State machine (per-class `B` exec budget + `C` liveness ceiling)

Recommended worker states:

- `RUNNING` / `BUSY` — normal.
- `SLOW` — **deadline `B` exceeded but alive** → report, **do NOT poison**.
- `SUSPECTED` — no liveness progress for `k` samples.
- `CONFIRMED_HUNG` — owned pid dead, **or** no liveness past `C + confirm` → **only now** `NEEDS_RECOVERY` + off-thread pid-scoped `reap_owned_pid` → `reset`.
- `RECOVERING` — rebuild in progress.

Watchdog runs on a **free thread** and polls `_pid_alive` (Win32, not COM). Kill is pid-scoped last resort with backoff + jitter + an attempt budget (anti-loop). Only `READ_ONLY` is retried (already enforced at `com_worker.py:656-675`).

### 6.5 Draft per-class defaults (to be validated against real timings in the PoC)

| Op class | Budget `B` (deadline) | Liveness `C` (ceiling) |
|---|---|---|
| small read | 15 s | 60 s |
| small mutate | 30 s | 90 s |
| **open** | **300 s** | **600 s** |
| save / close | 120 s | 300 s |
| diagnostic | 60 s | 120 s |
| macro | `macro_timeout_s` = 90 s / `max(300, 3×B)` | — |
| close-section (drain) | 10 s | 15 s |

New env knobs: `EXCEL_VBA_MCP_OPEN_TIMEOUT_S = 300`, `EXCEL_VBA_MCP_LIVENESS_CEILING_S = 600`.

### 6.6 Reusable primitives already in the codebase

- `ExcelAppAdapter._pid_alive` — Win32 `OpenProcess` + `GetExitCodeProcess`, **not COM** (`app_adapter.py:434-467`).
- `reap_owned_pid` / `_terminate_pid` — pid-scoped `taskkill /F /PID`, single-process blast radius (`app_adapter.py:178-214, 469-491`).
- `macro_service._handle_hung` — already does off-thread terminate-then-recover (`asyncio.to_thread(_terminate_owned_pid, ...)`, `macro_service.py:864-866`; recover at `924-926`).
- `reset()` queues `_do_reset` as a close-section item on the **same** worker (`com_worker.py:489-521, 767`) — confirming recovery **must** be off-thread to preempt a live call.

---

## 7. Priority Architectural Solutions

### P1 — Decouple queue-wait from execution deadline; add the SLOW/SUSPECTED/CONFIRMED state machine; off-thread recovery

**What changes (files/methods):**

- `adapters/excel/com_worker.py`
  - `submit()` (`:330-409`): split the single `timeout` into a **queue-wait** bound and a per-class **execution budget `B`**. A `TimeoutError` on *queue-wait* must **not** set `NEEDS_RECOVERY` — it returns a `busy` / `retry_after_s` result instead. Only an execution-budget breach moves the item to `SLOW`.
  - `_WorkItem` (`:133-147`): carry both `queue_deadline` and `exec_deadline`.
  - `_process_item` (`:596-640`): on exec-budget breach set `SLOW` (not `NEEDS_RECOVERY`); record liveness samples.
  - New watchdog on a **free thread** polling `_pid_alive` (`app_adapter.py:434-467`) to promote `SLOW → SUSPECTED → CONFIRMED_HUNG`.
  - `_clear_active` ordering fix (see P-D) so `busy` is truthful.
- `application/session_service.py`
  - `_maybe_auto_recover` (`:729-757`): extend beyond dead-pid to `CONFIRMED_HUNG`, and call an **off-thread** reap (`reap_owned_pid`, `app_adapter.py:178-214`) before `reset()` — never rely on the worker's own queue to preempt.
  - `open_workbook` (`:167-334`): pass the open-class budget (300 s) and liveness ceiling (600 s).
- New thin `excel_recover` tool (last resort) in `api/` + `server.py` registration, delegating to `recover_session` with off-thread reap. This is the missing "in-tool restart" the audit asked for (problem doc lines 96, 117, 147).

**Interaction with non-negotiables:** COM stays on the single STA thread; the watchdog is **non-COM** (`_pid_alive` is Win32). Originals never modified. Approval unchanged. Kill remains pid-scoped (`reap_owned_pid`), never a wildcard image kill.

**SCHEMA_VERSION:** **No bump.** New status fields (`state`, `retry_after_s`, `liveness_samples`) are **additive defaulted fields** on existing DTOs — allowed without a bump (`DomainModel` forbids *extra* fields but declared defaulted fields are fine; `models.py:62-68`).

**Tests to update:** `tests/unit/test_com_worker.py`, `tests/unit/test_recovery.py`, `tests/unit/test_session_service.py`, `tests/unit/test_health_tools.py` (all confirmed present; plus the `integration_excel/` variants).

**Risks + mitigations:** Risk that `SLOW` is reported for a genuinely hung call → mitigate with the `C + confirm` ceiling before `CONFIRMED_HUNG`. Risk of watchdog false positives → require `k` no-progress samples. Risk of off-thread reap racing a finishing call → reap is pid-scoped and idempotent (`app_adapter.py:178-214`).

**Effort:** Medium.

### P2 — Job-seam for genuinely long operations (open / save / test_suite / macro > 90 s)

**What changes:** introduce an **internal** `start/status/cancel` seam with a durable store, `poll_after_s`, long-poll ≤ 45 s, and a `lost` status — **mapped 1:1 to the Tasks extension shape** so a future swap is mechanical, without depending on the (experimental, SDK-absent) extension today (SEP-2663). Apply to `open`, `save`, `test_suite_run`, and `macro_run` when expected duration > 90 s.

**Interaction with non-negotiables:** the seam is application-layer; COM still runs only on the STA worker; the durable store is filesystem (existing `storage/` pattern). No new COM surface.

**SCHEMA_VERSION:** **No bump** (additive job-handle DTOs, defaulted).

**Tests to update:** `test_session_service.py`, `test_health_tools.py`, plus a new `test_job_seam.py`; pivot tests unaffected except where open is exercised (`test_pivot_tables_*`).

**Risks + mitigations:** documented small-model polling failure (print "waiting" and end turn) → expose async **only** for the listed op classes and add a harness guard; keep the synchronous path as the default.

**Effort:** Medium–Large.

---

## 8. Per-Block Disposition

- **Block A → P1 state-machine + off-thread recovery.** Adopt `SLOW/SUSPECTED/CONFIRMED_HUNG`; recovery via off-thread pid-scoped reap + thin `excel_recover`. **Reject** per-op workers (violates ARCH-009 single-STA) and auto-kill-on-timeout (kills healthy slow opens).
- **Block B → per-sheet cost guard returning bounded `too_large` + sheet-hint early-exit; `range_read_values`-only fallback.** **Reject** a file-size trigger (file size is unreliable — §4.2.1). Guard on **sheet dimension / row count**, returning a bounded `too_large` with the offending sheet name, and steer callers to `range_read_values` (the only path that works on `dashboard_data`, which is not a ListObject).
- **Block C → startup root log + additive `original_name` + doc `inspect`-is-not-readonly; defer the `inspect ⇒ read_only` contract change; reject hardlink/move/runtime-root-add.** Log the resolved roots at startup so CWD-dependence is visible (`server.py:752-774`); add an additive `original_name` field to `WorkbookInfo` (no SCHEMA bump) to reduce GUID confusion; document the writable-inspect footgun. **Reject** hardlinking/moving originals (breaks "originals never modified") and runtime root mutation (weakens deny-by-default).
- **Block D → reorder `_clear_active` before `_complete_result`; epoch-scope notes; new `list_workbooks_reconciled` without touching `list_workbooks_sync`; doc note.** In `_process_item` (`com_worker.py:629-640`) clear active **before** completing the result to kill the P6 `busy:true` race; tag notes with a session epoch and expose current-epoch notes in `excel_health` (P5); add `list_workbooks_reconciled` that cross-checks the owned pid, leaving `list_workbooks_sync` (used at `health_tools.py:165`) untouched (P7).

---

## 9. Risks, Open Questions, and PoC Validation Plan

**PoC (do this before committing to the defaults in §6.5):**

1. **Calibrate open `B`/`C`.** Open `base/CategoryData.xlsx` (124.81 MB, external cache `recordCount` 1,251,228) with a **300 s** budget and measure the **real** open time. Repeat for `weekly/report_w.xlsx` (5.51 MB, 566k-row `dashboard_data`). Confirm the 300 s / 600 s defaults or adjust.
2. **Confirm external-cache behaviour.** Verify `pivot_table_refresh` / `pivot_cache_refresh` / `pivot_field_*` return `unsupported_source` on these external caches (expected per `pivot_tables_service.py:29-31, 51-52`).
3. **Confirm `range_read_values` page count.** Verify `ceil(566386/50000) = 12` windows on `dashboard_data!A1:B566386` at `MAX_RANGE_CELLS = 100_000` (`models.py:666`, `ranges_adapter.py:3710-3735`).

**Open questions:**

- What is the real `SUSPECTED` sample count `k` and confirm window for this host's Excel build (16.0.14332)?
- Should `list_workbooks_reconciled` be the default for `excel_list_workbooks`, or additive-only to preserve the current contract?
- Does the job-seam's durable store need retention/GC policy, or is the existing `storage/` cleanup sufficient?

**Risks:** changing timeout semantics could mask a real hang if `C` is set too high → mitigate with the `C + confirm` ceiling and the attempt budget. Off-thread reap must remain pid-scoped (`app_adapter.py:178-214`) — never a wildcard kill.

---

## 10. Appendix

### 10.1 Sources (URLs)

- MCP specification (timeout / progress / cancellation): https://modelcontextprotocol.io/specification
- MCP Python SDK (2.x, no Tasks): https://github.com/modelcontextprotocol/python-sdk
- SEP-2663 (Tasks extension moved out of core): https://github.com/modelcontextprotocol/modelcontextprotocol/issues/2663
- Google AIP-151, Long-running operations: https://google.aip.dev/151
- Databricks Jobs API (`wait_timeout`): https://docs.databricks.com/api/workspace/jobs
- Microsoft Excel object model, `PivotCache` / `PivotTable`: https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache

### 10.2 Glossary

- **ARCH-009** — all COM on one `excel-com-sta` thread; only DTOs cross; services/api never import `pythoncom`/`win32com`.
- **Deadline (`B`)** — per-op-class budget for how long an operation may legitimately run.
- **Liveness (`C`)** — ceiling for how long the worker may show no progress before it is `CONFIRMED_HUNG`.
- **Head-of-line blocking** — one slow FIFO item delaying all later items on the single worker.
- **Working copy** — the per-`workbook_id` copy under `work/<GUID>`; the original is never opened or modified.
- **`unsupported_source`** — the pivot result for external/OLAP/Data Model caches on the target build.
- **Job-seam** — internal `start/status/cancel` + durable store + `poll_after_s` + `lost`, shaped 1:1 to the Tasks extension for a future swap.
- **SCHEMA_VERSION** — Pydantic contract version (`models.py:62`, currently `1`); additive defaulted fields do not require a bump.

### 10.3 Real-data measurement method

Read-only, in-memory `zipfile` inspection of each workbook's parts (no extraction to disk, no Excel launch): read `xl/pivotCache/pivotCacheRecords1.xml` for `recordCount` and approximate inflated size, `xl/pivotCache/pivotCacheDefinition*.xml` for `cacheSource@type` / `connectionId`, workbook `sheet*.xml` for the `dimension` of the large sheet, and `xl/workbook.xml` for sheet ordering. File sizes were independently re-confirmed with a directory listing (`Get-ChildItem`). No file under `.ai/tmp/theiss` was modified, and no macro was executed.

---

*End of Report 03. This report is advisory: it cites evidence and recommends evolution; it does not modify production code.*