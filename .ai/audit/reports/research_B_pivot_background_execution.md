# Research B — Pivot read execution, soft-timeout (B2) and background job (B3)

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `24201f4`
- **Scope:** RESEARCH ONLY. No production code/contracts changed. PoCs live under `.ai/tmp`.
- **Subject:** the shipped source-type gate (B1, WP1) on `pivot_read` `fields-list` / `field-get-filters`, and the OPTIONAL B2 (shorter soft timeout) / B3 (background job) on top of it.
- **Date:** 2026-11-27

---

## 1. Execution trace (file:line)

### 1.1 API surface — `api/pivot_tools.py`
- `pivot_read(...)` dispatcher: `fields-list` → `pivot_fields_list(workbook_id, pivot_table)` (`pivot_tools.py:624-625`); `field-get-filters` → `pivot_field_get_filters(...)` (`pivot_tools.py:626-629`).
- `pivot_fields_list` (`pivot_tools.py:277-288`) and `pivot_field_get_filters` (`pivot_tools.py:291-310`) are **blocking** — they `await service.<method>(...)` with **no `background` parameter**. Only `pivot_tables_list` has `background` (`pivot_tools.py:218-246`, `start_long_op(..., timeout=60.0)` at `240-245`). **B3 is therefore not wired for the two hanging verbs today.**

### 1.2 Service — `application/pivot_tables_service.py`
- `_DEFAULT_TIMEOUT_S = 60.0` (`pivot_tables_service.py:108`).
- `list_pivot_fields` → `_submit_read(..., "pivot_fields_list", lambda wb: adapter.enumerate_pivot_fields(wb, pivot_table), ...)` (`pivot_tables_service.py:264-288`).
- `get_pivot_field_filters` → `_submit_read(..., "pivot_field_get_filters", lambda wb: adapter.get_field_filters(...), ...)` (`pivot_tables_service.py:290-322`).
- `_submit_read` (`pivot_tables_service.py:869-888`) and `_submit_mutation` (`823-842`) both call `worker.submit(call, timeout=_DEFAULT_TIMEOUT_S, section="diagnostic", op_class=...)`.
- **No per-call timeout override, no job registry, no `background` path** for these two verbs.

### 1.3 The gate (B1) — `adapters/excel/pivot_tables_adapter.py`
- `_gate_supported_source_read(pivot, pivot_table)` (`pivot_tables_adapter.py:946-969`): resolves `pivot.PivotCache()`, classifies `SourceType`/`OLAP` via `_classify_source_type` (`327-346`), and **raises `UnsupportedPivotSourceError` for anything ≠ `DATABASE` (incl. `UNKNOWN`) BEFORE `PivotFields` is touched**.
- Call sites prove ordering:
  - `enumerate_pivot_fields`: gate at `1931`, **then** `_data_field_index` (`1932`) / `_describe_all_fields` (`1933`) which read `PivotFields()`.
  - `get_field_filters`: gate at `2611`, **then** `_resolve_field_target` (`2612`) → `_source_field_entries` → `PivotFields()` (`977`).
- The gate reads only **cache metadata** (`SourceType`, `OLAP`) — properties of the already-stored pivot-cache definition — not a refresh; `PivotFields` enumeration is what forces Excel to materialize external/OLAP items (the WP1 hang). Rationale documented at `pivot_tables_adapter.py:946-956`.

### 1.4 Worker — `adapters/excel/com_worker.py` (the timeout authority)
- Single `threading.Thread`, `queue.SimpleQueue` FIFO, one `Future` per item (`com_worker.py:266-322`).
- `submit` (`376-506`): clamps `timeout` to `liveness_ceiling_s` (`442-446`), enqueues, then `await asyncio.wait_for(asyncio.wrap_future(fut), timeout=effective_timeout)` (`464`). **The deadline is enforced client-side by `wait_for`; the worker thread is never told to stop.**
- On `TimeoutError` (`465-506`):
  - `not item.started` → **queue-wait** timeout → `ComBusyError` (`busy`, `retry_after_s=1.0`), **worker state unchanged** (`468-487`).
  - `item.started` → **execution** deadline → publish active triple, set `DEADLINE_EXCEEDED`, raise `ExcelHungError` (`SUSPECTED_STALLED`) (`488-506`). **The call is NOT cancelled and the pid is NOT terminated.**
- `_process_item` (`704-760`): sets `item.started=True` (`715`); if already past deadline on dequeue → `ExcelHungError` "expired before it could start" (`716-725`). The running `fn` runs to completion; on return `_clear_active` resumes `RUNNING` from `DEADLINE_EXCEEDED` (clear-before-complete, `984-1012`).
- Late result: `_complete_result`/`_complete_exception` set the abandoned future (`954-960`), but `wait_for` already cancelled the wrapped awaitable → **the late COM result is discarded, never reconciled at the worker layer.**

### 1.5 Jobs — `application/job_registry.py` + `storage/job_store.py` + `api/job_tools.py`
- `start_long_op` (`job_tools.py:56-93`) → `registry.start(...)` returns a prompt `RunningJob{status:"running", job_id, poll_after_s}` (`job_registry.py:241-297`); `poll_after_s` clamped `[1,30]`, default 5 (`63-65,101-104`).
- Sync callable → `_run` (`421-467`) wraps the **same single-STA `worker.submit`** (no second STA). `ExcelHungError` → job **`EXPIRED`**, `error_code=EXCEL_HUNG`, `result=None` (`436-444`). `ComBusyError` → `FAILED` + `retry_after_s` (`445-455`).
- `job_get` → `registry.status` (`job_tools.py:96-110`, `job_registry.py:309-326`); `job_request_cancel` → `registry.cancel` records **intent only** (`job_tools.py:113-128`, `job_registry.py:328-353`).
- Persistence: RUNNING marker at start, finalize on terminal (`job_registry.py:543-567`); `JobStore` startup-reconcile turns non-terminal disk records into `interrupted_unknown` (`job_store.py:162-194`).
- `mark_interrupted` (`job_registry.py:355-399`): after a successful recovery, non-terminal/expired/deadline-result jobs → `INTERRUPTED_UNKNOWN` (`outcome="unknown"`, `retryable=False`); a late completion cannot overwrite it (`_finalize` guard `533-534`).

### 1.6 Recovery — `application/session_service.py`
- `recover_owned_process` (`1174-1332`): RECOVERY approval auto-allowed (`1224-1229`); **fail-closed** if worker closed or no live identity-verified handle (`1242-1267`); `terminate_owned_handle` via `asyncio.to_thread` — a **free-thread kernel op that works while the STA is blocked in COM** (`1277`); then `worker.reset()` (`1282`); on success `_mark_jobs_interrupted()` (`1301`, `1334-1346`).

### 1.7 Config — `config.py`
- `macro_timeout_s` default `90.0` (`config.py:54-60`), `open_timeout_s` `300.0` (`61-66`), `liveness_ceiling_s` `600.0` (`67-73`). Pivot read/mutation use the service's own `_DEFAULT_TIMEOUT_S=60.0`, clamped by the 600 s ceiling → **effective 60 s**.

### 1.8 Exact behaviour on soft-timeout expiry (answer)
| Question | Answer | Evidence |
|---|---|---|
| Caller gets a timeout while COM continues? | **Yes.** `wait_for` raises `ExcelHungError` at the deadline; the STA `fn` keeps running. | `com_worker.py:464,488-506` |
| STA worker still occupied? | **Yes.** Single thread is inside `_run_with_retries(fn)`; `state=DEADLINE_EXCEEDED`, `accepting=true`, `active_operation` set. | `com_worker.py:493-494,704-760` |
| Job/operation state? | Background job → **`expired`** (`result=None`, `error_code=EXCEL_HUNG`). Direct op → caller sees `ExcelHungError`; worker `DEADLINE_EXCEEDED`. | `job_registry.py:436-444`; `com_worker.py:494` |
| Can queued ops run? | **No, not until the hung call returns** (FIFO). A queued op whose own deadline passes first → `ComBusyError`; otherwise it waits, then runs. | `com_worker.py:459,464,468-487`; PoC B |
| Can `excel_recover` run independently? | **Terminate: yes** (free-thread). **`worker.reset()`: no** — it is a `close`-section FIFO item behind the hang, bounded by the orphan join (production default 10 s; the PoC used a 3 s override) → `ExcelHungError`. | `session_service.py:1277,1282`; `com_worker.py:589-615`; PoC C |
| Late COM result retained/discarded/reconciled? | **Discarded** at the worker (abandoned future); job stays `expired`/`result=None`. Only a successful `excel_recover` reconciles the job to `interrupted_unknown`. | `com_worker.py:954-960`; `job_registry.py:533-534,355-399` |
| Enforced by worker (abandon) or client? | **Client** (`asyncio.wait_for` in `submit`). The worker never abandons/interrupts the COM call. | `com_worker.py:464` |

---

## 2. B1 (shipped gate) — does it fully prevent the external/OLAP hang?

**Yes — HIGH confidence.** Live verification against the disposable external workbook (`.ai/tmp/theiss/weekly/report_w.xlsx`, copied before open, inspect mode, macros disabled):

```
opened        1.77 s
inventory     0.07 s   pivot_tables=["PivotTable1"]  caches=[{cache_index:1, source_type:"external"}]
fields_list   rejected 0.05 s  code=UNSUPPORTED_PIVOT_SOURCE  reason=unsupported_source  source_type=external
worker_state  closed (clean)
```
(`research_B_live_b1.py`; the fields-list op ran under a short 12 s deadline as a safety bound — it returned in 0.05 s, so the deadline was never approached.)

- The gate rejects on `SourceType=external` **before** `PivotFields` is enumerated (`pivot_tables_adapter.py:1931` precedes `1932-1933`; `2611` precedes `2612`). The dangerous enumeration is never reached.
- The inventory sweep (gated, identity-only) is also fast (70 ms) and reports the cache identity without materializing fields.
- **Static guarantee** (file:line ordering) + **live confirmation** agree. No hang observed; worker closed cleanly.

**Conclusion:** B1 fully closes the WP1 external/OLAP hang for `fields-list` and `field-get-filters`. No B2/B3 is needed to fix that specific hang.

---

## 3. PoC / timing evidence (controlled, no dangerous hang)

Faithful harness: the **real** `ExcelComWorker` (real FIFO / two-phase deadline / `DEADLINE_EXCEEDED` / clear-before-complete) + the **real** `JobRegistry`, with a no-op `ExcelAppAdapter` (test-suite `FakeAdapter` pattern) so no Excel is launched. The submitted `fn` does a bounded `time.sleep` to model an uninterruptible COM call (the B1 hang shape, minus the danger). Scripts: `research_B_poc.py`, `research_B_live_b1.py`.

### Scenario A — B3 background job (20 s op, timeout 60 s)
```
A_jobid_returned   start_call_s=0.0   status=running  poll_after_s=5.0
A_poll t=0.0  queued   (started=false)
A_poll t=1.0  running  (started=true)
A_poll t=20.02 succeeded  result="LONG_DONE"
A_second_result  t=20.02  status=succeeded  second_op_total_wait_s=20.02
```
- `registry.start` (== `start_long_op`) returned **instantly** (0.0 s) with a `job_id` — the MCP request does **not** block.
- Job transitions `queued → running → succeeded`; **result retained** and retrievable via `job_get`.
- A second op queued behind the long op waited the **full 20.02 s** — the single STA is occupied the whole time (**FIFO unchanged**).

### Scenario B — B2 shorter soft timeout (12 s op, timeout 4 s)
```
B_terminal            t=4.01  status=expired  error_code=EXCEL_HUNG  result=null
B_queued_op(2 s)      outcome=ComBusyError  waited_s=2.0  reason=queue_wait_timeout  retry_after_s=1.0
B_worker_state_during_hang  state=running/deadline_exceeded  accepting=true  active_operation=<lambda>
B_next_op_after_hang  result="RAN_AFTER_HANG"  waited_s=6.0  (waited until the 12 s op released the STA)
B_final_job_state     status=expired  result=null  (late COM result discarded; never reconciled)
B_worker_state_after  state=running  accepting=true
```
- Job **expired at 4.01 s**; the STA stayed occupied until the 12 s op returned (a generous-timeout op waited 6 s more and then ran).
- A queued op with a 2 s own-timeout got `ComBusyError` (queue-wait) at 2 s — **worker state unchanged** by the busy.
- The late 12 s result was **discarded**; the job stayed `expired`/`result=None`.
- Worker resumed `RUNNING` after the hang cleared (not poisoned).

### Scenario C — recovery independence (12 s op, timeout 4 s)
```
C_terminate  terminated=true  took_s=0.005  (free-thread kernel op; does not need the STA)
C_reset      outcome=ExcelHungError(timeout)  took_s=3.0  (reset is FIFO behind the call; bounded by the orphan join — PoC used a 3 s override; production default is 10 s)
```
- `terminate_owned_handle` (the kill) is **independent** of the blocked STA (5 ms).
- `worker.reset()` is **FIFO behind the hang** and times out at the orphan join (production default 10 s; the PoC used a 3 s override) — recovery's rebuild must wait for the kill to unblock the COM call (which `TerminateProcess` does).

---

## 4. B1 / B2 / B3 comparison

| Dimension | **B1 (gate, shipped)** | **B2 (shorter soft timeout)** | **B3 (background job)** |
|---|---|---|---|
| Fixes the external/OLAP hang? | **Yes** (rejects in ~50 ms). | No (hang still occurs; only the caller's wait shortens). | No (hang still occurs; the request returns early but the STA is still stuck). |
| Caller wall-clock on a real hang | n/a (no hang) | ↓ (e.g. 60 s → 8 s) | **~0 s** (returns `job_id`). |
| STA freed by the change? | n/a | **No** — occupied until COM returns. | **No** — single-STA/FIFO unchanged. |
| FIFO starvation risk | none | **Worse**: shorter deadline → more `ComBusyError` for queued ops; the hung op still blocks FIFO. | **Same** underlying FIFO block, but the caller can keep polling instead of erroring. |
| Breaks slow *supported* pivots? | No (only non-`DATABASE` rejected). | **Yes** — a legitimate large `DATABASE` pivot read > new timeout is falsely `expired`. | No (timeout can stay 60 s; the op completes in the background). |
| Monitoring / recovery | n/a | none | **Yes** — `job_get` polling, `job_request_cancel` (intent), recovery → `interrupted_unknown`. |
| Late-result handling | n/a | discarded; job `expired` | discarded; job `expired`/`interrupted_unknown` (never a false success). |
| Code/contract change | already shipped | small (per-verb timeout constant) | moderate (wire `background` into `fields-list`/`field-get-filters` + `start_long_op`). |

---

## 5. Benefits / costs

### B2 — shorter soft timeout
- **Benefit:** the caller fails fast instead of waiting 60 s on a genuine hang. Marginal given B1 already prevents the *known* hang.
- **Cost / risk:** COM is uninterruptible and the worker stays occupied, so a shorter timeout does **not** free the STA and does **not** reduce FIFO starvation — it only converts a slow-but-valid `DATABASE` pivot read into a false `expired`/`EXCEL_HUNG`, and increases `ComBusyError` churn for queued ops. Net negative for supported slow pivots.

### B3 — background job (`start_long_op` → `job_id`)
- **Benefit:** removes the COM call from the MCP request (returns `job_id` in ~0 s) **without** changing single-STA/FIFO; prevents client wall-clock timeout; adds `job_get` monitoring, `job_request_cancel` (intent-only), and correct `interrupted_unknown` reconciliation on recovery; the result is retained and never a false success.
- **Cost / risk:** the STA is still occupied for the op's duration, so **FIFO starvation is not solved** — a queued op still waits (PoC A: 20 s). It changes *who waits* (the poller, not the request) and makes the wait observable/recoverable. Requires wiring `background` into the two verbs (today only `tables-list` has it). The 60 s per-call deadline should be kept (B3 lets you keep it without blocking the caller).

---

## 6. Single-STA / FIFO risks (explicit)
- The worker is one thread with a FIFO queue; an uninterruptible COM call blocks **all** subsequent ops until it returns (PoC A: 20 s; PoC B: 12 s). Neither B2 nor B3 changes this.
- B2 **amplifies** starvation: more queued ops hit their own queue-wait deadline → `ComBusyError`.
- B3 **hides** the block from the caller (early `job_id`) but the block is still real; the agent must poll and may still wait the full op duration for a dependent op.
- Recovery: the kill (`terminate_owned_handle`) is independent and fast; the rebuild (`worker.reset()`) is FIFO behind the hang and bounded by the orphan join — so recovery succeeds only once the kill unblocks the COM call.

---

## 7. Recommendation

**Implement B3. Do NOT implement B2. Keep B1 as the hang fix.**

1. **B1 (shipped) is sufficient** to prevent the external/OLAP `PivotFields` hang — verified live (reject in ~50 ms). No further work is required for that specific failure.
2. **B2 is net-negative**: it cannot free the uninterruptible STA, does not reduce FIFO starvation, and risks false `expired` on legitimate slow `DATABASE` pivots. Reject.
3. **B3 is the right optional enhancement** for the residual "a *supported* pivot read is slow" case: wire `background: bool = False` into `pivot_fields_list` / `pivot_field_get_filters` (mirroring `pivot_tables_list`, `pivot_tools.py:218-246`) via `start_long_op(partial(service.<method>, ...), timeout=60.0, section="pivot", op_class=READ_ONLY)`. Keep the 60 s per-call deadline (B3 makes it non-blocking). This yields prompt `job_id`, `job_get` monitoring, intent-only cancel, and correct `interrupted_unknown` on recovery — all already implemented in `JobRegistry`/`JobStore` — with **no change to single-STA/FIFO**.

**Risk note for B3:** it does not eliminate FIFO starvation (the STA is still occupied); it makes the wait observable and recoverable rather than a hard client timeout. Document this so agents poll with `poll_after_s` and do not assume a `job_id` means the STA is free.

---

## 8. Artifacts / how to reproduce
- `.ai/tmp/research_B_poc.py` — single-STA/FIFO + B2 + B3 + recovery (no Excel).
- `.ai/tmp/research_B_live_b1.py` — bounded live B1 verification on the disposable external workbook.
- Both run with `.\.venv\Scripts\python.exe <script>`; originals untouched (the external workbook is copied before open; the live check reaps the owned Excel at shutdown).

### Confidence
- B1 prevents the hang: **HIGH** (static file:line ordering + live 50 ms rejection).
- Soft-timeout semantics (client-enforced, STA stays occupied, late result discarded): **HIGH** (code + PoC A/B/C).
- B3 behaviour (prompt job_id, FIFO unchanged, retained result, recovery reconciliation): **HIGH** (real `JobRegistry` + `ExcelComWorker` PoC).
- B3 not wired for `fields-list`/`field-get-filters` today: **HIGH** (`pivot_tools.py` only `tables-list` has `background`).