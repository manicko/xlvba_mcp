# task_010 — Dedicated Excel COM STA Worker — Research Report

**Author:** researcher
**Date:** 2026-11-19
**Repo:** `C:\py_exp\xlvba_mcp`
**Status:** complete (research only; no code under `src/` or `tests/` modified)
**Target task:** `task_010` — *Implement the serialized Excel COM STA worker* (`plan.md` G3, critical path)
**Authoritative addendum:** `.ai/plans/plan_addendum_background_execution.md` (BG01–BG09, R11–R14)
**Binding audit constraints:** ARCH-009 (sync SDK tools run on anyio threads → `api/*` never touch COM; `submit()` awaitable, no blocking join), ARCH-005 (worker owned by `MCPServer(lifespan=...)` / `main()`), ARCH-013 (visibility policy), ARCH-003 (affinity probe must show `RPC_E_WRONG_THREAD`), ARCH-007 (canonical path passed to COM; revalidation on the worker)

**Confidence legend:** `[observed]` = executed live on the target build; `[sdk]` = verified in the installed wheel; `[doc]` = Microsoft Learn/Support; `[inference]` = reasoned conclusion from observed facts.

**Validated target:** Windows 10 Enterprise LTSC 2019 build 17763.9020 x64; Excel 16.0.14332 x64 (Office 2016 family); Office/VBE UI 1033 (en-US); Python 3.14.7; pywin32 312; MCP SDK 2.3.0 (`mcp.server.mcpserver.MCPServer`); pywinauto 0.6.9.

---

## 1. Scope

This report answers eight design questions for `task_010` inside the already-approved architecture (no new service, no VM/isolated environment). It specifies, implementation-ready:

1. Threading model (dedicated STA thread + queue + awaitable `Future` vs `anyio.to_thread` per call).
2. Awaitable `submit(fn, timeout)` that never blocks the MCP event loop; truthful `EXCEL_HUNG`/`needs_recovery` mapping; safe handling of a still-running call after timeout.
3. COM apartment + message pumping; `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER` retry policy.
4. Hidden instance + settings (BG03/BG04/BG08): `Visible=False`, narrowly scoped `DisplayAlerts`, `ScreenUpdating`, guaranteed restore in `finally`.
5. Shutdown ownership (ARCH-005) wired into `MCPServer(lifespan=...)` and `main()`.
6. Proxy lifetime (no proxy crosses into the MCP layer; DTOs; leak detection).
7. Coexistence with the addendum (owned instance never becomes visible/foreground; compile indeterminate; macro UI blocked).
8. Concrete public API sketch for `adapters/excel/com_worker.py`, `adapters/excel/app_adapter.py`, `application/session_service.py`.

Out of scope (per instructions): proposing a redesign, new features, modifying code, running macros on production data, Trust Center changes.

Prior evidence reused: task_002 PoC (`task_002_planner_v3.md`, `excel_instance_poc_20261119_215845.txt`, `CONDITIONAL_GO`), `.ai/audit/hidden-compile-research-2026-11-19.md`, `.ai/audit/architecture-audit-2026-11-19.md`, `.ai/audit/g1-poc-research-2026-11-19.md`.

---

## 2. Facts established by prior work (reused, not re-proven)

- **Single dedicated STA thread works on target.** `pythoncom.CoInitializeEx(COINIT_APARTMENTTHREADED)` → `S_OK`; `DispatchEx('Excel.Application')` creates a distinct owned PID; serialized multi-op (read `Version` then `Caption`) on the same thread id; `Quit` + `CoUninitialize` → owned PID gone, no orphan; user foreground unchanged (BG03/BG04/BG05) `[observed: excel_instance_poc_20261119_215845.txt]`.
- **Affinity is real.** A cross-thread read of the owned proxy fails with `0x8001010E RPC_E_WRONG_THREAD` `[observed: same artifact, step affinity_probe]`. ARCH-003 requires the probe thread to have initialized COM; the production worker simply relies on the same apartment confinement (all calls on the STA thread).
- **COM init is per-thread and fragile.** A thread that touches `DispatchEx` without `CoInitializeEx` fails `0x800401F0 CO_E_NOTINITIALIZED` `[observed: ARCH-010]`. Therefore the worker **must** call `CoInitializeEx` explicitly, and no other thread may create COM.
- **Sync MCP tools run on anyio pool threads (up to 40)** via `anyio.to_thread.run_sync` (`mcp/server/mcpserver/utilities/func_metadata.py:149-164`) `[sdk: ARCH-009]`. Direct COM in a sync tool would create an apartment on an arbitrary pool thread → violates the single-STA rule.
- **Windows process exit paths** `[observed: ARCH-005]`: stdin-EOF → graceful rc 0 (~0.17 s); SIGTERM idle → rc 1, no Python `finally`; SIGTERM while a sync tool runs → process stays alive until the tool returns (a hung COM call cannot be interrupted by the client); hard kill → no cleanup. Any worker cleanup in a `finally` runs only on EOF/normal return.
- **`MCPServer` supports `lifespan=`** `[sdk]`: `lifespan: Callable[[MCPServer], AbstractAsyncContextManager[LifespanResultT]] | None`, and `mcp.run(transport='stdio')` drives `stdio_server`. Verified live signature this pass.
- **Compile is not non-intrusively achievable** on 16.0.14332: only the UIA Debug-menu route compiles, and it shows a popup + steals foreground + non-dismissible modal on error → `COMPILE_INDETERMINATE` `[observed: hidden-compile-research]`. The worker must therefore treat compile as an ordinary (indeterminate) operation, never a reason to make the app visible.
- **Application.Run blocks on a modal** when a procedure raises a compile/runtime error (`#32770`), and the modal is not dismissible non-intrusively `[observed: probe_run_compile_error_result.json]`. This is the concrete "hung operation" scenario the timeout must handle.

---

## 3. Question-by-question analysis

### Q1 — Threading model

**Options compared.**

| Option | Single-STA guarantee | Event-loop safety (ARCH-009) | Message pumping | Verdict |
|---|---|---|---|---|
| **(a) Dedicated `threading.Thread` running a COM STA loop; `queue.Queue` of work items; `concurrent.futures.Future` per item; `await asyncio.wrap_future(fut)` in `submit()`** | **Strong.** One thread calls `CoInitializeEx(APARTMENTTHREADED)` once; every proxy is created/used/released there; affinity probe proves confinement. | **Safe.** `submit()` is `async def`; it awaits a `Future` via `asyncio.wrap_future` — the loop thread never blocks. Multiple `submit()` calls serialize on the queue. | Controlled: pump only while idle-waiting (`PumpWaitingMessages` with a short timeout) and optionally between retry backoffs; never pump *inside* a live COM call. | **Recommended** |
| (b) `await anyio.to_thread.run_sync(fn)` per call | **Broken.** Each call may land on a different pool thread; the first thread to touch COM initializes an MTA (or STA) implicitly; proxies then cross apartments → `RPC_E_WRONG_THREAD` / intermittent `CO_E_NOTINITIALIZED`. ARCH-009 explicitly warns this. | Safe for the loop, but wrong apartment. | None — cannot pump a chosen STA from the pool reliably. | **Rejected** |
| (c) One long-lived `anyio` worker thread per session + `run_sync` | Threads are anyio-managed and not guaranteed to be the same OS thread across calls; same apartment-crossing risk as (b). | Safe. | Fragile. | **Rejected** |
| (d) `asyncio.run_coroutine_threadsafe` to a loop owned by the STA thread | Workable but heavier; adds a second event loop and cross-loop marshaling for no benefit over (a). | Safe. | Yes. | Overkill (violates "avoid overengineering") |

**Recommendation: (a).** Exact primitives:

- **One** `threading.Thread(target=self._run, name="excel-com-sta", daemon=True)`. `daemon=True` is a *last-resort* guard only (prevents process-hang at interpreter exit if cleanup is skipped); it is **not** the shutdown mechanism — graceful `shutdown()` is always attempted first. (See §5 for why daemon must not be relied on.)
- **Apartment init (once, top of `_run`):**
  ```python
  pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)  # -> S_OK / S_FALSE
  ```
  `S_FALSE` (already initialized) is acceptable; `RPC_E_CHANGED_MODE` is a hard error (the thread must not switch apartments). Record the result; do **not** call `CoUninitialize` until the very end of `_run` on the same thread.
- **Queue:** `queue.SimpleQueue` (unbounded FIFO; `SimpleQueue.put` is non-blocking and lock-light) or `queue.Queue`. A private `_WorkItem` dataclass carries `fn`, `future`, `token`, `deadline_monotonic`, and a `section` tag (`"open" | "close" | "save" | "diagnostic"` — see §4) used by `app_adapter` to decide settings scope. `SimpleQueue` is preferred because the worker is the only consumer and producers never need `task_done`.
- **Awaitable submit:** `concurrent.futures.Future` per item; `submit()` does `await asyncio.wrap_future(fut)` (loop-affine, non-blocking). `asyncio.wrap_future` is created on the loop thread each call, so each caller gets its own awaitable.
- **Worker loop:**
  ```text
  while not stop:
      item = queue.get()            # blocks the STA thread only (never the loop)
      if item is None: break        # shutdown sentinel
      if now > item.deadline: item.future.set_exception(TimeoutMarker); continue  # do not start
      if recovery hold and item.section != "close": item.future.set_exception(NeedsRecovery); continue
      run item.fn under AppSettings(section=item.section)  # try/except/finally
      pump idle messages between items (bounded)
  ```
- **Serialization proof:** the single consumer guarantees FIFO, one-at-a-time execution; `test_requests_are_serialized` (task_011) asserts submission order == completion order and identical STA thread id in every step.

**Why not block with `join()`:** ARCH-009 — `submit()` must not block the event loop. `await asyncio.wrap_future(...)` is the only correct bridge from a thread `Future` to the loop.

---

### Q2 — Awaitable submit with bounded timeout

**`submit(fn, timeout)` contract.**

```python
async def submit(self, fn: Callable[[], T], *, timeout: float, section: Section = "diagnostic") -> T:
```

Steps:
1. If the worker is not started/stopped → raise `ExcelVbaMcpError` (or a dedicated `WorkerUnavailableError`, code `EXCEL_HUNG` per §6 mapping) *without* queueing.
2. If the worker is in `NEEDS_RECOVERY` (a prior call timed out and may still be running) → **fail fast** with `ExcelHungError`/`needs_recovery`; do **not** enqueue (prevents starting a new COM call on a possibly-busy worker — see step 6).
3. Create `fut = Future()`; `deadline = time.monotonic() + timeout`; `queue.put(_WorkItem(fn, fut, deadline, section))`.
4. `return await asyncio.wait_for(asyncio.wrap_future(fut), timeout=remaining)` — two independent timeouts:
   - **Queue-wait timeout** (`wait_for`): the call never even started (worker busy on an earlier item).
   - **Execution timeout** (the item's `deadline` checked by the worker before it starts, plus an optional watchdog — see step 6).
   `asyncio.wait_for` guarantees the loop is freed at the deadline even if `wrap_future` never resolves.

**Timeout is truthful, not a guess.** The MCP tool must not report a false result when a timeout fires:

- On `asyncio.TimeoutError`, the tool asks `worker.get_status()` (a *non-COM* snapshot of worker bookkeeping: `state`, `active_operation`, `active_section`, `last_hresult`) and returns a structured result with `state = WorkbookSessionState.NEEDS_RECOVERY` (serialized value `"running/needs_recovery"`) and `code = EXCEL_HUNG` in the error/context. The tool does **not** claim the operation failed — it claims *unknown/possibly-still-running*.
- The worker sets `self._state = NEEDS_RECOVERY`, records `self._active = {token, section, started_monotonic, deadline}` **before** invoking `fn`, and clears it in `finally` when `fn` returns. So `get_status()` can always distinguish "idle" from "a call started and has not returned".

**Still-running queued call after timeout (do not start a new COM call):**

- The item that timed out may still be executing on the STA thread (e.g. blocked in `Application.Run` behind a modal). The worker cannot preempt it. Therefore, after the first timeout:
  - Worker enters `NEEDS_RECOVERY` and **stops dequeuing non-`close` items**. New `submit()` calls fail fast (`needs_recovery`) instead of piling up COM calls behind a hung one.
  - The timed-out item is **never cancelled** (a `concurrent.futures.Future.cancel()` only works before it starts; once running it is uninterruptible). The `Future` is completed with a `TimeoutMarker` exception so the awaiting tool returns immediately; the running `fn` continues on the STA thread until it returns or the worker is reset.
  - A `close`-section item (shutdown drain) *is* allowed through even in `NEEDS_RECOVERY`, because it is the recovery action itself (close owned workbooks, `Quit`) — see §5.
- **No blind replay:** a timed-out MUTATING/EXECUTING call is **never** resubmitted. Only READ_ONLY items may be retried (bounded, see Q3). This is the ARCH-009/task_010 acceptance "hung worker → running/needs_recovery, not a stall" and "COM busy/rejected operations are not blindly replayed."

**`needs_recovery` exit:** the worker can leave `NEEDS_RECOVERY` only via (a) the hung `fn` finally returning (worker records `last_hresult`, re-checks `Application`/window state) or (b) an explicit `worker.reset()` that tears down and recreates the owned Excel instance on the **same** STA thread (close owned workbooks → `Quit` → drop refs → new `DispatchEx`), all inside the worker. The MCP layer never resets the apartment from another thread.

**Recommended timeouts (config, not hardcoded):** per-call ceiling supplied by the caller (`api` layer) from typed config; sane defaults: inventory/read ≤ 30 s, save/close ≤ 60 s, macro run ≤ 120 s, open ≤ 60 s. These are *soft* deadlines that produce `needs_recovery`, never a kill.

---

### Q3 — COM apartment + message pumping + busy/rejected retry

**Where `CoInitializeEx` is called:** exactly once, first statement of the worker thread `_run`, before any `DispatchEx`. No other code path initializes COM. `CoUninitialize` exactly once, last statement of `_run`, after all refs are dropped (§5).

**Message pumping strategy (concrete):**

- **Idle wait (between items):** the worker loop's `queue.get()` blocks. To service cross-apartment/queued calls and let Excel finalize its own bookkeeping, replace the bare blocking `get()` with a bounded `get(timeout=0.05)` loop and call `pythoncom.PumpWaitingMessages()` on each timeout. This keeps the STA responsive without reentering user code. (Equivalent: `MsgWaitForMultipleObjects` on a self-pipe/event; `PumpWaitingMessages` is the planner-v3-chosen primitive and is retained — `task_002_planner_v3.md §6.1`.)
- **During a long operation:** the worker must **not** pump inside a live synchronous COM call (that would reenter the worker and risk duplicate side effects — ARCH "message pump starvation / reentrancy"). Long ops (open/save/macro) are simply awaited by the loop thread via `wait_for`; the STA thread is inside the COM call. If Excel needs to service cross-apartment calls mid-operation, that is handled by COM's own STA message loop; the worker does not inject `PumpWaitingMessages` into `fn`.
- **Between retry backoffs only:** if a READ_ONLY item is retried, pump during the sleep window (`PumpWaitingMessages` in a short loop) so the server is not seen as hung — bounded, and only for READ_ONLY.
- **Shutdown orphan wait:** pump during the bounded orphan poll (§5) so Excel can finalize `Quit`.

Document in code: `message_pump_strategy = "PumpWaitingMessages(interval=0.05s) on idle queue wait and retry backoff; never inside fn"`.

**Busy/rejected handling (retry policy):**

| HRESULT | Constant | Meaning | Policy |
|---|---|---|---|
| `0x80010001` | `RPC_E_CALL_REJECTED` | Callee busy; call rejected. | **READ_ONLY:** bounded retry (≤3 attempts, exponential backoff 0.5→1→2 s, pump between). **MUTATING/EXECUTING:** never auto-replay; surface `RpcCallRejectedError` (code `RPC_E_CALL_REJECTED`) with `context={"attempts": 1, "class": op_class}`; session stays `RUNNING` (rejection is not a hang). |
| `0x8001010A` | `RPC_E_SERVERCALL_RETRYLATER` | Server busy; retry later. | Same classification as above; maps to `ComBusyError` (code `COM_BUSY`) — already in `domain/errors.py` `_HRESULT_TO_CODE`. |

The operation class comes from the existing shared enum `domain/statuses.py::OperationClass` (`READ_ONLY / MUTATING / EXECUTING`) — reused, not reinvented (satisfies ARCH "worker and task_009 share one enum"). The `_WorkItem` carries an `op_class: OperationClass`; the worker reads it to decide retry eligibility.

**Truthful mapping:** `RPC_E_CALL_REJECTED` after retries → `RpcCallRejectedError` (do **not** retry again). `RPC_E_SERVERCALL_RETRYLATER` → `ComBusyError`. A modal/hang (no HRESULT, no return) → timeout path → `ExcelHungError`/`needs_recovery` (§Q2). Never convert a rejection into a success or a silent retry of a mutation.

---

### Q4 — Hidden instance + settings (BG03/BG04/BG08)

`app_adapter.py::ExcelAppAdapter` centralizes all Application-state mutation. **All settings are applied and restored inside the worker thread**, inside `finally`, for every section.

**Visibility (resolves ARCH-013 for task_010):**

- The owned instance is created with `Visible = False` (set immediately after `DispatchEx`, before any workbook op) and **kept hidden for its entire lifetime** (addendum §2.1: open/read/write/compile/macro/save/snapshot/restore/cleanup). task_010 does **not** implement the old ARCH-013 "Visible=True for execute" branch — the authoritative addendum supersedes it (addendum §1/§4 task_010 row). Compile does not change this: on 16.0.14332 compile is `COMPILE_INDETERMINATE`, so no UIA/visible route is taken (§7).
- The worker never calls `.Activate`/`.Select`, `SetForegroundWindow`, `ShowWindow(SW_SHOW)`, `SendKeys`, or any focus-changing API. Workbook/worksheet/range access uses direct object-model references. This is a hard rule in `app_adapter` and asserted in task_011/task_012/task_013 tests.
- **Owned-identity proof:** at startup the adapter records `owned_pid = app.Hwnd`→pid delta and `owned_caption`; these are stored in worker bookkeeping (non-COM ints/strings) and exposed via `get_status().notes` so tests can prove the owned instance is hidden and separate from any user instance (BG05).

**`DisplayAlerts` — narrowly scoped, save/restore (BG08/R14):**

- Never set globally for the worker lifetime. `ExcelAppAdapter` exposes a synchronous context manager used **inside** a submitted `fn`:
  ```python
  @contextlib.contextmanager
  def alerts_suppressed(self, app):  # runs on the STA thread
      prior = app.DisplayAlerts
      app.DisplayAlerts = False
      try:
          yield
      finally:
          app.DisplayAlerts = prior   # restore on success/failure/timeout/cancel
  ```
- Used **only** around operations whose default alert response is known-safe (e.g. `Close(SaveChanges=False)` on an owned workbook, `Quit`). Documented in the adapter docstring that `DisplayAlerts=False` is **not** a universal dialog blocker (does not suppress `MsgBox`, `UserForm.Show`, `FileDialog`, add-in UI, arbitrary Win32 dialogs — addendum §2.3). The macro service (task_022) handles those separately.
- Because the context manager's `finally` runs on the STA thread even if `fn` raises, restoration is guaranteed on success, exception, and timeout of the *enclosing* operation *provided the worker thread is still executing `fn`*. The one case `finally` cannot cover is a worker that is **hung inside** `fn` and never returns — that is exactly why `needs_recovery`/`reset()` (Q2) rebuilds the whole instance rather than trying to restore a setting on a stuck app.

**`ScreenUpdating` — optimization only:**

- `ScreenUpdating = False` is set around bulk operations (open, write) to reduce repaint, and restored in `finally`. It is **never** treated as a visibility control or a dialog blocker (addendum §2.2 / MS docs). Same context-manager pattern as alerts.

**`AutomationSecurity` (task_012, referenced for completeness):** set per open (`ForceDisable` for inspect, `ByUI` immediately before `Open` for execute) and restored in `finally`; owned by `app_adapter`, not global.

**`finally` guarantee matrix (BG08):**

| Exit path | Settings restore | Who runs it |
|---|---|---|
| `fn` returns normally | `app_adapter` context managers restore alerts/screenupdating | STA thread |
| `fn` raises COM error | same `finally` restores | STA thread |
| Enclosing `submit` times out but `fn` still running | cannot restore a running/stuck app → `needs_recovery` → `reset()` rebuilds instance (clean state) | STA thread (reset) |
| Enclosing `submit` times out before `fn` starts | item skipped by deadline check; nothing to restore | STA thread |
| Worker `shutdown()` | close owned workbooks → `Quit` (no settings to restore on a dead app) | STA thread |

---

### Q5 — Shutdown ownership (ARCH-005)

**Ownership point.** The worker is created and owned by an `ApplicationRuntime` object whose lifetime is bound to `MCPServer(lifespan=...)`:

```python
# server.py (target shape, task_010+012)
@asynccontextmanager
async def lifespan(server: MCPServer) -> AsyncIterator[dict[str, Any]]:
    runtime = ApplicationRuntime(config)          # builds PathPolicy, ExcelAppAdapter, ExcelComWorker
    await runtime.start()                          # starts the STA thread, awaits ready
    try:
        yield {"runtime": runtime}                 # request_lifespan_context result
    finally:
        await runtime.shutdown()                    # bounded, idempotent
```

`MCPServer("excel-vba-mcp", lifespan=lifespan)` `[sdk: verified signature this pass]`. `main()` calls `mcp.run(transport="stdio")`; the SDK enters/exits the lifespan around the STDIO session, so shutdown runs on stdin-EOF and on normal return. This satisfies ARCH-005 without `atexit` heuristics.

**Shutdown sequence (exact, ordered, all on the STA thread except the final join):**

1. **Stop accepting new work.** Set an internal `_accepting = False`; `submit()` begins failing fast (`WorkerUnavailableError`). (Do this from the loop side before signalling.)
2. **Enqueue a `close`-section drain item** (allowed even in `NEEDS_RECOVERY`) and a `None` sentinel. The drain item runs **after** any already-running item finishes (FIFO), so it never races a live COM call.
3. **Drain/cancel queued items.** For every not-yet-started queued `_WorkItem`, complete its `Future` with `WorkerUnavailableError` (so any awaiting tool returns promptly); do not run them. Items already running are awaited (bounded) — they are not forcibly killed.
4. **Close owned workbooks.** On the STA thread, for each tracked owned workbook still open: `wb.Close(SaveChanges=False)` inside `alerts_suppressed` (restore handled by context manager). Catch/log COM errors per workbook and continue (best-effort, never swallow silently — record in worker `shutdown_errors`).
5. **Quit owned instance.** `app.Quit()` inside `alerts_suppressed`. Record result.
6. **Release references.** Drop all proxy refs (`app`, `wb`, adapter-held refs), `del` locals, `gc.collect()` (diagnostic aid only).
7. **`CoUninitialize()`** on the STA thread (last COM action).
8. **Bounded join + orphan poll.** Loop side does `thread.join(timeout=orphan_timeout)` (e.g. 10 s). If still alive, record `SHUTDOWN_TIMEOUT` in worker bookkeeping (do not hang the server); the daemon flag lets the process exit. Best-effort orphan poll for `owned_pid` absence (record, never kill a user PID).
9. **Idempotent.** `shutdown()` called twice is safe (second is a no-op).

**Exit-path behavior (honest limits, from ARCH-005):**

| Path | Shutdown runs? | Orphan risk |
|---|---|---|
| stdin-EOF | Yes (lifespan `finally`) | No — owned instance closed, no orphan. |
| SIGTERM idle | No Python `finally` (observed rc 1) | Owned `EXCEL.exe` may orphan. Documented; recovery = kill **owned PID only**. |
| SIGTERM while a tool runs / hung macro | No (process waits or is terminated) | Owned instance may orphan. Documented. |
| Hard kill | No | Orphan. Documented. |

Mitigation for the non-clean paths: the worker records its `owned_pid` to a small runtime file (e.g. `.ai/tmp/owned_excel_pid.txt` or a configured artifact path) at startup so a later run / operator can identify and clean the **owned** PID only (never user PIDs). This is a documented orphan-recovery aid, not a Windows-Service/auto-respawn fallback (forbidden by BG02).

**Why not `atexit`/daemon-thread-only:** `atexit` and daemon flags cannot close workbooks/Quit (those need the STA thread and the event loop); lifespan is the only place that can `await` the worker and then join. Daemon=True is retained only so a *documented* hard-kill path cannot hang the interpreter.

---

### Q6 — Proxy lifetime

- **Invariant:** no COM proxy (`Application`, `Workbook`, `Worksheet`, `Range`, `VBProject`, …) is ever returned from `submit()`, stored in `SessionService`, or referenced by `api/*`. `submit()` returns only **Pydantic DTOs / primitives** (`domain/models.py`, which already forbid COM objects by design: `DomainModel` is strict JSON). `[observed: models.py docstring]`.
- **How DTOs cross the boundary:** each `fn` runs entirely on the STA thread, reads COM, builds a `BaseModel` (e.g. `WorkbookInfo`, `SessionStatus`), and returns it. `concurrent.futures.Future` carries the DTO (a plain Python object) to the awaiting loop thread. Adapters (`ExcelAppAdapter`, `WorkbookAdapter`) are the **only** code that touches proxies, and they are instantiated and used inside the worker thread (they hold no cross-thread refs; they receive `app`/`wb` as locals per call or hold a ref that is only ever touched from the STA thread).
- **Leak detection / discard:** worker keeps a bounded set of owned workbook identities (`workbook_id → name`) for cleanup, and a counter of live proxy refs (best-effort: `sys.getrefcount(app)` before/after, diagnostic only). If a `submit()` result contains a non-JSON-serializable object (proxy leaked into a DTO), the adapter's return-boundary guard (`_ensure_dto`) raises `TypeError`/`ExcelVbaMcpError` and the worker records `PROXY_LEAK_SUSPECTED`, drops the item, and does **not** place the leaked proxy anywhere. A `gc.collect()` after shutdown releases residual cycles.
- **Test invariant (task_011):** after `VBA_PROJECT_LOCKED`-style failures and after shutdown, `EXCEL.exe` count for the owned PID returns to 0 (observed true in every PoC run) and no proxy is reachable from `SessionService`/`api` (assert via a sentinel object that raises on attribute access).

---

### Q7 — Coexistence with the addendum (owned instance never visible/foreground)

- **Compile:** worker treats `vba_compile` as an ordinary `submit()` with `section="diagnostic"`. On 16.0.14332 the compile service returns `COMPILE_INDETERMINATE(reason=NON_ACTIVATING_COMPILE_UNAVAILABLE)` **without** expanding the Debug menu or invoking UIA; `app_adapter` therefore never sets `Visible=True` for compile. The worker's hidden guarantee is preserved across compile/macro by construction (no code path sets `Visible=True`). `[observed: hidden-compile-research]`
- **Macro:** `macro_run` is `section="execute"`; it runs `Application.Run` on the hidden owned instance. If the macro shows UI (`MsgBox`/`UserForm.Show`), the call blocks; the bounded `submit` timeout → `needs_recovery` (no COM is sent while blocked; the dialog is **never** dismissed — BG09). Preflight/`background_safe` classification (task_022) blocks known UI-dependent macros **before** submit, so they never reach the worker in background mode (BG06). The worker itself adds a runtime guard: if `app.Visible` becomes `True` or an owned top-level window becomes foreground (observed via a read-only window check in the timeout path), the worker records `FOREGROUND_VIOLATION` diagnostics and refuses further non-`close` submits.
- **Open/save/close:** `Visible=False` set at creation and never changed; `DisplayAlerts`/`ScreenUpdating` scoped; `.Activate`/`.Select` forbidden. This directly satisfies BG03/BG04 for all session ops.
- **No VM/isolated session (BG01/BG02):** the worker runs under the current logged-in user in the current interactive session (the MCP process itself), owns one `DispatchEx` instance, and never attaches to the user's instance. This is ownership separation, not environment isolation (addendum §4 task_010 row).

---

### Q8 — Concrete module sketch

> Signatures only; no production code. Names match `plan.md` targets and existing DTOs/errors. `Section`/`OperationClass` reuse `domain/statuses.py`.

#### `adapters/excel/com_worker.py`

```python
Section = Literal["open", "close", "save", "diagnostic", "execute"]

@dataclass
class _WorkItem:
    fn: Callable[[], Any]
    future: Future[Any]
    deadline: float           # time.monotonic()
    section: Section
    op_class: OperationClass  # READ_ONLY / MUTATING / EXECUTING

class ExcelComWorker:
    """Single dedicated COM STA thread; serialized FIFO; awaitable submit."""

    def __init__(self, app_adapter: ExcelAppAdapter, *, orphan_timeout_s: float = 10.0) -> None: ...

    # lifecycle (owning side / loop thread)
    async def start(self) -> None: ...
        # spawn thread, await a "ready" event (worker signals after CoInitializeEx + DispatchEx + Visible=False)
    async def submit(self, fn, *, timeout: float, section: Section = "diagnostic",
                     op_class: OperationClass = OperationClass.READ_ONLY) -> Any: ...
        # fail fast if not started / stopped / NEEDS_RECOVERY; queue _WorkItem; await asyncio.wrap_future
    def get_status(self) -> WorkerStatus: ...   # NON-COM snapshot: state, active token/section, last_hresult
    async def shutdown(self) -> None: ...        # ordered §5 sequence; idempotent
    async def reset(self) -> None: ...           # rebuild owned instance on same STA thread (needs_recovery exit)

    # worker thread (internal)
    def _run(self) -> None: ...                  # CoInitializeEx → loop → CoUninitialize
    def _drain_close(self) -> None: ...          # close owned wbs, Quit, release (steps 4-7)

class WorkerStatus(DomainModel):                 # serializable, no COM
    state: WorkbookSessionState                  # STARTING/RUNNING/NEEDS_RECOVERY/CLOSING/CLOSED/ERROR
    accepting: bool
    active_section: Section | None
    active_operation: str | None
    last_hresult: int | None
    owned_pid: int | None
    last_error: str | None
```

#### `adapters/excel/app_adapter.py`

```python
class ExcelAppAdapter:
    """Centralized non-activating Excel settings; all methods run on the STA thread."""

    def create_hidden(self) -> None: ...         # DispatchEx, Visible=False, record owned_pid/caption
    def quit(self) -> None: ...                  # owned instance only
    def close_workbook(self, wb, *, save_changes: bool = False) -> None: ...
    def open_workbook(self, path: str, *, read_only: bool, update_links: int = 0,
                      notify: bool = False, add_to_mru: bool = False): ...  # returns proxy (STA-local)

    @contextlib.contextmanager
    def alerts_suppressed(self): ...             # save/restore DisplayAlerts in finally
    @contextlib.contextmanager
    def screenupdating_off(self): ...            # save/restore ScreenUpdating in finally (optimization only)
    @contextlib.contextmanager
    def automation_security(self, mso: int): ... # save/restore AutomationSecurity in finally

    def assert_non_intrusive(self) -> None: ...  # runtime guard: Visible==False, owned fg check
```

No `.Activate`/`.Select`/`ShowWindow`/`SetForegroundWindow`/`SendKeys` anywhere in this class (documented invariant + test).

#### `application/session_service.py` (task_012)

```python
class SessionService:
    """Hidden workbook open/close/save + session state; owns no COM proxies."""

    def __init__(self, worker: ExcelComWorker, app_adapter: ExcelAppAdapter,
                 path_policy: PathPolicy, approval: ApprovalPolicy, audit: AuditLogger) -> None: ...

    async def open_workbook(self, path: str, *, mode: SessionMode,
                            timeout_s: float = 60.0) -> WorkbookInfo: ...
        # PathPolicy.validate_path (canonical) -> lock -> materialize work/ copy ->
        # worker.submit(open section, MUTATING, alerts_suppressed, automation_security) -> WorkbookInfo DTO
    async def close_workbook(self, workbook_id: str, *, save_changes: bool = False,
                             timeout_s: float = 60.0) -> SessionStatus: ...
        # worker.submit(close section, MUTATING) ; default SaveChanges=False
    async def save_workbook(self, workbook_id: str, *, timeout_s: float = 60.0) -> WorkbookInfo: ...
        # worker.submit(save section, MUTATING) ; explicit save only
    async def get_status(self, workbook_id: str | None = None) -> SessionStatus: ...
        # merges worker.get_status() (non-COM) with session bookkeeping; no submit needed
    async def list_workbooks(self) -> list[WorkbookInfo]: ...   # from bookkeeping, no COM
```

**Error/status mapping table (worker → domain):**

| Worker condition | Exception raised | `ErrorCode` | `WorkbookSessionState` |
|---|---|---|---|
| `submit` before start / after stop | `WorkerUnavailableError` | `EXCEL_HUNG` | `ERROR` / `CLOSED` |
| `submit` while `NEEDS_RECOVERY` | `ExcelHungError` (fail fast) | `EXCEL_HUNG` | `NEEDS_RECOVERY` (`"running/needs_recovery"`) |
| `wait_for` queue/exec timeout | tool returns structured result | `EXCEL_HUNG` | `NEEDS_RECOVERY` |
| `RPC_E_CALL_REJECTED` (read-only, after ≤3 retries) | `RpcCallRejectedError` | `RPC_E_CALL_REJECTED` | `RUNNING` |
| `RPC_E_CALL_REJECTED` (mutating/executing, 1st) | `RpcCallRejectedError` (no retry) | `RPC_E_CALL_REJECTED` | `RUNNING` |
| `RPC_E_SERVERCALL_RETRYLATER` | `ComBusyError` | `COM_BUSY` | `RUNNING` |
| `com_error` other HRESULT | `error_from_hresult(hresult, msg)` | mapped via `_HRESULT_TO_CODE` | `ERROR` |
| `VBA_PROJECT_LOCKED` (from adapter) | `VbaProjectLockedError` | `VBA_PROJECT_LOCKED` | `RUNNING` |
| foreground/visible violation | `ExcelHungError` | `EXCEL_HUNG` (+`FOREGROUND_VIOLATION` note) | `NEEDS_RECOVERY` |
| clean shutdown | (none) | — | `CLOSING` → `CLOSED` |

All codes already exist in `domain/errors.py::ErrorCode`; `error_from_hresult` already builds the typed subclass. The worker adds only `WorkerUnavailableError` (or reuses `ExcelHungError`) — keep it minimal.

---

## 4. Options summary (Q1 focus, one line each)

- **(a) dedicated STA thread + `SimpleQueue` + `Future` + `asyncio.wrap_future`** — single-STA strong, loop-safe, controllable pumping. **Chosen.**
- (b) `anyio.to_thread.run_sync` per call — breaks single-STA (ARCH-009). Rejected.
- (c) anyio worker-per-session — apartment not guaranteed. Rejected.
- (d) second event loop in the STA thread — heavier, no benefit. Rejected.

---

## 5. Test strategy (feeds task_011/task_012/task_013)

1. **Queue ordering & serialization (unit, no Excel):** two `submit`s complete in FIFO order; `op_class` tag honored; `_WorkItem` deadline skips expired items (no `fn` invocation).
2. **Awaitable non-blocking (unit):** a 20 s `fn` submitted from a loop; `tools/list` still responds (mirror the ARCH-009 probe). `asyncio.wait_for` fires; loop never stalls.
3. **Timeout → `needs_recovery` (unit):** `submit(fn, timeout=0.1)` with a slow `fn`; result is `EXCEL_HUNG` + state `NEEDS_RECOVERY`; a second `submit` fails fast without starting; a `close`-section item is allowed through.
4. **Shutdown ordering (unit):** `shutdown()` while idle → no orphan; `shutdown()` with a running `fn` → drain waits (bounded), closes owned wb, `Quit`, `CoUninitialize`, join ≤ orphan_timeout; second `shutdown()` is a no-op.
5. **COM busy/rejected (integration, mocked HRESULT):** READ_ONLY retried ≤3; MUTATING not retried (assert call count == 1) — task_011 `test_com_busy_does_not_replay_mutation`.
6. **Owned-only shutdown (integration, task_011):** with a user Excel instance open, `shutdown()` leaves the user PID/window untouched and closes only the owned instance; owned PID gone after orphan poll.
7. **Proxy leak guard (unit):** a `fn` returning a fake proxy raises at the return boundary; worker records `PROXY_LEAK_SUSPECTED`, drops the item, never stores the proxy.
8. **Non-intrusive regression (integration, task_011/012/013):** owned app `Visible==False`, `DisplayAlerts`/`ScreenUpdating` restored after success **and** exception, foreground unchanged (BG03), no owned visible window (BG04), `.Activate`/`.Select` never called (static grep + runtime guard).

---

## 6. Risks (for the register / task_010 acceptance)

- **R-01 (HIGH) — Hung call cannot be preempted.** A macro/modal blocks the STA thread; timeout yields `needs_recovery`, but the thread is stuck until `fn` returns or `reset()`/`shutdown()` runs. Mitigation: bounded `submit`, fail-fast in recovery, `close`-section drain, PID-scoped cleanup; document that SIGTERM/hard-kill orphans are recoverable only by killing the **owned** PID.
- **R-02 (MED) — SIGTERM/hard-kill bypasses lifespan.** ARCH-005: no `finally` on those paths; orphan `EXCEL.exe` possible. Mitigation: record `owned_pid` to a runtime file; operator/nextrun cleans owned PID only; never a service/respawn (BG02).
- **R-03 (MED) — Reentrancy if pumping is misused.** Pumping inside `fn` can reenter the worker and duplicate mutations. Mitigation: pump only on idle queue wait + retry backoff; code review + test that `fn` is never reentered.
- **R-04 (MED) — `DisplayAlerts` restore not guaranteed on a stuck app.** A hung `fn` never reaches its `finally`. Mitigation: `reset()` rebuilds the instance (clean state) instead of restoring a setting on a stuck app; BG08 is satisfied by "restored or instance replaced".
- **R-05 (LOW) — `DispatchEx` identity/affinity drift on other builds.** Verified on 16.0.14332 only. Mitigation: affinity probe semantics (ARCH-003) recorded per run; other builds are a documented limitation (task_037).
- **R-06 (LOW) — `asyncio.wrap_future` requires the loop thread.** `submit()` is loop-affine; calling it off-loop would raise. Mitigation: `submit` is only ever called from `api/*` (loop) — enforced by convention + the ARCH-009 rule.

---

## 7. Recommendation (one paragraph)

Adopt **option (a): one dedicated `threading.Thread` running `CoInitializeEx(COINIT_APARTMENTTHREADED)` + a `SimpleQueue` FIFO + a `concurrent.futures.Future` per item, awaited in `submit()` via `asyncio.wrap_future` under `asyncio.wait_for`** — this is the only option that is simultaneously single-STA, event-loop-safe (ARCH-009), and controllable for pumping. `submit(fn, timeout)` fails fast in `NEEDS_RECOVERY`, never cancels a running call, and maps timeout to a truthful `EXCEL_HUNG`/`needs_recovery`; hung mutations are never replayed. `ExcelAppAdapter` owns `Visible=False`, narrowly scoped `DisplayAlerts`/`ScreenUpdating` with `finally` restore (or instance replacement via `reset()`), and forbids `.Activate`/`.Select`. The worker is owned by `MCPServer(lifespan=...)`; `shutdown()` runs the ordered close-owned-workbooks → `Quit` → release → `CoUninitialize` → bounded join on the STA thread, on stdin-EOF and normal return. Compile stays `COMPILE_INDETERMINATE` and macro UI is blocked, so the owned instance never becomes visible/foreground.

---

## 8. References

**Microsoft Learn / Support:**
- `Application.Visible` — https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible
- `Application.DisplayAlerts` (suppresses certain alerts; Excel picks a default response) — https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
- `Application.ScreenUpdating` (stops repaint; not a visibility guarantee) — https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating
- *Considerations for server-side Automation of Office* (interactive desktop assumed; modal hangs) — https://support.microsoft.com/en-gb/visio/considerations-for-server-side-automation-of-office

**Project evidence (reused):**
- task_002 PoC: `.ai/tmp/excel_instance_poc_20261119_215845.txt` (`CONDITIONAL_GO`; affinity `0x8001010E`; no orphan; user foreground unchanged).
- task_002 planner v3 (pump strategy, affinity, lifecycle): `.ai/plans/execution/task_002_planner_v3.md`.
- Architecture audit (ARCH-003/005/009/013): `.ai/audit/architecture-audit-2026-11-19.md`.
- Hidden-compile research (compile indeterminate; modal blocks): `.ai/audit/hidden-compile-research-2026-11-19.md`.
- Macro-run compile-on-demand/modal: `.ai/tmp/probe_run_compile_error_result.json`, `.ai/tmp/probe_run.json`.
- Domain contracts: `src/excel_vba_mcp/domain/{models,statuses,errors}.py` (DTOs, `OperationClass`, `ErrorCode`, `error_from_hresult`).
- Path policy: `src/excel_vba_mcp/security/path_policy.py` (canonical path for COM).
- SDK: `MCPServer(..., lifespan=...)`, `mcp/server/mcpserver/utilities/func_metadata.py:149-164` (`anyio.to_thread.run_sync` for sync tools) `[sdk]`.

**Confidence note:** Q1/Q2/Q5/Q6 mechanics are `[sdk]`/`[observed]`; Q3 pump placement is `[inference]` from ARCH-009 reentrancy analysis + planner-v3 choice; Q4 restore-on-stuck-app is `[inference]` from the impossibility of running `finally` on a hung thread (hence `reset()`).