# Orphaned hidden Excel processes — defect report & fix plan

**Task:** `FE-17` (orchestration follow-up)
**Date:** 2026-11-22
**Severity:** HIGH — resource leak; leaves hidden `EXCEL.EXE` processes after test runs / server shutdown
**Status:** Confirmed, fix designed (see §3)

---

## 1. Observed symptom

During the FE-00…FE-15 session, orphaned hidden `EXCEL.EXE` processes repeatedly appeared after test runs
(e.g. PID 2996 started 20:21, PID 23420 started 20:34, PID 13236). All had `MainWindowHandle = 0`
(i.e. **no visible window**) — exactly the "hidden orphan" class. They were leftovers from
`tests/integration_excel/test_mcp_error_surfacing_matrix.py` and `test_mcp_full_access_workflow.py`,
which spawn a real `python -m excel_vba_mcp` STDIO server per test; the server's owned Excel was not always
reaped before the next test asserted "user Excel PIDs survive".

The presence of these orphans caused a **test-isolation defect**: the e2e workflow test's final assertion
(`pre-existing user Excel PIDs survived`) misclassified a lingering owned PID as a user instance and failed
(`AssertionError: pre-existing user Excel PID(s) [13236] were terminated by the test`).

A cleanup pass (`Stop-Process` on hidden Excel) restored a clean state, and a bounded reap barrier was added to
the matrix test harness as a **test-level workaround**. That workaround is not sufficient: the **product** still
does not guarantee its owned Excel dies on shutdown.

## 2. Root cause (architecture)

The owned Excel instance is created in `ExcelAppAdapter.create_hidden()` via
`win32com.client.DispatchEx("Excel.Application")`; its PID is read from the window `Hwnd`
(`app_adapter.py:create_hidden`, `owned_pid`).

Teardown happens in `ExcelWorkerThread._drain_close()` (`com_worker.py`), which:

1. closes owned workbooks best-effort,
2. calls `adapter.quit()` — which is `app.Quit()` inside a `try/except Exception: pass`
   (`app_adapter.py:quit`), and
3. immediately sets `self._owned_pid = None` (`com_worker.py:_drain_close` and `shutdown`).

`Excel.Application.Quit()` is **not a reliable process terminator**:

- it can raise and is swallowed (`except Exception: pass`),
- it can return while the process is still alive (e.g. an open workbook, a pending macro, a modal dialog, or a
  COM reference keeps `EXCEL.EXE` resident),
- and there is **no PID-based fallback** (no `taskkill`/Win32 `TerminateProcess`) and **no post-quit liveness
  check** anywhere on the shutdown path.

Consequently, if `Quit()` does not fully terminate Excel, the worker drops `_owned_pid` and **nothing ever kills
the process**. The only `taskkill` in the codebase is in `macro_service.py` and only for a macro-run modal
dialog — it does not cover general shutdown.

### 2.1 Why the worker thread being a daemon makes it worse

The worker is a daemon thread. If `shutdown()` times out joining the thread
(`com_worker.py:shutdown` → `thread.is_alive()` → state `ERROR`), the process can exit with the daemon thread
still holding a live COM reference, leaving Excel orphaned.

## 3. Fix design (minimal, defence-in-depth)

Make "the owned Excel is dead" an **enforced invariant** of teardown, not a best-effort side effect. Keep the
existing `Quit()` first (clean COM shutdown), then add a bounded, PID-scoped fallback.

### 3.1 `ExcelAppAdapter.quit()` → `shutdown()` semantics

Change `ExcelAppAdapter.quit()` (and the `_drain_close`/`shutdown` callers) so that after `app.Quit()`:

1. **Capture** the owned PID *before* quitting (currently discarded).
2. **Verify death** with a bounded poll (e.g. up to ~5 s, 50–100 ms interval) using a PID-liveness check
   (Win32 `OpenProcess`/`GetExitCodeProcess`, or `tasklist /FI "PID eq <pid>"`).
3. If still alive after the grace period, **terminate only that PID** (`taskkill /F /PID <pid>`, already the
   pattern used in `macro_service.py`) and re-verify.
4. Only then clear `_owned_pid`.

No user PID is ever touched: the fallback targets **only** the adapter's own `owned_pid`.

### 3.2 Worker teardown ordering

In `com_worker.py:_drain_close` and `shutdown`:

- do **not** set `_owned_pid = None` until after the verified shutdown above;
- if `shutdown()` times out joining the thread, perform the same PID-scoped fallback on the main thread as the
  last resort (the PID is the authoritative identity; `Quit()` from a dead STA thread is not needed to kill the
  process).

### 3.3 Keep it non-intrusive

- The fallback must never make the owned instance visible (it kills, not shows).
- No registry / Trust Center / `AutomationSecurity` changes.
- Reuse the existing `taskkill` invocation style; keep English comments.

## 4. Acceptance criteria

1. `ExcelAppAdapter.quit()`/`shutdown()` guarantees the owned PID is dead before returning, or the process is
   terminated by the fallback; `_owned_pid` is cleared only after verification.
2. Unit tests (no Excel): a fake app whose `Quit()` is a no-op and a fake PID-liveness source prove the
   fallback terminates the PID and that a `Quit()` that works short-circuits the fallback.
3. Live integration: after `excel_close` / server shutdown, `excel_health` reports no owned PID and the process
   is reaped (the existing e2e PID assertions must pass **without** the test-level reap barrier).
4. Existing suites stay green: 872 unit, 5 e2e (both orders), 0 orphaned Excel.
5. Fixtures byte-identical; no production behavior change for the happy path.

## 5. Out of scope

- Worker-lifecycle hardening beyond PID reaping (F-2 / R-17 health liveness probe) — documented separately.
- Reworking `macro_service`'s modal-dialog `taskkill` (already correct, just narrow).