# task_002 — Post-Execution Validator Report (independent re-validation)

**Date:** 2026-11-19
**Role:** Validator (independent post-execution verification)
**Subject:** `task_002` PoC — `tests/integration_excel/test_excel_instance_poc.py` (commit `bc881ed`)
**Binding spec:** `.ai/plans/execution/task_002_planner_v3.md` (N1–N5, §4–§8)
**Plan AC:** `.ai/plans/plan.md` task_002 acceptance criteria (lines 156–162)
**Latest artifacts reviewed:** `.ai/tmp/excel_instance_poc_20261119_174032.json|.txt` (Implementor run) plus my own fresh runs below.

**Read-only discipline:** No source/test/plan file was modified. Only tests were executed and artifacts read.

---

## 1. Scope

Independently confirm the task_002 PoC proves, on the real target (Windows 10 x64, Excel 16.0.14332):
single dedicated STA ownership of Excel COM, `DispatchEx` process isolation, cross-thread affinity
confinement (negative probe), serialized multi-op on one STA boundary, pinned lifecycle open/close/Quit
with no orphan, and no impact on any user Excel session. Also re-run the regression gates and confirm the
`go_no_go_r01_task002` verdict against planner v3 §8. This is **not** a re-read of the Implementor's claims;
every result below is from my own execution/observation.

---

## 2. Checks performed

1. Read binding spec (planner v3), plan AC, implementation source (1416 lines), and latest artifact JSON.
2. Verified `git` state: `tests/integration_excel/test_excel_instance_poc.py` is byte-identical to commit
   `bc881ed` (`git diff bc881ed HEAD -- <file>` empty; working tree entry clean at validation start).
3. Confirmed no pre-existing `EXCEL.EXE` before running (`Get-Process EXCEL` → count 0).
4. **Independently re-ran** the exact planner command:
   `uv run pytest tests/integration_excel/test_excel_instance_poc.py -v -s` → observed live console.
5. Verified `Get-Process EXCEL` **after** the run → count 0 (no orphan) — my own check.
6. Re-verified artifact internal consistency (matrix, `steps[]`, `serialized_multiop_evidence`,
   `affinity_probe`, `orphan_poll`, `go_no_go`) on my fresh artifacts.
7. Anti-gaming review of the affinity probe and assertions (source + faulthandler evidence).
8. Regression gates: `uv run pytest` (full), `uv run ruff check src tests`,
   `uv run ruff format --check src tests`, `uv run mypy src`.
9. Environment-fact verification (Windows 10 vs Windows 11; VBOM/VBE readability).

---

## 3. Actual results (my own observations)

### 3.1 Independent PoC run (console observed live)

Command: `uv run pytest tests/integration_excel/test_excel_instance_poc.py -v -s`
Result: **2 passed in 23.99s** (no skips). Fresh artifacts written by the run:
`.ai/tmp/excel_instance_poc_20261119_174458.json|.txt` and `...174505.json|.txt`.

| Test | STA thread id (init) | Owned PID | Owned HWND | Affinity probe thread | Affinity HRESULT | Orphan |
|------|----------------------|-----------|------------|-----------------------|------------------|--------|
| `test_dispatch_ex_creates_isolated_excel_instance` | 24436 | **23804** | 6621220 | 21648 | **0x8001010E** (RPC_E_WRONG_THREAD) | gone (poll 1) |
| `test_excel_workbook_open_close_lifecycle` | 24436 | **2636** | 5639394 | 14160 | **0x8001010E** (RPC_E_WRONG_THREAD) | gone (poll 1) |

- All COM-touching steps in both runs carried `thread_id == 24436` (single distinct value).
- `sta_init` notes: `CoInitializeEx(APARTMENTTHREADED)=S_OK`; `couninitialize` logged on the same thread.
- Affinity probe console line (my run): `[step] affinity_probe thread=24436 status=expected_failure hr=0x8001010E RPC_E_WRONG_THREAD`.
- Python `faulthandler` independently printed `Windows fatal exception: code 0x8001010e` on thread
  `[poc_affinity_probe]` at `test_excel_instance_poc.py:792` (the raw `Invoke`) — OS-level confirmation the
  cross-thread call genuinely faults (see §4 anti-gaming).
- `FINAL go_no_go_r01_task002=CONDITIONAL_GO`.

**Post-run orphan check (mine):** `Get-Process EXCEL` → **count 0**. No orphaned owned PID (23804 / 2636 both gone).

### 3.2 Binding-requirement verification (evidence-backed)

| Requirement (planner v3) | Status | Evidence |
|--------------------------|--------|----------|
| Single dedicated STA thread; same `thread_id` on every COM step | PASS | `steps[].thread_id` all = 24436 in both runs; `sta_thread_id_init=24436`. |
| `CoInitializeEx(APARTMENTTHREADED)` + `CoUninitialize` on that thread | PASS | `sta_init` / `couninitialize` steps both `thread_id=24436`, notes confirm APARTMENTTHREADED / CoUninitialize. |
| No COM on pytest/main thread except intentional affinity probe | **PARTIAL / WARN** | STA thread *is* the pytest main thread (fixture `_init_sta` runs on main thread; no worker thread spawned except the probe). COM therefore runs on the main thread. Affinity confinement still genuinely proven (probe thread 21648/14160 ≠ 24436 fails). See §6 W2. |
| `DispatchEx` as creation method (not `GetActiveObject`) | PASS | `matrix.dispatch_method = "DispatchEx('Excel.Application')"`, `test…py:1210`. `GetActiveObject` used only for read-only ROT snapshot (`:522`). |
| Owned instance identified by PID delta; owned PID ∉ baseline | PASS | `owned_pid_resolution=pid_delta_single_new`; baseline_before `excel_processes=[]`; owned 23804/2636 not in baseline. |
| `Application.Hwnd`/`Caption`/`Version`/`Build` readable | PASS | `excel_hwnd=6621220`, `excel_caption="Excel"`, `excel_version="16.0"`, `excel_build="14332.0"`; `OperatingSystem="Windows (64-bit) NT 10.00"`. |
| Affinity probe: separate thread inits COM then fails `RPC_E_WRONG_THREAD 0x8001010E` | PASS | `affinity_probe.performed=true`, `outcome=failed_as_expected_rpc_e_wrong_thread`, `hresult_hex=0x8001010E`, probe thread 21648/14160 ≠ STA 24436, `_probe` calls `CoInitializeEx` before `Invoke`. |
| Serialized multi-op: ≥2 distinct COM ops, same boundary, order+thread recorded | PASS | `serialized_multiop_evidence`: ops = [read_application_version, read_application_caption], `submission_completion_order=[op1,op2]`, `sta_thread_id=24436`, `same_boundary=true`; `steps` has `serialized_op_1_read_version`/`serialized_op_2_read_caption`. |
| Lifecycle: temp workbook only; pinned `Workbooks.Open(UpdateLinks=0, ReadOnly=False, Notify=False, AddToMru=False)` | PASS | `matrix`/steps `wb_open` params exactly `{'UpdateLinks':0,'ReadOnly':False,'Notify':False,'AddToMru':False}`; temp dir `…\Temp\poc_excel_…`; assertion `params == {…}` passes. |
| `SaveChanges=False` on closes | PASS | `wb_close_after_save` / `wb_close_after_open` notes `SaveChanges=False`. |
| No repo files touched | PASS | Temp workbook under `%TEMP%\poc_excel_*`; no repo path in any step; `tmp.cleanup()` in `finally`. |
| Cleanup: owned WB closed, `Quit`, refs released, `CoUninitialize`; timed orphan poll | PASS | `app_quit ok`, `release_refs_gc ok`, `couninitialize ok`, `orphan_poll_1 status=gone` (poll_count=1). |
| No orphaned EXCEL.exe (my check) | PASS | `Get-Process EXCEL` count 0 after run. |
| User session not harmed; no unrelated process kill | PASS | `grep` found no `kill/terminate/taskkill/os.system/TerminateProcess/Stop-Process`; baseline_before empty; no user PID to harm; ROT snapshot read-only. |
| `go_no_go` matches §8 rubric; `CONDITIONAL_GO` correct (not overclaimed) | PASS | All §8.1 disqualifiers absent; §8.2(b) "not exercised" path → CONDITIONAL_GO for lifecycle/isolation/STA-core only; follow-ups task_004/005/010 listed. See §5. |
| N3/N4/N5 fields present | PASS | N3: `app_visible=false`, `app_display_alerts=false`, `workbook_open_params`. N4: `process_snapshot_method`, `window_snapshot_method`, `snapshot_partial_failures=[]`. N5: `steps[]` with name/thread_id/timestamp/status/notes (+hresult). |

### 3.3 Environment facts (recorded accurately; no Windows 11 claim)

- Artifact `matrix`: `platform="Windows-10-10.0.17763-SP0"`, `windows_display.current_build="17763"`, `ubr="9020"`,
  `windows_version_tuple=[10,0,17763,2]`. Excel `16.0` build `14332.0`, `office_bitness=x64`, Python `3.14.7` x64,
  pywin32 `312`.
- No artifact contains "Windows 11" / "win11" (grep confirmed).
- `README.md:11`: "Windows 10/11 x64 (spec target Windows 11; validated on Windows 10 Enterprise LTSC 2019, build 17763, x64)" — accurate.
- VBOM trust: `vbe_ui_language.methods` includes `VBE.MainWindow.Caption='Microsoft Visual Basic for Applications'`
  (VBE object reachable) → VBOM trust enabled; UI language `1033` confirmed via `LanguageSettings.LanguageID(2)`.
  Recorded accurately with confidence `confirmed_via_LanguageSettings`.

### 3.4 Regression gates

| Gate | Result | Attribution |
|------|--------|-------------|
| `uv run pytest` (full) | **10 passed** | Clean at my run time. |
| `uv run ruff check tests/integration_excel/test_excel_instance_poc.py` | **All checks passed** | task_002 scope clean. |
| `uv run ruff format --check tests/integration_excel/test_excel_instance_poc.py` | **1 file already formatted** | task_002 scope clean. |
| `uv run ruff check src tests` (aggregate) | **23 errors (UP042)** | All in **untracked** `src/excel_vba_mcp/domain/statuses.py` (concurrent task_007 work). `git log` shows no history; file mtime 17:45:24 — created *during* this validation window by the main agent. Not task_002. |
| `uv run ruff check <tracked src files>` | **All checks passed** (12 files) | Bootstrap src (at bc881ed) is clean. |
| `uv run ruff format --check <tracked src files>` | **12 files already formatted** | Bootstrap src clean. |
| `uv run mypy src` | **Success: no issues found in 13 source files** | Clean. |

**Conclusion on gates:** At the task_002 commit (`bc881ed`) the repo-wide `ruff`/`format`/`mypy`/`pytest` gates
were clean; the only current aggregate failures come from the main agent's concurrent untracked `src/domain/*.py`
(task_007), which is out of task_002 scope and did not exist at commit time.

---

## 4. Anti-gaming checks (all passed)

- **Affinity probe is real, not hardcoded.** `_affinity_probe` (`:757`–`847`) obtains the raw
  `app._oleobj_` (PyIDispatch), starts a real `threading.Thread` (`:803`), calls `CoInitializeEx` on it (`:786`),
  then `raw.GetIDsOfNames("Version")` + `raw.Invoke(..., DISPATCH_PROPERTYGET, 0)` (`:792`–`793`). The HRESULT
  is the actual COM return, not a literal. Python's `faulthandler` independently reported the OS-level exception
  `0x8001010e` on `[poc_affinity_probe]` at line 792 — impossible to fake without a real cross-thread COM fault.
- **"Expected failure" is genuinely cross-thread.** Probe thread id (21648 / 14160) ≠ STA thread id (24436); the
  failure is raised by `Invoke`, not by a `raise` statement. Probe initializes COM first, so it yields
  `RPC_E_WRONG_THREAD` (0x8001010E), not `CO_E_NOTINITIALIZED` (0x800401F0) — the exact discrimination ARCH-003 required.
- **Assertions are meaningful (no `assert True`).** Verified assertions: `dispatch_method == "DispatchEx('Excel.Application')"`;
  `owned_pid is not None`; `owned_pid not in baseline_pids`; `all(step.thread_id == sta_thread_id)`;
  `multiop["same_boundary"] and len(ops) >= 2`; `affinity["performed"] is True`;
  `affinity["outcome"] == "failed_as_expected_rpc_e_wrong_thread"`; `affinity["hresult_hex"] == "0x8001010E"`;
  `not (before_pids - after_pids)`; `lifecycle["ok"] is True`; `params == {…}`; `opened_name == saved_name`;
  `added_name == "Book1"`; per-step coverage loop. All non-trivial.
- **Skip logic does not hide failures.** `_excel_com_available()` returned available on this host (tests ran, not
  skipped). Skip is gated only on `sys.platform`/`pywin32` import; a broken Excel would make `DispatchEx` raise and
  fail the test, not skip it.
- **Evidence is produced by the run, not hand-edited.** My independent runs produced **different PIDs** (23804/2636)
  and **different probe thread ids** (21648/14160) than the Implementor artifact (PID 25552, probe thread 11872), while
  reproducing identical HRESULT/structure. Hardcoded evidence could not vary per run. `.txt` twins were emitted for
  each run (174458/174505, plus 174531/174536 from the full-suite run).

---

## 5. Go/No-Go rubric evaluation (planner v3 §8)

- **§8.1 disqualifiers → No-Go:** none present (no orphan; no user-session harm; no COM off-STA except the intentional
  probe; probe failed-as-expected; DispatchEx used; required matrix fields present; multi-op present).
- **§8.2 Conditional Go:** satisfied. Affinity probe performed and recorded failed-as-expected. Serialized multi-op
  recorded. Pinned lifecycle params recorded. Snapshot methods/partial failures recorded. RPC/modal **not exercised**
  → §8.2(b) "Conditional Go for task_002 lifecycle/isolation/STA-core evidence only" with follow-ups.
- **§8.3 Full Go:** not claimed — no RPC/modal observation context, so Full Go is correctly withheld.
- **Verdict `CONDITIONAL_GO` is correct and not overclaimed.** `follow_ups` lists task_004 (VBE compile + UIA modal/busy),
  task_005 (Application.Run modal/RPC), task_010 (full MCP request-queue serialization). `modal_rpcs_exercised=false`,
  `rpc_observed=false` — consistent with the "not exercised" note.

---

## 6. Defects / Warnings

No Critical or High defects. The PoC meets every plan task_002 acceptance criterion and reproduces reliably.
Non-blocking items:

- **W1 (MEDIUM, evidence gap vs binding spec) — missing explicit alignment field.** planner v3 §1 (line 26) requires the
  artifact to state "Aligned to com_worker single-STA design: yes/no (with notes)" and §11 (line 253) lists
  "alignment yes/no" among artifacts to record. The artifact/matrix has **no** `aligned_to_com_worker_single_sta` key
  (grep of source and JSON found only prose in rationale/scope_note). Alignment is *implicitly* yes (whole PoC is
  single-STA), and this does **not** trigger an §8.1 No-Go (alignment is not in the §8.1 enumerated field list), but the
  explicit required field should be added. Location: matrix built in `test_excel_instance_poc.py:1177`–`1262`; no such key.

- **W2 (LOW–MEDIUM, design fidelity) — STA thread is the pytest main thread, not a separate worker.** `sta_context`
  fixture calls `_init_sta` (`:1187`–`1191`, `_init_sta` at `:623`) directly on the pytest main thread; the only spawned
  thread is the affinity probe (`:803`). Consequence: all COM runs on the pytest/main thread, which is *literally* at odds
  with planner §1 "no access from pytest/main thread except the harness API." This does **not** break any plan AC
  ("COM proxies remain in the creating STA thread" is proven — the proxy is confined to thread 24436 and thread 21648/14160
  fails), but it does not model the production `com_worker` pattern of a *spawned* dedicated worker thread with main-thread
  marshalling. Recommend a follow-up (task_010) to exercise a spawned STA worker.

- **W3 (LOW, cosmetic) — faulthandler noise.** The real affinity fault prints `Windows fatal exception: code 0x8001010e`
  and the post-Quit proxy release prints `Windows fatal exception: code 0x80010108` (RPC_E_DISCONNECTED at
  `del app`, `:1022`). Benign (process not crashed; orphan still reaped; steps[] correct) but the 0x80010108 release is not
  captured in `rpc_observations`. Consider suppressing faulthandler for these expected COM faults in production harness.

- **W4 (LOW, cosmetic) — `paths` field always empty in persisted JSON.** `persist()` computes `payload = asdict(self)`
  (`:211`) before assigning `self.paths` (`:226`), so the JSON records `"paths": {}`. Cosmetic; file names are known from
  the run log.

- **W5 (LOW, robustness) — `tracked_workbooks` safety net unused.** `_lifecycle_temp_xlsx` (`:853`–`916`) closes its own
  `wb`/`wb2` but never appends to `ctx.tracked_workbooks`, so the cleanup loop (`:985`) iterates an empty list. Safe here
  because `app.Quit()` closes everything and orphan poll confirms reap, but the intended "close owned workbooks in finally"
  net is not actually exercised. Minor.

- **W6 (INFO, environment) — shared tree mutated by parallel agents during validation.** `src/excel_vba_mcp/domain/{statuses,errors,models}.py`
  are untracked (task_007, mtimes 17:45–17:46) and `execution-log.md` is modified (task_003/task_007 status). This is the
  sole cause of the aggregate `ruff check src tests` failure (UP042 on `statuses.py`). Not a task_002 regression; task_002
  scope + tracked src are ruff/format/mypy clean.

---

## 7. Unmet acceptance criteria

**None.** Plan task_002 acceptance criteria (plan.md:156–162) are all satisfied and independently reproduced:

1. Separate Excel instance created and identified — PASS (DispatchEx; PID delta 23804/2636; HWND; version readable).
2. Existing user Excel session unaffected — PASS (baseline empty; no kill logic; ROT read-only; no user PID touched).
3. open/close/Quit leave no orphaned owned Excel process — PASS (orphan=false, poll_count=1; my `Get-Process EXCEL`=0).
4. COM proxies remain in the creating STA thread — PASS (affinity probe on a different thread fails 0x8001010E).
5. Message-pump behavior and limitations documented — PASS (`message_pump_strategy` recorded; RPC/modal "not exercised" stated → CONDITIONAL_GO).

Binding planner v3 N1–N5: all satisfied (see §3.2). The only binding-spec shortfall is the W1 explicit alignment field
(documentation), which does not affect technical validity or the correctness of the `CONDITIONAL_GO` verdict.

---

## 8. Final result

**PASS WITH WARNINGS**

The task_002 PoC is technically sound and independently reproduced on the real target (Windows 10 x64 build 17763.9020,
Excel 16.0.14332). All plan acceptance criteria and binding N1–N5 requirements are met; `go_no_go_r01_task002=CONDITIONAL_GO`
is correct and not overclaimed; no orphan; user session untouched; regression gates are clean for task_002 scope.
Warnings are non-blocking documentation/design-fidelity items (W1–W6) and do not invalidate the gate.

**Required follow-ups (non-blocking):**
- Add explicit `aligned_to_com_worker_single_sta: yes` (+notes) to the artifact matrix (W1; planner v3 §1/§11).
- (Advisory) In task_010, exercise a *spawned* dedicated STA worker thread (not the request/main thread) to fully model the
  production `com_worker` pattern (W2).
- (Advisory) Persist `paths` in the artifact and log the benign RPC_E_DISCONNECTED on proxy release (W3/W4).
- (Advisory) task_004/task_005/task_010 remain the gates that close full R01 (modal/busy paths not exercised here).

---

## Appendix A — Required/Advisory classification (per role output format)

**Approved findings:** none new (this is a post-execution validation, not a design review).

**Rejected findings:** none.

**Merged findings:** W3+W4 (artifact hygiene) share the "persist() ordering / expected-COM-fault logging" root cause;
W2 is a distinct design-fidelity item; W1 is a distinct binding-spec evidence item.

**Rollout analysis:** task_002 is the critical-path gate (`task_001 → task_002 → task_010 → …`). CONDITIONAL_GO correctly
unblocks task_003/004/005/010 with the explicit condition that modal/busy/RPC handling and full MCP queue serialization are
validated downstream. No circular dependency; no rollout conflict. The only sequencing risk is environmental: parallel
agents share the working tree (W6), so per-task quality gates should be run on the task's own files or a clean checkout.

**Execution validation:** targets exist and are current (`test_excel_instance_poc.py` == `bc881ed`); plan not stale;
dependencies valid; architecture consistent (single-STA model aligns with `com_worker` intent). Ready.

**Required fixes (blocking):** none.

**Advisory recommendations:** W1 (add alignment field), W2 (spawned worker in task_010), W3/W4/W5 (artifact/robustness polish).