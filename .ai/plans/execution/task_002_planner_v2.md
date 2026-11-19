# task_002 — Planner output (v2, revised to satisfy Validator FAIL), 2026-11-19

This revision addresses Validator findings (see `task_002_validator.md`): concrete evidence protocol, STA↔`com_worker.py` alignment, environment/VBE discovery, isolation/shutdown/orphan/thread-affinity rules, message pumping + RPC/modal handling, and explicit Go/No-Go rubric. Any gaps discovered by Auditor/Validator in re-review must be fixed in a v3 before implementing `tests/integration_excel/test_excel_instance_poc.py`.

**Source anchors:** `plan.md` task_002/R01/G1; `spec_text.txt` §5.3, §7.1 context, §9.4, §11/11.1, §13.1; prior v1 in `task_002_planner.md` (preserved conceptually) plus Researcher/Auditor inputs.

---

## 1. Production alignment (single STA model) — must be explicit

- The PoC **must** demonstrate the same threading/COM model intended for production `src/excel_vba_mcp/adapters/excel/com_worker.py` (`task_010`): **one dedicated STA thread** owns every COM proxy created/used/released for Excel/VBE in that worker context.
- **PoC harness requirement:** all `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)`, all `win32com.client.DispatchEx`/workbook/`Application` proxy access, all reference releases, `CoUninitialize()`, and workbook close/`Quit` calls happen on **the same OS thread** (record thread id `threading.get_ident()` at init and assert/log it for each COM-touching step). No COM proxies are passed to other threads; no access from pytest/main thread except via the harness API that marshals only non-COM data/logs (if any).
- **If the implemented test helper uses a different pattern** (e.g. implicit MTA, or COM calls on multiple threads), the result is **No-Go for R01** regardless of “Excel opened” success.
- Evidence artifact must state: “Aligned to com_worker single-STA design: yes/no (with notes)”. If “no”, Go is forbidden.

---

## 2. Prerequisites / context (repeated, explicit)

- Windows 11 x64 target, **interactive desktop** (logged-in user session). No Windows Service / RDP-without-interactive-desktop assumptions. Verify: `sys.platform == "win32"`, record `platform.platform()`, `platform.version()`, `sys.getwindowsversion()`.
- Microsoft Excel desktop installed. Python/pywin32/Office bitness recorded; recommended x64 Python x64 Office per project; mismatches recorded as risk (may cause No-Go if COM fails).
- **Serial execution only**; do not run with `pytest -n`/xdist. One Office interaction run at a time.
- No Trust Center/registry changes by PoC. Do **not** touch files under `projects/`, `work/`, `snapshots/` except creating a **runtime temporary** workbook in an OS temp dir (or `.ai/tmp`-scoped temp) for lifecycle test; delete in teardown.
- No SendKeys/coordinate UI automation to dismiss dialogs as a primary strategy. If an unexpected modal/dialog appears: capture screenshot only when present (`pywinauto` Desktop screenshot to `.ai/tmp/excel_instance_poc_<ts>_dialog.png`) and record dialog text/title if readable; manual resolution may be required; do not automate dismissal.

---

## 3. Environment matrix (required evidence) — capture at run start

Record into evidence log/artifact (stderr + file under `.ai/tmp/excel_instance_poc_<YYYYMMDD_HHMMSS>.{json,txt}`; both human + machine-readable fields recommended):

| Field | How to obtain (PoC) | If unavailable |
|-------|---------------------|----------------|
| `timestamp_utc`, `timestamp_local` | `datetime.now(timezone.utc)` / local | — |
| `windows_product`, `windows_build`, `windows_display_version`, `windows_ubr`, `platform` | `platform.platform()`, `platform.version()`, `sys.getwindowsversion()`; optional registry `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion` (`DisplayVersion`, `CurrentBuildNumber`, `UBR`) via `winreg` | record `unknown` + exception text |
| `python_version`, `python_executable`, `python_bitness` | `sys.version`, `sys.executable`, `struct.calcsize("P")*8` | — |
| `pywin32_version`, `mcp_sdk_version` (optional) | `import win32api; win32api.GetFileVersionInfo(...)` or `__version__` patterns if exposed; `mcp.__version__` if importable | `unknown` |
| `excel_version`, `excel_build`, `excel_fullname_or_path_hint`, `excel_operating_system`, `excel_hwnd` (owned), `excel_caption` (owned) | From owned `DispatchEx` instance: `Application.Version`, `Application.Build`, `Application.OperatingSystem`, `Application.Hwnd`, `Application.Caption` (read after STA init, on STA thread) | record `unknown` + HRESULT/error |
| `office_bitness` | Best-effort: Click-to-Run registry `HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\Platform` (`x64`/`x86`); cross-check with Excel process bitness if safely inspectable; document method | `unknown` + methods tried |
| `vbe_ui_language` | **Required attempt (document method + result):** (a) if VBE object accessible on owned instance via COM (`excel.VBE`), record `excel.VBE.MainWindow.Caption` read **only** as a hint and/or attempt to read culture/locale hints from known VBE APIs; (b) additionally record **Windows UI language / Office UI language context**: query `GetUserDefaultUILanguage()`/registry or `excel.Application.LanguageSettings.LanguageID(msoLanguageIDUI=2)` if available — log the numeric ID and whether mapping was applied; (c) if neither is reliably obtainable, set `vbe_ui_language = "not_discoverable"` and list all methods attempted. **Do not** change VBE/Office language. | `not_discoverable` + attempted methods list |
| `sta_thread_id_init` | `threading.get_ident()` in the STA worker/test thread that calls `CoInitializeEx` | — |
| `com_init_result` | Result of `pythoncom.CoInitializeEx(COINIT_APARTMENTTHREADED)`: note `S_OK`/`S_FALSE`/`RPC_E_CHANGED_MODE` semantics in log | exception details |
| `dispatch_method` | Literal `DispatchEx("Excel.Application")` | — |

**VBE language note (explicit):** “VBE language” in AC is often interpreted as UI language of the VBA IDE; exact COM API is environment-dependent. The PoC must **not** claim certainty if only heuristics were used — record method and confidence (`confirmed_via_...` or `heuristic_...` or `not_discoverable`).

---

## 4. Isolation protocol (rigorous) — required

### 4.1 Precondition
- Before any `DispatchEx`, perform **baseline capture (label `baseline_before`)**. Preferred: run with **no user Excel instances** for clearest proof (operator may close Excel). If user Excel is open, baseline **must** capture it (PIDs, windows) and the PoC **must not** mutate any non-owned instance.

### 4.2 Baseline capture (read-only, non-intrusive)
Capture and persist:
- `user_excel_pids_before`: list of `EXCEL.EXE` PIDs (`psutil` if available, else `Get-CimInstance Win32_Process`/`tasklist` parsing) with start time/session if available.
- `user_excel_windows_before`: read-only enumeration of top-level windows owned by those PIDs (title, HWND) — no clicks, no focus stealing beyond enumeration APIs used read-only.
- **ROT/GetActiveObject snapshot (read-only, best-effort):** call `win32com.client.GetActiveObject("Excel.Application")` **only** to record whether any instance is registered and, if so, capture **read-only** observable facts (e.g. workbook count/names via read of `Workbooks.Count`/iterated names in a way that does not save/close/modify — list names only).  
  - **Must not:** `Quit`, `Close`, `Save`, run macros, change `DisplayAlerts` in a way that affects user workflow beyond local read? Better: if attached to a user instance by accident of API behavior, **do not assume ownership** — document. Prefer to treat ROT result as informational; **isolation proof is primarily PID/HWND delta + no mutation** of user PIDs/windows.  
  - If `GetActiveObject` fails (e.g. `MK_E_UNAVAILABLE`), record that explicitly.
- Record `timestamp_baseline_before`.

**Forbidden:** any write/close/save/quit/run on processes/workbooks not owned by the PoC after `DispatchEx` is created; any enumeration that requires modifying user state.

### 4.3 Owned instance creation and identity
- On STA thread: `app = win32com.client.DispatchEx("Excel.Application")`, set `app.Visible = False`, `app.DisplayAlerts = False` (and record; restore not required for owned instance beyond cleanup, but do not touch others).
- **Owned identity (required):**
  - Determine `owned_excel_pid`: compute PID set delta `after_dispatch_pids - user_excel_pids_before` intersected with processes that are the new Excel process; if ambiguous, enumerate Excel processes again and correlate by timing/`Application.Hwnd` if resolvable to a window process — record method. If still `unknown`, **AC for isolation identification fails** (must be recorded; Go is not possible for a clean pass).
  - Record `owned_excel_hwnd = app.Hwnd`, `owned_excel_caption = app.Caption`.
  - Assert `owned_excel_pid` is not in `user_excel_pids_before`.
  - Log `sta_thread_id_init` again for confirmation that all subsequent COM calls use the same thread context (documented).

### 4.4 Post-creation / during-run checks
- After creating owned instance and after each major step (open workbook, close workbook), re-check that no user PID disappeared unexpectedly unless user acted externally (record external changes as “inconclusive for user-session-unaffected” if observed) and no user windows changed titles in unexpected ways (best effort).

### 4.5 Post-cleanup baseline
- After full cleanup (section 5): capture `baseline_after` same method set; compare to `baseline_before`.
- **Pass for “existing user Excel session not touched”:** all PIDs present in `baseline_before` still present **or** the operator explicitly stated they closed them during run (recorded) — but for a strict PoC claim of “not touched by PoC”, the stronger claim is: PoC performed no operations on non-owned COM objects and no user PID disappeared solely due to PoC; any disappearance must be attributed or treated as fail. Default: if user had Excel open and any of those PIDs are gone after PoC cleanup without operator note, mark **FAIL** (isolation/user-session risk).
- If there were **no** user Excel instances before, `baseline_after` must show **no orphaned owned PID** and ideally no extra `EXCEL.EXE` processes beyond normal transient noise? Assert owned PID absent; any other new Excel PIDs require explanation (fail unless clearly external — recorded as inconclusive).

---

## 5. Lifecycle, shutdown, orphans, COM init/uninit — required procedure

### 5.1 Thread and COM init
- Run all COM in a **single dedicated STA thread** (fixture/worker thread in tests). Call `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)` on that thread. Log outcome.
- On teardown, call `pythoncom.CoUninitialize()` on the **same** thread (after releasing all COM references created there). Log.

### 5.2 Workbook/test file policy (lifecycle test)
- Create a unique temporary directory (e.g. `tempfile.TemporaryDirectory()` under system temp) and a file path `poc_lifecycle.xlsx`. Do **not** use any repo files.
- Sequence (owned instance only, on STA thread):
  1. `wb_new = app.Workbooks.Add()` (or documented equivalent) — record result.
  2. `wb_new.SaveAs(path, FileFormat=51)` (xlsx). Record HRESULT/errors.
  3. `wb_new.Close(SaveChanges=False)`. Release `wb_new` reference.
  4. `wb_open = app.Workbooks.Open(path, ...)` with explicit read-only/params as appropriate (document choices): at minimum avoid surprises; prefer `ReadOnly=True` or explicit `UpdateLinks=0, ReadOnly=False` only if necessary — record params used.  
     **SaveChanges policy:** for any workbook opened/saved by PoC, default `SaveChanges=False` on Close unless the step explicitly requires saving — record every Close with `SaveChanges` value.
  5. Verify open (e.g. `wb_open.Name`/`FullName` matches) — read-only checks.
  6. `wb_open.Close(SaveChanges=False)`. Release `wb_open`.

### 5.3 Shutdown (always in `finally`)
- In `finally`, on the **same STA thread**:
  1. Close **only owned** workbooks still referenced (iterate a tracked list; call `Close(SaveChanges=False)`) catching and logging COM errors (do not silently swallow in a way that hides orphans — log then continue cleanup best effort).
  2. Call `app.Quit()` (owned instance). Catch and log errors.
  3. Release all COM references (`del` Python vars / `pythoncom._CoUninitialize` not needed per ref; ensure no remaining strong references in the test scope for those proxies). Call `gc.collect()` if helpful for diagnostics (document; not a substitute for explicit release in logic).
  4. Call `pythoncom.CoUninitialize()` on STA thread.
  5. **Orphan check (timed):** after cleanup, poll for `owned_excel_pid` absence: e.g. check every 0.5 s for up to **10 seconds** (configurable constant, record value). If still present after timeout: record as **orphan** → test/PoC **FAIL** for AC “no orphaned owned Excel process”; attempt is made only to record (operator may kill owned PID manually outside automation in some environments — document; do not kill user PIDs).
- **No blind replay of mutating calls:** if `SaveAs`/`Close(SaveChanges=True)` style steps were ever attempted in a variant, do not retry automatically in a way that could duplicate side effects — PoC v2 uses non-mutating-close defaults except the explicit SaveAs followed by Close(False) as above.

---

## 6. Message pumping, COM/UI events, RPC_E_CALL_REJECTED, modals — required protocol

### 6.1 Pumping strategy (must be stated)
- The PoC must run on an STA thread. Choose and document a **concrete pumping approach** used during waits (one of):
  - Periodic `pythoncom.PumpWaitingMessages()` calls in loops that wait for bounded operations, **or**
  - A documented `MsgWaitForMultipleObjects`/CoWait-style loop if implemented (only if already standard in repo deps) — otherwise prefer `PumpWaitingMessages` for v2 scope.
- Record in evidence: `message_pump_strategy = "PumpWaitingMessages(interval=...)"` (or exact), and where applied (e.g. after `Quit`, during bounded orphan wait, if any COM busy retries are attempted in a non-mutating probe — see 6.2).

### 6.2 `RPC_E_CALL_REJECTED` / busy / retry policy
- **Classification (log):** catch `pywintypes.com_error`, record `hresult` (hex), message, `excepinfo` if present. Known codes to recognize: `RPC_E_CALL_REJECTED` (0x80010001), `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A). Treat others as COM errors per log.
- **Policy:** **Never blindly replay mutating COM calls** (SaveAs/Close-with-save in unexpected ways / Quit retries in loops that could be harmful — v2 avoids). For **read-only probes** only, a bounded retry with pump may be attempted: max **3** attempts, exponential backoff starting **0.5 s**, pumping between waits if strategy chosen — record every attempt and outcome. Mutating steps in the defined lifecycle are **not retried automatically** on rejection except operator judgment recorded separately (not in automated PoC flow).
- **If observed:** record `rpc_observations` entries (timestamp, operation, hresult, attempts, final result). This is evidence for R01/message-pump behavior.

### 6.3 Modals / Excel UI dialogs
- **Do not attempt** to dismiss programmatically via SendKeys/click coordinates in v2 PoC.
- If the test appears hung or a dialog is suspected (no progress beyond timeout, or operator reports): stop bounded wait, record state, take screenshot if a dialog window is visible (UIA/Desktop), log window titles/handles of Excel-related dialogs if discoverable read-only.
- **“Not exercised” is allowed only if explicitly stated**: e.g. “No modal dialogs observed during this run; pumping strategy exercised only in normal lifecycle and orphan wait.” This **weakens** R01 evidence but is not automatically No-Go **if** all other ACs pass and the rubric section 7 states consequences (see 7). A clean pass with “not exercised” means **conditional** for that sub-risk until a later gate (task_004/task_005/task_010 integration tests) demonstrates modal/busy handling under real flows — the evidence must say so.

---

## 7. Proxy thread-affinity evidence (required)

- Prove design + logs: single STA thread id `sta_thread_id_init` recorded; every COM-touching operation logs that thread id (or is executed only in that context). No COM proxy variables passed to other threads in the harness code.
- **Negative probe (recommended, bounded):** after obtaining an `app` proxy on STA thread, start a short-lived thread that attempts to call a trivial read (e.g. `app.Version`) and **expect failure** (marshaling/apartment error) — catch and log it. This demonstrates confinement risk in practice (do not suppress). Mark as “affinity_probe_performed: true/false” and results.
- If affinity probe is not performed, record “not performed (risk accepted for v2?)” — Auditor/Validator will treat this as a gap unless justified; prefer to perform it.

---

## 8. Go/No-Go rubric (explicit) — decision rules

Evaluate after a complete run with evidence artifacts saved.

### 8.1 Disqualifying conditions → **No-Go** (any one is sufficient)
- Orphaned owned Excel process remains after timed cleanup (section 5.3) and is not explained by external kill with recorded reason that was not the PoC’s fault? No: any orphan from PoC cleanup path is **No-Go**.
- User Excel session (PIDs/windows in `baseline_before`, when present) shown to be modified by PoC actions (closed, quit, workbooks closed/saved unexpectedly) — **No-Go**.
- Evidence shows COM operations occurred on a thread other than the recorded STA thread (or clear log contradiction) — **No-Go** (breaks R01 model).
- `DispatchEx` not used as the creation method, or `GetActiveObject`/attachment used to “obtain” the test instance as primary path — **No-Go**.
- Environment matrix incomplete for required fields (Windows build, Office Excel version/build, Python bitness, dispatch method, STA thread id, cleanup result) with no recorded “unknown + attempts” where allowed — **No-Go** for claiming a validated gate on target; may be **Inconclusive** if run aborted early but still not a Go.

### 8.2 **Conditional Go** (allowed only if all of the following)
- All Disqualifying conditions absent.
- All required AC evidence present: owned instance created via `DispatchEx` and identified (PID + HWND + version readable), lifecycle open/close/Quit executed on owned temp workbook (or documented alternative) with cleanup, no orphan after timed wait, STA model documented and consistent with logs, proxy confinement evidenced by design+logs (and preferably affinity probe), message pump strategy documented and applied where specified.
- **RPC_E_CALL_REJECTED / modal behavior:** either (a) observed and logged with outcomes per 6.2/6.3 (any failures handled per policy) **or** (b) explicitly **not exercised** with statement: “R01 busy/modal sub-risk not exercised on target in this run; dependent gates (`task_004`/`task_005`/`task_010` integration) must validate modal/busy handling before claiming full R01 closure.” 
  - In case (b), decision is **Conditional Go for task_002 lifecycle/isolation/STA-core evidence only**, **not** a full R01 closure. The execution log must record this condition and list follow-up gates.
- VBE language: if `not_discoverable`, record methods attempted; does **not** block Go by itself unless a future AC requires it, but must be recorded (transparent).

### 8.3 **Full Go** (preferred)
- Same as Conditional Go conditions satisfied **and** 6.2/6.3 produced at least one meaningful observation context logged (even if no rejection occurred, record “attempted safe read probe / normal path — no RPC rejection observed”) **and** pumping behavior described during orphan wait and any waits — evidence is strong enough that R01 core (STA + DispatchEx + lifecycle + no orphans) is proven on this target build; modal/busy edge cases still noted as lower-probability but exercised paths documented.

### 8.4 **Inconclusive**
- Run aborted, Excel crashed externally, or evidence incomplete despite operator effort — cannot claim Go/No-Go; document what was proven.

**Final field in artifact:** `go_no_go_r01_task002: GO | CONDITIONAL_GO (note follow-ups) | NO_GO | INCONCLUSIVE` with free-text rationale referencing this rubric.

---

## 9. Test harness outline (concrete, aligned to v2)

- File: `tests/integration_excel/test_excel_instance_poc.py` (temporary PoC, not production abstractions beyond what’s needed).
- Markers: `@pytest.mark.integration_excel`, skip on non-Windows or missing `win32com`/Excel; document “interactive session required”.
- Helpers (outline): `setup_poc_artifact_logger()` → writes `.ai/tmp/excel_instance_poc_<ts>.json` + `.txt` + stderr mirror; `collect_environment_matrix(...)`; `snapshot_excel_processes_windows()`; `snapshot_excel_windows()`; `readonly_rot_snapshot()`; `init_sta_thread()` returns `(thread_ident, coinit_result)`; `dispatch_owned_excel(...)` with delta PID resolution + affinity checks; `lifecycle_temp_xlsx(app, tmpdir)` per section 5; `safe_quit_and_cleanup(app, tracked_workbooks, owned_pid, sta_thread_ident)`, orphan poll; `run_affinity_probe(app, sta_thread_ident)`.
- Fixtures (function-scoped, serial): `poc_logger`, `sta_context` (init/uninit on same thread; yields context dict), `owned_excel_session` (baseline_before, creates app, yields bundle, finally runs cleanup + baseline_after comparison helpers).
- Tests (names from plan):
  - `test_dispatch_ex_creates_isolated_excel_instance`: asserts DispatchEx path, owned identity, env matrix fields captured, baseline-before/after user-session checks (no mutation), STA thread id consistent, no orphan, cleanup completed. Includes “not exercised” note for modals/RPC if none occurred (log it).
  - `test_excel_workbook_open_close_lifecycle`: temp `.xlsx`, Add/SaveAs/Open/Close per 5.2, params logged, same STA thread, orphan check, user-session unaffected. Records pumping/any RPC observations during waits if any.

---

## 10. Manual run checklist (operator)

- Interactive desktop, note open Excel workbooks/titles (or close Excel for clean baseline).
- Close modal dialogs/add-in prompts if possible; no Office updates in progress.
- `uv sync --locked` in repo env.
- Run: `uv run pytest tests/integration_excel/test_excel_instance_poc.py -v -s` (no xdist). Capture console output if desired.
- If dialog appears: screenshot only, record text/title, do not automate dismissal; resolve manually or abort and record as evidence.
- After run: verify no orphaned owned PID, user Excel unchanged (or changes documented), complete evidence artifact and apply rubric 8.

---

## 11. Handoff
- Persist this v2 as `task_002_planner_v2.md` (done). Auditor must re-review against plan/spec + v2; Validator must re-validate after Auditor. If either raises issues, produce v3 addressing them; repeat until Validator returns PASS or PASS WITH WARNINGS with only non-blocking noted items explicitly accepted by Tech Lead decision (documented).
