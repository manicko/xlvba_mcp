# task_002 — Planner output (v3, revised to satisfy Auditor v2 N1–N5 and Validator v2 conditions), 2026-11-19

This revision incorporates Auditor v2 findings N1–N5 and Validator v2 conditions into the task_002 PoC design/gate basis. It supersedes `task_002_planner_v2.md` for implementation of `tests/integration_excel/test_excel_instance_poc.py` and for evidence/rubric evaluation. Any gaps discovered by Auditor/Validator in re-review must be fixed in a v4 before implementation proceeds beyond an approved design.

**Source anchors:** `plan.md` task_002/R01/G1; `spec_text.txt` §5.3, §7.1 context, §9.4, §11/11.1, §13.1; `task_002_planner.md` (v1, preserved conceptually), `task_002_planner_v2.md`, `task_002_auditor_v2.md`, `task_002_validator_v2.md`.

---

## 0. Changes from v2 (N1–N5 addressed)

| ID | Summary | Where integrated |
|----|---------|------------------|
| **N1** (MEDIUM, Validator cond. 2) | **Affinity probe mandatory.** Bounded negative cross-thread probe required: attempt trivial read of owned proxy from another thread, expect/log failure, record `affinity_probe_performed=true` and outcome. If truly infeasible for a documented platform reason (rare), record reason and mark evidence insufficient for a clean pass (treat as **Conditional at best** and explain in §8). “Recommended only” removed. | §3 (artifact fields), §5.4 (required step), §7 (mandatory protocol), §8 (rubric: omission without documented platform infeasibility incompatible with Full Go / clean Conditional Go), §9, §10 |
| **N2** (LOW–MEDIUM) | **Serialized multi-op evidence required.** Execute ≥2 distinct COM operations through the same STA execution boundary in defined order (e.g. read `Application.Version` then perform lifecycle step, or two logged queued submits in same thread context), log submission/completion order and STA thread id. Stated as evidence for same-thread serialized execution on the worker thread (full MCP queue remains `task_010/011`). | §3, §5.3 (required sequence), §7, §8, §9 |
| **N3** (LOW) | **Pinned `Workbooks.Open` params.** Concrete params prescribed and recorded (default: `UpdateLinks=0`, `ReadOnly=False`, `Notify=False`, `AddToMru=False`, with justified alternative allowed only if recorded). App settings `Visible=False`, `DisplayAlerts=False` recorded. Capture actual values in evidence. | §3, §4.3, §5.2, §9 |
| **N4** (LOW) | **Snapshot method fields required.** `process_snapshot_method` (psutil/CIM/tasklist + notes), `window_snapshot_method`, and record partial failures in evidence. | §3, §4.2, §4.5, §9 |
| **N5** (LOW) | **Per-step logging schema defined.** JSON `steps[]` (or equivalent structured list) with fields `name`, `thread_id`, `timestamp`, `result`/`hresult` or status, `notes`. Every COM-touching step must be logged (init, CoInitialize result, Dispatch, workbook Add/SaveAs/Open/Close, Quit, affinity probe, uninit, orphan poll checkpoints as appropriate). | §3 (schema), §4, §5, §7, §9 |

---

## 1. Production alignment (single STA model) — must be explicit

- The PoC **must** demonstrate the same threading/COM model intended for production `src/excel_vba_mcp/adapters/excel/com_worker.py` (`task_010`): **one dedicated STA thread** owns every COM proxy created/used/released for Excel/VBE in that worker context.
- **PoC harness requirement:** all `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)`, all `win32com.client.DispatchEx`/workbook/`Application` proxy access, all reference releases, `CoUninitialize()`, and workbook close/`Quit` calls happen on **the same OS thread** (record thread id `threading.get_ident()` at init and assert/log it for each COM-touching step via `steps[]`). No COM proxies are passed to other threads; no access from pytest/main thread except via the harness API that marshals only non-COM data/logs (if any). The only exception is the **mandatory affinity negative probe** (§7), which intentionally attempts access from another thread and must **fail** as logged evidence.
- **If the implemented test helper uses a different pattern** (e.g. implicit MTA, or COM calls on multiple threads outside the documented probe), the result is **No-Go for R01** regardless of “Excel opened” success.
- Evidence artifact must state: “Aligned to com_worker single-STA design: yes/no (with notes)”. If “no”, Go is forbidden.

---

## 2. Prerequisites / context (repeated, explicit)

- Windows 11 x64 target, **interactive desktop** (logged-in user session). No Windows Service / RDP-without-interactive-desktop assumptions. Verify: `sys.platform == "win32"`, record `platform.platform()`, `platform.version()`, `sys.getwindowsversion()`.
- Microsoft Excel desktop installed. Python/pywin32/Office bitness recorded; recommended x64 Python x64 Office per project; mismatches recorded as risk (may cause No-Go if COM fails).
- **Serial execution only**; do not run with `pytest -n`/xdist. One Office interaction run at a time.
- No Trust Center/registry changes by PoC. Do **not** touch files under `projects/`, `work/`, `snapshots/` except creating a **runtime temporary** workbook in an OS temp dir (or `.ai/tmp`-scoped temp) for lifecycle test; delete in teardown.
- No SendKeys/coordinate UI automation to dismiss dialogs as a primary strategy. If an unexpected modal/dialog appears: capture screenshot only when present (`pywinauto` Desktop screenshot to `.ai/tmp/excel_instance_poc_<ts>_dialog.png`) and record dialog text/title if readable; manual resolution may be required; do not automate dismissal.

---

## 3. Environment matrix and artifact schema (required evidence) — capture at run start

Record into evidence log/artifact (stderr + files under `.ai/tmp/excel_instance_poc_<YYYYMMDD_HHMMSS>.{json,txt}`; both human + machine-readable fields recommended):

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
| `process_snapshot_method` | **Required (N4):** method/library/tool used for Excel PID/process snapshots (e.g. `psutil.process_iter`, `Get-CimInstance Win32_Process`, `tasklist /v /fo csv`) + version/notes | `unknown` + attempts (not allowed to omit field) |
| `window_snapshot_method` | **Required (N4):** method used for top-level window enumeration (e.g. `win32gui.EnumWindows` + PID mapping, UIA, other) + notes | `unknown` + attempts (not allowed to omit field) |
| `snapshot_partial_failures` | **Required (N4):** list/notes of partial failures during snapshot capture (empty if none) | empty list / explicit `none` |
| `app_visible`, `app_display_alerts` | **Required (N3):** actual values set/read on owned `Application` after config (`Visible=False`, `DisplayAlerts=False` expected) | record actual/error |
| `workbook_open_params` | **Required (N3):** concrete params actually used for lifecycle `Workbooks.Open` (see §5.2), recorded as structured fields/values | — |
| `affinity_probe_performed` | **Required (N1):** boolean; must be `true` unless a documented platform infeasibility is recorded (rare) | `false` only with documented reason + insufficiency note |
| `affinity_probe_outcome` | **Required when performed:** expected failure details (exception type/HRESULT/message) or unexpected success (fail condition) | documented reason if not performed |
| `serialized_multiop_evidence` | **Required (N2):** summary object/fields: operations executed, submission/completion order, `sta_thread_id` used, and confirmation that ≥2 distinct COM ops ran through the same STA execution boundary in defined order | — |
| `steps` | **Required (N5):** per-step log array (see schema below) | must be present (may be empty only if run aborted before any step, but then Go/No-Go cannot be claimed) |

**VBE language note (explicit):** “VBE language” in AC is often interpreted as UI language of the VBA IDE; exact COM API is environment-dependent. The PoC must **not** claim certainty if only heuristics were used — record method and confidence (`confirmed_via_...` or `heuristic_...` or `not_discoverable`).

### 3.1 Per-step logging schema (N5, required)

`steps` MUST be a JSON array (or equivalent structured list in `.txt` with the same fields) where **every COM-touching step** is logged. Each entry MUST include:

| Field | Description |
|-------|-------------|
| `name` | Stable step identifier (e.g. `sta_init`, `coinitialize`, `dispatch_ex`, `set_app_flags`, `serialized_op_1_version`, `serialized_op_2_lifecycle_probe`, `wb_add`, `wb_saveas`, `wb_close_after_save`, `wb_open`, `wb_close_after_open`, `affinity_probe`, `owned_wb_cleanup`, `app_quit`, `couninitialize`, `orphan_poll_1` … `orphan_poll_n`) |
| `thread_id` | OS thread id (`threading.get_ident()`) on which the step executed/submitted at the COM boundary (for affinity probe, log both STA thread id and probe thread id in `notes`/fields as applicable) |
| `timestamp` | ISO-8601 timestamp (UTC or local; be consistent and record timezone policy) |
| `result` / `hresult` or `status` | Outcome: success/failure plus HRESULT/exception code when applicable (hex preferred for COM HRESULTs) |
| `notes` | Free-text/details: params, submission/completion order markers, observed values, retry count, partial failure notes, etc. |

**Minimum required coverage:** `sta_init`/CoInitialize result, `dispatch_ex`, app flag configuration reads/writes, each serialized multi-op operation (≥2) with submission/completion order, workbook `Add`/`SaveAs`/`Open`/`Close` (each), affinity probe, `Quit`, `CoUninitialize`, and orphan poll checkpoints as appropriate (each poll or a summarized series with per-checkpoint evidence). Non-COM snapshot steps may be logged for traceability but COM-touching steps must not be omitted.

---

## 4. Isolation protocol (rigorous) — required

### 4.1 Precondition
- Before any `DispatchEx`, perform **baseline capture (label `baseline_before`)**. Preferred: run with **no user Excel instances** for clearest proof (operator may close Excel). If user Excel is open, baseline **must** capture it (PIDs, windows) and the PoC **must not** mutate any non-owned instance.

### 4.2 Baseline capture (read-only, non-intrusive)
Capture and persist (record `process_snapshot_method`, `window_snapshot_method`, and `snapshot_partial_failures` per N4 in the artifact; log capture as COM/non-COM steps where appropriate):
- `user_excel_pids_before`: list of `EXCEL.EXE` PIDs (`psutil` if available, else `Get-CimInstance Win32_Process`/`tasklist` parsing) with start time/session if available. Record which method was actually used and any partial failures.
- `user_excel_windows_before`: read-only enumeration of top-level windows owned by those PIDs (title, HWND) — no clicks, no focus stealing beyond enumeration APIs used read-only. Record method and partial failures.
- **ROT/GetActiveObject snapshot (read-only, best-effort):** call `win32com.client.GetActiveObject("Excel.Application")` **only** to record whether any instance is registered and, if so, capture **read-only** observable facts (e.g. workbook count/names via read of `Workbooks.Count`/iterated names in a way that does not save/close/modify — list names only).  
  - **Must not:** `Quit`, `Close`, `Save`, run macros, change `DisplayAlerts` in a way that affects user workflow beyond local read? Better: if attached to a user instance by accident of API behavior, **do not assume ownership** — document. Prefer to treat ROT result as informational; **isolation proof is primarily PID/HWND delta + no mutation** of user PIDs/windows.  
  - If `GetActiveObject` fails (e.g. `MK_E_UNAVAILABLE`), record that explicitly.
- Record `timestamp_baseline_before`.

**Forbidden:** any write/close/save/quit/run on processes/workbooks not owned by the PoC after `DispatchEx` is created; any enumeration that requires modifying user state.

### 4.3 Owned instance creation and identity
- On STA thread: `app = win32com.client.DispatchEx("Excel.Application")`, set `app.Visible = False`, `app.DisplayAlerts = False` (and record **actual values** in evidence, N3). Restore not required for owned instance beyond cleanup, but do not touch others.
- **Owned identity (required):**
  - Determine `owned_excel_pid`: compute PID set delta `after_dispatch_pids - user_excel_pids_before` intersected with processes that are the new Excel process; if ambiguous, enumerate Excel processes again and correlate by timing/`Application.Hwnd` if resolvable to a window process — record method. If still `unknown`, **AC for isolation identification fails** (must be recorded; Go is not possible for a clean pass).
  - Record `owned_excel_hwnd = app.Hwnd`, `owned_excel_caption = app.Caption`.
  - Assert `owned_excel_pid` is not in `user_excel_pids_before`.
  - Log `sta_thread_id_init` again for confirmation that all subsequent COM calls use the same thread context (documented, and present in `steps[]`).

### 4.4 Post-creation / during-run checks
- After creating owned instance and after each major step (open workbook, close workbook), re-check that no user PID disappeared unexpectedly unless user acted externally (record external changes as “inconclusive for user-session-unaffected” if observed) and no user windows changed titles in unexpected ways (best effort). Record snapshot methods/partial failures if re-captured.

### 4.5 Post-cleanup baseline
- After full cleanup (section 5): capture `baseline_after` using the same method set where possible; record `process_snapshot_method`, `window_snapshot_method`, and any `snapshot_partial_failures` (N4). Compare to `baseline_before`.
- **Pass for “existing user Excel session not touched”:** all PIDs present in `baseline_before` still present **or** the operator explicitly stated they closed them during run (recorded) — but for a strict PoC claim of “not touched by PoC”, the stronger claim is: PoC performed no operations on non-owned COM objects and no user PID disappeared solely due to PoC; any disappearance must be attributed or treated as fail. Default: if user had Excel open and any of those PIDs are gone after PoC cleanup without operator note, mark **FAIL** (isolation/user-session risk).
- If there were **no** user Excel instances before, `baseline_after` must show **no orphaned owned PID**. Assert owned PID absent; any other new Excel PIDs require explanation (fail unless clearly external — recorded as inconclusive).

---

## 5. Lifecycle, shutdown, orphans, COM init/uninit — required procedure

### 5.1 Thread and COM init
- Run all COM in a **single dedicated STA thread** (fixture/worker thread in tests). Call `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)` on that thread. Log outcome in `steps[]` (`sta_init`/`coinitialize`) and in matrix fields.
- On teardown, call `pythoncom.CoUninitialize()` on the **same** thread (after releasing all COM references created there). Log in `steps[]`.

### 5.2 Workbook/test file policy and pinned lifecycle params (N3)
- Create a unique temporary directory (e.g. `tempfile.TemporaryDirectory()` under system temp) and a file path `poc_lifecycle.xlsx`. Do **not** use any repo files.
- **Owned app settings (required, N3):** set and record actual values `app.Visible = False`, `app.DisplayAlerts = False` (see §4.3, §3).
- **Prescribed `Workbooks.Open` params (required, N3):** use and record concrete values:
  - `UpdateLinks=0`
  - `ReadOnly=False`
  - `Notify=False`
  - `AddToMru=False`
  - **Justified alternative:** allowed only if explicitly justified in evidence (why required) and actual values recorded in `workbook_open_params`.
- Sequence (owned instance only, on STA thread), logging every COM-touching step in `steps[]`:
  1. `wb_new = app.Workbooks.Add()` (or documented equivalent) — record result.
  2. `wb_new.SaveAs(path, FileFormat=51)` (xlsx). Record HRESULT/errors and actual path/format.
  3. `wb_new.Close(SaveChanges=False)`. Release `wb_new` reference. Record `SaveChanges=False`.
  4. `wb_open = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False, Notify=False, AddToMru=False)` (or justified alternative recorded). Capture actual params in `workbook_open_params`.
  5. Verify open (e.g. `wb_open.Name`/`FullName` matches) — read-only checks.
  6. `wb_open.Close(SaveChanges=False)`. Release `wb_open`. Record `SaveChanges=False`.

**SaveChanges policy:** for any workbook opened/saved by PoC, default `SaveChanges=False` on Close unless the step explicitly requires saving — record every Close with `SaveChanges` value.

### 5.3 Serialized multi-op evidence through same STA execution boundary (N2, required)

- Execute **≥2 distinct COM operations** through the **same STA execution boundary** in a **defined order**. Acceptable examples (choose one and log clearly): (a) read `Application.Version` (op 1) then read `Application.Caption` or `Workbooks.Count` (op 2) as logged queued/sequential submits in the same thread context; (b) read `Application.Version` then perform a subsequent lifecycle step (e.g. `Workbooks.Add`) as the ordered pair, provided both are distinct COM ops executed through the same boundary and individually logged.
- **Logging (required):** for each operation log `name`, `thread_id`, `timestamp`, `result`/`hresult` or `status`, `notes` in `steps[]`, and record **submission/completion order** explicitly (e.g. sequence indices or ordered timestamps/markers) and the **STA thread id**.
- **Evidence statement (required):** record in `serialized_multiop_evidence` that this demonstrates **same-thread serialized execution on the worker thread**. **Scope note:** the **full MCP request queue** remains `task_010/011` (PoC provides STA-boundary/worker-thread serialization evidence only, not end-to-end MCP queue semantics).

### 5.4 Mandatory affinity negative probe (N1, required)

- Perform a **bounded negative cross-thread probe** (see §7 for full protocol): attempt a trivial read of the owned proxy from another thread, **expect/log failure**, and record `affinity_probe_performed=true` and outcome in evidence and `steps[]`.
- If truly infeasible for a **documented platform reason** (rare), record the concrete reason and mark evidence **insufficient for a clean pass** (treat as **Conditional at best**; explain in §8). Do not omit silently.

### 5.5 Shutdown (always in `finally`)
- In `finally`, on the **same STA thread**:
  1. Close **only owned** workbooks still referenced (iterate a tracked list; call `Close(SaveChanges=False)`) catching and logging COM errors (do not silently swallow in a way that hides orphans — log then continue cleanup best effort). Log each in `steps[]`.
  2. Call `app.Quit()` (owned instance). Catch and log errors.
  3. Release all COM references (`del` Python vars; ensure no remaining strong references in the test scope for those proxies). Call `gc.collect()` if helpful for diagnostics (document; not a substitute for explicit release in logic).
  4. Call `pythoncom.CoUninitialize()` on STA thread. Log in `steps[]`.
  5. **Orphan check (timed):** after cleanup, poll for `owned_excel_pid` absence: e.g. check every 0.5 s for up to **10 seconds** (configurable constant, record value). Log orphan poll checkpoints in `steps[]` as appropriate. If still present after timeout: record as **orphan** → test/PoC **FAIL** for AC “no orphaned owned Excel process”; attempt is made only to record (operator may kill owned PID manually outside automation in some environments — document; do not kill user PIDs).
- **No blind replay of mutating calls:** if `SaveAs`/`Close(SaveChanges=True)` style steps were ever attempted in a variant, do not retry automatically in a way that could duplicate side effects — v3 uses non-mutating-close defaults except the explicit SaveAs followed by Close(False) as above.

---

## 6. Message pumping, COM/UI events, RPC_E_CALL_REJECTED, modals — required protocol

### 6.1 Pumping strategy (must be stated)
- The PoC must run on an STA thread. Choose and document a **concrete pumping approach** used during waits (one of):
  - Periodic `pythoncom.PumpWaitingMessages()` calls in loops that wait for bounded operations, **or**
  - A documented `MsgWaitForMultipleObjects`/CoWait-style loop if implemented (only if already standard in repo deps) — otherwise prefer `PumpWaitingMessages` for v3 scope.
- Record in evidence: `message_pump_strategy = "PumpWaitingMessages(interval=...)"` (or exact), and where applied (e.g. after `Quit`, during bounded orphan wait, if any COM busy retries are attempted in a non-mutating probe — see 6.2). Log relevant waits/pump usage in `steps[]` notes where applicable.

### 6.2 `RPC_E_CALL_REJECTED` / busy / retry policy
- **Classification (log):** catch `pywintypes.com_error`, record `hresult` (hex), message, `excepinfo` if present. Known codes to recognize: `RPC_E_CALL_REJECTED` (0x80010001), `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A). Treat others as COM errors per log.
- **Policy:** **Never blindly replay mutating COM calls** (SaveAs/Close-with-save in unexpected ways / Quit retries in loops that could be harmful — v3 avoids). For **read-only probes** only, a bounded retry with pump may be attempted: max **3** attempts, exponential backoff starting **0.5 s**, pumping between waits if strategy chosen — record every attempt and outcome in `steps[]`/notes. Mutating steps in the defined lifecycle are **not retried automatically** on rejection except operator judgment recorded separately (not in automated PoC flow).
- **If observed:** record `rpc_observations` entries (timestamp, operation, hresult, attempts, final result). This is evidence for R01/message-pump behavior.

### 6.3 Modals / Excel UI dialogs
- **Do not attempt** to dismiss programmatically via SendKeys/click coordinates in v3 PoC.
- If the test appears hung or a dialog is suspected (no progress beyond timeout, or operator reports): stop bounded wait, record state, take screenshot if a dialog window is visible (UIA/Desktop), log window titles/handles of Excel-related dialogs if discoverable read-only.
- **“Not exercised” is allowed only if explicitly stated**: e.g. “No modal dialogs observed during this run; pumping strategy exercised only in normal lifecycle and orphan wait.” This **weakens** R01 evidence but is not automatically No-Go **if** all other ACs pass and the rubric in section 8 states consequences. A run with “not exercised” means **Conditional Go for task_002 lifecycle/isolation/STA-core evidence only** (not full R01 closure) when all other rubric conditions for Conditional are met — the evidence must say so and list follow-up gates.

---

## 7. Proxy thread-affinity evidence (required, N1)

- **Design + logs (required):** single STA thread id `sta_thread_id_init` recorded; every COM-touching operation logs that thread id in `steps[]` (or is executed only in that context). No COM proxy variables passed to other threads in the harness code (except the intentional cross-thread failure in the mandatory probe, which must not reuse the proxy successfully).
- **Mandatory negative probe (N1, required):** after obtaining an owned `app` proxy on the STA thread, start a **short-lived separate thread** that attempts to call a **trivial read** on that proxy (e.g. `app.Version` or `app.Caption`) and **expect failure** (marshaling/apartment error or equivalent). Catch and log the exception type/HRESULT/message and outcome in `steps[]` and in `affinity_probe_outcome`. Record `affinity_probe_performed=true` (and probe thread id in notes). This demonstrates confinement risk in practice (do not suppress). An unexpected success (cross-thread call succeeds) is a **FAIL** for confinement evidence and must be treated per §8 (No-Go for R01 model alignment/safety claim).
- **If truly infeasible (rare, documented platform reason only):** record the concrete reason (environment/API limitation with details), set `affinity_probe_performed=false`, record `affinity_probe_outcome` with reason, and mark evidence **insufficient for a clean pass**. Per §8, this is **incompatible with Full Go** and **not a clean Conditional Go**; at best **Conditional Go** with explicit insufficiency noted and follow-up required (and may be **No-Go/Inconclusive** if other required evidence is weak). Silent omission is **not allowed**.
- **“Recommended only” is removed:** the probe is required in practice and in rubric evaluation except the documented-infeasibility case above.

---

## 8. Go/No-Go rubric (explicit) — decision rules

Evaluate after a complete run with evidence artifacts saved.

### 8.1 Disqualifying conditions → **No-Go** (any one is sufficient)
- Orphaned owned Excel process remains after timed cleanup (section 5.5) and is not explained by external kill with recorded reason that was not the PoC’s fault? No: any orphan from PoC cleanup path is **No-Go**.
- User Excel session (PIDs/windows in `baseline_before`, when present) shown to be modified by PoC actions (closed, quit, workbooks closed/saved unexpectedly) — **No-Go**.
- Evidence shows COM operations occurred on a thread other than the recorded STA thread (or clear log contradiction), **excluding** the intentional failure inside the mandatory affinity probe — **No-Go** (breaks R01 model).
- Affinity probe unexpectedly **succeeds** (cross-thread trivial read succeeds on owned proxy) — **No-Go** for confinement/R01 safety claim.
- `DispatchEx` not used as the creation method, or `GetActiveObject`/attachment used to “obtain” the test instance as primary path — **No-Go**.
- Environment matrix/artifacts incomplete for required fields (including N3/N4/N5 fields and `steps[]` coverage as applicable, Windows build, Office Excel version/build, Python bitness, dispatch method, STA thread id, cleanup result, `serialized_multiop_evidence`) with no recorded “unknown + attempts” where allowed — **No-Go** for claiming a validated gate on target; may be **Inconclusive** if run aborted early but still not a Go.

### 8.2 **Conditional Go** (allowed only if all of the following)
- All Disqualifying conditions absent.
- All required AC evidence present: owned instance created via `DispatchEx` and identified (PID + HWND + version readable), lifecycle open/close/Quit executed on owned temp workbook with pinned/recorded params per §5.2, no orphan after timed wait, STA model documented and consistent with logs/`steps[]`, **serialized multi-op evidence** recorded per §5.3, **affinity probe performed and recorded as failed-as-expected** per §7 (`affinity_probe_performed=true`) **or** documented platform infeasibility handled per §7/§8 note below, message pump strategy documented and applied where specified, snapshot methods/partial failures recorded per N4.
- **Affinity probe omission (N1):** omission **without documented platform infeasibility** is **incompatible with Conditional Go** (and with Full Go). If truly infeasible for a **documented platform reason** (rare), evidence is **insufficient for a clean pass**: the result may be **Conditional at best** with explicit insufficiency/follow-up noted, never Full Go; treat as **No-Go or Inconclusive** if other evidence is not strong.
- **RPC_E_CALL_REJECTED / modal behavior:** either (a) observed and logged with outcomes per 6.2/6.3 (any failures handled per policy) **or** (b) explicitly **not exercised** with statement: “R01 busy/modal sub-risk not exercised on target in this run; dependent gates (`task_004`/`task_005`/`task_010` integration) must validate modal/busy handling before claiming full R01 closure.” 
  - In case (b), decision is **Conditional Go for task_002 lifecycle/isolation/STA-core evidence only**, **not** a full R01 closure. The execution log must record this condition and list follow-up gates.
- VBE language: if `not_discoverable`, record methods attempted; does **not** block Go by itself unless a future AC requires it, but must be recorded (transparent).

### 8.3 **Full Go** (preferred)
- Same as Conditional Go conditions satisfied **and** the affinity probe was **performed** with expected failure recorded (`affinity_probe_performed=true`) — **not** omitted/infeasible.
- 6.2/6.3 produced at least one meaningful observation context logged (even if no rejection occurred, record “attempted safe read probe / normal path — no RPC rejection observed”) **and** pumping behavior described during orphan wait and any waits — evidence is strong enough that R01 core (STA + DispatchEx + lifecycle + no orphans + serialized same-thread evidence + confinement probe) is proven on this target build; modal/busy edge cases still noted as lower-probability but exercised paths documented.
- **Affinity probe omission (with or without reason) is incompatible with Full Go.**

### 8.4 **Inconclusive**
- Run aborted, Excel crashed externally, or evidence incomplete despite operator effort — cannot claim Go/No-Go; document what was proven.

**Final field in artifact:** `go_no_go_r01_task002: GO | CONDITIONAL_GO (note follow-ups/insufficiency) | NO_GO | INCONCLUSIVE` with free-text rationale referencing this rubric.

---

## 9. Test harness outline (concrete, aligned to v3)

- File: `tests/integration_excel/test_excel_instance_poc.py` (temporary PoC, not production abstractions beyond what’s needed).
- Markers: `@pytest.mark.integration_excel`, skip on non-Windows or missing `win32com`/Excel; document “interactive session required”.
- **Artifact/schema helpers (N4, N5):** `setup_poc_artifact_logger()` → writes `.ai/tmp/excel_instance_poc_<ts>.json` + `.txt` + stderr mirror; maintains `steps[]` with `name`, `thread_id`, `timestamp`, `result`/`hresult` or `status`, `notes`; records `process_snapshot_method`, `window_snapshot_method`, `snapshot_partial_failures`.
- Helpers (outline): `collect_environment_matrix(...)` including N3/N4 fields and placeholders for `affinity_probe_performed/outcome`, `serialized_multiop_evidence`, `steps`; `snapshot_excel_processes_windows()` and `snapshot_excel_windows()` with method/partial-failure capture; `readonly_rot_snapshot()`; `init_sta_thread()` returns `(thread_ident, coinit_result)` and logs steps; `dispatch_owned_excel(...)` with delta PID resolution + owned identity; `run_serialized_multiop_evidence(app, sta_thread_ident, logger)` executes/logs ≥2 distinct COM ops in defined order with submission/completion order and STA thread id (N2); `lifecycle_temp_xlsx(app, tmpdir, logger)` per §5.2 with **pinned** `Workbooks.Open` params and recorded actual values (N3), app `Visible/DisplayAlerts` recorded; `run_affinity_probe(app, sta_thread_ident, logger)` performs mandatory bounded negative cross-thread trivial read, expects/logs failure, records outcome (N1); `safe_quit_and_cleanup(app, tracked_workbooks, owned_pid, sta_thread_ident, logger)` including timed orphan poll checkpoints logged.
- Fixtures (function-scoped, serial): `poc_logger`, `sta_context` (init/uninit on same thread; yields context dict), `owned_excel_session` (baseline_before, creates app, yields bundle, finally runs cleanup + baseline_after comparison helpers).
- Tests (names from plan):
  - `test_dispatch_ex_creates_isolated_excel_instance`: asserts DispatchEx path, owned identity, env matrix fields captured (incl. N4), baseline-before/after user-session checks (no mutation), STA thread id consistent, serialized multi-op evidence (N2) and/or logged ordered ops present, mandatory affinity probe (N1) performed with expected failure recorded, no orphan, cleanup completed. Includes “not exercised” note for modals/RPC if none occurred (log it) and rubric consequence.
  - `test_excel_workbook_open_close_lifecycle`: temp `.xlsx`, Add/SaveAs/Open/Close per §5.2 with pinned/recorded Open params and `SaveChanges=False` policy (N3), same STA thread, `steps[]` coverage for COM ops, orphan check, user-session unaffected. Records pumping/any RPC observations during waits if any.

---

## 10. Manual run checklist (operator)

- Interactive desktop, note open Excel workbooks/titles (or close Excel for clean baseline).
- Close modal dialogs/add-in prompts if possible; no Office updates in progress.
- `uv sync --locked` in repo env.
- Run: `uv run pytest tests/integration_excel/test_excel_instance_poc.py -v -s` (no xdist). Capture console output if desired.
- If dialog appears: screenshot only, record text/title, do not automate dismissal; resolve manually or abort and record as evidence.
- After run: verify artifacts in `.ai/tmp/` contain env matrix (incl. N3/N4 fields), `steps[]` (N5), `serialized_multiop_evidence` (N2), affinity probe outcome (N1, mandatory), snapshot method/partial-failure notes, baseline before/after, pumping/RPC/modal notes, timed orphan checks, and `go_no_go_r01_task002` per §8. Verify no orphaned owned PID and user Excel unchanged (or changes documented).

---

## 11. Handoff

- Persist this v3 as `task_002_planner_v3.md` (done). Auditor must re-review against plan/spec + v3; Validator must re-validate after Auditor. If either raises issues, produce v4 addressing them; repeat until Validator returns PASS or PASS WITH WARNINGS with only non-blocking noted items explicitly accepted by Tech Lead decision (documented).
- Implementation must treat Planner v3 as binding: N1–N5 are required, affinity probe omission without documented platform infeasibility is incompatible with Full Go / clean Conditional Go, and “not exercised” RPC/modal limits closure to Conditional Go for lifecycle/isolation/STA-core only with follow-ups to `task_004`, `task_005`, `task_010` integration as applicable.

(End of file)
