# XLVBA MCP — Phase 5 Hidden Execution & Desktop Non-Interference

**Task:** `FE-14` / `FULL-EDIT-MODE-ORCHESTRATION-001` (Phase 5 acceptance)
**Date:** 2026-11-22
**Auditor role:** Independent, read-only acceptance. **Production code modified: NONE.**
**Repo / branch:** `C:\py_exp\xlvba_mcp` @ `master`, HEAD `dece703`.
**Method:** code grep audit of `src/` + an independent read-only Win32 focus probe driving the
**real MCP STDIO server** (`python -m excel_vba_mcp`) on a disposable copy of the fixture, plus
confirmation of the already-committed e2e harness assertions.

---

## 1. Environment (verified live)

| Item | Value |
|---|---|
| OS | Windows 10 Корпоративная LTSC (Enterprise LTSC 2019), NT 10.0.17763.0 (SP0) |
| Host session | **Interactive console, Session 1** (see §3). Not Session 0 / not a service. |
| Python | 3.14.7 (`.venv\Scripts\python.exe`) |
| pywin32 | import OK (`win32com`, `pythoncom`, `win32gui`, `win32process`) |
| Excel | COM `Application.Version = 16.0`; owned hidden `DispatchEx` probe returned `Visible=False` |
| Excel install | `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE` |
| MCP package | `mcp` stdio client + `ClientSession` |
| Access mode | `EXCEL_VBA_MCP_ACCESS_MODE` **unset** → default `full` |
| VBOM trust | Available (classify + `macro_run` succeeded) |
| Test root | per-run temp root (`tempfile.mkdtemp`); repo fixture never opened |

A manual `win32com.client.DispatchEx("Excel.Application")` probe confirmed an independent owned
instance starts with `Visible = False` and `Version = 16.0` (no focus, no visible window).

---

## 2. Phase-5 checklist (10 items)

Legend: **PASS** = objectively observed; **NOT VERIFIED** = not reliably observable / not covered.

| # | Item | Expected | Observed | Status | Evidence |
|---|---|---|---|---|---|
| 1 | Excel launched **hidden**, owned instance (owned PID ∉ pre-existing user PIDs) | Owned hidden EXCEL.exe; owned PID not a pre-existing user PID | Owned PID `13692`; `excel_health` `status=ok, worker_state=running, owned_pid=13692`; baseline user PID `21100`; `13692 ∉ {21100}`; hidden probe `Visible=False` | **PASS** | `.ai/tmp/fe14_focus_probe.py` run (this audit); `ExcelAppAdapter.create_hidden` (`DispatchEx` + `Visible=False`); harness `test_full_access_workflow_e2e` (`owned_pid not in baseline_pids`, `_assert_owned_hidden(...,"server start")`) |
| 2 | No VM/container/separate session/service/SYSTEM/remote host | Runs in the current interactive Windows user session | `qwinsta`: `console  MatveichevN  1  Active`; `Get-Process -Id $PID` → SessionId `1`; Excel PID `21100` in Session 1; server is a local `python -m excel_vba_mcp` stdio child; no container/VM APIs | **PASS** | `qwinsta`, `Get-Process`, tasklist; server env (`_env()` in probe) is a local subprocess |
| 3 | Existing user Excel instances not reused/controlled/closed | Baseline user PID set survives; server uses a fresh `DispatchEx` instance | Baseline `{21100}` → post `{21100}`; `baseline_user_pids_survived=true`; `owned_pid_reaped=true`; user instance never touched (owned PID differs) | **PASS** | probe report (`baseline_excel_pids=[21100]`, `post_excel_pids=[21100]`); `create_hidden` uses `DispatchEx` (not ROT reuse); harness `surviving_user_pids` empty |
| 4 | No owned visible window (per-PID `EnumWindows`+`IsWindowVisible` empty) at stable points | Owned PID has no visible top-level window | At **server_started / after_excel_open / after_showmessage_classify_veto / after_macro_run_1 / after_macro_run_2 / after_excel_close**: `owned_visible_windows = []` each | **PASS** | Independent probe `_visible_windows_for_pid(owned_pid)` (read-only `EnumWindows`+`IsWindowVisible`+`GetWindowThreadProcessId`); harness `_assert_owned_hidden` at start/open/cycles |
| 5 | No `SetForegroundWindow` / `ShowWindow` / `SendKeys` / foreground clicks in code path | Zero call sites | **Zero call sites** of `SetForegroundWindow`/`ShowWindow`/`SetActiveWindow`/`BringWindowToTop`/`SendKeys`/`AppActivate`/`keybd_event`/`mouse_event`/`SetCursorPos`/`SendInput`/`BlockInput`/`SwitchToThisWindow`/`PostMessage`/`SendMessage` in `src/`. Only `ctypes.windll.user32` uses are `GetForegroundWindow` and `GetWindowThreadProcessId` (both read-only). | **PASS** | grep sweep of 64 `src/**/*.py` files (§4); all matches are docstrings/denylist string literals |
| 6 | Interactive VBA rejected without displaying UI | `ShowMessage` classify veto; no dialog appears | `macro_classify_background("modWorkflow.ShowMessage", true)` → `INTERACTIVE_UI_NOT_SUPPORTED` (R13 veto); `after_showmessage_classify_veto` `owned_visible_windows=[]`, foreground unchanged | **PASS** | probe `showmessage_veto` payload; harness asserts `error_code(vetoed)=="INTERACTIVE_UI_NOT_SUPPORTED"`; Phase-4 §6 |
| 7 | No foreground UI automation to operate VBE | comtypes UIAutomationCore only; no pywinauto | `pywinauto` removed from `pyproject.toml`/`uv.lock` (commit `c28b3dd`); no `pywinauto`/`click_input`/`type_keys` imports; `windows_ui` uses comtypes `IUIAutomation` **read-only** (`ElementFromHandle`/`FindAll`/`GetText`), `DialogReader` never activates/clicks; compile path **fails closed** (`COMPILE_INDETERMINATE`, no compile, no UIA expansion) | **PASS** | `adapters/windows_ui/vbe_compile_ui.py` (`DialogReader`, `VBECompileUI.discover_command`); `application/compile_service.py` (fail-closed); `git show c28b3dd` |
| 8 | Worker health/recovery correct after repeated mutation + macro execution | `excel_health` ok pre/post | Pre: `status=ok, worker_state=running, owned_pid=13692`. Post (after 2× `table_update_rows` + 2× `macro_run` + classify + close): `status=ok, worker_state=running, accepting=true, needs_recovery=false, open_workbooks=1, owned_pid=13692` | **PASS** | probe `health_post`; harness `_wait_healthy`/`_assert_healthy_and_live` pre and per-cycle |
| 9 | Owned PID reaped after server shutdown; user PIDs alive | Owned gone; user PIDs survive | `owned_pid=13692` → post PIDs `[21100]`; `owned_pid_reaped=true`; `baseline_user_pids_survived=true` | **PASS** | probe `owned_pid_reaped`/`baseline_user_pids_survived`; harness `_wait_pid_dead(owned_pid)` + `surviving_user_pids` empty |
| 10 | Focus-theft: reliable observation (foreground-window logger) | If no reliable tool → NOT VERIFIED + evidence needed | Foreground was **unchanged and never owned by Excel** at all 7 sample points (was a Chrome window pre-start, then the VS Code / Kilo window PID 16376; owned Excel never appeared). But these are **discrete point samples**, not a continuous trace, so a transient focus grab between samples cannot be excluded. | **NOT VERIFIED** | Probe `_foreground()` (GetForegroundWindow+GetWindowText+PID) at 7 points; no continuous foreground logger. See §5 for evidence needed. |

**Counts:** 9 PASS, 1 NOT VERIFIED (item 10), 0 FAIL.

---

## 3. Host interactive-session confirmation

```
SESSIONNAME       USERNAME        ID  STATE   TYPE        DEVICE
 services                         0  Disc
>console           MatveichevN     1  Active
 rdp-tcp                          65536  Listen
```

* Current shell `Get-Process -Id $PID` → `SessionId = 1`, `ProcessName = pwsh`.
* Pre-existing user Excel `PID 21100` → `SessionId = 1` (interactive console, Office16).
* No service / SYSTEM / Session 0 / VM / container / remote host is involved. The MCP server is a
  local `python -m excel_vba_mcp` child process in the same session.

---

## 4. Forbidden-API grep audit (64 `src/**/*.py` files)

**Forbidden focus/input/foreground APIs — ZERO call sites:**

| API | Result |
|---|---|
| `SetForegroundWindow` | 0 calls (2 docstring mentions only) |
| `ShowWindow` | 0 calls (docstring mentions only) |
| `SetActiveWindow`, `BringWindowToTop`, `SwitchToThisWindow` | 0 |
| `SendKeys`, `AppActivate` | 0 calls (`SendKeys`/`AppActivate` appear only as string literals in `macro_service.INTERACTIVE_TOKENS`) |
| `keybd_event`, `mouse_event`, `SetCursorPos`, `SendInput`, `BlockInput` | 0 |
| `PostMessage`, `SendMessage`, `SetFocus` | 0 calls (1 docstring mention) |
| `.Activate()` / `.Select()` (Excel object model) | 0 calls (docstrings state they are never called) |
| `pywinauto`, `click_input`, `type_keys` | 0 (dependency removed, `c28b3dd`) |

**Read-only Win32 usage (NOT focus-stealing):**

| File | API | Purpose |
|---|---|---|
| `adapters/excel/app_adapter.py` `_owned_window_is_foreground` | `GetForegroundWindow` | read-only check "is owned hwnd the foreground window?" |
| `adapters/excel/app_adapter.py` `_pid_from_hwnd` | `GetWindowThreadProcessId` | read-only PID resolution for the owned instance |

**`INTERACTIVE_TOKENS` denylist** (`application/macro_service.py`): `MsgBox, InputBox, .Show,
FileDialog, Shell, CreateObject, FollowHyperlink, SendKeys, DoEvents, GetOpenFilename,
GetSaveAsFilename, CallByName, ExecuteGlobal, Execute, Eval, OnTime, DisplayAlerts, .Run,
CommandBars, AppActivate, Application.Visible, .Activate, WindowState` — this is a **veto list**,
not a call site. None of these are executed by the server; `SendKeys`/`AppActivate` are matched
against macro source text only.

---

## 5. Focus probe observations (independent, read-only)

Probe: `.ai/tmp/fe14_focus_probe.py` (read-only Win32 + real MCP STDIO server; temp copy of
`tests/fixtures/FullAccessWorkflow.xlsm`; no `SetForegroundWindow`/`ShowWindow`/`SendKeys`).

Foreground window (`GetForegroundWindow` + `GetWindowText` + PID) and owned visible windows
(`EnumWindows`+`IsWindowVisible`+`GetWindowThreadProcessId` for the owned PID) at each stable point:

| Sample point | Foreground window (PID) | Owned visible windows |
|---|---|---|
| pre_server_start | Chrome "Airbnb — Google Chrome" (21856) | — |
| server_started | VS Code "Kilo Settings - xlvba_mcp" (16376) | `[]` |
| after_excel_open | VS Code (16376) | `[]` |
| after_showmessage_classify_veto | VS Code (16376) | `[]` |
| after_macro_run_1 | VS Code (16376) | `[]` |
| after_macro_run_2 | VS Code (16376) | `[]` |
| after_excel_close | VS Code (16376) | `[]` |
| after_server_shutdown | VS Code (16376) | — |

Key observations:
* **Owned Excel was never the foreground window** at any sample; it had **no visible top-level
  window** at every stable point.
* The foreground window changed once, **before** the server started (Chrome → VS Code), which is
  attributable to the audit tooling/IDE, not the server. It remained constant (VS Code, PID 16376)
  across every server/MCP observation.
* Owned PID `13692` reaped after shutdown; user PID `21100` survived; fixture SHA-256 unchanged
  (`6B905837B8DEEB549505F9A8A51C4F55507D377E4160BB5D9D6363A19AC1D4D9`).

**Limitation (why item 10 is NOT VERIFIED):** this is a **discrete point sample**, not a
continuous foreground trace. A transient focus grab that occurred and reverted **between** two
samples would not be observed. No continuous foreground-window logger was available.

**Evidence needed to upgrade item 10 to PASS:**
1. A continuous foreground-window logger (e.g. a background thread calling `GetForegroundWindow`+
   `GetWindowThreadProcessId` every 10–50 ms, or a `SetWinEventHook(EVENT_SYSTEM_FOREGROUND)` out-
   of-context hook) running for the entire server lifecycle, recording `(timestamp, hwnd, pid,
   title)`; assert no record names the owned Excel PID / owned hwnd.
2. Ideally pair with `GetGUIThreadInfo`/`GetLastInputInfo` or an `EVENT_SYSTEM_FOREGROUND` log that
   captures **every** foreground transition, so a transient steal is provable either way.
3. The probe must run in the interactive session (as here) so foreground is observable at all.

> **Note on existing tooling:** the POC helpers `tests/integration_excel/test_vbide_access_poc.py`
> `_record_foreground` and this audit's `.ai/tmp/fe14_focus_probe.py` are **discrete point samplers**
> (before / after_dispatch / after_operations / after, and 7 lifecycle points respectively). Neither
> is a continuous logger, so neither can exclude a transient focus grab between samples. The
> `EVENT_SYSTEM_FOREGROUND` hook (or a 10–50 ms poll loop) above is the specific missing capability.

---

## 6. Existing e2e evidence cited

`tests/integration_excel/test_mcp_full_access_workflow.py` already asserts, per owned PID:
* owned PID ∉ pre-existing user PIDs (`owned_pid not in baseline_pids`);
* no owned visible window via `_assert_owned_hidden` (`EnumWindows`+`IsWindowVisible`) at
  server start, `excel_open`, unclassified macro_run, cycle-1 and cycle-2 macro_run;
* owned PID reaped after shutdown (`_wait_pid_dead`); every pre-existing user PID survives
  (`surviving_user_pids` empty);
* source fixture + temp copy byte-identical (PO-6).

Phase-4 report `.ai/audit/reports/xlvba-mcp-phase4-e2e-acceptance.md` recorded these as passing
(5/5 live tests, fixture SHA-256 unchanged). This audit **independently reproduced** the hidden/
non-interference behavior with a separate read-only probe rather than relying on that report.

---

## 7. Additional observations (advisory, non-blocking)

* **`assert_non_intrusive` is dead code.** `ExcelAppAdapter.assert_non_intrusive` (and its helper
  `_owned_window_is_foreground`) are defined but **never called** from any runtime path — the only
  grep hits are the definitions themselves. The `Visible=False` invariant is actually enforced at
  creation time (`create_hidden` quits the instance if `Visible` cannot be set to `False`), so hidden
  execution is genuinely guaranteed. The unused non-intrusive guard is either future-proofing or
  leftover; per the dead-code policy, investigate intent before removing. It does **not** weaken
  safety because `create_hidden` is the real enforcement point.
  *Effort:* trivial · *Priority:* advisory.
* **`macro_run` has no snapshot precondition** (CT-B7 / D-7) — unrelated to hidden execution; already
  documented as a known limitation in `119c303`. No action for FE-14.
* **`excel_health` has no PID-liveness probe** (D-2 / R-17) — documented limitation; the worker
  reaps and reports correctly in the normal lifecycle observed here.

---

## 8. Verdict

9 of 10 Phase-5 items are objectively **PASS** (hidden owned instance, interactive host session,
user instances untouched, zero owned visible windows, zero forbidden-API call sites, interactive-macro
veto with no dialog, comtypes read-only UI automation with fail-closed compile, worker health stable,
owned PID reaped / user PIDs alive). Item 10 (focus-theft) is **NOT VERIFIED**: point samples showed
the owned Excel never foreground and never visible, but no continuous foreground logger was available
to prove a transient focus grab cannot occur.

**PHASE 5: PARTIAL** — the only gap is the continuous focus-theft trace (item 10); the evidence needed
to close it is a continuous `EVENT_SYSTEM_FOREGROUND` / high-frequency `GetForegroundWindow` logger
over the full server lifecycle. All other desktop non-interference and hidden-execution checks pass.

---

## 9. Reproduction

```
# Focus probe (read-only; real MCP STDIO server; temp copy)
.venv\Scripts\python.exe .ai\tmp\fe14_focus_probe.py

# Host session
Get-Process -Id $PID | Select-Object Id, SessionId, ProcessName
qwinsta

# Forbidden-API sweep
Get-ChildItem -Recurse src -Filter *.py | Select-String `
  -Pattern "SetForegroundWindow|ShowWindow|SetActiveWindow|BringWindowToTop|SendKeys|AppActivate|keybd_event|mouse_event|SetCursorPos|PostMessage|SendMessage|SetFocus"
```