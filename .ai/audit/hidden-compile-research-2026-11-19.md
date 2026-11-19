# Hidden / Non-Activating VBA Compilation — Research Report

**Author:** researcher
**Date:** 2026-11-19
**Status:** complete (research only; no production code under `src/` or `tests/` modified)
**Scope:** Read-only investigation + throwaway probes under `.ai/tmp/`. Answers whether VBA
compilation can be **invoked and diagnosed non-intrusively** (hidden, non-activating) on the
validated target, and what the truthful `task_020` contract should be.

**Validated target (all [observed] evidence below was executed live here):**

| Attribute | Value |
|---|---|
| OS | Windows 10 Enterprise LTSC 2019, build 17763.9020 x64 |
| Excel | 16.0.14332 x64 (Office 2016 family / Version 16.0) |
| VBE | Microsoft Visual Basic for Applications 7.01 |
| Office / VBE UI language | 1033 (en-US) |
| VBOM trust | enabled (`AccessVBOM=1`, read-only) |
| Tooling | Python 3.14.7, pywin32 312, comtypes 1.5.0, pywinauto 0.6.9 |

**Confidence legend:** `[observed]` = executed live on the target build; `[doc]` = Microsoft
Learn / Microsoft Support; `[community]` = Stack Overflow / vendor blog (LOW-MEDIUM, unverified
here); `[inference]` = reasoned conclusion.

---

## 1. Scope

The question is **not** "can we compile VBA?" (we can, via UIA) but "can we compile VBA
**without showing or activating any Excel/VBE window, and without the automation stealing
foreground focus**?" The controlling requirement is the plan addendum
`.ai/plans/plan_addendum_background_execution.md` (BG03 no focus theft, BG04 no visible popup,
BG07 hidden compile evidence, BG09 no blind dialog handling).

Prior evidence base (all under `.ai/tmp/`):

- `probe_hidden_compile{,2,3,4,5}.py` + `*_result.json` — hidden-VBE compile behaviour.
- `probe_run.py/.json` — `Application.Run` arg/return/error semantics.
- `probe_vbide{,2}.py/.json` — VBIDE export, `References`/`IsBroken`, `Protection`.
- `probe_compile_deep.json`, `probe_compile_error.json`, `probe_compile_discovery2.json`,
  `probe_compile_tp.json`, `probe_vbe_compile_raw.json` — command/popup/dialog structure.
- **New this pass:** `probe_run_compile_error.py` + `probe_run_compile_error_result.json` —
  whether `Application.Run` is a compile diagnostic (Q3a).

No file under `src/` or `tests/` was created or modified.

---

## 2. Findings per question

### Q1 — Is there a documented, non-activating way to invoke VBE Compile on 16.0.14332 with `Application.Visible=False`?

**Answer: No. Every route either does not reach Compile, or activates/shows UI.** `[observed]`

#### 2.1 `VBE.CommandBars["Debug"]` — Compile control is absent on this build

`VBE.CommandBars["Debug"].Controls` has **14 controls and none is Compile** `[observed]`
(`probe_hidden_compile_result.json`, identical on the visible build in
`probe_vbe_compile_raw.json`):

```
212 Design Mode, 5415 Run Project, 186 Run Macro, 189 Break, 228 Reset,
51 Toggle Breakpoint, 188 Step Into, 194 Step Over, 2559 Step Out,
2555 Locals Window, 2554 Immediate Window, 2556 Watch Window,
229 Quick Watch..., 620 Call Stack...
```

The VBE `CommandBars` **are** populated even when the VBE main window was never shown
(`bars_count = 30`, Debug controls enumerable) — so "not populated when hidden" is **false**;
the bar simply has no Compile entry on this build. `[observed]`

#### 2.2 Control ID 578 (`FindControl`) — not usable on this build

The community-documented "Compile VBAProject" id is **578** `[community]`
(Stack Overflow 28001620, 51030992, 32390464, 55621735, 78356855). On the target build it does
**not** work:

- `bars["Debug"].FindControl(578)` and every argument variant `(578)`, `(578,"Compile")`,
  `(578,"",True)`, `(578,"",True,1)` → `com_error 0x80020009` with inner `-2147024809`
  (`E_INVALIDARG`). `[observed]`
- Even `FindControl(2554)` for a control that **is** in the Debug bar raises `E_INVALIDARG`
  (`probe_hidden_compile2_result.json` → `execute_2554_error`). So `FindControl` is effectively
  non-functional on this build's VBE CommandBars, not merely missing id 578. `[observed]`

Microsoft documents `CommandBar.FindControl` as "returns `Nothing` if no control fits"
`[doc]`; the `E_INVALIDARG` we get instead is a build-specific implementation difference.

#### 2.3 `CommandBarControl.Execute` — no control to execute

`Execute` runs a located control's action `[doc]`; because no Compile control can be located
(2.1/2.2), `Execute` has nothing to invoke. Community reports that on builds where 578 exists,
`Execute` still **shows a modal message box on compile error** and requires a watchdog thread
to detect success (control becomes disabled) `[community]`.

#### 2.4 `Application.CommandBars.ExecuteMso` — wrong object model

`ExecuteMso(idMso)` operates on the **host application's** CommandBars (the Ribbon), returns
`E_InvalidArg` for an invalid `idMso` and `E_Fail` for disabled/not-visible controls, and is
"useful when there is no object model for a particular command" `[doc]`. The VBE is a **separate
application** with its own `VBE.CommandBars` (no Ribbon); `Application.CommandBars` is Excel's.
There is no `idMso` for VBE Debug→Compile. **Not a viable route.** `[inference from doc]`

**Conclusion Q1:** No documented non-activating route exists for this build. The only route that
actually compiles is UIA popup expansion, which activates/shows UI (§Q4).

---

### Q2 — Can compile be triggered through the VBIDE object model or any Office API without UI?

**Answer: No.** `[doc]` + `[observed]`

- The VBIDE object model (`VBE`, `VBProject`, `VBComponents`, `VBComponent`, `CodeModule`,
  `Reference`, `References`) exposes **properties and source/CRUD operations only**. `VBProject`
  is documented as "set properties for the project, and … access the VBComponents and References
  collections" `[doc]`. There is **no `Compile` method, no `Build`, no p-code access** anywhere in
  the model `[doc]`.
- All VBIDE source access works fine while hidden and non-activating: `Workbook.VBProject`,
  `Protection`, `References`, `VBComponents`, `CodeModule.Lines/AddFromString`, `Export/Import`
  all succeeded with the VBE main window `Visible=False` and no foreground change
  `[observed]` (`probe_hidden_compile2_result.json`: `vbe_main_window_visible_after_vbproject:
  false`, `fg_after_vbproject` = the user's editor; `probe_vbide2.json` round-trip).
- `Application.Run` is the only "compile-ish" side effect, and it is **not** a compile trigger
  (see Q3a) — it compiles on demand per-procedure, and only the executed procedure's call tree.

**Conclusion Q2:** The VBIDE model gives full source access hidden, but **no compile entry point**.
Compilation is a VBE *command*, reachable only through the VBE UI/command layer.

---

### Q3 — Alternative non-intrusive diagnostics

#### 3a. Does `Application.Run` on a clean procedure surface compile errors (HRESULT) without UI?

**Answer: No — `Application.Run` is NOT a compile diagnostic.** `[observed]`

New probe `probe_run_compile_error.py` built one workbook with `modMix` (`Good()` clean +
`Bad()` = `x = UndefinedIdentifier`) and `modGood` (`Good2()` clean), opened it hidden, and ran
three cases (watchdog force-kills only the automation-owned PID if a `#32770` appears):

| Case | Macro | Result | Dialog? | Time |
|---|---|---|---|---|
| T1 | `modMix.Good` (clean proc, **same module** as broken `Bad`) | **ok** | none | 0.051 s |
| T2 | `modGood.Good2` (clean proc, **different** module) | **ok** | none | 0.004 s |
| T3 | `modMix.Bad` (the broken proc) | **blocked** | **`#32770` "Compile error:\n\nVariable not defined"** | 1.05 s (killed) |

Interpretation:

- **T1/T2 prove VBA "compile on demand":** a clean procedure runs successfully even when another
  procedure *in the same module* and other modules contain a compile error. `Application.Run`
  never compiles unrelated code, so it **cannot** detect a project-wide compile error.
- **T3 proves the failure path is a modal, not an HRESULT:** running the broken procedure blocked
  on the same `#32770` compile-error dialog used by VBE Debug→Compile. The `0x800706BE`
  (`RPC_S_CALL_FAILED`) we recorded is only because the watchdog **force-killed** the process —
  it is *not* a clean compile HRESULT. Excel showed the dialog instead of raising a catchable
  error (consistent with community reports: "Excel will display a MessageBox instead of throwing
  an exception" `[community]`).

> **Caveat (honesty):** this probe ran while the Windows session was **locked** (foreground =
> "Windows Default Lock Screen"), so its `visible:false` flag and "no focus change" for T3 are
> **not** representative of an active desktop. The compile-on-demand verdict (T1/T2) is engine
> behaviour and is lock-independent. Active-desktop evidence that the *same* dialog appears
> **visible** and takes foreground is in `probe_hidden_compile4_result.json` and
> `probe_hidden_compile5_result.json` (see Q4).

**Conclusion Q3a:** `Application.Run` must **not** be used to claim compile success, and its
failure path is a visible modal. Not a non-intrusive diagnostic.

#### 3b. Can missing/broken references be detected via `VBProject.References` (`IsBroken`)?

**Answer: Yes — this is a real, non-intrusive partial diagnostic.** `[observed]` + `[doc]`

- `VBProject.References` returned 5 references with `Name`, `GUID`, `Major.Minor`, `FullPath`,
  `IsBroken` all readable while hidden (`probe_vbide.json`). `[observed]`
- Microsoft documents: "The `IsBroken` property returns `True` if the reference no longer points
  to a valid reference." `[doc]`
- Reading references does not show/activate the VBE. This detects **missing/broken type-library
  references** — a distinct and common compile-blocking failure mode — but it is **not** a code
  compile and cannot detect syntax/semantic errors in VBA source.

#### 3c. Is there a p-code / `CodeModule` API that reports compile state?

**Answer: No.** `[doc]`

- `CodeModule` exposes `AddFromString`, `Lines`, `CountOfLines`, `ProcOfLine`, `ProcStartLine`,
  `ProcCountLines`, `InsertLines`, `DeleteLines`, `CreateEventProc`, etc. — source text and
  procedure boundaries only. No compiled state, no p-code, no `IsCompiled`.
- `VBProject` exposes `Name`, `Description`, `Protection`, `References`, `VBComponents`, `Type`,
  `Saved`, `Collection`, `Parent`, `VBE` — no compile state.
- The only observable proxy for "just compiled successfully" is behavioural: the **Compile menu
  item becomes disabled** after a successful compile (VBE built-in behaviour, `[community]` and
  `[observed]` in `probe_compile_deep.json`: `clean_compile_item_after.enabled: 0`). But reading
  that requires opening the menu (UI), and a fresh uncompiled project can also report it disabled
  — so it is not a safe standalone signal (§Q5).

---

### Q4 — Can the VBE be kept hidden but functional (menus/CommandBars usable) without showing/activating it?

**Answer: No.** `[observed]`

- `VBE.MainWindow.Visible = False` / `Application.Visible = False` does keep the VBE main window
  hidden (`vbe_main_window_visible: false` throughout `probe_hidden_compile2_result.json`), and
  accessing `VBProject` does **not** itself show or foreground the VBE (`fg_after_vbproject` =
  user editor). So *reading source* hidden is fine.
- But the Compile command lives in an **owner-drawn CommandBar popup**, not a Win32 menu:
  - `GetMenu(vbe_hwnd)` returns **0** — the VBE main window has no standard HMENU
    (`probe_hidden_compile3_result.json`), so `WM_COMMAND` routing is impossible.
  - Expanding the Debug menu (the only way to reveal `Compile VBAProject`) creates **4 visible
    windows owned by our PID** — `MsoCommandBarShadow` ×3 + `MsoCommandBarPopup` "Debug" at
    on-screen coordinates (e.g. rect `[206,55,456,334]`) — **and makes the hidden VBE the
    foreground window** (`fg_after_expand` = the VBE hwnd, `visible:false`). `[observed]`
    (`probe_hidden_compile2_result.json`, `probe_hidden_compile4_result.json`,
    `probe_hidden_compile5_result.json`).
- Invoking Compile on the **clean** fixture: popup closes, no dialog, but the hidden VBE is
  **left as the foreground window** (`probe_hidden_compile4_result.json`, CleanProject case →
  `fg_after_invoke` = VBE). That alone violates **BG03**.
- Invoking Compile on the **error** fixture: a **visible** `#32770` "Microsoft Visual Basic for
  Applications" dialog appears (on-screen rect, foreground), text "Compile error:\n\nVariable not
  defined". That violates **BG04**, and `PostMessage(WM_COMMAND, IDOK)` does **not** dismiss it
  (`dialog_still_open: true`). `[observed]` (`probe_hidden_compile4_result.json`,
  `probe_hidden_compile5_result.json`).
- Off-screen positioning is not a fix: the popup is created on `Expand()`, and positioning it /
  activating it is itself a focus-changing UI action; `Application.Visible=False` does not
  suppress these CommandBar popup windows.

**Conclusion Q4:** "Hidden but functional menus" is unachievable. Merely *expanding* the Debug
menu violates BG03 (focus steal) + BG04 (visible popup); a compile error additionally shows a
visible modal that cannot be dismissed non-intrusively.

---

### Q5 — Correct, truthful contract for `task_020` on this build

**Answer: compile status must be `COMPILE_INDETERMINATE` (fail-closed) on this build; report
non-compile partial diagnostics separately and clearly labelled.** `[inference from observed]`

Recommended contract for `CompileService.compile_project` in **background mode**:

1. **Do not** expand the Debug menu / invoke Compile (it violates BG03/BG04). Do **not** run
   `Application.Run` as a compile proxy (Q3a).
2. Return status = **`COMPILE_INDETERMINATE`** with a structured reason, e.g.
   `NON_ACTIVATING_COMPILE_UNAVAILABLE` (evidence: CommandBars has no Compile control;
   `FindControl` E_INVALIDARG; UIA route requires visible popup + foreground; error path shows a
   non-dismissible modal).
3. Attach a **`partial_diagnostics`** block (clearly labelled **"NOT a compile"**):
   - `references`: name/guid/version/fullpath/**`is_broken`** (Q3b) — real, non-intrusive.
   - `protection`: `VBProject.Protection` (locked → `VBA_PROJECT_LOCKED`).
   - `components`: name/type/`CountOfLines` inventory (already available from task_016).
   - `access_vbom`: `ACCESS_VBOM_DISABLED` when `VBProject` raises `0x800A17B4`.
4. Never return `compiled` or `compile_error` on this build through automation. `already_compiled`
   is not knowable non-intrusively (the only signal is the disabled menu item, which needs UI).
5. Keep any UI-based compile behind an **explicit, non-default "interactive dev diagnostic"**
   flag that is **never** used by background/automated flows, and is documented as violating
   BG03/BG04. (Optional; recommended default = absent.)

Pseudocode:

```text
compile_project(session) -> CompileResult:
    if not session.vbom_available: return INDETERMINATE(ACCESS_VBOM_DISABLED)
    partial = {
        "references": [ {name, guid, version, full_path, is_broken} ... ],   # Q3b, non-intrusive
        "protection": session.protection,                                    # 0 / locked
        "components": session.component_inventory(),                         # name/type/lines
        "note": "partial diagnostics only - NOT a VBA compile",
    }
    if session.protection != 0:
        return BLOCKED(VBA_PROJECT_LOCKED, partial)
    # Non-activating compile is NOT available on this build (see research report).
    return INDETERMINATE(NON_ACTIVATING_COMPILE_UNAVAILABLE, partial)

    # (Never reached in background mode) interactive dev-only path:
    #   requires opt-in flag AND accepts visible popup + foreground + possible modal.
```

This is exactly what the addendum prescribes: *"A dialog that cannot be read non-intrusively
produces `COMPILE_INDETERMINATE`; it must not be clicked or dismissed"* and *"If the target
Office build displays or activates VBE/dialog windows as an unavoidable part of compile-error
handling, the PoC fails the background requirement."*

**Note on task_004/task_020 spec text:** the current plan still lists `compiled` /
`compile_error` as required production outcomes (task_020 acceptance criteria). On this build
those two are **not achievable non-intrusively**; the acceptance criteria must be amended to
"success and error fixtures both yield a truthful `COMPILE_INDETERMINATE` with a documented
reason, and never a false `compiled`", plus the partial-diagnostics block. This is a spec
correction, not a scope expansion.

---

### Q6 — Other Office builds/versions where hidden compile is known to work (release matrix / limitations)

**Answer: There is no authoritative source for a supported non-activating hidden compile on any
build; the id-578 route works on some builds but still shows UI on error.** `[doc]` + `[community]`

- **`FindControl(Type:=msoControlButton, ID:=578).Execute`** is widely reported to compile VBA on
  VBA-capable Excel builds (2010–2019/365 era) **when run from inside Excel / with the VBE
  visible** `[community]` (SO 28001620, 51030992, 32390464, 55621735, 78356855, 49286336).
  Build-dependent: on **16.0.14332 it does not exist / `FindControl` raises `E_INVALIDARG`**
  `[observed]`. So 578 is *not* a universal identifier — treat as a candidate to verify per build.
- Even where 578 works, community reports the compile **error path shows a modal message box**
  (not an exception) and success must be inferred from the control becoming disabled via a
  watchdog thread `[community]` — i.e., **not non-intrusive**.
- **Microsoft does not support unattended/non-interactive Office automation** and warns that
  Office "assume[s] an interactive desktop", that "a modal dialog box on a non-interactive
  desktop cannot be dismissed … [the] thread stops responding (hangs) indefinitely"
  `[doc]`.
- No Microsoft Learn page documents a `Compile`/`CompileProject` object-model method or a
  supported headless compile.

**Release-matrix wording recommendation:**

> Hidden / non-activating VBA compilation is **not supported** on any tested Office build.
> Validated target: Windows 10 LTSC 2019 (17763) x64, Excel 16.0.14332 x64, UI 1033, VBOM on.
> On this build the only working compile route (UIA Debug menu expansion) shows a visible popup,
> steals foreground, and on error shows a non-dismissible modal → `COMPILE_INDETERMINATE`.
> Community id-578 `CommandBars` compile works on some builds with the VBE visible but still
> shows UI on error; unverified and out of scope. See `.ai/audit/hidden-compile-research-2026-11-19.md`.

---

## 3. Options comparison

| Option | Reaches Compile? | Non-activating (BG03/BG04)? | Error diagnostic non-intrusive? | Verdict |
|---|---|---|---|---|
| `VBE.CommandBars["Debug"]` control lookup | **No** (no Compile control) | n/a | n/a | Dead on 14332 |
| `FindControl(578)` + `Execute` | **No** (E_INVALIDARG on 14332; works on some builds) | No (VBE visible / modal on error) | No | Dead here; build-specific elsewhere |
| `Application.CommandBars.ExecuteMso` | **No** (host ribbon, not VBE) | n/a | n/a | Not applicable |
| Win32 `WM_COMMAND` to VBE hwnd | **No** (`GetMenu`=0) | n/a | n/a | Dead |
| **UIA Debug `Expand()` → `Compile VBAProject` `Invoke()`** | **Yes** | **No** — 4 visible popup windows + VBE becomes foreground; error → visible modal | **No** (dialog visible, not dismissible via PostMessage) | **Only working route; violates BG03/BG04** |
| `Application.Run(clean proc)` as compile check | **No** (compile-on-demand; unrelated errors missed) | Yes (clean proc) | N/A | **Not a compile diagnostic** |
| `Application.Run(broken proc)` | n/a | **No** — blocks on `#32770` modal | No | Not usable |
| **`VBProject.References` → `IsBroken`** | Partial (reference errors only) | **Yes** | **Yes** | **Recommended partial diagnostic** |
| `VBProject.Protection` read | n/a (locked detection) | **Yes** | Yes | Recommended partial diagnostic |
| `CodeModule`/p-code compile-state API | **No such API** | n/a | n/a | Does not exist |

---

## 4. Recommendation

1. **`task_020` contract on the validated build:** `vba_compile` / `compile_service` must
   **fail closed to `COMPILE_INDETERMINATE`** with reason `NON_ACTIVATING_COMPILE_UNAVAILABLE`,
   and must **not** invoke the UIA menu route (which shows UI / steals focus). Attach
   `partial_diagnostics` = broken-reference report (`References.IsBroken`), `Protection`, component
   inventory, and `ACCESS_VBOM_DISABLED` detection — explicitly labelled **"NOT a compile"**.
2. **Amend task_020 acceptance criteria** (and task_004 goal wording) so that "success fixture →
   compiled" and "error fixture → compile_error" are replaced by "both fixtures → truthful
   `COMPILE_INDETERMINATE` (or `VBA_PROJECT_LOCKED`) with documented reason; never a false
   `compiled`", plus the partial-diagnostics block. This matches addendum BG07/R12.
3. **Do not** use `Application.Run` as a compile proxy (proven unsound: compile-on-demand; error
   path is a modal).
4. **Keep an optional, non-default "interactive dev diagnostic"** (UIA route) for local
   diagnostics only, gated behind an explicit flag and documented as violating BG03/BG04; never
   reachable from background/automated flows.
5. **Document the limitation** in `docs/release-evidence.md` / README per the Q6 wording, and
   record that `task_004` is a **No-Go for non-activating compile** on this build (the compile
   capability is blocked, not silently substituted).

**Bottom line:** non-activating compile is **not feasible** on the validated build
(**No** for a full compile; **Partial** only for reference/locked/inventory diagnostics).

---

## 5. Constraints / Risks

- **Build/locale fragility:** the id-578 route and the literal caption `Compile VBAProject` are
  build/locale-specific. A future Office update or UI-language change can remove/rename them;
  the service must fail closed (`COMPILE_INDETERMINATE`), never guess.
- **Compile-on-demand:** any "run a proc to test compilation" idea is unsound (Q3a) — errors in
  unexecuted procedures are invisible. Do not ship it as a compile check.
- **Modal deadlock:** the compile-error `#32770` cannot be dismissed with
  `PostMessage(WM_COMMAND, IDOK)`; only a process kill (our own PID) clears it. Background code
  must **not** attempt to dismiss it and must not report success while it is open.
- **Partial diagnostics are not compile:** `IsBroken`/inventory can be all-green while the code
  has a syntax error. They must never be presented as "compiled".
- **Locked workstation confound:** some observations were taken with the session locked; the
  compile-on-demand result is lock-independent, but active-desktop focus/visibility claims rely on
  `probe_hidden_compile4/5` (captured with an active desktop).
- **Orphan process risk:** killing the automation instance is required to clear a stuck modal;
  ensure `finally` cleanup and PID-scoped kill only (never touch user-owned Excel).

---

## 6. References

Microsoft Learn / Support:

- VBIDE object model overview (VBProject has no Compile method): *Objects (Visual Basic Add-In
  Model)* — https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model
- VBIDE properties (`Protection`, `References`, `IsBroken` on Reference): *Properties (Visual
  Basic Add-In Model)* — https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model
- `Workbook.VBProject` (Excel): https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbproject
- `CommandBar.FindControl` (Office): https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.findcontrol
- `CommandBars.FindControl` (Office): https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.findcontrol
- `CommandBars.ExecuteMso` (Office, host ribbon): https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso
- `Application.CommandBars` (Excel): https://learn.microsoft.com/en-us/office/vba/api/excel.application.commandbars
- `Application.Run` (Excel): https://learn.microsoft.com/en-us/office/vba/api/Excel.Application.Run
- `Application.Visible`: https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible
- *Considerations for server-side Automation of Office* (Microsoft does not support unattended
  Office automation; modal dialog hangs): https://support.microsoft.com/en-gb/visio/considerations-for-server-side-automation-of-office
- *Programmatic access to Office VBA project is denied* (`0x800A17B4`): https://support.microsoft.com/en-gb/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445

Community (LOW-MEDIUM, unverified here, cited for the id-578 route and modal behaviour):

- SO 28001620 — *Can I compile VBA on workbook open?* (id 578 `FindControl(...).Execute`)
- SO 51030992 — *A Way to Automate the "Compile" Function of MS Office's VBA Code*
- SO 32390464 — *Trap VBE compile error when running editor under VBA control* (compile button
  disabled = success)
- SO 78356855 — *How to intercept VBE errors during automation?* (watchdog: button-disabled vs
  `#32770` dialog; localized text)
- SO 55621735 — *VBA - Method to Detect Compilation Failure* (errors surface as message boxes)
- SO 49286336 — *Programatically compiling VBA modules with C#* (external vs in-process Execute)

Project evidence (`.ai/tmp/`): `probe_hidden_compile{,2,3,4,5}_result.json`,
`probe_run_compile_error_result.json` (new), `probe_run.json`, `probe_vbide{,2}.json`,
`probe_compile_deep.json`, `probe_compile_error.json`, `probe_compile_discovery2.json`,
`probe_compile_tp.json`, `probe_vbe_compile_raw.json`.

Prior research: `.ai/audit/g1-poc-research-2026-11-19.md`;
addendum: `.ai/plans/plan_addendum_background_execution.md`.

---

## 7. One-line answer

Non-activating (hidden, non-focus-stealing) VBA compilation is **not feasible** on
Excel 16.0.14332 / Win10 LTSC 2019 / UI 1033: the only working route (UIA Debug-menu expansion)
shows a visible popup, steals foreground, and on error shows a non-dismissible modal. The correct
`task_020` contract is **fail-closed `COMPILE_INDETERMINATE`** plus clearly-labelled
**non-compile partial diagnostics** (broken references via `References.IsBroken`, `Protection`,
component inventory); `Application.Run` is not a compile diagnostic (compile-on-demand).