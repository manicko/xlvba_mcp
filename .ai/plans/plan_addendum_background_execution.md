# Plan Addendum — Non-Intrusive Background Excel Execution

**Applies to:** `plan.md` — *Excel VBA MCP Server — Implementation Execution Plan*  
**Status:** Corrective addendum; does not replace or regenerate `plan.md`.  
**Scope:** Stage 1 only.  
**Decision:** Excel/VBA operations must run in the user's existing interactive Windows session without displaying or activating Excel/VBE windows over the user's other applications. A VM, container, separate Windows logon/RDP session, alternate desktop, or Windows Service is not an acceptable implementation or fallback.

---

## 1. Purpose and normative decision

The implementation plan must be interpreted and executed with a strict **non-intrusive background execution** requirement.

The MCP server and its Excel automation worker run as the currently logged-in user in the user's normal Windows desktop session. Excel automation must not steal focus, activate Excel/VBE, show a workbook window, or display a dialog over the user's active applications.

A dedicated automation-owned `Excel.Application` instance may still be created through COM (the existing `DispatchEx` direction remains valid), provided that it runs in the **same logged-in Windows session and desktop** and remains hidden/non-activating. This is process/instance ownership for safe automation; it is **not** a separate execution environment.

### Explicitly prohibited

The implementation must not rely on:

- Hyper-V, VMware, VirtualBox, or any other virtual machine;
- a container or separate isolated Windows image;
- a separate Windows user account, logon session, RDP session, or alternate desktop;
- a Windows Service, non-interactive service desktop, or `SYSTEM`-scheduled Office automation;
- showing Excel/VBE and asking the user to leave it in the foreground;
- sending mouse/keyboard input to the user's desktop, using `SendKeys`, or bringing Excel/VBE to the foreground as a normal execution path.

There is no VM/isolated-environment fallback for a feature that fails the background requirement. If a capability cannot be made non-intrusive in the target Excel build, record it as blocked/indeterminate and continue with independent work; do not silently switch to a visible or isolated execution mode.

## 2. Required background execution contract

### 2.1 Excel instance ownership and session

- Start the MCP process as a normal process under the user's current logged-in Windows account.
- Create and own a dedicated Excel COM instance in that same session. Do not attach to or manipulate the user's already-open Excel instance.
- Keep all Excel COM objects and proxies inside the existing dedicated STA worker; all COM operations remain serialized.
- The automation-owned Excel instance must remain hidden for its entire lifetime, including open, read, write, compile, macro execution, save, snapshot, restore, and cleanup operations.
- The worker must close only workbooks and the Excel instance it owns. It must never close, change, or reconfigure user-owned Excel sessions.

### 2.2 No foreground activation or visible windows

The normal execution path must not:

- set Excel or VBE visible;
- activate an Excel workbook, worksheet, range, or VBE window;
- use `.Activate` or `.Select` to perform object-model operations;
- call `SetForegroundWindow`, `ShowWindow` (to show/restore), UIA `SetFocus`, `click_input`, `type_keys`, or equivalent focus-changing actions;
- show a workbook, VBE, modal dialog, UserForm, or helper window on the user's desktop.

Use direct object-model references rather than UI navigation. `Application.Visible = False` is the principal visibility control. `ScreenUpdating = False` is an additional rendering optimization, **not** a substitute for hiding the application or preventing dialogs.

The worker must preserve and restore any application-level settings it changes. Restoration must happen in `finally`-style cleanup on success, failure, timeout, and cancellation. Do not change settings on the user's existing Excel instance.

### 2.3 Alerts, dialogs, and macro security

`Application.DisplayAlerts = False` suppresses only certain Excel alerts and can cause Excel to choose a default response. It must therefore not be treated as a universal dialog blocker or set blindly for the whole worker lifetime.

- Use alert suppression only around a specific operation whose default response is known and safe.
- Save the prior value and restore it on every exit path.
- Never rely on `DisplayAlerts = False` to suppress VBA `MsgBox`, UserForms, `FileDialog`, third-party add-in UI, or arbitrary Win32 dialogs.
- Do not weaken Trust Center settings or enable “Trust access to the VBA project object model” programmatically.
- Preserve the existing macro approval/allowlist policy. Background execution is not permission to run unknown macros.
- Do not automatically press Enter, Escape, or click a dialog to make a blocked operation appear successful.

### 2.4 Interactive VBA behavior

A macro may itself display UI even when the automation-owned Excel instance is hidden. Examples include `MsgBox`, `UserForm.Show`, `Application.FileDialog`, controls/forms, and calls into add-ins or external applications.

The background execution contract applies to macro behavior as well as MCP code:

- Before execution, use available VBA source/inventory information for a **best-effort UI-interaction preflight**. Detection is heuristic; it must not be represented as complete static proof.
- Extend the existing macro allowlist/approval record so approval is bound not only to workbook identity and procedure, but also to an explicit `background_safe` classification. A macro is not eligible for strict background execution until that classification has been reviewed and its non-interactive behavior has been verified on a disposable copy. A VBA source/workbook hash change invalidates the prior classification.
- If known interactive UI calls are found, or the procedure has not been classified as background-safe, block it in background mode with a structured status such as `INTERACTIVE_UI_NOT_SUPPORTED` / `BACKGROUND_SAFETY_NOT_VERIFIED` (or equivalent domain errors agreed in the existing status model).
- During execution, monitor for Excel/VBE-owned modal or visible windows without activating or interacting with them.
- If a dialog or UI-dependent state is detected or suspected, stop issuing further COM commands to that session, capture only non-intrusive diagnostics, and return a truthful blocked/indeterminate result.
- Never dismiss a dialog by sending input. Never report a macro as passed when it is waiting for user interaction.

The preflight is a risk filter, not a guarantee: dynamic calls, COM add-ins, external executables, and indirect procedure calls may create UI without an obvious direct call in the inspected procedure. Runtime monitoring and truthful status reporting remain mandatory.

## 3. VBA compilation: revised non-intrusive requirement

VBA compilation remains in Stage 1 and remains a required capability. The compile implementation must satisfy the same background contract; making VBE visible is not an acceptable workaround.

### Required investigation and PoC

`task_004` must test the compile workflow against both a clean fixture and a fixture with a deliberate compile error while the automation-owned Excel instance remains hidden and non-activating.

Researcher must investigate, within the current architecture, command invocation through the VBE object model (including `VBE.CommandBars`, `FindControl`, control metadata, and `CommandBarControl.Execute`) and any supported non-activating diagnostic route. Command IDs found in examples are candidates to verify against the target Office build, not hard-coded universal identifiers. Microsoft documents `FindControl` and `Execute`; these references do not themselves guarantee that a VBE compile error can be collected without UI appearing.

### Compile PoC acceptance

The PoC must establish all of the following on the declared target Windows/Excel build:

1. The compile command can be invoked without setting Excel or VBE visible and without activating either window.
2. A successful compile can be positively distinguished from `already_compiled` and `indeterminate`.
3. A deliberate compile error can be detected and its available diagnostic information collected without bringing VBE or a modal dialog to the foreground.
4. No Excel/VBE-owned window steals foreground focus or appears over the user's active application during either test.
5. No mouse/keyboard injection, `SendKeys`, coordinate clicking, or user intervention is required.
6. If diagnostic text cannot be obtained non-intrusively, the result is `COMPILE_INDETERMINATE` (or equivalent), never a false `compiled` result.

UI Automation may be used only for non-activating observation/read-only diagnostics. It must not call focus, click, type, show, restore, or foreground operations. Screenshots are optional evidence only when they can be captured without making a window visible or active; they are not a mandatory acceptance artifact under this addendum.

If the target Office build displays or activates VBE/dialog windows as an unavoidable part of compile-error handling, the PoC fails the background requirement. Record the limitation and investigate another same-session, non-activating approach. Do not fall back to a VM, separate desktop, visible VBE, or user-operated dialog.

## 4. Amendments to existing tasks

The following task specifications in `plan.md` must be interpreted or updated as described below. These are targeted amendments only; all other task scope and acceptance criteria remain unchanged.

| Existing task | Required correction / review | Required evidence |
|---|---|---|
| `task_001` — Bootstrap | Launch/configuration instructions must run the MCP server as the logged-in user in the current interactive session. Document that Windows Service, `SYSTEM`, VM, container, and separate desktop/session execution are unsupported and prohibited for this project. | Startup proof under the normal user session; no service/alternate-session dependency. |
| `task_002` — Excel COM / STA PoC | Rename the semantic objective from “isolated Excel COM” to “automation-owned hidden Excel instance in the current interactive session.” Keep `DispatchEx` if PoC confirms a distinct automation-owned instance, but remove any wording that implies VM/session isolation. Verify the user's existing Excel instance is untouched and the automation instance never becomes visible/foreground. | Record process/session ownership, visibility state, foreground-window observations, and cleanup results while the user desktop remains usable. |
| `task_003` — VBIDE access PoC | Verify VBProject enumeration, source reading, and export while Excel/VBE remains hidden. AccessVBOM remains a manually configured Trust Center setting; do not reveal VBE to make source access work. | Read/export succeeds with no Excel/VBE foreground activation or visible window. |
| `task_004` — VBE compile/UIA PoC | Replace the assumption that UIA can interact with a visible compile dialog with the strict compile PoC in Section 3. Command invocation and diagnostic observation must be non-activating. UIA is read-only only. Screenshot capture becomes optional. | Success and compile-error fixtures pass the no-visible/no-foreground requirement; otherwise report a truthful failed/indeterminate PoC. |
| `task_005` — `Application.Run` PoC | Test successful, missing, and runtime-error macros with Excel hidden. Add cases for known interactive UI calls (for example `MsgBox` and `UserForm.Show`) and verify they are blocked/detected rather than surfaced over the desktop. | No focus steal or visible popup; structured runtime/UI-blocked results. |
| `task_006` — Snapshot PoC | Keep snapshot and restore in the same Windows session and same automation-owned Excel instance lifecycle. “Independent test copy” means a separate workbook file/copy, not a separate VM or Windows environment. | Snapshot/restore succeeds without showing or activating Excel. |
| `task_009` — Approval and audit | Extend the existing procedure allowlist/approval record with an explicit `background_safe` classification bound to workbook hash and exact procedure. Require this classification for strict background macro execution; source/hash changes invalidate it. Preserve macro approval and Trust Center rules. | Audit record distinguishes background-safety approval, blocked interactive UI, compile indeterminate, runtime failure, and success. |
| `task_010` — COM STA worker | Extend the existing `com_worker.py` / `app_adapter.py` responsibilities to initialize the owned Excel instance as hidden and non-activating in the current session. Centralize controlled application settings and restore them reliably. Do not introduce a VM worker or second desktop. | STA worker operates hidden, serializes COM, and never changes user-owned Excel state. |
| `task_011` — COM worker verification | Add visibility/focus tests across worker startup, workbook open, operation, error, shutdown, and recovery. “Isolation” means ownership separation from the user's existing Excel instance, not environment isolation. | No server-owned Excel/VBE window becomes foreground/visible; user's Excel instance is unchanged; no orphan owned process. |
| `task_012` / `task_013` — Workbook session | Ensure `excel_open`, reads, writes, save, close, and status calls never show/activate workbook windows. Preserve `Visible=False`; restore any temporary settings even on exception. | Hidden session lifecycle and unchanged user desktop/session state. |
| `task_014` / `task_015` — Snapshots | Keep the existing per-test disposable workbook-copy model. Do not implement snapshots by launching a VM or another Windows session. Snapshot operations must not activate Excel. | Separate cases cannot contaminate one another; all operations remain non-intrusive. |
| `task_016`–`task_019` — VBA inventory and source mutation | Read, export, search, and modify VBComponents/CodeModule through COM without opening or activating the VBE UI. Maintain the existing snapshot and approval gates for mutations. | Module operations work while VBE remains hidden; safety and source round-trip checks still pass. |
| `task_020` / `task_021` — Compile service and verification | Keep `compile_service.py` and the existing VBE/UIA adapters, but revise their contract: VBE command invocation and diagnostic collection must not display or activate UI. UIA is observation-only. Remove mandatory screenshot/visible-dialog interaction. A dialog that cannot be read non-intrusively produces `COMPILE_INDETERMINATE`; it must not be clicked or dismissed. | Both success and error compile fixtures meet the no-window/no-focus-steal criteria, or the feature is explicitly reported as blocked on the tested Office build. |
| `task_022` / `task_023` — Macro execution | Require the existing allowlist entry to identify the exact workbook hash/procedure as `background_safe`. Extend preflight and runtime handling for interactive UI. Block macros with known UI-dependent behavior or without verified background-safe classification. Never dismiss a popup or report a waiting macro as passed. | Approved background-safe macro completes with no visible UI; unknown/unverified or UI-dependent macro returns the documented blocked/indeterminate status. |
| `task_024`–`task_033` — Workbook objects and model aggregation | All Sheets, Tables, Names, Shapes/Controls, and model reads/mutations must use direct object references. Do not use `.Activate`, `.Select`, or visible navigation to access them. | Object operations pass with Excel hidden and no focus/visibility change. |
| `task_034` — Recovery and artifacts | Add recovery for background/UI state: detect owned modal/visible windows, stop issuing COM calls while the instance is blocked, record non-intrusive diagnostics, and clean up only the server-owned Excel instance. Do not recover by showing Excel or clicking through a dialog. | Fault-injection tests prove no false success, no focus-stealing recovery, and correct owned-process cleanup. |
| `task_035` / `task_036` — Test runner and E2E verification | “Test-case isolation” means a separate disposable workbook copy and baseline per case, executed serially through the single STA worker in the same Windows session. It does not mean VM/container/Windows-session isolation. Include a desktop non-interference assertion in the E2E workflow. | Multi-case suite passes without visible Excel/VBE windows, focus changes, or cross-case contamination. |
| `task_037` — Release verification | Add background execution to the release acceptance matrix. Document the tested Windows/Office build, any unsupported UI-dependent macros, and the exact limitations of hidden compilation/diagnostics. | Evidence confirms the normal Stage 1 workflow does not show or activate Excel/VBE windows. No VM/isolated-environment dependency exists. |

## 5. Risk register amendments

Add the following risks to the existing register; retain the existing risk IDs and content unless the description below explicitly updates them.

| Risk | Description | Required mitigation / execution rule |
|---|---|---|
| R11 — Desktop non-interference | Excel, VBE, a workbook window, or a modal dialog may become visible or steal foreground focus during COM/UI operations. | Prove hidden/non-activating behavior in `task_002`; verify throughout lifecycle in `task_011`, session tests, and `task_037`. Any focus steal or popup is a failed acceptance criterion. |
| R12 — Hidden VBE compilation | The VBE Compile command can be invoked via CommandBars, but compile errors may produce a modal message or activate the editor. | Research and PoC in `task_004`; no visible UI or VM fallback. Return `COMPILE_INDETERMINATE` if diagnostics cannot be obtained non-intrusively. |
| R13 — Interactive macro behavior | VBA code, forms, add-ins, or external programs may display UI even when Excel is hidden. | Preflight known UI calls, runtime observation, structured blocked/indeterminate status, and no automatic dialog interaction. Do not claim complete static detection. |
| R14 — Alert suppression side effects | `DisplayAlerts=False` selects default responses for some prompts and can cause unintended actions. | Apply narrowly around operations with known safe defaults; restore prior values on all paths; cover overwrite/save behavior in tests. |

Update existing risks as follows:

- **R01:** retain STA serialization and message pumping, but explicitly require a hidden automation-owned Excel instance in the current interactive session and no foreground activation.
- **R03–R04:** retain VBE command discovery and UIA diagnostics, but constrain both to non-activating invocation/observation. No visible VBE or focus-changing UI automation.
- **R05:** add detection/blocking of macro-generated UI; modal-window observation must not activate or dismiss the dialog.
- **R10:** retain interactive user identity and macro allowlist controls; explicitly prohibit VM/isolated-session/Windows-Service fallback for this project.

## 6. Execution DAG and gate amendments

The existing task dependencies remain unchanged. Add the following acceptance gates to the existing DAG:

1. `task_002` is the **Desktop Non-Interference Go/No-Go** gate. `task_010` must not be considered complete until this gate passes.
2. `task_004` is the **Hidden Compilation Go/No-Go** gate. `task_020` must not claim compile support until the clean and error fixtures have been tested under the no-window/no-focus-steal contract.
3. `task_005` is the **Background Macro Execution Go/No-Go** gate. `task_022` must distinguish normal macros from known UI-dependent behavior.
4. `task_011`, `task_013`, `task_015`, `task_021`, `task_023`, and `task_036` must include desktop non-interference regression checks relevant to their scope.
5. `task_037` must aggregate evidence from all these checks.

All Excel COM execution remains serialized through the single STA worker. Parallel code development and independent reviews remain permitted; parallel interaction with the same Excel instance is not permitted.

If a background PoC fails, the failure blocks only its dependent capability branch. Continue independent branches. Do not change the execution environment to bypass the failure.

## 7. Additional acceptance criteria for Stage 1

The following criteria supplement, rather than replace, the existing A01–A15 criteria:

- **BG01 — Same-session execution:** MCP and the automation-owned Excel instance run under the current logged-in user in the current interactive Windows session.
- **BG02 — No environment isolation:** no VM, container, separate Windows user/session, RDP desktop, alternate desktop, or Windows Service is required or used.
- **BG03 — No focus theft:** no Excel/VBE-owned window becomes the foreground window during normal operations.
- **BG04 — No visible popup:** Excel/VBE/workbook/modal windows created by the automation do not appear over the user's active applications.
- **BG05 — User-session preservation:** user-owned Excel instances and workbooks remain untouched.
- **BG06 — UI-dependent macro safety:** only procedures explicitly classified `background_safe` for the current workbook hash are eligible for strict background execution; macros that require UI or lack that classification return a truthful blocked/indeterminate result.
- **BG07 — Hidden compile evidence:** successful and failed compile cases are distinguished without showing or activating VBE/dialog UI; otherwise compilation remains explicitly unsupported/indeterminate on that tested build and is not reported as passed.
- **BG08 — Controlled settings:** all temporary Excel Application settings are restored after success, failure, timeout, and cancellation.
- **BG09 — No blind dialog handling:** no automatic keyboard/mouse input is used to dismiss unexpected dialogs.

## 8. Required work products / files to update during implementation

This addendum does not prescribe new architecture or require a new service. Prefer extending the existing modules and tests already named in `plan.md`.

Review/update these existing semantic targets:

- `src/excel_vba_mcp/adapters/excel/com_worker.py` — hidden instance startup, STA ownership, lifecycle, message handling, cleanup.
- `src/excel_vba_mcp/adapters/excel/app_adapter.py` — centralized non-activating Excel settings and safe restoration.
- `src/excel_vba_mcp/application/session_service.py` — hidden workbook open/close/save and session-state reporting.
- `src/excel_vba_mcp/application/compile_service.py` — compile invocation, truthful status, no-UI diagnostics.
- `src/excel_vba_mcp/adapters/windows_ui/vbe_compile_ui.py` and `dialog_reader.py` — read-only/non-activating observation only; no focus/click/type behavior.
- `src/excel_vba_mcp/application/macro_service.py` — background UI preflight, execution monitoring, blocked/indeterminate statuses.
- `src/excel_vba_mcp/domain/statuses.py` and `errors.py` — structured UI-required / non-interactive-execution outcomes using existing conventions.
- `src/excel_vba_mcp/security/macro_policy.py` and `audit.py` — preserve macro approval, bind `background_safe` to workbook hash/procedure, and record background/UI-blocked outcomes.
- Existing tests under `tests/integration_excel/` — assert hidden state, no foreground activation, no visible windows, safe dialog handling, and unchanged user-owned Excel state.
- `README.md` and `docs/release-evidence.md` — document same-session deployment, background limitations, supported Excel build, and no-VM policy.

Do not create a parallel Excel runtime, an alternate GUI automation stack, or a separate environment solely to satisfy this addendum.

## 9. Technical references and limits

- Microsoft documents `Application.Visible` as the application visibility property: https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible
- Microsoft documents that `ScreenUpdating=False` stops screen updates but does not provide a general guarantee that windows/dialogs cannot appear: https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating
- Microsoft documents that `DisplayAlerts=False` suppresses certain alerts and Excel may choose a default response; this can affect operations such as overwriting files: https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts
- Microsoft documents `CommandBars.FindControl`: https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.findcontrol
- Microsoft documents `CommandBarControl.Execute`: https://learn.microsoft.com/en-us/office/vba/api/office.commandbarcontrol.execute
- Microsoft warns that Office assumes an interactive desktop and can display modal dialogs; unattended/non-interactive Office automation is not recommended/supported as a general server-side pattern: https://support.microsoft.com/en-gb/visio/considerations-for-server-side-automation-of-office

These references establish individual API behaviors and limitations; they do **not** prove that VBE compilation and compile-error diagnosis will work without UI activation on every Office build. That remains an empirical acceptance test in `task_004`.

---

**Interpretation rule:** This addendum is authoritative for the background/non-interference requirement. It corrects the affected tasks in `plan.md` but does not replace the plan, change unrelated task scopes, or introduce a VM/isolated-environment implementation path.
