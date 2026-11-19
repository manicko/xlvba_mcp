# G1 Proof-of-Concept Research — tasks 003–006

**Author:** researcher
**Date:** 2026-11-19
**Status:** complete (research only; no production code modified)
**Scope:** Read-only investigation + throwaway probes under `.ai/tmp/`. No files under `src/` or `tests/` were created or modified.
**Target host (probed live):** Windows 10 Enterprise LTSC 2019 build 17763.9020 x64; Excel 16.0.14332 x64 (Office 2016 family); Office/VBE UI language 1033 (en-US); Python 3.14.7; pywin32 312; pywinauto 0.6.9; comtypes 1.5.0; VBOM trust **enabled** (`HKCU\...\16.0\Excel\Security\AccessVBOM = 1`, read-only).

---

## 0. Method

All findings below marked **[observed]** were executed live against the target build with `win32com.client.DispatchEx("Excel.Application")` on a dedicated STA thread. Raw evidence is persisted under `.ai/tmp/probe_*.json`. Findings marked **[doc]** come from Microsoft Learn / Microsoft Support and are cited in §7.

Probe scripts (throwaway, `.ai/tmp/`):
`probe_vbe_compile.py`, `probe_vbe_compile_uia.py`, `probe_vbe_compile_raw.py`, `probe_compile_discovery.py`, `probe_compile_discovery2.py`, `probe_compile_invoke.py`, `probe_compile_error.py`, `probe_compile_tp.py`, `probe_compile_deep.py`, `probe_tp_selection.py`, `probe_run.py`, `probe_snapshot.py`, `probe_vbide.py`, `probe_vbide2.py`, `probe_lock_modal.py`.

---

## 1. Fixture creation strategy

### 1.1 What was verified **[observed]**

| Operation | Result |
|---|---|
| `VBComponents.Add(1)` standard module | works; `CodeModule.AddFromString` writes text |
| `VBComponents.Add(2)` class module | works |
| `VBComponents.Add(3)` UserForm | works; `Designer` becomes non-`None` |
| `UserForm.Designer.Controls.Add("Forms.CommandButton.1", "btnProbe", True)` | works; control `btnProbe` present |
| Save as `.xlsm` (`FileFormat=52`) | works; produces valid macro-enabled file |
| Reopen saved `.xlsm` | UserForm + `.frx` intact; `Designer.Controls` re-read as `btnProbe` |
| `VBComponent.Export(path)` for UserForm | writes `.frm` **and** `.frx` side-by-side |
| `VBComponents.Import("frmProbe.frm")` | restores component + designer + control |
| Compile-error fixture (`Option Explicit` + undefined identifier) | VBE raises `Compile error: Variable not defined` dialog |
| **Compile-error fixture WITHOUT `Option Explicit`** | **no compile error** — an undeclared assignment is legal |
| `VBProject.Protection` | **read-only**: `AttributeError: Property '<unknown>.Protection' can not be set.` |

### 1.2 Critical gotchas

1. **`Option Explicit` is mandatory for a reliable compile error.** Without it, `x = UndefinedIdentifier` compiles cleanly. A guaranteed error is:
   ```vb
   Option Explicit
   Public Sub TriggerCompileError()
       Dim x As Long
       x = UndefinedIdentifierForCompileError   ' -> "Variable not defined"
   End Sub
   ```
   Alternatives that also fail compile: a syntax error (`If x Then` without `End If`), a call to a non-existent `Sub`/`Function` (`BuildSummary` with no declaration → `Sub or Function not defined`), or a `ByRef` type mismatch.

2. **A UserForm `.frx` is created implicitly by saving the workbook**, not by `Export`. `Export` only writes the `.frx` that already exists. Sequence that works:
   `VBComponents.Add(3)` → `Designer.Controls.Add(...)` → `Workbook.Save()` → `Component.Export("<dir>\frmProbe.frm")` → `.frm` + `.frx` both on disk. **[observed]**

3. **A locked project cannot be created through VBIDE.** `VBProject.Protection` is read-only from COM. The only programmatic route is to set the protection *inside VBE* (Tools → `<Project> Properties` → Protection tab → "Lock project for viewing" + password) — i.e. a one-time manual/UI step. The `.xlsm` must then be committed as a binary fixture. Setting a password on **our own** fixture is not "bypassing protection" and does not violate the spec; it is the only way to produce the fixture.

4. **`.frx` is not byte-reproducible.** Re-exporting the same UserForm produced a different SHA-256 (`b5685b8b…` → `dd1b9187…`) **[observed]**. Never hash `.frx` for source identity; hash the `CodeModule` text instead.

5. **Bitness/format:** the target is Excel **x64**, and pywin32 312 x64 matches. `.xlsm` (format 52) is the fixture format. 32-bit Excel is out of scope (spec §11.1); `xlsb`/`xlam` need a separate compatibility matrix.

### 1.3 Options comparison

| Option | Reproducibility | Tooling cost | `.frx` fidelity | Locked fixture | Verdict |
|---|---|---|---|---|---|
| **(a) Script-generated at test time** (COM+VBIDE, committed generator script) | High for standard/class/UserForm; **cannot** produce locked | Low–medium (needs Excel + VBOM) | Good | **No** | Best for CleanProject / CompileErrorProject |
| **(b) Manually built binary committed** | Medium (binary, opaque diff) | Medium (one-time manual) | Exact | **Yes** | Best for LockedProject |
| **(c) Hybrid: generator script + committed locked binary** | High where possible; explicit where not | Low | Good | Yes | **Recommended** |

**Recommendation: (c) hybrid.**
- `tests/fixtures/build_fixtures.py` — a committed, idempotent generator that (re)creates `CleanProject.xlsm`, `CompileErrorProject.xlsm`, and (later) `TableNamesProject.xlsm` / `ControlsProject.xlsm` via COM+VBIDE, with a `--verify` mode that reopens and asserts component inventory.
- `tests/fixtures/LockedProject.xlsm` — committed binary produced once by the manual VBE protection step, with a `README` documenting the exact VBE clicks and a `--verify` assertion that `VBProject.Protection <> 0`.
- Rationale: (a) is fully reproducible but structurally cannot express a locked project; (b) is opaque and drifts; (c) gives reproducibility where the object model allows it and a documented, reviewable binary only where it does not. This matches the plan's risk register (R02: "VBIDE is source of truth").

> Note: do **not** generate fixtures at import time on every test run. Gate generation behind `--generate` (or a session fixture with `scope="session"`), and have tests fail loudly if a fixture is missing, so a CI run on a machine without VBOM trust does not silently rewrite fixtures.

---

## 2. VBIDE access and export (task_003)

### 2.1 Component type mapping **[observed]**

| `VBComponent.Type` | Kind | `has_designer` | Export target |
|---|---|---|---|
| `1` | Standard module | False | `<name>.bas` |
| `2` | Class module | False | `<name>.cls` |
| `3` | UserForm | True | `<name>.frm` + `<name>.frx` |
| `100` | Document module (`ThisWorkbook`, `Sheet1`, …) | False | `<name>.cls` |

Confirmed inventory on a freshly built fixture: `ThisWorkbook(100)`, `Sheet1(100)`, `modStandard(1)`, `clsThing(2)`, `frmProbe(3)`. **[observed]**

### 2.2 Export/Import caveats **[observed]**

1. **`Export` does NOT append an extension.** `component.Export(r"C:\dir\modStandard")` produced a file literally named `modStandard` (no `.bas`), and `VBComponents.Import` then failed with `File not found: …\modStandard.bas`. **Always pass the full path with the correct extension.** This is the single most likely implementation bug in task_016/task_018.
2. `Export` writes a **companion `.frx`** for UserForms, at the same directory, named `<name>.frx`.
3. `Import("<name>.frm")` requires `<name>.frx` to sit beside it; it restores the designer and all controls. **[observed]**
4. **Document modules must never be removed/recreated.** They are type `100`, bound to `ThisWorkbook` / each worksheet. Export works (`ThisWorkbook.cls`, `Sheet1.cls`) and is read-only-safe. Removing and re-importing a document module is not a supported round-trip and will corrupt the workbook's code-name bindings.
5. **`CodeModule` text uses CRLF**; `.frm` headers use LF. Do not byte-compare raw `Lines()` output against an exported file for hash equality — normalize EOL and compare the code body only. **[observed]**
6. `CodeModule.CountOfLines`, `Lines(start, count)`, `ProcStartLine`, `ProcCountLines`, `ProcOfLine` are the VBIDE procedure APIs; use them instead of regex (spec §7.2). Not exercised in this pass, but they are the documented surface and match the plan.

### 2.3 References and `IsBroken` **[observed]**

`VBProject.References` on the clean fixture returned 5 entries with full metadata:

| Name | GUID | Major.Minor | FullPath | IsBroken |
|---|---|---|---|---|
| VBA | `{000204EF-…}` | 4.2 | `…\VBA7.1\VBE7.DLL` | False |
| Excel | `{00020813-…}` | 1.9 | `…\Office16\EXCEL.EXE` | False |
| stdole | `{00020430-…}` | 2.0 | `…\System32\stdole2.tlb` | False |
| Office | `{2DF8D04C-…}` | 2.8 | `…\OFFICE16\MSO.DLL` | False |
| MSForms | `{0D452EE1-…}` | 2.0 | `…\System32\FM20.DLL` | False |

`Name`, `GUID`, `Major`, `Minor`, `FullPath`, `IsBroken` are all readable. A missing reference surfaces as `IsBroken = True`; surface these as `missing_references` separately, before compile (spec §7.2 / §8.3).

### 2.4 Locked-project detection

- `VBProject.Protection` is **readable** (`0` = unprotected, `1` = locked for viewing) and **not settable** from COM. **[observed]**
- On a locked project, `CodeModule` access raises; the adapter must catch that and return `VBA_PROJECT_LOCKED` (spec §11) without attempting any password bypass.
- Recommended detection order: read `Protection` first (cheap, no side effects); if `<> 0` → `VBA_PROJECT_LOCKED`. If `Protection` read itself fails, fall back to catching the `CodeModule` access error.

### 2.5 `ACCESS_VBOM_DISABLED` detection without touching the registry

- The registry key is `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM` (confirmed `= 1` here, read-only). **[observed]**
- When VBOM trust is off, `Workbook.VBProject` raises a COM error whose HRESULT is **`0x800A17B4`** ("Programmatic Access to Visual Basic Project is not trusted"; VBA runtime error 6068 / 1004). **[doc]**
- Detection contract:
  ```python
  try:
      proj = workbook.VBProject
  except pythoncom.com_error as exc:
      if (exc.hresult & 0xFFFFFFFF) == 0x800A17B4:
          raise AccessVbOmDisabled(...)   # ACCESS_VBOM_DISABLED
      raise
  ```
- **Never** write `AccessVBOM`, `MacroSettings`, or Trusted Locations. The server only diagnoses and returns the user-facing instruction (spec §10.1). Reading the registry for *diagnostics only* is acceptable but unnecessary — the HRESULT is authoritative and cheaper.

### 2.6 Source-hash round-trip caveat

`CodeModule.Lines(1, N)` returns CRLF text. `.bas` exports also CRLF; `.cls` exports CRLF; `.frm` headers LF. Because `.frx` is non-deterministic (§1.2.4), the only stable identity for "same source" is the **normalized code body hash** (strip `\r`, strip trailing whitespace, hash UTF-8). Document this in task_016's `source_hash` definition so task_018's before/after revision check is stable.

---

## 3. VBE compile + UIA (task_004)

### 3.1 The primary finding — the plan's CommandBars route does not work here **[observed]**

On Excel 16.0.14332, UI language 1033:

- `Application.VBE.CommandBars["Debug"].Controls` contains **14 controls and none of them is Compile**:
  `212 Design Mode`, `5415 Run Project`, `186 Run Macro`, `189 Break`, `228 Reset`, `51 Toggle Breakpoint`, `188 Step Into`, `194 Step Over`, `2559 Step Out`, `2555 Locals Window`, `2554 Immediate Window`, `2556 Watch Window`, `229 Quick Watch...`, `620 Call Stack...`.
  `Run Project` is present but `Run Macro` is the enabled one; **no Compile entry exists in the `Controls` collection**.
- `CommandBars.FindControl(578)` raises `com_error 0x80020009` with inner `-2147024809` (`E_INVALIDARG`) for every argument form tried (`(578,)`, `(578,"Compile")`, `(578,"")`, `(578,"",True)`, `(578,"",True,1)`). **Control ID 578 does not exist in this VBE CommandBars implementation.**
- Enumerating **all** 30 VBE command bars and scanning captions for `compil` finds only `&Save <file>` (caption coincidence) and `&Remove <module>...` in the Project Window context menu. No Compile.

**Consequence:** the spec's "primary = localized CommandBars caption, fallback = control ID 578" plan is **inverted on this build**. Neither CommandBars route works. This is a material deviation from plan R03 and must be recorded as an open choice before task_020.

### 3.2 The working route — UIA menu expansion + Invoke **[observed]**

The Compile command *does* exist, but only in the **UIA popup** after the Debug menu is expanded:

1. `ElementFromHandle(VBE.MainWindow.Hwnd)`.
2. Find the `MenuItem` (control type `50011`) named `Debug` — note there are **multiple** `Debug` MenuItems (main menu bar + embedded toolbars); pick the visible/enabled one.
3. `GetCurrentPattern(UIA_ExpandCollapsePatternId=10005).QueryInterface(IUIAutomationExpandCollapsePattern).Expand()`.
4. Re-enumerate the subtree; find `MenuItem` named **`Compile VBAProject`** (caption is *not* `Compile <Project>`; it is the literal `Compile VBAProject`).
5. `GetCurrentPattern(UIA_InvokePatternId=10000).QueryInterface(IUIAutomationInvokePattern).Invoke()`.

The Compile item exposes **only the Invoke pattern** — `Selection`, `Value`, `RangeValue`, `Scroll`, `ExpandCollapse`, `Grid`, `Table`, `Text`, `Toggle`, `Window`, `LegacyIAccessible` are all `supported: False`. So there is no way to read its `Enabled` state from UIA other than `CurrentIsEnabled` on the element (which *is* readable and is the key signal, see §3.3).

Multiple `Compile VBAProject` items appear after expansion (one per `Debug` menu instance). Select the one whose `CurrentIsEnabled` is `True` and whose bounding rectangle is on-screen.

### 3.3 Distinguishing the four outcomes **[observed]**

| Outcome | Observable signal | Evidence |
|---|---|---|
| `compiled` | Invoke succeeds, **no `#32770` dialog appears** within the wait window, and after the call the `Compile VBAProject` menu item's `CurrentIsEnabled` becomes **`0`** (disabled) | `clean_compile_new_windows: []`, `clean_compile_item_after.enabled: 0` |
| `already_compiled` | `Compile VBAProject` is **disabled before** invoking (`CurrentIsEnabled == 0`) and the project is unchanged since the last successful compile | `clean_compile_item_after` shows `enabled: 0` after a successful compile; re-opening the menu on an unchanged project should show the same |
| `compile_error` | A new top-level `#32770` window titled `Microsoft Visual Basic for Applications` appears; its Text element reads `Compile error:\n\nVariable not defined` | `error_dialog`, `error_dialog_texts` |
| `indeterminate` | Menu/Compile item not found, `Expand()`/`Invoke()` raises, UIA tree unreadable, or the COM call hangs | fallthrough |

**Note on `already_compiled`:** the Compile item is *disabled* after a successful compile, but it is also disabled on a fresh project that has never been compiled in some states. Treat "disabled before invoke" as `already_compiled` **only** when combined with "project unchanged since last recorded compile" (a stored `source_hash` from task_016). Otherwise return `indeterminate`. Do not claim `compiled` merely because no dialog appeared.

### 3.4 Compile-error dialog structure **[observed]**

```
Window  class="#32770"  title="Microsoft Visual Basic for Applications"
  Button  name="OK"     automationId="2"
  Button  name="Help"   automationId="9"
  Image   name=""       automationId="20"
  Text    name="Compile error:\n\nVariable not defined"  automationId="65535"
  Button  name="Close"  automationId=""
```

- pywinauto UIA reads the text element reliably: `dlg.child_window(control_type="Text").window_text()` → `"Compile error:\n\nVariable not defined"`. **[observed]**
- The error text is **already fully qualified** with the VBA message (`Variable not defined`, `Sub or Function not defined`, …). Parse `error_message` from the `Compile error:\n\n<msg>` payload.

### 3.5 TextPattern / selected token

- **On the dialog: NOT available.** `GetCurrentPattern(10014).QueryInterface(IUIAutomationTextPattern)` on the dialog's Text element raises `ValueError: NULL COM pointer access`. `ValuePattern` is also absent. **The dialog's `TextPattern.GetSelection()` does not work on this build.** Use `CurrentName` instead.
- **On the VBE code pane: available.** The code pane is control type `50030` (Pane) with `AutomationId = "Bottom VBA Editor Code Pane"`. `TextPattern.DocumentRange.GetText(-1)` returns the module text, and `GetSelection().GetText(-1)` returns the selected token. **[observed]** — after `CodePane.SetSelection(4, 5, 4, 40)` the UIA selection read back `"x = UndefinedIdentifierForCompileE"`.
- Therefore: **dialog text → `CurrentName` on the Text element; selected token → TextPattern on the `Bottom VBA Editor Code Pane`.** The VBE selects the offending token on compile error, but this must be re-verified in the task_004 PoC because my first attempt filtered on control type `50020` (Text) and missed the `50030` Pane, returning an empty selection.

### 3.6 Modal / COM interaction

- The compile-error dialog is a **modal** `#32770` window owned by `EXCEL.EXE`.
- Spec §8.3 requires: UIA reads and dismisses the dialog **before** any further COM call. The server must set a "modal pending" flag the moment a dialog is detected and refuse to dispatch any other COM request until the dialog is resolved.
- My cross-thread `app.Version` probe while the modal was open was **inconclusive** (`AttributeError: Excel.Application.Version` from a proxy touched on the wrong thread, returning in 0.0 s). This is exactly the kind of false signal that must not be trusted: the correct way to prove "no COM runs while modal is open" is to **gate at the worker level** (the single-STA queue refuses new submissions while `compile` is in flight and a modal is pending), not to rely on a cross-thread probe. Mark this as a task_010/task_020 integration obligation.

### 3.7 pywinauto UIA vs raw `UIAutomationClient`

| | pywinauto 0.6.9 | raw comtypes `IUIAutomation` |
|---|---|---|
| Dialog text read | Works (`child_window(control_type="Text").window_text()`) | Works (`CurrentName`) |
| Code-pane TextPattern | `child_window(automation_id=…)` **failed** in this version (`find_elements() got an unexpected keyword argument 'automation_id'`) | Works (`ElementFromHandle` + `FindAll` + `GetCurrentPattern(10014)`) |
| Menu expansion | Not exercised for VBE menus | Works (`ExpandCollapsePattern.Expand()`) |
| COM-mode interference | Import emits `UserWarning: Revert to STA COM threading mode` | None |

**Recommendation:** use **raw comtypes `IUIAutomation`** (via `comtypes.client.GetModule("UIAutomationCore.dll")` + `CoCreateInstance(CUIAutomation)`) for the compile flow, and use **pywinauto only as a convenience wrapper for the dialog text** if desired. Reason: pywinauto 0.6.9's UIA wrapper has a keyword-argument bug for `automation_id` on this version and mutates COM threading mode on import, which is dangerous next to the STA worker. This is a deviation from the spec's "pywinauto is the primary interface" (R04); record it as an open choice with this evidence. pywinauto remains a fallback for simple dialog reads.

### 3.8 Screenshot

`PIL` is not installed in the venv (`ModuleNotFoundError: No module named 'PIL'`). For task_004's screenshot requirement, either add `pillow` as a dev dependency (screenshot only, not runtime) or use `pywinauto`'s `capture_as_image()` (which also needs Pillow) / Windows `PrintWindow` via pywin32. Recommend `pillow` in the dev group, used only by the test harness.

### 3.9 Single biggest risk

> **There is no CommandBar/ID route to Compile on this build; the only working path is localized UIA menu expansion + Invoke, and if UIA cannot see or expand the Debug menu the server can only return `indeterminate`.**

This is the highest-severity risk in G1. It means:
- The plan's R03 "localized caption, ID 578 fallback" must be rewritten to "UIA popup expansion is primary; CommandBars is not available on the validated build".
- A future Office update or a different UI language could rename `Compile VBAProject` or restructure the menu; the server must fail closed (`indeterminate`), never report `compiled` from a missing dialog.
- `task_020` (compile service) cannot be considered proven until this UIA route is validated on the actual target, which task_004 is meant to do.

---

## 4. Application.Run (task_005)

### 4.1 Argument and return semantics **[observed]**

| Case | Result |
|---|---|
| `(s, n, b)` positional | OK → `('hello', 42, True)` |
| `pywintypes.datetime` argument | OK → `datetime` |
| **`datetime.date` argument** | **`TypeError: must be a pywintypes time object (got datetime.date)`** — must convert to `datetime`/`pywintypes` before dispatch |
| 1-D list | OK → `tuple` |
| 2-D list | OK → nested `tuple` |
| `None` (null) | OK → `Empty` |
| 30 positional args | OK (`EchoArrayArgs` returned `29`) |
| 31 positional args | `TypeError: Run() takes from 1 to 32 positional arguments but 33 were given` (pywin32-level) |
| `Variant` return: scalar/date/array | OK → `int` / `datetime` / nested `tuple` |
| `Variant` return: empty | `None` |
| **Return type `Object` (Scripting.Dictionary)** | **Blocks ~51–175 s then `0x80020009`** — must be rejected as non-serializable |
| **Unsupported argument (`object()`)** | `TypeError: Objects of type 'object' can not be converted to a COM VARIANT` |

**Argument allowlist for the service:** `str`, `int`, `float`, `bool`, `None`, `datetime`/`pywintypes` time, 1-D `list` of the above, 2-D `list` of lists of the above. Reject `datetime.date` (convert to `datetime` at the boundary), reject arbitrary Python objects, reject COM references. Cap at 30 args.

### 4.2 Error classification

| Condition | HRESULT | Distinguishing evidence |
|---|---|---|
| Missing macro / macros disabled | `0x80020009` | description: `Cannot run the macro '<wb>'!<name>. The macro may not be available in this workbook or all macros may be disabled.` |
| Wrong argument count | `0x80020009` | inner HRESULT `-2147352562` (`DISP_E_BADPARAMCOUNT`) |
| Unhandled runtime error | `0x80020009` | inner HRESULT `-2146788248` (`0x800A03E8`), **no VBA description** |
| Object return | `0x80020009` | long delay (51–175 s), inner `-2146788248` |

- **The generic outer `0x80020009` cannot distinguish these alone.** Use the inner HRESULT (`exc.args[0][5]` in the pywintypes tuple) plus the message text.
- **Never claim the original VBA runtime description.** Excel returns a generic COM error; the spec explicitly forbids promising the original description (R05). Surface `error_text` as "Excel returned a generic COM error" with the HRESULT and, if available, the `Err.Description` via a companion UIA read of any modal.

### 4.3 Modal runtime error and session state

- A macro that shows `MsgBox` will block the STA worker indefinitely. The `Application.Run` call must be dispatched with a **soft timeout**; on timeout the session transitions to `running/needs_recovery` (spec §9.4), no parallel Excel operation is started, and a diagnostic trace is recorded.
- My `ModalProbe` test was **inconclusive**: the proxy was touched on a secondary thread (`AttributeError: Excel.Application.Run`), so it did not prove or disprove blocking. The correct proof is: dispatch `Application.Run` on the STA worker, wait on a `threading.Event` with timeout, and if it does not fire, flag `needs_recovery` and have UIA dismiss the modal (`{ENTER}`) as a *diagnostic* action before any further COM. Mark this as a task_005/task_022 obligation.
- **External side effects are not rolled back by snapshots.** `Application.Run` can write databases, send email, touch files. This must be stated in the macro-run result and in every snapshot restore message (spec §9.4).

---

## 5. SaveCopyAs and restore (task_006)

### 5.1 `SaveCopyAs` semantics **[observed]**

| Property | Result |
|---|---|
| Snapshot of an **open** workbook with **unsaved** changes | **Captures the unsaved in-memory state** — snapshot `A1 = "unsaved-mutation"` while the source file on disk still had `"baseline"` |
| Source file after `SaveCopyAs` | **Unchanged** — identical SHA-256 and identical mtime |
| Open workbook after `SaveCopyAs` | Still open; `Workbook.Saved` remains `False` |
| Works on a never-saved new workbook | Yes |
| Overwrites an existing target file | Yes, no error |
| Elapsed | ~0.03 s |
| Snapshot VBA available | Yes — `ThisWorkbook`, `Sheet1`, `modSnap` all readable |
| File copy (`shutil.copy2`) of a **closed** workbook | Identical SHA-256 |
| Restore (close without saving + reopen snapshot) | Baseline restored, SHA-256 matches snapshot |

**Key correction to the task framing:** the spec says "Snapshot is the file". In practice, `SaveCopyAs` writes the **current in-memory state**, which may include unsaved edits. If the intent is "snapshot the last-saved baseline file", the service must either:
- (a) use `shutil.copy2` for a closed workbook, or
- (b) `SaveCopyAs` **before** any mutation, when `Workbook.Saved == True`, or
- (c) `SaveCopyAs` after an explicit `Save()` if "snapshot current state" is intended.

Document this choice explicitly in task_014. My probe proves (b) is the safe default for "baseline before mutation".

### 5.2 Restore

- Restore = close the mutated working copy **without saving** (`Close(SaveChanges=False)`), then open the chosen snapshot. `restore opens baseline copy without saving mutated copy` is satisfied. **[observed]**
- Hash the **copy** and the **source** separately; `manifest.json` should record both `source_sha256` and `copy_sha256` plus timestamp, source path, Excel version, and (if VBProject is available) component exports.
- Restore **does not** undo external side effects (DB/email/network/files). State this in the result and in the test assertion output.

---

## 6. Ordered recommendation for the G1 sequence

### 6.1 Minimal fixture set

| Fixture | How | Needed by |
|---|---|---|
| `tests/fixtures/CleanProject.xlsm` | generator script (`VBComponents.Add` ×3, `Designer.Controls.Add`, `Save` as format 52) | task_003, task_005, task_006 |
| `tests/fixtures/CompileErrorProject.xlsm` | generator script (`Option Explicit` + undefined identifier) | task_004 |
| `tests/fixtures/LockedProject.xlsm` | one-time manual VBE protection, committed binary + README | task_003 |
| `tests/fixtures/TableNamesProject.xlsm` | generator script (defer to G7) | task_028 |
| `tests/fixtures/ControlsProject.xlsm` | generator script (defer to G7) | task_030 |

### 6.2 Ordered G1 sequence

1. **`task_003` — VBIDE access + export.** Lowest risk; produces the fixtures the other PoCs need. Exit: component type mapping, `.bas`/`.cls`/`.frm`+`.frx` export, document-module preservation, `References`/`IsBroken`, `VBA_PROJECT_LOCKED`, `ACCESS_VBOM_DISABLED` (HRESULT `0x800A17B4`).
2. **`task_006` — SaveCopyAs + restore.** Depends only on `task_002`; independent of VBIDE. Do it second so `task_004`/`task_005` can snapshot before mutating. Exit: open/closed branches, hash/manifest, restore.
3. **`task_004` — VBE compile + UIA.** Depends on `task_002` + `task_003` (needs `CompileErrorProject.xlsm`). Highest risk; execute after fixtures exist. Exit: `compiled` / `already_compiled` / `compile_error` / `indeterminate`, dialog text + screenshot, selected token from the code pane.
4. **`task_005` — Application.Run.** Depends on `task_002` + `task_003` (needs `CleanProject.xlsm`). Exit: scalar/date/1-D/2-D args, `Variant` return, `MISSING` vs runtime error vs modal, `running/needs_recovery`.
5. **Gate review.** Record Go/No-Go for R02–R06 in the execution log. `task_004`'s outcome gates `task_020`; a No-Go there must not be silently substituted.

### 6.3 Risks that could force a No-Go

| Risk | Trigger | Consequence |
|---|---|---|
| **Compile command unreachable** | UIA cannot expand the Debug menu / `Compile VBAProject` not found on a future build or locale | `task_004` → `indeterminate`; `task_020` blocked; **No-Go for the compile branch** |
| **`Compile VBAProject` item always disabled** | Project-state detection unreliable | cannot distinguish `already_compiled` from `indeterminate`; downgrade to `indeterminate` |
| **Dialog TextPattern unavailable** | Confirmed on this build (`NULL COM pointer`) | selected-token capture must come from the code pane; if the pane TextPattern also fails → no `selected_text` (partial credit only) |
| **Object return blocks 51–175 s** | Any macro returning an object | worker hangs; must reject non-serializable returns before dispatch, and enforce a soft timeout |
| **Locked fixture cannot be created** | `VBProject.Protection` read-only; no UI automation of VBE Tools dialog | `VBA_PROJECT_LOCKED` path untestable; accept as "code-reviewed, not executed" or build via manual step |
| **Modal blocks COM** | MsgBox in a test macro | worker must be gated; if not, the whole session deadlocks → `needs_recovery` |
| **`SaveCopyAs` captures unsaved state** | Snapshot taken after mutation | baseline hash mismatch; must snapshot before mutation or copy the closed file |
| **`.frx` non-determinism** | Byte-hashing `.frx` | false source-revision mismatch; hash normalized code text instead |

---

## 7. References

- Microsoft Support — *Programmatic access to Office VBA project is denied* (HRESULT `0x800A17B4`, error 6068/1004): https://support.microsoft.com/en-gb/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445
- Microsoft Learn — `Workbook.SaveCopyAs`: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas
- Microsoft Learn — `Application.Run`: https://learn.microsoft.com/en-us/office/vba/api/Excel.Application.Run
- Microsoft Learn — `Workbook.VBProject`: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbproject
- Microsoft Learn — *Debug menu* (VBE Compile): https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu
- Microsoft Learn — COM apartments / STA: https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments
- Microsoft Support — *Considerations for server-side automation of Office* (interactive-session limits): https://support.microsoft.com/en-us/visio/considerations-for-server-side-automation-of-office
- pywinauto — long-operation waits: https://pywinauto.readthedocs.io/en/uia/wait_long_operations.html
- MCP Python SDK: https://py.sdk.modelcontextprotocol.io/
- UI Automation TextPattern (COM): `IUIAutomationTextPattern`, pattern ID `10014`; code pane `AutomationId = "Bottom VBA Editor Code Pane"`, control type `50030` (Pane).

---

## 8. Evidence index (`.ai/tmp/`)

| File | Content |
|---|---|
| `probe_compile_discovery2.json` | Debug menu lacks Compile; `FindControl(578)` E_INVALIDARG; popup contains `Compile VBAProject` |
| `probe_compile_deep.json` | Compile item patterns (Invoke only); clean compile → item disabled; code pane TextPattern `Bottom VBA Editor Code Pane` |
| `probe_compile_tp.json` | Dialog text via pywinauto; dialog TextPattern `NULL COM pointer`; second compile re-shows dialog |
| `probe_tp_selection.json` | `GetSelection()` on the code pane returns the selected token |
| `probe_run.json` | `Application.Run` arg/return/error matrix |
| `probe_snapshot.json` | `SaveCopyAs` captures unsaved state; source unchanged; restore hash match |
| `probe_vbide2.json` | `.bas`/`.cls`/`.frm`/`.frx` export; import round-trip; `.frx` non-determinism |
| `probe_lock_modal.json` | `Protection` read-only; `AccessVBOM=1`; object-return hang |
| `probe_compile_error_dialog.png` | (screenshot skipped — `PIL` not installed) |

---

## 9. Handoff notes to Planner / Auditor

1. **Plan deviation (R03):** the CommandBars/ID-578 route is unavailable on Excel 16.0.14332 / en-US. task_004/task_020 must adopt UIA menu expansion as primary. This is a technical correction, not a scope change; record it in the execution log and re-run Auditor/Planner on task_004 with this evidence.
2. **Plan deviation (R04):** pywinauto 0.6.9's UIA wrapper has an `automation_id` keyword bug and flips COM threading on import; raw comtypes `IUIAutomation` is safer next to the STA worker. Keep pywinauto as a fallback for dialog text only.
3. **`Export` extension bug** must be an explicit assertion in task_016/task_018 tests.
4. **`.frx` non-determinism** must be documented in task_016's `source_hash` definition.
5. **`SaveCopyAs` semantics** must be pinned in task_014 (snapshot before mutation, or closed-file copy).
6. **`datetime.date` and object returns** must be rejected at the task_005 boundary.