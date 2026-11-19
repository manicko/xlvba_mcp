# G1 Proof-of-Concept Independent Validation — task_003 / task_004 / task_005 / task_006

**Validator:** Kilo (independent)
**Date:** 2026-11-19
**Targets:** G1 risk-reduction PoCs against `.ai/plans/plan.md` and the authoritative
addendum `.ai/plans/plan_addendum_background_execution.md` (BG01–BG09).
**Mode:** READ-ONLY on source/plan. Tests were executed (serially; Excel is a shared resource). No source, plan, fixture, or committed file was modified.

## Environment (verified)

| Item | Value |
|---|---|
| OS | Windows 10, Version 10.0.17763 (LTSC 2019), x64 |
| Excel | 16.0 (Build 14332.0) |
| VBE | 7.01 |
| Office UI language | 1033 (en-US); user UI language 2057 (en-GB) |
| Python | 3.14.7 (`.venv`), uv 0.13.6 |
| Console | `chcp` reports 65001, but **piped stdout uses the ANSI code page cp1251** (this is the crux of Finding 1) |
| Foreground window during all runs | VS Code, hwnd 1050262, pid 17520, title `✓ Sub-agent: Implement task_005 macro Run PoC - xlvba_mcp - Visual Studio Code` (the `✓` = U+2713) |
| Excel process baseline | none running before/after every run (no orphaned EXCEL.EXE) |

## Commits verified (git)

| Task | Commit(s) | Files | Status |
|---|---|---|---|
| task_003 | `a8af5a2` + BG `c927e6e` | `tests/integration_excel/test_vbide_access_poc.py`, `tests/fixtures/*`, `tests/fixtures/build_fixtures.py` | tracked, unmodified |
| task_004 | `9bbc6eb` | `tests/integration_excel/test_vbe_compile_poc.py` | tracked, unmodified |
| task_005 | `62a4ee8` | `tests/integration_excel/test_macro_run_poc.py`, `CleanProject.xlsm`, `build_fixtures.py` | tracked, unmodified |
| task_006 | `4c31ef3` + BG `92b4afd` | `tests/integration_excel/test_snapshot_poc.py` | tracked, unmodified |

`git diff --stat` over all four PoC modules and the fixtures is **empty** — the validation
target is exactly the committed G1 code. All six commits map to the described files.

> **Concurrency note (material to repo-wide checks, not to G1 scope):** the working tree was
> clean at HEAD `4cc1188` when this validation began. During the session a concurrent agent
> advanced HEAD (`4cc1188 → a5d6cca → 11f93cf → adc4de4`, all `docs(...)` commits) and is
> mid-implementation of **task_010** as **untracked** files (`src/excel_vba_mcp/adapters/excel/app_adapter.py`,
> `.../com_worker.py`, `tests/unit/test_app_adapter.py`, `tests/unit/test_com_worker.py`) plus a
> modified (uncommitted) `adapters/excel/__init__.py`. All repo-wide ruff/mypy failures below are
> located **exclusively** in that untracked task_010 work-in-progress — not in any G1 PoC module and
> not in any other committed file. The committed tree (as of `adc4de4`) is ruff/format/mypy clean.

---

## task_003 — VBIDE access and component export (R02)

### Scope
Prove, on the target build, that `Workbook.VBProject` is reachable with VBOM trust, that
component types 1/2/3/100 are distinguished, that `.frm`+`.frx` export as a pair, that the
normalized code hash is stable, that Document modules survive, and that disabled VBOM is
diagnosed as `ACCESS_VBOM_DISABLED` — all with no VBE reveal (BG03/BG04/BG05).

### Checks performed
- Read the module and the latest `.ai/tmp/vbide_access_poc_*.json` artifacts (diagnostic + export).
- Confirmed `get_vbproject` maps the simulated `0x800A17B4` `com_error` to `AccessVbomDisabledError`.
- Independently re-ran the module (serial). First with the documented command, then with `PYTHONIOENCODING=utf-8`.
- Anti-gaming: read-only Win32 usage, owned-PID-scoped cleanup, no focus/input APIs.

### Actual results (independently observed)
- **Functional PASS (2 passed)** with `PYTHONIOENCODING=utf-8` (19.9s). Observed:
  - VBProject enabled; `Protection=0`; `HasPassword=False`.
  - Components: `ThisWorkbook(100)`, `Sheet1(100)`, `modTest(1)`, `clsTest(2)`, `frmTest(3, has_designer=True)` → **types 1/2/3/100 all present and distinguished**.
  - CodeModule hashes for `modTest`(30 lines)/`clsTest`(4 lines) present; markers `AddNumbers/NoArgs/Echo` and `FullName` present.
  - References: VBA, Excel, stdole, Office, MSForms — all `is_broken=False`.
  - Export: `frmTest.frm`+`frmTest.frx` both exist; `code_hash_match=True` across two independent exports (frx bytes differ, code body stable); `modTest.bas` and `ThisWorkbook.cls` exported; `document_modules_intact=True`.
  - Orphan reaped = True; no orphaned EXCEL.EXE.
- **BG evidence valid:** foreground hwnd **1050262** (VS Code, pid 17520, `is_excel_owned=False`) held constant across before/after_dispatch/after_operations/after; `visible_windows=[]` for the owned PID at both checkpoints; `app_visible_start=False`, `app_visible_after_operations=False`; `vbe_visible={check: reachable, visible: False}`; `bg03/bg04/bg05=True`, `bg_verdict=GO`. Artifact evidence matches the implementor's committed artifacts (which were produced with a cp1251-safe foreground and are therefore valid for their run).
- **Logging defect reproduced (see Finding 1):** the *documented* command `uv run pytest tests/integration_excel/test_vbide_access_poc.py -v -s` **FAILED** on this host with `AssertionError: STA worker error: UnicodeEncodeError: 'charmap' codec can't encode character '\u2713'` — the worker's `log()`→`print()` crashed on the `✓` in the VS Code title.

### Defects
- **[HIGH] F1 — Logging robustness (also affects task_004, task_006).**
  `tests/integration_excel/test_vbide_access_poc.py:352-355` — `PocArtifacts.log()` calls `print(line, flush=True)` with **no** encoding guard. When stdout is piped (any harness / CI), Python uses cp1251 and a foreground-window title containing a non-cp1251 glyph (e.g. `✓` U+2713) raises `UnicodeEncodeError`, which propagates through `_record_foreground` into the STA worker `finally`, is caught by `except BaseException`, and fails the test via `holder["error"]` (spurious failure). task_005's `log()` (`test_macro_run_poc.py:343-353`) already has the correct `try/except` + ascii-replace fallback; task_003/004/006 do not.
  **Repro:** with the VS Code window (title contains `✓`) in the foreground, run the documented command above → fail. **Workaround (no source change):** `PYTHONIOENCODING=utf-8 uv run pytest ...` → 2 passed. **Fix:** copy task_005's guard into the three `log()` methods.

### Unmet acceptance criteria
None for the capability (VBIDE round-trip fully proven). The acceptance criterion "exported source read back with a consistent hash" and "ACCESS_VBOM_DISABLED detected without registry change" are met (simulation verified; live disabled host correctly skips).

### Final result: **PASS WITH WARNINGS** (functionality + BG evidence valid; F1 makes the documented invocation flaky on non-UTF-8 consoles).

---

## task_004 — VBE compile and UIA diagnostics (R03/R12, BG07/BG09)

### Scope
Prove whether VBE Compile can be invoked — and a compile error diagnosed — while the owned Excel stays hidden and non-activating; otherwise return truthful `indeterminate` (never false `compiled`), with no click/type/SetFocus/dialog dismissal and no visible-VBE/VM fallback.

### Checks performed
- Read the module and both compile artifacts (success + error).
- Independently re-ran (serial) with `PYTHONIOENCODING=utf-8`.
- Verified CommandBars discovery, the UIA expand/invoke route, the watcher, and the modal-safety guard in code; grepped for focus/input APIs.

### Actual results (independently observed)
- **Functional PASS (2 passed)** with `PYTHONIOENCODING=utf-8` (21.9s). Both fixtures recorded:
  - `outcome = indeterminate`; `outcome_detail.reason = ui_unavoidable_on_expand`.
  - `commandbars_evidence`: `bars=30`, `debug_controls=14`, `has_compile_entry=False`, `FindControl(578)` → `E_INVALIDARG` (`0x80020009`), `compile_caption_scan=[]`, `route_usable=False` → **no CommandBar route** (R03 correction confirmed).
  - The only compile route, `uia.expand_debug()` (`ExpandCollapsePattern.Expand()`), created a **visible** `MsoCommandBarPopup` ("Debug") and moved foreground to the **hidden VBE main window** (e.g. hwnd 525124 pid 23668 "…[clsTest (Code)]"). The `_UiWatcher` caught it (`owned_foreground=true`, `max_visible_owned_windows=4`) and the PoC **aborted before any `InvokePattern.Invoke()`** — so compile was never run on either fixture.
  - `background_compile_supported=False`, `background_compile_ui_unavoidable=True`.
  - `app_visible_final=False` — **Excel stayed hidden for the entire test**.
  - `dialog=None`, `dialog_text=[]`, `selected_token="unavailable"` — the compile-error dialog was **never reached** (aborted at expand).
  - `no_com_while_modal=True` (no Excel COM while a modal could be open); `bg05=True`; `orphan_reaped=True`.
- **No false positives:** success fixture asserts `outcome != compile_error`; error fixture asserts `outcome != compiled`; when UI-unavoidable both assert `outcome == indeterminate`. Confirmed by both code and observed artifacts.
- **No click/type/SetFocus/dialog dismissal:** grep found no `SetForegroundWindow`/`ShowWindow`/`SetFocus`/`click_input`/`type_keys`/`SendKeys`/`keybd_event`/`mouse_event` calls — only docstrings stating they are absent. The only UI actions are `Expand()`/`Invoke()`, gated behind the watcher and aborted on any visibility violation; `read_dialog`/`read_code_pane` are read-only (`CurrentName`, `TextPattern.GetText/GetSelection`). No dialog is ever dismissed by input.

### Defects
- **[HIGH] F1 — Logging robustness** (same root cause as task_003). `tests/integration_excel/test_vbe_compile_poc.py:757-760` — unguarded `print()`; the `expand_debug` log line embeds the watch dict containing the `✓` foreground title, so the documented command fails on cp1251.
- **[LOW] F2 — `bg03_no_focus_theft` derived from end-boundary samples, not the in-flight watcher.** `test_vbe_compile_poc.py:1288-1289` (and `:1611-1612`): `bg03 = not (owned_fg_after or fg_changed)` uses `foreground_before`/`foreground_after` (end state), while the watcher's `violation.owned_foreground=true` (a real focus steal to the hidden VBE during `expand_debug`) is **not** folded into bg03. Result: `bg03=True` in the artifact even though the watcher recorded an owned-foreground sample. This does **not** hide the failure — `bg04=False`, `outcome=indeterminate`, and `go_no_go.verdict=NO_GO` all correctly capture the violation, and the assertion `bg03 is False or bg04 is False` passes via bg04. Advisory clarity issue only.
- **[LOW] F3 — Compile-error fixture's hardcoded `go_no_go.rationale` is inconsistent with the run it records.** `test_vbe_compile_poc.py:1628-1635` states "Compile-error fixture produces a visible, foreground #32770 dialog … Dialog text was read read-only …", but in the actual passing run the test aborted at `expand_debug` before invoking compile, so **no #32770 dialog was produced** (`dialog=[]`, `outcome_detail.reason=ui_unavoidable_on_expand`). The structured fields are accurate; only the free-text rationale is stale. (Doc-accuracy only.)
- **[INFO] F4 — Compile-error dialog-capture path is unexercised on this build.** Because `expand_debug` is itself UI-unavoidable, `invoke_compile`/`read_dialog`/`read_code_pane`/token extraction never run in the passing artifacts. This is correct fail-closed behaviour, but the optional "selected token captured" acceptance item is therefore not demonstrated here (honestly marked `unavailable`).
- **[INFO] F5 — Benign COM first-chance noise.** faulthandler prints `Windows fatal exception: code 0x800706be/0x800706ba/0x80010108` during teardown; all are handled (tests pass). Cosmetic.

### Unmet acceptance criteria
- The capability "hidden compile" is **not achievable on this build** — but that is the correct, truthful PoC outcome, not a defect: the addendum (§3, BG07) requires `COMPILE_INDETERMINATE` / `background_compile_supported=false` when compile is UI-unavoidable. The PoC satisfies every procedural acceptance criterion (hidden Excel, non-activating attempts, read-only UIA, no dismissal, no false compiled/compile_error) and returns the honest gate result **NO_GO for background compilation on 16.0.14332**, which correctly blocks task_020. The "selected token captured" item is optional and is honestly unavailable.

### Final result: **PASS WITH WARNINGS** (gate correctly determined `NO_GO`/indeterminate; F1 robustness; F2/F3 clarity; F4 coverage note).

---

## task_005 — Application.Run behaviour (R05/R13, BG06/BG09)

### Scope
Prove `Application.Run` argument/return semantics, missing-macro vs runtime-error classification, and that known interactive-UI macros (`MsgBox`/`UserForm.Show`) are preflight-blocked and that a runtime modal is handled by owned-PID termination (no dismissal, not reported as success), with Excel hidden.

### Checks performed
- Read the module and the latest macro artifacts (args + runtime).
- Independently re-ran (serial) **without** `PYTHONIOENCODING` to test the guard.
- Verified preflight heuristic, error classification, and watchdog in code.

### Actual results (independently observed)
- **Functional PASS (2 passed)** *without* any encoding override (12.0s) — confirming task_005's `log()` guard works (the `✓` printed as `?` via ascii-replace).
  - `AddNumbers(2,3)=5.0`, `NoArgs()="no-args"`, `Echo("hi")=success(None)`, `MakeArray()=(1,2,3)` — all `MACRO_SUCCESS`, serializable.
  - `Echo(object())` → `REJECTED_UNSUPPORTED_ARGUMENT` (TypeError "Objects of type 'object' can not be converted to a COM VARIANT"), `outer_hresult=None`, duration ~0.03ms → **rejected before any COM dispatch**.
  - `modTest.Nope` → `MACRO_NOT_FOUND` (outer `0x80020009`, inner `0x800A03EC` on this build) classified via EXCEPINFO description "Cannot run the macro … may not be available".
  - Preflight: `ShowMsg` flagged `MsgBox`, `ShowForm` flagged `.Show` → both `INTERACTIVE_UI_NOT_SUPPORTED`, `executed=False`; neither dispatched.
  - `RaiseError` → watchdog detected a **visible owned modal** (`visible_titles=['Microsoft Visual Basic']`), **terminated only the owned PID**, `foreground_owned=False`, `dismissed=False`, `owned_pid_terminated=True`; classified `MACRO_RUNTIME_ERROR`, **not success**. (`app.Run` then returned `0x800706BE` only because the server was killed.)
  - Orphan reaped = True; no orphaned EXCEL.EXE.
- **BG evidence valid:** Run A gate — foreground hwnd 1050262 (VS Code) constant, `visible_windows=[]`, `app_visible=False`, `vbe_visible=false`, `bg03/04/05=True`, `GO`.
- **Cleanup scoped to owned PID:** `test_macro_run_poc.py:719-737` uses `taskkill /F /PID <pid>` only (no `/IM`). Never touches user Excel.

### Defects
- **[LOW] F1 — Logging robustness:** task_005's `log()` (`test_macro_run_poc.py:343-353`) **has** the correct guard; no defect here (this is the reference fix for F1).
- **[INFO] BG04 runtime-modal limitation (honestly recorded):** `RaiseError` raises a visible automation-owned VBA modal. The BG gate (`_assert_bg`) is computed from Run A (clean) only, and the modal is recorded as `bg04_runtime_error_limitation` (`test_macro_run_poc.py:1391-1396`) rather than asserted away. This is consistent with addendum R13 (detect + truthful status, never dismiss). Not a defect; noted for completeness — strictly, BG04 is locally violated by the runtime-error case and the artifact says so.

### Unmet acceptance criteria
None. All task_005 acceptance criteria met (serializable returns, unsupported arg rejected, missing/runtime distinguished, modal detected + session diagnosable + not success, no snapshot claim).

### Final result: **PASS** (clean; the only observation is the inherent, honestly-recorded BG04 runtime-modal limitation).

---

## task_006 — SaveCopyAs snapshot and restore (R06, spec §9)

### Scope
Prove `SaveCopyAs` on an open `.xlsm` captures a separate snapshot leaving the source byte-identical; the snapshot reopens as a separate workbook; restore returns the baseline; manifest is correct; and external side effects are explicitly not undone.

### Checks performed
- Read the module and both snapshot artifacts.
- Independently re-ran (serial) with `PYTHONIOENCODING=utf-8`.

### Actual results (independently observed)
- **Functional PASS (2 passed)** with `PYTHONIOENCODING=utf-8` (8.2s).
  - `SaveCopyAs` ok; **source unchanged** (sha `2a7c3c0b…` and mtime identical before/after); snapshot sha differs (OOXML re-serialized).
  - Snapshot reopens as a **separate workbook** (`Workbooks.Count=2`, `FullName` differs), sheet model matches, file hash stable across read-only reopen, VBA components preserved (`ThisWorkbook, Sheet1, modTest, clsTest, frmTest`).
  - Manifest: `snapshot_method="Workbook.SaveCopyAs"`, `copy_sha256` (snapshot) distinct from `source_sha256` (working file), `external_side_effects_not_undone=True`.
  - Restore: mutate A1→`MUTATED-A`, `Close(SaveChanges=False)` → on-disk sha equals baseline; reopen sha equals baseline (A1 restored); snapshot is an immutable restore baseline (hash matches, A1 not mutated).
  - `external_side_effects_not_undone=True` emitted in both tests.
  - Orphan reaped = True; no orphaned EXCEL.EXE.
- **BG evidence valid:** foreground hwnd 1050262 (VS Code) constant, `visible_windows=[]`, `app_visible=False`, `vbe_visible=false`, `bg03/04/05=True`, `GO`.

### Defects
- **[HIGH] F1 — Logging robustness** (same root cause). `tests/integration_excel/test_snapshot_poc.py:363-366` — unguarded `print()`; fails on cp1251 when the foreground title contains `✓`.

### Unmet acceptance criteria
None. All task_006 acceptance criteria met.

### Final result: **PASS WITH WARNINGS** (semantics fully correct; F1 robustness).

---

## Anti-gaming assessment

- **BG checks are read-only.** Grep across all four modules found **no** calls to `SetForegroundWindow`, `ShowWindow`, `SetFocus`, `SetCursorPos`, `keybd_event`, `mouse_event`, `click_input`, `type_keys`, `SendKeys`, `PostMessage`, `SendMessage`, `.Activate`, or `.Select`. All 8 textual matches are comments/docstrings asserting their absence. Foreground/visible/VBE reads use `GetForegroundWindow`/`EnumWindows`/`IsWindowVisible`/`GetWindowThreadProcessId` and a read-only `VBE.MainWindow.Visible` — none activate or show. (task_004's `Expand()`/`Invoke()` are the deliberate, watcher-gated compile route, not focus manipulation; they abort on any visibility violation.)
- **Assertions are meaningful.** They pin concrete invariants: component types `{1,2,3,100}`, stable normalized code hash, `.frm`+`.frx` presence, reference `IsBroken=False`, exact macro return values, pre-dispatch `TypeError`, `MACRO_NOT_FOUND` vs `MACRO_RUNTIME_ERROR`, source-hash identity, `app_visible_final=False`, `bg03/04/05` booleans, `orphan_reaped`, and `outcome != compiled/compile_error` guards.
- **No failure is hidden.** task_004's green pytest correctly coexists with `go_no_go.verdict=NO_GO` (the PoC's job is to determine feasibility; "indeterminate + unsupported" *is* the truthful result). BG04 violations and the runtime modal are recorded, not asserted away.
- **Cleanup never touches non-owned Excel.** task_003/005/006 close only the owned `wb`/`app` and poll only the resolved owned PID; task_004/005 terminate **only** the owned PID via `taskkill /F /PID` / `TerminateProcess` on the owned handle (never `/IM`, never a wildcard). Owned-PID resolution is `current − baseline` and is skipped when the delta ≠ 1.
- **No dialog is dismissed by input.** No input APIs exist; task_004 never reaches the dialog; task_005 terminates the owned process for the runtime modal with `dismissed=False`.
- **Coverage limitation (not gaming):** `baseline_excel_pids` was empty in every run (no user Excel open), so BG05 "user session preserved" is satisfied trivially here; the PID-diff logic would catch a disappearing user Excel if one existed.

---

## Repo-wide checks (required work item 5)

Run at ~22:08 local on the working tree (which, per the concurrency note, includes untracked task_010 WIP):

| Check | Result | Attribution |
|---|---|---|
| `uv run pytest` (full `tests/`) | **171 passed in 75.8s**; no orphaned EXCEL.EXE | Includes the concurrent untracked `test_com_worker.py` / `test_app_adapter.py` (task_010/011 WIP) — all green. |
| `uv run ruff check src tests` | **16 errors** (`UP035`, `F401`, `UP041`, `N802`, `UP037`, `SIM117`) | **All** in `src/.../adapters/excel/app_adapter.py`, `.../com_worker.py`, `tests/unit/test_app_adapter.py`, `tests/unit/test_com_worker.py` (untracked task_010/011 WIP). |
| `uv run ruff format --check src tests` | **5 files would be reformatted** | Same untracked/modified task_010 files (`adapters/excel/__init__.py`, `app_adapter.py`, `com_worker.py`, `test_app_adapter.py`, `test_com_worker.py`). |
| `uv run mypy src` | **6 errors** (`import-untyped` ×3, `unused-ignore` ×2, `has-type` ×1) | All in `com_worker.py` / `app_adapter.py` (untracked task_010 WIP). |
| G1 PoC files only — `ruff check` | **All checks passed** | The four PoC modules + `build_fixtures.py` are ruff-clean. |
| G1 PoC files only — `ruff format --check` | **5 files already formatted** | The four PoC modules + `build_fixtures.py` are format-clean. |
| Orphaned EXCEL.EXE after every run | **none** | Verified after each module and after the full suite. |

**Interpretation:** the repo-wide ruff/mypy red is **not a G1 defect** — it is entirely attributable to
concurrent, uncommitted task_010 work by another agent. The committed codebase (as of `adc4de4`, which
does not yet contain `com_worker.py`/`app_adapter.py` and has an empty `adapters/excel/__init__.py`)
would be ruff/mypy clean. A definitive "repo-wide green" claim must be re-established once the concurrent
task_010 work is committed and its own lint/type debt is resolved. The four G1 PoC modules themselves are
lint/format/type clean.

---

## Required Fixes (blocking F1 before G1 sign-off)

1. **[HIGH] Add the encoding guard to `PocArtifacts.log()` in three modules** so the documented
   `uv run pytest …` invocation is robust on legacy-code-page consoles (matches task_005's existing fix):
   - `tests/integration_excel/test_vbide_access_poc.py:352-355`
   - `tests/integration_excel/test_vbe_compile_poc.py:757-760`
   - `tests/integration_excel/test_snapshot_poc.py:363-366`
   Pattern (from `test_macro_run_poc.py:349-353`): wrap `print(...)` in `try/except` and fall back to
   `print(line.encode("ascii","replace").decode("ascii"))`. The full UTF-8 evidence is already persisted
   by `persist()`, so nothing is lost.

## Advisory Recommendations (non-blocking)

2. **[LOW] task_004 `bg03`** — fold the watcher's `violation.owned_foreground` into `bg03_no_focus_theft`
   (or rename/annotate it as "foreground restored at end boundary") so the artifact does not read
   `bg03=True` while the watcher captured an owned-foreground sample. (`test_vbe_compile_poc.py:1288-1289`, `:1611-1612`.)
3. **[LOW] task_004 compile-error rationale** — make `go_no_go.rationale` reflect the actual code path
   (it currently claims a #32770 dialog was read when the run aborted at `expand_debug`). (`test_vbe_compile_poc.py:1628-1635`.)
4. **[INFO] task_004 coverage** — document that the compile-error dialog-capture/token-extraction path is
   unreachable on this build (expand is UI-unavoidable), so `selected_token`/`dialog_text` are always
   empty here; this is fail-closed and acceptable for Stage 1.
5. **[INFO] Domain layer (out of G1 scope)** — `src/excel_vba_mcp/domain/errors.py:48` maps `0x8002801C`
   to `ACCESS_VBOM_DISABLED` ("best effort"), whereas the real VBOM-disabled HRESULT exercised by task_003
   is `0x800A17B4`. The PoC is unaffected (it constructs `AccessVbomDisabledError` with an explicit
   hresult), but the domain map should be reconciled with `0x800A17B4` for task_007/task_016.
6. **[INFO] Test runtime noise** — the benign COM first-chance exceptions (`0x80010108`/`0x800706ba`/`0x800706be`)
   printed by faulthandler during teardown are alarming but harmless; consider suppressing faulthandler for
   these well-understood codes or documenting them in the module docstrings.

---

## Overall G1 verdict: **PASS WITH WARNINGS**

All four G1 proof-of-concepts are fundamentally sound and achieve their gates on the declared target
(Windows 10 17763 / Excel 16.0.14332 / VBE 7.01 / UI 1033):

- **task_003 (R02 VBIDE)** — **GO**: VBIDE round-trip, component-type mapping, `.frm`/`.frx` export,
  stable normalized code hash, Document-module preservation, and VBOM-disabled diagnosis all proven
  with Excel/VBE hidden and no focus theft.
- **task_004 (R12 hidden compile)** — **NO_GO (correct, fail-closed)**: on 16.0.14332 the only compile
  route (UIA Debug-menu expansion) is itself UI-unavoidable; both fixtures return truthful `indeterminate`,
  `background_compile_supported=false`, Excel stays `Visible=False`, no false `compiled`/`compile_error`,
  no click/type/SetFocus, no dialog dismissal. This correctly blocks task_020 until a non-activating route exists.
- **task_005 (R13 macro execution)** — **GO**: argument/return semantics, pre-dispatch argument rejection,
  missing-vs-runtime error classification, interactive-UI preflight blocking, and runtime-modal handling
  (owned-PID termination, `dismissed=False`, not success) all proven with Excel hidden.
- **task_006 (R06 snapshot)** — **GO**: `SaveCopyAs` source-identity, separate-workbook reopen, restore
  baseline identity, correct manifest, and explicit "external side effects not undone" all proven.

**BG compliance:** BG03/BG04/BG05 hold in every clean (non-compile) run (foreground unchanged, no visible
owned windows, `Application.Visible=False`, user session preserved); task_004 records the unavoidable
BG03/BG04 violation honestly (bg04=false, indeterminate) rather than hiding it; task_005 records the
runtime-modal BG04 limitation honestly. No VM/container/separate-session/Windows-Service fallback is used
anywhere; cleanup is owned-PID-scoped; no dialog is ever dismissed by input.

**Conditions for unconditional PASS:**
1. Fix F1 (the logging guard) in task_003/004/006 so `uv run pytest` is reproducible on cp1251 consoles.
2. Re-run the repo-wide `ruff check` / `ruff format --check` / `mypy` once the concurrent task_010 work is
   committed and its (currently untracked, non-G1) lint/type debt is resolved; the G1 modules are already clean.

No Critical findings. One High finding (F1, robustness/reproducibility). All other findings are Low/Info.