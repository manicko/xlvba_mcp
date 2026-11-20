---
id: stage1-acceptance-matrix
domain: 03-Operations
tags:
  - release-gate
  - acceptance
  - stage-1
  - excel-vba-mcp
related:
  - release-notes
---

# Stage-1 Acceptance Matrix — Excel VBA MCP Server

## Purpose

This document is the Stage-1 release-gate record for the Excel VBA MCP Server (task_037).
It maps every spec acceptance criterion **A01–A15** (spec §12.2) and every background-execution
criterion **BG01–BG09** (addendum §7) to observable evidence, and records the known limitations
that a reviewer must accept before sign-off.

It is evidence-based: a capability is only marked `PASS` when a test, commit, or artifact in this
repository demonstrates it. Where the validated target build cannot provide a capability, the row
says so and points to the research that proves it. No capability is claimed beyond what the
evidence supports.

## Main Concepts

- **Status vocabulary:** `PASS` (criterion met with direct evidence), `PASS WITH WARNINGS` (criterion met but with a documented, non-blocking limitation), `PARTIAL` (criterion only partially satisfiable on the validated build; the truthful degraded result is reported), `NOT VERIFIED` (no evidence available).
- **Validated target:** Windows 10 Enterprise LTSC 2019 (build 17763.9020, x64), Excel 16.0.14332 x64, VBE 7.01, Office/VBE UI language 1033, Python 3.14.7, uv 0.13.6. This is the *tested* configuration; it is not the spec's *recommended* configuration (spec asks Windows 11 / Microsoft 365 Apps).
- **Repository state:** commit `1227458` (`test(e2e): verify end-to-end declarative test runner workflows (task_036)`), branch `master`.
- **Evidence hierarchy:** committed test module > recorded execution-log entry > audit report > `.ai/tmp` artifact.

## How to read the evidence columns

Test paths are relative to the repository root. "Commit" is the primary implementation commit;
verification/test commits are listed where they add material evidence. Counts are from the declared
environment (see §"Reproducible local run").

---

## A01–A15 — Spec acceptance criteria (spec §12.2)

| ID | Requirement (short) | Status | Observable evidence | Notes / limitations |
|----|--------------------|--------|---------------------|---------------------|
| **A01** | Launch: MCP server starts over STDIO; a client discovers tools; stdout carries only protocol. | PASS WITH WARNINGS | task_001 (`cfcc156`); architecture audit verified `initialize` → `tools/list` → `tools/call server_health` over a real child process and confirmed stdout discipline (`docs/03-Operations` see `.ai/audit/architecture-audit-2026-11-19.md` ARCH-011, §"MCP SDK / runtime risks"); `tests/unit/test_workbook_tools.py::test_api_package_never_imports_com`; `tests/test_smoke.py`. | MCP Inspector discovery was recorded as a manual step (not executed headless). Protocol correctness is proven via real STDIO probes; `server_health` returns structured output. |
| **A02** | Excel isolation: a separate Excel instance is created; the user's existing Excel session is not intercepted. | PASS WITH WARNINGS | task_002 (`bc881ed`, BG evidence `b110a8e`), `go_no_go_r01_task002=CONDITIONAL_GO`; task_011 (`abe4ffc`) `tests/integration_excel/test_com_worker.py` (serialization / owned-shutdown / busy-no-replay / recovery); `tests/integration_excel/test_excel_instance_poc.py` (affinity probe `RPC_E_WRONG_THREAD 0x8001010E`, single-STA ownership). | `DispatchEx` instance is separate and owned; user session untouched. The "user Excel already open" branch was not exercised because no user Excel existed during the PoC (baseline PID set empty); the PID-diff cleanup logic would catch a vanishing user Excel if one existed. |
| **A03** | Workbook model: all sheets, tables, names, VBA components and shapes/controls returned in a structured, serializable response. | PASS | task_032 (`4f08e2d`) `application/workbook_model_service.py`; `tests/unit/test_workbook_model_service.py` (20) + `tests/integration_excel/test_workbook_model_service.py`; MCP resources registered in `server.py` (`excel://workbooks/{id}/model`, `.../vba/components`, ...). | Aggregation of the underlying Stage-1 inventory services; paginated/selective detail, no COM objects leak. |
| **A04** | VBA read/export: standard, class, document modules and UserForms exported with `.frx`; source hash matches export. | PASS | task_016 (`9f3a383`) read-only VBA inventory/export; task_017 (`f7be788`) fidelity verification; `tests/integration_excel/test_vba_project_read.py`, `test_vbide_access_poc.py` (G1 **PASS WITH WARNINGS**); `.ai/audit/g1-validator-2026-11-19.md` (types 1/2/3/100 distinguished, `code_hash_match=True`, `.frm`+`.frx` pair, `document_modules_intact=True`). | Locked-project case not covered: `LockedProject.xlsm` is not present (see Limitations); related tests skip with a reason. |
| **A05** | VBA edit: a standard module is modified on a working copy; snapshot created before the change; before/after hash and log recorded. | PASS | task_018 (`e6fef99`) snapshot+revision+approval guarded mutation; task_019 (`85d086a`) mutation-safety verification; `tests/integration_excel/test_vba_project_write.py`, `test_vba_mutation.py`, `tests/unit/test_vba_mutation.py`. | Revision conflicts surface as `COM_BUSY`+context (no dedicated code); Document/UserForm source replacement is rejected by design. |
| **A06** | Compile success: on a known-compilable fixture VBE compile returns `compiled` / `already_compiled` with a correct reason. | PARTIAL | task_020 (`e078d56`) `application/compile_service.py` returns `COMPILE_INDETERMINATE` with `reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`; task_021 (`822eb6b`) `tests/integration_excel/test_compile_service.py`; research `.ai/audit/hidden-compile-research-2026-11-19.md`. | On the validated build (Excel 16.0.14332) **non-activating compile is not achievable**: the only working route (UIA Debug-menu expansion) shows a visible popup and steals foreground, so the service fails closed to `COMPILE_INDETERMINATE` and never returns a false `compiled`. Positive `compiled` cannot be confirmed in the background; see Limitations. |
| **A07** | Compile error: on a syntax/reference-error fixture returns `compile_error` or `indeterminate` with a real dialog/selection; no false `compiled`. | PASS WITH WARNINGS | task_020/task_021 as above; G1 task_004 **PASS WITH WARNINGS** (`NO_GO` for background compile, fail-closed); `.ai/audit/g1-validator-2026-11-19.md` (error fixture asserts `outcome != compiled`). | `indeterminate` is returned (criterion allows it) and no false `compiled`/`compile_error` is produced; however the compile-error dialog/token is never reached on this build, so no dialog text or selected token is captured (marked `unavailable`). Partial diagnostics (broken references, protection, inventory, VBOM status) are returned and clearly labelled "NOT a compile". |
| **A08** | Macro run: an allowed test macro executes, return value and elapsed time returned; a blocked macro is not started. | PASS | task_022 (`ab418cf`) `application/macro_service.py`; task_023 (`00d1b53`) `tests/integration_excel/test_macro_execution.py`, `tests/unit/test_macro_service.py`; G1 task_005 **PASS**. | Runtime-error macros can raise a visible owned VBA modal (BG04 limitation) — the owned PID is terminated, never dismissed, and reported `MACRO_RUNTIME_ERROR`, not success. |
| **A09** | Snapshots: a baseline is created before the test; restore opens a byte-identical copy (SHA-256) of the pre-change state. | PASS | task_014 (`7ee9cfc`) `application/snapshot_service.py` + `storage/snapshot_store.py`; task_015 (`ac0c11a`) integrity/isolation; task_035 (`16d10d4`) declarative runner; task_036 (`1227458`) E2E; `tests/integration_excel/test_snapshot_service.py`, `test_end_to_end_test_runner.py`; G1 task_006 **PASS WITH WARNINGS**. | `SaveCopyAs` (open) + file-copy (closed); manifest + SHA-256; restore closes without saving. Snapshots do **not** roll back external side effects (documented, not a defect). |
| **A10** | Controls: Shape/Form buttons and OnAction enumerated; ActiveX inventoried via OLEObjects; orphaned macro mapping flagged. | PASS WITH WARNINGS | task_030 (`727c1c9`) `adapters/excel/controls_adapter.py` + `application/controls_service.py`; `tests/unit/test_shapes_controls.py` (23) + `tests/integration_excel/test_shapes_controls.py`. | ActiveX event mapping is **probable/inferred** from `<sheet_code_name>_<control>_<Event>` naming (`macro_mapping=PROBABLE`), not confirmed; no ActiveX click-testing in Stage 1 (out of scope per spec §14). Orphaned OnAction is flagged. |
| **A11** | Sheets: create/rename/copy/visibility work; destructive delete requires approval and snapshot. | PASS | task_024 (`f4724ca`) `adapters/excel/sheets_adapter.py` + `application/sheets_service.py`; `tests/unit/test_sheets_service.py` (18) + `tests/integration_excel/test_sheets_service.py`. | Delete requires snapshot+approval+last-visible-sheet guard; `CodeName` returned read-only; `xlSheetVisible` handled. |
| **A12** | Tables: add/modify/delete rows and add/delete columns mutate the ListObject (not only adjacent cells); postconditions checked. | PASS | task_026 (`80349c4`) `adapters/excel/tables_adapter.py` + `application/tables_service.py`; `tests/unit/test_tables_service.py` (30) + `tests/integration_excel/test_tables_service.py`. | Uses `ListRows`/`ListColumns`; postcondition row/column counts and header-conflict checks. |
| **A13** | Names: workbook- and worksheet-scope names read/create/update/delete correctly; non-range names do not break model reads. | PASS | task_028 (`995e067`) `adapters/excel/names_adapter.py` + `application/names_service.py`; `tests/unit/test_names_service.py` + `tests/integration_excel/test_names_service.py`. | `RefersToRange` attempted in an isolated branch; non-range names retain `RefersTo` (`target_kind`); no `eval`. |
| **A14** | Recovery: on COM exception or modal window the server does not start a parallel Excel operation and leaves a diagnostic trail. | PASS | task_034 (`c833c72`) recovery hardening + artifact storage; `tests/unit/test_recovery.py` (9) + `tests/integration_excel/test_recovery.py`; `storage/artifact_store.py`. | Single-STA serialization; busy/modal → `running/needs_recovery` fail-fast; diagnostics captured non-intrusively. |
| **A15** | Security: files outside the allowlist are blocked; inspect mode opens with macros disabled; the server never changes Trust Center/registry. | PASS | task_008 (`e6a4dd9`) `security/path_policy.py` (`tests/unit/test_path_policy.py` 12); task_009 (`72ddf93`) `security/{macro_policy,approval,audit}.py` (`tests/unit/test_approval.py`, `test_macro_policy.py`, `test_audit.py`, `test_security_policy_static.py`); task_013 (`8d64300`) `tests/integration_excel/test_session_service.py` (inspect/execute modes, forbidden path, settings restored). | Deny-by-default approval; inspect=`ForceDisable`, execute=`ByUI`; no `winreg`/AccessVBOM/TrustedLocation writes (static test enforces); VBA protection never bypassed. Two low defense-in-depth warnings (W1/W2) recorded in `.ai/audit/task_009_validator-2026-11-19.md`. |

### A01–A15 status counts

| Status | Count | IDs |
|--------|-------|-----|
| PASS | 10 | A03, A04, A05, A08, A09, A11, A12, A13, A14, A15 |
| PASS WITH WARNINGS | 4 | A01, A02, A07, A10 |
| PARTIAL | 1 | A06 |
| NOT VERIFIED | 0 | — |
| **Total** | **15** | |

---

## BG01–BG09 — Background execution (addendum §7)

These supplement A01–A15 and are mandatory for Stage-1 acceptance. The controlling rule is that
Excel/VBA runs hidden and non-activating in the current interactive user session; VM/container/
separate-session/Windows-Service fallbacks are prohibited.

| ID | Requirement (short) | Status | Observable evidence | Notes / limitations |
|----|--------------------|--------|---------------------|---------------------|
| **BG01** | Same-session execution: MCP and the owned Excel run as the current logged-in user in the current interactive session. | PASS | `README.md` §"Execution model"; `adapters/excel/com_worker.py` (`DispatchEx`, `CoInitializeEx(APARTMENTTHREADED)`); G1 task_002 evidence (`.ai/tmp/excel_instance_poc_20261119_175217.json`). | No service/SYSTEM/alternate-session dependency. |
| **BG02** | No environment isolation: no VM, container, separate user/session, RDP/alternate desktop, or Windows Service. | PASS | `README.md` §"Execution model"; addendum adoption in execution log; no such code path exists in `src/`. | Prohibited by policy and absent by construction. |
| **BG03** | No focus theft: no Excel/VBE-owned window becomes foreground during normal operations. | PASS WITH WARNINGS | G1 task_002/003/005/006 `bg03=True` in clean runs (foreground hwnd constant, `visible_windows=[]`); task_011/013/015/021/023 regression checks. | Normal Stage-1 operations are clean. The **compile** route (UIA Debug expansion) would steal foreground, so it is never invoked in background — compile fails closed instead (see BG07). |
| **BG04** | No visible popup: automation-created Excel/VBE/workbook/modal windows do not appear over the user's active applications. | PARTIAL | G1 task_005 records the runtime-modal limitation (`visible_titles=['Microsoft Visual Basic']`, owned PID terminated, `dismissed=False`); `macro_service.py` §"Honest modal handling". | Normal operations show no popup. **Known limitation:** a macro runtime error can raise a visible VBA modal even while Excel is hidden; the server terminates only the owned PID and reports `MACRO_RUNTIME_ERROR` (it does not dismiss the dialog). |
| **BG05** | User-session preservation: user-owned Excel instances and workbooks remain untouched. | PASS WITH WARNINGS | G1 task_002/003/005/006 `bg05=True`, orphan reaped, cleanup scoped to owned PID (`taskkill /F /PID`, never `/IM`); task_011 owned-shutdown test. | Verified with no user Excel running (empty baseline). PID-diff logic would detect a disappearing user Excel if one were open. |
| **BG06** | UI-dependent macro safety: only procedures classified `background_safe` for the current workbook hash are eligible; others return truthful blocked/indeterminate. | PASS | `application/macro_service.py` `BackgroundSafetyRegistry` (bound to `(workbook_sha256, "module.procedure")`); `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`; G1 task_005 preflight blocked `MsgBox`/`UserForm.Show`; `tests/integration_excel/test_macro_execution.py`. | Hash change invalidates classification. Preflight is heuristic (risk filter), not complete static proof; dynamic/indirect UI is caught at runtime and reported blocked. |
| **BG07** | Hidden compile evidence: success and failure distinguished without showing/activating VBE/dialog UI; otherwise compilation stays explicitly unsupported/indeterminate and is not reported as passed. | PASS | task_020/task_021; `.ai/audit/hidden-compile-research-2026-11-19.md`; G1 task_004 `go_no_go.verdict=NO_GO` for background compile. | The build cannot compile non-intrusively, so the service returns `COMPILE_INDETERMINATE` (`NON_ACTIVATING_COMPILE_UNAVAILABLE`) and never claims `compiled`. This is the criterion's allowed honest outcome. |
| **BG08** | Controlled settings: all temporary Excel Application settings are restored after success, failure, timeout, and cancellation. | PASS | `adapters/excel/app_adapter.py` centralized settings + `finally` restore; task_012/013 settings-restored tests (`tests/integration_excel/test_session_service.py`); `DisplayAlerts` scoped narrowly, never global. | Settings restored on all exit paths. |
| **BG09** | No blind dialog handling: no automatic keyboard/mouse input is used to dismiss unexpected dialogs. | PASS | G1 anti-gaming grep: no `SetForegroundWindow`/`ShowWindow`/`SetFocus`/`click_input`/`type_keys`/`SendKeys`/`keybd_event`/`mouse_event`; `macro_service.py` terminates owned PID with `dismissed=false`; `adapters/windows_ui/vbe_compile_ui.py` is read-only observation. | No input APIs exist in the codebase. |

### BG01–BG09 status counts

| Status | Count | IDs |
|--------|-------|-----|
| PASS | 6 | BG01, BG02, BG06, BG07, BG08, BG09 |
| PASS WITH WARNINGS | 2 | BG03, BG05 |
| PARTIAL | 1 | BG04 |
| NOT VERIFIED | 0 | — |
| **Total** | **9** | |

---

## Known limitations

These are truthful, evidence-backed limitations. A reviewer accepts the Stage-1 gate with them
present; none is a hidden failure.

1. **VBA compilation is NOT available in the background on the validated build (Excel 16.0.14332).**
   `compile_project` / `vba_compile` returns `status=COMPILE_INDETERMINATE` with reason
   `NON_ACTIVATING_COMPILE_UNAVAILABLE`. Only **partial, non-intrusive diagnostics** are provided:
   broken references (`References.IsBroken`), project protection (`VBA_PROJECT_LOCKED`), component
   inventory, and VBOM access status (`ACCESS_VBOM_DISABLED`) — explicitly labelled **"NOT a
   compile"**. The only working compile route on this build (UIA `Debug` → `Compile VBAProject`
   expansion) creates a visible popup, steals foreground, and on error shows a non-dismissible modal;
   invoking it would violate BG03/BG04/BG07. Positive `compiled`/`already_compiled` cannot be
   produced in the background. `Application.Run` is **not** a compile proxy (compile-on-demand;
   errors in unexecuted procedures are invisible). Evidence:
   `.ai/audit/hidden-compile-research-2026-11-19.md`; `.ai/audit/g1-validator-2026-11-19.md` (task_004).

2. **Macro runtime errors show a visible VBA modal even when Excel is hidden.**
   A runtime-error macro can raise a `#32770`/VBA modal that appears over the user's desktop. The
   server does **not** dismiss it (BG09); it terminates **only the owned Excel PID** and reports
   `MACRO_RUNTIME_ERROR` (`dismissed=false`). This is the BG04 limitation recorded in G1 task_005 and
   `application/macro_service.py`. Known interactive-UI macros (`MsgBox`/`InputBox`/`UserForm.Show`/
   `FileDialog`) are preflight-blocked before dispatch.

3. **ActiveX event mapping is probable/inferred, not confirmed.**
   ActiveX controls are inventoried via `Worksheet.OLEObjects`; their event handler is *inferred*
   from `<sheet_code_name>_<control>_<Event>` naming and surfaced as `macro_mapping=PROBABLE`. It is
   never treated as a confirmed event mapping, and there is no ActiveX click-testing in Stage 1
   (spec §14 defers UI click tests). Form-control `OnAction` mappings are read directly and are
   confirmed when the target procedure exists (orphans are flagged).

4. **`LockedProject.xlsm` / `TableNamesProject.xlsm` manual fixtures are not present.**
   Only `CleanProject.xlsm` and `CompileErrorProject.xlsm` are committed (both generated by
   `tests/fixtures/build_fixtures.py`). A locked project cannot be produced through VBIDE
   (`VBProject.Protection` is read-only from COM), so `LockedProject.xlsm` is a manual binary that is
   intentionally absent. `TableNamesProject.xlsm` does not exist; the table-specific end-to-end case
   is intentionally excluded rather than fabricating a table in an unrelated fixture (see
   `tests/integration_excel/test_end_to_end_test_runner.py` header). Tests that need these fixtures
   skip with an explicit reason (`pytest.skip(f"fixture {name} is missing; build fixtures first")`).

5. **Validated runtime is Python 3.14.7, not the plan's 3.12.**
   The plan's "Python 3.12" target was **relaxed to 3.14.7** (documented deviation, 2026-11-19; the
   only installed interpreter; `pyproject.toml` allows `>=3.11,<3.15`). `.python-version` pins `3.14`;
   ruff/mypy still target `py311` as the compatibility floor. 3.11–3.13 are installable but
   **unvalidated**.

6. **Validated host is Windows 10 Enterprise LTSC 2019 (build 17763), not Windows 11.**
   The spec's recommended matrix is Windows 11 x64 / Microsoft 365 Apps; the tested host is Windows 10
   LTSC 2019 x64 with Excel 16.0.14332 (Office 2016 family) and VBE UI 1033. Other Office
   builds/locales are untested. Only `.xlsm` is validated; `.xlam`/`.xlsb`/`.xls` are **not** claimed
   (spec §2.2 requires a separate compatibility test before supporting them).

7. **`xlsm`-only scope; no Stage-2 features.**
   Ranges/Cells, PivotTables, Charts/Slicers, Connections/Power Query, Data Model, and UI click tests
   are out of Stage-1 scope. No Stage-2 feature is present in this release.

---

## Reproducible local run

All commands run from the repository root in PowerShell. Integration tests require Windows + Excel +
VBOM trust and drive a real Excel instance via COM; they are serialized and skip cleanly when Excel
is unavailable.

### 1. Environment setup

```powershell
uv sync --locked
```

Creates `.venv/` from `uv.lock` and the `excel-vba-mcp` console script. Validated with Python 3.14.7
(`.python-version` = `3.14`); project floor is 3.11.

### 2. Start the MCP server (STDIO)

```powershell
uv run excel-vba-mcp
# or:
uv run python -m excel_vba_mcp
```

Diagnostics go to stderr; stdout carries only MCP protocol messages. Connect any MCP-capable client
(Kilo, Claude Desktop, Cline) as a local (stdio) server; see `README.md` §"Use as an MCP Server".

### 3. Unit tests and static gates

```powershell
uv run pytest tests/unit
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src
```

Declared-environment result: **465 unit tests passed**, ruff clean, format clean (108 files),
mypy clean (57 source files).

### 4. Integration tests (live Excel)

```powershell
uv run pytest -m integration_excel
# or explicitly:
uv run pytest tests/integration_excel
```

Declared-environment collection: **125 integration tests** (plus 1 smoke test; 591 collected in
total). These require Excel + VBOM trust and run serially on the owned, hidden Excel instance.

### 5. Build / verify test fixtures

```powershell
uv run python tests/fixtures/build_fixtures.py
uv run python tests/fixtures/build_fixtures.py --verify   # rebuild + reopen
```

Generates `CleanProject.xlsm` and `CompileErrorProject.xlsm` via COM + VBIDE (Windows + Excel + VBOM
trust required; run serially). `LockedProject.xlsm` is a manual binary, not generated.

### 6. Trust Center / VBOM note

Access to `Workbook.VBProject` requires the user to enable **Trust Center → Macro Settings →
"Trust access to the VBA project object model"** manually. The server **only diagnoses** a disabled
setting (`ACCESS_VBOM_DISABLED`); it never enables it, and never writes the registry or Trust Center.
VBA project protection is never bypassed. Macros are not run in inspect mode (AutomationSecurity
`ForceDisable`); execute mode uses `ByUI` immediately before open and restores prior settings.

---

## Related docs

- [Release notes](release-notes.md) — supported configuration, Stage-1 feature summary, quality gates, security notes, upgrade/rollback, Definition of Done.
- Evidence: `.ai/plans/execution/execution-log.md`; `.ai/audit/g1-validator-2026-11-19.md`; `.ai/audit/hidden-compile-research-2026-11-19.md`; `.ai/audit/task_009_validator-2026-11-19.md`; `.ai/audit/architecture-audit-2026-11-19.md`.