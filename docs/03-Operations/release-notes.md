---
id: release-notes
domain: 03-Operations
tags:
  - release-notes
  - stage-1
  - excel-vba-mcp
related:
  - stage1-acceptance-matrix
---

# Release Notes — Excel VBA MCP Server (Stage 1)

## Purpose

Stage-1 release notes for the local Excel VBA MCP Server: a Windows-only, STDIO Model Context
Protocol server that gives an AI agent controlled access to desktop Microsoft Excel for workbook
inspection, VBA read/modify/compile, macro execution, snapshots, declarative testing, and
Sheets/Tables/Names/Shapes/Controls operations.

This is the first Stage-1 release candidate. It records the supported configuration, feature set,
quality gates, known limitations, security posture, and upgrade/rollback guidance. It is consistent
with the [Stage-1 acceptance matrix](stage1-acceptance-matrix.md).

## Main Concepts

- **Local, interactive, single-user:** the MCP server and its automation-owned Excel run as the
  logged-in user in the current interactive Windows session. No VM, container, separate session/
  desktop, RDP, or Windows Service.
- **Non-intrusive background execution:** the owned Excel instance stays hidden and non-activating
  for its whole lifetime; no focus theft, no blind dialog dismissal.
- **Structured tools, not generic COM:** a fixed catalog of typed MCP tools and read-only resources;
  COM objects never cross the STA worker boundary; only DTOs are returned.
- **Fail-closed compile:** on builds where hidden compilation is not achievable, compile returns a
  truthful `COMPILE_INDETERMINATE` instead of a false success.

## Supported configuration

| Component | Validated value | Notes |
|-----------|-----------------|-------|
| OS | Windows 10 Enterprise LTSC 2019, build 17763.9020, x64 | Spec *recommends* Windows 11 x64; this is the **tested** host. |
| Excel | Microsoft Excel 16.0.14332 x64 (Office 2016 family) | Tested via COM/VBIDE/UIA. Other Office builds/locales untested. |
| VBE | Microsoft Visual Basic for Applications 7.01 | Office/VBE UI language 1033 (en-US). |
| Python | 3.14.7 x64 | Project floor 3.11 (`requires-python >=3.11,<3.15`); 3.11–3.13 installable but unvalidated. `.python-version` = 3.14. |
| MCP SDK | `mcp` 2.3.0 | Official Python SDK v2 (`MCPServer`, `add_tool`, `add_resource`, STDIO). |
| Key deps | `pywin32` 312, `pywinauto` 0.6.9, `pydantic` 2.14.4 | Pinned in `uv.lock`. |
| Tooling | `pytest` 9.2.0, `ruff` 0.17.5, `mypy` 2.4.1 | `uv` 0.13.6. |
| Package manager | `uv` (lockfile `uv.lock`) | Reproducible install. |
| Repository state | commit `1227458` on `master` | task_036 (E2E test-runner verification). |

**Format scope:** `.xlsm` only. `.xlam`, `.xlsb`, and legacy `.xls` are **not** claimed (each needs a
separate compatibility test per spec §2.2).

## Stage-1 feature summary

The server registers **56 MCP tools** and **4 read-only MCP resources**.

- **Workbook sessions** — `excel_open`, `excel_close` (explicit `save_changes`), `excel_save`,
  `excel_get_status`, `excel_list_workbooks`, `excel_health`. Dedicated `DispatchEx` Excel instance;
  inspect mode (`AutomationSecurity=ForceDisable`) vs execute mode (`ByUI` before open); working copies
  under `work/`; per-canonical-path + `workbook_id` locking.
- **Snapshots** — `snapshot_create` (`SaveCopyAs` for open, file copy for closed), `snapshot_list`,
  `snapshot_restore` (closes working copy without saving), `snapshot_diff`. Manifest + SHA-256 +
  timestamp + source path + Excel version + model; restore returns a byte-identical baseline.
- **VBA inventory & mutations** — `vba_project_info`, `vba_list_components`, `vba_list_procedures`,
  `vba_read_component`, `vba_export_project` (`.bas`/`.cls`/`.frm`+`.frx`), `vba_get_references`,
  `vba_search_code`; guarded `vba_replace_component_source`, `vba_insert_code` (snapshot + revision
  check + approval; Document/UserForm source replacement rejected).
- **Compile [indeterminate]** — `vba_compile`, `vba_get_compile_diagnostics`. On the validated build
  this fails closed to `COMPILE_INDETERMINATE` (`NON_ACTIVATING_COMPILE_UNAVAILABLE`) with partial,
  clearly-labelled non-compile diagnostics (broken references, protection, inventory, VBOM status).
  Never a false `compiled`/`compile_error`.
- **Macro execution** — `macro_list`, `macro_run` (positional args, serializable types, Variant
  return, `run_id`/duration), `macro_get_last_result`. Approval + `background_safe` gate (bound to
  workbook hash + procedure) + interactive-UI preflight; runtime modals handled by owned-PID
  termination (`MACRO_RUNTIME_ERROR`, `dismissed=false`).
- **Sheets** — `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`
  (approval + snapshot + last-visible guard), `sheet_set_visibility`. `CodeName` read-only.
- **Tables (ListObjects)** — `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`,
  `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`,
  `table_delete_column`, `table_create`, `table_resize`. Mutations via `ListRows`/`ListColumns` with
  postcondition checks.
- **Names** — `names_list`, `name_get`, `name_create`, `name_update`, `name_delete`; workbook- and
  worksheet-scope; non-range names keep `RefersTo` (no `eval`).
- **Shapes & controls** — `shapes_list`, `shape_get`, `shape_set_on_action`, `controls_list`,
  `control_get`, `control_set_on_action`. Form-control `OnAction` confirmed when the procedure
  exists; ActiveX event mapping is `PROBABLE`/inferred; orphans flagged.
- **Workbook model + resources** — `get_workbook_model` aggregates sheets/tables/names/VBA/controls
  into a serializable, paginated DTO. Read-only resources: `excel://workbooks/{id}/model`,
  `.../vba/components`, `.../vba/components/{name}`, `excel://test-runs/{run_id}/summary`.
- **Declarative test runner** — `test_run_suite`, `test_get_run`. Bounded schema (allowlisted
  actions/assertions only; no arbitrary Python/shell/COM/VBA). Per-case disposable copy, compile-gated
  macro steps, step/case statuses, persisted run manifest + artifacts.
- **Recovery & artifacts** — single-STA serialization; busy/modal → `running/needs_recovery`
  fail-fast (no parallel Excel op); non-intrusive diagnostics; `storage/artifact_store.py` +
  `storage/run_repository.py`.

## Quality gates passed

Declared environment (Windows 10 LTSC 2019 / Excel 16.0.14332 / Python 3.14.7, commit `1227458`):

| Gate | Command | Result |
|------|---------|--------|
| Unit tests | `uv run pytest tests/unit` | **465 passed** |
| Integration tests (collected) | `uv run pytest -m integration_excel` | **125 collected** (require live Excel; 591 total incl. 1 smoke) |
| Lint | `uv run ruff check src tests` | **All checks passed!** |
| Format | `uv run ruff format --check src tests` | **108 files already formatted** |
| Types | `uv run mypy src` | **Success: no issues found in 57 source files** |

G1 proof-of-concept gates (task_002–006) are recorded **PASS WITH WARNINGS** overall
(`.ai/audit/g1-validator-2026-11-19.md`); the hidden-compile gate (task_004) is a truthful
`NO_GO` for background compile on this build. task_009 security validation is **PASS WITH WARNINGS**
(two low defense-in-depth items W1/W2; `.ai/audit/task_009_validator-2026-11-19.md`).

## Known limitations

1. **No background VBA compilation on the validated build (Excel 16.0.14332).** `compile_project`
   returns `COMPILE_INDETERMINATE` (`NON_ACTIVATING_COMPILE_UNAVAILABLE`); only partial non-intrusive
   diagnostics are provided (broken references, protection, inventory, VBOM status). Positive
   `compiled`/`already_compiled` cannot be produced in the background. See
   `.ai/audit/hidden-compile-research-2026-11-19.md` and the [acceptance matrix](stage1-acceptance-matrix.md).
2. **Macro runtime errors show a visible VBA modal even when Excel is hidden.** The server terminates
   only the owned PID and reports `MACRO_RUNTIME_ERROR` (`dismissed=false`); it never dismisses the
   dialog (BG04/BG09). Known interactive-UI macros are preflight-blocked.
3. **ActiveX event mapping is probable/inferred**, from `<sheet_code_name>_<control>_<Event>` naming;
   not confirmed and not click-tested (deferred per spec §14).
4. **Missing manual fixtures:** `LockedProject.xlsm` and `TableNamesProject.xlsm` are absent; related
   tests skip with explicit reasons. Only `CleanProject.xlsm` and `CompileErrorProject.xlsm` are
   committed.
5. **Runtime/host deviation:** Python 3.14.7 (plan's 3.12 relaxed) on Windows 10 LTSC 2019 build 17763
   (not Windows 11). `.xlsm`-only.

## Security notes

- **No Trust Center / registry changes.** The server never enables VBOM, never writes `winreg`, never
  touches `TrustedLocations`/`AccessVBOM`/`VBAWarnings` (enforced by a static policy test). Disabled
  VBOM is diagnosed as `ACCESS_VBOM_DISABLED`.
- **No protection bypass.** Locked VBA projects return `VBA_PROJECT_LOCKED`; no password/protection
  bypass path exists.
- **Approval deny-by-default.** Mutating/executing operations require an explicit grant bound to
  canonical path + sha256; unknown tool classification fails closed to `EXECUTING`.
- **Read-first / inspect mode.** Opening in inspect mode disables macros (`ForceDisable`); execute
  mode uses `ByUI` immediately before open and restores prior settings in `finally`.
- **No VM / container / service.** No isolated-environment fallback exists or is permitted.
- **Hidden, non-activating Excel.** The owned instance stays `Visible=False`; no `.Activate`/`.Select`,
  no `SetForegroundWindow`/`ShowWindow`/`SetFocus`/`click_input`/`type_keys`/`SendKeys`, no blind
  dialog dismissal.
- **Audit with redaction.** JSONL audit records workbook identity, tool, approval, result and HRESULT;
  secrets/args redacted; writes never go to stdout (protocol channel).
- **Secrets hygiene.** Passwords/secrets are redacted; credentials are not logged.

## Upgrade / rollback notes

- **Install/upgrade:** `uv sync --locked` from the repository root. No global install or Node.js is
  required. If `pyproject.toml` changed, run `uv lock` then `uv sync` deliberately and commit the
  updated lockfile.
- **Client connection:** point the MCP client at `uv run --directory <repo-root> excel-vba-mcp` (see
  `README.md` for Kilo/Claude Desktop/Cline examples). Only `EXCEL_VBA_MCP_LOG_LEVEL` is read today.
- **Prerequisites:** Windows + desktop Excel + **manual** VBOM trust for VBA features. Integration
  tests require Excel and run serially.
- **Rollback:** this is a local, file-based tool with no persistent server state, schema, or database.
  Roll back by checking out the previous commit and re-running `uv sync --locked`. Snapshots/work copies
  live under `snapshots/` and `work/` and are not touched by a code rollback. No migration step exists.
- **Compatibility:** validated only on the configuration above. Newer/older Excel builds, other
  locales, `.xlam`/`.xlsb`/`.xls`, and non-Windows hosts are untested and unsupported for this release.

## Definition of Done (Stage 1)

Mapping to spec §12.2 / §14 and task_037 acceptance criteria:

- [x] Every applicable **A01–A15** has pass / pass-with-warnings / partial evidence with a reason
      (see [acceptance matrix](stage1-acceptance-matrix.md)).
- [x] Every **BG01–BG09** background-execution criterion has evidence (no VM/isolated-environment
      dependency).
- [x] Full unit + integration regression runs on the declared target (unit 465 passed; integration 125
      collected; ruff/format/mypy clean).
- [x] Clean setup (`uv sync --locked`) and MCP STDIO startup are reproducible.
- [x] Original fixture files remain unchanged (mutations only on `work/` copies and disposable
      snapshots).
- [x] No Stage-2 feature is included (Ranges/Cells, PivotTables, Charts/Slicers, Connections/Power
      Query, Data Model, UI click tests all out of scope).
- [x] Limitations and unsupported configurations are explicit (compile indeterminate; runtime-modal
      BG04; ActiveX inference; missing manual fixtures; runtime/host deviation; `.xlsm`-only).
- [x] README covers environment, Trust Center/VBOM setup, run, and limitations.
- [ ] Independent final release-gate review (task_037 auditor/planner/validator) — pending; this
      document is the evidence package for that review.

## Related docs

- [Stage-1 acceptance matrix](stage1-acceptance-matrix.md) — A01–A15 / BG01–BG09 evidence and limitations.
- Evidence: `.ai/plans/execution/execution-log.md`; `.ai/audit/g1-validator-2026-11-19.md`;
  `.ai/audit/hidden-compile-research-2026-11-19.md`; `.ai/audit/task_009_validator-2026-11-19.md`;
  `.ai/audit/architecture-audit-2026-11-19.md`.