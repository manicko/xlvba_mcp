---
id: xlvba-limitations-and-compatibility
domain: 09-limitations-and-compatibility
tags:
  - limitations
  - compatibility
  - excel
  - windows
  - vba
  - pivottables
related:
  - xlvba-overview
  - xlvba-installation-and-setup
  - xlvba-configuration
  - xlvba-architecture
  - xlvba-mcp-tools-reference
  - xlvba-excel-capabilities
  - xlvba-workflows-and-examples
  - xlvba-security
  - xlvba-troubleshooting
  - xlvba-development-and-maintenance
---

## Purpose

This document consolidates what `xlvba` does **not** support, where it is
validated, and which behaviors are suspected rather than proven. It complements
[Excel capabilities](06-excel-capabilities.md) (which maps the capability
surface) and [Troubleshooting](10-troubleshooting.md) (which explains how to
diagnose failures).

Two labels are used throughout:

- **Verified** — the behavior is determined by code or was observed on the
  validated build; the statement is a fact about this revision.
- **Suspected / not verified** — the path is allowed or inferred, but the
  behavior is not validated on the target build. Treat it as unproven.

## Compatibility

| Area | Supported / validated | Notes |
|------|-----------------------|-------|
| Operating system | Windows 10/11 x64; validated on Windows 10 Enterprise LTSC 2019 (build 17763) | Windows only. COM (`pywin32`) and VBIDE are available only on desktop Windows. |
| Excel | Desktop Microsoft Excel with COM/VBE; validated on 16.0.14332 x64 | The universal (Microsoft Store / UWP) version is generally not suitable. Office must be activated and launched manually at least once. |
| Python | `>=3.11,<3.15` (project floor 3.11); validated on 3.14 x64 | `.python-version` pins 3.14. `pywin32` 312 ships a `cp314` wheel. |
| Package manager | [uv](https://docs.astral.sh/uv/) | `uv sync --locked` creates `.venv/` and the `excel-vba-mcp` console script. |
| Transport | MCP over STDIO | Protocol on `stdout`; logs and tracebacks on `stderr`. No network listener. |
| MCP clients | Kilo, Claude Desktop / generic MCP clients | The client spawns the server as a local child process; tools are exposed with a client prefix (for example `xlvba_excel_open`). |
| Workbook formats | `.xlsm` validated; `.xlsb`, `.xlsa`, `.xlam`, `.xltm`, `.xltx`, `.xlsx`, `.xls` allowed by the path policy but **not** verified | Only `.xlsm` is the validated scope. Behavior with other extensions is not verified. |
| VBA project access | VBOM trust enabled manually per machine | Without it VBA tools return `ACCESS_VBOM_DISABLED`. Locked projects return `VBA_PROJECT_LOCKED`. |
| Desktop integration | Hidden, non-activating Excel instance via `DispatchEx` | No GUI, no `.Activate`/`.Select`, no `SendKeys`. |

### Workbook format details

The path policy accepts the following extensions. Only `.xlsm` is validated.

| Extension | Excel role | Validation status |
|-----------|------------|-------------------|
| `.xlsm` | Macro-enabled workbook | Validated |
| `.xlsb` | Binary workbook (macro-capable) | Allowed, not verified |
| `.xlsa` | Excel add-in | Allowed, not verified |
| `.xlam` | Macro-enabled add-in | Allowed, not verified |
| `.xltm` | Macro-enabled template | Allowed, not verified |
| `.xltx` | Template | Allowed, not verified |
| `.xlsx` | Standard workbook | Allowed, not verified |
| `.xls` | Legacy workbook | Allowed, not verified |

Add-ins and templates may not open as ordinary workbooks; treat any non-`.xlsm`
behavior as unproven until tested on a copy.

## Functional limitations

| Limitation | Detail | Verified/Suspected |
|------------|--------|--------------------|
| No workbook creation | There is no new-workbook tool. `excel_open` opens an existing file only. | Verified |
| No PivotTable/PivotCache create or delete | Only existing PivotTables and caches can be inspected, refreshed and reconfigured. | Verified |
| No charts, PivotCharts or slicers | No tool lists, creates, formats or deletes them. A chart sheet appears only as a `chart` entry in `sheets_list`. | Verified |
| No row/column sizing, freeze panes or page setup | No tool exposes row height, column width, freeze panes, print/page setup or zoom. `range_get_format` covers cells only. | Verified |
| No worksheet protection changes | `sheets_list` / `sheet_get` report a read-only `protected` flag. There is no protect/unprotect tool; mutating a protected target returns `PROTECTED_TARGET`. | Verified |
| No forced recalculation | `Application.Calculate` (and `CalculateFull` / `CalculateRow` / calculation properties) is not exposed. `range_read_values` returns whatever Excel currently holds; nothing forces a recalculation. | Verified |
| No clipboard / cross-workbook copy | `range_copy` is same-workbook and never uses the clipboard. Sheet copy is in-workbook only. | Verified |
| No external links, connections or queries | No tool lists, refreshes or edits external links, data connections or Power Query definitions. | Verified |
| No import/export, print, PDF or email | No CSV import/export, PDF export, print, email or file-conversion tool. `vba_export_project` exports VBA source only. | Verified |
| No events | No tool registers, lists or edits workbook/worksheet events. ActiveX event procedures are only inferred for reading. | Verified |
| No approval-grant tool | Approval grants and allowlist entries cannot be granted or edited over MCP. `background_safe` classification **is** available through `macro_classify_background`. | Verified |
| Compile is always `indeterminate` | On the validated build `vba_compile` returns `status="indeterminate"`, `reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"`, `diagnostics.is_compile=false`. It never returns `compiled` or `compile_error`. | Verified on 16.0.14332; suspected on other builds |
| Macros require `background_safe` classification | `macro_run` runs only a procedure classified `background_safe` for the current workbook hash and VBA source revision. An unclassified procedure returns `BACKGROUND_SAFETY_NOT_VERIFIED`; an interactive one returns `INTERACTIVE_UI_NOT_SUPPORTED`. | Verified |
| `EXCEL_VBA_MCP_LOG_LEVEL` is parsed but not applied | Logging is configured at import time with a fixed `INFO` level on `stderr`; the variable has no effect in this revision. | Verified |
| ActiveX event mapping is inferred | The probable event procedure is inferred from `<sheet_code_name>_<control>_<Event>` naming. It is never click-tested and may be absent or wrong. | Suspected / not verified |
| Session recovery is not exposed as an MCP tool | `SessionService.recover_session` exists internally, but no MCP tool triggers it. The single exception is a `macro_run` hang/timeout, which terminates the owned PID and rebuilds the instance in-process (F2b), reporting `session_recovery` and discarding open workbooks; `excel_health` surfaces `needs_recovery` for every other case, where the remedy is close/reopen or restarting the server. | Verified |
| `pywinauto` is not a dependency | `pywinauto` was removed from `pyproject.toml` and `uv.lock`; VBE compile UI diagnostics use `comtypes` `IUIAutomation` (read-only). Do not expect a `pywinauto` API. | Verified |
| Missing manual fixtures | `LockedProject.xlsm` and `TableNamesProject.xlsm` are not committed; the related tests skip with explicit reasons. | Verified |
| Non-`.xlsm` formats are allowed but not validated | The path allowlist accepts the extensions above, but only `.xlsm` is validated on the target build. | Suspected / not verified |
| `macro_run` has no snapshot precondition | Unlike content-mutating tools, `macro_run` does not require a prior `snapshot_create` (documented spec deviation). | Verified |
| R13 preflight is heuristic | The token scan is a veto, not static proof; dynamic `CallByName`/`ExecuteGlobal`/`Eval` indirection and transitive UI are not fully covered. | Verified |
| The macro dialog guard is defense-in-depth, not a dialog blocker | F2a sets `Application.Interactive=False` / `DisplayAlerts=False` around the single `Application.Run` dispatch and restores them in `finally`; it suppresses only Excel's built-in alerts, not `MsgBox`/`UserForm.Show`/`FileDialog`/add-in UI/Win32 dialogs. | Verified |
| `excel_health` can lag a dead worker when the Win32 layer is unavailable | The owned-pid liveness probe requires the Win32 layer; where it is unavailable the check degrades to state-based reporting. | Verified |

## Why these limits exist

| Reason | Examples |
|--------|----------|
| Deliberate scope boundary | PivotTable/PivotCache create-delete; charts, PivotCharts, slicers; workbook creation |
| Ownership boundary | Structural range operations over a ListObject (`LISTOBJECT_INTERSECTION`) are routed to the Tables tools |
| Platform / build limitation | Non-activating compile; `PivotFilters.Add` unsupported; `Workbook.PivotTables` not a usable inventory |
| Security posture | No approval-grant tool; `macro_run` requires `background_safe`; no GUI automation |
| Not implemented | Row/column sizing, freeze panes, page setup; external links; events; import/export/print/email |

## VBA and COM limitations

- **VBOM trust is required.** Without "Trust access to the VBA project object
  model", every VBA/VBIDE tool and `vba_compile` returns
  `ACCESS_VBOM_DISABLED`.
- **Locked projects are never bypassed.** A password-protected project returns
  `VBA_PROJECT_LOCKED`; the server never attempts to unlock it.
- **No component add/remove.** Components cannot be created or deleted through
  MCP.
- **No form designer.** UserForms cannot be designed or edited; `UserForm`
  components and Document modules are rejected by the source-replacement tools.
- **No reference changes.** References can be listed (with broken ones flagged)
  but not added or removed.
- **No project protection changes.** Protection is read-only diagnostics.
- **Single hidden, non-activating instance.** The server owns one
  `DispatchEx("Excel.Application")` instance kept `Visible=False`. It never
  attaches to a user instance, never activates a window and never sends input.
- **No `.Activate` / `.Select` / `SendKeys`.** There is no GUI automation.
- **`DisplayAlerts=False` is not a universal dialog blocker.** It does not block
  `MsgBox`, `UserForm.Show`, `FileDialog`, add-in UI or Win32 dialogs.

## Concurrency and performance

- **One Excel instance.** All COM access is serialized on a single dedicated STA
  worker thread (`ExcelComWorker`); tools never touch COM directly.
- **Per-path lock.** An `asyncio.Lock` is held for the life of an open workbook,
  so reopening the same canonical path is serialized.
- **Soft timeouts.** Most operations use a 60-second default
  (`_DEFAULT_TIMEOUT_S = 60.0`). The `macro_run` dispatch uses a 90-second
  default (`macro_timeout_s`, configurable via `EXCEL_VBA_MCP_MACRO_TIMEOUT_S`,
  F2b); the test runner uses
  `_CALL_TIMEOUT_S = 60.0` and `_MACRO_TIMEOUT_S = 120.0`. The worker's
  `orphan_timeout_s` defaults to 10 seconds.
- **No blind retry of mutations.** `READ_ONLY` calls may be retried at most 3
  times for `RPC_E_CALL_REJECTED` (0.5 / 1.0 / 2.0-second pauses). `MUTATING` and
  `EXECUTING` calls are never replayed.
- **Bounded reads.** Range reads are capped at 100,000 cells per target and
  50,000 cells per response; table row reads are additionally capped at 2,000
  cells per response. Oversized requests are rejected, not silently clipped.
- **Large-workbook limits.** Very large workbooks increase memory use and COM
  call latency; a call that exceeds its soft deadline moves the worker to
  `NEEDS_RECOVERY` and raises `EXCEL_HUNG`. Read large ranges in chunks with
  `offset` / `limit`.

## External dependencies and known issues

- **`mcp` resolves to 2.3.x.** `pyproject.toml` declares `mcp>=2.3,<3`; the lock
  resolves 2.3.0. The code uses the SDK v2 `MCPServer` API.
- **Windows-only dependencies.** `pywin32` and `comtypes` are Windows-only, which
  is the main reason the server is Windows-only.
- **Excel build-specific behaviors.** `Workbook.PivotTables` is not a usable
  inventory on the target build, `PivotFilters.Add` is unsupported, and
  non-activating compilation is unavailable. These are properties of the
  validated build, not configuration errors.
- **`macro_policy.classify_operation` is non-authoritative.** The function is
  retained for illustration, is not wired into enforcement, and its table is
  stale. Use the service code or
  [MCP tools reference](05-mcp-tools-reference.md).
- **`MODAL_UI_BLOCKED` is declared but unused.** It is part of the stable
  `ErrorCode` vocabulary but is not emitted by the current adapters.
- **No general in-band recovery.** The server exposes no recovery tool; the single
  automatic path is a `macro_run` hang/timeout, which rebuilds the owned instance
  in-process (F2b). In every other case the remedy is close/reopen or restarting
  the process.
- **R13 preflight is a heuristic.** See [Security](08-security.md) for the
  residual false-safe boundary.

## Verified vs suspected

| Category | Items |
|----------|-------|
| Verified (code or validated build) | Tool inventory; path policy; approval modes; snapshot-before-mutation; compile `indeterminate`; macro gating; timeouts; bounded reads; hidden non-activating instance; no GUI automation |
| Suspected / not verified | Non-`.xlsm` formats; ActiveX event mapping; compile behavior on builds other than 16.0.14332; behavior under non-English Office UI; behavior when corporate policy blocks VBOM or macros |

If a behavior is not listed here and not in the code, treat it as unsupported
rather than assumed. When in doubt, verify with a read-only session on a copy of
the workbook.

## Operational notes

These behaviors follow from the session model in
[Architecture](04-architecture.md):

- **Originals are never modified.** `excel_open` copies the source into `work/`
  and opens only the copy.
- **Edits live in the working copy.** `excel_close` deletes it; a later
  `excel_open` re-copies the untouched original. Use
  `snapshot_create` → `excel_close` → `snapshot_restore` to persist and return to
  an earlier state.
- **Saves are explicit.** `excel_close` does not save by default; `excel_save` or
  `excel_close(save_changes=true)` writes the working copy only.
- **Snapshots restore files, not external effects.** A snapshot never undoes
  database writes, emails, network calls, files or printing.
- **One workbook per path.** A per-path lock serializes opens; reopening the same
  canonical path is rejected while it is open.

## Pre-deployment compatibility checklist

- [ ] Windows 10/11 x64 with desktop Excel (not the UWP/Store version), activated.
- [ ] Excel build recorded (validated: 16.0.14332); re-test if it differs.
- [ ] Python 3.11–3.14 x64 and `uv` on `PATH`.
- [ ] `EXCEL_VBA_MCP_ROOTS` points at a dedicated directory.
- [ ] VBOM trust enabled if VBA tools are required.
- [ ] A read-only `excel_open` → `range_read_values` → `excel_close` smoke test
  passes.

## Related

- [Overview](01-overview.md) — scope and non-goals.
- [Installation and setup](02-installation-and-setup.md) — requirements.
- [Configuration](03-configuration.md) — environment variables.
- [Architecture](04-architecture.md) — worker and session model.
- [MCP tools reference](05-mcp-tools-reference.md) — tool-level errors.
- [Excel capabilities](06-excel-capabilities.md) — capability matrix.
- [Workflows and examples](07-workflows-and-examples.md) — practical recipes.
- [Security](08-security.md) — security and permission model.
- [Troubleshooting](10-troubleshooting.md) — diagnosing failures.
- [Development and maintenance](11-development-and-maintenance.md) — contributor guide.