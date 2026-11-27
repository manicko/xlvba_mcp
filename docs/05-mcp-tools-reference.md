---
id: xlvba-mcp-tools-reference
domain: 05-mcp-tools-reference
tags:
  - mcp
  - tools
  - reference
  - excel
  - vba
related:
  - xlvba-overview
  - xlvba-architecture
  - xlvba-excel-capabilities
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This reference describes the **41 tools** of the `xlvba` MCP server: purpose, `action` values,
parameters, return value, side effects, typical errors and truthful annotations. The document is
based on the actual tool schemas obtained from a live `MCPServer` and on the source modules
`api/*`, `application/*` and `adapters/*`.

The surface was consolidated from an earlier **92-tool** catalogue (the intermediate "91 exposed"
count in older revisions was reconciled to 92 once `server_health` was counted) down to the final
**41 tools** in a single switchover. Non-destructive verbs per domain were merged into
`action`-dispatch tools (with `_read` companions where reads exist); destructive, high-risk and
lifecycle operations were kept standalone with unchanged names. See
[Migration (single switchover)](02-installation-and-setup.md#tool-surface-migration-single-switchover)
for the client migration note.

This document does not restate the architecture — that is covered in
[Architecture](04-architecture.md). Settings are in [Configuration](03-configuration.md),
the security model is in [Architecture](04-architecture.md), and limits are in
[Excel capabilities](06-excel-capabilities.md).

## How to read this reference

| Element | Meaning |
|---------|---------|
| `Tool` | The exact MCP name. Clients display it with a prefix (for example, Kilo uses `xlvba_workbook`). Tables show it without the prefix. |
| `Actions` | The `action` values the tool dispatches on. A standalone tool has a single implicit action equal to its name. |
| `Parameters` | `name:type`; `?` means optional; `=value` is the default. `—` means no parameters. The `action` selector plus the action-specific fields are validated in the handler (per-action required params). |
| `Returns` | The DTO type or response shape. All DTOs serialize to JSON and contain no COM references. |
| `Side effects` | What changes outside the response: files, snapshots, workbook state, artifacts, audit. |
| `Errors` | Stable `ErrorCode` values or other codes the tool may return. |
| `R/D/I` | Truthful MCP annotations: `readOnly` / `destructive` / `idempotent` (`T`/`F`), taken from the live tool annotations. |

Shared conventions:

1. **Access mode.** In the default `full` mode (`EXCEL_VBA_MCP_ACCESS_MODE=full`), `ApprovalPolicy`
   authorizes mutating and executing operations without a per-action grant. In `restricted` mode
   the deny-by-default gate applies: those operations are refused with `APPROVAL_DENIED` unless an
   in-process allowlist entry or live grant matches. There is no MCP tool to grant approval or edit
   the allowlist, so a restricted server exposes only `read-only` work over MCP. Per the B4 error
   taxonomy: an approval refusal is `APPROVAL_DENIED`; `WORKBOOK_READ_ONLY` is raised only for a
   read-only session; an unknown or no-longer-open `workbook_id` is `WORKBOOK_NOT_FOUND` (a caller
   error, not contention); a stale `expected_source_hash` is `REVISION_CONFLICT`; a missing
   shape/control addressed by name is `NOT_FOUND` (F6); and
   `COM_BUSY`/`RPC_E_CALL_REJECTED` are reserved for genuine COM contention only. `MACRO_BLOCKED`
   is only for Office macro-security / non-execute contexts. A lost owned Excel instance is reported
   as `WORKBOOK_NOT_FOUND` on the plain read/mutation paths and as `EXCEL_HUNG` (`needs_recovery`)
   from `macro_run` / `vba_compile_diagnostics`, which require recovery — except that `macro_run` recovers the
   session itself on a hang/timeout (F2b). In the per-tool tables, `COM_BUSY` on a
   mutating row denotes genuine contention, not approval denial, a bad id, or a revision conflict.
2. **`workbook` `action="open"` first.** Any tool that needs a `workbook_id` requires the workbook to
   be opened first. `workbook_id` is the identifier of the working copy issued by the server, not a
   file path.
3. **Snapshot before mutation.** Mutating tools require a successful snapshot
   (`snapshot_create`); otherwise — `SNAPSHOT_FAILED`.
4. **Bounds.** Range and table reads are bounded (`offset`/`limit`, the `truncated` flag);
   requests beyond the limit are rejected, not silently truncated.
5. **No GUI.** Tools never activate windows, never show dialogs and never send keystrokes.

## Annotations (truthful hints)

Every tool advertises the MCP `readOnlyHint` / `destructiveHint` / `idempotentHint` triple computed
from the per-action truth matrix (MCP default semantics: absent is not neutral). The hints are
**hints, not enforcement** — Kilo approval remains the `xlvba_*` name-glob (see
[Configuration](03-configuration.md) and the migration note). The distribution across the 41 tools:

| Annotation | Count | Tools |
|------------|-------|-------|
| `readOnly=true` | 13 | `workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `sheet_read`, `table_read`, `pivot_read`, `range_read`, `name_read`, `server_health`, `xlvba_environment` |
| `destructive=true` | 16 | `workbook`, `excel_save_as`, `excel_recover`, `snapshot_restore`, `test`, `vba_replace_component_source`, `vba_insert_code`, `macro_run`, `sheet_delete`, `table_delete_rows`, `table_delete_column`, `pivot_field`, `range_structural`, `range_delete`, `range_merge`, `name_delete` |
| `idempotent=true` | 21 | `workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `job_request_cancel`, `sheet_read`, `table_read`, `pivot_read`, `pivot_refresh`, `range_read`, `range_merge`, `name`, `name_read`, `shape`, `control`, `server_health`, `xlvba_environment`, `macro`, `pivot_field` |

Notable, matrix-proven cases:

- **`workbook` is `destructive=true`** because its `close` action has a destructive
  `discard_changes=true` path; `close` **refuses by default** (D1) and `open`/`save` are
  non-destructive. See [D1 close disposition](#workbook).
- **`range_structural` is NOT idempotent** (H4): re-`insert` duplicates rows/cells and `unmerge`
  discards merges.
- **Dual-hint tools are exactly `range_merge` and `pivot_field`** (`destructive=true` **and**
  `idempotent=true`), both proven by the matrix (`pivot_field` via `field-set-layout(delete=True)`).
- **`macro` is `readOnly=false, destructive=false, idempotent=true`**: `list`/`get-last-result` are
  reads and `classify-background` is idempotent bookkeeping; the tool is not read-only overall.
- **`vba_compile_diagnostics` is `readOnly=false, destructive=false, idempotent=false`**: it does
  **not** compile (always `indeterminate` on the validated build) and is non-mutating, but it is not
  advertised read-only. Its STA probe is tagged `op_class=EXECUTING` for retry/audit only and does
  not drive the approval gate.

## Abbreviations

| Abbreviation | Meaning |
|--------------|---------|
| `STA` | Single-Threaded Apartment — the COM apartment model in which Excel lives. |
| `DTO` | Data Transfer Object — a typed structure without COM references. |
| `VBOM` | VBA project object model — the VBA project object model (requires trust in Trust Center). |
| `ListObject` | Excel Table — a structured worksheet table. |
| `BG06` | The `background_safe` classification rule for background macro execution. |
| `R13` | Interactive-UI preflight (case-insensitive): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`, `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`, `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`, `DisplayAlerts`, `.Run`, `CommandBars`, `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`. |
| `VBE` | VBA Editor — the VBA development environment. |
| `UIA` | UI Automation — read-only access to the Windows element tree (no activation). |

## Group summary

| Group | Count | Tools |
|-------|-------|-------|
| Lifecycle / session | 5 | `workbook`, `workbook_read`, `excel_save_as`, `excel_health`, `excel_recover` |
| Snapshots | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` |
| Jobs | 2 | `job_get`, `job_request_cancel` |
| VBA | 5 | `vba_read`, `vba_export_project`, `vba_replace_component_source`, `vba_insert_code`, `vba_compile_diagnostics` |
| Macro / test | 3 | `macro`, `macro_run`, `test` |
| Sheets / tables / names | 10 | `sheet`, `sheet_read`, `sheet_delete`, `table`, `table_read`, `table_delete_rows`, `table_delete_column`, `name`, `name_read`, `name_delete` |
| Ranges | 5 | `range_read`, `range_write`, `range_structural`, `range_delete`, `range_merge` |
| Pivot | 3 | `pivot_read`, `pivot_refresh`, `pivot_field` |
| Objects | 2 | `shape`, `control` |
| Capability / health | 2 | `xlvba_environment`, `server_health` |
| **Total** | **41** | — |

Sum across groups: 5 + 4 + 2 + 5 + 3 + 10 + 5 + 3 + 2 + 2 = **41**.

## Lifecycle / session

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `workbook` | `open`, `close`, `save` | Opens a workbook as a hidden working copy; closes it; saves it. | `action`, `path?`, `mode?=execute`, `read_only?=false`, `background?=false`, `workbook_id?`, `save_changes?`, `discard_changes?` | `WorkbookInfo` / `SessionStatus` / `RunningJob` | File copy in `work/`; close deletes the copy; save writes the working copy; audit | `PATH_NOT_ALLOWED`, `APPROVAL_DENIED`, `EXCEL_HUNG`, `CONFIRMATION_REQUIRED` | F/T/F |
| `workbook_read` | `status`, `list`, `model` | Session/workbook state; open-workbook list; aggregated workbook model. | `action`, `workbook_id?`, `include_vba?=true` | `SessionStatus` / `list[WorkbookInfo]` / `WorkbookModel` | — | — (VBA state is in `vba_access`) | T/F/T |
| `excel_save_as` | (standalone) | Saves the working copy to a (possibly new) target path without rebinding the session. | `workbook_id`, `path`, `overwrite_existing?=false`, `overwrite_source?=false` | `SaveAsResult` | Flushes the working copy (`wb.Save`), then an atomic temp→replace copy + single `<target>.bak`; audit | `PATH_NOT_ALLOWED`, `PROTECTED_TARGET`, `STALE_TARGET`, `APPROVAL_DENIED` | F/T/F |
| `excel_health` | (standalone) | Worker/session state without forcing a restart and without COM. | — | `ExcelHealth` | — | — (always a payload) | T/F/T |
| `excel_recover` | (standalone) | Agent-directed recovery of the owned Excel process (fail-closed; originals safe). | `reason?=agent_directed`, `confirm?=false` | `RecoveryResult` | Terminates the wedged owned `EXCEL.EXE` via the held handle; discards working copies; rebuilds a fresh hidden instance; audit | `CONFIRMATION_REQUIRED`; otherwise reports `ok=false` / `recovery_possible=false` when refused | F/T/F |

### workbook

`workbook` dispatches on `action`:

- `open` — `mode="execute"` (default) opens a writable, execution-capable session; `mode="inspect"`
  and/or `read_only=true` opens a macros-disabled or read-only session. Authorized in `full` mode;
  refused with `APPROVAL_DENIED` in `restricted` mode for a writable/execute open.
- `close` — **refuse-by-default (D1).** A clean workbook closes normally. If there are unsaved
  changes, `close` **refuses by default**; discarding requires explicit `discard_changes=true`.
  Preserve an explicit save-on-close option (`save_changes=true`) where supported; **contradictory
  flags** (e.g. `discard_changes=true` together with `save_changes=true`) are rejected with a
  structured `ExcelVbaMcpError`.
- `save` — explicitly saves the working copy and returns an updated `WorkbookInfo` with
  `dirty=false`. Changes are never saved implicitly.

`workbook` is `readOnly=false, destructive=true, idempotent=false`: the `discard_changes=true` path
is destructive, so the tool carries a destructive hint even though `open`/`save` are non-destructive.

### excel_save_as

Save the working copy to a new (or the original) path — the fail-closed Save As (standalone, own
permission/overwrite policy). `excel_save_as` flushes the working copy with `wb.Save` on the STA
worker, then copies it to `path` by a pure-filesystem atomic temp→`os.replace` (never `wb.SaveAs`,
so the session stays bound to its working copy). A single `<target>.bak` preserves any previous
target across an overwrite.

- A **new** path is written directly (`outcome="ok"`).
- An **existing** target is refused with `PROTECTED_TARGET/target_exists` unless
  `overwrite_existing=true` (which writes the `.bak`).
- The **original source** as target is refused with `PROTECTED_TARGET/source_protected` unless
  `overwrite_source=true`; the source identity (SHA-256 + size) captured at open is re-verified, so
  a source replaced or modified since open is `STALE_TARGET`.
- The **open working copy** as target is always refused (`PROTECTED_TARGET/target_is_working_copy`).
- `outcome` is `"ok"` only after the final size + SHA-256 verify; an interrupted/unprovable write
  reports `"unknown"`/`"error"` (never success) with `backup_path` when a `.bak` exists.

### excel_recover

`excel_recover` is agent-directed: it is allow-by-default in **both** `full` and `restricted` mode,
audited, pid-scoped (never a wildcard) and discards **working copies only** — originals are never
modified. It is fail-closed: when no live, identity-verified handle is held it is refused
(`confirm` must be `true`) and a server restart is the documented fallback. See
[Architecture](04-architecture.md) and [Security](08-security.md).

`ExcelHealth.worker_state` / `session_state` include
`running/deadline_exceeded` (a single COM call exceeded its execution deadline; the call is **not**
cancelled and the owned pid is **not** terminated) — distinct from `running/needs_recovery`, which
is a genuine recovery state.

## Snapshots

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `snapshot_create` | (standalone) | Creates an immutable snapshot of an open or known workbook. | `workbook_id` | `SnapshotManifest` | Snapshot files + `manifest.json`; audit | `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `EXCEL_HUNG` | F/F/F |
| `snapshot_list` | (standalone) | Lists snapshots and their metadata. | — | `list[SnapshotManifest]` | — | `SNAPSHOT_FAILED` | T/F/T |
| `snapshot_restore` | (standalone) | Restores a working session from a snapshot. | `snapshot_id`, `workbook_id?`, `confirm?=false` | `WorkbookInfo` | Closes the current copy without saving; opens a new working copy; audit | `CONFIRMATION_REQUIRED`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `EXCEL_HUNG` | F/T/F |
| `snapshot_diff` | (standalone) | Compares two snapshots by metadata and hash. | `snapshot_id_a`, `snapshot_id_b` | `SnapshotDiff` | — | `SNAPSHOT_FAILED` | T/F/T |

`snapshot_create` and `vba_export_project` write files and are approval-gated (F-4): auto-allowed
in `full` mode, refused with `APPROVAL_DENIED` in `restricted` mode. `snapshot_restore` adds no
second gate — its approval point is the inherited writable open it performs on the snapshot
baseline. `confirm` is **required** (default `false`): a missing or `false` `confirm` is rejected
with `CONFIRMATION_REQUIRED` **before** the working copy is closed. The snapshot restores files
only — it does **not** undo external macro side effects.

## Jobs

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `job_get` | (standalone) | Returns one background job's status/result by `job_id` (poll; never touches COM). | `job_id` | `JobStatus` | — | `NOT_FOUND` | T/F/T |
| `job_request_cancel` | (standalone) | Records a job's cancel **intent** (intent-only; never interrupts the running COM call). | `job_id` | `JobStatus` | — | `NOT_FOUND` | F/F/T |

`job_get` and `job_request_cancel` never touch COM. `job_request_cancel` only sets `cancel_requested`
on the job so a poller can observe it: the single-STA COM call runs to completion and the job still
reaches a terminal state. See [Background jobs](#background-jobs).

## VBA

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `vba_read` | `project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics` | VBA project summary; component/procedure/reference inventory; module source; code search; cached compile diagnostics. | `action`, `workbook_id`, `component?`, `pattern?`, `regex?=false` | `VbaProjectInfo` / `list[VbaComponentInfo]` / `list[VbaProcedureInfo]` / `VbaComponentSource` / `list[VbaReferenceInfo]` / `list[VbaSearchMatch]` / `CompileDiagnostics` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | T/F/T |
| `vba_export_project` | (standalone) | Exports every component and returns a manifest. | `workbook_id`, `out_dir?` | `VbaExportManifest` | `.bas`/`.cls`/`.frm`/`.frx` files; artifact root; audit | `ACCESS_VBOM_DISABLED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED` | F/F/F |
| `vba_replace_component_source` | (standalone) | Fully replaces a Standard/Class module. | `workbook_id`, `component`, `new_source`, `expected_source_hash` | `VbaMutationResult` | Changes the VBA project; before/after export; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `REVISION_CONFLICT`, `PROTECTED_TARGET`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | F/T/F |
| `vba_insert_code` | (standalone) | Inserts code at a semantic position in a Standard/Class module. | `workbook_id`, `component`, `code`, `position`, `expected_source_hash` | `VbaMutationResult` | Changes the VBA project; audit | same; an invalid `position` — `COM_BUSY` | F/T/F |
| `vba_compile_diagnostics` | (standalone) | Fail-closed compile diagnostic; always `indeterminate` on the validated build. | `workbook_id` | `CompileResult` | — (does not launch UI) | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | F/F/F |

`vba_read` is read-only. `vba_replace_component_source` and `vba_insert_code` work only with
Standard/Class modules; Document modules and UserForms are rejected with `PROTECTED_TARGET`
(`reason=non_replaceable_component`). A stale `expected_source_hash` is rejected with
`REVISION_CONFLICT` (`reason=revision_conflict`); re-calling with the same hash fails the revision
guard, so `vba_replace_component_source` is **not** idempotent. `vba_compile_diagnostics` never
returns `compiled` or `compile_error`, and never shows or activates UI.

## Macro / test

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `macro` | `list`, `classify-background`, `get-last-result` | Lists callable macros with `background_safe`; records a procedure's classification; returns the cached last run. | `action`, `workbook_id`, `procedure?`, `background_safe?`, `expected_project_source_hash?` | `list[MacroInfo]` / `BackgroundSafetyClassification` / `MacroRunResult` | `classify-background` mutates in-memory classification state (bookkeeping); audit | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `MACRO_NOT_FOUND`, `STALE_TARGET`, `INTERACTIVE_UI_NOT_SUPPORTED`, `EXCEL_HUNG` | F/F/T |
| `macro_run` | (standalone) | Runs one approved `background_safe` macro. | `workbook_id`, `macro_name`, `args?`, `background?=false` | `MacroRunResult` or `RunningJob` | VBA execution; possible external effects; audit | `APPROVAL_DENIED`, `MACRO_BLOCKED`, `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED` | F/T/F |
| `test` | `run-suite`, `get-run` | Runs a declarative test suite (canonical entry point) or reads a persisted run. | `action`, `suite?`, `background?=false`, `run_id?` | `TestRunResult` / `RunningJob` | Baseline snapshot, a working copy per case, artifacts, `runs/`, audit | `INVALID_RANGE`, `APPROVAL_DENIED`, `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED` | F/T/F |

`macro` is not read-only overall (`classify-background` is a bookkeeping write) but is not
destructive; `classify-background` is **ungated** and is not the security boundary — `macro_run`
keeps the approval gate and the deterministic R13 veto. The classification is bound to
`(workbook_file_sha256, project_source_hash, procedure)` and is invalidated when the VBA source
changes. `test` is a mixed-profile tool (`run-suite` executes, `get-run` reads) and is
`destructive=true` because `run-suite` can dispatch `macro_run` steps with external side effects.
`test` accepts `background=true` to return a `RunningJob` promptly.

## Sheets / tables / names

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `sheet` | `add`, `rename`, `copy`, `set-visibility` | Adds, renames, copies a sheet; sets visibility. | `action`, `workbook_id`, `name?`, `new_name?`, `sheet_type?=worksheet`, `visible?` | `SheetInfo` | New/renamed/copied sheet; visibility; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/F |
| `sheet_read` | `list`, `get` | All sheets, or one sheet by name. | `action`, `workbook_id`, `name?` | `list[SheetInfo]` / `SheetInfo` | — | `COM_BUSY` (not found) | T/F/T |
| `sheet_delete` | (standalone) | Deletes a sheet (destructive). | `workbook_id`, `name` | `SheetInfo` | Sheet deletion; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |
| `table` | `create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column` | Creates/resizes a table; appends/inserts/updates rows; adds a column. | `action`, `workbook_id`, `sheet?`, `range_or_name?`, `name?`, `new_range?`, `rows?`, `at?`, `values?`, `updates?`, `header?`, `position?`, `style?` | `TableMutationResult` | Table rows/columns/range; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/F |
| `table_read` | `list`, `get`, `read-rows` | All tables, one table, or a bounded slice of rows. | `action`, `workbook_id`, `name?`, `offset?=0`, `limit?` | `list[TableInfo]` / `TableInfo` / `TableRows` | — | `COM_BUSY` | T/F/T |
| `table_delete_rows` | (standalone) | Deletes rows by indices or an unambiguous `where` (XOR guard). | `workbook_id`, `name`, `indices?`, `where?` | `TableMutationResult` | Table rows; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |
| `table_delete_column` | (standalone) | Deletes a column by its unique header (own permission). | `workbook_id`, `name`, `header` | `TableMutationResult` | Table column; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |
| `name` | `create`, `update` | Creates a name in an explicit scope; updates `RefersTo` (scope preserved). | `action`, `workbook_id`, `name`, `refers_to?`, `scope`, `sheet?` | `DefinedNameInfo` | New/updated name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/T |
| `name_read` | `list`, `get` | All names, or one name by name and explicit scope. | `action`, `workbook_id`, `name?`, `scope?`, `sheet?` | `list[DefinedNameInfo]` / `DefinedNameInfo` | — | `COM_BUSY` | T/F/T |
| `name_delete` | (standalone) | Deletes a name and returns its previous value. | `workbook_id`, `name`, `scope`, `sheet?` | `DefinedNameInfo` | Name; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |

`sheet_delete` is rejected if no visible sheet would remain after the deletion;
`sheet set-visibility` rejects hiding the last visible sheet. `table` `update-rows` accepts an
explicit index or an unambiguous `where`; `table_delete_rows` requires exactly one of `indices` /
`where` (XOR). `name` `create` rejects a duplicate (same scope) or a cross-scope conflict; `name`
`update` never re-scopes a name.

## Ranges

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `range_read` | `get`, `read-values`, `read-formulas`, `get-format`, `get-validation`, `get-conditional-formatting` | Range metadata; typed values/formulas; format; validation; conditional formatting. | `action`, `workbook_id`, `sheet`, `address`, `offset?=0`, `limit?` | `RangeAddress` / `RangeValues` / `RangeFormulas` / `FormatInfo` / `ValidationRule` / `list[ConditionalFormatRule]` | — | `INVALID_RANGE` | T/F/T |
| `range_write` | `write-values`, `write-formulas`, `set-format`, `set-validation`, `set-conditional-formatting` | Writes typed values/formulas; sets format, validation or a conditional-format rule. | `action`, `workbook_id`, `sheet`, `address`, `payload?`, `patch?`, `spec?`, `mode?=add` | `RangeMutationResult` / `FormatMutationResult` / `ValidationRule` / `ConditionalFormatMutationResult` | Cell values/formulas/format/rules; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/F/F |
| `range_structural` | `copy`, `insert`, `unmerge` | Copies a range; inserts cells (`right`/`down`); unmerges. | `action`, `workbook_id`, `sheet`, `source?`, `destination_sheet?`, `destination?`, `address?`, `shift?`, `allow_merged?=false` | `StructuralMutationResult` | Destination data / cell shift / unmerge; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/F |
| `range_delete` | (standalone) | Deletes cells with a `left`/`up` shift. | `workbook_id`, `sheet`, `address`, `shift`, `allow_merged?=false` | `StructuralMutationResult` | Cell deletion; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/F |
| `range_merge` | (standalone) | Merges a range into a single cell. | `workbook_id`, `sheet`, `address`, `across?=false`, `allow_data_loss?=false` | `StructuralMutationResult` | Merge; audit | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/T |

`range_structural` is **not idempotent** (H4): re-`insert` duplicates and `unmerge` discards merges.
`range_write` is not idempotent because `set-conditional-formatting(mode="add")` appends a rule
(renumbering priorities). Structural operations that intersect an Excel Table (ListObject) are
rejected with `LISTOBJECT_INTERSECTION` and must be performed through the table tools.

## Pivot

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `pivot_read` | `tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters` | PivotTable/cache inventory; table/cache metadata; field list; field filter state. | `action`, `workbook_id`, `pivot_table?`, `cache_index?`, `field?`, `background?=false` | `PivotInventory` / `PivotTableInfo` / `PivotCacheInfo` / `list[PivotFieldInfo]` / `PivotFieldFilters` (or `RunningJob`) | — | `COM_BUSY`, `AMBIGUOUS_FILTER` | T/F/T |
| `pivot_refresh` | `refresh-table`, `refresh-cache` | Refreshes one PivotTable or one PivotCache. | `action`, `workbook_id`, `pivot_table?`, `cache_index?` | `PivotRefreshResult` | PivotTable/cache data; audit | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/T |
| `pivot_field` | `field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters` | Sets a field's orientation/position/aggregation/properties; applies/clears filters. | `action`, `workbook_id`, `pivot_table`, `field?`, `orientation?`, `delete?=false`, `position?`, `function?`, `caption?`, `properties?`, `spec?`, `scope?=field` | `PivotFieldInfo` / `PivotFieldFilters` | PivotTable structure/data; audit | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/T |

`pivot_read` is read-only; `pivot_refresh` and `pivot_field` are not. `pivot_field` is
`destructive=true` **and** `idempotent=true` (dual-hint): `field-set-layout(delete=True)` discards a
field, while re-applying the same layout/position/aggregation is idempotent. `cache-get` /
`refresh-cache` require `cache_index`; field actions require `pivot_table` + `field`
(`field-clear-filters` also accepts table scope). Creating/deleting PivotTables/PivotCaches is not
supported. `PivotFilters.Add` (label/value/date/top-bottom) is not supported on the validated build:
it returns an explicit `unsupported` result rather than silently converting to a manual filter. A
page (report-filter) field's state is classified from `CurrentPage` / `AllItemsVisible`;
`field-apply-filter` on a page field is rejected before any COM mutation.

## Objects

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `shape` | `list`, `get`, `set-on-action` | All shapes; one shape by name; assigns a macro to a shape. | `action`, `workbook_id`, `name?`, `macro_name?` | `list[ShapeInfo]` / `ShapeInfo` | `OnAction`; audit | `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED` | F/F/T |
| `control` | `list`, `get`, `set-on-action` | All Form Controls/ActiveX; one control by name; assigns a macro to a Form Control. | `action`, `workbook_id`, `name?`, `macro_name?` | `list[ControlEntry]` / `ControlEntry` | `OnAction`; audit | `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED` | F/F/T |

`shape` and `control` are mixed tools (list/get read, set-on-action mutates) and are
`readOnly=false, destructive=false, idempotent=true` (all actions idempotent per the matrix). An
empty `macro_name` clears the assignment. ActiveX controls are rejected: they do not have Form
Control `OnAction` semantics. A missing shape/control returns `NOT_FOUND` (F6); an ambiguous
(duplicate) name is a plain `ExcelVbaMcpError` with `COM_BUSY` semantics and
`context.reason="ambiguous"`.

## Capability / health

| Tool | Actions | Purpose | Parameters | Returns | Side effects | Errors | R/D/I |
|------|---------|---------|------------|---------|--------------|--------|-------|
| `xlvba_environment` | (standalone) | Reports the server access mode, policy, build and capabilities without touching COM. | — | `EnvironmentReport` | — | — | T/F/T |
| `server_health` | (standalone) | Read-only liveness probe; never touches Excel/COM. | — | `{"status", "service", "version"}` | — | — | T/F/T |

`xlvba_environment` is registered with an **inline** `read_only` annotation (not via the shared
annotation set); `server_health` is likewise inline. Both are read-only and never touch COM.

## Complex tools in detail

### workbook `action="open"`

`workbook` `open` defaults to `mode="execute", read_only=False` — a writable, execution-capable
session. Pass `mode="inspect"` and/or `read_only=True` to opt into a macros-disabled or read-only
session:

```json
{ "action": "open", "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
```

Expected response shape (`WorkbookInfo`):

```json
{
  "workbook_id": "1f0c…",
  "name": "report.xlsm",
  "format": "xlsm",
  "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm",
  "source_path": "C:/data/report.xlsm",
  "read_only": true,
  "has_vba": true,
  "excel_version": "16.0.14332",
  "session_id": "9b2a…",
  "mode": "inspect"
}
```

`WorkbookInfo` carries **two paths** and deliberately **no** `original_name`:

- `path` is the **working copy** actually opened by the server — the materialized file under the
  server `work/` directory. Every read, write and mutation targets this copy; the original is never
  modified.
- `source_path` is the **original, allowlisted** path the working copy was materialized from, or
  `null` when the workbook has no on-disk source.
- `original_name` is intentionally absent (redundant with `name`).

### workbook `action="close"` / `action="save"`

```json
{ "action": "close", "workbook_id": "1f0c…", "save_changes": false }
```

`close` without saving is read-only and allowed; the response is a `SessionStatus` with
`state="closed"`. **Refuse-by-default (D1):** with unsaved changes, `close` refuses unless
`discard_changes=true`; `save_changes=true` saves on close (a mutation, `APPROVAL_DENIED` in
`restricted` mode); contradictory flags are rejected. Explicit save:

```json
{ "action": "save", "workbook_id": "1f0c…" }
```

Requires approval (authorized in `full` mode; `APPROVAL_DENIED` in `restricted` mode) and returns an
updated `WorkbookInfo` with `dirty=false`. Changes are never saved implicitly.

### range_read `action="read-values"` / range_write `action="write-values"`

Reading a bounded window:

```json
{
  "action": "read-values",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "offset": 0,
  "limit": 10
}
```

Response — `RangeValues`:

```json
{
  "range": { "sheet": "Sheet1", "address": "A1:C10", "rows": 10, "cols": 3 },
  "shape": "block",
  "rows": 10,
  "cols": 3,
  "values": [
    [ { "kind": "string", "value": "Header" }, { "kind": "empty" }, { "kind": "empty" } ]
  ],
  "truncated": false,
  "notes": []
}
```

Values are typed: `string`, `number`, `boolean`, `date`, `error`, `empty`; precision loss from
coercion to string is excluded. A request beyond the cell limit is rejected, not truncated;
`truncated=true` appears when the `offset`/`limit` window is clipped.

Writing requires an exact rectangular `payload`:

```json
{
  "action": "write-values",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "B2:C3",
  "payload": {
    "rows": 2,
    "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "=1+1" } ],
      [ { "kind": "boolean", "value": true }, { "kind": "empty" } ]
    ]
  }
}
```

A string beginning with `=` is stored as **text**, not as a formula (use
`range_write` `action="write-formulas"`). A wrong payload shape is rejected with `INVALID_RANGE`
and no changes. The response is a `RangeMutationResult` with `range_before`, `range_after`,
`verified` and `postcondition_values`.

### range_write `action="set-format"`

Sparse write: only the fields explicitly set in `patch` are applied.

```json
{
  "action": "set-format",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C1",
  "patch": {
    "font_bold": true,
    "font_size": 11,
    "fill_color": 15773696,
    "number_format": "0.00",
    "alignment_horizontal": "center"
  }
}
```

Missing fields stay unchanged; `null` resets a value to automatic only where Excel allows it.

### table `action="update-rows"` / table_delete_rows

An update accepts either an **explicit index** or an **unambiguous `where`**:

```json
{
  "action": "update-rows",
  "workbook_id": "1f0c…",
  "name": "Sales",
  "updates": [
    { "index": 2, "values": { "Amount": 120 } },
    { "where": { "Region": "North" }, "values": { "Flag": "Y" } }
  ]
}
```

If `where` does not resolve to exactly one row, the tool returns `COM_BUSY` with a message about the
ambiguity and a preview of the matched rows — there is no automatic "guessing". Deleting rows uses
the standalone `table_delete_rows` with `indices` or `where`; omitting both is a validation error.

### snapshot_create / snapshot_restore

```json
{ "workbook_id": "1f0c…" }
```

Response — `SnapshotManifest`:

```json
{
  "snapshot_id": "snap-…",
  "created_at": "2026-11-21T18:41:56Z",
  "source_path": "C:/data/report.xlsm",
  "file_name": "report.xlsm",
  "stored_path": "<snapshots>/snap-…/report.xlsm",
  "sha256": "…",
  "size": 48213,
  "snapshot_method": "SaveCopyAs",
  "vba_exports": [
    {
      "component_name": "Module1",
      "type": 1,
      "path": "exports/Module1.bas",
      "export_file_sha256": "…"
    }
  ],
  "model_version": 2,
  "external_side_effects_not_undone": true
}
```

The top-level `sha256` is the snapshot **file** hash. Each `vba_exports` entry's
`export_file_sha256` is the SHA-256 of the **exported component file bytes** (F7) — **not** a valid
`expected_source_hash` for the VBA mutation tools, which require the normalized `source_hash` from
`vba_read` `action="read-component"`. `snapshot_list` returns the same manifests; `snapshot_diff`
compares the full manifests and separately reports the top-level file-hash comparison.

Restore:

```json
{ "snapshot_id": "snap-…", "workbook_id": "1f0c…", "confirm": true }
```

The current working copy is closed **without saving**, then a new copy is opened from the baseline;
a `WorkbookInfo` is returned. External macro effects are not rolled back. `confirm` is required.

### vba_replace_component_source

Before replacing, two `vba_read` calls are needed: `list-components` (confirm Standard/Class) and
`read-component` (obtain `source_hash`). Then:

```json
{
  "workbook_id": "1f0c…",
  "component": "Module1",
  "new_source": "Option Explicit\n\nPublic Sub Hello()\nEnd Sub",
  "expected_source_hash": "<source_hash from vba_read read-component>"
}
```

`expected_source_hash` is the stale-revision guard: the server compares it with the current
normalized hash before and **inside** the STA worker. A mismatch raises `REVISION_CONFLICT`. The
guard token is the **normalized `source_hash`** from `vba_read` `read-component` — never the
snapshot export **file** hash. Document modules and UserForms are rejected.

### macro_run

```json
{ "workbook_id": "1f0c…", "macro_name": "Module1.Hello", "args": [] }
```

Gate order:

1. execute mode + writable session (`_require_execute_context`);
2. procedure resolution from the inventory (`Module.Procedure`);
3. approval (key: canonical path + SHA-256 + action + procedure) — authorized in `full` mode,
   `APPROVAL_DENIED` in `restricted` mode;
4. `background_safe` for the current workbook hash **and** `project_source_hash` (BG06);
5. interactive-UI preflight (R13) — the token list in [Abbreviations](#abbreviations);
6. one `Application.Run` on the STA worker.

Possible `MacroRunResult` statuses: `success`, `not_found`, `runtime_error`, `blocked`, `timeout`,
`needs_recovery`. When blocked the macro **is not started**. On a deadline the owned PID is inspected
read-only for a visible modal (modal → `MACRO_RUNTIME_ERROR`, no modal → `MACRO_TIMEOUT`); in both
branches only the automation-owned Excel PID is terminated and the session is recovered in-process
(`MacroRunResult.session_recovery` reports `"recovered"` / `"needs_recovery"` / `null`). Recovery
**discards open workbooks**: the prior `workbook_id` is invalid afterwards and the macro is never
auto-replayed.

### macro `action="classify-background"`

```json
{ "action": "classify-background", "workbook_id": "1f0c…", "procedure": "Module1.Hello", "background_safe": true }
```

The optional `expected_project_source_hash` is an optimistic-concurrency guard (mismatch →
`STALE_TARGET`). Classifying a heuristic or interactive-token procedure as `background_safe` is
rejected with `INTERACTIVE_UI_NOT_SUPPORTED` (the R13 veto). The tool is **ungated** and is not the
security boundary — `macro_run` still applies the approval gate and the R13 veto.

### test `action="run-suite"`

```json
{
  "action": "run-suite",
  "suite": {
    "suite": "smoke",
    "workbook_id": "1f0c…",
    "baseline": "create",
    "cases": [
      {
        "id": "reads-range",
        "steps": [
          { "action": "range_read", "sheet": "Sheet1", "range_address": "A1:B2", "limit": 2 }
        ],
        "assertions": [
          { "type": "range_structure", "sheet": "Sheet1", "address": "A1:B2", "rows": 2, "cols": 2 }
        ]
      }
    ]
  }
}
```

`action` and `type` come from **closed** vocabularies: `TestAction` and `AssertionType`; the schema
forbids extra fields (`extra="forbid"`). Isolation: one immutable baseline snapshot, then **a
disposable working copy per case**. A `macro_run` step may carry an optional data-only
`background_safe` field, and the runner classifies the procedure through the guarded
`MacroService.classify_background` path before dispatch. Compile-before-macro is explicit — include
a `vba_compile` step. The run is persisted (`runs/`) and read via `test` `action="get-run"`.

### vba_compile_diagnostics

```json
{ "workbook_id": "1f0c…" }
```

Fail-closed response:

```json
{
  "workbook_id": "1f0c…",
  "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": {
    "is_compile": false,
    "vbom_available": true,
    "protection": 0,
    "locked": false,
    "component_count": 3,
    "broken_references": []
  },
  "command_discovery": { "route_usable": false },
  "notes": [
    "NOT a compile: non-activating VBA compilation is unavailable on this build.",
    "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"
  ]
}
```

The tool never returns `compiled` or `compile_error`, and never shows or activates UI.
`vba_read` `action="get-compile-diagnostics"` returns the same data from the cache without
re-collecting.

## Background jobs

Four long-running tools accept `background=true`: `workbook` (`open`), `macro_run`,
`test` (`run-suite`) and `pivot_read` (`tables-list`). With `background=true` the tool schedules the
operation on the in-process job registry and returns a prompt `RunningJob` handle:

```json
{ "status": "running", "job_id": "job-…", "poll_after_s": 5 }
```

The client then **pulls/poll**s `job_get(job_id)` no sooner than `poll_after_s` (clamped to
`[1, 30]`, default `5`). The registry **owns the result**, so a client timeout or disconnect never
drops it; a completed result stays retrievable by `job_id` (and, after a restart, from the persisted
`JobStore`).

`JobStatus` fields: `job_id`, `status`, `started`, `cancel_requested`, `poll_after_s`, `result`
(once terminal), `error_code`, `error_text`, `retry_after_s`, `phase`, `outcome`, `retryable` and an
additive `events` history.

`status` is one of:

| Status | Meaning |
|--------|---------|
| `queued` | Scheduled but not yet started on the STA worker. |
| `running` | The operation is executing on the single-STA worker. |
| `succeeded` | The wrapped op returned a value (stored on the job). |
| `failed` | The wrapped op raised a domain or unexpected error. |
| `expired` | The op exceeded its execution deadline (`EXCEL_HUNG`); the call is left running on the STA (never cancelled). |
| `interrupted_unknown` | A persisted non-terminal job reconciled at startup, or a job finalized after a successful `excel_recover`. `outcome="unknown"`, `retryable=false`; never completed, never auto-retried. |
| `unknown` | Reserved; never produced by the registry. |

`job_request_cancel` records cancel **intent** only (`cancel_requested=true`); it cannot interrupt the
uninterruptible COM call, so the job still runs to a terminal state. The lifecycle and persistence
model are described in [Architecture](04-architecture.md); the `excel_recover` interaction is in
[Troubleshooting](10-troubleshooting.md).

## Tool dependencies

1. **`workbook` `action="open"` before any workbook tool.** Every tool with a `workbook_id` requires
   an open working copy; otherwise — `ExcelVbaMcpError` "unknown workbook_id".
2. **`snapshot_create` before any guarded mutation.** Without a successful snapshot the mutation is
   rejected with `SNAPSHOT_FAILED`.
3. **`vba_read` (`list-components` / `read-component`) before `vba_replace_component_source` and
   `vba_insert_code`.** They provide the component name and the current `source_hash` for
   `expected_source_hash`.
4. **`vba_compile_diagnostics` before `vba_read` `get-compile-diagnostics`** — only to seed the
   cache; without the first call the diagnostics are collected again (same result).
5. **`macro` `list` before `macro_run`.** It reveals the exact `Module.Procedure` name and the
   `background_safe` status before execution.
6. **`test` `get-run` after `test` `run-suite`.** The `run_id` comes from
   `TestRunResult.test_run_id`; the tool only reads the persisted run.

Additionally, `workbook` `close` (without saving) and `workbook` `save` are incompatible with
subsequent calls using the same `workbook_id` — after closing, the identifier is no longer valid.

## Mutating and potentially dangerous operations

| Tool | What changes | Why it is gated | How to use it safely |
|------|--------------|-----------------|----------------------|
| `workbook` `open` (writable/execute) | Creates a working copy and opens it with macros | Default open; authorized in `full`, `APPROVAL_DENIED` in `restricted` | Use `mode="inspect", read_only=true` for read-only work; apply `execute` only to verified copies |
| `workbook` `close` (`discard_changes=true`) | Discards unsaved changes and closes | Destructive; refuse-by-default (D1) | Prefer `save_changes=true` or a prior `snapshot_create`; only pass `discard_changes=true` deliberately |
| `workbook` `save` | Writes the working-copy file | Mutating | Combine with a snapshot; do not save "just in case" |
| `excel_save_as` | Writes the working copy to a target path (own permission) | Standalone high-risk; overwrite policy | Check the target; use `overwrite_existing` / `overwrite_source` only after review |
| `snapshot_restore` | Discards the current copy and opens the baseline | Mutating; requires `confirm=true` | Take a fresh `snapshot_create` first; a missing/false `confirm` is rejected before anything is discarded |
| `sheet_delete` | Deletes a sheet with its data | Destructive, standalone | Check `sheet_read` `list`; rely on a snapshot |
| `table_delete_rows` | Deletes table rows | Destructive, standalone (XOR guard) | Use explicit `indices`; use `where` only when unambiguous |
| `table_delete_column` | Deletes a table column | Destructive, standalone (own permission) | Check `table_read` `get`; do not pass an ambiguous header |
| `range_write` | Overwrites values/formulas/format/rules | Mutating | Check the payload size; take a snapshot |
| `range_structural` | Copies/inserts/unmerges cells | Destructive; not idempotent | Pass `shift` explicitly for `insert`; check table intersections |
| `range_delete` | Deletes cells with a shift | Destructive, standalone | Always pass `shift` explicitly; check table intersections |
| `range_merge` | Merges cells, discarding data outside the anchor | Destructive (dual-hint) | Review preconditions; pass `allow_data_loss=true` only after review |
| `name_delete` | Deletes a defined name | Destructive, standalone | Check `name_read` and dependencies; rely on a snapshot |
| `vba_replace_component_source` | Fully replaces a module's code | Destructive, standalone | Pass the exact `expected_source_hash`; check `vba_compile_diagnostics` afterwards |
| `vba_insert_code` | Inserts code into a module | Destructive, standalone | Specify an exact `position`; do not insert code blindly |
| `shape` `set-on-action` | Assigns a macro to a shape | Mutating | Check `shape` `list`; an empty `macro_name` clears the assignment |
| `control` `set-on-action` | Assigns a macro to a Form Control | Mutating | Work only with Form Controls; ActiveX is unsupported |
| `macro_run` | Executes VBA code | Executing; gated by approval (restricted) + `background_safe` + R13 | Only after `macro` `list` / `macro` `classify-background` and preflight; do not run macros with external effects |
| `test` `run-suite` | Runs a declarative suite (may dispatch `macro_run`) | Executing; destructive | Review cases; a `macro_run` step needs a classified procedure |
| `excel_recover` | Terminates the owned Excel process and discards working copies | Recovery; allow-by-default (both modes), audited, pid-scoped, fail-closed; requires `confirm=true` | Use only when `excel_health` shows `needs_recovery` / `running/deadline_exceeded`; originals are never modified, but open working copies are lost |

## Read-only resources

| URI template | Name | MIME | Content |
|--------------|------|------|---------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` | Stage 1 workbook model (sheets, tables, names, shapes, controls, optional VBA) |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` | VBA component inventory + `vba_access` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` | Source text of one VBA component |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` | Persisted declarative test run |

Resources are read-only: they do not mutate the workbook and do not execute VBA.

## Known tool limitations

1. **Compilation is always `indeterminate`.** `vba_compile_diagnostics` cannot confirm compilation on
   the validated build; `is_compile=false`.
2. **Macros require background-safety classification.** `BackgroundSafetyRegistry` starts empty
   (fail-closed); classify a procedure with `macro` `classify-background`. An unclassified procedure
   returns `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`); an interactive one is vetoed
   (`INTERACTIVE_UI_NOT_SUPPORTED`). The R13 token scan is a heuristic veto, not static proof.
3. **No approval grant tool.** Approval/allowlist management cannot be granted through MCP. In the
   default `full` access mode mutations and execute mode are authorized without a grant; in
   `restricted` mode they are refused with `APPROVAL_DENIED`. Consequently, in `restricted` mode
   there is no in-band path to grant approvals or manage the allowlist, so the default execute
   `workbook` `open` is denied and only `inspect`+`read_only` sessions are usable; the remedy is to
   run in `full` mode or pre-authorize via the library.
4. **No PivotTable/PivotCache creation or deletion.** Only existing objects can be managed;
   external/OLAP/Data-Model sources are rejected.
5. **Range and ListObject intersection.** Structural range operations over tables are rejected with
   `LISTOBJECT_INTERSECTION`; use the table tools instead.
6. **Bounded reads.** `range_read` (`read-values` / `read-formulas`) and `table_read` (`read-rows`)
   return bounded windows; read large ranges in chunks via `offset`/`limit`.
7. **No charts, PivotCharts or slicers.** No tools are provided for them.
8. **No workbook creation.** `workbook` `open` opens only existing files; a new workbook cannot be
   created through MCP.
9. **No full recalculation or cross-workbook copy.** `Application.Calculate` is not exposed;
   `range_structural` `copy` works only within one workbook and without the clipboard.
10. **No list pagination.** `table_read` `list` and `pivot_read` `tables-list` return the full
    bounded list without paging.
11. **A dead owned Excel is detected as `needs_recovery`; recovery is agent-directed.**
    `excel_health` reads worker/session bookkeeping without touching COM and additionally runs a
    read-only owned-pid liveness probe: if the hidden `EXCEL.EXE` the worker owns has died while the
    bookkeeping still says `running`/`accepting`, `excel_health` reports `needs_recovery` instead of
    a false `ok`. A `macro_run` hang/timeout rebuilds the instance automatically (F2b). For every
    other `needs_recovery` / `running/deadline_exceeded` state the agent calls `excel_recover`
    (agent-directed, fail-closed, originals-safe); the server never auto-terminates a live pid on a
    deadline. If recovery is refused the fallback is a server restart.
12. **`macro_run` has no snapshot precondition (spec deviation, tracked).** Unlike content-mutating
    tools, `macro_run` does not require a prior `snapshot_create`. The recommended workflow still
    snapshots first so a macro's external side effects have a rollback baseline.
13. **`excel_recover` is the only recovery tool; there is no `excel_recover_session` tool.**
    `SessionService.recover_session`/`close_session` exist internally but are not exposed as MCP
    tools. The `macro_run` hang/timeout path also recovers automatically (F2b) and reports it in
    `MacroRunResult.session_recovery`. An interrupted operation is reported `prior_job="unknown"`
    and is never auto-retried.
14. **`macro_policy.classify_operation` is non-authoritative (dead/illustrative table).** The
    `classify_operation` function in `security/macro_policy.py` is retained for illustration and is
    **not** wired into the runtime enforcement path; the live operation class of each tool is defined
    per-service. Its table predates the current 41-tool catalogue and must not be read as the
    authoritative tool classification — use this reference or the service code for that.