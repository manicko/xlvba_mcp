# Report 10 — Remaining Usability Reconciliation (AUDIT ONLY)

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d7a1e5a7406bbd9f074319e112711e60fc` (verified via `git rev-parse HEAD`).
- **Working tree:** clean except pre-existing untracked `.ai/audit/**` report/plan files (this report included). No source/test/config file was modified by this audit.
- **Role:** Independent Auditor. **No code changes.** Backward compatibility NOT required. Lazy loading excluded.
- **Inputs reconciled:**
  - `.ai/audit/reports/remaining_usability_baseline.md`
  - `.ai/audit/reports/research_remaining_1_contracts_errors.md`
  - `.ai/audit/reports/research_remaining_2_save_as.md`
- **Package root note:** the import package is `excel_vba_mcp` (path `src/excel_vba_mcp/...`), not `xlvba_mcp`. All `file:line` below were read directly from the working tree at HEAD.

## Methodology

1. Read all three input reports in full.
2. Re-read every cited `file:line` in the actual source (primary source of truth). Each row of the findings table records what the source *actually* shows, not what the reports assert.
3. Verified the MCP SDK error contract directly from `.venv/Lib/site-packages/mcp/server/mcpserver/exceptions.py` (text-only `ToolError` / `UnexpectedToolError`).
4. Verified the "no backup mechanism" claim by full-tree search (`0` hits for `.bak`/`backup`; `13` `os.replace` atomic-write sites).
5. Classified each item: **CONFIRMED** (source matches the report's factual claims), **PARTIALLY_CONFIRMED**, **ALREADY_FIXED**, **REFUTED**. No item was found ALREADY_FIXED or REFUTED; two items carry a precision note (marked CONFIRMED WITH NOTE).

## Findings table

| id | Item | file:line (verified) | Current behavior (actual source) | Classification | Recommended change (advisory) | Compat risk |
|---|---|---|---|---|---|---|
| 1a | `vba_compile` → `vba_compile_diagnostics` | handler `api/compile_tools.py:63`; `__all__` `:102`; register `server.py:1979-1984` (name `:1981`); desc `_DESC_VBA_COMPILE` `server.py:407`; import `server.py:55` | Tool is named `vba_compile`; docstring already states "never compiles" (`:64-77`). Not in any `_ANN_*` set. | **CONFIRMED** | Rename handler + registration + `_DESC_*` + `api/__init__.py` + service audit labels (`compile_service.py:198,261,276,352`) + tests + docs (EN+RU) + add benchmark task. | Low — no alias (approved). `TestAction.VBA_COMPILE` (`test_spec.py:73`) is a **different** surface and must NOT be renamed. |
| 1b | `job_cancel` → `job_request_cancel` | handler `api/job_tools.py:113`; `__all__` `:133`; register `server.py:2015-2020` (name `:2017`); desc `_DESC_JOB_CANCEL` `server.py:489`; import `server.py:63` | Tool named `job_cancel`; docstring "intent-only; never interrupts COM" (`:114-124`). No `start_long_op(partial(job_cancel,…))` caller. | **CONFIRMED** | Rename handler + registration + desc + tests (`test_job_tools.py`, `test_job_registry.py`, `test_description_contract.py`) + docs + benchmark `benchmark_tasks.yaml:471-477`. | Low. |
| 1c | `test_run_suite` vs `test_suite_run` — keep superset | `test_run_suite` `api/test_tools.py:163`; `test_suite_run` `:189`; register `:1907-1912` / `:1913-1918`; internal `partial(test_run_suite, spec)` `:229`; direct `test_run_suite(spec)` `:234` | `test_suite_run` is the **superset**: accepts `TestSuiteSpec \| dict` + `background`; validates mapping pre-Excel (`_parse_suite` `:237`). `test_run_suite` is typed-spec, blocking only. `test_suite_run` calls `test_run_suite` directly (`:234`) and via `start_long_op(partial(test_run_suite, spec))` (`:228-233`). | **CONFIRMED** | Canonical = `test_suite_run`. Remove `test_run_suite`; inline/rename the typed-spec core to a private helper so one execution path remains. Update audit label `test_runner.py:1763`, tests, docs, benchmark `must_not` (`:569`). | Low. `TestAction.VBA_COMPILE` and `_EXECUTING_TOOLS` are NOT this tool. |
| 1d | `TestAction.VBA_COMPILE` is NOT the tool name | `domain/test_spec.py:73` `VBA_COMPILE = "vba_compile"`; dispatched `test_runner.py:622` | A declarative-suite **action enum** member, semantically distinct from the MCP tool; shares the string only. | **CONFIRMED** | Do NOT rename as part of the tool rename (separate, larger change to suite schema + many tests). Flag only. | High if wrongly renamed (breaks suite contract). |
| 2a | Error `to_dict` keys are `code`/`context` | `domain/errors.py:335-353` (keys `code` `:345`, `context` `:349`) | Emits 8 keys: `code, message, hresult, hresult_hex, context, errorCategory, isRetryable, suggestedNextActions`. Target wants `errorCode`/`details`. | **CONFIRMED** | Rename `code`→`errorCode`, `context`→`details` (no alias). Update `test_error_envelope.py:30,41-42`. | Low (compat not required). |
| 2b | `call_tool` re-raise hatch leaks bare error | `server.py:2484-2502`; re-raise at `:2500-2501` | Domain cause → structured `ToolError(json)`. Validation cause → `INVALID_RANGE` envelope. **No cause → bare `raise`** → SDK `UnexpectedToolError("Error executing tool <name>")`. Proven by `test_mcp_error_surfacing.py:75-80`. | **CONFIRMED** | Replace the `raise` with a synthesized fallback envelope (additive `INTERNAL_ERROR`, see U-1) so every `isError` carries the 6 fields. | Low. |
| 2c | SDK is text-only; no structuredContent | `.venv/.../mcp/server/mcpserver/exceptions.py` (`ToolError`, `UnexpectedToolError`) | `ToolError` → `is_error=True` with message in `content` (text). `UnexpectedToolError` message is only `"Error executing tool <name>"`. No structured error channel. | **CONFIRMED** | Guarantee = additive **text-JSON** envelope at `call_tool`; **no SDK change**, **no error `structuredContent`**. | None. |
| 3 | `WorkbookInfo` path vs source_path; `original_name` redundant | `domain/models.py:86-99` (`path` `:92`, `source_path` `:93`) | `path` = actual working path (`wb.FullName`); `source_path` = original allowlisted path (nullable). No `original_name` field anywhere in `src` (0 hits). `name` already = `wb.Name` (original file name). | **CONFIRMED** | Do NOT add `original_name` (redundant with `name`; original *path* is `source_path`). Document the two-path model. | None. |
| 4 | Cross-scope name conflict + update scope preservation | `names_adapter.py` `_guard_no_conflict:576-618` (called `create_name:422`); `update_name:439-470`; `_resolve_name:535-574` | Same scope+sheet → `reason="duplicate"` (`:596-606`); same id at different (scope,sheet) → `reason="scope_conflict"` (`:607-618`, stricter than raw Excel). `update_name` resolves in given scope and writes `RefersTo` (`:457-459`), scope preserved. Resolve: 0→not-found, >1→`ambiguous`, 1→resolved. | **CONFIRMED** | Keep as intended product semantics (stricter than Excel). Document examples; no code change. | None (doc-only). |
| 5 | `snapshot_restore` / `excel_recover` lack `confirm` | `snapshot_restore` `api/test_tools.py:122-143`; `excel_recover` `api/recovery_tools.py:46-69` | Neither has a `confirm` param. `snapshot_restore` closes working copy **without saving** (discards unsaved) + opens fresh copy from baseline. `excel_recover` terminates owned pid, discards working copies (originals safe). Both in `_ANN_DESTRUCTIVE`. **Note:** `RecoveryResult.confirmed` (`recovery_tools.py:36`) is an *output* field, not an input gate. | **CONFIRMED WITH NOTE** | Add `confirm: bool=False` (agent intent, not approval); missing/false → reject before any destructive action (additive `CONFIRMATION_REQUIRED`, see U-2). Descriptions must state discard vs preserve and "does not undo external macro side effects". | Low. |
| 6 | Save As lifecycle facts | `_SessionRecord` `session_service.py:83-94`; `save_workbook` `:474-538` (`wb.Save()` `:512`); `_materialize_working_copy` `:1199-1206` (`shutil.copy2` `:1205`); `_find_open_workbook` `:1236-1252`; `PathPolicy.validate_path` `path_policy.py:110-185`; atomic `artifact_store._atomic_write:321-350` | In-place `wb.Save()`, no path arg; working copy via `shutil.copy2`; **no source hash** on record; **no backup mechanism** (0 hits). `_find_open_workbook` matches `str(wb.FullName).lower()==working_path.lower()` (`:1242-1246`). | **CONFIRMED** | Implement Save As as `wb.Save()` (STA) + pure-FS atomic copy (event loop); never `wb.SaveAs` (see Save As section). | Low (new tool). |
| 7 | No new `ErrorCode`; `SCHEMA_VERSION=1` additive | `domain/errors.py:28-67` (29 members incl. `REVISION_CONFLICT:62`, `PROTECTED_TARGET:45`, `STALE_TARGET:47`); `models.py:65` `SCHEMA_VERSION=1` | 29 codes; additive Phase-2 precedent exists. `SCHEMA_VERSION=1`. Reuse `REVISION_CONFLICT`/`PROTECTED_TARGET`/`PATH_NOT_ALLOWED` for Save As. | **CONFIRMED** | Reuse existing codes for Save As; keep `SCHEMA_VERSION=1` (additive DTOs only). New codes only if U-1/U-2 approved. | Low. |

## Per-item verification detail

### Item 1 — Rename / consolidation surface
- All four registration sites match the reports exactly (`vba_compile` name at `server.py:1981`; `job_cancel` at `:2017`; `test_run_suite` at `:1909`; `test_suite_run` at `:1915`).
- `test_suite_run` is the superset (mapping + `background`); `test_run_suite` is the typed-spec blocking core. Internal coupling is exactly two references: `partial(test_run_suite, spec)` at `test_tools.py:229` and direct `await test_run_suite(spec)` at `:234`.
- `TestAction.VBA_COMPILE = "vba_compile"` (`test_spec.py:73`) is a suite **action**, not the tool — confirmed distinct surface.
- `_EXECUTING_TOOLS` (`macro_policy.py:161-169`) contains `vba_compile`/`test_run_suite` but is explicitly documented non-wired/stale (`:186-191`); updating it is optional cleanup.
- Benchmark references confirmed: `job_cancel` at `benchmark_tasks.yaml:471-477`; `test_run_suite` in `must_not` at `:569` (plus `test_suite_run` at `:554-570`).

### Item 2 — Guaranteed error contract
- `to_dict` emits `code`/`context` (not `errorCode`/`details`) — confirmed at `errors.py:345,349`.
- Boundary: `call_tool` (`server.py:2484-2502`) re-raises at `:2500-2501` when neither a domain cause nor a `ValidationError` is found. That is the leak hatch that produces a bare `"Error executing tool <name>"`.
- SDK confirmed text-only: `ToolError` carries the message in `content`; `UnexpectedToolError` withholds the original. **No SDK change and no error `structuredContent` are possible/needed** — the guarantee is an additive text-JSON envelope built at `call_tool`.
- To guarantee the 6 fields (`errorCode, errorCategory, message, isRetryable, suggestedNextActions, details`) on **every** `isError` response you need **both**: (a) rename in `to_dict`, and (b) close the `raise` hatch with a fallback envelope. The Pydantic class alone is insufficient.

### Item 3 — WorkbookInfo
- `path` = working path, `source_path` = original (nullable) — confirmed `models.py:92-93`.
- No `original_name` in `src` (0 hits) — redundant with `name`; do not add.
- Snapshot restore sets `source_path = manifest.source_path` (`snapshot_service.py:346`), i.e. the snapshot's stored source, **not** the pre-restore original — confirmed (resolves research question §3/§10.6).

### Item 4 — Names
- Conflict rule, update-preserves-scope, and resolution order all match the report (`names_adapter.py:576-618`, `:439-470`, `:535-574`). The cross-scope rejection is stricter than raw Excel and is intended product behavior.

### Item 5 — Confirm flags
- `snapshot_restore(snapshot_id, workbook_id=None)` and `excel_recover(reason="agent_directed")` have no `confirm` — confirmed. Discard semantics: restore discards the working copy without saving; recover terminates the owned pid and discards working copies (originals never modified). `RecoveryResult.confirmed` is an output echo, not an input gate (precision note).

### Item 6 — Save As lifecycle (see dedicated section).

### Item 7 — Error codes / schema
- 29 `ErrorCode` members; `REVISION_CONFLICT`, `PROTECTED_TARGET`, `STALE_TARGET`, `PATH_NOT_ALLOWED` all present and reusable. `SCHEMA_VERSION=1` (`models.py:65`); additive DTOs do not bump it. No new members required for Save As.

## Save As lifecycle verification

All facts below were read directly and **CONFIRMED**:

1. **No source hash on the session record.** `_SessionRecord` (`session_service.py:83-94`) fields = `workbook_id, canonical, working_path, mode, read_only, dirty, info, lock`. No hash/stat/backup/target field. ⇒ no way today to detect on-disk source drift between open and write-back.
2. **In-place save only.** `save_workbook(workbook_id, *, timeout_s)` (`:474-479`) has **no path arg**; STA `save_fn` resolves the open proxy by `record.working_path` (`:505`) and calls `wb.Save()` (`:512`) — writes the working copy in place, never the original.
3. **Working copy materialization.** `_materialize_working_copy` (`:1199-1206`) = `shutil.copy2(canonical, work/<workbook_id><ext>)` (`:1205`); filename is `workbook_id`+suffix (no source basename/hash). Original opened only as the copy.
4. **Path canonicalization.** `PathPolicy.validate_path` (`path_policy.py:110-185`) rejects empty/device/UNC, canonicalizes via `_resolve_candidate` (resolve + realpath, prefer realpath), enforces root containment (SYMLINK_ESCAPE/TRAVERSAL/OUTSIDE_ROOTS), `expected_root`, and extension allowlist; returns the resolved `Path`. Pure gate, no writes — reusable for the Save As **target**; compare target↔source with `os.path.normcase` (not `Path` equality).
5. **Atomic primitive.** `artifact_store._atomic_write` (`:321-350`) = `mkstemp(dir=target.parent)` → write → `flush` → `os.fsync` → `os.replace`; wraps `OSError` as `ArtifactWriteError`; unlinks temp in `finally`. This is the canonical shape to reuse (temp on the **target** filesystem for same-volume atomic replace).
6. **No backup mechanism.** Full-tree search: `0` hits for `.bak`/`backup`/`preserve.*existing`; `13` `os.replace` sites (all atomic writes, none preserve the prior destination). Backup-on-overwrite must be designed new.
7. **`wb.SaveAs` would break session resolution — CONFIRMED (mechanism) / HIGH (Excel semantics).** `_find_open_workbook` (`:1236-1252`) matches the open proxy **only** by `str(wb.FullName).lower() == working_path.lower()` (`:1242-1246`). `wb.SaveAs(target)` repoints the open Workbook object's `FullName` to the target; every later tool that resolves by `record.working_path` (`save_workbook:505`, plus `vba_project_service`/`snapshot_service`/`macro_service` `_find_open_workbook`) would then fail to find the workbook, and the "original never modified" invariant would be destroyed for the rest of the session. ⇒ Save As must be `wb.Save()` (flush working copy) + pure-FS atomic copy; the session stays bound to the working copy (`workbook_id`/`working_path`/`canonical` unchanged).

**Net:** the research-2 lifecycle model is accurate and internally consistent with the invariants (single hidden Excel, single-STA FIFO, working-copy model, originals protected). The only design gaps to close are: (a) capture a source identity (hash, or size+mtime) at open on `_SessionRecord`; (b) add the atomic-copy + backup + verify algorithm; (c) reuse existing codes (`PROTECTED_TARGET`/`REVISION_CONFLICT`/`PATH_NOT_ALLOWED`); (d) keep `SCHEMA_VERSION=1`.

## Unresolved decisions for Tech Lead

- **U-1 — Fallback error code for the closed hatch (item 2b).** Approve an **additive** `ErrorCode.INTERNAL_ERROR` (category `internal`/`validation`, `isRetryable=False`, actions `["inspect server logs; the failure was not a domain error"]`) vs reusing `COM_BUSY`/`INVALID_RANGE`. Additive codes have precedent (`errors.py:43-67`); if approved, `category_for`/`retryable_for`/`actions_for` must map it (totality test `test_error_envelope.py:47` will then require it). Recommendation: additive `INTERNAL_ERROR`.
- **U-2 — Confirm-flag error code (item 5).** Additive `ErrorCode.CONFIRMATION_REQUIRED` (category `validation`, not retryable) vs reuse `APPROVAL_DENIED` (semantically wrong — that is policy denial). Recommendation: additive `CONFIRMATION_REQUIRED`. Also decide whether `confirm` should later gate `excel_close(save_changes=true)`/`excel_save` (out of current scope).
- **U-3 — Save As overwrite-rejection code (item 6/7).** Confirm `PROTECTED_TARGET` for "existing target / source protected" (with `context.reason`) vs a dedicated code (forbidden by closed `SCHEMA_VERSION=1`). `REVISION_CONFLICT` for source-modified-since-open; `PATH_NOT_ALLOWED` for path/extension. Recommendation: reuse `PROTECTED_TARGET`.
- **U-4 — Source identity at open (item 6).** Hash the whole source at `excel_open` (strong, workbooks bounded) vs stat-only (size+mtime) at open + hash at Save-As (cheaper). Recommendation: hash at open; confirm cost is acceptable.
- **U-5 — `source_sha256` visibility (item 3/6).** Keep internal to `_SessionRecord` (no DTO change) vs expose additively on `WorkbookInfo`. Recommendation: keep internal for now.
- **U-6 — Backup retention / location (item 6).** Exactly one newest `.bak` per target (prune older) vs retain all; and whether `.bak` beside user files in the destination dir is acceptable vs an artifact root. Product decision.
- **U-7 — Interrupted-state recovery (item 6).** "Destination absent + `.bak` present ⇒ report `outcome="unknown"`, require explicit retry" vs auto-resume (rename `.bak` back). Auto-resume risks masking a partial write; recommendation: report unknown, no auto-resume.
- **U-8 — `macro_policy._EXECUTING_TOOLS` (item 1).** Update the stale illustrative set (`macro_policy.py:161-169`) to the new names or leave (documented non-wired). Optional cleanup; validator to confirm leaving stale names is acceptable.
- **U-9 — `TestAction.VBA_COMPILE` (item 1d).** Confirm the tool rename must NOT touch the declarative-suite action enum (`test_spec.py:73`/`test_runner.py:622`). Renaming the action is a separate, larger change (suite schema + docs + many tests).
- **U-10 — Description-contract test coupling.** `test_description_contract.py` hard-codes the full guard sentence and the `_CRITICAL`/`_TRUTH`/`_EXAMPLE_TOOLS` name lists; any rename or boilerplate move must update it in lockstep (single commit) to avoid a red suite mid-migration.

## Classification summary

- **CONFIRMED:** 1a, 1b, 1c, 1d, 2a, 2b, 2c, 3, 4, 6, 7.
- **CONFIRMED WITH NOTE:** 5 (output field `RecoveryResult.confirmed` ≠ input `confirm`).
- **PARTIALLY_CONFIRMED / ALREADY_FIXED / REFUTED:** none.

No source, test, or config file was modified. This report is the only artifact written.