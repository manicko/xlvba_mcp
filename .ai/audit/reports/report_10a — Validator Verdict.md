# Report 10a — Validator Verdict (Remaining-Usability Remediation)

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d7a1e5a7406bbd9f074319e112711e60fc` (re-verified via `git rev-parse HEAD`).
- **Working tree:** clean except pre-existing untracked `.ai/audit/**` report/plan files. No source/test/config modified by this validation.
- **Role:** Independent Validator. **No changes.** I did not research or plan; I challenged the proposed contracts against actual source.
- **Inputs reconciled:** `report_10 — Remaining Usability Reconciliation.md`, `research_remaining_1_contracts_errors.md`, `research_remaining_2_save_as.md`, `remaining_usability_baseline.md`.
- **Package root:** import package is `excel_vba_mcp` (`src/excel_vba_mcp/...`). All `file:line` read directly at HEAD.

## Verdict table

| # | Contract under validation | Verdict | Primary evidence (file:line) |
|---|---|---|---|
| 1 | Renames `vba_compile`→`vba_compile_diagnostics`, `job_cancel`→`job_request_cancel`; canonical `test_suite_run`, `test_run_suite` fully removed (no alias); migration surface complete | **APPROVE** (with blocking gaps) | compile_tools.py:63,102; server.py:1979-1984; job_tools.py:113,133; server.py:2015-2020; test_tools.py:163,189,229,234; server.py:1907-1918 |
| 2 | Error envelope guarantees 6 fields on EVERY `isError`; hatch closed; text-JSON only; `interrupted_unknown`/stall/permission semantics truthful | **APPROVE** (with blocking gaps) | errors.py:335-353,168-178,162-164; server.py:2484-2502; exceptions.py (SDK) |
| 3 | `xlvba_environment` routing is capability-based, does not claim to detect other servers | **APPROVE** | environment_tools.py:1-234; server.py:1004-1019 |
| 4 | Description compression removes boilerplate WITHOUT removing safety-critical text; no misleading text | **APPROVE** (conditional) | server.py:407-421,489-500,455-472; test_description_contract.py:54-59,110-137 |
| 5 | `WorkbookInfo` semantics + name-scope conflict spec match code | **APPROVE** | models.py:86-99; names_adapter.py:576-618,439-470,535-574; snapshot_service.py:346 |
| 6 | `confirm` flags reject before destructive action; descriptions state discard vs preserve, no external-macro-rollback implication | **APPROVE** (with blocking gaps) | test_tools.py:122-143; recovery_tools.py:36,46; session_service.py:748,865,947; server.py:292-301,272-286 |
| 7 | Save As: default safe; overwrite only via explicit flags; normcase/realpath compare; hash-at-open + drift detection; atomic temp→replace + backup; failed/interrupted NEVER success; session stays bound (no `wb.SaveAs`) | **APPROVE** (design is safe; blocking impl requirements) | session_service.py:83-94,474-538,1199-1206,1236-1252; path_policy.py:110-185,253-273; artifact_store.py:321-350 |
| 8 | Additive only; `SCHEMA_VERSION=1`; no new `ErrorCode` required (reuse); no lazy loading | **APPROVE** | models.py:65,68-71; errors.py:28-67; server.py:1845+ |

**Final verdict: APPROVE the remediation direction.** All eight contracts are consistent with actual source. The Save As design does **not** silently overwrite a source and does **not** report a failed save as successful (see Rejected-proposals section). Approval is conditional on the Planner closing the blocking gaps below — chiefly the **omitted error-envelope test surface** (item 2) and the **mandatory source-identity capture** for Save As (item 7).

---

## Per-item evidence and reasoning

### Item 1 — Renames + `test_run_suite` removal — APPROVE (blocking gaps)
Verified facts:
- `vba_compile` handler `compile_tools.py:63`; `__all__` `:102`; registered `server.py:1979-1984` (name `:1981`); desc `_DESC_VBA_COMPILE` `server.py:407-421` (truthful "does NOT compile" `:409`); import `server.py:55`; re-export `api/__init__.py:10,59`; audit labels `compile_service.py:198,261,276,352`. **Not** in `_ANN_READ_ONLY` (`server.py:1747-1790`; comment `:1746`).
- `job_cancel` handler `job_tools.py:113`; `__all__` `:133`; registered `server.py:2015-2020` (name `:2017`); desc `server.py:489-500` ("intent-only; never interrupts" `:493`); import `server.py:63`. **No** `start_long_op(partial(job_cancel,…))` caller exists (grep: only docstrings `job_registry.py:29`). Not in any `_ANN_*` set.
- `test_run_suite` `test_tools.py:163`; registered `server.py:1907-1912` (name `:1909`); desc `server.py:303-319`. `test_suite_run` `:189` is the **superset** (mapping + `background`, validated pre-Excel by `_parse_suite` `:237`), registered `:1913-1918`. Internal coupling is exactly two sites: `partial(test_run_suite, spec)` `:229` and `await test_run_suite(spec)` `:234`; audit label `test_runner.py:1763`.
- `TestAction.VBA_COMPILE = "vba_compile"` (`test_spec.py:73`) is a declarative-suite **action**, dispatched at `test_runner.py:622` — a different surface; must NOT be renamed (tests `test_domain_models.py:144,184`, `test_test_runner.py:891,928,958`, `test_end_to_end_test_runner.py:870,883` depend on it).
- `_EXECUTING_TOOLS` (`macro_policy.py:161-169`) is documented stale/non-wired (`:186-191`); updating optional.

Migration surface is **complete in the research** but the following are under-specified and the Planner MUST update in the same commit:
- **Cross-reference descriptions:** `server.py:486` (`_DESC_JOB_GET` "Next actions: job_cancel …") and `server.py:330` (`_DESC_TEST_SUITE_RUN` "use test_run_suite") are not called out in research §1.2/§1.3 tables.
- **Tool-count test:** `test_description_contract.py:84-87` hard-codes `len(tools) == 92`. Removing `test_run_suite` (−1) and adding `excel_save_as` (+1) nets 92 **only if both land together**; otherwise the constant must change. The Planner must reconcile this explicitly.
- **Benchmark:** `job_cancel` `.ai/tmp/benchmark_tasks.yaml:471-477`; `test_run_suite` in `must_not` `:569` (+`test_suite_run` `:554-570`).

### Item 2 — Guaranteed error envelope — APPROVE (blocking gaps)
- `to_dict` emits `code`/`context` (not `errorCode`/`details`) among 8 keys — `errors.py:344-353` (`code` `:345`, `context` `:349`). Confirmed gap.
- Boundary hatch: `call_tool` `server.py:2484-2502`; `_find_domain_cause` `:2383-2398`; `_validation_error_as_domain` `:2432-2461`; **bare `raise` at `:2500-2501`** is the leak. Proven by `tests/unit/test_mcp_error_surfacing.py:75-80` (`test_non_domain_error_surfaced_as_unexpected_tool_error`).
- SDK is **text-only**: `ToolError` → `is_error=True` with message in `content`; `UnexpectedToolError` message is only `"Error executing tool <name>"` (`.venv/.../mcp/server/mcpserver/exceptions.py`). No `structuredContent` channel — the text-JSON envelope at `call_tool` is the only mechanism. Confirmed.
- Truthful semantics already in code: `interrupted_unknown` → category `interrupted_unknown`, `retryable_for` False, actions `["never resubmit; report status"]` (`errors.py:168-178`, esp. `:177`); only `COM_BUSY`/`RPC_E_CALL_REJECTED` retryable (`:162-164`); `APPROVAL_DENIED` → `permission_denied`, not retryable (`:127`; test `test_error_envelope.py:96-99`); `EXCEL_HUNG`/`MODAL_UI_BLOCKED` → category `suspected_stall` (`:151-152`), and `_DESC_JOB_GET` says "SUSPECTED_STALLED … not confirmed hung" (`server.py:482-484`). Confirmed — stall is never labelled a confirmed hang; permission denial is never auto-retryable.

**Blocking gap (the single most-missed surface):** the research (§2.3-2) only lists `test_error_envelope.py:30-31,41-42` for the `code`→`errorCode` / `context`→`details` rename. It **omits** `tests/unit/test_mcp_error_surfacing.py:71-72` and `:110-115`, which assert `payload["code"]` / `payload["context"]`. The Planner MUST update these in lockstep, and MUST rewrite `:75-80` (which currently asserts the leak) once the hatch is closed. The fallback envelope must be built from the caught exception's `str()` only (never the live object), reusing the `_jsonable` pattern (`server.py:2420-2427`).

### Item 3 — Environment routing — APPROVE
`xlvba_environment` (`environment_tools.py:1-234`) is non-COM, side-effect-free, and makes **no** claim to detect other servers; `_DESC_ENVIRONMENT` (`server.py:1004-1019`) reports only this server's mode/policy/build/capabilities. The proposed capability-based routing is therefore additive and safe **iff** it (a) phrases guidance as "this server provides capability X" and (b) never asserts presence/absence of `excel-mcp_*` or any other server. The current text already satisfies (b); the Planner must keep it that way.

### Item 4 — Description compression — APPROVE (conditional)
Safety-critical phrases that MUST survive compression are real and test-enforced: `_TRUTH` `test_description_contract.py:55-59` (`vba_compile` "does NOT compile"; `macro_classify_background` "never grants run permission"; `job_cancel` "intent-only; never interrupts"); the uniform guard sentence is asserted at `:110-137`. The 133-char guard sentence and the `(full-access default; restricted requires approval)` clause are genuine repeated boilerplate (research §6) and safe to externalize to `xlvba://glossary` **provided** the test at `:110-137` is changed to assert the token in the same commit. Compression must NOT touch per-tool "Do not use when" routing (`server.py:307-308,329-330,396-397,433-435`), tool-specific preconditions (e.g. `range_write_values` "does NOT snapshot"), or the one `Example:` per `_EXAMPLE_TOOLS` (`test_description_contract.py:47-52`). No misleading text is introduced by the proposal.

### Item 5 — WorkbookInfo + name-scope — APPROVE
- `WorkbookInfo` `models.py:86-99`: `path` = working path (`:92`), `source_path` = original allowlisted, nullable (`:93`). **No `original_name`** anywhere in `src` (grep 0 hits) — redundant with `name`. Do not add.
- Snapshot restore sets `source_path = manifest.source_path` (`snapshot_service.py:346`) — the snapshot's stored source, **not** the pre-restore original. Matches the spec.
- Names: `_guard_no_conflict` `names_adapter.py:576-618` (duplicate `596-606`, cross-scope `scope_conflict` `607-618`), called only by `create_name` `:422`; `update_name` `:439-470` resolves in-scope and writes `RefersTo` `:459` with scope preserved; `_resolve_name` `:535-574` (0→not-found `:558`, >1→ambiguous `:563`, 1→resolved). Stricter-than-Excel cross-scope rejection is intended product semantics; document, do not relax.

### Item 6 — Confirm flags — APPROVE (blocking gaps)
- `snapshot_restore(snapshot_id, workbook_id=None)` `test_tools.py:122-143` — no `confirm`; discards working copy without saving (`snapshot_service.py:331-336`), external side effects not undone (`:309-310,359`).
- `excel_recover(reason="agent_directed")` `recovery_tools.py:46` — no `confirm`; `RecoveryResult.confirmed` (`:36`) is an **output** echo, not an input gate; recover discards working copies only, originals never modified (`session_service.py:748,865,947`). Both in `_ANN_DESTRUCTIVE` (`server.py:1795-1796`).
- Descriptions: `_DESC_SNAPSHOT_RESTORE` already states discard + "does NOT undo external macro side effects" (`server.py:296-298`) — truthful. `_DESC_EXCEL_RECOVER` (`:272-286`) states "discards the working copies" but does **not** yet state the external-side-effect caveat and does not claim rollback — the Planner MUST add the explicit "does not roll back external side effects of any macro that already ran" and the `confirm=true` requirement, and place the `confirm` check **before** `recover_owned_process` / `restore_snapshot` so no termination/close occurs when `confirm` is missing/false.

### Item 7 — Save As — APPROVE (design is safe; blocking impl requirements)
Current state (all confirmed): `_SessionRecord` has **no** source hash/stat/backup/target field (`session_service.py:83-94`); `save_workbook` has no path arg and calls `wb.Save()` in place on the working copy (`:474-538`, `:512`); `_materialize_working_copy` = `shutil.copy2` (`:1199-1206`); `_find_open_workbook` matches by lowercased `FullName` (`:1236-1252`, `:1246`) — so `wb.SaveAs` would repoint the open object and break every later working-path resolution, destroying original-protection; `PathPolicy.validate_path` canonicalizes via resolve+realpath (`path_policy.py:110-185`, `_resolve_candidate` `:253-260`) and compares with `normcase` (`_normalize` `:272-273`), performs no writes; `artifact_store._atomic_write` is the canonical `mkstemp(target.parent)→fsync→os.replace` with `finally` unlink (`:321-350`) and **no** backup; no backup mechanism exists anywhere (grep 0 hits). No `excel_save_as`/`overwrite_*` exists yet.

The research-2 design satisfies the two non-negotiables:
- **No silent source overwrite:** default `overwrite_existing=False`/`overwrite_source=False` rejects existing target and source-overwrite; source is written only with `overwrite_source=True` AND a source-hash match; target↔source compared by `normcase`, never string equality; target==working-copy rejected.
- **No ambiguous success:** `outcome="ok"` only after final size+hash verify; any unprovable/interrupted state is `outcome="unknown"` with `backup_path` and a suggested action; never "no exception ⇒ success".

**Blocking implementation requirements (MUST):**
1. Capture source identity (`source_sha256`, optionally size+mtime) on `_SessionRecord` **at open, before** `_materialize_working_copy` (`session_service.py:306`), and **re-hash + normcase-compare** at Save-As; mismatch ⇒ `REVISION_CONFLICT`. Without this, source-overwrite could silently clobber an externally-modified file — the central safety property.
2. Create the `.bak` of the existing destination **before** `os.replace(tmp→dest)`; stage temp on the **target** filesystem; verify size+sha before reporting `ok`.
3. Interrupted rule: destination absent + matching `.bak` present ⇒ `outcome="unknown"`, **no auto-resume**, require explicit retry (never orphan a second backup).
4. Session stays bound: `wb.Save()` (flush working copy) + pure-FS atomic copy; **never** `wb.SaveAs`. `workbook_id`/`working_path`/`canonical` unchanged; a later `excel_save` writes the working copy, not the target.
5. Reuse existing codes (`PROTECTED_TARGET`/`REVISION_CONFLICT`/`PATH_NOT_ALLOWED`); do not bump `SCHEMA_VERSION`.

### Item 8 — SDK/Kilo compatibility — APPROVE
`SCHEMA_VERSION = 1` (`models.py:65`); DTOs are additive `extra="forbid"` (`:68-71`) and additive fields do not bump the schema. `ErrorCode` has 29 members with a closed Stage-1 set plus **additive** Phase-2 precedent (`errors.py:28-67`); Save As reuses `PROTECTED_TARGET` (`:45`), `REVISION_CONFLICT` (`:62`), `PATH_NOT_ALLOWED` (`:42`) — **no new member required**. Tools are registered eagerly at import (`server.py:1845+`; `test_description_contract.py:3-5`) — no lazy loading. Additive `INTERNAL_ERROR`/`CONFIRMATION_REQUIRED` (U-1/U-2) are **permitted** (additive, no schema bump) but **not required**; if adopted, `category_for`/`retryable_for`/`actions_for` and the totality test (`test_error_envelope.py:47-62`) must be extended.

---

## Rejected proposals
None of the proposed contracts is rejected. For the record, I **reject** any Save As variant that: (a) defaults to overwriting an existing target or the source; (b) compares target↔source by raw string/`Path` equality instead of `normcase`+realpath; (c) omits source-identity capture/re-hash (silent clobber of an externally-modified source); (d) emits `outcome="ok"` from "no exception" rather than a final size+hash verify; (e) uses `wb.SaveAs` (rebinds the session and breaks original-protection); or (f) auto-resumes an interrupted `.bak` in a way that masks a partial write. The research-2 design avoids all six; the Planner must not regress any of them during implementation.

## Blocking concerns the Planner MUST address
1. **Item 2 test surface (highest risk):** update `tests/unit/test_mcp_error_surfacing.py:71-72,110-115` (key rename) and rewrite `:75-80` (leak assertion) in the **same** commit as the `to_dict` rename + hatch closure. Research omitted this file.
2. **Item 1 lockstep edits:** update `test_description_contract.py` `_CRITICAL`/`_TRUTH`/`_EXAMPLE_TOOLS` and the `==92` count test; rename the cross-refs at `server.py:486` and `:330`; update benchmark `job_cancel`/`test_run_suite`. Do **not** touch `TestAction.VBA_COMPILE` (`test_spec.py:73`) or its tests.
3. **Item 7 source identity:** add `source_sha256` to `_SessionRecord` at open (before materialize) and enforce re-hash + normcase compare at Save-As; `outcome="ok"` only after final verify; backup-before-replace; interrupted ⇒ `unknown`, no auto-resume; `wb.Save()`+copy, never `wb.SaveAs`.
4. **Item 6 confirm ordering + text:** reject missing/false `confirm` **before** any termination/close; add the external-side-effect caveat to `_DESC_EXCEL_RECOVER`.
5. **Item 2 fallback code:** choose reuse vs additive `INTERNAL_ERROR` deliberately; if additive, extend the totality maps + test.

## Advisory recommendations
- Keep `source_sha256` internal to `_SessionRecord` (no DTO change) for now (U-5).
- Decide backup retention/location (U-6) and target-keyed locking for concurrent Save-As/open (research §10.7) — out of current scope but note the race.
- Leave `_EXECUTING_TOOLS` stale names acceptable (documented non-wired); optional cleanup.
- Update `_DESC_SAVE` (`server.py:249` "there is no save-as") when `excel_save_as` lands, and add `excel_save_as` to `_CRITICAL`/docs/benchmark.

## Final verdict
**APPROVE.** The contracts are architecturally sound, match source at `40c8b4d`, and the Save As design is fail-closed against both silent source overwrite and ambiguous success. Approval is conditional on the five blocking concerns above; the item-2 omitted test file and the item-7 mandatory source-identity capture are the two that, if missed, would invalidate the safety guarantees.