# Plan 05 — Remaining Usability & Save As Implementation

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d7a1e5a7406bbd9f074319e112711e60fc` (verified via `git rev-parse HEAD`).
- **Role:** Planner. **PLAN ONLY — no code/test/config change in this document.**
- **Package root:** import package is `excel_vba_mcp` (`src/excel_vba_mcp/...`). All `file:line` anchors were read directly at HEAD.
- **Authority order:** the task brief (binding Tech Lead decisions D-R1..D-R7) > `report_10a — Validator Verdict` (binding conditions) > `report_10` / `research_remaining_1` / `research_remaining_2` > source at HEAD. Where a research recommendation conflicts with a D-R decision, **D-R wins** (notably: Save As source-drift code is `STALE_TARGET` per D-R4, not `REVISION_CONFLICT`; backup is a single `<target>.bak` per D-R5, not timestamped).
- **Global constraints:** Backward compatibility **NOT required** (old names removed outright, no aliases). Lazy loading **excluded** (tools stay eagerly registered at import). `SCHEMA_VERSION` stays **1** (all DTO/code additions are additive).

---

## 1. Executive summary

This plan closes the remaining-usability gaps for the `xlvba` MCP server in six dependency-ordered work packages (U6-A..U6-E, U7). The work is deliberately split by change class so each package is independently executable, reviewable, and revertible:

- **U6-A — Canonical tool naming & consolidation** (tool-lifecycle + description + tests + benchmark): rename `vba_compile`→`vba_compile_diagnostics`, `job_cancel`→`job_request_cancel`; **remove** `test_run_suite` entirely (no alias), keep `test_suite_run` as the single canonical entry point.
- **U6-B — Guaranteed structured error envelope** (schema-additive + runtime): rename envelope keys `code`→`errorCode`, `context`→`details`; add additive `ErrorCode.INTERNAL_ERROR` and `ErrorCode.CONFIRMATION_REQUIRED` with totality maps; **close the `call_tool` leak hatch** so *every* `isError` carries the 6-field envelope (text-JSON only).
- **U6-C — Capability-based routing guidance** (description-only): add routing guidance to `xlvba_environment` phrased as "this server provides capability X"; never detect/claim other servers.
- **U6-D — Description boilerplate compression** (description-only + glossary): move the repeated 133-char guard sentence and the approval clause to `xlvba://glossary` tokens; keep all safety-critical text.
- **U6-E — Semantics docs + confirm flags** (doc + runtime + schema): document the `WorkbookInfo` two-path model and the cross-scope name-conflict rules; add agent-intent `confirm` flags to `snapshot_restore` and `excel_recover` that reject **before** any destructive action.
- **U7 — Save As new capability** (new-capability + schema-additive + runtime): one tool `excel_save_as(workbook_id, path, *, overwrite_existing=False, overwrite_source=False)` with source-identity capture at open, re-hash + normcase/realpath compare, atomic temp→replace, single `.bak`, fail-closed `outcome`.

The two non-negotiables the Validator flagged as safety-critical are first-class acceptance gates: (a) the error-envelope test surface (`test_mcp_error_surfacing.py`) is updated **in the same change** as the hatch closure (U6-B); (b) Save As captures `source_sha256`+size **at open, before materialize** and re-verifies at write-back (U7).

---

## 2. Verified baseline (read at HEAD `40c8b4d`)

- **Tool surface:** `list_tools()` = **92**; descriptions **37,125** chars; `tools/list` JSON **481,641** chars; annotations **92/92**; static resource `xlvba://glossary` (14 terms) + 4 per-workbook resource templates. Artifact: `.ai/tmp/remaining_tools_baseline.json`.
- **Runtime:** Python 3.14.7, mcp 2.3.0, pydantic 2.14.4.
- **Prior remediation:** U1–U5 accepted (HEAD `81bcf2c` + `40c8b4d`).
- **Error envelope today:** `ExcelVbaMcpError.to_dict()` (`domain/errors.py:335-353`) emits 8 keys `code, message, hresult, hresult_hex, context, errorCategory, isRetryable, suggestedNextActions`; `ErrorCode` = 29 members (`errors.py:28-67`); `category_for`/`retryable_for`/`actions_for` are total (`errors.py:249-299`), proven by `tests/unit/test_error_envelope.py:47-62`.
- **Boundary hatch:** `ExcelVbaMcpServer.call_tool` (`server.py:2484-2502`) re-raises bare at `:2500-2501` when neither a domain cause (`_find_domain_cause` `:2383-2398`) nor a validation cause (`_validation_error_as_domain` `:2432-2461`) is found → SDK `UnexpectedToolError("Error executing tool <name>")` (proven by `test_mcp_error_surfacing.py:75-80`).
- **SDK:** text-only `ToolError`/`UnexpectedToolError` (`.venv/.../mcp/server/mcpserver/exceptions.py`); no `structuredContent` error channel — the text-JSON envelope at `call_tool` is the only mechanism.
- **Save As lifecycle:** `_SessionRecord` (`session_service.py:83-94`) has **no** source hash/stat/backup/target field; `save_workbook` (`:474-538`) has no path arg and calls `wb.Save()` in place (`:512`); `_materialize_working_copy` = `shutil.copy2` (`:1199-1206`); `_find_open_workbook` matches by lowercased `FullName` (`:1236-1252`) so `wb.SaveAs` would repoint the open object and break original-protection; `PathPolicy.validate_path` (`security/path_policy.py:110-185`) canonicalizes (resolve+realpath, prefer realpath) and compares with `normcase` (`_normalize` `:272-273`), performs no writes; atomic primitive `artifact_store._atomic_write` (`storage/artifact_store.py:321-350`) = `mkstemp(dir=target.parent)→write→flush→fsync→os.replace` with `finally` unlink; **no backup mechanism exists** (0 hits for `.bak`/`backup`).
- **Confirm today:** `snapshot_restore(snapshot_id, workbook_id=None)` (`api/test_tools.py:122-143`) and `excel_recover(reason="agent_directed")` (`api/recovery_tools.py:46-69`) have **no** `confirm`; `RecoveryResult.confirmed` (`recovery_tools.py:36`) is an **output** echo, not an input gate. Both in `_ANN_DESTRUCTIVE` (`server.py:1795-1796`).
- **Descriptions:** 96 `_DESC_*` constants; 27 tools use the full 8-part template; the 133-char guard sentence appears on 27 tools (≈3,591 chars); `(full-access default; restricted requires approval)` ×20; `xlvba://glossary` referenced ×7. `test_description_contract.py` hard-codes `_CRITICAL`/`_TRUTH`/`_EXAMPLE_TOOLS`, the full guard sentence (`:110-137`), and `len(tools)==92` (`:84-87`).

---

## 3. Findings + severity

| id | Finding | Severity | Package |
|---|---|---|---|
| F1 | `vba_compile` name misleads (never compiles) | Medium (usability) | U6-A |
| F2 | `job_cancel` name implies interruption (intent-only) | Medium (usability) | U6-A |
| F3 | Duplicate test tools `test_run_suite` + `test_suite_run` | Medium (surface bloat) | U6-A |
| F4 | Envelope keys `code`/`context` ≠ target `errorCode`/`details` | High (contract) | U6-B |
| F5 | Non-domain exceptions leak bare `"Error executing tool"` (no envelope) | **High (safety/contract)** | U6-B |
| F6 | No routing guidance in `xlvba_environment` | Low (usability) | U6-C |
| F7 | Repeated guard/approval boilerplate (~3.6k chars) | Low (token cost) | U6-D |
| F8 | `WorkbookInfo.path` vs `source_path` undocumented; `original_name` a trap | Low (doc) | U6-E |
| F9 | Cross-scope name rejection stricter than Excel, undocumented | Low (doc) | U6-E |
| F10 | `snapshot_restore`/`excel_recover` lack `confirm`; recover desc omits external-side-effect caveat | **High (destructive-action safety)** | U6-E |
| F11 | No Save As; no source-identity capture; no backup; `wb.SaveAs` would break session | **High (data-loss risk)** | U7 |

---

## 4. Non-negotiables (binding)

1. **N1 — Total error envelope.** Every `isError=true` response carries exactly the 6 keys `errorCode, errorCategory, message, isRetryable, suggestedNextActions, details` (plus additive `hresult`/`hresult_hex`). Achieved by **both** the `to_dict` rename **and** closing the `call_tool` hatch with a synthesized `INTERNAL_ERROR` fallback. Text-JSON only; **no SDK change; no error `structuredContent`**.
2. **N2 — Additive schema only.** `SCHEMA_VERSION` stays `1`. New `ErrorCode` members (`INTERNAL_ERROR`, `CONFIRMATION_REQUIRED`) and the `SaveAsResult` DTO are additive; totality maps extended.
3. **N3 — Canonical names, no aliases.** `vba_compile_diagnostics`, `job_request_cancel`, `test_suite_run` (sole test entry). `test_run_suite` removed. `TestAction.VBA_COMPILE` (`test_spec.py:73`) is a **different surface** and is **never** renamed.
4. **N4 — Save As is fail-closed.** Default `overwrite_existing=False`/`overwrite_source=False`; target↔source compared by `normcase`+realpath (never string/`Path` equality); source identity captured at open and re-verified; `outcome="ok"` only after final size+hash verify; interrupted/unprovable ⇒ `outcome="unknown"`/`"error"`, **never** success; **no auto-resume**; session stays bound (`wb.Save()` + pure-FS copy, **never** `wb.SaveAs`).
5. **N5 — Confirm rejects before action.** Missing/false `confirm` on `snapshot_restore`/`excel_recover` raises `CONFIRMATION_REQUIRED` **before** any termination/close; descriptions state discard vs preserve and the external-side-effect caveat.
6. **N6 — Routing is capability-based.** Guidance says "this server provides capability X"; never asserts presence/absence of `excel-mcp_*` or any other server.
7. **N7 — Lockstep test edits.** The `to_dict` rename + hatch closure + `test_mcp_error_surfacing.py` edits land in **one** commit (U6-B). The tool-count constant and `_CRITICAL`/`_TRUTH`/`_EXAMPLE_TOOLS` change in the **same** commit as the rename/removal (U6-A) and the Save As add (U7).
8. **N8 — No lazy loading; no backward-compat shims.**

---

## 5. Final contracts

### 5.1 Tool names (final)

| Old | Final | Notes |
|---|---|---|
| `vba_compile` | `vba_compile_diagnostics` | handler + registration + `_DESC_*` + `api/__init__` + audit labels + tests + docs (EN+RU) |
| `job_cancel` | `job_request_cancel` | handler + registration + `_DESC_*` + tests + docs + benchmark |
| `test_run_suite` | **removed** | typed-spec core inlined/renamed to a private helper; one execution path |
| `test_suite_run` | **kept (canonical)** | superset (mapping + `background`); sole test entry |
| — | `excel_save_as` | **new** (U7) |

`TestAction.VBA_COMPILE = "vba_compile"` (`domain/test_spec.py:73`, dispatched `test_runner.py:622`) is **out of scope** — do not touch it or its tests (`test_domain_models.py:144,184`, `test_test_runner.py:891,928,958`, `test_end_to_end_test_runner.py:870,883`).

### 5.2 Error envelope (final)

`to_dict()` emits (in order): `errorCode, message, hresult, hresult_hex, details, errorCategory, isRetryable, suggestedNextActions`.

- `errorCode` ← `code.value`; `details` ← `context`; `errorCategory`/`isRetryable`/`suggestedNextActions` unchanged pure functions.
- New codes (additive):
  - `INTERNAL_ERROR` — category `internal`, `isRetryable=False`, actions `["inspect server logs; the failure was not a domain error"]`.
  - `CONFIRMATION_REQUIRED` — category `validation`, `isRetryable=False`, actions `["re-call with confirm=true after reviewing what is discarded"]`.
- Totality: `_CATEGORY_BY_CODE`, `_ACTIONS_BY_CODE` extended for both; `_RETRYABLE_CODES` unchanged (neither retryable). `test_error_envelope.py` totality parametrize already covers new members; add explicit per-code assertions.
- Hatch closure (`server.py:2492-2502`): when `cause is None` after `_find_domain_cause` and `_validation_error_as_domain`, build a fallback `ExcelVbaMcpError(code=INTERNAL_ERROR, message=<str(exc) only>, context={"reason":"unhandled_exception","exception": exc.__class__.__name__})` and `raise ToolError(json.dumps(fallback.to_dict(), ensure_ascii=False)) from exc`. **Never** embed the live exception object; reuse the `_jsonable` sanitizer pattern (`server.py:2420-2427`) for any non-primitive.

### 5.3 Routing (final)

`_DESC_ENVIRONMENT` gains a capability sentence, e.g.: *"This server provides: VBA project access, macros, pivot tables, background jobs, recovery, and non-activating compile diagnostics (not compilation)."* Capability source = `EnvironmentCapabilities` (`environment_tools.py:112-124`). No other-server detection. Keep the existing "Do not use when" routing.

### 5.4 Confirm (final)

- `snapshot_restore(snapshot_id, workbook_id=None, *, confirm: bool=False)` and `excel_recover(reason="agent_directed", *, confirm: bool=False)`.
- Guard in the **handler** before delegation: `if not confirm: raise ExcelVbaMcpError("... requires confirm=true ...", code=ErrorCode.CONFIRMATION_REQUIRED, context={"tool": ..., "discard": ...})`.
- `_DESC_SNAPSHOT_RESTORE` keeps "discards the current working copy … does NOT undo external macro side effects" + adds "Requires `confirm=true`."
- `_DESC_EXCEL_RECOVER` adds "does not roll back external side effects of any macro that already ran" + "Requires `confirm=true`."
- `RecoveryResult.confirmed` stays an output echo.

### 5.5 Save As (final)

```
excel_save_as(workbook_id: str, path: str, *,
              overwrite_existing: bool = False,
              overwrite_source: bool = False) -> SaveAsResult
```

- One tool; `excel_save` unchanged. Session stays bound to the working copy.
- Decision matrix (target canonical `T`, source canonical `S`, working copy `W`):

| T==S? | T==W? | `overwrite_source` | `overwrite_existing` | source hash | Result |
|---|---|---|---|---|---|
| no | no | any | false (T absent) | any | **write** (new file) |
| no | no | any | false (T exists) | any | **reject** `PROTECTED_TARGET/target_exists` |
| no | no | any | true (T exists) | any | **write** + `.bak` of old T |
| no | yes | any | any | any | **reject** `PROTECTED_TARGET/target_is_working_copy` |
| yes | no | false | any | any | **reject** `PROTECTED_TARGET/source_protected` |
| yes | no | true | (implied) | match | **write** (source-overwrite) + `.bak` of source |
| yes | no | true | (implied) | mismatch | **reject** `STALE_TARGET/source_modified_since_open` |

- Source-replaced (re-canonical of `S` ≠ stored `S` by normcase) → `STALE_TARGET/source_replaced`.
- Path/extension/symlink/UNC/device → `PATH_NOT_ALLOWED` (from `validate_path`).
- Approval class **MUTATING** on the target canonical (and source when `T==S`); FULL auto-allows, RESTRICTED requires grant → `APPROVAL_DENIED`.
- **Skip `require_snapshot`** (Save As does not mutate the open working copy; state this in the description).
- `SaveAsResult` (additive DTO, `extra="forbid"`, `SCHEMA_VERSION=1`): `workbook_id, target_path, source_path, source_overwritten, overwrote_existing, backup_path (str|None), sha256, size, outcome Literal["ok","unknown","error"], mode, notes, created_at, updated_at`.

---

## 6. Exact source code map

| Concern | File / anchor |
|---|---|
| Rename `vba_compile` | `api/compile_tools.py:63` (def), `:102` (`__all__`); `server.py:55` (import), `:1979-1984` (register, name `:1981`), `:407-421` (`_DESC_VBA_COMPILE`); `api/__init__.py:10,59`; audit labels `application/compile_service.py:198,261,276,352` |
| Rename `job_cancel` | `api/job_tools.py:113` (def), `:133` (`__all__`); `server.py:63` (import), `:2015-2020` (register, name `:2017`), `:489-500` (`_DESC_JOB_CANCEL`); cross-ref `server.py:486` (`_DESC_JOB_GET`) |
| Remove `test_run_suite` | `api/test_tools.py:163` (def), `:229` (`partial`), `:234` (direct call), `:277` (`__all__`); `server.py:144` (import), `:303-319` (`_DESC_TEST_RUN_SUITE`, delete), `:1907-1912` (register, delete); cross-ref `server.py:330` (`_DESC_TEST_SUITE_RUN`); audit label `application/test_runner.py:1763` |
| Error `to_dict` | `domain/errors.py:335-353` (keys `:345` code, `:349` context) |
| Error maps | `domain/errors.py:121-159` (`_CATEGORY_BY_CODE`), `:161-164` (`_RETRYABLE_CODES`), `:182-236` (`_ACTIONS_BY_CODE`); codes enum `:28-67` |
| Boundary hatch | `server.py:2484-2502`; helpers `:2383-2398`, `:2420-2427`, `:2432-2461` |
| Routing | `server.py:1004-1019` (`_DESC_ENVIRONMENT`); `api/environment_tools.py:112-124` (`_capabilities`) |
| Glossary | `server.py:2626-2661` (`_GLOSSARY_MD`, `glossary_resource`); `_DESC_RES_GLOSSARY` `:1020-1025` |
| Guard sentence (compress) | `server.py` `_DESC_*` ×27; test `tests/unit/test_description_contract.py:110-137` |
| Confirm | `api/test_tools.py:122-143` (`snapshot_restore`); `api/recovery_tools.py:46-69` (`excel_recover`), `:36` (`RecoveryResult.confirmed`); descs `server.py:292-301`, `:272-286` |
| Save As | `_SessionRecord` `application/session_service.py:83-94`; `open_workbook` `:223-349` (materialize `:306`); `save_workbook` `:474-538`; `_materialize_working_copy` `:1199-1206`; `_find_open_workbook` `:1236-1252`; `_record_audit` `:1215-1233`; `PathPolicy.validate_path` `security/path_policy.py:110-185`, `_normalize` `:272-273`; atomic `storage/artifact_store.py:321-350`; handler `api/workbook_tools.py` (add `excel_save_as`); DTO `domain/models.py` (add `SaveAsResult`, `SCHEMA_VERSION` `:65`) |
| Description-contract test | `tests/unit/test_description_contract.py` (`_CRITICAL` `:18-44`, `_EXAMPLE_TOOLS` `:47-52`, `_TRUTH` `:55-59`, count `:84-87`) |
| Error tests | `tests/unit/test_error_envelope.py:30-31,41-42`; `tests/unit/test_mcp_error_surfacing.py:71-72,110-115` + rewrite `:75-80` |
| Benchmark | `.ai/tmp/benchmark_tasks.yaml:471-477` (`job_cancel`), `:554-570` (`test_suite_run`/`must_not` `:569`), `:573-588` (T28 save-export-gap) |

---

## 7. Work-package DAG

### 7.1 Graph

```
            U6-B (error envelope: errors.py + hatch + error tests)
           /      \
     (unblocks)   (unblocks)
          v          v
   U6-E (confirm)   U7 (Save As)  ──(count 91→92)──┐
          ^   ^          ^                          │
          |   \         |                          |
   U6-A (renames)      |                           |
          ^            |                           |
          |            |                           |
   U6-C (routing)      |                           |
          ^            |                           |
          \___________/                            |
                 \                               /
                  v                             v
                     U6-D (description compression)  [last]
```

### 7.2 Ordering table

| Package | blocked_by | Rationale |
|---|---|---|
| U6-B | — (root) | Foundational; unblocks U6-E (`CONFIRMATION_REQUIRED`) and U7 (envelope consistency). |
| U6-A | U6-B | Shares `server.py`; B's small hatch edit lands first. Sets tool count 92→91. |
| U6-C | U6-A | `server.py` description edit; avoid textual conflict. |
| U6-E | U6-B, U6-A | Needs `CONFIRMATION_REQUIRED`; `server.py` desc edits after A. |
| U7 | U6-B, U6-A | Needs envelope; count 91→92 coordinated with A. |
| U6-D | U6-A, U6-C, U6-E, U7 | Compresses the guard sentence across all `_DESC_*` after every other `server.py` desc/registration edit has landed. |

**Shared-file note:** `server.py` is the serialization hotspot. Land server.py-touching packages in the order A → C → E → U7 → D. U6-B's `errors.py`/test work may proceed in parallel with U6-A's `api/*` renames up to the `server.py` merge. Each package is independently green (the count constant is updated per step: 91 after A, 92 after U7).

---

## 8. Work packages

### U6-A — Canonical tool naming & consolidation

- **Class:** tool-lifecycle + description + tests + benchmark.
- **Scope:**
  - Rename handler `vba_compile`→`vba_compile_diagnostics` (`compile_tools.py:63`), `__all__` `:102`, import `server.py:55`, registration `server.py:1979-1984` (name `:1981`), `_DESC_VBA_COMPILE` `:407-421` (constant rename to `_DESC_VBA_COMPILE_DIAGNOSTICS`; keep "does NOT compile"), `api/__init__.py:10,59`, audit labels `compile_service.py:198,261,276,352`.
  - Rename handler `job_cancel`→`job_request_cancel` (`job_tools.py:113`), `__all__` `:133`, import `server.py:63`, registration `server.py:2015-2020` (name `:2017`), `_DESC_JOB_CANCEL` `:489-500` (constant rename; keep "intent-only; never interrupts"), cross-ref `server.py:486` (`_DESC_JOB_GET` "job_cancel"→"job_request_cancel").
  - Remove `test_run_suite`: delete def `test_tools.py:163`, inline/rename typed-spec core to a private `_run_suite_blocking(spec)` (or inline into `test_suite_run`), repoint `partial(...)` `:229` and direct call `:234`, delete `_DESC_TEST_RUN_SUITE` `server.py:303-319`, delete registration `server.py:1907-1912`, delete import `server.py:144`, update `__all__` `test_tools.py:277`. Rewrite `_DESC_TEST_SUITE_RUN` `:330` "Do not use when … use test_run_suite" (obsolete cross-ref removed; `test_suite_run` is sole entry). Audit label `test_runner.py:1763` "test_run_suite"→"test_suite_run".
  - Update `test_description_contract.py`: `_CRITICAL` (rename `job_cancel`→`job_request_cancel`, `vba_compile`→`vba_compile_diagnostics`, remove `test_run_suite`; keep `test_suite_run`), `_TRUTH` (rename keys, values unchanged), count `92`→`91`.
  - Update unit tests: `test_compile_service.py:391,517`, `test_job_tools.py:20,122-165`, `test_job_registry.py:11,160-162`, `test_macro_policy.py:92` (if touched), `test_mcp_error_surfacing_matrix.py` `test_run_suite` refs, `test_description_contract.py`.
  - Update docs (EN + RU): `docs/01,04,05,06,07,09,10,11` per research §1.4.
  - Update benchmark `.ai/tmp/benchmark_tasks.yaml:471-477` (`job_cancel`→`job_request_cancel`), `:569` (`must_not` remove `test_run_suite` wording).
- **Non-scope:** `TestAction.VBA_COMPILE` (`test_spec.py:73`) and its tests; `macro_policy._EXECUTING_TOOLS` stale set (documented non-wired; leave as-is — optional cleanup only, not required); no alias.
- **Changed files:** `api/compile_tools.py`, `api/job_tools.py`, `api/test_tools.py`, `api/__init__.py`, `server.py`, `application/compile_service.py`, `application/test_runner.py`, `tests/unit/test_compile_service.py`, `tests/unit/test_job_tools.py`, `tests/unit/test_job_registry.py`, `tests/unit/test_description_contract.py`, `tests/integration_excel/test_mcp_error_surfacing_matrix.py`, `docs/**` (EN+RU), `.ai/tmp/benchmark_tasks.yaml`.
- **Acceptance:** `list_tools()` contains `vba_compile_diagnostics`, `job_request_cancel`, `test_suite_run`; **not** `vba_compile`, `job_cancel`, `test_run_suite`. `test_suite_run` handles both mapping and typed spec and `background`. `_DESC_*` truthful phrases intact. Count test = 91.
- **Tests:** `pytest tests/unit/test_description_contract.py tests/unit/test_compile_service.py tests/unit/test_job_tools.py tests/unit/test_job_registry.py`; grep confirms zero `test_run_suite`/`job_cancel`/`vba_compile` tool-name residues in `src` (excluding `TestAction.VBA_COMPILE` and stale `_EXECUTING_TOOLS`).
- **Compat:** backward compat not required; removals are intentional. `SCHEMA_VERSION` unchanged.
- **Rollback:** revert the U6-A commit; count returns to 92. No data migration.

### U6-B — Guaranteed structured error envelope

- **Class:** schema-additive (ErrorCode) + runtime (boundary) + tests.
- **Scope:**
  - `domain/errors.py`: add `INTERNAL_ERROR`, `CONFIRMATION_REQUIRED` to `ErrorCode` (`:28-67` block); extend `_CATEGORY_BY_CODE` (`internal`/`validation`), `_ACTIONS_BY_CODE`; leave `_RETRYABLE_CODES` unchanged. Rename `to_dict` keys `code`→`errorCode` (`:345`), `context`→`details` (`:349`); keep `hresult`/`hresult_hex` additive.
  - `server.py`: close hatch `:2500-2501` — synthesize `INTERNAL_ERROR` fallback envelope from `str(exc)` only (reuse `_jsonable`); `raise ToolError(json.dumps(fallback.to_dict(), ensure_ascii=False)) from exc`.
  - Tests (same commit): `test_error_envelope.py:30-31,41-42` (`_ORIGINAL_KEYS` code→errorCode, context→details) + explicit `INTERNAL_ERROR`/`CONFIRMATION_REQUIRED` category/retryable assertions; `test_mcp_error_surfacing.py:71-72,110-115` (key rename) **and rewrite `:75-80`** to assert a structured `INTERNAL_ERROR` envelope with all 6 keys (no `UnexpectedToolError` leak, no `"Error executing tool"`).
- **Non-scope:** SDK change; `structuredContent`; `SCHEMA_VERSION` bump; `_EXECUTING_TOOLS`.
- **Changed files:** `domain/errors.py`, `server.py`, `tests/unit/test_error_envelope.py`, `tests/unit/test_mcp_error_surfacing.py`.
- **Acceptance:** every `isError` path (domain, validation, non-domain) yields the 6 keys; `INTERNAL_ERROR` for the non-domain hatch; totality test green over 31 codes; no bare `"Error executing tool"` reaches clients.
- **Tests:** `pytest tests/unit/test_error_envelope.py tests/unit/test_mcp_error_surfacing.py`.
- **Compat:** key rename is breaking-by-design (compat not required). Additive codes; `SCHEMA_VERSION=1`.
- **Rollback:** revert U6-B commit (hatch + keys + codes + tests together). U6-E/U7 depend on it; do not revert in isolation after they land.

### U6-C — Capability-based routing guidance

- **Class:** description-only.
- **Scope:** extend `_DESC_ENVIRONMENT` (`server.py:1004-1019`) with a capability sentence sourced from `EnvironmentCapabilities` (`environment_tools.py:112-124`). No other-server detection. Keep "Do not use when" routing.
- **Non-scope:** no new tool; no capability logic change; no other-server probing.
- **Changed files:** `server.py` (`_DESC_ENVIRONMENT`), `tests/unit/test_description_contract.py` (if a routing assertion is added).
- **Acceptance:** description states "this server provides …"; contains no `excel-mcp_*`/other-server claim; `test_description_contract.py::test_environment_and_glossary_constants_exist` still passes.
- **Tests:** `pytest tests/unit/test_description_contract.py`.
- **Compat:** additive text; `SCHEMA_VERSION` unchanged.
- **Rollback:** revert the `_DESC_ENVIRONMENT` edit.

### U6-D — Description boilerplate compression

- **Class:** description-only + glossary.
- **Scope:**
  - Add glossary terms **guarded mutation** and **approval gate** to `_GLOSSARY_MD` (`server.py:2626-2649`); add them to `_GLOSSARY_TERMS` in the test (`:61-76`).
  - Replace the 133-char guard sentence on the 27 guarded-mutation tools with the token `"Guarded mutation (see xlvba://glossary)."`; replace `(full-access default; restricted requires approval)` with the glossary token.
  - Update `test_description_contract.py:110-137` to assert the token (not the full sentence).
  - **Keep** all `_TRUTH` phrases, per-tool "Do not use when" routing, tool-specific preconditions (e.g. `range_write_values` "does NOT snapshot"), and the one `Example:` per `_EXAMPLE_TOOLS`.
- **Non-scope:** no safety-text removal; no per-tool routing rewrite; no lazy loading.
- **Changed files:** `server.py` (`_DESC_*` ×27, `_GLOSSARY_MD`, `_DESC_RES_GLOSSARY`), `tests/unit/test_description_contract.py`.
- **Acceptance:** description chars measurably reduced (target ≈ −3,000); glossary contains the two new terms; all `_TRUTH`/routing/example assertions green; guard-token assertion green.
- **Tests:** `pytest tests/unit/test_description_contract.py`; char-count delta recorded.
- **Compat:** text-only; `SCHEMA_VERSION` unchanged.
- **Rollback:** revert U6-D commit (restores full sentences).

### U6-E — Semantics docs + confirm flags

- **Class:** doc (WorkbookInfo, Names) + runtime + schema (confirm).
- **Scope:**
  - **E-A (doc):** document `WorkbookInfo.path` (working copy) vs `source_path` (original, nullable) and that `original_name` is intentionally absent (redundant with `name`) — in `docs/05`/`docs/09` and the `WorkbookInfo` docstring (`models.py:86-99`). No code change.
  - **E-B (doc):** document cross-scope name conflict (`names_adapter.py:576-618`), update-preserves-scope (`:439-470`), resolution order (`:535-574`) with the reject examples — in `docs/05`. No code change.
  - **E-C (runtime):** add `confirm: bool=False` to `snapshot_restore` (`test_tools.py:122-143`) and `excel_recover` (`recovery_tools.py:46-69`); guard `if not confirm: raise ExcelVbaMcpError(..., code=CONFIRMATION_REQUIRED, ...)` **before** `service.restore_snapshot` / `session.recover_owned_process`. Update `_DESC_SNAPSHOT_RESTORE` (`server.py:292-301`) and `_DESC_EXCEL_RECOVER` (`:272-286`) with discard/preserve + external-side-effect caveat + "Requires `confirm=true`."
- **Non-scope:** gating `excel_close(save_changes=true)`/`excel_save` (future); `RecoveryResult.confirmed` semantics (stays output echo); relaxing name strictness.
- **Changed files:** `api/test_tools.py`, `api/recovery_tools.py`, `server.py` (descs), `domain/models.py` (docstring only), `docs/05`, `docs/09`, `tests/unit/test_description_contract.py` (if `_TRUTH`/`_CRITICAL` touched), new unit tests for confirm rejection.
- **Acceptance:** missing/false `confirm` → `CONFIRMATION_REQUIRED` with **no** termination/close (assert via fake service that the delegate is not called); descriptions state discard vs preserve + caveat; WorkbookInfo/Names docs present.
- **Tests:** new `tests/unit/test_confirm_flags.py` (reject-before-action with fakes); `pytest tests/unit/test_description_contract.py`.
- **Compat:** additive param (default `False`); `SCHEMA_VERSION=1`.
- **Rollback:** revert E-C (param + guards + descs); docs revert independently.

### U7 — Save As (new capability)

- **Class:** new-capability + schema-additive (SaveAsResult) + runtime.
- **Scope:**
  - **Source identity:** add `source_sha256: str`, `source_size: int` to `_SessionRecord` (`session_service.py:83-94`); capture in `open_workbook` **after** `validate_path` and **before** `_materialize_working_copy` (`:306`) via a streaming SHA-256 (reuse `snapshot_store._sha256_file` shape) + `stat().st_size`.
  - **Service:** add `SessionService.save_as_workbook(workbook_id, path, *, overwrite_existing=False, overwrite_source=False, timeout_s=...)` implementing §9 algorithm; `wb.Save()` on STA (flush working copy) then pure-FS atomic copy on the event loop; **never** `wb.SaveAs`. Audit `tool="excel_save_as"` with `params={workbook_id, target_path, overwrite_existing, overwrite_source}`, `result={target_path, sha256, size, backup_path, source_overwritten, outcome}`.
  - **Handler:** add `excel_save_as` to `api/workbook_tools.py` (thin, no COM) + `__all__`; export via `api/__init__.py`.
  - **DTO:** add `SaveAsResult` to `domain/models.py` (additive, `extra="forbid"`); `SCHEMA_VERSION` stays 1.
  - **Registration:** `server.py` import + `add_tool(excel_save_as, name="excel_save_as", description=_DESC_SAVE_AS, annotations=_annotations("excel_save_as"))` (mutating; not read-only/idempotent; destructive off — rely on flags). Add `_DESC_SAVE_AS`. Update `_DESC_SAVE` (`:246-255`) "there is no save-as" → reference `excel_save_as`. Add `excel_save_as` to `_CRITICAL`; count `91`→`92`.
  - **Tests:** 14-case matrix (research §9) as `tests/unit/test_save_as.py` with fake STA + real temp FS; optional `tests/integration_excel` round-trip (L3).
  - **Benchmark:** update T28 (`:573-588`) to use `excel_save_as` (new path, source-overwrite, rejection, external-change variants); add tasks per §10.
- **Non-scope:** `export_fixed_format`/PDF (PathPolicy rejects `.pdf/.xps`); `vba_export_project` (unchanged); session rebinding; target-keyed concurrency lock (note race, out of scope); `destructive_hint` (static per name; leave off).
- **Changed files:** `application/session_service.py`, `api/workbook_tools.py`, `api/__init__.py`, `domain/models.py`, `server.py`, `tests/unit/test_save_as.py`, `tests/unit/test_description_contract.py`, `.ai/tmp/benchmark_tasks.yaml`, `docs/05`, `docs/09`.
- **Acceptance:** all 14 matrix cases pass; `outcome="ok"` only after final size+hash verify; source-overwrite only with `overwrite_source=True` + hash match; existing target only with `overwrite_existing=True`; interrupted ⇒ `unknown` + `backup_path`, no auto-resume; session `workbook_id`/`working_path`/`canonical` unchanged after Save As; later `excel_save` writes the working copy (target hash unchanged).
- **Tests:** `pytest tests/unit/test_save_as.py tests/unit/test_description_contract.py`.
- **Compat:** new tool + additive DTO; `SCHEMA_VERSION=1`.
- **Rollback:** revert U7 commit (tool + DTO + service + session fields + count). Because U6-D depends on U7, land U6-D after U7 or revert both together.

---

## 9. Save As safety model (detail)

**Path gate:** `canonical_t = path_policy.validate_path(path, MUTATING)` (allowlist + extension + symlink/UNC/device). Reject `normcase(canonical_t)==normcase(record.working_path)` → `PROTECTED_TARGET/target_is_working_copy`.

**Identity:** at open, `record.source_sha256 = sha256(canonical)`, `record.source_size = stat.st_size` (before materialize). At Save As: re-`validate_path(record.canonical, READ_ONLY)`; if `normcase(re_canonical) != normcase(record.canonical)` → `STALE_TARGET/source_replaced`; re-hash source; if `!= record.source_sha256` → `STALE_TARGET/source_modified_since_open`.

**Overwrite decision:** `same = normcase(canonical_t)==normcase(record.canonical)`; `exists = canonical_t.exists()`; apply §5.5 matrix (`PROTECTED_TARGET` for source_protected/target_exists; `STALE_TARGET` for drift).

**Atomic + backup (single `<target>.bak`, overwritten each save):**
1. `wb.Save()` on STA (flush working copy in place).
2. Stage temp on target FS: `mkstemp(prefix=f".{name}.", suffix=".tmp", dir=canonical_t.parent)`; copy working→tmp; `flush`+`fsync`; verify staged size==working size and staged sha==working sha (else close+unlink temp, destination untouched, typed error).
3. If `exists`: copy current `canonical_t` → `backup_path = f"{canonical_t}.bak"` (overwrite any prior `.bak`), `flush`+`fsync`.
4. `os.replace(tmp, canonical_t)` (same volume → atomic); `tmp=None`.
5. Verify final size+sha; only then `outcome="ok"`.
6. `finally` unlink leftover temp.

**Interruption:** if the process dies after step 3/4 boundary (destination absent + `.bak` present) → `outcome="unknown"` with `backup_path` and a suggested action; **no auto-resume**; a retry must not orphan a second backup (detect existing `.bak` for the same target and stop). Never report `ok` from "no exception". Final-verify mismatch → attempt best-effort rollback from `.bak`, report `unknown`/`error`.

**Session binding:** `record.workbook_id`/`working_path`/`canonical` unchanged; later `excel_save` writes the working copy, not the target. `wb.SaveAs` is rejected (rebinds `FullName`, breaks `_find_open_workbook` and original-protection).

---

## 10. Benchmark task additions

Update existing:
- `T22` (`:465-478`): `job_cancel`→`job_request_cancel` (all `expected_*`/`token_measure`).
- `T27` (`:548-571`): `must_not` remove `test_run_suite` wording (sole entry `test_suite_run`).
- `T28` (`:573-588`): replace "NO save_as exists" gap with real `excel_save_as` usage (new path); adjust `expected_sequence`/`success_criteria`/`must_not`.

Add:
- Renamed-tool task for `vba_compile_diagnostics` (asserts "does NOT compile").
- Canonical-test-tool task (mapping + `background`; `must_not` arbitrary code).
- Structured-error task (non-domain failure surfaces 6-key envelope; `errorCode`).
- WorkbookInfo task (`path` vs `source_path`; no `original_name`).
- Name-conflict task (duplicate vs scope_conflict rejects).
- Confirm-flag task (missing `confirm` → `CONFIRMATION_REQUIRED`, no termination).
- Save As tasks: new path; existing target + `overwrite_existing`; source-overwrite + `overwrite_source` + hash match; source drift → `STALE_TARGET`; target==working copy → `PROTECTED_TARGET`; interrupted → `unknown` + `backup_path`.

---

## 11. Rollback (global)

- Each package is a single revertible commit; `server.py` serialization order A→C→E→U7→D means a mid-chain revert must also revert downstream server.py edits (or re-apply them).
- U6-B is foundational: do not revert in isolation after U6-E/U7 land (they depend on `CONFIRMATION_REQUIRED`/envelope).
- No data migration, no schema bump, no workbook/artifact format change → rollback is code-only.
- The `==92`/`==91` count constant is the only cross-package coupling; it is updated per step so each commit is green.

---

## 12. Change-class separation (summary)

| Class | Packages |
|---|---|
| Description-only | U6-C, U6-D (and the doc parts of U6-E A/B) |
| Schema (additive) | U6-B (`INTERNAL_ERROR`, `CONFIRMATION_REQUIRED`), U7 (`SaveAsResult`) |
| Runtime | U6-B (hatch), U6-E-C (confirm guards), U7 (save-as service) |
| Tool-lifecycle | U6-A (renames/removal), U7 (new `excel_save_as` registration) |
| New-capability | U7 (Save As) |

**Glossary:** all jargon externalized here is defined in `xlvba://glossary` (`server.py:2626-2661`), extended in U6-D with **guarded mutation** and **approval gate**. Descriptions reference the URI, never re-explain jargon.