# report_10f — Final Auditor Spot-Check (Remaining-Usability Remediation)

- **Role:** Auditor (source-to-contract final spot-check). AUDIT ONLY — no files modified.
- **Repo:** `C:\py_exp\xlvba_mcp`
- **Working tree:** baseline `40c8b4d` + uncommitted U6-A..E + U7 + U6-D + alignment.
- **Date:** 2026-11-26
- **Method:** static source reading + `capture_baseline.py` + `pytest` (unit run, integration collect). No Excel/COM started; no production code changed.

## PASS/FAIL summary

| # | Check | Verdict |
|---|-------|---------|
| 1 | Tool surface = 92; required names present; forbidden names absent; no aliases | **PASS** |
| 2 | Error envelope: `to_dict` keys; `call_tool` hatch → structured `INTERNAL_ERROR`; every isError path has the 6 keys | **PASS** |
| 3 | `snapshot_restore` / `excel_recover` reject missing/false `confirm` before the delegate | **PASS** |
| 4 | Save As: hash-at-open, normcase/realpath, decision matrix, atomic temp→replace, single non-overwritten `.bak`, `outcome="ok"` only after verify, no `wb.SaveAs`, session unchanged, `_find_open_workbook` intact | **PASS** |
| 5 | Routing capability-based (no other-server detection); description compression kept `_TRUTH`/routing/examples; glossary has guarded-mutation + approval-gate | **PASS** |
| 6 | Non-negotiables: no SDK change; no error `structuredContent`; `SCHEMA_VERSION=1`; only 2 new ErrorCodes; single-STA/Job Object/working-copy unchanged; no lazy loading | **PASS** |
| 7 | `pytest tests/unit` green; integration collects | **PASS** |
| 8 | No stray production edits outside the packages | **PASS** |

**Blocking inconsistencies:** none.

## Evidence

### Item 1 — Tool surface (PASS)
- `& ".venv\Scripts\python.exe" .ai\tmp\capture_baseline.py` → `TOOL_COUNT=92`, `TOOLS_WITH_ANNOTATIONS=92`, `TOOLS_WITH_OUTPUT_SCHEMA=92`.
- Parsed `.ai/tmp/tools_list_baseline.json` (`tools` array, 92 entries):
  - present: `vba_compile_diagnostics`, `job_request_cancel`, `test_suite_run`, `excel_save_as` → all `True`.
  - absent: `vba_compile`, `job_cancel`, `test_run_suite` → all `False`.
  - duplicates: none; alias-like names (`*_v2`/`_old`/`_alias`/`legacy*`): none.
- Registration is explicit, one `add_tool` per name: `server.py:1784-2319` (e.g. `excel_save_as` `:1798-1803`, `vba_compile_diagnostics` `:1918-1923`, `job_request_cancel` `:1954-1959`, `test_suite_run` `:1852-1857`). No alias/duplicate registrations.
- Contract test `test_description_contract.py:86-89` asserts `len(tools) == 92`.

### Item 2 — Error envelope (PASS)
- `to_dict` emits the fixed 8-key order `errorCode, message, hresult, hresult_hex, details, errorCategory, isRetryable, suggestedNextActions` — `errors.py:352-373` (return `:364-373`). The 6 required keys (`errorCode, errorCategory, message, isRetryable, suggestedNextActions, details`) are always present; `hresult`/`hresult_hex` are additive/nullable. `errorCode ← code.value` `:365`, `details ← context` `:369`.
- `call_tool` hatch is closed: domain cause → JSON `ToolError` (`server.py:2481`); pydantic `ValidationError` → translated domain (`server.py:2459`, `_validation_error_as_domain` `:2390-2419`); **non-domain fallback** synthesizes `ExcelVbaMcpError(code=INTERNAL_ERROR)` from the originating cause and raises JSON (`server.py:2460-2480`, `origin=_originating_cause` `:2469`, `INTERNAL_ERROR` `:2472`, `raise ToolError(json.dumps(...))` `:2478-2480`). No bare re-raise; no "Error executing tool" text reaches the client.
- Only two `raise ToolError` exist in `src`, both JSON (`server.py:2478`, `:2481`); grep found **no** tool-side `CallToolResult(`/`is_error`/`isError` producer that could bypass the envelope.
- `_originating_cause` is cycle-safe (`server.py:2340-2356`); `_find_domain_cause` cycle-safe (`:2322-2337`).
- Tests: `test_error_envelope.py:34-55` (key set + fixed order), `:113-120` INTERNAL_ERROR internal/non-retryable, `:123-130` CONFIRMATION_REQUIRED validation/non-retryable, `:58-73` totality over `list(ErrorCode)`.

### Item 3 — Confirm gating before delegate (PASS)
- `excel_recover`: `if not confirm: raise ... CONFIRMATION_REQUIRED` at `recovery_tools.py:81-94`, **before** `get_session_service()` `:95` and `recover_owned_process` `:96`.
- `snapshot_restore`: `if not confirm: raise ... CONFIRMATION_REQUIRED` at `test_tools.py:150-163`, **before** `get_snapshot_service()` `:164` and `restore_snapshot` `:165`.
- Both use `ErrorCode.CONFIRMATION_REQUIRED` (added `errors.py:75`), category `validation`, non-retryable (`errors.py:167`, `:250-252`).

### Item 4 — Save As (PASS)
- **Source hash captured at open before materialize:** `source_sha256, source_size = _sha256_file(canonical)` at `session_service.py:336`, then `working_path = self._materialize_working_copy(...)` `:337`; stored on the record `:366-367`.
- **normcase / realpath compare:** target canonical from `validate_path` (realpath/resolve canonicalization, `path_policy.py:248-260`); comparisons use `os.path.normcase` at `session_service.py:630` (target==working copy), `:648` (source canonical unchanged), `:668` (target==source).
- **Decision matrix** (`session_service.py:671-681`): source target without `overwrite_source` → `ProtectedTargetError` `:672-676`; existing non-source target without `overwrite_existing` → `ProtectedTargetError` `:677-681`; working-copy target rejected `:630-637`.
- **Atomic temp→replace:** `_atomic_save_as_copy` `:754-869` — `mkstemp` `:783`, copy+fsync `:788-793`, staged verify `:795-804`, `os.replace(temp_path, canonical_t)` `:828`, final verify `:831-842`, temp cleanup `finally` `:863-869`.
- **Single `.bak` NOT overwritten if pre-existing → unknown:** `:806-823` — if `<target>.bak` already exists, returns `outcome="unknown"` with `backup_path`, destination untouched, no auto-resume.
- **`outcome="ok"` only after final verify:** `:843-850` (ok returned only after size+sha match); any `OSError` after backup boundary → `"unknown"`, before → `"error"` `:851-862`; staged mismatch → `"error"` `:798-804`.
- **NO `wb.SaveAs`:** grep over `src` finds only doc/comment references (`server.py:268`, `session_service.py:591,683`, `models.py:1493`); the flush uses `wb.Save()` at `session_service.py:742` (`_flush_working_copy` `:729-752`).
- **Session unchanged:** `save_as_workbook` returns `SaveAsResult` (`models.py:1485-1523`) and never rebinds `workbook_id`/`working_path`/`canonical` (docstring `:590-594`; result build `:693-707`).
- **`_find_open_workbook` still resolves:** unchanged, matches by `working_path` (`session_service.py:1601-1617`), used by flush `:735` and save `:538`.

### Item 5 — Routing / description compression / glossary (PASS)
- `_DESC_ENVIRONMENT` is capability-based and explicitly states no other-server detection: "Routing is capability-based advice only: this server does NOT detect whether any other MCP server is installed or connected." `server.py:945-946`; capability list `:941-944`; routing matrix `:953-957`; examples `:956-957`.
- Compression kept the non-negotiable truths: `_TRUTH` phrases present in descriptions (`vba_compile_diagnostics` "does NOT compile" `:408-422`; `macro_classify_background` "never grants run permission" `:456-473`; `job_request_cancel` "intent-only; never interrupts" `:490-501`); examples kept for the 4 example tools (`_DESC_RANGE_WRITE_VALUES` `:636-637`, `_DESC_TABLE_UPDATE_ROWS` `:572-574`, `_DESC_VBA_REPLACE_COMPONENT_SOURCE` `:393-395`, `_DESC_TEST_SUITE_RUN` `:351-354`).
- Uniform guard token "Guarded mutation (see xlvba://glossary)." applied across guarded mutations (e.g. `:517,529,660,666,684,690,753,764,776,898`).
- Glossary (`server.py:2607-2632`) includes **guarded mutation** `:2628-2629` and **approval gate** `:2630-2631` (plus R13/BG06/STA/indeterminate/working copy/snapshot gate).
- Contract tests: `test_description_contract.py:54-59` (_TRUTH), `:46-52` (examples), `:112-137` (guard token), `:140-147` (glossary terms incl. guarded mutation/approval gate), `:150-158` (environment read-only).

### Item 6 — Non-negotiables (PASS)
- **No SDK change:** `git diff 40c8b4d -- pyproject.toml uv.lock` empty; no `.venv/.../mcp` diff.
- **No error `structuredContent`:** grep for `structuredContent`/`structured_content` in `src` → none.
- **`SCHEMA_VERSION=1`:** `models.py:65` (used `:313,769,793`); snapshot manifest `model_version=2` is a separate pre-existing constant (`snapshot_store.py:64`), not the Stage-1 DTO version.
- **No new ErrorCode beyond the two:** `ErrorCode` count = 31 (baseline 29 + 2). Diff of `errors.py` adds only `INTERNAL_ERROR` `:74` and `CONFIRMATION_REQUIRED` `:75` (plus their category/actions entries).
- **Single-STA / Job Object / working-copy unchanged:** `com_worker.py` and job-object modules are **not** in the diff; the single-STA FIFO submit is still used (`session_service.py:747-752`); job_registry change is a docstring rename only (`job_cancel`→`job_request_cancel`).
- **No lazy loading:** tools registered eagerly at import (`server.py:2503 _register_tools(mcp)`); services wired in `lifespan` (`:1270-1284`).

### Item 7 — Tests (PASS)
- `pytest tests/unit` → **1152 passed in 29.28s** (0 failed, 0 errors).
- `pytest tests/integration_excel --collect-only` → **384 tests collected in 1.80s**, no collection errors.

### Item 8 — No stray production edits (PASS)
- All production (non-test, non-doc) changes are inside `src/excel_vba_mcp/` (13 files): `api/{__init__,compile_tools,job_tools,recovery_tools,test_tools,workbook_tools}.py`, `application/{compile_service,job_registry,session_service,test_runner}.py`, `domain/{errors,models}.py`, `server.py`.
- `com_worker.py` / job-object / storage / security modules untouched. `pyproject.toml` / `uv.lock` untouched. Remaining changes are tests and docs only.

## Blocking inconsistencies
None. All eight spot-checks pass with cited source evidence; the two new error codes are additive and non-retryable; the Save As path is fail-closed and never rebinds the session; the tool surface is exactly 92 with the renamed/added tools and no aliases.

## Final recommendation
**ACCEPT** the remaining-usability remediation (U6-A..E + U7 + U6-D + alignment). The source-to-contract spot-check is clean: tool surface, error envelope, confirm gating, Save As, capability-based routing, and all non-negotiables hold, with unit tests green and integration collecting. No changes required.