# Remaining Usability Baseline (plan_05 cycle)

- **Git HEAD:** `40c8b4d7a1e5a7406bbd9f074319e112711e60fc` (`docs(xlvba): agent usability audit, plan, and results`).
- **Working tree:** clean except pre-existing untracked `plan_01`/`plan_02` and three untracked files from a prior session (`agent_usability_phase0_environment.md`, `research_eval_1_benchmark_methodology.md`, `research_eval_2_measurement_design.md`) — not part of this task.
- **Runtime:** Python 3.14.7, mcp 2.3.0, pydantic 2.14.4.
- **Prior remediation:** U1–U5 accepted (HEAD `81bcf2c` + `40c8b4d`).

## 1. Tool surface (current)

- `list_tools()` = **92**; `tools/list` JSON = **481,641 chars**.
- Descriptions **37,125** chars; input schema **82,373**; output schema **350,758** (72.8%); annotations **92/92**.
- Static resource `xlvba://glossary` (14 terms) + 4 per-workbook resource templates.
- Artifact: `.ai/tmp/remaining_tools_baseline.json`.

## 2. Rename / duplicate targets (U6-A)

| Current | Location | Required canonical |
|---|---|---|
| `vba_compile` | `api/compile_tools.py:63`; registered `server.py:1981` | `vba_compile_diagnostics` |
| `job_cancel` | `api/job_tools.py:113`; registered `server.py:2017` | `job_request_cancel` |
| `test_run_suite` | `api/test_tools.py:163`; registered `server.py:1909` | REMOVE (keep `test_suite_run`) |
| `test_suite_run` | `api/test_tools.py:189`; registered `server.py:1915` | KEEP as canonical |

- `test_suite_run` is the superset: accepts `TestSuiteSpec | dict` + `background`; `test_run_suite` accepts only a typed spec and blocks. `test_suite_run` internally calls `test_run_suite` and `start_long_op(partial(test_run_suite, spec))` — consolidation must keep one request schema + one execution contract.
- No backward-compat aliases (approved). Update registration, schemas, internal refs, docs, examples, tests, benchmark tasks.

## 3. Error envelope (U6-B) — current state

- `domain/errors.py`: `ExcelVbaMcpError.to_dict()` (line 335) emits `code, message, hresult, hresult_hex, context` + `errorCategory, isRetryable, suggestedNextActions` via pure `category_for`/`retryable_for`/`actions_for` (lines 249–300). `ErrorCode` = 29 members.
- Gaps vs target contract: key is `code` not `errorCode`; `details` not present (context is); need to guarantee EVERY server `isError=true` path (including non-`ExcelVbaMcpError` exceptions and validation) yields the full envelope at the MCP boundary (`server.py` `call_tool`), not just the Pydantic class.

## 4. `WorkbookInfo` (U6-E A) — current fields

`domain/models.py:86`: `workbook_id, name, format, path` (working path), `source_path` (original, nullable), `read_only, has_vba, excel_version, session_id, mode, model_version`. No source hash captured at open (relevant to U7 source-overwrite safety).

## 5. Confirmation semantics (U6-E C) — current state

- `snapshot_restore(snapshot_id, workbook_id=None)` (`test_tools.py:122`) — no `confirm`; discards current working copy without saving.
- `excel_recover` (`api/recovery_tools.py`) — no `confirm`; terminates owned session, discards working copies.
- Need explicit agent-intent `confirm` flags (not human approval); missing/false → clear rejection, no destructive action.

## 6. Save / persistence (U7) — current state

- `excel_save(workbook_id)` (`workbook_tools.py:147` → `session_service.save_workbook:474`) saves the working copy **in place** via `wb.Save()`; no path arg; no save-as.
- Working copy materialized by `shutil.copy2(canonical, working)` (`session_service._materialize_working_copy:1199`); original never modified.
- `_SessionRecord` holds `canonical` (source), `working_path`, `info`; **no source hash**.
- `PathPolicy.validate_path` (`security/path_policy.py:110`) canonicalizes (realpath + resolve), enforces allowlisted roots, extension allowlist, symlink/UNC/device rejection. Reuse for target validation.

## 7. Environment / routing (U6-C)

- `xlvba_environment` (`api/environment_tools.py`) reports server/excel/capabilities/policy/workbooks; no `xlvba_*` vs `excel-mcp_*` routing guidance yet.

## 8. Description boilerplate (U6-D)

- 37,125 desc chars across 92 tools; U1 standard applied but repeated guard/glossary text remains. Measure before/after; no lazy loading (excluded).

## 9. Benchmark

- Existing `.ai/tmp/benchmark_tasks.yaml` (29 entries) + `research_2b` harness design. Must add tasks for renamed tools, canonical test tool, routing, structured errors, WorkbookInfo, name conflicts, confirm flags, macro/pivot examples, Save As (new path, source overwrite, rejection, external change).