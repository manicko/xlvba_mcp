# Report 10c — U6-A Validator Verdict

- **Role:** Independent Validator (did NOT implement U6-A).
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d` (working tree carries U6-B + U6-A, uncommitted).
- **Contract:** `.ai/audit/plans/plan_05 — Remaining Usability & Save As Implementation.md` §5.1, §4 N3/N7, §8 U6-A.
- **Method:** read source at working tree, live `list_tools()` introspection, `git diff`/`git status`, full `pytest tests/unit`. Validation only — no files changed except this report.
- **Runtime for introspection:** `C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe` (imports `excel_vba_mcp.server`, registers tools at import).

---

## Per-item verdicts

### Item 1 — Tool surface (count 91, names, no aliases) — **APPROVE**

Live introspection (`mcp.list_tools()`):

```
COUNT 91
HAS vba_compile_diagnostics True
HAS job_request_cancel      True
HAS test_suite_run          True
HAS vba_compile             False
HAS job_cancel              False
HAS test_run_suite          False
```

Registrations confirm the canonical names and the absence of the old ones:
- `server.py:1956-1961` registers `vba_compile_diagnostics` (handler `vba_compile_diagnostics`, `_DESC_VBA_COMPILE_DIAGNOSTICS`).
- `server.py:1992-1997` registers `job_request_cancel` (`_DESC_JOB_REQUEST_CANCEL`).
- `server.py:1890-1895` registers `test_suite_run` (`_DESC_TEST_SUITE_RUN`).
- No `add_tool(... name="vba_compile"|"job_cancel"|"test_run_suite")` anywhere in `server.py` (grep 0). No alias/duplicate registration exists.

Count 91 = baseline 92 minus the removed `test_run_suite` (plan §7.2 "count 92→91 after A"). Matches.

### Item 2 — Descriptions truthful — **APPROVE**

- `vba_compile_diagnostics` description still states it does NOT compile: `server.py:390-404` (`_DESC_VBA_COMPILE_DIAGNOSTICS`) — `"diagnostic only, does NOT compile"` (`:392`), `"never returns compiled or compile_error"` (`:400-401`). Handler docstring `compile_tools.py:63-77` repeats "never compiles".
- `job_request_cancel` description states intent-only / never interrupts: `server.py:472-483` (`_DESC_JOB_REQUEST_CANCEL`) — `"intent-only; never interrupts"` (`:476`), `"the single-STA COM call runs to completion"` (`:479-480`). Handler docstring `job_tools.py:113-124` repeats it.
- `test_suite_run` is the sole test entry with no obsolete cross-ref: `server.py:309-331` (`_DESC_TEST_SUITE_RUN`) describes mapping + typed spec + `background`, and contains **no** `test_run_suite` mention. Module docstring `test_tools.py:11-17` states "single canonical test entry point". `_DESC_TEST_GET_RUN` (`:305-308`) is clean.
- `_TRUTH` contract phrases present (verified by `test_description_contract.py:54-58` + green test run).

### Item 3 — `test_run_suite` fully removed — **APPROVE**

- Def removed; replaced by private `_run_suite_blocking(suite)` at `test_tools.py:160-186`.
- `partial` repointed: `test_tools.py:228` `partial(_run_suite_blocking, spec)` (background path).
- Direct call repointed: `test_tools.py:233` `return await _run_suite_blocking(spec)`.
- `__all__` (`test_tools.py:266-277`) lists `test_suite_run`, not `test_run_suite`.
- Import `server.py:141-148` imports `test_suite_run` only; registration `server.py:1890-1895` only; `_DESC_TEST_RUN_SUITE` deleted (grep 0).
- Audit label repointed: `test_runner.py:1763` `tool="test_suite_run"`.
- **src grep residual:** exactly one hit — `security/macro_policy.py:167` `"test_run_suite"` inside `_EXECUTING_TOOLS` (`:161-169`). This is the **explicitly excluded** stale, non-wired set (plan §8 U6-A Non-scope; `classify_operation` docstring `:186-189` states it is "not called at runtime … also stale"). Excluding `TestAction.VBA_COMPILE` and `_EXECUTING_TOOLS`, residual = 0. Matches the contract's grep exclusion exactly.

### Item 4 — `TestAction.VBA_COMPILE` + its tests unchanged — **APPROVE**

- `domain/test_spec.py:73` `VBA_COMPILE = "vba_compile"` intact.
- Protected tests unchanged (not in `git status` modified set; verified by grep + diff):
  - `tests/unit/test_domain_models.py:144,184` — `{"action": "vba_compile"}`. The only diff in this file is U6-B envelope keys at `:214-220` (`code`→`errorCode`, `context`→`details`); VBA_COMPILE lines untouched.
  - `tests/unit/test_test_runner.py:891,928,958` — unmodified.
  - `tests/integration_excel/test_end_to_end_test_runner.py:870,883` — unmodified.
  - `tests/unit/test_macro_policy.py:92` `("vba_compile", EX)` — unmodified (stale-set test).
- The `vba_compile`→`vba_compile_diagnostics` test edits are confined to the U6-A tool-rename targets (`test_compile_service.py:391,517`; `test_description_contract.py:39,55`) — correct, not the step.

### Item 5 — `test_description_contract.py` — **APPROVE**

- Count: `test_description_contract.py:83-86` `assert len(tools) == 91` with comment "U6-A removes the duplicate test_run_suite (92 -> 91)".
- `_CRITICAL` (`:18-43`) renamed: contains `job_request_cancel` (`:37`), `vba_compile_diagnostics` (`:39`), `test_suite_run` (`:42`); no `job_cancel`/`vba_compile`/`test_run_suite`.
- `_TRUTH` (`:54-58`) renamed keys: `vba_compile_diagnostics` (`:55`), `job_request_cancel` (`:57`); values unchanged. `test_run_suite` removed from `_EXAMPLE_TOOLS` (`:46-51` uses `test_suite_run`).

### Item 6 — Docs EN + RU updated; residual gap noted — **APPROVE (with non-blocking gap)**

- EN + RU docs updated for renames/removal/count 91: `git status` shows modified `docs/01,04,05,06,07,09,10,11` and `docs/ru/01,04,05,06,07,09,10,11`. Count `91` present in `docs/01:39`, `docs/04:60`, `docs/05:21,31,136,138,991`, `docs/11:66`, and RU `01:39`, `05:21,31,138,140,931`.
- **Non-blocking residual gap:** `docs/08-security.md:105` (EN) and `docs/ru/08-security.md:99` still say the old TOOL name `vba_compile` ("every VBA/VBIDE tool and `vba_compile` returns `ACCESS_VBOM_DISABLED`"). `docs/08` is **not** in the U6-A modified set. This is the exact gap class the brief anticipated (docs/08). Non-blocking: docs/08 is security-narrative, not the tool reference, and the change is documentation-only.
- Note: other `vba_compile` doc hits (`docs/05:806-807`, `docs/06:114,321`, `docs/07:525,546,568`, RU equivalents) refer to the **`vba_compile` step** (`TestAction.VBA_COMPILE`), which is intentionally kept — NOT the tool. Correct.
- `docs/.tmp/2026-11-26-autonomous-excel-doc-update.md:7` references `job_cancel` — a transient scratch file under `docs/.tmp`, not a shipped doc; non-blocking.

### Item 7 — Benchmark T22 / T27 — **APPROVE**

- T22 (`.ai/tmp/benchmark_tasks.yaml:465-478`): `expected_first_tool: job_request_cancel` (`:471`), `expected_sequence: [job_request_cancel]` (`:472`), `expected_args.job_request_cancel` (`:473-474`), `token_measure: tool_description_tokens(job_request_cancel)` (`:477`). No `job_cancel` residual.
- T27 (`:547-571`): `must_not` (`:569`) = `["express arbitrary code in a step", "express a shell/COM payload", "skip the closed schema"]` — `test_run_suite` wording removed. `expected_sequence: [excel_open, test_suite_run]` (`:554`).

### Item 8 — `pytest tests/unit -q` green — **APPROVE**

Command: `& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" -m pytest tests/unit -q`
Result: **`1131 passed in 27.69s`**, exit code `0`. (Output: `.ai/tmp/pytest_u6a2.txt`.)

---

## Residual-reference gaps (non-blocking)

1. `docs/08-security.md:105` + `docs/ru/08-security.md:99` — old tool name `vba_compile` in security narrative (docs/08 outside U6-A scope). **Non-blocking.**
2. `src/excel_vba_mcp/security/macro_policy.py:167` — `"test_run_suite"` in the stale, non-wired `_EXECUTING_TOOLS` set. **Explicitly excluded** by contract (N3 / §8 non-scope). **Non-blocking.**
3. `docs/.tmp/2026-11-26-autonomous-excel-doc-update.md:7` — `job_cancel` in a scratch file. **Non-blocking (not a shipped doc).**

No residual old tool names in `src` outside the two contract-excluded locations. No aliases. `SCHEMA_VERSION` untouched by U6-A (no DTO change).

---

## Final U6-A verdict: **ACCEPT**

All 8 acceptance items pass with cited evidence. The single `test_run_suite` src residual is the contract-excluded `_EXECUTING_TOOLS` stale set; the only shipped-doc residual is `docs/08` (anticipated, non-blocking). `TestAction.VBA_COMPILE` and its tests are provably unchanged. Unit suite green (1131 passed). U6-A is correctly implemented per plan_05 §5.1 / §4 N3,N7 / §8 and is accepted.

**Advisory (non-blocking, do not gate U6-A):** a follow-up doc pass should update `docs/08-security.md:105` / `docs/ru/08-security.md:99` to `vba_compile_diagnostics` (tool) while preserving the `vba_compile` step references elsewhere; optionally prune `_EXECUTING_TOOLS` stale entries (`vba_compile`, `test_run_suite`, `compile`, `test_run`) since the set is documented non-wired.