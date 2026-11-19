# Architecture and Code Audit — Excel VBA MCP Server (pre-task_002 / pre-task_010)

**Author:** auditor
**Date:** 2026-11-19
**Repo:** `C:\py_exp\xlvba_mcp`
**Commit measured:** `018ae73` (`chore(agents): narrow format command deny pattern`)
**Working-tree state:** dirty
**Dirty paths:** `uv.lock` (pydantic 2.14.3→2.14.4, pydantic-core 2.50.4→2.50.5 auto-bump), `.ai/plans/plan.md` (3.12→3.14 deviation note), `.ai/plans/execution/execution-log.md` (runtime decision + Office probe), untracked `.ai/tmp/`, untracked `tests/integration_excel/`
**Plan under audit:** `.ai/plans/plan.md` (37 tasks, Stage 1)
**Spec under audit:** `docs/01-Spec/excel_vba_mcp_technical_spec.docx` (text at `.ai/tmp/spec_text.txt`)

---

## Scope

This audit answers one question: **does the current skeleton plus the task_002 design package form a sound, spec-aligned foundation to proceed with task_002 (COM/STA PoC) and then tasks 007–016?**

In scope:

1. Module layout `src/excel_vba_mcp/**` vs plan task_001 file list and spec §5.1 directory structure.
2. `server.py`, `__main__.py`, `config.py` correctness: MCP SDK 2.3.0 API usage, STDIO lifecycle, stdout discipline, blocking behaviour, config typing.
3. task_002 design-gate artifacts (`task_002_planner_v3.md`, auditor v1–v3, validator v1–v3, researcher) and the on-disk `tests/integration_excel/test_excel_instance_poc.py` fragment.
4. Planned STA-worker design (plan task_010) vs spec §5.3 COM-concurrency rules: dedicated STA, `CoInitializeEx(COINIT_APARTMENTTHREADED)`, `DispatchEx("Excel.Application")`, serialized queue, no proxy escape, busy/`RPC_E_CALL_REJECTED`.
5. Security boundaries: plan task_008 (path policy), task_009 (macro/approval policy), task_014 (snapshots) vs spec §9–§10.
6. Python runtime 3.14.7 vs plan target 3.12 / `pyproject.toml` (`requires-python >=3.11,<3.15`, ruff/mypy `py311`) and MCP SDK compatibility.
7. Integration points later tasks depend on, and ordering/contract risks.

Out of scope (per instructions): proposing a redesign, new features, modifying code, executing macros, mutating workbooks, Trust Center changes. Excel was launched only read-only through `DispatchEx` for environment probing, and every probe process was shut down (`EXCEL.exe` count verified 0 afterwards).

Method summary is in **Method** below; runtime commands actually executed are listed so each `observed` finding is reproducible.

---

## Method

1. **Read** `src/excel_vba_mcp/{__init__,__main__,server,config}.py`, `pyproject.toml`, `README.md`, `.gitignore`, `uv.lock` header, `tests/test_smoke.py`, and the 262-line `tests/integration_excel/test_excel_instance_poc.py`.
2. **Read** plan §1–§7 and task_001–task_016 YAML specs with line numbers (`plan.md:66-868`), plus spec §3–§13 extracted text.
3. **Read** all task_002 gate artifacts under `.ai/plans/execution/` and the execution log (including the uncommitted runtime-decision block).
4. **Executed read-only probes** (documented for reproducibility):
   - `uv run python` → Python 3.14.7, `mcp 2.3.0`, `pydantic 2.14.4`, `pywin32 312`, `pywinauto 0.6.9`, `pytest 9.2.0`, `ruff 0.17.5`, `mypy 2.4.1`.
   - `uv run pytest --collect-only -q` → collection error in the PoC file (`NameError: dataclass`).
   - `uv run ruff check src tests` → 56 × F821 in the PoC; `uv run ruff format --check` → PoC would be reformatted; `uv run mypy src` → clean.
   - STDIO protocol probes: `initialize`, `tools/list`, `tools/call server_health` over a real child process; stdin-EOF shutdown; SIGTERM during an in-flight sync tool; stray-print routing.
   - COM probes: `win32com.client.DispatchEx("Excel.Application")` on the main thread (implicit COM init works); on an explicitly `CoInitializeEx(APARTMENTTHREADED)` thread; cross-thread/cross-apartment access from MTA and STA probe threads; no orphaned `EXCEL.exe` after each probe.
5. **Verified SDK internals** in the installed wheel: `mcp/server/mcpserver/server.py` (`run`, `run_stdio_async`), `mcp/server/mcpserver/utilities/func_metadata.py:149-164` (sync tools run via `anyio.to_thread.run_sync`), `mcp/server/stdio.py` (fd claim/diversion on Windows), `mcp/os/win32/utilities.py` (`rebind_std_handle_to_fd`).
6. **Classified** findings as confirmed defects (with file:line and observed evidence) vs risks/uncertainties; each finding carries a `verification` field (`observed`/`unverified`) and a `classification` (`mandatory`/`advisory`).

No source, test, config, or plan file was modified. The only file written is this report.

---

## Findings

Format per finding: ID, severity, type, classification, verification, affected modules, evidence (file:line + reproduction), recommendation.

### ARCH-001: task_002 PoC is non-importable and non-collectable; it is a mid-function fragment, not a test module

| Field | Value |
|---|---|
| **ID** | ARCH-001 |
| **Severity** | CRITICAL |
| **Type** | RUNTIME-ERROR |
| **Classification** | mandatory |
| **Verification** | observed |
| **Affected Modules** | `tests/integration_excel/test_excel_instance_poc.py` |

**Description:** `task_002` is the critical Go/No-Go gate for the whole Excel branch (`plan.md:118-167`, R01 at `plan.md:29`, DAG G1 at `plan.md:47`). The only artifact on disk for it is untracked (`git status` shows `?? tests/integration_excel/`; `git ls-files tests` returns only `tests/test_smoke.py`) and is syntactically a fragment: it begins at line 1 with `def snapshot_isolation_baseline(...)`, contains no import block, no `PocArtifacts` definition, and no `test_*` function. It cannot be imported or collected, so **no part of task_002 is executable today**.

**Evidence (file:line):**
- `tests/integration_excel/test_excel_instance_poc.py:1` — file starts with a function definition, not a module docstring/import. Whole file is 262 lines; `Select-String '^def |^class '` returns only helper functions and two dataclasses, zero `test_` functions.
- Undefined names (all from `ruff check`, 56 F821): `dataclass`, `field` (line 63, 65-66, 225, 246, 250, 255), `Any` (1, 6, 41-42, 65, 80, 230-255), `PocArtifacts` (1, 43, 80, 140, 227), `pythoncom` (57, 60), `win32com` (8), `EXCEL_PROGID` (8, 243), `contextlib` (24, 59), `json` (118), `tempfile` (244), `Path` (245), `threading` (91), `_now_local_iso` (30), `_now_utc_iso` (143, 183), `_list_excel_processes_cim` (3), `_list_excel_windows` (5), `_com_error_details` (89, 169, 209), `_hresult_hex_from_exc` (173, 213).
- Reproduction:
  ```
  uv run pytest --collect-only -q
  # tests\integration_excel\test_excel_instance_poc.py:63: in <module>
  # @dataclass
  # E   NameError: name 'dataclass' is not defined
  # !!! Interrupted: 1 error during collection !!!
  ```
  ```
  uv run python -c "import importlib.util as u; s=u.spec_from_file_location('poc','tests/integration_excel/test_excel_instance_poc.py'); m=u.module_from_spec(s); s.loader.exec_module(m)"
  # NameError: name 'dataclass' is not defined
  ```
- Consequence: a full `uv run pytest` is red (`exit=2`), so the task_001 acceptance criterion "package import/startup smoke test" still passes alone (`uv run pytest tests/test_smoke.py` → 1 passed) but the repository-level test command does not. `ruff check` and `ruff format --check` also fail (56 errors; 1 file would be reformatted), so any CI gate mirroring README development commands fails now.

**Missing pieces required by `task_002_planner_v3.md` §3–§9 (enumerate):**
1. Module docstring + imports (`contextlib`, `dataclasses.dataclass/field`, `json`, `tempfile`, `threading`, `pathlib.Path`, `typing.Any`, `datetime`, `pythoncom`, `win32com.client`, `pytest`).
2. `EXCEL_PROGID` constant and `PocArtifacts` class (artifact schema: environment matrix, `steps[]`, `serialized_multiop_evidence`, affinity fields, snapshot methods, baselines, `go_no_go_r01_task002`) with `append_log`/`add_step` methods.
3. Helper implementations `_list_excel_processes_cim`, `_list_excel_windows`, `_com_error_details`, `_hresult_hex_from_exc`, `_now_local_iso`, `_now_utc_iso`, `setup_poc_artifact_logger`.
4. Lifecycle/driver code: STA fixture with `CoInitializeEx`/`CoUninitialize`, `DispatchEx` + owned-PID resolution, `Workbooks.Add/SaveAs/Open/Close` per planner §5.2, timed orphan poll, `finally` cleanup.
5. The two required tests named in `plan.md:147-148`: `test_dispatch_ex_creates_isolated_excel_instance`, `test_excel_workbook_open_close_lifecycle`, plus the `@pytest.mark.integration_excel` marker (planner §9) — and the marker must be registered in `pyproject.toml` (no `[tool.pytest.ini_options] markers` today).
6. Evidence writers to `.ai/tmp/excel_instance_poc_<ts>.{json,txt}` with the §3 matrix, `steps[]` (§3.1), and final rubric field `go_no_go_r01_task002` (§8).

**Recommendation:** Treat ARCH-001 as a **blocking precondition for task_002 execution**, not a defect to patch silently. Two admissible options, both spec-compatible:
- **(a) Finish the fragment** under planner v3 §3–§9 binding (recommended if the prior session's helpers are to be preserved). Minimum: restore imports/constants/`PocArtifacts`, implement helpers, add the two named tests and marker, verify `pytest --collect-only` is green, then run on the target Excel build and persist evidence.
- **(b) Delete the untracked fragment** and implement `tests/integration_excel/test_excel_instance_poc.py` from planner v3 directly.
Either way, `uv run pytest`, `ruff check src tests`, `ruff format --check src tests` must be green *before* the PoC is run, otherwise the gate's own evidence is produced from a red baseline. Do **not** commit the fragment as-is.

### ARCH-002: task_002 gate is "design-cleared" but no executable evidence exists; the gate's own Go/No-Go field is never produced

| Field | Value |
|---|---|
| **ID** | ARCH-002 |
| **Severity** | HIGH |
| **Type** | SPEC-DEVIATION |
| **Classification** | mandatory |
| **Verification** | observed |
| **Affected Modules** | `.ai/plans/execution/execution-log.md`, `.ai/plans/execution/task_002_*` |

**Description:** The execution log records task_002 as `REVIEW_REQUIRED (design gate cleared for implementation under conditions)` with validator v3 `PASS WITH WARNINGS` (`execution-log.md` task tracker row). The plan's task_002 is `verification_mode: dedicated_gate` and its acceptance criteria require *runtime* evidence ("отдельный Excel instance создан и идентифицирован", "open/close/Quit не оставляют orphaned owned Excel process", "COM proxies остаются в создавшем их STA thread", "поведение message pump и ограничения задокументированы", `plan.md:154-159`). None of that evidence exists yet: no `.ai/tmp/excel_instance_poc_*` artifact exists (`.ai/tmp` contains only `spec_text.txt`), no `go_no_go_r01_task002` value has been produced, and the execution log's own change log ends at "No PoC code implemented".

This is not an accusation that the reviewers erred — auditor v1→v3 and validator v1→v3 explicitly gate on *runtime evidence* and forbid treating the design as proof. It is a **process-state finding**: downstream tasks (003/004/005/006, 010) are marked BLOCKED/PENDING on a gate that currently has design approval but zero execution evidence. Any attempt to start task_007/010 implementation now would violate plan §6 rule 1-2 (DAG order; PoC is Go/No-Go for the dependent branch) and risk building `com_worker.py` on unproven STA/modal assumptions.

**Evidence:**
- `execution-log.md` tracker: `task_002 | ... | REVIEW_REQUIRED (design gate cleared for implementation under conditions) | cee9755 ... | No executable PoC tests run yet.`
- `execution-log.md` change log last entry: `**No PoC code implemented** for tests/integration_excel/test_excel_instance_poc.py pending closure and re-validation.`
- `task_002_validator_v3.md`: "Implementation ... is **permitted** ... Final Go/No-Go is target-run evidence driven."
- `.ai/tmp/` contains only `spec_text.txt`; no `excel_instance_poc_*.json|txt`.
- `plan.md:60`: "A failed gate blocks its dependent implementation; it is not permission to silently substitute a different architecture or to report an unverified capability as complete."

**Recommendation:** Keep task_002 open as `IN_PROGRESS` until the artifact exists and `go_no_go_r01_task002` is one of `GO | CONDITIONAL_GO | NO_GO | INCONCLUSIVE` per planner §8. Update the execution log to reflect "design gate PASSED; execution evidence PENDING" rather than "gate cleared", so a later reader cannot mistake design approval for a Go. Do not start task_010 (`com_worker.py`) before this field is recorded.

### ARCH-003: task_002 design package remains internally good, but two evidence rules are materially weaker than the audit needs (affinity probe semantics; "not exercised" modal path)

| Field | Value |
|---|---|
| **ID** | ARCH-003 |
| **Severity** | MEDIUM |
| **Type** | BEST-PRACTICE |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `.ai/plans/execution/task_002_planner_v3.md` |

**Description:** Planner v3 (binding) closes auditor v1 D1–D9 / v2 N1–N5 in a traceable way (§0 table, §3.1 step schema, §4 isolation, §5 lifecycle/orphans, §6 pump/RPC/modal, §7 mandatory affinity probe, §8 rubric, §9 harness outline, §10 operator checklist). This audit confirms the design is **sufficient as a design gate** and aligned with `com_worker.py` intent. Two residual weaknesses should be closed *during implementation* rather than by a v4 rewrite:

1. **Affinity probe result depends on how the probe thread initializes COM, and planner v3 does not pin it.** Planner §7 says "attempt a trivial read on that proxy ... expect failure", without specifying whether the probe thread calls `CoInitializeEx` first. On this host both variants fail, but for different reasons, which matters for the evidence's meaning:
   - probe thread with **no** `CoInitializeEx` → `com_error(-2147221008, 'CoInitialize...')` / `0x800401F0` (`CO_E_NOTINITIALIZED`) or, through pywin32 dynamic dispatch, `AttributeError: Excel.Application.Version` (no HRESULT surfaced);
   - probe thread with `CoInitializeEx(APARTMENTTHREADED)` or `COINIT_MULTITHREADED` → `com_error(hresult=0x8001010E, 'Ошибка из-за ...' )` (`RPC_E_WRONG_THREAD` — the call was "rejected because the callee is not in the right apartment").
   An affinity probe that fails with `CO_E_NOTINITIALIZED` does **not** demonstrate apartment confinement; it demonstrates the probe thread never joined COM. For the gate to prove confinement, the probe thread must initialize COM (either apartment) and then fail with `RPC_E_WRONG_THREAD (0x8001010E)`. Reproduction: see the two probes run for this audit (`probe2.py`, `probe_sta2.py`, results quoted above).
2. **"Not exercised" RPC/modal is allowed to yield `CONDITIONAL_GO` for "lifecycle/isolation/STA-core" (planner §8.2).** That is reasonable, but it means the critical path's real unknown (message pump under a modal VBE/compile dialog) is deferred to task_004/task_005/task_010. The execution log records the condition but not the follow-up obligations as first-class checklist items.

**Recommendation (advisory, not a v4 rewrite):**
- Add one sentence to the *implementation* (not necessarily the plan): the affinity probe thread MUST call `pythoncom.CoInitializeEx(COINIT_MULTITHREADED or APARTMENTTHREADED)` before touching the proxy; the pass condition is `RPC_E_WRONG_THREAD (0x8001010E)`. Record the HRESULT in `steps[]`. If the probe fails with `CO_E_NOTINITIALIZED`, mark the evidence `inconclusive`, not "expected failure".
- Record the "not exercised" caveat and the three follow-up gates (`task_004`, `task_005`, `task_010` integration) as explicit `CONDITIONAL_GO` notes in the artifact and execution log.

### ARCH-004: `server.py` is functionally correct for MCP 2.3.0 STDIO, but `config.load_config()` ignores the environment despite its docstring, and log level is unvalidated

| Field | Value |
|---|---|
| **ID** | ARCH-004 |
| **Severity** | MEDIUM |
| **Type** | SPEC-DEVIATION |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `src/excel_vba_mcp/config.py`, `src/excel_vba_mcp/server.py` |

**Description:** `config.py:1-5` states "Configuration is loaded from environment variables." `config.py:24-31` returns `Config()` unconditionally; no `os.environ` read, no validation. `server.py:19-23` feeds `_config.server.log_level` directly into `logging.basicConfig(level=...)`; because the value is the raw `str` `"INFO"`, an invalid level (if env loading is later added) would raise `ValueError: Unknown level` at import time, before the MCP server can report a clean error. The typing is also weaker than the project's own standard ("Use Pydantic for all data models and validation", `.ai/context/python-code-standards.md`): `log_level: str` instead of a `Literal`/enum.

This is not a runtime defect today (only the default is used), but it is a **spec deviation with forward risk**: the plan's task_001 goal is "typed configuration" (`plan.md:99`, `source_section` "5.1/6"), and every later task will need path roots, policy files (`config/default.toml`, `config/macro-policy.example.toml` per spec §5.1), artifact root, log level and lock timeouts. Leaving the loader a stub makes task_008/task_009/task_012 each invent their own config access.

**Evidence:**
- `config.py:24-31` — `return Config()`; no environment access anywhere in the file (verified by grep: only `pydantic` import).
- `server.py:17-23` — `_config = load_config()` at import time; `logging.basicConfig(level=_config.server.log_level, ...)`.
- Reproduction: `uv run python -c "from excel_vba_mcp.config import load_config; print(load_config())"` → `log_level='INFO' server=...`; setting any env var (e.g. `LOG_LEVEL=DEBUG`) has no effect.
- MCP protocol itself verified working: `initialize` → `serverInfo.name="excel-vba-mcp"`, `tools/list` → `server_health` with correct `inputSchema`/`outputSchema`, `tools/call` → `structuredContent {"status":"ok","server":"excel-vba-mcp"}`; diagnostics only on stderr (`2026-11-19 ... INFO excel_vba_mcp.server: Starting ...`). So ARCH-004 is **not** a transport/stdout defect.

**Recommendation:** Either (a) implement env loading now (small): read a documented set of env vars, validate with a Pydantic model / `Literal["DEBUG","INFO","WARNING","ERROR","CRITICAL"]`, fail with a clear message; or (b) if deferring, correct the docstring to "Configuration is currently defaults-only; environment loading is added with path/security config" so the code does not claim behaviour it lacks. Do not let task_008/task_009/task_012 each add ad-hoc env reads — the config module is the single place for roots and policy paths (spec §5.1/§10). This is advisory for task_001, but becomes a **coordination risk** for tasks 008/009/012 if unresolved.

### ARCH-005: no `MCPServer` lifespan / shutdown hook, so the STA worker (task_010) has no defined ownership point

| Field | Value |
|---|---|
| **ID** | ARCH-005 |
| **Severity** | MEDIUM |
| **Type** | BEST-PRACTICE |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `src/excel_vba_mcp/server.py`, `src/excel_vba_mcp/__main__.py` |

**Description:** `server.py:26` creates `MCPServer("excel-vba-mcp")` with no `lifespan=`. `MCPServer.__init__` supports `lifespan: Callable[[MCPServer], AbstractAsyncContextManager[LifespanResultT]] | None` (verified in the installed SDK). Today there is no resource to own, so this is harmless. But task_010 introduces a dedicated STA thread that owns an `Excel.Application` and must run `Quit`/release/`CoUninitialize` in `finally`; task_012 adds session state; task_014 adds snapshot store handles. Without a lifespan or an explicit shutdown path, those resources will either (a) be tied to `atexit`/daemon-thread heuristics, or (b) leak if the client disconnects.

Observed shutdown behaviour to design against:
- **stdin EOF** → graceful exit rc 0 in ~0.17 s (`stdio_server` restores fds; verified).
- **SIGTERM while no tool is running** → exits immediately rc 1 (Windows `TerminateProcess`-like semantics; no Python `finally` runs).
- **SIGTERM while a synchronous tool is in flight** → process stays alive until the tool returns (a 20 s `time.sleep` tool kept the process alive >10 s after `terminate()`; `p.wait(timeout=10)` timed out). This is expected (Windows has no tree-wide SIGTERM), but it means a hung COM call (modal dialog) cannot be interrupted by the client, and any worker cleanup in a `finally` will not run.
- **Hard kill** → rc 1, no stderr, no cleanup.

**Recommendation:** When implementing task_010/task_012, add an explicit lifecycle object owned by `MCPServer(lifespan=...)` (or a module-level `SessionService` started/stopped by `main()` around `mcp.run`) that: starts the STA worker, registers `atexit`/finally cleanup, and defines who may call `worker.shutdown()` on client disconnect. Record in task_011's test matrix the three observed exit paths (EOF / SIGTERM idle / SIGTERM busy) and assert "no orphaned `EXCEL.exe`" for the paths where Python cleanup can run, and "documented orphan risk + recovery" for hard kill. This is advisory now, **mandatory before task_011 sign-off**.

### ARCH-006: package layout matches spec §5.1 for task_001; only `config/` and two test dirs are missing, and the api filename for Shapes deviates

| Field | Value |
|---|---|
| **ID** | ARCH-006 |
| **Severity** | LOW |
| **Type** | SPEC-DEVIATION |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `src/excel_vba_mcp/**`, `config/`, `tests/` |

**Description:** Verified against spec §5.1 (spec_text.txt:186) and plan task_001 file list (`plan.md:90-102`):

- **Present and correctly named:** `pyproject.toml`, `uv.lock`, `README.md`, `src/excel_vba_mcp/__main__.py`, `server.py`, `config.py`, `api/`, `application/`, `domain/`, `adapters/excel/`, `adapters/windows_ui/`, `security/`, `storage/`, `projects/`, `work/`, `snapshots/`, `artifacts/`, `tests/fixtures/`, `tests/integration_excel/`.
- **Missing vs spec §5.1:** top-level `config/` with `default.toml` and `macro-policy.example.toml`; `tests/unit/`; `tests/test_data/`.
- **Empty package dirs:** all of `api`, `application`, `domain`, `security`, `storage`, `adapters/excel`, `adapters/windows_ui` contain only `__init__.py` (0 bytes). This is correct for task_001 (skeleton) and is *not* dead code — the plan's later tasks fill exactly these files (task_007 domain, task_010 com_worker/app_adapter, task_012 session_service/workbook_adapter, etc.). No action beyond noting that `adapters/excel/__init__.py` etc. must remain importable packages.
- **Naming mismatch:** spec §5.1 lists `shape_tools.py` (singular) among `api/*_tools.py`; plan task_030 (`plan.md:1531`) specifies `src/excel_vba_mcp/api/shape_tools.py` with function `shapes_list`. That matches spec. However spec §6.5 tool names are `shapes_list`, `controls_list`, `control_get`, `control_set_on_action` — task_030's `api/shape_tools.py` only lists `shapes_list`, `shape_get`, `shape_set_on_action` (`plan.md:1533-1535`). `controls_list`/`control_get`/`control_set_on_action` are **not assigned to any `api/*` file** in task_030's `files:` block, though the goals require them. This is a plan-internal gap, not a code gap; see ARCH-014.
- **`tests/fixtures/` is empty** and `tests/unit/` absent: fixtures are task_003+ deliverables (`CleanProject.xlsm` etc.), so expected now.

**Recommendation:** Create `config/` with `default.toml` + `macro-policy.example.toml` when task_008/task_009 land (they need roots/policy paths), and create `tests/unit/`, `tests/test_data/` at the same time so the plan's `tests_to_run` paths resolve. No code change needed now.

### ARCH-007: security design (tasks 008/009/014) is spec-aligned; two enforcement points are under-specified and must be resolved before implementation

| Field | Value |
|---|---|
| **ID** | ARCH-007 |
| **Severity** | MEDIUM |
| **Type** | BEST-PRACTICE |
| **Classification** | advisory |
| **Verification** | unverified (design review; no path-policy code exists yet) |
| **Affected Modules** | `src/excel_vba_mcp/security/*`, `src/excel_vba_mcp/application/*` (future) |

**Description:** Plan task_008 (`path_policy.py`, canonical path, allowlist roots `projects/`/`work/`/`snapshots/`, traversal/symlink/reparse/UNC denial, `PATH_NOT_ALLOWED`) and task_009 (`macro_policy.py`, `approval.py`, `audit.py`, read/mutate/execute classification, workbook path+hash allowlist, no Trust Center/registry change, no VBA password bypass, redacted audit) map cleanly onto spec §10 (table rows "Доступ к файлам", "Разделение режимов", "VBOM access", "Macro security", "Approval", "VBA project lock", "Аудит", "MCP transport", "Секреты") and §10.1 (safe open order). Plan task_014 snapshot model matches spec §9.1/§9.4: `SaveCopyAs` for open, file copy for closed, `manifest.json` + SHA-256 + timestamp + source path + Excel version + workbook model + VBA exports; restore closes working copy without save; explicit "snapshot is not a sandbox". No spec deviation found in the *design*.

Two enforcement points are not yet pinned and will cause defects if implementers guess:

1. **Path policy scope vs COM boundary.** task_008 acceptance says "arbitrary filesystem paths are never sent directly to Excel COM" (`plan.md:447`), but `Workbooks.Open` receives a path. The policy must return the *canonical* path and the COM call must use **only** that value, with the check performed on the STA worker (or on a value captured before queueing) — otherwise a TOCTOU swap between validation and `Open` is possible. No task currently states *where* `validate_path` is called relative to the worker queue.
2. **`work/` copy semantics vs "original unchanged".** Spec §7.1: "По умолчанию работать с копией в `work/`"; task_012 goals repeat it. But task_012 also requires `excel_open` to return "original path and actual working path", and task_014 restore opens a snapshot. The copy step (original → `work/`) is a filesystem mutation that itself needs path policy + audit, and it is not assigned to a named method in task_012's targets (`open_workbook/close_workbook/save_workbook/get_status` only). If `open_workbook` silently copies, the "no save without explicit instruction" guarantee (task_012 acceptance) needs an explicit definition of whether the *copy* counts as the working file.

**Recommendation:** When planning task_008, add to its acceptance/tests: "canonical path is resolved once and the same value is passed to COM; revalidation occurs on the worker for queued operations". When planning task_012, name the copy step explicitly (e.g. `SessionService._materialize_working_copy`) and state in the test matrix that the original file hash is unchanged after open/mutate/save/close. Both are spec-compatible; neither requires new architecture.

### ARCH-008: Python 3.14.7 runtime — acceptable deviation, but the repo does not pin it and the docs/tooling still claim 3.12/`py311`

| Field | Value |
|---|---|
| **ID** | ARCH-008 |
| **Severity** | MEDIUM |
| **Type** | SPEC-DEVIATION |
| **Classification** | mandatory (documentation/config consistency, not code) |
| **Verification** | observed |
| **Affected Modules** | `pyproject.toml`, `README.md`, `.python-version` (absent), `uv.lock`, `.ai/plans/plan.md` |

**Description:** Runtime facts: only `C:\Python314\python.exe` (3.14.7) is on PATH; `.venv` is 3.14.7 (`pyvenv.cfg` `version_info = 3.14.7`); `uv python list` shows a managed **3.12.16 already installed** at `C:\Users\MatveichevN\AppData\Roaming\uv\python\cpython-3.12.16-windows-x86_64-none`. The plan target is 3.12 (`plan.md:82-89`), the spec stack says "Python 3.12 x64 (минимум 3.11)" (spec_text.txt:148-149), and `pyproject.toml` allows `>=3.11,<3.15` with ruff/mypy `target-version`/`python_version = "py311"`.

Compatibility verified for 3.14:
- `mcp 2.3.0` requires `>=3.10`; its metadata has explicit 3.14 markers (`anyio>=4.10; python_version >= '3.14'`), so the SDK supports 3.14.
- `pywin32 312` wheel is `cp314-cp314-win_amd64` with `pythoncom314.dll`/`pywintypes314.dll`; `DispatchEx` works on 3.14.7 (verified: `16.0`, build `14332.0`). `pywinauto 0.6.9` imports.
- The environment is **GIL-enabled** (`sysconfig.get_config_var("Py_GIL_DISABLED") == 0`), so free-threading caveats (pywin32 "knowingly unsafe" no-GIL wheel, anyio free-threading fixes) do **not** apply.
- `mypy` clean under both default and `--python-version 3.14`; ruff/mypy still *target* 3.11 syntax, which 3.14 accepts.

The deviation itself is **technically sound** and already documented in the working tree (`execution-log.md`: "Runtime decision ... relaxed to Python 3.14.7 ... do not install/pin 3.12"; `plan.md:82-89` edited to 3.14). But three inconsistencies remain:
1. `README.md:12` still says "Python 3.12 x64 recommended (minimum 3.11)".
2. `pyproject.toml` still has `target-version = "py311"` and `python_version = "3.11"`, and no `.python-version` pin exists, so `uv sync` on another machine may create a 3.11/3.12/3.13 venv and produce different `pywin32`/`pydantic` wheels than the validated 3.14 set. `uv.lock` `requires-python = ">=3.11, <3.15"` with 3.14 resolution markers means the lock is 3.14-aware but not 3.14-*pinned*.
3. `uv.lock` is modified but uncommitted (pydantic 2.14.3→2.14.4), and the execution log itself is uncommitted; a fresh clone would resolve 2.14.3.

**Recommendation:** Pick one and make the repo consistent:
- **Preferred (recommended): keep 3.14.7 and pin it.** Add `.python-version` containing `3.14`, change `[tool.ruff] target-version` to `py314` (or leave `py311` only if 3.11 compatibility is a real goal — then say so), update `README.md` requirements, and commit `uv.lock`. Keep `requires-python = ">=3.11,<3.15"` if 3.11 must remain installable, but state in README that the validated runtime is 3.14.7 and 3.11/3.12/3.13 are unvalidated. Rationale: the only interpreter present is 3.14, the SDK/pywin32 both ship 3.14 artifacts, and task_002 will generate COM evidence on 3.14 — pinning avoids "works on my venv" drift in the critical path.
- **Alternative (only if 3.12 is a hard requirement): provision it via uv.** `uv python pin 3.12` + `uv sync --python 3.12` would use the already-installed 3.12.16. This is cheap here (no download) but contradicts the recorded decision and would invalidate the 3.14 COM probe results (must be redone on 3.12). Do not do both.
- Do **not** leave ruff/mypy at `py311` while running 3.14 without a comment: it means the linters will not catch 3.12+-only syntax and the "minimum 3.11" claim is unenforced. Either enforce 3.11 compatibility (and test it) or target 3.14.
- Also commit or discard the `uv.lock` bump deliberately; an uncommitted lock change means task_001's "environment reproduced from uv.lock" acceptance is only true for the dirty tree.

### ARCH-009: sync MCP tools execute on SDK-managed worker threads — direct COM use in a tool would violate the single-STA rule; async tools must not touch COM

| Field | Value |
|---|---|
| **ID** | ARCH-009 |
| **Severity** | HIGH |
| **Type** | BEST-PRACTICE |
| **Classification** | mandatory (constraint for task_010+) |
| **Verification** | observed |
| **Affected Modules** | `src/excel_vba_mcp/server.py`, future `api/*_tools.py`, `adapters/excel/com_worker.py` |

**Description:** Verified in the installed SDK: `FuncMetadata.call_fn` (`mcp/server/mcpserver/utilities/func_metadata.py:149-164`) does
```python
if fn_is_async:
    return await fn(**kwargs)
return await anyio.to_thread.run_sync(functools.partial(fn, **kwargs))
```
So a **synchronous** `@mcp.tool()` function runs on an `anyio` worker thread from a default limiter pool of 40 threads. Empirically confirmed: with a 20 s sleeping sync tool, the server still answered `tools/list` (id 3) during the sleep, proving the event loop is not blocked and the tool is off-thread. This has two consequences that the plan implies but no task states explicitly:

1. **A sync tool that calls `win32com` directly would create a COM apartment on an arbitrary pool thread** (pywin32 initializes COM implicitly on first use — verified: `DispatchEx` on the main thread works with no explicit `CoInitialize`), and different calls could land on different pool threads. That is exactly the "COM proxies cross threads / multiple apartments" failure the spec §5.3 forbids and R01 exists to prevent. `api/*_tools.py` must therefore be **thin**: validate DTOs, then `await`/`submit` to the single `ExcelComWorker`; never hold a `Workbook`/`Application`/`Range` reference.
2. **An `async def` tool runs on the event loop thread.** If such a tool blocks on a `queue.get()`/`threading.Event` waiting for the STA worker, it blocks the entire MCP server (including cancellation and `tools/list`). Task_010's `submit()` contract must be async (awaitable) and must not be implemented with a blocking `join()` on the event loop.

**Recommendation:** Record this as an explicit implementation constraint for task_010/task_012 (and reference it in task_011 verification): "`api/*` handlers are sync or async but never touch COM; all COM goes through `ExcelComWorker.submit()` which is awaitable and bounded by a timeout; a blocked worker yields `running/needs_recovery`, not an event-loop stall." Add a unit test that `tools/list` responds while an Excel operation is in flight (the same probe used in this audit is a suitable template). This is not a current defect (`server_health` is pure Python), but it is the highest-probability architectural regression once task_012 lands.

### ARCH-010: PoC COM facts confirm the plan's core assumptions on this host, with two caveats for the evidence record

| Field | Value |
|---|---|
| **ID** | ARCH-010 |
| **Severity** | LOW |
| **Type** | SPEC-DEVIATION |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | task_002 evidence, `plan.md` R01 |

**Description:** Read-only COM probes on this host (Windows 11, Excel 16.0 build 14332, Python 3.14.7, pywin32 312) confirm the plan's primitives:
- `win32com.client.DispatchEx("Excel.Application")` creates a separate instance and returns `Version="16.0"`, `Build="14332.0"`; no orphaned `EXCEL.exe` after `Quit` (verified `Get-Process EXCEL` count 0 after each probe).
- `pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)` succeeds on a dedicated thread; `CoUninitialize` on the same thread returns cleanly; the thread exits.
- Cross-apartment access fails with `RPC_E_WRONG_THREAD (0x8001010E)` when the probe thread has initialized COM — good evidence for the affinity probe (see ARCH-003).
- `pythoncom` import does **not** set `sys.coinit_flags` in this environment (`getattr(sys, "coinit_flags", "absent") == "absent"`), and a **second thread that calls `DispatchEx` without `CoInitializeEx` fails** with `CO_E_NOTINITIALIZED (0x800401F0)` / `com_error(-2147221008)`. This is important: it means "implicit COM init" is per-thread and fragile, and every worker/probe thread must call `CoInitializeEx` explicitly. It also means a naive "just use a thread pool" implementation will fail intermittently depending on which thread initializes first.

Caveats to record in the task_002 artifact (not defects):
- Excel is **16.0 build 14332** (Office 2016 family), not the spec's "recommended matrix: Microsoft 365 Apps x64" (spec_text.txt:147). Plan task_037 requires release evidence to state the *tested* configuration; this is a legitimate tested configuration but must not be described as the recommended one.
- No user Excel session existed during probes, so the "existing user session untouched" branch of task_002 has **not** been exercised. The PoC must run once with a user workbook open (read-only baseline) to satisfy `plan.md:156` / planner §4.1.

**Recommendation:** In the task_002 artifact, record: Office build 14332, Python 3.14.7, pywin32 312, `CoInitializeEx` required per thread (with the observed `0x800401F0` failure as evidence), affinity probe `0x8001010E`, and a note that the "user session open" isolation case is pending. Add a line to `docs/release-evidence.md` (task_037) draft: "validated on Excel 16.0.14332; other builds untested."

### ARCH-011: `server.py` import-time side effects and no `if __name__` guard in `main` are acceptable but create a subtle test/tooling constraint

| Field | Value |
|---|---|
| **ID** | ARCH-011 |
| **Severity** | LOW |
| **Type** | BEST-PRACTICE |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `src/excel_vba_mcp/server.py`, `src/excel_vba_mcp/__main__.py` |

**Description:** `server.py` calls `load_config()` and `logging.basicConfig()` at import time (lines 17–23) and constructs the global `mcp` at line 26. Importing `excel_vba_mcp.server` therefore configures root logging as a side effect. This is a common MCP pattern (the SDK docs do the same), and it is not a stdout/transport bug — but it means unit tests cannot import the module without capturing logging, and any future config error surfaces as an import failure rather than a startup error. `__main__.py` correctly imports `logger, mcp` and calls `mcp.run(transport="stdio")` inside `main()` with an `if __name__ == "__main__"` guard.

**Recommendation:** No change needed for task_001. When adding tests that import `server` (e.g. tool registration tests), use `caplog` or restructure logging into a `configure_logging()` function called from `main()`; keep `mcp` module-level because `@mcp.tool()` decorators require it.

### ARCH-012: `pywinauto` is a declared dependency but unused; the task_004 UIA path has no spike yet

| Field | Value |
|---|---|
| **ID** | ARCH-012 |
| **Severity** | LOW |
| **Type** | BEST-PRACTICE |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `pyproject.toml`, `adapters/windows_ui/*` (future) |

**Description:** `pyproject.toml` lists `pywinauto>=0.6.8`; the venv has 0.6.9 and it imports. No source file imports it (only `pyproject.toml`). This is expected at task_001 (no UI code yet) and is **not dead code** — task_004 and task_020 depend on it. However, task_004 is a `dedicated_gate` that requires pywinauto UIA to read a VBE compile dialog; if the 0.6.9 UIA backend has issues on this Office build, the gate will discover it late (after tasks 003/002). A 10-minute read-only spike (open VBE, enumerate windows, read a dialog-less VBE caption) would de-risk task_004 without changing code.

**Recommendation:** Before task_004, record in the task_004 evidence: `pywinauto` version, UIA backend availability, and whether `Application.VBE.MainWindow.Caption` is readable on build 14332. If not readable, task_004 must record the fallback (`indeterminate`) rather than promise a compile result.

### ARCH-013: spec §7.1 "show Excel UI by default" vs PoC/worker hidden-window default — a policy decision must be recorded, not left to implementers

| Field | Value |
|---|---|
| **ID** | ARCH-013 |
| **Severity** | MEDIUM |
| **Type** | SPEC-DEVIATION |
| **Classification** | advisory |
| **Verification** | unverified (design comparison) |
| **Affected Modules** | `adapters/excel/com_worker.py` (task_010), `application/session_service.py` (task_012), `adapters/windows_ui/*` (task_004/020) |

**Description:** Spec §7.1 (spec_text.txt:320): "По умолчанию показывать Excel UI для диагностики VBE и модальных ошибок; режим скрытого окна допустим только в тех действиях, где не требуется UIA." Planner v3 §4.3 and §5.2 require the PoC to set `app.Visible = False` and `app.DisplayAlerts = False`. These are not contradictory (the PoC is not the product), but no task states how the *production* worker chooses `Visible`. task_010's goals never mention `Visible`/`DisplayAlerts`; task_012 says "working with copies in work/" and "restore temporary Application settings in finally", and spec §7.1 also says "Не отключать DisplayAlerts глобально на всё время сессии. Любые изменения Application state делать в context manager с восстановлением исходных значений."

If task_010 hardcodes `Visible=False` (natural for a background worker), then task_004/task_020 compile diagnostics via UIA may fail because VBE windows are not rendered, and the spec's "show UI by default" requirement is silently violated. If it hardcodes `Visible=True`, every `excel_open` pops an Excel window on the user's desktop, which may be undesirable for inspection but is what the spec asks for.

**Recommendation:** Resolve explicitly in task_010's planner/auditor gate (not in code): define a `visible` policy — e.g. `Visible=True` for `execute`/compile/macro sessions and `Visible=False` only for pure inventory actions that do not need UIA — and make it a parameter of `ExcelComWorker.start()` / `SessionService.open_workbook`, with `DisplayAlerts` restored via context manager. Record the decision in the task_010 design artifact so task_004/020 do not inherit an unstated assumption. This is spec-aligned either way; the defect would be *silence*.

### ARCH-014: plan task_030 api file list omits the `controls_*` tools required by its own goals

| Field | Value |
|---|---|
| **ID** | ARCH-014 |
| **Severity** | LOW |
| **Type** | SPEC-DEVIATION |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `.ai/plans/plan.md` (task_030), future `api/shape_tools.py` |

**Description:** Plan task_030 goals require `controls_list`, `control_get`, `control_set_on_action` (goal 1–2, `plan.md:1503-1505`), and spec §6.5 names `shapes_list, controls_list, control_get, control_set_on_action`. But task_030's `files:` block (`plan.md:1531-1535`) lists only `api/shape_tools.py` with `shapes_list`, `shape_get`, `shape_set_on_action`. The `controls_*` functions have no declared file target. This is a planning omission that will cause either (a) an undeclared file `api/control_tools.py`, or (b) `controls_*` squeezed into `shape_tools.py`, contradicting the spec's module naming (`shape_tools.py` vs `sheet_tools.py` vs `table_tools.py` are per-object modules).

**Recommendation:** When task_030 is planned, add `src/excel_vba_mcp/api/control_tools.py` (or explicitly state that `controls_*` live in `shape_tools.py` and update the spec cross-reference). Do not implement `controls_*` without a declared target; plan §6 rule 8 requires semantic targets to exist by name. This is a plan-text fix, not a code change.

### ARCH-015: execution log says "task_001 DONE ... tests passed" but repo-level test/lint/format are red because of the untracked PoC

| Field | Value |
|---|---|
| **ID** | ARCH-015 |
| **Severity** | LOW |
| **Type** | DOC-UPDATE |
| **Classification** | advisory |
| **Verification** | observed |
| **Affected Modules** | `.ai/plans/execution/execution-log.md` |

**Description:** The task_001 row records "`uv run pytest tests/test_smoke.py` passed; `ruff check`/`ruff format --check`/`mypy src` passed". Those commands were presumably run scoped to `src`/`test_smoke.py`. Mirroring the README development commands at repo scope fails today:
- `uv run pytest` → collection error (ARCH-001).
- `uv run ruff check src tests` → 56 F821 in the PoC.
- `uv run ruff format --check src tests` → 1 file would be reformatted (the PoC, trailing paren style).
- `uv run mypy src` → clean (12 files).

**Recommendation:** Update the execution log to state the exact scoped commands used for task_001 sign-off and note that repo-wide checks are red solely due to the untracked PoC fragment. When ARCH-001 is resolved, re-run the repo-wide commands and record the green result; do not mark a later task complete on a red repo-wide baseline.

---

## Architecture / risks

### STA worker design (task_010) vs spec §5.3 — assessment

**Alignment (confirmed design-level):** The plan's task_010 (`plan.md:508-560`) matches spec §5.3 point-for-point: single dedicated STA thread, `CoInitializeEx(COINIT_APARTMENTTHREADED)`, `DispatchEx("Excel.Application")`, serialized request queue, only DTOs returned, shutdown closes owned workbooks → `Quit` → release in `finally`. `ExcelComWorker(start/submit/shutdown)` + `ExcelAppAdapter` is the right module split, and it keeps COM out of `api/*` (spec §5.2 "MCP-обработчики не должны содержать COM-логику").

**Risks that must be designed for explicitly (in priority order):**

1. **Message pump starvation / reentrancy.** An STA thread that blocks in a synchronous COM call does not pump messages; Excel expects its STA to service cross-apartment calls. `pythoncom.PumpWaitingMessages()` inside a `while` loop is the planner-v3 approach, but calling it *inside* a COM call (via a callback/event sink) can reenter the worker. Task_010 must define exactly where pumping happens: (a) only while waiting for a queued item (safe), (b) during long operations via `IMessageFilter`/`CoWaitForMultipleHandles` (correct but more code), or (c) never (risk `RPC_E_CALL_REJECTED`/hang). Planner v3 §6.1 picks `PumpWaitingMessages`; task_010 must not silently change that without evidence.
2. **`RPC_E_CALL_REJECTED` (0x80010001) and `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A).** Spec §11 lists `COM_BUSY / RPC_E_CALL_REJECTED` with "повторять только после проверки busy state; не дублировать изменяющий вызов." Task_010 acceptance: "COM busy/rejected operations are not blindly replayed." A generic `retry` wrapper around `submit()` would be a **correctness bug** for mutations (duplicate `ListRows.Add`, duplicate `SaveAs`). The worker needs an operation classification (`read_only` vs `mutating`) and must never auto-retry the latter. This classification is also required by task_009 (read/mutate/execute), so the two should share one enum in `domain/statuses.py` (task_007) rather than each inventing one.
3. **Modal dialogs and the "no new COM operation" rule.** Spec §5.3/§8.3: while a VBE compile dialog is open, no COM operation may run; UIA reads/closes it first. The worker queue must therefore support an **exclusive "diagnostic hold"** state: after a COM call raises a modal-indicating error (or the UIA watcher signals a dialog), `submit()` must fail fast with `running/needs_recovery` (spec §9.4) instead of queueing. Task_010/011 must test this; task_020 (compile) depends on it.
4. **Shutdown ownership.** The worker owns the `Application` and every workbook it opened. On `shutdown()`, it must close only its own workbooks (`Close(SaveChanges=False)` unless a save was explicitly requested), `Quit()`, release references, `CoUninitialize()` on the same thread, and then signal completion with a bounded wait. Task_011 acceptance "shutdown closes only the Excel instance owned by the worker" is correct but needs the converse test: **a user's Excel instance must remain open and untouched**. The PoC (task_002) proves the `DispatchEx` half; task_011 must prove it for the *production* worker with a user instance running.
5. **Deadlock matrix to document:** (a) `submit()` from the event loop while the worker waits for UIA; (b) UIA call that blocks on the same STA; (c) `Quit()` while a workbook is dirty/`DisplayAlerts=True`; (d) client disconnect while a macro runs; (e) two `excel_open` calls on the same path (spec §5.3 lock: "lock на нормализованный путь + workbook_id"). Task_010's planner must enumerate these with the intended status (`running/needs_recovery`, `WORKBOOK_LOCKED`, etc.).
6. **Proxy lifetime / GC.** `del app` is not enough if a DTO or an exception traceback holds a proxy; task_010 must state that no proxy is stored in `SessionService` beyond the worker, and task_011 should test `VBA_PROJECT_LOCKED`-style failures do not leak references. Observed: after each audit probe, `EXCEL.exe` count was 0, so the basic `Quit` path is clean on this host.

### Security boundaries — assessment

- **Path policy (task_008):** design is correct and matches spec §10. Two gaps: (i) validation must be tied to the exact canonical value passed to COM (TOCTOU), (ii) UNC denial "by default" needs a config switch if a user legitimately stores projects on a share — spec says "запрет ... произвольных UNC-путей по умолчанию", so a config option is spec-compliant, but must default to deny.
- **Macro/approval policy (task_009):** design is correct and matches spec §10/§10.1 (inspect=`ForceDisable`, execute=`ByUI` immediately before `Open`, restore in `finally`; no `AutomationSecurity=Low`; no registry/AccessVBOM change; no VBA password bypass; audit redaction). The risk is **policy bypass via the test runner** (task_035): the declarative runner must not be able to express "open arbitrary path" or "run arbitrary procedure" — task_035 acceptance already says "неподдерживаемые actions/assertions отвергаются до запуска Excel operation", which is the right invariant; task_009 should expose a single `ApprovalPolicy.decide(operation, workbook_identity, action)` that task_035 reuses, not a second checker.
- **Snapshot model (task_014):** matches spec §9.1/§9.4. Risks: (i) `SaveCopyAs` on an open workbook does not include unsaved in-memory state — the snapshot is the *file*, so a mutation made but not saved before snapshotting is not captured; task_014 must define whether `snapshot_create` flushes or requires a prior save; (ii) `manifest.json` hash must be computed on the *copy*, not the source, and the source hash recorded separately; (iii) restore must not save the mutated copy (task_014 acceptance says this, good); (iv) snapshot is explicitly **not** a sandbox (spec §9.4) — task_015/035 must repeat this in evidence so no one claims rollback of DB/email/file side effects.
- **Trust Center / VBOM:** both plan tasks 003/009/012 correctly diagnose `ACCESS_VBOM_DISABLED` and never mutate registry/Trust Center. No deviation.
- **STDIO transport:** no HTTP/SSE endpoint is opened (`mcp.run(transport="stdio")` only), logs go to stderr (`logging.basicConfig(stream=sys.stderr)`), and the SDK's Windows stdio implementation additionally diverts fd 1 so stray `print()` cannot corrupt the wire (verified in `mcp/server/stdio.py`). No arbitrary COM/shell/Python-eval tool exists (only `server_health` today). Compliant with spec §10 "MCP transport".

### MCP SDK / runtime risks

- **Sync vs async tool execution** (ARCH-009) is the main structural risk for tasks 010–016.
- **`MCPServer` global + decorators**: tools must be registered at import time; task_012's `server.py` target "workbook_tools registration" must import `api/workbook_tools.py` from `server.py` (or a registration module) before `mcp.run`. No circular-import risk if `api/*` imports only `domain`/`application` types, never `server`.
- **Structured output**: `server_health` already returns `structuredContent` (verified). Later tools must return Pydantic-serializable DTOs (task_007) so the SDK can emit `structuredContent`; returning raw `dict[str, Any]` with COM-derived values would break serialization. Task_007's acceptance ("DTO сериализуются в JSON и не содержат COM objects") is the right gate.
- **Cancellation**: `anyio.to_thread.run_sync` without `abandon_on_cancel=True` shields the thread from cancellation (anyio docs); a cancelled `tools/call` will not stop the underlying Excel operation. Task_010/022 must define behaviour: cancellation returns to the client but the worker finishes the operation (and the session records `running/needs_recovery` if it was a macro). This should be documented, not discovered.

### Structural gaps summary (task_001 vs spec §5.1)

| Item | Status | Note |
|---|---|---|
| `pyproject.toml`, `uv.lock`, `README.md` | present | `uv.lock` uncommitted bump |
| `src/excel_vba_mcp/{__main__,server,config}.py` | present | config loader is a stub (ARCH-004) |
| `api/`, `application/`, `domain/`, `security/`, `storage/`, `adapters/{excel,windows_ui}/` | present, empty | correct for task_001; not dead code |
| `config/default.toml`, `config/macro-policy.example.toml` | **missing** | needed by tasks 008/009 |
| `tests/unit/`, `tests/test_data/` | **missing** | needed by plan `tests_to_run` |
| `projects/`, `work/`, `snapshots/`, `artifacts/` | present, empty | correct |
| `tests/integration_excel/` | present, untracked, **broken** | ARCH-001 |
| `tests/fixtures/` | present, empty | fixtures are task_003+ |

---

## Runtime recommendation

**Recommendation: keep Python 3.14.7, pin it explicitly, and document the deviation — do not provision 3.12.**

Reasoning:

1. **The validated environment is 3.14.** All runtime evidence in this audit (MCP handshake, tool call, sync-tool threading, `DispatchEx`, `CoInitializeEx`, `RPC_E_WRONG_THREAD`, pywin32 wheel `cp314`) was produced on 3.14.7. Switching to 3.12 now would invalidate that evidence and require re-running the task_002 PoC on a different interpreter, adding a variable to the critical path for no functional gain.
2. **3.14 is supported by every dependency that matters.** `mcp 2.3.0` declares `>=3.10` and has explicit `python_version >= '3.14'` markers; `pywin32 312` ships a `cp314` wheel with `pythoncom314.dll`; `pywinauto 0.6.9` imports; `pydantic 2.14.4` supports 3.14. The environment is GIL-enabled, so free-threading caveats do not apply.
3. **3.12 is available but not needed.** `uv python list` shows 3.12.16 already installed; `uv python pin 3.12 && uv sync` would work. But the execution log records an explicit decision ("do not install/pin 3.12"), and the plan/spec minimum is 3.11 — 3.14 satisfies the spec.
4. **The real risk is drift, not the minor version.** Today the repo claims 3.12 (README), targets `py311` (ruff/mypy), allows `>=3.11,<3.15` (pyproject), and runs 3.14 (venv) with no `.python-version`. That is the worst of both worlds. Pinning 3.14 and making the docs/linters agree is a **trivial** change (add `.python-version`, update README line 12, decide `target-version`), and it makes the task_002 evidence reproducible.

Actions (all small):
- Add `.python-version` with `3.14`.
- Update `README.md:12` to "Python 3.14 x64 (validated); 3.11–3.13 installable but unvalidated".
- Decide ruff/mypy target: `py314` if 3.14 is the only supported runtime; keep `py311` only if 3.11 compatibility is a stated goal, and add a comment.
- Commit `uv.lock` (pydantic 2.14.4 bump) so "reproduced from uv.lock" is true for a clean clone.
- In `docs/release-evidence.md` (task_037), record "validated on Python 3.14.7 / Excel 16.0.14332 / pywin32 312; other configurations untested" per plan task_037's rule against claiming untested support.

If a reviewer insists on 3.12 because the spec text names it: treat that as a **DOC-UPDATE** to the spec/plan (the code and SDK support 3.14; the spec's "рекомендуемый" matrix is not a hard requirement), not as a reason to redo the PoC.

---

## Integration / ordering notes

Dependency facts that later tasks rely on (from plan DAG `plan.md:40-60` and task `depends_on`):

1. **task_002 is the gate for task_003/004/005/006 and task_010.** No production COM code should be written before `go_no_go_r01_task002` is recorded. ARCH-002/ARCH-003 are the closing conditions.
2. **task_007 (DTOs/statuses) is the contract for everything.** `domain/models.py`, `errors.py`, `statuses.py`, `test_spec.py` must exist before task_008 (path policy returns `PATH_NOT_ALLOWED`), task_009 (read/mutate/execute enum), task_010 (`COM_BUSY`/`RPC_E_CALL_REJECTED` mapping), task_012 (session DTOs), task_014 (snapshot manifest), task_035 (test spec). Recommend task_007 be implemented immediately after ARCH-001 is cleared and **before** task_010, because both task_009 and task_010 need the shared operation-class enum (ARCH-009 risk 2).
3. **task_008 + task_009 before task_010/012.** `SessionService.open_workbook` (task_012) calls `PathPolicy.validate_path`, `MacroPolicy`/`ApprovalPolicy.decide`, and writes `AuditLogger` events. If task_010 is built first, it will either duplicate these checks or take a path/DTO it should not trust. The plan's DAG already orders 008/009 before 010/012; keep it.
4. **task_010 contract that task_011/012/014/016/020/022/024/026/028/030 all consume:** `ExcelComWorker.submit(operation) -> DTO` (awaitable, bounded, classified read/mutate/execute), `ExcelComWorker.start()/shutdown()`, and `ExcelAppAdapter` as the only COM surface. Any later task that needs a new COM call must add it to `app_adapter.py`/`workbook_adapter.py`/`vbide_adapter.py` etc., never to `api/*`. This is the single most important boundary to enforce in reviews.
5. **task_012 → task_014 → task_016 ordering.** `task_014` (snapshots) depends on `task_006` (PoC) + `task_012` (session); `task_016` (VBA read) depends on `task_003` (PoC) + `task_012` + `task_007`. `task_016` is read-only and can proceed in parallel with task_014, but both need task_012's `workbook_id`/session registry. Do not start task_016 against a raw `DispatchEx` app; it must go through the worker.
6. **Compile/macro chain.** task_020 (compile) depends on task_004 (PoC) + task_010 + task_016; task_022 (macro) depends on task_005 (PoC) + task_009 + task_016 + task_020. The "compile before macro" rule (spec §8.3, §7.3) must be enforced in `MacroService`, not only in the test runner.
7. **Task_032/034/035 are aggregation/hardening, not new COM paths.** task_032 explicitly says "не дублируя COM traversal в API handlers"; task_034 says "cleanup освобождает только принадлежащий серверу Excel instance"; task_035 says "Тест-кейс не получает прямого доступа к Python interpreter, произвольным путям, COM objects, shell". These are correct invariants; reviewers should check them as *negative* constraints (no new COM entry points, no new path bypass, no new eval).
8. **Ordering risk not in the DAG:** task_009's `AuditLogger` and task_034's `AuditLogger` target the same file (`security/audit.py`). Task_009 creates it; task_034 extends it. Whoever implements task_034 must not rewrite task_009's event schema. Similarly `domain/statuses.py` is created by task_007 and extended by task_034; `domain/errors.py` likewise. Recommend task_007 define the *full* Stage 1 status/error enum (spec §11 table has 16 codes) rather than a minimal set, to avoid churn.
9. **Fixture dependency:** tasks 003/004/005/006/014/016/020/022/024/026/028/030/032/035/036 all reference `tests/fixtures/*.xlsm` that do not exist. `CleanProject.xlsm` (task_003), `CompileErrorProject.xlsm` (task_004), `ControlsProject.xlsm` (task_030), `TableNamesProject.xlsm` (task_026/028), `LockedProject.xlsm` (task_016). These are on the critical path for *verification*, not implementation; creating them is a prerequisite for the first `dedicated_gate` that needs them (task_003). Do not let task_003's PoC become a fixture-creation task by accident — call it out explicitly.

---

## Final result

**FAIL**

The skeleton (task_001) is a sound, spec-aligned foundation: MCP 2.3.0 STDIO works, stdout is clean, logging goes to stderr, module layout matches spec §5.1, and the runtime (3.14.7) is compatible with every dependency including pywin32/Excel COM. The task_002 **design** package (researcher → auditor v1–v3 → planner v3 → validator v3) is rigorous and closes the auditor's D1–D9/N1–N5 findings; this audit found no blocking design defect in planner v3.

However, the audit cannot return PASS because:

1. **ARCH-001 (CRITICAL):** the only task_002 artifact is a non-importable, non-collectable fragment with 56 undefined names and no `test_*` functions; the repo-wide test/lint/format commands are red, and the PoC cannot produce the evidence the gate requires.
2. **ARCH-002 (HIGH):** the task_002 gate has design approval but **zero runtime evidence**; the plan's own rule forbids treating design approval as a Go, and downstream tasks 003–006/010 are blocked on this.
3. **ARCH-009 (HIGH):** the SDK runs sync tools on arbitrary worker threads, so the "single STA" invariant is not enforced by the current skeleton; without an explicit design constraint, task_012 can regress into per-call COM apartments.
4. **ARCH-008 (MEDIUM, mandatory consistency):** the repo simultaneously claims 3.12 (README), targets `py311` (ruff/mypy), allows `>=3.11,<3.15` (pyproject), and runs 3.14 with no pin; the lockfile bump is uncommitted.

None of these require redesign. The path to PASS is: finish or replace the PoC under planner v3 (ARCH-001) → run it on Excel 16.0.14332 and record `go_no_go_r01_task002` (ARCH-002) → record the sync-tool/STA constraint in task_010's design (ARCH-009) → pin 3.14 and commit the lock/docs (ARCH-008) → re-run `pytest`/`ruff`/`mypy` green. With those closed, the architecture is ready for task_002 execution and the task_007–task_016 sequence.

---

## Appendix A — Evidence commands (reproducible)

```
# Environment
uv run python -c "import sys,importlib.metadata as md; print(sys.version); print({n: md.version(n) for n in ['mcp','pydantic','pywin32','pywinauto','pytest','ruff','mypy']})"
# 3.14.7; mcp 2.3.0; pydantic 2.14.4; pywin32 312; pywinauto 0.6.9; pytest 9.2.0; ruff 0.17.5; mypy 2.4.1

# PoC non-collectable
uv run pytest --collect-only -q
# ERROR tests\integration_excel\test_excel_instance_poc.py - NameError: name 'dataclass' is not defined

# PoC undefined names
uv run ruff check src tests --output-format concise
# 56 F821 in tests/integration_excel/test_excel_instance_poc.py

# Format / types
uv run ruff format --check src tests   # 1 file would be reformatted (PoC)
uv run mypy src                        # Success: no issues found in 12 source files

# STDIO protocol (initialize / tools/list / tools/call)
# serverInfo.name="excel-vba-mcp"; tools=[server_health]; structuredContent={"status":"ok","server":"excel-vba-mcp"}
# stderr: "INFO excel_vba_mcp.server: Starting Excel VBA MCP server (STDIO)"; stdout: 2 JSON-RPC lines only

# STDIO shutdown
# stdin EOF -> rc 0 in ~0.17 s; SIGTERM idle -> rc 1; SIGTERM during 20 s sync tool -> alive >10 s

# COM
# DispatchEx main thread (implicit init) -> Version 16.0, Build 14332.0; no orphan EXCEL.exe
# CoInitializeEx(APARTMENTTHREADED) on dedicated thread -> OK; CoUninitialize same thread -> OK
# cross-thread without CoInitialize -> com_error(-2147221008) CO_E_NOTINITIALIZED
# cross-thread with CoInitialize (STA or MTA) -> com_error(hresult=0x8001010E) RPC_E_WRONG_THREAD
```

## Appendix B — File:line index

| Finding | Primary evidence |
|---|---|
| ARCH-001 | `tests/integration_excel/test_excel_instance_poc.py:1-262` (fragment; `dataclass` at 63/71/225; no imports; no `test_*`) |
| ARCH-002 | `.ai/plans/execution/execution-log.md` task_002 row; `plan.md:118-167`; `.ai/tmp/` (no artifact) |
| ARCH-003 | `task_002_planner_v3.md` §7, §8.2; observed `0x800401F0` vs `0x8001010E` |
| ARCH-004 | `src/excel_vba_mcp/config.py:1-5,24-31`; `src/excel_vba_mcp/server.py:17-23` |
| ARCH-005 | `src/excel_vba_mcp/server.py:26`; `mcp/server/mcpserver/server.py` lifespan param; observed shutdown matrix |
| ARCH-006 | `src/excel_vba_mcp/**` (empty `__init__.py` only); missing `config/`, `tests/unit/`, `tests/test_data/` |
| ARCH-007 | `plan.md:413-506` (tasks 008/009), `plan.md:707-763` (task 014); spec §9–§10 |
| ARCH-008 | `pyproject.toml` (`requires-python`, `target-version`, `python_version`); `README.md:12`; `.venv/pyvenv.cfg`; `uv.lock` header; no `.python-version` |
| ARCH-009 | `mcp/server/mcpserver/utilities/func_metadata.py:149-164`; observed `tools/list` answered during 20 s sync tool |
| ARCH-010 | observed `DispatchEx`/`CoInitializeEx`/RPC probe results; spec_text.txt:147 |
| ARCH-011 | `server.py:17-26`; `__main__.py:8-13` |
| ARCH-012 | `pyproject.toml` pywinauto dep; no source import |
| ARCH-013 | spec_text.txt:320; `task_002_planner_v3.md` §4.3/§5.2; task_010 goals |
| ARCH-014 | `plan.md:1503-1505` vs `plan.md:1531-1535`; spec §6.5 |
| ARCH-015 | `execution-log.md` task_001 row; observed repo-wide command results |