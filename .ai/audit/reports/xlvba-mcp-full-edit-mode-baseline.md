# Excel VBA MCP — Full Edit Mode Baseline Audit

**Report type:** Phase 0 baseline audit (read-only; no source code changed)
**Task:** `FULL-EDIT-MODE-ORCHESTRATION-001` (Auditor deliverable, §8.2)
**Repository:** `C:\py_exp\xlvba_mcp` (branch `master`, HEAD `8c22f11` — "docs: align with full-access default, macro classification, and error taxonomy (task_b8_docs)")
**Audit date:** 2026-11-22
**Auditor role:** independent read-only architecture/security/code review + live MCP probing
**Verdict of this document:** baseline only. It records what the product *actually does today*, before any corrective work of this orchestration task.

> Scope discipline: no production file was modified. All live Excel work ran on disposable copies inside a fresh
> `%TEMP%\xlvba_audit*_*` root. The repository fixture `tests/fixtures/FullAccessWorkflow.xlsm` was hashed before and
> after every probe and remained byte-identical (`6B905837B8DE…A19AC1D4D9`). No registry, Trust Center,
> `AutomationSecurity = LOW`, `SendKeys`, `SetForegroundWindow` or dialog dismissal was used.

---

## 1. Environment facts

| Item | Value | How observed |
|---|---|---|
| OS | Windows 10, build 10.0.17763 (64-bit), `win32` | `platform.platform()` in the venv interpreter; process runs in the interactive console session (no service / no separate session) |
| Python (runtime) | 3.14.7 (tags/v3.14.7, MSC v.1944, AMD64) | `.venv\Scripts\python.exe -c "import sys; print(sys.version)"` |
| Declared floor | `>=3.11,<3.15` | `pyproject.toml` |
| Excel | Microsoft Excel **16.0**, build **14332.0**, `Visible=False` on creation | `win32com.client.DispatchEx("Excel.Application")` → `app.Version` / `app.Build` / `app.Visible` |
| Excel bitness / Office path | `C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE` | live `vba_compile` diagnostics reference path |
| MCP SDK | **mcp 2.3.0** | `importlib.metadata.version("mcp")` |
| Protocol version negotiated | `2025-11-25` | live `initialize` result |
| Server identity | name `excel-vba-mcp`, version `0.1.0` | live `initialize` result `serverInfo` |
| Pydantic | 2.14.4 | venv |
| `pywinauto` | installed (0.6.8 floor) but **not imported anywhere in `src/`** | `pywinauto` grep over `src/` → 0 hits; still declared in `pyproject.toml` |
| Excel launch model | `win32com.client.DispatchEx("Excel.Application")` on a dedicated **single-STA** worker thread; `app.Visible = False` set immediately after `DispatchEx`, before any workbook is touched; no code path ever sets `Visible = True` | `adapters/excel/app_adapter.py::ExcelAppAdapter.create_hidden`, `assert_non_intrusive`; `adapters/excel/com_worker.py::ExcelComWorker.start/_run` |
| VBOM trust assumption | "Trust access to the VBA project object model" must be enabled **manually** per Windows profile; the server never reads/writes the registry or Trust Center and only *diagnoses* state (`ACCESS_VBOM_DISABLED`) | `security/macro_policy.py::diagnose_vbom`; live `vba_compile` reported `vbom_available: true`, `protection: 0` — VBOM is available in this profile |
| Macro-security handling | per-open `AutomationSecurity`: `INSPECT → msoAutomationSecurityForceDisable (3)`, `EXECUTE → msoAutomationSecurityByUI (2)`; `LOW (1)` is never selected and is restored on every exit path | `security/macro_policy.py::automation_security_for`, `app_adapter.automation_security` |
| Working copy model | original → `shutil.copy2` into `<work_dir>/<workbook_id>.xlsm`; only the copy is opened; `excel_close` always deletes the working copy | `application/session_service.py::_materialize_working_copy` / `_cleanup_working_copy` |

**Orphan observation (environment hygiene):** one hidden automation Excel instance
(`EXCEL.EXE` PID 3848, `/automation -Embedding`, created 2026-11-22 11:57:12) predates this audit and was left
behind by an earlier run. It is not owned by this session and was not touched. It is evidence that an owned
instance can survive an abruptly-terminated server process (see §8, F-8).

---

## 2. Actual live MCP surface (derived from code **and** confirmed by launching the server)

### 2.1 Commands used

All probes used the venv interpreter and the same pattern as
`tests/integration_excel/test_mcp_full_access_workflow.py` (a real `stdio_client` + `ClientSession` handshake over
`python -m excel_vba_mcp`).

```powershell
# Discovery (tools / resources / templates / prompts / health / error surfacing)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" `
  "C:\py_exp\xlvba_mcp\.ai\tmp\audit_live_tool_catalogue.py" `
  1> ".ai\tmp\audit_probe.json" 2> ".ai\tmp\audit_probe_stderr.log"

# Functional probe on a disposable copy (full mode, restricted mode, read-only session)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_functional.py" `
  1> ".ai\tmp\audit_functional.json" 2> ".ai\tmp\audit_functional_stderr.log"

# Full-payload evidence (values read back after each mutation/macro)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_evidence.py" `
  1> ".ai\tmp\audit_evidence.json" 2> ".ai\tmp\audit_evidence_stderr.log"

# Post-restore session mode + restricted-mode ungated tools
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_restore.py" `
  1> ".ai\tmp\audit_restore.json" 2> ".ai\tmp\audit_restore_stderr.log"

# Foreground-window / owned-window observation (read-only Win32)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_focus.py" `
  1> ".ai\tmp\audit_focus.json" 2> ".ai\tmp\audit_focus_stderr.log"

# Owned-Excel loss vs excel_health truthfulness (R-17 reproduction)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_stale.py" `
  1> ".ai\tmp\audit_stale.json" 2> ".ai\tmp\audit_stale_stderr.log"

# Snapshot -> save -> close -> restore persistence (+ post-restore mode)
& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" ".ai\tmp\audit_live_persistence.py" `
  1> ".ai\tmp\audit_persistence.json" 2> ".ai\tmp\audit_persistence_stderr.log"
```

Each probe launches the server with
`StdioServerParameters(command=sys.executable, args=["-m","excel_vba_mcp"], cwd=<temp base>, env=...)`
where `env` scopes `EXCEL_VBA_MCP_ROOTS / _WORK_DIR / _SNAPSHOT_DIR / _ARTIFACT_DIR / _RUN_DIR / _AUDIT_LOG`
to a fresh `%TEMP%\xlvba_audit*_*` directory, and **deliberately does not override**
`USERPROFILE / HOMEDRIVE / HOMEPATH / LOCALAPPDATA / APPDATA` (VBOM trust is per profile).

### 2.2 Counts (live, not assumed)

| Surface | Count | Evidence |
|---|---|---|
| **Tools** | **88** | `len((await session.list_tools()).tools) == 88` |
| Static resources (`resources`) | **0** | `list_resources()` → `resources: []` |
| Resource **templates** | **4** | `list_resource_templates()` |
| Prompts | **0** | `list_prompts()` → empty |

Code-side reconciliation: `server.py::_register_tools` performs **87** `server.add_tool(...)` registrations, plus
`server_health` registered by the `@mcp.tool()` decorator = **88**. This matches the live count and matches the
documentation (`README.md`, `docs/01-overview.md`, `docs/04-architecture.md`, `docs/05-mcp-tools-reference.md` all
state 88 tools / 4 resources).

> **Correction to the prior baseline.** `01-plan-corrective-full-access-execution.md` (C-audit-14) records
> "86 `add_tool` + `server_health` = 87 tools". That number is now **stale**: B5 added
> `macro_classify_background` and B8 updated the docs, so the live count is **88**. The B5 execution note (W1)
> that flagged "docs already say 87" is resolved by commit `8c22f11`.

### 2.3 Live tool catalogue (88 names, alphabetical)

```
control_get, control_set_on_action, controls_list,
excel_close, excel_get_status, excel_health, excel_list_workbooks, excel_open, excel_save,
macro_classify_background, macro_get_last_result, macro_list, macro_run,
name_create, name_delete, name_get, name_update, names_list,
pivot_cache_get, pivot_cache_refresh, pivot_field_apply_filter, pivot_field_clear_filters,
pivot_field_get_filters, pivot_field_set_aggregation, pivot_field_set_layout, pivot_field_set_position,
pivot_field_set_properties, pivot_fields_list, pivot_table_get, pivot_table_refresh, pivot_tables_list,
range_copy, range_delete, range_get, range_get_conditional_formatting, range_get_format,
range_get_validation, range_insert, range_merge, range_read_formulas, range_read_values,
range_set_conditional_formatting, range_set_format, range_set_validation, range_unmerge,
range_write_formulas, range_write_values,
server_health,
shape_get, shape_set_on_action, shapes_list,
sheet_add, sheet_copy, sheet_delete, sheet_get, sheet_rename, sheet_set_visibility, sheets_list,
snapshot_create, snapshot_diff, snapshot_list, snapshot_restore,
table_add_column, table_append_rows, table_create, table_delete_column, table_delete_rows, table_get,
table_insert_row, table_read_rows, table_resize, table_update_rows, tables_list,
test_get_run, test_run_suite, test_suite_run,
vba_compile, vba_export_project, vba_get_compile_diagnostics, vba_get_references, vba_insert_code,
vba_list_components, vba_list_procedures, vba_project_info, vba_read_component,
vba_replace_component_source, vba_search_code,
workbook_get_model
```

### 2.4 Resources and resource templates (read-only, non-tool surface)

All four are **URI templates** (no static resource URIs are registered):

| Name | URI template | MIME | Handler |
|---|---|---|---|
| `workbook_model` | `excel://workbooks/{workbook_id}/model` | `application/json` | `server.py::workbook_model_resource` → `WorkbookModelService.get_workbook_model` |
| `workbook_vba_components` | `excel://workbooks/{workbook_id}/vba/components` | `application/json` | `server.py::workbook_vba_components_resource` → `VBAProjectService.list_components` (reports `vba_access` instead of failing) |
| `workbook_vba_component_source` | `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `text/plain` | `server.py::workbook_vba_component_source_resource` → `VBAProjectService.read_component` |
| `test_run_summary` | `excel://test-runs/{run_id}` | `application/json` | `server.py::test_run_resource` → `TestRunner.get_run` |

**Prompts:** none.

---

## 3. Configuration defaults

| Question | Answer | Evidence |
|---|---|---|
| Effective access mode when `EXCEL_VBA_MCP_ACCESS_MODE` is unset | **`full`** (full access). | `config.py::ServerConfig.access_mode` default `AccessMode.FULL`; `load_config()` only sets the key when the env var is present. Live: a server started with no `EXCEL_VBA_MCP_ACCESS_MODE` returned `excel_get_status.access_mode == "full"` and `excel_open` succeeded without any grant. |
| Valid values / invalid handling | `full` \| `restricted`; any other value raises `pydantic.ValidationError` at lifespan entry (`server.py::lifespan` logs and re-raises → process exits). | `config.py`, `domain/statuses.py::AccessMode`; `server.py::lifespan` |
| Is `config.py` wired into `server.py`? | **Yes (fixed in B2).** `server.py::lifespan` calls `load_config()`; `_build_runtime(config)` constructs the single `ApprovalPolicy(mode=config.server.access_mode)` and audits `server_startup` with `access_mode`. | `server.py::_build_runtime`, `server.py::lifespan`, `security/approval.py::ApprovalPolicy.__init__` |
| Is the mode observable by a client? | Yes — `SessionStatus.access_mode` (additive field, `domain/models.py`) is returned by `excel_get_status`. | live `excel_get_status` → `"access_mode": "full"` |
| Default roots when `EXCEL_VBA_MCP_ROOTS` is unset | `Path.cwd()` **+** `tempfile.gettempdir()` (both registered as `root_0`, `root_1`), plus the artifact root under the artifact logical name. | `server.py::_configured_roots`. Live: with `EXCEL_VBA_MCP_ROOTS` **unset**, `excel_open` on a copy under `%TEMP%` succeeded and a traversal path (`…\..\..\audit_evil.xlsm`) was rejected `PATH_NOT_ALLOWED`. |
| Artifact / work / snapshot / run / audit dirs when their env vars are unset | `%TEMP%\excel_vba_mcp\{artifacts,work,snapshots,runs,audit.jsonl}` | `server.py::_default_*_dir`, `_default_audit_path` |
| Default `excel_open` | `mode = SessionMode.EXECUTE`, `read_only = False` → **writable, execution-capable** session. | `api/workbook_tools.py::excel_open` signature; live result `{"mode":"execute","read_only":false}` |
| Server-level defaults for `ApprovalPolicy` when constructed directly (no mode) | `AccessMode.RESTRICTED` (deny-by-default) — only the server passes `FULL`. | `security/approval.py::ApprovalPolicy.__init__` |
| Grants / allowlist over MCP | **None.** No MCP tool grants approvals or manages an allowlist; `ApprovalPolicy.grant()` / `AllowlistEntry` are library-only. | grep for `grant|allowlist` in `src/excel_vba_mcp/api/*` → only docstrings |

---

## 4. Call-path traces (module → class → method → guard → adapter → COM)

All Excel/VBA/COM access is serialized on one `ExcelComWorker` STA thread; tool handlers are thin and never touch
COM (ARCH-009). The generic shape is:

```
MCP client
  → mcp.server.mcpserver.MCPServer.call_tool                (server.py::ExcelVbaMcpServer.call_tool — structured-error remap)
    → excel_vba_mcp.api.<domain>_tools.<tool_fn>            (thin handler; validates args; reads a module-global service)
      → excel_vba_mcp.application.<domain>_service.<method> (guard chain, audit)
        → ExcelComWorker.submit(fn, timeout, section, op_class) (single-STA queue; READ_ONLY retried, MUTATING/EXECUTING never replayed)
          → adapters/excel/<x>_adapter.py                   (COM object model)
            → Excel.Application (hidden, owned, via ExcelAppAdapter)
```

### 4.1 Workbook / session

| Tool(s) | Path |
|---|---|
| `excel_open` | `api/workbook_tools.py::excel_open` → `application/session_service.py::SessionService.open_workbook`: `_require_accepting_new_work` → `PathPolicy.validate_path(path, MUTATING)` → **approval** (`INSPECT+read_only → READ_ONLY`; `EXECUTE → EXECUTING`; else `MUTATING`) via `_require_approval` → per-canonical-path `asyncio.Lock` held for the workbook lifetime → `_materialize_working_copy` (`shutil.copy2` into `work/`) → `_open_on_worker` → `ExcelComWorker.submit(..., section="open", op_class=MUTATING)` → `app_adapter.open_workbook` (per-open `automation_security` context) |
| `excel_close` | `SessionService.close_workbook`: `_require_record` → **approval** (`MUTATING` only when `save_changes=True`) → `worker.submit(section="close")` → `app_adapter.close_workbook` → `_cleanup_working_copy` (deletes the copy, **always**) |
| `excel_save` | `SessionService.save_workbook`: `_require_record` → **approval `MUTATING`** → `worker.submit(section="save", op_class=MUTATING)` → `app_adapter.save` |
| `excel_get_status`, `excel_list_workbooks` | `SessionService.get_status` / `list_workbooks` — pure bookkeeping, **no COM** |
| `excel_health` | `api/health_tools.py::excel_health` → `SessionService.worker_status` + `session.accepting`; **no COM submit** (deliberate: a health check must not block on the hang it reports) |
| `server_health` | `server.py::server_health` — literal dict, no COM |

### 4.2 Snapshots

| Tool(s) | Path |
|---|---|
| `snapshot_create` | `api/test_tools.py::snapshot_create` → `application/snapshot_service.py::SnapshotService.create_snapshot` → open branch `worker.submit(..., op_class=READ_ONLY)` + `SaveCopyAs` (or file copy for closed books) → `SnapshotStore` (manifest + SHA-256). **No approval gate, no writable-session requirement** (audited `read_only`). |
| `snapshot_list`, `snapshot_diff` | `SnapshotService.list_snapshots` / `diff_snapshots` — filesystem only |
| `snapshot_restore` | `api/test_tools.py::snapshot_restore` → `SnapshotService.restore_snapshot`: `SnapshotStore.load` → optional `SessionService.close_workbook(save_changes=False)` → **`SessionService.open_workbook(snapshot_file, mode=SessionMode.INSPECT, read_only=False)`** (which is where the only approval check happens, classified `MUTATING`) → registers the manifest as the new workbook's latest snapshot. Explicitly documented as having **no approval gate of its own** (R-19). |

### 4.3 Sheets / Tables / Names / Shapes / Controls / Ranges / Pivots

Uniform guard chain, implemented per service as `_require_writable` → `_require_snapshot` → `_require_approval`:

| Service (module) | `_require_writable` | `_require_snapshot` | `_require_approval` |
|---|---|---|---|
| `sheets_service.py` | `WORKBOOK_READ_ONLY` | `SnapshotService.require_snapshot` → `SNAPSHOT_FAILED` | `ApprovalPolicy.authorize_vba_write` → `ApprovalDeniedError` (`APPROVAL_DENIED`) |
| `tables_service.py` | idem | idem | idem |
| `ranges_service.py` | idem | idem | idem |
| `names_service.py` | idem | idem | idem |
| `shapes_service.py` | idem | idem | idem |
| `controls_service.py` | idem | idem | idem |
| `pivot_tables_service.py` | idem (injected `approval: ApprovalPolicy \| None` → `APPROVAL_DENIED` when `None`) | idem | idem |
| `vba_project_service.py` | idem | `_require_snapshot_service` + `SnapshotService.require_snapshot` | `_require_approval_service` (→ `ApprovalDeniedError` when unwired) + `authorize_vba_write` |

After the guard chain each service submits one non-retried `MUTATING` operation to the worker, then (for range /
pivot / table mutations) performs a **postcondition re-read** and records an audit entry.

Read-only tools (`*_list`, `*_get`, `range_read_*`, `table_read_rows`, `range_get_*`, `pivot_*_get/list`,
`pivot_field_get_filters`) submit `op_class=READ_ONLY` (safe to retry after `RPC_E_CALL_REJECTED`).

### 4.4 VBA project / compile

| Tool(s) | Path |
|---|---|
| `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_get_references`, `vba_search_code` | `api/vba_tools.py` → `application/vba_project_service.py` (`_submit`, `READ_ONLY`) → `adapters/excel/vbide_adapter.py` (VBIDE `VBProject.VBComponents`, `CodeModule`, `References`). VBOM/protection failures surface `ACCESS_VBOM_DISABLED` / `VBA_PROJECT_LOCKED`. |
| `vba_replace_component_source` | `_require_writable` → `_require_snapshot_service` → `require_snapshot` → **pre-check read** (`_submit`, READ_ONLY) → `_require_replaceable` (Standard/Class only) → `_require_revision` (stale hash) → `authorize_vba_write` → `_submit_mutation` (`MUTATING`, export-before/after, audit). |
| `vba_insert_code` | same chain with `insert_code` semantics (`declarations` / `before_procedure:` / `after_procedure:`). |
| `vba_export_project` | `VBAProjectService.export_project` — **no approval, no writable check**; writes `.bas/.cls/.frm+.frx` to a server-chosen `%TEMP%\excel_vba_mcp_vba_export_*` directory, or to a caller `out_dir` validated by `PathPolicy` against the artifact roots. |
| `vba_compile` | `api/compile_tools.py::vba_compile` → `application/compile_service.py::CompileService.compile_project` → read-only CommandBar discovery on the STA worker (`op_class=EXECUTING`, **no approval gate**) → on this build returns `status="indeterminate"`, `reason="NON_ACTIVATING_COMPILE_UNAVAILABLE"` (no `CommandBarControl.Execute`, UIA expansion refused). |
| `vba_get_compile_diagnostics` | returns the cached diagnostics; no compile, no COM mutation. |

### 4.5 Macros

| Tool(s) | Path |
|---|---|
| `macro_list` | `api/macro_tools.py::macro_list` → `application/macro_service.py::MacroService.list_macros` → `_collect_project_state` (one `READ_ONLY` STA scan returning **both** the callable inventory and the live `project_source_hash`) → annotate each public procedure with `BackgroundSafetyRegistry.is_safe(workbook_sha256, project_source_hash, qualified)` |
| `macro_classify_background` | `MacroService.classify_background`: same single scan → `_resolve_procedure` (`MACRO_NOT_FOUND`) → optional `expected_project_source_hash` check (`STALE_TARGET`) → **R13 veto** (`INTERACTIVE_UI_NOT_SUPPORTED`) when `background_safe=True` and (heuristic procedure **or** interactive token found) → `BackgroundSafetyRegistry.classify(...)` (in-memory) → audit. **Ungated** by design (bookkeeping, not the security boundary). |
| `macro_run` | `MacroService.run_macro`, in order: (1) `_require_execute_context` (`MACRO_BLOCKED` if `mode != EXECUTE`; `WORKBOOK_READ_ONLY` if read-only) → (2) `_collect_project_state` + `_resolve_procedure` (`NOT_FOUND`) → (3) `ApprovalPolicy.check(EXECUTING)` → `APPROVAL_DENIED` → (4) `BackgroundSafetyRegistry.is_safe` → `BACKGROUND_SAFETY_NOT_VERIFIED` / `INTERACTIVE_UI_NOT_SUPPORTED` → (5) R13 token veto + heuristic fail-closed → `INTERACTIVE_UI_NOT_SUPPORTED` → (6) single `worker.submit(MacroRunner(app).run("'name'!Module.Proc", args), section="execute", op_class=EXECUTING)`; timeout → `_handle_hung` (`EXCEL_HUNG`, `NEEDS_RECOVERY`). |
| `macro_get_last_result` | returns `MacroService._last[workbook_id]`, no COM. |

### 4.6 Test runner

| Tool(s) | Path |
|---|---|
| `test_run_suite` / `test_suite_run` | `api/test_tools.py` → `application/test_runner.py::TestRunner.run_suite`: Pydantic-validated closed suite → one immutable baseline snapshot → per-case fresh disposable copy opened `EXECUTE`/writable → each `TestStep` dispatched to the **real** services (so writable/snapshot/approval guards apply per step) → `background_safe` steps recorded through `MacroService.classify_background` before `macro_run` → artifacts + `runs/` persistence. Compile failure blocks later `macro_run` steps; `NEEDS_RECOVERY` aborts. |
| `test_get_run` | `RunRepository` read; no Excel. |

### 4.7 Server error boundary

`server.py::ExcelVbaMcpServer.call_tool` wraps `MCPServer.call_tool`, and on `ToolError` walks the `__cause__`
chain (`_find_domain_cause`, cycle-safe). If an `ExcelVbaMcpError` is found it re-raises
`ToolError(json.dumps(cause.to_dict()))`, so the client receives a structured payload
(`code`, `message`, `hresult`, `hresult_hex`, `context`). Anything else is re-raised unchanged.

---

## 5. Tool classification matrix (88 tools)

Legend for **Precondition**: `RO` = read-only path (no approval); `W` = writable session (`WORKBOOK_READ_ONLY`);
`S` = successful snapshot required (`SNAPSHOT_FAILED`); `A` = approval policy consulted (full mode auto-allows;
restricted mode denies with `APPROVAL_DENIED`); `BG` = background-safe classification required; `R13` =
interactive-UI/heuristic veto; `--` = no guard.

Denial codes: `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `MACRO_BLOCKED`,
`BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`, `EXCEL_HUNG`, `COM_BUSY`, `STALE_TARGET`.

### 5.1 READ_ONLY — 37 tools (no mutation of workbook or config)

| Tool | Precondition | Denial code | Notes |
|---|---|---|---|
| `excel_get_status` | `--` | — | no COM; exposes `access_mode` |
| `excel_list_workbooks` | `--` | — | bookkeeping only |
| `workbook_get_model` | `--` | — | aggregates sections; `vba_access` instead of failing |
| `snapshot_list` | `--` | — | filesystem |
| `snapshot_diff` | `--` | `SNAPSHOT_FAILED` | |
| `sheets_list`, `sheet_get` | `--` | — | |
| `tables_list`, `table_get`, `table_read_rows` | `--` | — | |
| `range_get`, `range_read_values`, `range_read_formulas`, `range_get_format`, `range_get_validation`, `range_get_conditional_formatting` | `--` | — | bounded windows |
| `names_list`, `name_get` | `--` | — | |
| `shapes_list`, `shape_get` | `--` | — | |
| `controls_list`, `control_get` | `--` | — | |
| `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_field_get_filters` | `--` | — | never materializes pivot output |
| `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_get_references`, `vba_search_code` | `--` (VBOM must be available) | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | |
| `vba_get_compile_diagnostics` | `--` | — | returns cache; `is_compile=false` |
| `macro_list` | `--` | `ACCESS_VBOM_DISABLED` / `EXCEL_HUNG` | annotates `background_safe` |
| `macro_get_last_result` | `--` | — | in-memory cache |
| `test_get_run` | `--` | — | run repository only |

### 5.2 MUTATING — 42 tools

| Tool(s) | Precondition | Denial code |
|---|---|---|
| `excel_save` | `A` | `APPROVAL_DENIED` |
| `snapshot_create` | `--` (filesystem write; no approval, no writable requirement) | `SNAPSHOT_FAILED` on failure |
| `snapshot_restore` | `A` (inherited from the internal `open_workbook`, classed `MUTATING`); no own gate | `APPROVAL_DENIED`, `SNAPSHOT_FAILED` |
| `vba_export_project` | `--` (ungated filesystem write; `out_dir` path-policy checked when supplied) | `PATH_NOT_ALLOWED`, `ACCESS_VBOM_DISABLED` |
| `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` | `W` + `S` + `A` | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED` |
| `table_create`, `table_add_column`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_delete_column`, `table_resize` | `W` + `S` + `A` | idem |
| `name_create`, `name_update`, `name_delete` | `W` + `S` + `A` | idem |
| `shape_set_on_action`, `control_set_on_action` | `W` + `S` + `A` | idem |
| `range_write_values`, `range_write_formulas`, `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`, `range_set_validation`, `range_set_conditional_formatting` | `W` + `S` + `A` (+ postcondition re-read) | idem |
| `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_apply_filter`, `pivot_field_clear_filters` | `W` + `S` + `A` (+ postcondition re-read) | idem, `UNSUPPORTED_PIVOT_SOURCE`, `STALE_TARGET` |
| `vba_replace_component_source`, `vba_insert_code` | `W` + `S` + `A` + revision check | idem, `COM_BUSY` (revision conflict — see §6.3) |

### 5.3 EXECUTING — 4 tools

| Tool | Precondition | Denial code |
|---|---|---|
| `macro_run` | execute-mode + writable + `A` + `BG` + `R13` | `MACRO_BLOCKED`, `WORKBOOK_READ_ONLY`, `APPROVAL_DENIED`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`; `EXCEL_HUNG` on timeout |
| `test_run_suite`, `test_suite_run` | no entry-level gate; inner steps apply the service guards | per-step `APPROVAL_DENIED` / `WORKBOOK_READ_ONLY` / `SNAPSHOT_FAILED` / `BACKGROUND_SAFETY_NOT_VERIFIED`; run `blocked` |
| `vba_compile` | none (no approval); submits `op_class=EXECUTING`; **never mutates** | returns `indeterminate`, never raises for compile state |

### 5.4 MIXED — 2 tools

| Tool | Modes | Precondition | Denial code |
|---|---|---|---|
| `excel_open` | read/write | `A` (`INSPECT+read_only → READ_ONLY`; `EXECUTE → EXECUTING`; else `MUTATING`); `PathPolicy` | `PATH_NOT_ALLOWED`, `APPROVAL_DENIED`, `EXCEL_HUNG`, already-open guard |
| `excel_close` | read (no save) / write (save) | `A` only when `save_changes=True` | `APPROVAL_DENIED`, `EXCEL_HUNG` |

### 5.5 ADMIN-DIAGNOSTIC / bookkeeping — 3 tools

| Tool | Category | Precondition | Notes |
|---|---|---|---|
| `excel_health`, `server_health` | ADMIN-DIAGNOSTIC | none | never touch COM |
| `macro_classify_background` | **MUTATING bookkeeping** (in-memory registry) | none (ungated) | `MACRO_NOT_FOUND`, `STALE_TARGET`, `INTERACTIVE_UI_NOT_SUPPORTED` (R13 veto) |

**Totals:** 37 READ_ONLY (§5.1) + 42 MUTATING (§5.2) + 4 EXECUTING (§5.3) + 2 MIXED (§5.4) + 3
ADMIN/bookkeeping (§5.5) = **88** ✔ (matches the live `list_tools` count in §2.2)

---

## 6. Error taxonomy

### 6.1 `ErrorCode` members (complete, `domain/errors.py`)

Stage-1 closed set: `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `MISSING_REFERENCE`, `COM_BUSY`,
`RPC_E_CALL_REJECTED`, `COMPILE_ERROR`, `COMPILE_INDETERMINATE`, `MACRO_BLOCKED`, `MACRO_NOT_FOUND`,
`MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `EXCEL_HUNG`, `SNAPSHOT_FAILED`, `WORKBOOK_READ_ONLY`, `PATH_NOT_ALLOWED`.

Phase-2 additive: `INVALID_RANGE`, `PROTECTED_TARGET`, `UNSUPPORTED_PIVOT_SOURCE`, `STALE_TARGET`,
`AMBIGUOUS_FILTER`, `REFRESH_FAILED`, `MODAL_UI_BLOCKED`, `LISTOBJECT_INTERSECTION`.

Task_b4/b5 additive: `APPROVAL_DENIED`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED`.
(Both `BACKGROUND_*` codes are **result codes only** — deliberately absent from `_ERROR_BY_CODE`; they appear in
`MacroRunResult.error_code`.)

### 6.2 Which paths raise which code

| Code | Raised by |
|---|---|
| `ACCESS_VBOM_DISABLED` | `diagnose_vbom` / VBIDE adapter when `VBProject` is inaccessible |
| `VBA_PROJECT_LOCKED` | `diagnose_protection` (protection == 1) |
| `MISSING_REFERENCE` | `error_from_hresult(0x80040154)` |
| `COM_BUSY` | `ExcelVbaMcpError.default_code` (generic/unclassified errors); HRESULT `0x8001010A`; worker busy; **and** VBA revision conflict / non-replaceable component (see §6.3) |
| `RPC_E_CALL_REJECTED` | HRESULT `0x80010001` |
| `COMPILE_ERROR` / `COMPILE_INDETERMINATE` | compile service (the latter is the live outcome on this build) |
| `MACRO_BLOCKED` | `_require_execute_context` (non-execute session); Office macro-security |
| `MACRO_NOT_FOUND` | `run_macro` / `classify_background` unresolved or ambiguous procedure; HRESULT `0x800A03ED` |
| `MACRO_RUNTIME_ERROR` / `MACRO_TIMEOUT` | macro runner HRESULT `0x800A03EC` / `0x80004005`; soft deadline |
| `EXCEL_HUNG` | worker timeout / `NEEDS_RECOVERY` fail-fast / owned instance missing |
| `SNAPSHOT_FAILED` | `require_snapshot` (no snapshot), snapshot create/restore failure, artifact write failure |
| `WORKBOOK_READ_ONLY` | every `_require_writable`; `_require_execute_context` |
| `PATH_NOT_ALLOWED` | `PathPolicy.validate_path` |
| `APPROVAL_DENIED` | `ApprovalPolicy` denial in every mutation service, `SessionService._require_approval`, `macro_run`'s approval branch |
| `BACKGROUND_SAFETY_NOT_VERIFIED` / `INTERACTIVE_UI_NOT_SUPPORTED` | `MacroService.run_macro` blocked result; `classify_background` R13 veto (as an exception) |
| Phase-2 codes | `ranges_service` / `pivot_tables_service` pre-mutation validation |

### 6.3 Do structured errors reach the MCP client?

**Yes — this was fixed in B2 and is verified live.** `server.py::ExcelVbaMcpServer.call_tool` re-raises a
`ToolError` whose message is `json.dumps(cause.to_dict())`. Observed over the wire:

```
call_tool("vba_read_component", {"workbook_id":"__audit_nonexistent__","component":"Nope"})
→ is_error: true
→ {"code": "COM_BUSY", "message": "Unknown workbook_id '__audit_nonexistent__' (not open).", ...}

call_tool("excel_open", {"path": "…\\..\\..\\audit_evil.xlsm"})
→ {"code": "PATH_NOT_ALLOWED", "message": "Path not allowed (traversal).",
   "context": {"reason":"traversal","operation":"mutating","name":"audit_evil.xlsm"}}

call_tool("snapshot_restore", {"snapshot_id": "__audit_missing__"})
→ {"code": "SNAPSHOT_FAILED", "message": "snapshot does not exist.", ...}
```

`excel_health` / `macro_run` blocks are **not** exceptions: they are truthful result payloads
(`{"status":"blocked","error_code":"BACKGROUND_SAFETY_NOT_VERIFIED"}` / `{"status":"needs_recovery",
"error_code":"EXCEL_HUNG"}`), so a client must inspect both `is_error` **and** `payload.status/error_code`.

### 6.4 Taxonomy defects observed

* **Generic domain errors masquerade as `COM_BUSY`.** Any `ExcelVbaMcpError` built without an explicit `code`
  inherits `default_code = COM_BUSY`. Live examples: unknown `workbook_id`, VBA **revision conflict**
  (`"Revision conflict for component 'modWorkflow' …"` → `COM_BUSY`), non-replaceable component. A client cannot
  distinguish caller error / lost session / contention. There is no `WORKBOOK_NOT_FOUND`, `REVISION_CONFLICT` or
  `INVALID_ARGUMENT` code.
* **Lost owned Excel instance surfaces as `COM_BUSY`** for reads (`"The working copy is no longer open in the owned
  instance."` → `COM_BUSY`) but correctly as `EXCEL_HUNG` + `status=needs_recovery` from `macro_run`
  (inconsistent classification of the same underlying condition).
* **Resources are not structured.** `workbook_vba_component_source_resource` / `test_run_resource` raise
  `ResourceError(str(exc))` — plain text, no `code`.

---

## 7. Macro / background-safety design as implemented

### 7.1 Keying

`application/macro_service.py::BackgroundSafetyRegistry` — in-memory, **server-lifetime**, not disk-persisted,
deny-by-default, keyed by the 3-tuple:

```
(workbook_sha256, project_source_hash, procedure.strip().lower())
```

* `workbook_sha256` = SHA-256 of the **working-copy file on disk** (`_sha256_file(record.path)`); cell edits
  without a save do not change it, a save does.
* `project_source_hash` = deterministic hash over **every** VBA component's normalized source, including Document
  modules and UserForm code-behind, computed by `_compute_project_source_hash` from per-component hashes
  (`_scan_project` / `_read_module_lines` / `_hash_module_lines`), with a distinct `<unreadable>` sentinel so an
  unreadable module is never conflated with an empty one (`hash_code("")`).
* Both hashes are derived from **one** STA scan (`MacroService._collect_project_state`) used identically by
  `list_macros`, `run_macro` and `classify_background`, so classify/run/list cannot disagree.

### 7.2 `macro_classify_background` (MCP tool, B5)

* Signature: `(workbook_id, procedure, background_safe, expected_project_source_hash=None)`.
* **Ungated** (explicitly documented as bookkeeping; `macro_run` remains the security boundary).
* Rejects with `MACRO_NOT_FOUND` (unknown/ambiguous), `STALE_TARGET` (optimistic-concurrency mismatch), and
  `INTERACTIVE_UI_NOT_SUPPORTED` when `background_safe=True` and the procedure is heuristic **or** its body
  matches an interactive token (**R13 veto — the deterministic layer can only veto, never certify**).
* Audit entry `tool="macro_classify_background"`, `operation_class="config"`.
* Live: `ShowMessage` → `INTERACTIVE_UI_NOT_SUPPORTED`; `Recalculate` → `{"background_safe": true,
  "project_source_hash": "3817b504…"}`.

### 7.3 `macro_run` gates (order matters)

1. `_require_execute_context` — non-execute → `MACRO_BLOCKED`; read-only → `WORKBOOK_READ_ONLY`.
2. Inventory + source hash scan; unresolved procedure → `status=not_found`, `MACRO_NOT_FOUND`.
3. `ApprovalPolicy.check(operation_class=EXECUTING)` → `status=blocked`, `APPROVAL_DENIED`.
4. `BackgroundSafetyRegistry.is_safe(...)` → `None` → `BACKGROUND_SAFETY_NOT_VERIFIED`; `False` →
   `INTERACTIVE_UI_NOT_SUPPORTED`.
5. **R13 deterministic veto** — `_find_interactive_tokens` (case-insensitive `casefold()` substring match over
   `INTERACTIVE_TOKENS = MsgBox, InputBox, .Show, FileDialog, Shell, CreateObject, FollowHyperlink, SendKeys,
   DoEvents, GetOpenFilename, GetSaveAsFilename, CallByName, ExecuteGlobal, Execute, Eval, OnTime, DisplayAlerts,
   .Run, CommandBars`) **plus heuristic-procedure fail-closed** → `INTERACTIVE_UI_NOT_SUPPORTED`.
6. Single `Application.Run` dispatch (`section="execute"`, `op_class=EXECUTING`, workbook-qualified
   `'<name>'!Module.Proc`); no retry; timeout → `_handle_hung` → `EXCEL_HUNG` / `needs_recovery`.
7. No dialog dismissal anywhere; `_owned_modal_present` + `_terminate_owned_pid` are used only for
   post-hang containment (read-only detection + owned-PID termination), never input injection.

### 7.4 `project_source_hash` invalidation

Because the hash is recomputed live on every list/run/classify, any unsaved **VBA** edit (via
`vba_replace_component_source` / `vba_insert_code`) changes it and silently invalidates prior classifications
(the key no longer matches → `BACKGROUND_SAFETY_NOT_VERIFIED`). Cell-only edits do **not** invalidate — verified
live: one classification survived a `table_update_rows` plus two `macro_run` cycles in the same session. A
`snapshot_restore` creates a *new* working copy and therefore a new registry key (new workbook hash and, as
observed, a new `project_source_hash`), so classifications must be re-applied after a restore.

### 7.5 Residual risks (as implemented)

| ID | Risk | Detail |
|---|---|---|
| BG-R1 | **Heuristic token scan is an under-approximation.** | `AppActivate`, `Application.Visible`, `ThisWorkbook.FollowHyperlink` variants, `On Error`-guarded UI, `CallByName`/`ExecuteGlobal`/`Eval` indirection, and UI invoked from an un-hashed *transitive* procedure are not reliably detected. The code documents this and fails closed on `is_heuristic`, but the docs/plan already record `AppActivate`/`Application.Visible` as a residual (B5 W2). |
| BG-R2 | **Classification is agent-asserted.** | The MCP client decides `background_safe`; the server only vetoes. A false assertion is caught only by the R13 veto + runtime modal observation. |
| BG-R3 | **`workbook_sha256` is the working copy, not the original.** | Two sessions over the same original share the same file hash only if their working copies are byte-identical; after a save they diverge, which *splits* registry keys (fail-closed, but surprising). |
| BG-R4 | **Registry is server-lifetime only.** | Restarting the server drops all classifications; every macro is `BACKGROUND_SAFETY_NOT_VERIFIED` again until re-classified (by design, documented). |
| BG-R5 | **Modal detection is bounded.** | `_owned_modal_present` only recognises the VBIDE dialog class `#32770`; other UI surfaces rely on the process remaining non-visible plus honest status reporting. |

---

## 8. Known limitations / discrepancies confirmed against **current** code

Prior findings come from `01-plan-corrective-full-access-execution.md` §5 (C-audit-1…24) and
`.ai/audit/problems/doc-vs-impl-discrepancies-2026-11-21.md` (D-001…D-005). Status is re-verified against HEAD
`8c22f11`.

| ID | Prior finding | Status now | Evidence |
|---|---|---|---|
| C-audit-1 / R-01 | Config not wired; `server.py` read `os.environ` directly | **STALE — fixed (B2)** | `server.py::lifespan` → `load_config()`; `_build_runtime(config)`; `ApprovalPolicy(mode=config.server.access_mode)` |
| C-audit-2 / R-10 | Default roots are not `projects/work/snapshots`; they are cwd + temp | **CONFIRMED (still true)** | `server.py::_configured_roots`; live open from `%TEMP%` with `EXCEL_VBA_MCP_ROOTS` unset |
| C-audit-3 | `ApprovalPolicy` had no access-mode concept | **STALE — fixed (B2)** | `ApprovalPolicy.__init__(mode=...)`, `check()` full-mode branch |
| C-audit-4 / R-18 | `macro_policy.classify_operation` is fail-closed but **not called at runtime** and is stale | **CONFIRMED (still true)** | grep: `classify_operation` referenced only in `macro_policy.py` + `tests/unit/test_macro_policy.py`. The table still lists `excel_open` as READ_ONLY and omits pivot/range tools. |
| C-audit-5 | `open_workbook` signature / approval classes; no unconditional rejection of `read_only=False` | **CONFIRMED (still true), but the *default* is now EXECUTE** | `SessionService.open_workbook(mode=INSPECT, read_only=False)` default is INSPECT; the **API tool** `excel_open` defaults to `EXECUTE`/writable (B3) |
| C-audit-6 | `MacroInfo` / registry live in `application/macro_service.py`; `run_macro` order | **CONFIRMED, extended (B5)** | registry now 3-tuple hash-bound; `macro_classify_background` exists; `run_macro` order as in §7.3 |
| C-audit-7 | Test runner opens `EXECUTE`/writable copies | **CONFIRMED (B6 reworked)** | `test_runner.py` per-case copies + `classify_background` before `macro_run` |
| C-audit-8 | `excel_open` default was inspect + writable | **STALE — fixed (B3)** | `api/workbook_tools.py::excel_open(mode=SessionMode.EXECUTE, read_only=False)` |
| C-audit-9 | Services mapped approval denial to `COM_BUSY` | **STALE — fixed (B4)** | every service now raises `ApprovalDeniedError` |
| C-audit-10 | No `APPROVAL_DENIED` code; `default_code = COM_BUSY` | **PARTIALLY FIXED** | `APPROVAL_DENIED` exists and is used; `default_code = COM_BUSY` remains (see §6.4) |
| C-audit-11 | No access-mode in domain models | **STALE — fixed (B2/B3)** | `domain/statuses.py::AccessMode`; `SessionStatus.access_mode` (optional field, `extra="forbid"` respected) |
| C-audit-12 | Adapter `describe()` mapping | **CONFIRMED** | `adapters/excel/workbook_adapter.py::describe` |
| C-audit-13 | Test-spec enums closed (`extra="forbid"`) | **CONFIRMED** | `domain/test_spec.py` |
| C-audit-14 | "86 + server_health = 87 tools" | **STALE — now 88** | §2.2 |
| C-audit-15 / R-09 | No MCP client harness; `pywinauto` unused | **HARNESS FIXED (B7); pywinauto STILL UNUSED** | `tests/integration_excel/test_mcp_full_access_workflow.py` uses `stdio_client` + `ClientSession`; `pywinauto` has 0 hits in `src/` but remains a declared dependency (D-003 open) |
| C-audit-16 | Macro-trust static guards (`winreg`, `AccessVBOM =`, `VBAWarnings =`, `AutomationSecurity =` in security modules) | **CONFIRMED intact** | `tests/unit/test_security_policy_static.py` unchanged; `security/*` contains no registry/Trust-Center writes |
| C-audit-17 | BG01–BG09 normative | **CONFIRMED; BG03/BG04 observed (see §9)** | foreground-window probe |
| C-audit-19 | Non-existent docs (`08-security.md`, `release-evidence.md`, `config/*.toml`) | **CONFIRMED (docs updated to 88 tools)** | `docs/05-mcp-tools-reference.md` states 88 |
| C-audit-20 | Dirty tree | **CONFIRMED** | `git status`: 5 deleted docs, untracked `.ai/plans/*`, `.ai/audit/problems/`, `.ai/tasks/todo/` |
| C-audit-21 / R-16 | MCP SDK masked domain errors as `"Error executing tool <name>"` | **STALE — fixed (B2)** | `ExcelVbaMcpServer.call_tool` + `_find_domain_cause`; live structured payloads in §6.3 |
| C-audit-22 / R-17 | Stale worker: `excel_health.status=ok` while owned PID is dead; `NEEDS_RECOVERY` only on timeout | **CONFIRMED — reproduced live (F-2)** | §10 probe 6 |
| C-audit-23 | 87-tool classification (41 RO / 40 MUT / 4 EXEC / 2 MIXED) | **UPDATED** | §5: 37 RO / 42 MUT / 4 EXEC / 2 MIXED / 3 ADMIN |
| C-audit-24 | Macro denial not reachable end-to-end (blocked upstream) | **STALE — resolved by B2/B5** | `macro_run` unclassified → `BACKGROUND_SAFETY_NOT_VERIFIED` observed live |
| D-001 | `background_safe` not settable over MCP | **STALE — fixed (B5)** | `macro_classify_background` live |
| D-002 | `recover_session` / `close_session` not exposed as MCP tools | **CONFIRMED (still true)** | no MCP tool; `excel_health` only reports `needs_recovery` |
| D-003 | `pywinauto` unused | **CONFIRMED (still true)** | 0 imports |
| D-004/D-005 | Doc drift | **RESOLVED for tool count; other links re-verified by B8** | docs state 88/4 |

### 8.1 New / still-open defects found in this baseline (ranked)

| ID | Severity | Type | Finding |
|---|---|---|---|
| **F-1** | **HIGH** | SPEC-DEVIATION | **`snapshot_restore` leaves the session in `inspect` mode → macros are disabled after a restore.** `SnapshotService.restore_snapshot` calls `open_workbook(snapshot_file, mode=SessionMode.INSPECT, read_only=False)`. Live: after `snapshot_restore`, `excel_get_status.mode == "inspect"` and `macro_run` returns `MACRO_BLOCKED` ("Macro execution requires an execute-mode session"). The required workflow "save / snapshot / restore as supported" therefore cannot continue with macro execution; the client must close and re-open the restored content. |
| **F-2** | **HIGH** | BEST-PRACTICE (environment) | **Stale owned-Excel / health is not truthful (R-17) — reproduced.** After `taskkill /F /PID <owned_pid>`: `excel_health` still reports `status=ok, worker_state=running, accepting=true, needs_recovery=false, owned_pid=<dead pid>`, even after two failing calls (only `last_error` is set). Reads then fail `COM_BUSY` ("The working copy is no longer open…") while `macro_run` correctly reports `needs_recovery`/`EXCEL_HUNG`. Remedy today is process restart; no MCP-callable recovery exists (D-002). |
| **F-3** | **MEDIUM** | SPEC-DEVIATION | **`macro_run` has no snapshot precondition.** `MacroService.__init__` takes no `SnapshotService` and `run_macro` never calls `require_snapshot`, although `snapshot_service.py::require_snapshot` documents "mutation entry points (VBA replace/insert, **macro execution**, destructive sheet operations) must call this". Tracked as `CT-B7-macronapshot-precondition`, still open. |
| **F-4** | **MEDIUM** | BEST-PRACTICE | **Ungated filesystem writes.** `vba_export_project` and `snapshot_create` write files with **no approval check and no writable-session requirement**; both succeed in `restricted` mode (live). `snapshot_restore` likewise has no own approval gate (only inherited via `open_workbook`). Documented as R-19, but it means "restricted" is not a complete deny-by-default boundary for filesystem effects. |
| **F-5** | **MEDIUM** | SPEC-DEVIATION | **Generic domain errors surface as `COM_BUSY`.** Unknown `workbook_id`, VBA revision conflicts and non-replaceable components all inherit `ExcelVbaMcpError.default_code = COM_BUSY`. A client cannot distinguish "you passed a bad id" from "Excel is busy". |
| **F-6** | **MEDIUM** | BEST-PRACTICE | **Restricted mode has no in-band remedy.** No MCP tool grants approvals or manages the allowlist, so in `restricted` mode `excel_open` (the default execute open) is permanently denied and the only usable session is `inspect`+`read_only`. Combined with F-1 this makes `restricted` mode effectively read-only-only. |
| **F-7** | **LOW** | BEST-PRACTICE | `macro_policy.classify_operation` is stale, unused production code that contradicts the live enforcement path (lists `excel_open` as READ_ONLY; omits pivot/range tools). Recommend either deleting it or documenting it as non-authoritative (project rule: dead code is only dead when undocumented — here it is documented as "the" classification source in its own docstring, which is the real problem). |
| **F-8** | **LOW** | BEST-PRACTICE | **Orphaned owned Excel.** An `EXCEL.EXE /automation -Embedding` instance from an earlier run was still alive during this audit (PID 3848, created 11:57). `ExcelComWorker.shutdown` only runs on graceful lifespan exit; a killed server leaves a hidden Excel behind with no marker. |
| **F-9** | **LOW** | BEST-PRACTICE | **Audit-classification inconsistencies.** `snapshot_create` is audited as `read_only` while creating a snapshot file; `macro_classify_background` is audited as `config`; docs classify `macro_classify_background` as "read-only" although it mutates in-memory state. |
| **F-10** | **LOW** | BEST-PRACTICE | `pywinauto` remains a declared dependency with zero imports (D-003 open). |
| **F-11** | **LOW** | DOC-UPDATE | `docs/05-mcp-tools-reference.md` counts `executing` = 3 and classifies `vba_compile`/`vba_get_compile_diagnostics` as non-mutating; the live classification also has `vba_compile` submitting `op_class=EXECUTING` with no approval gate. Worth an explicit sentence. |

---

## 9. Baseline function matrix (required acceptance capabilities vs current reality)

| # | Required capability | Status | Evidence (code + live) |
|---|---|---|---|
| 1 | **Full access is the default** (`EXCEL_VBA_MCP_ACCESS_MODE` unset) | **PASS** | `config.py` default `FULL`; live server with the var unset returned `excel_get_status.access_mode="full"`; mutation and macro execution succeeded with no grant. |
| 2 | **Default `excel_open` is writable + execution-capable** | **PASS** | `api/workbook_tools.py::excel_open(mode=EXECUTE, read_only=False)`; live `{"mode":"execute","read_only":false}`. |
| 3 | **Table mutation through MCP** | **PASS** | `table_update_rows(workbook_id, "tblSettings", [{"index":1,"values":{"SettingValue":20}}])` → success; `table_read_rows` showed `factor/5.0`, `label/alpha`; post-update macro output changed accordingly. |
| 4 | **Two consecutive settings-edit → macro-run cycles in one session** | **PASS** | Same `workbook_id`: cycle 1 → `Result!A1 = 10` (factor 5×2); cycle 2 → `Result!A1 = 40` (factor 20×2), `Result!A2 = "factor=20"` — proves the macro read live table state, not a cached value. |
| 5 | **Background-safe classification through MCP** | **PASS** | `macro_classify_background` → `{"background_safe": true, "project_source_hash": "3817b504…"}`; `macro_list` then reports `background_safe: true, background_safe_classified: true`. |
| 6 | **Interactive macro rejection without UI** | **PARTIAL (PASS for classify/run veto; focus NOT VERIFIED for the veto path)** | `macro_classify_background(ShowMessage, true)` → `INTERACTIVE_UI_NOT_SUPPORTED`; unclassified `Recalculate` → `status=blocked, BACKGROUND_SAFETY_NOT_VERIFIED` and **not dispatched**. No dialog was observed (owned instance has no visible window), but no test drove a macro that *would* show a MsgBox through to the veto — the veto is pre-execution, so focus cannot be observed. |
| 7 | **Read-only session denies mutations** | **PASS** | `excel_open(read_only=True)` then `range_write_values` / `table_update_rows` / `macro_run` → `WORKBOOK_READ_ONLY` each. |
| 8 | **Restricted mode preserved** | **PASS (with F-6 caveat)** | `EXCEL_VBA_MCP_ACCESS_MODE=restricted`: default `excel_open` → `APPROVAL_DENIED`; `inspect`+`read_only` open allowed (READ_ONLY auto-approved). No grant path exists (F-6). |
| 9 | **Snapshot / save / restore** | **PARTIAL (F-1)** | `snapshot_create` → `snap-…` with SHA-256 + VBA exports; `excel_save` → success; `excel_close` deletes the working copy; `snapshot_restore` → new `workbook_id` and content **persisted** (`Result!A1=40`, `Sheet1!B2=20`, `Result!A2="factor=20"`, components `ThisWorkbook/Sheet1/Sheet2/modWorkflow`, name `SettingsKey`, macros `Recalculate`/`ShowMessage`). **But** the restored session is `mode=inspect` → `macro_run` fails `MACRO_BLOCKED` (F-1). |
| 10 | **Hidden Excel in the interactive session, no focus theft** | **PARTIAL — observed PASS for the sampled path** | Owned instance via `DispatchEx` with `Visible=False`; `excel_health.owned_pid` is a dedicated PID. Live foreground-window observation across server start → open → snapshot → classify → `macro_run` (success) → close: foreground window stayed `task_tech_lead_full_edit_mode_orchestrator.md - xlvba_mcp - Visual Studio Code` (hwnd 131732) at every sample, and `EnumWindows` found **no visible window** for the owned PID at start and after the macro run. Code: no `SetForegroundWindow`/`ShowWindow`/`SendKeys`/`.Activate`/`.Select` in `src/`. **Residual NOT VERIFIED:** behaviour under a macro that *does* call UI (blocked pre-execution by design), and behaviour when the user's own Excel is in the foreground (my probes ran with VS Code foreground). |
| 11 | **Originals unchanged** | **PASS** | `tests/fixtures/FullAccessWorkflow.xlsm` SHA-256 identical before/after every probe (`6B905837B8DEEB549505F9A8A51C4F55507D377E4160BB5D9D6363A19AC1D4D9`); temp copies also byte-identical after the run. |
| 12 | **Structured errors through MCP** | **PASS** | §6.3 live payloads. |
| 13 | **Auditability** | **PASS** | `audit.jsonl` per tool with `operation_class`, params, approval decision, result and error code (29 entries in one probe run). |

---

## 10. Reproducible findings (exact tools, arguments, observed results)

All probes used a **fresh server process** (avoids R-17) and a **disposable copy** of
`tests/fixtures/FullAccessWorkflow.xlsm`. Full raw JSON: `.ai/tmp/audit_probe.json`,
`audit_functional.json`, `audit_evidence.json`, `audit_restore.json`, `audit_focus.json`, `audit_stale.json`
(plus `*_stderr.log`).

**P1 — Default full mode + writable open.**
`excel_open {"path": "<temp>/input/FullAccessWorkflow.xlsm"}` →
`{"workbook_id":"c3b5a4a7-…","name":"c3b5a4a7-….xlsm","format":"xlsm","read_only":false,"has_vba":true,
"excel_version":"16.0","mode":"execute","model_version":1}`.

**P2 — Snapshot precondition on mutations.**
`range_write_values {Result!B9 = 7}` **before** any snapshot → `is_error`, `SNAPSHOT_FAILED`
("A successful snapshot is required before mutating this workbook.").
`macro_run {"macro_name":"modWorkflow.Recalculate"}` **without** a snapshot → `status=blocked`,
`BACKGROUND_SAFETY_NOT_VERIFIED` (i.e. no snapshot error — see F-3).

**P3 — R13 veto and background-safety deny.**
`macro_classify_background {"procedure":"modWorkflow.ShowMessage","background_safe":true}` →
`INTERACTIVE_UI_NOT_SUPPORTED`.
`macro_run {"macro_name":"modWorkflow.Recalculate"}` before classification → `status=blocked`,
`BACKGROUND_SAFETY_NOT_VERIFIED`.

**P4 — Two-cycle workflow (one `workbook_id`).**
`snapshot_create` → `snap-1e5dc8a1…` → `range_write_values Result!B9=7` → `range_read_values Result!B9` →
`7.0`; `macro_classify_background Recalculate=true` → `macro_run` → `status=success` → `Result!A1 = 10.0`;
`table_update_rows tblSettings index 1 {SettingValue: 20}` → `affected_rows [1]` → `macro_run` →
`status=success` → `Result!A1 = 40.0`, `Result!A2 = "factor=20"`.

**P5 — Read-only session.**
`excel_open {"read_only": true}` → `read_only: true`; `snapshot_create` → OK; `range_write_values` →
`WORKBOOK_READ_ONLY`; `table_update_rows` → `WORKBOOK_READ_ONLY`; `macro_run` → `WORKBOOK_READ_ONLY`.

**P6 — Restricted mode.**
`EXCEL_VBA_MCP_ACCESS_MODE=restricted`: `excel_open {path}` → `APPROVAL_DENIED`
("denied by default: no matching allowlist entry or live grant"); `excel_open {mode:"inspect", read_only:true}` →
success. In the same server, `snapshot_create` and `vba_export_project` **succeeded** (no approval) — F-4.

**P7 — Path policy.**
`excel_open {"path": "…\\..\\..\\audit_evil.xlsm"}` → `PATH_NOT_ALLOWED`
(`{"reason":"traversal","operation":"mutating","name":"audit_evil.xlsm"}`).

**P8 — Post-restore mode (F-1).**
`excel_open` → `mode=execute`; `snapshot_create`; `snapshot_restore {"snapshot_id":…,"workbook_id":…}` →
`mode=inspect, read_only=false`; `excel_get_status` → `"mode":"inspect","access_mode":"full"`;
`macro_run` → `MACRO_BLOCKED` ("Macro execution requires an execute-mode session; macros are disabled in inspect
mode."); `macro_list` still lists macros and `classify` still succeeds.

**P9 — Lost owned Excel vs health (F-2).**
`taskkill /F /PID 24196` (the owned PID reported by `excel_health`) → `alive_after_kill=false`;
`excel_health` → `status=ok, worker_state=running, accepting=true, needs_recovery=false, owned_pid=24196`;
`range_read_values` → `COM_BUSY` ("The working copy is no longer open in the owned instance.");
`macro_run` → `status=needs_recovery, error_code=EXCEL_HUNG`;
`excel_health` afterwards → still `status=ok` with `last_error` set.

**P10 — Focus / visibility.**
Foreground window before start, after start, after `macro_run`, and after settle:
`[131732, "task_tech_lead_full_edit_mode_orchestrator.md - xlvba_mcp - Visual Studio Code"]` (unchanged).
Owned PID visible top-level windows: `[]` at start and `[]` after `macro_run`.

**P11 — Ungated VBA export / compile.**
`vba_export_project {"workbook_id":…}` → 4 components exported to
`%TEMP%\excel_vba_mcp_vba_export_*` (`export_status: ok`); works identically in `restricted` mode.
`vba_compile {"workbook_id":…}` → `status=indeterminate`, `reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`,
`diagnostics.is_compile=false`, `vbom_available=true`, `protection=0`; `vba_get_compile_diagnostics` returns the
same payload.

**P12 — Revision conflict code (F-5).**
`vba_replace_component_source {"component":"modWorkflow","expected_source_hash":"deadbeef", …}` →
`COM_BUSY` "Revision conflict for component 'modWorkflow': the source has changed since it was read."
(`current_source_hash: 3ee58a8e…`).

**P13 — Tool discovery.**
`list_tools` → 88; `list_resources` → 0; `list_resource_templates` → 4 (`excel://workbooks/{workbook_id}/model`,
`…/vba/components`, `…/vba/components/{component_name}`, `excel://test-runs/{run_id}`); `list_prompts` → 0;
`server_health` → `{"status":"ok","service":"excel-vba-mcp","version":"0.1.0"}`.

**P14 — Snapshot → save → close → restore persistence (F-1 reconfirmed).**
`excel_open` → `snapshot_create` → classify → `table_update_rows {SettingValue: 20}` → `macro_run`
(`Result!A1=40`) → `snapshot_create` (edited) → `excel_save` → `excel_close {save_changes:true}` →
`snapshot_restore {edited}` → new `workbook_id`, `mode="inspect"`, then:
`Result!A1 = 40.0`, `Sheet1!B2 = 20.0`, `Result!A2 = "factor=20"`, components
`["ThisWorkbook","Sheet1","Sheet2","modWorkflow"]`, name `SettingsKey → =Sheet1!$E$1`, macros
`["modWorkflow.Recalculate","modWorkflow.ShowMessage"]`; `macro_run` → `MACRO_BLOCKED`
("Macro execution requires an execute-mode session; macros are disabled in inspect mode.").

---

## 11. Method / limitations of this audit

* **Not performed (out of read-only scope):** no unit/integration suite run; no source edit; no `-SyncExports` /
  `-DeployToDataTheiss`; no macro executed against a repository fixture; no registry or Trust Center change; no
  `AutomationSecurity = LOW`.
* **NOT VERIFIED items and the tooling needed:**
  * Focus-steal during an *interactive* macro — requires a fixture macro that actually calls `MsgBox` to be run
    to the point of UI creation, which the design forbids; the honest check is the pre-execution veto (P3) plus
    the foreground-window observation (P10). Needed for a stronger claim: a dedicated probe that classifies an
    interactive macro `background_safe=false`, attempts `macro_run`, and samples `GetForegroundWindow` +
    `EnumWindows` at 50 ms intervals for the duration.
  * Focus behaviour when the **user's** Excel is foreground — not reproducible deterministically in this session.
  * Compile *success/error* paths — unavailable on this build (indeterminate by design).
  * Multi-workbook / multi-session interference — only one session was exercised per server.
* **Environment caveat:** all Excel-backed probes ran on one host (Windows 10 17763, Excel 16.0 14332). Results
  involving VBE/CommandBars are build-specific by the project's own admission.

---

## 12. Bottom line for the Tech Lead

The full-edit MCP surface is **real and usable today**: 88 live tools, 4 read-only resource templates, structured
errors at the MCP boundary, full-access default with a writable execute session, working table/range/name/VBA
mutations, an MCP-accessible background-safety classification, working macro execution with deny-by-default and an
R13 veto, and observed hidden/non-focus-stealing Excel.

The material gaps that block a clean "full-edit verified" claim are, in order: **F-1** (restore drops the session
to inspect mode, so macros cannot continue after a restore), **F-2** (stale worker + untruthful `excel_health`,
no MCP-callable recovery), **F-3** (`macro_run` without a snapshot precondition), **F-4** (ungated filesystem
writes undermine `restricted` mode), **F-5** (`COM_BUSY` overload for caller errors), and **F-6** (no in-band
grant path in `restricted` mode).
---

## 13. Baseline findings disposition (Tech Lead, 2026-11-22)

All §8.1 findings F-1…F-11 and the §6.4 / D-00x items have been re-checked against HEAD and dispositioned:

| ID | Finding | Status | Resolution | Commit |
|---|---|---|---|---|
| **F-1** | `snapshot_restore` drops session to `inspect` (macros disabled) | **FIXED** | `restore_snapshot` opens the restored copy `SessionMode.EXECUTE` + writable; post-restore `macro_run` works. | `0993370` |
| **F-2** | Stale worker / untruthful `excel_health` (dead owned PID) | **FIXED** | `excel_health` additive read-only owned-pid liveness probe → `needs_recovery`; + unit/integration. | `9cf4346` |
| **F-3** | `macro_run` has no snapshot precondition | **ACCEPTED (documented limitation)** | Tech Lead decision: not a defect — a snapshot cannot roll back external macro side effects; e2e snapshots before mutating. Documented (05 item 12). | — |
| **F-4** | `snapshot_create` / `vba_export_project` / `snapshot_restore` ungated fs writes | **FIXED** | `snapshot_create` + `vba_export_project` now approval-gated (APPROVAL_DENIED in restricted); `snapshot_restore` relies on inherited writable `open_workbook` gate. | `5c1e6c4` |
| **F-5** | Generic domain errors surface as `COM_BUSY`; resources unstructured | **FIXED** | Added `WORKBOOK_NOT_FOUND` / `REVISION_CONFLICT`; remapped unknown-id / lost-instance / revision-conflict / non-replaceable call sites; resource errors now `json.dumps(exc.to_dict())`. | `d8e588f` |
| **F-6** | Restricted mode has no in-band grant/recovery tool | **ACCEPTED (documented)** | No MCP grant tool by design; documented remedy (run `full` mode or pre-authorize via library). | `99f26ed` |
| **F-7** | `classify_operation` docstring claims authority; dead code | **FIXED** | Module + function docstrings now state it is illustrative/non-authoritative, not wired, stale. | `7bc43cb` |
| **F-8** | Orphaned owned Excel after killed server | **FIXED** | `ExcelAppAdapter.quit()`/`reap_owned_pid()` enforce death (liveness poll + `taskkill /F /PID` fallback + re-verify); worker mirrors PID only after verified quit; shutdown last-resort reap. | `9131f08` |
| **F-9** | `snapshot_create` audited `read_only`; `macro_classify_background` "read-only" docs | **FIXED** | `snapshot_create` audit → `mutating`; docs describe `macro_classify_background` as ungated bookkeeping. (`export_project` audit label left `read_only` — noted, low.) | `5c1e6c4` |
| **F-10** | `pywinauto` unused declared dependency | **FIXED** | Removed from `pyproject.toml` + lock + docs. | `c28b3dd` |
| **F-11** | `vba_compile` `op_class=EXECUTING` vs docs "read-only"/`executing=3` | **FIXED (docs)** | Docs reconcile: `vba_compile` submits `EXECUTING` (retry/audit tag) but is non-mutating/non-activating and not approval-gated; stays read-only; `executing=3` unchanged. | `99f26ed` |
| §6.4 | Resource errors unstructured | **FIXED** | `ResourceError(json.dumps(exc.to_dict()))`. | `d8e588f` |
| D-002 | `recover_session`/`close_session` not exposed | **ACCEPTED (documented)** | Close/reopen remedy documented (05 item 13). | `119c303` |
| D-003 | `pywinauto` unused | **FIXED** | (same as F-10) | `c28b3dd` |
| C-audit-19 | Broken `08/09/10/11-*.md` doc links | **FIXED** | All repointed to real docs; scan clean. | `92c988c`, `66a61a3` |
| R-17 | Stale worker health | **FIXED** | (same as F-2) | `9cf4346` |
| R-19 | `vba_compile` misclassified EXECUTING | **RESOLVED** | Now `READ_ONLY`. | (pre-existing) |

**Accepted-as-limitation (non-blocking, documented):** F-3 (macro_run snapshot), F-6 (restricted no in-band grant), D-002 (recovery close/reopen).

### Verification at disposition (HEAD)
- `pytest tests/unit -o addopts="" -q` → **896 passed**
- `pytest test_mcp_full_access_workflow.py test_mcp_error_surfacing_matrix.py -o addopts="" -q -p no:xdist` → **7 passed**, both orders, **0 orphaned Excel**
- `ruff check src tests` → clean; `mypy src/excel_vba_mcp` → **Success: 64 files**
- Fixture `FullAccessWorkflow.xlsm` SHA-256 `6B905837…A19AC1D4D9` unchanged

**Baseline audit status: all findings F-1…F-11 dispositioned (fixed or accepted-as-documented); no open defects remain.**
