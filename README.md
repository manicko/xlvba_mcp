# Excel VBA MCP Server

Local MCP server (STDIO) for controlled inspection and automation of Microsoft Excel workbooks on Windows.

> [!IMPORTANT]
> This server drives a real, hidden Excel instance through COM. It is Windows-only and validated on a specific Excel build. By default the server runs in **full access mode**: workbook mutations and macro execution are authorized without per-action grants. Set `EXCEL_VBA_MCP_ACCESS_MODE=restricted` to opt into the deny-by-default approval policy (see [Approval and access model](#approval-and-access-model)).

> [!TIP]
> New to the server? Start with the documentation entry point: [docs/01-overview.md](docs/01-overview.md).

## Features

- Workbook sessions with a hidden, non-activating Excel worker and per-workbook working copies.
- Read-only workbook model, sheets, tables, names, shapes, controls, ranges, and PivotTables.
- VBA inventory, source read/search/export, guarded source replacement/insertion, and compile diagnostics.
- Snapshots with manifest and SHA-256, restore, and diff.
- Bounded declarative test runner with one disposable copy per case.
- JSONL audit trail, path allowlist, and an access-mode policy: full access by default, restricted mode as an explicit opt-in.
- Policy-controlled macro execution with `macro_classify_background` background-safety classification and a non-bypassable interactive-UI veto.

## Requirements

| Component | Requirement |
|-----------|-------------|
| OS | Windows 10/11 x64 (validated on Windows 10 Enterprise LTSC 2019, build 17763) |
| Excel | Desktop Microsoft Excel with COM/VBE (validated on 16.0.14332 x64) |
| Python | 3.14 x64 validated; project floor 3.11 (`>=3.11,<3.15`) |
| Package manager | [uv](https://docs.astral.sh/uv/) |
| VBA features | Excel Trust Center: "Trust access to the VBA project object model" (manual, per machine) |

## Installation

```powershell
uv sync --locked
```

This creates `.venv/` and the `excel-vba-mcp` console script. The `dev` group (pytest, ruff, mypy) is installed by default. No global install or Node.js is required.

If `pyproject.toml` changed, refresh the lock deliberately:

```powershell
uv lock
uv sync
```

### Dependencies

Runtime dependencies from `pyproject.toml` (locked versions in parentheses):

| Package | Constraint | Locked | Purpose |
|---------|------------|--------|---------|
| `mcp` | `>=1.0.0` | 2.3.0 | MCP server, STDIO transport, tool/resource registration |
| `pydantic` | `>=2.0.0` | 2.14.4 | Typed configuration and DTOs |
| `pywin32` | `>=306` | 312 | Excel COM automation |
| `comtypes` | `>=1.2.0` | 1.5.0 | UI Automation (read-only VBE compile diagnostics) |

> [!NOTE]
> The code uses the MCP SDK v2 `MCPServer` API (`mcp.server.mcpserver`), which is why the lock resolves `mcp` to 2.3.0 despite the lower bound in `pyproject.toml`.

Development dependencies (`dev` group): `pytest>=8.0.0` (locked 9.2.0), `ruff>=0.6.0` (locked 0.17.5), `mypy>=1.11.0` (locked 2.4.1).

## Run

```powershell
uv run excel-vba-mcp
```

Or:

```powershell
uv run python -m excel_vba_mcp
```

The server speaks MCP over **STDIO**: protocol messages go to stdout; diagnostics and tool tracebacks go to stderr. Startup is non-fatal if Excel cannot be reached: `server_health` and the tool listing still work, and workbook tools report that the session is unavailable rather than crashing the server.

## Connect an MCP client

### Kilo

Kilo's `mcp` block is runner-agnostic: `type: local` spawns the `command` array as a child process. `uv run --directory` points at the repository root regardless of the client's working directory.

```json
{
  "$schema": "https://app.kilo.ai/config.json",
  "mcp": {
    "xlvba": {
      "type": "local",
      "command": [
        "uv",
        "run",
        "--directory",
        "<absolute-path-to-repo>",
        "excel-vba-mcp"
      ],
      "enabled": true,
      "timeout": 60000
    }
  },
  "permission": {
    "xlvba_*": "ask"
  }
}
```

- Replace `<absolute-path-to-repo>` with the absolute path to this repository's root.
- Tools are exposed as `xlvba_<tool>` (for example `xlvba_excel_open`).
- In the Kilo TUI, `/mcps` toggles loaded servers.
- `permission` defaults to `"ask"` because the tools drive real Excel. Whitelist read-only tools later, for example `"xlvba_server_health": "allow"`.
- Add an `environment` block for the variables in [Configuration](#configuration); `EXCEL_VBA_MCP_ROOTS` is the most important in production.
- `timeout`: allow extra time on first launch (dependency sync / venv warm-up) and for slow Excel COM calls.

### Claude Desktop / generic MCP client

```json
{
  "mcpServers": {
    "xlvba": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "<absolute-path-to-repo>",
        "excel-vba-mcp"
      ]
    }
  }
}
```

Kilo also accepts the module form (`uv run --directory <repo> python -m excel_vba_mcp`), the console script after `uv sync` (`<repo>\.venv\Scripts\excel-vba-mcp.exe`), or a global install (`uv tool install --editable <repo>`, then `excel-vba-mcp`).

## Configuration

Configuration is read from environment variables at server startup.

| Variable | Default | Purpose |
|----------|---------|---------|
| `EXCEL_VBA_MCP_ACCESS_MODE` | `full` | Server access policy: `full` authorizes mutations and macro execution without per-action grants; `restricted` keeps the deny-by-default approval gate. Any other value fails startup fast (a configuration error). |
| `EXCEL_VBA_MCP_ROOTS` | Current working directory and system temp directory | `os.pathsep`-separated allowlist of directories that may contain workbooks. Paths outside these roots are rejected with `PATH_NOT_ALLOWED`. |
| `EXCEL_VBA_MCP_WORK_DIR` | `<temp>/excel_vba_mcp/work` | Working-copy directory. Originals are never modified. |
| `EXCEL_VBA_MCP_SNAPSHOT_DIR` | `<temp>/excel_vba_mcp/snapshots` | Snapshot storage. |
| `EXCEL_VBA_MCP_ARTIFACT_DIR` | `<temp>/excel_vba_mcp/artifacts` | Test-run artifacts and evidence. |
| `EXCEL_VBA_MCP_RUN_DIR` | `<temp>/excel_vba_mcp/runs` | Persisted declarative test runs. |
| `EXCEL_VBA_MCP_AUDIT_LOG` | `<temp>/excel_vba_mcp/audit.jsonl` | Append-only JSONL audit trail. |
| `EXCEL_VBA_MCP_LOG_LEVEL` | `INFO` | Parsed by `config.load_config()` (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`) but not applied; see the warning below. |

`<temp>` is the system temp directory (`%TEMP%` on Windows).

> [!WARNING]
> `EXCEL_VBA_MCP_LOG_LEVEL` is parsed by `config.load_config()` but not applied: logging is configured at import time with a fixed `INFO` level (`_configure_logging()`), so in this revision the server always logs at `INFO` and the variable has no effect. (`config.load_config()` is now called during server startup for `EXCEL_VBA_MCP_ACCESS_MODE`; an invalid access mode or log level fails startup fast.)

> [!TIP]
> Set `EXCEL_VBA_MCP_ROOTS` explicitly. The default (current working directory + temp) is convenient locally but broad.

## Usage

Tool names below are MCP names; Kilo exposes them with its server prefix, for example `xlvba_excel_open`.

### Read-only inspection

1. Open the workbook as a read-only inspect session:

   ```json
   {
     "path": "C:/data/report.xlsm",
     "mode": "inspect",
     "read_only": true
   }
   ```

   Call `excel_open`. The result contains a server-assigned `workbook_id`, the original `source_path`, the working-copy `path`, and `read_only: true`. The path must be inside `EXCEL_VBA_MCP_ROOTS`.

2. Inspect the workbook model or a sheet:

   ```json
   { "workbook_id": "<workbook_id>", "include_vba": false }
   ```

   Call `workbook_get_model`, or use `sheets_list`, `tables_list`, `names_list`, `shapes_list`, `controls_list`, or `pivot_tables_list`.

3. Read a bounded range:

   ```json
   {
     "workbook_id": "<workbook_id>",
     "sheet": "Sheet1",
     "address": "A1:C10",
     "offset": 0,
     "limit": 10
   }
   ```

   Call `range_read_values`, `range_read_formulas`, `range_get_format`, `range_get_validation`, or `range_get_conditional_formatting`.

4. Create a snapshot:

   ```json
   { "workbook_id": "<workbook_id>" }
   ```

   Call `snapshot_create`; it returns a `snapshot_id` and SHA-256. `snapshot_list` and `snapshot_diff` are also available.

5. Close without saving:

   ```json
   { "workbook_id": "<workbook_id>", "save_changes": false }
   ```

   Call `excel_close`.

### Approval and access model

The server has two access modes, selected by `EXCEL_VBA_MCP_ACCESS_MODE`:

- **`full` (default).** Mutating and executing operations are authorized by policy without a per-action grant. `excel_open` opens a writable, execution-capable session by default, guarded mutations run after a snapshot, and `macro_run` needs no per-run grant. Path policy, snapshot-before-mutation, the interactive-UI veto and audit still apply.
- **`restricted` (opt-in).** The deny-by-default gate applies: `MUTATING` and `EXECUTING` operations are refused with `APPROVAL_DENIED` unless an in-process allowlist entry or live grant matches. No MCP tool grants approval or edits the allowlist, so through MCP alone a restricted server exposes only read-only work, and `excel_open` with `mode="execute"` or `read_only=false` surfaces as `APPROVAL_DENIED`.

Other invariants:

- Read-only operations are always auto-allowed.
- Every mutating tool requires a writable working copy and a baseline snapshot, then performs a postcondition re-read and writes an audit entry.
- `excel_open` defaults to `mode="execute", read_only=False` (writable). Pass `mode="inspect"` and/or `read_only=True` for a macros-disabled or read-only session.
- `macro_run` runs only a procedure classified `background_safe` for the current workbook hash and VBA source revision; classify a procedure with `macro_classify_background`. An unclassified procedure returns `BACKGROUND_SAFETY_NOT_VERIFIED`; an interactive one is blocked by the preflight with `INTERACTIVE_UI_NOT_SUPPORTED`.
- `test_run_suite` / `test_suite_run` open disposable copies; a macro step on an unclassified procedure is blocked with `BACKGROUND_SAFETY_NOT_VERIFIED`.
- Inspect mode opens with macros disabled (`AutomationSecurity = ForceDisable`); execute mode uses `ByUI` immediately before open. Neither mode changes Excel Trust Center or registry settings.
- Tool failures surface to MCP clients as a structured JSON error carrying a stable `code` and `context` (for example `APPROVAL_DENIED`, `PATH_NOT_ALLOWED`); the same error is logged to stderr and recorded in the audit trail.

### Declarative test suite shape

`test_suite_run` accepts a plain JSON object (or `test_run_suite` with a `TestSuiteSpec`). The schema is closed: unknown actions, assertions, or extra fields are rejected before any Excel call.

```json
{
  "suite": {
    "suite": "smoke",
    "workbook_id": "<workbook_id>",
    "baseline": "create",
    "cases": [
      {
        "id": "sheet-present",
        "steps": [],
        "assertions": [
          { "type": "sheet_exists", "sheet": "Sheet1" }
        ]
      }
    ]
  }
}
```

Call `test_suite_run` with this object as the `suite` argument. `test_get_run` returns a persisted run by `run_id`.

## Tool catalog

The server registers **88 MCP tools** and **4 read-only MCP resources**.

| Group | Count | Tools |
|-------|-------|-------|
| Workbook/session | 6 | `excel_open`, `excel_close`, `excel_save`, `excel_get_status`, `excel_list_workbooks`, `excel_health` |
| Snapshots | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` |
| VBA inventory/editing | 9 | `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code`, `vba_replace_component_source`, `vba_insert_code` |
| Compile | 2 | `vba_compile`, `vba_get_compile_diagnostics` |
| Macros | 4 | `macro_list`, `macro_run`, `macro_get_last_result`, `macro_classify_background` |
| Sheets | 7 | `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` |
| Tables | 11 | `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`, `table_delete_column`, `table_create`, `table_resize` |
| Ranges & cells | 16 | `range_get`, `range_read_values`, `range_read_formulas`, `range_write_values`, `range_write_formulas`, `range_get_format`, `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`, `range_get_validation`, `range_set_validation`, `range_get_conditional_formatting`, `range_set_conditional_formatting` |
| PivotTables | 13 | `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters` |
| Names | 5 | `names_list`, `name_get`, `name_create`, `name_update`, `name_delete` |
| Shapes & controls | 6 | `shapes_list`, `shape_get`, `shape_set_on_action`, `controls_list`, `control_get`, `control_set_on_action` |
| Workbook model | 1 | `workbook_get_model` |
| Test runner | 3 | `test_run_suite`, `test_suite_run`, `test_get_run` |
| Health | 1 | `server_health` |

Read-only resources:

| URI template | Name | MIME type |
|--------------|------|-----------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` |

## Project structure

```text
.
├─ pyproject.toml
├─ uv.lock
├─ README.md
├─ src/excel_vba_mcp/
│  ├─ __main__.py              # console entry point (mcp.run stdio)
│  ├─ server.py                # MCPServer, tool/resource registration, runtime lifespan
│  ├─ config.py                # typed environment configuration
│  ├─ api/                     # thin MCP tool handlers (validation + delegation)
│  ├─ application/             # session, snapshot, workbook model, sheets, tables, ranges, pivots, names, shapes, controls, VBA, compile, macro, and test-runner services
│  ├─ adapters/excel/          # single-STA COM worker and Excel/VBIDE adapters
│  ├─ adapters/windows_ui/     # VBE compile UI adapter
│  ├─ domain/                  # models, errors, statuses, test-suite schema
│  ├─ security/                # approval, audit, path policy, macro policy
│  └─ storage/                 # snapshot store, artifact store, run repository
├─ tests/
│  ├─ unit/                    # no COM/Excel
│  ├─ integration_excel/       # live Excel COM tests (marker: integration_excel)
│  └─ fixtures/                # workbook fixtures and deterministic generators
├─ docs/                       # technical specification, scope matrix, documentation rules
├─ .ai/                        # plans, audits, context, execution logs
└─ .kilo/                      # agents, commands, rules
```

All COM access is serialized on a single-STA worker; tools never touch COM directly. Tool handlers validate arguments and delegate to services; services enforce the guard chain and return DTOs.

Runtime data is written outside the repository by default (`<temp>/excel_vba_mcp/`). `work/`, `snapshots/`, `artifacts/`, and `runs/` are runtime directories, not source.

## Development

```powershell
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src
uv run pytest tests/unit
```

Integration tests require Windows, Excel, and (for VBA fixtures) VBOM trust. They are marked `integration_excel` and run serially:

```powershell
uv run pytest -m integration_excel
```

Rebuild fixtures when the binary fixtures are missing or stale:

```powershell
uv run python tests/fixtures/build_fixtures.py
uv run python tests/fixtures/build_phase2_fixtures.py
```

## Known limitations

- **Compile is fail-closed.** On the validated build (Excel 16.0.14332), `vba_compile` returns `status: "indeterminate"` with `reason: "NON_ACTIVATING_COMPILE_UNAVAILABLE"`; it never reports a false success. Partial diagnostics (broken references, protection, inventory, VBOM status) are still returned.
- **Access mode defaults to full.** In the default `full` mode mutations and macro execution are authorized without per-action grants; set `EXCEL_VBA_MCP_ACCESS_MODE=restricted` for deny-by-default (`APPROVAL_DENIED`). See [Approval and access model](#approval-and-access-model).
- **Macros require background-safety classification.** `macro_run` runs only a procedure classified `background_safe` for the current workbook hash and VBA source revision (via `macro_classify_background`); an unclassified procedure returns `BACKGROUND_SAFETY_NOT_VERIFIED`. Interactive macros are blocked by the preflight (`INTERACTIVE_UI_NOT_SUPPORTED`). The preflight token scan is a heuristic veto, not static proof: dynamic `CallByName`/`ExecuteGlobal`/`Eval` indirection and transitive UI are not fully covered (the veto list does include `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`, so direct use is caught).
- **Macro runtime errors can show a modal.** The server terminates only its owned Excel PID and reports `MACRO_RUNTIME_ERROR` (`dismissed=false`); it never dismisses dialogs.
- **ActiveX event mapping is inferred** from `<sheet_code_name>_<control>_<Event>` naming; it is not click-tested.
- **PivotTables:** management of existing PivotTables only; creating or deleting PivotTables or PivotCaches is out of scope. External/OLAP/Data-Model sources are rejected. `PivotFilters.Add` is unsupported on the validated build; filters are limited to manual visible-items filters and clear-all.
- **Range structure vs Tables:** range operations that intersect an Excel Table (ListObject) are rejected and routed to the Tables tools.
- **Formats:** `.xlsm` is the validated scope. Other Excel extensions pass the path allowlist but are not validated.
- **Missing manual fixtures:** `LockedProject.xlsm` and `TableNamesProject.xlsm` are not committed; related tests skip with explicit reasons.

## Documentation

Documentation entry point: [docs/01-overview.md](docs/01-overview.md).

- [Overview](docs/01-overview.md) — what the server is, for whom, and its non-goals
- [Installation and setup](docs/02-installation-and-setup.md) — install, configure Windows/Excel, first run
- [Configuration](docs/03-configuration.md) — environment variables, runtime paths, path policy, audit
- [Architecture](docs/04-architecture.md) — STDIO transport, single-STA worker, layers, guard chain
- [MCP tools reference](docs/05-mcp-tools-reference.md) — all 88 tools and 4 resources
- [Excel capabilities](docs/06-excel-capabilities.md) — sheets, tables, ranges, PivotTables, VBA, macros
- [Workflows and examples](docs/07-workflows-and-examples.md) — practical read-only and mutating recipes
- [Security](docs/08-security.md) — path allowlist, access modes (full/restricted), macro safety, audit
- [Limitations and compatibility](docs/09-limitations-and-compatibility.md) — fail-closed compile, PivotTable limits, validated scope
- [Troubleshooting](docs/10-troubleshooting.md) — Excel unavailable, VBOM access, path policy, blocked macros
- [Development and maintenance](docs/11-development-and-maintenance.md) — tests, fixtures, doc rules

Russian version: [docs/ru/](docs/ru/) — same content in Russian.

Additional references:

- [Technical specification](docs/01-Spec/excel_vba_mcp_technical_spec.docx)
- [Scope matrix](docs/01-Spec/excel_vba_mcp_scope_matrix.xlsx)
- [Documentation rules](docs/00-overview/doc-maintenance-rules.md)
- [Phase 2 implementation plan](.ai/plans/plan_phase2_ranges_cells_pivottables.md)
- [Architecture audit](.ai/audit/architecture-audit-2026-11-19.md)
