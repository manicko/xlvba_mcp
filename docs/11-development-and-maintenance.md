---
id: xlvba-development-and-maintenance
domain: 11-development-and-maintenance
tags:
  - development
  - maintenance
  - architecture
  - testing
  - contributing
related:
  - xlvba-overview
  - xlvba-installation-and-setup
  - xlvba-configuration
  - xlvba-architecture
  - xlvba-mcp-tools-reference
  - xlvba-excel-capabilities
  - xlvba-workflows-and-examples
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
---

## Purpose

This document is for **maintainers and contributors** of `xlvba`. It explains
the repository layout, the layer boundaries, how to add a new MCP tool without
breaking the architecture, the error-handling contract, and the test, logging
and release workflow.

If you only want to use the server, start with
[Overview](01-overview.md) and [Installation and setup](02-installation-and-setup.md).

Documentation governance (frontmatter contract, kebab-case, relative links, split
thresholds) is defined in
[doc-maintenance-rules](00-overview/doc-maintenance-rules.md); follow it for every
change under `docs/`.

## Repository structure

| Path | Responsibility |
|------|----------------|
| `pyproject.toml` | Package metadata, dependencies, `excel-vba-mcp` console script, tool config (ruff, mypy, pytest) |
| `uv.lock` | Locked dependency versions |
| `README.md` | Entry point, requirements, install/run, configuration, tool catalog |
| `src/excel_vba_mcp/__main__.py` | Console entry point (`mcp.run(transport="stdio")`) |
| `src/excel_vba_mcp/server.py` | `MCPServer`, lifespan, tool/resource registration, structured-error boundary, logging |
| `src/excel_vba_mcp/config.py` | Typed environment configuration (Pydantic) |
| `src/excel_vba_mcp/api/` | Thin async MCP tool handlers (validation + delegation) |
| `src/excel_vba_mcp/application/` | Session, snapshot, workbook model, sheets, tables, ranges, pivots, names, shapes, controls, VBA, compile, macro, test runner |
| `src/excel_vba_mcp/adapters/excel/` | Single-STA COM worker and Excel/VBIDE adapters |
| `src/excel_vba_mcp/adapters/windows_ui/` | Read-only VBE compile UI adapter (`comtypes`) |
| `src/excel_vba_mcp/domain/` | Models, statuses/enums, stable error codes, declarative test schema |
| `src/excel_vba_mcp/security/` | Approval, audit, path policy, macro policy |
| `src/excel_vba_mcp/storage/` | Snapshot store, artifact store, run repository |
| `tests/unit/` | No COM / Excel |
| `tests/integration_excel/` | Live Excel COM tests (marker `integration_excel`) |
| `tests/fixtures/` | Workbook fixtures and deterministic generators |
| `docs/` | Project documentation |
| `.ai/` | Plans, audits, context, execution logs |
| `.kilo/` | Agents, commands, rules |

## Main source modules

| Module | Responsibility |
|--------|----------------|
| `server.py` | Builds `MCPServer`, owns the runtime via `lifespan`, registers 88 tools and 4 resources, re-raises structured `ToolError`s, configures `stderr` logging |
| `config.py` | Loads `EXCEL_VBA_MCP_LOG_LEVEL`, `EXCEL_VBA_MCP_ACCESS_MODE` and `EXCEL_VBA_MCP_MACRO_TIMEOUT_S` into a Pydantic `Config`; validates them together |
| `api/*_tools.py` | One module per tool group; each handler is thin: validate arguments, delegate to a service, return a DTO |
| `application/session_service.py` | Open/close/save working copies, per-path lock, session status, recovery |
| `application/snapshot_service.py` | Immutable snapshots and restore |
| `application/*_service.py` | Sheets, tables, ranges, pivots, names, shapes, controls |
| `application/vba_project_service.py` | VBA project/component/source read, export, guarded replace/insert |
| `application/compile_service.py` | Fail-closed compile diagnostics |
| `application/macro_service.py` | Macro inventory, `background_safe` registry, R13 preflight, `Application.Run` dispatch |
| `application/test_runner.py` | Declarative test runner over disposable copies |
| `application/workbook_model_service.py` | Aggregated workbook model |
| `adapters/excel/com_worker.py` | Single STA thread, FIFO queue, soft timeouts, read-only retry, recovery |
| `adapters/excel/app_adapter.py` | Hidden `Excel.Application`, `DispatchEx`, narrowly scoped settings, pid-scoped teardown |
| `adapters/excel/*.py` | Sheet/table/range/pivot/VBIDE/shape/control/macro adapters |
| `adapters/windows_ui/vbe_compile_ui.py` | Read-only VBE command discovery and dialog observation (`comtypes`) |
| `domain/models.py` | Pydantic DTOs and size ceilings |
| `domain/statuses.py` | Closed status vocabularies and enums |
| `domain/errors.py` | `ExcelVbaMcpError` hierarchy, `ErrorCode`, HRESULT mapping |
| `domain/test_spec.py` | Declarative test schema (closed action/assertion enums) |
| `security/path_policy.py` | Allowlist and canonicalisation |
| `security/approval.py` | `full` / `restricted` approval policy |
| `security/audit.py` | Append-only JSONL audit with redaction |
| `security/macro_policy.py` | `AutomationSecurity` modes and macro diagnostics |
| `storage/snapshot_store.py` | Immutable snapshot store |
| `storage/artifact_store.py` | Artifacts and evidence |
| `storage/run_repository.py` | Persisted declarative runs |

## Development environment

```powershell
uv sync --locked
```

- `.python-version` pins **3.14**; the project floor is **3.11**
  (`requires-python = ">=3.11,<3.15"`).
- The `dev` group (`pytest`, `ruff`, `mypy`) is installed by default.
- Run the server with:

```powershell
uv run excel-vba-mcp
```

- The console script is also available at
  `<repo>\.venv\Scripts\excel-vba-mcp.exe` after `uv sync`.

## How to add a new MCP tool

Follow the existing layers. A tool is not a single function; it is a DTO + an
adapter method + a service method + a thin handler + registration + tests + docs.

1. **Add DTOs and enums to `domain/`.** Define the request/response models in
   `domain/models.py` and any new status vocabulary in `domain/statuses.py`.
   Domain models are Pydantic `BaseModel`s and must stay COM-free.
2. **Add an adapter method in `adapters/excel/`.** COM stays on the STA worker.
   The adapter method runs inside a submitted callable and returns **DTOs only**
   (never a COM proxy).
3. **Add a service method in `application/`.** The service must go through
   `worker.submit(...)` and enforce the guard chain:
   **writable session → snapshot → approval → operation-specific checks →
   postcondition re-read + audit**. A service never holds a COM proxy.
4. **Add a thin async handler in `api/<group>_tools.py`.** Give it typed
   parameters (the MCP SDK derives the input schema from the signature), validate
   values with the module's explicit guards (`_require_*`), use a strict Pydantic
   payload/DTO model with `extra="forbid"` for structured inputs, delegate to the
   service, and return the DTO. The API layer must not import
   `pythoncom`/`win32com` and must not touch COM.
5. **Register it in `server.py` `_register_tools`.** Use
   `server.add_tool(handler, name="...", description=_DESC_...)` with a stable
   name and a `_DESC_*` description constant. A dependency-free tool (like
   `server_health`) can use `@mcp.tool()` directly.
6. **Add tests.** Unit tests go under `tests/unit` (no COM/Excel). If the tool
   touches Excel, add an integration test under `tests/integration_excel`
   (marker `integration_excel`).
7. **Update the docs.** Add the tool to
   [MCP tools reference](05-mcp-tools-reference.md) and, if it adds Excel
   capability, to [Excel capabilities](06-excel-capabilities.md).

## Implementation conventions

- **English only.** All comments, logs, docstrings, error messages and docs are in
  English.
- **Single source of truth per fact.** A fact lives in exactly one module or doc.
- **Layer boundaries.**
  - `api` must not import `pythoncom`/`win32com` and must not touch COM.
  - Services never hold COM proxies; every operation is `await worker.submit(...)`.
  - `domain` has no COM dependency; an HRESULT is only an integer.
- **DTO boundary.** `_ensure_dto` validates every worker return value
  (Pydantic model, primitive, or container of those). A COM proxy or any other
  non-serializable object raises and is recorded as `PROXY_LEAK_SUSPECTED`.
- **Stable `ErrorCode` vocabulary.** Codes are additive only; never remove or
  change an existing code.
- **Pydantic `extra="forbid"`** on tool inputs.
- **Explicit `shift`, never guessed.** `range_insert` / `range_delete` reject a
  missing or unknown shift.
- **No blind retry of mutations.** Only `READ_ONLY` calls are retried.
- **Audit every operation.** Every mutating and executing path writes an audit
  record with its `operation_class` and result.
- **Stable MCP tool names.** Tool names are a public contract; do not rename them
  casually.

## Error handling requirements

- Raise `ExcelVbaMcpError` subclasses with a stable `ErrorCode`
  (`domain/errors.py`).
- Include `hresult` when the error came from COM; `hresult_hex` is derived.
- Provide a structured `context` dictionary (path, `workbook_id`, `reason`,
  `section`, `token`, `class`, …).
- Never leak COM objects across the boundary; only DTOs cross the worker.
- Argument-validation failures map to `INVALID_RANGE` with
  `reason=invalid_arguments` (see `server._validation_error_as_domain`).
- **Never claim success on an indeterminate compile.** `vba_compile` returns
  `status="indeterminate"` with `diagnostics.is_compile=false` on the validated
  build.

## Testing approach

- **Unit tests** (`tests/unit`) run without COM/Excel and must stay hermetic.
- **Integration tests** (`tests/integration_excel`) drive a live Excel COM
  instance. They are marked `integration_excel` and skipped when Excel is
  unavailable. They run serially because there is one Excel instance.
- **Fixtures** are built by `tests/fixtures/build_fixtures.py` and
  `tests/fixtures/build_phase2_fixtures.py`. Rebuild them when the binary
  fixtures are missing or stale. `LockedProject.xlsm` and
  `TableNamesProject.xlsm` are not committed, so those tests skip with an
  explicit reason.

```powershell
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src
uv run pytest tests/unit
uv run pytest -m integration_excel
```

Run all four static/unit commands before proposing a change; run the integration
suite on Windows with Excel when the change touches COM behavior.

## Logging and debugging

- Logs go **only to `stderr`** at a fixed `INFO` level
  (`server._configure_logging()`); `stdout` is reserved for MCP protocol traffic.
- `EXCEL_VBA_MCP_LOG_LEVEL` is parsed and validated but **not applied** in this
  revision; do not rely on it to change verbosity.
- The audit trail (`audit.jsonl`) is the authoritative record of what happened.
- Use `server_health` (liveness) and `excel_health` (worker/session state) to
  diagnose without touching COM.
- Structured errors carry `code`/`context`; prefer them over parsing log text.

## Git and release workflow

- Inspect the history for the commit-message style:

```powershell
git log --oneline -10
```

- Use short, conventional subjects (for example `docs: ...`, `fix: ...`,
  `feat: ...`, `test: ...`).
- Stage and commit **only the intended files**; review `git status` and
  `git diff` first. Never commit secrets.
- Do not amend or rebase shared history. If a commit fails a hook, fix the issue
  and create a new commit rather than amending the failed one.
- Do not commit runtime data (`<temp>/excel_vba_mcp/`) or build artifacts.

## Dependency maintenance

- Edit `pyproject.toml`, then refresh the lock and environment deliberately:

```powershell
uv lock
uv sync
```

- Runtime dependencies: `mcp`, `pydantic`, `pywin32`, `comtypes`. Development
  dependencies: `pytest`, `ruff`, `mypy`.
- `pyproject.toml` declares `mcp>=2.3,<3`; the lock resolves 2.3.0. The code uses
  the SDK v2 `MCPServer` API (`mcp.server.mcpserver`).
- Pin exact versions through `uv.lock` and review dependency updates; a
  `pywin32`/`comtypes` update can change COM behavior on the validated build.

## Related

- [Overview](01-overview.md) — what the server is.
- [Installation and setup](02-installation-and-setup.md) — environment setup.
- [Configuration](03-configuration.md) — environment variables.
- [Architecture](04-architecture.md) — layers, worker, guard chain.
- [MCP tools reference](05-mcp-tools-reference.md) — tool contract.
- [Excel capabilities](06-excel-capabilities.md) — capability matrix.
- [Workflows and examples](07-workflows-and-examples.md) — usage recipes.
- [Security](08-security.md) — security model.
- [Limitations and compatibility](09-limitations-and-compatibility.md) — known limits.
- [Troubleshooting](10-troubleshooting.md) — diagnosing failures.
- [Documentation rules](00-overview/doc-maintenance-rules.md) — doc governance.