---
id: xlvba-configuration
domain: 03-configuration
tags:
  - configuration
  - environment
  - security
  - paths
  - audit
related:
  - xlvba-overview
  - xlvba-installation-and-setup
  - xlvba-security
  - xlvba-troubleshooting
  - xlvba-architecture
---

## Purpose

This document is a reference for every configuration parameter that actually exists in `xlvba`:
environment variables, runtime directories, path policy, session parameters, values set in code,
logging and audit.

## Configuration Methods

There is no configuration file. All settings are defined **exclusively through environment
variables**, which are read when the server starts. The configuration model is typed with Pydantic
(`src/excel_vba_mcp/config.py`) and is now wired into the server: `config.load_config()` is called
in `server.lifespan` at startup, so `EXCEL_VBA_MCP_ACCESS_MODE`, `EXCEL_VBA_MCP_LOG_LEVEL`,
`EXCEL_VBA_MCP_MACRO_TIMEOUT_S`, `EXCEL_VBA_MCP_OPEN_TIMEOUT_S` and
`EXCEL_VBA_MCP_LIVENESS_CEILING_S` are validated together and an invalid value for any of them
fails startup with a `ValidationError`.

The remaining variables (`EXCEL_VBA_MCP_ROOTS`, `EXCEL_VBA_MCP_JOB_DIR`,
`EXCEL_VBA_MCP_OWNED_PID_DIR` and the other runtime-directory variables) are read separately by
`server._configured_roots()` / `server._resolve_runtime_path()` / `server._resolve_owned_pid_dir()`,
not by `config.load_config()`; they are still read only at startup.

## Environment Variables

| Variable | Purpose | Default | Allowed Values | Effect |
|----------|---------|---------|----------------|--------|
| `EXCEL_VBA_MCP_ACCESS_MODE` | Server access policy (approval behavior) | `full` | `full`, `restricted` | `full` authorizes `MUTATING`/`EXECUTING` operations without a per-action grant; `restricted` keeps deny-by-default approval. Any other value fails startup fast (Pydantic `ValidationError`). |
| `EXCEL_VBA_MCP_MACRO_TIMEOUT_S` | Macro dispatch deadline (seconds) | `90` | Positive number | Soft per-call ceiling for the `Application.Run` dispatch inside `macro_run` (F2b). Exceeding it classifies the run `MACRO_TIMEOUT` (or `MACRO_RUNTIME_ERROR` when the owned PID shows a modal), terminates only the owned Excel PID and recovers the session in-process. A non-positive value fails startup fast (Pydantic `ValidationError`). |
| `EXCEL_VBA_MCP_OPEN_TIMEOUT_S` | Server-side deadline for the `workbook` (`open`) STA open (seconds) | `300` | Positive number | The per-op-class server authority for the open path (Block E). A non-positive value fails startup fast (Pydantic `ValidationError`). |
| `EXCEL_VBA_MCP_LIVENESS_CEILING_S` | Server-wide upper bound applied to every per-call COM deadline (seconds) | `600` | Positive number | Every per-call timeout is clamped to this ceiling so no single operation's deadline exceeds it (Block E). A non-positive value fails startup fast (Pydantic `ValidationError`). |
| `EXCEL_VBA_MCP_ROOTS` | Allowlist of root directories for workbook paths | CWD + system temp | A path string separated by `os.pathsep` (on Windows `;`) | Paths outside the roots are rejected with `PATH_NOT_ALLOWED`. Empty/missing value → default. |
| `EXCEL_VBA_MCP_WORK_DIR` | Working-copy directory | `<temp>/excel_vba_mcp/work` | Absolute path | Where working copies of opened workbooks are created. |
| `EXCEL_VBA_MCP_SNAPSHOT_DIR` | Snapshot store | `<temp>/excel_vba_mcp/snapshots` | Absolute path | Snapshot directory with manifest and SHA-256. |
| `EXCEL_VBA_MCP_ARTIFACT_DIR` | Artifacts directory | `<temp>/excel_vba_mcp/artifacts` | Absolute path | Artifacts and evidence; also registered as the logical root `artifacts`. |
| `EXCEL_VBA_MCP_RUN_DIR` | Persisted-runs directory | `<temp>/excel_vba_mcp/runs` | Absolute path | Storage for declarative test runs. |
| `EXCEL_VBA_MCP_JOB_DIR` | Background-job store directory | `<temp>/excel_vba_mcp/jobs` | Absolute path | One JSON document per background job (`JobStore`): a `RUNNING` marker at start, finalized on a terminal state, and reconciled to `interrupted_unknown` at startup. |
| `EXCEL_VBA_MCP_OWNED_PID_DIR` | Ownership-record directory (startup reconcile) | `%LOCALAPPDATA%\excel-vba-mcp` | Absolute path | Stable per-user directory for `owned_pid.json` (owner + owned-Excel identity: pid + create-time + exe + session). Deliberately non-volatile so the marker survives a restart / logoff-logon. Falls back to the system temp dir only when `LOCALAPPDATA` is unset. |
| `EXCEL_VBA_MCP_AUDIT_LOG` | Audit-log path | `<temp>/excel_vba_mcp/audit.jsonl` | Absolute path | Append-only JSONL journal. |
| `EXCEL_VBA_MCP_LOG_LEVEL` | Logging level | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` | Parsed in `config.load_config()`, but **not applied**. See the "Logging" section. |

`EXCEL_VBA_MCP_ROOTS` is a list of directories separated by `os.pathsep` (on Windows — `;`).
Each element becomes a logical root `root_N`; the artifact directory is additionally registered
under the name `artifacts`. If the variable is empty or missing, the current working directory
(`Path.cwd()`) and the system temp directory are used.

`vba_export_project` treats the resolved `artifacts` root (and `snapshots`, when configured) as the
only valid location for a caller-supplied `out_dir`. When `out_dir` is outside that root the tool
fails closed with `PATH_NOT_ALLOWED`; the error `context.allowed_root` (and `context.expected_root`)
reports the **absolute** artifacts-root path — by default `<temp>/excel_vba_mcp/artifacts`, or the
directory named by `EXCEL_VBA_MCP_ARTIFACT_DIR` — so a retry can use a valid path. Omitting `out_dir`
exports to a fresh temporary directory instead (absolute paths in the manifest).

> [!IMPORTANT]
> `EXCEL_VBA_MCP_ROOTS` is read **once at server startup** by `server._configured_roots()`.
> After changing it you **must restart the MCP server** (reload the server in Kilo / your MCP
> client) for the new roots to take effect. A running session keeps its original allowlist.

> [!TIP]
> `EXCEL_VBA_MCP_ROOTS` is the supported way to allowlist additional project trees — no source
> change is required. Paths outside every configured root are rejected with `PATH_NOT_ALLOWED`
> (reason `outside_roots`); the error `context` also lists the currently-allowed roots under
> `allowed_roots` so a missing entry is easy to diagnose.

`EXCEL_VBA_MCP_LOG_LEVEL` is parsed by `config.load_config()` and validated through Pydantic,
but it effectively **does not affect** logging: logging is configured once at import time by
`_configure_logging()` with a fixed `INFO` level, before the lifespan runs `load_config()`.
The variable is accepted (and validated) but has no effect in this revision.

## Paths and Directories

| Directory | Variable | Purpose |
|-----------|------------|------------|
| `work/` | `EXCEL_VBA_MCP_WORK_DIR` | Working copies of opened workbooks; only the copy is opened, the original is not modified. |
| `snapshots/` | `EXCEL_VBA_MCP_SNAPSHOT_DIR` | Snapshots: manifest `manifest.json` and SHA-256; snapshots are immutable. |
| `artifacts/` | `EXCEL_VBA_MCP_ARTIFACT_DIR` | Test-run artifacts and evidence; registered as the logical root `artifacts`. |
| `runs/` | `EXCEL_VBA_MCP_RUN_DIR` | Persisted declarative test runs (by `run_id`). |
| `jobs/` | `EXCEL_VBA_MCP_JOB_DIR` | Background-job documents (one JSON file per job); RUNNING marker + startup reconcile to `interrupted_unknown`. |
| `owned_pid.json` | `EXCEL_VBA_MCP_OWNED_PID_DIR` | Ownership record (owner + owned-Excel identity) for startup reconcile; default `%LOCALAPPDATA%\excel-vba-mcp` (non-volatile). |
| `audit.jsonl` | `EXCEL_VBA_MCP_AUDIT_LOG` | Append-only audit journal in JSONL format. |

In addition to its user-facing purpose, the `artifacts` directory is registered in the path policy
under the logical name `artifacts`. The snapshot and artifact directories additionally accept the
`.json` and `.sha256` extensions (for manifests and checksums).

## Path Policy

All paths reaching Excel COM pass through `PathPolicy.validate_path`
(`src/excel_vba_mcp/security/path_policy.py`). The policy canonicalizes the path via
`Path.resolve` and cross-checks it via `os.path.realpath`.

Rejected categories (the `reason` field in the error context, code `PATH_NOT_ALLOWED`):

| `reason` | Meaning |
|----------|---------|
| `empty_path` | Empty path. |
| `device_path_denied` | Device path (`\\?\`, `\\.\` and their variants). |
| `unc_denied` | UNC path (`\\server\share`) when `allow_unc` is disabled. |
| `traversal` | `..` traversal that escapes the root. |
| `outside_roots` | Path outside all allowed roots. |
| `symlink_escape` | Lexically inside a root but resolves outside (symlink/junction/reparse point). |
| `root_mismatch` | Path not in the expected logical root. |
| `extension_not_allowed` | File extension is not in the allowed set. |

Allowed workbook extensions: `.xlsm`, `.xlsb`, `.xlsa`, `.xlam`, `.xltm`, `.xltx`, `.xlsx`, `.xls`.
For the `snapshots` and `artifacts` roots `.json` and `.sha256` are additionally allowed.

The `allow_unc` parameter is always `False` (the setting is defined in code and not available via
environment variables), so UNC and device paths are denied by default.

## Session Parameters

`workbook` (`open`) accepts:

| Parameter | Values | Meaning |
|----------|--------|---------|
| `mode` | `execute` (default), `inspect` | `execute` applies `ByUI` immediately before opening (macros enabled by Office UI policy); `inspect` opens the workbook with macros disabled (`AutomationSecurity = ForceDisable`). |
| `read_only` | `false` (default), `true` | A read-only session opens read-only; mutations in such a session are unavailable. |

Effect on approval (depends on `EXCEL_VBA_MCP_ACCESS_MODE`):

- In the default `full` mode, `workbook` (`open`) with `read_only: false` or `mode: "execute"` is
  authorized by policy; no per-action grant is needed.
- In `restricted` mode, `workbook` (`open`) with `read_only: false` or `mode: "execute"` requires an
  allowlist entry or live grant and is otherwise **denied** with `APPROVAL_DENIED`.
- `mode: "inspect"` + `read_only: true` passes as read-only and requires no grant in either mode.
- Neither `inspect` nor `execute` changes Trust Center or registry settings.

## Values Set in Code

The following values exist in code but are **not configurable through environment variables or MCP**.
They are given as reference, not as settings:

- **Worker soft timeouts.** For most operations `_DEFAULT_TIMEOUT_S = 60.0` seconds
  (session, snapshot, sheets, tables, ranges, names, shapes, controls, pivots, vba_project,
  compile). In `macro_service` the dispatch deadline is `macro_timeout_s`, default `90.0`
  seconds and configurable via `EXCEL_VBA_MCP_MACRO_TIMEOUT_S` (F2b); in
  `test_runner`, `_CALL_TIMEOUT_S = 60.0` and `_MACRO_TIMEOUT_S = 120.0` are used.
- **`orphan_timeout_s`.** The timeout for an "orphaned" worker/thread is set when
  `ExcelComWorker` is constructed (default `10.0` seconds) and is not surfaced in configuration.
- **Path policy `allow_unc`.** Always `False`.
- **Approval allowlist.** Empty (`ApprovalPolicy()` with no allowlist entries), and no grant tool is
  provided over MCP. In the default `full` access mode (`EXCEL_VBA_MCP_ACCESS_MODE=full`) this is
  sufficient because `MUTATING`/`EXECUTING` operations are authorized by policy; `restricted` mode
  re-enables the deny-by-default gate.

## Logging

- Logs are written **only to `stderr`**; `stdout` is reserved for the MCP protocol.
- The level is fixed at `INFO` (`_configure_logging()` calls `logging.basicConfig(level=INFO, stream=sys.stderr, ...)` at import time).
- `EXCEL_VBA_MCP_LOG_LEVEL` is parsed and validated in `config.load_config()` (called at server
  startup) but not applied: the level is set once at import time. In this revision the variable has
  no effect.
- Sensitive values must not reach logs or artifacts.

## Audit

Audit is kept in an append-only JSONL journal (by default `<temp>/excel_vba_mcp/audit.jsonl`,
overridden by `EXCEL_VBA_MCP_AUDIT_LOG`). Each record contains (where available):

- `schema_version`, `timestamp` (UTC), `event_id`, `correlation_id`, `session_id`;
- `workbook_id` and `workbook` (canonical path + SHA-256, file name);
- `tool`, `operation_class`;
- `params` — parameters with mandatory redaction;
- `approval` — approval decision (source, grant id, reason);
- `snapshot_id`, `result`, `user`, `process`, `excel_version`.

Redaction is mandatory and deliberately "compresses" the data: values under "secret" keys
(for example `password`, `token`, `secret`) are replaced with `"<redacted>"`, paths are reduced to
their basename, and macro arguments are logged as typed, length-capped summaries without COM
references.

## Configuration Examples

### Local Development (explicit roots)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "C:\data\workbooks;D:\projects\excel"
uv run excel-vba-mcp
```

### Local project tree (`C:\py_exp\mko_treport`)

Allowlist the whole `mko_treport` tree so `xlvba_workbook` accepts workbooks under it. Set the
variable **before** starting the server, then start/reload the MCP server:

```powershell
# Allow only the mko_treport tree (plus any other roots you need, pathsep-separated).
$env:EXCEL_VBA_MCP_ROOTS = "C:\py_exp\mko_treport"
uv run excel-vba-mcp
```

To keep the existing defaults (CWD + temp) *and* add the project, join them explicitly with
`;` (Windows) — the default is not merged with an explicit value:

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "$env:TEMP;C:\py_exp\mko_treport"
```

> [!NOTE]
> Restart / reload the MCP server after setting the variable. The value is read at startup only;
> a live session continues to use its original allowlist. If `xlvba_workbook` still returns
> `PATH_NOT_ALLOWED / outside_roots`, confirm the server was restarted and that the path is under
> one of the configured roots (see `context.allowed_roots` in the error).

### Production-ish (dedicated directories + audit)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "D:\xlvba\roots"
$env:EXCEL_VBA_MCP_WORK_DIR = "D:\xlvba\work"
$env:EXCEL_VBA_MCP_SNAPSHOT_DIR = "D:\xlvba\snapshots"
$env:EXCEL_VBA_MCP_ARTIFACT_DIR = "D:\xlvba\artifacts"
$env:EXCEL_VBA_MCP_RUN_DIR = "D:\xlvba\runs"
$env:EXCEL_VBA_MCP_AUDIT_LOG = "D:\xlvba\audit.jsonl"
uv run excel-vba-mcp
```

### Restricted Configuration (single root)

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "C:\data\readonly"
uv run excel-vba-mcp
```

### Restricted access mode (deny-by-default)

```powershell
$env:EXCEL_VBA_MCP_ACCESS_MODE = "restricted"
uv run excel-vba-mcp
```

With `restricted`, mutating and executing tools are refused with `APPROVAL_DENIED` unless an
in-process allowlist entry or live grant matches; only read-only work is available over MCP.
Omitting the variable (or setting `full`) keeps the default full-access behavior.

## Configuration Diagnostics

If tools return `PATH_NOT_ALLOWED`, check that the workbook path is inside
`EXCEL_VBA_MCP_ROOTS` (by default CWD and temp). The error `context.allowed_roots` lists the
roots the running server actually loaded; if your newly configured root is missing, the server was
not restarted after the change (see the restart note above). If mutating or executing tools return
`APPROVAL_DENIED`, the server is running in `restricted` mode
(`EXCEL_VBA_MCP_ACCESS_MODE=restricted`); unset it (or set `full`) for the default full-access
behavior. If VBA features are unavailable, check the VBOM setting in Trust Center. For details —
[Excel capabilities](06-excel-capabilities.md).