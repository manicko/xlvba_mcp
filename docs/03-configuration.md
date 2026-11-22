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
in `server.lifespan` at startup, so `EXCEL_VBA_MCP_ACCESS_MODE` and `EXCEL_VBA_MCP_LOG_LEVEL` are
validated together and an invalid value for either fails startup with a `ValidationError`.

## Environment Variables

| Variable | Purpose | Default | Allowed Values | Effect |
|----------|---------|---------|----------------|--------|
| `EXCEL_VBA_MCP_ACCESS_MODE` | Server access policy (approval behavior) | `full` | `full`, `restricted` | `full` authorizes `MUTATING`/`EXECUTING` operations without a per-action grant; `restricted` keeps deny-by-default approval. Any other value fails startup fast (Pydantic `ValidationError`). |
| `EXCEL_VBA_MCP_ROOTS` | Allowlist of root directories for workbook paths | CWD + system temp | A path string separated by `os.pathsep` (on Windows `;`) | Paths outside the roots are rejected with `PATH_NOT_ALLOWED`. Empty/missing value → default. |
| `EXCEL_VBA_MCP_WORK_DIR` | Working-copy directory | `<temp>/excel_vba_mcp/work` | Absolute path | Where working copies of opened workbooks are created. |
| `EXCEL_VBA_MCP_SNAPSHOT_DIR` | Snapshot store | `<temp>/excel_vba_mcp/snapshots` | Absolute path | Snapshot directory with manifest and SHA-256. |
| `EXCEL_VBA_MCP_ARTIFACT_DIR` | Artifacts directory | `<temp>/excel_vba_mcp/artifacts` | Absolute path | Artifacts and evidence; also registered as the logical root `artifacts`. |
| `EXCEL_VBA_MCP_RUN_DIR` | Persisted-runs directory | `<temp>/excel_vba_mcp/runs` | Absolute path | Storage for declarative test runs. |
| `EXCEL_VBA_MCP_AUDIT_LOG` | Audit-log path | `<temp>/excel_vba_mcp/audit.jsonl` | Absolute path | Append-only JSONL journal. |
| `EXCEL_VBA_MCP_LOG_LEVEL` | Logging level | `INFO` | `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` | Parsed in `config.load_config()`, but **not applied**. See the "Logging" section. |

`EXCEL_VBA_MCP_ROOTS` is a list of directories separated by `os.pathsep` (on Windows — `;`).
Each element becomes a logical root `root_N`; the artifact directory is additionally registered
under the name `artifacts`. If the variable is empty or missing, the current working directory
(`Path.cwd()`) and the system temp directory are used.

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

`excel_open(path, mode, read_only)` accepts:

| Parameter | Values | Meaning |
|----------|--------|---------|
| `mode` | `execute` (default), `inspect` | `execute` applies `ByUI` immediately before opening (macros enabled by Office UI policy); `inspect` opens the workbook with macros disabled (`AutomationSecurity = ForceDisable`). |
| `read_only` | `false` (default), `true` | A read-only session opens read-only; mutations in such a session are unavailable. |

Effect on approval (depends on `EXCEL_VBA_MCP_ACCESS_MODE`):

- In the default `full` mode, `excel_open` with `read_only: false` or `mode: "execute"` is
  authorized by policy; no per-action grant is needed.
- In `restricted` mode, `excel_open` with `read_only: false` or `mode: "execute"` requires an
  allowlist entry or live grant and is otherwise **denied** with `APPROVAL_DENIED`.
- `mode: "inspect"` + `read_only: true` passes as read-only and requires no grant in either mode.
- Neither `inspect` nor `execute` changes Trust Center or registry settings.

## Values Set in Code

The following values exist in code but are **not configurable through environment variables or MCP**.
They are given as reference, not as settings:

- **Worker soft timeouts.** For most operations `_DEFAULT_TIMEOUT_S = 60.0` seconds
  (session, snapshot, sheets, tables, ranges, names, shapes, controls, pivots, vba_project,
  compile). For `macro_service` `_DEFAULT_TIMEOUT_S = 120.0`; in `test_runner`,
  `_CALL_TIMEOUT_S = 60.0` and `_MACRO_TIMEOUT_S = 120.0` are used.
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
`EXCEL_VBA_MCP_ROOTS` (by default CWD and temp). If mutating or executing tools return
`APPROVAL_DENIED`, the server is running in `restricted` mode
(`EXCEL_VBA_MCP_ACCESS_MODE=restricted`); unset it (or set `full`) for the default full-access
behavior. If VBA features are unavailable, check the VBOM setting in Trust Center. For details —
[Troubleshooting](10-troubleshooting.md).