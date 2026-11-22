---
id: xlvba-security
domain: 08-security
tags:
  - security
  - permissions
  - approval
  - paths
  - audit
  - macros
related:
  - xlvba-overview
  - xlvba-installation-and-setup
  - xlvba-configuration
  - xlvba-architecture
  - xlvba-mcp-tools-reference
  - xlvba-excel-capabilities
  - xlvba-workflows-and-examples
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
  - xlvba-development-and-maintenance
---

## Purpose

This document describes the security and permission model of `xlvba`: where the
trust boundaries are, which controls are implemented in code, which depend on
Windows and Office settings, and which are recommendations the application does
**not** enforce. It is deliberately explicit about residual risk: the server
drives a real Excel instance and can execute real VBA, so "secure" here means
"bounded, observable and honest", not "safe by default".

For configuration variables see [Configuration](03-configuration.md). For the
internal request path and the guard chain see [Architecture](04-architecture.md).
For the full tool contract see [MCP tools reference](05-mcp-tools-reference.md).

## MCP transport and communication boundary

- The server speaks **MCP over STDIO**. Protocol messages (JSON-RPC) travel on
  `stdout`; diagnostics, logs and tracebacks travel on `stderr`
  (`server._configure_logging()` sets a fixed `INFO` level on `sys.stderr`).
- The MCP client **spawns the server as a local child process** (for example
  Kilo's `type: local`, or Claude Desktop's `command`/`args`). The server does
  not open a network listener and exposes no TCP/HTTP endpoint, so there is no
  remote access path in the server itself.
- Because the transport is a local pipe, the effective trust boundary is the
  machine's user account and the client that launches the process. There is no
  authentication or authorization layer inside the server: whoever can start the
  child process and write to its stdin can call the tools.
- `stdout` is reserved for protocol traffic. The audit log and all logging go to
  files or `stderr`, never to `stdout`, so JSON-RPC framing is never corrupted.

## Windows user permissions

`xlvba` runs as the **interactive Windows user** that launched the MCP client
(and therefore the child server process). It does not switch accounts, elevate,
or install a service. Consequently it needs:

| Requirement | Used for |
|-------------|----------|
| Read access to the workbook files | Opening the source file to copy it into `work/` |
| Write access to the runtime directories | `work/`, `snapshots/`, `artifacts/`, `runs/`, and `audit.jsonl` (by default under `%TEMP%\excel_vba_mcp\`) |
| Write access to the snapshot/artifact roots | Snapshot manifests, SHA-256 files, VBA export, test-run evidence |
| COM/VBE access to the desktop Excel installation | Driving the hidden Excel instance and reading the VBA project |

If the runtime directories are not writable, tools fail with a structured
`SNAPSHOT_FAILED` or a filesystem error; they do not silently fall back to
another location. If the user cannot read a workbook, `excel_open` fails when
copying the file. The server never runs elevated and never writes to the
original workbook.

## Excel macro security

Excel's own macro security remains in force. The server does not weaken it and
does not change it permanently.

- **Trust Center macro settings, Protected View, and "block macros from the
  internet" (Mark-of-the-Web)** are evaluated by Excel when the working copy is
  opened. A file downloaded from the internet may open with macros disabled or
  in Protected View regardless of the server's session mode.
- The server **never** changes Trust Center settings, the registry, Group Policy,
  or the application's `AutomationSecurity` permanently. It only sets
  `AutomationSecurity` inside a narrow context manager and restores the previous
  value in `finally`:
  - `mode="inspect"` sets `AutomationSecurity = ForceDisable` (3) for the open, so
    macros never run while inspecting.
  - `mode="execute"` applies `ByUI` (2) immediately before `Workbooks.Open`, so
    Excel's normal UI/macro policy decides whether macros are enabled.
- `DisplayAlerts = False` is set only around close/quit style operations and only
  suppresses Excel's built-in alert prompts. It does **not** block `MsgBox`,
  `UserForm.Show`, `FileDialog`, add-in UI, or Win32 dialogs.
- The server never calls `.Activate`, `.Select`, `ShowWindow`,
  `SetForegroundWindow` or `SendKeys`; the owned instance stays `Visible=False`.

Corporate policy (Group Policy, Intune, registry) can still block macros or VBOM
on a managed machine. When that happens the server reports it truthfully (for
example `ACCESS_VBOM_DISABLED` or a macro failure) rather than bypassing it.

## VBA project object model trust (VBOM)

Reading or editing VBA requires **"Trust access to the VBA project object
model"** (VBOM) to be enabled in Excel Trust Center → Macro Settings. This is a
**manual, per-machine** setting.

- If VBOM trust is off, every VBA/VBIDE tool and `vba_compile` returns
  `ACCESS_VBOM_DISABLED`; the workbook model still returns its non-VBA sections
  with `vba_access = "vbom_disabled"`.
- The server never enables VBOM, never edits the registry, and never attempts to
  bypass the setting. The remedy is always manual: enable the setting, restart
  Excel, and retry. See [Troubleshooting](10-troubleshooting.md).

## VBA code execution risks

This is the highest-risk capability in the server:

- `macro_run` executes **real VBA** through `Application.Run`. A macro can write
  files, call Win32 APIs, access the network, send email, or modify anything the
  Windows user can reach.
- `vba_replace_component_source` and `vba_insert_code` modify the real VBA
  project stored in the working copy. A later `excel_save` persists those changes
  into the working-copy file (never the original).
- **Snapshots do not undo external macro side effects.** A snapshot restores the
  workbook file only; database writes, emails, network calls, files and printing
  performed by a macro are not rolled back. This is recorded explicitly in the
  snapshot manifest (`external_side_effects_not_undone: true`), the audit log and
  test-runner results.
- `macro_run` does **not** require a prior `snapshot_create` (a documented spec
  deviation). The recommended workflow still snapshots first so there is a
  rollback baseline for the workbook file.

Mitigations that are always active: the macro must be classified
`background_safe` for the current workbook hash **and** VBA source revision, and
must pass the interactive-UI preflight (see below). These reduce the risk of
accidental dialogs and unreviewed macros; they do not make arbitrary VBA safe.

## Filesystem access and path restrictions

Every path that reaches Excel COM passes through
`PathPolicy.validate_path` (`security/path_policy.py`). The policy:

1. Canonicalises the path with `Path.resolve` and cross-checks it with
   `os.path.realpath`, so symlink/junction/reparse-point escapes are detected.
2. Requires the resolved target to stay inside one of the configured roots.
3. Restricts workbook files to an explicit extension set.

The root allowlist comes from `EXCEL_VBA_MCP_ROOTS` (an `os.pathsep`-separated
list). When unset it defaults to the current working directory **and** the system
temp directory, which is convenient locally but broad; set it explicitly in
production. The artifact directory is additionally registered under the logical
root name `artifacts`.

Allowed workbook extensions:

```text
.xlsm  .xlsb  .xlsa  .xlam  .xltm  .xltx  .xlsx  .xls
```

Under the `snapshots` and `artifacts` roots, `.json` and `.sha256` are also
accepted (manifests and checksums).

Rejected categories, reported as `PATH_NOT_ALLOWED` with a `reason` in the error
`context`:

| `reason` | Meaning |
|----------|---------|
| `empty_path` | Empty path. |
| `device_path_denied` | Device path (`\\?\`, `\\.\` and variants). |
| `unc_denied` | UNC path (`\\server\share`) when `allow_unc` is disabled. |
| `traversal` | `..` traversal that escapes a root. |
| `outside_roots` | Path outside all allowed roots. |
| `symlink_escape` | Lexically inside a root but resolves outside (symlink/junction/reparse point). |
| `root_mismatch` | Path not in the expected logical root. |
| `extension_not_allowed` | File extension is not in the allowed set. |

The `allow_unc` option is always `False` (it is defined in code and is not
exposed as an environment variable), so UNC and device paths are always denied.

## Access model

The server has two access modes, selected by `EXCEL_VBA_MCP_ACCESS_MODE`:

| Mode | Behavior | Benefit | Risk |
|------|----------|---------|------|
| `full` (default) | `MUTATING` and `EXECUTING` operations are authorized by policy without a per-action grant. `excel_open` opens writable/execute by default; guarded mutations run after a snapshot; `macro_run` needs no per-run grant. | Usable out of the box; no in-band grant step is needed. | Any client that can call the tools can mutate workbooks and run classified macros without a human approval step. |
| `restricted` (opt-in) | Deny-by-default: `MUTATING`/`EXECUTING` operations are refused with `APPROVAL_DENIED` unless an in-process allowlist entry or live grant matches. | Strong default-deny posture for untrusted automation. | No MCP tool grants approval or edits the allowlist, so through MCP alone only read-only work is available; `excel_open` with `mode="execute"` or `read_only=false` surfaces as `APPROVAL_DENIED`. |

What applies in **both** modes:

- **Path policy** (`PATH_NOT_ALLOWED`) — no path escapes the allowlist.
- **Snapshot-before-mutation** — a mutating tool without a successful snapshot
  fails with `SNAPSHOT_FAILED`.
- **Postcondition re-read** — mutations re-read the target and set `verified`
  from the actual state, not merely from the absence of an exception.
- **Interactive-UI veto** — a macro whose source hits the R13 token scan is
  blocked with `INTERACTIVE_UI_NOT_SUPPORTED` and is never dispatched.
- **Audit** — every operation is written to the append-only JSONL audit log.

`READ_ONLY` operations are always auto-allowed. Neither mode bypasses the path
policy, Office trust, snapshot requirement or audit.

## Input validation

Tool arguments are validated before any Excel call:

- **Typed handlers with explicit guards.** Each tool is a typed async handler
  whose arguments the MCP SDK turns into an input schema; the handler then
  validates values with explicit guards (`_require_non_empty`, `_require_int`,
  `_require_bool`, `_require_range_text`, `_require_shift`) before any Excel
  call, and rejects bad input with `INVALID_RANGE`.
- **Strict payload/DTO models.** Structured nested inputs and outputs use
  Pydantic models with `extra="forbid"` (domain models, `RangeWriteRequest`,
  `FormulaWriteRequest`, the declarative suite schema), so unknown fields in a
  payload are rejected rather than ignored.
- **Bounded 2-D range payloads** (`RangeWriteRequest`, `FormulaWriteRequest`):
  the payload shape must match the target, and the cell count is capped
  (`MAX_WRITE_PAYLOAD_CELLS = 100_000`).
- **Explicit `shift`** for `range_insert` / `range_delete` (`right`/`down` and
  `left`/`up`). A missing or unknown shift is rejected and never guessed.
- **255-character cap** on `Formula1`/`Formula2` for data validation; the
  formulas are treated as opaque strings and never evaluated in Python.
- **Unique-header / unambiguous-filter checks**: `table_add_column` and
  `table_delete_column` resolve one unique header; `table_update_rows` /
  `table_delete_rows` reject a `where` filter that matches zero or multiple rows
  with `AMBIGUOUS_FILTER` instead of guessing.
- **Size ceilings**: range reads are capped at 100,000 cells per target and
  50,000 cells per response; table row reads are additionally capped at 2,000
  cells per response. Oversized requests are rejected, not silently clipped.

## Logging and sensitive data

The audit trail is an **append-only JSONL** file at `EXCEL_VBA_MCP_AUDIT_LOG`
(default `<temp>/excel_vba_mcp/audit.jsonl`). Writes are serialized with a
`threading.Lock`; the file is never written to `stdout`.

Redaction is mandatory and intentionally lossy:

- Values under secret-looking keys are replaced with `"<redacted>"`. The
  reserved keys are `password`, `passwd`, `pwd`, `secret`, `token`,
  `connection_string`, `conn_str`, `api_key`, `credential` and `auth`
  (case-insensitive).
- Path-like values are reduced to their basename.
- Macro arguments are logged as typed, length-capped summaries (no COM
  references).
- Nested mappings and lists are scrubbed recursively.

The audit record captures `schema_version`, `timestamp` (UTC), `event_id`,
`correlation_id`, `session_id`, `workbook_id`, `workbook` (canonical path +
SHA-256 + file name), `tool`, `operation_class`, redacted `params`, `approval`,
`snapshot_id`, `result`, `user`, `process` and `excel_version` where available.

## Credential handling

- The server **stores no credentials** and has no credential store.
- Configuration contains **no secrets**; it is limited to paths and the access
  mode (see [Configuration](03-configuration.md)).
- **Do not pass secrets as tool arguments.** They are redacted in the audit log,
  but they still travel through the MCP client, the server process and possibly
  the client's own logs. Use the operating system's secret storage or an
  external secret manager if a macro needs a credential.

## Third-party dependencies

Runtime dependencies (`pyproject.toml`, locked in `uv.lock`):

| Package | Constraint | Purpose |
|---------|------------|---------|
| `mcp` | `>=2.3,<3` (locked 2.3.0) | MCP server, STDIO transport, tool/resource registration |
| `pydantic` | `>=2.0.0` | Typed configuration and DTOs |
| `pywin32` | `>=306` | Excel COM automation |
| `comtypes` | `>=1.2.0` | Read-only VBE compile UI diagnostics |

Development dependencies (`dev` group): `pytest`, `ruff`, `mypy`.

Dependency risk is inherited from these projects. Pin exact versions through
`uv.lock` and review dependency updates deliberately; do not use a wildcard
upgrade in production. `pywin32` and `comtypes` are Windows-only and are the
main reason the server is Windows-only.

## Recommended deployment practices

These are **not enforced** by the application; they are operator choices:

- Set `EXCEL_VBA_MCP_ROOTS` to a **dedicated directory** that contains only the
  workbooks the server should touch. Do not rely on the CWD + temp default.
- Run the server under a **dedicated Windows account** with the minimum
  privileges needed (read workbooks, write runtime dirs).
- Use `EXCEL_VBA_MCP_ACCESS_MODE=restricted` for untrusted or automated callers,
  accepting that only read-only work is available over MCP.
- **Review VBA before running it.** Classify a procedure `background_safe` only
  after reading it (`vba_read_component` / `vba_list_procedures`).
- Keep Excel and Windows **patched**; macro and Protected View behavior changes
  across builds.
- Protect the runtime directories and `audit.jsonl` from tampering (filesystem
  ACLs); the audit log is append-only by convention, not tamper-evident.
- Prefer `mode="inspect", read_only=true` for any workflow that does not need
  mutations or macros.
- Treat the MCP client's own logs and configuration as sensitive: they may
  contain workbook paths and tool arguments.

## Risk table

| Risk | Cause | Potential impact | Existing protection | Recommendation |
|------|-------|------------------|---------------------|----------------|
| Malicious macro runs | `macro_run` executes real VBA | Data theft, file/network/registry changes, lateral movement | `background_safe` classification bound to workbook hash + source revision; R13 interactive-UI veto; audit | Review VBA first; run under a least-privilege account; keep `restricted` mode for untrusted callers |
| Destructive mutation | A mutating tool changes or deletes workbook content | Data loss in the working copy | Snapshot-before-mutation (`SNAPSHOT_FAILED`), postcondition re-read, audit | Snapshot explicitly; prefer `inspect` for read-only work; save only deliberately |
| Path escape | A path outside the allowlist (traversal, symlink, UNC, device) | Read/write outside intended roots | `PathPolicy.validate_path`, canonicalisation, `allow_unc=False`, extension allowlist | Set `EXCEL_VBA_MCP_ROOTS` to a dedicated root; keep `allow_unc` off |
| Data exfiltration via working copies | `work/` copies and snapshots/artifacts may contain sensitive data | Sensitive data left on disk under `%TEMP%` | Originals are never modified; runtime dirs are operator-controlled | Use a dedicated runtime dir; restrict ACLs; clean up after use |
| Credential leakage in arguments | Secrets passed as tool arguments | Secrets in client logs, process memory, or audit context | Audit redaction of secret-looking keys; basename-only paths | Never pass secrets as tool arguments; use an external secret store |
| Macro dialog hang | A macro shows `MsgBox`/`UserForm`/`FileDialog` and blocks COM | Worker `NEEDS_RECOVERY`, hung server, possible orphaned Excel | R13 preflight veto; soft timeout; only the owned PID is terminated (`dismissed=false`) | Keep macros background-safe; monitor `excel_health`; close/reopen or restart |
| Untrusted workbook | Opening a workbook from an untrusted source | Macro/Protected View exploitation, external links | `inspect` disables macros; `execute` honors Office policy; no GUI | Open untrusted files with `inspect` + `read_only`; keep Office patched |
| Dependency compromise | A compromised `mcp`/`pywin32`/`comtypes` release | Arbitrary code in the server process | Pinned versions in `uv.lock` | Pin and review updates; rebuild the venv from the lock |
| Log tampering | Someone edits or truncates `audit.jsonl` | Loss of accountability | Append-only writes; serialized with a lock | Protect with filesystem ACLs; ship logs off-host; do not rely on it as tamper-evident |
| Restricted-mode misconfiguration | Expecting deny-by-default while running `full` (default) | Mutations/macros run without per-action approval | Startup validation of `EXCEL_VBA_MCP_ACCESS_MODE` | Set `EXCEL_VBA_MCP_ACCESS_MODE=restricted` explicitly for untrusted callers |
| Stale worker / lost Excel | The hidden `EXCEL.EXE` dies or hangs | Calls fail or hang; `needs_recovery` | Soft timeouts; `excel_health` owned-pid liveness probe; truthful `EXCEL_HUNG` | Monitor `excel_health`; close/reopen or restart the server (no in-band recovery tool) |
| Over-broad roots | Default roots are CWD + temp | Any workbook under CWD/temp is reachable | Path policy still canonicalises and checks roots | Always set an explicit, dedicated `EXCEL_VBA_MCP_ROOTS` |

## Control categories

### Implemented security controls (in code)

- Path allowlist and canonicalisation (`security/path_policy.py`).
- Approval policy with `full` / `restricted` modes (`security/approval.py`).
- Snapshot-before-mutation and postcondition re-read in every content-mutating
  service (`macro_run` is the documented exception; see the gaps below).
- DTO return boundary (`_ensure_dto`, `PROXY_LEAK_SUSPECTED`) — COM proxies never
  cross the worker boundary.
- `background_safe` registry bound to `(workbook hash, source hash, procedure)`.
- R13 interactive-UI preflight (heuristic veto, never a certificate).
- Append-only JSONL audit with mandatory redaction.
- Structured errors with stable `code`/`context` (`ExcelVbaMcpServer.call_tool`).
- Non-activating compile (fail-closed `indeterminate`).
- Bounded reads and payload ceilings.
- Hidden, non-activating Excel instance; no `.Activate`/`.Select`/`SendKeys`.
- Single-STA COM worker with bounded retry for reads only and no blind replay of
  mutations.

### Controls that depend on Windows / Office settings

- **VBOM trust** ("Trust access to the VBA project object model") — manual, per
  machine; without it VBA tools return `ACCESS_VBOM_DISABLED`.
- **Macro settings / Trust Center** — determine whether `execute` mode actually
  enables macros.
- **Protected View and Mark-of-the-Web** — can block macros or opening entirely.
- **Corporate policy** (Group Policy, Intune, registry) — can override or disable
  VBOM and macro execution.

### Recommended practices NOT enforced by the application

- Dedicated roots and runtime directories.
- A dedicated, least-privilege Windows account.
- `restricted` mode for untrusted automation.
- Reviewing VBA before classifying or running it.
- Patching Windows and Office.
- Protecting the runtime directories and audit log with ACLs.
- Shipping audit logs to an append-only or off-host sink.

### Security gaps requiring attention

- **Default `full` mode.** Mutations and macro execution are authorized without a
  per-action grant. This is a deliberate usability choice, not a sandbox.
- **No per-action grant tool.** In `restricted` mode there is no in-band way to
  grant approval or edit the allowlist; only an in-process embedder can
  pre-authorize. Over MCP, restricted mode means read-only.
- **R13 token scan is heuristic.** Dynamic `CallByName` / `ExecuteGlobal` / `Eval`
  indirection and UI reached transitively through other procedures are not fully
  covered. The veto list does include window/focus and visibility tokens
  (`AppActivate`, `Application.Visible`, `.Activate`, `WindowState`), but a
  substring scan can still miss indirect UI.
- **`EXCEL_VBA_MCP_LOG_LEVEL` is parsed but not applied.** Logging is fixed at
  `INFO` on `stderr`; the variable has no effect in this revision.
- **`macro_run` has no snapshot precondition.** Unlike content-mutating tools it
  does not require a prior `snapshot_create`; this is a documented spec
  deviation.
- **Session recovery is not exposed over MCP.** `excel_health` reports
  `needs_recovery`, but a client cannot trigger recovery in-band; the remedy is
  close/reopen or restarting the server.
- **`macro_policy.classify_operation` is non-authoritative.** The function is
  retained for illustration, is not wired into enforcement, and its table is
  stale. Use the service code or [MCP tools reference](05-mcp-tools-reference.md)
  for the authoritative classification.
- **Audit is append-only by convention, not tamper-evident.** There is no
  cryptographic chaining or remote sink.

## Related

- [Overview](01-overview.md) — what the server is and its non-goals.
- [Installation and setup](02-installation-and-setup.md) — Windows/Excel preparation.
- [Configuration](03-configuration.md) — environment variables and runtime paths.
- [Architecture](04-architecture.md) — request lifecycle and guard chain.
- [MCP tools reference](05-mcp-tools-reference.md) — tool-level errors and access classes.
- [Excel capabilities](06-excel-capabilities.md) — capability coverage and limits.
- [Workflows and examples](07-workflows-and-examples.md) — safe workflow recipes.
- [Limitations and compatibility](09-limitations-and-compatibility.md) — known limits.
- [Troubleshooting](10-troubleshooting.md) — diagnosing security-related errors.
- [Development and maintenance](11-development-and-maintenance.md) — contributor guide.