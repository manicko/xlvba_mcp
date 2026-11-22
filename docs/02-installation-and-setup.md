---
id: xlvba-installation-and-setup
domain: 02-installation-and-setup
tags:
  - installation
  - setup
  - windows
  - excel
  - uv
related:
  - xlvba-overview
  - xlvba-configuration
  - xlvba-troubleshooting
  - xlvba-mcp-tools-reference
---

## Purpose

This document lets a new user install `xlvba` from scratch and verify that the server
starts and basic read-only operations work. It describes requirements, Windows and Excel
preparation, installation via `uv`, connecting an MCP client, first launch and runtime data
location.

All commands are given for PowerShell (Windows). None of the installation commands modifies the
user's source workbooks.

## System Requirements

| Component | Requirement | Note |
|-----------|------------|------------|
| OS | Windows 10/11 x64 | Validated on Windows 10 Enterprise LTSC 2019 (build 17763). Windows only — COM and VBOM are needed. |
| Excel | Desktop Microsoft Excel with COM/VBE | Validated on 16.0.14332 x64. The universal (Microsoft Store/UWP) version is generally not suitable. |
| Python | `>=3.11,<3.15` x64 | Validated on 3.14 x64 (`.python-version`). |
| uv | Installed [uv](https://docs.astral.sh/uv/) | Package and environment manager. |
| Permissions | Write to the temp directory, access to workbook files | Workbooks must be inside the allowed roots (see configuration). |
| Disk | A few MB for `.venv` and runtime data | `.venv` and `<temp>/excel_vba_mcp/` grow as you work. |

Why Windows only: the server drives Excel via COM (`pywin32`) and uses VBIDE
(VBA project object model). Both mechanisms are available only in desktop Excel on Windows.

## Preparing Windows and Excel

1. Install desktop Microsoft Excel (not the UWP/Store version in general) and make sure the
   app has been launched manually at least once and is activated.
2. Enable trust in the VBA project object model (VBOM):
   - Excel → `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Macro Settings`;
   - check **Trust access to the VBA project object model**;
   - confirm and restart Excel.
3. Note that corporate policies (Group Policy, Intune, registry) may block VBOM or macro
   execution. If the setting is unavailable or reverts, this blocks VBA features.
4. Close all open workbooks and Excel itself before operations that require exclusive file
   access (for example, before restoring from a snapshot).
5. Check macro security settings: Excel may block macros in files
   received from the internet (Protected View, "Block macros from running in Office files from
   the Internet"). This affects execute mode and macro execution.

> [!NOTE]
> `excel_open` defaults to `mode="execute", read_only=False` — a writable, execution-capable
> session. Pass `mode="inspect"` and/or `read_only=True` to opt into a macros-disabled or
> read-only session. `inspect` opens with macros disabled
> (`AutomationSecurity = ForceDisable`); `execute` applies `ByUI` immediately before
> opening. Neither mode changes Trust Center or registry settings.

## Installation

1. Open a terminal at the repository root (where `pyproject.toml` lives).
2. Install the dependencies and create the environment:

   ```powershell
   uv sync --locked
   ```

   The command creates the `.venv/` directory and the `excel-vba-mcp` console script. The `dev`
   group (pytest, ruff, mypy) is installed by default. No global install and no Node.js are required.

3. If `pyproject.toml` changed, deliberately refresh the lock and the environment:

   ```powershell
   uv lock
   uv sync
   ```

## Verifying the Installation

1. Run the server directly — it starts as a STDIO server and waits for JSON-RPC on `stdin`:

   ```powershell
   uv run excel-vba-mcp
   ```

   Expected behavior: the process starts and waits for input; with no incoming messages it
   exits (or keeps waiting). This is normal for a STDIO server. If Excel is unavailable,
   the server still starts (see [Health Checks](#health-checks)).

2. Alternative launch through the module:

   ```powershell
   uv run python -m excel_vba_mcp
   ```

3. Checking for CLI flags. In the current implementation (`src/excel_vba_mcp/__main__.py`)
   there are no CLI flags: the entry point `main()` simply calls `mcp.run(transport="stdio")`
   without parsing arguments. Therefore the `uv run excel-vba-mcp --help` command is not
   documented — there is no `--help` flag.

## Connecting an MCP Client

### Kilo

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

- Replace `<absolute-path-to-repo>` with the absolute path to the repository root.
- Tools are shown with the `xlvba_` prefix (for example, `xlvba_excel_open`).
- The `permission` block sets the client's access policy for the tools: `"ask"` asks for
  permission before each call. Read-only tools can be added to an allowlist later,
  for example `"xlvba_server_health": "allow"`.

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

Kilo also accepts the module form (`uv run --directory <repo> python -m excel_vba_mcp`),
the console script after `uv sync` (`<repo>\.venv\Scripts\excel-vba-mcp.exe`) or a global
install (`uv tool install --editable <repo>`).

## First Launch

The server runs in **full access mode** by default: `excel_open` opens a writable,
execution-capable session and mutating tools and `macro_run` are authorized without a per-action
grant (a baseline snapshot is still required before any mutation). To opt into the
deny-by-default policy, set `EXCEL_VBA_MCP_ACCESS_MODE=restricted` before starting the server; any
other value for that variable fails startup fast. See [Configuration](03-configuration.md).

Minimal read-only sequence (all arguments are JSON). Pass `mode="inspect", read_only=true` to open
read-only even though the default open is writable:

1. Open the workbook as a read-only inspect session. Call **`excel_open`** with arguments:

   ```json
   { "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
   ```

   Expected response contains: `workbook_id` (a server-assigned identifier),
   `source_path` (the original path), `path` (the working-copy path) and `read_only: true`.
   The path must be inside the allowed roots (by default CWD and system temp;
   see [Configuration](03-configuration.md)).

2. Get the list of sheets. Call **`sheets_list`** with arguments:

   ```json
   { "workbook_id": "<workbook_id>" }
   ```

   Expected response — an array of sheets with fields `name`, `index`, `object_type`,
   `visibility`, and for worksheets also `code_name`, `protected`, `table_count`, `used_range`.

3. Read a bounded range. Call **`range_read_values`** with arguments:

   ```json
   { "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
   ```

   Expected response — typed cell values and a `truncated` flag.

4. Close the workbook without saving. Call **`excel_close`** with arguments:

   ```json
   { "workbook_id": "<workbook_id>", "save_changes": false }
   ```

   Expected response confirms the session is closed.

## Health Checks

- **`server_health`** — a read-only liveness probe that never touches Excel/COM.
  Shows that the server is alive and responding.
- **`excel_health`** — worker/session state (`state`, `accepting`, `owned Excel pid`,
  `active operation`, `needs_recovery`) without a forced restart and without touching COM.
- **`excel_list_workbooks`** — the list of open workbooks from session bookkeeping.

What "Excel unavailable" means: the server still starts over STDIO (a worker start failure is
logged to `stderr`, but the process does not crash), `server_health` and the tool listing keep
working, and workbook tools return a truthful error about the session being unavailable rather
than a "false success".

## Runtime Data Location

By default runtime data is written outside the repository, into the system temp directory:

```text
<temp>/excel_vba_mcp/
├─ work/          # working copies of open workbooks
├─ snapshots/     # snapshot store (manifest + SHA-256)
├─ artifacts/     # test-run artifacts and evidence
├─ runs/          # persisted declarative test runs
└─ audit.jsonl    # append-only audit journal (JSONL)
```

`<temp>` is the system temp directory (`%TEMP%` on Windows). The user's source files are never
modified: only the copy from `work/` is opened.

## Updating

1. Pull the changes and refresh the lock if needed:

   ```powershell
   uv lock
   uv sync
   ```

2. Restart the MCP client so it picks up the new server version.

## Uninstalling

1. Remove the virtual environment:

   ```powershell
   Remove-Item -Recurse -Force .venv
   ```

2. If needed, remove the runtime directory:

   ```powershell
   Remove-Item -Recurse -Force "$env:TEMP\excel_vba_mcp"
   ```

Source workbooks and user files are not affected when the server is removed.

## Next Steps

- [Configuration](03-configuration.md) — environment variables, paths, path policy, sessions.
- [MCP Tools Reference](05-mcp-tools-reference.md) — the full list of tools.
- [Excel capabilities](06-excel-capabilities.md) — if Excel is unavailable or VBOM is blocked.