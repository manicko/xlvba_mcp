# Excel VBA MCP Server (Stage 1)

Local MCP server (STDIO) for controlled inspection and automation of Microsoft Excel workbooks on Windows.

## Scope

This repository follows the approved technical specification for Stage 1 (see `docs/01-Spec/excel_vba_mcp_technical_spec.docx`). Stage 2 is out of scope for implementation planning in this repository.

## Requirements

- Windows 10/11 x64 (spec target Windows 11; validated on Windows 10 Enterprise LTSC 2019, build 17763, x64)
- Python 3.14 x64 (validated runtime; project floor is 3.11 — 3.11–3.13 are installable but unvalidated)
- [uv](https://docs.astral.sh/uv/)
- Microsoft Excel installed for integration work (COM/VBE/pywinauto paths are Windows-only)

## Setup

```powershell
uv sync --locked
```

If `uv.lock` is missing or you need to refresh dependencies after changing `pyproject.toml`, run `uv lock` then `uv sync` in a controlled way and commit the updated lockfile per project rules.

## Run (STDIO)

```powershell
uv run excel-vba-mcp
```

Or:

```powershell
uv run python -m excel_vba_mcp
```

The server writes diagnostics to stderr only; stdout is reserved for MCP protocol messages.

## Use as an MCP Server

The server speaks the Model Context Protocol over **STDIO**: protocol messages go to
stdout, diagnostics go to stderr. Any MCP-capable client (Kilo, Claude Desktop,
Cline, etc.) can launch it as a local (stdio) server.

### Install / prepare

```powershell
uv sync --locked
```

That creates `.venv/` and the `excel-vba-mcp` console script. No global install or
Node.js/`npx` is involved — this is a Python project managed by `uv`.

### Connect to an MCP client (Kilo example)

Kilo's `mcp` block is runner-agnostic: `type: local` spawns the `command` array as a
child process. The Python analogue of an `npx` command is `uv run` with
`--directory` pointing at the repository root (this works regardless of the client's
working directory):

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
      "environment": {
        "EXCEL_VBA_MCP_LOG_LEVEL": "INFO"
      },
      "enabled": true,
      "timeout": 60000
    }
  },
  "permission": {
    "xlvba_*": "ask"
  }
}
```

Notes:

- Replace `<absolute-path-to-repo>` with the absolute path to this repository's root.
- `environment`: only `EXCEL_VBA_MCP_LOG_LEVEL` is supported today (`DEBUG`, `INFO`,
  `WARNING`, `ERROR`, `CRITICAL`; default `INFO`). It controls stderr logging only.
- `timeout`: allow extra time on first launch (dependency sync / venv warm-up) and for
  slow Excel COM calls; raise it further if you hit startup timeouts.
- `permission`: tools drive real Excel COM automation and can modify workbooks, so
  default them to `"ask"`. Whitelist safe tools later, e.g. `"xlvba_server_health":
  "allow"`.
- In the Kilo TUI, `/mcps` toggles loaded servers; their tools are exposed as
  `xlvba_<tool>`.

### Alternate launch commands

All of the following are valid `type: local` commands:

- Direct console script (requires a prior `uv sync`):
  ```json
  "command": ["<repo-root>\\.venv\\Scripts\\excel-vba-mcp.exe"]
  ```
- Module form:
  ```json
  "command": ["uv", "run", "--directory", "<absolute-path-to-repo>", "python", "-m", "excel_vba_mcp"]
  ```
- Global install (optional), then a clean command:
  ```powershell
  uv tool install --editable <absolute-path-to-repo>
  ```
  ```json
  "command": ["excel-vba-mcp"]
  ```

## Development

```powershell
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src
uv run pytest
```

## Notes

- Excel COM automation requires an interactive Windows session in practice for many UI flows (VBE compile dialogs, etc.). Headless/server-style guarantees are limited per spec.
- Do not open or modify original workbooks in `projects/` for experiments; use disposable copies in `work/` per project policy.
