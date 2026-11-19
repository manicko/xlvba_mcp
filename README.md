# Excel VBA MCP Server (Stage 1)

Local MCP server (STDIO) for controlled inspection and automation of Microsoft Excel workbooks on Windows.

## Scope

This repository follows the approved technical specification for Stage 1 (see `docs/01-Spec/excel_vba_mcp_technical_spec.docx`). Stage 2 is out of scope for implementation planning in this repository.

## Requirements

- Windows 11 x64 (target)
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
