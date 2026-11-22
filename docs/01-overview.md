---
id: xlvba-overview
domain: 01-overview
tags:
  - overview
  - excel
  - vba
  - mcp
  - windows
related:
  - xlvba-installation-and-setup
  - xlvba-configuration
  - xlvba-architecture
  - xlvba-mcp-tools-reference
  - xlvba-excel-capabilities
  - xlvba-workflows-and-examples
  - xlvba-security
  - xlvba-limitations-and-compatibility
  - xlvba-troubleshooting
  - xlvba-development-and-maintenance
---

## Purpose

This document describes what `xlvba` (package `excel-vba-mcp`) is, who it is for,
what it does and what it does not do. It is the entry point to the project documentation;
details of installation, configuration and individual tool groups are moved to child documents.

## What is xlvba

`xlvba` is a local **MCP server** that runs over the **STDIO** transport and
drives a real, **hidden** instance of Microsoft Excel via COM on Windows.
The server has no user interface of its own: it exposes a set of MCP tools
and resources that an MCP client (for example, Kilo or Claude Desktop) invokes.

Key facts:

- Package: `excel-vba-mcp`, version `0.1.0` (Proprietary license).
- The server registers **88 tools** and **4 read-only resources**.
- Transport — MCP over STDIO: protocol messages go to `stdout`, diagnostics and
  tracebacks go to `stderr`.
- The Excel instance is created via `win32com.client.DispatchEx` and stays `Visible=False`
  for its entire lifetime.
- Clients expose the published tool names with a prefix
  (for example, in Kilo — `xlvba_excel_open`).

A detailed description of the tools is in [MCP Tools Reference](05-mcp-tools-reference.md),
and of Excel capabilities in [Excel Capabilities](06-excel-capabilities.md).

## Key Capabilities

- **Workbook sessions and working copies.** `excel_open` copies the source file into the `work/`
  directory and opens only the copy; the original is never modified. By default a session is
  writable and execution-capable (`mode="execute"`, `read_only=False`); pass `mode="inspect"`
  and/or `read_only=True` for a macros-disabled or read-only session.
- **Read-only workbook model.** `workbook_get_model` aggregates sheets, tables, names, shapes,
  controls and (optionally) VBA components and procedures.
- **Sheets.** `sheets_list`, `sheet_get`, plus the mutating operations `sheet_add`,
  `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility`.
- **Tables (ListObject).** `tables_list`, `table_get`, `table_read_rows`,
  `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`,
  `table_add_column`, `table_delete_column`, `table_create`, `table_resize`.
- **Ranges and cells.** Reading and writing values and formulas, reading and partial writing of
  formatting, data validation, conditional formatting, plus
  `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`.
- **PivotTables (existing only).** Enumeration, reading metadata and caches, refreshing,
  field management (orientation, position, aggregation, properties), filters via visible items
  and clearing filters. Creating and deleting PivotTables/PivotCaches is not supported.
- **Names (Named Ranges).** `names_list`, `name_get` and the mutating `name_create`,
  `name_update`, `name_delete`.
- **Shapes and controls.** `shapes_list`, `shape_get`, `shape_set_on_action`,
  `controls_list`, `control_get`, `control_set_on_action`.
- **VBA.** Project and component inventory, source reading, code search, project export,
  reference reading, plus guarded source replacement and insertion.
- **Compile diagnostics (fail-closed).** `vba_compile` and `vba_get_compile_diagnostics`
  return partial diagnostic data but never report a false success.
- **Macros (under a gate).** `macro_list`, `macro_run`, `macro_get_last_result` and
  `macro_classify_background`; running requires an execute-mode session, a `background_safe`
  classification for the current workbook hash and VBA source revision, and a passing
  interactive-UI preflight. In the default `full` access mode no per-run approval grant is needed;
  in `restricted` mode the approval gate also applies.
- **Snapshots.** `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`
  with a manifest and SHA-256.
- **Declarative test runner.** `test_run_suite`, `test_suite_run`, `test_get_run`
  with a closed vocabulary of actions and assertions.
- **Audit and security.** The `audit.jsonl` journal (append-only), a path allowlist and an
  access-mode policy: full access by default, restricted mode (deny-by-default approval) as an
  explicit opt-in.

Each group is described in more detail in [Excel Capabilities](06-excel-capabilities.md);
threats and boundaries are in [Architecture](04-architecture.md) and
[Excel capabilities](06-excel-capabilities.md).

## How an AI Agent Interacts with Excel via xlvba

The agent never sees a visible Excel window and does not drive the GUI. Instead, this loop runs:

1. The MCP client sends a tool call over STDIO (JSON-RPC).
2. The server validates the arguments (Pydantic models, paths through the path policy, approval gate).
3. The request is handed to a single dedicated **STA thread** (`ExcelComWorker`), where all
   COM access happens.
4. Excel code performs the operation in a hidden, non-activating Excel instance.
5. The result is returned as a DTO (typed data, no "raw" COM objects), then serialized into the
   MCP response.

Thus all COM access is serialized on one STA thread; tools never touch COM directly, and only
DTOs cross between the worker and the rest of the code. Error handling and worker recovery are
described in [Architecture](04-architecture.md).

## Typical Scenarios

1. **Inspect a workbook.** `excel_open` (mode `inspect`, `read_only=true`) → `workbook_get_model`
   or `sheets_list` → `excel_close`.
2. **Read a range.** `excel_open` → `range_read_values` / `range_read_formulas` /
   `range_get_format` → `excel_close`.
3. **Inventory VBA.** `excel_open` (inspect) → `vba_project_info` / `vba_list_components`
   / `vba_list_procedures` → `excel_close`.
4. **Snapshot before a change.** `snapshot_create` (for an open workbook `Workbook.SaveCopyAs`
   is used, for a closed one a file copy).
5. **Export VBA.** `vba_export_project` with a manifest and file hashes saved.
6. **Run a declarative test.** `test_suite_run` (or `test_run_suite`) → if needed
   `test_get_run` to read a persisted run.
7. **Diagnose compile state.** `vba_compile` / `vba_get_compile_diagnostics` —
   obtaining status and partial diagnostic data without launching the UI.

Step-by-step examples with arguments are in [Workflows and Examples](07-workflows-and-examples.md).

## What xlvba Does NOT Do

- It does not automate the GUI and does not show a visible Excel window.
- It does not create new workbooks (`excel_open` only opens existing files).
- It does not create or delete PivotTables/PivotCaches; it does not create charts, pivot charts or slicers.
- It does not expose a tool to grant per-action approvals or edit an allowlist over MCP. In the
  default `full` access mode no grant is needed (mutations and execute mode are authorized); in
  `restricted` mode the deny-by-default gate refuses `MUTATING`/`EXECUTING` operations with
  `APPROVAL_DENIED`.
- It does not report a successful compile on the validated build: `vba_compile` always returns
  `status=indeterminate` (`reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`).
- It does not run macros without a `background_safe` classification bound to the current workbook
  hash and VBA source revision; classify a procedure with `macro_classify_background`. Interactive
  macros are vetoed by the preflight (`INTERACTIVE_UI_NOT_SUPPORTED`).
- It does not support cloud scenarios or Office 365 (it works only with desktop Excel via COM).
- It does not support non-Windows platforms (COM and VBOM are available only on Windows).

The full list of limitations is in [Excel capabilities](06-excel-capabilities.md).

## Core Dependencies

| Component | Requirement |
|-----------|------------|
| OS | Windows 10/11 x64 (validated on Windows 10 Enterprise LTSC 2019, build 17763) |
| Excel | Desktop Microsoft Excel with COM/VBE (validated on 16.0.14332 x64) |
| Python | `>=3.11,<3.15` (validated on 3.14 x64) |
| Package manager | [uv](https://docs.astral.sh/uv/) |
| `mcp` | `>=1.0.0` (locked to 2.3.0) |
| `pydantic` | `>=2.0.0` (locked to 2.14.4) |
| `pywin32` | `>=306` (locked to 312) |
| `comtypes` | `>=1.2.0` (locked to 1.5.0) |
| Excel setting | Trust Center: "Trust access to the VBA project object model" (VBOM), set manually on each machine |

The installation and verification procedure is in [Installation and Setup](02-installation-and-setup.md),
environment variables are in [Configuration](03-configuration.md).

## Documentation

- [Installation and Setup](02-installation-and-setup.md)
- [Configuration](03-configuration.md)
- [Architecture](04-architecture.md)
- [MCP Tools Reference](05-mcp-tools-reference.md)
- [Excel Capabilities](06-excel-capabilities.md)
- [Workflows and Examples](07-workflows-and-examples.md)
- [Root README](../README.md)