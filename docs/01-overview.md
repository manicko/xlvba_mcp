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

This document explains what `xlvba` (package `excel-vba-mcp`) is, who it is for, which Excel
problems it solves, and when to use it instead of an alternative. It is the entry point to the
project documentation; installation, configuration and per-tool detail live in child documents.

## What is xlvba

`xlvba` is a local **MCP server** that gives an AI agent the tools to develop, run and test VBA
macros directly in Microsoft Excel on Windows. The server has no user interface of its own: it
exposes **43 MCP tools** and **5 read-only MCP resources** (4 data-resource templates plus a static
glossary) that an MCP client (for example, Kilo or Claude Desktop) invokes.

The Excel instance the server drives is **hidden** — there is no visible window and no keystroke
input — so the agent works on a real workbook without disturbing your desktop.

Key facts:

- Package: `excel-vba-mcp`, version `0.1.0` (Proprietary license).
- **43 MCP tools**: 20 dispatch tools (with an `action`) plus 23 standalone tools.
- **5 read-only MCP resources**: the workbook model, VBA component inventory, VBA component source,
  test-run summary, and the static `xlvba://glossary`.
- Transport: MCP over STDIO (protocol on `stdout`, diagnostics on `stderr`).
- Clients show the published tool names with a prefix (for example, in Kilo — `xlvba_workbook`).

A detailed description of the tools is in [MCP Tools Reference](05-mcp-tools-reference.md) and of
Excel capabilities in [Excel Capabilities](06-excel-capabilities.md).

## Who it is for

- **Analysts and PMs** who want an AI agent to build and verify a macro without pasting code into
  Excel by hand.
- **Developers** who want controlled, revision-safe VBA edits and repeatable tests against real
  workbooks.
- **Teams automating Excel on Windows** who need the original file to stay untouched, writes to be
  verified, and macros to be gated before they run.

You describe the task and the expected result; the agent handles the technical steps.

## Which Excel problems it solves

- **Stop copy-pasting VBA.** The agent writes code, inserts it into a workbook, runs it, reads the
  result and fixes issues — without a human moving code and errors between Excel and the assistant.
- **Work on real Excel, safely.** Every open works on a copy; the original is never modified.
- **Bulk-edit without manual cell work.** Update thousands of rows in an Excel Table, page through a
  huge range, or export a large dataset to a file.
- **Verify before you trust.** Snapshots let you roll back a bad change; a declarative test runner
  checks edit→macro→assert scenarios on disposable copies.
- **Run macros under control.** A macro runs only when it is classified `background_safe` and passes
  an interactive-UI preflight, so an accidental dialog cannot hang the server.

## Key capabilities

- **Workbook sessions and working copies.** `workbook` (`open`, `close`, `save`) copies the source
  into a working copy and opens only the copy; the original is never modified. `workbook_read`
  (`status`, `list`, `model`) reports session state or an aggregated model.
- **Save As copy.** `workbook_save_copy` writes a verified copy to a new (or the original) path.
- **Sheets, tables and names.** `sheet_read` / `sheet` / `sheet_delete`; `table_read` / `table` /
  `table_delete_rows` / `table_delete_column`; `defined_name_read` / `defined_name` /
  `defined_name_delete`.
- **Ranges and cells.** `range_read` / `range_write` for values, formulas, format, validation and
  conditional formatting; `range_structural` / `range_delete` / `range_merge` for structure.
- **PivotTables (existing only).** `pivot_read` (including `result-read`), `pivot_refresh`,
  `pivot_field`. Creating/deleting PivotTables/PivotCaches is not supported.
- **Shapes and controls.** `shape` and `control` (list/get, assign `OnAction`).
- **VBA.** `vba_read`, `vba_export_project`, and guarded `vba_replace_component_source` /
  `vba_insert_code`.
- **Compile diagnostics (fail-closed).** `vba_compile_diagnostics` returns partial diagnostics but
  never reports a false success.
- **Macros (under a gate).** `macro` (`list`, `classify-background`, `get-last-result`) and
  `macro_run`.
- **Declarative testing.** `test_suite` (`run`, `get-run`).
- **Snapshots.** `snapshot_create` / `snapshot_list` / `snapshot_restore` / `snapshot_diff`.
- **Large data and jobs.** `data_export`, `workbook_snapshot_read`, `job_get`, `job_request_cancel`.
- **Environment and recovery.** `xlvba_environment`, `server_health`, `excel_health`,
  `excel_recover`.
- **Audit and security.** An append-only `audit.jsonl` journal, a path allowlist and an access-mode
  policy (full access by default; restricted mode as an explicit opt-in).

Each group is described in more detail in [Excel Capabilities](06-excel-capabilities.md); the
security model is in [Security](08-security.md).

## How an AI agent interacts with Excel via xlvba

You never see a visible Excel window and the agent does not drive the GUI. Instead:

1. The MCP client sends a tool call over STDIO.
2. The server validates the arguments, the path allowlist and the approval gate.
3. The operation runs in a hidden, non-activating Excel instance.
4. The result is returned as typed data (never a raw Excel object) and serialized into the MCP
   response.

Because there is **one** Excel instance, a long operation delays later ones; plan long reads,
exports or macros accordingly.

## Typical scenarios

1. **Inspect a workbook.** `workbook` (`open`, `mode="inspect"`, `read_only=true`) →
   `workbook_read` (`model`) or `sheet_read` (`list`) → `workbook` (`close`).
2. **Read a range.** `workbook` (`open`) → `range_read` (`read-values` / `read-formulas`) →
   `workbook` (`close`).
3. **Inventory VBA.** `workbook` (`open`, inspect) → `vba_read` (`project-info` / `list-components`
   / `list-procedures`) → `workbook` (`close`).
4. **Snapshot before a change.** `snapshot_create`.
5. **Export VBA.** `vba_export_project` with a manifest and file hashes.
6. **Run a declarative test.** `test_suite` (`run`) → `test_suite` (`get-run`).
7. **Diagnose compile state.** `vba_compile_diagnostics` — status and partial diagnostics without
   launching the UI.

Step-by-step examples with arguments are in [Workflows and Examples](07-workflows-and-examples.md).

## When to use xlvba — and when an alternative fits

`xlvba` is **not universally better**. It is strongest for *agent-driven, safety-first automation on
real files where the original must stay untouched, writes must be verified, and macros must be
gated*. The table below compares it, objectively, with two common alternatives (facts verified live
from each project's README, 2026-11-28).

| User requirement | xlvba | haris-musa (openpyxl) | sbroenne (COM/.NET) | Fits better |
|---|---|---|---|---|
| Real Excel + VBA execution | Yes (gated) | No | Yes | sbroenne / xlvba |
| Original provably untouched | Yes (copy + SHA) | Atomic, opens file | Opens live file | **xlvba** |
| Snapshots / restore / diff | Yes | No | No | **xlvba** |
| Controlled mutations | Yes | Read-only + formula check | Worker isolation | **xlvba** |
| Macro safeguards | Yes | N/A | Fewer gates | **xlvba** |
| VBA edit + compile diagnostics | Yes (fail-closed) | Write opt-in | Edit / run | xlvba (honesty) |
| Structured testing | Yes (`test_suite`) | No | No | **xlvba** |
| Health / recovery | Yes | No | Worker pool | **xlvba** |
| Large-data paging / export | Yes | Yes | Unbounded | xlvba / haris |
| External / Data-Model pivot read | Displayed result only | Yes | Yes (full) | sbroenne / haris |
| Capability breadth | Narrow | Broad | Broadest | sbroenne / haris |
| Cross-platform / no-Excel | No | Yes | No | **haris** |

**In short:** choose `xlvba` for guarded, original-preserving automation with snapshots, structured
testing and recovery; choose **haris-musa** for no-Excel, cross-platform, broad file-feature work;
choose **sbroenne** for the broadest live-Excel feature set (DAX, Power Query, Data Model) and fewer
safety gates.

## What xlvba does NOT do

- It does not automate the GUI or show a visible Excel window.
- It does not create new workbooks (`workbook` `open` only opens existing files).
- It does not create or delete PivotTables/PivotCaches, charts, pivot charts or slicers.
- It does not expose a tool to grant per-action approvals or edit an allowlist over MCP.
- It does not report a successful compile: `vba_compile_diagnostics` always returns
  `status=indeterminate` (`reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`).
- It does not run a macro without a `background_safe` classification; interactive macros are vetoed
  (`INTERACTIVE_UI_NOT_SUPPORTED`).
- It does not support cloud scenarios or Office 365 (desktop Excel via COM only).
- It does not support non-Windows platforms.

The full list of limitations is in [Limitations and compatibility](09-limitations-and-compatibility.md).

## Requirements

| Component | Requirement |
|-----------|------------|
| OS | Windows 10/11 x64 (validated on Windows 10 Enterprise LTSC 2019, build 17763) |
| Excel | Desktop Microsoft Excel with COM/VBE (validated on 16.0.14332 x64) |
| Python | `>=3.11,<3.15` (validated on 3.14 x64) |
| Package manager | [uv](https://docs.astral.sh/uv/) |
| `mcp` | `>=2.3,<3` (locked to 2.3.0) |
| `pydantic` | `>=2.0.0` (locked to 2.14.4) |
| `pywin32` | `>=306` (locked to 312) |
| `comtypes` | `>=1.2.0` (locked to 1.5.0) |
| Excel setting | Trust Center: "Trust access to the VBA project object model" (VBOM), set manually on each machine |

The installation and verification procedure is in [Installation and Setup](02-installation-and-setup.md);
environment variables are in [Configuration](03-configuration.md).

## Documentation

- [Installation and Setup](02-installation-and-setup.md)
- [Configuration](03-configuration.md)
- [Architecture](04-architecture.md)
- [MCP Tools Reference](05-mcp-tools-reference.md)
- [Excel Capabilities](06-excel-capabilities.md)
- [Workflows and Examples](07-workflows-and-examples.md)
- [Security](08-security.md) — trust boundaries, access model and risk table
- [Limitations and compatibility](09-limitations-and-compatibility.md) — validated platform and known limits
- [Troubleshooting](10-troubleshooting.md) — error codes, diagnostics and recovery
- [Development and maintenance](11-development-and-maintenance.md) — contributor and maintainer guide
- [Root README](../README.md)