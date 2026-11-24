# xlvba_mcp

**Stop copy-pasting VBA. Let your AI agent work in Excel.**

AI can write VBA code. But turning that code into a working macro usually requires a human to paste it into Excel, run it, check the results, and report errors back to the AI.

**xlvba_mcp closes that gap.**

A local MCP server that gives AI agents the tools to develop and test VBA macros directly in Microsoft Excel — from writing and inserting code to running tests, checking results, and fixing issues.

**You describe the task. Your AI agent handles the development cycle. No VBA knowledge required.**

## What can your AI agent do?

- **Write and edit VBA** — insert, read, search, and modify code in Excel workbooks.
- **Work with real Excel** — manage worksheets, cells, formulas, tables, named ranges, and other workbook elements.
- **Run and test macros** — prepare test data, execute supported macros, inspect results, and validate expected behavior.
- **Debug and iterate** — use diagnostics, snapshots, and repeatable test scenarios to identify and fix issues.

All without manually transferring VBA code, errors, or test results between Excel and your AI assistant.

## See it in action

Give your AI agent a task like this:

> Write a VBA macro that combines data from multiple worksheets into a single summary report, removes duplicates, and calculates totals.
>
> Use the **xlvba_mcp** tools to insert your code into an Excel workbook, prepare test data, run the macro, and verify the results. Test different scenarios, identify and fix any issues, and repeat until the expected behavior is confirmed.
>
> Complete the entire development and testing cycle yourself, then provide me with the finished, tested macro.

**Write → Insert → Test → Verify → Fix → Repeat**

The agent can write VBA code independently and then use an existing macro-enabled Excel workbook to test it. You define the goal and expected results; the agent handles the technical steps.


## Quick Start

### Requirements

| Component | Requirement |
|---|---|
| **OS** | Windows 10/11 x64 (validated on Windows 10 Enterprise LTSC 2019, build 17763) |
| **Excel** | Desktop Microsoft Excel with COM/VBE (validated on 16.0.14332 x64) |
| **Python** | 3.14 x64 validated; supported range `>=3.11,<3.15` |
| **Package manager** | [uv](https://docs.astral.sh/uv/) |
| **VBA features** | Excel Trust Center: **Trust access to the VBA project object model** |

### Step 1. Prepare Excel for VBA automation

**This is a required step.** xlvba_mcp uses Excel's VBA project object model (VBOM) to inspect and modify VBA code programmatically.

By default, Excel restricts external applications from accessing VBA projects. You must explicitly allow this access.

1. Open Microsoft Excel.
2. Go to **File → Options**.
3. Select **Trust Center**.
4. Click **Trust Center Settings**.
5. Select **Macro Settings**.
6. Enable **Trust access to the VBA project object model**.
7. Click **OK** to close both dialogs.
8. Restart Excel to apply the setting.

> [!IMPORTANT]
> Without this setting, xlvba_mcp cannot reliably read or modify VBA projects, preventing the agent from inserting and editing macro code.
>
> This option allows programmatic access to the VBA project; it does **not** require enabling all macros. Keep your normal macro security protections in place.

> [!NOTE]
> On company-managed computers, this setting may be restricted by Group Policy or your IT administrator.
>
> See [Microsoft's macro security documentation](https://support.microsoft.com/en-us/excel/change-macro-security-settings-in-excel).

### Step 2. Install xlvba_mcp

Clone the repository and install the dependencies using `uv`:

```powershell
git clone https://github.com/manicko/xlvba_mcp.git
cd xlvba_mcp
uv sync --locked
```

This creates a local `.venv` environment and installs the `excel-vba-mcp` entry point.

No global Python package installation or Node.js is required.

### Step 3. Configure workbook access

xlvba_mcp restricts workbook access to configured directories.

Set `EXCEL_VBA_MCP_ROOTS` to the directory containing your Excel files.

For example, in PowerShell:

```powershell
$env:EXCEL_VBA_MCP_ROOTS = "C:\ExcelProjects"
```

Multiple directories can be separated with a semicolon (`;`) on Windows.

> [!TIP]
> For MCP clients, define `EXCEL_VBA_MCP_ROOTS` in the client's server configuration, as shown below. Environment variables are loaded when the server starts, so restart the MCP server after changing them.

For more options, see [Configuration](docs/03-configuration.md).

### Step 4. Run the server

Start the MCP server:

```powershell
uv run excel-vba-mcp
```

Alternatively:

```powershell
uv run python -m excel_vba_mcp
```

The server communicates through MCP over **STDIO**.

> [!NOTE]
> The server may appear to do nothing when started manually. This is expected: it waits for requests from an MCP client.
>
> Once configured, your MCP client starts the server automatically.

### Step 5. Connect your MCP client

xlvba_mcp works with compatible STDIO MCP clients, including **Kilo** and **Claude Desktop**.

Replace `<absolute-path-to-repo>` with the full path to your local repository and adjust `C:\\ExcelProjects` to your workbook directory.

#### Option A: Kilo

Add the following to your Kilo configuration:

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
      "timeout": 60000,
      "environment": {
        "EXCEL_VBA_MCP_ROOTS": "C:\\ExcelProjects"
      }
    }
  },
  "permission": {
    "xlvba_*": "ask"
  }
}
```

The permission setting asks for confirmation before tool calls.

In the Kilo TUI, use `/mcps` to view and manage loaded MCP servers.

#### Option B: Claude Desktop

Open the Claude Desktop configuration file on Windows:

`%APPDATA%\Claude\claude_desktop_config.json`

Add the server definition:

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
      ],
      "env": {
        "EXCEL_VBA_MCP_ROOTS": "C:\\ExcelProjects"
      }
    }
  }
}
```

Save the configuration and fully restart Claude Desktop.

> [!TIP]
> If your MCP client cannot find `uv`, replace `"uv"` in the `command` field with the absolute path to `uv.exe`.

### Step 6. Verify the connection

Once your MCP client is connected, ask your AI agent:

> Check the xlvba_mcp server health and list its available tools.

The agent can use `server_health` to verify that the server is responding and `excel_health` to inspect the Excel worker status.

You can now give your agent a VBA development or testing task.

> [!WARNING]
> VBA macros can modify or delete files and affect resources outside the workbook. Only run trusted code, keep backups of important files, and review actions before authorizing them.

## Tool catalog

xlvba_mcp provides **88 MCP tools** and **4 read-only MCP resources**, covering VBA development, workbook automation, macro execution, and testing.

| Group | Count | Tools |
|---|---:|---|
| **Workbook / session** | 6 | `excel_open`, `excel_close`, `excel_save`, `excel_get_status`, `excel_list_workbooks`, `excel_health` |
| **Snapshots** | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` |
| **VBA inventory / editing** | 9 | `vba_project_info`, `vba_list_components`, `vba_list_procedures`, `vba_read_component`, `vba_export_project`, `vba_get_references`, `vba_search_code`, `vba_replace_component_source`, `vba_insert_code` |
| **Compile** | 2 | `vba_compile`, `vba_get_compile_diagnostics` |
| **Macros** | 4 | `macro_list`, `macro_run`, `macro_get_last_result`, `macro_classify_background` |
| **Sheets** | 7 | `sheets_list`, `sheet_get`, `sheet_add`, `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility` |
| **Tables** | 11 | `tables_list`, `table_get`, `table_read_rows`, `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`, `table_add_column`, `table_delete_column`, `table_create`, `table_resize` |
| **Ranges & cells** | 16 | `range_get`, `range_read_values`, `range_read_formulas`, `range_write_values`, `range_write_formulas`, `range_get_format`, `range_set_format`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`, `range_get_validation`, `range_set_validation`, `range_get_conditional_formatting`, `range_set_conditional_formatting` |
| **PivotTables** | 13 | `pivot_tables_list`, `pivot_table_get`, `pivot_cache_get`, `pivot_fields_list`, `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_set_layout`, `pivot_field_set_position`, `pivot_field_set_aggregation`, `pivot_field_set_properties`, `pivot_field_get_filters`, `pivot_field_apply_filter`, `pivot_field_clear_filters` |
| **Names** | 5 | `names_list`, `name_get`, `name_create`, `name_update`, `name_delete` |
| **Shapes & controls** | 6 | `shapes_list`, `shape_get`, `shape_set_on_action`, `controls_list`, `control_get`, `control_set_on_action` |
| **Workbook model** | 1 | `workbook_get_model` |
| **Test runner** | 3 | `test_run_suite`, `test_suite_run`, `test_get_run` |
| **Health** | 1 | `server_health` |

> [!NOTE]
> Tool names shown above are MCP registration names. Some clients, including Kilo, display them with a server prefix, such as `xlvba_excel_open`.
>
> The compile tools expose diagnostics; availability of a conclusive compilation result depends on the Excel environment.

### Read-only MCP resources

| URI template | Name | MIME type |
|---|---|---|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` |

For tool parameters, return values, and usage examples, see the [complete MCP tools reference](docs/05-mcp-tools-reference.md).

## Documentation

Detailed documentation is available in the `docs/` directory.

- [Overview](docs/01-overview.md) — project goals, concepts, and scope
- [Installation and setup](docs/02-installation-and-setup.md) — Windows/Excel preparation, installation, and first run
- [Configuration](docs/03-configuration.md) — environment variables, paths, and runtime settings
- [Architecture](docs/04-architecture.md) — MCP transport, Excel COM worker, and internal design
- [MCP tools reference](docs/05-mcp-tools-reference.md) — all 88 tools and 4 resources
- [Excel capabilities](docs/06-excel-capabilities.md) — supported Excel and VBA operations
- [Workflows and examples](docs/07-workflows-and-examples.md) — practical automation recipes
- [Security](docs/08-security.md) — access policies, trust boundaries, and risk management
- [Limitations and compatibility](docs/09-limitations-and-compatibility.md) — supported environments and technical constraints
- [Troubleshooting](docs/10-troubleshooting.md) — errors, diagnostics, and recovery
- [Development and maintenance](docs/11-development-and-maintenance.md) — testing, contributions, and maintenance

**Additional resources**

- [Russian documentation](docs/ru/)
- [Architecture diagrams](docs/assets/)
- [Technical specification](docs/01-Spec/excel_vba_mcp_technical_spec.docx)
- [Scope matrix](docs/01-Spec/excel_vba_mcp_scope_matrix.xlsx)
- [Documentation maintenance rules](docs/00-overview/doc-maintenance-rules.md)