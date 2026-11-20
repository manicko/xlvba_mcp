"""MCP tool handlers (task_012, task_016, task_020).

Thin, async tool functions that validate arguments and delegate to the
application layer. This package **never touches COM** (ARCH-009): no
``pythoncom`` / ``win32com`` import lives under ``api/``.
"""

from excel_vba_mcp.api.compile_tools import (
    set_compile_service,
    vba_compile,
    vba_get_compile_diagnostics,
)
from excel_vba_mcp.api.macro_tools import (
    macro_get_last_result,
    macro_list,
    macro_run,
    set_macro_service,
)
from excel_vba_mcp.api.vba_tools import (
    set_vba_project_service,
    vba_export_project,
    vba_get_references,
    vba_list_components,
    vba_list_procedures,
    vba_project_info,
    vba_read_component,
    vba_search_code,
)
from excel_vba_mcp.api.workbook_tools import (
    excel_close,
    excel_get_status,
    excel_list_workbooks,
    excel_open,
    excel_save,
    get_workbook_model_service,
    set_workbook_model_service,
    workbook_get_model,
)

__all__ = [
    "excel_close",
    "excel_get_status",
    "excel_list_workbooks",
    "excel_open",
    "excel_save",
    "get_workbook_model_service",
    "macro_get_last_result",
    "macro_list",
    "macro_run",
    "set_compile_service",
    "set_macro_service",
    "set_vba_project_service",
    "set_workbook_model_service",
    "vba_compile",
    "vba_export_project",
    "vba_get_compile_diagnostics",
    "vba_get_references",
    "vba_list_components",
    "vba_list_procedures",
    "vba_project_info",
    "vba_read_component",
    "vba_search_code",
    "workbook_get_model",
]
