"""MCP tool handlers (task_012).

Thin, async tool functions that validate arguments and delegate to the
application layer. This package **never touches COM** (ARCH-009): no
``pythoncom`` / ``win32com`` import lives under ``api/``.
"""

from excel_vba_mcp.api.workbook_tools import (
    excel_close,
    excel_get_status,
    excel_list_workbooks,
    excel_open,
    excel_save,
)

__all__ = [
    "excel_close",
    "excel_get_status",
    "excel_list_workbooks",
    "excel_open",
    "excel_save",
]
