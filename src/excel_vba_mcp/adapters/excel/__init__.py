"""Excel COM adapters (task_010, task_016).

Exposes the hidden, non-activating application wrapper, the single-STA COM
worker and the STA-local VBIDE adapter. Submodules import ``pywin32`` lazily, so
importing this package never requires Windows / COM at import time.
"""

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import (
    ExcelComWorker,
    Section,
    WorkerStatus,
    WorkerUnavailableError,
)
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter

__all__ = [
    "ExcelAppAdapter",
    "ExcelComWorker",
    "Section",
    "VBIDEAdapter",
    "WorkerStatus",
    "WorkerUnavailableError",
]
