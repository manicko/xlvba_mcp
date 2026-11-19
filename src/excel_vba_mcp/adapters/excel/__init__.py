"""Excel COM adapters (task_010).

Exposes the hidden, non-activating application wrapper and the single-STA COM
worker. Submodules import ``pywin32`` lazily, so importing this package never
requires Windows / COM at import time.
"""

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import (
    ExcelComWorker,
    Section,
    WorkerStatus,
    WorkerUnavailableError,
)

__all__ = [
    "ExcelAppAdapter",
    "ExcelComWorker",
    "Section",
    "WorkerStatus",
    "WorkerUnavailableError",
]
