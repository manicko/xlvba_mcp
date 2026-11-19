"""Application services (task_012, task_020).

The application layer orchestrates use cases over the domain and adapters. It
owns no COM proxies: all Excel access is delegated to the single-STA worker via
``await worker.submit(...)``.
"""

from excel_vba_mcp.application.compile_service import (
    CompileDiagnostics,
    CompileResult,
    CompileService,
)
from excel_vba_mcp.application.macro_service import (
    BackgroundSafetyRegistry,
    MacroInfo,
    MacroService,
)
from excel_vba_mcp.application.session_service import SessionService

__all__ = [
    "BackgroundSafetyRegistry",
    "CompileDiagnostics",
    "CompileResult",
    "CompileService",
    "MacroInfo",
    "MacroService",
    "SessionService",
]
