"""Application services (task_012).

The application layer orchestrates use cases over the domain and adapters. It
owns no COM proxies: all Excel access is delegated to the single-STA worker via
``await worker.submit(...)``.
"""

from excel_vba_mcp.application.session_service import SessionService

__all__ = ["SessionService"]
