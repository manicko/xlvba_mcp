"""MCP server bootstrap and lifecycle (STDIO).

This module defines the MCP server instance. It must not write non-protocol data to
stdout; all diagnostics go to stderr via logging.
"""

from __future__ import annotations

import logging
import sys
from typing import Any

from mcp.server.mcpserver import MCPServer

from excel_vba_mcp.config import load_config

_config = load_config()

logging.basicConfig(
    level=_config.server.log_level,
    stream=sys.stderr,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

mcp = MCPServer("excel-vba-mcp")


@mcp.tool()
def server_health() -> dict[str, Any]:
    """Return a minimal health payload for discovery smoke tests.

    Returns:
        Dictionary with server status.
    """

    return {"status": "ok", "server": "excel-vba-mcp"}
