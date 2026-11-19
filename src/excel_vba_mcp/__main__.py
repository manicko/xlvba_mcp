"""Process entry point for the Excel VBA MCP server."""

from __future__ import annotations

from excel_vba_mcp.server import logger, mcp


def main() -> None:
    """Run the MCP server over STDIO."""

    logger.info("Starting Excel VBA MCP server (STDIO)")
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
