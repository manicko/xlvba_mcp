"""Unit tests for structured MCP error surfacing at the SDK boundary (task_b2, R-16).

These tests never touch Excel/COM. They exercise ``ExcelVbaMcpServer.call_tool``
directly to prove that a domain :class:`ExcelVbaMcpError` raised by a tool reaches
the client as a ``ToolError`` whose message is the JSON of ``to_dict()`` (stable
``code`` + ``context``) with no generic "Error executing tool" masking, while a
non-domain exception still surfaces as ``UnexpectedToolError`` and an unknown tool
raises ``ToolError``.
"""

from __future__ import annotations

import json

import pytest
from mcp.server.mcpserver.exceptions import ToolError, UnexpectedToolError

from excel_vba_mcp.domain.errors import PathNotAllowedError
from excel_vba_mcp.server import ExcelVbaMcpServer

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


async def _raise_path_not_allowed() -> None:
    raise PathNotAllowedError("path rejected", context={"path": "C:/secret/x.xlsm"})


async def _raise_runtime_error() -> None:
    raise RuntimeError("boom")


def _make_server() -> ExcelVbaMcpServer:
    return ExcelVbaMcpServer("test-excel-vba-mcp", version="0.0.0")


async def test_domain_error_surfaced_as_structured_tool_error() -> None:
    server = _make_server()
    server.add_tool(_raise_path_not_allowed, name="raise_path")

    with pytest.raises(ToolError) as exc_info:
        await server.call_tool("raise_path", {})

    # A domain error is not the generic SDK crash wrapper, and not an
    # UnexpectedToolError: it is a plain ToolError carrying the JSON payload.
    assert not isinstance(exc_info.value, UnexpectedToolError)
    message = str(exc_info.value)
    assert "Error executing tool" not in message
    payload = json.loads(message)
    assert payload["code"] == "PATH_NOT_ALLOWED"
    assert payload["context"] == {"path": "C:/secret/x.xlsm"}


async def test_non_domain_error_surfaced_as_unexpected_tool_error() -> None:
    server = _make_server()
    server.add_tool(_raise_runtime_error, name="raise_runtime")

    with pytest.raises(UnexpectedToolError):
        await server.call_tool("raise_runtime", {})


async def test_unknown_tool_raises_tool_error() -> None:
    server = _make_server()

    with pytest.raises(ToolError):
        await server.call_tool("does_not_exist", {})