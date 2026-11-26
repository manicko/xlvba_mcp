"""Unit tests for structured MCP error surfacing at the SDK boundary (task_b2, R-16).

These tests never touch Excel/COM. They exercise ``ExcelVbaMcpServer.call_tool``
directly to prove that a domain :class:`ExcelVbaMcpError` raised by a tool reaches
the client as a ``ToolError`` whose message is the JSON of ``to_dict()`` (stable
``errorCode`` + ``details``) with no generic "Error executing tool" masking, while a
non-domain exception still surfaces a guaranteed structured ``INTERNAL_ERROR``
envelope (U6-B) and an unknown tool raises ``ToolError``.
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


async def _raise_model_validator_error() -> None:
    """Trigger a ``model_validator`` ``ValueError`` (audit finding C-1).

    ``RangeWriteRequest`` raises a plain ``ValueError`` from its ``model_validator``
    when the grid exceeds ``MAX_RANGE_ROWS``; pydantic then embeds that live
    exception in ``errors()[i]['ctx']['error']``, which is not JSON-serializable.
    """

    from excel_vba_mcp.domain.models import CellValue, RangeWriteRequest

    RangeWriteRequest(
        rows=100_001,
        cols=1,
        values=[[CellValue(kind="number", value=1)]],  # type: ignore[list-item]
    )


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
    assert payload["errorCode"] == "PATH_NOT_ALLOWED"
    assert payload["details"] == {"path": "C:/secret/x.xlsm"}


async def test_non_domain_error_surfaced_as_structured_internal_error() -> None:
    # U6-B: a non-domain exception must NOT leak as a bare UnexpectedToolError /
    # "Error executing tool"; it surfaces a guaranteed INTERNAL_ERROR envelope.
    server = _make_server()
    server.add_tool(_raise_runtime_error, name="raise_runtime")

    with pytest.raises(ToolError) as exc_info:
        await server.call_tool("raise_runtime", {})

    assert not isinstance(exc_info.value, UnexpectedToolError)
    message = str(exc_info.value)
    assert "Error executing tool" not in message
    payload = json.loads(message)
    # All 6 required envelope keys (plus additive hresult/hresult_hex).
    assert set(payload) >= {
        "errorCode",
        "message",
        "details",
        "errorCategory",
        "isRetryable",
        "suggestedNextActions",
    }
    assert payload["errorCode"] == "INTERNAL_ERROR"
    assert payload["errorCategory"] == "internal"
    assert payload["isRetryable"] is False
    assert payload["details"]["reason"] == "unhandled_exception"
    assert payload["details"]["exception"] == "RuntimeError"
    assert payload["message"] == "boom"


async def test_unknown_tool_raises_tool_error() -> None:
    server = _make_server()

    with pytest.raises(ToolError):
        await server.call_tool("does_not_exist", {})


async def test_model_validator_error_surfaced_structured() -> None:
    """A ``model_validator`` ``ValueError`` still yields a structured code (C-1).

    Regression: ``ValidationError.errors()`` embeds the live ``ValueError`` in
    ``ctx['error']``; embedding it verbatim made ``json.dumps`` raise
    ``TypeError`` inside the handler, so the client received an unstructured
    error instead of ``INVALID_RANGE``.
    """

    server = _make_server()
    server.add_tool(_raise_model_validator_error, name="raise_model_validator")

    with pytest.raises(ToolError) as exc_info:
        await server.call_tool("raise_model_validator", {})

    # Must be the structured boundary translation, not the generic SDK wrapper.
    message = str(exc_info.value)
    assert "Error executing tool" not in message
    # The payload must be valid JSON (no TypeError leak) with a stable code.
    payload = json.loads(message)
    assert payload["errorCode"] == "INVALID_RANGE"
    assert payload["details"]["reason"] == "invalid_arguments"
    # The non-serializable exception must have been stringified, not dropped.
    ctx_errors = payload["details"]["validation"]
    assert ctx_errors, "validation details must be preserved"
    assert all(isinstance(entry.get("ctx", {}).get("error", ""), str) for entry in ctx_errors)


def test_jsonable_validation_errors_stringifies_live_exception() -> None:
    """``_jsonable_validation_errors`` makes ``errors()`` JSON-serializable (C-1)."""

    from pydantic import ValidationError

    from excel_vba_mcp.domain.models import CellValue, RangeWriteRequest
    from excel_vba_mcp.server import _jsonable_validation_errors

    try:
        RangeWriteRequest(
            rows=100_001,
            cols=1,
            values=[[CellValue(kind="number", value=1)]],  # type: ignore[list-item]
        )
    except ValidationError as exc:
        sanitized = _jsonable_validation_errors(exc)
    else:  # pragma: no cover - the guard must fire
        pytest.fail("expected a ValidationError for an oversized grid")

    # Must not raise, and must be JSON-serializable.
    json.dumps(sanitized)
    error_entries = [entry for entry in sanitized if entry.get("ctx", {}).get("error")]
    assert error_entries, "the ValueError ctx entry must be preserved as a string"
    assert isinstance(error_entries[0]["ctx"]["error"], str)