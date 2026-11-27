"""Unit tests for the WP-A ``STALE_CURSOR`` error code.

Static, read-only: never touches Excel/COM. Proves that ``StaleCursorError``
carries ``ErrorCode.STALE_CURSOR``, classifies as a non-retryable ``validation``
error, and is registered in the code -> exception map.
"""

from __future__ import annotations

from excel_vba_mcp.domain import errors as err
from excel_vba_mcp.domain.errors import (
    ErrorCode,
    StaleCursorError,
    category_for,
    retryable_for,
)


def test_stale_cursor_error_carries_code() -> None:
    error = StaleCursorError("x")

    assert error.code is ErrorCode.STALE_CURSOR
    assert error.to_dict()["errorCode"] == "STALE_CURSOR"


def test_stale_cursor_is_validation_and_not_retryable() -> None:
    assert category_for(ErrorCode.STALE_CURSOR) == "validation"
    assert retryable_for(ErrorCode.STALE_CURSOR) is False


def test_stale_cursor_is_registered() -> None:
    assert err._ERROR_BY_CODE[ErrorCode.STALE_CURSOR] is StaleCursorError