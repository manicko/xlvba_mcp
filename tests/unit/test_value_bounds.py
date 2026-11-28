"""Unit tests for the large-data value bounds (Fix 3).

Hermetic: exercises the pure :func:`bound_value_grid` helper that enforces the
per-field ``max_text_len`` cap and the ``max_response_bytes`` page byte budget at
the range / OOXML value-serialization boundary. No Excel / no COM.
"""

from __future__ import annotations

from excel_vba_mcp.adapters.excel.cell_values import bound_value_grid
from excel_vba_mcp.domain.models import CellValue
from excel_vba_mcp.domain.statuses import CellValueKind


def _text(value: str) -> CellValue:
    return CellValue(kind=CellValueKind.STRING, value=value)


def _num(value: float) -> CellValue:
    return CellValue(kind=CellValueKind.NUMBER, value=value)


# ---------------------------------------------------------------------------
# max_text_len
# ---------------------------------------------------------------------------
def test_text_field_truncated_to_cap() -> None:
    rows = [[_text("a" * 100)]]
    out, text_truncated, bytes_truncated = bound_value_grid(
        rows, max_text_len=10, max_response_bytes=0
    )

    assert text_truncated is True
    assert bytes_truncated is False
    assert len(out[0][0].value) == 10  # type: ignore[arg-type]
    assert out[0][0].value == "a" * 10


def test_text_field_within_cap_unchanged() -> None:
    rows = [[_text("short")]]
    out, text_truncated, _ = bound_value_grid(
        rows, max_text_len=100, max_response_bytes=0
    )

    assert text_truncated is False
    assert out[0][0].value == "short"


def test_non_text_fields_not_truncated() -> None:
    # A long numeric/error-free value is not a text field; numbers are untouched.
    rows = [[_num(123456789.0)]]
    out, text_truncated, _ = bound_value_grid(
        rows, max_text_len=1, max_response_bytes=0
    )

    assert text_truncated is False
    assert out[0][0].value == 123456789.0


# ---------------------------------------------------------------------------
# max_response_bytes
# ---------------------------------------------------------------------------
def test_bytes_budget_drops_trailing_rows() -> None:
    # Each row is ~10 bytes of value payload; a 25-byte budget keeps 2 rows.
    rows = [[_text("0123456789")] for _ in range(5)]
    out, text_truncated, bytes_truncated = bound_value_grid(
        rows, max_text_len=0, max_response_bytes=25
    )

    assert bytes_truncated is True
    assert text_truncated is False
    assert len(out) == 2  # 2 * 10 = 20 <= 25; a 3rd row would exceed.


def test_bytes_budget_keeps_at_least_one_row() -> None:
    # A single row larger than the budget is still kept (cannot split a row).
    rows = [[_text("x" * 100)], [_text("y" * 100)]]
    out, _, bytes_truncated = bound_value_grid(
        rows, max_text_len=0, max_response_bytes=10
    )

    assert bytes_truncated is True
    assert len(out) == 1


def test_bytes_budget_not_exceeded_keeps_all() -> None:
    rows = [[_text("ab")], [_text("cd")]]
    out, _, bytes_truncated = bound_value_grid(
        rows, max_text_len=0, max_response_bytes=1000
    )

    assert bytes_truncated is False
    assert len(out) == 2


def test_zero_byte_budget_disables_byte_bound() -> None:
    rows = [[_text("x" * 50)] for _ in range(3)]
    out, _, bytes_truncated = bound_value_grid(
        rows, max_text_len=0, max_response_bytes=0
    )

    assert bytes_truncated is False
    assert len(out) == 3