"""Shared, STA-local cell-value classification helpers (task_026 / task_042).

These pure helpers classify a raw Excel cell value (as read from
``Range.Value`` / ``Range.Value2`` through pywin32) into the typed Stage-1
:class:`~excel_vba_mcp.domain.models.CellValue`. They live here, free of any COM
import, so both the Table adapter (task_026) and the Range adapter (task_042)
share one classification order and never leak a COM proxy or a lossy string
coercion.

Classification order (task_038 §1.4):

1. ``None`` -> :attr:`CellValueKind.EMPTY` (``value=None``; never ``""`` or ``0``).
2. ``bool`` -> :attr:`CellValueKind.BOOLEAN` (checked **before** ``int`` because
   ``bool`` is an ``int`` subclass).
3. ``datetime`` -> :attr:`CellValueKind.DATE` (ISO-8601 string; the domain
   :class:`CellValue` ``value`` union has no ``datetime`` member).
4. An OLE error code -> :attr:`CellValueKind.ERROR` with the Excel literal.
5. ``int`` / ``float`` -> :attr:`CellValueKind.NUMBER`.
6. ``str``: an Excel error literal -> :attr:`CellValueKind.ERROR`; otherwise
   :attr:`CellValueKind.STRING`.
7. Any other COM variant -> its string form as :attr:`CellValueKind.STRING`
   (never leaked).
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from excel_vba_mcp.domain.models import CellValue
from excel_vba_mcp.domain.statuses import CellValueKind

#: Excel error literals returned by ``Range.Value`` for error cells (spec §7.6:
#: cell errors are surfaced as a distinct typed value).
EXCEL_ERROR_CODES: frozenset[str] = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
    }
)

#: Excel error cells read through ``Range.Value`` / ``Range.Value2`` surface as
#: signed OLE error codes (``xlErr*``), not as the literal text. Map the code to
#: the literal so error cells are typed as :attr:`CellValueKind.ERROR` instead of
#: a bogus number.
EXCEL_ERROR_CODE_TO_LITERAL: dict[int, str] = {
    -2147221244: "#NULL!",
    -2146826281: "#DIV/0!",
    -2146826273: "#VALUE!",
    -2146826265: "#REF!",
    -2146826259: "#NAME?",
    -2146826252: "#NUM!",
    -2146826246: "#N/A",
    -2146826245: "#GETTING_DATA",
}


def classify_cell_value(value: Any, *, address: str) -> CellValue:
    """Classify a raw cell value into a typed :class:`CellValue`.

    Order matters: ``bool`` is checked before ``int`` (``bool`` is an ``int``
    subclass), and error codes/literals before the generic number / string case.
    """

    if value is None:
        return CellValue(kind=CellValueKind.EMPTY, value=None, address=address)
    if isinstance(value, bool):
        return CellValue(kind=CellValueKind.BOOLEAN, value=value, address=address)
    if isinstance(value, datetime):
        return CellValue(kind=CellValueKind.DATE, value=value.isoformat(), address=address)
    if isinstance(value, int) and value in EXCEL_ERROR_CODE_TO_LITERAL:
        return CellValue(
            kind=CellValueKind.ERROR, value=EXCEL_ERROR_CODE_TO_LITERAL[value], address=address
        )
    if isinstance(value, (int, float)):
        return CellValue(kind=CellValueKind.NUMBER, value=value, address=address)
    if isinstance(value, str):
        if value in EXCEL_ERROR_CODES:
            return CellValue(kind=CellValueKind.ERROR, value=value, address=address)
        return CellValue(kind=CellValueKind.STRING, value=value, address=address)
    # Unknown COM variant: fall back to its string form rather than leaking it.
    return CellValue(kind=CellValueKind.STRING, value=str(value), address=address)


def classify_cell_kind_name(value: Any) -> str:
    """Return a coarse type name for a sampled cell (``TableColumnInfo.data_type``)."""

    if value is None:
        return "empty"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, datetime):
        return "date"
    if isinstance(value, int) and value in EXCEL_ERROR_CODE_TO_LITERAL:
        return "error"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str) and value in EXCEL_ERROR_CODES:
        return "error"
    return "string"


#: Cell kinds whose ``value`` is free text subject to ``max_text_len`` truncation.
_TEXT_KINDS: frozenset[CellValueKind] = frozenset(
    {CellValueKind.STRING, CellValueKind.ERROR}
)


def _cell_text_bytes(cell: CellValue) -> int:
    """Estimated UTF-8 payload size of one cell's value (0 for empty)."""

    value = cell.value
    if value is None:
        return 0
    return len(str(value).encode("utf-8"))


def bound_value_grid(
    rows: list[list[CellValue]],
    *,
    max_text_len: int,
    max_response_bytes: int,
) -> tuple[list[list[CellValue]], bool, bool]:
    """Enforce the per-field text cap and a page byte budget on a value grid.

    Returns ``(rows, text_truncated, bytes_truncated)``:

    * ``text_truncated`` — a text field longer than ``max_text_len`` was cut to
      ``max_text_len`` characters (the caller marks the page ``truncated`` so the
      cut is never silent, contract §6).
    * ``bytes_truncated`` — trailing rows were dropped so the estimated UTF-8
      value payload stays within ``max_response_bytes`` (at least one row is
      always kept; a single oversized row is bounded only by ``max_text_len``).

    The byte estimate is the sum of the UTF-8 length of each non-empty cell value
    (the value payload, not the full JSON envelope) — a conservative proxy for the
    serialized response size (``max_response_bytes``). A non-positive limit
    disables that bound.
    """

    text_truncated = False
    if max_text_len > 0:
        bounded: list[list[CellValue]] = []
        for row in rows:
            new_row: list[CellValue] = []
            for cell in row:
                if (
                    cell.kind in _TEXT_KINDS
                    and isinstance(cell.value, str)
                    and len(cell.value) > max_text_len
                ):
                    new_row.append(
                        cell.model_copy(update={"value": cell.value[:max_text_len]})
                    )
                    text_truncated = True
                else:
                    new_row.append(cell)
            bounded.append(new_row)
        rows = bounded

    bytes_truncated = False
    if max_response_bytes > 0 and rows:
        row_bytes = [sum(_cell_text_bytes(cell) for cell in row) for row in rows]
        if sum(row_bytes) > max_response_bytes:
            bytes_truncated = True
            kept = 0
            running = 0
            for size in row_bytes:
                if kept >= 1 and running + size > max_response_bytes:
                    break
                running += size
                kept += 1
            rows = rows[:kept]

    return rows, text_truncated, bytes_truncated


__all__ = [
    "EXCEL_ERROR_CODES",
    "EXCEL_ERROR_CODE_TO_LITERAL",
    "bound_value_grid",
    "classify_cell_kind_name",
    "classify_cell_value",
]
