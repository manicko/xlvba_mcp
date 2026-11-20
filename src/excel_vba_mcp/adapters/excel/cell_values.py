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


__all__ = [
    "EXCEL_ERROR_CODES",
    "EXCEL_ERROR_CODE_TO_LITERAL",
    "classify_cell_kind_name",
    "classify_cell_value",
]
