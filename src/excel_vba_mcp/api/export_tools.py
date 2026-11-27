"""Large-data file export MCP tool contract (WP-F).

Exposes ONE tool, ``data_export``, that streams a large dataset — a worksheet
Range, an Excel Table (ListObject) data body or a PivotTable's rendered result —
to a **file artifact** (CSV or JSONL) with bounded memory and returns only compact
metadata (:class:`~excel_vba_mcp.domain.models.DataExportResult`): the artifact
path, row count, byte size, SHA-256 and completeness. The data itself is never
returned in the MCP response.

This module imports **no Excel/COM** — only the service and domain DTOs. The
active service is published by the server lifespan via :func:`set_data_export_service`
(module global + setter, mirroring ``pivot_tools`` / ``ooxml_tools``).

The tool is read-only with respect to the workbook (it only reads pages and writes
to the artifact store) but is **not** read-only in the MCP annotation sense because
it creates a file; it is non-destructive and idempotent (each call writes a new
artifact).
"""

from __future__ import annotations

from typing import Literal

from excel_vba_mcp.application.data_export_service import DataExportService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import DataExportResult

# The active data export service, published by the server lifespan. ``None``
# before startup and after shutdown; the tool raises a clear error if used out of
# band (e.g. a unit test that has not installed a fake).
_data_export_service: DataExportService | None = None


def set_data_export_service(service: DataExportService | None) -> None:
    """Publish (or clear) the active data export service used by the tool."""

    global _data_export_service
    _data_export_service = service


def get_data_export_service() -> DataExportService:
    """Return the active data export service or raise if it is not configured."""

    if _data_export_service is None:
        raise ExcelVbaMcpError(
            "The data export service is not initialized.",
            context={"tool": "data_export"},
        )
    return _data_export_service


def _require_non_empty(value: object, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "reason": "invalid_argument"},
        )
    return value


def _require_int(value: object, field: str, *, minimum: int | None = None) -> int:
    """Validate that an argument is an integer (optionally with a minimum)."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be an integer.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "reason": "invalid_argument"},
        )
    if minimum is not None and value < minimum:
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be >= {minimum}.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "value": value, "reason": "invalid_argument"},
        )
    return value


async def data_export(
    workbook_id: str,
    source_kind: Literal["range", "table", "pivot-result"],
    sheet: str | None = None,
    address: str | None = None,
    table: str | None = None,
    pivot_table: str | None = None,
    fmt: Literal["csv", "jsonl"] = "csv",
    limit_rows: int | None = None,
) -> DataExportResult:
    """Stream a large dataset to a CSV / JSONL file artifact (bounded memory).

    Returns only compact metadata (path, rows, bytes, sha256, completeness); the
    data is written to the artifact store, never returned inline. The source is
    paged by ``offset`` and each page is written straight to the artifact stream,
    so memory stays bounded regardless of the total row count.

    Args:
        workbook_id: The id returned by ``excel_open``.
        source_kind: ``"range"`` (needs ``sheet`` + ``address``), ``"table"`` (needs
            ``table``) or ``"pivot-result"`` (needs ``pivot_table``).
        sheet: Worksheet name (``range``).
        address: A1 address (``range``).
        table: Table name (``table``).
        pivot_table: PivotTable name (``pivot-result``).
        fmt: ``"csv"`` (default) or ``"jsonl"``.
        limit_rows: Optional cap on the number of data rows exported; when the cap
            cuts the source short the result reports ``truncated=True`` /
            ``complete=False``.

    Returns:
        A :class:`DataExportResult`.

    Raises:
        ExcelVbaMcpError: ``INVALID_RANGE`` for a missing/invalid argument.
    """

    _require_non_empty(workbook_id, "workbook_id")
    if limit_rows is not None:
        _require_int(limit_rows, "limit_rows", minimum=0)
    service = get_data_export_service()
    return await service.export(
        workbook_id,
        source_kind=source_kind,
        sheet=sheet,
        address=address,
        table=table,
        pivot_table=pivot_table,
        fmt=fmt,
        limit_rows=limit_rows,
    )


__all__ = [
    "data_export",
    "get_data_export_service",
    "set_data_export_service",
]