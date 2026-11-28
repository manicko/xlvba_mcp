"""Non-COM OOXML snapshot-read MCP tool contract (WP-E1).

Exposes ONE read-only tool, ``workbook_file_read``, that reads worksheet rows or
workbook metadata from an **immutable snapshot file** (a ``Workbook.SaveCopyAs``
image) without occupying the single-STA COM worker and without materializing the
whole sheet. Parsing runs in a worker thread inside
:class:`~excel_vba_mcp.application.ooxml_service.OoxmlService`.

This module imports **no Excel/COM** — only the service and domain DTOs. The
active service is published by the server lifespan via :func:`set_ooxml_service`
(module global + setter, mirroring ``pivot_tools``).
"""

from __future__ import annotations

from typing import Literal

from excel_vba_mcp.application.ooxml_service import OoxmlService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    OoxmlMetadata,
    OoxmlPivotCachePage,
    OoxmlSheetPage,
)

# The active OOXML read service, published by the server lifespan. ``None``
# before startup and after shutdown; the tool raises a clear error if used out of
# band (e.g. a unit test that has not installed a fake).
_ooxml_service: OoxmlService | None = None


def set_ooxml_service(service: OoxmlService | None) -> None:
    """Publish (or clear) the active OOXML read service used by the tool."""

    global _ooxml_service
    _ooxml_service = service


def get_ooxml_service() -> OoxmlService:
    """Return the active OOXML read service or raise if it is not configured."""

    if _ooxml_service is None:
        raise ExcelVbaMcpError(
            "The OOXML read service is not initialized.",
            context={"tool": "workbook_snapshot_read"},
        )
    return _ooxml_service


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


async def workbook_file_read(
    action: Literal["metadata", "sheet-rows", "pivot-cache-records"],
    workbook_id: str,
    sheet: str | None = None,
    offset: int = 0,
    limit: int | None = None,
    snapshot_id: str | None = None,
    cache_index: int | None = None,
) -> OoxmlMetadata | OoxmlSheetPage | OoxmlPivotCachePage:
    """Read workbook metadata, worksheet rows, or saved PivotCache records from a snapshot.

    Non-COM, streaming, bounded and read-only: the data is parsed from an
    immutable snapshot file (never the live working copy or the original) in a
    worker thread, so the single-STA COM worker is never occupied. When
    ``snapshot_id`` is omitted a fresh snapshot of ``workbook_id`` is created and
    used as the immutable input.

    Args:
        action: ``"metadata"`` (workbook structure), ``"sheet-rows"`` (a page of
            typed rows) or ``"pivot-cache-records"`` (a page of **saved** PivotCache
            records — not the displayed result and not the live source).
        workbook_id: The id returned by ``excel_open`` (required).
        sheet: The worksheet name (``sheet-rows``; required).
        offset: 0-based row/record index of the first item to return.
        limit: Optional cap on the number of rows/records returned.
        snapshot_id: Optional explicit snapshot to read; when omitted a fresh
            snapshot is created from the open workbook.
        cache_index: The 1-based PivotCache index (``pivot-cache-records``; required).

    Returns:
        An :class:`OoxmlMetadata` (``metadata``), an :class:`OoxmlSheetPage`
        (``sheet-rows``) or an :class:`OoxmlPivotCachePage` (``pivot-cache-records``).
        A cache with no saved local records yields an empty page with
        ``page.reason == "unsupported_source"``.

    Raises:
        ExcelVbaMcpError: ``INVALID_RANGE`` for a missing/invalid argument or an
            unknown action; ``NOT_FOUND`` for a missing cache index;
            ``WORKBOOK_NOT_FOUND`` when the snapshot file is absent.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_ooxml_service()

    if action == "metadata":
        return await service.read_metadata(workbook_id, snapshot_id=snapshot_id)

    if action == "sheet-rows":
        sheet = _require_non_empty(sheet, "sheet")
        _require_int(offset, "offset", minimum=0)
        if limit is not None:
            _require_int(limit, "limit", minimum=1)
        return await service.read_sheet_rows(
            workbook_id,
            sheet,
            offset=offset,
            limit=limit,
            snapshot_id=snapshot_id,
        )

    if action == "pivot-cache-records":
        if cache_index is None:
            raise ExcelVbaMcpError(
                "Argument 'cache_index' is required for 'pivot-cache-records'.",
                code=ErrorCode.INVALID_RANGE,
                context={"field": "cache_index", "reason": "invalid_argument"},
            )
        _require_int(cache_index, "cache_index", minimum=1)
        _require_int(offset, "offset", minimum=0)
        if limit is not None:
            _require_int(limit, "limit", minimum=1)
        return await service.read_pivot_cache_records(
            workbook_id,
            cache_index,
            offset=offset,
            limit=limit,
            snapshot_id=snapshot_id,
        )

    raise ExcelVbaMcpError(
        f"Unknown workbook_snapshot_read action {action!r}; expected 'metadata', "
        "'sheet-rows' or 'pivot-cache-records'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "workbook_snapshot_read", "action": action},
    )


__all__ = [
    "get_ooxml_service",
    "set_ooxml_service",
    "workbook_file_read",
]