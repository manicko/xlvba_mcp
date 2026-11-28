"""Large-data file export service (WP-F).

:class:`DataExportService` streams a large dataset — a worksheet Range, an Excel
Table (ListObject) data body or a PivotTable's rendered result — to a **file
artifact** (CSV or JSONL) with bounded memory, and returns only compact metadata
(:class:`~excel_vba_mcp.domain.models.DataExportResult`): the artifact path, row
count, byte size, SHA-256 and completeness. The data itself is **never** returned
in the MCP response.

The service is deliberately COM-free: it delegates every per-page read to the
injected :class:`~excel_vba_mcp.application.ranges_service.RangesService`,
:class:`~excel_vba_mcp.application.tables_service.TablesService` and
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`, and
writes each page straight to an :class:`~excel_vba_mcp.storage.artifact_store.ArtifactStream`
so only one page is ever held in memory. The async signature matches the other
application services; the per-page reads are ``await``-ed.
"""

from __future__ import annotations

import asyncio
import csv
import io
import json
from collections.abc import Sequence
from typing import Any, Literal, cast
from uuid import uuid4

from excel_vba_mcp.adapters.excel.tables_adapter import TableRows
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    CellValue,
    DataExportResult,
    PageInfo,
    PivotResultPage,
    RangeValues,
)
from excel_vba_mcp.domain.statuses import CellValueKind
from excel_vba_mcp.storage.artifact_store import ArtifactStore, ArtifactType

#: Number of rows requested per page read. The bounded readers cap cells per page
#: independently, so a wide sheet simply returns fewer rows with ``has_more`` set;
#: the loop then advances by the returned count. One page is held in memory at a
#: time, so this stays bounded regardless of the total row count.
_PAGE_ROWS = 1000

#: Accepted source kinds for an export.
_SOURCE_KINDS: frozenset[str] = frozenset({"range", "table", "pivot-result"})

#: Accepted output formats.
_FORMATS: frozenset[str] = frozenset({"csv", "jsonl"})


class DataExportService:
    """Stream a Range / Table / PivotTable result to a CSV / JSONL artifact.

    Args:
        ranges: The Range read service (``read_values``).
        tables: The Table read service (``read_table_rows``).
        pivots: The PivotTable read service (``read_pivot_result``).
        artifacts: The filesystem artifact store (streaming writer + index).
        limits: The unified large-data limits (``max_export_rows`` ceiling).
        session: The workbook session service; its content ``version`` is captured
            at export start and threaded as ``cursor_version`` into every page read
            so a concurrent mutation fails the export with ``STALE_CURSOR``.
        snapshot: The snapshot service; an immutable snapshot (SaveCopyAs) is
            created at export start and its ``sha256`` is recorded on the result so
            the export is bound to a stable input (contract §5).
    """

    def __init__(
        self,
        ranges: RangesService,
        tables: TablesService,
        pivots: PivotTablesService,
        artifacts: ArtifactStore,
        limits: LargeDataConfig,
        *,
        session: Any,
        snapshot: Any,
    ) -> None:
        self._ranges = ranges
        self._tables = tables
        self._pivots = pivots
        self._artifacts = artifacts
        self._limits = limits
        self._session = session
        self._snapshot = snapshot
        # Bound concurrent large-file exports (contract §5 / ``max_file_jobs``).
        self._file_jobs = asyncio.Semaphore(max(1, limits.max_file_jobs))

    async def export(
        self,
        workbook_id: str,
        *,
        source_kind: Literal["range", "table", "pivot-result"],
        sheet: str | None = None,
        address: str | None = None,
        table: str | None = None,
        pivot_table: str | None = None,
        fmt: Literal["csv", "jsonl"] = "csv",
        limit_rows: int | None = None,
    ) -> DataExportResult:
        """Export a large dataset to a file artifact and return compact metadata.

        The chosen source is paged by ``offset`` with a bounded row chunk; each page
        is converted to CSV rows (a header from the first page's column count, or the
        Table headers when available) or JSONL objects (``col0..colN`` keys, or the
        Table headers) and written straight to the artifact stream. Only one page is
        held in memory at a time. Paging stops when ``page.has_more`` is ``False`` or
        ``limit_rows`` / ``max_export_rows`` is reached (``truncated`` / ``complete``
        reflect which).

        Args:
            workbook_id: The id returned by ``excel_open``.
            source_kind: ``"range"``, ``"table"`` or ``"pivot-result"``.
            sheet: Worksheet name (``range``; required).
            address: A1 address (``range``; required).
            table: Table name (``table``; required).
            pivot_table: PivotTable name (``pivot-result``; required).
            fmt: ``"csv"`` (default) or ``"jsonl"``.
            limit_rows: Optional cap on the number of data rows exported.

        Returns:
            A :class:`DataExportResult` (path, rows, bytes, sha256, completeness).

        Raises:
            ExcelVbaMcpError: ``INVALID_RANGE`` for a missing/invalid argument;
                ``STALE_CURSOR`` when the workbook content version advances between
                pages (the partial artifact is aborted, never reported complete).
        """

        normalized_kind = _require_source_kind(source_kind)
        normalized_fmt = _require_format(fmt)
        _validate_source_args(normalized_kind, sheet, address, table, pivot_table)
        _require_limit_rows(limit_rows)

        # Bound concurrent large-file exports (contract §5 / ``max_file_jobs``).
        async with self._file_jobs:
            # Bind the export to an immutable snapshot + a fixed content version
            # BEFORE the first page: the snapshot sha256 is recorded on the result
            # and the version is threaded as ``cursor_version`` into every page read
            # so a concurrent mutation fails the export with ``STALE_CURSOR`` rather
            # than silently mixing versions (contract §3 / §5).
            source_version = self._session.get_version(workbook_id)
            manifest = await self._snapshot.create_snapshot(workbook_id)
            input_sha256 = manifest.sha256

            artifact_type = (
                ArtifactType.CSV if normalized_fmt == "csv" else ArtifactType.JSONL
            )
            stream = self._artifacts.open_stream(
                artifact_type,
                suggested_name=f"data-export-{normalized_kind}-{uuid4().hex}",
                workbook_id=workbook_id,
            )

            max_rows = self._limits.max_export_rows
            rows_written = 0
            headers: list[str] | None = None
            truncated = False
            complete = False
            offset = 0

            try:
                while True:
                    page_rows, cols, page_info, page_headers = await self._read_page(
                        workbook_id,
                        normalized_kind,
                        sheet,
                        address,
                        table,
                        pivot_table,
                        offset,
                        cursor_version=source_version,
                    )
                    if headers is None:
                        headers = page_headers or [
                            f"col{index}" for index in range(cols)
                        ]
                        if normalized_fmt == "csv":
                            stream.write_text(_csv_line(headers))
                    for row in page_rows:
                        if limit_rows is not None and rows_written >= limit_rows:
                            truncated = True
                            break
                        if rows_written >= max_rows:
                            truncated = True
                            break
                        scalar_row = [_cell_to_scalar(cell) for cell in row]
                        stream.write_text(
                            _format_row(normalized_fmt, headers, scalar_row)
                        )
                        rows_written += 1
                    if truncated:
                        break
                    if not page_info.has_more:
                        complete = True
                        break
                    if page_info.returned <= 0:
                        # No progress possible (defensive): stop rather than spin.
                        break
                    offset += page_info.returned
            except BaseException:
                stream.abort()
                raise

            info = stream.finalize(
                kind="data_export",
                source={
                    "source_kind": normalized_kind,
                    "format": normalized_fmt,
                    "rows": rows_written,
                    "complete": complete,
                    "truncated": truncated,
                },
            )
            return DataExportResult(
                artifact_id=info.artifact_id,
                path=info.stored_path,
                format=normalized_fmt,
                rows=rows_written,
                bytes=info.size,
                sha256=info.sha256,
                complete=complete,
                source=normalized_kind,
                source_version=source_version,
                truncated=truncated,
                input_sha256=input_sha256,
            )

    # ------------------------------------------------------------------
    # Per-page read (delegated to the injected services; no COM here)
    # ------------------------------------------------------------------
    async def _read_page(
        self,
        workbook_id: str,
        source_kind: str,
        sheet: str | None,
        address: str | None,
        table: str | None,
        pivot_table: str | None,
        offset: int,
        *,
        cursor_version: int,
    ) -> tuple[list[list[CellValue]], int, PageInfo, list[str] | None]:
        """Read one bounded page and normalize it to rows / cols / PageInfo / headers.

        ``headers`` is the Table column names for a ``table`` source (so the export
        carries real headers) and ``None`` for ``range`` / ``pivot-result`` (the
        caller synthesizes ``col0..colN``). ``cursor_version`` is the export's bound
        content version; each underlying read raises ``STALE_CURSOR`` if the workbook
        version advanced since the export started.
        """

        if source_kind == "range":
            result: RangeValues = await self._ranges.read_values(
                workbook_id,
                sheet or "",
                address or "",
                offset=offset,
                limit=_PAGE_ROWS,
                cursor_version=cursor_version,
            )
            return result.values, result.cols, _coalesce_page(result.page, result.values), None

        if source_kind == "table":
            table_result: TableRows = await self._tables.read_table_rows(
                workbook_id,
                table or "",
                offset=offset,
                limit=_PAGE_ROWS,
                cursor_version=cursor_version,
            )
            return (
                table_result.rows,
                len(table_result.columns),
                _coalesce_page(table_result.page, table_result.rows),
                list(table_result.columns),
            )

        pivot_result: PivotResultPage = await self._pivots.read_pivot_result(
            workbook_id,
            pivot_table or "",
            offset=offset,
            limit=_PAGE_ROWS,
            cursor_version=cursor_version,
        )
        return pivot_result.values, pivot_result.cols, pivot_result.page, None


# ---------------------------------------------------------------------------
# Pure helpers (no COM)
# ---------------------------------------------------------------------------
def _require_source_kind(value: object) -> str:
    """Validate the export source kind."""

    if isinstance(value, str) and value.strip().lower() in _SOURCE_KINDS:
        return value.strip().lower()
    raise ExcelVbaMcpError(
        f"Argument 'source_kind' must be one of {sorted(_SOURCE_KINDS)}, got {value!r}.",
        code=ErrorCode.INVALID_RANGE,
        context={
            "source_kind": repr(value),
            "accepted": sorted(_SOURCE_KINDS),
            "reason": "invalid_source_kind",
        },
    )


def _require_format(value: object) -> Literal["csv", "jsonl"]:
    """Validate the export output format."""

    if isinstance(value, str) and value.strip().lower() in _FORMATS:
        return cast('Literal["csv", "jsonl"]', value.strip().lower())
    raise ExcelVbaMcpError(
        f"Argument 'fmt' must be one of {sorted(_FORMATS)}, got {value!r}.",
        code=ErrorCode.INVALID_RANGE,
        context={
            "fmt": repr(value),
            "accepted": sorted(_FORMATS),
            "reason": "invalid_format",
        },
    )


def _validate_source_args(
    source_kind: str,
    sheet: str | None,
    address: str | None,
    table: str | None,
    pivot_table: str | None,
) -> None:
    """Enforce the per-source-kind required arguments before any read."""

    if source_kind == "range":
        _require_text(sheet, "sheet")
        _require_text(address, "address")
    elif source_kind == "table":
        _require_text(table, "table")
    else:  # pivot-result
        _require_text(pivot_table, "pivot_table")


def _require_text(value: object, field: str) -> str:
    """Reject a blank / non-string argument before any read."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} is required for this source and must be a non-empty string.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "reason": "invalid_argument"},
        )
    return value.strip()


def _require_limit_rows(value: object) -> None:
    """Validate an optional ``limit_rows`` cap."""

    if value is None:
        return
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ExcelVbaMcpError(
            "Argument 'limit_rows' must be a non-negative integer when supplied.",
            code=ErrorCode.INVALID_RANGE,
            context={"limit_rows": repr(value), "reason": "invalid_argument"},
        )


def _coalesce_page(page: PageInfo | None, rows: list[list[CellValue]]) -> PageInfo:
    """Return the page's :class:`PageInfo`, synthesizing one when absent."""

    if page is not None:
        return page
    returned = len(rows)
    return PageInfo(offset=0, returned=returned, has_more=False, complete=True)


def _cell_to_scalar(cell: CellValue) -> str | int | float | bool:
    """Convert a typed :class:`CellValue` to a plain JSON/CSV scalar.

    ``empty`` becomes ``""``; ``boolean`` / ``number`` keep their native type;
    ``string`` / ``date`` / ``error`` become text (never ``None``).
    """

    if cell.kind == CellValueKind.EMPTY:
        return ""
    if cell.kind == CellValueKind.BOOLEAN:
        return bool(cell.value)
    if cell.kind == CellValueKind.NUMBER:
        return cell.value if isinstance(cell.value, (int, float)) else 0
    return "" if cell.value is None else str(cell.value)


def _csv_line(row: Sequence[object]) -> str:
    """Render one row as a CSV line (RFC 4180 quoting, ``\\r\\n`` terminator)."""

    buffer = io.StringIO()
    csv.writer(buffer).writerow(row)
    return buffer.getvalue()


def _format_row(fmt: str, headers: Sequence[str], row: Sequence[object]) -> str:
    """Render one data row as a CSV line or a JSONL object line."""

    if fmt == "jsonl":
        record = {
            header: row[index] if index < len(row) else ""
            for index, header in enumerate(headers)
        }
        return json.dumps(record, ensure_ascii=False) + "\n"
    return _csv_line(row)


__all__ = ["DataExportService"]