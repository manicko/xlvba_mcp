"""Non-COM, streaming OOXML read service (WP-E1).

:class:`OoxmlService` reads worksheet rows and workbook metadata from an
**immutable snapshot file** (a ``Workbook.SaveCopyAs`` image captured by
:class:`~excel_vba_mcp.application.snapshot_service.SnapshotService`) using the
streaming, bounded adapters in :mod:`excel_vba_mcp.adapters.ooxml`. It owns **no
COM proxies** and never touches the live working copy or the original: the only
input is the snapshot's ``stored_path``.

Event-loop safety: the actual parsing (pure file I/O + ``xml.etree``) runs in a
worker thread via :func:`asyncio.to_thread`, never inline on the loop, so a large
sheet read cannot block the single-STA worker or the async runtime.

Error mapping: an :class:`~excel_vba_mcp.adapters.ooxml.zip_safety.OoxmlSafetyError`
becomes a structured ``INVALID_RANGE`` (``reason="unsafe_ooxml"``); a missing
snapshot file becomes ``WORKBOOK_NOT_FOUND``.
"""

from __future__ import annotations

import asyncio
import os
from typing import Any

from excel_vba_mcp.adapters.ooxml.pivotcache import (
    iter_cache_records,
    read_cache_definition,
)
from excel_vba_mcp.adapters.ooxml.shared_strings import load_shared_strings
from excel_vba_mcp.adapters.ooxml.workbook_meta import (
    pivot_caches,
    sheet_part,
    workbook_parts,
)
from excel_vba_mcp.adapters.ooxml.worksheet import iter_sheet_rows
from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    open_safe,
)
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, WorkbookNotFoundError
from excel_vba_mcp.domain.models import (
    CellValue,
    OoxmlInputInfo,
    OoxmlMetadata,
    OoxmlPivotCacheMeta,
    OoxmlPivotCachePage,
    OoxmlSheetPage,
    OxmlSheetMeta,
    PageInfo,
)

#: Default number of rows returned when the caller does not supply a ``limit``.
#: Chosen so the defensive per-row cell cap (``max_response_cells // limit``)
#: stays reasonable for a typical worksheet.
_DEFAULT_ROW_LIMIT = 1000


class OoxmlService:
    """Read-only, non-COM OOXML reads over an immutable snapshot file.

    Args:
        session: The workbook session service (reserved for E2; not used by the
            snapshot-file read path).
        snapshot: The snapshot service used to create / look up the immutable
            input file.
        limits: The unified large-data limits (file/ZIP/XML/response bounds).
    """

    def __init__(
        self,
        session: Any,
        snapshot: SnapshotService,
        limits: LargeDataConfig,
    ) -> None:
        self._session = session
        self._snapshot = snapshot
        self._limits = limits
        self._zip_limits = ZipLimits.from_config(limits)

    # ------------------------------------------------------------------
    # Input resolution (async: create_snapshot runs on the STA worker)
    # ------------------------------------------------------------------
    async def _resolve_input(
        self, workbook_id: str, snapshot_id: str | None
    ) -> tuple[str, str, str | None]:
        """Resolve the immutable snapshot file to ``(path, sha256, snapshot_id)``.

        When ``snapshot_id`` is given, that snapshot is loaded from the store;
        otherwise a fresh snapshot of ``workbook_id`` is created (SaveCopyAs for
        an open workbook). The resolved file must exist on disk.

        Raises:
            WorkbookNotFoundError: If the resolved snapshot file is missing.
        """

        if snapshot_id is not None:
            manifest = self._snapshot.store.load(snapshot_id)
            path, sha, used_id = manifest.stored_path, manifest.sha256, snapshot_id
        else:
            snap = await self._snapshot.create_snapshot(workbook_id)
            path, sha, used_id = snap.stored_path, snap.sha256, snap.snapshot_id

        if not os.path.isfile(path):
            raise WorkbookNotFoundError(
                "The snapshot file to read does not exist.",
                context={"workbook_id": workbook_id, "snapshot_id": used_id},
            )
        return path, sha, used_id

    # ------------------------------------------------------------------
    # Public async API (parsing dispatched off the event loop)
    # ------------------------------------------------------------------
    async def read_metadata(
        self, workbook_id: str, snapshot_id: str | None = None
    ) -> OoxmlMetadata:
        """Read workbook metadata (sheets, defined names, table count) from a snapshot."""

        path, sha, used_id = await self._resolve_input(workbook_id, snapshot_id)
        return await asyncio.to_thread(
            self._read_metadata_sync, path, sha, used_id, workbook_id
        )

    async def read_sheet_rows(
        self,
        workbook_id: str,
        sheet: str,
        offset: int = 0,
        limit: int | None = None,
        snapshot_id: str | None = None,
    ) -> OoxmlSheetPage:
        """Read a bounded page of typed worksheet rows from a snapshot."""

        resolved_limit = limit if (limit is not None and limit > 0) else _DEFAULT_ROW_LIMIT
        path, sha, used_id = await self._resolve_input(workbook_id, snapshot_id)
        return await asyncio.to_thread(
            self._read_sheet_rows_sync,
            path,
            sha,
            used_id,
            workbook_id,
            sheet,
            offset,
            resolved_limit,
        )

    async def read_pivot_cache_records(
        self,
        workbook_id: str,
        cache_index: int,
        offset: int = 0,
        limit: int | None = None,
        snapshot_id: str | None = None,
    ) -> OoxmlPivotCachePage:
        """Read a bounded page of **saved** PivotCache records from a snapshot.

        The records are the data Excel saved with the workbook (not the PivotTable
        rendered result and not the live source). A cache with no local records
        part (``saveData="0"`` / OLAP / Data Model) yields an empty page with
        ``page.reason == "unsupported_source"`` — the records are never fabricated.
        """

        resolved_limit = limit if (limit is not None and limit > 0) else _DEFAULT_ROW_LIMIT
        path, sha, used_id = await self._resolve_input(workbook_id, snapshot_id)
        return await asyncio.to_thread(
            self._read_pivot_cache_records_sync,
            path,
            sha,
            used_id,
            workbook_id,
            cache_index,
            offset,
            resolved_limit,
        )

    # ------------------------------------------------------------------
    # Sync parsing helpers (pure file parsing; run via asyncio.to_thread)
    # ------------------------------------------------------------------
    def _read_metadata_sync(
        self, path: str, sha: str, snapshot_id: str | None, workbook_id: str
    ) -> OoxmlMetadata:
        try:
            with open_safe(path, self._zip_limits) as zf:
                parts = workbook_parts(zf, self._zip_limits)
        except OoxmlSafetyError as exc:
            raise self._unsafe_error(exc, workbook_id=workbook_id) from exc

        sheets = [
            OxmlSheetMeta(
                name=entry["name"],
                part=entry["part"],
                hidden=entry["hidden"],
                rows_hint=None,
            )
            for entry in parts["sheets"]
        ]
        return OoxmlMetadata(
            workbook_id=workbook_id,
            input=OoxmlInputInfo(input_sha256=sha, snapshot_id=snapshot_id),
            sheets=sheets,
            defined_names=list(parts["defined_names"]),
            table_count=len(parts["table_parts"]),
        )

    def _read_sheet_rows_sync(
        self,
        path: str,
        sha: str,
        snapshot_id: str | None,
        workbook_id: str,
        sheet: str,
        offset: int,
        limit: int,
    ) -> OoxmlSheetPage:
        try:
            with open_safe(path, self._zip_limits) as zf:
                parts = workbook_parts(zf, self._zip_limits)
                part = self._resolve_sheet_part(parts, sheet, workbook_id)
                shared = load_shared_strings(zf, self._zip_limits)
                rows, truncated, total_scanned = iter_sheet_rows(
                    zf, part, shared, self._zip_limits, offset, limit
                )
        except OoxmlSafetyError as exc:
            raise self._unsafe_error(
                exc, workbook_id=workbook_id, sheet=sheet
            ) from exc

        cols = max((len(row) for row in rows), default=0)
        returned = len(rows)
        page = PageInfo(
            offset=offset,
            limit=limit,
            returned=returned,
            total=total_scanned,
            has_more=truncated,
            complete=not truncated,
            truncated=truncated,
            version=0,
            source="snapshot",
        )
        return OoxmlSheetPage(
            workbook_id=workbook_id,
            sheet=sheet,
            input=OoxmlInputInfo(input_sha256=sha, snapshot_id=snapshot_id),
            rows=returned,
            cols=cols,
            values=rows,
            page=page,
        )

    def _read_pivot_cache_records_sync(
        self,
        path: str,
        sha: str,
        snapshot_id: str | None,
        workbook_id: str,
        cache_index: int,
        offset: int,
        limit: int,
    ) -> OoxmlPivotCachePage:
        try:
            with open_safe(path, self._zip_limits) as zf:
                caches = pivot_caches(zf, self._zip_limits)
                entry = self._resolve_cache(caches, cache_index, workbook_id)
                definition_part = entry.get("definition_part") or ""
                records_part = entry.get("records_part")
                definition = (
                    read_cache_definition(zf, definition_part, self._zip_limits)
                    if definition_part
                    else {}
                )
                save_data = bool(definition.get("save_data", True))
                has_records_part = records_part is not None
                rows: list[list[CellValue]] = []
                truncated = False
                total_scanned: int | None = None
                if has_records_part and save_data is not False:
                    assert records_part is not None  # narrowed by the branch above
                    shared_items = definition.get("shared_items", [])
                    rows, truncated, total_scanned, _scanned = iter_cache_records(
                        zf,
                        records_part,
                        shared_items,  # type: ignore[arg-type]
                        self._zip_limits,
                        offset,
                        limit,
                    )
        except OoxmlSafetyError as exc:
            raise self._unsafe_error(
                exc, workbook_id=workbook_id, cache_index=cache_index
            ) from exc

        meta = OoxmlPivotCacheMeta(
            cache_index=cache_index,
            record_count=definition.get("record_count"),  # type: ignore[arg-type]
            save_data=save_data,
            refresh_on_load=definition.get("refresh_on_load"),  # type: ignore[arg-type]
            refresh_date=definition.get("refresh_date"),  # type: ignore[arg-type]
            has_records_part=has_records_part,
        )

        unsupported = not has_records_part or save_data is False
        cols = max((len(row) for row in rows), default=0)
        returned = len(rows)
        total = meta.record_count if meta.record_count is not None else total_scanned
        page = PageInfo(
            offset=offset,
            limit=limit,
            returned=returned,
            total=total if not unsupported else 0,
            has_more=truncated,
            complete=not truncated and not unsupported,
            truncated=truncated,
            version=0,
            source="cache",
            reason="unsupported_source" if unsupported else None,
        )
        return OoxmlPivotCachePage(
            workbook_id=workbook_id,
            cache_index=cache_index,
            input=OoxmlInputInfo(input_sha256=sha, snapshot_id=snapshot_id),
            meta=meta,
            rows=returned,
            cols=cols,
            values=rows,
            page=page,
        )

    # ------------------------------------------------------------------
    # Error mapping
    # ------------------------------------------------------------------
    @staticmethod
    def _resolve_cache(
        caches: list[dict[str, Any]], cache_index: int, workbook_id: str
    ) -> dict[str, Any]:
        """Resolve a 1-based ``cache_index`` to its cache entry, or raise ``NOT_FOUND``.

        A missing cache is a caller error (``NOT_FOUND``), distinct from an unsafe
        archive (``INVALID_RANGE`` / ``unsafe_ooxml``).
        """

        for entry in caches:
            if entry.get("cache_index") == cache_index:
                return entry
        raise ExcelVbaMcpError(
            f"PivotCache index not found in the snapshot workbook: {cache_index}.",
            code=ErrorCode.NOT_FOUND,
            context={
                "reason": "cache_not_found",
                "workbook_id": workbook_id,
                "cache_index": cache_index,
            },
        )

    @staticmethod
    def _resolve_sheet_part(parts: dict[str, Any], sheet: str, workbook_id: str) -> str:
        """Resolve a sheet name to its part, or raise a clear ``sheet_not_found``.

        A missing sheet is a caller error (``INVALID_RANGE`` / ``sheet_not_found``),
        distinct from an unsafe archive; the underlying
        :func:`~excel_vba_mcp.adapters.ooxml.workbook_meta.sheet_part` still guards
        the (unexpected) no-part case as ``unsafe_ooxml``.
        """

        names = [entry["name"] for entry in parts.get("sheets", [])]
        lowered = sheet.lower()
        if sheet not in names and lowered not in {name.lower() for name in names}:
            raise ExcelVbaMcpError(
                f"Sheet not found in the snapshot workbook: {sheet!r}.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "reason": "sheet_not_found",
                    "workbook_id": workbook_id,
                    "sheet": sheet,
                },
            )
        return sheet_part(parts, sheet)

    @staticmethod
    def _unsafe_error(exc: OoxmlSafetyError, **context: Any) -> ExcelVbaMcpError:
        return ExcelVbaMcpError(
            str(exc),
            code=ErrorCode.INVALID_RANGE,
            context={"reason": "unsafe_ooxml", **context},
        )


__all__ = ["OoxmlService"]