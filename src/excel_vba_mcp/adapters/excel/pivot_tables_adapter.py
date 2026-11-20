"""STA-local Excel PivotTables / PivotCaches / PivotFields adapter (task_048).

Phase 2 — PivotTables inventory (read-only).

:class:`PivotTablesAdapter` is the single place where the Excel ``PivotTables`` /
``PivotCache`` / ``PivotField`` object model is touched for the Phase-2 inventory
surface (``pivot_tables_list``, ``pivot_table_get``, ``pivot_cache_get``,
``pivot_fields_list``). It is a **pure STA-local helper**: every method is
synchronous, receives the live ``Workbook`` COM proxy (``Any``) and returns only
Pydantic DTOs / primitives, so results cross the single-STA worker boundary
safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never
imports ``pywin32`` at module import time and never calls ``.Activate`` /
``.Select`` on any object (addendum BG03/BG04/BG05).

Inventory semantics (task_039 decision record, ``docs/research/pivot_tables_api.md``):

* **Worksheet sweep.** ``Workbook.PivotTables`` is unusable on the target build
  (returns an empty collection), so every PivotTable is discovered by sweeping
  ``wb.Worksheets`` and reading each ``ws.PivotTables()`` collection. The result
  is sorted deterministically by ``(sheet_index, pivot_table_name)``.
* **Cache identity.** ``PivotCaches`` is unusable for enumeration on this build;
  a cache is resolved through ``PivotTable.PivotCache()`` and identified by its
  ``CacheIndex``. Two PivotTables built from the same cache report the same
  ``CacheIndex`` (one cache N:1 PivotTables).
* **Field identity.** A field is ``(owning pivot table, source name)`` for source
  fields; data fields carry ``(owning pivot table, data-field caption)`` plus
  ``source_name``. ``PivotField.Position`` is 1-based **within its orientation**
  and returns the sentinel ``-2146826246`` (``#N/A``) for a hidden/unplaced field
  — surfaced as ``None`` in the DTO.
* **Data fields.** ``Function`` is readable only via ``PivotTable.DataFields``,
  never via ``PivotFields(...)``. Data fields are matched to source fields by
  ``SourceName`` to populate ``is_data_field`` / ``data_field_name`` /
  ``function`` / ``number_format``.

Source-type / unsupported policy (task_039 §7): only ``xlDatabase``
(``SourceType == 1``) is a supported refresh/source type. External,
consolidation, scenario and pivot-table sources are determinable but
``unsupported``; OLAP / Data Model is detected via ``SourceType``/``OLAP`` and is
never refreshed through this increment. ``SourceType`` text mapping:

===  ==========================  =========================
Int  ``XlPivotTableSourceType``  ``PivotSourceType``
===  ==========================  =========================
1    ``xlDatabase``              ``DATABASE``
2    ``xlExternal``              ``EXTERNAL``
3    ``xlConsolidation``         ``CONSOLIDATION``
4    ``xlScenario``              ``SCENARIO``
-4148 ``xlPivotTable``           ``PIVOT_TABLE``
—    OLAP indicator             ``OLAP`` / ``DATA_MODEL``
—    anything else              ``UNKNOWN``
===  ==========================  =========================

This adapter is **read-only**: it never creates/deletes PivotTables, never
refreshes, never mutates fields or filters. ``Workbook.RefreshAll`` and
Connections/QueryTables/Power Query/Data Model are never touched.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldInfo,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotSourceType,
)

# ``PivotField.Position`` sentinel returned by Excel for a hidden/unplaced field
# (``#N/A`` as a signed OLE error code). Surfaced as ``None`` in the DTO.
_POSITION_SENTINEL = -2146826246

# ``XlPivotFieldOrientation`` values (Microsoft Learn).
_XL_ORIENTATION_BY_INT: dict[int, PivotFieldOrientation] = {
    0: PivotFieldOrientation.HIDDEN,
    1: PivotFieldOrientation.ROW,
    2: PivotFieldOrientation.COLUMN,
    3: PivotFieldOrientation.PAGE,
    4: PivotFieldOrientation.DATA,
}

# ``XlPivotTableSourceType`` values (Microsoft Learn). ``-4148`` is
# ``xlPivotTable``; the remaining values are the local database/external/
# consolidation/scenario sources.
_XL_SOURCE_TYPE_BY_INT: dict[int, PivotSourceType] = {
    1: PivotSourceType.DATABASE,
    2: PivotSourceType.EXTERNAL,
    3: PivotSourceType.CONSOLIDATION,
    4: PivotSourceType.SCENARIO,
    -4148: PivotSourceType.PIVOT_TABLE,
}

# ``XlConsolidationFunction`` values accepted by the target build (task_039 §5).
# Only these eleven codes are modelled; legacy codes the build rejects are
# intentionally absent and surface as ``None``.
_XL_FUNCTION_BY_INT: dict[int, PivotAggregation] = {
    -4157: PivotAggregation.SUM,
    -4106: PivotAggregation.AVERAGE,
    -4112: PivotAggregation.COUNT,
    -4113: PivotAggregation.COUNT_NUMS,
    -4136: PivotAggregation.MAX,
    -4139: PivotAggregation.MIN,
    -4149: PivotAggregation.PRODUCT,
    -4155: PivotAggregation.STDEV,
    -4156: PivotAggregation.STDEVP,
    -4164: PivotAggregation.VAR,
    -4165: PivotAggregation.VARP,
}

# Source types that are determinable but unsupported for refresh / field mutation
# (task_039 §7). ``DATABASE`` is the only supported source type.
_UNSUPPORTED_SOURCE_TYPES: frozenset[PivotSourceType] = frozenset(
    {
        PivotSourceType.EXTERNAL,
        PivotSourceType.CONSOLIDATION,
        PivotSourceType.SCENARIO,
        PivotSourceType.PIVOT_TABLE,
        PivotSourceType.OLAP,
        PivotSourceType.DATA_MODEL,
        PivotSourceType.UNSUPPORTED,
        PivotSourceType.UNKNOWN,
    }
)


# ---------------------------------------------------------------------------
# Safe COM read helpers (never raise; defensive per-property)
# ---------------------------------------------------------------------------
def _safe_str(value: Any, default: str | None = None) -> str | None:
    """Read a COM string property defensively (never raises)."""

    try:
        text = value
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    if text is None:
        return default
    if isinstance(text, str):
        return text or default
    return str(text)


def _safe_int(value: Any, default: int | None = None) -> int | None:
    """Read a COM integer property defensively (never raises)."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_bool(value: Any, default: bool | None = None) -> bool | None:
    """Read a COM boolean property defensively (never raises)."""

    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_address(range_obj: Any) -> str | None:
    """Return ``range_obj.Address`` as a string, or ``None`` when absent."""

    if range_obj is None:
        return None
    try:
        address = range_obj.Address
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    if address is None:
        return None
    text = str(address)
    return text or None


def _safe_datetime(value: Any) -> datetime | None:
    """Read a COM date property defensively into a ``datetime`` (or ``None``).

    Excel returns a ``pywintypes``/``datetime`` for a date property; a non-date
    value (or an absent property) yields ``None`` rather than a guess.
    """

    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value))
    except (TypeError, ValueError):
        return None


# The ``_get_*`` helpers below read a COM property by name and **swallow the
# ``com_error`` raised by the property getter itself**. This matters for
# ``PivotField`` properties such as ``AllItemsVisible`` / ``ShowAllItems`` that
# raise ``DISP_E_EXCEPTION`` (0x80020009) when read from a plain
# ``PivotFields(...)`` handle on the target build (task_039 §3): ``getattr(obj,
# name, default)`` does NOT swallow a ``com_error`` (only ``AttributeError``), so
# the attribute access must happen inside the try.
def _get_str(obj: Any, attr: str, default: str | None = None) -> str | None:
    """Read ``obj.<attr>`` as a string, swallowing a raising COM getter."""

    try:
        return _safe_str(getattr(obj, attr))
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _get_int(obj: Any, attr: str, default: int | None = None) -> int | None:
    """Read ``obj.<attr>`` as an int, swallowing a raising COM getter."""

    try:
        return _safe_int(getattr(obj, attr))
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _get_bool(obj: Any, attr: str, default: bool | None = None) -> bool | None:
    """Read ``obj.<attr>`` as a bool, swallowing a raising COM getter."""

    try:
        return _safe_bool(getattr(obj, attr))
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _get_address(obj: Any, attr: str) -> str | None:
    """Read ``obj.<attr>.Address`` as a string, swallowing a raising COM getter."""

    try:
        return _safe_address(getattr(obj, attr))
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _get_datetime(obj: Any, attr: str) -> datetime | None:
    """Read ``obj.<attr>`` as a datetime, swallowing a raising COM getter."""

    try:
        return _safe_datetime(getattr(obj, attr))
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


# ---------------------------------------------------------------------------
# Source-type classification
# ---------------------------------------------------------------------------
def _classify_source_type(cache: Any) -> tuple[PivotSourceType, bool]:
    """Classify a ``PivotCache`` source type (task_039 §7).

    Returns ``(source_type, olap)``. ``olap`` is ``True`` when the cache exposes
    an OLAP / Data Model indicator; such a cache is reported as ``OLAP`` /
    ``DATA_MODEL`` and never refreshed.
    """

    raw = _get_int(cache, "SourceType")
    olap = bool(_get_bool(cache, "OLAP", default=False))
    if raw is not None and raw in _XL_SOURCE_TYPE_BY_INT:
        source_type = _XL_SOURCE_TYPE_BY_INT[raw]
    else:
        source_type = PivotSourceType.UNKNOWN

    if olap:
        # OLAP / Data Model indicators override the numeric source type: both are
        # unsupported/read-only through this increment (task_039 §7).
        return PivotSourceType.OLAP, True
    return source_type, olap


def _source_status(source_type: PivotSourceType) -> str:
    """Return the ``source_status`` string for a source type (task_039 §7)."""

    if source_type in _UNSUPPORTED_SOURCE_TYPES:
        return "unsupported_source"
    return "supported"


# ---------------------------------------------------------------------------
# PivotField description (read-only)
# ---------------------------------------------------------------------------
def _orientation_of(field: Any) -> PivotFieldOrientation:
    """Read a field's ``Orientation`` as a :class:`PivotFieldOrientation`."""

    raw = _get_int(field, "Orientation")
    if raw is None:
        return PivotFieldOrientation.HIDDEN
    return _XL_ORIENTATION_BY_INT.get(raw, PivotFieldOrientation.HIDDEN)


def _position_of(field: Any) -> int | None:
    """Read a field's 1-based ``Position``; ``None`` when hidden/unplaced."""

    raw = _get_int(field, "Position")
    if raw is None or raw == _POSITION_SENTINEL:
        return None
    return raw


def _is_synthetic_data_placeholder(field: Any) -> bool:
    """Return whether ``field`` is the synthetic ``"Data"`` placeholder.

    When a PivotTable has two or more data fields, Excel inserts a synthetic
    ``"Data"`` field into the row/column orientation collections. Its
    ``SourceName`` is the ``#N/A`` error code (the ``-2146826246`` sentinel)
    rather than a real source field name (task_039 §2, §3). It is a layout
    artifact, not a real source field, so it is excluded from the reported field
    layout. The raw ``SourceName`` is read as an int here — a real source name is
    a non-empty string, so an int sentinel identifies the placeholder.
    """

    try:
        raw = field.SourceName
    except Exception:  # noqa: BLE001 - defensive COM read
        return False
    # A real source name is a non-empty string; the placeholder is the int #N/A.
    if isinstance(raw, str):
        return raw == ""
    return bool(raw == _POSITION_SENTINEL)


def _describe_data_field(data_field: Any) -> dict[str, Any]:
    """Read one ``DataFields`` item's aggregation metadata (STA-local).

    Returns a plain dict keyed by the source name so the owning PivotField DTO can
    be enriched without re-reading ``PivotFields(...)`` (which cannot expose
    ``Function`` on this build — task_039 §3).
    """

    function_code = _get_int(data_field, "Function")
    return {
        "source_name": _get_str(data_field, "SourceName"),
        "caption": _get_str(data_field, "Name"),
        "function": _XL_FUNCTION_BY_INT.get(function_code) if function_code is not None else None,
        "number_format": _get_str(data_field, "NumberFormat"),
        "position": _position_of(data_field),
    }


def _data_field_index(pivot: Any) -> dict[str, dict[str, Any]]:
    """Map ``source_name`` -> data-field metadata for one PivotTable."""

    index: dict[str, dict[str, Any]] = {}
    try:
        data_fields = pivot.DataFields
        count = _safe_int(data_fields.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        return index
    for position in range(1, count + 1):
        try:
            item = data_fields.Item(position)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        described = _describe_data_field(item)
        source_name = described.get("source_name")
        if isinstance(source_name, str) and source_name:
            index[source_name] = described
    return index


def _filter_kind_of(field: Any) -> PivotFilterKind:
    """Classify a field's manual filter state (task_039 §6).

    Only the manual (``PivotItems.Visible``) filter state is classified here;
    label/value/date filters are unsupported on the target build and are never
    created, so a field with all items visible reports ``NONE``. A field with at
    least one hidden item reports ``MANUAL``. ``AllItemsVisible`` raises on a
    plain ``PivotFields(...)`` handle for some fields, so it is read defensively
    and a missing value defaults to "all visible".
    """

    if _get_bool(field, "AllItemsVisible", default=True):
        return PivotFilterKind.NONE
    return PivotFilterKind.MANUAL


def _describe_field(
    pivot: Any,
    field: Any,
    data_field_index: dict[str, dict[str, Any]],
) -> PivotFieldInfo:
    """Build a :class:`PivotFieldInfo` for one ``PivotField`` (STA-local)."""

    source_name = _get_str(field, "SourceName") or ""
    caption = _get_str(field, "Caption")
    orientation = _orientation_of(field)
    position = _position_of(field)
    show_all_items = _get_bool(field, "ShowAllItems")
    all_items_visible = _get_bool(field, "AllItemsVisible")
    filter_kind = _filter_kind_of(field)

    data_meta = data_field_index.get(source_name)
    is_data_field = data_meta is not None
    function: PivotAggregation | None = None
    number_format: str | None = None
    data_field_name: str | None = None
    if data_meta is not None:
        function = data_meta.get("function")
        number_format = data_meta.get("number_format")
        data_field_name = data_meta.get("caption")
        # A data field's orientation is DATA even when the source field is also
        # placed in a row/column/page bucket.
        orientation = PivotFieldOrientation.DATA
        if position is None:
            position = data_meta.get("position")

    return PivotFieldInfo(
        source_name=source_name,
        caption=caption,
        orientation=orientation,
        position=position,
        function=function,
        show_all_items=show_all_items,
        all_items_visible=all_items_visible,
        filter_state=filter_kind,
        is_data_field=is_data_field,
        data_field_name=data_field_name,
        number_format=number_format,
    )


# ---------------------------------------------------------------------------
# PivotTable description (read-only)
# ---------------------------------------------------------------------------
def _describe_pivot_table(
    wb: Any,
    sheet: Any,
    sheet_index: int,
    pivot: Any,
) -> PivotTableInfo:
    """Build a :class:`PivotTableInfo` for one PivotTable (STA-local).

    The description reads the PivotTable's identity, range metadata, its cache's
    source type and its field layout. It never materializes the PivotTable cell
    output — only metadata.
    """

    name = _get_str(pivot, "Name") or ""
    cache = pivot.PivotCache()
    source_type, _olap = _classify_source_type(cache)

    data_field_index = _data_field_index(pivot)

    # Enumerate every source field once and bucket it by orientation.
    data_fields: list[PivotFieldInfo] = []
    row_fields: list[PivotFieldInfo] = []
    column_fields: list[PivotFieldInfo] = []
    page_fields: list[PivotFieldInfo] = []
    hidden_fields: list[PivotFieldInfo] = []

    try:
        pivot_fields = pivot.PivotFields()
        field_count = _safe_int(pivot_fields.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        pivot_fields = None
        field_count = 0

    for position in range(1, field_count + 1):
        if pivot_fields is None:
            break
        try:
            field = pivot_fields.Item(position)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        # Skip the synthetic "Data" placeholder inserted when a PivotTable has
        # two or more data fields (its SourceName is the #N/A sentinel).
        if _is_synthetic_data_placeholder(field):
            continue
        info = _describe_field(pivot, field, data_field_index)
        if info.is_data_field or info.orientation is PivotFieldOrientation.DATA:
            data_fields.append(info)
        elif info.orientation is PivotFieldOrientation.ROW:
            row_fields.append(info)
        elif info.orientation is PivotFieldOrientation.COLUMN:
            column_fields.append(info)
        elif info.orientation is PivotFieldOrientation.PAGE:
            page_fields.append(info)
        else:
            hidden_fields.append(info)

    return PivotTableInfo(
        name=name,
        sheet_name=_get_str(sheet, "Name") or "",
        sheet_index=sheet_index,
        index=_get_int(pivot, "Index", default=0) or 0,
        cache_index=_get_int(pivot, "CacheIndex", default=0) or 0,
        table_range1=_get_address(pivot, "TableRange1"),
        table_range2=_get_address(pivot, "TableRange2"),
        data_body_range=_get_address(pivot, "DataBodyRange"),
        row_grand=_get_bool(pivot, "RowGrand"),
        column_grand=_get_bool(pivot, "ColumnGrand"),
        refresh_date=_get_datetime(pivot, "RefreshDate"),
        source_type=source_type,
        source_status=_source_status(source_type),
        data_fields=data_fields,
        row_fields=row_fields,
        column_fields=column_fields,
        page_fields=page_fields,
        hidden_fields=hidden_fields,
    )


# ---------------------------------------------------------------------------
# Full-inventory enumeration (worksheet sweep)
# ---------------------------------------------------------------------------
def _iter_pivot_tables(wb: Any) -> list[tuple[Any, int, Any]]:
    """Return ``(sheet, sheet_index, pivot)`` for every PivotTable on every sheet.

    ``Workbook.PivotTables`` is unusable on the target build (returns an empty
    collection), so the sweep goes through each worksheet's ``PivotTables()``
    collection (task_039 §2).
    """

    triples: list[tuple[Any, int, Any]] = []
    try:
        worksheets = wb.Worksheets
    except Exception:  # noqa: BLE001 - defensive COM read
        return triples
    for sheet_index, sheet in enumerate(worksheets, start=1):
        try:
            pivot_tables = sheet.PivotTables()
            count = _safe_int(pivot_tables.Count, default=0) or 0
        except Exception:  # noqa: BLE001 - sheet without PivotTables
            continue
        for position in range(1, count + 1):
            try:
                pivot = pivot_tables.Item(position)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
            triples.append((sheet, sheet_index, pivot))
    return triples


def _describe_cache(
    wb: Any, cache_index: int, pivot_tables: list[PivotTableInfo]
) -> PivotCacheInfo:
    """Build a :class:`PivotCacheInfo` for ``cache_index`` (STA-local).

    The cache is resolved through the first PivotTable whose ``CacheIndex``
    matches (the ``PivotCaches`` collection is unusable for enumeration on this
    build — task_039 §3). The shared-cache relationship is computed from the
    already-described PivotTables: every table with the same ``cache_index`` is
    listed in ``affected_pivot_tables``.
    """

    cache = None
    for _sheet, _sheet_index, pivot in _iter_pivot_tables(wb):
        pivot_cache_index = _get_int(pivot, "CacheIndex")
        if pivot_cache_index == cache_index:
            cache = pivot.PivotCache()
            break
    if cache is None:
        raise ExcelVbaMcpError(
            f"PivotCache {cache_index} was not found in the workbook.",
            context={"cache_index": cache_index},
        )

    source_type, _olap = _classify_source_type(cache)
    source_data_raw = getattr(cache, "SourceData", None)
    source_data: str | None
    source_data_kind = "text"
    if isinstance(source_data_raw, str):
        source_data = source_data_raw
    elif source_data_raw is None:
        source_data, source_data_kind = None, "unavailable"
    else:
        # External sources return a Variant array; a bounded text rendering is
        # exposed rather than a raw nested array (task_039 §7).
        source_data, source_data_kind = str(source_data_raw), "array"

    affected = [pt.name for pt in pivot_tables if pt.cache_index == cache_index]
    return PivotCacheInfo(
        cache_index=cache_index,
        source_type=source_type,
        source_data=source_data,
        source_data_kind=source_data_kind,
        refresh_date=_get_datetime(cache, "RefreshDate"),
        refresh_on_file_open=_get_bool(cache, "RefreshOnFileOpen"),
        enable_refresh=_get_bool(cache, "EnableRefresh"),
        record_count=_get_int(cache, "RecordCount"),
        olap=_olap,
        shared_cache=len(affected) > 1,
        affected_pivot_tables=affected,
    )


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class PivotTablesAdapter:
    """STA-local PivotTables / PivotCaches / PivotFields inventory (no COM at import).

    All methods are synchronous, operate on the live ``Workbook`` proxy and
    return only Pydantic DTOs. They never create, delete, refresh or mutate a
    PivotTable, PivotCache or PivotField (task_048 is read-only).
    """

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def enumerate_pivot_tables(self, wb: Any) -> list[PivotTableInfo]:
        """Return one :class:`PivotTableInfo` per PivotTable on every worksheet.

        The sweep is worksheet-scoped (``ws.PivotTables()``) because
        ``Workbook.PivotTables`` is unusable on the target build. The result is
        sorted deterministically by ``(sheet_index, pivot_table_name)``.
        """

        described = [
            _describe_pivot_table(wb, sheet, sheet_index, pivot)
            for sheet, sheet_index, pivot in _iter_pivot_tables(wb)
        ]
        described.sort(key=lambda pt: (pt.sheet_index, pt.name))
        return described

    def read_pivot_cache(self, wb: Any, cache_index: int) -> PivotCacheInfo:
        """Return the :class:`PivotCacheInfo` for a 1-based ``cache_index``.

        The cache is resolved through the PivotTable that references it (the
        ``PivotCaches`` collection is unusable for enumeration). The shared-cache
        relationship is computed from the full PivotTable inventory so
        ``affected_pivot_tables`` lists every table sharing this cache.
        """

        pivot_tables = self.enumerate_pivot_tables(wb)
        return _describe_cache(wb, cache_index, pivot_tables)

    def enumerate_pivot_fields(self, wb: Any, pivot_table: str) -> list[PivotFieldInfo]:
        """Return every :class:`PivotFieldInfo` of one PivotTable (read-only).

        Fields are returned in ``PivotFields`` order, each enriched with its
        orientation, 1-based position within that orientation, data-field
        aggregation (from ``DataFields``) and manual filter state. The
        ``pivot_table`` argument is the stable PivotTable ``Name``; a missing name
        raises a structured error.
        """

        target = pivot_table.strip().lower()
        for _sheet, _sheet_index, pivot in _iter_pivot_tables(wb):
            name = _get_str(pivot, "Name") or ""
            if name.lower() == target:
                data_field_index = _data_field_index(pivot)
                return self._describe_all_fields(pivot, data_field_index)
        raise ExcelVbaMcpError(
            f"PivotTable {pivot_table!r} was not found in the workbook.",
            context={"pivot_table": pivot_table},
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _describe_all_fields(
        pivot: Any, data_field_index: dict[str, dict[str, Any]]
    ) -> list[PivotFieldInfo]:
        """Describe every ``PivotFields`` entry of one PivotTable (STA-local)."""

        fields: list[PivotFieldInfo] = []
        try:
            pivot_fields = pivot.PivotFields()
            count = _safe_int(pivot_fields.Count, default=0) or 0
        except Exception:  # noqa: BLE001 - defensive COM read
            return fields
        for position in range(1, count + 1):
            try:
                field = pivot_fields.Item(position)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
            # Skip the synthetic "Data" placeholder (SourceName is the #N/A
            # sentinel) inserted when a PivotTable has two or more data fields.
            if _is_synthetic_data_placeholder(field):
                continue
            fields.append(_describe_field(pivot, field, data_field_index))
        return fields


__all__ = ["PivotTablesAdapter"]
