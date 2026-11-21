"""STA-local Excel PivotTables / PivotCaches / PivotFields adapter (task_048, task_049, task_050).

Phase 2 — PivotTables inventory (read-only), targeted refresh (task_049) and
PivotField configuration (task_050).

:class:`PivotTablesAdapter` is the single place where the Excel ``PivotTables`` /
``PivotCache`` / ``PivotField`` object model is touched for the Phase-2
PivotTables surface (``pivot_tables_list``, ``pivot_table_get``,
``pivot_cache_get``, ``pivot_fields_list``, ``pivot_table_refresh``,
``pivot_cache_refresh``, ``pivot_field_set_layout``,
``pivot_field_set_position``, ``pivot_field_set_aggregation``,
``pivot_field_set_properties``). It is a **pure STA-local helper**: every method
is synchronous, receives the live ``Workbook`` COM proxy (``Any``) and returns
only Pydantic DTOs / primitives / plain containers, so results cross the
single-STA worker boundary safely
(:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never imports
``pywin32`` at module import time and never calls ``.Activate`` / ``.Select`` on
any object (addendum BG03/BG04/BG05).

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

This adapter never creates/deletes PivotTables and never touches
``Workbook.RefreshAll``, Connections/QueryTables/Power Query/Data Model. A
refresh is confirmed only by Excel's truth signal (the ``RefreshTable`` Boolean
and/or an advanced ``RefreshDate``), never by the absence of a COM exception.

Field-configuration semantics (task_050, ``docs/research/pivot_tables_api.md``
§5): ``Orientation`` accepts only the layout buckets — ``xlDataField`` (4) is
rejected because assigning it does **not** create a data field, which is
``PivotTable.AddDataField``. Removal is ``Orientation = 0`` (non-destructive,
the default) or ``PivotField.Delete`` (destructive, explicit). ``Position`` is
1-based within one orientation and validated before the write. Aggregation is
read/written only through ``DataFields[i].Function`` (eleven accepted codes;
legacy codes are rejected) and ``DataFields[i].Caption`` is the data-field
caption (unique). Every mutation is followed by a post-mutation re-read; a
mismatch is reported through ``verified=False`` rather than claimed as success.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

from excel_vba_mcp.domain.errors import (
    AmbiguousFilterError,
    ErrorCode,
    ExcelVbaMcpError,
    InvalidRangeError,
    UnsupportedPivotSourceError,
)
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldInfo,
    PivotFieldProperties,
    PivotRefreshResult,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotRefreshMethod,
    PivotRefreshStatus,
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

# The layout buckets a field's ``Orientation`` may be **assigned** to (task_039
# §5). ``xlDataField`` (4) is deliberately absent: assigning it does not create a
# data field, so ``PivotTable.AddDataField`` is the only way into the data area.
_XL_ORIENTATION_CODES: dict[PivotFieldOrientation, int] = {
    PivotFieldOrientation.HIDDEN: 0,
    PivotFieldOrientation.ROW: 1,
    PivotFieldOrientation.COLUMN: 2,
    PivotFieldOrientation.PAGE: 3,
}

# The accepted ``DataFields[i].Function`` codes (task_039 §5). Legacy codes the
# target build rejects (``9``, ``-4114``, ``-4153``, ``-4154``) are absent, so an
# unknown code is rejected rather than silently written.
_XL_FUNCTION_CODES: dict[PivotAggregation, int] = {
    PivotAggregation.SUM: -4157,
    PivotAggregation.AVERAGE: -4106,
    PivotAggregation.COUNT: -4112,
    PivotAggregation.COUNT_NUMS: -4113,
    PivotAggregation.MAX: -4136,
    PivotAggregation.MIN: -4139,
    PivotAggregation.PRODUCT: -4149,
    PivotAggregation.STDEV: -4155,
    PivotAggregation.STDEVP: -4156,
    PivotAggregation.VAR: -4164,
    PivotAggregation.VARP: -4165,
}

# The orientation collections used to validate a position and to report the
# resulting order. A hidden field has no such collection (its order is the
# ``PivotFields`` order).
_ORIENTATION_COLLECTION: dict[PivotFieldOrientation, str] = {
    PivotFieldOrientation.ROW: "RowFields",
    PivotFieldOrientation.COLUMN: "ColumnFields",
    PivotFieldOrientation.PAGE: "PageFields",
    PivotFieldOrientation.DATA: "DataFields",
}

# ``PivotFieldProperties`` fields accepted by ``set_field_properties``.
_FIELD_PROPERTY_NAMES: frozenset[str] = frozenset({"caption", "number_format", "show_all_items"})

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
# ---------------------------------------------------------------------------
# Targeted refresh (task_049)
# ---------------------------------------------------------------------------
def _affected_pivot_tables(pivot_tables: list[PivotTableInfo], cache_index: int) -> list[str]:
    """Return the names of every PivotTable sharing ``cache_index`` (sorted)."""

    return sorted(pt.name for pt in pivot_tables if pt.cache_index == cache_index)


def _find_pivot(wb: Any, pivot_table: str) -> Any:
    """Return the live ``PivotTable`` proxy named ``pivot_table`` (STA-local)."""

    target = pivot_table.strip().lower()
    for _sheet, _sheet_index, pivot in _iter_pivot_tables(wb):
        name = _get_str(pivot, "Name") or ""
        if name.lower() == target:
            return pivot
    raise ExcelVbaMcpError(
        f"PivotTable {pivot_table!r} was not found in the workbook.",
        context={"pivot_table": pivot_table},
    )


def _find_cache(wb: Any, cache_index: int) -> Any:
    """Return the live ``PivotCache`` proxy for ``cache_index`` (STA-local).

    ``Workbook.PivotCaches`` is unusable for enumeration on the target build
    (task_039 §3), so the cache is resolved through the PivotTable that
    references it.
    """

    for _sheet, _sheet_index, pivot in _iter_pivot_tables(wb):
        if _get_int(pivot, "CacheIndex") == cache_index:
            try:
                return pivot.PivotCache()
            except Exception as exc:  # noqa: BLE001 - defensive COM read
                raise ExcelVbaMcpError(
                    f"The PivotCache for index {cache_index} could not be resolved.",
                    context={"cache_index": cache_index},
                ) from exc
    raise ExcelVbaMcpError(
        f"PivotCache {cache_index} was not found in the workbook.",
        context={"cache_index": cache_index},
    )


def _unsupported_reason(source_type: PivotSourceType) -> str:
    """Return the human-readable reason a source type cannot be refreshed."""

    return (
        f"Pivot source type {source_type.value!r} is external, OLAP, Data Model "
        "or otherwise unsupported; only worksheet/Table-backed (database) caches "
        "are refreshed."
    )


def _unsupported_result(
    *,
    pivot_table: str,
    cache_index: int,
    method: PivotRefreshMethod,
    affected: list[str],
    source_type: PivotSourceType,
    before: datetime | None,
) -> PivotRefreshResult:
    """Build the structured ``unsupported_source`` result (no refresh invoked)."""

    return PivotRefreshResult(
        pivot_table=pivot_table,
        cache_index=cache_index,
        method=method,
        affected_pivot_tables=affected,
        refreshed=False,
        refresh_return=None,
        refresh_date_before=before,
        refresh_date_after=before,
        status=PivotRefreshStatus.UNSUPPORTED_SOURCE,
        error={
            "code": "UNSUPPORTED_PIVOT_SOURCE",
            "source_type": source_type.value,
            "reason": _unsupported_reason(source_type),
        },
    )


def _failed_result(
    *,
    pivot_table: str,
    cache_index: int,
    method: PivotRefreshMethod,
    affected: list[str],
    before: datetime | None,
    after: datetime | None,
    refresh_return: bool | None,
    started_at: datetime,
    finished_at: datetime,
    reason: str,
) -> PivotRefreshResult:
    """Build the structured ``failed`` result (no truth signal from Excel)."""

    duration_ms = int((finished_at - started_at).total_seconds() * 1000)
    return PivotRefreshResult(
        pivot_table=pivot_table,
        cache_index=cache_index,
        method=method,
        affected_pivot_tables=affected,
        refreshed=False,
        refresh_return=refresh_return,
        started_at=started_at,
        finished_at=finished_at,
        duration_ms=duration_ms,
        refresh_date_before=before,
        refresh_date_after=after,
        status=PivotRefreshStatus.FAILED,
        error={"code": "REFRESH_FAILED", "reason": reason},
    )


def _refreshed_result(
    *,
    pivot_table: str,
    cache_index: int,
    method: PivotRefreshMethod,
    affected: list[str],
    before: datetime | None,
    after: datetime | None,
    refresh_return: bool | None,
    started_at: datetime,
    finished_at: datetime,
) -> PivotRefreshResult:
    """Build the successful ``refreshed`` result."""

    duration_ms = int((finished_at - started_at).total_seconds() * 1000)
    return PivotRefreshResult(
        pivot_table=pivot_table,
        cache_index=cache_index,
        method=method,
        affected_pivot_tables=affected,
        refreshed=True,
        refresh_return=refresh_return,
        started_at=started_at,
        finished_at=finished_at,
        duration_ms=duration_ms,
        refresh_date_before=before,
        refresh_date_after=after,
        status=PivotRefreshStatus.REFRESHED,
    )


def _advance_confirms(before: datetime | None, after: datetime | None) -> bool:
    """Return whether ``after`` shows the refresh advanced past ``before``.

    A refresh is confirmed only by a truth signal: ``RefreshDate`` moved forward,
    or a previously-absent ``RefreshDate`` is now present. Equal timestamps are
    not proof (task_039 §4 verification rule).
    """

    if after is None:
        return False
    if before is None:
        return True
    return after > before


# ---------------------------------------------------------------------------
# PivotField configuration helpers (task_050; STA-local, no COM at import)
# ---------------------------------------------------------------------------
def _require_supported_source(pivot: Any) -> PivotSourceType:
    """Reject an OLAP / Data Model / external cache before any field mutation.

    ``DATABASE`` (a worksheet or ListObject range) is the only source type whose
    PivotFields may be reconfigured (task_039 §7). The check runs before the
    first mutation, so an unsupported field is never touched.
    """

    cache = pivot.PivotCache()
    source_type, _olap = _classify_source_type(cache)
    if source_type is not PivotSourceType.DATABASE:
        raise UnsupportedPivotSourceError(
            _unsupported_reason(source_type),
            context={
                "source_type": source_type.value,
                "reason": "unsupported_pivot_source",
            },
        )
    return source_type


def _source_field_entries(pivot: Any) -> list[tuple[Any, str]]:
    """Return ``(field, source_name)`` for every non-synthetic ``PivotFields`` entry."""

    entries: list[tuple[Any, str]] = []
    try:
        pivot_fields = pivot.PivotFields()
        count = _safe_int(pivot_fields.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        return entries
    for position in range(1, count + 1):
        try:
            field = pivot_fields.Item(position)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        if _is_synthetic_data_placeholder(field):
            continue
        source_name = _get_str(field, "SourceName") or ""
        entries.append((field, source_name))
    return entries


def _data_field_entries(pivot: Any) -> list[tuple[int, str, str]]:
    """Return ``(position, caption, source_name)`` for every ``DataFields`` entry."""

    entries: list[tuple[int, str, str]] = []
    try:
        data_fields = pivot.DataFields
        count = _safe_int(data_fields.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        return entries
    for position in range(1, count + 1):
        try:
            item = data_fields.Item(position)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        caption = _get_str(item, "Name") or ""
        source_name = _get_str(item, "SourceName") or ""
        entries.append((position, caption, source_name))
    return entries


def _raw_orientation(field: Any) -> PivotFieldOrientation | None:
    """Read a field's own ``Orientation`` (no data-field override)."""

    code = _get_int(field, "Orientation")
    if code is None:
        return None
    return _XL_ORIENTATION_BY_INT.get(code)


def _raw_position(field: Any) -> int | None:
    """Read a field's own 1-based ``Position`` (sentinel -> ``None``)."""

    return _position_of(field)


def _field_dto(
    pivot: Any,
    field: Any,
    data_field_index: dict[str, dict[str, Any]],
    *,
    source_layout: bool,
) -> PivotFieldInfo:
    """Describe a field, optionally keeping its **own** orientation/position.

    ``_describe_field`` reports a source field that also has a data field as
    ``DATA`` (the inventory convention). A layout operation is explicitly about
    the source field's placement, so it re-reads the field's own orientation and
    position instead; every other attribute (aggregation, caption, number
    format, filter state) is unchanged.
    """

    described = _describe_field(pivot, field, data_field_index)
    if not source_layout:
        return described
    raw_orientation = _raw_orientation(field)
    if raw_orientation is None:
        return described
    return described.model_copy(
        update={"orientation": raw_orientation, "position": _raw_position(field)}
    )


def _resolve_source_field(pivot: Any, field: str) -> tuple[Any, PivotFieldInfo]:
    """Resolve ``field`` to its unique ``PivotField`` and describe its state.

    The identity is the source field ``SourceName`` (task_039 §8). A name that
    matches no source field, or more than one, is rejected as an ambiguous
    identity rather than guessed.
    """

    target_name = field.strip().lower()
    matches = [
        (candidate, name)
        for candidate, name in _source_field_entries(pivot)
        if name.strip().lower() == target_name
    ]
    if not matches:
        raise AmbiguousFilterError(
            f"Field {field!r} was not found in the PivotTable.",
            context={"field": field, "reason": "field_not_found"},
        )
    if len(matches) > 1:
        raise AmbiguousFilterError(
            f"Field {field!r} resolved to {len(matches)} source fields; the identity is ambiguous.",
            context={"field": field, "match_count": len(matches), "reason": "ambiguous_field"},
        )
    candidate, source_name = matches[0]
    info = _describe_field(pivot, candidate, _data_field_index(pivot))
    return candidate, info


def _resolve_field_target(pivot: Any, field: str) -> tuple[Any, PivotFieldInfo]:
    """Resolve ``field`` by data-field caption first, then by source name.

    ``set_field_properties`` accepts either a data-field caption (the display
    ``Name`` of a ``DataFields`` entry) or a source field ``SourceName``; a
    caption match is preferred because it is the data field's own identity
    (task_039 §8, P12).
    """

    target_name = field.strip().lower()
    caption_matches = [
        entry for entry in _data_field_entries(pivot) if entry[1].strip().lower() == target_name
    ]
    if len(caption_matches) == 1:
        return _resolve_source_field(pivot, caption_matches[0][2])
    return _resolve_source_field(pivot, field)


def _field_absent(pivot: Any, field: str) -> bool:
    """Return whether ``field`` is gone from ``PivotFields`` (postcondition)."""

    target_name = field.strip().lower()
    return not any(name.strip().lower() == target_name for _f, name in _source_field_entries(pivot))


def _orientation_collection(pivot: Any, orientation: PivotFieldOrientation) -> Any:
    """Return the orientation's field collection, or ``None`` for a hidden field."""

    collection_name = _ORIENTATION_COLLECTION.get(orientation)
    if collection_name is None:
        return None
    try:
        return getattr(pivot, collection_name)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _orientation_count(pivot: Any, orientation: PivotFieldOrientation) -> int:
    """Return the number of fields in one orientation (0 when unreadable)."""

    collection = _orientation_collection(pivot, orientation)
    if collection is None:
        return 0
    try:
        return _safe_int(collection.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        return 0


def _orientation_order(pivot: Any, orientation: PivotFieldOrientation) -> list[str]:
    """Return the field names of one orientation in their current order.

    Used to surface the resulting order after a position change so the caller can
    verify the swap semantics without materializing the PivotTable output
    (task_039 §5).
    """

    if orientation is PivotFieldOrientation.HIDDEN:
        return [name for _f, name in _source_field_entries(pivot)]
    collection = _orientation_collection(pivot, orientation)
    if collection is None:
        return []
    try:
        count = _safe_int(collection.Count, default=0) or 0
    except Exception:  # noqa: BLE001 - defensive COM read
        return []
    entries: list[tuple[int | None, str]] = []
    for position in range(1, count + 1):
        try:
            item = collection.Item(position)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        # ``Name`` is the data-field caption for a DataFields entry and the field
        # name for a source field; fall back to ``SourceName``.
        name = _get_str(item, "Name") or _get_str(item, "SourceName") or ""
        entries.append((_get_int(item, "Position"), name))
    entries.sort(key=lambda pair: (pair[0] is None, pair[0] or 0))
    return [name for _position, name in entries]


def _require_layout_code(orientation: PivotFieldOrientation, *, delete: bool) -> int:
    """Validate a target orientation and return its ``Orientation`` code."""

    if delete and orientation is not PivotFieldOrientation.HIDDEN:
        raise InvalidRangeError(
            "A destructive field removal (delete=True) requires orientation "
            "'hidden'; set the orientation to 'hidden' to remove the field.",
            context={
                "orientation": orientation.value,
                "reason": "invalid_orientation",
            },
        )
    if orientation is PivotFieldOrientation.DATA:
        raise InvalidRangeError(
            "Orientation 'data' cannot be assigned: setting Orientation = "
            "xlDataField does not create a data field. Use "
            "pivot_field_set_aggregation (PivotTable.AddDataField) instead.",
            context={"orientation": "data", "reason": "invalid_orientation"},
        )
    code = _XL_ORIENTATION_CODES.get(orientation)
    if code is None:
        raise InvalidRangeError(
            f"Orientation {orientation.value!r} is not a layout bucket; use one "
            "of 'hidden', 'row', 'column' or 'page'.",
            context={"orientation": orientation.value, "reason": "invalid_orientation"},
        )
    return code


def _require_function_code(function: PivotAggregation | int | str) -> int:
    """Validate an aggregation and return its ``XlConsolidationFunction`` code.

    Only the eleven accepted codes are allowed. Legacy codes (``9``, ``-4114``,
    ``-4153``, ``-4154``) that the target build rejects are refused with the
    accepted set in the error context rather than silently written.
    """

    accepted = sorted(code for code in _XL_FUNCTION_BY_INT)
    if isinstance(function, PivotAggregation):
        return _XL_FUNCTION_CODES[function]
    if isinstance(function, str):
        try:
            return _XL_FUNCTION_CODES[PivotAggregation(function.strip().lower())]
        except ValueError as exc:
            raise InvalidRangeError(
                f"Aggregation {function!r} is not accepted; accepted values are "
                f"{[a.value for a in PivotAggregation]}.",
                context={
                    "function": function,
                    "accepted": [a.value for a in PivotAggregation],
                    "reason": "invalid_function",
                },
            ) from exc
    if isinstance(function, int) and not isinstance(function, bool):
        if function in _XL_FUNCTION_BY_INT:
            return function
        raise InvalidRangeError(
            f"Aggregation code {function} is not accepted by the target build; "
            f"accepted codes are {accepted}. Legacy codes (9, -4114, -4153, "
            "-4154) are rejected.",
            context={
                "function": function,
                "accepted": accepted,
                "reason": "invalid_function",
            },
        )
    raise InvalidRangeError(
        f"Aggregation {function!r} is not accepted; expected a PivotAggregation, "
        "its string name or one of the accepted codes.",
        context={"function": repr(function), "reason": "invalid_function"},
    )


def validate_field_properties(properties: PivotFieldProperties) -> list[str]:
    """Validate a sparse PivotField property patch before any COM mutation.

    Returns the sorted names of the properties the caller explicitly set. An
    empty patch, a blank caption / number format, or a non-boolean
    ``show_all_items`` leaves the workbook untouched (task_039 §5).
    """

    applied = sorted(properties.model_fields_set)
    unknown = set(applied) - _FIELD_PROPERTY_NAMES
    if unknown:
        raise InvalidRangeError(
            f"Unknown PivotField property/ies {sorted(unknown)}; supported: "
            f"{sorted(_FIELD_PROPERTY_NAMES)}.",
            context={"unknown": sorted(unknown), "reason": "invalid_property"},
        )
    if not applied:
        raise InvalidRangeError(
            "The PivotField property patch is empty; set at least one of "
            "'caption', 'number_format' or 'show_all_items'.",
            context={"reason": "empty_patch"},
        )
    if "caption" in applied:
        caption = properties.caption
        if not isinstance(caption, str) or not caption.strip():
            raise InvalidRangeError(
                "Argument 'caption' must be a non-empty string.",
                context={"property": "caption", "reason": "invalid_caption"},
            )
    if "number_format" in applied:
        number_format = properties.number_format
        if not isinstance(number_format, str) or not number_format.strip():
            raise InvalidRangeError(
                "Argument 'number_format' must be a non-empty format string.",
                context={"property": "number_format", "reason": "invalid_number_format"},
            )
    if "show_all_items" in applied and not isinstance(properties.show_all_items, bool):
        raise InvalidRangeError(
            "Argument 'show_all_items' must be a boolean.",
            context={"property": "show_all_items", "reason": "invalid_property"},
        )
    return applied


def _write_data_field_property(
    pivot: Any, position: int, attribute: str, value: Any, pivot_table: str
) -> None:
    """Write one property of a ``DataFields`` entry, raising a truthful error."""

    try:
        data_field = pivot.DataFields.Item(position)
        setattr(data_field, attribute, value)
    except ExcelVbaMcpError:
        raise
    except Exception as exc:  # noqa: BLE001 - a blocked/failed write
        raise ExcelVbaMcpError(
            f"DataFields[{position}].{attribute} could not be set: "
            f"{str(exc) or exc.__class__.__name__}",
            code=ErrorCode.COM_BUSY,
            context={
                "pivot_table": pivot_table,
                "data_field": position,
                "attribute": attribute,
            },
        ) from exc


class PivotTablesAdapter:
    """STA-local PivotTables / PivotCaches / PivotFields adapter (no COM at import).

    All methods are synchronous, operate on the live ``Workbook`` proxy and
    return only Pydantic DTOs / plain containers, so results cross the
    single-STA worker boundary safely. They never create a PivotTable or
    PivotCache (task_048 is read-only); task_049 adds the bounded targeted
    refresh of a worksheet/Table-backed PivotTable or its cache; task_050 adds
    the bounded configuration of an **existing** PivotField (layout, position,
    data-field aggregation and properties) with a pre-mutation source check and
    a post-mutation re-read. No ``.Activate`` / ``.Select`` is ever called.
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
    # Targeted refresh (task_049)
    # ------------------------------------------------------------------
    def refresh_pivot_table(self, wb: Any, pivot_table: str) -> PivotRefreshResult:
        """Refresh one PivotTable via ``PivotTable.RefreshTable`` (table-scoped).

        The cache source type is checked **before** invocation: an external, OLAP,
        Data Model or otherwise unsupported cache returns ``unsupported_source``
        and is never refreshed. A successful refresh is confirmed only by Excel's
        truth signal — the ``RefreshTable`` Boolean return and/or an advanced
        ``RefreshDate`` — never by the absence of a COM exception (task_039 §4).

        ``affected_pivot_tables`` always lists every table sharing the cache,
        because ``RefreshTable`` refreshes the shared cache.
        """

        pivot = _find_pivot(wb, pivot_table)
        name = _get_str(pivot, "Name") or pivot_table
        cache_index = _get_int(pivot, "CacheIndex", default=0) or 0
        inventory = self.enumerate_pivot_tables(wb)
        affected = _affected_pivot_tables(inventory, cache_index)
        before = _get_datetime(pivot, "RefreshDate")

        cache = pivot.PivotCache()
        source_type, _olap = _classify_source_type(cache)
        if source_type is not PivotSourceType.DATABASE:
            return _unsupported_result(
                pivot_table=name,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_TABLE,
                affected=affected,
                source_type=source_type,
                before=before,
            )

        started_at = datetime.now(UTC)
        refresh_return: bool | None
        try:
            raw = pivot.RefreshTable()
            refresh_return = bool(raw)
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - a blocked/failed refresh
            return _failed_result(
                pivot_table=name,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_TABLE,
                affected=affected,
                before=before,
                after=None,
                refresh_return=None,
                started_at=started_at,
                finished_at=datetime.now(UTC),
                reason=str(exc) or exc.__class__.__name__,
            )
        finished_at = datetime.now(UTC)
        after = _get_datetime(pivot, "RefreshDate")

        if refresh_return or _advance_confirms(before, after):
            return _refreshed_result(
                pivot_table=name,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_TABLE,
                affected=affected,
                before=before,
                after=after,
                refresh_return=refresh_return,
                started_at=started_at,
                finished_at=finished_at,
            )
        return _failed_result(
            pivot_table=name,
            cache_index=cache_index,
            method=PivotRefreshMethod.REFRESH_TABLE,
            affected=affected,
            before=before,
            after=after,
            refresh_return=refresh_return,
            started_at=started_at,
            finished_at=finished_at,
            reason=("RefreshTable returned no success signal and RefreshDate did not advance."),
        )

    def refresh_pivot_cache(self, wb: Any, cache_index: int) -> PivotRefreshResult:
        """Refresh one PivotCache via ``PivotCache.Refresh`` (cache-scoped).

        A cache is shared N:1 with PivotTables, so this refreshes every table on
        the cache; ``affected_pivot_tables`` reports them. The source type is
        checked before invocation and an unsupported cache returns
        ``unsupported_source``. ``PivotCache.Refresh`` returns no Boolean, so
        success is confirmed by the shared cache's advanced ``RefreshDate``
        (task_039 §4).

        ``pivot_table`` is the sorted-first affected table name (the cache has no
        single owning PivotTable).
        """

        inventory = self.enumerate_pivot_tables(wb)
        affected = _affected_pivot_tables(inventory, cache_index)
        if not affected:
            raise ExcelVbaMcpError(
                f"PivotCache {cache_index} was not found in the workbook.",
                context={"cache_index": cache_index},
            )
        pivot_table = affected[0]

        cache = _find_cache(wb, cache_index)
        source_type, _olap = _classify_source_type(cache)
        before = _get_datetime(cache, "RefreshDate")

        if source_type is not PivotSourceType.DATABASE:
            return _unsupported_result(
                pivot_table=pivot_table,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_CACHE,
                affected=affected,
                source_type=source_type,
                before=before,
            )

        started_at = datetime.now(UTC)
        try:
            cache.Refresh()
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - a blocked/failed refresh
            return _failed_result(
                pivot_table=pivot_table,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_CACHE,
                affected=affected,
                before=before,
                after=None,
                refresh_return=None,
                started_at=started_at,
                finished_at=datetime.now(UTC),
                reason=str(exc) or exc.__class__.__name__,
            )
        finished_at = datetime.now(UTC)
        after = _get_datetime(cache, "RefreshDate")

        if _advance_confirms(before, after):
            return _refreshed_result(
                pivot_table=pivot_table,
                cache_index=cache_index,
                method=PivotRefreshMethod.REFRESH_CACHE,
                affected=affected,
                before=before,
                after=after,
                refresh_return=None,
                started_at=started_at,
                finished_at=finished_at,
            )
        return _failed_result(
            pivot_table=pivot_table,
            cache_index=cache_index,
            method=PivotRefreshMethod.REFRESH_CACHE,
            affected=affected,
            before=before,
            after=after,
            refresh_return=None,
            started_at=started_at,
            finished_at=finished_at,
            reason="Refresh returned but the shared RefreshDate did not advance.",
        )

    # ------------------------------------------------------------------
    # PivotField configuration (task_050; guarded mutations)
    # ------------------------------------------------------------------
    def set_field_layout(
        self,
        wb: Any,
        pivot_table: str,
        field: str,
        orientation: PivotFieldOrientation,
        *,
        delete: bool = False,
    ) -> dict[str, Any]:
        """Set one field's layout bucket, or remove it, and verify the result.

        ``orientation`` must be one of the layout buckets (``hidden`` / ``row`` /
        ``column`` / ``page``): ``xlDataField`` (4) is rejected because assigning
        ``Orientation = 4`` never creates a data field — that is
        ``PivotTable.AddDataField`` via :meth:`set_field_aggregation`
        (task_039 §5).

        Removal is exposed two ways (task_039 §5):

        * ``orientation="hidden"`` and ``delete=False`` (default) is the
          **non-destructive** removal: ``PivotField.Orientation = 0`` leaves the
          field in ``PivotFields`` and its items and filter state intact.
        * ``delete=True`` (which requires ``orientation="hidden"``) is the
          **destructive** removal: ``PivotField.Delete`` drops the field from
          ``PivotFields``. The returned DTO then describes the last-known state
          (``orientation="hidden"``, ``position=None``); the field is verified
          absent from ``PivotFields``.

        The cache source type is checked before any mutation: an OLAP / Data
        Model / external cache is rejected with ``UNSUPPORTED_PIVOT_SOURCE``.
        The post-mutation state is re-read from the PivotTable as a
        postcondition; a mismatch is reported through ``verified=False`` and
        never claimed as success.
        """

        pivot = _find_pivot(wb, pivot_table)
        _require_supported_source(pivot)
        target, before = _resolve_source_field(pivot, field)
        code = _require_layout_code(orientation, delete=delete)

        if delete:
            try:
                target.Delete()
            except ExcelVbaMcpError:
                raise
            except Exception as exc:  # noqa: BLE001 - a blocked/failed delete
                raise ExcelVbaMcpError(
                    f"PivotField.Delete failed for field {field!r}: "
                    f"{str(exc) or exc.__class__.__name__}",
                    code=ErrorCode.COM_BUSY,
                    context={"pivot_table": pivot_table, "field": field},
                ) from exc
            data_field_index = _data_field_index(pivot)
            absent = _field_absent(pivot, field)
            notes: list[str] = []
            if not absent:
                notes.append(
                    "PivotField.Delete returned but the field is still present in "
                    "PivotFields; the removal is not confirmed."
                )
            removed = before.model_copy(
                update={"orientation": PivotFieldOrientation.HIDDEN, "position": None},
                deep=True,
            )
            return {
                "field": removed,
                "verified": absent,
                "notes": notes,
                "orientation_order": [],
                "data_field_index": data_field_index,
                "before": before,
            }

        try:
            target.Orientation = code
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - a blocked/failed write
            raise ExcelVbaMcpError(
                f"PivotField.Orientation could not be set for field {field!r}: "
                f"{str(exc) or exc.__class__.__name__}",
                code=ErrorCode.COM_BUSY,
                context={"pivot_table": pivot_table, "field": field},
            ) from exc

        data_field_index = _data_field_index(pivot)
        after = _field_dto(pivot, target, data_field_index, source_layout=True)
        if orientation is PivotFieldOrientation.HIDDEN:
            verified = _raw_position(target) is None
        else:
            verified = _raw_orientation(target) is orientation
        layout_notes: list[str] = []
        if not verified:
            layout_notes.append(
                "Postcondition mismatch: the re-read orientation differs from the "
                f"request ({orientation.value!r})."
            )
        return {
            "field": after,
            "verified": verified,
            "notes": layout_notes,
            "orientation_order": _orientation_order(pivot, orientation),
            "data_field_index": data_field_index,
            "before": before,
        }

    def set_field_position(
        self, wb: Any, pivot_table: str, field: str, position: int
    ) -> dict[str, Any]:
        """Set one field's 1-based position within its current orientation.

        The field must already be placed in a layout bucket (``row`` /
        ``column`` / ``page``) or be a data field; a hidden field has no
        position and is rejected before any COM mutation. ``position`` is
        validated against the target orientation's field count before the write,
        so an out-of-range position never reaches Excel.

        The post-mutation position is re-read and the resulting order of the
        whole orientation is returned (``orientation_order``) so the caller can
        verify the swap semantics (task_039 §5).
        """

        pivot = _find_pivot(wb, pivot_table)
        _require_supported_source(pivot)
        target, before = _resolve_source_field(pivot, field)
        orientation = before.orientation
        data_matches = [
            entry
            for entry in _data_field_entries(pivot)
            if entry[2].strip().lower() == before.source_name.strip().lower()
        ]

        if orientation is PivotFieldOrientation.HIDDEN:
            if not data_matches:
                raise ExcelVbaMcpError(
                    f"Field {field!r} is hidden and has no position to set; set its "
                    "orientation first.",
                    code=ErrorCode.INVALID_RANGE,
                    context={
                        "pivot_table": pivot_table,
                        "field": field,
                        "reason": "field_not_placed",
                    },
                )
            orientation = PivotFieldOrientation.DATA

        count = _orientation_count(pivot, orientation)
        if count == 0 or position > count:
            raise ExcelVbaMcpError(
                f"Position {position} is out of range for orientation "
                f"{orientation.value!r} ({count} field(s)).",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "pivot_table": pivot_table,
                    "field": field,
                    "orientation": orientation.value,
                    "position": position,
                    "count": count,
                    "reason": "position_out_of_range",
                },
            )

        try:
            if orientation is PivotFieldOrientation.DATA:
                pivot.DataFields.Item(data_matches[0][0]).Position = position
            else:
                target.Position = position
        except ExcelVbaMcpError:
            raise
        except Exception as exc:  # noqa: BLE001 - a blocked/failed write
            raise ExcelVbaMcpError(
                f"PivotField.Position could not be set for field {field!r}: "
                f"{str(exc) or exc.__class__.__name__}",
                code=ErrorCode.COM_BUSY,
                context={"pivot_table": pivot_table, "field": field},
            ) from exc

        after = _field_dto(
            pivot,
            target,
            _data_field_index(pivot),
            source_layout=orientation is not PivotFieldOrientation.DATA,
        )
        written = (
            pivot.DataFields.Item(data_matches[0][0])
            if orientation is PivotFieldOrientation.DATA
            else target
        )
        verified = _raw_position(written) == position
        notes: list[str] = []
        if not verified:
            notes.append(
                "Postcondition mismatch: the re-read position differs from the "
                f"request ({position})."
            )
        return {
            "field": after,
            "verified": verified,
            "notes": notes,
            "orientation_order": _orientation_order(pivot, orientation),
            "data_field_index": _data_field_index(pivot),
            "before": before,
        }

    def set_field_aggregation(
        self,
        wb: Any,
        pivot_table: str,
        field: str,
        function: PivotAggregation,
        *,
        caption: str | None = None,
    ) -> dict[str, Any]:
        """Add or update one data field's aggregation and verify the result.

        An existing data field on the same source field has its
        ``DataFields[i].Function`` updated in place. A source field without a
        data field is added with ``PivotTable.AddDataField`` (assigning
        ``Orientation = xlDataField`` does **not** create one — task_039 §5).

        ``caption`` identifies the data field when it is supplied; it must be
        unique among the PivotTable's data-field captions and must not equal a
        source field name, otherwise the request is rejected **before** COM
        (Excel itself rejects duplicate captions). When ``caption`` is omitted
        the single existing data field of that source is updated; zero or several
        are rejected as an ambiguous identity.

        Only the eleven accepted ``XlConsolidationFunction`` codes are accepted;
        legacy codes (``9``, ``-4114``, ``-4153``, ``-4154``) are rejected.
        """

        pivot = _find_pivot(wb, pivot_table)
        _require_supported_source(pivot)
        target, info = _resolve_source_field(pivot, field)
        code = _require_function_code(function)
        source_name = info.source_name

        entries = _data_field_entries(pivot)
        by_caption = {entry[1].strip().lower(): entry for entry in entries}
        by_source: dict[str, list[tuple[int, str, str]]] = {}
        for entry in entries:
            by_source.setdefault(entry[2].strip().lower(), []).append(entry)

        target_position: int
        if caption is not None:
            clean_caption = caption.strip()
            if not clean_caption:
                raise ExcelVbaMcpError(
                    "Argument 'caption' must be a non-empty string.",
                    code=ErrorCode.INVALID_RANGE,
                    context={"pivot_table": pivot_table, "reason": "empty_caption"},
                )
            existing = by_caption.get(clean_caption.lower())
            if existing is not None and existing[2].strip().lower() == source_name.strip().lower():
                target_position = existing[0]
                try:
                    pivot.DataFields.Item(target_position).Function = code
                except ExcelVbaMcpError:
                    raise
                except Exception as exc:  # noqa: BLE001 - a blocked/failed write
                    raise ExcelVbaMcpError(
                        f"The data field caption {clean_caption!r} could not be "
                        f"updated: {str(exc) or exc.__class__.__name__}",
                        code=ErrorCode.COM_BUSY,
                        context={"pivot_table": pivot_table, "field": field},
                    ) from exc
            else:
                source_names = {name.strip().lower() for _f, name in _source_field_entries(pivot)}
                if clean_caption.lower() in source_names:
                    raise ExcelVbaMcpError(
                        f"The data-field caption {clean_caption!r} duplicates a "
                        "source field name; Excel rejects that caption.",
                        code=ErrorCode.AMBIGUOUS_FILTER,
                        context={
                            "pivot_table": pivot_table,
                            "field": field,
                            "caption": clean_caption,
                            "reason": "duplicate_caption",
                        },
                    )
                try:
                    pivot.AddDataField(target, clean_caption, code)
                except ExcelVbaMcpError:
                    raise
                except Exception as exc:  # noqa: BLE001 - a blocked/failed add
                    raise ExcelVbaMcpError(
                        f"AddDataField failed for field {field!r} with caption "
                        f"{clean_caption!r}: {str(exc) or exc.__class__.__name__}",
                        code=ErrorCode.AMBIGUOUS_FILTER,
                        context={
                            "pivot_table": pivot_table,
                            "field": field,
                            "caption": clean_caption,
                            "reason": "duplicate_caption",
                        },
                    ) from exc
                new_entries = _data_field_entries(pivot)
                created = [
                    entry
                    for entry in new_entries
                    if entry[1].strip().lower() == clean_caption.lower()
                ]
                if not created:
                    raise ExcelVbaMcpError(
                        "AddDataField returned but no data field with the requested "
                        "caption exists; the aggregation is not confirmed.",
                        code=ErrorCode.AMBIGUOUS_FILTER,
                        context={
                            "pivot_table": pivot_table,
                            "field": field,
                            "caption": clean_caption,
                            "reason": "postcondition_mismatch",
                        },
                    )
                target_position = created[0][0]
        else:
            positions = by_source.get(source_name.strip().lower(), [])
            if len(positions) != 1:
                raise ExcelVbaMcpError(
                    f"Field {field!r} resolves to {len(positions)} data field(s); "
                    "supply 'caption' to identify the data field.",
                    code=ErrorCode.AMBIGUOUS_FILTER,
                    context={
                        "pivot_table": pivot_table,
                        "field": field,
                        "data_field_count": len(positions),
                        "reason": "ambiguous_data_field",
                    },
                )
            target_position = positions[0][0]
            try:
                pivot.DataFields.Item(target_position).Function = code
            except ExcelVbaMcpError:
                raise
            except Exception as exc:  # noqa: BLE001 - a blocked/failed write
                raise ExcelVbaMcpError(
                    f"The data field {target_position} Function could not be set: "
                    f"{str(exc) or exc.__class__.__name__}",
                    code=ErrorCode.COM_BUSY,
                    context={"pivot_table": pivot_table, "field": field},
                ) from exc

        final_entries = _data_field_entries(pivot)
        final = next((entry for entry in final_entries if entry[0] == target_position), None)
        verified = final is not None
        notes: list[str] = []
        if final is None:
            notes.append(
                "Postcondition mismatch: the data field disappeared after the aggregation change."
            )
        else:
            final_function = _get_int(pivot.DataFields.Item(target_position), "Function")
            if final_function != code:
                notes.append(
                    "Postcondition mismatch: the re-read Function differs from the "
                    "requested aggregation."
                )
        after = _describe_field(pivot, target, _data_field_index(pivot))
        return {
            "field": after,
            "verified": verified and not notes,
            "notes": notes,
            "orientation_order": _orientation_order(pivot, PivotFieldOrientation.DATA),
            "data_field_index": _data_field_index(pivot),
            "before": info,
        }

    def set_field_properties(
        self, wb: Any, pivot_table: str, field: str, properties: PivotFieldProperties
    ) -> dict[str, Any]:
        """Set bounded PivotField properties and verify the targeted ones.

        The sparse patch is validated before any COM mutation (an empty patch,
        a blank caption or number format, or a caption that duplicates another
        data-field caption leaves the workbook untouched). Only the properties
        the caller explicitly set are written:

        * ``caption`` / ``number_format`` target the **data field** (a source
          field without a data field is rejected — task_039 §5);
        * ``show_all_items`` targets the source ``PivotField`` itself.

        Each targeted property is re-read after the write; a mismatch is
        reported through ``verified=False`` and ``notes`` rather than claimed as
        success.
        """

        pivot = _find_pivot(wb, pivot_table)
        _require_supported_source(pivot)
        applied = validate_field_properties(properties)
        target, info = _resolve_field_target(pivot, field)

        needs_data_field = bool({"caption", "number_format"} & set(applied))
        data_position: int | None = None
        if needs_data_field:
            if not info.is_data_field:
                raise ExcelVbaMcpError(
                    f"Field {field!r} is not a data field; 'caption' and "
                    "'number_format' apply to data fields only (add it with "
                    "pivot_field_set_aggregation first).",
                    code=ErrorCode.INVALID_RANGE,
                    context={
                        "pivot_table": pivot_table,
                        "field": field,
                        "reason": "not_a_data_field",
                    },
                )
            matches = [
                entry
                for entry in _data_field_entries(pivot)
                if entry[2].strip().lower() == info.source_name.strip().lower()
            ]
            if len(matches) != 1:
                raise ExcelVbaMcpError(
                    f"Field {field!r} resolves to {len(matches)} data field(s); "
                    "identify the data field by its caption.",
                    code=ErrorCode.AMBIGUOUS_FILTER,
                    context={
                        "pivot_table": pivot_table,
                        "field": field,
                        "data_field_count": len(matches),
                        "reason": "ambiguous_data_field",
                    },
                )
            data_position = matches[0][0]

        if "caption" in applied:
            new_caption = str(properties.caption).strip()
            duplicate = any(
                entry[0] != data_position and entry[1].strip().lower() == new_caption.lower()
                for entry in _data_field_entries(pivot)
            )
            if duplicate:
                raise ExcelVbaMcpError(
                    f"The data-field caption {new_caption!r} duplicates another "
                    "data-field caption.",
                    code=ErrorCode.AMBIGUOUS_FILTER,
                    context={
                        "pivot_table": pivot_table,
                        "field": field,
                        "caption": new_caption,
                        "reason": "duplicate_caption",
                    },
                )

        if "caption" in applied and data_position is not None:
            _write_data_field_property(
                pivot, data_position, "Caption", str(properties.caption).strip(), pivot_table
            )
        if "number_format" in applied and data_position is not None:
            _write_data_field_property(
                pivot,
                data_position,
                "NumberFormat",
                str(properties.number_format),
                pivot_table,
            )
        if "show_all_items" in applied:
            try:
                target.ShowAllItems = bool(properties.show_all_items)
            except ExcelVbaMcpError:
                raise
            except Exception as exc:  # noqa: BLE001 - a blocked/failed write
                raise ExcelVbaMcpError(
                    f"PivotField.ShowAllItems could not be set for field {field!r}: "
                    f"{str(exc) or exc.__class__.__name__}",
                    code=ErrorCode.COM_BUSY,
                    context={"pivot_table": pivot_table, "field": field},
                ) from exc

        final_data_field_index = _data_field_index(pivot)
        after = _describe_field(pivot, target, final_data_field_index)
        notes: list[str] = []
        if "caption" in applied and after.data_field_name != str(properties.caption).strip():
            notes.append(
                "Postcondition mismatch: the re-read data-field caption differs from the request."
            )
        if "number_format" in applied and after.number_format != str(properties.number_format):
            notes.append(
                "Postcondition mismatch: the re-read number format differs from the request."
            )
        if "show_all_items" in applied and after.show_all_items != bool(properties.show_all_items):
            notes.append(
                "Postcondition mismatch: the re-read ShowAllItems differs from the request."
            )
        return {
            "field": after,
            "verified": not notes,
            "notes": notes,
            "orientation_order": _orientation_order(pivot, after.orientation),
            "data_field_index": final_data_field_index,
            "before": info,
        }

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
