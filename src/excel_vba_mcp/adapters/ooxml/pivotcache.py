"""Streaming PivotCache reader for OOXML snapshot files (WP-E2).

Reads the **saved** PivotCache of an ``.xlsx`` / ``.xlsm`` snapshot without
occupying the COM worker and without materializing the whole cache in memory:

* :func:`read_cache_definition` parses a ``pivotCacheDefinition`` part (bounded)
  for ``recordCount`` / ``saveData`` / ``refreshOnLoad`` / ``refreshDate`` and the
  per-field shared-items table needed to resolve ``<s v="i"/>`` record indices.
* :func:`iter_cache_records` streams a ``pivotCacheRecords`` part with
  :func:`xml.etree.ElementTree.iterparse`, building a bounded window of typed rows
  and stopping as soon as ``limit`` rows are collected.

The cache records are the data Excel **saved** with the workbook — they are not the
PivotTable's displayed result and not the live source. A cache with no local
records part (OLAP / Data Model / ``saveData="0"``) is reported as an unsupported
source by the service, never fabricated. The part bytes are first read through the
bounded :func:`~excel_vba_mcp.adapters.ooxml.zip_safety.read_part`, so the
uncompressed ceiling is enforced before parsing.
"""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime

from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    read_part,
)
from excel_vba_mcp.domain.models import CellValue
from excel_vba_mcp.domain.statuses import CellValueKind

#: Record / shared-item value child tags (``<n>``/``<s>``/``<d>``/``<m/>``/...).
_VALUE_TAGS = frozenset({"n", "s", "d", "m", "e", "b"})


def _local_name(tag: object) -> str:
    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _parse_int(raw: str) -> int | float:
    """Parse a numeric payload as ``int`` when integral, else ``float``."""

    try:
        return int(raw)
    except ValueError:
        return float(raw)


def _bool_attr(raw: str | None, *, default: bool | None) -> bool | None:
    """Parse an OOXML boolean attribute (``"1"``/``"true"`` vs ``"0"``/``"false"``)."""

    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true"}


def _datetime_attr(raw: str | None) -> datetime | None:
    """Parse an ISO-8601 ``refreshDate`` attribute, or ``None`` when absent/invalid."""

    if not raw:
        return None
    text = raw.strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def _int_or_none(raw: str | None) -> int | None:
    """Parse an integer attribute, or ``None`` when absent/empty/non-numeric."""

    if raw is None or raw == "":
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _typed_value(tag: str, raw: str | None) -> CellValue:
    """Build a typed :class:`CellValue` from a non-shared value element.

    ``m`` (missing) and any unknown tag map to :attr:`CellValueKind.EMPTY`; ``n``
    parses to :attr:`~CellValueKind.NUMBER` (falling back to ``STRING`` on a
    non-numeric payload); ``d``/``b``/``e`` map to date/boolean/error.
    """

    if tag == "n":
        if raw is None or raw == "":
            return CellValue(kind=CellValueKind.EMPTY)
        try:
            return CellValue(kind=CellValueKind.NUMBER, value=_parse_int(raw))
        except ValueError:
            return CellValue(kind=CellValueKind.STRING, value=raw)
    if tag == "d":
        return CellValue(kind=CellValueKind.DATE, value=raw or "")
    if tag == "b":
        return CellValue(kind=CellValueKind.BOOLEAN, value=(raw == "1"))
    if tag == "e":
        return CellValue(kind=CellValueKind.ERROR, value=raw or "")
    # ``m`` (missing) and any unknown tag.
    return CellValue(kind=CellValueKind.EMPTY)


def _shared_item_value(tag: str, raw: str | None) -> CellValue:
    """Build a typed :class:`CellValue` for one ``<sharedItems>`` entry.

    A shared ``<s v="text"/>`` carries the literal string (not an index), so it is
    resolved directly; every other tag defers to :func:`_typed_value`.
    """

    if tag == "s":
        return CellValue(kind=CellValueKind.STRING, value=raw or "")
    return _typed_value(tag, raw)


def read_cache_definition(
    zf: zipfile.ZipFile, definition_part: str, limits: ZipLimits
) -> dict[str, object]:
    """Parse a ``pivotCacheDefinition`` part for freshness metadata + shared items.

    Args:
        zf: An already safety-vetted archive.
        definition_part: The cache definition part (e.g.
            ``xl/pivotCache/pivotCacheDefinition1.xml``).
        limits: The bounded ZIP/XML limits.

    Returns:
        A dict with:

        * ``record_count``: ``recordCount`` attribute (``int`` or ``None``);
        * ``save_data``: ``saveData`` attribute (default ``True``);
        * ``refresh_on_load``: ``refreshOnLoad`` attribute (``bool`` or ``None``);
        * ``refresh_date``: parsed ``refreshDate`` (``datetime`` or ``None``);
        * ``shared_items``: ``list[list[CellValue]]`` — the per-field shared-items
          table (index-aligned by field position) used to resolve ``<s v="i"/>``.

    Raises:
        OoxmlSafetyError: If the part is missing, exceeds the XML depth limit, or
            the shared-item count exceeds ``limits.max_shared_strings``.
    """

    data = read_part(zf, definition_part, limits)

    record_count: int | None = None
    save_data: bool = True
    refresh_on_load: bool | None = None
    refresh_date: datetime | None = None
    shared_items: list[list[CellValue]] = []

    current_shared: list[CellValue] | None = None
    in_shared = False
    total_shared = 0
    root_seen = False
    depth = 0

    context = ET.iterparse(io.BytesIO(data), events=("start", "end"))
    for event, elem in context:
        tag = _local_name(elem.tag)
        if event == "start":
            depth += 1
            if depth > limits.max_xml_depth:
                raise OoxmlSafetyError(
                    f"PivotCache definition XML nesting depth exceeds the limit "
                    f"({limits.max_xml_depth})."
                )
            if not root_seen:
                root_seen = True
                record_count = _int_or_none(elem.get("recordCount"))
                save_data = bool(_bool_attr(elem.get("saveData"), default=True))
                refresh_on_load = _bool_attr(elem.get("refreshOnLoad"), default=None)
                refresh_date = _datetime_attr(elem.get("refreshDate"))
            elif tag == "cacheField":
                current_shared = []
            elif tag == "sharedItems":
                in_shared = True
            continue

        # ``end`` events.
        if tag == "sharedItems":
            in_shared = False
        elif in_shared and current_shared is not None and tag in _VALUE_TAGS:
            current_shared.append(_shared_item_value(tag, elem.get("v")))
            total_shared += 1
            if total_shared > limits.max_shared_strings:
                raise OoxmlSafetyError(
                    f"PivotCache shared-item count exceeds the limit "
                    f"({limits.max_shared_strings})."
                )
        elif tag == "cacheField" and current_shared is not None:
            shared_items.append(current_shared)
            current_shared = None

        if tag != "pivotCacheDefinition":
            elem.clear()
        depth -= 1

    return {
        "record_count": record_count,
        "save_data": save_data,
        "refresh_on_load": refresh_on_load,
        "refresh_date": refresh_date,
        "shared_items": shared_items,
    }


def _record_cell(
    elem: ET.Element, tag: str, shared_items: list[list[CellValue]], field_index: int
) -> CellValue:
    """Build a typed cell for one record child, resolving ``<s>`` shared indices."""

    raw = elem.get("v")
    if tag == "s":
        if raw is None or raw == "":
            return CellValue(kind=CellValueKind.STRING, value="")
        try:
            index = int(raw)
        except ValueError:
            return CellValue(kind=CellValueKind.STRING, value=raw)
        if 0 <= field_index < len(shared_items):
            field_shared = shared_items[field_index]
            if 0 <= index < len(field_shared):
                return field_shared[index]
        # Unresolved shared index: report the raw index string (never fabricate).
        return CellValue(kind=CellValueKind.STRING, value=raw)
    return _typed_value(tag, raw)


def _build_record_row(
    r_elem: ET.Element, shared_items: list[list[CellValue]], max_cells: int
) -> list[CellValue]:
    """Flatten a ``<r>`` record element to a field-aligned list of typed cells."""

    cells: list[CellValue] = []
    field_index = 0
    for child in r_elem:
        tag = _local_name(child.tag)
        if tag not in _VALUE_TAGS:
            continue
        if field_index < max_cells:
            cells.append(_record_cell(child, tag, shared_items, field_index))
        field_index += 1
    return cells


def iter_cache_records(
    zf: zipfile.ZipFile,
    records_part: str,
    shared_items: list[list[CellValue]],
    limits: ZipLimits,
    offset: int,
    limit: int,
) -> tuple[list[list[CellValue]], bool, int | None, int]:
    """Stream a bounded window of typed records from ``records_part``.

    Records before ``offset`` are skipped; up to ``limit`` records are collected,
    then parsing stops early (the remainder of the part is not parsed). A depth
    guard enforces ``limits.max_xml_depth`` and a per-row cell cap
    (``max_response_cells // limit``) bounds the response defensively.

    Args:
        zf: An already safety-vetted archive.
        records_part: The ``pivotCacheRecords`` part path.
        shared_items: The per-field shared-items table (from
            :func:`read_cache_definition`).
        limits: The bounded ZIP/XML limits.
        offset: 0-based record index of the first record to return.
        limit: Maximum number of records to collect (``>= 1``).

    Returns:
        ``(rows, truncated, total_record_count, scanned)``. ``truncated`` is
        ``True`` when collection stopped at ``limit`` (so
        ``total_record_count`` is ``None``); otherwise ``total_record_count`` is
        the exact number of records streamed. ``scanned`` is the number of
        records inspected (whether or not returned).

    Raises:
        OoxmlSafetyError: If the part is missing, exceeds the XML depth limit, or
            cannot be read.
    """

    data = read_part(zf, records_part, limits)
    max_cells = max(1, limits.max_response_cells // max(1, limit))

    collected: list[list[CellValue]] = []
    row_index = 0
    stopped_early = False
    depth = 0

    context = ET.iterparse(io.BytesIO(data), events=("start", "end"))
    for event, elem in context:
        if event == "start":
            depth += 1
            if depth > limits.max_xml_depth:
                raise OoxmlSafetyError(
                    f"PivotCache records XML nesting depth exceeds the limit "
                    f"({limits.max_xml_depth})."
                )
            continue

        if _local_name(elem.tag) == "r":
            row_index += 1
            if row_index > offset:
                collected.append(_build_record_row(elem, shared_items, max_cells))
                if len(collected) >= limit:
                    stopped_early = True
                    elem.clear()
                    break
            elem.clear()

        depth -= 1

    scanned = row_index
    if stopped_early:
        return collected, True, None, scanned
    return collected, False, scanned, scanned


__all__ = ["iter_cache_records", "read_cache_definition"]