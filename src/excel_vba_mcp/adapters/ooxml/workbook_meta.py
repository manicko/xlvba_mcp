"""Workbook metadata reader for OOXML snapshot files (WP-E1).

Parses ``xl/workbook.xml`` and its relationship part
``xl/_rels/workbook.xml.rels`` to map each sheet name to its worksheet part
(e.g. ``xl/worksheets/sheet3.xml``), preserving sheet order and hidden state,
and collects the workbook defined names plus the table parts. Only metadata is
read (never cell data), and every part is read through the bounded
:func:`~excel_vba_mcp.adapters.ooxml.zip_safety.read_part`.

The returned ``dict`` is JSON-friendly and consumed by the application service to
build :class:`~excel_vba_mcp.domain.models.OoxmlMetadata` and to resolve a sheet
name to its part for row streaming.
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile
from typing import Any

from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    read_part,
)

_WORKBOOK_PART = "xl/workbook.xml"
_WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels"

# Office relationship namespaces (the ``r:id`` attribute and the rels root).
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

#: Relationship ``Type`` suffix that identifies a PivotCache records part.
_PIVOT_RECORDS_REL_SUFFIX = "/pivotCacheRecords"


def _local_name(tag: object) -> str:
    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _rel_id_attr(elem: ET.Element) -> str | None:
    """Return the ``r:id`` relationship id of a ``<sheet>`` element."""

    for key, value in elem.attrib.items():
        if key.startswith(f"{{{_REL_NS}}}") and _local_name(key) == "id":
            return value
    return None


def _parse_rels(zf: zipfile.ZipFile, limits: ZipLimits) -> dict[str, str]:
    """Parse the workbook rels into ``{Id: resolved_part_path}``.

    Targets are resolved relative to the ``xl/`` package directory (the rels part
    lives at ``xl/_rels/workbook.xml.rels``); an absolute target (``/xl/...``) is
    taken as package-absolute.
    """

    return _parse_rels_at(zf, _WORKBOOK_RELS_PART, "xl", limits)


def _parse_rels_at(
    zf: zipfile.ZipFile, rels_part: str, base_dir: str, limits: ZipLimits
) -> dict[str, str]:
    """Parse a rels part into ``{Id: resolved_part_path}`` relative to ``base_dir``.

    ``base_dir`` is the package directory the relationship targets are relative to
    (e.g. ``"xl"`` for the workbook rels, ``"xl/pivotCache"`` for a cache
    definition's rels). An absolute target (``/xl/...``) is package-absolute. A
    missing rels part yields an empty mapping.
    """

    if rels_part not in zf.namelist():
        return {}
    data = read_part(zf, rels_part, limits)
    rels: dict[str, str] = {}
    root = ET.fromstring(data)
    for rel in root:
        if _local_name(rel.tag) != "Relationship":
            continue
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        rels[rid] = _resolve_target(target, base_dir)
    return rels


def _resolve_target(target: str, base_dir: str = "xl") -> str:
    """Resolve a relationship target to a package part path relative to ``base_dir``."""

    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def workbook_parts(zf: zipfile.ZipFile, limits: ZipLimits) -> dict[str, Any]:
    """Return workbook structural metadata (sheets, names, table parts).

    Args:
        zf: An already safety-vetted archive.
        limits: The bounded ZIP/XML limits.

    Returns:
        A dict with:

        * ``sheets``: ``list[dict]`` each ``{name, part, hidden, order}`` in
          workbook order (``order`` is 1-based);
        * ``defined_names``: ``list[str]`` of defined-name identifiers;
        * ``table_parts``: ``list[str]`` of ``xl/tables/*.xml`` part paths.

    Raises:
        OoxmlSafetyError: If ``xl/workbook.xml`` is missing or unreadable.
    """

    data = read_part(zf, _WORKBOOK_PART, limits)
    rels = _parse_rels(zf, limits)
    root = ET.fromstring(data)

    sheets: list[dict[str, Any]] = []
    defined_names: list[str] = []

    for elem in root:
        tag = _local_name(elem.tag)
        if tag == "sheets":
            for order, sheet in enumerate(elem, start=1):
                if _local_name(sheet.tag) != "sheet":
                    continue
                name = sheet.get("name") or ""
                rid = _rel_id_attr(sheet)
                part = rels.get(rid, "") if rid else ""
                state = (sheet.get("state") or "visible").lower()
                sheets.append(
                    {
                        "name": name,
                        "part": part,
                        "hidden": state != "visible",
                        "order": order,
                    }
                )
        elif tag == "definedNames":
            for dn in elem:
                if _local_name(dn.tag) == "definedName":
                    identifier = dn.get("name")
                    if identifier:
                        defined_names.append(identifier)

    table_parts = sorted(
        name
        for name in zf.namelist()
        if name.startswith("xl/tables/") and name.endswith(".xml")
    )

    return {
        "sheets": sheets,
        "defined_names": defined_names,
        "table_parts": table_parts,
    }


def pivot_caches(zf: zipfile.ZipFile, limits: ZipLimits) -> list[dict[str, Any]]:
    """Map each workbook PivotCache to its definition and records parts (WP-E2).

    Parses ``xl/workbook.xml`` ``<pivotCaches>`` (1-based ``cache_index`` in
    document order) and the workbook rels to resolve each ``pivotCacheDefinition``
    part (e.g. ``xl/pivotCache/pivotCacheDefinition1.xml``). For each definition
    the cache's own rels (``xl/pivotCache/_rels/<defname>.rels``) are read to find
    the optional ``pivotCacheRecords`` part. Only metadata is read (never record
    data), and every part is read through the bounded
    :func:`~excel_vba_mcp.adapters.ooxml.zip_safety.read_part`.

    Returns:
        A list of ``{cache_index, definition_part, records_part}`` (1-based
        ``cache_index``); ``records_part`` is ``None`` when the cache saved no
        local records (e.g. OLAP / Data Model / ``saveData="0"``).

    Raises:
        OoxmlSafetyError: If ``xl/workbook.xml`` is missing or unreadable.
    """

    data = read_part(zf, _WORKBOOK_PART, limits)
    rels = _parse_rels(zf, limits)
    root = ET.fromstring(data)

    definition_parts: list[str] = []
    for elem in root:
        if _local_name(elem.tag) != "pivotCaches":
            continue
        for cache in elem:
            if _local_name(cache.tag) != "pivotCache":
                continue
            rid = _rel_id_attr(cache)
            definition_parts.append(rels.get(rid, "") if rid else "")

    caches: list[dict[str, Any]] = []
    for index, definition_part in enumerate(definition_parts, start=1):
        records_part = _records_part_for(zf, definition_part, limits)
        caches.append(
            {
                "cache_index": index,
                "definition_part": definition_part,
                "records_part": records_part,
            }
        )
    return caches


def _records_part_for(
    zf: zipfile.ZipFile, definition_part: str, limits: ZipLimits
) -> str | None:
    """Resolve a cache definition's ``pivotCacheRecords`` part, or ``None``.

    The definition's rels live at ``xl/pivotCache/_rels/<defname>.rels`` and their
    targets are relative to the definition's directory (``xl/pivotCache``). The
    records relationship is identified by its ``Type`` suffix
    (``/pivotCacheRecords``), never by target name.
    """

    if not definition_part:
        return None
    directory = posixpath.dirname(definition_part)
    defname = posixpath.basename(definition_part)
    rels_part = posixpath.join(directory, "_rels", f"{defname}.rels")
    if rels_part not in zf.namelist():
        return None
    data = read_part(zf, rels_part, limits)
    root = ET.fromstring(data)
    for rel in root:
        if _local_name(rel.tag) != "Relationship":
            continue
        rel_type = rel.get("Type") or ""
        target = rel.get("Target")
        if target and rel_type.endswith(_PIVOT_RECORDS_REL_SUFFIX):
            return _resolve_target(target, directory)
    return None


def sheet_part(parts: dict[str, Any], sheet_name: str) -> str:
    """Resolve a sheet name to its worksheet part path (case-insensitive).

    Raises:
        OoxmlSafetyError: If no sheet with that name exists or it has no part.
    """

    sheets = parts.get("sheets", [])
    for entry in sheets:
        if entry["name"] == sheet_name:
            return _require_part(entry, sheet_name)
    lowered = sheet_name.lower()
    for entry in sheets:
        if entry["name"].lower() == lowered:
            return _require_part(entry, sheet_name)
    raise OoxmlSafetyError(f"Sheet not found in workbook metadata: {sheet_name!r}.")


def _require_part(entry: dict[str, Any], sheet_name: str) -> str:
    part = entry.get("part") or ""
    if not part:
        raise OoxmlSafetyError(
            f"Sheet {sheet_name!r} has no resolvable worksheet part."
        )
    return part


__all__ = ["pivot_caches", "sheet_part", "workbook_parts"]