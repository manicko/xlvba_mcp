"""Streaming worksheet row reader for OOXML snapshot files (WP-E1).

Reads a bounded window of typed rows from a worksheet part with
:func:`xml.etree.ElementTree.iterparse`, stopping as soon as ``limit`` rows are
collected (the rest of the part is never parsed). Each ``<row>`` is flattened to
a list of :class:`~excel_vba_mcp.domain.models.CellValue` aligned by column index
(the ``r`` cell reference, or sequential when absent). A depth guard enforces
``limits.max_xml_depth`` and a per-row cell cap (``max_response_cells // limit``)
bounds the response defensively. The part bytes are read through the bounded
:func:`~excel_vba_mcp.adapters.ooxml.zip_safety.read_part` first.
"""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

from excel_vba_mcp.adapters.excel.cell_values import bound_value_grid
from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    read_part,
)
from excel_vba_mcp.domain.models import CellValue
from excel_vba_mcp.domain.statuses import CellValueKind


def _local_name(tag: object) -> str:
    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _col_index(ref: str) -> int:
    """Return the 0-based column index from an A1 cell reference (``"AA10"`` -> 26)."""

    index = 0
    for char in ref:
        if char.isalpha():
            index = index * 26 + (ord(char.upper()) - ord("A") + 1)
        else:
            break
    return index - 1


def _find_child(elem: ET.Element, local: str) -> ET.Element | None:
    for child in elem:
        if _local_name(child.tag) == local:
            return child
    return None


def _descendant_text(elem: ET.Element, local: str) -> str:
    parts: list[str] = []
    for node in elem.iter():
        if _local_name(node.tag) == local:
            parts.append(node.text or "")
    return "".join(parts)


def _parse_number(raw: str) -> int | float:
    """Parse a numeric ``<v>`` payload as ``int`` when integral, else ``float``."""

    try:
        return int(raw)
    except ValueError:
        return float(raw)


def _cell_value(cell: ET.Element, shared: list[str]) -> CellValue:
    """Build a typed :class:`CellValue` from a ``<c>`` element."""

    cell_type = cell.get("t")

    if cell_type == "inlineStr":
        inline = _find_child(cell, "is")
        text = _descendant_text(inline, "t") if inline is not None else ""
        return CellValue(kind=CellValueKind.STRING, value=text)

    value_elem = _find_child(cell, "v")
    raw = value_elem.text if value_elem is not None else None

    if cell_type == "s":
        if raw is None or raw == "":
            return CellValue(kind=CellValueKind.EMPTY)
        try:
            index = int(raw)
        except ValueError:
            return CellValue(kind=CellValueKind.STRING, value=raw)
        value = shared[index] if 0 <= index < len(shared) else ""
        return CellValue(kind=CellValueKind.STRING, value=value)
    if cell_type == "str":
        return CellValue(kind=CellValueKind.STRING, value=raw or "")
    if cell_type == "b":
        return CellValue(kind=CellValueKind.BOOLEAN, value=(raw == "1"))
    if cell_type == "e":
        return CellValue(kind=CellValueKind.ERROR, value=raw or "")
    if cell_type == "d":
        return CellValue(kind=CellValueKind.DATE, value=raw or "")

    # Numeric (no type or ``t="n"``).
    if raw is None or raw == "":
        return CellValue(kind=CellValueKind.EMPTY)
    try:
        return CellValue(kind=CellValueKind.NUMBER, value=_parse_number(raw))
    except ValueError:
        return CellValue(kind=CellValueKind.STRING, value=raw)


def _build_row(row_elem: ET.Element, shared: list[str], max_cells: int) -> list[CellValue]:
    """Flatten a ``<row>`` element to a column-aligned list of typed cells."""

    cells: list[CellValue] = []
    next_col = 0
    for cell in row_elem:
        if _local_name(cell.tag) != "c":
            continue
        ref = cell.get("r")
        col = _col_index(ref) if ref else next_col
        next_col = col + 1
        if col < 0 or col >= max_cells:
            continue  # defensive per-row cell cap
        while len(cells) <= col:
            cells.append(CellValue(kind=CellValueKind.EMPTY))
        cells[col] = _cell_value(cell, shared)

    if len(cells) > max_cells:
        del cells[max_cells:]
    return cells


def iter_sheet_rows(
    zf: zipfile.ZipFile,
    sheet_part: str,
    shared: list[str],
    limits: ZipLimits,
    offset: int,
    limit: int,
) -> tuple[list[list[CellValue]], bool, int | None]:
    """Stream a bounded window of typed rows from ``sheet_part``.

    Rows before ``offset`` are skipped; up to ``limit`` rows are collected, then
    parsing stops early (the remainder of the part is not parsed).

    Args:
        zf: An already safety-vetted archive.
        sheet_part: The worksheet part path (e.g. ``xl/worksheets/sheet1.xml``).
        shared: The shared-string table (from
            :func:`~excel_vba_mcp.adapters.ooxml.shared_strings.load_shared_strings`).
        limits: The bounded ZIP/XML limits.
        offset: 0-based row index of the first row to return.
        limit: Maximum number of rows to collect (``>= 1``).

    Returns:
        ``(rows, truncated, total_rows_scanned)``. ``truncated`` is ``True`` when
        collection stopped at ``limit`` (so ``total_rows_scanned`` is ``None``);
        otherwise ``total_rows_scanned`` is the exact number of rows in the sheet.

    Raises:
        OoxmlSafetyError: If the part is missing, exceeds the XML depth limit, or
            cannot be read.
    """

    data = read_part(zf, sheet_part, limits)
    max_cells = max(1, limits.max_response_cells // max(1, limit))

    collected: list[list[CellValue]] = []
    rows_seen = 0
    row_index = 0
    stopped_early = False
    depth = 0

    context = ET.iterparse(io.BytesIO(data), events=("start", "end"))
    for event, elem in context:
        if event == "start":
            depth += 1
            if depth > limits.max_xml_depth:
                raise OoxmlSafetyError(
                    f"Worksheet XML nesting depth exceeds the limit "
                    f"({limits.max_xml_depth})."
                )
            continue

        if _local_name(elem.tag) == "row":
            if row_index >= offset:
                collected.append(_build_row(elem, shared, max_cells))
                if len(collected) >= limit:
                    stopped_early = True
                    elem.clear()
                    break
            row_index += 1
            rows_seen = row_index
            elem.clear()

        depth -= 1

    # Enforce the per-field text cap and the page byte budget at the
    # serialization boundary (contract §6). A page cut here is reported as
    # truncated with an unknown total so the caller pages on from
    # ``offset + returned`` and no row is skipped.
    collected, text_truncated, bytes_truncated = bound_value_grid(
        collected,
        max_text_len=limits.max_text_len,
        max_response_bytes=limits.max_response_bytes,
    )
    if stopped_early or text_truncated or bytes_truncated:
        return collected, True, None
    return collected, False, rows_seen


__all__ = ["iter_sheet_rows"]