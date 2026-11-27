"""Streaming sharedStrings reader for OOXML snapshot files (WP-E1).

``xl/sharedStrings.xml`` holds the deduplicated string table an ``.xlsx`` /
``.xlsm`` worksheet references by index (``<c t="s"><v>0</v></c>``). The table
can be very large, so it is streamed with
:func:`xml.etree.ElementTree.iterparse` (clearing each ``<si>`` after it is
resolved) and bounded by ``limits.max_shared_strings`` rather than loaded as a
DOM. The part bytes are first read through the bounded
:func:`~excel_vba_mcp.adapters.ooxml.zip_safety.read_part`, so the uncompressed
ceiling is enforced before parsing.
"""

from __future__ import annotations

import io
import xml.etree.ElementTree as ET
import zipfile

from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    read_part,
)

#: The shared-strings part path inside an OOXML package.
_SHARED_STRINGS_PART = "xl/sharedStrings.xml"


def _local_name(tag: object) -> str:
    """Return the local tag name without an ``{namespace}`` prefix."""

    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _si_text(si_elem: ET.Element) -> str:
    """Concatenate every ``<t>`` descendant text of an ``<si>`` (rich text aware).

    A shared-string item is either a single ``<t>`` or a run of rich-text
    ``<r><t>`` fragments; both are flattened to one plain string.
    """

    parts: list[str] = []
    for elem in si_elem.iter():
        if _local_name(elem.tag) == "t":
            parts.append(elem.text or "")
    return "".join(parts)


def load_shared_strings(zf: zipfile.ZipFile, limits: ZipLimits) -> list[str]:
    """Stream and return the shared-string table (bounded).

    Args:
        zf: An already safety-vetted archive (from
            :func:`~excel_vba_mcp.adapters.ooxml.zip_safety.open_safe`).
        limits: The bounded ZIP/XML limits.

    Returns:
        The ordered list of shared strings; empty when the part is absent.

    Raises:
        OoxmlSafetyError: If the entry count exceeds ``limits.max_shared_strings``
            or the part cannot be read.
    """

    if _SHARED_STRINGS_PART not in zf.namelist():
        return []

    data = read_part(zf, _SHARED_STRINGS_PART, limits)
    strings: list[str] = []
    # Only ``end`` events are needed; each ``<si>`` is resolved and cleared so at
    # most one item's subtree is retained at a time (plus the bounded result).
    context = ET.iterparse(io.BytesIO(data), events=("end",))
    for _event, elem in context:
        if _local_name(elem.tag) != "si":
            continue
        strings.append(_si_text(elem))
        if len(strings) > limits.max_shared_strings:
            raise OoxmlSafetyError(
                f"Shared-string count exceeds the limit "
                f"({limits.max_shared_strings})."
            )
        elem.clear()
    return strings


__all__ = ["load_shared_strings"]