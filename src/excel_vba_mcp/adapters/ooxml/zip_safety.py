"""Bounded, defensive ZIP access for OOXML snapshot files (WP-E1).

An OOXML ``.xlsx`` / ``.xlsm`` file is a ZIP archive. Before any part is read
this module enforces the unified large-data limits (member count, total
uncompressed size, per-member compression ratio) and rejects path-traversal
member names, so a hostile or accidental "ZIP bomb" cannot exhaust memory or
escape the snapshot store. ``extractall`` is never used.

All functions are pure filesystem / stdlib (``zipfile``) and run off the event
loop (the service dispatches them via ``asyncio.to_thread``). They raise
:class:`OoxmlSafetyError` (a plain exception) on any violation; the application
service maps it to a structured :class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError`.
"""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass

from excel_vba_mcp.config import LargeDataConfig

#: One mebibyte in bytes (limits are expressed in MiB).
_MIB = 1024 * 1024

#: Members smaller than this (uncompressed) skip the compression-ratio check so
#: tiny, legitimately-compressible parts do not trip a false positive. They are
#: still counted toward the member count and total uncompressed size.
_RATIO_SKIP_BYTES = 1024


class OoxmlSafetyError(Exception):
    """A ZIP-safety violation (unsafe archive); mapped to ``INVALID_RANGE``."""


@dataclass(frozen=True)
class ZipLimits:
    """Bounded ZIP limits derived from :class:`LargeDataConfig` (WP-E1).

    A frozen snapshot of the relevant ``LargeDataConfig`` fields so the parsing
    helpers do not depend on the live config object.
    """

    max_file_mb: int
    max_uncompressed_mb: int
    max_compression_ratio: int
    max_zip_members: int
    max_shared_strings: int
    max_xml_depth: int
    max_response_cells: int
    max_text_len: int = 32_767
    max_response_bytes: int = 8 * 1024 * 1024

    @classmethod
    def from_config(cls, config: LargeDataConfig) -> ZipLimits:
        """Build :class:`ZipLimits` from the unified large-data configuration."""

        return cls(
            max_file_mb=config.max_file_mb,
            max_uncompressed_mb=config.max_uncompressed_mb,
            max_compression_ratio=config.max_compression_ratio,
            max_zip_members=config.max_zip_members,
            max_shared_strings=config.max_shared_strings,
            max_xml_depth=config.max_xml_depth,
            max_response_cells=config.max_response_cells,
            max_text_len=config.max_text_len,
            max_response_bytes=config.max_response_bytes,
        )

    @property
    def max_uncompressed_bytes(self) -> int:
        """Total uncompressed size ceiling in bytes."""

        return self.max_uncompressed_mb * _MIB

    @property
    def max_file_bytes(self) -> int:
        """Input file size ceiling in bytes (``max_file_mb`` converted MiB -> bytes)."""

        return self.max_file_mb * _MIB


def _is_unsafe_member_name(name: str) -> bool:
    """Return whether a ZIP member name is a path-traversal / absolute path.

    OOXML parts are always relative, forward-slash paths (e.g.
    ``xl/worksheets/sheet1.xml``). A backslash, a leading slash, a drive-letter
    prefix (``C:``) or a ``..`` path segment is rejected.
    """

    if not name:
        return True
    if "\\" in name:
        return True
    if name.startswith("/"):
        return True
    if len(name) >= 2 and name[1] == ":":
        return True
    return ".." in name.split("/")


def open_safe(path: str, limits: ZipLimits) -> zipfile.ZipFile:
    """Open ``path`` as a :class:`zipfile.ZipFile` after enforcing ZIP limits.

    The archive's central directory is inspected **before** any part is read:

    * the member count must not exceed ``limits.max_zip_members``;
    * the total uncompressed size must not exceed ``limits.max_uncompressed_mb``;
    * no member name may be a traversal / absolute path;
    * each member's compression ratio (``file_size / compress_size``) must not
      exceed ``limits.max_compression_ratio`` (tiny ``< 1 KiB`` members are
      exempt from the ratio check but still counted).

    Args:
        path: Filesystem path to the ``.xlsx`` / ``.xlsm`` snapshot file.
        limits: The bounded ZIP limits.

    Returns:
        The opened, vetted :class:`zipfile.ZipFile` (the caller closes it).

    Raises:
        OoxmlSafetyError: On any violation (with a precise reason) or when the
            file is not a valid ZIP archive.
    """

    # Pre-open size reject (contract §1 ``max_file_mb``): stat the file BEFORE
    # opening the central directory so a pathologically large input is refused
    # without scanning it.
    try:
        file_size = os.stat(path).st_size
    except OSError as exc:
        raise OoxmlSafetyError(f"Cannot stat OOXML input file {path!r}: {exc}") from exc
    if file_size > limits.max_file_bytes:
        raise OoxmlSafetyError(
            f"OOXML input file size {file_size} bytes exceeds the limit "
            f"({limits.max_file_mb} MiB)."
        )

    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        raise OoxmlSafetyError(f"Not a valid OOXML (ZIP) archive: {exc}") from exc

    infos = zf.infolist()
    if len(infos) > limits.max_zip_members:
        raise OoxmlSafetyError(
            f"ZIP member count {len(infos)} exceeds the limit "
            f"({limits.max_zip_members})."
        )

    total_uncompressed = 0
    for info in infos:
        name = info.filename
        if _is_unsafe_member_name(name):
            raise OoxmlSafetyError(f"Unsafe ZIP member name rejected: {name!r}.")
        total_uncompressed += info.file_size
        if total_uncompressed > limits.max_uncompressed_bytes:
            raise OoxmlSafetyError(
                f"Total uncompressed size exceeds the limit "
                f"({limits.max_uncompressed_mb} MiB)."
            )
        if info.file_size >= _RATIO_SKIP_BYTES and info.compress_size > 0:
            ratio = info.file_size / info.compress_size
            if ratio > limits.max_compression_ratio:
                raise OoxmlSafetyError(
                    f"ZIP member {name!r} compression ratio {ratio:.1f} exceeds "
                    f"the limit ({limits.max_compression_ratio})."
                )

    return zf


def read_part(zf: zipfile.ZipFile, name: str, limits: ZipLimits) -> bytes:
    """Read one member with a defensive uncompressed-size bound.

    Reads at most ``max_uncompressed_bytes + 1`` bytes and raises if the member
    is larger, so a single part can never exceed the configured ceiling even if
    the central directory understated its size.

    Raises:
        OoxmlSafetyError: If the part exceeds the uncompressed-size limit, or the
            member is missing / unreadable.
    """

    try:
        info = zf.getinfo(name)
    except KeyError as exc:
        raise OoxmlSafetyError(f"Required OOXML part is missing: {name!r}.") from exc

    limit = limits.max_uncompressed_bytes
    try:
        with zf.open(info) as handle:
            data = handle.read(limit + 1)
    except (zipfile.BadZipFile, OSError) as exc:
        raise OoxmlSafetyError(f"Failed to read OOXML part {name!r}: {exc}") from exc

    if len(data) > limit:
        raise OoxmlSafetyError(
            f"OOXML part {name!r} exceeds the uncompressed-size limit "
            f"({limits.max_uncompressed_mb} MiB)."
        )
    return data


__all__ = ["OoxmlSafetyError", "ZipLimits", "open_safe", "read_part"]