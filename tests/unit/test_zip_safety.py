"""Unit tests for the bounded OOXML ZIP reader (WP-E1).

Hermetic: builds tiny ZIP archives with :mod:`zipfile` in ``tmp_path`` and
asserts the ZIP-bomb / traversal guards in
:func:`excel_vba_mcp.adapters.ooxml.zip_safety.open_safe`. No Excel/COM, no
openpyxl, no repository data files.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    open_safe,
    read_part,
)


def _limits(**overrides: int) -> ZipLimits:
    base = dict(
        max_file_mb=100,
        max_uncompressed_mb=1024,
        max_compression_ratio=1000,
        max_zip_members=10_000,
        max_shared_strings=1_000_000,
        max_xml_depth=64,
        max_response_cells=50_000,
    )
    base.update(overrides)
    return ZipLimits(**base)  # type: ignore[arg-type]


def test_normal_tiny_archive_passes(tmp_path: Path) -> None:
    path = tmp_path / "ok.xlsx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", b"<workbook/>")
        zf.writestr("xl/worksheets/sheet1.xml", b"<worksheet/>")

    with open_safe(str(path), _limits()) as zf:
        assert "xl/workbook.xml" in zf.namelist()
        data = read_part(zf, "xl/workbook.xml", _limits())
    assert data == b"<workbook/>"


def test_member_count_rejected(tmp_path: Path) -> None:
    path = tmp_path / "many.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for index in range(3):
            zf.writestr(f"p{index}.txt", b"x")

    with pytest.raises(OoxmlSafetyError):
        open_safe(str(path), _limits(max_zip_members=2))


def test_uncompressed_size_rejected(tmp_path: Path) -> None:
    path = tmp_path / "big.xlsx"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("a.txt", b"x" * 4096)

    # max_uncompressed_mb=0 -> any non-empty member exceeds the ceiling.
    with pytest.raises(OoxmlSafetyError):
        open_safe(str(path), _limits(max_uncompressed_mb=0))


def test_compression_ratio_rejected(tmp_path: Path) -> None:
    path = tmp_path / "bomb.xlsx"
    # 4096 zero bytes compress to a tiny payload -> a very high ratio.
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("zeros.bin", b"\x00" * 4096)

    with pytest.raises(OoxmlSafetyError):
        open_safe(str(path), _limits(max_compression_ratio=5))


def test_path_traversal_rejected(tmp_path: Path) -> None:
    path = tmp_path / "evil.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("../evil.txt", b"nope")

    with pytest.raises(OoxmlSafetyError):
        open_safe(str(path), _limits())


def test_read_part_missing_member_raises(tmp_path: Path) -> None:
    path = tmp_path / "ok2.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", b"<workbook/>")

    with open_safe(str(path), _limits()) as zf, pytest.raises(OoxmlSafetyError):
        read_part(zf, "xl/missing.xml", _limits())


def test_not_a_zip_raises(tmp_path: Path) -> None:
    path = tmp_path / "notzip.xlsx"
    path.write_bytes(b"this is not a zip archive")

    with pytest.raises(OoxmlSafetyError):
        open_safe(str(path), _limits())


def test_file_size_pre_open_rejected(tmp_path: Path) -> None:
    # max_file_mb=0 -> any non-empty file exceeds the byte ceiling and is
    # rejected BEFORE the central directory is opened (contract §1 max_file_mb).
    path = tmp_path / "toobig.xlsx"
    path.write_bytes(b"x" * 4096)

    with pytest.raises(OoxmlSafetyError, match="exceeds the limit"):
        open_safe(str(path), _limits(max_file_mb=0))


def test_file_size_within_limit_passes(tmp_path: Path) -> None:
    # A small file under a 1 MiB cap opens normally (size pre-check passes).
    path = tmp_path / "small.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", b"<workbook/>")

    with open_safe(str(path), _limits(max_file_mb=1)) as zf:
        assert "xl/workbook.xml" in zf.namelist()