"""Unit tests for the streaming OOXML PivotCache reader (WP-E2).

Hermetic: builds a minimal ``.xlsx`` (workbook + rels + a ``pivotCacheDefinition``
+ its rels + a ``pivotCacheRecords`` part) with :mod:`zipfile` and hand-written
XML in ``tmp_path`` and exercises :func:`pivot_caches`,
:func:`read_cache_definition` and :func:`iter_cache_records`. No Excel/COM, no
openpyxl, no repository data files.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from excel_vba_mcp.adapters.ooxml.pivotcache import (
    iter_cache_records,
    read_cache_definition,
)
from excel_vba_mcp.adapters.ooxml.workbook_meta import pivot_caches
from excel_vba_mcp.adapters.ooxml.zip_safety import (
    OoxmlSafetyError,
    ZipLimits,
    open_safe,
)
from excel_vba_mcp.domain.statuses import CellValueKind

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_DEF_PART = "xl/pivotCache/pivotCacheDefinition1.xml"
_DEF_RELS = "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"
_REC_PART = "xl/pivotCache/pivotCacheRecords1.xml"


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


def _definition_xml(*, record_count: str = "3", save_data: str = "1") -> str:
    return (
        f'<pivotCacheDefinition xmlns="{_NS}" recordCount="{record_count}" '
        f'saveData="{save_data}" refreshOnLoad="0" '
        'refreshDate="2026-01-15T10:30:00">'
        '<cacheFields count="2">'
        '<cacheField name="Name"><sharedItems count="2">'
        '<s v="Alice"/><s v="Bob"/>'
        "</sharedItems></cacheField>"
        '<cacheField name="Age"><sharedItems containsNumber="1"/>'
        "</cacheField>"
        "</cacheFields>"
        "</pivotCacheDefinition>"
    )


def _records_xml() -> str:
    return (
        f'<pivotCacheRecords xmlns="{_NS}" count="3">'
        "<r><s v=\"0\"/><n v=\"30\"/></r>"
        "<r><s v=\"1\"/><n v=\"25\"/></r>"
        "<r><s v=\"0\"/><d v=\"2026-02-01\"/></r>"
        "</pivotCacheRecords>"
    )


def _build_xlsx(path: Path, *, definition: str, records: str | None) -> None:
    workbook_xml = (
        f'<workbook xmlns="{_NS}" xmlns:r="{_REL_NS}">'
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
        '<pivotCaches><pivotCache cacheId="1" r:id="rId2"/></pivotCaches>'
        "</workbook>"
    )
    rels_xml = (
        f'<Relationships xmlns="{_PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_NS}/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{_REL_NS}/pivotCacheDefinition" '
        'Target="pivotCache/pivotCacheDefinition1.xml"/>'
        "</Relationships>"
    )
    def_rels_xml = (
        f'<Relationships xmlns="{_PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_NS}/pivotCacheRecords" '
        'Target="pivotCacheRecords1.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr(_DEF_PART, definition)
        if records is not None:
            zf.writestr(_DEF_RELS, def_rels_xml)
            zf.writestr(_REC_PART, records)


def test_pivot_caches_maps_index_to_definition_and_records(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    with open_safe(str(path), _limits()) as zf:
        caches = pivot_caches(zf, _limits())

    assert len(caches) == 1
    assert caches[0]["cache_index"] == 1
    assert caches[0]["definition_part"] == _DEF_PART
    assert caches[0]["records_part"] == _REC_PART


def test_pivot_caches_records_part_none_when_absent(tmp_path: Path) -> None:
    path = tmp_path / "no_records.xlsx"
    _build_xlsx(path, definition=_definition_xml(save_data="0"), records=None)

    with open_safe(str(path), _limits()) as zf:
        caches = pivot_caches(zf, _limits())

    assert caches[0]["records_part"] is None


def test_read_cache_definition_reads_metadata_and_shared_items(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    with open_safe(str(path), _limits()) as zf:
        definition = read_cache_definition(zf, _DEF_PART, _limits())

    assert definition["record_count"] == 3
    assert definition["save_data"] is True
    assert definition["refresh_on_load"] is False
    assert definition["refresh_date"] is not None
    shared = definition["shared_items"]
    assert isinstance(shared, list)
    assert [cell.value for cell in shared[0]] == ["Alice", "Bob"]
    assert shared[1] == []  # numeric field: no shared string items


def test_iter_cache_records_streams_typed_rows_with_shared_resolution(
    tmp_path: Path,
) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    with open_safe(str(path), _limits()) as zf:
        definition = read_cache_definition(zf, _DEF_PART, _limits())
        shared = definition["shared_items"]
        rows, truncated, total, scanned = iter_cache_records(
            zf,
            _REC_PART,
            shared,  # type: ignore[arg-type]
            _limits(),
            0,
            10,
        )

    assert truncated is False
    assert total == 3
    assert scanned == 3
    assert len(rows) == 3

    assert rows[0][0].kind == CellValueKind.STRING
    assert rows[0][0].value == "Alice"
    assert rows[0][1].kind == CellValueKind.NUMBER
    assert rows[0][1].value == 30

    assert rows[1][0].value == "Bob"
    assert rows[1][1].value == 25

    assert rows[2][0].value == "Alice"
    assert rows[2][1].kind == CellValueKind.DATE
    assert rows[2][1].value == "2026-02-01"


def test_iter_cache_records_offset_limit_truncates(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    with open_safe(str(path), _limits()) as zf:
        definition = read_cache_definition(zf, _DEF_PART, _limits())
        shared = definition["shared_items"]
        rows, truncated, total, scanned = iter_cache_records(
            zf,
            _REC_PART,
            shared,  # type: ignore[arg-type]
            _limits(),
            1,
            1,
        )

    assert truncated is True
    assert total is None
    assert scanned == 2  # offset 1 + limit 1 => two records inspected
    assert len(rows) == 1
    assert rows[0][0].value == "Bob"
    assert rows[0][1].value == 25


def test_iter_cache_records_per_row_cell_cap(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    # max_response_cells // limit = 2 // 1 = 2 cells per row (both fields kept);
    # a tighter cap of 1 keeps only the first field.
    with open_safe(str(path), _limits()) as zf:
        definition = read_cache_definition(zf, _DEF_PART, _limits())
        shared = definition["shared_items"]
        rows, _truncated, _total, _scanned = iter_cache_records(
            zf,
            _REC_PART,
            shared,  # type: ignore[arg-type]
            _limits(max_response_cells=1),
            0,
            10,
        )

    assert all(len(row) == 1 for row in rows)
    assert rows[0][0].value == "Alice"


def test_iter_cache_records_depth_guard(tmp_path: Path) -> None:
    path = tmp_path / "deep.xlsx"
    _build_xlsx(path, definition=_definition_xml(), records=_records_xml())

    with open_safe(str(path), _limits()) as zf:
        definition = read_cache_definition(zf, _DEF_PART, _limits())
        shared = definition["shared_items"]
        with pytest.raises(OoxmlSafetyError):
            iter_cache_records(
                zf,
                _REC_PART,
                shared,  # type: ignore[arg-type]
                _limits(max_xml_depth=1),
                0,
                10,
            )