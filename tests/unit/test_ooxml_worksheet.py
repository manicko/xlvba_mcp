"""Unit tests for the streaming OOXML worksheet reader (WP-E1).

Hermetic: builds a minimal ``.xlsx`` (workbook + rels + one sheet +
sharedStrings) with :mod:`zipfile` and hand-written XML in ``tmp_path`` and
exercises :func:`load_shared_strings`, :func:`iter_sheet_rows` and
:func:`workbook_parts`. No Excel/COM, no openpyxl, no repository data files.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from excel_vba_mcp.adapters.ooxml.shared_strings import load_shared_strings
from excel_vba_mcp.adapters.ooxml.workbook_meta import sheet_part, workbook_parts
from excel_vba_mcp.adapters.ooxml.worksheet import iter_sheet_rows
from excel_vba_mcp.adapters.ooxml.zip_safety import ZipLimits, open_safe
from excel_vba_mcp.domain.statuses import CellValueKind

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


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


def _build_xlsx(path: Path, *, sheet_name: str, rows_xml: str, shared: list[str]) -> None:
    shared_items = "".join(f"<si><t>{value}</t></si>" for value in shared)
    shared_xml = f'<sst xmlns="{_NS}">{shared_items}</sst>'
    workbook_xml = (
        f'<workbook xmlns="{_NS}" xmlns:r="{_REL_NS}">'
        f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        f'<Relationships xmlns="{_PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_NS}/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    sheet_xml = f'<worksheet xmlns="{_NS}"><sheetData>{rows_xml}</sheetData></worksheet>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/sharedStrings.xml", shared_xml)


def _sample_rows() -> str:
    return (
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="b"><v>1</v></c></row>'
        '<row r="3"><c r="A3" t="inlineStr"><is><t>Inline</t></is></c>'
        '<c r="B3" t="str"><v>F</v></c></row>'
    )


def test_workbook_parts_maps_sheet_name_to_part(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, sheet_name="Data", rows_xml=_sample_rows(), shared=["Hello", "World"])

    with open_safe(str(path), _limits()) as zf:
        parts = workbook_parts(zf, _limits())

    assert [entry["name"] for entry in parts["sheets"]] == ["Data"]
    assert parts["sheets"][0]["part"] == "xl/worksheets/sheet1.xml"
    assert parts["sheets"][0]["hidden"] is False
    assert sheet_part(parts, "Data") == "xl/worksheets/sheet1.xml"
    assert parts["defined_names"] == []
    assert parts["table_parts"] == []


def test_load_shared_strings_resolves(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, sheet_name="Data", rows_xml=_sample_rows(), shared=["Hello", "World"])

    with open_safe(str(path), _limits()) as zf:
        shared = load_shared_strings(zf, _limits())

    assert shared == ["Hello", "World"]


def test_iter_sheet_rows_returns_typed_rows(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, sheet_name="Data", rows_xml=_sample_rows(), shared=["Hello", "World"])

    with open_safe(str(path), _limits()) as zf:
        shared = load_shared_strings(zf, _limits())
        rows, truncated, total = iter_sheet_rows(
            zf, "xl/worksheets/sheet1.xml", shared, _limits(), 0, 10
        )

    assert truncated is False
    assert total == 3
    assert len(rows) == 3

    assert rows[0][0].kind == CellValueKind.STRING
    assert rows[0][0].value == "Hello"
    assert rows[0][1].kind == CellValueKind.NUMBER
    assert rows[0][1].value == 42

    assert rows[1][0].kind == CellValueKind.STRING
    assert rows[1][0].value == "World"
    assert rows[1][1].kind == CellValueKind.BOOLEAN
    assert rows[1][1].value is True

    assert rows[2][0].kind == CellValueKind.STRING
    assert rows[2][0].value == "Inline"
    assert rows[2][1].kind == CellValueKind.STRING
    assert rows[2][1].value == "F"


def test_iter_sheet_rows_offset_limit_truncates(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, sheet_name="Data", rows_xml=_sample_rows(), shared=["Hello", "World"])

    with open_safe(str(path), _limits()) as zf:
        shared = load_shared_strings(zf, _limits())
        rows, truncated, total = iter_sheet_rows(
            zf, "xl/worksheets/sheet1.xml", shared, _limits(), 1, 1
        )

    assert truncated is True
    assert total is None
    assert len(rows) == 1
    # offset=1 skips the first row; the second row is "World" / True.
    assert rows[0][0].value == "World"
    assert rows[0][1].value is True


def test_shared_strings_absent_returns_empty(tmp_path: Path) -> None:
    path = tmp_path / "no_shared.xlsx"
    sheet_xml = (
        f'<worksheet xmlns="{_NS}"><sheetData>'
        '<row r="1"><c r="A1"><v>1</v></c></row>'
        "</sheetData></worksheet>"
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    with open_safe(str(path), _limits()) as zf:
        assert load_shared_strings(zf, _limits()) == []