"""Unit tests for the non-COM OOXML read service (WP-E1).

Hermetic: a tiny ``.xlsx`` is built with :mod:`zipfile` in ``tmp_path`` and a
fake session + fake snapshot (returning that path + a SHA-256) are injected into
the production :class:`~excel_vba_mcp.application.ooxml_service.OoxmlService`.
The full parse path runs (via ``asyncio.to_thread``) without Excel/COM. No
openpyxl, no repository data files.
"""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_vba_mcp.application.ooxml_service import OoxmlService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.models import OoxmlMetadata, OoxmlSheetPage
from excel_vba_mcp.domain.statuses import CellValueKind

pytestmark = pytest.mark.anyio

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _build_xlsx(path: Path) -> None:
    shared_xml = f'<sst xmlns="{_NS}"><si><t>Hello</t></si><si><t>World</t></si></sst>'
    workbook_xml = (
        f'<workbook xmlns="{_NS}" xmlns:r="{_REL_NS}">'
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels_xml = (
        f'<Relationships xmlns="{_PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_NS}/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    rows_xml = (
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="b"><v>1</v></c></row>'
    )
    sheet_xml = f'<worksheet xmlns="{_NS}"><sheetData>{rows_xml}</sheetData></worksheet>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", rels_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/sharedStrings.xml", shared_xml)


class _FakeStore:
    def __init__(self, path: str, sha: str) -> None:
        self._path = path
        self._sha = sha

    def load(self, snapshot_id: str) -> SimpleNamespace:
        return SimpleNamespace(stored_path=self._path, sha256=self._sha, snapshot_id=snapshot_id)


class _FakeSnapshot:
    """Fake SnapshotService: create + store lookup return the tmp xlsx (no COM)."""

    def __init__(self, path: str, sha: str) -> None:
        self._path = path
        self._sha = sha
        self.store = _FakeStore(path, sha)

    async def create_snapshot(
        self, workbook_id: str, *, timeout_s: float = 60.0
    ) -> SimpleNamespace:
        return SimpleNamespace(
            stored_path=self._path, sha256=self._sha, snapshot_id="snap-created"
        )


def _service(tmp_path: Path) -> tuple[OoxmlService, str]:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path)
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    snapshot = _FakeSnapshot(str(path), sha)
    service = OoxmlService(session=object(), snapshot=snapshot, limits=LargeDataConfig())  # type: ignore[arg-type]
    return service, sha


async def test_read_sheet_rows_returns_snapshot_page(tmp_path: Path) -> None:
    service, sha = _service(tmp_path)

    page = await service.read_sheet_rows("wb1", "Data", offset=0, limit=10, snapshot_id="snap-1")

    assert isinstance(page, OoxmlSheetPage)
    assert page.page.source == "snapshot"
    assert page.page.version == 0
    assert page.input.input_sha256 == sha
    assert page.input.snapshot_id == "snap-1"
    assert page.rows == 2
    assert page.cols == 2
    assert page.values[0][0].kind == CellValueKind.STRING
    assert page.values[0][0].value == "Hello"
    assert page.values[0][1].value == 42
    assert page.values[1][1].kind == CellValueKind.BOOLEAN
    assert page.values[1][1].value is True
    assert page.page.complete is True
    assert page.page.has_more is False


async def test_read_sheet_rows_creates_snapshot_when_none(tmp_path: Path) -> None:
    service, sha = _service(tmp_path)

    page = await service.read_sheet_rows("wb1", "Data", offset=0, limit=1)

    assert page.input.snapshot_id == "snap-created"
    assert page.input.input_sha256 == sha
    assert page.rows == 1
    assert page.page.truncated is True
    assert page.page.has_more is True


async def test_read_metadata_lists_sheet(tmp_path: Path) -> None:
    service, sha = _service(tmp_path)

    meta = await service.read_metadata("wb1", snapshot_id="snap-1")

    assert isinstance(meta, OoxmlMetadata)
    assert meta.input.input_sha256 == sha
    assert [sheet.name for sheet in meta.sheets] == ["Data"]
    assert meta.sheets[0].part == "xl/worksheets/sheet1.xml"
    assert meta.table_count == 0


async def test_read_sheet_rows_unknown_sheet_errors(tmp_path: Path) -> None:
    from excel_vba_mcp.domain.errors import ExcelVbaMcpError

    service, _sha = _service(tmp_path)

    with pytest.raises(ExcelVbaMcpError):
        await service.read_sheet_rows("wb1", "Missing", snapshot_id="snap-1")