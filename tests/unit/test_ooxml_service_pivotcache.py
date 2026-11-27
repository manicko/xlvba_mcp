"""Unit tests for the PivotCache records path of the OOXML read service (WP-E2).

Hermetic: a tiny ``.xlsx`` with a ``pivotCacheDefinition`` + ``pivotCacheRecords``
is built with :mod:`zipfile` in ``tmp_path`` and a fake session + fake snapshot
(returning that path + a SHA-256) are injected into the production
:class:`~excel_vba_mcp.application.ooxml_service.OoxmlService`. The full parse path
runs (via ``asyncio.to_thread``) without Excel/COM. No openpyxl, no repository data
files.
"""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pytest

from excel_vba_mcp.application.ooxml_service import OoxmlService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import OoxmlPivotCachePage
from excel_vba_mcp.domain.statuses import CellValueKind

pytestmark = pytest.mark.anyio

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _definition_xml(*, save_data: str = "1") -> str:
    return (
        f'<pivotCacheDefinition xmlns="{_NS}" recordCount="2" '
        f'saveData="{save_data}" refreshDate="2026-01-15T10:30:00">'
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
        f'<pivotCacheRecords xmlns="{_NS}" count="2">'
        '<r><s v="0"/><n v="30"/></r>'
        '<r><s v="1"/><n v="25"/></r>'
        "</pivotCacheRecords>"
    )


def _build_xlsx(path: Path, *, save_data: str, with_records: bool) -> None:
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
        zf.writestr(
            "xl/pivotCache/pivotCacheDefinition1.xml",
            _definition_xml(save_data=save_data),
        )
        if with_records:
            zf.writestr(
                "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", def_rels_xml
            )
            zf.writestr("xl/pivotCache/pivotCacheRecords1.xml", _records_xml())


class _FakeStore:
    def __init__(self, path: str, sha: str) -> None:
        self._path = path
        self._sha = sha

    def load(self, snapshot_id: str) -> SimpleNamespace:
        return SimpleNamespace(
            stored_path=self._path, sha256=self._sha, snapshot_id=snapshot_id
        )


class _FakeSnapshot:
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


def _service(tmp_path: Path, *, save_data: str, with_records: bool) -> tuple[OoxmlService, str]:
    path = tmp_path / "book.xlsx"
    _build_xlsx(path, save_data=save_data, with_records=with_records)
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    snapshot = _FakeSnapshot(str(path), sha)
    service = OoxmlService(session=object(), snapshot=snapshot, limits=LargeDataConfig())  # type: ignore[arg-type]
    return service, sha


async def test_read_pivot_cache_records_returns_cache_page(tmp_path: Path) -> None:
    service, sha = _service(tmp_path, save_data="1", with_records=True)

    page = await service.read_pivot_cache_records(
        "wb1", 1, offset=0, limit=10, snapshot_id="snap-1"
    )

    assert isinstance(page, OoxmlPivotCachePage)
    assert page.page.source == "cache"
    assert page.page.version == 0
    assert page.input.input_sha256 == sha
    assert page.input.snapshot_id == "snap-1"
    assert page.cache_index == 1
    assert page.meta.has_records_part is True
    assert page.meta.record_count == 2
    assert page.rows == 2
    assert page.cols == 2
    assert page.values[0][0].kind == CellValueKind.STRING
    assert page.values[0][0].value == "Alice"
    assert page.values[0][1].value == 30
    assert page.values[1][0].value == "Bob"
    assert page.page.complete is True
    assert page.page.reason is None


async def test_read_pivot_cache_records_unsupported_when_save_data_false(
    tmp_path: Path,
) -> None:
    service, _sha = _service(tmp_path, save_data="0", with_records=False)

    page = await service.read_pivot_cache_records("wb1", 1, snapshot_id="snap-1")

    assert page.page.source == "cache"
    assert page.page.reason == "unsupported_source"
    assert page.page.complete is False
    assert page.meta.has_records_part is False
    assert page.values == []
    assert page.rows == 0


async def test_read_pivot_cache_records_unknown_cache_errors(tmp_path: Path) -> None:
    service, _sha = _service(tmp_path, save_data="1", with_records=True)

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await service.read_pivot_cache_records("wb1", 99, snapshot_id="snap-1")

    assert excinfo.value.code.value == "NOT_FOUND"