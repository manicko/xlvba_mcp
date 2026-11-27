"""Unit tests for the ``workbook_file_read`` dispatch tool (WP-E1).

Hermetic: installs a fake ``OoxmlService`` via the ``ooxml_tools`` setter and
asserts the flat-signature ``workbook_file_read`` dispatcher routes ``metadata`` /
``sheet-rows`` to the service, enforces per-action required parameters as
structured :class:`ExcelVbaMcpError`, and rejects an unknown action. No Excel/COM.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_vba_mcp.api import ooxml_tools
from excel_vba_mcp.api.ooxml_tools import workbook_file_read
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    OoxmlInputInfo,
    OoxmlMetadata,
    OoxmlPivotCacheMeta,
    OoxmlPivotCachePage,
    OoxmlSheetPage,
    PageInfo,
)

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _metadata(workbook_id: str) -> OoxmlMetadata:
    return OoxmlMetadata(
        workbook_id=workbook_id, input=OoxmlInputInfo(input_sha256="sha")
    )


def _page(workbook_id: str, sheet: str) -> OoxmlSheetPage:
    return OoxmlSheetPage(
        workbook_id=workbook_id,
        sheet=sheet,
        input=OoxmlInputInfo(input_sha256="sha"),
        rows=0,
        cols=0,
        values=[],
        page=PageInfo(offset=0, returned=0, source="snapshot"),
    )


def _pivot_page(workbook_id: str, cache_index: int) -> OoxmlPivotCachePage:
    return OoxmlPivotCachePage(
        workbook_id=workbook_id,
        cache_index=cache_index,
        input=OoxmlInputInfo(input_sha256="sha"),
        meta=OoxmlPivotCacheMeta(cache_index=cache_index),
        rows=0,
        cols=0,
        values=[],
        page=PageInfo(offset=0, returned=0, source="cache"),
    )


class FakeOoxmlService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def read_metadata(
        self, workbook_id: str, snapshot_id: str | None = None
    ) -> OoxmlMetadata:
        self.calls.append(("read_metadata", (workbook_id, snapshot_id)))
        return _metadata(workbook_id)

    async def read_sheet_rows(
        self,
        workbook_id: str,
        sheet: str,
        offset: int = 0,
        limit: int | None = None,
        snapshot_id: str | None = None,
    ) -> OoxmlSheetPage:
        self.calls.append(
            ("read_sheet_rows", (workbook_id, sheet, offset, limit, snapshot_id))
        )
        return _page(workbook_id, sheet)

    async def read_pivot_cache_records(
        self,
        workbook_id: str,
        cache_index: int,
        offset: int = 0,
        limit: int | None = None,
        snapshot_id: str | None = None,
    ) -> OoxmlPivotCachePage:
        self.calls.append(
            (
                "read_pivot_cache_records",
                (workbook_id, cache_index, offset, limit, snapshot_id),
            )
        )
        return _pivot_page(workbook_id, cache_index)


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeOoxmlService()
    ooxml_tools.set_ooxml_service(fake)  # type: ignore[arg-type]
    yield fake
    ooxml_tools.set_ooxml_service(None)


async def test_metadata_routes(_install_fake: FakeOoxmlService) -> None:
    result = await workbook_file_read("metadata", workbook_id="wb1", snapshot_id="snap-1")

    verb, args = _install_fake.calls[-1]
    assert verb == "read_metadata"
    assert args == ("wb1", "snap-1")
    assert isinstance(result, OoxmlMetadata)


async def test_sheet_rows_routes(_install_fake: FakeOoxmlService) -> None:
    result = await workbook_file_read(
        "sheet-rows", workbook_id="wb1", sheet="Data", offset=2, limit=5
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "read_sheet_rows"
    assert args == ("wb1", "Data", 2, 5, None)
    assert isinstance(result, OoxmlSheetPage)


async def test_sheet_rows_requires_sheet() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_file_read("sheet-rows", workbook_id="wb1")


async def test_pivot_cache_records_routes(_install_fake: FakeOoxmlService) -> None:
    result = await workbook_file_read(
        "pivot-cache-records", workbook_id="wb1", cache_index=2, offset=1, limit=5
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "read_pivot_cache_records"
    assert args == ("wb1", 2, 1, 5, None)
    assert isinstance(result, OoxmlPivotCachePage)


async def test_pivot_cache_records_requires_cache_index() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_file_read("pivot-cache-records", workbook_id="wb1")


async def test_sheet_rows_rejects_negative_offset() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_file_read(
            "sheet-rows", workbook_id="wb1", sheet="Data", offset=-1
        )


async def test_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_file_read("metadata", workbook_id="")


async def test_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_file_read("destroy", workbook_id="wb1")  # type: ignore[arg-type]