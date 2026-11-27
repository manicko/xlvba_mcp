"""Unit tests for the ``data_export`` dispatch tool (WP-F).

Hermetic: installs a fake ``DataExportService`` via the ``export_tools`` setter
and asserts ``data_export`` routes to the service with the validated arguments,
enforces the required ``workbook_id`` / ``limit_rows`` and rejects an invalid
``limit_rows``. No Excel / no COM.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_vba_mcp.api import export_tools
from excel_vba_mcp.api.export_tools import data_export
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import DataExportResult

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _result() -> DataExportResult:
    return DataExportResult(
        artifact_id="a1",
        path="C:/artifacts/export.csv",
        format="csv",
        rows=6,
        bytes=42,
        sha256="deadbeef",
        complete=True,
        source="range",
        source_version=7,
    )


class FakeDataExportService:
    """Records the export call and returns a canned result; no COM."""

    def __init__(self) -> None:
        self.calls: list[dict[str, Any]] = []

    async def export(self, workbook_id: str, **kwargs: Any) -> DataExportResult:
        self.calls.append({"workbook_id": workbook_id, **kwargs})
        return _result()


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeDataExportService()
    export_tools.set_data_export_service(fake)  # type: ignore[arg-type]
    yield fake
    export_tools.set_data_export_service(None)


async def test_data_export_routes_to_service(_install_fake: FakeDataExportService) -> None:
    result = await data_export(
        "wb1",
        "range",
        sheet="S",
        address="A1:B10",
        fmt="csv",
        limit_rows=100,
    )

    assert isinstance(result, DataExportResult)
    call = _install_fake.calls[-1]
    assert call["workbook_id"] == "wb1"
    assert call["source_kind"] == "range"
    assert call["sheet"] == "S"
    assert call["address"] == "A1:B10"
    assert call["fmt"] == "csv"
    assert call["limit_rows"] == 100


async def test_data_export_requires_workbook_id(_install_fake: FakeDataExportService) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await data_export("", "range", sheet="S", address="A1:B10")


async def test_data_export_rejects_negative_limit_rows(
    _install_fake: FakeDataExportService,
) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await data_export(
            "wb1", "range", sheet="S", address="A1:B10", limit_rows=-1
        )


async def test_data_export_requires_service_when_unset() -> None:
    export_tools.set_data_export_service(None)
    with pytest.raises(ExcelVbaMcpError):
        await data_export("wb1", "range", sheet="S", address="A1:B10")