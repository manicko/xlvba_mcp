"""Unit tests for the WP2 sheet dispatch tools (plan B3).

These tests never touch COM / Excel. They install a fake ``SheetsService`` via
the ``sheet_tools`` setter and assert that the flat-signature ``sheet`` /
``sheet_read`` dispatchers route to the existing handlers, enforce per-action
required parameters as structured :class:`ExcelVbaMcpError` (option (a) —
handler validation), and advertise the truthful annotation triples from the WP0
matrix. ``sheet_delete`` stays standalone and destructive.
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import sheet_tools
from excel_vba_mcp.api.sheet_tools import sheet, sheet_delete, sheet_read
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import SheetInfo
from excel_vba_mcp.domain.statuses import SheetObjectType, SheetVisibility

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _sheet_info(name: str) -> SheetInfo:
    return SheetInfo(
        name=name,
        index=1,
        object_type=SheetObjectType.WORKSHEET,
        visibility=SheetVisibility.VISIBLE,
    )


class FakeSheetsService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_sheets(self, workbook_id: str) -> list[SheetInfo]:
        self.calls.append(("list", (workbook_id,)))
        return [_sheet_info("Sheet1")]

    async def get_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        self.calls.append(("get", (workbook_id, name)))
        return _sheet_info(name)

    async def add_sheet(
        self, workbook_id: str, name: str, sheet_type: str
    ) -> SheetInfo:
        self.calls.append(("add", (workbook_id, name, sheet_type)))
        return _sheet_info(name)

    async def rename_sheet(self, workbook_id: str, name: str, new_name: str) -> SheetInfo:
        self.calls.append(("rename", (workbook_id, name, new_name)))
        return _sheet_info(new_name)

    async def copy_sheet(
        self, workbook_id: str, name: str, *, target_name: str | None = None
    ) -> SheetInfo:
        self.calls.append(("copy", (workbook_id, name, target_name)))
        return _sheet_info(target_name or name)

    async def delete_sheet(self, workbook_id: str, name: str) -> SheetInfo:
        self.calls.append(("delete", (workbook_id, name)))
        return _sheet_info(name)

    async def set_visibility(self, workbook_id: str, name: str, visible: bool) -> SheetInfo:
        self.calls.append(("set_visibility", (workbook_id, name, visible)))
        return _sheet_info(name)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeSheetsService()
    sheet_tools.set_sheets_service(fake)  # type: ignore[arg-type]
    yield fake
    sheet_tools.set_sheets_service(None)


# ---------------------------------------------------------------------------
# sheet: routing
# ---------------------------------------------------------------------------
async def test_sheet_add_routes_to_add(_install_fake: FakeSheetsService) -> None:
    info = await sheet("add", workbook_id="wb1", name="Data")

    assert info.name == "Data"
    name, args = _install_fake.calls[-1]
    assert name == "add"
    assert args == ("wb1", "Data", "worksheet")


async def test_sheet_rename_routes_to_rename(_install_fake: FakeSheetsService) -> None:
    await sheet("rename", workbook_id="wb1", name="Old", new_name="New")

    name, args = _install_fake.calls[-1]
    assert name == "rename"
    assert args == ("wb1", "Old", "New")


async def test_sheet_copy_routes_to_copy(_install_fake: FakeSheetsService) -> None:
    await sheet("copy", workbook_id="wb1", name="Data")

    name, args = _install_fake.calls[-1]
    assert name == "copy"
    assert args == ("wb1", "Data", None)


async def test_sheet_copy_with_target_name_routes(_install_fake: FakeSheetsService) -> None:
    info = await sheet("copy", workbook_id="wb1", name="Data", target_name="DataCopy")

    assert info.name == "DataCopy"
    name, args = _install_fake.calls[-1]
    assert name == "copy"
    assert args == ("wb1", "Data", "DataCopy")


async def test_sheet_set_visibility_routes(_install_fake: FakeSheetsService) -> None:
    await sheet("set-visibility", workbook_id="wb1", name="Data", visible=False)

    name, args = _install_fake.calls[-1]
    assert name == "set_visibility"
    assert args == ("wb1", "Data", False)


# ---------------------------------------------------------------------------
# sheet: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_sheet_add_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("add", workbook_id="wb1")


async def test_sheet_rename_missing_names_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("rename", workbook_id="wb1")


async def test_sheet_rename_missing_new_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("rename", workbook_id="wb1", name="Old")


async def test_sheet_copy_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("copy", workbook_id="wb1")


async def test_sheet_set_visibility_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("set-visibility", workbook_id="wb1", visible=True)


async def test_sheet_set_visibility_missing_visible_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("set-visibility", workbook_id="wb1", name="Data")


async def test_sheet_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("add", name="Data")


async def test_sheet_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet("delete", workbook_id="wb1", name="Data")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# sheet_read: routing + validation
# ---------------------------------------------------------------------------
async def test_sheet_read_list_routes(_install_fake: FakeSheetsService) -> None:
    listed = await sheet_read("list", workbook_id="wb1")

    assert [s.name for s in listed] == ["Sheet1"]
    name, args = _install_fake.calls[-1]
    assert name == "list"
    assert args == ("wb1",)


async def test_sheet_read_get_routes(_install_fake: FakeSheetsService) -> None:
    info = await sheet_read("get", workbook_id="wb1", name="Data")

    assert info.name == "Data"
    name, args = _install_fake.calls[-1]
    assert name == "get"
    assert args == ("wb1", "Data")


async def test_sheet_read_get_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet_read("get", workbook_id="wb1")


async def test_sheet_read_list_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet_read("list")


async def test_sheet_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet_read("destroy", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# sheet_delete: standalone, unchanged behaviour (destructive)
# ---------------------------------------------------------------------------
async def test_sheet_delete_routes(_install_fake: FakeSheetsService) -> None:
    info = await sheet_delete("wb1", "Data")

    assert info.name == "Data"
    name, args = _install_fake.calls[-1]
    assert name == "delete"
    assert args == ("wb1", "Data")


async def test_sheet_delete_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await sheet_delete("wb1", "")


# ---------------------------------------------------------------------------
# Annotation triples (WP0 truth matrix; None treated as "not True")
# ---------------------------------------------------------------------------
async def test_sheet_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["sheet"].annotations
    assert ann is not None
    assert ann.destructive_hint is not True
    assert ann.read_only_hint is not True
    assert ann.idempotent_hint is not True


async def test_sheet_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["sheet_read"].annotations
    assert ann is not None
    assert ann.read_only_hint is True
    assert ann.idempotent_hint is True
    assert ann.destructive_hint is not True


async def test_sheet_delete_is_destructive_standalone() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["sheet_delete"].annotations
    assert ann is not None
    assert ann.destructive_hint is True


async def test_old_sheet_tool_names_are_not_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for removed in (
        "sheets_list",
        "sheet_get",
        "sheet_add",
        "sheet_rename",
        "sheet_copy",
        "sheet_set_visibility",
    ):
        assert removed not in names