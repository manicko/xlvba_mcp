"""Unit tests for the WP4 name dispatch tools (plan B5).

These tests never touch COM / Excel. They install a fake ``NamesService`` via
the ``name_tools`` setter and assert that the flat-signature ``name`` /
``name_read`` dispatchers route to the existing handlers, enforce per-action
required parameters as structured :class:`ExcelVbaMcpError` (option (a) —
handler validation), and advertise the truthful annotation triples from the WP0
matrix. ``name_delete`` stays standalone and destructive.
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import name_tools
from excel_vba_mcp.api.name_tools import name, name_delete, name_read
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import DefinedNameInfo
from excel_vba_mcp.domain.statuses import NameScope

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _name_info(name: str) -> DefinedNameInfo:
    return DefinedNameInfo(name=name, scope=NameScope.WORKBOOK, refers_to="=42")


class FakeNamesService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_names(self, workbook_id: str) -> list[DefinedNameInfo]:
        self.calls.append(("list", (workbook_id,)))
        return [_name_info("Sales")]

    async def get_name(
        self,
        workbook_id: str,
        name: str,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        self.calls.append(("get", (workbook_id, name, scope, sheet)))
        return _name_info(name)

    async def create_name(
        self,
        workbook_id: str,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        self.calls.append(("create", (workbook_id, name, refers_to, scope, sheet)))
        return _name_info(name)

    async def update_name(
        self,
        workbook_id: str,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        self.calls.append(("update", (workbook_id, name, refers_to, scope, sheet)))
        return _name_info(name)

    async def delete_name(
        self,
        workbook_id: str,
        name: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        self.calls.append(("delete", (workbook_id, name, scope, sheet)))
        return _name_info(name)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeNamesService()
    name_tools.set_names_service(fake)  # type: ignore[arg-type]
    yield fake
    name_tools.set_names_service(None)


# ---------------------------------------------------------------------------
# name: routing
# ---------------------------------------------------------------------------
async def test_name_create_routes(_install_fake: FakeNamesService) -> None:
    await name(
        "create",
        workbook_id="wb1",
        name="Sales",
        refers_to="=Sheet1!$A$1",
        scope="workbook",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "create"
    assert args == ("wb1", "Sales", "=Sheet1!$A$1", NameScope.WORKBOOK, None)


async def test_name_update_routes(_install_fake: FakeNamesService) -> None:
    await name(
        "update",
        workbook_id="wb1",
        name="Sales",
        refers_to="=Sheet1!$B$2",
        scope="workbook",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "update"
    assert args == ("wb1", "Sales", "=Sheet1!$B$2", NameScope.WORKBOOK, None)


async def test_name_create_worksheet_scope_forwards_sheet(
    _install_fake: FakeNamesService,
) -> None:
    await name(
        "create",
        workbook_id="wb1",
        name="Local",
        refers_to="=42",
        scope="worksheet",
        sheet="Data",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "create"
    assert args == ("wb1", "Local", "=42", NameScope.WORKSHEET, "Data")


# ---------------------------------------------------------------------------
# name: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_name_create_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("create", workbook_id="wb1", refers_to="=42", scope="workbook")


async def test_name_create_missing_refers_to_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("create", workbook_id="wb1", name="Sales", scope="workbook")


async def test_name_create_missing_scope_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("create", workbook_id="wb1", name="Sales", refers_to="=42")


async def test_name_update_missing_refers_to_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("update", workbook_id="wb1", name="Sales", scope="workbook")


async def test_name_update_missing_scope_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("update", workbook_id="wb1", name="Sales", refers_to="=42")


async def test_name_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("create", name="Sales", refers_to="=42", scope="workbook")


async def test_name_worksheet_scope_missing_sheet_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name(
            "create", workbook_id="wb1", name="Local", refers_to="=42", scope="worksheet"
        )


async def test_name_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("destroy", workbook_id="wb1", name="Sales")  # type: ignore[arg-type]


async def test_name_delete_is_not_a_name_action() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name("delete", workbook_id="wb1", name="Sales", scope="workbook")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# name_read: routing + validation
# ---------------------------------------------------------------------------
async def test_name_read_list_routes(_install_fake: FakeNamesService) -> None:
    listed = await name_read("list", workbook_id="wb1")

    assert [n.name for n in listed] == ["Sales"]
    verb, args = _install_fake.calls[-1]
    assert verb == "list"
    assert args == ("wb1",)


async def test_name_read_get_routes(_install_fake: FakeNamesService) -> None:
    info = await name_read("get", workbook_id="wb1", name="Sales", scope="workbook")

    assert info.name == "Sales"
    verb, args = _install_fake.calls[-1]
    assert verb == "get"
    assert args == ("wb1", "Sales", NameScope.WORKBOOK, None)


async def test_name_read_get_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name_read("get", workbook_id="wb1", scope="workbook")


async def test_name_read_get_missing_scope_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name_read("get", workbook_id="wb1", name="Sales")


async def test_name_read_list_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name_read("list")


async def test_name_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name_read("destroy", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# name_delete: standalone destructive
# ---------------------------------------------------------------------------
async def test_name_delete_routes(_install_fake: FakeNamesService) -> None:
    await name_delete("wb1", "Sales", scope="workbook")

    verb, args = _install_fake.calls[-1]
    assert verb == "delete"
    assert args == ("wb1", "Sales", NameScope.WORKBOOK, None)


async def test_name_delete_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await name_delete("wb1", "", scope="workbook")


# ---------------------------------------------------------------------------
# Annotation triples (WP0 truth matrix; None treated as "not True")
# ---------------------------------------------------------------------------
async def test_name_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["defined_name"].annotations
    assert ann is not None
    assert ann.destructive_hint is not True
    assert ann.read_only_hint is not True
    assert ann.idempotent_hint is True


async def test_name_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["defined_name_read"].annotations
    assert ann is not None
    assert ann.read_only_hint is True
    assert ann.idempotent_hint is True
    assert ann.destructive_hint is not True


async def test_name_delete_is_destructive_standalone() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["defined_name_delete"].annotations
    assert ann is not None
    assert ann.destructive_hint is True


async def test_old_name_tool_names_are_not_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for removed in ("names_list", "name_get", "name_create", "name_update"):
        assert removed not in names