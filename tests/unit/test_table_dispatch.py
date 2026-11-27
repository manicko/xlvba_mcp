"""Unit tests for the WP3 table dispatch tools (plan B4).

These tests never touch COM / Excel. They install a fake ``TablesService`` via
the ``table_tools`` setter and assert that the flat-signature ``table`` /
``table_read`` dispatchers route to the existing handlers, enforce per-action
required parameters as structured :class:`ExcelVbaMcpError` (option (a) —
handler validation), and advertise the truthful annotation triples from the WP0
matrix. ``table_delete_rows`` (XOR guard) and ``table_delete_column`` stay
standalone and destructive.
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.adapters.excel.tables_adapter import TableMutationResult, TableRows
from excel_vba_mcp.api import table_tools
from excel_vba_mcp.api.table_tools import (
    table,
    table_delete_column,
    table_delete_rows,
    table_read,
)
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import TableInfo

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _table_info(name: str) -> TableInfo:
    return TableInfo(name=name, sheet_name="Sheet1", range_address="A1:B2")


def _mutation(action: str, name: str) -> TableMutationResult:
    return TableMutationResult(
        name=name, action=action, before=None, after=_table_info(name)
    )


def _table_rows(name: str, offset: int, limit: int | None) -> TableRows:
    return TableRows(
        name=name,
        sheet_name="Sheet1",
        columns=["A"],
        offset=offset,
        limit=limit,
        total_rows=1,
        returned_rows=0,
        truncated=False,
        rows=[],
    )


class FakeTablesService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_tables(self, workbook_id: str) -> list[TableInfo]:
        self.calls.append(("list", (workbook_id,)))
        return [_table_info("Sales")]

    async def get_table(self, workbook_id: str, name: str) -> TableInfo:
        self.calls.append(("get", (workbook_id, name)))
        return _table_info(name)

    async def read_table_rows(
        self,
        workbook_id: str,
        name: str,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
    ) -> TableRows:
        self.calls.append(("read_rows", (workbook_id, name, offset, limit)))
        return _table_rows(name, offset, limit)

    async def append_table_rows(
        self, workbook_id: str, name: str, rows: list[dict[str, Any]]
    ) -> TableMutationResult:
        self.calls.append(("append_rows", (workbook_id, name, rows)))
        return _mutation("append_rows", name)

    async def insert_table_row(
        self,
        workbook_id: str,
        name: str,
        at: int,
        values: dict[str, Any] | None = None,
    ) -> TableMutationResult:
        self.calls.append(("insert_row", (workbook_id, name, at, values)))
        return _mutation("insert_row", name)

    async def update_table_rows(
        self, workbook_id: str, name: str, updates: list[dict[str, Any]]
    ) -> TableMutationResult:
        self.calls.append(("update_rows", (workbook_id, name, updates)))
        return _mutation("update_rows", name)

    async def delete_table_rows(
        self,
        workbook_id: str,
        name: str,
        indices: list[int] | None,
        where: dict[str, Any] | None,
    ) -> TableMutationResult:
        self.calls.append(("delete_rows", (workbook_id, name, indices, where)))
        return _mutation("delete_rows", name)

    async def add_table_column(
        self, workbook_id: str, name: str, header: str, position: int | None = None
    ) -> TableMutationResult:
        self.calls.append(("add_column", (workbook_id, name, header, position)))
        return _mutation("add_column", name)

    async def delete_table_column(
        self, workbook_id: str, name: str, header: str
    ) -> TableMutationResult:
        self.calls.append(("delete_column", (workbook_id, name, header)))
        return _mutation("delete_column", name)

    async def create_table(
        self,
        workbook_id: str,
        sheet: str,
        range_or_name: str,
        style: str | None = None,
    ) -> TableMutationResult:
        self.calls.append(("create", (workbook_id, sheet, range_or_name, style)))
        return _mutation("create", range_or_name)

    async def resize_table(
        self, workbook_id: str, name: str, new_range: str
    ) -> TableMutationResult:
        self.calls.append(("resize", (workbook_id, name, new_range)))
        return _mutation("resize", name)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeTablesService()
    table_tools.set_tables_service(fake)  # type: ignore[arg-type]
    yield fake
    table_tools.set_tables_service(None)


# ---------------------------------------------------------------------------
# table: routing
# ---------------------------------------------------------------------------
async def test_table_create_routes(_install_fake: FakeTablesService) -> None:
    await table("create", workbook_id="wb1", sheet="Data", range_or_name="A1:B2")

    name, args = _install_fake.calls[-1]
    assert name == "create"
    assert args == ("wb1", "Data", "A1:B2", None)


async def test_table_resize_routes(_install_fake: FakeTablesService) -> None:
    await table("resize", workbook_id="wb1", name="Sales", new_range="A1:C3")

    name, args = _install_fake.calls[-1]
    assert name == "resize"
    assert args == ("wb1", "Sales", "A1:C3")


async def test_table_append_rows_routes(_install_fake: FakeTablesService) -> None:
    rows = [{"Amount": 100}]
    await table("append-rows", workbook_id="wb1", name="Sales", rows=rows)

    name, args = _install_fake.calls[-1]
    assert name == "append_rows"
    assert args == ("wb1", "Sales", rows)


async def test_table_insert_row_routes(_install_fake: FakeTablesService) -> None:
    await table("insert-row", workbook_id="wb1", name="Sales", at=2, values={"A": 1})

    name, args = _install_fake.calls[-1]
    assert name == "insert_row"
    assert args == ("wb1", "Sales", 2, {"A": 1})


async def test_table_update_rows_routes(_install_fake: FakeTablesService) -> None:
    updates = [{"index": 1, "values": {"A": 2}}]
    await table("update-rows", workbook_id="wb1", name="Sales", updates=updates)

    name, args = _install_fake.calls[-1]
    assert name == "update_rows"
    assert args == ("wb1", "Sales", updates)


async def test_table_add_column_routes(_install_fake: FakeTablesService) -> None:
    await table("add-column", workbook_id="wb1", name="Sales", header="Cost")

    name, args = _install_fake.calls[-1]
    assert name == "add_column"
    assert args == ("wb1", "Sales", "Cost", None)


# ---------------------------------------------------------------------------
# table: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_table_create_missing_sheet_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("create", workbook_id="wb1", range_or_name="A1:B2")


async def test_table_create_missing_range_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("create", workbook_id="wb1", sheet="Data")


async def test_table_resize_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("resize", workbook_id="wb1", new_range="A1:C3")


async def test_table_append_rows_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("append-rows", workbook_id="wb1", rows=[{"A": 1}])


async def test_table_append_rows_missing_rows_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("append-rows", workbook_id="wb1", name="Sales")


async def test_table_insert_row_missing_at_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("insert-row", workbook_id="wb1", name="Sales")


async def test_table_update_rows_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("update-rows", workbook_id="wb1", updates=[{"index": 1}])


async def test_table_update_rows_missing_updates_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("update-rows", workbook_id="wb1", name="Sales")


async def test_table_add_column_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("add-column", workbook_id="wb1", header="Cost")


async def test_table_add_column_missing_header_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("add-column", workbook_id="wb1", name="Sales")


async def test_table_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("append-rows", name="Sales", rows=[{"A": 1}])


async def test_table_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("destroy", workbook_id="wb1", name="Sales")  # type: ignore[arg-type]


async def test_table_delete_column_is_not_a_table_action() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table("delete-column", workbook_id="wb1", name="Sales", header="Cost")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# table_read: routing + validation
# ---------------------------------------------------------------------------
async def test_table_read_list_routes(_install_fake: FakeTablesService) -> None:
    listed = await table_read("list", workbook_id="wb1")

    assert [t.name for t in listed] == ["Sales"]
    name, args = _install_fake.calls[-1]
    assert name == "list"
    assert args == ("wb1",)


async def test_table_read_get_routes(_install_fake: FakeTablesService) -> None:
    info = await table_read("get", workbook_id="wb1", name="Sales")

    assert info.name == "Sales"
    name, args = _install_fake.calls[-1]
    assert name == "get"
    assert args == ("wb1", "Sales")


async def test_table_read_read_rows_routes(_install_fake: FakeTablesService) -> None:
    await table_read("read-rows", workbook_id="wb1", name="Sales", offset=0, limit=50)

    name, args = _install_fake.calls[-1]
    assert name == "read_rows"
    assert args == ("wb1", "Sales", 0, 50)


async def test_table_read_get_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_read("get", workbook_id="wb1")


async def test_table_read_read_rows_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_read("read-rows", workbook_id="wb1")


async def test_table_read_list_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_read("list")


async def test_table_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_read("destroy", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# table_delete_rows: standalone XOR guard (indices XOR where)
# ---------------------------------------------------------------------------
async def test_table_delete_rows_both_selectors_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_delete_rows("wb1", "Sales", indices=[1], where={"A": 1})


async def test_table_delete_rows_neither_selector_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_delete_rows("wb1", "Sales")


async def test_table_delete_rows_indices_only_ok(
    _install_fake: FakeTablesService,
) -> None:
    await table_delete_rows("wb1", "Sales", indices=[2])

    name, args = _install_fake.calls[-1]
    assert name == "delete_rows"
    assert args == ("wb1", "Sales", [2], None)


async def test_table_delete_rows_where_only_ok(_install_fake: FakeTablesService) -> None:
    await table_delete_rows("wb1", "Sales", where={"SKU": "A-1"})

    name, args = _install_fake.calls[-1]
    assert name == "delete_rows"
    assert args == ("wb1", "Sales", None, {"SKU": "A-1"})


# ---------------------------------------------------------------------------
# table_delete_column: standalone, header required (destructive)
# ---------------------------------------------------------------------------
async def test_table_delete_column_routes(_install_fake: FakeTablesService) -> None:
    await table_delete_column("wb1", "Sales", "Cost")

    name, args = _install_fake.calls[-1]
    assert name == "delete_column"
    assert args == ("wb1", "Sales", "Cost")


async def test_table_delete_column_missing_header_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await table_delete_column("wb1", "Sales", "")


# ---------------------------------------------------------------------------
# Annotation triples (WP0 truth matrix; None treated as "not True")
# ---------------------------------------------------------------------------
async def test_table_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["table"].annotations
    assert ann is not None
    assert ann.destructive_hint is not True
    assert ann.read_only_hint is not True
    assert ann.idempotent_hint is not True


async def test_table_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["table_read"].annotations
    assert ann is not None
    assert ann.read_only_hint is True
    assert ann.idempotent_hint is True
    assert ann.destructive_hint is not True


async def test_table_delete_rows_is_destructive_standalone() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["table_delete_rows"].annotations
    assert ann is not None
    assert ann.destructive_hint is True


async def test_table_delete_column_is_destructive_standalone() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["table_delete_column"].annotations
    assert ann is not None
    assert ann.destructive_hint is True


async def test_old_table_tool_names_are_not_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for removed in (
        "tables_list",
        "table_get",
        "table_read_rows",
        "table_append_rows",
        "table_insert_row",
        "table_update_rows",
        "table_add_column",
        "table_create",
        "table_resize",
    ):
        assert removed not in names