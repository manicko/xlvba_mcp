"""Unit tests for the WP8 pivot dispatch tools (plan B9).

These tests never touch COM / Excel. They install a fake ``PivotTablesService``
via the ``pivot_tools`` setter and assert that the flat-signature
``pivot_read`` / ``pivot_refresh`` / ``pivot_field`` dispatchers route to the
existing per-verb handlers, enforce per-action required parameters as structured
:class:`ExcelVbaMcpError` (option (a) — handler validation), and advertise the
truthful annotation triples from the WP0 matrix: ``pivot_read`` is (T,F,T),
``pivot_refresh`` is (F,F,F) — NOT idempotent because a re-refresh may read
different source data — and ``pivot_field`` is (F,T,T), the intentional dual-hint
(destructive via ``field-set-layout(delete=True)`` + idempotent for the same
target state, R-Q4).
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import pivot_tools
from excel_vba_mcp.api.pivot_tools import pivot_field, pivot_read, pivot_refresh
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFieldProperties,
    PivotFilterSpec,
    PivotInventory,
    PivotRefreshResult,
    PivotTableInfo,
)
from excel_vba_mcp.domain.statuses import (
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterRequestKind,
    PivotRefreshMethod,
    PivotSourceType,
)

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _inventory() -> PivotInventory:
    return PivotInventory(workbook_id="wb1")


def _table_info() -> PivotTableInfo:
    return PivotTableInfo(
        name="pt", sheet_name="Sheet1", sheet_index=1, index=1, cache_index=1
    )


def _cache_info() -> PivotCacheInfo:
    return PivotCacheInfo(cache_index=1, source_type=PivotSourceType.DATABASE)


def _field_info() -> PivotFieldInfo:
    return PivotFieldInfo(source_name="Region")


def _field_filters() -> PivotFieldFilters:
    return PivotFieldFilters(pivot_table="pt", field="Region")


def _refresh_result(method: PivotRefreshMethod) -> PivotRefreshResult:
    return PivotRefreshResult(
        pivot_table="pt", cache_index=1, method=method, refreshed=True
    )


class FakePivotTablesService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_pivot_tables(self, workbook_id: str) -> PivotInventory:
        self.calls.append(("list_pivot_tables", (workbook_id,)))
        return _inventory()

    async def get_pivot_table(self, workbook_id: str, pivot_table: str) -> PivotTableInfo:
        self.calls.append(("get_pivot_table", (workbook_id, pivot_table)))
        return _table_info()

    async def get_pivot_cache(self, workbook_id: str, cache_index: int) -> PivotCacheInfo:
        self.calls.append(("get_pivot_cache", (workbook_id, cache_index)))
        return _cache_info()

    async def list_pivot_fields(
        self, workbook_id: str, pivot_table: str
    ) -> list[PivotFieldInfo]:
        self.calls.append(("list_pivot_fields", (workbook_id, pivot_table)))
        return [_field_info()]

    async def get_pivot_field_filters(
        self, workbook_id: str, pivot_table: str, field: str
    ) -> PivotFieldFilters:
        self.calls.append(("get_pivot_field_filters", (workbook_id, pivot_table, field)))
        return _field_filters()

    async def refresh_pivot_table(
        self, workbook_id: str, pivot_table: str
    ) -> PivotRefreshResult:
        self.calls.append(("refresh_pivot_table", (workbook_id, pivot_table)))
        return _refresh_result(PivotRefreshMethod.REFRESH_TABLE)

    async def refresh_pivot_cache(
        self, workbook_id: str, cache_index: int
    ) -> PivotRefreshResult:
        self.calls.append(("refresh_pivot_cache", (workbook_id, cache_index)))
        return _refresh_result(PivotRefreshMethod.REFRESH_CACHE)

    async def set_pivot_field_layout(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        orientation: PivotFieldOrientation,
        *,
        delete: bool = False,
    ) -> PivotFieldInfo:
        self.calls.append(
            ("set_pivot_field_layout", (workbook_id, pivot_table, field, orientation, delete))
        )
        return _field_info()

    async def set_pivot_field_position(
        self, workbook_id: str, pivot_table: str, field: str, position: int
    ) -> PivotFieldInfo:
        self.calls.append(
            ("set_pivot_field_position", (workbook_id, pivot_table, field, position))
        )
        return _field_info()

    async def set_pivot_field_aggregation(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        function: PivotAggregation,
        *,
        caption: str | None = None,
    ) -> PivotFieldInfo:
        self.calls.append(
            (
                "set_pivot_field_aggregation",
                (workbook_id, pivot_table, field, function, caption),
            )
        )
        return _field_info()

    async def set_pivot_field_properties(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        properties: PivotFieldProperties,
    ) -> PivotFieldInfo:
        self.calls.append(
            ("set_pivot_field_properties", (workbook_id, pivot_table, field, properties))
        )
        return _field_info()

    async def apply_pivot_field_filter(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str,
        spec: PivotFilterSpec,
    ) -> PivotFieldFilters:
        self.calls.append(
            ("apply_pivot_field_filter", (workbook_id, pivot_table, field, spec))
        )
        return _field_filters()

    async def clear_pivot_field_filters(
        self,
        workbook_id: str,
        pivot_table: str,
        field: str | None,
        *,
        scope: str = "field",
    ) -> PivotFieldFilters:
        self.calls.append(
            ("clear_pivot_field_filters", (workbook_id, pivot_table, field, scope))
        )
        return _field_filters()


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakePivotTablesService()
    pivot_tools.set_pivot_tables_service(fake)  # type: ignore[arg-type]
    yield fake
    pivot_tools.set_pivot_tables_service(None)


# ---------------------------------------------------------------------------
# pivot_read: routing
# ---------------------------------------------------------------------------
async def test_pivot_read_tables_list_routes(_install_fake: FakePivotTablesService) -> None:
    await pivot_read("tables-list", workbook_id="wb1")

    verb, args = _install_fake.calls[-1]
    assert verb == "list_pivot_tables"
    assert args == ("wb1",)


async def test_pivot_read_table_get_routes(_install_fake: FakePivotTablesService) -> None:
    await pivot_read("table-get", workbook_id="wb1", pivot_table="pt")

    verb, args = _install_fake.calls[-1]
    assert verb == "get_pivot_table"
    assert args == ("wb1", "pt")


async def test_pivot_read_cache_get_routes(_install_fake: FakePivotTablesService) -> None:
    await pivot_read("cache-get", workbook_id="wb1", cache_index=1)

    verb, args = _install_fake.calls[-1]
    assert verb == "get_pivot_cache"
    assert args == ("wb1", 1)


async def test_pivot_read_fields_list_routes(_install_fake: FakePivotTablesService) -> None:
    await pivot_read("fields-list", workbook_id="wb1", pivot_table="pt")

    verb, args = _install_fake.calls[-1]
    assert verb == "list_pivot_fields"
    assert args == ("wb1", "pt")


async def test_pivot_read_field_get_filters_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_read(
        "field-get-filters", workbook_id="wb1", pivot_table="pt", field="Region"
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "get_pivot_field_filters"
    assert args == ("wb1", "pt", "Region")


# ---------------------------------------------------------------------------
# pivot_read: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_pivot_read_cache_get_missing_cache_index_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("cache-get", workbook_id="wb1")


async def test_pivot_read_table_get_missing_pivot_table_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("table-get", workbook_id="wb1")


async def test_pivot_read_fields_list_missing_pivot_table_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("fields-list", workbook_id="wb1")


async def test_pivot_read_field_get_filters_missing_pivot_table_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("field-get-filters", workbook_id="wb1", field="Region")


async def test_pivot_read_field_get_filters_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("field-get-filters", workbook_id="wb1", pivot_table="pt")


async def test_pivot_read_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("tables-list")


async def test_pivot_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_read("destroy", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# pivot_refresh: routing
# ---------------------------------------------------------------------------
async def test_pivot_refresh_refresh_table_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_refresh("refresh-table", workbook_id="wb1", pivot_table="pt")

    verb, args = _install_fake.calls[-1]
    assert verb == "refresh_pivot_table"
    assert args == ("wb1", "pt")


async def test_pivot_refresh_refresh_cache_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_refresh("refresh-cache", workbook_id="wb1", cache_index=1)

    verb, args = _install_fake.calls[-1]
    assert verb == "refresh_pivot_cache"
    assert args == ("wb1", 1)


# ---------------------------------------------------------------------------
# pivot_refresh: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_pivot_refresh_refresh_table_missing_pivot_table_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_refresh("refresh-table", workbook_id="wb1")


async def test_pivot_refresh_refresh_cache_missing_cache_index_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_refresh("refresh-cache", workbook_id="wb1")


async def test_pivot_refresh_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_refresh("destroy", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# pivot_field: routing
# ---------------------------------------------------------------------------
async def test_pivot_field_set_layout_routes(_install_fake: FakePivotTablesService) -> None:
    await pivot_field(
        "field-set-layout", workbook_id="wb1", pivot_table="pt", field="Region",
        orientation="row",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_pivot_field_layout"
    assert args == ("wb1", "pt", "Region", PivotFieldOrientation.ROW, False)


async def test_pivot_field_set_layout_delete_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_field(
        "field-set-layout", workbook_id="wb1", pivot_table="pt", field="Region",
        orientation="hidden", delete=True,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_pivot_field_layout"
    assert args == ("wb1", "pt", "Region", PivotFieldOrientation.HIDDEN, True)


async def test_pivot_field_set_position_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_field(
        "field-set-position", workbook_id="wb1", pivot_table="pt", field="Region",
        position=2,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_pivot_field_position"
    assert args == ("wb1", "pt", "Region", 2)


async def test_pivot_field_set_aggregation_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_field(
        "field-set-aggregation", workbook_id="wb1", pivot_table="pt", field="Amount",
        function="sum",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_pivot_field_aggregation"
    assert args == ("wb1", "pt", "Amount", PivotAggregation.SUM, None)


async def test_pivot_field_set_properties_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    patch = PivotFieldProperties(caption="Total")
    await pivot_field(
        "field-set-properties", workbook_id="wb1", pivot_table="pt", field="Amount",
        properties=patch,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_pivot_field_properties"
    assert args == ("wb1", "pt", "Amount", patch)


async def test_pivot_field_apply_filter_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    spec = PivotFilterSpec(kind=PivotFilterRequestKind.MANUAL, visible_items=["East"])
    await pivot_field(
        "field-apply-filter", workbook_id="wb1", pivot_table="pt", field="Region",
        spec=spec,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "apply_pivot_field_filter"
    assert args == ("wb1", "pt", "Region", spec)


async def test_pivot_field_clear_filters_field_scope_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    await pivot_field(
        "field-clear-filters", workbook_id="wb1", pivot_table="pt", field="Region"
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "clear_pivot_field_filters"
    assert args == ("wb1", "pt", "Region", "field")


async def test_pivot_field_clear_filters_table_scope_routes(
    _install_fake: FakePivotTablesService,
) -> None:
    # Table scope does not require a field (preserved behaviour).
    await pivot_field(
        "field-clear-filters", workbook_id="wb1", pivot_table="pt", scope="table"
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "clear_pivot_field_filters"
    assert args == ("wb1", "pt", None, "table")


# ---------------------------------------------------------------------------
# pivot_field: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_pivot_field_set_layout_missing_pivot_table_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-set-layout", workbook_id="wb1", field="Region", orientation="row"
        )


async def test_pivot_field_set_layout_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-set-layout", workbook_id="wb1", pivot_table="pt", orientation="row"
        )


async def test_pivot_field_set_position_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-set-position", workbook_id="wb1", pivot_table="pt", position=2
        )


async def test_pivot_field_set_aggregation_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-set-aggregation", workbook_id="wb1", pivot_table="pt", function="sum"
        )


async def test_pivot_field_set_properties_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-set-properties", workbook_id="wb1", pivot_table="pt",
            properties=PivotFieldProperties(caption="Total"),
        )


async def test_pivot_field_apply_filter_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-apply-filter", workbook_id="wb1", pivot_table="pt",
            spec=PivotFilterSpec(kind=PivotFilterRequestKind.MANUAL),
        )


async def test_pivot_field_clear_filters_field_scope_missing_field_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "field-clear-filters", workbook_id="wb1", pivot_table="pt", scope="field"
        )


async def test_pivot_field_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await pivot_field(
            "destroy", workbook_id="wb1", pivot_table="pt", field="Region"
        )  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Registration + annotation triples (WP0 truth matrix)
# ---------------------------------------------------------------------------
async def test_pivot_tools_registered_old_names_removed() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for present in ("pivot_read", "pivot_refresh", "pivot_field"):
        assert present in names
    for removed in (
        "pivot_tables_list",
        "pivot_table_get",
        "pivot_cache_get",
        "pivot_fields_list",
        "pivot_field_get_filters",
        "pivot_table_refresh",
        "pivot_cache_refresh",
        "pivot_field_set_layout",
        "pivot_field_set_position",
        "pivot_field_set_aggregation",
        "pivot_field_set_properties",
        "pivot_field_apply_filter",
        "pivot_field_clear_filters",
    ):
        assert removed not in names


async def test_pivot_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["pivot_read"].annotations
    assert ann is not None
    # (T, F, T): read-only and idempotent.
    assert ann.read_only_hint is True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_pivot_refresh_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["pivot_refresh"].annotations
    assert ann is not None
    # (F, F, T): not read-only, not destructive, idempotent.
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_pivot_field_annotation_dual_hint() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["pivot_field"].annotations
    assert ann is not None
    # (F, T, T): intentional dual-hint (destructive via delete=True + idempotent).
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is True