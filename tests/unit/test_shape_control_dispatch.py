"""Unit tests for the WP5 shape/control dispatch tools (plan B6).

These tests never touch COM / Excel. They install fake ``ShapesService`` /
``ControlsService`` instances via the tool-module setters and assert that the
flat-signature ``shape`` / ``control`` dispatchers route to the existing
handlers, enforce per-action required parameters as structured
:class:`ExcelVbaMcpError` (option (a) — handler validation), and advertise the
truthful annotation triples from the WP0 matrix (both are (F,F,T): neither
read-only nor destructive, idempotent only).
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import control_tools, shape_tools
from excel_vba_mcp.api.control_tools import control
from excel_vba_mcp.api.shape_tools import shape
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import ControlEntry, ShapeInfo
from excel_vba_mcp.domain.statuses import ControlKind, ShapeKind

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _shape_info(name: str) -> ShapeInfo:
    return ShapeInfo(name=name, kind=ShapeKind.AUTOSHAPE, on_action=None)


def _control_entry(name: str) -> ControlEntry:
    return ControlEntry(name=name, kind=ControlKind.FORM_CONTROL)


class FakeShapesService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_shapes(self, workbook_id: str) -> list[ShapeInfo]:
        self.calls.append(("list", (workbook_id,)))
        return [_shape_info("rect1")]

    async def get_shape(self, workbook_id: str, name: str) -> ShapeInfo:
        self.calls.append(("get", (workbook_id, name)))
        return _shape_info(name)

    async def set_on_action(self, workbook_id: str, name: str, macro_name: str) -> ShapeInfo:
        self.calls.append(("set_on_action", (workbook_id, name, macro_name)))
        return _shape_info(name)


class FakeControlsService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_controls(self, workbook_id: str) -> list[ControlEntry]:
        self.calls.append(("list", (workbook_id,)))
        return [_control_entry("cmd1")]

    async def get_control(self, workbook_id: str, name: str) -> ControlEntry:
        self.calls.append(("get", (workbook_id, name)))
        return _control_entry(name)

    async def set_on_action(self, workbook_id: str, name: str, macro_name: str) -> ControlEntry:
        self.calls.append(("set_on_action", (workbook_id, name, macro_name)))
        return _control_entry(name)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fakes() -> Any:
    shapes = FakeShapesService()
    controls = FakeControlsService()
    shape_tools.set_shapes_service(shapes)  # type: ignore[arg-type]
    control_tools.set_controls_service(controls)  # type: ignore[arg-type]
    yield shapes, controls
    shape_tools.set_shapes_service(None)
    control_tools.set_controls_service(None)


# ---------------------------------------------------------------------------
# shape: routing
# ---------------------------------------------------------------------------
async def test_shape_list_routes(_install_fakes: tuple[Any, Any]) -> None:
    shapes, _ = _install_fakes
    listed = await shape("list", workbook_id="wb1")

    assert [s.name for s in listed] == ["rect1"]
    verb, args = shapes.calls[-1]
    assert verb == "list"
    assert args == ("wb1",)


async def test_shape_get_routes(_install_fakes: tuple[Any, Any]) -> None:
    shapes, _ = _install_fakes
    info = await shape("get", workbook_id="wb1", name="rect1")

    assert info.name == "rect1"
    verb, args = shapes.calls[-1]
    assert verb == "get"
    assert args == ("wb1", "rect1")


async def test_shape_set_on_action_routes(_install_fakes: tuple[Any, Any]) -> None:
    shapes, _ = _install_fakes
    info = await shape(
        "set-on-action", workbook_id="wb1", name="rect1", macro_name="SaveData"
    )

    assert info.name == "rect1"
    verb, args = shapes.calls[-1]
    assert verb == "set_on_action"
    assert args == ("wb1", "rect1", "SaveData")


# ---------------------------------------------------------------------------
# shape: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_shape_get_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await shape("get", workbook_id="wb1")


async def test_shape_set_on_action_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await shape("set-on-action", workbook_id="wb1", macro_name="SaveData")


async def test_shape_set_on_action_missing_macro_name_clears(
    _install_fakes: tuple[Any, Any],
) -> None:
    # macro_name is optional: an omitted value clears the assignment (handler
    # contract) rather than raising.
    shapes, _ = _install_fakes
    await shape("set-on-action", workbook_id="wb1", name="rect1")

    verb, args = shapes.calls[-1]
    assert verb == "set_on_action"
    assert args == ("wb1", "rect1", "")


async def test_shape_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await shape("list")


async def test_shape_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await shape("destroy", workbook_id="wb1", name="rect1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# control: routing
# ---------------------------------------------------------------------------
async def test_control_list_routes(_install_fakes: tuple[Any, Any]) -> None:
    _, controls = _install_fakes
    listed = await control("list", workbook_id="wb1")

    assert [c.name for c in listed] == ["cmd1"]
    verb, args = controls.calls[-1]
    assert verb == "list"
    assert args == ("wb1",)


async def test_control_get_routes(_install_fakes: tuple[Any, Any]) -> None:
    _, controls = _install_fakes
    entry = await control("get", workbook_id="wb1", name="cmd1")

    assert entry.name == "cmd1"
    verb, args = controls.calls[-1]
    assert verb == "get"
    assert args == ("wb1", "cmd1")


async def test_control_set_on_action_routes(_install_fakes: tuple[Any, Any]) -> None:
    _, controls = _install_fakes
    entry = await control(
        "set-on-action", workbook_id="wb1", name="cmd1", macro_name="RunReport"
    )

    assert entry.name == "cmd1"
    verb, args = controls.calls[-1]
    assert verb == "set_on_action"
    assert args == ("wb1", "cmd1", "RunReport")


# ---------------------------------------------------------------------------
# control: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_control_get_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await control("get", workbook_id="wb1")


async def test_control_set_on_action_missing_name_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await control("set-on-action", workbook_id="wb1", macro_name="RunReport")


async def test_control_set_on_action_missing_macro_name_clears(
    _install_fakes: tuple[Any, Any],
) -> None:
    # macro_name is optional: an omitted value clears the assignment (handler
    # contract) rather than raising.
    _, controls = _install_fakes
    await control("set-on-action", workbook_id="wb1", name="cmd1")

    verb, args = controls.calls[-1]
    assert verb == "set_on_action"
    assert args == ("wb1", "cmd1", "")


async def test_control_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await control("list")


async def test_control_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await control("destroy", workbook_id="wb1", name="cmd1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Annotation triples (WP0 truth matrix: shape/control == (F, F, T))
# ---------------------------------------------------------------------------
async def test_shape_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["shape"].annotations
    assert ann is not None
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_control_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["control"].annotations
    assert ann is not None
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_old_shape_control_tool_names_are_not_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for removed in (
        "shapes_list",
        "shape_get",
        "shape_set_on_action",
        "controls_list",
        "control_get",
        "control_set_on_action",
    ):
        assert removed not in names