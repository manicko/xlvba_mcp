"""Unit tests for the WP1 workbook dispatch tools (plan B2).

These tests never touch COM / Excel. They install fake ``SessionService`` /
``WorkbookModelService`` instances via the ``workbook_tools`` setters and assert
that the flat-signature ``workbook`` / ``workbook_read`` dispatchers route to the
existing handlers, enforce per-action required parameters as structured
:class:`ExcelVbaMcpError` (option (a) — handler validation), and advertise the
truthful annotation triples from the WP0 matrix.
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.workbook_tools import workbook, workbook_read
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    SessionStatus,
    WorkbookInfo,
    WorkbookModel,
)
from excel_vba_mcp.domain.statuses import SessionMode, WorkbookSessionState

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class FakeSessionService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...], dict[str, Any]]] = []
        self._info = WorkbookInfo(
            workbook_id="wb1",
            name="Book.xlsm",
            format="xlsm",
            path="C:/work/wb1.xlsm",
            read_only=False,
            has_vba=True,
            excel_version="16.0",
            session_id="s1",
            mode=SessionMode.EXECUTE,
        )
        self._status = SessionStatus(
            session_id="s1",
            workbook_id="wb1",
            state=WorkbookSessionState.RUNNING,
            mode=SessionMode.EXECUTE,
            created_at="2026-01-01T00:00:00Z",
            updated_at="2026-01-01T00:00:00Z",
        )

    async def open_workbook(
        self, path: str, *, mode: SessionMode, read_only: bool
    ) -> WorkbookInfo:
        self.calls.append(("open", (path,), {"mode": mode, "read_only": read_only}))
        return self._info

    async def close_workbook(
        self,
        workbook_id: str,
        *,
        save_changes: bool,
        discard_changes: bool = False,
    ) -> SessionStatus:
        self.calls.append(
            (
                "close",
                (workbook_id,),
                {"save_changes": save_changes, "discard_changes": discard_changes},
            )
        )
        return self._status

    async def save_workbook(self, workbook_id: str) -> WorkbookInfo:
        self.calls.append(("save", (workbook_id,), {}))
        return self._info

    async def get_status(self, workbook_id: str | None = None) -> SessionStatus:
        self.calls.append(("status", (workbook_id,), {}))
        return self._status

    async def list_workbooks(self) -> list[WorkbookInfo]:
        self.calls.append(("list", (), {}))
        return [self._info]


class FakeModelService:
    def __init__(self, model: WorkbookModel) -> None:
        self._model = model
        self.calls: list[tuple[str, bool]] = []

    async def get_workbook_model(
        self, workbook_id: str, *, include_vba: bool = True
    ) -> WorkbookModel:
        self.calls.append((workbook_id, include_vba))
        return self._model


def _empty_model() -> WorkbookModel:
    from datetime import UTC, datetime

    now = datetime.now(UTC)
    return WorkbookModel(workbook=FakeSessionService()._info, created_at=now, updated_at=now)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fakes() -> Any:
    fake = FakeSessionService()
    model = FakeModelService(_empty_model())
    workbook_tools.set_session_service(fake)  # type: ignore[arg-type]
    workbook_tools.set_workbook_model_service(model)  # type: ignore[arg-type]
    yield fake, model
    workbook_tools.set_session_service(None)
    workbook_tools.set_workbook_model_service(None)


# ---------------------------------------------------------------------------
# workbook: routing
# ---------------------------------------------------------------------------
async def test_workbook_open_routes_to_open(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    info = await workbook("open", path="C:/data/Book.xlsm")

    assert info.workbook_id == "wb1"
    name, args, kwargs = fake.calls[-1]
    assert name == "open"
    assert args == ("C:/data/Book.xlsm",)
    assert kwargs == {"mode": SessionMode.EXECUTE, "read_only": False}


async def test_workbook_close_routes_to_close(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    await workbook("close", workbook_id="wb1", discard_changes=True)

    name, args, kwargs = fake.calls[-1]
    assert name == "close"
    assert args == ("wb1",)
    assert kwargs == {"save_changes": False, "discard_changes": True}


async def test_workbook_save_routes_to_save(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    await workbook("save", workbook_id="wb1")

    name, args, _kwargs = fake.calls[-1]
    assert name == "save"
    assert args == ("wb1",)


# ---------------------------------------------------------------------------
# workbook: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_workbook_open_missing_path_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook("open")


async def test_workbook_close_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook("close")


async def test_workbook_save_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook("save")


async def test_workbook_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook("rename")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# workbook_read: routing + validation
# ---------------------------------------------------------------------------
async def test_workbook_read_status_without_id(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    await workbook_read("status")

    name, args, _kwargs = fake.calls[-1]
    assert name == "status"
    assert args == (None,)


async def test_workbook_read_status_with_id(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    await workbook_read("status", workbook_id="wb1")

    name, args, _kwargs = fake.calls[-1]
    assert args == ("wb1",)


async def test_workbook_read_list(_install_fakes: tuple[Any, Any]) -> None:
    fake, _model = _install_fakes
    listed = await workbook_read("list")

    assert [w.workbook_id for w in listed] == ["wb1"]
    name, _args, _kwargs = fake.calls[-1]
    assert name == "list"


async def test_workbook_read_model_routes(_install_fakes: tuple[Any, Any]) -> None:
    _fake, model = _install_fakes
    result = await workbook_read("model", workbook_id="wb1", include_vba=False)

    assert isinstance(result, WorkbookModel)
    assert model.calls == [("wb1", False)]


async def test_workbook_read_model_missing_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_read("model")


async def test_workbook_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await workbook_read("destroy")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Annotation triples (WP0 truth matrix; None treated as "not True")
# ---------------------------------------------------------------------------
async def test_workbook_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["workbook"].annotations
    assert ann is not None
    assert ann.destructive_hint is True
    assert ann.read_only_hint is not True
    assert ann.idempotent_hint is not True


async def test_workbook_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["workbook_read"].annotations
    assert ann is not None
    assert ann.read_only_hint is True
    assert ann.idempotent_hint is True
    assert ann.destructive_hint is not True


async def test_workbook_save_copy_is_destructive_standalone() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["workbook_save_copy"].annotations
    assert ann is not None
    assert ann.destructive_hint is True


async def test_old_workbook_tool_names_are_not_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for removed in (
        "excel_open",
        "excel_close",
        "excel_save",
        "excel_get_status",
        "excel_list_workbooks",
        "workbook_get_model",
    ):
        assert removed not in names