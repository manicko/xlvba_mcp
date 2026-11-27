"""Unit tests for the base workbook/session MCP tools (task_012).

These tests never touch COM / Excel. They install a fake ``SessionService`` via
``workbook_tools.set_session_service`` and assert that each thin tool validates
its arguments and delegates to the service. A static check asserts the ``api``
package source contains no ``pythoncom`` / ``win32com`` import (ARCH-009).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.api import workbook_tools
from excel_vba_mcp.api.workbook_tools import (
    excel_close,
    excel_get_status,
    excel_list_workbooks,
    excel_open,
    excel_save,
)
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import SessionStatus, WorkbookInfo
from excel_vba_mcp.domain.statuses import (
    SessionMode,
    WorkbookSessionState,
)

pytestmark = pytest.mark.anyio

_API_DIR = Path(__file__).resolve().parents[2] / "src" / "excel_vba_mcp" / "api"


# ---------------------------------------------------------------------------
# Fake service
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
            read_only=True,
            has_vba=True,
            excel_version="16.0",
            session_id="s1",
            mode=SessionMode.INSPECT,
        )
        self._status = SessionStatus(
            session_id="s1",
            workbook_id="wb1",
            state=WorkbookSessionState.RUNNING,
            mode=SessionMode.INSPECT,
            created_at="2026-01-01T00:00:00Z",
            updated_at="2026-01-01T00:00:00Z",
        )

    async def open_workbook(self, path: str, *, mode: SessionMode, read_only: bool) -> WorkbookInfo:
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


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake_service() -> Any:
    fake = FakeSessionService()
    workbook_tools.set_session_service(fake)  # type: ignore[arg-type]
    yield fake
    workbook_tools.set_session_service(None)


# ---------------------------------------------------------------------------
# Delegation + defaults
# ---------------------------------------------------------------------------
async def test_excel_open_delegates_with_mode_and_read_only(
    _install_fake_service: FakeSessionService,
) -> None:
    info = await excel_open("C:/data/Book.xlsm", mode=SessionMode.EXECUTE, read_only=True)

    assert info.workbook_id == "wb1"
    name, args, kwargs = _install_fake_service.calls[-1]
    assert name == "open"
    assert args == ("C:/data/Book.xlsm",)
    assert kwargs == {"mode": SessionMode.EXECUTE, "read_only": True}


async def test_excel_open_defaults_to_execute_writable(
    _install_fake_service: FakeSessionService,
) -> None:
    await excel_open("C:/data/Book.xlsm")

    _name, _args, kwargs = _install_fake_service.calls[-1]
    assert kwargs == {"mode": SessionMode.EXECUTE, "read_only": False}


async def test_excel_close_defaults_to_no_save(
    _install_fake_service: FakeSessionService,
) -> None:
    status = await excel_close("wb1")

    assert status.workbook_id == "wb1"
    name, args, kwargs = _install_fake_service.calls[-1]
    assert name == "close"
    assert args == ("wb1",)
    assert kwargs == {"save_changes": False, "discard_changes": False}


async def test_excel_close_passes_save_changes(
    _install_fake_service: FakeSessionService,
) -> None:
    await excel_close("wb1", save_changes=True)
    _name, _args, kwargs = _install_fake_service.calls[-1]
    assert kwargs == {"save_changes": True, "discard_changes": False}


async def test_excel_close_passes_discard_changes(
    _install_fake_service: FakeSessionService,
) -> None:
    await excel_close("wb1", discard_changes=True)
    _name, _args, kwargs = _install_fake_service.calls[-1]
    assert kwargs == {"save_changes": False, "discard_changes": True}


async def test_excel_save_delegates(_install_fake_service: FakeSessionService) -> None:
    info = await excel_save("wb1")
    assert info.workbook_id == "wb1"
    name, args, _kwargs = _install_fake_service.calls[-1]
    assert name == "save"
    assert args == ("wb1",)


async def test_excel_get_status_without_id(_install_fake_service: FakeSessionService) -> None:
    status = await excel_get_status()
    assert isinstance(status, SessionStatus)
    name, args, _kwargs = _install_fake_service.calls[-1]
    assert name == "status"
    assert args == (None,)


async def test_excel_get_status_with_id(_install_fake_service: FakeSessionService) -> None:
    await excel_get_status("wb1")
    _name, args, _kwargs = _install_fake_service.calls[-1]
    assert args == ("wb1",)


async def test_excel_list_workbooks(_install_fake_service: FakeSessionService) -> None:
    listed = await excel_list_workbooks()
    assert [w.workbook_id for w in listed] == ["wb1"]
    name, _args, _kwargs = _install_fake_service.calls[-1]
    assert name == "list"


# ---------------------------------------------------------------------------
# Argument validation
# ---------------------------------------------------------------------------
async def test_excel_open_rejects_empty_path() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await excel_open("")


async def test_excel_close_rejects_empty_id() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await excel_close("")


async def test_excel_save_rejects_empty_id() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await excel_save("   ")


async def test_excel_get_status_rejects_empty_id() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await excel_get_status("")


async def test_tools_raise_when_service_unset() -> None:
    workbook_tools.set_session_service(None)
    with pytest.raises(ExcelVbaMcpError):
        await excel_list_workbooks()


# ---------------------------------------------------------------------------
# Static guard: api/* never imports COM (ARCH-009)
# ---------------------------------------------------------------------------
_FORBIDDEN_IMPORTS = re.compile(r"\b(import|from)\s+(pythoncom|win32com)\b")


def test_api_package_never_imports_com() -> None:
    offenders: list[str] = []
    for py in _API_DIR.rglob("*.py"):
        text = py.read_text(encoding="utf-8")
        if _FORBIDDEN_IMPORTS.search(text):
            offenders.append(str(py))
    assert offenders == [], f"api/* must not import COM: {offenders}"
