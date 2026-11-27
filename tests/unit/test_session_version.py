"""Unit tests for the WP-A workbook content-version primitive.

These tests never open Excel. They inject a fake STA worker (which runs the
submitted callable inline) and a real :class:`ExcelAppAdapter` whose ``_app`` is
a fake ``Application`` object, mirroring ``test_session_service``. They prove:

* a freshly opened workbook has version ``0``;
* ``bump_version`` increments the version by one (and only on demand);
* an unknown ``workbook_id`` raises :class:`WorkbookNotFoundError`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ErrorCode, WorkbookNotFoundError
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no Excel / no COM)
# ---------------------------------------------------------------------------
class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeWorkbook:
    def __init__(self, app: FakeApp, path: str, *, read_only: bool) -> None:
        self._app = app
        self.FullName = path
        self.Name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.ReadOnly = read_only
        self.Saved = True
        self.Worksheets = [FakeSheet("Sheet1")]

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM casing
        self.Saved = SaveChanges

    def Save(self) -> None:  # noqa: N802 - COM casing
        self.Saved = True

    @property
    def Application(self) -> FakeApp:  # noqa: N802 - COM casing
        return self._app


class FakeWorkbooks:
    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        wb = FakeWorkbook(self._app, path, read_only=bool(kwargs.get("ReadOnly", False)))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)

    def __len__(self) -> int:
        return len(self.items)


class FakeApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(self)


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self, status: WorkerStatus | None = None) -> None:
        self.submits: list[dict[str, Any]] = []
        self._status = status or WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append({"name": getattr(fn, "__name__", "?"), "section": section})
        return fn()

    def get_status(self) -> WorkerStatus:
        return self._status


# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _make_service(tmp_path: Path) -> tuple[SessionService, Path]:
    original = tmp_path / "Book.xlsm"
    original.write_bytes(b"original-bytes")
    path_policy = PathPolicy({str(tmp_path): str(tmp_path)})
    adapter = ExcelAppAdapter()
    adapter._app = FakeApp()  # inject without DispatchEx
    service = SessionService(
        FakeWorker(),
        adapter,
        path_policy,
        ApprovalPolicy(),
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )
    return service, original


# ---------------------------------------------------------------------------
# Version primitive
# ---------------------------------------------------------------------------
async def test_version_starts_at_zero_and_bumps(tmp_path: Path) -> None:
    service, original = _make_service(tmp_path)

    info = await service.open_workbook(str(original), mode=SessionMode.INSPECT, read_only=True)

    assert service.get_version(info.workbook_id) == 0

    service.bump_version(info.workbook_id)
    assert service.get_version(info.workbook_id) == 1

    service.bump_version(info.workbook_id)
    assert service.get_version(info.workbook_id) == 2


async def test_version_unknown_workbook_raises(tmp_path: Path) -> None:
    service, _original = _make_service(tmp_path)

    with pytest.raises(WorkbookNotFoundError) as exc_info:
        service.get_version("nope")
    assert exc_info.value.code is ErrorCode.WORKBOOK_NOT_FOUND

    with pytest.raises(WorkbookNotFoundError) as exc_info:
        service.bump_version("nope")
    assert exc_info.value.code is ErrorCode.WORKBOOK_NOT_FOUND