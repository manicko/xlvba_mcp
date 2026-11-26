"""Unit tests for the U6-E agent-intent ``confirm`` gates.

``snapshot_restore`` and ``excel_recover`` are destructive to working copies, so
both are gated by ``confirm`` (default ``False``). A missing or ``False``
``confirm`` must raise ``CONFIRMATION_REQUIRED`` **before** any destructive
action, and the delegate (fake snapshot/session service) must **not** be called
(reject-before-action).

These tests never touch COM / Excel: they install fakes via
``test_tools.set_snapshot_service`` / ``workbook_tools.set_session_service`` and
record whether the delegate was invoked.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_vba_mcp.api import test_tools, workbook_tools
from excel_vba_mcp.api.recovery_tools import excel_recover
from excel_vba_mcp.api.test_tools import snapshot_restore
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import WorkbookInfo
from excel_vba_mcp.domain.statuses import SessionMode

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _info() -> WorkbookInfo:
    return WorkbookInfo(
        workbook_id="wb1",
        name="Book.xlsm",
        format="xlsm",
        path="C:/work/wb1.xlsm",
        source_path="C:/data/Book.xlsm",
        read_only=False,
        has_vba=True,
        excel_version="16.0",
        session_id="s1",
        mode=SessionMode.EXECUTE,
    )


class _FakeSnapshotService:
    """Records ``restore_snapshot`` calls; returns a canned WorkbookInfo."""

    def __init__(self) -> None:
        self.restore_calls: list[tuple[str, str | None]] = []

    async def restore_snapshot(
        self, snapshot_id: str, *, workbook_id: str | None = None
    ) -> WorkbookInfo:
        self.restore_calls.append((snapshot_id, workbook_id))
        return _info()


class _FakeSessionService:
    """Records ``recover_owned_process`` calls; returns a canned outcome dict."""

    def __init__(self) -> None:
        self.recover_calls: list[str] = []

    async def recover_owned_process(self, *, reason: str = "agent_directed") -> dict[str, Any]:
        self.recover_calls.append(reason)
        return {
            "ok": True,
            "terminated": True,
            "confirmed": True,
            "recovery_possible": True,
            "process_restarted": True,
            "prior_job": "none",
            "what_lost": "working copies only; originals are never modified",
            "reason": reason,
        }


# ---------------------------------------------------------------------------
# snapshot_restore
# ---------------------------------------------------------------------------
async def test_snapshot_restore_missing_confirm_rejects_before_action() -> None:
    fake = _FakeSnapshotService()
    test_tools.set_snapshot_service(fake)  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await snapshot_restore("snap-1", "wb1")
    finally:
        test_tools.set_snapshot_service(None)

    assert excinfo.value.code is ErrorCode.CONFIRMATION_REQUIRED
    assert excinfo.value.context.get("tool") == "snapshot_restore"
    assert "discard" in excinfo.value.context
    assert fake.restore_calls == [], "delegate must not be called on reject"


async def test_snapshot_restore_false_confirm_rejects_before_action() -> None:
    fake = _FakeSnapshotService()
    test_tools.set_snapshot_service(fake)  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await snapshot_restore("snap-1", "wb1", confirm=False)
    finally:
        test_tools.set_snapshot_service(None)

    assert excinfo.value.code is ErrorCode.CONFIRMATION_REQUIRED
    assert fake.restore_calls == [], "delegate must not be called on reject"


async def test_snapshot_restore_true_confirm_delegates() -> None:
    fake = _FakeSnapshotService()
    test_tools.set_snapshot_service(fake)  # type: ignore[arg-type]
    try:
        result = await snapshot_restore("snap-1", "wb1", confirm=True)
    finally:
        test_tools.set_snapshot_service(None)

    assert isinstance(result, WorkbookInfo)
    assert fake.restore_calls == [("snap-1", "wb1")]


# ---------------------------------------------------------------------------
# excel_recover
# ---------------------------------------------------------------------------
async def test_excel_recover_missing_confirm_rejects_before_action() -> None:
    fake = _FakeSessionService()
    workbook_tools.set_session_service(fake)  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await excel_recover()
    finally:
        workbook_tools.set_session_service(None)

    assert excinfo.value.code is ErrorCode.CONFIRMATION_REQUIRED
    assert excinfo.value.context.get("tool") == "excel_recover"
    assert "discard" in excinfo.value.context
    assert fake.recover_calls == [], "delegate must not be called on reject"


async def test_excel_recover_false_confirm_rejects_before_action() -> None:
    fake = _FakeSessionService()
    workbook_tools.set_session_service(fake)  # type: ignore[arg-type]
    try:
        with pytest.raises(ExcelVbaMcpError) as excinfo:
            await excel_recover("agent_directed", confirm=False)
    finally:
        workbook_tools.set_session_service(None)

    assert excinfo.value.code is ErrorCode.CONFIRMATION_REQUIRED
    assert fake.recover_calls == [], "delegate must not be called on reject"


async def test_excel_recover_true_confirm_delegates() -> None:
    fake = _FakeSessionService()
    workbook_tools.set_session_service(fake)  # type: ignore[arg-type]
    try:
        result = await excel_recover("agent_directed", confirm=True)
    finally:
        workbook_tools.set_session_service(None)

    assert result.ok is True
    assert result.confirmed is True
    assert fake.recover_calls == ["agent_directed"]