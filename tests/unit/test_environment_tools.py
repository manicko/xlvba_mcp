"""Unit tests for the read-only ``xlvba_environment`` tool (U2; plan_03 §6).

Static, read-only: never touches Excel/COM and never starts the worker. Proves:

* ``EnvironmentReport`` is strict (``extra="forbid"``) and ``excel.build`` is
  always ``None`` (never fabricated);
* the tool reports ``excel.available=false`` when the worker is down and the
  correct ``access_mode`` / capabilities when it is up;
* ``workbooks[].has_snapshot`` and ``SessionStatus.policy.has_snapshot`` read the
  **in-memory** snapshot gate (``SnapshotService.has_snapshot`` ->
  ``_latest_snapshot``), not the on-disk store.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.api import environment_tools
from excel_vba_mcp.api.environment_tools import xlvba_environment
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.config import ServerConfig
from excel_vba_mcp.domain.models import (
    EnvironmentReport,
    SessionPolicy,
    SessionStatus,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _clear_environment() -> Any:
    """Ensure the published runtime is cleared before and after each test."""

    environment_tools.clear_environment_runtime()
    yield
    environment_tools.clear_environment_runtime()


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _worker_status(
    *,
    state: WorkbookSessionState = WorkbookSessionState.RUNNING,
    accepting: bool = True,
    owned_pid: int | None = 4242,
) -> WorkerStatus:
    return WorkerStatus(state=state, accepting=accepting, owned_pid=owned_pid)


def _workbook_info(workbook_id: str = "w1", *, excel_version: str = "16.0") -> WorkbookInfo:
    return WorkbookInfo(
        workbook_id=workbook_id,
        name="book.xlsm",
        format="xlsm",
        path="C:/work/book.xlsm",
        source_path="C:/orig/book.xlsm",
        read_only=False,
        has_vba=True,
        excel_version=excel_version,
        session_id="sess",
        mode=SessionMode.EXECUTE,
    )


class _FakeSession:
    """Duck-typed session exposing only what ``xlvba_environment`` reads."""

    def __init__(
        self,
        worker_status: WorkerStatus,
        infos: list[WorkbookInfo],
        *,
        has_snapshot: bool = False,
    ) -> None:
        self._worker_status = worker_status
        self._infos = infos
        self.session_id = "sess"
        self._has_snapshot = has_snapshot

    def worker_status(self) -> WorkerStatus:
        return self._worker_status

    def list_workbooks_sync(self) -> list[WorkbookInfo]:
        return list(self._infos)

    async def get_status(self, workbook_id: str) -> SessionStatus:
        now = datetime.now(UTC)
        return SessionStatus(
            session_id=self.session_id,
            workbook_id=workbook_id,
            state=WorkbookSessionState.RUNNING,
            mode=SessionMode.EXECUTE,
            read_only=False,
            dirty=True,
            access_mode=AccessMode.FULL,
            policy=SessionPolicy(
                approval_mode=AccessMode.FULL,
                mutating_requires_grant=False,
                executing_requires_grant=False,
                recovery_auto_allowed=True,
                snapshot_required_for_mutations=True,
                has_snapshot=self._has_snapshot,
            ),
            created_at=now,
            updated_at=now,
        )


# ---------------------------------------------------------------------------
# EnvironmentReport DTO
# ---------------------------------------------------------------------------
def test_environment_report_rejects_extra_fields() -> None:
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        EnvironmentReport.model_validate(
            {
                "server": {
                    "name": "excel-vba-mcp",
                    "version": "0.1.0",
                    "access_mode": "full",
                    "macro_timeout_s": 90.0,
                    "open_timeout_s": 300.0,
                    "liveness_ceiling_s": 600.0,
                    "log_level": "INFO",
                },
                "excel": {"available": False},
                "capabilities": {
                    "vba": True,
                    "macros": True,
                    "pivot_tables": True,
                    "background_jobs": True,
                    "recovery": True,
                },
                "policy": {
                    "approval": {
                        "mode": "full",
                        "mutating_requires_grant": False,
                        "executing_requires_grant": False,
                        "recovery_auto_allowed": True,
                    }
                },
                "unexpected_key": True,
            }
        )


# ---------------------------------------------------------------------------
# Worker down -> available False, build None
# ---------------------------------------------------------------------------
async def test_environment_worker_down_is_truthful() -> None:
    environment_tools.set_environment_runtime(
        config=ServerConfig(access_mode=AccessMode.FULL),
        session=None,
        snapshot=None,
        path_policy=None,
    )

    report = await xlvba_environment()

    assert report.excel.available is False
    assert report.excel.build is None
    assert report.excel.owned_pid is None
    assert report.workbooks == []
    assert report.glossary_uri == "xlvba://glossary"
    assert report.policy.macro.classification_is_authorization is False
    assert report.policy.macro.r13_preflight == "veto"


# ---------------------------------------------------------------------------
# Worker up -> available True, access_mode, capabilities, has_snapshot via gate
# ---------------------------------------------------------------------------
async def test_environment_worker_up_reports_access_mode_and_capabilities() -> None:
    environment_tools.set_environment_runtime(
        config=ServerConfig(access_mode=AccessMode.RESTRICTED),
        session=_FakeSession(_worker_status(), [_workbook_info()], has_snapshot=True),
        snapshot=None,
        path_policy=None,
    )

    report = await xlvba_environment()

    assert report.excel.available is True
    assert report.excel.owned_pid == 4242
    assert report.excel.build is None
    assert report.excel.application_version == "16.0"
    assert report.server.access_mode is AccessMode.RESTRICTED
    assert report.policy.approval.mutating_requires_grant is True
    assert report.policy.approval.executing_requires_grant is True
    assert report.policy.approval.recovery_auto_allowed is True
    assert report.capabilities.non_activating_compile is False
    assert report.capabilities.pivot_filter_kinds == ["manual"]
    assert report.capabilities.pivot_source_types == ["database"]
    # has_snapshot is propagated from the in-memory gate.
    assert report.workbooks[0].has_snapshot is True
    assert report.workbooks[0].dirty is True


async def test_environment_full_mode_does_not_require_grant() -> None:
    environment_tools.set_environment_runtime(
        config=ServerConfig(access_mode=AccessMode.FULL),
        session=_FakeSession(_worker_status(), []),
        snapshot=None,
        path_policy=None,
    )

    report = await xlvba_environment()

    assert report.policy.approval.mutating_requires_grant is False
    assert report.policy.approval.executing_requires_grant is False


# ---------------------------------------------------------------------------
# has_snapshot reads the in-memory gate, not the store
# ---------------------------------------------------------------------------
def test_has_snapshot_reads_in_memory_gate_not_store() -> None:
    svc = object.__new__(SnapshotService)
    svc._latest_snapshot = {"w-inmem": object()}  # type: ignore[attr-defined]

    class _FakeStore:
        def list(self) -> list[str]:
            return ["w-store-only"]

    svc._store = _FakeStore()  # type: ignore[attr-defined]

    assert svc.has_snapshot("w-inmem") is True
    # A store-only entry (no in-memory snapshot) must be False — the gate is read.
    assert svc.has_snapshot("w-store-only") is False


# ---------------------------------------------------------------------------
# SessionStatus.policy is populated from the in-memory gate
# ---------------------------------------------------------------------------
def _make_session_service(tmp_path: Path, mode: AccessMode) -> SessionService:
    worker = type("W", (), {"get_status": lambda self: _worker_status()})()
    return SessionService(
        worker,  # type: ignore[arg-type]
        object(),  # app_adapter unused by get_status
        PathPolicy({str(tmp_path): str(tmp_path)}),
        ApprovalPolicy(mode=mode),
        AuditLogger(tmp_path / "audit.jsonl"),
        work_dir=tmp_path / "work",
    )


async def test_session_status_policy_uses_in_memory_gate(tmp_path: Path) -> None:
    service = _make_session_service(tmp_path, AccessMode.FULL)
    service.set_snapshot_gate(lambda workbook_id: workbook_id == "w-inmem")

    in_mem = await service.get_status("w-inmem")
    assert in_mem.policy is not None
    assert in_mem.policy.has_snapshot is True
    assert in_mem.policy.classification_is_authorization is False
    assert in_mem.policy.r13_preflight == "veto"
    assert in_mem.policy.mutating_requires_grant is False

    other = await service.get_status("w-other")
    assert other.policy is not None
    assert other.policy.has_snapshot is False


async def test_session_status_policy_restricted_requires_grant(tmp_path: Path) -> None:
    service = _make_session_service(tmp_path, AccessMode.RESTRICTED)

    status = await service.get_status("w1")

    assert status.policy is not None
    assert status.policy.approval_mode is AccessMode.RESTRICTED
    assert status.policy.mutating_requires_grant is True
    assert status.policy.executing_requires_grant is True
    assert status.policy.recovery_auto_allowed is True
    assert status.policy.snapshot_required_for_mutations is True