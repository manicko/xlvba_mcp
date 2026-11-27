"""Live verification of the WP1 read-path source-type gate (integration_excel).

``pivot_read(fields-list)`` and ``pivot_read(field-get-filters)`` enumerate
``PivotFields``; on an external / OLAP / Data Model cache that enumeration forces
Excel to resolve external data, blocking the single-STA worker past the client
deadline (the WP1 hang -> ``-32001`` -> recover). The fix resolves the cache
source type **before** any ``PivotFields`` access and rejects an unsupported
cache with ``UNSUPPORTED_PIVOT_SOURCE`` (fail closed on ``UNKNOWN``).

This module drives the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
against a hidden, automation-created Excel instance and asserts:

* the problematic external/OLAP workbook (``.ai/tmp/theiss/weekly/report_w.xlsx``,
  skipped when absent) makes ``list_pivot_fields`` raise
  ``UNSUPPORTED_PIVOT_SOURCE`` **promptly** (elapsed < 10 s), with the worker
  still ``accepting`` afterwards (no recovery needed);
* a supported ``DATABASE`` pivot (the committed ``PivotTablesProject.xlsm``
  fixture) still returns its fields.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_pivot_read_source_gate.py -v -s
"""

from __future__ import annotations

import shutil
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotFieldInfo
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode, WorkbookSessionState
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_EXTERNAL_FIXTURE = _REPO_ROOT / ".ai" / "tmp" / "theiss" / "weekly" / "report_w.xlsx"
_DATABASE_FIXTURE = _REPO_ROOT / "tests" / "fixtures" / "PivotTablesProject.xlsm"

_TEST_TIMEOUT_S = 60.0
_PROMPT_DEADLINE_S = 10.0


def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


class Rig:
    """Real COM worker + Excel + PivotTablesService for one verification run."""

    def __init__(self, tmp_path: Path, project: Path) -> None:
        self.tmp_path = tmp_path
        self.project = project
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        canonical = self.path_policy.validate_path(str(project), OperationClass.READ_ONLY)
        self.canonical = str(canonical)
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter, orphan_timeout_s=30.0)
        self.approval = ApprovalPolicy()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )
        self.pivots = PivotTablesService(
            self.worker,
            self.session,
            self.audit,
            snapshot_service=self.snapshot,
            approval=self.approval,
        )

    async def start(self) -> None:
        await self.worker.start()

    async def open(self) -> Any:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action="excel_open"),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(UTC),
                granted_by="wp1-verifier",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )
        return await self.session.open_workbook(
            str(self.project), mode=SessionMode.INSPECT, read_only=True
        )

    async def close(self) -> None:
        owned_pid = self.worker.get_status().owned_pid
        await self.worker.shutdown()
        self.audit.close()
        if owned_pid is not None:
            time.sleep(0.1)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _copy_under_tmp(src: Path, tmp_path: Path) -> Path:
    target = tmp_path / src.name
    shutil.copy2(src, target)
    return target


async def _first_pivot_name(rig: Rig, workbook_id: str) -> str:
    """Resolve the first PivotTable name via the (gated) inventory read."""

    inventory = await rig.pivots.list_pivot_tables(workbook_id)
    names = [pt.name for pt in inventory.pivot_tables]
    assert names, "the workbook must expose at least one PivotTable"
    return names[0]


# ---------------------------------------------------------------------------
# External / OLAP workbook: fields-list is rejected promptly, no recovery
# ---------------------------------------------------------------------------
@requires_excel
async def test_external_source_fields_list_rejected_promptly(tmp_path: Path) -> None:
    if not _EXTERNAL_FIXTURE.is_file():
        pytest.skip(f"external/OLAP workbook not available: {_EXTERNAL_FIXTURE}")

    project = _copy_under_tmp(_EXTERNAL_FIXTURE, tmp_path)
    rig = Rig(tmp_path, project)
    try:
        await rig.start()
        info = await rig.open()
        pivot_name = await _first_pivot_name(rig, info.workbook_id)

        started = time.monotonic()
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.pivots.list_pivot_fields(info.workbook_id, pivot_name)
        elapsed = time.monotonic() - started

        assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
        assert exc_info.value.context["reason"] == "unsupported_source"
        # The gate must short-circuit well inside the client deadline.
        assert elapsed < _PROMPT_DEADLINE_S, (
            f"fields-list took {elapsed:.2f}s (>= {_PROMPT_DEADLINE_S}s); the gate did "
            "not short-circuit before PivotFields enumeration"
        )

        # The worker is still healthy and accepting — no recovery was needed.
        status = rig.worker.get_status()
        assert status.accepting is True
        assert status.state is WorkbookSessionState.RUNNING

        await rig.session.close_workbook(info.workbook_id)
    finally:
        await rig.close()


# ---------------------------------------------------------------------------
# Supported DATABASE workbook: fields-list still returns fields
# ---------------------------------------------------------------------------
@requires_excel
async def test_database_source_fields_list_returns_fields(tmp_path: Path) -> None:
    if not _DATABASE_FIXTURE.is_file():
        pytest.skip(f"database fixture not available: {_DATABASE_FIXTURE}")

    project = _copy_under_tmp(_DATABASE_FIXTURE, tmp_path)
    rig = Rig(tmp_path, project)
    try:
        await rig.start()
        info = await rig.open()

        fields = await rig.pivots.list_pivot_fields(info.workbook_id, "ptA")
        assert isinstance(fields, list)
        assert fields, "the DATABASE pivot must expose source fields"
        assert all(isinstance(field, PivotFieldInfo) for field in fields)
        assert {field.source_name for field in fields} >= {"Region", "Product", "Sales"}

        status = rig.worker.get_status()
        assert status.accepting is True

        await rig.session.close_workbook(info.workbook_id)
    finally:
        await rig.close()