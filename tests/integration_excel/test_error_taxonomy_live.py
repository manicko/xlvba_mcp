"""Live-Excel error-taxonomy classification over the real service/adapter path (WP1).

These tests drive the production :class:`~excel_vba_mcp.application.names_service.
NamesService` / :class:`~excel_vba_mcp.application.tables_service.TablesService`
and their STA-local adapters against an **automation-created hidden Excel
instance** (never the user's Excel), on a disposable working copy of the committed
``CleanProject.xlsm`` fixture. They assert the *serialized* error envelope
(``ExcelVbaMcpError.to_dict()``) — not merely the exception class — for the WP1
deterministic raise sites:

* a duplicate defined-name create (same name + scope) is ``NAME_CONFLICT`` /
  ``validation`` / not retryable, with a corrective ``suggestedNextActions``
  (never a blind "retry"), and the workbook is **not** partially mutated (the
  defined-name count is unchanged);
* a table append that references an unknown header is ``NOT_FOUND`` /
  ``not_found`` / not retryable, with corrective actions and **no** partial
  mutation (the table row count is unchanged);
* a missing table and a missing sheet are ``NOT_FOUND`` / ``not_found`` / not
  retryable.

In every case the error must NOT be the generic ``COM_BUSY`` fallback and must
NOT classify as ``transient_com_failure`` / retryable. The original fixture on
disk is byte-identical after each test (the session only ever mutates its own
working copy under ``work/`` and the worker is shut down without saving).

Run serially (no xdist) on a host with Excel + VBOM trust::

    .venv\\Scripts\\python.exe -m pytest tests/integration_excel/test_error_taxonomy_live.py -v -s
"""

from __future__ import annotations

import hashlib
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-untyped]
    import win32com  # type: ignore[import-untyped]
    import win32gui  # type: ignore[import-untyped]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    NameScope,
    OperationClass,
    SessionMode,
)
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

_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

_TEST_TIMEOUT_S = 120.0


def _pywin32_available() -> bool:
    return (
        sys.platform == "win32"
        and pythoncom is not None
        and win32com is not None
        and win32gui is not None
    )


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(f"fixture not available: {path}")
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# Envelope assertion (the contract under test): assert to_dict(), not the class.
# ---------------------------------------------------------------------------
def _assert_non_retryable_taxonomy(
    error: ExcelVbaMcpError, *, code: str, category: str
) -> dict[str, Any]:
    """Assert the serialized envelope is a specific, non-retryable, corrective error.

    Mirrors the WP1 unit-test helper so the live path is held to the same contract:
    a deterministic caller error must never masquerade as the generic ``COM_BUSY``
    fallback or as a transient/retryable condition, and must carry corrective
    (never blind-retry) ``suggestedNextActions``.
    """

    payload = error.to_dict()
    assert payload["errorCode"] == code
    assert payload["errorCode"] not in {"COM_BUSY", "RPC_E_CALL_REJECTED"}
    assert payload["errorCategory"] == category
    assert payload["errorCategory"] != "transient_com_failure"
    assert payload["isRetryable"] is False
    actions = payload["suggestedNextActions"]
    assert actions and all(isinstance(action, str) and action for action in actions)
    # Corrective, never a blind retry recommendation.
    assert "retry" not in " ".join(actions).lower()
    return payload


# ---------------------------------------------------------------------------
# Fixtures / rig (mirrors test_names_service.py / test_tables_service.py)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack + names & tables."""

    def __init__(self, tmp_path: Path, project: Path) -> None:
        self.tmp_path = tmp_path
        self.project = project
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        canonical = self.path_policy.validate_path(str(project), OperationClass.MUTATING)
        self.canonical = str(canonical)
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter)
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
        self.names = NamesService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self.tables = TablesService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

    @property
    def app(self) -> ExcelAppAdapter:
        return self.adapter

    async def start(self) -> None:
        await self.worker.start()

    async def open(self, *, read_only: bool = False) -> Any:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action="excel_open"),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )
        return await self.session.open_workbook(
            str(self.project), mode=SessionMode.INSPECT, read_only=read_only
        )

    def grant(self, action: str) -> None:
        self.approval.grant(
            ApprovalGrant(
                key=ApprovalKey(canonical_path=self.canonical, action=action),
                scope=ApprovalScope.SESSION,
                granted_at=datetime.now(),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def create_snapshot(self, workbook_id: str) -> None:
        await self.snapshot.create_snapshot(workbook_id)

    async def close(self) -> None:
        # Shut the worker down without saving: the disposable working copy is
        # discarded and the original fixture on disk is never touched.
        await self.worker.shutdown()
        self.audit.close()


@pytest.fixture
async def rig(tmp_path: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every test mutates a disposable working copy and never the source.
    clean_project = _fixture_or_skip(_CLEAN_PROJECT)
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


# ---------------------------------------------------------------------------
# STA helpers to seed a worksheet / table on the working copy
# ---------------------------------------------------------------------------
def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


def _record_path(rig: Rig, workbook_id: str) -> str:
    return rig.session._records[workbook_id].working_path


async def _add_worksheet(rig: Rig, workbook_id: str, name: str) -> str:
    """Create a real worksheet on the working copy and return its name."""

    working_path = _record_path(rig, workbook_id)

    def seed() -> str:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Add()
        sheet.Name = name
        return str(sheet.Name)

    seed.__name__ = "seed_worksheet"
    return await rig.worker.submit(
        seed, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.MUTATING
    )


async def _seed_table(rig: Rig, workbook_id: str) -> str:
    """Create a real ListObject on a fresh sheet and return its name (STA-local)."""

    working_path = _record_path(rig, workbook_id)

    def seed() -> str:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Add()
        sheet.Name = "TableData"
        headers = ["Name", "Qty", "Price"]
        rows = [
            ["widget", 3, 1.5],
            ["gizmo", 7, 2.25],
        ]
        for col, header in enumerate(headers, start=1):
            sheet.Cells(1, col).Value = header
        for r, row in enumerate(rows, start=2):
            for col, value in enumerate(row, start=1):
                sheet.Cells(r, col).Value = value
        last_col = chr(ord("A") + len(headers) - 1)
        source = sheet.Range(f"A1:{last_col}{len(rows) + 1}")
        list_object = sheet.ListObjects.Add(SourceType=1, Source=source, XlListObjectHasHeaders=1)
        list_object.Name = "SampleTable"
        wb.Save()
        return str(list_object.Name)

    seed.__name__ = "seed_table"
    return await rig.worker.submit(
        seed, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.MUTATING
    )


# ---------------------------------------------------------------------------
# Case 1: duplicate defined-name create -> NAME_CONFLICT (validation, not retryable)
# ---------------------------------------------------------------------------
@requires_excel
async def test_duplicate_name_create_is_name_conflict(rig: Rig) -> None:
    before = _sha256(rig.project)
    committed_hash = _sha256(_CLEAN_PROJECT)

    info = await rig.open()
    wid = info.workbook_id
    await _add_worksheet(rig, wid, "NamesData")
    await rig.create_snapshot(wid)
    rig.grant("name_create")

    created = await rig.names.create_name(wid, "Dup", "=NamesData!$A$1", scope=NameScope.WORKBOOK)
    assert created.name == "Dup"

    count_before = len(await rig.names.list_names(wid))

    # Duplicate create (same name + scope) must be a structured NAME_CONFLICT.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.names.create_name(wid, "Dup", "=NamesData!$B$1", scope=NameScope.WORKBOOK)

    payload = _assert_non_retryable_taxonomy(
        exc_info.value, code="NAME_CONFLICT", category="validation"
    )
    # Corrective action tells the caller to choose a unique name / delete the existing.
    assert payload["suggestedNextActions"] == ["choose a unique name or delete the existing one"]
    assert exc_info.value.context.get("reason") == "duplicate"

    # No partial mutation: the defined-name count is unchanged.
    count_after = len(await rig.names.list_names(wid))
    assert count_after == count_before, (
        f"duplicate create must not add a name: before={count_before} after={count_after}"
    )

    # BG03: original + committed fixture never modified (worker closed without saving).
    assert _sha256(rig.project) == before
    assert _sha256(_CLEAN_PROJECT) == committed_hash


# ---------------------------------------------------------------------------
# Case 2: table append with unknown header -> NOT_FOUND (not_found, not retryable)
# ---------------------------------------------------------------------------
@requires_excel
async def test_table_append_unknown_header_is_not_found(rig: Rig) -> None:
    before = _sha256(rig.project)

    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_append_rows")

    row_count_before = (await rig.tables.get_table(wid, name)).row_count

    # An unknown header must be rejected before any row is added.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.append_table_rows(wid, name, [{"Nope": 1}])

    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context.get("reason") == "not_found"

    # No partial mutation: the table row count is unchanged.
    row_count_after = (await rig.tables.get_table(wid, name)).row_count
    assert row_count_after == row_count_before, (
        f"unknown-header append must not add a row: before={row_count_before} "
        f"after={row_count_after}"
    )

    assert _sha256(rig.project) == before


# ---------------------------------------------------------------------------
# Case 3a: missing table -> NOT_FOUND (not_found, not retryable)
# ---------------------------------------------------------------------------
@requires_excel
async def test_missing_table_is_not_found(rig: Rig) -> None:
    before = _sha256(rig.project)

    info = await rig.open()
    wid = info.workbook_id

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.get_table(wid, "NoSuchTable")

    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context.get("reason") == "not_found"

    assert _sha256(rig.project) == before


# ---------------------------------------------------------------------------
# Case 3b: missing sheet (table create over an absent sheet) -> NOT_FOUND
# ---------------------------------------------------------------------------
@requires_excel
async def test_missing_sheet_is_not_found(rig: Rig) -> None:
    before = _sha256(rig.project)

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("table_create")

    tables_before = {t.name for t in await rig.tables.list_tables(wid)}

    # Creating a table over a sheet that does not exist must surface NOT_FOUND
    # (the adapter's _require_sheet), not the generic COM_BUSY fallback.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.create_table(wid, "NoSuchSheet", "A1:B2")

    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context.get("reason") == "not_found"

    # No partial mutation: no table was created.
    tables_after = {t.name for t in await rig.tables.list_tables(wid)}
    assert tables_after == tables_before, (
        f"missing-sheet create must not add a table: before={sorted(tables_before)} "
        f"after={sorted(tables_after)}"
    )

    assert _sha256(rig.project) == before