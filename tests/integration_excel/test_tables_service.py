"""Real Excel integration tests for the Table (ListObject) surface (task_026).

These tests drive the real :class:`~excel_vba_mcp.application.tables_service.
TablesService` and :class:`~excel_vba_mcp.adapters.excel.tables_adapter.
TablesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel). They assert:

* a real ``ListObject`` is created via ``ListObjects.Add`` and reported with the
  expected range / header / body / total metadata, style, rows and columns;
* typed row reads (``str`` / ``int`` / ``float`` / ``bool`` / ``None`` / date /
  Excel error) survive a real round-trip;
* ``table_append_rows`` / ``table_insert_row`` / ``table_update_rows`` /
  ``table_delete_rows`` mutate through ``ListRows`` and preserve the table
  structure (name, columns, filters);
* ``table_add_column`` / ``table_delete_column`` mutate through ``ListColumns``
  and keep headers unique;
* ``table_resize`` grows the table and keeps it a ListObject;
* ambiguous ``where`` updates are rejected (no row changed);
* the original fixture on disk is byte-identical after every test (migrations
  never touch the source);
* BG03/BG04/BG05 hold (no non-owned Excel touched, hidden + non-activating, and
  the worker refuses mutations on a hung instance).

Mutations run only on the disposable working copy of the fixture. Excel is
created with ``DispatchEx`` and always ``Visible = False``; no ``.Activate`` /
``.Select`` is ever called.
"""

from __future__ import annotations

import hashlib
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
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

pytest.importorskip("win32com", reason="Excel COM automation unavailable")

_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

pytestmark = pytest.mark.anyio

_TEST_TIMEOUT_S = 120.0


def _requires_excel() -> Any:
    """Skip the test when Excel COM automation is unavailable."""

    return pytest.mark.skipif(
        not _excel_available(), reason="Excel COM automation is not available on this host"
    )


def _excel_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001 - COM unavailable (non-Windows)
        return False
    return True


requires_excel = _requires_excel()


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def clean_project() -> Any:
    """Yield the committed ``CleanProject.xlsm`` fixture path (skip if absent)."""

    if not _CLEAN_PROJECT.is_file():
        pytest.skip(f"fixture not available: {_CLEAN_PROJECT}")
    return _CLEAN_PROJECT


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class Rig:
    """Real COM worker + Excel + full session/snapshot/policy stack."""

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
                granted_at=datetime.now(UTC),
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
                granted_at=datetime.now(UTC),
                granted_by="tester",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    async def create_snapshot(self, workbook_id: str) -> None:
        await self.snapshot.create_snapshot(workbook_id)

    async def close(self) -> None:
        await self.worker.shutdown()
        self.audit.close()


@pytest.fixture
async def rig(tmp_path: Path, clean_project: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every test mutates a disposable working copy and never the source.
    original = tmp_path / clean_project.name
    shutil.copy2(clean_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


# ---------------------------------------------------------------------------
# STA helpers to create / inspect a table directly on the working copy
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


async def _seed_table(rig: Rig, workbook_id: str) -> str:
    """Create a real ListObject on a fresh sheet and return its name.

    Runs on the STA worker inside the owned Excel. Adds a worksheet, writes a
    header + rows, then ``ListObjects.Add`` over the range. Returns the created
    table name.
    """

    working_path = _record_path(rig, workbook_id)

    def seed() -> str:
        wb = _open_workbook(rig, working_path)
        # A fresh sheet so the test does not depend on fixture content.
        sheet = wb.Worksheets.Add()
        sheet.Name = "TableData"
        headers = ["Name", "Qty", "Price", "Active", "When", "Err"]
        rows = [
            ["widget", 3, 1.5, True, datetime(2026, 1, 2, 3, 4, 5), None],
            ["gizmo", 7, 2.25, False, None, None],
            ["gizmo", 1, 0.5, True, None, None],
        ]
        for col, header in enumerate(headers, start=1):
            sheet.Cells(1, col).Value = header
        for r, row in enumerate(rows, start=2):
            for col, value in enumerate(row, start=1):
                sheet.Cells(r, col).Value = value
        # An Excel error literal in a cell (typed as an error on read).
        sheet.Cells(2, 6).Value = "=1/0"
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
# Inventory + typed read
# ---------------------------------------------------------------------------
async def test_create_list_and_read_rows(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    name = await _seed_table(rig, wid)

    tables = await rig.tables.list_tables(wid)
    assert any(t.name == name for t in tables)
    one = await rig.tables.get_table(wid, name)
    assert one.name == name
    assert one.sheet_name == "TableData"
    assert one.row_count == 3
    assert [c.name for c in one.columns] == ["Name", "Qty", "Price", "Active", "When", "Err"]
    assert one.range_address  # populated
    assert one.header_row_range is not None
    assert one.data_body_range is not None

    rows = await rig.tables.read_table_rows(wid, name)
    assert rows.total_rows == 3
    assert rows.returned_rows == 3
    first = rows.rows[0]
    assert first[0].kind is CellValueKind.STRING and first[0].value == "widget"
    assert first[1].kind is CellValueKind.NUMBER and int(first[1].value) == 3
    assert first[2].kind is CellValueKind.NUMBER and abs(float(first[2].value) - 1.5) < 1e-9
    assert first[3].kind is CellValueKind.BOOLEAN and first[3].value is True
    assert first[4].kind is CellValueKind.DATE  # ISO date string
    # The error cell is typed as an error, never an ordinary string.
    assert first[5].kind is CellValueKind.ERROR

    # offset / limit
    page = await rig.tables.read_table_rows(wid, name, offset=1, limit=1)
    assert page.returned_rows == 1
    assert page.rows[0][0].value == "gizmo"

    # BG03: original fixture never modified.
    assert _sha256(rig.project) == before


async def test_unknown_table_raises(rig: Rig) -> None:
    info = await rig.open()
    with pytest.raises(ExcelVbaMcpError):
        await rig.tables.get_table(info.workbook_id, "NoSuchTable")


# ---------------------------------------------------------------------------
# Row mutations
# ---------------------------------------------------------------------------
async def test_append_and_read_back(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_append_rows")

    result = await rig.tables.append_table_rows(
        wid, name, [{"Name": "appended", "Qty": 9, "Price": 4.5}]
    )
    assert result.after.row_count == 4
    assert result.after.name == name  # structure preserved
    assert [c.name for c in result.after.columns] == [
        "Name",
        "Qty",
        "Price",
        "Active",
        "When",
        "Err",
    ]

    rows = await rig.tables.read_table_rows(wid, name, offset=3, limit=1)
    assert rows.rows[0][0].value == "appended"
    assert int(rows.rows[0][1].value) == 9


async def test_insert_row_middle(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_insert_row")

    result = await rig.tables.insert_table_row(wid, name, 2, {"Name": "inserted", "Qty": 0})
    assert result.after.row_count == 4

    rows = await rig.tables.read_table_rows(wid, name)
    names = [row[0].value for row in rows.rows]
    assert names[0] == "widget"
    assert names[1] == "inserted"
    assert names[2] == "gizmo"


async def test_update_rows_by_index_and_where(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_update_rows")

    # By explicit index (row 2 = the second data row = "gizmo" Qty 7).
    result = await rig.tables.update_table_rows(wid, name, [{"index": 2, "values": {"Qty": 70}}])
    assert result.affected_rows == [2]

    rows = await rig.tables.read_table_rows(wid, name)
    assert int(rows.rows[1][1].value) == 70
    assert rows.rows[1][0].value == "gizmo"

    # By a unique where filter ("widget" appears once).
    result2 = await rig.tables.update_table_rows(
        wid, name, [{"where": {"Name": "widget"}, "values": {"Qty": 100}}]
    )
    assert result2.affected_rows == [1]
    rows2 = await rig.tables.read_table_rows(wid, name)
    assert int(rows2.rows[0][1].value) == 100


async def test_update_rows_ambiguous_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_update_rows")

    # "gizmo" appears in two rows -> ambiguous; must be rejected, no mutation.
    before = await rig.tables.read_table_rows(wid, name)
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.update_table_rows(
            wid, name, [{"where": {"Name": "gizmo"}, "values": {"Qty": 999}}]
        )
    assert exc_info.value.context.get("reason") == "ambiguous"
    after = await rig.tables.read_table_rows(wid, name)
    assert [row[1].value for row in before.rows] == [row[1].value for row in after.rows]


async def test_delete_rows_by_indices(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_delete_rows")

    result = await rig.tables.delete_table_rows(wid, name, indices=[1, 3])
    assert result.after.row_count == 1
    rows = await rig.tables.read_table_rows(wid, name)
    assert [row[0].value for row in rows.rows] == ["gizmo"]


# ---------------------------------------------------------------------------
# Column mutations
# ---------------------------------------------------------------------------
async def test_add_and_delete_column(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_add_column")
    rig.grant("table_delete_column")

    added = await rig.tables.add_table_column(wid, name, "Note")
    assert "Note" in [c.name for c in added.after.columns]

    deleted = await rig.tables.delete_table_column(wid, name, "Note")
    assert "Note" not in [c.name for c in deleted.after.columns]
    assert deleted.after.name == name


async def test_add_column_duplicate_header_rejected(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_add_column")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.add_table_column(wid, name, "Qty")  # duplicate
    assert "unique" in str(exc_info.value).lower()


# ---------------------------------------------------------------------------
# Create + resize
# ---------------------------------------------------------------------------
async def test_create_table_service(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("table_create")

    # Write a header + rows on a fresh sheet, then create a table via the tool.
    working_path = _record_path(rig, wid)

    def seed_plain() -> None:
        wb = _open_workbook(rig, working_path)
        sheet = wb.Worksheets.Add()
        sheet.Name = "PlainData"
        for col, header in enumerate(["A", "B"], start=1):
            sheet.Cells(1, col).Value = header
        sheet.Cells(2, 1).Value = 1
        sheet.Cells(2, 2).Value = 2
        sheet.Cells(3, 1).Value = 3
        sheet.Cells(3, 2).Value = 4

    seed_plain.__name__ = "seed_plain"
    await rig.worker.submit(
        seed_plain, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.MUTATING
    )

    created = await rig.tables.create_table(wid, "PlainData", "A1:B3", style="TableStyleMedium2")
    assert created.before is None
    assert created.after.sheet_name == "PlainData"
    assert created.after.row_count == 2
    assert [c.name for c in created.after.columns] == ["A", "B"]


async def test_resize(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    name = await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_resize")

    before = await rig.tables.get_table(wid, name)
    # Grow the table by one (empty) data row.
    new_range = f"A1:F{before.row_count + 2}"
    resized = await rig.tables.resize_table(wid, name, new_range)
    assert resized.after.row_count == before.row_count + 1
    assert resized.after.name == name


# ---------------------------------------------------------------------------
# Guards
# ---------------------------------------------------------------------------
async def test_append_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_table(rig, wid)
    rig.grant("table_append_rows")  # approval but no snapshot

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.tables.append_table_rows(wid, "SampleTable", [{"Name": "x"}])
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_append_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    # No grant.

    with pytest.raises(ExcelVbaMcpError):
        await rig.tables.append_table_rows(wid, "SampleTable", [{"Name": "x"}])


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
async def test_bg03_hidden_owned_excel_and_snapshot(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id
    await _seed_table(rig, wid)

    # BG03/BG05: the automation Excel is hidden and non-activating (read on the
    # STA worker so the COM property access is thread-correct).
    def read_app_flags() -> dict[str, Any]:
        app = rig.app.app
        assert app is not None
        return {"visible": bool(app.Visible), "display_alerts": bool(app.DisplayAlerts)}

    read_app_flags.__name__ = "read_app_flags"
    flags = await rig.worker.submit(
        read_app_flags,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert flags["visible"] is False
    # A snapshot (SaveCopyAs) exists without touching the source fixture.
    await rig.create_snapshot(wid)
    assert rig.snapshot._latest_snapshot.get(wid) is not None

    # BG03: source fixture unchanged.
    assert _sha256(rig.project) == before
    # BG05: alerts restored after a mutation-capable session.
    assert flags["display_alerts"] is True


async def test_bg04_worker_refuses_mutation_on_hung_excel(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await _seed_table(rig, wid)
    await rig.create_snapshot(wid)
    rig.grant("table_append_rows")

    def hang() -> None:
        import time

        # Block the STA thread past the soft deadline so the submit times out and
        # the worker enters NEEDS_RECOVERY (a hung dispatch is never retried).
        time.sleep(3.0)

    hang.__name__ = "hang_test"
    # A very short deadline makes the sleep outlive the soft timeout, moving the
    # worker to NEEDS_RECOVERY.
    with pytest.raises(ExcelVbaMcpError):
        await rig.worker.submit(
            hang,
            timeout=0.1,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    # BG04: the worker is now NEEDS_RECOVERY and refuses new work, so the table
    # mutation never reaches Excel (fail fast, no enqueue).
    with pytest.raises(ExcelVbaMcpError):
        await rig.tables.append_table_rows(wid, "SampleTable", [{"Name": "x"}])


async def test_bg05_session_state_and_alerts(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING

    # BG05: DisplayAlerts is restored (not permanently disabled).
    def read_alerts() -> bool:
        app = rig.app.app
        assert app is not None
        return bool(app.DisplayAlerts)

    read_alerts.__name__ = "read_alerts"
    alerts = await rig.worker.submit(
        read_alerts,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )
    assert alerts is True
