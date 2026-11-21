"""Unit tests for the Sheets inventory + worksheet operations (task_024).

These tests exercise the production
:class:`~excel_vba_mcp.application.sheets_service.SheetsService` and
:class:`~excel_vba_mcp.adapters.excel.sheets_adapter.SheetsAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* ``Sheets`` (all types) vs ``Worksheets`` (Worksheet only) distinction.
* ``CodeName`` is returned read-only and never changed by rename.
* ``list_sheets`` / ``get_sheet`` metadata (object_type, visibility,
  protected, table_count, used_range).
* ``add`` / ``rename`` / ``copy`` / ``set_visibility`` happy paths.
* ``delete`` guards: no snapshot -> ``SNAPSHOT_FAILED``; no approval ->
  denied; deleting the last visible sheet -> rejected.
* Postconditions: the returned DTO reflects the actual collection state.
* Non-mutating reads submit ``READ_ONLY``; mutations submit ``MUTATING``.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.sheets_adapter import SheetsAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ApprovalDeniedError, ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    SheetObjectType,
    SheetVisibility,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph
# ---------------------------------------------------------------------------
_XL_WORKSHEET = -4167
_XL_CHART = -4109
# Excel visibility codes: xlSheetVisible = -1, xlSheetHidden = 0, xlSheetVeryHidden = 2.
_XL_VISIBLE = -1
_XL_HIDDEN = 0


class FakeListObjects:
    def __init__(self, count: int) -> None:
        self.Count = count


class FakeUsedRange:
    def __init__(self, address: str) -> None:
        self.Address = address


class FakeSheet:
    """A fake ``Sheets`` item (Worksheet or Chart)."""

    def __init__(
        self,
        name: str,
        *,
        code_name: str | None = None,
        sheet_type: int = _XL_WORKSHEET,
        visible: int = _XL_VISIBLE,
        protected: bool = False,
        table_count: int = 0,
        used_range: str | None = "A1:C10",
    ) -> None:
        self.Name = name
        self.CodeName = code_name if code_name is not None else name
        self.Type = sheet_type
        self.Visible = visible
        self.ProtectContents = protected
        self.ListObjects = FakeListObjects(table_count)
        self.UsedRange = FakeUsedRange(used_range or "")
        self.Index = 0  # set by the workbook when materialised

    def Copy(self, Before: Any = None, After: Any = None) -> None:  # noqa: N802, N803 - COM
        wb = _current_workbook()
        source_name = self.Name
        copy_name = _copy_name(wb, source_name)
        clone = FakeSheet(
            copy_name,
            code_name=f"{copy_name}_Code",
            sheet_type=self.Type,
            visible=self.Visible,
            protected=self.ProtectContents,
            table_count=self.ListObjects.Count,
            used_range=self.UsedRange.Address,
        )
        if After is not None:
            wb._insert_after(After, clone)
        elif Before is not None:
            wb._insert_before(Before, clone)
        else:
            # Excel Copy() with no args creates a NEW workbook; emulate by not
            # inserting into this workbook.
            return
        wb._reindex()

    def Delete(self) -> None:  # noqa: N802 - COM
        wb = _current_workbook()
        wb._sheets = [s for s in wb._sheets if s is not self]
        wb._reindex()


# The adapter is stateless; a module-level "current workbook" lets the fake
# sheet Copy/Delete reach the owning workbook without threading it through COM.
_CURRENT: list[Any] = []


def _current_workbook() -> Any:
    return _CURRENT[-1]


def _copy_name(wb: Any, source_name: str) -> str:
    existing = {s.Name for s in wb._sheets}
    if f"{source_name} (2)" not in existing:
        return f"{source_name} (2)"
    index = 2
    while f"{source_name} ({index})" in existing:
        index += 1
    return f"{source_name} ({index})"


class FakeSheets:
    """A fake ``Sheets`` / ``Worksheets`` collection (all sheet types)."""

    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._items = sheets

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, key: Any) -> FakeSheet:  # noqa: N802 - COM casing
        if isinstance(key, int):
            return self._items[key - 1]
        for sheet in self._items:
            if sheet.Name == key:
                return sheet
        raise KeyError(key)

    def __iter__(self):
        return iter(self._items)

    def Add(self, Type: Any = None) -> FakeSheet:  # noqa: N802, N803 - COM casing
        wb = _current_workbook()
        default = f"Sheet{len(self._items) + 1}"
        sheet = FakeSheet(default, sheet_type=Type if Type is not None else _XL_WORKSHEET)
        wb._sheets.append(sheet)
        wb._reindex()
        return sheet


class FakeWorkbook:
    """A fake ``Workbook`` exposing ``Sheets`` (all) + ``Worksheets`` (subset)."""

    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets
        self._reindex()
        _CURRENT.append(self)

    def _reindex(self) -> None:
        for index, sheet in enumerate(self._sheets, start=1):
            sheet.Index = index

    def _insert_after(self, anchor: FakeSheet, clone: FakeSheet) -> None:
        idx = self._sheets.index(anchor)
        self._sheets.insert(idx + 1, clone)

    def _insert_before(self, anchor: FakeSheet, clone: FakeSheet) -> None:
        idx = self._sheets.index(anchor)
        self._sheets.insert(idx, clone)

    @property
    def Sheets(self) -> FakeSheets:  # noqa: N802 - COM casing
        return FakeSheets(self._sheets)

    @property
    def Worksheets(self) -> FakeSheets:  # noqa: N802 - COM casing
        return FakeSheets([s for s in self._sheets if s.Type == _XL_WORKSHEET])

    @property
    def FullName(self) -> str:  # noqa: N802 - COM casing
        return self._full_name

    @FullName.setter
    def FullName(self, value: str) -> None:  # noqa: N802 - COM casing
        self._full_name = value


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **_kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook(
            [
                FakeSheet("Sheet1", code_name="Sheet1_Code"),
                FakeSheet("Data", code_name="Data_Code"),
                FakeSheet("Chart1", code_name=None, sheet_type=_XL_CHART),
            ]
        )
        wb.FullName = path
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)


class FakeApp:
    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks()


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self) -> None:
        self.submits: list[dict[str, Any]] = []

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append(
            {
                "name": getattr(fn, "__name__", "?"),
                "section": section,
                "op_class": op_class,
            }
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# ---------------------------------------------------------------------------
# Rig
# ---------------------------------------------------------------------------
class Rig:
    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy()
        self.app = FakeApp()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = FakeWorker()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,  # type: ignore[arg-type]
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.snapshot = SnapshotService(
            SnapshotStore(tmp_path / "snapshots"),
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
        )
        self.service = SheetsService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

    async def open(self) -> Any:
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
        info = await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=False
        )
        _CURRENT.append(self.worker and self._open_workbook(info.workbook_id))
        return info

    def _open_workbook(self, workbook_id: str) -> Any:
        record = self.session._records[workbook_id]
        for wb in self.app.Workbooks:
            if str(wb.FullName).lower() == record.working_path.lower():
                return wb
        raise AssertionError("working copy not open")

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

    def add_snapshot(self, workbook_id: str) -> None:
        # Inject a successful snapshot without touching the filesystem/COM.
        manifest = SnapshotManifest(
            snapshot_id="snap-test",
            created_at=datetime.now(UTC),
            source_path=self.canonical,
            file_name="snap.xlsm",
            stored_path=str(self.tmp_path / "snap.xlsm"),
            sha256="0" * 64,
            size=0,
        )
        self.snapshot._latest_snapshot[workbook_id] = manifest
        self.snapshot._known_sources[workbook_id] = self.canonical

    def workbook(self) -> Any:
        return _current_workbook()


# ---------------------------------------------------------------------------
# Adapter-level: Sheets vs Worksheets, CodeName, metadata
# ---------------------------------------------------------------------------
def test_list_sheets_distinguishes_worksheets_and_chart() -> None:
    wb = FakeWorkbook(
        [
            FakeSheet("Sheet1", code_name="Sheet1_Code"),
            FakeSheet("Chart1", code_name=None, sheet_type=_XL_CHART),
        ]
    )
    adapter = SheetsAdapter()

    sheets = adapter.list_sheets(wb)

    by_name = {s.name: s for s in sheets}
    assert by_name["Sheet1"].object_type is SheetObjectType.WORKSHEET
    assert by_name["Sheet1"].code_name == "Sheet1_Code"
    assert by_name["Sheet1"].visibility is SheetVisibility.VISIBLE
    assert by_name["Chart1"].object_type is SheetObjectType.CHART
    assert by_name["Chart1"].code_name is None


def test_list_sheets_reports_metadata_fields() -> None:
    wb = FakeWorkbook(
        [
            FakeSheet(
                "Data",
                code_name="Data_Code",
                visible=_XL_HIDDEN,
                protected=True,
                table_count=2,
                used_range="A1:D42",
            )
        ]
    )
    adapter = SheetsAdapter()

    info = adapter.list_sheets(wb)[0]

    assert info.name == "Data"
    assert info.index == 1
    assert info.visibility is SheetVisibility.HIDDEN
    assert info.protected is True
    assert info.table_count == 2
    assert info.used_range_address == "A1:D42"


def test_get_sheet_missing_raises() -> None:
    wb = FakeWorkbook([FakeSheet("Sheet1")])
    adapter = SheetsAdapter()

    with pytest.raises(ExcelVbaMcpError):
        adapter.get_sheet(wb, "Nope")


def test_rename_keeps_code_name() -> None:
    wb = FakeWorkbook([FakeSheet("Old", code_name="Old_Code")])
    adapter = SheetsAdapter()

    info = adapter.rename_sheet(wb, "Old", "New")

    assert info.name == "New"
    assert info.code_name == "Old_Code"


def test_add_sheet_and_copy_sheet() -> None:
    wb = FakeWorkbook([FakeSheet("Sheet1")])
    adapter = SheetsAdapter()

    added = adapter.add_sheet(wb, "Summary")
    assert added.name == "Summary"
    assert added.object_type is SheetObjectType.WORKSHEET

    copied = adapter.copy_sheet(wb, "Sheet1", to_end=True)
    assert copied.name.startswith("Sheet1")
    assert copied.name != "Sheet1"


def test_set_visibility_and_last_visible_guard() -> None:
    # A single visible sheet cannot be hidden.
    wb = FakeWorkbook([FakeSheet("Only")])
    adapter = SheetsAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.set_visibility(wb, "Only", False)

    # With two visible sheets, hiding one is allowed.
    wb2 = FakeWorkbook([FakeSheet("A"), FakeSheet("B")])
    hidden = adapter.set_visibility(wb2, "A", False)
    assert hidden.visibility is SheetVisibility.HIDDEN


# ---------------------------------------------------------------------------
# Service-level: read paths
# ---------------------------------------------------------------------------
async def test_service_list_and_get_sheets(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()

    sheets = await rig.service.list_sheets(info.workbook_id)
    names = {s.name for s in sheets}
    assert "Sheet1" in names and "Chart1" in names

    one = await rig.service.get_sheet(info.workbook_id, "Sheet1")
    assert one.object_type is SheetObjectType.WORKSHEET

    read_submit = next(s for s in rig.worker.submits if s["name"] == "sheets_list")
    assert read_submit["op_class"] == OperationClass.READ_ONLY
    assert read_submit["section"] == "diagnostic"


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.list_sheets("missing")


# ---------------------------------------------------------------------------
# Service-level: mutation guards
# ---------------------------------------------------------------------------
async def test_add_sheet_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("sheet_add")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.add_sheet(info.workbook_id, "New")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "sheet_add"] == []


async def test_add_sheet_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.add_sheet(info.workbook_id, "New")
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "sheet_add"] == []


async def test_add_sheet_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_add")

    added = await rig.service.add_sheet(info.workbook_id, "Summary")

    assert added.name == "Summary"
    assert added.object_type is SheetObjectType.WORKSHEET
    mutate = next(s for s in rig.worker.submits if s["name"] == "sheet_add")
    assert mutate["op_class"] == OperationClass.MUTATING
    assert {s.name for s in rig.service._adapter.list_sheets(rig.workbook())} >= {
        "Sheet1",
        "Summary",
    }


async def test_rename_sheet_success_keeps_code_name(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_rename")

    renamed = await rig.service.rename_sheet(info.workbook_id, "Sheet1", "Renamed")

    assert renamed.name == "Renamed"
    assert renamed.code_name == "Sheet1_Code"


async def test_copy_sheet_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_copy")

    copied = await rig.service.copy_sheet(info.workbook_id, "Sheet1")

    assert copied.name.startswith("Sheet1")
    assert copied.name != "Sheet1"


async def test_set_visibility_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_set_visibility")

    hidden = await rig.service.set_visibility(info.workbook_id, "Data", False)

    assert hidden.visibility is SheetVisibility.HIDDEN


# ---------------------------------------------------------------------------
# Service-level: delete guards
# ---------------------------------------------------------------------------
async def test_delete_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("sheet_delete")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.delete_sheet(info.workbook_id, "Sheet1")
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "sheet_delete"] == []


async def test_delete_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.delete_sheet(info.workbook_id, "Sheet1")
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "sheet_delete"] == []


async def test_delete_last_visible_sheet_rejected(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    # Remove all but one visible sheet, then try to delete the last visible one.
    wb = rig.workbook()
    for sheet in list(wb._sheets):
        if sheet.Name not in ("Sheet1", "Chart1"):
            wb._sheets.remove(sheet)
    wb._reindex()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_delete")

    # Sheet1 is the only visible worksheet; Chart1 is not visible.
    chart = next(s for s in wb._sheets if s.Name == "Chart1")
    chart.Visible = _XL_HIDDEN

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.delete_sheet(info.workbook_id, "Sheet1")
    assert "visible" in str(exc_info.value).lower()
    # The visible-sheet guard fires on the STA thread; the delete itself never
    # runs, so the sheet is still present.
    assert "Sheet1" in {s.Name for s in rig.workbook()._sheets}


async def test_delete_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("sheet_delete")

    deleted = await rig.service.delete_sheet(info.workbook_id, "Sheet1")

    assert deleted.name == "Sheet1"
    remaining = {s.name for s in rig.service._adapter.list_sheets(rig.workbook())}
    assert "Sheet1" not in remaining
    mutate = next(s for s in rig.worker.submits if s["name"] == "sheet_delete")
    assert mutate["op_class"] == OperationClass.MUTATING
