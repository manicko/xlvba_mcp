"""Unit tests for the defined-name (Named Range) surface (task_028).

These tests exercise the production
:class:`~excel_vba_mcp.application.names_service.NamesService` and
:class:`~excel_vba_mcp.adapters.excel.names_adapter.NamesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

The fake models workbook-scoped names in ``wb.Names`` and worksheet-scoped names
in ``ws.Names``. A ``FakeName`` exposes a settable ``RefersTo`` and a
``RefersToRange`` that either returns a ``FakeRange`` (range target) or raises
(non-range target), so the adapter's separate error-handling branch is exercised.

Covered:

* ``names_list`` field mapping (scope, NameLocal, Visible, RefersTo,
  RefersToLocal) and ``target_kind`` (range / constant / formula / external /
  error / unknown) without any formula evaluation.
* ``name_get`` workbook vs worksheet scope and missing-name errors.
* ``name_create`` / ``name_update`` / ``name_delete`` scoped CRUD.
* Conflict detection: duplicate in the requested scope and cross-scope collision.
* Guard chain: no snapshot -> ``SNAPSHOT_FAILED``; no approval -> denied;
  read-only session -> ``WORKBOOK_READ_ONLY``.
* Non-mutating reads submit ``READ_ONLY``; mutations submit ``MUTATING``.
"""

from __future__ import annotations

import re
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.names_adapter import NamesAdapter
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ErrorCode,
    ExcelVbaMcpError,
    NameConflictError,
    category_for,
)
from excel_vba_mcp.domain.models import DefinedNameInfo
from excel_vba_mcp.domain.statuses import (
    NameScope,
    NameTargetKind,
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
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph — Names / Name / Name.Parent / Name.RefersToRange
# ---------------------------------------------------------------------------
# Matches a simple same-workbook range reference (``=Sheet!$A$1:$B$2``) so the
# fake can re-resolve ``RefersToRange`` after a create/update, mirroring Excel.
_RANGE_REF = re.compile(
    r"^=(?:'([^']+)'|([^'!]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$"
)


def _maybe_range(owner: Any, refers_to: str) -> FakeRange | None:
    """Return a ``FakeRange`` when ``refers_to`` is a resolvable range reference."""

    match = _RANGE_REF.match(refers_to.strip())
    if match is None:
        return None
    sheet_name = match.group(1) or match.group(2)
    finder = getattr(owner, "_find_sheet", None)
    if finder is None:
        return None
    sheet = finder(sheet_name)
    if sheet is None:
        return None
    return FakeRange(sheet, match.group(3))


class FakeRange:
    """A fake ``Range`` proxy carrying only the sheet + address."""

    def __init__(self, sheet: FakeSheet, address: str) -> None:
        self._sheet = sheet
        self.Address = address

    @property
    def Worksheet(self) -> FakeSheet:  # noqa: N802 - COM casing
        return self._sheet


class FakeName:
    """A fake ``Name`` object; ``RefersToRange`` either returns or raises."""

    def __init__(
        self,
        name: str,
        refers_to: str,
        parent: Any,
        *,
        name_local: str | None = None,
        visible: bool = True,
        refers_to_range: FakeRange | BaseException | None = None,
    ) -> None:
        self.Name = name
        self.NameLocal = name_local if name_local is not None else name
        self._refers_to = refers_to
        self.Visible = visible
        self.Parent = parent
        self._refers_to_range = refers_to_range
        self._collections: list[FakeNames] = []
        self.deleted = False

    @property
    def RefersTo(self) -> str:  # noqa: N802 - COM casing
        return self._refers_to

    @RefersTo.setter
    def RefersTo(self, value: str) -> None:  # noqa: N802 - COM casing
        self._refers_to = value
        resolved = _maybe_range(self.Parent, value)
        if resolved is not None:
            self._refers_to_range = resolved

    @property
    def RefersToLocal(self) -> str:  # noqa: N802 - COM casing
        return self._refers_to

    @property
    def RefersToRange(self) -> FakeRange:  # noqa: N802 - COM casing
        target = self._refers_to_range
        if isinstance(target, BaseException):
            raise target
        if target is None:
            raise RuntimeError("RefersToRange is not available")
        return target

    def attach(self, collection: FakeNames) -> None:
        """Record that this name is listed by ``collection`` (Excel dual listing)."""

        self._collections.append(collection)

    def Delete(self) -> None:  # noqa: N802 - COM
        self.deleted = True
        for collection in self._collections:
            collection.remove(self)
        self._collections = []


class FakeNames:
    """A fake ``Names`` collection owned by a workbook or worksheet."""

    def __init__(self, owner: Any) -> None:
        self._owner = owner
        self._items: list[FakeName] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Add(self, Name: str, RefersTo: str) -> FakeName:  # noqa: N802, N803 - COM
        if any(existing.Name.lower() == Name.lower() for existing in self._items):
            raise RuntimeError(f"duplicate name {Name}")
        created = FakeName(Name, RefersTo, self._owner)
        resolved = _maybe_range(self._owner, RefersTo)
        if resolved is not None:
            created._refers_to_range = resolved
        created.attach(self)
        self._items.append(created)
        return created

    def Item(self, key: Any) -> FakeName:  # noqa: N802 - COM casing
        if isinstance(key, int):
            return self._items[key - 1]
        for existing in self._items:
            if existing.Name.lower() == str(key).lower():
                return existing
        raise KeyError(key)

    def __call__(self, key: Any) -> FakeName:
        return self.Item(key)

    def __iter__(self):
        return iter(list(self._items))

    def remove(self, name: FakeName) -> None:
        self._items = [existing for existing in self._items if existing is not name]

    def append(self, name: FakeName) -> None:
        """Test helper: attach a pre-built name (emulates Excel's dual listing)."""

        name.attach(self)
        self._items.append(name)


class FakeSheet:
    """A fake worksheet (deliberately without ``FullName`` for scope tests)."""

    def __init__(self, name: str) -> None:
        self.Name = name
        self.Names = FakeNames(self)

    def _find_sheet(self, name: str) -> FakeSheet | None:
        return self if self.Name.lower() == str(name).lower() else None


class FakeSheets:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._sheets)

    def Item(self, key: Any) -> FakeSheet:  # noqa: N802 - COM casing
        if isinstance(key, int):
            return self._sheets[key - 1]
        for sheet in self._sheets:
            if sheet.Name.lower() == str(key).lower():
                return sheet
        raise KeyError(key)

    def __iter__(self):
        return iter(list(self._sheets))


class FakeWorkbook:
    def __init__(self, sheets: list[FakeSheet] | None = None) -> None:
        self.Name = "Book.xlsm"
        self.FullName = "C:/tmp/Book.xlsm"
        self.ReadOnly = False
        self._sheets = sheets if sheets is not None else [FakeSheet("Data")]
        self.Names = FakeNames(self)

    @property
    def Worksheets(self) -> FakeSheets:  # noqa: N802 - COM casing
        return FakeSheets(self._sheets)

    def sheet(self, name: str) -> FakeSheet:
        for sheet in self._sheets:
            if sheet.Name == name:
                return sheet
        raise KeyError(name)

    def _find_sheet(self, name: str) -> FakeSheet | None:
        for sheet in self._sheets:
            if sheet.Name.lower() == str(name).lower():
                return sheet
        return None


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook([FakeSheet("Data"), FakeSheet("Other")])
        wb.FullName = path
        wb.ReadOnly = bool(kwargs.get("ReadOnly", False))
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
            {"name": getattr(fn, "__name__", "?"), "section": section, "op_class": op_class}
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)


# ---------------------------------------------------------------------------
# Seed helpers
# ---------------------------------------------------------------------------
def _seed_name(
    collection: FakeNames,
    name: str,
    refers_to: str,
    *,
    name_local: str | None = None,
    visible: bool = True,
    refers_to_range: FakeRange | BaseException | None = None,
) -> FakeName:
    created = FakeName(
        name,
        refers_to,
        collection._owner,
        name_local=name_local,
        visible=visible,
        refers_to_range=refers_to_range,
    )
    collection.append(created)
    return created


def _add_range_name(
    collection: FakeNames,
    name: str,
    sheet: FakeSheet,
    address: str,
) -> FakeName:
    return _seed_name(
        collection,
        name,
        f"={sheet.Name}!{address}",
        refers_to_range=FakeRange(sheet, address),
    )


def _populated_workbook() -> FakeWorkbook:
    wb = FakeWorkbook([FakeSheet("Data"), FakeSheet("Other")])
    data = wb.sheet("Data")
    _add_range_name(wb.Names, "WbRange", data, "$A$1:$B$2")
    sheet_name = _add_range_name(data.Names, "SheetRange", data, "$C$1")
    # Excel lists worksheet-scoped names in ``Workbook.Names`` too; emulate that
    # so the adapter's Parent-based de-duplication is exercised.
    wb.Names.append(sheet_name)
    _seed_name(wb.Names, "Pi", "=3.14")
    _seed_name(wb.Names, "Greeting", '="hi"')
    _seed_name(wb.Names, "Flag", "=TRUE")
    _seed_name(wb.Names, "Total", "=SUM(Data!$A$1:$A$2)")
    _seed_name(wb.Names, "Broken", "=#REF!")
    _seed_name(wb.Names, "Ext", "=[Book1.xlsx]Sheet1!$A$1")
    _seed_name(wb.Names, "HiddenName", "=1", visible=False)
    return wb


# ---------------------------------------------------------------------------
# Adapter: inventory + target_kind
# ---------------------------------------------------------------------------
def test_list_names_maps_fields_and_target_kinds() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()

    infos = adapter.list_names(wb)
    by_key = {(info.name, info.scope): info for info in infos}

    wb_range = by_key[("WbRange", NameScope.WORKBOOK)]
    assert wb_range.scope is NameScope.WORKBOOK
    assert wb_range.sheet_name is None
    assert wb_range.target_kind is NameTargetKind.RANGE
    assert wb_range.target_sheet == "Data"
    assert wb_range.target_address == "$A$1:$B$2"
    assert wb_range.refers_to == "=Data!$A$1:$B$2"

    sheet_range = by_key[("SheetRange", NameScope.WORKSHEET)]
    assert sheet_range.scope is NameScope.WORKSHEET
    assert sheet_range.sheet_name == "Data"
    assert sheet_range.target_kind is NameTargetKind.RANGE
    assert sheet_range.target_sheet == "Data"
    assert sheet_range.target_address == "$C$1"

    assert by_key[("Pi", NameScope.WORKBOOK)].target_kind is NameTargetKind.CONSTANT
    assert by_key[("Greeting", NameScope.WORKBOOK)].target_kind is NameTargetKind.CONSTANT
    assert by_key[("Flag", NameScope.WORKBOOK)].target_kind is NameTargetKind.CONSTANT
    assert by_key[("Total", NameScope.WORKBOOK)].target_kind is NameTargetKind.FORMULA
    assert by_key[("Broken", NameScope.WORKBOOK)].target_kind is NameTargetKind.ERROR
    assert by_key[("Ext", NameScope.WORKBOOK)].target_kind is NameTargetKind.EXTERNAL
    assert by_key[("HiddenName", NameScope.WORKBOOK)].visible is False

    # The worksheet-scoped name listed in both ``wb.Names`` and ``ws.Names`` is
    # returned exactly once (no duplicate scope entry).
    assert sum(1 for info in infos if info.name == "SheetRange") == 1


def test_list_names_local_and_visible_defaults() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    _seed_name(wb.Names, "Localized", "=1", name_local="Локализованное")

    info = NamesAdapter().list_names(wb)[0]
    assert info.name == "Localized"
    assert info.name_local == "Локализованное"
    assert info.visible is True


def test_worksheet_scope_name_is_sheet_qualified_by_excel() -> None:
    """Excel reports sheet-scoped ``Name``/``NameLocal`` as ``Sheet!Name``."""

    wb = FakeWorkbook([FakeSheet("Data")])
    data = wb.sheet("Data")
    created = _seed_name(
        data.Names, "SheetRange", "=Data!$C$1", refers_to_range=FakeRange(data, "$C$1")
    )
    # Excel's real behaviour: the Name object carries the sheet qualifier.
    created.Name = "Data!SheetRange"
    created.NameLocal = "Data!SheetRange"

    infos = NamesAdapter().list_names(wb)
    assert len(infos) == 1
    assert infos[0].name == "SheetRange"
    assert infos[0].name_local == "SheetRange"
    assert infos[0].scope is NameScope.WORKSHEET
    assert infos[0].sheet_name == "Data"

    adapter = NamesAdapter()
    resolved = adapter.get_name(wb, "SheetRange", NameScope.WORKSHEET, "Data")
    assert resolved.name == "SheetRange"


def test_worksheet_scope_name_with_quoted_sheet_prefix() -> None:
    """A quoted ``'Sheet Name'!Name`` prefix is stripped to the identifier."""

    wb = FakeWorkbook([FakeSheet("Data Sheet")])
    data = wb.sheet("Data Sheet")
    created = _seed_name(data.Names, "Local", "=Data!$A$1", refers_to_range=FakeRange(data, "$A$1"))
    created.Name = "'Data Sheet'!Local"
    created.NameLocal = "'Data Sheet'!Local"

    info = NamesAdapter().list_names(wb)[0]
    assert info.name == "Local"
    assert info.sheet_name == "Data Sheet"


def test_get_name_scopes() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()

    wb_info = adapter.get_name(wb, "WbRange", NameScope.WORKBOOK)
    assert wb_info.scope is NameScope.WORKBOOK
    assert wb_info.target_kind is NameTargetKind.RANGE

    sheet_info = adapter.get_name(wb, "SheetRange", NameScope.WORKSHEET, "Data")
    assert sheet_info.scope is NameScope.WORKSHEET
    assert sheet_info.sheet_name == "Data"

    # Case-insensitive resolution.
    assert adapter.get_name(wb, "wbrange", NameScope.WORKBOOK).name == "WbRange"


def test_get_name_missing_raises() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.get_name(wb, "NoSuchName", NameScope.WORKBOOK)


def test_get_name_missing_is_not_found() -> None:
    # A missing defined name is a caller-addressed missing identity: NOT_FOUND
    # (not the generic COM_BUSY / transient_com_failure fallback).
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.get_name(wb, "NoSuchName", NameScope.WORKBOOK)
    assert exc_info.value.code is ErrorCode.NOT_FOUND
    assert category_for(exc_info.value.code) == "not_found"
    assert exc_info.value.context["reason"] == "not_found"
    assert exc_info.value.context["name"] == "NoSuchName"


def test_worksheet_scope_missing_sheet_is_not_found() -> None:
    # A worksheet-scoped name whose owning sheet does not exist is a missing
    # identity: NOT_FOUND (not COM_BUSY).
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.create_name(
            wb, "Local", "=1", scope=NameScope.WORKSHEET, sheet="NoSuchSheet"
        )
    assert exc_info.value.code is ErrorCode.NOT_FOUND
    assert category_for(exc_info.value.code) == "not_found"
    assert exc_info.value.context["reason"] == "not_found"
    assert exc_info.value.context["sheet"] == "NoSuchSheet"


def test_get_name_worksheet_requires_sheet() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.get_name(wb, "SheetRange", NameScope.WORKSHEET)


# ---------------------------------------------------------------------------
# Adapter: scoped CRUD + conflicts
# ---------------------------------------------------------------------------
def test_create_workbook_scope_name() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()

    created = adapter.create_name(wb, "NewRange", "Data!$A$1", scope=NameScope.WORKBOOK)
    assert created.name == "NewRange"
    assert created.scope is NameScope.WORKBOOK
    assert created.refers_to == "=Data!$A$1"
    assert created.target_kind is NameTargetKind.RANGE
    assert adapter.get_name(wb, "NewRange", NameScope.WORKBOOK).refers_to == "=Data!$A$1"


def test_create_worksheet_scope_name() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()

    created = adapter.create_name(
        wb, "LocalRange", "Data!$C$1", scope=NameScope.WORKSHEET, sheet="Data"
    )
    assert created.scope is NameScope.WORKSHEET
    assert created.sheet_name == "Data"
    assert adapter.get_name(wb, "LocalRange", NameScope.WORKSHEET, "Data").name == "LocalRange"


def test_create_duplicate_same_scope_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()
    adapter.create_name(wb, "Dup", "=1", scope=NameScope.WORKBOOK)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.create_name(wb, "Dup", "=2", scope=NameScope.WORKBOOK)
    assert exc_info.value.context["reason"] == "duplicate"


def test_create_cross_scope_conflict_rejected() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()
    adapter.create_name(wb, "Shared", "=1", scope=NameScope.WORKBOOK)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.create_name(wb, "Shared", "=Data!$A$1", scope=NameScope.WORKSHEET, sheet="Data")
    assert exc_info.value.context["reason"] == "scope_conflict"


def test_create_worksheet_scope_requires_sheet() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.create_name(wb, "NoSheet", "=1", scope=NameScope.WORKSHEET)


def test_update_name_preserves_scope() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()

    updated = adapter.update_name(wb, "WbRange", "Data!$D$1:$D$4", scope=NameScope.WORKBOOK)
    assert updated.scope is NameScope.WORKBOOK
    assert updated.refers_to == "=Data!$D$1:$D$4"
    assert updated.target_kind is NameTargetKind.RANGE
    assert updated.target_address == "$D$1:$D$4"

    updated_sheet = adapter.update_name(
        wb, "SheetRange", "Data!$E$1", scope=NameScope.WORKSHEET, sheet="Data"
    )
    assert updated_sheet.scope is NameScope.WORKSHEET
    assert updated_sheet.sheet_name == "Data"


def test_update_missing_name_rejected() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.update_name(wb, "NoSuchName", "=1", scope=NameScope.WORKBOOK)


def test_delete_name_removes_it() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()

    deleted = adapter.delete_name(wb, "WbRange", scope=NameScope.WORKBOOK)
    assert deleted.name == "WbRange"
    assert deleted.scope is NameScope.WORKBOOK
    with pytest.raises(ExcelVbaMcpError):
        adapter.get_name(wb, "WbRange", NameScope.WORKBOOK)

    adapter.delete_name(wb, "SheetRange", scope=NameScope.WORKSHEET, sheet="Data")
    with pytest.raises(ExcelVbaMcpError):
        adapter.get_name(wb, "SheetRange", NameScope.WORKSHEET, "Data")


def test_delete_missing_name_rejected() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()
    with pytest.raises(ExcelVbaMcpError):
        adapter.delete_name(wb, "NoSuchName", scope=NameScope.WORKBOOK)


# ---------------------------------------------------------------------------
# Service-level rig
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
        self.service = NamesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

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
            str(self.original), mode=SessionMode.INSPECT, read_only=read_only
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

    def add_snapshot(self, workbook_id: str) -> None:
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

    def workbook(self) -> FakeWorkbook:
        for wb in self.app.Workbooks:
            return wb
        raise AssertionError("no workbook open")


async def test_service_list_and_get(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    wid = info.workbook_id
    data = rig.workbook().sheet("Data")
    _add_range_name(rig.workbook().Names, "WbRange", data, "$A$1:$B$2")

    names = await rig.service.list_names(wid)
    assert any(n.name == "WbRange" and n.scope is NameScope.WORKBOOK for n in names)

    one = await rig.service.get_name(wid, "WbRange", NameScope.WORKBOOK)
    assert one.target_kind is NameTargetKind.RANGE

    read_submit = next(s for s in rig.worker.submits if s["name"] == "names_list")
    assert read_submit["op_class"] == OperationClass.READ_ONLY


async def test_service_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.list_names("missing")


async def test_create_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("name_create")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.create_name(info.workbook_id, "X", "=1", scope=NameScope.WORKBOOK)
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    assert [s for s in rig.worker.submits if s["name"] == "name_create"] == []


async def test_create_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await rig.service.create_name(info.workbook_id, "X", "=1", scope=NameScope.WORKBOOK)
    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED
    assert [s for s in rig.worker.submits if s["name"] == "name_create"] == []


async def test_create_requires_writable(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.add_snapshot(info.workbook_id)
    rig.grant("name_create")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.create_name(info.workbook_id, "X", "=1", scope=NameScope.WORKBOOK)
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_create_and_delete_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    wid = info.workbook_id
    rig.add_snapshot(wid)
    rig.grant("name_create")
    rig.grant("name_delete")

    created = await rig.service.create_name(wid, "X", "=Data!$A$1", scope=NameScope.WORKBOOK)
    assert isinstance(created, DefinedNameInfo)
    assert created.scope is NameScope.WORKBOOK
    assert created.refers_to == "=Data!$A$1"
    mutate = next(s for s in rig.worker.submits if s["name"] == "name_create")
    assert mutate["op_class"] == OperationClass.MUTATING

    deleted = await rig.service.delete_name(wid, "X", scope=NameScope.WORKBOOK)
    assert deleted.name == "X"
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.get_name(wid, "X", NameScope.WORKBOOK)


async def test_update_success(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    wid = info.workbook_id
    data = rig.workbook().sheet("Data")
    _add_range_name(rig.workbook().Names, "WbRange", data, "$A$1")
    rig.add_snapshot(wid)
    rig.grant("name_update")

    updated = await rig.service.update_name(
        wid, "WbRange", "=Data!$B$1:$B$2", scope=NameScope.WORKBOOK
    )
    assert updated.refers_to == "=Data!$B$1:$B$2"
    assert updated.target_kind is NameTargetKind.RANGE


# ---------------------------------------------------------------------------
# WP1: raise-site error taxonomy — deterministic cases must NOT be COM_BUSY
# ---------------------------------------------------------------------------
def _assert_non_retryable_taxonomy(
    error: ExcelVbaMcpError, *, code: str, category: str
) -> None:
    """Assert the serialized envelope is a specific, non-retryable, corrective error."""

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


def test_invalid_scope_is_invalid_range() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        NamesAdapter().get_name(wb, "WbRange", "bogus_scope")  # type: ignore[arg-type]
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_scope"


def test_missing_worksheet_for_sheet_scope_is_not_found() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        NamesAdapter().create_name(
            wb, "Local", "=1", scope=NameScope.WORKSHEET, sheet="NoSuchSheet"
        )
    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context["reason"] == "not_found"


def test_invalid_name_syntax_is_invalid_range() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        NamesAdapter().create_name(wb, "   ", "=1", scope=NameScope.WORKBOOK)
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_name"


def test_invalid_refers_to_is_invalid_range() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        NamesAdapter().create_name(wb, "X", "   ", scope=NameScope.WORKBOOK)
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "invalid_refers_to"


def test_duplicate_same_scope_is_name_conflict() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()
    adapter.create_name(wb, "Dup", "=1", scope=NameScope.WORKBOOK)
    with pytest.raises(NameConflictError) as exc_info:
        adapter.create_name(wb, "Dup", "=2", scope=NameScope.WORKBOOK)
    _assert_non_retryable_taxonomy(exc_info.value, code="NAME_CONFLICT", category="validation")
    assert exc_info.value.context["reason"] == "duplicate"


def test_cross_scope_conflict_is_name_conflict() -> None:
    wb = FakeWorkbook([FakeSheet("Data")])
    adapter = NamesAdapter()
    adapter.create_name(wb, "Shared", "=1", scope=NameScope.WORKBOOK)
    with pytest.raises(NameConflictError) as exc_info:
        adapter.create_name(wb, "Shared", "=Data!$A$1", scope=NameScope.WORKSHEET, sheet="Data")
    _assert_non_retryable_taxonomy(exc_info.value, code="NAME_CONFLICT", category="validation")
    assert exc_info.value.context["reason"] == "scope_conflict"


def test_ambiguous_name_resolution_is_invalid_range() -> None:
    wb = _populated_workbook()
    adapter = NamesAdapter()
    dup1 = FakeName("Dup", "=1", wb)
    dup2 = FakeName("Dup", "=2", wb)
    # Force two same-scope matches so the ambiguity guard is exercised.
    adapter._iter_entries = lambda _wb: iter(  # type: ignore[method-assign]
        [(dup1, NameScope.WORKBOOK, None), (dup2, NameScope.WORKBOOK, None)]
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.get_name(wb, "Dup", NameScope.WORKBOOK)
    _assert_non_retryable_taxonomy(exc_info.value, code="INVALID_RANGE", category="validation")
    assert exc_info.value.context["reason"] == "ambiguous_name"


def test_missing_name_is_not_found() -> None:
    wb = _populated_workbook()
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        NamesAdapter().get_name(wb, "NoSuchName", NameScope.WORKBOOK)
    _assert_non_retryable_taxonomy(exc_info.value, code="NOT_FOUND", category="not_found")
    assert exc_info.value.context["reason"] == "not_found"
