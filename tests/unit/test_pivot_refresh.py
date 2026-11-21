"""Unit tests for the targeted PivotTable / PivotCache refresh (task_049).

These tests exercise the production
:class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService` and
:class:`~excel_vba_mcp.adapters.excel.pivot_tables_adapter.PivotTablesAdapter`
against a pure-Python fake ``Workbook`` object graph (a fake STA worker that runs
the submitted callable inline). No Excel / COM is required.

Covered:

* ``refresh_pivot_table`` dispatch — ``PivotTable.RefreshTable`` is invoked (not
  ``PivotCache.Refresh``) and its Boolean return is reported;
* ``refresh_pivot_cache`` dispatch — ``PivotCache.Refresh`` is invoked and the
  shared ``RefreshDate`` advances;
* affected-table reporting — every table sharing the cache is listed, for both
  entry points;
* ``RefreshDate`` confirmation — a refresh whose ``RefreshDate`` does not advance
  is reported as ``failed`` (not success by absence of an exception);
* unsupported source rejection — external / OLAP caches return
  ``unsupported_source`` and are never refreshed;
* the guard chain — read-only workbook, missing snapshot and missing approval
  each block the mutation before any STA dispatch.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.pivot_tables_adapter import PivotTablesAdapter
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import PivotRefreshResult
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    PivotRefreshMethod,
    PivotRefreshStatus,
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

# ``XlPivotFieldOrientation`` / ``XlPivotTableSourceType`` values (adapter maps).
XL_HIDDEN = 0
XL_ROW = 1
XL_SOURCE_DATABASE = 1
XL_SOURCE_EXTERNAL = 2

#: Approval action names used by the service (must match the service constants).
_ACTION_REFRESH_TABLE = "pivot_table_refresh"
_ACTION_REFRESH_CACHE = "pivot_cache_refresh"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel PivotTables object graph (refresh-aware)
# ---------------------------------------------------------------------------
class FakePivotField:
    def __init__(self, source_name: str, *, orientation: int = XL_HIDDEN) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = source_name
        self.Orientation = orientation
        self.Position = -2146826246
        self.ShowAllItems = True
        self.AllItemsVisible = True

    def PivotItems(self) -> list[Any]:  # noqa: N802 - COM casing
        return []


class FakePivotFields:
    def __init__(self, fields: list[FakePivotField]) -> None:
        self._fields = fields

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> FakePivotField:  # noqa: N802 - COM casing
        return self._fields[index - 1]


class FakeDataFields:
    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return 0


class FakePivotCache:
    """A fake ``PivotCache`` whose ``Refresh`` advances ``RefreshDate``.

    ``advance_on_refresh`` controls whether ``Refresh`` moves the date forward;
    ``False`` models a refresh Excel accepted but that produced no state change
    (used to prove success is not inferred from the absence of an exception).
    """

    def __init__(
        self,
        index: int,
        *,
        source_type: int = XL_SOURCE_DATABASE,
        source_data: Any = "Data!R1C1:R6C5",
        refresh_date: datetime | None = None,
        olap: bool = False,
        advance_on_refresh: bool = True,
    ) -> None:
        self.Index = index
        self.SourceType = source_type
        self.SourceData = source_data
        self.RefreshDate = refresh_date
        self.RefreshOnFileOpen = False
        self.EnableRefresh = True
        self.RecordCount = 6
        self.OLAP = olap
        self._advance = advance_on_refresh
        self.refresh_count = 0

    def Refresh(self) -> None:  # noqa: N802 - COM casing
        self.refresh_count += 1
        if self._advance:
            base = self.RefreshDate or datetime(2026, 1, 1, tzinfo=UTC)
            self.RefreshDate = base + timedelta(seconds=1)


class FakeRange:
    def __init__(self, address: str) -> None:
        self.Address = address


class FakePivotTable:
    """A fake ``PivotTable`` whose ``RefreshTable`` returns a Boolean.

    ``refresh_returns`` models Excel's ``RefreshTable`` return value and
    ``advance_on_refresh`` models whether the refresh moved ``RefreshDate``.
    """

    def __init__(
        self,
        name: str,
        sheet_name: str,
        cache: FakePivotCache,
        *,
        index: int = 1,
        refresh_date: datetime | None = None,
        refresh_returns: bool = True,
        advance_on_refresh: bool = True,
        fields: list[FakePivotField] | None = None,
    ) -> None:
        self.Name = name
        self._sheet_name = sheet_name
        self._cache = cache
        self.Index = index
        self.CacheIndex = cache.Index
        self.TableRange1 = FakeRange("$A$1:$B$7")
        self.TableRange2 = FakeRange("$A$3:$B$7")
        self.DataBodyRange = FakeRange("$B$5:$B$7")
        self.RowGrand = True
        self.ColumnGrand = True
        self.RefreshDate = refresh_date
        self._fields = fields or []
        self._refresh_returns = refresh_returns
        self._advance = advance_on_refresh
        self.refresh_table_count = 0

    @property
    def Parent(self) -> FakeSheet:  # noqa: N802 - COM casing
        return FakeSheet(self._sheet_name)

    def PivotCache(self) -> FakePivotCache:  # noqa: N802 - COM casing
        return self._cache

    def PivotFields(self, *_args: Any) -> FakePivotFields:  # noqa: N802 - COM casing
        return FakePivotFields(self._fields)

    @property
    def DataFields(self) -> FakeDataFields:  # noqa: N802 - COM casing
        return FakeDataFields()

    def RefreshTable(self) -> bool:  # noqa: N802 - COM casing
        self.refresh_table_count += 1
        if self._advance:
            base = self.RefreshDate or datetime(2026, 1, 1, tzinfo=UTC)
            self.RefreshDate = base + timedelta(seconds=1)
        return self._refresh_returns


class FakePivotTables:
    def __init__(self, pivots: list[FakePivotTable]) -> None:
        self._pivots = pivots

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._pivots)

    def Item(self, index: int) -> FakePivotTable:  # noqa: N802 - COM casing
        return self._pivots[index - 1]

    def __iter__(self):
        return iter(self._pivots)


class FakeSheet:
    def __init__(self, name: str, pivots: list[FakePivotTable] | None = None) -> None:
        self.Name = name
        self._pivots = pivots or []

    def PivotTables(self) -> FakePivotTables:  # noqa: N802 - COM casing
        return FakePivotTables(self._pivots)


class FakeSheetCollection:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets

    def Item(self, key: Any) -> FakeSheet:  # noqa: N802 - COM casing
        if isinstance(key, int):
            return self._sheets[key - 1]
        for sheet in self._sheets:
            if sheet.Name == key:
                return sheet
        raise KeyError(key)

    def __iter__(self):
        return iter(self._sheets)


class FakeWorkbook:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets
        self.FullName = ""
        self.ReadOnly = False

    @property
    def Worksheets(self) -> FakeSheetCollection:  # noqa: N802 - COM casing
        return FakeSheetCollection(self._sheets)

    def sheet(self, name: str) -> FakeSheet:
        for sheet in self._sheets:
            if sheet.Name == name:
                return sheet
        raise KeyError(name)


class FakeWorkbooks:
    def __init__(self) -> None:
        self.items: list[FakeWorkbook] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM
        wb = FakeWorkbook([FakeSheet("Data")])
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
# Rig — real service stack, fake worker/Excel, real snapshot + approval
# ---------------------------------------------------------------------------
class Rig:
    def __init__(
        self,
        tmp_path: Path,
        *,
        wire_snapshot: bool = True,
        wire_approval: bool = True,
    ) -> None:
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
        self.service = PivotTablesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.audit,
            snapshot_service=self.snapshot if wire_snapshot else None,
            approval=self.approval if wire_approval else None,
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


# ---------------------------------------------------------------------------
# Fixture builders
# ---------------------------------------------------------------------------
def _refresh_submits(rig: Rig) -> list[dict[str, Any]]:
    """Return only the refresh dispatches (ignores the session's open_workbook)."""

    return [
        s for s in rig.worker.submits if s["name"] in ("pivot_table_refresh", "pivot_cache_refresh")
    ]


def _seed_pivots(
    wb: FakeWorkbook,
    *,
    shared_advance: bool = True,
    external: bool = False,
    olap: bool = False,
) -> tuple[FakePivotCache, FakePivotTable, FakePivotTable, FakePivotTable]:
    """Seed ptA/ptB (cache 1, shared) and ptC (cache 2), plus an optional ptExt.

    Returns ``(cache1, ptA, ptB, ptC)``.
    """

    cache1 = FakePivotCache(
        1,
        refresh_date=datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC),
        advance_on_refresh=shared_advance,
    )
    pt_a = FakePivotTable(
        "ptA",
        "PivotReport",
        cache1,
        index=1,
        refresh_date=datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC),
        fields=[FakePivotField("Region", orientation=XL_ROW)],
    )
    pt_b = FakePivotTable(
        "ptB",
        "PivotReport",
        cache1,
        index=2,
        refresh_date=datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC),
        fields=[FakePivotField("Region", orientation=XL_ROW)],
    )
    cache2 = FakePivotCache(2, source_data="TableData!R1C1:R6C5")
    pt_c = FakePivotTable(
        "ptC",
        "PivotReport",
        cache2,
        index=3,
        refresh_date=datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC),
        fields=[FakePivotField("Product", orientation=XL_ROW)],
    )
    report_pivots = [pt_a, pt_b, pt_c]
    if external or olap:
        cache_ext = FakePivotCache(
            3, source_type=XL_SOURCE_EXTERNAL, olap=olap, source_data="OLAP$" * 3
        )
        report_pivots.append(
            FakePivotTable(
                "ptExt",
                "PivotReport",
                cache_ext,
                index=4,
                fields=[FakePivotField("Region", orientation=XL_ROW)],
            )
        )
    wb._sheets = [FakeSheet("Data"), FakeSheet("PivotReport", report_pivots)]
    return cache1, pt_a, pt_b, pt_c


# ---------------------------------------------------------------------------
# Adapter-level dispatch
# ---------------------------------------------------------------------------
def test_refresh_pivot_table_invokes_refreshtable() -> None:
    wb = FakeWorkbook([])
    cache, pt_a, _pt_b, _pt_c = _seed_pivots(wb)
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_table(wb, "ptA")

    assert pt_a.refresh_table_count == 1
    assert cache.refresh_count == 0  # PivotCache.Refresh must NOT be used here
    assert result.method is PivotRefreshMethod.REFRESH_TABLE
    assert result.refresh_return is True
    assert result.refreshed is True
    assert result.status is PivotRefreshStatus.REFRESHED
    assert result.refresh_date_after is not None
    assert result.refresh_date_after > result.refresh_date_before  # type: ignore[operator]


def test_refresh_pivot_cache_invokes_cache_refresh() -> None:
    wb = FakeWorkbook([])
    cache, pt_a, pt_b, _pt_c = _seed_pivots(wb)
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_cache(wb, 1)

    assert cache.refresh_count == 1
    assert pt_a.refresh_table_count == 0  # RefreshTable must NOT be used here
    assert pt_b.refresh_table_count == 0
    assert result.method is PivotRefreshMethod.REFRESH_CACHE
    assert result.refresh_return is None  # PivotCache.Refresh returns no Boolean
    assert result.refreshed is True
    assert result.status is PivotRefreshStatus.REFRESHED


# ---------------------------------------------------------------------------
# Affected-table reporting
# ---------------------------------------------------------------------------
def test_refresh_table_reports_every_table_sharing_the_cache() -> None:
    wb = FakeWorkbook([])
    _cache, _pt_a, _pt_b, _pt_c = _seed_pivots(wb)
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_table(wb, "ptA")

    # ptA and ptB share cache 1; ptC is on cache 2.
    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]
    assert "ptC" not in result.affected_pivot_tables


def test_refresh_cache_reports_every_table_sharing_the_cache() -> None:
    wb = FakeWorkbook([])
    _cache, _pt_a, _pt_b, _pt_c = _seed_pivots(wb)
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_cache(wb, 1)

    assert sorted(result.affected_pivot_tables) == ["ptA", "ptB"]


# ---------------------------------------------------------------------------
# RefreshDate confirmation (no success by absence of an exception)
# ---------------------------------------------------------------------------
def test_refresh_table_without_date_advance_is_failed() -> None:
    wb = FakeWorkbook([])
    _cache, pt_a, _pt_b, _pt_c = _seed_pivots(wb, shared_advance=False)
    # ptA explicitly returns True but its RefreshDate never advances -> the
    # Boolean alone is the truth signal here, so it is still a success.
    pt_a._refresh_returns = True
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_table(wb, "ptA")

    assert result.refresh_return is True
    assert result.refreshed is True  # Boolean truth signal confirms the refresh


def test_refresh_table_false_and_no_advance_is_failed() -> None:
    wb = FakeWorkbook([])
    _cache, pt_a, _pt_b, _pt_c = _seed_pivots(wb)
    pt_a._refresh_returns = False
    pt_a._advance = False
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_table(wb, "ptA")

    assert result.refresh_return is False
    assert result.refreshed is False
    assert result.status is PivotRefreshStatus.FAILED
    assert result.error is not None
    assert result.error["code"] == "REFRESH_FAILED"


def test_refresh_cache_without_date_advance_is_failed() -> None:
    wb = FakeWorkbook([])
    _cache, _pt_a, _pt_b, _pt_c = _seed_pivots(wb, shared_advance=False)
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_cache(wb, 1)

    assert result.refreshed is False
    assert result.status is PivotRefreshStatus.FAILED
    assert result.error is not None
    assert result.error["code"] == "REFRESH_FAILED"


# ---------------------------------------------------------------------------
# Unsupported source rejection
# ---------------------------------------------------------------------------
def test_refresh_external_source_is_unsupported_and_not_invoked() -> None:
    wb = FakeWorkbook([])
    _cache1, _pt_a, _pt_b, _pt_c = _seed_pivots(wb, external=True)
    ext = next(pt for pt in wb.sheet("PivotReport")._pivots if pt.Name == "ptExt")
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_table(wb, "ptExt")

    assert result.status is PivotRefreshStatus.UNSUPPORTED_SOURCE
    assert result.refreshed is False
    assert ext.refresh_table_count == 0  # never invoked
    assert result.error is not None
    assert result.error["code"] == "UNSUPPORTED_PIVOT_SOURCE"
    assert result.error["source_type"] == "external"


def test_refresh_olap_source_is_unsupported_and_not_invoked() -> None:
    wb = FakeWorkbook([])
    _cache1, _pt_a, _pt_b, _pt_c = _seed_pivots(wb, olap=True)
    ext = next(pt for pt in wb.sheet("PivotReport")._pivots if pt.Name == "ptExt")
    adapter = PivotTablesAdapter()

    result = adapter.refresh_pivot_cache(wb, ext.CacheIndex)

    assert result.status is PivotRefreshStatus.UNSUPPORTED_SOURCE
    assert result.refreshed is False
    assert result.error is not None
    assert result.error["code"] == "UNSUPPORTED_PIVOT_SOURCE"


# ---------------------------------------------------------------------------
# Service-level: dispatch, guard chain, audit
# ---------------------------------------------------------------------------
async def test_service_refresh_table_dispatches_mutating(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant(_ACTION_REFRESH_TABLE)
    _seed_pivots(rig.workbook())

    result = await rig.service.refresh_pivot_table(info.workbook_id, "ptA")

    assert isinstance(result, PivotRefreshResult)
    assert result.status is PivotRefreshStatus.REFRESHED
    names = [s["name"] for s in rig.worker.submits]
    assert "pivot_table_refresh" in names
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_table_refresh")
    assert submit["op_class"] is OperationClass.MUTATING


async def test_service_refresh_cache_dispatches_mutating(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant(_ACTION_REFRESH_CACHE)
    _seed_pivots(rig.workbook())

    result = await rig.service.refresh_pivot_cache(info.workbook_id, 1)

    assert result.status is PivotRefreshStatus.REFRESHED
    submit = next(s for s in rig.worker.submits if s["name"] == "pivot_cache_refresh")
    assert submit["op_class"] is OperationClass.MUTATING


async def test_service_refresh_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant(_ACTION_REFRESH_TABLE)
    _seed_pivots(rig.workbook())

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.refresh_pivot_table(info.workbook_id, "ptA")

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED
    # No STA dispatch happened: the guard blocked it first.
    assert _refresh_submits(rig) == []


async def test_service_refresh_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    _seed_pivots(rig.workbook())

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.refresh_pivot_cache(info.workbook_id, 1)

    assert exc_info.value.code is ErrorCode.COM_BUSY
    assert _refresh_submits(rig) == []


async def test_service_refresh_rejects_read_only_workbook(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)
    rig.add_snapshot(info.workbook_id)
    rig.grant(_ACTION_REFRESH_TABLE)
    _seed_pivots(rig.workbook())

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.refresh_pivot_table(info.workbook_id, "ptA")

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY
    assert _refresh_submits(rig) == []


async def test_service_refresh_fails_closed_without_snapshot_service(tmp_path: Path) -> None:
    rig = Rig(tmp_path, wire_snapshot=False)
    info = await rig.open()
    rig.grant(_ACTION_REFRESH_TABLE)
    _seed_pivots(rig.workbook())

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.refresh_pivot_table(info.workbook_id, "ptA")

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_service_refresh_records_mutating_audit(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant(_ACTION_REFRESH_TABLE)
    _seed_pivots(rig.workbook())

    await rig.service.refresh_pivot_table(info.workbook_id, "ptA")

    entries = (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
    refresh_entries = [
        line for line in entries if "pivot_table_refresh" in line and "mutating" in line
    ]
    assert refresh_entries, "expected a mutating audit record for the refresh"
