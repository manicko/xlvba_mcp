"""Unit tests for the WP-C PivotTable rendered-result reader.

These tests never touch COM / Excel. The service-level tests install a fake
adapter (returning a canned :class:`PivotResultPage`) and a fake session/worker
so the production :class:`~excel_vba_mcp.application.pivot_tables_service.PivotTablesService`
orchestration is exercised without Excel:

* ``offset`` / ``limit`` / ``cursor_version`` are forwarded to the adapter and
  ``max_cells`` comes from ``large_data.max_pivot_result_cells``;
* a ``cursor_version`` mismatch raises ``STALE_CURSOR``;
* the :class:`~excel_vba_mcp.domain.models.PivotResultPage` DTO rejects an
  oversized ``rows*cols`` grid and a ``len(values) != rows`` mismatch, accepts a
  valid page and preserves the ``fingerprint`` format;
* the pure ``_top_row`` address parser (header-row derivation) is hermetic.

The live COM path (fake PivotTable objects exposing ``TableRange1`` /
``DataBodyRange`` / ``Parent`` / ``RefreshDate`` / ``PivotCache`` plus a fake
sheet ``Range.Value2``) is intentionally not unit-tested here — the Validator
covers it against the real Excel object model.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.pivot_tables_adapter import _top_row
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    MAX_PIVOT_RESULT_CELLS,
    CellValue,
    PageInfo,
    PivotResultPage,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    OperationClass,
    PivotSourceType,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.audit import AuditLogger

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class _FakeWb:
    def __init__(self, full_name: str) -> None:
        self.FullName = full_name


class _FakeWorkbooks:
    def __init__(self, workbooks: list[_FakeWb]) -> None:
        self._workbooks = workbooks

    def __iter__(self):
        return iter(self._workbooks)


class _FakeApp:
    def __init__(self, workbooks: list[_FakeWb]) -> None:
        self.Workbooks = _FakeWorkbooks(workbooks)


class _FakeAppAdapter:
    def __init__(self, app: _FakeApp) -> None:
        self.app = app


class _FakeSession:
    """Minimal session exposing the surface the read path uses (no COM)."""

    def __init__(self, record: WorkbookInfo, version: int) -> None:
        self._record = record
        self._version = version
        self.session_id = "sess"
        self._app_adapter = _FakeAppAdapter(_FakeApp([_FakeWb(record.path)]))

    def list_workbooks_sync(self) -> list[WorkbookInfo]:
        return [self._record]

    def get_version(self, workbook_id: str) -> int:
        return self._version

    def bump_version(self, workbook_id: str) -> None:
        self._version += 1


class _FakeWorker:
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
            {"name": getattr(fn, "__name__", "?"), "op_class": op_class}
        )
        return fn()

    def get_status(self) -> Any:
        return WorkbookStatus()


class WorkbookStatus:
    state = WorkbookSessionState.RUNNING
    accepting = True


class _FakeAdapter:
    """Fake PivotTablesAdapter returning a canned page and recording the call."""

    def __init__(self, page: PivotResultPage) -> None:
        self._page = page
        self.calls: list[dict[str, Any]] = []

    def read_pivot_result(
        self,
        wb: Any,
        pivot_table: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        version_getter: Any = None,
        max_cells: int = MAX_PIVOT_RESULT_CELLS,
    ) -> PivotResultPage:
        self.calls.append(
            {
                "pivot_table": pivot_table,
                "offset": offset,
                "limit": limit,
                "version": version_getter() if version_getter is not None else None,
                "max_cells": max_cells,
            }
        )
        return self._page


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------
def _record() -> WorkbookInfo:
    return WorkbookInfo(
        workbook_id="wb1",
        name="Book.xlsm",
        format="xlsm",
        path="C:/work/Book.xlsm",
        read_only=False,
        has_vba=False,
        excel_version="16.0",
        session_id="sess",
        mode=SessionMode.INSPECT,
    )


def _page(
    version: int = 0,
    *,
    rows: int = 1,
    cols: int = 1,
    table_range: str = "$A$1:$B$7",
    refresh: datetime | None = None,
    fingerprint: str | None = None,
) -> PivotResultPage:
    values = [
        [CellValue(kind=CellValueKind.STRING, value="v") for _ in range(cols)]
        for _ in range(rows)
    ]
    if fingerprint is None:
        fingerprint = f"{table_range}|{refresh.isoformat() if refresh else ''}"
    return PivotResultPage(
        pivot_table="pt",
        sheet_name="PivotReport",
        source_type=PivotSourceType.DATABASE,
        table_range=table_range,
        rows=rows,
        cols=cols,
        values=values,
        page=PageInfo(offset=0, limit=None, returned=rows, total=rows, version=version),
        fingerprint=fingerprint,
    )


def _build_service(
    tmp_path: Path, *, page_version: int, large_data: LargeDataConfig
) -> tuple[PivotTablesService, _FakeAdapter]:
    record = _record()
    session = _FakeSession(record, page_version)
    worker = _FakeWorker()
    audit = AuditLogger(tmp_path / "audit.jsonl")
    service = PivotTablesService(
        worker,  # type: ignore[arg-type]
        session,  # type: ignore[arg-type]
        audit,
        large_data=large_data,
    )
    adapter = _FakeAdapter(_page(page_version))
    service._adapter = adapter  # type: ignore[assignment]
    return service, adapter


# ---------------------------------------------------------------------------
# Service-level: forwarding + max_cells + op_class
# ---------------------------------------------------------------------------
async def test_service_forwards_paging_args_and_max_cells(tmp_path: Path) -> None:
    large_data = LargeDataConfig(max_pivot_result_cells=12_345)
    service, adapter = _build_service(tmp_path, page_version=7, large_data=large_data)

    result = await service.read_pivot_result(
        "wb1", "pt", offset=3, limit=11, cursor_version=7
    )

    assert isinstance(result, PivotResultPage)
    call = adapter.calls[-1]
    assert call["pivot_table"] == "pt"
    assert call["offset"] == 3
    assert call["limit"] == 11
    # version_getter resolves to the session's current content version.
    assert call["version"] == 7
    # max_cells comes from large_data.max_pivot_result_cells.
    assert call["max_cells"] == 12_345


async def test_service_read_is_read_only_op(tmp_path: Path) -> None:
    service, _adapter = _build_service(
        tmp_path, page_version=0, large_data=LargeDataConfig()
    )
    worker = service._worker

    await service.read_pivot_result("wb1", "pt")

    submit = next(s for s in worker.submits if s["name"] == "pivot_result_read")
    assert submit["op_class"] == OperationClass.READ_ONLY


# ---------------------------------------------------------------------------
# Service-level: stale cursor
# ---------------------------------------------------------------------------
async def test_service_stale_cursor_raises(tmp_path: Path) -> None:
    service, _adapter = _build_service(
        tmp_path, page_version=7, large_data=LargeDataConfig()
    )

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await service.read_pivot_result("wb1", "pt", cursor_version=5)

    assert exc_info.value.code is ErrorCode.STALE_CURSOR
    assert exc_info.value.context["expected_version"] == 5
    assert exc_info.value.context["actual_version"] == 7
    assert exc_info.value.context["kind"] == "pivot_result"


# ---------------------------------------------------------------------------
# DTO: bounds + fingerprint
# ---------------------------------------------------------------------------
def test_dto_rejects_too_many_cells() -> None:
    cols = MAX_PIVOT_RESULT_CELLS + 1
    values = [[CellValue(kind=CellValueKind.STRING, value="v") for _ in range(cols)]]
    with pytest.raises(ValueError, match="MAX_PIVOT_RESULT_CELLS"):
        PivotResultPage(
            pivot_table="pt",
            sheet_name="Sheet1",
            source_type=PivotSourceType.DATABASE,
            table_range="$A$1",
            rows=1,
            cols=cols,
            values=values,
            page=PageInfo(offset=0, returned=1, total=1),
            fingerprint="$A$1|",
        )


def test_dto_rejects_row_count_mismatch() -> None:
    values = [[CellValue(kind=CellValueKind.STRING, value="v")]]
    with pytest.raises(ValueError, match="row count must equal rows"):
        PivotResultPage(
            pivot_table="pt",
            sheet_name="Sheet1",
            source_type=PivotSourceType.DATABASE,
            table_range="$A$1",
            rows=2,
            cols=1,
            values=values,
            page=PageInfo(offset=0, returned=1, total=2),
            fingerprint="$A$1|",
        )


def test_dto_accepts_valid_page_and_fingerprint() -> None:
    refresh = datetime(2026, 1, 2, 3, 4, 5, tzinfo=UTC)
    page = _page(version=4, rows=2, cols=2, table_range="$A$1:$B$7", refresh=refresh)

    assert page.rows == 2
    assert page.cols == 2
    assert page.page.version == 4
    # fingerprint format: "<table_range>|<refresh isoformat or empty>".
    assert page.fingerprint == "$A$1:$B$7|2026-01-02T03:04:05+00:00"


def test_dto_fingerprint_empty_when_no_refresh() -> None:
    page = _page(table_range="$A$1:$B$7", refresh=None)
    assert page.fingerprint == "$A$1:$B$7|"


# ---------------------------------------------------------------------------
# _top_row: header-row derivation (pure, hermetic)
# ---------------------------------------------------------------------------
def test_top_row_parses_first_cell_row() -> None:
    assert _top_row("$A$1:$B$7") == 1
    assert _top_row("$B$5:$B$7") == 5
    assert _top_row("$AA$10") == 10
    assert _top_row("A1:B2") == 1


def test_top_row_none_on_unparseable() -> None:
    assert _top_row(None) is None
    assert _top_row("") is None
    assert _top_row("notanaddress") is None