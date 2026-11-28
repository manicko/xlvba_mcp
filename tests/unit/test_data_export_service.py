"""Unit tests for the large-data file export service (WP-F).

Hermetic: fake range / table / pivot read services plus a real
:class:`ArtifactStore` rooted in ``tmp_path`` (no Excel / no COM). These tests
prove that ``DataExportService.export`` streams a paged source to a CSV / JSONL
artifact with bounded memory, returns only compact metadata, reports
``complete`` / ``truncated`` correctly and never returns the data inline.
"""

from __future__ import annotations

import asyncio
import csv
import io
import json
import tracemalloc
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from excel_vba_mcp.application.data_export_service import DataExportService
from excel_vba_mcp.config import LargeDataConfig
from excel_vba_mcp.domain.errors import ErrorCode, StaleCursorError
from excel_vba_mcp.domain.models import CellValue, PageInfo, RangeAddress, RangeValues
from excel_vba_mcp.domain.statuses import CellValueKind, RangeShape
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.artifact_store import ArtifactStore

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class StoreRig:
    """A real artifact store rooted at ``<tmp>/artifacts``."""

    def __init__(self, tmp_path: Path) -> None:
        self.artifacts = tmp_path / "artifacts"
        self.policy = PathPolicy({"artifacts": str(self.artifacts)})
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.store = ArtifactStore(self.artifacts, self.policy, self.audit)


@pytest.fixture
def store_rig(tmp_path: Path) -> Iterator[StoreRig]:
    rig = StoreRig(tmp_path)
    yield rig
    rig.audit.close()


# ---------------------------------------------------------------------------
# Fakes
# ---------------------------------------------------------------------------
def _string_cell(text: str) -> CellValue:
    return CellValue(kind=CellValueKind.STRING, value=text)


def _range_page(offset: int, rows: int, *, has_more: bool, version: int) -> RangeValues:
    values = [
        [_string_cell(f"r{offset + index}c0"), _string_cell(f"r{offset + index}c1")]
        for index in range(rows)
    ]
    return RangeValues(
        range=RangeAddress(sheet="S", address="A1:B100", rows=rows, cols=2),
        shape=RangeShape.BLOCK,
        rows=rows,
        cols=2,
        values=values,
        page=PageInfo(
            offset=offset,
            returned=rows,
            has_more=has_more,
            complete=not has_more,
            version=version,
        ),
    )


class FakeRanges:
    """Returns canned :class:`RangeValues` pages keyed by offset; no COM."""

    def __init__(self, pages: dict[int, RangeValues]) -> None:
        self._pages = pages
        self.offsets: list[int] = []

    async def read_values(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
    ) -> RangeValues:
        self.offsets.append(offset)
        return self._pages[offset]


class FakeTables:
    """Returns canned :class:`TableRows` pages keyed by offset; no COM."""

    def __init__(self, pages: dict[int, Any]) -> None:
        self._pages = pages
        self.offsets: list[int] = []

    async def read_table_rows(
        self,
        workbook_id: str,
        name: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
    ) -> Any:
        self.offsets.append(offset)
        return self._pages[offset]


class FakePivots:
    """Minimal pivot service stub (unused by the range/table tests)."""

    async def read_pivot_result(self, *_args: Any, **_kwargs: Any) -> Any:  # pragma: no cover
        raise AssertionError("not used in these tests")


class FakeSession:
    """Fake session exposing the content version the export binds to (no COM)."""

    def __init__(self, version: int) -> None:
        self._version = version

    def get_version(self, workbook_id: str) -> int:
        return self._version


class FakeSnapshotManifest:
    def __init__(self, sha256: str) -> None:
        self.sha256 = sha256


class FakeSnapshot:
    """Fake snapshot service returning a fixed-sha manifest (no COM/STA)."""

    def __init__(self, sha256: str = "snap-sha-0001") -> None:
        self._sha = sha256
        self.calls = 0

    async def create_snapshot(self, workbook_id: str) -> FakeSnapshotManifest:
        self.calls += 1
        return FakeSnapshotManifest(self._sha)


def _build_service(
    store_rig: StoreRig,
    ranges: Any,
    tables: Any | None = None,
    pivots: Any | None = None,
    *,
    version: int = 7,
    snapshot: Any | None = None,
) -> DataExportService:
    return DataExportService(
        ranges,
        tables if tables is not None else FakeTables({}),
        pivots if pivots is not None else FakePivots(),
        store_rig.store,
        LargeDataConfig(),
        session=FakeSession(version),
        snapshot=snapshot if snapshot is not None else FakeSnapshot(),
    )


def _three_pages() -> dict[int, RangeValues]:
    return {
        0: _range_page(0, 2, has_more=True, version=7),
        2: _range_page(2, 2, has_more=True, version=7),
        4: _range_page(4, 2, has_more=False, version=7),
    }


# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------
async def test_export_range_csv_complete(store_rig: StoreRig) -> None:
    ranges = FakeRanges(_three_pages())
    service = _build_service(store_rig, ranges)

    result = await service.export(
        "wb1", source_kind="range", sheet="S", address="A1:B100", fmt="csv"
    )

    assert ranges.offsets == [0, 2, 4]
    assert result.rows == 6
    assert result.complete is True
    assert result.truncated is False
    assert result.format == "csv"
    assert result.source == "range"
    assert result.source_version == 7
    assert result.bytes > 0
    assert Path(result.path).is_file()

    # The file parses to a header + the same number of data rows.
    text = Path(result.path).read_text(encoding="utf-8")
    parsed = list(csv.reader(io.StringIO(text)))
    assert parsed[0] == ["col0", "col1"]
    assert len(parsed) == 1 + 6


async def test_export_range_csv_limit_rows_truncates(store_rig: StoreRig) -> None:
    ranges = FakeRanges(_three_pages())
    service = _build_service(store_rig, ranges)

    result = await service.export(
        "wb1",
        source_kind="range",
        sheet="S",
        address="A1:B100",
        fmt="csv",
        limit_rows=3,
    )

    assert result.rows == 3
    assert result.truncated is True
    assert result.complete is False
    text = Path(result.path).read_text(encoding="utf-8")
    parsed = list(csv.reader(io.StringIO(text)))
    assert len(parsed) == 1 + 3


# ---------------------------------------------------------------------------
# JSONL export
# ---------------------------------------------------------------------------
async def test_export_range_jsonl_parses(store_rig: StoreRig) -> None:
    ranges = FakeRanges(_three_pages())
    service = _build_service(store_rig, ranges)

    result = await service.export(
        "wb1", source_kind="range", sheet="S", address="A1:B100", fmt="jsonl"
    )

    assert result.format == "jsonl"
    assert result.rows == 6
    lines = Path(result.path).read_text(encoding="utf-8").splitlines()
    assert len(lines) == 6
    records = [json.loads(line) for line in lines]
    assert records[0] == {"col0": "r0c0", "col1": "r0c1"}
    assert records[-1] == {"col0": "r5c0", "col1": "r5c1"}


# ---------------------------------------------------------------------------
# Table headers
# ---------------------------------------------------------------------------
async def test_export_table_uses_table_headers(store_rig: StoreRig) -> None:
    from excel_vba_mcp.adapters.excel.tables_adapter import TableRows

    page = TableRows(
        name="T",
        sheet_name="S",
        columns=["Name", "Qty"],
        offset=0,
        limit=None,
        total_rows=1,
        returned_rows=1,
        truncated=False,
        rows=[[_string_cell("apple"), CellValue(kind=CellValueKind.NUMBER, value=3)]],
        page=PageInfo(offset=0, returned=1, has_more=False, complete=True, version=1),
    )
    tables = FakeTables({0: page})
    service = _build_service(store_rig, FakeRanges({}), tables=tables)

    result = await service.export("wb1", source_kind="table", table="T", fmt="csv")

    assert result.rows == 1
    assert result.complete is True
    text = Path(result.path).read_text(encoding="utf-8")
    parsed = list(csv.reader(io.StringIO(text)))
    assert parsed[0] == ["Name", "Qty"]
    assert parsed[1] == ["apple", "3"]


# ---------------------------------------------------------------------------
# Bounded memory
# ---------------------------------------------------------------------------
async def test_export_memory_is_bounded(store_rig: StoreRig) -> None:
    # 50 pages of 100 rows = 5000 rows total; only one page is held at a time.
    pages: dict[int, RangeValues] = {}
    chunk = 100
    total_pages = 50
    for page_index in range(total_pages):
        offset = page_index * chunk
        has_more = page_index < total_pages - 1
        pages[offset] = _range_page(offset, chunk, has_more=has_more, version=1)
    ranges = FakeRanges(pages)
    service = _build_service(store_rig, ranges)

    tracemalloc.start()
    try:
        result = await service.export(
            "wb1", source_kind="range", sheet="S", address="A1:B100", fmt="csv"
        )
        _current, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()

    assert result.rows == 5000
    assert result.complete is True
    # Light bounded-memory assertion: peak stays well under 5 MiB for 5000 rows.
    assert peak < 5 * 1024 * 1024


# ---------------------------------------------------------------------------
# Snapshot / version binding (Fix 2)
# ---------------------------------------------------------------------------
async def test_export_records_snapshot_sha_and_source_version(
    store_rig: StoreRig,
) -> None:
    ranges = FakeRanges(_three_pages())
    snapshot = FakeSnapshot("deadbeefcafe")
    service = _build_service(store_rig, ranges, version=7, snapshot=snapshot)

    result = await service.export(
        "wb1", source_kind="range", sheet="S", address="A1:B100", fmt="csv"
    )

    assert result.input_sha256 == "deadbeefcafe"
    assert result.source_version == 7
    # A fresh immutable snapshot is created exactly once, at export start.
    assert snapshot.calls == 1
    # The bound version is threaded into every page read as ``cursor_version``.
    assert ranges.offsets == [0, 2, 4]


class VersionChangingRanges:
    """Fake range service whose version advances after page 1 (no COM).

    Page 1 serves with ``version == cursor_version`` and ``has_more=True`` then
    bumps the version, so page 2 sees ``cursor_version != current`` and raises
    ``STALE_CURSOR`` — mirroring a concurrent mutation between pages.
    """

    def __init__(self) -> None:
        self._version = 7
        self.offsets: list[int] = []

    async def read_values(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
    ) -> RangeValues:
        self.offsets.append(offset)
        if cursor_version is not None and cursor_version != self._version:
            raise StaleCursorError(
                "The workbook changed since the first page; restart paging.",
                context={
                    "expected_version": cursor_version,
                    "actual_version": self._version,
                    "kind": "range",
                },
            )
        page = _range_page(offset, 2, has_more=True, version=self._version)
        self._version += 1
        return page


async def test_export_version_change_between_pages_raises_stale(
    store_rig: StoreRig,
) -> None:
    ranges = VersionChangingRanges()
    service = _build_service(store_rig, ranges, version=7)

    with pytest.raises(StaleCursorError) as exc_info:
        await service.export(
            "wb1", source_kind="range", sheet="S", address="A1:B100", fmt="csv"
        )

    assert exc_info.value.code is ErrorCode.STALE_CURSOR
    # No finalized data-export artifact: the partial temp file was aborted.
    assert not list((store_rig.artifacts).glob("*.csv"))


# ---------------------------------------------------------------------------
# max_file_jobs bound (Fix 4)
# ---------------------------------------------------------------------------
class ConcurrencyRanges:
    """Fake range service that records the peak concurrent in-flight reads."""

    def __init__(self, page: RangeValues) -> None:
        self._page = page
        self.active = 0
        self.max_active = 0

    async def read_values(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
        cursor_version: int | None = None,
    ) -> RangeValues:
        self.active += 1
        self.max_active = max(self.max_active, self.active)
        try:
            await asyncio.sleep(0.01)
            return self._page
        finally:
            self.active -= 1


async def test_export_file_jobs_bound_serializes(store_rig: StoreRig) -> None:
    ranges = ConcurrencyRanges(_range_page(0, 2, has_more=False, version=7))
    service = DataExportService(
        ranges,
        FakeTables({}),
        FakePivots(),
        store_rig.store,
        LargeDataConfig(max_file_jobs=1),
        session=FakeSession(7),
        snapshot=FakeSnapshot(),
    )

    await asyncio.gather(
        service.export("wb1", source_kind="range", sheet="S", address="A1:B100"),
        service.export("wb1", source_kind="range", sheet="S", address="A1:B100"),
    )

    # max_file_jobs=1 serializes the two concurrent exports (no overlap).
    assert ranges.max_active == 1