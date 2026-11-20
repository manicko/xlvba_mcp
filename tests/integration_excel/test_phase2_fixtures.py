"""Fixture smoke test for the Phase 2 Range and PivotTables workbooks (task_041).

These tests open the committed ``RangesCellsProject.xlsm`` and
``PivotTablesProject.xlsm`` fixtures through an **automation-created hidden Excel
instance** (never the user's Excel), read them back on the single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, and assert the
**manifest's key expectations**:

* representative scalar cells (string / number / bool / blank / date / Excel
  error / an ``='...'`` text cell / a real formula);
* the 2-D block, the merged area, both data-validation rules and both
  conditional-formatting rules;
* the ``tblCells`` ListObject range, headers, and the inside / overlap / outside
  ranges used by the ListObject structural guards;
* the PivotTables' cache indices, field orientations, aggregations, page field,
  and the manual (``PivotItems.Visible``) filter on ``ptB``;
* the two PivotTables (``ptA`` / ``ptB``) sharing cache 1, while ``ptC`` uses the
  ListObject-backed cache 2;
* every cache is local ``xlDatabase`` (no external / OLAP / Data Model source).

It also asserts **BG03/BG04/BG05**: the automation Excel stays hidden and
non-activating, the owned PID never holds foreground, no owned window is
visible, the user's session is untouched, and the owned PID is reaped after
shutdown (no orphan). The committed fixtures are opened **read-only** and are
byte-identical afterwards.

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_phase2_fixtures.py -v -s
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.domain.statuses import OperationClass, WorkbookSessionState

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_FIXTURES_DIR = _REPO_ROOT / "tests" / "fixtures"
_RANGES_FIXTURE = _FIXTURES_DIR / "RangesCellsProject.xlsm"
_PIVOTS_FIXTURE = _FIXTURES_DIR / "PivotTablesProject.xlsm"
_MANIFEST = _FIXTURES_DIR / "phase2_manifest.json"

_RANGES_FIXTURE_NAME = "RangesCellsProject.xlsm"
_PIVOTS_FIXTURE_NAME = "PivotTablesProject.xlsm"

#: Excel ``XlPivotFieldOrientation`` / ``XlConsolidationFunction`` values used by
#: the assertions (mirrors the fixture generator; this test never imports Excel).
XL_FIELD_ROW = 1
XL_FIELD_COLUMN = 2
XL_FIELD_PAGE = 3
XL_FUNC_SUM = -4157
XL_FUNC_AVERAGE = -4106
XL_FUNC_COUNT = -4112
XL_SOURCE_DATABASE = 1

_TEST_TIMEOUT_S = 120.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


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
# BG03/04/05 helpers (read-only; never activate/focus/click)
# ---------------------------------------------------------------------------
def _pid_alive(pid: int) -> bool:
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001 - best effort
        return True
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:  # pragma: no cover - depends on host
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        title = win32gui.GetWindowText(hwnd)
        return {"hwnd": hwnd, "pid": pid, "title": title}
    except Exception:  # noqa: BLE001 - best effort
        return None


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    if win32gui is None or win32process is None:
        return []
    windows: list[dict[str, Any]] = []

    def _cb(hwnd: int, _extra: Any) -> bool:
        try:  # pragma: no cover - depends on host
            _tid, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            if win_pid == pid:
                windows.append(
                    {
                        "hwnd": hwnd,
                        "title": win32gui.GetWindowText(hwnd),
                        "visible": bool(win32gui.IsWindowVisible(hwnd)),
                    }
                )
        except Exception:  # noqa: BLE001 - best effort
            pass
        return True

    with contextlib.suppress(Exception):  # pragma: no cover - depends on host
        win32gui.EnumWindows(_cb, None)
    return windows


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at "
        f"{label} (title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at "
        f"{label}: {[w['title'] for w in visible]}"
    )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def manifest() -> dict[str, Any]:
    """Load the committed phase-2 manifest (skip if absent)."""

    if not _MANIFEST.is_file():
        pytest.skip(f"manifest not available: {_MANIFEST}")
    return json.loads(_MANIFEST.read_text(encoding="utf-8"))


class Rig:
    """Minimal COM worker + Excel used to read the committed fixtures."""

    def __init__(self) -> None:
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)

    @property
    def app(self) -> ExcelAppAdapter:
        return self.adapter

    async def start(self) -> None:
        await self.worker.start()

    async def close(self) -> None:
        await self.worker.shutdown()


@pytest.fixture
async def rig() -> Any:
    r = Rig()
    try:
        await r.start()
        yield r
    finally:
        await r.close()


def _open_read_only(rig: Rig, path: Path) -> Any:
    """Open ``path`` read-only on the owned hidden Excel instance."""

    app = rig.app.app
    assert app is not None
    return app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, Notify=False, AddToMru=False)


def _cell(sheet: Any, address: str) -> dict[str, Any]:
    """Read one cell as a JSON-safe dict (no COM proxy escapes)."""

    cell = sheet.Range(address)
    raw = cell.Value2
    return {
        "raw": None
        if raw is None
        else (raw if isinstance(raw, (str, bool, int, float)) else str(raw)),
        "number_format": str(cell.NumberFormat),
        "formula": str(cell.Formula),
        "has_formula": bool(cell.HasFormula),
        "merge_area": str(cell.MergeArea.Address),
        "merge_cells": bool(cell.MergeCells),
    }


# ---------------------------------------------------------------------------
# RangesCellsProject.xlsm — representative cells, block, merge, rules, table
# ---------------------------------------------------------------------------
@requires_excel
async def test_ranges_cells_representative_cells(rig: Rig, manifest: dict[str, Any]) -> None:
    expected = manifest["fixtures"][_RANGES_FIXTURE_NAME]
    path = _fixture_or_skip(_RANGES_FIXTURE)
    before = _sha256(path)

    def read() -> dict[str, Any]:
        wb = _open_read_only(rig, path)
        try:
            sheet = wb.Worksheets("Cells")
            return {address: _cell(sheet, address) for address in expected["cells"]}
        finally:
            wb.Close(SaveChanges=False)

    read.__name__ = "read_ranges_cells"
    actual = await rig.worker.submit(
        read, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.READ_ONLY
    )

    # A representative slice covering every value kind the fixture encodes.
    for address in (
        "B2",
        "B3",
        "B4",
        "B5",
        "B6",
        "B7",
        "B8",
        "B9",
        "B10",
        "B11",
        "B12",
        "B13",
        "B14",
        "B15",
    ):
        expected_cell = expected["cells"][address]
        observed = actual[address]
        assert observed["number_format"] == expected_cell["number_format"], (
            f"{address}: number format {observed['number_format']!r} != "
            f"{expected_cell['number_format']!r}"
        )
        assert observed["has_formula"] == expected_cell["has_formula"], address
        # ``Range.Formula`` echoes the value for a non-formula cell, so it is only
        # meaningful (as the formula string) when ``HasFormula`` is True.
        if expected_cell["has_formula"]:
            assert observed["formula"] == expected_cell["formula"], address
        else:
            assert expected_cell["formula"] is None, address

    # Specific typed checks (the manifest records the exact raw value).
    assert actual["B2"]["raw"] == "hello world"  # string
    assert actual["B3"]["raw"] == 42  # number
    assert actual["B5"]["raw"] is True  # boolean True
    assert actual["B6"]["raw"] is False  # boolean False
    assert actual["B7"]["raw"] is None  # blank
    # '='-prefixed *text* cell: stored as text, not a formula.
    assert actual["B8"]["raw"] == "=1+1"
    assert actual["B8"]["has_formula"] is False
    # Real formulas.
    assert actual["B9"]["formula"] == "=1+2" and actual["B9"]["has_formula"] is True
    assert actual["B10"]["formula"] == "=1/0" and actual["B10"]["has_formula"] is True
    assert actual["B11"]["formula"] == "=NA()" and actual["B11"]["has_formula"] is True
    # Dates keep their date number format; the date-shaped serial stays a number.
    assert actual["B13"]["number_format"] == "yyyy-mm-dd"
    assert actual["B14"]["number_format"] == "yyyy-mm-dd hh:mm"
    assert actual["B15"]["number_format"] == "General" and actual["B15"]["has_formula"] is False

    # BG03/BG05: the committed fixture is never modified.
    assert _sha256(path) == before


@requires_excel
async def test_ranges_cells_block_merge_validation_cf_and_table(
    rig: Rig, manifest: dict[str, Any]
) -> None:
    expected = manifest["fixtures"][_RANGES_FIXTURE_NAME]
    path = _fixture_or_skip(_RANGES_FIXTURE)
    before = _sha256(path)

    def read() -> dict[str, Any]:
        wb = _open_read_only(rig, path)
        try:
            sheet = wb.Worksheets("Cells")

            # 2-D block.
            top_left = sheet.Range("F1").Cells(1, 1)
            bottom_right = sheet.Range("F1").Cells(3, 3)
            block = sheet.Range(top_left, bottom_right).Value2

            # Merged area.
            merge_area = str(sheet.Range("J1").MergeArea.Address)
            merge_value = sheet.Range("J1").Value2
            merge_non_anchor = sheet.Range("K1").Value2

            # Data validation.
            validation = []
            for address in ("M4", "N4"):
                v = sheet.Range(address).Validation
                validation.append(
                    {
                        "address": address,
                        "type": int(v.Type),
                        "operator": int(v.Operator),
                        "formula1": str(v.Formula1),
                        "formula2": str(v.Formula2),
                        "ignore_blank": bool(v.IgnoreBlank),
                        "in_cell_dropdown": bool(v.InCellDropdown),
                    }
                )

            # Conditional formatting.
            cf = []
            for address in ("Q1", "R1"):
                conditions = sheet.Range(address).FormatConditions
                rules = []
                for index in range(1, int(conditions.Count) + 1):
                    rule = conditions.Item(index)
                    try:
                        operator = int(rule.Operator)
                    except Exception:  # noqa: BLE001 - xlExpression ignores Operator
                        operator = None
                    rules.append(
                        {
                            "type": int(rule.Type),
                            "operator": operator,
                            "formula1": str(rule.Formula1),
                            "priority": int(rule.Priority),
                            "stop_if_true": bool(rule.StopIfTrue),
                            "applies_to": str(rule.AppliesTo.Address),
                        }
                    )
                cf.append({"address": address, "rules": rules})

            # ListObject.
            table = sheet.ListObjects("tblCells")
            headers = [
                str(table.HeaderRowRange.Cells(1, col).Value2)
                for col in range(1, int(table.ListColumns.Count) + 1)
            ]

            return {
                "block": [[float(cell) for cell in row] for row in block],
                "merge": {
                    "area": merge_area,
                    "value": merge_value,
                    "non_anchor_value": merge_non_anchor,
                },
                "validation": validation,
                "cf": cf,
                "table": {
                    "name": str(table.Name),
                    "range": str(table.Range.Address),
                    "headers": headers,
                },
            }
        finally:
            wb.Close(SaveChanges=False)

    read.__name__ = "read_ranges_structures"
    observed = await rig.worker.submit(
        read, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.READ_ONLY
    )

    # 2-D block matches the manifest.
    assert observed["block"] == expected["block"]["values"]

    # Merged area: anchor holds the value, a non-anchor cell is empty.
    assert observed["merge"]["area"].replace("$", "") == "J1:L2"
    assert observed["merge"]["value"] == "merged title"
    assert observed["merge"]["non_anchor_value"] is None

    # Data validation: list + whole-number between, with the expected operators.
    expected_validation = {
        rule["range"].split(":")[0]: rule for rule in expected["data_validation"]
    }
    by_address = {rule["address"]: rule for rule in observed["validation"]}
    assert by_address["M4"]["type"] == expected_validation["M4"]["type"]  # list
    assert by_address["M4"]["formula1"] == expected_validation["M4"]["formula1"]
    assert by_address["M4"]["in_cell_dropdown"] is True
    assert by_address["N4"]["type"] == expected_validation["N4"]["type"]  # whole number
    assert by_address["N4"]["operator"] == expected_validation["N4"]["operator"]  # between
    assert by_address["N4"]["formula1"] == "1" and by_address["N4"]["formula2"] == "10"

    # Conditional formatting: one cell-value rule and one expression rule.
    expected_cf = {rule["range"].split(":")[0]: rule for rule in expected["conditional_formatting"]}
    cf_by_address = {rule["address"]: rule["rules"][0] for rule in observed["cf"]}
    assert cf_by_address["Q1"]["type"] == expected_cf["Q1"]["type"]  # cell_value
    assert cf_by_address["Q1"]["formula1"] == expected_cf["Q1"]["formula1"]
    assert cf_by_address["Q1"]["operator"] == expected_cf["Q1"]["operator"]
    assert cf_by_address["R1"]["type"] == expected_cf["R1"]["type"]  # expression
    assert cf_by_address["R1"]["formula1"] == expected_cf["R1"]["formula1"]
    # xlExpression rules ignore Operator, so both sides are None.
    assert cf_by_address["R1"]["operator"] is None

    # ListObject range, name and headers.
    design_table = expected["design"]["list_object"]
    assert observed["table"]["name"] == design_table["name"]
    assert observed["table"]["range"].replace("$", "") == design_table["range"]
    assert observed["table"]["headers"] == design_table["headers"]

    # The manifest documents inside / overlap / outside ranges for the guards.
    assert design_table["inside_ranges"]
    assert design_table["overlap_ranges"]
    assert design_table["outside_ranges"]

    assert _sha256(path) == before


# ---------------------------------------------------------------------------
# PivotTablesProject.xlsm — layout, cache sharing, aggregation, manual filter
# ---------------------------------------------------------------------------
@requires_excel
async def test_pivots_layout_cache_sharing_and_filter(rig: Rig, manifest: dict[str, Any]) -> None:
    expected = manifest["fixtures"][_PIVOTS_FIXTURE_NAME]
    path = _fixture_or_skip(_PIVOTS_FIXTURE)
    before = _sha256(path)

    def read() -> dict[str, Any]:
        wb = _open_read_only(rig, path)
        try:
            report = wb.Worksheets("PivotReport")
            pivots: dict[str, Any] = {}
            for name in ("ptA", "ptB", "ptC"):
                pivot = report.PivotTables(name)
                cache = pivot.PivotCache()
                pivot_ref = pivot

                def field_names(collection_name: str, *, _pivot: Any = pivot_ref) -> list[str]:
                    collection = getattr(_pivot, collection_name)
                    names: list[str] = []
                    for index in range(1, int(collection.Count) + 1):
                        source_name = collection.Item(index).SourceName
                        if isinstance(source_name, str):
                            names.append(str(source_name))
                    return names

                data_fields = []
                for index in range(1, int(pivot.DataFields.Count) + 1):
                    field = pivot.DataFields.Item(index)
                    data_fields.append(
                        {
                            "source_name": str(field.SourceName),
                            "caption": str(field.Name),
                            "function": int(field.Function),
                        }
                    )

                page_current = None
                if int(pivot.PageFields.Count) > 0:
                    page_source = field_names("PageFields")[0]
                    page_current = str(pivot.PivotFields(page_source).CurrentPage.Value)

                filter_state = None
                if name == "ptB":
                    field = pivot.PivotFields("Region")
                    visible, hidden = [], []
                    items = field.PivotItems()
                    for index in range(1, int(items.Count) + 1):
                        item = items.Item(index)
                        if bool(item.Visible):
                            visible.append(str(item.Value))
                        else:
                            hidden.append(str(item.Value))
                    filter_state = {
                        "all_items_visible": bool(field.AllItemsVisible),
                        "visible_items": visible,
                        "hidden_items": hidden,
                    }

                pivots[name] = {
                    "cache_index": int(pivot.CacheIndex),
                    "cache_source_type": int(cache.SourceType),
                    "cache_source_data": str(cache.SourceData),
                    "row_fields": field_names("RowFields"),
                    "column_fields": field_names("ColumnFields"),
                    "page_fields": field_names("PageFields"),
                    "page_current_page": page_current,
                    "data_fields": data_fields,
                    "filter": filter_state,
                }

            cache_collection = wb.PivotCaches()
            caches = []
            for index in range(1, int(cache_collection.Count) + 1):
                cache = cache_collection.Item(index)
                caches.append(
                    {
                        "index": int(cache.Index),
                        "source_type": int(cache.SourceType),
                        "source_data": str(cache.SourceData),
                    }
                )

            return {"pivots": pivots, "caches": caches}
        finally:
            wb.Close(SaveChanges=False)

    read.__name__ = "read_pivots"
    observed = await rig.worker.submit(
        read, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.READ_ONLY
    )

    # All three PivotTables are present and each matches the manifest design.
    expected_by_name = {spec["name"]: spec for spec in expected["design"]["pivot_tables"]}
    assert set(observed["pivots"]) == set(expected_by_name)

    for name, spec in expected_by_name.items():
        actual = observed["pivots"][name]
        assert actual["cache_index"] == spec["cache_index"], name
        assert actual["row_fields"] == spec["row_fields"], name
        assert actual["column_fields"] == spec["column_fields"], name
        assert actual["page_fields"] == spec["page_fields"], name
        expected_data = [
            (f["source_name"], f["caption"], f["function"]) for f in spec["data_fields"]
        ]
        observed_data = [
            (f["source_name"], f["caption"], f["function"]) for f in actual["data_fields"]
        ]
        assert observed_data == expected_data, name

    # ptA: row + column + page + two data fields with distinct aggregations.
    pta = observed["pivots"]["ptA"]
    assert pta["row_fields"] == ["Region"]
    assert pta["column_fields"] == ["Product"]
    assert pta["page_fields"] == ["Quarter"]
    assert pta["page_current_page"] == "(All)"
    assert {f["function"] for f in pta["data_fields"]} == {XL_FUNC_SUM, XL_FUNC_AVERAGE}

    # ptB: shares cache 1 with ptA and carries a manual filter (West hidden).
    ptb = observed["pivots"]["ptB"]
    assert ptb["cache_index"] == pta["cache_index"] == 1
    assert ptb["filter"]["all_items_visible"] is False
    assert "West" in ptb["filter"]["hidden_items"]
    assert "West" not in ptb["filter"]["visible_items"]

    # ptC: ListObject-backed cache 2 with a count aggregation.
    ptc = observed["pivots"]["ptC"]
    assert ptc["cache_index"] == 2
    assert ptc["data_fields"][0]["function"] == XL_FUNC_COUNT
    assert "TableData" in ptc["cache_source_data"]

    # Shared-cache invariant from the manifest.
    shared = expected["design"]["shared_cache"]
    shared_indices = {observed["pivots"][name]["cache_index"] for name in shared["pivot_tables"]}
    assert shared_indices == {shared["cache_index"]}

    # Every cache is local xlDatabase (no external / OLAP / Data Model source).
    for cache in observed["caches"]:
        assert cache["source_type"] == XL_SOURCE_DATABASE, cache
    for name, actual in observed["pivots"].items():
        assert actual["cache_source_type"] == XL_SOURCE_DATABASE, name

    assert _sha256(path) == before


# ---------------------------------------------------------------------------
# Manifest integrity
# ---------------------------------------------------------------------------
def test_manifest_matches_committed_fixture_hashes(manifest: dict[str, Any]) -> None:
    """The manifest SHA-256 must equal the committed fixture bytes."""

    assert manifest["task"] == "task_041"
    assert manifest["target_excel_build"] == "16.0.14332"
    assert manifest["generated_at"]

    for name, block in manifest["fixtures"].items():
        path = _REPO_ROOT / block["path"]
        assert path.is_file(), f"missing fixture: {path}"
        assert _sha256(path) == block["sha256"], f"{name}: sha256 mismatch"
        assert block["size_bytes"] == path.stat().st_size
        assert not block.get("cell_mismatches"), f"{name}: cell mismatches recorded"
        assert not block.get("pivot_mismatches"), f"{name}: pivot mismatches recorded"

    # The design oracle is present and non-empty.
    ranges_design = manifest["fixtures"][_RANGES_FIXTURE_NAME]["design"]
    assert ranges_design["scalar_cells"]
    assert ranges_design["list_object"]["inside_ranges"]
    assert ranges_design["list_object"]["overlap_ranges"]

    pivots_design = manifest["fixtures"][_PIVOTS_FIXTURE_NAME]["design"]
    assert len(pivots_design["pivot_tables"]) == 3


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05) + orphan check
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_04_05_hidden_owned_excel(rig: Rig, manifest: dict[str, Any]) -> None:
    ranges_path = _fixture_or_skip(_RANGES_FIXTURE)
    pivots_path = _fixture_or_skip(_PIVOTS_FIXTURE)
    ranges_before = _sha256(ranges_path)
    pivots_before = _sha256(pivots_path)

    owned_pid = rig.worker.get_status().owned_pid
    assert owned_pid is not None

    def read_all() -> dict[str, Any]:
        result: dict[str, Any] = {}
        for path in (ranges_path, pivots_path):
            wb = _open_read_only(rig, path)
            try:
                result[str(path)] = int(wb.Worksheets.Count)
            finally:
                wb.Close(SaveChanges=False)
        return result

    read_all.__name__ = "read_all"
    await rig.worker.submit(
        read_all, timeout=_TEST_TIMEOUT_S, section="diagnostic", op_class=OperationClass.READ_ONLY
    )

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

    # BG03/BG04: owned Excel hidden + non-activating.
    assert flags["visible"] is False
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING
    _assert_owned_never_foreground(owned_pid, "phase2-fixtures")
    _assert_no_visible_owned_windows(owned_pid, "phase2-fixtures")

    # BG05: the committed fixtures are byte-identical after the reads.
    assert _sha256(ranges_path) == ranges_before
    assert _sha256(pivots_path) == pivots_before


@requires_excel
async def test_owned_excel_reaped_after_shutdown(manifest: dict[str, Any]) -> None:
    """The automation Excel PID is gone after ``shutdown`` (no orphan)."""

    _fixture_or_skip(_RANGES_FIXTURE)
    r = Rig()
    await r.start()
    try:

        def open_one() -> int:
            wb = _open_read_only(r, _RANGES_FIXTURE)
            try:
                return int(wb.Worksheets.Count)
            finally:
                wb.Close(SaveChanges=False)

        open_one.__name__ = "open_one"
        await r.worker.submit(
            open_one,
            timeout=_TEST_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
        owned_pid = r.worker.get_status().owned_pid
        assert owned_pid is not None
    finally:
        await r.close()

    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
