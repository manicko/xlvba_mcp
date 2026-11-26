"""Independent end-to-end verification of the Ranges & Cells workflow (task_047).

This module drives the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and a disposable copy of the committed ``RangesCellsProject.xlsm``
fixture (task_041). It is deliberately written as an *independent* verification
of the behaviour implemented by task_042..task_046 rather than a copy of those
modules' assertions: the targets, payloads and postconditions below are chosen
fresh so a regression in any single implementation task is caught here.

Coverage (task_047 acceptance criteria):

1. value / formula round-trips — string, number, bool, blank, date, error and a
   2-D block, plus a wrong-sized payload rejection and a postcondition match;
2. formatting — read + set of number format / font / fill / alignment, partial
   update leaves the other properties untouched, and mixed-format reporting;
3. structural — copy / insert / delete / merge / unmerge postconditions, the
   merge data-loss gate and the ListObject structural guard for every operation;
4. data validation + conditional formatting — read the fixture rules, add /
   modify / delete, unsupported-type rejection and priority readback;
5. snapshot / approval / recovery — the guard chain across every mutating tool,
   a forced hung COM call (``EXCEL_HUNG`` + ``NEEDS_RECOVERY``) and the audit
   trail of mutating operations;
6. response limits — oversized ranges rejected before a read and typed errors
   for an invalid sheet / address / multi-area target;
8. Stage-1 sanity — ``TablesService`` / ``NamesService`` still behave.

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called and every mutation runs on a
disposable working copy.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import shutil
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest
from pydantic import ValidationError

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    MAX_RANGE_CELLS,
    CellValue,
    ConditionalFormatSpec,
    FormulaWriteRequest,
    RangeFormatPatch,
    RangeWriteRequest,
    ValidationSpec,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    ConditionalFormatType,
    DataValidationType,
    NameScope,
    OperationClass,
    RangeShape,
    SessionMode,
    ValidationOperator,
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

_RANGES_PROJECT = (
    Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "RangesCellsProject.xlsm"
)

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_TEST_TIMEOUT_S = 120.0
_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


def _excel_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001 - COM unavailable (non-Windows)
        return False
    return True


requires_excel = pytest.mark.skipif(
    not _excel_available(), reason="Excel COM automation is not available on this host"
)


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
def ranges_project() -> Any:
    """Yield the committed ``RangesCellsProject.xlsm`` fixture path (skip if absent)."""

    if not _RANGES_PROJECT.is_file():
        pytest.skip(f"fixture not available: {_RANGES_PROJECT}")
    return _RANGES_PROJECT


# ---------------------------------------------------------------------------
# Helpers (read-only / test-local)
# ---------------------------------------------------------------------------
def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _pid_alive(pid: int) -> bool:
    """Return whether ``pid`` still exists (best-effort ``tasklist`` probe)."""

    import subprocess

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


def _cv(kind: CellValueKind, value: Any = None) -> CellValue:
    return CellValue(kind=kind, value=value)


def _values_payload(grid: list[list[CellValue]]) -> RangeWriteRequest:
    return RangeWriteRequest(rows=len(grid), cols=len(grid[0]), values=grid)


def _single_value(
    kind: CellValueKind = CellValueKind.NUMBER, value: Any = None
) -> RangeWriteRequest:
    return RangeWriteRequest(rows=1, cols=1, values=[[_cv(kind, value)]])


def _single_formula(formula: str) -> FormulaWriteRequest:
    return FormulaWriteRequest(rows=1, cols=1, formulas=[[formula]])


def _list_validation(formula1: str = '"Yes,No"') -> ValidationSpec:
    return ValidationSpec(
        type=DataValidationType.LIST,
        operator=ValidationOperator.BETWEEN,
        formula1=formula1,
        in_cell_dropdown=True,
        ignore_blank=True,
    )


def _cell_value_cf(formula1: str = "10") -> ConditionalFormatSpec:
    return ConditionalFormatSpec(
        type=ConditionalFormatType.CELL_VALUE,
        operator=ValidationOperator.GREATER,
        formula1=formula1,
    )


def _read_audit(path: Path) -> list[dict[str, Any]]:
    """Parse the append-only JSONL audit trail (best effort)."""

    if not path.is_file():
        return []
    records: list[dict[str, Any]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if stripped:
            records.append(json.loads(stripped))
    return records


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for the range tools."""

    def __init__(self, tmp_path: Path, project: Path) -> None:
        self.tmp_path = tmp_path
        self.project = project
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        canonical = self.path_policy.validate_path(str(project), OperationClass.MUTATING)
        self.canonical = str(canonical)
        self.adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
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
        self.ranges = RangesService(
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
        self.names = NamesService(
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
                granted_by="task_047",
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
                granted_by="task_047",
                source=ApprovalSource.CLIENT,
                grant_id=uuid4().hex,
            )
        )

    def grant_all_range_actions(self) -> None:
        for action in (
            "range_write_values",
            "range_write_formulas",
            "range_set_format",
            "range_copy",
            "range_insert",
            "range_delete",
            "range_merge",
            "range_unmerge",
            "range_set_validation",
            "range_set_conditional_formatting",
        ):
            self.grant(action)

    async def create_snapshot(self, workbook_id: str) -> None:
        await self.snapshot.create_snapshot(workbook_id)

    def working_path(self, workbook_id: str) -> str:
        return self.session._records[workbook_id].working_path

    async def close(self) -> None:
        owned_pid = getattr(self.worker, "_owned_pid", None)
        await self.worker.shutdown()
        self.audit.close()
        # Never leave a half-dead owned EXCEL.exe behind: wait for our own PID to
        # exit, but never touch a non-owned process.
        if owned_pid:
            await asyncio.to_thread(_wait_pid_gone, int(owned_pid))


@pytest.fixture
async def rig(tmp_path: Path, ranges_project: Path) -> Any:
    # Copy the committed fixture into the allowlisted temp root as the "original"
    # so every test mutates a disposable working copy and never the source.
    original = tmp_path / ranges_project.name
    shutil.copy2(ranges_project, original)
    r = Rig(tmp_path, original)
    try:
        await r.start()
        yield r
    finally:
        await r.close()


# ---------------------------------------------------------------------------
# STA probes (test setup / independent postcondition reads)
# ---------------------------------------------------------------------------
def _open_workbook(rig: Rig, working_path: str) -> Any:
    app = rig.app.app
    assert app is not None
    for wb in app.Workbooks:
        if str(wb.FullName).lower() == working_path.lower():
            return wb
    raise AssertionError(f"working copy not open: {working_path}")


async def _probe_cell(rig: Rig, workbook_id: str, sheet: str, address: str) -> dict[str, Any]:
    """Read one cell's raw value / formula / merge state through the STA worker."""

    working_path = rig.working_path(workbook_id)

    def probe() -> dict[str, Any]:
        wb = _open_workbook(rig, working_path)
        cell = wb.Worksheets(sheet).Range(address)
        return {
            "value": cell.Value2,
            "formula": cell.Formula,
            "merge_cells": bool(cell.MergeCells),
            "merge_area": str(cell.MergeArea.Address),
        }

    probe.__name__ = "task047_probe_cell"
    return await rig.worker.submit(
        probe,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.READ_ONLY,
    )


async def _protect_sheet(rig: Rig, workbook_id: str, sheet: str) -> None:
    """Protect ``sheet`` on the disposable working copy (test setup only)."""

    working_path = rig.working_path(workbook_id)

    def protect() -> None:
        wb = _open_workbook(rig, working_path)
        wb.Worksheets(sheet).Protect()

    protect.__name__ = "task047_protect_sheet"
    await rig.worker.submit(
        protect,
        timeout=_TEST_TIMEOUT_S,
        section="diagnostic",
        op_class=OperationClass.MUTATING,
    )


async def _open_for_mutation(rig: Rig) -> str:
    """Open the disposable copy writable, snapshot it and grant every range action."""

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant_all_range_actions()
    return wid


def _mutation_calls(rig: Rig, wid: str) -> list[tuple[str, Any]]:
    """Return ``(tool, awaitable-factory)`` for every mutating range tool.

    Each entry is a zero-argument callable returning a coroutine so the guard
    chain can be exercised uniformly across the whole surface.
    """

    return [
        (
            "range_write_values",
            lambda: rig.ranges.write_values(wid, "Cells", "A40", _single_value(value=1.0)),
        ),
        (
            "range_write_formulas",
            lambda: rig.ranges.write_formulas(wid, "Cells", "A40", _single_formula("=1+1")),
        ),
        (
            "range_set_format",
            lambda: rig.ranges.set_format(wid, "Cells", "A40", RangeFormatPatch(font_bold=True)),
        ),
        (
            "range_copy",
            lambda: rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2"),
        ),
        ("range_insert", lambda: rig.ranges.insert_range(wid, "StructOps", "B2", "right")),
        ("range_delete", lambda: rig.ranges.delete_range(wid, "StructOps", "B2", "up")),
        (
            "range_merge",
            lambda: rig.ranges.merge_range(
                wid, "StructOps", "J2:L2", across=False, allow_data_loss=True
            ),
        ),
        ("range_unmerge", lambda: rig.ranges.unmerge_range(wid, "StructOps", "J5:L5")),
        (
            "range_set_validation",
            lambda: rig.ranges.set_validation(wid, "Cells", "W4:W6", _list_validation()),
        ),
        (
            "range_set_conditional_formatting",
            lambda: rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf()),
        ),
    ]


# ===========================================================================
# 1. Value & formula round-trips
# ===========================================================================
@requires_excel
async def test_value_round_trip_string_number_bool_blank(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_mutation(rig)

    payload = _values_payload(
        [
            [
                _cv(CellValueKind.STRING, "task047"),
                _cv(CellValueKind.NUMBER, 7.5),
                _cv(CellValueKind.BOOLEAN, True),
                _cv(CellValueKind.EMPTY),
            ]
        ]
    )
    result = await rig.ranges.write_values(wid, "Cells", "A20:D20", payload)

    assert result.verified is True
    assert result.range_after is not None
    assert result.range_after.address == "$A$20:$D$20"
    assert result.range_after.shape is RangeShape.ROW

    post = await rig.ranges.read_values(wid, "Cells", "A20:D20")
    assert [cell.kind for cell in post.values[0]] == [
        CellValueKind.STRING,
        CellValueKind.NUMBER,
        CellValueKind.BOOLEAN,
        CellValueKind.EMPTY,
    ]
    assert post.values[0][0].value == "task047"
    assert float(post.values[0][1].value) == 7.5
    assert post.values[0][2].value is True
    assert post.values[0][3].value is None

    assert _sha256(rig.project) == before


@requires_excel
async def test_value_round_trip_2d_block(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    payload = _values_payload(
        [
            [_cv(CellValueKind.NUMBER, 11.0), _cv(CellValueKind.NUMBER, 12.0)],
            [_cv(CellValueKind.NUMBER, 13.0), _cv(CellValueKind.NUMBER, 14.0)],
        ]
    )
    result = await rig.ranges.write_values(wid, "Cells", "A30:B31", payload)
    assert result.verified is True

    post = await rig.ranges.read_values(wid, "Cells", "A30:B31")
    assert post.shape is RangeShape.BLOCK
    matrix = [[float(cell.value) for cell in row] for row in post.values]
    assert matrix == [[11.0, 12.0], [13.0, 14.0]]


@requires_excel
async def test_value_round_trip_date_and_error_reads(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    # A date only reads back as a DATE when the cell carries a date number format.
    await rig.ranges.set_format(wid, "Cells", "A34", RangeFormatPatch(number_format="yyyy-mm-dd"))

    payload = _values_payload([[_cv(CellValueKind.DATE, "2026-03-04T00:00:00")]])
    result = await rig.ranges.write_values(wid, "Cells", "A34", payload)
    assert result.verified is True

    post = await rig.ranges.read_values(wid, "Cells", "A34")
    assert post.values[0][0].kind is CellValueKind.DATE
    assert post.values[0][0].value == "2026-03-04T00:00:00"

    # An Excel error is produced by the formula channel and typed on read.
    formula = FormulaWriteRequest(rows=1, cols=1, formulas=[["=1/0"]])
    assert (await rig.ranges.write_formulas(wid, "Cells", "C34", formula)).verified is True
    errors = await rig.ranges.read_values(wid, "Cells", "C34")
    assert errors.values[0][0].kind is CellValueKind.ERROR
    assert errors.values[0][0].value == "#DIV/0!"

    # The fixture's own pre-existing error cells also read as typed errors.
    fixture_errors = await rig.ranges.read_values(wid, "Cells", "B10:B11")
    assert [row[0].kind for row in fixture_errors.values] == [
        CellValueKind.ERROR,
        CellValueKind.ERROR,
    ]
    assert [row[0].value for row in fixture_errors.values] == ["#DIV/0!", "#N/A"]


@requires_excel
async def test_formula_round_trip_and_evaluation(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    constants = _values_payload(
        [
            [_cv(CellValueKind.NUMBER, 3.0), _cv(CellValueKind.NUMBER, 4.0)],
            [_cv(CellValueKind.NUMBER, 5.0), _cv(CellValueKind.NUMBER, 6.0)],
        ]
    )
    assert (await rig.ranges.write_values(wid, "Cells", "F20:G21", constants)).verified is True

    formulas = FormulaWriteRequest(rows=1, cols=2, formulas=[["=SUM(F20:G21)", "=F20*G20"]])
    result = await rig.ranges.write_formulas(wid, "Cells", "F22:G22", formulas)
    assert result.verified is True

    read_back = await rig.ranges.read_formulas(wid, "Cells", "F22:G22")
    assert [cell.formula2 for cell in read_back.formulas[0]] == ["=SUM(F20:G21)", "=F20*G20"]

    values = await rig.ranges.read_values(wid, "Cells", "F22:G22")
    assert float(values.values[0][0].value) == 18.0
    assert float(values.values[0][1].value) == 12.0


@requires_excel
async def test_wrong_size_payload_rejected_before_mutation(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    original = _values_payload(
        [
            [_cv(CellValueKind.NUMBER, 1.0), _cv(CellValueKind.NUMBER, 2.0)],
            [_cv(CellValueKind.NUMBER, 3.0), _cv(CellValueKind.NUMBER, 4.0)],
        ]
    )
    assert (await rig.ranges.write_values(wid, "Cells", "A40:B41", original)).verified is True
    before = await rig.ranges.read_values(wid, "Cells", "A40:B41")

    # A 1x1 payload against a 2x2 target must be rejected before any mutation.
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_values(wid, "Cells", "A40:B41", _single_value(value=99.0))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.read_values(wid, "Cells", "A40:B41")
    before_matrix = [[float(cell.value) for cell in row] for row in before.values]
    after_matrix = [[float(cell.value) for cell in row] for row in after.values]
    assert before_matrix == after_matrix


@requires_excel
async def test_write_postcondition_carries_re_read_values(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    payload = _values_payload([[_cv(CellValueKind.STRING, "post"), _cv(CellValueKind.NUMBER, 2.5)]])
    result = await rig.ranges.write_values(wid, "Cells", "H20:I20", payload)

    assert result.verified is True
    assert result.postcondition_values is not None
    assert result.range_before is not None
    assert result.postcondition_values.rows == 1
    assert result.postcondition_values.values[0][0].value == "post"

    fresh = await rig.ranges.read_values(wid, "Cells", "H20:I20")
    assert [cell.value for cell in fresh.values[0]] == [
        cell.value for cell in result.postcondition_values.values[0]
    ]


@requires_excel
async def test_formula_write_rejects_non_formula_entry(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.write_formulas(wid, "Cells", "A40", _single_formula("F20*2"))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


# ===========================================================================
# 2. Formatting
# ===========================================================================
@requires_excel
async def test_get_format_number_font_alignment_from_fixture(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    fmt = await rig.ranges.get_format(wid, "Cells", "B4")

    assert fmt.number_format == "0.00"
    assert fmt.mixed_properties == []
    assert isinstance(fmt.font.name, str) and fmt.font.name
    assert isinstance(fmt.font.size, float) and fmt.font.size > 0
    assert fmt.font.bold is False
    assert fmt.alignment.horizontal in {"xlGeneral", "xlLeft", "xlCenter", "xlRight"}
    assert fmt.alignment.vertical in {"xlTop", "xlBottom", "xlCenter"}
    # Unsupported aspects are reported, never silently dropped.
    assert "borders" in fmt.unsupported

    assert _sha256(rig.project) == before


@requires_excel
async def test_set_format_all_supported_properties_round_trip(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    patch = RangeFormatPatch(
        font_bold=True,
        font_italic=True,
        font_size=14.0,
        font_color=255,
        fill_color=65280,
        number_format="0.000",
        alignment_horizontal="xlCenter",
        alignment_vertical="xlTop",
        alignment_wrap_text=True,
    )
    result = await rig.ranges.set_format(wid, "Cells", "C20", patch)
    assert result.verified is True
    assert set(result.applied) == set(patch.applied_fields())

    fmt = await rig.ranges.get_format(wid, "Cells", "C20")
    assert fmt.font.bold is True
    assert fmt.font.italic is True
    assert fmt.font.size == 14.0
    assert fmt.font.color == 255
    assert fmt.fill.color == 65280
    assert fmt.number_format == "0.000"
    assert fmt.alignment.horizontal == "xlCenter"
    assert fmt.alignment.vertical == "xlTop"
    assert fmt.alignment.wrap_text is True


@requires_excel
async def test_set_format_partial_update_leaves_other_properties(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    await rig.ranges.set_format(
        wid, "Cells", "D20", RangeFormatPatch(font_bold=True, font_italic=True, font_size=13.0)
    )
    before = await rig.ranges.get_format(wid, "Cells", "D20")

    result = await rig.ranges.set_format(
        wid, "Cells", "D20", RangeFormatPatch(number_format="0.00")
    )
    assert result.verified is True
    assert result.applied == ["number_format"]

    after = await rig.ranges.get_format(wid, "Cells", "D20")
    assert after.number_format == "0.00"
    assert after.font.bold is True
    assert after.font.italic is True
    assert after.font.size == before.font.size


@requires_excel
async def test_get_format_mixed_range_is_not_flattened(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    await rig.ranges.set_format(wid, "Cells", "E20", RangeFormatPatch(number_format="0.00"))
    await rig.ranges.set_format(wid, "Cells", "E20", RangeFormatPatch(font_bold=True))
    await rig.ranges.set_format(wid, "Cells", "E21", RangeFormatPatch(number_format="0.000"))

    mixed = await rig.ranges.get_format(wid, "Cells", "E20:E21")
    assert mixed.number_format is None
    assert mixed.font.bold is None
    assert "number_format" in mixed.mixed_properties
    assert "font.bold" in mixed.mixed_properties


@requires_excel
async def test_set_format_invalid_enum_rejected_without_mutation(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    before = await rig.ranges.get_format(wid, "Cells", "F20")
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_format(
            wid, "Cells", "F20", RangeFormatPatch(alignment_horizontal="xlBogus")
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_format(wid, "Cells", "F20")
    assert after.alignment.horizontal == before.alignment.horizontal


# ===========================================================================
# 3. Structural operations
# ===========================================================================
@requires_excel
async def test_copy_same_sheet_postcondition(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")

    assert result.verified is True
    assert result.range_after is not None
    assert result.range_after.address == "$D$2:$E$4"
    assert result.table_metadata_unchanged is True

    post = await rig.ranges.read_values(wid, "StructOps", "D2:E4")
    matrix = [[float(cell.value) for cell in row] for row in post.values]
    assert matrix == [[1.0, 2.0], [3.0, 4.0], [5.0, 6.0]]


@requires_excel
async def test_copy_cross_sheet_postcondition(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.copy_range(wid, "StructOps", "G2:H4", "Cells", "Z1")

    assert result.verified is True
    assert result.source is not None and result.source.sheet == "StructOps"
    assert result.range_after is not None and result.range_after.sheet == "Cells"

    post = await rig.ranges.read_values(wid, "Cells", "Z1:AA3")
    matrix = [[float(cell.value) for cell in row] for row in post.values]
    assert matrix == [[10.0, 20.0], [30.0, 40.0], [50.0, 60.0]]


@requires_excel
async def test_insert_shift_right_and_down_postconditions(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    right = await rig.ranges.insert_range(wid, "StructOps", "B2", "right")
    assert right.verified is True
    assert right.shift == "right"
    assert (await _probe_cell(rig, wid, "StructOps", "B2"))["value"] is None
    assert float((await _probe_cell(rig, wid, "StructOps", "C2"))["value"]) == 2.0

    down = await rig.ranges.insert_range(wid, "StructOps", "G3", "down")
    assert down.verified is True
    assert down.shift == "down"
    assert (await _probe_cell(rig, wid, "StructOps", "G3"))["value"] is None
    assert float((await _probe_cell(rig, wid, "StructOps", "G4"))["value"]) == 30.0


@requires_excel
async def test_delete_shift_up_and_left_postconditions(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    up = await rig.ranges.delete_range(wid, "StructOps", "B2", "up")
    assert up.verified is True
    assert up.shift == "up"
    assert float((await _probe_cell(rig, wid, "StructOps", "B2"))["value"]) == 4.0
    assert float((await _probe_cell(rig, wid, "StructOps", "B3"))["value"]) == 6.0

    left = await rig.ranges.delete_range(wid, "StructOps", "G2", "left")
    assert left.verified is True
    assert left.shift == "left"
    assert float((await _probe_cell(rig, wid, "StructOps", "G2"))["value"]) == 20.0


@requires_excel
async def test_merge_data_loss_gate_then_merge(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.merge_range(wid, "StructOps", "J2:L2")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert exc_info.value.context["cells_with_data"] == ["K2", "L2"]
    assert (await _probe_cell(rig, wid, "StructOps", "K2"))["merge_cells"] is False

    result = await rig.ranges.merge_range(wid, "StructOps", "J2:L2", allow_data_loss=True)
    assert result.verified is True
    assert result.affected_addresses == ["K2", "L2"]
    assert result.notes and any("allow_data_loss" in note for note in result.notes)

    anchor = await _probe_cell(rig, wid, "StructOps", "J2")
    assert anchor["value"] == "anchor"
    assert anchor["merge_area"] == "$J$2:$L$2"
    assert (await _probe_cell(rig, wid, "StructOps", "K2"))["value"] is None


@requires_excel
async def test_merge_across_keeps_row_anchors(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.merge_range(
        wid, "StructOps", "J2:L3", across=True, allow_data_loss=True
    )

    assert result.verified is True
    assert result.preconditions["across"] is True
    assert (await _probe_cell(rig, wid, "StructOps", "J2"))["merge_area"] == "$J$2:$L$2"
    assert (await _probe_cell(rig, wid, "StructOps", "J3"))["merge_area"] == "$J$3:$L$3"


@requires_excel
async def test_unmerge_restores_merged_area(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.unmerge_range(wid, "StructOps", "J5:L5")

    assert result.verified is True
    assert result.affected_addresses == ["$J$5:$L$5"]

    anchor = await _probe_cell(rig, wid, "StructOps", "J5")
    assert anchor["value"] == "unmerge me"
    assert anchor["merge_cells"] is False
    assert anchor["merge_area"] == "$J$5"

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.unmerge_range(wid, "StructOps", "A2:B4")
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


@requires_excel
async def test_get_range_reports_merged_geometry(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    merged = await rig.ranges.merge_range(wid, "StructOps", "N2:P2", allow_data_loss=True)
    assert merged.verified is True
    assert merged.range_after.is_merged is True
    assert merged.range_after.merged_area == "$N$2:$P$2"

    # The descriptor of the merged block reports the anchor area.
    anchor = await rig.ranges.get_range(wid, "StructOps", "N2:P2")
    assert anchor.is_merged is True
    assert anchor.merged_area == "$N$2:$P$2"

    # An interior (non-anchor) cell resolves to the same merged area.
    interior = await rig.ranges.get_range(wid, "StructOps", "O2")
    assert interior.is_merged is True
    assert interior.merged_area == "$N$2:$P$2"

    # range_unmerge precondition agrees with the descriptor's merged_area.
    unmerged = await rig.ranges.unmerge_range(wid, "StructOps", "N2")
    assert unmerged.verified is True
    assert unmerged.preconditions["merged_areas"] == ["$N$2:$P$2"]

    after = await rig.ranges.get_range(wid, "StructOps", "N2:P2")
    assert after.is_merged is False
    assert after.merged_area is None


@requires_excel
async def test_get_range_fragmented_merge_is_not_merged(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    # Two separate single-row merges; the union A1:B2 is NOT one merged block.
    assert (await rig.ranges.merge_range(wid, "StructOps", "N5:O5", allow_data_loss=True)).verified
    assert (await rig.ranges.merge_range(wid, "StructOps", "N6:O6", allow_data_loss=True)).verified

    combined = await rig.ranges.get_range(wid, "StructOps", "N5:O6")
    assert combined.is_merged is False
    assert combined.merged_area is None

    row = await rig.ranges.get_range(wid, "StructOps", "N5:O5")
    assert row.is_merged is True
    assert row.merged_area == "$N$5:$O$5"


@requires_excel
async def test_range_structure_descriptor_tracks_merge_flag(rig: Rig) -> None:
    """The Range descriptor that backs ``range_structure`` must observe geometry.

    The declarative ``range_structure {merged: ...}`` assertion reads
    ``RangesService.get_range`` -> ``RangeAddress.is_merged`` / ``merged_area``;
    this asserts that surface directly (the runner-backed case lives in
    ``test_test_runner_ranges_pivots``).
    """

    wid = await _open_for_mutation(rig)
    assert (await rig.ranges.merge_range(wid, "StructOps", "N9:P9", allow_data_loss=True)).verified

    merged = await rig.ranges.get_range(wid, "StructOps", "N9:P9")
    assert merged.is_merged is True
    assert merged.merged_area == "$N$9:$P$9"

    unmerged = await rig.ranges.get_range(wid, "StructOps", "A1:B2")
    assert unmerged.is_merged is False
    assert unmerged.merged_area is None


@requires_excel
async def test_structural_operations_reject_table_intersection(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    before = await rig.ranges.get_range(wid, "Cells", "T1:V4")

    calls = [
        ("insert", lambda: rig.ranges.insert_range(wid, "Cells", "T2:V4", "right")),
        ("delete", lambda: rig.ranges.delete_range(wid, "Cells", "U2:V3", "up")),
        (
            "merge",
            lambda: rig.ranges.merge_range(wid, "Cells", "T1:V4", allow_data_loss=True),
        ),
        ("unmerge", lambda: rig.ranges.unmerge_range(wid, "Cells", "T1:V4")),
        ("copy source", lambda: rig.ranges.copy_range(wid, "Cells", "T2:V4", "StructOps", "D2")),
        (
            "copy destination",
            lambda: rig.ranges.copy_range(wid, "StructOps", "A2:B4", "Cells", "T1"),
        ),
    ]
    for label, call in calls:
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.LISTOBJECT_INTERSECTION, label
        assert exc_info.value.context["table"] == "tblCells", label

        # The table geometry is unchanged after every rejected structural call.
        after = await rig.ranges.get_range(wid, "Cells", "T1:V4")
        assert (after.rows, after.cols, after.address) == (
            before.rows,
            before.cols,
            before.address,
        ), label


@requires_excel
async def test_protected_sheet_rejects_structural_mutations(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)
    await _protect_sheet(rig, wid, "StructOps")

    with pytest.raises(ExcelVbaMcpError) as insert_info:
        await rig.ranges.insert_range(wid, "StructOps", "B2", "right")
    assert insert_info.value.code is ErrorCode.PROTECTED_TARGET

    with pytest.raises(ExcelVbaMcpError) as merge_info:
        await rig.ranges.merge_range(wid, "StructOps", "J2:L2", allow_data_loss=True)
    assert merge_info.value.code is ErrorCode.PROTECTED_TARGET


# ===========================================================================
# 4. Data validation + conditional formatting
# ===========================================================================
@requires_excel
async def test_get_validation_reads_fixture_rules(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    list_rule = await rig.ranges.get_validation(wid, "Cells", "M4:M6")
    assert list_rule.present is True
    assert list_rule.type is DataValidationType.LIST
    assert list_rule.formula1 == '"Red,Green,Blue"'
    assert list_rule.apply_to == "$M$4:$M$6"

    whole_rule = await rig.ranges.get_validation(wid, "Cells", "N4:N6")
    assert whole_rule.type is DataValidationType.WHOLE_NUMBER
    assert whole_rule.formula1 == "1"
    assert whole_rule.formula2 == "10"

    absent = await rig.ranges.get_validation(wid, "Cells", "D1:D2")
    assert absent.present is False

    mixed = await rig.ranges.get_validation(wid, "Cells", "M4:N4")
    assert {"mixed_validation", "non_uniform_validation"} & set(mixed.unsupported)

    assert _sha256(rig.project) == before


@requires_excel
async def test_set_validation_adds_list_rule_and_reads_back(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    rule = await rig.ranges.set_validation(wid, "Cells", "W4:W6", _list_validation())

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rule.formula1 == '"Yes,No"'
    assert rule.in_cell_dropdown is True
    assert "postcondition_mismatch" not in rule.unsupported

    read_back = await rig.ranges.get_validation(wid, "Cells", "W4:W6")
    assert read_back.type is DataValidationType.LIST
    assert read_back.formula1 == '"Yes,No"'


@requires_excel
async def test_set_validation_modifies_existing_rule_preserving_unspecified(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    rule = await rig.ranges.set_validation(
        wid,
        "Cells",
        "N4:N6",
        ValidationSpec(
            type=DataValidationType.WHOLE_NUMBER,
            operator=ValidationOperator.BETWEEN,
            formula1="2",
            formula2="20",
        ),
    )

    assert rule.type is DataValidationType.WHOLE_NUMBER
    assert rule.formula1 == "2"
    assert rule.formula2 == "20"
    # An unspecified property keeps its previous value (Modify, not Add).
    assert rule.ignore_blank is True


@requires_excel
async def test_set_validation_rejects_unsupported_rule_without_mutation(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)
    before = await rig.ranges.get_validation(wid, "Cells", "N4:N6")

    with pytest.raises(ExcelVbaMcpError) as type_info:
        await rig.ranges.set_validation(
            wid, "Cells", "N4:N6", ValidationSpec(type=DataValidationType.INPUT_ONLY)
        )
    assert type_info.value.code is ErrorCode.INVALID_RANGE

    overlong = ValidationSpec.model_construct(
        type=DataValidationType.CUSTOM, formula1="x" * 256, formula2=None
    )
    with pytest.raises(ExcelVbaMcpError) as length_info:
        await rig.ranges.set_validation(wid, "Cells", "N4:N6", overlong)
    assert length_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_validation(wid, "Cells", "N4:N6")
    assert after.formula1 == before.formula1
    assert after.formula2 == before.formula2


@requires_excel
async def test_get_conditional_formatting_reads_fixture_rules(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    cell_value_rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "Q1:Q5")
    assert len(cell_value_rules) == 1
    assert cell_value_rules[0].type is ConditionalFormatType.CELL_VALUE
    assert cell_value_rules[0].operator is ValidationOperator.GREATER
    assert cell_value_rules[0].formula1 == "=5"
    assert cell_value_rules[0].priority == 1
    assert cell_value_rules[0].interior_color == 255

    expression_rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "R1:R5")
    assert len(expression_rules) == 1
    assert expression_rules[0].type is ConditionalFormatType.EXPRESSION
    assert expression_rules[0].formula1 == "=R1>1"
    assert expression_rules[0].interior_color == 65280


@requires_excel
async def test_set_conditional_formatting_add_and_priority_readback(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    result = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        _cell_value_cf(),
    )

    assert result.action == "add"
    assert result.verified is True
    # Sheet-wide priorities: Q=1, R=2, the new rule=3.
    assert result.applied_priorities == [3]

    read_back = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")
    assert len(read_back) == 1
    assert read_back[0].formula1 == "=10"
    assert read_back[0].priority == 3

    # The fixture rules keep their own priorities.
    q_rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "Q1:Q5")
    r_rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "R1:R5")
    assert [rule.priority for rule in q_rules] == [1]
    assert [rule.priority for rule in r_rules] == [2]


@requires_excel
async def test_set_conditional_formatting_modify_and_delete(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)
    await rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf())
    original = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")
    priority = original[0].priority

    modified = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.LESS,
            formula1="2",
            priority=priority,
        ),
        "modify",
    )
    assert modified.action == "modify"
    assert modified.verified is True
    assert modified.rules[0].operator is ValidationOperator.LESS
    assert modified.rules[0].formula1 == "=2"

    deleted = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.LESS,
            formula1="2",
            priority=priority,
        ),
        "delete",
    )
    assert deleted.action == "delete"
    assert deleted.verified is True
    assert deleted.rules == []
    assert await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5") == []


async def test_set_conditional_formatting_rejects_unsupported_type() -> None:
    # U3: color_scale/data_bar/icon_set/unsupported are rejected at the pydantic
    # request layer and never reach validation or COM.
    with pytest.raises(ValidationError):
        ConditionalFormatSpec(type=ConditionalFormatType.COLOR_SCALE)


# ===========================================================================
# 5. Snapshot / approval / recovery / audit
# ===========================================================================
@requires_excel
async def test_every_mutation_requires_a_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    # Approval is present, but no snapshot was taken.
    rig.grant_all_range_actions()

    for tool, call in _mutation_calls(rig, wid):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED, tool


@requires_excel
async def test_every_mutation_requires_a_writable_workbook(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant_all_range_actions()

    for tool, call in _mutation_calls(rig, wid):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY, tool


@requires_excel
async def test_every_mutation_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    # No range action was granted: the deny-by-default policy must reject.
    assert rig.approval is not None

    for tool, call in _mutation_calls(rig, wid):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await call()
        assert exc_info.value.code is ErrorCode.APPROVAL_DENIED, tool
        assert "approval" in exc_info.value.message.lower(), tool


@requires_excel
async def test_forced_hung_com_marks_needs_recovery(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    def hang() -> None:
        time.sleep(3.0)

    hang.__name__ = "task047_hang"
    with pytest.raises(ExcelVbaMcpError) as hung_info:
        await rig.worker.submit(
            hang,
            timeout=0.1,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )
    assert hung_info.value.code is ErrorCode.EXCEL_HUNG

    # Block E: a deadline-exceeded call is NOT a poison state — the worker moves to
    # running/deadline_exceeded and stays accepting (no fail-fast); the call is
    # never cancelled and the owned pid is never terminated.
    status = rig.worker.get_status()
    assert status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert status.accepting is True

    # No fail-fast: the single-STA FIFO serializes the follow-up behind the still-
    # running call (no parallel COM); once the hang returns the read succeeds.
    rng = await rig.ranges.get_range(wid, "Cells", "A1")
    assert rng.address
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


@requires_excel
async def test_audit_trail_records_mutating_operations(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    assert (
        await rig.ranges.write_values(
            wid, "Cells", "A20", _single_value(CellValueKind.STRING, "audit")
        )
    ).verified
    assert (
        await rig.ranges.set_format(wid, "Cells", "C20", RangeFormatPatch(font_bold=True))
    ).verified
    assert (await rig.ranges.copy_range(wid, "StructOps", "A2:B4", "StructOps", "D2")).verified
    assert (await rig.ranges.set_validation(wid, "Cells", "W4:W6", _list_validation())).present
    assert (
        await rig.ranges.set_conditional_formatting(wid, "Cells", "X1:X5", _cell_value_cf())
    ).verified
    await rig.ranges.read_values(wid, "Cells", "B2")

    rig.audit.flush()
    records = _read_audit(rig.audit.path)

    mutating = [record for record in records if record.get("operation_class") == "mutating"]
    tools = {record["tool"] for record in mutating}
    assert {
        "range_write_values",
        "range_set_format",
        "range_copy",
        "range_set_validation",
        "range_set_conditional_formatting",
    } <= tools

    assert any(
        record["tool"] == "range_read_values" and record.get("operation_class") == "read_only"
        for record in records
    )
    assert all(record.get("result") is not None for record in mutating)


# ===========================================================================
# 6. Response limits / invalid input
# ===========================================================================
@requires_excel
async def test_oversized_range_rejected_before_read(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    with pytest.raises(ExcelVbaMcpError) as values_info:
        await rig.ranges.read_values(wid, "Cells", "A1:XFD1048576")
    assert values_info.value.code is ErrorCode.INVALID_RANGE
    assert values_info.value.context["cells"] > MAX_RANGE_CELLS

    with pytest.raises(ExcelVbaMcpError) as formulas_info:
        await rig.ranges.read_formulas(wid, "Cells", "A1:XFD1048576")
    assert formulas_info.value.code is ErrorCode.INVALID_RANGE
    assert formulas_info.value.context["cells"] > MAX_RANGE_CELLS


@requires_excel
async def test_invalid_sheet_and_address_rejected(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    with pytest.raises(ExcelVbaMcpError) as sheet_info:
        await rig.ranges.get_range(wid, "NoSuchSheet", "A1")
    assert sheet_info.value.code is ErrorCode.INVALID_RANGE

    with pytest.raises(ExcelVbaMcpError) as address_info:
        await rig.ranges.get_range(wid, "Cells", "NotAnAddress")
    assert address_info.value.code is ErrorCode.INVALID_RANGE

    with pytest.raises(ExcelVbaMcpError) as read_info:
        await rig.ranges.read_values(wid, "Cells", "NotAnAddress")
    assert read_info.value.code is ErrorCode.INVALID_RANGE

    with pytest.raises(ExcelVbaMcpError) as offset_info:
        await rig.ranges.read_values(wid, "Cells", "B2:B15", offset=-1)
    assert offset_info.value.code is ErrorCode.INVALID_RANGE


@requires_excel
async def test_bounded_window_offset_and_limit(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)

    page = await rig.ranges.read_values(wid, "Cells", "B2:B15", offset=1, limit=2)
    assert page.rows == 2
    assert page.cols == 1
    assert float(page.values[0][0].value) == 42.0
    assert float(page.values[1][0].value) == 3.14

    beyond = await rig.ranges.read_values(wid, "Cells", "B2:B15", offset=100)
    assert beyond.rows == 0
    assert beyond.values == []


# ===========================================================================
# 8. Stage-1 sanity (TablesService / NamesService unchanged)
# ===========================================================================
@requires_excel
async def test_stage1_tables_service_still_lists_fixture_table(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    tables = await rig.tables.list_tables(wid)
    table = next(entry for entry in tables if entry.name == "tblCells")

    assert table.sheet_name == "Cells"
    assert table.range_address.replace("$", "") == "T1:V4"
    assert table.row_count == 3
    assert [column.name for column in table.columns] == ["id", "name", "qty"]

    fetched = await rig.tables.get_table(wid, "tblCells")
    assert fetched.name == "tblCells"
    assert [column.name for column in fetched.columns] == ["id", "name", "qty"]


@requires_excel
async def test_stage1_names_service_create_delete_round_trip(rig: Rig) -> None:
    wid = await _open_for_mutation(rig)
    rig.grant("name_create")
    rig.grant("name_delete")

    created = await rig.names.create_name(
        wid, "Task047Verify", "Cells!$A$1", scope=NameScope.WORKBOOK
    )
    assert created.name == "Task047Verify"
    assert created.scope is NameScope.WORKBOOK

    after_create = await rig.names.list_names(wid)
    assert "Task047Verify" in {entry.name for entry in after_create}

    deleted = await rig.names.delete_name(wid, "Task047Verify", scope=NameScope.WORKBOOK)
    assert deleted.name == "Task047Verify"

    after_delete = await rig.names.list_names(wid)
    assert "Task047Verify" not in {entry.name for entry in after_delete}
