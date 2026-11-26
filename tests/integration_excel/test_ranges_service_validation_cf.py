"""Real Excel integration tests for validation + conditional formatting (task_046).

These tests drive the real :class:`~excel_vba_mcp.application.ranges_service.
RangesService` and :class:`~excel_vba_mcp.adapters.excel.ranges_adapter.
RangesAdapter` against an **automation-created hidden Excel instance** (never the
user's Excel) and a disposable copy of the committed ``RangesCellsProject.xlsm``
fixture (task_041). They assert:

* the fixture's data validation reads back (``M4:M6`` list, ``N4:N6`` whole
  number) and a target without a rule reports ``present=false``;
* a non-uniform target is flagged instead of being flattened;
* a new list rule is added to a free range and an existing rule is modified in
  place (``Formula1``/``Formula2``/operator, unspecified properties preserved);
* an over-long formula / unsupported type is rejected before any mutation;
* the fixture's conditional formats read back (``Q1:Q5`` cell value, ``R1:R5``
  expression) with priority / ``stop_if_true`` / fill;
* add / modify / delete of a plain ``FormatCondition`` rule, priority readback and
  sheet-wide renumbering, and an explicit rejection of ``xlColorScale``;
* BG03 / BG04 / BG05 hold (owned Excel hidden + non-activating, source fixture
  byte-identical, worker refuses work on a hung instance, session RUNNING).

Excel is created with ``DispatchEx`` and always ``Visible = False``; no
``.Activate`` / ``.Select`` is ever called.
"""

from __future__ import annotations

import hashlib
import shutil
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest
from pydantic import ValidationError

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import ConditionalFormatSpec, ValidationSpec
from excel_vba_mcp.domain.statuses import (
    ConditionalFormatType,
    DataValidationType,
    OperationClass,
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


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class Rig:
    """Real COM worker + Excel + session/snapshot/policy stack for task_046."""

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
        self.ranges = RangesService(
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


async def _open_for_writes(rig: Rig) -> str:
    """Open the disposable copy writable, snapshot it and grant both write actions."""

    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_set_validation")
    rig.grant("range_set_conditional_formatting")
    return wid


# ---------------------------------------------------------------------------
# Read validation
# ---------------------------------------------------------------------------
@requires_excel
async def test_get_validation_list_from_fixture(rig: Rig) -> None:
    before = _sha256(rig.project)
    info = await rig.open()
    wid = info.workbook_id

    rule = await rig.ranges.get_validation(wid, "Cells", "M4:M6")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rule.operator is ValidationOperator.BETWEEN
    assert rule.formula1 == '"Red,Green,Blue"'
    assert rule.ignore_blank is True
    assert rule.in_cell_dropdown is True
    assert rule.alert_style == "xlValidAlertInformation"
    assert rule.apply_to == "$M$4:$M$6"
    assert rule.unsupported == []

    # BG03: a read never touches the source fixture.
    assert _sha256(rig.project) == before


@requires_excel
async def test_get_validation_whole_number_from_fixture(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    rule = await rig.ranges.get_validation(wid, "Cells", "N4:N6")

    assert rule.present is True
    assert rule.type is DataValidationType.WHOLE_NUMBER
    assert rule.formula1 == "1"
    assert rule.formula2 == "10"
    assert rule.in_cell_dropdown is False


@requires_excel
async def test_get_validation_absent_rule(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    rule = await rig.ranges.get_validation(wid, "Cells", "D1:D2")

    assert rule.present is False
    assert rule.type is DataValidationType.NONE
    assert rule.formula1 is None


@requires_excel
async def test_get_validation_non_uniform_target_is_flagged(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    # M4 has the list rule, N4 the whole-number rule: Excel refuses the
    # range-level read, so the adapter must report the target as non-uniform.
    rule = await rig.ranges.get_validation(wid, "Cells", "M4:N4")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert {"mixed_validation", "non_uniform_validation"} & set(rule.unsupported)


# ---------------------------------------------------------------------------
# Write validation
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_validation_adds_a_list_rule(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_writes(rig)

    rule = await rig.ranges.set_validation(
        wid,
        "Cells",
        "W1:W3",
        ValidationSpec(
            type=DataValidationType.LIST,
            operator=ValidationOperator.BETWEEN,
            formula1='"Yes,No"',
            in_cell_dropdown=True,
            ignore_blank=True,
        ),
    )

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rule.formula1 == '"Yes,No"'
    assert rule.in_cell_dropdown is True
    assert "postcondition_mismatch" not in rule.unsupported

    read_back = await rig.ranges.get_validation(wid, "Cells", "W1:W3")
    assert read_back.type is DataValidationType.LIST
    assert read_back.formula1 == '"Yes,No"'

    # BG03: the source fixture is unchanged after a validation mutation.
    assert _sha256(rig.project) == before


@requires_excel
async def test_set_validation_modifies_the_existing_rule(rig: Rig) -> None:
    wid = await _open_for_writes(rig)

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
async def test_set_validation_rejects_overlong_formula_without_mutation(rig: Rig) -> None:
    wid = await _open_for_writes(rig)
    before = await rig.ranges.get_validation(wid, "Cells", "N4:N6")

    spec = ValidationSpec.model_construct(
        type=DataValidationType.CUSTOM, formula1="x" * 256, formula2=None
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_validation(wid, "Cells", "N4:N6", spec)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_validation(wid, "Cells", "N4:N6")
    assert after.formula1 == before.formula1
    assert after.formula2 == before.formula2


@requires_excel
async def test_set_validation_rejects_unsupported_type(rig: Rig) -> None:
    wid = await _open_for_writes(rig)
    before = await rig.ranges.get_validation(wid, "Cells", "N4:N6")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_validation(
            wid, "Cells", "N4:N6", ValidationSpec(type=DataValidationType.INPUT_ONLY)
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_validation(wid, "Cells", "N4:N6")
    assert after.formula1 == before.formula1


# ---------------------------------------------------------------------------
# Read conditional formatting
# ---------------------------------------------------------------------------
@requires_excel
async def test_get_conditional_formatting_cell_value_from_fixture(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "Q1:Q5")

    assert len(rules) == 1
    rule = rules[0]
    assert rule.type is ConditionalFormatType.CELL_VALUE
    assert rule.operator is ValidationOperator.GREATER
    assert rule.formula1 == "=5"
    assert rule.priority == 1
    assert rule.stop_if_true is True
    assert rule.applies_to == "$Q$1:$Q$5"
    assert rule.interior_color == 255
    assert rule.supported is True


@requires_excel
async def test_get_conditional_formatting_expression_from_fixture(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id

    rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "R1:R5")

    assert len(rules) == 1
    rule = rules[0]
    assert rule.type is ConditionalFormatType.EXPRESSION
    assert rule.operator is None
    assert rule.formula1 == "=R1>1"
    assert rule.interior_color == 65280
    assert rule.supported is True


# ---------------------------------------------------------------------------
# Write conditional formatting
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_conditional_formatting_add_and_readback(rig: Rig) -> None:
    before = _sha256(rig.project)
    wid = await _open_for_writes(rig)

    result = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
            stop_if_true=True,
            interior_color=255,
        ),
    )

    assert result.action == "add"
    assert result.verified is True
    assert result.applied_priorities == [3]  # sheet-wide: Q=1, R=2, new=3
    assert len(result.rules) == 1

    rules = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")
    assert len(rules) == 1
    assert rules[0].type is ConditionalFormatType.CELL_VALUE
    assert rules[0].operator is ValidationOperator.GREATER
    assert rules[0].formula1 == "=10"
    assert rules[0].interior_color == 255
    assert rules[0].stop_if_true is True

    assert _sha256(rig.project) == before


@requires_excel
async def test_set_conditional_formatting_priority_readback_and_order(rig: Rig) -> None:
    wid = await _open_for_writes(rig)

    await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
        ),
    )
    result = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(type=ConditionalFormatType.EXPRESSION, formula1="=X1>0", priority=1),
    )

    assert result.applied_priorities == [1]
    priorities = [rule.priority for rule in result.rules]
    assert priorities == sorted(priorities)
    assert priorities[0] == 1
    assert result.rules[0].type is ConditionalFormatType.EXPRESSION
    assert any("sheet-wide" in note for note in result.notes)

    read_back = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")
    assert [rule.priority for rule in read_back] == priorities


@requires_excel
async def test_set_conditional_formatting_modify_rule(rig: Rig) -> None:
    wid = await _open_for_writes(rig)
    await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
        ),
    )
    original = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")

    result = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.LESS,
            formula1="2",
            priority=original[0].priority,
        ),
        "modify",
    )

    assert result.action == "modify"
    assert result.verified is True
    assert len(result.rules) == len(original)
    assert result.rules[0].operator is ValidationOperator.LESS
    assert result.rules[0].formula1 == "=2"


@requires_excel
async def test_set_conditional_formatting_delete_rule(rig: Rig) -> None:
    wid = await _open_for_writes(rig)
    await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
        ),
    )
    original = await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5")

    result = await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
            priority=original[0].priority,
        ),
        "delete",
    )

    assert result.action == "delete"
    assert result.verified is True
    assert result.rules == []
    assert any("gap" in note for note in result.notes)

    assert await rig.ranges.get_conditional_formatting(wid, "Cells", "X1:X5") == []


async def test_set_conditional_formatting_rejects_unsupported_type() -> None:
    # U3: color_scale/data_bar/icon_set/unsupported are rejected at the pydantic
    # request layer and never reach validation or COM.
    with pytest.raises(ValidationError):
        ConditionalFormatSpec(type=ConditionalFormatType.COLOR_SCALE)


@requires_excel
async def test_set_conditional_formatting_rejects_type_conversion(rig: Rig) -> None:
    wid = await _open_for_writes(rig)
    original = await rig.ranges.get_conditional_formatting(wid, "Cells", "Q1:Q5")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_conditional_formatting(
            wid,
            "Cells",
            "Q1:Q5",
            ConditionalFormatSpec(
                type=ConditionalFormatType.EXPRESSION,
                formula1="=Q1>0",
                priority=original[0].priority,
            ),
            "modify",
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE

    after = await rig.ranges.get_conditional_formatting(wid, "Cells", "Q1:Q5")
    assert after[0].type is ConditionalFormatType.CELL_VALUE
    assert after[0].formula1 == original[0].formula1


# ---------------------------------------------------------------------------
# Gates
# ---------------------------------------------------------------------------
@requires_excel
async def test_set_validation_requires_snapshot(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    rig.grant("range_set_validation")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_validation(
            wid,
            "Cells",
            "W1:W3",
            ValidationSpec(type=DataValidationType.LIST, formula1='"Yes,No"'),
        )
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


@requires_excel
async def test_set_validation_requires_approval(rig: Rig) -> None:
    info = await rig.open()
    wid = info.workbook_id
    await rig.create_snapshot(wid)

    with pytest.raises(ExcelVbaMcpError):
        await rig.ranges.set_validation(
            wid,
            "Cells",
            "W1:W3",
            ValidationSpec(type=DataValidationType.LIST, formula1='"Yes,No"'),
        )


@requires_excel
async def test_set_conditional_formatting_requires_writable(rig: Rig) -> None:
    info = await rig.open(read_only=True)
    wid = info.workbook_id
    await rig.create_snapshot(wid)
    rig.grant("range_set_conditional_formatting")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.ranges.set_conditional_formatting(
            wid,
            "Cells",
            "X1:X5",
            ConditionalFormatSpec(
                type=ConditionalFormatType.CELL_VALUE,
                operator=ValidationOperator.GREATER,
                formula1="10",
            ),
        )
    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


# ---------------------------------------------------------------------------
# Background safety (BG03 / BG04 / BG05)
# ---------------------------------------------------------------------------
@requires_excel
async def test_bg03_hidden_owned_excel_and_original_unchanged(rig: Rig) -> None:
    before = _sha256(rig.project)
    committed_before = _sha256(_RANGES_PROJECT)
    wid = await _open_for_writes(rig)

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
    assert flags["display_alerts"] is True

    await rig.ranges.set_validation(
        wid,
        "Cells",
        "W1:W3",
        ValidationSpec(type=DataValidationType.LIST, formula1='"Yes,No"'),
    )
    await rig.ranges.set_conditional_formatting(
        wid,
        "Cells",
        "X1:X5",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
        ),
    )

    # BG03: both the disposable original and the committed fixture are unchanged.
    assert _sha256(rig.project) == before
    assert _sha256(_RANGES_PROJECT) == committed_before


@requires_excel
async def test_bg04_worker_refuses_work_on_hung_excel(rig: Rig) -> None:
    wid = await _open_for_writes(rig)

    def hang() -> None:
        import time

        time.sleep(3.0)

    hang.__name__ = "hang_test"
    with pytest.raises(ExcelVbaMcpError):
        await rig.worker.submit(
            hang,
            timeout=0.1,
            section="diagnostic",
            op_class=OperationClass.MUTATING,
        )

    # Block E: a deadline-exceeded call is NOT a poison state — the worker stays
    # accepting (running/deadline_exceeded) and does NOT fail-fast. The single-STA
    # FIFO serializes the follow-up behind the still-running call (no parallel COM),
    # so once the hang returns the worker accepts the validation write again on its
    # own (no auto-kill, no recovery needed).
    hung_status = rig.worker.get_status()
    assert hung_status.state is WorkbookSessionState.DEADLINE_EXCEEDED
    assert hung_status.accepting is True

    applied = await rig.ranges.set_validation(
        wid,
        "Cells",
        "W1:W3",
        ValidationSpec(type=DataValidationType.LIST, formula1='"Yes,No"'),
    )
    assert applied.present
    assert rig.worker.get_status().state is WorkbookSessionState.RUNNING


@requires_excel
async def test_bg05_session_state_running(rig: Rig) -> None:
    await rig.open()
    assert rig.worker.get_status().state == WorkbookSessionState.RUNNING
