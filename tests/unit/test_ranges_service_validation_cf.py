"""Unit tests for the Range validation + conditional-formatting surface (task_046).

These tests exercise the production
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and
:class:`~excel_vba_mcp.adapters.excel.ranges_adapter.RangesAdapter` against a
pure-Python fake ``Workbook`` object graph (a fake STA worker that runs the
submitted callable inline). No Excel / COM is required.

Covered:

* schema / adapter validation of a validation spec (rule type, 255-char formulas,
  per-type operands) and of a conditional-format spec (unsupported rule type,
  missing identity, missing operator/formula, colour bounds);
* reading a rule, an absent rule, a non-uniform target and a mixed target;
* add vs modify of a validation rule (unspecified properties preserved);
* conditional-format add / modify / delete with sheet-wide priority renumbering,
  rule-order preservation and unsupported-type reporting;
* the snapshot / read-only / approval gates and the postcondition flags.
"""

from __future__ import annotations

import inspect
import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest
from pydantic import ValidationError

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.adapters.excel.ranges_adapter import (
    RangesAdapter,
    validate_conditional_format_spec,
    validate_validation_spec,
)
from excel_vba_mcp.api import range_tools
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    ConditionalFormatSpec,
    FontFormat,
    ValidationSpec,
)
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
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest, SnapshotStore

pytestmark = pytest.mark.anyio

# Excel enum values mirrored by the fake (the adapter owns the real constants).
XL_VALIDATE_WHOLE_NUMBER = 1
XL_VALIDATE_DECIMAL = 2
XL_VALIDATE_LIST = 3
XL_COND_BETWEEN = 1
XL_COND_GREATER = 3
XL_CF_CELL_VALUE = 1
XL_CF_EXPRESSION = 2
XL_CF_COLOR_SCALE = 3


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fake Excel object graph — validation rules live per cell; conditional-format
# rules live per range key with sheet-wide priorities (the real Excel semantics).
# ---------------------------------------------------------------------------
def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _col_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _parse_ref(ref: str) -> tuple[int, int]:
    cleaned = ref.replace("$", "")
    letters = "".join(ch for ch in cleaned if ch.isalpha())
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    return int(digits), _col_index(letters)


def _cell(row: int, col: int) -> str:
    return f"{_col_letter(col)}{row}"


def _validation_rule(
    type_value: int,
    operator_value: int,
    formula1: str,
    formula2: str = "",
    **overrides: Any,
) -> dict[str, Any]:
    """Build a fake validation rule with the same shape as Excel's defaults."""

    rule: dict[str, Any] = {
        "type": type_value,
        "operator": operator_value,
        "formula1": formula1,
        "formula2": formula2,
        "ignore_blank": True,
        "in_cell_dropdown": True,
        "alert_style": 2,
        "error_title": "",
        "error_message": "",
        "input_title": "",
        "input_message": "",
        "show_error": True,
        "show_input": True,
    }
    rule.update(overrides)
    return rule


class FakeValidationError(Exception):
    """Raised where Excel raises ``1004`` reading a ``Validation`` property."""


class FakeInterior:
    def __init__(self) -> None:
        self.Color = None  # noqa: N815 - COM casing

    @property
    def ColorIndex(self) -> int | None:  # noqa: N802 - COM casing
        return None if self.Color is None else 1


class FakeFont:
    def __init__(self) -> None:
        self.Name = "None"  # noqa: N815 - COM casing
        self.Bold = False  # noqa: N815 - COM casing
        self.Italic = False  # noqa: N815 - COM casing
        self.Size = None  # noqa: N815 - COM casing
        self.Color = None  # noqa: N815 - COM casing


class FakeAppliesTo:
    def __init__(self, address: str) -> None:
        self.Address = address  # noqa: N815 - COM casing


class FakeFormatCondition:
    """A plain ``FormatCondition`` double (the only supported rule shape)."""

    def __init__(
        self,
        sheet: FakeSheet,
        owner: FakeFormatConditions,
        type_value: int,
        operator_value: int | None,
        formula1: str | None,
        formula2: str | None,
        priority: int,
        stop_if_true: bool | None,
    ) -> None:
        self._sheet = sheet
        self._owner = owner
        self.Type = type_value  # noqa: N815 - COM casing
        self.Operator = operator_value  # noqa: N815 - COM casing
        self.Formula1 = formula1  # noqa: N815 - COM casing
        self.Formula2 = formula2  # noqa: N815 - COM casing
        self._priority = priority
        self.StopIfTrue = stop_if_true  # noqa: N815 - COM casing
        self.Interior = FakeInterior()
        self.Font = FakeFont()

    @property
    def Priority(self) -> int:  # noqa: N802 - COM casing
        return self._priority

    @Priority.setter
    def Priority(self, value: int) -> None:  # noqa: N802 - COM casing
        self._sheet.renumber_cf_rule(self, value)

    @property
    def AppliesTo(self) -> FakeAppliesTo:  # noqa: N802 - COM casing
        return FakeAppliesTo(self._owner.address)

    def Delete(self) -> None:  # noqa: N802 - COM casing
        self._owner.remove(self)


class FakeFormatConditions:
    def __init__(self, sheet: FakeSheet, address: str) -> None:
        self._sheet = sheet
        self.address = address
        self._rules: list[FakeFormatCondition] = sheet.cf_rules_for(address)

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._rules)

    def Item(self, index: int) -> FakeFormatCondition:  # noqa: N802 - COM casing
        return self._rules[index - 1]

    def Add(  # noqa: N802 - COM casing
        self,
        Type: int,  # noqa: N803 - COM casing
        Operator: int | None = None,  # noqa: N803 - COM casing
        Formula1: str | None = None,  # noqa: N803 - COM casing
        Formula2: str | None = None,  # noqa: N803 - COM casing
    ) -> FakeFormatCondition:
        # ``corrupt_cf`` is the postcondition test hook.
        formula1 = "TAMPERED" if self._sheet.corrupt_cf else Formula1
        if isinstance(formula1, str) and not formula1.startswith("="):
            # Excel always stores a rule formula with a leading "=".
            formula1 = f"={formula1}"
        rule = FakeFormatCondition(
            self._sheet,
            self,
            Type,
            Operator,
            formula1,
            Formula2,
            self._sheet.next_cf_priority(),
            None,
        )
        self._rules.append(rule)
        return rule

    def remove(self, rule: FakeFormatCondition) -> None:
        self._rules.remove(rule)


class FakeValidation:
    """A ``Validation`` double over the cells of one range.

    Reading a property raises :class:`FakeValidationError` when the cells carry
    no rule or disagree — the observable Excel behaviour the adapter relies on.
    """

    def __init__(self, sheet: FakeSheet, cells: list[tuple[int, int]]) -> None:
        self._sheet = sheet
        self._cells = cells

    # Internal rule access ------------------------------------------------
    def _rules(self) -> list[dict[str, Any]]:
        return [
            rule
            for rule in (self._sheet.validation_at(cell) for cell in self._cells)
            if rule is not None
        ]

    def _rule(self) -> dict[str, Any]:
        rules = self._rules()
        # Excel answers a range-level read only when *every* cell carries the same
        # rule; a hole (no rule) or a disagreement both raise 1004.
        if not rules or len(rules) != len(self._cells):
            raise FakeValidationError("no uniform validation rule")
        first = rules[0]
        if any(rule != first for rule in rules[1:]):
            raise FakeValidationError("mixed validation rules")
        return first

    # Properties ----------------------------------------------------------
    @property
    def Type(self) -> int:  # noqa: N802 - COM casing
        return int(self._rule()["type"])

    @Type.setter
    def Type(self, value: int) -> None:  # noqa: N802 - COM casing
        self._rule()["type"] = value

    @property
    def Operator(self) -> int:  # noqa: N802 - COM casing
        return int(self._rule()["operator"])

    @Operator.setter
    def Operator(self, value: int) -> None:  # noqa: N802 - COM casing
        self._rule()["operator"] = value

    @property
    def Formula1(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["formula1"])

    @Formula1.setter
    def Formula1(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["formula1"] = value

    @property
    def Formula2(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["formula2"])

    @Formula2.setter
    def Formula2(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["formula2"] = value

    @property
    def IgnoreBlank(self) -> bool:  # noqa: N802 - COM casing
        return bool(self._rule()["ignore_blank"])

    @IgnoreBlank.setter
    def IgnoreBlank(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._rule()["ignore_blank"] = value

    @property
    def InCellDropdown(self) -> bool:  # noqa: N802 - COM casing
        return bool(self._rule()["in_cell_dropdown"])

    @InCellDropdown.setter
    def InCellDropdown(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._rule()["in_cell_dropdown"] = value

    @property
    def AlertStyle(self) -> int:  # noqa: N802 - COM casing
        return int(self._rule()["alert_style"])

    @AlertStyle.setter
    def AlertStyle(self, value: int) -> None:  # noqa: N802 - COM casing
        self._rule()["alert_style"] = value

    @property
    def ErrorTitle(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["error_title"])

    @ErrorTitle.setter
    def ErrorTitle(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["error_title"] = value

    @property
    def ErrorMessage(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["error_message"])

    @ErrorMessage.setter
    def ErrorMessage(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["error_message"] = value

    @property
    def InputTitle(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["input_title"])

    @InputTitle.setter
    def InputTitle(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["input_title"] = value

    @property
    def InputMessage(self) -> str:  # noqa: N802 - COM casing
        return str(self._rule()["input_message"])

    @InputMessage.setter
    def InputMessage(self, value: str) -> None:  # noqa: N802 - COM casing
        self._rule()["input_message"] = value

    @property
    def ShowError(self) -> bool:  # noqa: N802 - COM casing
        return bool(self._rule()["show_error"])

    @ShowError.setter
    def ShowError(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._rule()["show_error"] = value

    @property
    def ShowInput(self) -> bool:  # noqa: N802 - COM casing
        return bool(self._rule()["show_input"])

    @ShowInput.setter
    def ShowInput(self, value: bool) -> None:  # noqa: N802 - COM casing
        self._rule()["show_input"] = value

    # Methods -------------------------------------------------------------
    def Add(  # noqa: N802 - COM casing
        self,
        Type: int,  # noqa: N803 - COM casing
        AlertStyle: int,  # noqa: N803 - COM casing
        Operator: int,  # noqa: N803 - COM casing
        Formula1: str,  # noqa: N803 - COM casing
        Formula2: str = "",  # noqa: N803 - COM casing
    ) -> None:
        if self._rules():
            raise FakeValidationError("a validation rule already exists")
        rule = _validation_rule(Type, Operator, Formula1, Formula2, alert_style=AlertStyle)
        if self._sheet.corrupt_validation:
            rule["formula1"] = "TAMPERED"
        for cell in self._cells:
            self._sheet.set_validation(cell, rule)

    def Modify(  # noqa: N802 - COM casing
        self,
        Type: int,  # noqa: N803 - COM casing
        AlertStyle: int,  # noqa: N803 - COM casing
        Operator: int,  # noqa: N803 - COM casing
        Formula1: str,  # noqa: N803 - COM casing
        Formula2: str = "",  # noqa: N803 - COM casing
    ) -> None:
        rule = self._rule()
        rule["type"] = Type
        rule["alert_style"] = AlertStyle
        rule["operator"] = Operator
        rule["formula1"] = Formula1
        rule["formula2"] = Formula2

    def Delete(self) -> None:  # noqa: N802 - COM casing
        for cell in self._cells:
            self._sheet.clear_validation(cell)


class FakeRange:
    def __init__(self, sheet: FakeSheet, top: int, left: int, height: int, width: int) -> None:
        self._sheet = sheet
        self.Top = top
        self.Left = left
        self.Height = height
        self.Width = width

    @property
    def Rows(self) -> FakeRange:  # noqa: N802 - COM casing
        return FakeRange(self._sheet, self.Top, self.Left, self.Height, 1)

    @property
    def Columns(self) -> FakeRange:  # noqa: N802 - COM casing
        return FakeRange(self._sheet, self.Top, self.Left, 1, self.Width)

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self.Height * self.Width

    @property
    def Address(self) -> str:  # noqa: N802 - COM casing
        if self.Height == 1 and self.Width == 1:
            return f"${_col_letter(self.Left)}${self.Top}"
        return (
            f"${_col_letter(self.Left)}${self.Top}:"
            f"${_col_letter(self.Left + self.Width - 1)}${self.Top + self.Height - 1}"
        )

    @property
    def Areas(self) -> FakeAreas:  # noqa: N802 - COM casing
        return FakeAreas(1)

    @property
    def Validation(self) -> FakeValidation:  # noqa: N802 - COM casing
        return FakeValidation(self._sheet, self._cells())

    @property
    def FormatConditions(self) -> FakeFormatConditions:  # noqa: N802 - COM casing
        return FakeFormatConditions(self._sheet, self.Address)

    def Cells(self, row: int, col: int) -> FakeCell:  # noqa: N802 - COM casing
        return FakeCell(self._sheet, self.Top + row - 1, self.Left + col - 1)

    def _cells(self) -> list[tuple[int, int]]:
        return [
            (row, col)
            for row in range(self.Top, self.Top + self.Height)
            for col in range(self.Left, self.Left + self.Width)
        ]


class FakeAreas:
    def __init__(self, count: int) -> None:
        self._count = count

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return self._count


class FakeCell:
    def __init__(self, sheet: FakeSheet, row: int, col: int) -> None:
        self._sheet = sheet
        self.row = row
        self.col = col

    @property
    def Validation(self) -> FakeValidation:  # noqa: N802 - COM casing
        return FakeValidation(self._sheet, [(self.row, self.col)])


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name
        self.ProtectContents = False
        self.corrupt_validation = False
        self.corrupt_cf = False
        self._validations: dict[tuple[int, int], dict[str, Any]] = {}
        self._cf_rules: dict[str, list[FakeFormatCondition]] = {}
        self._cf_priority = 0

    # Validation ----------------------------------------------------------
    def validation_at(self, cell: tuple[int, int]) -> dict[str, Any] | None:
        return self._validations.get(cell)

    def set_validation(self, cell: tuple[int, int], rule: dict[str, Any]) -> None:
        self._validations[cell] = rule

    def clear_validation(self, cell: tuple[int, int]) -> None:
        self._validations.pop(cell, None)

    # Conditional formatting ---------------------------------------------
    def cf_rules_for(self, address: str) -> list[FakeFormatCondition]:
        return self._cf_rules.setdefault(address, [])

    def next_cf_priority(self) -> int:
        self._cf_priority += 1
        return self._cf_priority

    def all_cf_rules(self) -> list[FakeFormatCondition]:
        return [rule for rules in self._cf_rules.values() for rule in rules]

    def renumber_cf_rule(self, rule: FakeFormatCondition, priority: int) -> None:
        """Replicate Excel's sheet-wide priority renumbering on assignment."""

        rules = [item for item in self.all_cf_rules() if item is not rule]
        rules.sort(key=lambda item: item.Priority)
        index = min(max(priority, 1), len(rules) + 1) - 1
        rules.insert(index, rule)
        for position, item in enumerate(rules, start=1):
            item._priority = position  # noqa: SLF001 - test double

    # COM-ish range resolution -------------------------------------------
    def Range(self, address: str) -> FakeRange:  # noqa: N802 - COM casing
        start, _, end = address.replace("$", "").partition(":")
        top, left = _parse_ref(start)
        bottom, right = _parse_ref(end) if end else (top, left)
        return FakeRange(self, top, left, bottom - top + 1, right - left + 1)


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

    def __iter__(self) -> Any:
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
        wb = FakeWorkbook([FakeSheet("Data"), FakeSheet("Other")])
        wb.FullName = path
        wb.ReadOnly = bool(kwargs.get("ReadOnly", False))
        self.items.append(wb)
        return wb

    def __iter__(self) -> Any:
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


class Rig:
    """Real service stack over the fake worker/Excel graph."""

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
        self.service = RangesService(
            self.worker,  # type: ignore[arg-type]
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )

    async def open(self, *, read_only: bool = False) -> Any:
        self.grant("excel_open")
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


def _fake_workbook() -> FakeWorkbook:
    """A workbook with one sheet carrying the fake validation/CF scaffolding."""

    return FakeWorkbook([FakeSheet("Data"), FakeSheet("Other")])


def _list_rule(formula1: str = '"Red,Green,Blue"') -> dict[str, Any]:
    return _validation_rule(XL_VALIDATE_LIST, XL_COND_BETWEEN, formula1)


def _whole_rule() -> dict[str, Any]:
    return _validation_rule(XL_VALIDATE_WHOLE_NUMBER, XL_COND_BETWEEN, "1", "10")


# ---------------------------------------------------------------------------
# Schema / adapter validation (pure, pre-COM)
# ---------------------------------------------------------------------------
def test_validation_spec_rejects_overlong_formula_at_model_level() -> None:
    with pytest.raises(ValidationError):
        ValidationSpec(type=DataValidationType.CUSTOM, formula1="x" * 256)


def test_validate_validation_spec_accepts_a_normal_rule() -> None:
    validate_validation_spec(
        ValidationSpec(
            type=DataValidationType.LIST,
            operator=ValidationOperator.BETWEEN,
            formula1='"A,B,C"',
            in_cell_dropdown=True,
        )
    )


def test_validate_validation_spec_rejects_ruleless_types() -> None:
    for rule_type in (DataValidationType.NONE, DataValidationType.INPUT_ONLY):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            validate_validation_spec(ValidationSpec(type=rule_type, formula1="1"))
        assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_validation_spec_rejects_missing_formula() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_validation_spec(ValidationSpec(type=DataValidationType.WHOLE_NUMBER))
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_validation_spec_rejects_overlong_formula() -> None:
    # ``model_construct`` bypasses the model validator so the adapter guard is tested.
    spec = ValidationSpec.model_construct(
        type=DataValidationType.CUSTOM, formula1="x" * 256, formula2=None
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_validation_spec(spec)
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_validation_spec_rejects_formula2_for_single_formula_types() -> None:
    for rule_type in (DataValidationType.LIST, DataValidationType.CUSTOM):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            validate_validation_spec(ValidationSpec(type=rule_type, formula1='"A"', formula2="1"))
        assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_validation_spec_requires_both_operands_for_between() -> None:
    for operator in (ValidationOperator.BETWEEN, ValidationOperator.NOT_BETWEEN):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            validate_validation_spec(
                ValidationSpec(
                    type=DataValidationType.WHOLE_NUMBER, operator=operator, formula1="1"
                )
            )
        assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_rejects_unsupported_types() -> None:
    # U3: color_scale/data_bar/icon_set/unsupported are rejected at the pydantic
    # request layer and never reach validation or COM.
    for rule_type in (
        ConditionalFormatType.COLOR_SCALE,
        ConditionalFormatType.DATA_BAR,
        ConditionalFormatType.ICON_SET,
        ConditionalFormatType.UNSUPPORTED,
    ):
        with pytest.raises(ValidationError):
            ConditionalFormatSpec(type=rule_type, formula1="1")


def test_validate_conditional_format_spec_requires_identity_for_modify_and_delete() -> None:
    for mode in ("modify", "delete"):
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            validate_conditional_format_spec(
                ConditionalFormatSpec(type=ConditionalFormatType.CELL_VALUE, formula1="1"),
                mode,
            )
        assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_requires_operator_for_cell_value() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_conditional_format_spec(
            ConditionalFormatSpec(type=ConditionalFormatType.CELL_VALUE, formula1="1"), "add"
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_requires_formula_for_expression() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_conditional_format_spec(
            ConditionalFormatSpec(
                type=ConditionalFormatType.EXPRESSION, operator=ValidationOperator.BETWEEN
            ),
            "add",
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_rejects_unknown_mode() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_conditional_format_spec(
            ConditionalFormatSpec(type=ConditionalFormatType.EXPRESSION, formula1="=1"),
            "replace",
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_rejects_out_of_range_colour() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        validate_conditional_format_spec(
            ConditionalFormatSpec(
                type=ConditionalFormatType.CELL_VALUE,
                operator=ValidationOperator.GREATER,
                formula1="1",
                interior_color=0x1000000,
            ),
            "add",
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


def test_validate_conditional_format_spec_accepts_font_subset() -> None:
    validate_conditional_format_spec(
        ConditionalFormatSpec(
            type=ConditionalFormatType.EXPRESSION,
            formula1="=A1>0",
            font=FontFormat(name="Calibri", bold=True, size=11),
        ),
        "add",
    )


# ---------------------------------------------------------------------------
# Adapter — validation reads
# ---------------------------------------------------------------------------
def test_get_validation_reports_absent_rule() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()

    rule = adapter.get_validation(wb, "Data", "A1:A2")

    assert rule.present is False
    assert rule.type is DataValidationType.NONE
    assert rule.apply_to == "$A$1:$A$2"


def test_get_validation_reads_list_rule() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    shared = _list_rule()
    for row in (1, 2):
        sheet.set_validation((row, 1), shared)
    adapter = RangesAdapter()

    rule = adapter.get_validation(wb, "Data", "A1:A2")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rule.operator is ValidationOperator.BETWEEN
    assert rule.formula1 == '"Red,Green,Blue"'
    assert rule.formula2 == ""
    assert rule.ignore_blank is True
    assert rule.in_cell_dropdown is True
    assert rule.alert_style == "xlValidAlertStop"
    assert rule.unsupported == []


def test_get_validation_reads_whole_number_rule() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    shared = _whole_rule()
    for row in (1, 2):
        sheet.set_validation((row, 1), shared)

    rule = RangesAdapter().get_validation(wb, "Data", "A1:A2")

    assert rule.type is DataValidationType.WHOLE_NUMBER
    assert rule.formula1 == "1"
    assert rule.formula2 == "10"


def test_get_validation_flags_mixed_target() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    sheet.set_validation((1, 1), _list_rule())
    sheet.set_validation((1, 2), _whole_rule())

    rule = RangesAdapter().get_validation(wb, "Data", "A1:B1")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert "mixed_validation" in rule.unsupported
    assert any(item.startswith("distinct_rules=") for item in rule.unsupported)


def test_get_validation_flags_non_uniform_coverage() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    sheet.set_validation((1, 1), _list_rule())

    rule = RangesAdapter().get_validation(wb, "Data", "A1:A3")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert "non_uniform_validation" in rule.unsupported


# ---------------------------------------------------------------------------
# Adapter — validation writes
# ---------------------------------------------------------------------------
def test_set_validation_adds_a_list_rule() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()

    rule = adapter.set_validation(
        wb,
        "Data",
        "A1:A3",
        ValidationSpec(
            type=DataValidationType.LIST,
            operator=ValidationOperator.BETWEEN,
            formula1='"A,B,C"',
            in_cell_dropdown=True,
        ),
    )

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rule.formula1 == '"A,B,C"'
    assert "postcondition_mismatch" not in rule.unsupported
    sheet = wb.sheet("Data")
    assert sheet.validation_at((1, 1)) is not None
    assert sheet.validation_at((3, 1)) is not None


def test_set_validation_modifies_the_existing_rule() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    shared = _whole_rule()
    shared["ignore_blank"] = False
    for row in (1, 2):
        sheet.set_validation((row, 1), shared)
    adapter = RangesAdapter()

    rule = adapter.set_validation(
        wb,
        "Data",
        "A1:A2",
        ValidationSpec(
            type=DataValidationType.DECIMAL,
            operator=ValidationOperator.BETWEEN,
            formula1="0",
            formula2="1",
        ),
    )

    assert rule.type is DataValidationType.DECIMAL
    assert rule.formula1 == "0"
    assert rule.formula2 == "1"
    # An unspecified property keeps its previous value (Modify, not Add).
    assert rule.ignore_blank is False
    assert len(sheet._validations) == 2  # noqa: SLF001 - test double


def test_set_validation_applies_explicit_message_and_flags() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()

    rule = adapter.set_validation(
        wb,
        "Data",
        "B2",
        ValidationSpec(
            type=DataValidationType.WHOLE_NUMBER,
            operator=ValidationOperator.BETWEEN,
            formula1="1",
            formula2="10",
            ignore_blank=True,
            in_cell_dropdown=False,
            error_title="Too big",
            error_message="Enter 1..10",
            show_error=True,
        ),
    )

    assert rule.ignore_blank is True
    assert rule.in_cell_dropdown is False
    assert rule.error_title == "Too big"
    assert rule.error_message == "Enter 1..10"
    assert rule.show_error is True


def test_set_validation_rejects_a_malformed_rule_before_com() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.set_validation(
            wb,
            "Data",
            "A1",
            ValidationSpec.model_construct(type=DataValidationType.CUSTOM, formula1="x" * 256),
        )
    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert wb.sheet("Data").validation_at((1, 1)) is None


def test_set_validation_flags_postcondition_mismatch() -> None:
    wb = _fake_workbook()
    wb.sheet("Data").corrupt_validation = True
    adapter = RangesAdapter()

    rule = adapter.set_validation(
        wb,
        "Data",
        "A1",
        ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
    )

    assert rule.formula1 == "TAMPERED"
    assert "postcondition_mismatch" in rule.unsupported


# ---------------------------------------------------------------------------
# Adapter — conditional-formatting reads
# ---------------------------------------------------------------------------
def _seed_cf(sheet: FakeSheet, address: str, *specs: tuple[int, int | None, str | None]) -> None:
    """Seed conditional-format rules through the fake's COM surface."""

    conditions = sheet.Range(address).FormatConditions
    for type_value, operator_value, formula1 in specs:
        conditions.Add(type_value, operator_value, formula1)


def test_get_conditional_formatting_returns_rules_in_priority_order() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(
        sheet,
        "A1:A3",
        (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"),
        (XL_CF_EXPRESSION, None, "=A1>1"),
    )

    rules = RangesAdapter().get_conditional_formatting(wb, "Data", "A1:A3")

    assert [rule.type for rule in rules] == [
        ConditionalFormatType.CELL_VALUE,
        ConditionalFormatType.EXPRESSION,
    ]
    assert [rule.priority for rule in rules] == [1, 2]
    assert rules[0].operator is ValidationOperator.GREATER
    assert rules[0].formula1 == "=5"
    assert rules[0].applies_to == "$A$1:$A$3"
    assert rules[1].operator is None
    assert rules[1].formula1 == "=A1>1"


def test_get_conditional_formatting_reports_unsupported_types() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(sheet, "A1:A3", (XL_CF_COLOR_SCALE, None, None))

    rules = RangesAdapter().get_conditional_formatting(wb, "Data", "A1:A3")

    assert len(rules) == 1
    assert rules[0].type is ConditionalFormatType.COLOR_SCALE
    assert rules[0].supported is False
    assert "ColorScale" in (rules[0].unsupported_reason or "")
    assert rules[0].interior_color is None
    assert rules[0].font is None


def test_get_conditional_formatting_reads_fill_and_font() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    sheet.Range("A1:A3").FormatConditions.Add(XL_CF_CELL_VALUE, XL_COND_GREATER, "=5")
    rule = sheet.Range("A1:A3").FormatConditions.Item(1)
    rule.Interior.Color = 255
    rule.Font.Bold = True
    rule.Font.Size = 12
    rule.StopIfTrue = True

    read = RangesAdapter().get_conditional_formatting(wb, "Data", "A1:A3")[0]

    assert read.interior_color == 255
    assert read.stop_if_true is True
    assert read.font is not None
    assert read.font.bold is True
    assert read.font.size == 12


# ---------------------------------------------------------------------------
# Adapter — conditional-formatting writes
# ---------------------------------------------------------------------------
def test_set_conditional_formatting_adds_a_rule() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()

    result = adapter.set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
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
    assert result.applied_priorities == [1]
    assert len(result.rules) == 1
    assert result.rules[0].interior_color == 255
    assert result.rules[0].stop_if_true is True
    assert result.unsupported == []


def test_set_conditional_formatting_add_preserves_priority_order() -> None:
    wb = _fake_workbook()
    adapter = RangesAdapter()
    adapter.set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="1",
        ),
    )

    result = adapter.set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(type=ConditionalFormatType.EXPRESSION, formula1="=A1>0", priority=1),
    )

    assert [rule.priority for rule in result.rules] == [1, 2]
    assert result.rules[0].type is ConditionalFormatType.EXPRESSION
    assert result.applied_priorities == [1]
    assert any("sheet-wide" in note for note in result.notes)


def test_set_conditional_formatting_modify_replaces_the_rule() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(
        sheet,
        "A1:A3",
        (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"),
        (XL_CF_EXPRESSION, None, "=A1>1"),
    )
    adapter = RangesAdapter()

    result = adapter.set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.LESS,
            formula1="2",
            priority=1,
        ),
        "modify",
    )

    assert result.action == "modify"
    assert result.verified is True
    assert len(result.rules) == 2
    first = result.rules[0]
    assert first.priority == 1
    assert first.type is ConditionalFormatType.CELL_VALUE
    assert first.operator is ValidationOperator.LESS
    assert first.formula1 == "=2"


def test_set_conditional_formatting_modify_rejects_type_conversion() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(sheet, "A1:A3", (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"))

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        RangesAdapter().set_conditional_formatting(
            wb,
            "Data",
            "A1:A3",
            ConditionalFormatSpec(
                type=ConditionalFormatType.EXPRESSION, formula1="=A1>1", priority=1
            ),
            "modify",
        )

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert "never silently converted" in str(exc_info.value)
    # The existing rule is untouched.
    assert len(sheet.cf_rules_for("$A$1:$A$3")) == 1


def test_set_conditional_formatting_delete_rejects_unknown_priority() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(sheet, "A1:A3", (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"))

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        RangesAdapter().set_conditional_formatting(
            wb,
            "Data",
            "A1:A3",
            ConditionalFormatSpec(
                type=ConditionalFormatType.CELL_VALUE,
                operator=ValidationOperator.GREATER,
                formula1="5",
                priority=2,
            ),
            "delete",
        )

    assert exc_info.value.code is ErrorCode.INVALID_RANGE
    assert "priority 2" in str(exc_info.value)
    # The existing rule is untouched.
    assert len(sheet.cf_rules_for("$A$1:$A$3")) == 1


def test_set_conditional_formatting_add_accepts_excel_formula_normalization() -> None:
    wb = _fake_workbook()

    result = RangesAdapter().set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="10",
        ),
    )

    # Excel stores the operand with a leading "="; the postcondition comparison
    # normalizes both spellings instead of reporting a false mismatch.
    assert result.verified is True
    assert result.rules[0].formula1 == "=10"


def test_set_conditional_formatting_delete_removes_the_rule() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(
        sheet,
        "A1:A3",
        (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"),
        (XL_CF_EXPRESSION, None, "=A1>1"),
    )

    result = RangesAdapter().set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(type=ConditionalFormatType.EXPRESSION, formula1="=A1>1", priority=2),
        "delete",
    )

    assert result.action == "delete"
    assert result.verified is True
    assert len(result.rules) == 1
    assert result.rules[0].priority == 1
    assert any("gap" in note for note in result.notes)


def test_set_conditional_formatting_rejects_unsupported_type() -> None:
    # U3: data_bar is rejected at the pydantic request layer; the adapter is never
    # reached and no rule is written.
    with pytest.raises(ValidationError):
        ConditionalFormatSpec(type=ConditionalFormatType.DATA_BAR, formula1="1")


def test_set_conditional_formatting_reports_unsupported_rules_in_range() -> None:
    wb = _fake_workbook()
    sheet = wb.sheet("Data")
    _seed_cf(sheet, "A1:A3", (XL_CF_COLOR_SCALE, None, None))

    result = RangesAdapter().set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="1",
        ),
    )

    assert result.unsupported == ["color_scale"]
    assert len(result.rules) == 2


def test_set_conditional_formatting_flags_postcondition_mismatch() -> None:
    wb = _fake_workbook()
    wb.sheet("Data").corrupt_cf = True

    result = RangesAdapter().set_conditional_formatting(
        wb,
        "Data",
        "A1:A3",
        ConditionalFormatSpec(
            type=ConditionalFormatType.CELL_VALUE,
            operator=ValidationOperator.GREATER,
            formula1="1",
        ),
    )

    assert result.verified is False
    assert any("Postcondition mismatch" in note for note in result.notes)


def test_set_conditional_formatting_rejects_a_protected_sheet() -> None:
    wb = _fake_workbook()
    wb.sheet("Data").ProtectContents = True

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        RangesAdapter().set_conditional_formatting(
            wb,
            "Data",
            "A1:A3",
            ConditionalFormatSpec(
                type=ConditionalFormatType.CELL_VALUE,
                operator=ValidationOperator.GREATER,
                formula1="1",
            ),
        )

    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET


def test_set_validation_rejects_a_protected_sheet() -> None:
    wb = _fake_workbook()
    wb.sheet("Data").ProtectContents = True

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        RangesAdapter().set_validation(
            wb,
            "Data",
            "A1",
            ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
        )

    assert exc_info.value.code is ErrorCode.PROTECTED_TARGET


# ---------------------------------------------------------------------------
# Service — guards, submission classification and validation-before-gates
# ---------------------------------------------------------------------------
async def test_service_get_validation_reads_through_the_worker(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    sheet = rig.workbook().sheet("Data")
    shared = _list_rule()
    for row in (1, 2):
        sheet.set_validation((row, 1), shared)

    rule = await rig.service.get_validation(info.workbook_id, "Data", "A1:A2")

    assert rule.present is True
    assert rule.type is DataValidationType.LIST
    assert rig.worker.submits[-1]["op_class"] is OperationClass.READ_ONLY
    assert rig.worker.submits[-1]["name"] == "range_get_validation"


async def test_service_get_conditional_formatting_reads_through_the_worker(
    tmp_path: Path,
) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    _seed_cf(rig.workbook().sheet("Data"), "A1:A3", (XL_CF_CELL_VALUE, XL_COND_GREATER, "=5"))

    rules = await rig.service.get_conditional_formatting(info.workbook_id, "Data", "A1:A3")

    assert [rule.priority for rule in rules] == [1]
    assert rig.worker.submits[-1]["op_class"] is OperationClass.READ_ONLY
    assert rig.worker.submits[-1]["name"] == "range_get_conditional_formatting"


async def test_service_set_validation_requires_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.grant("range_set_validation")

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_validation(
            info.workbook_id,
            "Data",
            "A1",
            ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
        )

    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED


async def test_service_set_validation_requires_approval(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_validation(
            info.workbook_id,
            "Data",
            "A1",
            ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
        )

    assert exc_info.value.code is ErrorCode.APPROVAL_DENIED


async def test_service_set_validation_rejects_read_only_session(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_validation(
            info.workbook_id,
            "Data",
            "A1",
            ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
        )

    assert exc_info.value.code is ErrorCode.WORKBOOK_READ_ONLY


async def test_service_set_validation_validates_before_the_gates(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=True)

    spec = ValidationSpec.model_construct(
        type=DataValidationType.CUSTOM, formula1="x" * 256, formula2=None
    )
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_validation(info.workbook_id, "Data", "A1", spec)

    # INVALID_RANGE (payload) wins over WORKBOOK_READ_ONLY (session).
    assert exc_info.value.code is ErrorCode.INVALID_RANGE


async def test_service_set_validation_submits_mutating_and_audits(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    rig.add_snapshot(info.workbook_id)
    rig.grant("range_set_validation")

    rule = await rig.service.set_validation(
        info.workbook_id,
        "Data",
        "A1",
        ValidationSpec(type=DataValidationType.LIST, formula1='"A,B"'),
    )

    assert rule.present is True
    assert rig.worker.submits[-1]["op_class"] is OperationClass.MUTATING
    assert rig.worker.submits[-1]["name"] == "range_set_validation"
    audit_lines = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    actions = [entry["tool"] for entry in audit_lines]
    assert "range_set_validation" in actions
    entry = next(item for item in audit_lines if item["tool"] == "range_set_validation")
    assert entry["operation_class"] == "mutating"
    assert entry["result"]["verified"] is True


async def test_service_set_conditional_formatting_guard_chain(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open()
    spec = ConditionalFormatSpec(
        type=ConditionalFormatType.CELL_VALUE,
        operator=ValidationOperator.GREATER,
        formula1="1",
    )

    rig.grant("range_set_conditional_formatting")
    with pytest.raises(ExcelVbaMcpError) as exc_info:
        await rig.service.set_conditional_formatting(info.workbook_id, "Data", "A1:A3", spec)
    assert exc_info.value.code is ErrorCode.SNAPSHOT_FAILED

    rig.add_snapshot(info.workbook_id)
    result = await rig.service.set_conditional_formatting(info.workbook_id, "Data", "A1:A3", spec)
    assert result.verified is True
    assert rig.worker.submits[-1]["op_class"] is OperationClass.MUTATING
    assert rig.worker.submits[-1]["name"] == "range_set_conditional_formatting"


async def test_service_set_conditional_formatting_validates_before_the_gates(
    tmp_path: Path,
) -> None:
    # U3: color_scale is rejected at the pydantic request layer — before the
    # service is called and before any snapshot/approval gate.
    with pytest.raises(ValidationError):
        ConditionalFormatSpec(type=ConditionalFormatType.COLOR_SCALE, formula1="1")


# ---------------------------------------------------------------------------
# API surface
# ---------------------------------------------------------------------------
def test_api_tools_are_async_and_delegate() -> None:
    for name in (
        "range_get_validation",
        "range_set_validation",
        "range_get_conditional_formatting",
        "range_set_conditional_formatting",
    ):
        tool = getattr(range_tools, name)
        assert inspect.iscoroutinefunction(tool)
    assert range_tools.get_ranges_service.__name__ == "get_ranges_service"


def test_api_tools_reject_blank_arguments() -> None:
    for name, args in (
        ("range_get_validation", ("wb", "", "A1")),
        ("range_get_validation", ("wb", "Data", "  ")),
        ("range_set_validation", ("", "Data", "A1")),
        ("range_get_conditional_formatting", ("wb", "Data", "")),
        ("range_set_conditional_formatting", ("wb", "Data", "A1")),
    ):
        tool = getattr(range_tools, name)
        kwargs: dict[str, Any] = {}
        if name == "range_set_validation":
            kwargs["spec"] = ValidationSpec(type=DataValidationType.LIST, formula1='"A"')
        if name == "range_set_conditional_formatting":
            kwargs["spec"] = ConditionalFormatSpec(
                type=ConditionalFormatType.EXPRESSION, formula1="=1"
            )
        with pytest.raises(ExcelVbaMcpError):
            import asyncio

            asyncio.run(tool(*args, **kwargs))
