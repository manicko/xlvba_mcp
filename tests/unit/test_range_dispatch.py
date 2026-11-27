"""Unit tests for the WP7 range dispatch tools (plan B8).

These tests never touch COM / Excel. They install a fake ``RangesService`` via
the ``range_tools`` setter and assert that the flat-signature ``range_read`` /
``range_write`` / ``range_structural`` dispatchers route to the existing
per-verb handlers, enforce per-action required parameters as structured
:class:`ExcelVbaMcpError` (option (a) — handler validation), and advertise the
truthful annotation triples from the WP0 matrix: ``range_read`` is (T,F,T),
``range_write`` is (F,F,F) — NOT idempotent because
``set-conditional-formatting`` mode='add' appends a rule — ``range_structural``
is (F,T,F) and NOT idempotent (H4). ``range_delete`` (shift in {left,up}) and
``range_merge`` stay standalone; ``range_merge`` keeps the intentional dual-hint
(F,T,T).
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import range_tools
from excel_vba_mcp.api.range_tools import (
    range_delete,
    range_merge,
    range_read,
    range_structural,
    range_write,
)
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    CellValue,
    ConditionalFormatMutationResult,
    ConditionalFormatSpec,
    FormatInfo,
    FormatMutationResult,
    FormulaCell,
    FormulaWriteRequest,
    RangeAddress,
    RangeFormatPatch,
    RangeFormulas,
    RangeMutationResult,
    RangeShape,
    RangeValues,
    RangeWriteRequest,
    StructuralMutationResult,
    ValidationRule,
    ValidationSpec,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    ConditionalFormatRequestType,
    DataValidationType,
)

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _addr() -> RangeAddress:
    return RangeAddress(sheet="Sheet1", address="A1", rows=1, cols=1)


def _values() -> RangeValues:
    return RangeValues(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        values=[[CellValue(kind=CellValueKind.NUMBER, value=1)]],
    )


def _formulas() -> RangeFormulas:
    return RangeFormulas(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        formulas=[[FormulaCell(address="A1", formula2="=1+1")]],
    )


def _write_payload() -> RangeWriteRequest:
    return RangeWriteRequest(
        rows=1,
        cols=2,
        values=[
            [
                CellValue(kind=CellValueKind.NUMBER, value=1),
                CellValue(kind=CellValueKind.NUMBER, value=2),
            ]
        ],
    )


def _formula_payload() -> FormulaWriteRequest:
    return FormulaWriteRequest(rows=1, cols=1, formulas=[["=1+1"]])


def _range_mutation(action: str) -> RangeMutationResult:
    return RangeMutationResult(action=action)


def _format_mutation() -> FormatMutationResult:
    return FormatMutationResult(range=_addr())


def _cf_mutation(action: str) -> ConditionalFormatMutationResult:
    return ConditionalFormatMutationResult(action=action, range=_addr())


def _structural(action: str) -> StructuralMutationResult:
    return StructuralMutationResult(action=action)


class FakeRangesService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def get_range(self, workbook_id: str, sheet: str, address: str) -> RangeAddress:
        self.calls.append(("get_range", (workbook_id, sheet, address)))
        return _addr()

    async def read_values(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeValues:
        self.calls.append(("read_values", (workbook_id, sheet, address, offset, limit)))
        return _values()

    async def read_formulas(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeFormulas:
        self.calls.append(("read_formulas", (workbook_id, sheet, address, offset, limit)))
        return _formulas()

    async def get_format(self, workbook_id: str, sheet: str, address: str) -> FormatInfo:
        self.calls.append(("get_format", (workbook_id, sheet, address)))
        return FormatInfo(range=_addr())

    async def get_validation(self, workbook_id: str, sheet: str, address: str) -> ValidationRule:
        self.calls.append(("get_validation", (workbook_id, sheet, address)))
        return ValidationRule()

    async def get_conditional_formatting(
        self, workbook_id: str, sheet: str, address: str
    ) -> list[Any]:
        self.calls.append(("get_conditional_formatting", (workbook_id, sheet, address)))
        return []

    async def write_values(
        self, workbook_id: str, sheet: str, address: str, payload: RangeWriteRequest
    ) -> RangeMutationResult:
        self.calls.append(("write_values", (workbook_id, sheet, address, payload)))
        return _range_mutation("write_values")

    async def write_formulas(
        self, workbook_id: str, sheet: str, address: str, payload: FormulaWriteRequest
    ) -> RangeMutationResult:
        self.calls.append(("write_formulas", (workbook_id, sheet, address, payload)))
        return _range_mutation("write_formulas")

    async def set_format(
        self, workbook_id: str, sheet: str, address: str, patch: RangeFormatPatch
    ) -> FormatMutationResult:
        self.calls.append(("set_format", (workbook_id, sheet, address, patch)))
        return _format_mutation()

    async def set_validation(
        self, workbook_id: str, sheet: str, address: str, spec: ValidationSpec
    ) -> ValidationRule:
        self.calls.append(("set_validation", (workbook_id, sheet, address, spec)))
        return ValidationRule(present=True, type=DataValidationType.LIST)

    async def set_conditional_formatting(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        spec: ConditionalFormatSpec,
        mode: str = "add",
    ) -> ConditionalFormatMutationResult:
        self.calls.append(
            ("set_conditional_formatting", (workbook_id, sheet, address, spec, mode))
        )
        return _cf_mutation(mode)

    async def copy_range(
        self,
        workbook_id: str,
        sheet: str,
        source: str,
        destination_sheet: str,
        destination: str,
    ) -> StructuralMutationResult:
        self.calls.append(
            ("copy_range", (workbook_id, sheet, source, destination_sheet, destination))
        )
        return _structural("copy")

    async def insert_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        self.calls.append(
            ("insert_range", (workbook_id, sheet, address, shift, allow_merged))
        )
        return _structural("insert")

    async def delete_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        self.calls.append(
            ("delete_range", (workbook_id, sheet, address, shift, allow_merged))
        )
        return _structural("delete")

    async def merge_range(
        self,
        workbook_id: str,
        sheet: str,
        address: str,
        *,
        across: bool = False,
        allow_data_loss: bool = False,
    ) -> StructuralMutationResult:
        self.calls.append(
            ("merge_range", (workbook_id, sheet, address, across, allow_data_loss))
        )
        return _structural("merge")

    async def unmerge_range(
        self, workbook_id: str, sheet: str, address: str
    ) -> StructuralMutationResult:
        self.calls.append(("unmerge_range", (workbook_id, sheet, address)))
        return _structural("unmerge")


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fake() -> Any:
    fake = FakeRangesService()
    range_tools.set_ranges_service(fake)  # type: ignore[arg-type]
    yield fake
    range_tools.set_ranges_service(None)


# ---------------------------------------------------------------------------
# range_read: routing
# ---------------------------------------------------------------------------
async def test_range_read_get_routes(_install_fake: FakeRangesService) -> None:
    await range_read("get", workbook_id="wb1", sheet="Sheet1", address="A1:C10")

    verb, args = _install_fake.calls[-1]
    assert verb == "get_range"
    assert args == ("wb1", "Sheet1", "A1:C10")


async def test_range_read_read_values_routes(_install_fake: FakeRangesService) -> None:
    await range_read(
        "read-values", workbook_id="wb1", sheet="Sheet1", address="A1", offset=0, limit=50
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "read_values"
    assert args == ("wb1", "Sheet1", "A1", 0, 50)


async def test_range_read_read_formulas_routes(_install_fake: FakeRangesService) -> None:
    await range_read("read-formulas", workbook_id="wb1", sheet="Sheet1", address="A1")

    verb, args = _install_fake.calls[-1]
    assert verb == "read_formulas"
    assert args == ("wb1", "Sheet1", "A1", 0, None)


async def test_range_read_get_format_routes(_install_fake: FakeRangesService) -> None:
    await range_read("get-format", workbook_id="wb1", sheet="Sheet1", address="A1")

    verb, args = _install_fake.calls[-1]
    assert verb == "get_format"
    assert args == ("wb1", "Sheet1", "A1")


async def test_range_read_get_validation_routes(_install_fake: FakeRangesService) -> None:
    await range_read("get-validation", workbook_id="wb1", sheet="Sheet1", address="A1")

    verb, args = _install_fake.calls[-1]
    assert verb == "get_validation"
    assert args == ("wb1", "Sheet1", "A1")


async def test_range_read_get_conditional_formatting_routes(
    _install_fake: FakeRangesService,
) -> None:
    await range_read(
        "get-conditional-formatting", workbook_id="wb1", sheet="Sheet1", address="A1"
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "get_conditional_formatting"
    assert args == ("wb1", "Sheet1", "A1")


# ---------------------------------------------------------------------------
# range_read: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_range_read_get_missing_address_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_read("get", workbook_id="wb1", sheet="Sheet1")


async def test_range_read_read_values_missing_address_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_read("read-values", workbook_id="wb1", sheet="Sheet1")


async def test_range_read_get_format_missing_address_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_read("get-format", workbook_id="wb1", sheet="Sheet1")


async def test_range_read_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_read("get", sheet="Sheet1", address="A1")


async def test_range_read_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_read("destroy", workbook_id="wb1", sheet="Sheet1", address="A1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# range_write: routing
# ---------------------------------------------------------------------------
async def test_range_write_write_values_routes(_install_fake: FakeRangesService) -> None:
    payload = _write_payload()
    await range_write(
        "write-values",
        workbook_id="wb1",
        sheet="Sheet1",
        address="A1",
        payload=payload,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "write_values"
    assert args == ("wb1", "Sheet1", "A1", payload)


async def test_range_write_write_formulas_routes(_install_fake: FakeRangesService) -> None:
    payload = _formula_payload()
    await range_write(
        "write-formulas",
        workbook_id="wb1",
        sheet="Sheet1",
        address="A1",
        payload=payload,
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "write_formulas"
    assert args == ("wb1", "Sheet1", "A1", payload)


async def test_range_write_set_format_routes(_install_fake: FakeRangesService) -> None:
    patch = RangeFormatPatch(font_bold=True)
    await range_write(
        "set-format", workbook_id="wb1", sheet="Sheet1", address="A1", patch=patch
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_format"
    assert args == ("wb1", "Sheet1", "A1", patch)


async def test_range_write_set_validation_routes(_install_fake: FakeRangesService) -> None:
    spec = ValidationSpec(type=DataValidationType.LIST, formula1='"A,B,C"')
    await range_write(
        "set-validation", workbook_id="wb1", sheet="Sheet1", address="A1", spec=spec
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_validation"
    assert args == ("wb1", "Sheet1", "A1", spec)


async def test_range_write_set_conditional_formatting_routes(
    _install_fake: FakeRangesService,
) -> None:
    spec = ConditionalFormatSpec(type=ConditionalFormatRequestType.CELL_VALUE)
    await range_write(
        "set-conditional-formatting",
        workbook_id="wb1",
        sheet="Sheet1",
        address="A1",
        spec=spec,
        mode="add",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "set_conditional_formatting"
    assert args == ("wb1", "Sheet1", "A1", spec, "add")


async def test_range_write_set_conditional_formatting_default_mode_is_add(
    _install_fake: FakeRangesService,
) -> None:
    spec = ConditionalFormatSpec(type=ConditionalFormatRequestType.CELL_VALUE)
    await range_write(
        "set-conditional-formatting",
        workbook_id="wb1",
        sheet="Sheet1",
        address="A1",
        spec=spec,
    )

    _, args = _install_fake.calls[-1]
    assert args[-1] == "add"


# ---------------------------------------------------------------------------
# range_write: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_range_write_write_values_missing_payload_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "write-values", workbook_id="wb1", sheet="Sheet1", address="A1"
        )


async def test_range_write_write_formulas_missing_payload_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "write-formulas", workbook_id="wb1", sheet="Sheet1", address="A1"
        )


async def test_range_write_set_format_missing_patch_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "set-format", workbook_id="wb1", sheet="Sheet1", address="A1"
        )


async def test_range_write_set_format_missing_address_errors() -> None:
    # patch is present, so the missing address is caught by the target handler.
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "set-format",
            workbook_id="wb1",
            sheet="Sheet1",
            patch=RangeFormatPatch(font_bold=True),
        )


async def test_range_write_set_validation_missing_spec_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "set-validation", workbook_id="wb1", sheet="Sheet1", address="A1"
        )


async def test_range_write_set_conditional_formatting_missing_spec_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "set-conditional-formatting",
            workbook_id="wb1",
            sheet="Sheet1",
            address="A1",
        )


async def test_range_write_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write(
            "write-values", sheet="Sheet1", address="A1", payload=_write_payload()
        )


async def test_range_write_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_write("destroy", workbook_id="wb1", sheet="Sheet1", address="A1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# range_structural: routing
# ---------------------------------------------------------------------------
async def test_range_structural_copy_routes(_install_fake: FakeRangesService) -> None:
    await range_structural(
        "copy",
        workbook_id="wb1",
        sheet="Sheet1",
        source="A1",
        destination_sheet="Sheet2",
        destination="B2",
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "copy_range"
    assert args == ("wb1", "Sheet1", "A1", "Sheet2", "B2")


async def test_range_structural_insert_routes(_install_fake: FakeRangesService) -> None:
    await range_structural(
        "insert", workbook_id="wb1", sheet="Sheet1", address="A1", shift="down"
    )

    verb, args = _install_fake.calls[-1]
    assert verb == "insert_range"
    assert args == ("wb1", "Sheet1", "A1", "down", False)


async def test_range_structural_unmerge_routes(_install_fake: FakeRangesService) -> None:
    await range_structural("unmerge", workbook_id="wb1", sheet="Sheet1", address="A1")

    verb, args = _install_fake.calls[-1]
    assert verb == "unmerge_range"
    assert args == ("wb1", "Sheet1", "A1")


# ---------------------------------------------------------------------------
# range_structural: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_range_structural_insert_missing_shift_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_structural(
            "insert", workbook_id="wb1", sheet="Sheet1", address="A1"
        )


async def test_range_structural_insert_disallowed_shift_errors() -> None:
    # 'left' is a delete shift, not an insert shift -> rejected (never guessed).
    with pytest.raises(ExcelVbaMcpError):
        await range_structural(
            "insert", workbook_id="wb1", sheet="Sheet1", address="A1", shift="left"
        )


async def test_range_structural_copy_missing_source_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_structural(
            "copy",
            workbook_id="wb1",
            sheet="Sheet1",
            destination_sheet="Sheet2",
            destination="B2",
        )


async def test_range_structural_copy_missing_destination_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_structural(
            "copy",
            workbook_id="wb1",
            sheet="Sheet1",
            source="A1",
            destination_sheet="Sheet2",
        )


async def test_range_structural_unmerge_missing_address_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_structural("unmerge", workbook_id="wb1", sheet="Sheet1")


async def test_range_structural_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await range_structural("destroy", workbook_id="wb1", sheet="Sheet1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# range_delete: standalone, shift in {left, up}
# ---------------------------------------------------------------------------
async def test_range_delete_left_routes(_install_fake: FakeRangesService) -> None:
    await range_delete("wb1", "Sheet1", "A1", "left")

    verb, args = _install_fake.calls[-1]
    assert verb == "delete_range"
    assert args == ("wb1", "Sheet1", "A1", "left", False)


async def test_range_delete_up_routes(_install_fake: FakeRangesService) -> None:
    await range_delete("wb1", "Sheet1", "A1", "up")

    verb, args = _install_fake.calls[-1]
    assert verb == "delete_range"
    assert args == ("wb1", "Sheet1", "A1", "up", False)


async def test_range_delete_missing_shift_errors() -> None:
    # An empty shift is treated as omitted and rejected (never guessed).
    with pytest.raises(ExcelVbaMcpError):
        await range_delete("wb1", "Sheet1", "A1", "")


async def test_range_delete_disallowed_shift_errors() -> None:
    # 'right' is an insert shift, not a delete shift -> rejected.
    with pytest.raises(ExcelVbaMcpError):
        await range_delete("wb1", "Sheet1", "A1", "right")


# ---------------------------------------------------------------------------
# range_merge: standalone (dual-hint)
# ---------------------------------------------------------------------------
async def test_range_merge_routes(_install_fake: FakeRangesService) -> None:
    await range_merge("wb1", "Sheet1", "A1:B2")

    verb, args = _install_fake.calls[-1]
    assert verb == "merge_range"
    assert args == ("wb1", "Sheet1", "A1:B2", False, False)


# ---------------------------------------------------------------------------
# Registration + annotation triples (WP0 truth matrix)
# ---------------------------------------------------------------------------
async def test_range_tools_registered_old_names_removed() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for present in (
        "range_read",
        "range_write",
        "range_structural",
        "range_delete",
        "range_merge",
    ):
        assert present in names
    for removed in (
        "range_get",
        "range_read_values",
        "range_read_formulas",
        "range_write_values",
        "range_write_formulas",
        "range_get_format",
        "range_set_format",
        "range_copy",
        "range_insert",
        "range_unmerge",
        "range_get_validation",
        "range_set_validation",
        "range_get_conditional_formatting",
        "range_set_conditional_formatting",
    ):
        assert removed not in names


async def test_range_read_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["range_read"].annotations
    assert ann is not None
    # (T, F, T): read-only and idempotent.
    assert ann.read_only_hint is True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_range_write_annotation_triple_not_idempotent() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["range_write"].annotations
    assert ann is not None
    # (F, F, F): not read-only, not destructive, NOT idempotent
    # (set-conditional-formatting mode='add' appends a rule -> R-Q6).
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is not True


async def test_range_structural_annotation_triple_not_idempotent() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["range_structural"].annotations
    assert ann is not None
    # (F, T, F): destructive and NOT idempotent (H4).
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is not True


async def test_range_delete_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["range_delete"].annotations
    assert ann is not None
    # (F, T, F): destructive, not idempotent.
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is not True


async def test_range_merge_annotation_dual_hint() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["range_merge"].annotations
    assert ann is not None
    # (F, T, T): intentional dual-hint (destructive + idempotent).
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is True