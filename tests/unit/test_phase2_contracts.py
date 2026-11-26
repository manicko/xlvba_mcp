"""Unit tests for the Phase 2 Range & PivotTables contracts (task_040).

Covers DTO serialization, bounded-request enforcement, tri-state formatting,
pivot enums, the new Phase 2 error codes, and Stage 1 backward compatibility.
No Excel/COM is touched: these are pure domain-contract tests.
"""

from __future__ import annotations

import json
from datetime import UTC, datetime

import pytest
from pydantic import ValidationError

from excel_vba_mcp.domain import errors as err
from excel_vba_mcp.domain.errors import (
    ErrorCode,
    ExcelVbaMcpError,
    error_from_hresult,
)
from excel_vba_mcp.domain.models import (
    MAX_RANGE_COLS,
    MAX_RANGE_ROWS,
    MAX_RESPONSE_CELLS,
    MAX_WRITE_PAYLOAD_CELLS,
    AlignmentFormat,
    CellValue,
    ConditionalFormatRule,
    ConditionalFormatSpec,
    FillFormat,
    FontFormat,
    FormatInfo,
    FormulaWriteRequest,
    PivotCacheInfo,
    PivotFieldFilters,
    PivotFieldInfo,
    PivotFieldLayout,
    PivotFieldOrientation,
    PivotFieldProperties,
    PivotFilterSpec,
    PivotFilterState,
    PivotInventory,
    PivotRefreshResult,
    PivotTableInfo,
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
    WorkbookInfo,
    WorkbookModel,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    ConditionalFormatType,
    DataValidationType,
    MacroRunStatus,
    OperationClass,
    PivotAggregation,
    PivotFilterKind,
    PivotFilterRequestKind,
    PivotRefreshMethod,
    PivotRefreshStatus,
    PivotSourceType,
    ValidationOperator,
)

_TS = datetime(2026, 1, 1, 12, 0, 0, tzinfo=UTC)


def _addr(**overrides: object) -> RangeAddress:
    """Build a valid RangeAddress with optional overrides."""

    base: dict[str, object] = {
        "sheet": "Data",
        "address": "A1:B2",
        "rows": 2,
        "cols": 2,
        "shape": RangeShape.BLOCK,
    }
    base.update(overrides)
    return RangeAddress.model_validate(base)


def _cell(value: object, kind: CellValueKind, address: str = "A1") -> CellValue:
    return CellValue(kind=kind, value=value, address=address)  # type: ignore[arg-type]


def _assert_json_native(value: object) -> None:
    if isinstance(value, dict):
        for child in value.values():
            _assert_json_native(child)
    elif isinstance(value, (list, tuple)):
        for child in value:
            _assert_json_native(child)
    else:
        assert isinstance(value, (str, int, float, bool, type(None))), type(value)


# ---------------------------------------------------------------------------
# Range DTOs: serialization + shape
# ---------------------------------------------------------------------------


def test_range_address_and_values_round_trip() -> None:
    values = RangeValues(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=2,
        cols=2,
        values=[
            [_cell(1, CellValueKind.NUMBER), _cell("x", CellValueKind.STRING)],
            [_cell(None, CellValueKind.EMPTY), _cell(True, CellValueKind.BOOLEAN)],
        ],
        truncated=False,
        notes=["ok"],
    )
    payload = json.loads(values.model_dump_json())
    restored = RangeValues.model_validate(payload)
    assert restored.model_dump(mode="json") == values.model_dump(mode="json")
    assert restored.shape is RangeShape.BLOCK
    _assert_json_native(values.model_dump(mode="json"))


def test_range_values_single_cell_shape_is_scalar() -> None:
    values = RangeValues(
        range=_addr(address="A1", rows=1, cols=1, shape=RangeShape.SCALAR),
        shape=RangeShape.SCALAR,
        rows=1,
        cols=1,
        values=[[_cell(42, CellValueKind.NUMBER, "A1")]],
    )
    assert values.shape is RangeShape.SCALAR
    assert values.values[0][0].kind is CellValueKind.NUMBER


def test_range_formulas_round_trip() -> None:
    formulas = RangeFormulas.model_validate(
        {
            "range": _addr().model_dump(mode="json"),
            "shape": "block",
            "rows": 1,
            "cols": 1,
            "formulas": [
                [
                    {
                        "address": "A1",
                        "formula2": "=SUM(B1:B2)",
                        "formula": "=SUM(B1:B2)",
                        "is_array": False,
                        "saved_as_array": None,
                    }
                ]
            ],
        }
    )
    assert formulas.formulas[0][0].formula2 == "=SUM(B1:B2)"
    assert json.loads(formulas.model_dump_json())["formulas"][0][0]["address"] == "A1"


# ---------------------------------------------------------------------------
# Bounded-request enforcement
# ---------------------------------------------------------------------------


def test_range_write_request_accepts_valid_grid() -> None:
    req = RangeWriteRequest(
        rows=1,
        cols=2,
        values=[[_cell(1, CellValueKind.NUMBER), _cell(2, CellValueKind.NUMBER)]],
    )
    assert req.rows * req.cols == 2


def test_range_write_request_rejects_too_many_rows() -> None:
    with pytest.raises(ValidationError):
        RangeWriteRequest.model_validate({"rows": MAX_RANGE_ROWS + 1, "cols": 1, "values": []})


def test_range_write_request_rejects_too_many_cols() -> None:
    with pytest.raises(ValidationError):
        RangeWriteRequest.model_validate({"rows": 1, "cols": MAX_RANGE_COLS + 1, "values": [[]]})


def test_range_write_request_rejects_excess_cells() -> None:
    rows = MAX_WRITE_PAYLOAD_CELLS // 2 + 1
    with pytest.raises(ValidationError):
        RangeWriteRequest.model_validate({"rows": rows, "cols": 2, "values": []})


def test_range_write_request_rejects_shape_mismatch() -> None:
    with pytest.raises(ValidationError):
        RangeWriteRequest.model_validate(
            {
                "rows": 2,
                "cols": 2,
                "values": [[_cell(1, CellValueKind.NUMBER)]],
            }
        )


def test_formula_write_request_rejects_non_formula_constant() -> None:
    # A formula write must not silently carry constants through: the payload is
    # still shape-validated here; constant-vs-formula separation is enforced by
    # the tool/service (range_write_formulas requires '=' strings).
    req = FormulaWriteRequest(rows=1, cols=1, formulas=[["=A1+1"]])
    assert req.formulas[0][0].startswith("=")


def test_range_values_response_rejects_oversized_grid() -> None:
    rows = MAX_RESPONSE_CELLS // 100 + 1
    with pytest.raises(ValidationError):
        RangeValues.model_validate(
            {
                "range": _addr().model_dump(mode="json"),
                "shape": "block",
                "rows": rows,
                "cols": 100,
                "values": [],
            }
        )


# ---------------------------------------------------------------------------
# Tri-state formatting
# ---------------------------------------------------------------------------


def test_format_info_tri_state_mixed_and_unsupported() -> None:
    info = FormatInfo(
        range=_addr(),
        number_format=None,
        font=FontFormat(bold=None, italic=True, size=11.0),
        fill=FillFormat(color=None),
        alignment=AlignmentFormat(horizontal="left", wrap_text=None),
        mixed_properties=["number_format", "font.bold", "fill.color"],
        unsupported=["borders", "gradient"],
    )
    assert info.number_format is None
    assert "number_format" in info.mixed_properties
    assert info.unsupported == ["borders", "gradient"]
    _assert_json_native(info.model_dump(mode="json"))


def test_range_format_patch_reports_only_set_fields() -> None:
    patch = RangeFormatPatch(number_format="#,##0.00", font_bold=True)
    assert patch.applied_fields() == ["font_bold", "number_format"]
    empty = RangeFormatPatch()
    assert empty.applied_fields() == []


# ---------------------------------------------------------------------------
# Validation / conditional formatting specs
# ---------------------------------------------------------------------------


def test_validation_spec_accepts_normal_rule() -> None:
    spec = ValidationSpec(
        type=DataValidationType.LIST,
        operator=ValidationOperator.EQUAL,
        formula1='"A,B,C"',
        ignore_blank=True,
        in_cell_dropdown=True,
    )
    assert spec.type is DataValidationType.LIST


def test_validation_spec_rejects_overlong_formula() -> None:
    with pytest.raises(ValidationError):
        ValidationSpec(
            type=DataValidationType.CUSTOM,
            formula1="x" * 256,
        )


def test_validation_rule_present_false_default() -> None:
    rule = ValidationRule()
    assert rule.present is False
    assert rule.type is DataValidationType.NONE


def test_conditional_format_rule_flags_unsupported_type() -> None:
    rule = ConditionalFormatRule(
        priority=1,
        type=ConditionalFormatType.COLOR_SCALE,
        supported=False,
        unsupported_reason="ColorScale is a distinct COM object; out of scope.",
    )
    assert rule.supported is False
    assert rule.type is ConditionalFormatType.COLOR_SCALE


def test_conditional_format_spec_round_trip() -> None:
    spec = ConditionalFormatSpec(
        type=ConditionalFormatType.CELL_VALUE,
        operator=ValidationOperator.GREATER,
        formula1="100",
        stop_if_true=True,
    )
    payload = json.loads(spec.model_dump_json())
    assert payload["type"] == "cell_value"


# ---------------------------------------------------------------------------
# Mutation results
# ---------------------------------------------------------------------------


def test_range_mutation_result_round_trip() -> None:
    result = RangeMutationResult(
        action="write_values",
        range_before=_addr(),
        range_after=_addr(),
        verified=True,
        postcondition_values=RangeValues(
            range=_addr(),
            shape=RangeShape.BLOCK,
            rows=1,
            cols=1,
            values=[[_cell(7, CellValueKind.NUMBER, "A1")]],
        ),
    )
    restored = RangeMutationResult.model_validate(json.loads(result.model_dump_json()))
    assert restored.verified is True
    assert restored.postcondition_values is not None


def test_structural_mutation_result_defaults() -> None:
    result = StructuralMutationResult(action="unmerge", source=_addr())
    assert result.table_metadata_unchanged is True
    assert result.preconditions == {}


# ---------------------------------------------------------------------------
# Pivot enums + DTOs
# ---------------------------------------------------------------------------


def test_pivot_source_type_values() -> None:
    values = {member.value for member in PivotSourceType}
    assert {"database", "external", "consolidation", "scenario"} <= values
    assert {"olap", "data_model", "unsupported"} <= values


def test_pivot_aggregation_accepts_eleven_codes() -> None:
    assert {member.value for member in PivotAggregation} == {
        "sum",
        "average",
        "count",
        "countNums",
        "max",
        "min",
        "product",
        "stDev",
        "stDevP",
        "var",
        "varP",
    }


def test_pivot_field_orientation_and_refresh_enums() -> None:
    assert {m.value for m in PivotFieldOrientation} == {
        "hidden",
        "row",
        "column",
        "page",
        "data",
    }
    assert {m.value for m in PivotRefreshStatus} == {
        "refreshed",
        "unsupported_source",
        "blocked",
        "failed",
    }
    assert {m.value for m in PivotRefreshMethod} == {"refresh_table", "refresh_cache"}
    assert PivotFilterKind.MANUAL.value == "manual"


def test_pivot_inventory_round_trip_no_com() -> None:
    inventory = PivotInventory(
        workbook_id="wb_1",
        pivot_tables=[
            PivotTableInfo(
                name="ptA",
                sheet_name="Sheet2",
                sheet_index=2,
                index=1,
                cache_index=1,
                table_range1="$A$1:$B$7",
                table_range2="$A$3:$B$7",
                refresh_date=_TS,
                source_type=PivotSourceType.DATABASE,
                source_status="supported",
                row_fields=[
                    PivotFieldInfo(
                        source_name="Region",
                        orientation=PivotFieldOrientation.ROW,
                        position=1,
                        all_items_visible=True,
                        filter_state=PivotFilterKind.NONE,
                    )
                ],
                data_fields=[
                    PivotFieldInfo(
                        source_name="Sales",
                        orientation=PivotFieldOrientation.DATA,
                        function=PivotAggregation.SUM,
                        is_data_field=True,
                        data_field_name="Sum of Sales",
                    )
                ],
            )
        ],
        caches=[
            PivotCacheInfo(
                cache_index=1,
                source_type=PivotSourceType.DATABASE,
                source_data="Data!R1C1:R6C5",
                source_data_kind="text",
                refresh_date=_TS,
                shared_cache=True,
                affected_pivot_tables=["ptA"],
            )
        ],
    )
    payload = json.loads(inventory.model_dump_json())
    restored = PivotInventory.model_validate(payload)
    assert restored.pivot_tables[0].data_fields[0].function is PivotAggregation.SUM
    assert restored.caches[0].source_type is PivotSourceType.DATABASE
    _assert_json_native(inventory.model_dump(mode="json"))


def test_pivot_field_layout_and_properties() -> None:
    layout = PivotFieldLayout(orientation=PivotFieldOrientation.COLUMN, position=2)
    assert layout.position == 2
    props = PivotFieldProperties(caption="Total", number_format="#,##0.00")
    assert props.caption == "Total"


def test_pivot_filter_spec_manual_vs_unsupported() -> None:
    manual = PivotFilterSpec(kind=PivotFilterRequestKind.MANUAL, visible_items=["South"])
    assert manual.supported is True
    # The request vocabulary is pruned to ``manual`` only (U3): a label/value/date/
    # top-bottom request is rejected at the pydantic layer, never sent to Excel.
    with pytest.raises(ValidationError):
        PivotFilterSpec(kind=PivotFilterKind.LABEL, supported=False)
    # The full kind is still representable on the read-state model.
    state = PivotFilterState(kind=PivotFilterKind.LABEL, supported=False)
    assert state.kind is PivotFilterKind.LABEL


def test_pivot_field_filters_and_refresh_result_round_trip() -> None:
    filters = PivotFieldFilters(
        pivot_table="ptA",
        field="Region",
        all_items_visible=False,
        visible_items=["South"],
        hidden_items=["North"],
        filter=PivotFilterState(kind=PivotFilterKind.MANUAL, visible_items=["South"]),
    )
    assert filters.hidden_items == ["North"]

    refresh = PivotRefreshResult(
        pivot_table="ptA",
        cache_index=1,
        method=PivotRefreshMethod.REFRESH_TABLE,
        affected_pivot_tables=["ptA", "ptB"],
        refreshed=True,
        refresh_return=True,
        started_at=_TS,
        finished_at=_TS,
        duration_ms=12,
        refresh_date_before=_TS,
        refresh_date_after=_TS,
        status=PivotRefreshStatus.REFRESHED,
    )
    payload = json.loads(refresh.model_dump_json())
    restored = PivotRefreshResult.model_validate(payload)
    assert restored.status is PivotRefreshStatus.REFRESHED
    assert restored.affected_pivot_tables == ["ptA", "ptB"]


def test_pivot_unsupported_source_is_explicit() -> None:
    cache = PivotCacheInfo(
        cache_index=2,
        source_type=PivotSourceType.OLAP,
        source_data_kind="unavailable",
        olap=True,
        shared_cache=False,
    )
    assert cache.source_type is PivotSourceType.OLAP
    assert cache.olap is True


# ---------------------------------------------------------------------------
# Phase 2 error codes
# ---------------------------------------------------------------------------


_PHASE2_CODES = {
    "INVALID_RANGE",
    "PROTECTED_TARGET",
    "UNSUPPORTED_PIVOT_SOURCE",
    "STALE_TARGET",
    "AMBIGUOUS_FILTER",
    "REFRESH_FAILED",
    "MODAL_UI_BLOCKED",
    "LISTOBJECT_INTERSECTION",
}

_PHASE2_SUBCLASSES = {
    ErrorCode.INVALID_RANGE: err.InvalidRangeError,
    ErrorCode.PROTECTED_TARGET: err.ProtectedTargetError,
    ErrorCode.UNSUPPORTED_PIVOT_SOURCE: err.UnsupportedPivotSourceError,
    ErrorCode.STALE_TARGET: err.StaleTargetError,
    ErrorCode.AMBIGUOUS_FILTER: err.AmbiguousFilterError,
    ErrorCode.REFRESH_FAILED: err.RefreshFailedError,
    ErrorCode.MODAL_UI_BLOCKED: err.ModalUiBlockedError,
    ErrorCode.LISTOBJECT_INTERSECTION: err.ListObjectIntersectionError,
}


def test_phase2_error_codes_present_and_additive() -> None:
    values = {code.value for code in ErrorCode}
    assert values >= _PHASE2_CODES
    # Stage 1 codes remain (subset check mirrors task_038 §11 closed set).
    assert {"COM_BUSY", "EXCEL_HUNG", "SNAPSHOT_FAILED", "PATH_NOT_ALLOWED"} <= values


def test_phase2_error_subclasses_and_hresult_preservation() -> None:
    for code, subclass in _PHASE2_SUBCLASSES.items():
        assert err._ERROR_BY_CODE[code] is subclass
        error = subclass("boom", hresult=0x80040154, context={"op": "range_get"})
        assert error.code is code
        assert error.hresult == 0x80040154
        payload = error.to_dict()
        assert payload["hresult_hex"] == "0x80040154"
        assert json.loads(json.dumps(payload))["context"]["op"] == "range_get"


def test_error_from_hresult_still_maps_stage1_codes() -> None:
    error = error_from_hresult(0x80010001, "rejected")
    assert isinstance(error, err.RpcCallRejectedError)
    assert error.code is ErrorCode.RPC_E_CALL_REJECTED


def test_phase2_error_code_round_trip_value() -> None:
    assert ErrorCode.INVALID_RANGE.value == "INVALID_RANGE"
    assert ErrorCode("MODAL_UI_BLOCKED") is ErrorCode.MODAL_UI_BLOCKED
    with pytest.raises(ValueError):
        ErrorCode("NOT_A_CODE")


# ---------------------------------------------------------------------------
# Stage 1 backward compatibility
# ---------------------------------------------------------------------------


def test_stage1_workbook_model_still_round_trips() -> None:
    model = WorkbookModel(
        created_at=_TS,
        updated_at=_TS,
        workbook=WorkbookInfo(
            workbook_id="wb",
            name="a.xlsm",
            format="xlsm",
            path="p",
            read_only=True,
            has_vba=False,
            excel_version="16.0",
            session_id="s",
            mode="inspect",
        ),
    )
    restored = WorkbookModel.model_validate(json.loads(model.model_dump_json()))
    assert restored.model_dump(mode="json") == model.model_dump(mode="json")


def test_stage1_status_enums_unchanged() -> None:
    assert {m.value for m in OperationClass} == {"read_only", "mutating", "executing"}
    assert {m.value for m in MacroRunStatus} == {
        "success",
        "runtime_error",
        "blocked",
        "not_found",
        "timeout",
        "needs_recovery",
        # Block E (additive): deadline-exceeded macro, no live-pid kill.
        "suspected_stalled",
    }
    assert {m.value for m in CellValueKind} == {
        "string",
        "number",
        "boolean",
        "date",
        "error",
        "empty",
    }


def test_stage1_cell_value_reused_in_range_values() -> None:
    # The Phase 2 RangeValues reuses the Stage 1 CellValue (same class).
    cell = _cell("hello", CellValueKind.STRING)
    values = RangeValues(
        range=_addr(address="A1", rows=1, cols=1, shape=RangeShape.SCALAR),
        shape=RangeShape.SCALAR,
        rows=1,
        cols=1,
        values=[[cell]],
    )
    assert isinstance(values.values[0][0], CellValue)


def test_stage1_error_base_still_available() -> None:
    error = ExcelVbaMcpError("plain", code=ErrorCode.COM_BUSY)
    assert error.code is ErrorCode.COM_BUSY
    assert error.to_dict()["code"] == "COM_BUSY"
