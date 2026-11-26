"""Unit tests for the Phase 2 bounded test-spec schema (task_053).

The declarative suite schema (``excel_vba_mcp.domain.test_spec``) is the single
gate that rejects arbitrary code before any Excel call. These tests cover the
Range/Cells and PivotTables actions and assertions added by task_053:

* every new Range/Pivot action and assertion is a documented, accepted member;
* valid Range/Pivot step and assertion documents validate and round-trip through
  ``model_dump`` / ``model_validate``;
* an unknown Range/Pivot action or assertion is rejected with ``extra="forbid"``
  semantics **before** any Excel call;
* non-scalar / arbitrary-code payloads (a shell string, a COM-ish dict, a bare
  string where a typed grid is required) are rejected;
* the declarative bounds are enforced (payload rows, filter item count, negative
  offsets / row-column expectations).

These are pure schema tests: no Excel, COM, worker or service is involved.
"""

from __future__ import annotations

from typing import Any

import pytest
from pydantic import ValidationError

from excel_vba_mcp.domain.models import (
    CellValue,
    FormulaWriteRequest,
    PivotFieldProperties,
    PivotFilterSpec,
    RangeFormatPatch,
    RangeWriteRequest,
)
from excel_vba_mcp.domain.statuses import (
    ConditionalFormatType,
    DataValidationType,
    PivotAggregation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotFilterRequestKind,
    ValidationOperator,
)
from excel_vba_mcp.domain.test_spec import (
    MAX_PIVOT_FILTER_ITEMS,
    Assertion,
    AssertionType,
    TestAction,
    TestSuiteSpec,
    validate_action,
    validate_assertion,
)


def _suite(cases: list[dict[str, Any]]) -> TestSuiteSpec:
    """Build a validated suite around one or more case documents."""

    return TestSuiteSpec.model_validate(
        {"suite": "phase2-schema", "workbook_id": "wb", "cases": cases}
    )


# ---------------------------------------------------------------------------
# The closed vocabularies contain the Phase 2 members
# ---------------------------------------------------------------------------
def test_phase2_actions_are_documented_members() -> None:
    expected = {
        "range_get",
        "range_read_values",
        "range_read_formulas",
        "range_get_format",
        "range_write_values",
        "range_write_formulas",
        "range_set_format",
        "range_copy",
        "range_insert",
        "range_delete",
        "range_merge",
        "range_unmerge",
        "pivot_table_refresh",
        "pivot_field_set_layout",
        "pivot_field_set_aggregation",
        "pivot_field_apply_filter",
    }
    values = {member.value for member in TestAction}
    assert expected <= values


def test_phase2_assertions_are_documented_members() -> None:
    expected = {
        "cell_value_equals",
        "range_formula_equals",
        "range_format_equals",
        "range_structure",
        "validation_exists",
        "conditional_format_exists",
        "pivot_table_exists",
        "pivot_field_layout",
        "pivot_field_aggregation",
        "pivot_filter_state",
        "pivot_refresh_status",
        "pivot_output_range",
    }
    values = {member.value for member in AssertionType}
    assert expected <= values


def test_validate_action_and_assertion_roundtrip() -> None:
    assert validate_action("range_write_values") is TestAction.RANGE_WRITE_VALUES
    assert validate_action("pivot_table_refresh") is TestAction.PIVOT_TABLE_REFRESH
    assert validate_assertion("range_structure") is AssertionType.RANGE_STRUCTURE
    assert validate_assertion("pivot_refresh_status") is AssertionType.PIVOT_REFRESH_STATUS


@pytest.mark.parametrize(
    "action",
    [
        "range_execute",
        "range_shell",
        "pivot_delete",
        "pivot_create",
        "os_system",
        "eval_python",
    ],
)
def test_unknown_action_is_rejected(action: str) -> None:
    with pytest.raises(ValueError):
        validate_action(action)


@pytest.mark.parametrize(
    "kind",
    ["range_value_equals_typo", "pivot_field_orientation", "shell_exec", "python_eval"],
)
def test_unknown_assertion_is_rejected(kind: str) -> None:
    with pytest.raises(ValueError):
        validate_assertion(kind)


# ---------------------------------------------------------------------------
# Valid Range step documents
# ---------------------------------------------------------------------------
def test_range_write_values_step_validates() -> None:
    suite = _suite(
        [
            {
                "id": "write-values",
                "steps": [
                    {
                        "action": "range_write_values",
                        "sheet": "Cells",
                        "range_address": "A1:B2",
                        "values_grid": {
                            "rows": 2,
                            "cols": 2,
                            "values": [
                                [
                                    {"kind": "number", "value": 1},
                                    {"kind": "number", "value": 2},
                                ],
                                [
                                    {"kind": "string", "value": "x"},
                                    {"kind": "empty"},
                                ],
                            ],
                        },
                    }
                ],
                "assertions": [
                    {"type": "cell_value_equals", "sheet": "Cells", "cell": "A1", "expected": 1}
                ],
            }
        ]
    )
    step = suite.cases[0].steps[0]
    assert step.action is TestAction.RANGE_WRITE_VALUES
    assert step.values_grid is not None
    assert step.values_grid.rows == 2
    assert step.values_grid.values[0][0].kind == "number"


def test_range_write_formulas_step_validates() -> None:
    suite = _suite(
        [
            {
                "id": "write-formulas",
                "steps": [
                    {
                        "action": "range_write_formulas",
                        "sheet": "Cells",
                        "range_address": "C1",
                        "formulas_grid": {"rows": 1, "cols": 1, "formulas": [["=1+1"]]},
                    }
                ],
                "assertions": [
                    {
                        "type": "range_formula_equals",
                        "sheet": "Cells",
                        "address": "C1",
                        "expected": ["=1+1"],
                    }
                ],
            }
        ]
    )
    step = suite.cases[0].steps[0]
    assert step.formulas_grid is not None
    assert step.formulas_grid.formulas[0][0] == "=1+1"


def test_range_structural_steps_validate() -> None:
    suite = _suite(
        [
            {
                "id": "structural",
                "steps": [
                    {
                        "action": "range_copy",
                        "sheet": "S",
                        "source_range": "A1:B2",
                        "destination_sheet": "S",
                        "destination": "D1:E2",
                    },
                    {
                        "action": "range_insert",
                        "sheet": "S",
                        "range_address": "A1",
                        "shift": "down",
                    },
                    {"action": "range_delete", "sheet": "S", "range_address": "A1", "shift": "up"},
                    {
                        "action": "range_merge",
                        "sheet": "S",
                        "range_address": "A1:B2",
                        "across": False,
                        "allow_data_loss": True,
                    },
                    {"action": "range_unmerge", "sheet": "S", "range_address": "A1:B2"},
                ],
                "assertions": [],
            }
        ]
    )
    kinds = [step.action for step in suite.cases[0].steps]
    assert kinds == [
        TestAction.RANGE_COPY,
        TestAction.RANGE_INSERT,
        TestAction.RANGE_DELETE,
        TestAction.RANGE_MERGE,
        TestAction.RANGE_UNMERGE,
    ]
    insert = suite.cases[0].steps[1]
    assert insert.shift == "down"


def test_range_format_and_read_steps_validate() -> None:
    suite = _suite(
        [
            {
                "id": "format-and-read",
                "steps": [
                    {"action": "range_get", "sheet": "S", "range_address": "A1:B2"},
                    {
                        "action": "range_read_values",
                        "sheet": "S",
                        "range_address": "A1:B2",
                        "offset": 0,
                        "limit": 5,
                    },
                    {"action": "range_read_formulas", "sheet": "S", "range_address": "A1:B2"},
                    {"action": "range_get_format", "sheet": "S", "range_address": "A1:B2"},
                    {
                        "action": "range_set_format",
                        "sheet": "S",
                        "range_address": "A1:B2",
                        "format_patch": {"font_bold": True, "number_format": "0.00"},
                    },
                ],
                "assertions": [
                    {
                        "type": "range_format_equals",
                        "sheet": "S",
                        "address": "A1",
                        "property": "font.bold",
                        "expected": True,
                    }
                ],
            }
        ]
    )
    patch = suite.cases[0].steps[4].format_patch
    assert isinstance(patch, RangeFormatPatch)
    assert patch.applied_fields() == ["font_bold", "number_format"]


# ---------------------------------------------------------------------------
# Valid Pivot step / assertion documents
# ---------------------------------------------------------------------------
def test_pivot_steps_validate() -> None:
    suite = _suite(
        [
            {
                "id": "pivots",
                "steps": [
                    {"action": "pivot_table_refresh", "pivot_table": "ptA"},
                    {
                        "action": "pivot_field_set_layout",
                        "pivot_table": "ptA",
                        "field": "Region",
                        "orientation": "row",
                    },
                    {
                        "action": "pivot_field_set_aggregation",
                        "pivot_table": "ptA",
                        "field": "Sales",
                        "function": "sum",
                        "caption": "Sum of Sales",
                    },
                    {
                        "action": "pivot_field_apply_filter",
                        "pivot_table": "ptB",
                        "field": "Region",
                        "filter_spec": {"kind": "manual", "visible_items": ["East", "North"]},
                    },
                ],
                "assertions": [
                    {"type": "pivot_table_exists", "pivot_table": "ptA"},
                    {
                        "type": "pivot_field_layout",
                        "pivot_table": "ptA",
                        "field": "Region",
                        "orientation": "row",
                        "position": 1,
                    },
                    {
                        "type": "pivot_field_aggregation",
                        "pivot_table": "ptA",
                        "field": "Sales",
                        "function": "sum",
                    },
                    {
                        "type": "pivot_filter_state",
                        "pivot_table": "ptB",
                        "field": "Region",
                        "hidden_items": ["West"],
                    },
                    {"type": "pivot_refresh_status", "pivot_table": "ptA"},
                    {"type": "pivot_output_range", "pivot_table": "ptA", "output_range": "A3:F14"},
                ],
            }
        ]
    )
    layout = suite.cases[0].steps[1]
    assert layout.orientation is PivotFieldOrientation.ROW
    aggregation = suite.cases[0].steps[2]
    assert aggregation.function is PivotAggregation.SUM
    filt = suite.cases[0].steps[3]
    assert isinstance(filt.filter_spec, PivotFilterSpec)
    assert filt.filter_spec.kind is PivotFilterRequestKind.MANUAL


# ---------------------------------------------------------------------------
# Serialization round-trips
# ---------------------------------------------------------------------------
def test_suite_serialization_roundtrip() -> None:
    suite = _suite(
        [
            {
                "id": "roundtrip",
                "steps": [
                    {
                        "action": "range_write_values",
                        "sheet": "S",
                        "range_address": "A1",
                        "values_grid": {
                            "rows": 1,
                            "cols": 1,
                            "values": [[{"kind": "number", "value": 7}]],
                        },
                    },
                    {"action": "pivot_table_refresh", "pivot_table": "ptA"},
                ],
                "assertions": [
                    {"type": "cell_value_equals", "sheet": "S", "cell": "A1", "expected": 7},
                    {"type": "pivot_refresh_status", "pivot_table": "ptA", "expected": "refreshed"},
                ],
            }
        ]
    )
    dumped = suite.model_dump(mode="json")
    rebuilt = TestSuiteSpec.model_validate(dumped)
    assert rebuilt == suite
    assert rebuilt.cases[0].steps[0].action is TestAction.RANGE_WRITE_VALUES


# ---------------------------------------------------------------------------
# Rejections: unknown actions / assertions and arbitrary payloads
# ---------------------------------------------------------------------------
def test_unknown_range_action_in_a_suite_is_rejected() -> None:
    with pytest.raises(ValidationError):
        _suite(
            [{"id": "c", "steps": [{"action": "range_execute", "sheet": "S"}], "assertions": []}]
        )


def test_unknown_pivot_action_in_a_suite_is_rejected() -> None:
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [{"action": "pivot_create", "pivot_table": "pt"}],
                    "assertions": [],
                }
            ]
        )


def test_unknown_assertion_in_a_suite_is_rejected() -> None:
    with pytest.raises(ValidationError):
        _suite([{"id": "c", "steps": [], "assertions": [{"type": "range_shell_exec"}]}])


@pytest.mark.parametrize(
    "step",
    [
        {"action": "range_write_values", "sheet": "S", "range_address": "A1", "values_grid": "A1"},
        {
            "action": "range_write_values",
            "sheet": "S",
            "range_address": "A1",
            "values_grid": {"rows": 1, "cols": 1, "values": [["=1+1"]]},
            "code": "Shell 'calc'",
        },
        {
            "action": "pivot_field_apply_filter",
            "pivot_table": "pt",
            "field": "f",
            "filter_spec": {"kind": "manual", "visible_items": ["a"]},
            "evaluate": "1+1",
        },
    ],
)
def test_arbitrary_or_extra_payloads_are_rejected(step: dict[str, Any]) -> None:
    with pytest.raises(ValidationError):
        _suite([{"id": "c", "steps": [step], "assertions": []}])


def test_non_scalar_grid_cells_are_rejected() -> None:
    """A grid cell must be a typed :class:`CellValue`, never an arbitrary object."""

    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "S",
                            "range_address": "A1",
                            "values_grid": {
                                "rows": 1,
                                "cols": 1,
                                "values": [[{"kind": "number", "value": 1, "com": "x"}]],
                            },
                        }
                    ],
                    "assertions": [],
                }
            ]
        )


def test_extra_step_field_is_rejected() -> None:
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [
                        {"action": "range_get", "sheet": "S", "range_address": "A1", "shell": "ls"}
                    ],
                    "assertions": [],
                }
            ]
        )


# ---------------------------------------------------------------------------
# Bounds
# ---------------------------------------------------------------------------
def test_range_payload_row_bound_is_enforced() -> None:
    """A declarative write grid above MAX_RANGE_WRITE_ROWS is rejected pre-Excel."""

    from excel_vba_mcp.domain.test_spec import MAX_RANGE_WRITE_ROWS

    rows = MAX_RANGE_WRITE_ROWS + 1
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [
                        {
                            "action": "range_write_values",
                            "sheet": "S",
                            "range_address": "A1",
                            "values_grid": {
                                "rows": rows,
                                "cols": 1,
                                "values": [[{"kind": "number", "value": 1}] for _ in range(rows)],
                            },
                        }
                    ],
                    "assertions": [],
                }
            ]
        )


def test_pivot_filter_item_bound_is_enforced() -> None:
    too_many = ["item"] * (MAX_PIVOT_FILTER_ITEMS + 1)
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [
                        {
                            "action": "pivot_field_apply_filter",
                            "pivot_table": "pt",
                            "field": "f",
                            "filter_spec": {"kind": "manual", "visible_items": too_many},
                        }
                    ],
                    "assertions": [],
                }
            ]
        )


def test_negative_offset_or_bound_is_rejected() -> None:
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [
                        {
                            "action": "range_read_values",
                            "sheet": "S",
                            "range_address": "A1",
                            "offset": -1,
                        }
                    ],
                    "assertions": [],
                }
            ]
        )
    with pytest.raises(ValidationError):
        _suite(
            [
                {
                    "id": "c",
                    "steps": [],
                    "assertions": [
                        {"type": "range_structure", "sheet": "S", "address": "A1", "rows": -1}
                    ],
                }
            ]
        )


def test_empty_suite_is_rejected() -> None:
    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate({"suite": "s", "workbook_id": "w", "cases": []})


# ---------------------------------------------------------------------------
# Direct model construction (typed DTOs accepted as-is)
# ---------------------------------------------------------------------------
def test_typed_dtos_are_accepted_directly() -> None:
    step = Assertion(
        type=AssertionType.RANGE_STRUCTURE, sheet="S", address="A1:B2", rows=2, cols=2, merged=True
    )
    assert step.type is AssertionType.RANGE_STRUCTURE
    grid = RangeWriteRequest(rows=1, cols=1, values=[[CellValue(kind="number", value=1)]])
    assert grid.rows == 1
    formulas = FormulaWriteRequest(rows=1, cols=1, formulas=[["=NOW()"]])
    assert formulas.formulas[0][0] == "=NOW()"
    props = PivotFieldProperties(caption="Sum of Sales")
    assert props.caption == "Sum of Sales"
    spec = PivotFilterSpec(kind=PivotFilterKind.MANUAL, visible_items=["East"])
    assert spec.visible_items == ["East"]
    # The enum vocabularies referenced by the schema stay bounded.
    assert ConditionalFormatType.CELL_VALUE.value == "cell_value"
    assert DataValidationType.LIST.value == "list"
    assert ValidationOperator.EQUAL.value == "equal"
