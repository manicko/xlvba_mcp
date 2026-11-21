"""Bounded declarative test-suite schema (spec §9.2–9.3, Appendix B).

The suite format is declarative and closed: actions and assertions are modelled
as closed :class:`enum.StrEnum` vocabularies with ``extra="forbid"`` step and
assertion models, so arbitrary Python / shell / COM code cannot be expressed.
Unknown action or assertion kinds are rejected at validation time.

This module defines the schema only. It does not execute anything; the runner
(task_035, extended by task_053) consumes these models.

Phase 2 (task_053) adds bounded Range/Cells and PivotTables actions and
assertions. Every new parameter is a **bounded scalar or a typed domain DTO**:
a step still cannot carry executable code, a shell command, a COM expression or
an unbounded cell matrix. Pivot specs are validated as real domain models so an
unknown filter kind or an oversized item list is rejected here, before any
Excel call.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from excel_vba_mcp.domain.models import (
    FormulaWriteRequest,
    PivotFilterSpec,
    RangeFormatPatch,
    RangeWriteRequest,
)
from excel_vba_mcp.domain.statuses import BaselinePolicy, PivotAggregation, PivotFieldOrientation

# Maximum number of positional arguments accepted by a ``macro_run`` step, kept
# here (not in steps.py) to mirror the Excel Application.Run 30-argument limit.
MAX_MACRO_ARGS = 30

# Maximum number of ``visible_items`` accepted by a declarative pivot filter
# step (task_039 §6). Excel itself caps a field's item count (32,768), so this
# is a deliberately conservative declarative ceiling.
MAX_PIVOT_FILTER_ITEMS = 1000

# Maximum number of typed rows accepted by a declarative ``range_write_values``
# step. Excel's own ceilings are enforced by :class:`RangeWriteRequest`; this
# bound keeps a declarative suite document small enough to validate cheaply.
MAX_RANGE_WRITE_ROWS = 10_000


class TestAction(StrEnum):
    """Closed set of bounded test steps (spec §9.2, §6.5)."""

    __test__ = False  # not a pytest test class despite the "Test" prefix

    SHEET_ADD = "sheet_add"
    SHEET_RENAME = "sheet_rename"
    SHEET_COPY = "sheet_copy"
    SHEET_DELETE = "sheet_delete"
    SHEET_SET_VISIBILITY = "sheet_set_visibility"
    TABLE_CREATE = "table_create"
    TABLE_APPEND_ROWS = "table_append_rows"
    TABLE_INSERT_ROW = "table_insert_row"
    TABLE_UPDATE_ROWS = "table_update_rows"
    TABLE_DELETE_ROWS = "table_delete_rows"
    TABLE_ADD_COLUMN = "table_add_column"
    TABLE_DELETE_COLUMN = "table_delete_column"
    TABLE_RESIZE = "table_resize"
    NAME_CREATE = "name_create"
    NAME_UPDATE = "name_update"
    NAME_DELETE = "name_delete"
    SHAPE_SET_ON_ACTION = "shape_set_on_action"
    VBA_REPLACE_COMPONENT_SOURCE = "vba_replace_component_source"
    VBA_INSERT_CODE = "vba_insert_code"
    VBA_COMPILE = "vba_compile"
    MACRO_RUN = "macro_run"
    # Phase 2 — Range/Cells (task_053).
    RANGE_GET = "range_get"
    RANGE_READ_VALUES = "range_read_values"
    RANGE_READ_FORMULAS = "range_read_formulas"
    RANGE_GET_FORMAT = "range_get_format"
    RANGE_WRITE_VALUES = "range_write_values"
    RANGE_WRITE_FORMULAS = "range_write_formulas"
    RANGE_SET_FORMAT = "range_set_format"
    RANGE_COPY = "range_copy"
    RANGE_INSERT = "range_insert"
    RANGE_DELETE = "range_delete"
    RANGE_MERGE = "range_merge"
    RANGE_UNMERGE = "range_unmerge"
    # Phase 2 — PivotTables (task_053).
    PIVOT_TABLE_REFRESH = "pivot_table_refresh"
    PIVOT_FIELD_SET_LAYOUT = "pivot_field_set_layout"
    PIVOT_FIELD_SET_AGGREGATION = "pivot_field_set_aggregation"
    PIVOT_FIELD_APPLY_FILTER = "pivot_field_apply_filter"


class AssertionType(StrEnum):
    """Closed set of bounded read-only assertions (spec §9.3)."""

    COMPILE_STATUS = "compile_status"
    MACRO_STATUS = "macro_status"
    MACRO_RETURN_VALUE = "macro_return_value"
    MACRO_DURATION_LTE = "macro_duration_lte"
    SHEET_EXISTS = "sheet_exists"
    SHEET_VISIBILITY = "sheet_visibility"
    TABLE_EXISTS = "table_exists"
    TABLE_ROW_COUNT_GTE = "table_row_count_gte"
    TABLE_ROW_COUNT = "table_row_count"
    TABLE_COLUMN_NAMES = "table_column_names"
    TABLE_COLUMN_COUNT = "table_column_count"
    NAME_EXISTS = "name_exists"
    NAME_REFERS_TO = "name_refers_to"
    CELL_VALUE_EQUALS = "cell_value_equals"
    READ_CELL = "read_cell"
    SNAPSHOT_EXISTS = "snapshot_exists"
    FILE_HASH_EXISTS = "file_hash_exists"
    ARTIFACT_EXISTS = "artifact_exists"
    # Phase 2 — Range/Cells (task_053).
    RANGE_FORMULA_EQUALS = "range_formula_equals"
    RANGE_FORMAT_EQUALS = "range_format_equals"
    RANGE_STRUCTURE = "range_structure"
    VALIDATION_EXISTS = "validation_exists"
    CONDITIONAL_FORMAT_EXISTS = "conditional_format_exists"
    # Phase 2 — PivotTables (task_053).
    PIVOT_TABLE_EXISTS = "pivot_table_exists"
    PIVOT_FIELD_LAYOUT = "pivot_field_layout"
    PIVOT_FIELD_AGGREGATION = "pivot_field_aggregation"
    PIVOT_FILTER_STATE = "pivot_filter_state"
    PIVOT_REFRESH_STATUS = "pivot_refresh_status"
    PIVOT_OUTPUT_RANGE = "pivot_output_range"


class BoundedModel(BaseModel):
    """Base for closed test-spec models: forbid extra fields, require enums."""

    __test__ = False  # not a pytest test class despite the "Test" prefix

    model_config = ConfigDict(extra="forbid", use_enum_values=False)


class TestStep(BoundedModel):
    """A bounded test step with a typed action and typed params (spec §9.2).

    The field set is a **closed, typed** vocabulary: a step can only describe an
    object selector (``sheet`` / ``table`` / ``name`` / ``macro`` / ``field`` …)
    plus scalar or string-keyed parameters of an allowlisted service operation.
    There is no field that can carry executable code, a shell command, a COM
    expression or an arbitrary path, so a suite cannot smuggle execution into a
    step. Phase 2 (task_053) adds the Range/Pivot selectors and the bounded typed
    payloads (``values`` / ``formulas`` / ``format`` / ``orientation`` /
    ``function`` / ``filter``), never an arbitrary COM object.
    """

    model_config = ConfigDict(extra="forbid", use_enum_values=False)

    action: TestAction
    # Object selectors (bounded scalars / string-keyed maps).
    sheet: str | None = None
    sheet_name: str | None = None
    new_name: str | None = None
    copy_to: str | None = None
    table: str | None = None
    key: dict[str, Any] | None = None
    values: dict[str, Any] | None = None
    rows: list[dict[str, Any]] | None = Field(
        default=None,
        description="Row payloads for table_append_rows (header -> value per row).",
    )
    updates: list[dict[str, Any]] | None = Field(
        default=None,
        description="Explicit table_update_rows entries ({index|where, values}).",
    )
    indices: list[int] | None = Field(
        default=None,
        description="Explicit 1-based row indices for table_delete_rows.",
    )
    row_index: int | None = None
    column_name: str | None = None
    position: int | None = Field(
        default=None,
        description="1-based column position for table_add_column.",
    )
    range_address: str | None = Field(
        default=None,
        description=(
            "A1 range or defined name for table_create / table_resize, or the "
            "A1 target address of a range_* step."
        ),
    )
    style: str | None = Field(default=None, description="Table style for table_create.")
    name: str | None = None
    scope: Literal["workbook", "worksheet"] | None = None
    refers_to: str | None = None
    visible: bool | None = None
    component: str | None = None
    source: str | None = None
    insert_at: str | None = Field(
        default=None,
        description=(
            "Insertion position for vba_insert_code: 'declarations', "
            "'before_procedure:<name>' or 'after_procedure:<name>'."
        ),
    )
    macro: str | None = None
    args: list[Any] = Field(default_factory=list, max_length=MAX_MACRO_ARGS)
    # Phase 2 — Range/Cells (task_053). All bounded scalars or typed DTOs.
    source_range: str | None = Field(
        default=None,
        description="Source A1 address for range_copy (never the clipboard).",
    )
    destination_sheet: str | None = Field(
        default=None, description="Destination worksheet for range_copy."
    )
    destination: str | None = Field(
        default=None, description="Destination A1 address for range_copy."
    )
    shift: Literal["left", "right", "up", "down"] | None = Field(
        default=None,
        description=(
            "Explicit shift direction for range_insert ('right'/'down') and "
            "range_delete ('left'/'up'); never guessed."
        ),
    )
    allow_merged: bool | None = Field(
        default=None, description="Permit a range insert/delete over merged cells."
    )
    across: bool | None = Field(
        default=None, description="Merge each row separately for range_merge."
    )
    allow_data_loss: bool | None = Field(
        default=None,
        description="Permit range_merge when non-anchor cells hold data (after preview).",
    )
    values_grid: RangeWriteRequest | None = Field(
        default=None,
        description="Typed 2-D value payload for range_write_values (bounded).",
    )
    formulas_grid: FormulaWriteRequest | None = Field(
        default=None,
        description="Bounded 2-D formula payload for range_write_formulas.",
    )
    format_patch: RangeFormatPatch | None = Field(
        default=None,
        description="Sparse format update for range_set_format (only set fields applied).",
    )
    offset: int | None = Field(
        default=None,
        ge=0,
        description="0-based row offset for range_read_values / range_read_formulas.",
    )
    limit: int | None = Field(
        default=None,
        ge=1,
        description="Maximum rows for range_read_values / range_read_formulas.",
    )
    # Phase 2 — PivotTables (task_053).
    pivot_table: str | None = Field(
        default=None, description="Stable PivotTable Name for pivot_* steps."
    )
    field: str | None = Field(
        default=None, description="PivotField SourceName for pivot_field_* steps."
    )
    orientation: PivotFieldOrientation | None = Field(
        default=None,
        description="Target layout bucket for pivot_field_set_layout (hidden/row/column/page).",
    )
    delete: bool | None = Field(
        default=None,
        description="Destructive PivotField.Delete for pivot_field_set_layout (needs hidden).",
    )
    function: PivotAggregation | None = Field(
        default=None, description="Target aggregation for pivot_field_set_aggregation."
    )
    caption: str | None = Field(
        default=None, description="Unique data-field caption for pivot_field_set_aggregation."
    )
    filter_spec: PivotFilterSpec | None = Field(
        default=None,
        description="Bounded manual (visible-items) filter for pivot_field_apply_filter.",
    )

    @model_validator(mode="after")
    def _check_declarative_size(self) -> TestStep:
        """Reject a declarative Range/Pivot payload that is too large (pre-COM).

        Excel's own ceilings are enforced by
        :class:`~excel_vba_mcp.domain.models.RangeWriteRequest` /
        :class:`~excel_vba_mcp.domain.models.FormulaWriteRequest`; this check only
        keeps the declarative suite document itself bounded so validation stays
        cheap and a suite cannot smuggle a whole-sheet matrix.
        """

        grid = self.values_grid or self.formulas_grid
        if grid is not None and grid.rows > MAX_RANGE_WRITE_ROWS:
            raise ValueError(
                f"{self.action.value} payload rows {grid.rows} exceed "
                f"MAX_RANGE_WRITE_ROWS ({MAX_RANGE_WRITE_ROWS})."
            )
        if self.filter_spec is not None and len(self.filter_spec.visible_items) > (
            MAX_PIVOT_FILTER_ITEMS
        ):
            raise ValueError(
                f"{self.action.value} filter lists {len(self.filter_spec.visible_items)} "
                f"items, exceeding MAX_PIVOT_FILTER_ITEMS ({MAX_PIVOT_FILTER_ITEMS})."
            )
        return self


class Assertion(BoundedModel):
    """A bounded assertion with a typed kind and typed params (spec §9.3).

    Assertions are **read-only** by construction: the model carries only the
    selectors (``sheet`` / ``table`` / ``name`` / ``cell`` / ``field`` /
    ``pivot_table``) and an ``expected`` value to compare against. There is no
    way to express an executable payload. Phase 2 (task_053) adds the
    Range/Pivot selectors and keeps every expectation a bounded scalar, a list of
    strings or a dotted property path — never a COM expression.
    """

    model_config = ConfigDict(extra="forbid", use_enum_values=False)

    type: AssertionType
    sheet: str | None = None
    table: str | None = None
    column: str | None = None
    name: str | None = None
    cell: str | None = Field(
        default=None,
        description="A1 address of a single cell for read_cell / cell_value_equals.",
    )
    expected: Any | None = Field(
        default=None,
        description="Expected value: scalar, list of column names, or status string.",
    )
    value: int | float | None = Field(
        default=None,
        description=(
            "Numeric bound for table_row_count(_gte) / table_column_count / macro_duration_lte."
        ),
    )
    visibility: str | None = None
    refers_to: str | None = None
    hash: str | None = Field(
        default=None,
        description="Expected SHA-256 for file_hash_exists (defaults to the case baseline).",
    )
    path: str | None = None
    # Phase 2 — Range/Cells (task_053).
    address: str | None = Field(
        default=None,
        description=(
            "A1 range for range_formula_equals / range_format_equals / "
            "range_structure / validation_exists / conditional_format_exists."
        ),
    )
    property: str | None = Field(
        default=None,
        min_length=1,
        max_length=64,
        description=(
            "Dotted bounded property path for range_format_equals / range_structure / "
            "pivot_filter_state (e.g. 'number_format', 'font.bold', 'merged_area', "
            "'visible_items'). Never a COM property expression."
        ),
    )
    rows: int | None = Field(
        default=None,
        ge=0,
        description="Expected row count for range_structure (ge=0 rejects a negative expectation).",
    )
    cols: int | None = Field(
        default=None,
        ge=0,
        description="Expected column count for range_structure (ge=0).",
    )
    merged: bool | None = Field(
        default=None,
        description="Expected merged state for range_structure.",
    )
    # Phase 2 — PivotTables (task_053).
    pivot_table: str | None = Field(
        default=None, description="Stable PivotTable Name for pivot_* assertions."
    )
    field: str | None = Field(
        default=None, description="PivotField SourceName for pivot_field_* assertions."
    )
    orientation: PivotFieldOrientation | None = Field(
        default=None,
        description="Expected layout bucket for pivot_field_layout.",
    )
    position: int | None = Field(
        default=None,
        ge=1,
        description="Expected 1-based position for pivot_field_layout.",
    )
    function: PivotAggregation | None = Field(
        default=None,
        description="Expected aggregation for pivot_field_aggregation.",
    )
    caption: str | None = Field(
        default=None,
        description="Expected data-field caption for pivot_field_aggregation.",
    )
    visible_items: list[str] | None = Field(
        default=None,
        max_length=MAX_PIVOT_FILTER_ITEMS,
        description="Expected visible PivotItems for pivot_filter_state.",
    )
    hidden_items: list[str] | None = Field(
        default=None,
        max_length=MAX_PIVOT_FILTER_ITEMS,
        description="Expected hidden PivotItems for pivot_filter_state.",
    )
    refresh_status: str | None = Field(
        default=None,
        description=(
            "Expected PivotRefreshStatus for pivot_refresh_status (defaults to 'refreshed')."
        ),
    )
    refreshed: bool | None = Field(
        default=None,
        description="Expected truth signal for pivot_refresh_status.",
    )
    output_range: str | None = Field(
        default=None,
        description="Expected PivotTable output range for pivot_output_range.",
    )


class TestCaseSpec(BoundedModel):
    """A single declarative test case (spec §9.2, Appendix B)."""

    id: str = Field(min_length=1)
    steps: list[TestStep] = Field(default_factory=list)
    assertions: list[Assertion] = Field(default_factory=list)


class TestSuiteSpec(BoundedModel):
    """A declarative test suite (spec §9.2, Appendix B)."""

    suite: str = Field(min_length=1)
    workbook_id: str = Field(min_length=1)
    baseline: BaselinePolicy = BaselinePolicy.CREATE
    cases: list[TestCaseSpec] = Field(min_length=1)


def validate_action(action: str | TestAction) -> TestAction:
    """Return the :class:`TestAction` for ``action`` or raise ``ValueError``.

    Args:
        action: A documented action string (or already a :class:`TestAction`).

    Raises:
        ValueError: If ``action`` is not one of the documented bounded actions.
    """

    try:
        return TestAction(action)
    except ValueError as exc:
        valid = ", ".join(member.value for member in TestAction)
        raise ValueError(f"Unknown test action {action!r}; allowed: {valid}") from exc


def validate_assertion(assertion_type: str | AssertionType) -> AssertionType:
    """Return the :class:`AssertionType` for ``assertion_type`` or raise ``ValueError``.

    Args:
        assertion_type: A documented assertion kind (or already an :class:`AssertionType`).

    Raises:
        ValueError: If the kind is not one of the documented bounded assertions.
    """

    try:
        return AssertionType(assertion_type)
    except ValueError as exc:
        valid = ", ".join(member.value for member in AssertionType)
        raise ValueError(f"Unknown test assertion {assertion_type!r}; allowed: {valid}") from exc
