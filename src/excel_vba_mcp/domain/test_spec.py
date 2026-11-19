"""Bounded declarative test-suite schema (spec §9.2–9.3, Appendix B).

The suite format is declarative and closed: actions and assertions are modelled
as closed :class:`enum.StrEnum` vocabularies with ``extra="forbid"`` step and
assertion models, so arbitrary Python / shell / COM code cannot be expressed.
Unknown action or assertion kinds are rejected at validation time.

This module defines the schema only. It does not execute anything; the runner
(task_035) consumes these models.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.domain.statuses import BaselinePolicy

# Maximum number of positional arguments accepted by a ``macro_run`` step, kept
# here (not in steps.py) to mirror the Excel Application.Run 30-argument limit.
MAX_MACRO_ARGS = 30


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
    VBA_REPLACE_COMPONENT_SOURCE = "vba_replace_component_source"
    VBA_INSERT_CODE = "vba_insert_code"
    VBA_COMPILE = "vba_compile"
    MACRO_RUN = "macro_run"


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
    TABLE_COLUMN_NAMES = "table_column_names"
    NAME_EXISTS = "name_exists"
    NAME_REFERS_TO = "name_refers_to"
    CELL_VALUE_EQUALS = "cell_value_equals"
    SNAPSHOT_EXISTS = "snapshot_exists"
    FILE_HASH_EXISTS = "file_hash_exists"
    ARTIFACT_EXISTS = "artifact_exists"


class BoundedModel(BaseModel):
    """Base for closed test-spec models: forbid extra fields, require enums."""

    __test__ = False  # not a pytest test class despite the "Test" prefix

    model_config = ConfigDict(extra="forbid", use_enum_values=False)


class TestStep(BoundedModel):
    """A bounded test step with a typed action and typed params (spec §9.2)."""

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
    row_index: int | None = None
    column_name: str | None = None
    name: str | None = None
    scope: Literal["workbook", "worksheet"] | None = None
    refers_to: str | None = None
    visible: bool | None = None
    component: str | None = None
    source: str | None = None
    insert_at: int | None = None
    macro: str | None = None
    args: list[Any] = Field(default_factory=list, max_length=MAX_MACRO_ARGS)


class Assertion(BoundedModel):
    """A bounded assertion with a typed kind and typed params (spec §9.3)."""

    model_config = ConfigDict(extra="forbid", use_enum_values=False)

    type: AssertionType
    sheet: str | None = None
    table: str | None = None
    column: str | None = None
    name: str | None = None
    cell: str | None = None
    expected: Any | None = None
    value: int | float | None = None
    visibility: str | None = None
    refers_to: str | None = None
    hash: str | None = None
    path: str | None = None


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
