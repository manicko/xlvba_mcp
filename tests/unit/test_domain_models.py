"""Unit tests for Phase 1 domain DTOs, statuses, errors and test spec (task_007)."""

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
    format_hresult,
)
from excel_vba_mcp.domain.models import (
    SCHEMA_VERSION,
    WorkbookInfo,
    WorkbookModel,
)
from excel_vba_mcp.domain.statuses import (
    MacroRunStatus,
    OperationClass,
    WorkbookSessionState,
)
from excel_vba_mcp.domain.test_spec import (
    AssertionType,
    TestAction,
    TestStep,
    TestSuiteSpec,
    validate_action,
    validate_assertion,
)

_TS = datetime(2026, 1, 1, 12, 0, 0, tzinfo=UTC)

_SPEC_11_CODES = {
    "ACCESS_VBOM_DISABLED",
    "VBA_PROJECT_LOCKED",
    "MISSING_REFERENCE",
    "COM_BUSY",
    "RPC_E_CALL_REJECTED",
    "COMPILE_ERROR",
    "COMPILE_INDETERMINATE",
    "MACRO_BLOCKED",
    "MACRO_NOT_FOUND",
    "MACRO_RUNTIME_ERROR",
    "MACRO_TIMEOUT",
    "EXCEL_HUNG",
    "SNAPSHOT_FAILED",
    "WORKBOOK_READ_ONLY",
    "PATH_NOT_ALLOWED",
}


# ---------------------------------------------------------------------------
# JSON serialization round-trip
# ---------------------------------------------------------------------------


def test_workbook_model_json_round_trip() -> None:
    model = WorkbookModel(
        created_at=_TS,
        updated_at=_TS,
        workbook=WorkbookInfo(
            workbook_id="wb_7f3a",
            name="Report.xlsm",
            format="xlsm",
            path="C:/work/Report.xlsm",
            source_path="C:/projects/Report.xlsm",
            read_only=False,
            has_vba=True,
            excel_version="16.0",
            session_id="sess_1",
            mode="inspect",
        ),
    )

    payload = json.loads(model.model_dump_json())

    assert payload["schema_version"] == SCHEMA_VERSION
    restored = WorkbookModel.model_validate(payload)
    assert restored.model_dump(mode="json") == model.model_dump(mode="json")
    assert restored.workbook.mode.value == "inspect"


def test_models_have_no_com_object_fields() -> None:
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
            mode="execute",
        ),
    )
    dumped = model.model_dump(mode="json")
    # Only JSON-native types may appear; nothing that could hold a COM proxy.
    _assert_json_native(dumped)


def test_workbook_info_has_no_original_name_field() -> None:
    # Block C: ``source_path`` is the original allowlisted path; ``original_name``
    # must stay absent (no duplicate/confusing field).
    assert "original_name" not in WorkbookInfo.model_fields
    assert "source_path" in WorkbookInfo.model_fields


def _assert_json_native(value: object) -> None:
    if isinstance(value, dict):
        for child in value.values():
            _assert_json_native(child)
    elif isinstance(value, (list, tuple)):
        for child in value:
            _assert_json_native(child)
    else:
        assert isinstance(value, str | int | float | bool | None), type(value)


# ---------------------------------------------------------------------------
# Test spec: bounded actions / assertions
# ---------------------------------------------------------------------------


def test_test_suite_accepts_documented_actions() -> None:
    suite = TestSuiteSpec.model_validate(
        {
            "suite": "report-smoke",
            "workbook_id": "wb_7f3a",
            "baseline": "create",
            "cases": [
                {
                    "id": "monthly-sales",
                    "steps": [
                        {"action": "table_update_rows", "table": "t", "key": {"Report": "Sales"}},
                        {"action": "vba_compile"},
                        {"action": "macro_run", "macro": "modReports.GenerateSales", "args": []},
                    ],
                    "assertions": [
                        {"type": "sheet_exists", "name": "SalesReport"},
                        {"type": "table_row_count_gte", "table": "t", "value": 1},
                        {"type": "macro_status", "expected": "success"},
                    ],
                }
            ],
        }
    )
    assert suite.cases[0].steps[0].action is TestAction.TABLE_UPDATE_ROWS
    assert suite.cases[0].assertions[0].type is AssertionType.SHEET_EXISTS


def test_test_spec_rejects_unknown_action() -> None:
    with pytest.raises(ValidationError):
        TestStep.model_validate({"action": "os.system", "sheet": "x"})


def test_test_spec_rejects_unknown_assertion() -> None:
    with pytest.raises(ValidationError):
        TestSuiteSpec.model_validate(
            {
                "suite": "s",
                "workbook_id": "w",
                "cases": [
                    {
                        "id": "c",
                        "steps": [],
                        "assertions": [{"type": "python_eval", "expected": "1"}],
                    }
                ],
            }
        )


def test_test_spec_rejects_extra_step_field() -> None:
    with pytest.raises(ValidationError):
        TestStep.model_validate({"action": "vba_compile", "code": "Shell 'calc'"})


def test_validate_action_and_assertion_reject_unknown() -> None:
    assert validate_action("macro_run") is TestAction.MACRO_RUN
    assert validate_assertion("macro_status") is AssertionType.MACRO_STATUS
    with pytest.raises(ValueError):
        validate_action("shell_exec")
    with pytest.raises(ValueError):
        validate_assertion("eval_python")


# ---------------------------------------------------------------------------
# Errors: §11 coverage and HRESULT preservation
# ---------------------------------------------------------------------------


def test_all_spec_11_error_codes_exist() -> None:
    # Stage 1 keeps the closed spec §11 set; Phase 2 (task_040) appends codes
    # additively, so the §11 set must remain a subset of the full enum.
    assert {code.value for code in ErrorCode} >= _SPEC_11_CODES


def test_error_preserves_hresult_and_hex() -> None:
    error = ExcelVbaMcpError(
        "busy",
        code=ErrorCode.COM_BUSY,
        hresult=0x8001010A,
        context={"operation": "macro_run"},
    )
    assert error.hresult == 0x8001010A
    assert error.hresult_hex == "0x8001010A"
    payload = error.to_dict()
    assert payload["code"] == "COM_BUSY"
    assert payload["hresult_hex"] == "0x8001010A"
    assert payload["context"]["operation"] == "macro_run"
    # to_dict must be JSON serializable.
    assert json.loads(json.dumps(payload))["hresult"] == 0x8001010A


def test_error_from_hresult_maps_known_code() -> None:
    error = error_from_hresult(0x80010001, "rejected")
    assert isinstance(error, err.RpcCallRejectedError)
    assert error.code is ErrorCode.RPC_E_CALL_REJECTED
    assert error.hresult == 0x80010001
    assert error.hresult_hex == "0x80010001"


def test_error_from_hresult_unknown_defaults_to_com_busy() -> None:
    error = error_from_hresult(0x12345678, "unknown")
    assert error.code is ErrorCode.COM_BUSY
    assert error.hresult == 0x12345678


def test_format_hresult_is_eight_hex_digits() -> None:
    assert format_hresult(0x1) == "0x00000001"
    assert format_hresult(0x800A03EC) == "0x800A03EC"


# ---------------------------------------------------------------------------
# Statuses
# ---------------------------------------------------------------------------


def test_operation_class_values() -> None:
    assert {member.value for member in OperationClass} == {"read_only", "mutating", "executing"}


def test_session_needs_recovery_state() -> None:
    assert WorkbookSessionState.NEEDS_RECOVERY.value == "running/needs_recovery"


def test_session_deadline_exceeded_state_value() -> None:
    # Block E (additive): pin the wire value of the deadline-exceeded state.
    assert WorkbookSessionState.DEADLINE_EXCEEDED.value == "running/deadline_exceeded"


def test_macro_status_values() -> None:
    assert {member.value for member in MacroRunStatus} == {
        "success",
        "runtime_error",
        "blocked",
        "not_found",
        "timeout",
        "needs_recovery",
        # Block E (additive): deadline-exceeded macro, no live-pid kill.
        "suspected_stalled",
    }
