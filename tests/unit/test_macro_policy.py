"""Unit tests for the macro / operation-classification policy (task_009)."""

from __future__ import annotations

import pytest

from excel_vba_mcp.domain.errors import AccessVbomDisabledError, VbaProjectLockedError
from excel_vba_mcp.domain.statuses import MacroRunStatus, OperationClass, SessionMode
from excel_vba_mcp.security.macro_policy import (
    AutomationSecurityMode,
    MacroPolicy,
    automation_security_for,
    classify_macro_failure,
    classify_operation,
    diagnose_protection,
    diagnose_vbom,
)

RO = OperationClass.READ_ONLY
MU = OperationClass.MUTATING
EX = OperationClass.EXECUTING


@pytest.mark.parametrize(
    ("tool", "expected"),
    [
        # --- read-only operations ---
        ("excel_open", RO),
        ("excel_get_status", RO),
        ("excel_list_workbooks", RO),
        ("workbook_get_model", RO),
        ("snapshot_create", RO),  # SaveCopyAs does not mutate the open workbook
        ("snapshot_list", RO),
        ("snapshot_diff", RO),
        ("vba_project_info", RO),
        ("vba_list_components", RO),
        ("vba_get_components", RO),
        ("vba_read_component", RO),
        ("vba_get_component_source", RO),
        ("vba_list_procedures", RO),
        ("vba_get_references", RO),
        ("vba_search_code", RO),
        ("vba_get_compile_diagnostics", RO),
        ("vba_export", RO),
        ("vba_export_project", RO),
        ("vba_export_component", RO),
        ("macro_list", RO),
        ("macro_get_last_result", RO),
        ("sheets_list", RO),
        ("sheet_get", RO),
        ("tables_list", RO),
        ("table_get", RO),
        ("table_read_rows", RO),
        ("names_list", RO),
        ("name_get", RO),
        ("controls_list", RO),
        ("control_get", RO),
        ("shapes_list", RO),
        ("shape_get", RO),
        # --- mutating operations ---
        ("excel_save", MU),
        ("snapshot_restore", MU),
        ("vba_write", MU),
        ("vba_insert_code", MU),
        ("vba_replace_component_source", MU),
        ("vba_add_component", MU),
        ("vba_remove_component", MU),
        ("sheet_add", MU),
        ("sheet_rename", MU),
        ("sheet_copy", MU),
        ("sheet_delete", MU),
        ("sheet_set_visibility", MU),
        ("table_create", MU),
        ("table_add_column", MU),
        ("table_append_rows", MU),
        ("table_insert_row", MU),
        ("table_update_rows", MU),
        ("table_delete_rows", MU),
        ("table_delete_column", MU),
        ("table_resize", MU),
        ("table_row_add", MU),
        ("table_row_delete", MU),
        ("table_column_delete", MU),
        ("name_create", MU),
        ("name_update", MU),
        ("name_delete", MU),
        ("control_set_on_action", MU),
        ("shape_set_on_action", MU),
        # --- executing operations ---
        ("macro_run", EX),
        ("compile", EX),
        ("vba_compile", EX),
        ("test_run", EX),
        ("test_run_suite", EX),
    ],
)
def test_classify_table(tool: str, expected: OperationClass) -> None:
    assert classify_operation(tool, {}) is expected


def test_excel_close_without_save_is_read_only() -> None:
    assert classify_operation("excel_close", {}) is RO
    assert classify_operation("excel_close") is RO
    assert classify_operation("excel_close", {"save_changes": False}) is RO


@pytest.mark.parametrize("flag", [True, "true", "1", "yes"])
def test_excel_close_with_save_is_mutating(flag: object) -> None:
    assert classify_operation("excel_close", {"save_changes": flag}) is MU


def test_unknown_tool_fails_closed_as_executing() -> None:
    # Documented fail-closed behaviour: unknown tools are the most restrictive.
    assert classify_operation("totally_unknown_tool", {}) is EX
    assert classify_operation("", {}) is EX


def test_classify_via_macro_policy_class() -> None:
    assert MacroPolicy.classify_operation("macro_run", {}) is EX
    assert MacroPolicy.classify_operation("excel_save", {}) is MU


@pytest.mark.parametrize(
    ("mode", "expected"),
    [
        (SessionMode.INSPECT, AutomationSecurityMode.FORCE_DISABLE),
        (SessionMode.EXECUTE, AutomationSecurityMode.BY_UI),
    ],
)
def test_automation_security_for(mode: SessionMode, expected: AutomationSecurityMode) -> None:
    assert automation_security_for(mode) is expected
    assert MacroPolicy.automation_security_for(mode) is expected


def test_automation_security_never_low() -> None:
    for mode in SessionMode:
        assert automation_security_for(mode) is not AutomationSecurityMode.LOW


def test_automation_security_enum_values() -> None:
    assert AutomationSecurityMode.LOW == 1
    assert AutomationSecurityMode.BY_UI == 2
    assert AutomationSecurityMode.FORCE_DISABLE == 3


def test_diagnose_vbom() -> None:
    assert diagnose_vbom(True) is None
    assert MacroPolicy.diagnose_vbom(True) is None
    err = diagnose_vbom(False)
    assert isinstance(err, AccessVbomDisabledError)
    assert "Trust Center" in err.message


def test_diagnose_protection() -> None:
    assert diagnose_protection(0) is None
    assert diagnose_protection(2) is None
    err = diagnose_protection(1)  # vbext_pp_locked
    assert isinstance(err, VbaProjectLockedError)
    assert MacroPolicy.diagnose_protection(1) is not None


def test_classify_macro_failure_blocked() -> None:
    assert classify_macro_failure(None, approved=False) is MacroRunStatus.BLOCKED
    assert classify_macro_failure(0x0, approved=False) is MacroRunStatus.BLOCKED


def test_classify_macro_failure_not_found() -> None:
    assert classify_macro_failure(0x800A03ED, approved=True) is MacroRunStatus.NOT_FOUND


@pytest.mark.parametrize("hresult", [0x800A03EC, 0x80004005])
def test_classify_macro_failure_runtime_error(hresult: int) -> None:
    assert classify_macro_failure(hresult, approved=True) is MacroRunStatus.RUNTIME_ERROR


def test_classify_macro_failure_indeterminate() -> None:
    # No HRESULT (soft timeout / recovery path) and unknown codes must not be a
    # false success; they need recovery.
    assert classify_macro_failure(None, approved=True) is MacroRunStatus.NEEDS_RECOVERY
    assert classify_macro_failure(0x12345678, approved=True) is MacroRunStatus.NEEDS_RECOVERY


def test_classify_macro_failure_never_success() -> None:
    # A bare call must never be reported as success by the classifier.
    for hresult in (None, 0x0, 0x800A03EC, 0x800A03ED, 0x80004005, 0xDEADBEEF):
        assert classify_macro_failure(hresult, approved=True) is not MacroRunStatus.SUCCESS
