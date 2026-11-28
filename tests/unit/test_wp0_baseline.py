"""Completeness tests for the WP0 baseline fixture (plan B0 / WP0).

Static, data-only: loads ``tests/fixtures/wp0_baseline.json`` and asserts the
four baseline matrices are complete and internally consistent. It never imports
the server, touches Excel/COM, or starts the worker.

The fixture is the "before" snapshot every later block (B2-B12) is verified
against: a 94-row mapping ledger, a 94-row security/approval matrix, a per-action
annotation truth matrix, and one derived annotation triple per distinct target.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

_FIXTURE = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "wp0_baseline.json"

# The four high-risk tools that must keep an explicit, standalone disposition (H1).
_STANDALONE_HIGH_RISK = {
    "macro_run": "macro_run",
    "excel_save_as": "workbook_save_copy",
    "table_delete_column": "table_delete_column",
}
# The fourth high-risk tool is excel_close, folded into workbook with refuse-by-default.
_CLOSE_HIGH_RISK = "excel_close"

_KINDS = {"consolidated", "standalone"}


@pytest.fixture(scope="module")
def baseline() -> dict:
    return json.loads(_FIXTURE.read_text(encoding="utf-8"))


def test_fixture_exists_and_has_schema_version(baseline: dict) -> None:
    assert _FIXTURE.is_file()
    assert baseline["schema_version"] == 1


def test_baseline_count_and_bytes(baseline: dict) -> None:
    meta = baseline["baseline"]
    assert meta["tool_count"] == 43
    # WP0 "before" ledger count is 94 (92 + the two justified large-data standalone
    # tools workbook_file_read / data_export); serialized_bytes is the live-recomputed
    # 43-tool surface measurement (the consolidated live surface, not the 92-tool WP0
    # pre-consolidation snapshot).
    assert meta["serialized_bytes"] == 352155
    assert meta["sdk"] == {"mcp": "2.3.0", "mcp-types": "2.3.0", "pydantic": "2.14.4"}


def test_ledger_totalness_and_resolution(baseline: dict) -> None:
    ledger = baseline["ledger"]
    assert len(ledger) == 94

    old_names = [row["old_name"] for row in ledger]
    assert len(set(old_names)) == 94, "ledger contains a duplicate old_name"

    for row in ledger:
        target = row["target"]
        assert target["kind"] in _KINDS, f"{row['old_name']} has an undefined target kind"
        assert target["tool"], f"{row['old_name']} target is missing a tool name"
        if target["kind"] == "consolidated":
            assert target.get("action"), f"{row['old_name']} consolidated target lacks an action"


def test_ledger_distinct_target_count(baseline: dict) -> None:
    targets = {row["target"]["tool"] for row in baseline["ledger"]}
    assert len(targets) == 43


def test_security_matrix_covers_every_tool(baseline: dict) -> None:
    security = baseline["security"]
    assert len(security) == 94

    ledger_names = {row["old_name"] for row in baseline["ledger"]}
    security_names = {row["tool"] for row in security}
    assert security_names == ledger_names, "security matrix does not cover the ledger exactly"

    for row in security:
        assert row["kilo_policy"] == "ask (xlvba_*)"
        assert row["permission_profile"].startswith("P_")
        assert row["risk"] in {"read", "additive-mutation", "destructive", "lifecycle", "high-risk"}


def test_security_new_name_matches_ledger_target(baseline: dict) -> None:
    target_by_old = {row["old_name"]: row["target"]["tool"] for row in baseline["ledger"]}
    for row in baseline["security"]:
        assert row["new_name"] == target_by_old[row["tool"]], (
            f"{row['tool']} new_name disagrees with its ledger target"
        )


def test_high_risk_tools_have_explicit_disposition(baseline: dict) -> None:
    by_name = {row["tool"]: row for row in baseline["security"]}

    for old_name, new_name in _STANDALONE_HIGH_RISK.items():
        row = by_name[old_name]
        assert row["risk"] == "high-risk", f"{old_name} must be high-risk"
        assert row["new_name"] == new_name, f"{old_name} standalone mismatch"

    close = by_name[_CLOSE_HIGH_RISK]
    assert close["risk"] == "high-risk", "excel_close must carry the high-risk disposition"
    assert close["new_name"] == "workbook", "excel_close folds into the workbook close action"
    assert close.get("note"), "excel_close must document the D1 refuse-by-default disposition"


def test_test_tool_is_documented_mixed_profile(baseline: dict) -> None:
    by_name = {row["tool"]: row for row in baseline["security"]}
    run = by_name["test_suite_run"]
    get = by_name["test_get_run"]
    assert run["new_name"] == get["new_name"] == "test_suite"
    assert run["permission_profile"] == "P_EXEC"
    assert get["permission_profile"] == "P_READ"
    assert run.get("note"), "the accepted H1 mixed-profile deviation must be noted"


def test_truth_covers_every_consolidated_action(baseline: dict) -> None:
    ledger = baseline["ledger"]
    truth = baseline["truth"]

    expected: dict[str, set[str]] = {}
    for row in ledger:
        target = row["target"]
        if target["kind"] == "consolidated":
            expected.setdefault(target["tool"], set()).add(target["action"])

    actual: dict[str, set[str]] = {}
    for row in truth:
        actual.setdefault(row["tool"], set()).add(row["action"])

    for tool, actions in expected.items():
        assert tool in actual, f"truth matrix is missing consolidated tool {tool}"
        assert actual[tool] == actions, f"truth actions for {tool} do not match the ledger"


def test_tool_annotations_one_row_per_target(baseline: dict) -> None:
    annotations = baseline["tool_annotations"]
    targets = {row["target"]["tool"] for row in baseline["ledger"]}

    names = [row["tool"] for row in annotations]
    assert len(names) == len(set(names)) == 43, "tool_annotations must have one row per target"
    assert set(names) == targets


def test_tool_annotations_match_truth_derivation(baseline: dict) -> None:
    # Tool-level triple is derived from per-action truth (MCP default semantics):
    # readOnly = all actions read-only; destructive = any action destructive;
    # idempotent = all actions idempotent.
    by_tool: dict[str, list[dict]] = {}
    for row in baseline["truth"]:
        by_tool.setdefault(row["tool"], []).append(row)

    for ann in baseline["tool_annotations"]:
        actions = by_tool[ann["tool"]]
        assert ann["read_only"] == all(a["read_only"] for a in actions)
        assert ann["destructive"] == any(a["destructive"] for a in actions)
        assert ann["idempotent"] == all(a["idempotent"] for a in actions)


def test_range_structural_is_not_idempotent(baseline: dict) -> None:
    ann = {row["tool"]: row for row in baseline["tool_annotations"]}
    structural = ann["range_structural"]
    assert structural["read_only"] is False
    assert structural["destructive"] is True
    assert structural["idempotent"] is False, "range_structural must NOT be idempotent (H4)"


def test_dual_hint_set_is_exactly_range_merge_and_pivot_field(baseline: dict) -> None:
    dual = {
        row["tool"]
        for row in baseline["tool_annotations"]
        if row["destructive"] and row["idempotent"]
    }
    assert dual == {"range_merge", "pivot_field"}


def test_locked_annotation_triples(baseline: dict) -> None:
    ann = {row["tool"]: row for row in baseline["tool_annotations"]}

    def triple(tool: str) -> tuple[bool, bool, bool]:
        row = ann[tool]
        return row["read_only"], row["destructive"], row["idempotent"]

    assert triple("workbook") == (False, True, False)
    assert triple("test_suite") == (False, True, False)
    assert triple("range_write") == (False, False, False)
    assert triple("vba_replace_component_source") == (False, True, False)
    assert triple("shape") == (False, False, True)
    assert triple("control") == (False, False, True)
    for read_tool in (
        "workbook_read",
        "sheet_read",
        "table_read",
        "pivot_read",
        "range_read",
        "defined_name_read",
        "vba_read",
    ):
        assert triple(read_tool) == (True, False, True), f"{read_tool} must be (T,F,T)"