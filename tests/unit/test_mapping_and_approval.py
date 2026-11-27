"""Mapping-totalness and security/approval-matrix completeness tests (B12 / WP11).

Static: loads the WP0 baseline fixture (``tests/fixtures/wp0_baseline.json`` — the
source of truth) and, for the mapping totalness, cross-checks the live registered
tool surface. It never touches Excel/COM and never starts the worker.

These tests lock the WP11 count/name/mapping/approval contract:
  * the ledger has exactly 92 rows and every old_name resolves to one target;
  * the live registered names equal the ledger's distinct target set (41);
  * the security/approval matrix covers all 92 tools;
  * the four high-risk tools keep an explicit standalone/protected disposition;
  * no consolidation merges two distinct *approval* permission profiles.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import excel_vba_mcp.server as server

pytestmark = pytest.mark.anyio

_FIXTURE = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "wp0_baseline.json"

# The three high-risk tools that stay standalone with unchanged names (H1).
_STANDALONE_HIGH_RISK = {"macro_run", "excel_save_as", "table_delete_column"}
# The fourth high-risk tool is excel_close, folded into the workbook close action
# with the D1 refuse-by-default disposition (protected, not standalone).
_CLOSE_OLD_NAME = "excel_close"

# Approval-gated permission profiles require an ApprovalPolicy decision. The
# non-gated profiles are auto-allowed / never grant run permission: P_READ (reads)
# and P_BOOK (ungated bookkeeping). A consolidation "merges two distinct
# permission profiles" only when it crosses this gated <-> non-gated boundary.
_GATED_PROFILES = {"P_MUT", "P_EXEC", "P_HIGH", "P_LIFE"}
_NON_GATED_PROFILES = {"P_READ", "P_BOOK"}

# The only consolidated tools that cross the approval boundary are the documented
# accepted deviations: ``test`` (R-Q2: run-suite P_EXEC + get-run P_READ) and
# ``shape`` / ``control`` (R-Q5: guarded set-on-action P_MUT + list/get P_READ).
# ``macro`` merges P_BOOK (ungated bookkeeping) + P_READ and therefore does NOT
# cross the boundary — classify-background never grants run permission.
_EXPECTED_BOUNDARY_MERGES = {"test", "shape", "control"}


@pytest.fixture(scope="module")
def baseline() -> dict:
    return json.loads(_FIXTURE.read_text(encoding="utf-8"))


def test_ledger_has_exactly_92_rows(baseline: dict) -> None:
    assert len(baseline["ledger"]) == 92


def test_every_old_name_maps_to_exactly_one_target(baseline: dict) -> None:
    ledger = baseline["ledger"]
    old_names = [row["old_name"] for row in ledger]
    assert len(set(old_names)) == 92, "ledger contains a duplicate old_name"

    for row in ledger:
        target = row["target"]
        assert target["kind"] in {"consolidated", "standalone"}, (
            f"{row['old_name']} has an undefined target kind"
        )
        assert target["tool"], f"{row['old_name']} target is missing a tool name"
        if target["kind"] == "consolidated":
            assert target.get("action"), f"{row['old_name']} consolidated target lacks an action"


async def test_registered_names_equal_ledger_target_set(baseline: dict) -> None:
    targets = {row["target"]["tool"] for row in baseline["ledger"]}
    assert len(targets) == 41
    names = {tool.name for tool in await server.mcp.list_tools()}
    assert names == targets, "registered tool set differs from the ledger target set"


def test_security_matrix_covers_all_92_tools(baseline: dict) -> None:
    security = baseline["security"]
    assert len(security) == 92

    ledger_names = {row["old_name"] for row in baseline["ledger"]}
    security_names = {row["tool"] for row in security}
    assert security_names == ledger_names, "security matrix does not cover the ledger exactly"


def test_high_risk_tools_have_explicit_disposition(baseline: dict) -> None:
    by_name = {row["tool"]: row for row in baseline["security"]}

    for name in _STANDALONE_HIGH_RISK:
        row = by_name[name]
        assert row["risk"] == "high-risk", f"{name} must be high-risk"
        assert row["new_name"] == name, f"{name} must stay standalone (new_name == tool)"

    close = by_name[_CLOSE_OLD_NAME]
    assert close["risk"] == "high-risk", "excel_close must carry the high-risk disposition"
    assert close["new_name"] == "workbook", "excel_close folds into the workbook close action"
    assert close.get("note"), "excel_close must document the D1 refuse-by-default disposition"


def test_no_consolidation_merges_approval_boundary(baseline: dict) -> None:
    by_name = {row["tool"]: row for row in baseline["security"]}

    members: dict[str, list[str]] = {}
    for row in baseline["ledger"]:
        target = row["target"]
        if target["kind"] == "consolidated":
            members.setdefault(target["tool"], []).append(row["old_name"])

    boundary_merges: set[str] = set()
    raw_multi_profile: set[str] = set()
    for tool, old_names in members.items():
        profiles = {by_name[old]["permission_profile"] for old in old_names}
        if len(profiles) > 1:
            raw_multi_profile.add(tool)
        gated = profiles & _GATED_PROFILES
        non_gated = profiles & _NON_GATED_PROFILES
        # A boundary merge combines a gated (approval-required) profile with a
        # non-gated (auto-allowed / ungated) one under a single tool name.
        if gated and non_gated:
            boundary_merges.add(tool)
        # No consolidation may merge two distinct gated profiles either.
        assert len(gated) <= 1, f"{tool} merges distinct gated profiles {sorted(gated)}"

    # Raw distinct-profile tools are exactly the three documented exceptions plus
    # ``macro`` (whose extra profile is the ungated P_BOOK bookkeeping class).
    assert raw_multi_profile == {"test", "shape", "control", "macro"}, (
        f"unexpected raw multi-profile tools: {sorted(raw_multi_profile)}"
    )
    # Only test/shape/control cross the gated <-> non-gated approval boundary; the
    # accepted H1 mixed-profile exceptions are exactly these three.
    assert boundary_merges == _EXPECTED_BOUNDARY_MERGES, (
        f"unexpected approval-boundary merges: {sorted(boundary_merges)}"
    )


def test_macro_merges_only_non_gated_profiles(baseline: dict) -> None:
    # Documents why ``macro`` is not an approval-boundary merge: classify-background
    # is P_BOOK (ungated bookkeeping, never grants run permission) + P_READ reads.
    by_name = {row["tool"]: row for row in baseline["security"]}
    macro_profiles = {
        by_name[old]["permission_profile"]
        for old in ("macro_list", "macro_classify_background", "macro_get_last_result")
    }
    assert macro_profiles <= _NON_GATED_PROFILES