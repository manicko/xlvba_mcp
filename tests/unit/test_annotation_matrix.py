"""Annotation-matrix contract tests (B11 / WP10 + WP11 lockstep part 1).

Static, read-only: imports the server module (which registers tools at import
time) and compares every registered tool's ``ToolAnnotations`` against the
authoritative per-tool triple in ``tests/fixtures/wp0_baseline.json``
(``tool_annotations``). It never touches Excel/COM and never starts the worker.

The fixture is the source of truth (H3). These tests lock:
  * exact triple equality for every final tool (no ``None`` / ``False`` drift);
  * ``range_structural`` is destructive-only and NOT idempotent (H4);
  * the dual-hint set (destructive AND idempotent) is exactly
    ``{range_merge, pivot_field}`` (R-Q4);
  * no tool is over-flagged relative to the matrix.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import excel_vba_mcp.server as server

pytestmark = pytest.mark.anyio

_FIXTURE = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "wp0_baseline.json"


def _matrix() -> dict[str, tuple[bool, bool, bool]]:
    data = json.loads(_FIXTURE.read_text(encoding="utf-8"))
    return {
        row["tool"]: (row["read_only"], row["destructive"], row["idempotent"])
        for row in data["tool_annotations"]
    }


async def _registered() -> dict[str, tuple[bool, bool, bool]]:
    tools = await server.mcp.list_tools()
    result: dict[str, tuple[bool, bool, bool]] = {}
    for tool in tools:
        annotations = tool.annotations
        assert annotations is not None, f"{tool.name} has no annotations"
        read_only = annotations.read_only_hint
        destructive = annotations.destructive_hint
        idempotent = annotations.idempotent_hint
        assert read_only is not None, f"{tool.name} emits a None read_only hint"
        assert destructive is not None, f"{tool.name} emits a None destructive hint"
        assert idempotent is not None, f"{tool.name} emits a None idempotent hint"
        result[tool.name] = (read_only, destructive, idempotent)
    return result


async def test_every_registered_tool_matches_matrix_exactly() -> None:
    matrix = _matrix()
    registered = await _registered()

    assert set(registered) == set(matrix), "registered tool set differs from the matrix"
    for name, triple in registered.items():
        assert triple == matrix[name], f"{name} annotations {triple} != matrix {matrix[name]}"


async def test_all_hints_are_explicit_booleans() -> None:
    # R-Q12 / A2: no hint may be ``None`` (which would fall back to the MCP
    # ``destructiveHint=true`` default and mis-advertise additive tools).
    for tool in await server.mcp.list_tools():
        annotations = tool.annotations
        assert annotations is not None, f"{tool.name} has no annotations"
        for hint in (
            annotations.read_only_hint,
            annotations.destructive_hint,
            annotations.idempotent_hint,
        ):
            assert hint is not None, f"{tool.name} emits a None hint"


async def test_range_structural_is_not_idempotent() -> None:
    registered = await _registered()
    read_only, destructive, idempotent = registered["range_structural"]
    assert read_only is False
    assert destructive is True
    assert idempotent is False, "range_structural must NOT be idempotent (H4)"


async def test_dual_hint_set_is_exactly_range_merge_and_pivot_field() -> None:
    registered = await _registered()
    dual = {
        name
        for name, (_read_only, destructive, idempotent) in registered.items()
        if destructive and idempotent
    }
    assert dual == {"range_merge", "pivot_field"}


async def test_no_tool_is_over_flagged_relative_to_matrix() -> None:
    # A tool is over-flagged when it advertises a ``True`` hint the matrix denies.
    matrix = _matrix()
    registered = await _registered()
    for name, (read_only, destructive, idempotent) in registered.items():
        exp_read_only, exp_destructive, exp_idempotent = matrix[name]
        assert not (read_only and not exp_read_only), f"{name} over-flagged read_only"
        assert not (destructive and not exp_destructive), f"{name} over-flagged destructive"
        assert not (idempotent and not exp_idempotent), f"{name} over-flagged idempotent"