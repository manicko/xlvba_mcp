"""Description-contract lint for Work Package U1 (plan_03 §5.1 / §12.1 / §14).

Static, read-only: imports the server module (which registers tools at import
time) and inspects tool descriptions and the static glossary resource. It never
touches Excel/COM and never starts the worker.
"""

from __future__ import annotations

import pytest

import excel_vba_mcp.server as server

pytestmark = pytest.mark.anyio

# Workflow-critical tools that must follow the §5.1 template (Purpose / Use when /
# Do not use when / Preconditions / Parameters / Behavior / Result / Next actions).
_CRITICAL = [
    "workbook",
    "workbook_read",
    "excel_save_as",
    "range_write",
    "range_read",
    "range_structural",
    "macro_run",
    "macro",
    "table",
    "table_delete_rows",
    "name",
    "sheet",
    "shape",
    "control",
    "snapshot_restore",
    "excel_recover",
    "job_request_cancel",
    "job_get",
    "vba_compile_diagnostics",
    "vba_replace_component_source",
    "vba_insert_code",
    "test",
]

# Tools that must carry exactly one request example (R-D8).
_EXAMPLE_TOOLS = [
    "range_write",
    "table",
    "vba_replace_component_source",
    "test",
]

# Truthful, non-negotiable phrases (R-D1).
_TRUTH = {
    "vba_compile_diagnostics": "does NOT compile",
    "macro": "never grants run permission",
    "job_request_cancel": "intent-only; never interrupts",
}

_GLOSSARY_TERMS = [
    "R13",
    "BG06",
    "preflight",
    "truth signal",
    "fail-closed",
    "RECOVERY",
    "STA",
    "Stage 1",
    "normalized source hash",
    "SUSPECTED_STALLED",
    "interrupted_unknown",
    "indeterminate",
    "working copy",
    "snapshot gate",
    "guarded mutation",
    "approval gate",
]


async def _descriptions() -> dict[str, str]:
    tools = await server.mcp.list_tools()
    return {tool.name: (tool.description or "") for tool in tools}


async def test_tool_count_is_92_exposed() -> None:
    tools = await server.mcp.list_tools()
    # B2 (WP1) consolidated the six workbook verbs into workbook + workbook_read
    # (excel_save_as stays standalone): 92 -> 88. B3 (WP2) consolidates the six
    # non-destructive sheet verbs into sheet + sheet_read (sheet_delete stays
    # standalone): 88 -> 84. B4 (WP3) consolidates the nine non-destructive table
    # verbs into table + table_read (table_delete_rows and table_delete_column stay
    # standalone): 84 -> 77. B5 (WP4) consolidates the four non-destructive name
    # verbs into name + name_read (name_delete stays standalone): 77 -> 75. B6
    # (WP5) consolidates the three shape verbs into shape and the three control
    # verbs into control: 75 -> 71. B7 (WP6) consolidates the three macro
    # read/bookkeeping verbs into macro (macro_run stays standalone) and the two
    # test verbs into test: 71 -> 68. B8 (WP7) splits the sixteen range verbs
    # into range_read / range_write / range_structural (range_delete and
    # range_merge stay standalone): 68 -> 57. B12 finalizes the count.
    assert len(tools) == 57


async def test_critical_tools_follow_template() -> None:
    descriptions = await _descriptions()
    for name in _CRITICAL:
        text = descriptions[name]
        for marker in ("Purpose:", "Use when:", "Preconditions", "Result:", "Next actions:"):
            assert marker in text, f"{name} missing {marker!r}"


async def test_truthful_phrases_present() -> None:
    descriptions = await _descriptions()
    for name, needle in _TRUTH.items():
        assert needle in descriptions[name], f"{name} missing truthful phrase {needle!r}"


async def test_examples_present() -> None:
    descriptions = await _descriptions()
    for name in _EXAMPLE_TOOLS:
        assert "Example:" in descriptions[name], f"{name} missing an example"


async def test_guarded_mutations_use_uniform_guard_sentence() -> None:
    descriptions = await _descriptions()
    # U6-D: the repeated 133-char guard sentence is compressed to a glossary token.
    guard = "Guarded mutation (see xlvba://glossary)."
    # A representative set of guarded mutations (U3-owned enum tools excluded).
    for name in [
        "sheet",
        "sheet_delete",
        "table",
        "name",
        "name_delete",
        "shape",
        "control",
        "range_write",
        "range_structural",
        "range_delete",
        "range_merge",
        "pivot_table_refresh",
        "pivot_cache_refresh",
    ]:
        assert guard in descriptions[name], f"{name} missing the uniform R-D3 guard token"


async def test_glossary_resource_registered_and_complete() -> None:
    resources = await server.mcp.list_resources()
    uris = {str(resource.uri) for resource in resources}
    assert "xlvba://glossary" in uris

    text = await server.glossary_resource()
    for term in _GLOSSARY_TERMS:
        assert term in text, f"glossary missing term {term!r}"


async def test_environment_and_glossary_constants_exist() -> None:
    assert "Purpose:" in server._DESC_ENVIRONMENT
    assert server._DESC_RES_GLOSSARY.strip()
    # U2 registers the environment tool as a read-only, non-COM capability.
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    assert "xlvba_environment" in tools
    annotations = tools["xlvba_environment"].annotations
    assert annotations is not None
    assert annotations.read_only_hint is True