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
    "excel_open",
    "excel_close",
    "excel_save",
    "excel_get_status",
    "range_write_values",
    "range_read_values",
    "range_write_formulas",
    "macro_run",
    "macro_classify_background",
    "table_update_rows",
    "table_delete_rows",
    "table_append_rows",
    "name_create",
    "name_update",
    "sheet_add",
    "sheet_copy",
    "snapshot_restore",
    "excel_recover",
    "job_cancel",
    "job_get",
    "vba_compile",
    "vba_replace_component_source",
    "vba_insert_code",
    "test_run_suite",
    "test_suite_run",
]

# Tools that must carry exactly one request example (R-D8).
_EXAMPLE_TOOLS = [
    "range_write_values",
    "table_update_rows",
    "vba_replace_component_source",
    "test_suite_run",
]

# Truthful, non-negotiable phrases (R-D1).
_TRUTH = {
    "vba_compile": "does NOT compile",
    "macro_classify_background": "never grants run permission",
    "job_cancel": "intent-only; never interrupts",
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
]


async def _descriptions() -> dict[str, str]:
    tools = await server.mcp.list_tools()
    return {tool.name: (tool.description or "") for tool in tools}


async def test_tool_count_is_92_exposed() -> None:
    tools = await server.mcp.list_tools()
    # U2 adds xlvba_environment (91 -> 92).
    assert len(tools) == 92


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
    guard = (
        "Guarded mutation: writable session + snapshot + approval "
        "(full-access default; restricted requires approval) + postcondition re-read."
    )
    # A representative set of guarded mutations (U3-owned enum tools excluded).
    for name in [
        "sheet_rename",
        "sheet_delete",
        "sheet_set_visibility",
        "table_add_column",
        "table_create",
        "table_resize",
        "name_delete",
        "shape_set_on_action",
        "control_set_on_action",
        "range_set_format",
        "range_copy",
        "range_insert",
        "range_delete",
        "range_merge",
        "range_unmerge",
        "range_set_validation",
        "pivot_table_refresh",
        "pivot_cache_refresh",
    ]:
        assert guard in descriptions[name], f"{name} missing the uniform R-D3 guard sentence"


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