"""Documentation consistency guard (WP7).

Static, read-only: imports the server (registers tools at import time) and
checks that the human-facing docs stay aligned with the live MCP surface. It
never touches Excel/COM and never starts the worker.

This is a small consistency check, not a documentation-generation framework.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import excel_vba_mcp.server as server

pytestmark = pytest.mark.anyio

_ROOT = Path(__file__).resolve().parents[2]
_README = _ROOT / "README.md"
_TOOLS_REF = _ROOT / "docs" / "05-mcp-tools-reference.md"

# The six renamed tools (canonical -> retired old public name).
_RENAMED_OLD = [
    "excel_save_as",
    "workbook_file_read",
]
# Clearly retired pre-consolidation per-verb names that must not reappear in the
# refreshed active docs (the docs/02 migration table is intentionally excluded).
_RETIRED_NAMES = [
    "excel_open",
    "workbook_get_model",
    "sheets_list",
    "names_list",
    "tables_list",
    "pivot_tables_list",
    "range_get",
    "shapes_list",
    "vba_project_info",
    "macro_list",
    "test_suite_run",
    "excel_list_workbooks",
    "excel_get_status",
    "run-suite",
]

# Active docs that must be free of retired names (excludes docs/ru mirrors and the
# docs/02 historical migration table).
_ACTIVE_DOCS = [
    "01-overview.md",
    "04-architecture.md",
    "05-mcp-tools-reference.md",
    "06-excel-capabilities.md",
    "07-workflows-and-examples.md",
    "08-security.md",
    "09-limitations-and-compatibility.md",
    "10-troubleshooting.md",
]


async def _tool_names() -> list[str]:
    return [tool.name for tool in await server.mcp.list_tools()]


async def test_tool_count_is_43() -> None:
    assert len(await _tool_names()) == 43


async def test_every_tool_documented_in_reference_and_readme() -> None:
    names = await _tool_names()
    ref = _TOOLS_REF.read_text(encoding="utf-8")
    readme = _README.read_text(encoding="utf-8")
    missing_ref = [n for n in names if n not in ref]
    missing_readme = [n for n in names if n not in readme]
    assert not missing_ref, f"tools missing from docs/05: {missing_ref}"
    assert not missing_readme, f"tools missing from README catalog: {missing_readme}"


def test_readme_states_43_tools() -> None:
    assert "43 MCP tools" in _README.read_text(encoding="utf-8")


def test_pivot_result_read_action_documented() -> None:
    assert "result-read" in _TOOLS_REF.read_text(encoding="utf-8")


def test_no_retired_names_in_active_docs() -> None:
    offenders: list[str] = []
    for rel in _ACTIVE_DOCS:
        text = (_ROOT / "docs" / rel).read_text(encoding="utf-8")
        for name in _RETIRED_NAMES + _RENAMED_OLD:
            if name in text:
                offenders.append(f"{rel}: {name}")
    assert not offenders, f"retired names still in active docs: {offenders}"


def test_no_stale_table_cap_2000_in_active_docs() -> None:
    # The table read cap is 50,000 (config-driven); a "2,000 cells" claim is stale.
    offenders: list[str] = []
    for rel in _ACTIVE_DOCS:
        text = (_ROOT / "docs" / rel).read_text(encoding="utf-8")
        if "2,000 cells" in text or "2000 cells" in text:
            offenders.append(rel)
    assert not offenders, f"stale 2,000-cell table cap in: {offenders}"