"""Tool-rename contract tests for the MCP naming refactor.

Static + hermetic: imports the server (registers tools at import time) and the
tool modules, and installs a fake ``TestRunner``. It never touches Excel/COM and
never starts the worker.
"""

from __future__ import annotations

from typing import Any, get_args, get_type_hints

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import name_tools, ooxml_tools, test_tools, workbook_tools
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError

pytestmark = pytest.mark.anyio

# canonical (new) name -> retired (old) public name
_RENAMED = {
    "workbook_save_copy": "excel_save_as",
    "test_suite": "test",
    "defined_name": "name",
    "defined_name_read": "name_read",
    "defined_name_delete": "name_delete",
    "workbook_snapshot_read": "workbook_file_read",
}
_OLD_NAMES = set(_RENAMED.values())

_EXPECTED_TRIPLES = {
    "workbook_save_copy": (False, True, False),
    "test_suite": (False, True, False),
    "defined_name": (False, False, True),
    "defined_name_read": (True, False, True),
    "defined_name_delete": (False, True, False),
    "workbook_snapshot_read": (True, False, True),
}


async def _tools() -> dict[str, Any]:
    return {tool.name: tool for tool in await server.mcp.list_tools()}


def _action_enum(func: Any) -> set[str]:
    return set(get_args(get_type_hints(func)["action"]))


async def test_all_six_canonical_names_registered() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for canonical in _RENAMED:
        assert canonical in names, f"missing canonical tool {canonical}"


async def test_all_six_old_names_absent() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for old in _OLD_NAMES:
        assert old not in names, f"retired tool name still registered: {old}"


async def test_no_duplicate_names_and_count_is_43() -> None:
    names = [tool.name for tool in await server.mcp.list_tools()]
    assert len(names) == len(set(names)), "duplicate tool names registered"
    assert len(names) == 43


async def test_renamed_tools_map_to_expected_handlers() -> None:
    assert server.excel_save_as is workbook_tools.excel_save_as
    assert server.test is test_tools.test
    assert server.name is name_tools.name
    assert server.name_read is name_tools.name_read
    assert server.name_delete is name_tools.name_delete
    assert server.workbook_file_read is ooxml_tools.workbook_file_read


async def test_renamed_tools_have_description_and_schema() -> None:
    tools = await _tools()
    for canonical in _RENAMED:
        tool = tools[canonical]
        assert tool.description and tool.description.strip(), f"{canonical} empty description"
        assert getattr(tool, "input_schema", None), f"{canonical} missing input schema"


async def test_renamed_tools_annotation_triples() -> None:
    tools = await _tools()
    for canonical, expected in _EXPECTED_TRIPLES.items():
        ann = tools[canonical].annotations
        assert ann is not None, f"{canonical} has no annotations"
        actual = (ann.read_only_hint, ann.destructive_hint, ann.idempotent_hint)
        assert actual == expected, f"{canonical} annotations {actual} != {expected}"


def test_test_suite_action_enum_is_run_get_run() -> None:
    actions = _action_enum(test_tools.test)
    assert actions == {"run", "get-run"}
    assert "run-suite" not in actions


def test_defined_name_action_enum_unchanged() -> None:
    assert _action_enum(name_tools.name) == {"create", "update"}


def test_defined_name_read_action_enum_unchanged() -> None:
    assert _action_enum(name_tools.name_read) == {"list", "get"}


def test_workbook_snapshot_read_action_enum_unchanged() -> None:
    assert _action_enum(ooxml_tools.workbook_file_read) == {
        "metadata",
        "sheet-rows",
        "pivot-cache-records",
    }


async def test_run_suite_action_rejected_with_invalid_range() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc:
        await test_tools.test("run-suite")  # type: ignore[arg-type]
    assert exc.value.code == ErrorCode.INVALID_RANGE


class _FakeRunner:
    def __init__(self) -> None:
        self.calls: list[tuple[str, Any]] = []

    async def run_suite(self, suite: Any) -> str:
        self.calls.append(("run_suite", suite))
        return "run-result"

    def get_run(self, run_id: str) -> str:
        self.calls.append(("get_run", run_id))
        return "persisted-run"


@pytest.fixture
def runner() -> Any:
    fake = _FakeRunner()
    test_tools.set_test_runner(fake)  # type: ignore[arg-type]
    yield fake
    test_tools.set_test_runner(None)


async def test_run_dispatches_to_run_path(runner: _FakeRunner) -> None:
    suite = {"suite": "s", "workbook_id": "wb1", "cases": [{"id": "c1", "steps": []}]}
    result = await test_tools.test("run", suite=suite)
    assert result == "run-result"
    assert runner.calls[-1][0] == "run_suite"


async def test_get_run_dispatches_to_get_run_path(runner: _FakeRunner) -> None:
    result = await test_tools.test("get-run", run_id="run-1")
    assert result == "persisted-run"
    assert runner.calls[-1] == ("get_run", "run-1")


async def test_run_missing_suite_is_invalid_range() -> None:
    with pytest.raises(ExcelVbaMcpError) as exc:
        await test_tools.test("run")
    assert exc.value.code == ErrorCode.INVALID_RANGE


async def test_get_run_missing_run_id_is_structured_error() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await test_tools.test("get-run")