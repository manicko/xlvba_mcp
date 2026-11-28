"""Unit tests for the WP6 macro/test dispatch tools (plan B7).

These tests never touch COM / Excel. They install fake ``MacroService`` /
``TestRunner`` instances via the tool-module setters and assert that the
flat-signature ``macro`` / ``test`` dispatchers route to the existing handlers,
enforce per-action required parameters as structured :class:`ExcelVbaMcpError`
(option (a) — handler validation), and advertise the truthful annotation triples
from the WP0 matrix: ``macro`` is (F,F,T) idempotent-only, ``test`` is (F,T,F)
destructive, and ``macro_run`` stays standalone and destructive (distinct Kilo
permission — H1).
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import macro_tools, test_tools
from excel_vba_mcp.api.macro_tools import macro
from excel_vba_mcp.domain.errors import ExcelVbaMcpError

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class FakeMacroService:
    """Records calls and returns canned sentinels; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def list_macros(self, workbook_id: str) -> list[str]:
        self.calls.append(("list_macros", (workbook_id,)))
        return ["Module1.Main"]

    async def classify_background(
        self,
        workbook_id: str,
        procedure: str,
        background_safe: bool,
        expected_project_source_hash: str | None = None,
    ) -> str:
        self.calls.append(
            (
                "classify_background",
                (workbook_id, procedure, background_safe, expected_project_source_hash),
            )
        )
        return "classification"

    async def get_last_result(self, workbook_id: str) -> str | None:
        self.calls.append(("get_last_result", (workbook_id,)))
        return None


class FakeTestRunner:
    """Records calls and returns canned sentinels; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def run_suite(self, suite: Any) -> str:
        self.calls.append(("run_suite", (suite,)))
        return "run-result"

    def get_run(self, run_id: str) -> str:
        self.calls.append(("get_run", (run_id,)))
        return "persisted-run"


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _install_fakes() -> Any:
    macros = FakeMacroService()
    runner = FakeTestRunner()
    macro_tools.set_macro_service(macros)  # type: ignore[arg-type]
    test_tools.set_test_runner(runner)  # type: ignore[arg-type]
    yield macros, runner
    macro_tools.set_macro_service(None)
    test_tools.set_test_runner(None)


# ---------------------------------------------------------------------------
# macro: routing
# ---------------------------------------------------------------------------
async def test_macro_list_routes(_install_fakes: tuple[Any, Any]) -> None:
    macros, _ = _install_fakes
    listed = await macro("list", workbook_id="wb1")

    assert listed == ["Module1.Main"]
    verb, args = macros.calls[-1]
    assert verb == "list_macros"
    assert args == ("wb1",)


async def test_macro_get_last_result_routes(_install_fakes: tuple[Any, Any]) -> None:
    macros, _ = _install_fakes
    result = await macro("get-last-result", workbook_id="wb1")

    assert result is None
    verb, args = macros.calls[-1]
    assert verb == "get_last_result"
    assert args == ("wb1",)


async def test_macro_classify_background_routes(_install_fakes: tuple[Any, Any]) -> None:
    macros, _ = _install_fakes
    result = await macro(
        "classify-background",
        workbook_id="wb1",
        procedure="Module1.Recalculate",
        background_safe=True,
        expected_project_source_hash="abc",
    )

    assert result == "classification"
    verb, args = macros.calls[-1]
    assert verb == "classify_background"
    assert args == ("wb1", "Module1.Recalculate", True, "abc")


# ---------------------------------------------------------------------------
# macro: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_macro_classify_background_missing_background_safe_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await macro(
            "classify-background",
            workbook_id="wb1",
            procedure="Module1.Recalculate",
        )


async def test_macro_classify_background_missing_procedure_errors() -> None:
    # procedure is required by the target handler; an omitted value surfaces as
    # a structured error rather than a crash.
    with pytest.raises(ExcelVbaMcpError):
        await macro("classify-background", workbook_id="wb1", background_safe=True)


async def test_macro_missing_workbook_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await macro("list")


async def test_macro_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await macro("destroy", workbook_id="wb1")  # type: ignore[arg-type]


async def test_macro_run_is_not_a_macro_action() -> None:
    # macro_run is intentionally standalone (distinct Kilo permission); it must
    # NOT be reachable through the macro dispatcher.
    with pytest.raises(ExcelVbaMcpError):
        await macro("run", workbook_id="wb1", procedure="Module1.Main")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# test: routing
# ---------------------------------------------------------------------------
async def test_test_run_suite_routes(_install_fakes: tuple[Any, Any]) -> None:
    _, runner = _install_fakes
    suite = {"suite": "smoke", "workbook_id": "wb1", "cases": [{"id": "c1", "steps": []}]}
    result = await test_tools.test("run", suite=suite)

    assert result == "run-result"
    verb, args = runner.calls[-1]
    assert verb == "run_suite"
    # The dict is validated into a TestSuiteSpec before dispatch.
    assert args[0].suite == "smoke"


async def test_test_get_run_routes(_install_fakes: tuple[Any, Any]) -> None:
    _, runner = _install_fakes
    result = await test_tools.test("get-run", run_id="run-1")

    assert result == "persisted-run"
    verb, args = runner.calls[-1]
    assert verb == "get_run"
    assert args == ("run-1",)


# ---------------------------------------------------------------------------
# test: per-action required-parameter validation (structured errors)
# ---------------------------------------------------------------------------
async def test_test_run_suite_missing_suite_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await test_tools.test("run")


async def test_test_get_run_missing_run_id_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await test_tools.test("get-run")


async def test_test_unknown_action_errors() -> None:
    with pytest.raises(ExcelVbaMcpError):
        await test_tools.test("destroy")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Registration + annotation triples (WP0 truth matrix)
# ---------------------------------------------------------------------------
async def test_macro_and_test_registered_old_names_removed() -> None:
    names = {tool.name for tool in await server.mcp.list_tools()}
    for present in ("macro", "macro_run", "test_suite"):
        assert present in names
    for removed in (
        "macro_list",
        "macro_classify_background",
        "macro_get_last_result",
        "test_suite_run",
        "test_get_run",
    ):
        assert removed not in names


async def test_macro_run_is_standalone_and_destructive() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    # macro_run keeps its own distinct registered name (H1: distinct permission).
    assert "macro_run" in tools
    ann = tools["macro_run"].annotations
    assert ann is not None
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is not True


async def test_macro_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["macro"].annotations
    assert ann is not None
    # (F, F, T): not read-only, not destructive, idempotent only.
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is not True
    assert ann.idempotent_hint is True


async def test_test_annotation_triple() -> None:
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}
    ann = tools["test_suite"].annotations
    assert ann is not None
    # (F, T, F): not read-only, destructive, not idempotent.
    assert ann.read_only_hint is not True
    assert ann.destructive_hint is True
    assert ann.idempotent_hint is not True
