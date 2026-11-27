"""Unit tests for the WP9 VBA read dispatch tool (plan B10).

These tests never touch COM / Excel. They install a fake ``VBAProjectService``
via the ``vba_tools`` setter and a fake compile service via the
``compile_tools`` setter, then assert that the flat-signature ``vba_read``
dispatcher routes to the existing per-verb handlers, enforces per-action
required parameters as structured :class:`ExcelVbaMcpError` (option (a) —
handler validation), and advertises the truthful annotation triples from the
WP0 matrix:

* ``vba_read`` is (T,F,T) — read-only + idempotent.
* ``vba_replace_component_source`` is (F,T,F) — destructive only, NOT
  idempotent (R-Q3: re-call with the same hash fails the revision guard).
* ``vba_insert_code`` is (F,T,F) — destructive only.
* ``vba_export_project`` is (F,F,T) — idempotent, not destructive.
* ``vba_compile_diagnostics`` is (F,F,T) — idempotent, NOT read-only, not
  destructive; keeps the "does NOT compile" phrase.

The four standalone tools keep their names/signatures and hash guards.
"""

from __future__ import annotations

from typing import Any

import pytest

import excel_vba_mcp.server as server
from excel_vba_mcp.api import compile_tools, vba_tools
from excel_vba_mcp.api.vba_tools import (
    vba_export_project,
    vba_insert_code,
    vba_read,
    vba_replace_component_source,
)
from excel_vba_mcp.application.compile_service import CompileDiagnostics
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaProcedureInfo,
    VbaProjectInfo,
    VbaReferenceInfo,
    VbaSearchMatch,
)

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
def _project_info() -> VbaProjectInfo:
    return VbaProjectInfo(available=True, name="VBAProject")


def _component_info() -> VbaComponentInfo:
    return VbaComponentInfo(component_name="Module1", component_type="standard_module")


def _procedure_info() -> VbaProcedureInfo:
    return VbaProcedureInfo(
        component_name="Module1",
        name="Main",
        procedure_type="sub",
        start_line=1,
        end_line=2,
    )


def _component_source() -> VbaComponentSource:
    return VbaComponentSource(
        created_at="2026-01-01T00:00:00Z",
        updated_at="2026-01-01T00:00:00Z",
        component_name="Module1",
        component_type="standard_module",
        source="Sub Main()\nEnd Sub",
        line_count=2,
        source_hash="hash",
        revision=1,
    )


def _reference_info() -> VbaReferenceInfo:
    return VbaReferenceInfo(name="stdole", status="available")


def _search_match() -> VbaSearchMatch:
    return VbaSearchMatch(
        component_name="Module1",
        component_type="standard_module",
        line_number=1,
        line_text="Sub Main()",
    )


class FakeVbaProjectService:
    """Records calls and returns canned DTOs; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def get_project_info(self, workbook_id: str) -> VbaProjectInfo:
        self.calls.append(("get_project_info", (workbook_id,)))
        return _project_info()

    async def list_components(self, workbook_id: str) -> list[VbaComponentInfo]:
        self.calls.append(("list_components", (workbook_id,)))
        return [_component_info()]

    async def list_procedures(
        self, workbook_id: str, component: str
    ) -> list[VbaProcedureInfo]:
        self.calls.append(("list_procedures", (workbook_id, component)))
        return [_procedure_info()]

    async def read_component(
        self, workbook_id: str, component: str
    ) -> VbaComponentSource:
        self.calls.append(("read_component", (workbook_id, component)))
        return _component_source()

    async def get_references(self, workbook_id: str) -> list[VbaReferenceInfo]:
        self.calls.append(("get_references", (workbook_id,)))
        return [_reference_info()]

    async def search_code(
        self, workbook_id: str, pattern: str, regex: bool = False
    ) -> list[VbaSearchMatch]:
        self.calls.append(("search_code", (workbook_id, pattern, regex)))
        return [_search_match()]

    async def replace_component_source(
        self,
        workbook_id: str,
        component: str,
        new_source: str,
        expected_source_hash: str,
    ) -> Any:
        self.calls.append(
            ("replace_component_source", (workbook_id, component, expected_source_hash))
        )
        return object()

    async def insert_code(
        self,
        workbook_id: str,
        component: str,
        code: str,
        position: str,
        expected_source_hash: str,
    ) -> Any:
        self.calls.append(
            ("insert_code", (workbook_id, component, position, expected_source_hash))
        )
        return object()

    async def export_project(self, workbook_id: str, out_dir: str | None = None) -> Any:
        self.calls.append(("export_project", (workbook_id, out_dir)))
        return object()


class FakeCompileService:
    """Records get_last_diagnostics calls; no COM."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []

    async def get_last_diagnostics(self, workbook_id: str) -> CompileDiagnostics:
        self.calls.append(("get_last_diagnostics", (workbook_id,)))
        return CompileDiagnostics()


@pytest.fixture
def fake_services() -> tuple[FakeVbaProjectService, FakeCompileService]:
    vba = FakeVbaProjectService()
    compile_svc = FakeCompileService()
    vba_tools.set_vba_project_service(vba)  # type: ignore[arg-type]
    compile_tools.set_compile_service(compile_svc)  # type: ignore[arg-type]
    yield vba, compile_svc
    vba_tools.set_vba_project_service(None)
    compile_tools.set_compile_service(None)


# ---------------------------------------------------------------------------
# Routing (all seven actions)
# ---------------------------------------------------------------------------
async def test_vba_read_routes_all_actions(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    vba, compile_svc = fake_services

    assert isinstance(await vba_read("project-info", workbook_id="wb1"), VbaProjectInfo)
    assert await vba_read("list-components", workbook_id="wb1")
    assert await vba_read("list-procedures", workbook_id="wb1", component="Module1")
    assert isinstance(
        await vba_read("read-component", workbook_id="wb1", component="Module1"),
        VbaComponentSource,
    )
    assert await vba_read("get-references", workbook_id="wb1")
    assert await vba_read("search-code", workbook_id="wb1", pattern="Main")
    assert isinstance(
        await vba_read("get-compile-diagnostics", workbook_id="wb1"),
        CompileDiagnostics,
    )

    assert [c[0] for c in vba.calls] == [
        "get_project_info",
        "list_components",
        "list_procedures",
        "read_component",
        "get_references",
        "search_code",
    ]
    assert compile_svc.calls == [("get_last_diagnostics", ("wb1",))]


async def test_vba_read_search_code_forwards_regex(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    vba, _ = fake_services
    await vba_read("search-code", workbook_id="wb1", pattern="Sub.*", regex=True)
    assert vba.calls[-1] == ("search_code", ("wb1", "Sub.*", True))


# ---------------------------------------------------------------------------
# Per-action required-parameter validation (option (a))
# ---------------------------------------------------------------------------
async def test_vba_read_missing_required_params(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await vba_read("read-component", workbook_id="wb1")  # missing component
    with pytest.raises(ExcelVbaMcpError):
        await vba_read("list-procedures", workbook_id="wb1")  # missing component
    with pytest.raises(ExcelVbaMcpError):
        await vba_read("search-code", workbook_id="wb1")  # missing pattern
    with pytest.raises(ExcelVbaMcpError):
        await vba_read("project-info")  # missing workbook_id


async def test_vba_read_unknown_action(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await vba_read("bogus", workbook_id="wb1")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# Standalone mutation hash guards
# ---------------------------------------------------------------------------
async def test_replace_component_source_requires_hash(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await vba_replace_component_source(
            "wb1", "Module1", "Sub Main()\nEnd Sub", ""
        )


async def test_insert_code_requires_hash(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    with pytest.raises(ExcelVbaMcpError):
        await vba_insert_code("wb1", "Module1", "code", "declarations", "")


async def test_export_project_standalone_callable(
    fake_services: tuple[FakeVbaProjectService, FakeCompileService],
) -> None:
    vba, _ = fake_services
    await vba_export_project("wb1")
    assert vba.calls[-1] == ("export_project", ("wb1", None))


# ---------------------------------------------------------------------------
# Annotation triples (from the WP0 matrix) + registration surface
# ---------------------------------------------------------------------------
async def test_vba_annotation_triples() -> None:
    tools = {t.name: t for t in await server.mcp.list_tools()}

    vba_read_ann = tools["vba_read"].annotations
    assert vba_read_ann is not None
    assert vba_read_ann.read_only_hint is True
    assert vba_read_ann.idempotent_hint is True
    assert vba_read_ann.destructive_hint is not True

    replace_ann = tools["vba_replace_component_source"].annotations
    assert replace_ann is not None
    assert replace_ann.destructive_hint is True
    assert replace_ann.idempotent_hint is not True  # R-Q3
    assert replace_ann.read_only_hint is not True

    insert_ann = tools["vba_insert_code"].annotations
    assert insert_ann is not None
    assert insert_ann.destructive_hint is True
    assert insert_ann.idempotent_hint is not True
    assert insert_ann.read_only_hint is not True

    export_ann = tools["vba_export_project"].annotations
    assert export_ann is not None
    assert export_ann.idempotent_hint is not True
    assert export_ann.destructive_hint is not True

    compile_ann = tools["vba_compile_diagnostics"].annotations
    assert compile_ann is not None
    assert compile_ann.read_only_hint is not True
    assert compile_ann.destructive_hint is not True
    assert compile_ann.idempotent_hint is not True


async def test_old_vba_read_names_removed() -> None:
    names = {t.name for t in await server.mcp.list_tools()}
    for old in (
        "vba_project_info",
        "vba_list_components",
        "vba_list_procedures",
        "vba_read_component",
        "vba_get_references",
        "vba_search_code",
        "vba_get_compile_diagnostics",
    ):
        assert old not in names
    # Standalone tools remain registered under their unchanged names.
    for kept in (
        "vba_read",
        "vba_replace_component_source",
        "vba_insert_code",
        "vba_export_project",
        "vba_compile_diagnostics",
    ):
        assert kept in names


async def test_compile_diagnostics_keeps_truthful_phrase() -> None:
    tools = {t.name: t for t in await server.mcp.list_tools()}
    assert "does NOT compile" in (tools["vba_compile_diagnostics"].description or "")
