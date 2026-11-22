"""Unit tests for the FE-22 (finding F-5) error-taxonomy additions.

These tests never touch Excel/COM. They prove that:

* the new ``WORKBOOK_NOT_FOUND`` / ``REVISION_CONFLICT`` codes exist and their
  typed subclasses carry them as ``default_code`` (a), while the generic
  ``ExcelVbaMcpError.default_code = COM_BUSY`` fallback is preserved;
* an unknown ``workbook_id`` is a ``WORKBOOK_NOT_FOUND`` (not ``COM_BUSY``) —
  covered at the service level in ``test_session_service``;
* the VBA optimistic-concurrency (revision) path raises ``REVISION_CONFLICT``
  and a non-replaceable component raises ``PROTECTED_TARGET`` (c);
* a read-only resource surfaces a structured JSON error carrying a ``code``,
  matching the tool boundary, instead of a plain-text ``str(exc)`` (d).
"""

from __future__ import annotations

import json
from datetime import UTC, datetime
from typing import Any

import pytest
from mcp.server.mcpserver.exceptions import ResourceError

from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import (
    ErrorCode,
    ExcelVbaMcpError,
    ProtectedTargetError,
    RevisionConflictError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import VbaComponentSource
from excel_vba_mcp.domain.statuses import VbaComponentType

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _ts() -> datetime:
    return datetime(2026, 1, 1, 12, 0, 0, tzinfo=UTC)


# ---------------------------------------------------------------------------
# (a) New codes + subclasses
# ---------------------------------------------------------------------------
def test_new_error_codes_are_additive_members() -> None:
    values = {code.value for code in ErrorCode}
    assert "WORKBOOK_NOT_FOUND" in values
    assert "REVISION_CONFLICT" in values
    # The Stage-1 closed set is untouched (subset check).
    assert {
        "ACCESS_VBOM_DISABLED",
        "COM_BUSY",
        "EXCEL_HUNG",
        "SNAPSHOT_FAILED",
        "WORKBOOK_READ_ONLY",
        "PATH_NOT_ALLOWED",
    } <= values


def test_workbook_not_found_error_carries_code() -> None:
    error = WorkbookNotFoundError()
    assert error.code is ErrorCode.WORKBOOK_NOT_FOUND
    assert error.to_dict()["code"] == "WORKBOOK_NOT_FOUND"
    assert error.context == {}


def test_revision_conflict_error_carries_code() -> None:
    error = RevisionConflictError()
    assert error.code is ErrorCode.REVISION_CONFLICT
    assert error.to_dict()["code"] == "REVISION_CONFLICT"


def test_generic_fallback_default_code_is_still_com_busy() -> None:
    # F-5 constraint: the generic fallback must remain COM_BUSY.
    assert ExcelVbaMcpError.default_code is ErrorCode.COM_BUSY
    assert ExcelVbaMcpError("boom").code is ErrorCode.COM_BUSY


# ---------------------------------------------------------------------------
# (c) VBA revision conflict + non-replaceable component
# ---------------------------------------------------------------------------
def test_require_revision_raises_revision_conflict() -> None:
    with pytest.raises(RevisionConflictError) as exc_info:
        VBAProjectService._require_revision("aaa", "bbb", "modTest")
    error = exc_info.value
    assert error.code is ErrorCode.REVISION_CONFLICT
    assert error.context.get("reason") == "revision_conflict"
    assert error.context.get("component") == "modTest"


def test_require_revision_passes_when_hashes_match() -> None:
    # No raise when the hashes agree.
    VBAProjectService._require_revision("same", "same", "modTest")


def test_non_replaceable_component_raises_protected_target() -> None:
    source = VbaComponentSource(
        component_name="ThisWorkbook",
        component_type=VbaComponentType.DOCUMENT_MODULE,
        source="Option Explicit",
        line_count=1,
        source_hash="0" * 64,
        revision=1,
        created_at=_ts(),
        updated_at=_ts(),
    )
    with pytest.raises(ProtectedTargetError) as exc_info:
        VBAProjectService._require_replaceable(source)
    error = exc_info.value
    assert error.code is ErrorCode.PROTECTED_TARGET
    assert error.context.get("reason") == "non_replaceable_component"


# ---------------------------------------------------------------------------
# (d) Resources surface structured JSON errors with a code
# ---------------------------------------------------------------------------
class _FakeVbaService:
    """Minimal VBA service whose read_component raises a domain error."""

    async def read_component(self, workbook_id: str, component: str) -> Any:
        raise WorkbookNotFoundError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            context={"workbook_id": workbook_id},
        )


class _FakeRunner:
    """Minimal test runner whose get_run raises a domain error."""

    def get_run(self, run_id: str) -> Any:
        raise ExcelVbaMcpError(
            "run not found.",
            code=ErrorCode.SNAPSHOT_FAILED,
            context={"run_id": run_id},
        )


async def test_component_source_resource_surfaces_structured_json() -> None:
    import excel_vba_mcp.api.vba_tools as vba_tools
    import excel_vba_mcp.server as server

    vba_tools.set_vba_project_service(_FakeVbaService())  # type: ignore[arg-type]
    try:
        with pytest.raises(ResourceError) as exc_info:
            await server.mcp.read_resource("excel://workbooks/__missing__/vba/components/modTest")
    finally:
        vba_tools.set_vba_project_service(None)

    # The ResourceError message must be the structured JSON payload (not str(exc)).
    payload = json.loads(str(exc_info.value))
    assert payload["code"] == "WORKBOOK_NOT_FOUND"
    assert payload["context"] == {"workbook_id": "__missing__"}


async def test_test_run_resource_surfaces_structured_json() -> None:
    import excel_vba_mcp.api.test_tools as test_tools
    import excel_vba_mcp.server as server

    test_tools.set_test_runner(_FakeRunner())  # type: ignore[arg-type]
    try:
        with pytest.raises(ResourceError) as exc_info:
            await server.mcp.read_resource("excel://test-runs/__missing__")
    finally:
        test_tools.set_test_runner(None)

    payload = json.loads(str(exc_info.value))
    assert payload["code"] == "SNAPSHOT_FAILED"
    assert payload["context"] == {"run_id": "__missing__"}
