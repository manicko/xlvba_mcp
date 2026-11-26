"""Snapshot + declarative test-run MCP tools (task_014, task_035; spec §6.4, §9).

Thin, async MCP tool handlers for immutable workbook snapshots
(``snapshot_create``, ``snapshot_list``, ``snapshot_restore``, ``snapshot_diff``)
and for the bounded declarative test runner (``test_suite_run``,
``test_get_run``). They are deliberately **thin**: each
validates its arguments, delegates to an application service and returns a domain
DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` / ``win32com`` import
lives in this module.

``test_suite_run`` is the single canonical test entry point. It accepts the suite
either as a :class:`~excel_vba_mcp.domain.test_spec.TestSuiteSpec` instance or as a
plain JSON-style mapping (the shape an MCP caller sends) and validates it here, so
the closed action / assertion vocabularies are checked by Pydantic at the protocol
boundary — before any Excel call — and an unknown Range/Pivot action or assertion
is rejected up front. The runner then isolates every case on a disposable copy of
one baseline snapshot.

The handlers need configured services. Like
:mod:`excel_vba_mcp.api.workbook_tools`, the active services are published here via
``set_snapshot_service`` / ``set_test_runner`` (called from the server lifespan) and
read via ``get_snapshot_service`` / ``get_test_runner``.
"""

from __future__ import annotations

from functools import partial
from typing import Any

from pydantic import ValidationError

from excel_vba_mcp.api.job_tools import start_long_op
from excel_vba_mcp.application.job_registry import RunningJob
from excel_vba_mcp.application.snapshot_service import SnapshotDiff, SnapshotService
from excel_vba_mcp.application.test_runner import TestRunner
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import TestRunResult, WorkbookInfo
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.domain.test_spec import TestSuiteSpec
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest

# The active snapshot service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_snapshot_service: SnapshotService | None = None

# The active declarative test runner, published by the server lifespan.
_test_runner: TestRunner | None = None


def set_snapshot_service(service: SnapshotService | None) -> None:
    """Publish (or clear) the active snapshot service used by the tools."""

    global _snapshot_service
    _snapshot_service = service


def get_snapshot_service() -> SnapshotService:
    """Return the active snapshot service or raise if it is not configured."""

    if _snapshot_service is None:
        raise ExcelVbaMcpError(
            "The snapshot service is not initialized.",
            context={"tool": "test_tools"},
        )
    return _snapshot_service


def set_test_runner(runner: TestRunner | None) -> None:
    """Publish (or clear) the active declarative test runner."""

    global _test_runner
    _test_runner = runner


def get_test_runner() -> TestRunner:
    """Return the active test runner or raise if it is not configured."""

    if _test_runner is None:
        raise ExcelVbaMcpError(
            "The test runner is not initialized.",
            context={"tool": "test_tools"},
        )
    return _test_runner


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


async def snapshot_create(workbook_id: str) -> SnapshotManifest:
    """Create an immutable snapshot of an open (or known) workbook.

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        The :class:`SnapshotManifest` describing the new snapshot.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_snapshot_service()
    return await service.create_snapshot(workbook_id)


async def snapshot_list() -> list[SnapshotManifest]:
    """List every snapshot and its manifest metadata."""

    service = get_snapshot_service()
    return await service.list_snapshots()


async def snapshot_restore(
    snapshot_id: str,
    workbook_id: str | None = None,
    *,
    confirm: bool = False,
) -> WorkbookInfo:
    """Restore a working session from an immutable snapshot.

    Closes the current working copy without saving (when ``workbook_id`` is
    given) and opens a fresh working copy from the snapshot baseline.

    This is destructive to the current working copy, so it is **gated by
    ``confirm``**: a missing or ``False`` ``confirm`` raises
    ``CONFIRMATION_REQUIRED`` **before** any close/termination and the snapshot
    service is never called. The snapshot restores files only — it does **not**
    undo external macro side effects (database writes, emails, network calls,
    files, printing).

    Args:
        snapshot_id: The snapshot to restore.
        workbook_id: The currently-open working copy to discard (no save).
        confirm: Must be ``True`` to proceed; otherwise the call is rejected
            with ``CONFIRMATION_REQUIRED`` before any destructive action.

    Returns:
        A :class:`WorkbookInfo` describing the freshly opened working copy.
    """

    _require_non_empty(snapshot_id, "snapshot_id")
    if workbook_id is not None:
        _require_non_empty(workbook_id, "workbook_id")
    if not confirm:
        raise ExcelVbaMcpError(
            "snapshot_restore requires confirm=true: it discards the current "
            "working copy without saving and does NOT undo external macro side "
            "effects.",
            code=ErrorCode.CONFIRMATION_REQUIRED,
            context={
                "tool": "snapshot_restore",
                "discard": (
                    "the current working copy (unsaved changes) when workbook_id "
                    "is given"
                ),
            },
        )
    service = get_snapshot_service()
    return await service.restore_snapshot(snapshot_id, workbook_id=workbook_id)


async def snapshot_diff(snapshot_id_a: str, snapshot_id_b: str) -> SnapshotDiff:
    """Compare two snapshots by manifest metadata and file hash.

    Args:
        snapshot_id_a: First snapshot id.
        snapshot_id_b: Second snapshot id.

    Returns:
        A :class:`~excel_vba_mcp.application.snapshot_service.SnapshotDiff`.
    """

    _require_non_empty(snapshot_id_a, "snapshot_id_a")
    _require_non_empty(snapshot_id_b, "snapshot_id_b")
    service = get_snapshot_service()
    return await service.diff_snapshots(snapshot_id_a, snapshot_id_b)


async def _run_suite_blocking(suite: TestSuiteSpec) -> TestRunResult:
    """Run a validated declarative suite on disposable copies of a workbook.

    This is the single blocking execution core shared by :func:`test_suite_run`
    (both the direct and the ``background=True`` paths). It is private: the only
    public test entry point is :func:`test_suite_run`.

    Every case starts from its own disposable working copy of one immutable
    baseline snapshot, so one case's failure cannot contaminate the next. Steps
    are dispatched to the existing application services; a failed compile step
    blocks the case's macro steps. External macro side effects are **not** rolled
    back by the snapshot (spec §9.4).

    Args:
        suite: The already-validated declarative suite. Actions and assertions are
            restricted to the closed vocabularies of
            :class:`~excel_vba_mcp.domain.test_spec.TestAction` /
            :class:`~excel_vba_mcp.domain.test_spec.AssertionType`.

    Returns:
        A :class:`~excel_vba_mcp.domain.models.TestRunResult` with per-case and
        per-step statuses. The run is persisted and can be read back with
        :func:`test_get_run`.
    """

    runner = get_test_runner()
    return await runner.run_suite(suite)


async def test_suite_run(
    suite: TestSuiteSpec | dict[str, Any],
    *,
    background: bool = False,
) -> TestRunResult | RunningJob:
    """Validate and run a declarative suite (task_053 explicit entry point).

    Accepts the suite either as a
    :class:`~excel_vba_mcp.domain.test_spec.TestSuiteSpec` or as a plain
    JSON-style mapping (the shape an MCP caller sends). The mapping is validated
    here with ``extra="forbid"`` and the closed Range/Pivot action / assertion
    vocabularies, so an unknown action or assertion — or an unsupported payload
    such as arbitrary code, a shell command or an unbounded COM expression — is
    rejected **before** any Excel call. The run then uses one immutable baseline
    snapshot, one disposable working copy per case, steps dispatched to the
    existing application services, and external macro side effects not rolled back.

    Args:
        suite: The declarative suite document (or an already-built
            :class:`TestSuiteSpec`).
        background: When ``True`` (Block F), schedule the run on the job registry
            and return ``{status: "running", job_id, poll_after_s}`` promptly
            instead of blocking; retrieve the :class:`TestRunResult` with
            ``job_get``. Default ``False`` (unchanged blocking behaviour).

    Returns:
        A :class:`~excel_vba_mcp.domain.models.TestRunResult` with per-case and
        per-step statuses, or a :class:`RunningJob` handle when
        ``background=True``. The run is persisted and can be read back with
        :func:`test_get_run`.

    Raises:
        ExcelVbaMcpError: With ``code='INVALID_RANGE'`` when the document does
            not satisfy the closed schema.
    """

    spec = suite if isinstance(suite, TestSuiteSpec) else _parse_suite(suite)
    if background:
        return start_long_op(
            partial(_run_suite_blocking, spec),
            timeout=600.0,
            section="test",
            op_class=OperationClass.EXECUTING,
        )
    return await _run_suite_blocking(spec)


def _parse_suite(document: dict[str, Any]) -> TestSuiteSpec:
    """Validate a plain suite mapping against the closed schema (pre-Excel)."""

    try:
        return TestSuiteSpec.model_validate(document)
    except ValidationError as exc:
        raise ExcelVbaMcpError(
            "The test suite is not a valid bounded declarative suite; "
            "unknown actions/assertions and non-scalar payloads are rejected "
            "before any Excel call.",
            code=ErrorCode.INVALID_RANGE,
            context={"reason": "invalid_test_suite", "validation": exc.errors()},
        ) from exc


async def test_get_run(run_id: str) -> TestRunResult:
    """Return a previously persisted test run.

    Args:
        run_id: The ``test_run_id`` reported by :func:`test_suite_run`.

    Returns:
        The persisted :class:`~excel_vba_mcp.domain.models.TestRunResult`.
    """

    _require_non_empty(run_id, "run_id")
    runner = get_test_runner()
    return runner.get_run(run_id)


__all__ = [
    "get_snapshot_service",
    "get_test_runner",
    "set_snapshot_service",
    "set_test_runner",
    "snapshot_create",
    "snapshot_diff",
    "snapshot_list",
    "snapshot_restore",
    "test_get_run",
    "test_suite_run",
]
