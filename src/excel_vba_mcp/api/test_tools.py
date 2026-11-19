"""Snapshot MCP tools (task_014, spec §6.4).

Thin, async MCP tool handlers for immutable workbook snapshots:
``snapshot_create``, ``snapshot_list``, ``snapshot_restore`` and
``snapshot_diff``. They are deliberately **thin**: each validates its arguments,
delegates to :class:`~excel_vba_mcp.application.snapshot_service.SnapshotService`
and returns a domain DTO. They **never touch COM** (ARCH-009): no
``pythoncom`` / ``win32com`` import lives in this module.

The handlers need a configured :class:`SnapshotService`. Like
:mod:`excel_vba_mcp.api.workbook_tools`, the active service is published here via
:func:`set_snapshot_service` (called from the server lifespan) and read via
:func:`get_snapshot_service`.
"""

from __future__ import annotations

from excel_vba_mcp.application.snapshot_service import SnapshotDiff, SnapshotService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import WorkbookInfo
from excel_vba_mcp.storage.snapshot_store import SnapshotManifest

# The active snapshot service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_snapshot_service: SnapshotService | None = None


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
) -> WorkbookInfo:
    """Restore a working session from an immutable snapshot.

    Closes the current working copy without saving (when ``workbook_id`` is
    given) and opens a fresh working copy from the snapshot baseline.

    Args:
        snapshot_id: The snapshot to restore.
        workbook_id: The currently-open working copy to discard (no save).

    Returns:
        A :class:`WorkbookInfo` describing the freshly opened working copy.
    """

    _require_non_empty(snapshot_id, "snapshot_id")
    if workbook_id is not None:
        _require_non_empty(workbook_id, "workbook_id")
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


__all__ = [
    "get_snapshot_service",
    "set_snapshot_service",
    "snapshot_create",
    "snapshot_diff",
    "snapshot_list",
    "snapshot_restore",
]
