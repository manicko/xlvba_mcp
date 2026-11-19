"""Storage layer: immutable snapshot persistence (task_014).

The storage layer is pure filesystem and owns no COM / Excel access. See
:mod:`excel_vba_mcp.storage.snapshot_store`.
"""

from excel_vba_mcp.storage.snapshot_store import (
    SnapshotManifest,
    SnapshotMethod,
    SnapshotStore,
)

__all__ = ["SnapshotManifest", "SnapshotMethod", "SnapshotStore"]
