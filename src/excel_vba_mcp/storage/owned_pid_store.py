"""Startup-reconcile persistence for the owned Excel pid (Block A).

A tiny, pure-filesystem store that remembers the **owned** ``EXCEL.EXE`` pid the
server launched, so the next server start can detect a leftover live instance and
**refuse to launch a second Excel** (report_07 §5, plan_02 Work item A #4). It
never touches Excel / COM and holds no proxies.

Guarantees:

* **Atomic.** The state document is written to a temp file in the destination
  directory, flushed / ``fsync``-ed and then moved into place with ``os.replace``
  (reusing the run-repository helper); a crash never leaves a half-written file.
* **Truthful reads.** :meth:`OwnedPidStore.read` returns ``None`` for a missing /
  unreadable / malformed document — the caller then proceeds with a normal start.
* **Fail-safe clear.** :meth:`OwnedPidStore.clear` removes the document (a dead
  prior pid must not block a fresh start).

The document is a single small JSON object ``{"owned_pid": <int>, ...}`` under the
work / artifact root; it is **not** a sync source and holds no workbook data.
"""

from __future__ import annotations

import contextlib
import json
from datetime import UTC, datetime
from pathlib import Path

from excel_vba_mcp.storage.run_repository import _atomic_write_json

__all__ = ["OwnedPidStore"]

_STATE_FILE_NAME = "owned_pid.json"


class OwnedPidStore:
    """Persist the last launched owned Excel pid for startup reconcile (Block A).

    Args:
        root: Directory that holds the ``owned_pid.json`` state document (the
            configured work / artifact root). Created on first write.
    """

    def __init__(self, root: Path) -> None:
        self._root = Path(root)
        self._path = self._root / _STATE_FILE_NAME

    @property
    def path(self) -> Path:
        """The state document path (for diagnostics / tests)."""

        return self._path

    def write(self, owned_pid: int) -> None:
        """Atomically persist ``owned_pid`` as the last launched owned pid."""

        payload = {
            "owned_pid": int(owned_pid),
            "updated_at": datetime.now(UTC).isoformat(),
        }
        _atomic_write_json(self._path, json.dumps(payload).encode("utf-8"))

    def read(self) -> int | None:
        """Return the persisted owned pid, or ``None`` when absent / unreadable."""

        if not self._path.is_file():
            return None
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        if not isinstance(raw, dict):
            return None
        pid = raw.get("owned_pid")
        if isinstance(pid, bool) or not isinstance(pid, int):
            return None
        return pid if pid > 0 else None

    def clear(self) -> None:
        """Remove the state document (best-effort; a dead pid must not block start)."""

        with contextlib.suppress(OSError):
            self._path.unlink(missing_ok=True)
