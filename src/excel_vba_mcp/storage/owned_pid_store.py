"""Startup-reconcile persistence for the owned Excel process (Block A, lifecycle).

A tiny, pure-filesystem store that remembers the **owned** ``EXCEL.EXE`` the server
launched (and the server/owner process that launched it) so the next server start
can reconcile precisely:

* **Owner alive + identity match** => refuse to launch a second Excel (the prior
  server still owns its instance).
* **Owner dead + owned Excel alive + identity verified** => the previous server was
  tree-killed / crashed and left an orphan; the caller terminates exactly that
  orphan (pid + create-time + image verified) and clears the marker.
* **Pid reuse** (create-time mismatch) => the recorded pid now belongs to a
  different process; never kill on pid-alone.

It never touches Excel / COM and holds no proxies; all Win32 identity probing is
done by the caller (the adapter) and passed in as plain values.

Guarantees:

* **Atomic.** The state document is written to a temp file in the destination
  directory, flushed / ``fsync``-ed and then moved into place with ``os.replace``
  (reusing the run-repository helper); a crash never leaves a half-written file.
* **Truthful reads.** :meth:`OwnedPidStore.read_record` returns ``None`` for a
  missing / unreadable / malformed document — the caller then proceeds with a
  normal start. A legacy ``{"owned_pid": <int>}`` document is still read (as a
  record with only the Excel pid) so an old marker is never silently ignored.
* **Fail-safe clear.** :meth:`OwnedPidStore.clear` removes the document (a dead
  prior pid must not block a fresh start).

The document is a single small JSON object under a **stable per-user** directory
(``%LOCALAPPDATA%\\excel-vba-mcp`` by default, or a configured non-volatile root);
it is **not** a sync source and holds no workbook data. ``SCHEMA_VERSION`` is
additive: a reader that does not know a field ignores it.
"""

from __future__ import annotations

import contextlib
import dataclasses
import json
import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excel_vba_mcp.storage.run_repository import _atomic_write_json

__all__ = [
    "SCHEMA_VERSION",
    "ExcelIdentity",
    "OwnedPidStore",
    "OwnedProcessRecord",
    "OwnerIdentity",
    "build_owned_record",
    "default_owned_pid_dir",
]

_STATE_FILE_NAME = "owned_pid.json"

#: Additive schema version. Bump only when the on-disk shape changes incompatibly;
#: readers ignore unknown fields so older/newer documents stay readable.
SCHEMA_VERSION = 1

#: Record states (the lifecycle phase the marker was written in).
STATE_RUNNING = "running"
STATE_RECOVERING = "recovering"


@dataclasses.dataclass(frozen=True)
class OwnerIdentity:
    """Identity of the server/owner process that launched the owned Excel."""

    pid: int
    create_time_filetime: int | None = None
    exe: str | None = None
    session_id: int | None = None


@dataclasses.dataclass(frozen=True)
class ExcelIdentity:
    """Identity of the owned Excel process (adds the image-verified flag)."""

    pid: int
    create_time_filetime: int | None = None
    exe: str | None = None
    session_id: int | None = None
    image_verified: bool = False


@dataclasses.dataclass(frozen=True)
class OwnedProcessRecord:
    """The full ownership marker persisted for startup reconcile."""

    owner: OwnerIdentity
    excel: ExcelIdentity
    state: str = STATE_RUNNING
    schema_version: int = SCHEMA_VERSION
    updated_at: str = ""

    def to_dict(self) -> dict[str, Any]:
        """Serialize to the on-disk JSON object (stable field order)."""

        return {
            "schema_version": self.schema_version,
            "owner": {
                "pid": self.owner.pid,
                "create_time_filetime": self.owner.create_time_filetime,
                "exe": self.owner.exe,
                "session_id": self.owner.session_id,
            },
            "excel": {
                "pid": self.excel.pid,
                "create_time_filetime": self.excel.create_time_filetime,
                "exe": self.excel.exe,
                "session_id": self.excel.session_id,
                "image_verified": self.excel.image_verified,
            },
            "state": self.state,
            "updated_at": self.updated_at or datetime.now(UTC).isoformat(),
        }

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> OwnedProcessRecord | None:
        """Parse a record from a JSON object; return ``None`` when unusable.

        Accepts both the current schema and the legacy ``{"owned_pid": <int>}``
        shape (mapped to an Excel-only record with no owner/create-time, so the
        reconcile treats it conservatively — never kill on pid-alone).
        """

        if not isinstance(raw, dict):
            return None
        # Legacy shape: {"owned_pid": <int>, "updated_at": ...}
        if "excel" not in raw and "owner" not in raw:
            legacy_pid = raw.get("owned_pid")
            if isinstance(legacy_pid, bool) or not isinstance(legacy_pid, int):
                return None
            if legacy_pid <= 0:
                return None
            return cls(
                owner=OwnerIdentity(pid=0),
                excel=ExcelIdentity(pid=legacy_pid),
                state=STATE_RUNNING,
                schema_version=SCHEMA_VERSION,
                updated_at=str(raw.get("updated_at", "")),
            )
        owner_raw = raw.get("owner")
        excel_raw = raw.get("excel")
        if not isinstance(owner_raw, dict) or not isinstance(excel_raw, dict):
            return None
        owner_pid = _as_pid(owner_raw.get("pid"))
        excel_pid = _as_pid(excel_raw.get("pid"))
        if owner_pid is None or excel_pid is None:
            return None
        schema = raw.get("schema_version", SCHEMA_VERSION)
        if isinstance(schema, bool) or not isinstance(schema, int):
            schema = SCHEMA_VERSION
        return cls(
            owner=OwnerIdentity(
                pid=owner_pid,
                create_time_filetime=_as_int(owner_raw.get("create_time_filetime")),
                exe=_as_str(owner_raw.get("exe")),
                session_id=_as_int(owner_raw.get("session_id")),
            ),
            excel=ExcelIdentity(
                pid=excel_pid,
                create_time_filetime=_as_int(excel_raw.get("create_time_filetime")),
                exe=_as_str(excel_raw.get("exe")),
                session_id=_as_int(excel_raw.get("session_id")),
                image_verified=bool(excel_raw.get("image_verified", False)),
            ),
            state=str(raw.get("state", STATE_RUNNING)),
            schema_version=schema,
            updated_at=str(raw.get("updated_at", "")),
        )


def _as_pid(value: Any) -> int | None:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value if value > 0 else None


def _as_int(value: Any) -> int | None:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value


def _as_str(value: Any) -> str | None:
    return value if isinstance(value, str) and value else None


def _is_excel_image(exe: str | None) -> bool:
    """Whether ``exe`` is the ``EXCEL.EXE`` image (basename, case-insensitive)."""

    if not exe:
        return False
    base = exe.replace("/", "\\").rsplit("\\", 1)[-1].strip().lower()
    return base == "excel.exe"


def build_owned_record(
    *,
    owner_pid: int,
    excel_pid: int,
    owner_create_time: int | None,
    owner_exe: str | None,
    owner_session_id: int | None,
    excel_create_time: int | None,
    excel_exe: str | None,
    excel_session_id: int | None,
    state: str = STATE_RUNNING,
) -> OwnedProcessRecord:
    """Assemble an :class:`OwnedProcessRecord` from already-probed primitives.

    Pure (no Win32): the caller gathers the identity values (via the adapter's
    read-only probes) and passes them in, so this stays unit-testable on any host.
    ``image_verified`` is derived from the Excel image name being ``EXCEL.EXE``.
    """

    return OwnedProcessRecord(
        owner=OwnerIdentity(
            pid=owner_pid,
            create_time_filetime=owner_create_time,
            exe=owner_exe,
            session_id=owner_session_id,
        ),
        excel=ExcelIdentity(
            pid=excel_pid,
            create_time_filetime=excel_create_time,
            exe=excel_exe,
            session_id=excel_session_id,
            image_verified=_is_excel_image(excel_exe),
        ),
        state=state,
    )


def default_owned_pid_dir() -> Path:
    """Return the stable per-user directory for the ownership marker.

    Prefers ``%LOCALAPPDATA%\\excel-vba-mcp`` (non-volatile, per-user); falls back
    to the system temp dir only when ``LOCALAPPDATA`` is unset (e.g. a stripped
    service context). This is deliberately **not** the volatile per-work ``%TEMP%``
    working-copy dir so the marker survives a server restart / logoff-logon.
    """

    local = os.environ.get("LOCALAPPDATA", "").strip()
    if local:
        return Path(local) / "excel-vba-mcp"
    import tempfile

    return Path(tempfile.gettempdir()) / "excel-vba-mcp"


class OwnedPidStore:
    """Persist the owned Excel + owner identity for startup reconcile (Block A).

    Args:
        root: Directory that holds the ``owned_pid.json`` state document (a stable
            per-user root in production; a temp dir in tests). Created on first
            write.
    """

    def __init__(self, root: Path) -> None:
        self._root = Path(root)
        self._path = self._root / _STATE_FILE_NAME

    @property
    def path(self) -> Path:
        """The state document path (for diagnostics / tests)."""

        return self._path

    def write_record(self, record: OwnedProcessRecord) -> None:
        """Atomically persist the full ownership ``record``."""

        payload = json.dumps(record.to_dict()).encode("utf-8")
        _atomic_write_json(self._path, payload)

    def write(self, owned_pid: int) -> None:
        """Backward-compat: persist only the owned Excel pid (minimal record).

        Retained for embedders/tests that only track a pid; production uses
        :meth:`write_record` with full owner+excel identity.
        """

        record = OwnedProcessRecord(
            owner=OwnerIdentity(pid=0),
            excel=ExcelIdentity(pid=int(owned_pid)),
            state=STATE_RUNNING,
        )
        self.write_record(record)

    def read_record(self) -> OwnedProcessRecord | None:
        """Return the persisted ownership record, or ``None`` when absent/invalid."""

        if not self._path.is_file():
            return None
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None
        if not isinstance(raw, dict):
            return None
        return OwnedProcessRecord.from_dict(raw)

    def read(self) -> int | None:
        """Return the persisted owned Excel pid, or ``None`` when absent/unreadable.

        Backward-compat convenience over :meth:`read_record`.
        """

        record = self.read_record()
        return record.excel.pid if record is not None else None

    def clear(self) -> None:
        """Remove the state document (best-effort; a dead pid must not block start)."""

        with contextlib.suppress(OSError):
            self._path.unlink(missing_ok=True)