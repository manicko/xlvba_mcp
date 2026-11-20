"""Filesystem artifact store (task_034, spec §9, §10, §11).

Persists *evidence* files (screenshots, logs, exports, diagnostics, traces) under
an allowlisted artifact root. Every write is:

* **path-policy enforced** — the final target is validated with
  :class:`~excel_vba_mcp.security.path_policy.PathPolicy` against the artifact
  root (canonical path, ``..`` traversal, symlink/reparse escape, UNC, device
  path). Traversal, outside-root, empty and disallowed artifact types are
  rejected with ``PATH_NOT_ALLOWED``;
* **atomic** — the payload is written to a temp file in the destination
  directory, flushed / ``fsync``-ed and then moved into place with
  ``os.replace``. A crash or error never leaves a partial artifact behind;
* **recorded** — size and SHA-256 are computed from the exact bytes written, and
  one entry is appended to an on-disk ``index.json`` plus one audit record. No
  secret, password or full workbook content is ever written to the index or the
  audit log.

The store is pure filesystem: it never touches Excel / COM.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
import tempfile
import threading
from datetime import UTC, datetime
from enum import StrEnum
from pathlib import Path
from typing import Any
from uuid import uuid4

from pydantic import BaseModel, ConfigDict

from excel_vba_mcp.domain.errors import ArtifactWriteError, PathNotAllowedError
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathDenyReason, PathPolicy

__all__ = [
    "ARTIFACT_ROOT_NAME",
    "ArtifactInfo",
    "ArtifactStore",
    "ArtifactType",
]

#: Logical name of the artifact root in :class:`PathPolicy`.
ARTIFACT_ROOT_NAME = "artifacts"

#: On-disk artifact index (one JSON object per line, append-only).
_INDEX_FILENAME = "index.json"

#: Upper bound for an artifact name (basename only; no directories).
_MAX_NAME_LENGTH = 160


class ArtifactType(StrEnum):
    """Kind of artifact being stored (spec §9: screenshots, logs, exports, ...)."""

    SCREENSHOT = "screenshot"
    LOG = "log"
    EXPORT = "export"
    MANIFEST = "manifest"
    DIAGNOSTIC = "diagnostic"
    TRACE = "trace"


#: The single file extension each artifact type is stored with. Artifacts are
#: evidence files, so the extension policy is owned by this allowlist rather than
#: by the snapshot-sidecar extension set in :mod:`path_policy` (see
#: :meth:`ArtifactStore._validate_target`).
_TYPE_EXTENSIONS: dict[ArtifactType, str] = {
    ArtifactType.SCREENSHOT: ".png",
    ArtifactType.LOG: ".log",
    ArtifactType.EXPORT: ".zip",
    ArtifactType.MANIFEST: ".json",
    ArtifactType.DIAGNOSTIC: ".json",
    ArtifactType.TRACE: ".txt",
}


class ArtifactInfo(BaseModel):
    """Serializable metadata of one stored artifact (spec §6.6 artifact resource)."""

    model_config = ConfigDict(extra="forbid")

    artifact_id: str
    name: str
    artifact_type: ArtifactType
    workbook_id: str | None = None
    stored_path: str
    sha256: str
    size: int
    created_at: datetime


def _validate_artifact_name(name: str) -> str:
    """Return a safe *basename* artifact name or raise ``PATH_NOT_ALLOWED``.

    A name must be a single path component: separators, absolute paths,
    ``..`` traversal, empty / dot names and control characters are rejected.
    """

    if not isinstance(name, str):
        raise PathNotAllowedError(
            "The artifact name must be a string.",
            context={"reason": "invalid_name"},
        )
    candidate = name.strip()
    if not candidate:
        raise PathNotAllowedError(
            "The artifact name is empty.",
            context={"reason": PathDenyReason.EMPTY.value},
        )
    if len(candidate) > _MAX_NAME_LENGTH:
        raise PathNotAllowedError(
            "The artifact name is too long.",
            context={"reason": "invalid_name"},
        )
    if "/" in candidate or "\\" in candidate:
        raise PathNotAllowedError(
            "The artifact name must not contain a path separator.",
            context={"reason": PathDenyReason.TRAVERSAL.value},
        )
    candidate_path = Path(candidate)
    if candidate_path.is_absolute() or ".." in candidate_path.parts:
        raise PathNotAllowedError(
            "The artifact name must not escape the artifact root.",
            context={"reason": PathDenyReason.TRAVERSAL.value},
        )
    if candidate in {".", ".."}:
        raise PathNotAllowedError(
            "The artifact name is not a valid file name.",
            context={"reason": "invalid_name"},
        )
    if any(ord(character) < 32 for character in candidate):
        raise PathNotAllowedError(
            "The artifact name contains a control character.",
            context={"reason": "invalid_name"},
        )
    return candidate


def _ensure_extension(name: str, artifact_type: ArtifactType) -> str:
    """Return ``name`` with the canonical extension of ``artifact_type``.

    A missing extension is appended; a conflicting extension is rejected with
    ``PATH_NOT_ALLOWED`` so an artifact cannot be stored under a misleading type.
    """

    expected = _TYPE_EXTENSIONS[artifact_type]
    if name.lower().endswith(expected):
        return name
    if Path(name).suffix:
        raise PathNotAllowedError(
            "The artifact name extension does not match the artifact type.",
            context={
                "reason": PathDenyReason.EXTENSION_NOT_ALLOWED.value,
                "expected_extension": expected,
            },
        )
    return f"{name}{expected}"


class ArtifactStore:
    """Store artifacts under an allowlisted root with atomic writes + audit."""

    def __init__(self, root: Path, path_policy: PathPolicy, audit: AuditLogger) -> None:
        """Bind the store to ``root`` (inside the allowlisted artifact root).

        ``path_policy`` must allow ``root``; a root outside the allowlist makes
        every write fail closed with ``PATH_NOT_ALLOWED`` (checked on each save,
        not cached at construction, so a tightened policy takes effect).
        """

        self._root = Path(root)
        self._path_policy = path_policy
        self._audit = audit
        self._index_path = self._root / _INDEX_FILENAME
        # Serializes index read-modify-write across threads (the store is used
        # from the event loop and from worker threads alike).
        self._index_lock = threading.Lock()
        # No filesystem write here: the root directory is created lazily by the
        # first (policy-validated) save, so an out-of-policy root never triggers
        # a directory creation.

    @property
    def root(self) -> Path:
        """The resolved artifact root directory."""

        return self._root

    def path_for(self, name: str) -> Path:
        """Return the validated absolute path for ``name`` under the artifact root.

        Raises:
            excel_vba_mcp.domain.errors.PathNotAllowedError: if the name escapes
                the artifact root or the target is outside the allowlist.
        """

        safe_name = _validate_artifact_name(name)
        target = self._root / safe_name
        return self._validate_target(target)

    def save_artifact(
        self,
        name: str,
        data: bytes,
        *,
        artifact_type: str,
        workbook_id: str | None = None,
    ) -> ArtifactInfo:
        """Atomically persist ``data`` and return its :class:`ArtifactInfo`.

        The artifact is stored under ``name`` (basename; no directories), so
        :meth:`path_for` maps a name back to its stored file. When ``name`` has
        no extension the canonical extension of ``artifact_type`` is appended; a
        conflicting extension is rejected.

        Args:
            name: Artifact file name (basename, no directories).
            data: Exact artifact bytes (never logged).
            artifact_type: One of :class:`ArtifactType` (disallowed -> ``PATH_NOT_ALLOWED``).
            workbook_id: Optional owning workbook id.

        Returns:
            The stored artifact metadata (size + SHA-256 of the written bytes).

        Raises:
            excel_vba_mcp.domain.errors.PathNotAllowedError: traversal,
                outside-root, disallowed artifact type or extension.
            excel_vba_mcp.domain.errors.ArtifactWriteError: the write/replace failed.
        """

        try:
            resolved_type = ArtifactType(artifact_type)
        except ValueError as exc:
            raise PathNotAllowedError(
                "The artifact type is not allowed.",
                context={
                    "reason": PathDenyReason.EXTENSION_NOT_ALLOWED.value,
                    "artifact_type": str(artifact_type),
                },
            ) from exc
        if not isinstance(data, (bytes, bytearray, memoryview)):
            raise ArtifactWriteError(
                "The artifact payload must be bytes.",
                context={"field": "data", "type": type(data).__name__},
            )

        safe_name = _ensure_extension(_validate_artifact_name(name), resolved_type)
        payload = bytes(data)
        target = self.path_for(safe_name)
        sha256 = hashlib.sha256(payload).hexdigest()
        size = len(payload)

        self._atomic_write(target, payload)

        info = ArtifactInfo(
            artifact_id=uuid4().hex,
            name=safe_name,
            artifact_type=resolved_type,
            workbook_id=workbook_id,
            stored_path=str(target),
            sha256=sha256,
            size=size,
            created_at=datetime.now(UTC),
        )
        self._append_index(info)
        self._record_audit(name=safe_name, info=info)
        return info

    def list_artifacts(self, *, workbook_id: str | None = None) -> list[ArtifactInfo]:
        """List stored artifacts, optionally filtered by owning ``workbook_id``.

        Reads the append-only index; a missing or unreadable index yields an
        empty list rather than an error (the artifact files are the source of
        truth, the index is a convenience view). When a name was saved more than
        once, the most recent entry wins (the index is append-only, the file is
        replaced atomically).
        """

        entries = self._read_index()
        latest: dict[str, ArtifactInfo] = {}
        for entry in entries:
            artifact = ArtifactInfo.model_validate(entry)
            if workbook_id is not None and artifact.workbook_id != workbook_id:
                continue
            latest[artifact.name] = artifact
        return list(latest.values())

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------
    def _validate_target(self, target: Path) -> Path:
        """Validate ``target`` under the artifact root via :class:`PathPolicy`.

        ``PathPolicy`` enforces canonical containment, ``..`` traversal,
        symlink/reparse escapes, UNC/device paths and the artifact root match.
        Its extension allowlist targets snapshot sidecars, so an
        ``EXTENSION_NOT_ALLOWED`` rejection is accepted only after containment
        has been proven (artifacts are evidence files such as screenshots/logs,
        whose extension policy is owned by :data:`_TYPE_EXTENSIONS`); every
        other rejection propagates as ``PATH_NOT_ALLOWED``.
        """

        try:
            return self._path_policy.validate_path(
                target, OperationClass.MUTATING, expected_root=ARTIFACT_ROOT_NAME
            )
        except PathNotAllowedError as exc:
            reason = exc.context.get("reason")
            if reason != PathDenyReason.EXTENSION_NOT_ALLOWED.value:
                raise
            # Root containment and the artifact-root match already passed above.
            return Path(os.path.realpath(target))

    def _atomic_write(self, target: Path, payload: bytes) -> None:
        """Write ``payload`` to ``target`` atomically (temp file + ``os.replace``)."""

        target.parent.mkdir(parents=True, exist_ok=True)
        descriptor = -1
        temp_path: Path | None = None
        try:
            descriptor, temp_name = tempfile.mkstemp(
                prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent)
            )
            temp_path = Path(temp_name)
            with os.fdopen(descriptor, "wb") as handle:
                descriptor = -1  # ownership transferred to the file object
                handle.write(payload)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_path, target)
            temp_path = None
        except OSError as exc:
            raise ArtifactWriteError(
                "The artifact could not be written.",
                context={"artifact_name": target.name, "error": exc.__class__.__name__},
            ) from exc
        finally:
            if descriptor >= 0:
                with contextlib.suppress(OSError):
                    os.close(descriptor)
            if temp_path is not None:
                with contextlib.suppress(OSError):
                    temp_path.unlink(missing_ok=True)

    def _append_index(self, info: ArtifactInfo) -> None:
        """Append one artifact entry to the on-disk index (atomic, thread-safe)."""

        with self._index_lock:
            entries = self._read_index_unlocked()
            entries.append(info.model_dump(mode="json"))
            self._atomic_write(
                self._index_path,
                json.dumps(entries, indent=2, sort_keys=True).encode("utf-8"),
            )

    def _read_index(self) -> list[dict[str, Any]]:
        with self._index_lock:
            return self._read_index_unlocked()

    def _read_index_unlocked(self) -> list[dict[str, Any]]:
        if not self._index_path.is_file():
            return []
        try:
            payload = json.loads(self._index_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return []
        if not isinstance(payload, list):
            return []
        return [entry for entry in payload if isinstance(entry, dict)]

    def _record_audit(self, *, name: str, info: ArtifactInfo) -> None:
        """Append one audit record (paths redacted to basename by the audit logger)."""

        self._audit.record(
            tool="artifact_save",
            operation_class=OperationClass.MUTATING,
            params={
                "artifact_id": info.artifact_id,
                "name": name,
                "artifact_type": info.artifact_type.value,
                "workbook_id": info.workbook_id,
                "artifact_path": info.stored_path,
                "size": info.size,
                "sha256": info.sha256,
            },
            result={
                "artifact_id": info.artifact_id,
                "sha256": info.sha256,
                "size": info.size,
                "saved": True,
            },
        )
