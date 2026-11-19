"""Immutable workbook snapshot store (task_014, spec §9.1 / §9.4).

:class:`SnapshotStore` is the pure-filesystem persistence layer for immutable
workbook snapshots. It owns **no COM** and performs **no Excel access**: every
Excel operation (e.g. ``Workbook.SaveCopyAs``) is orchestrated by the
application-layer :class:`~excel_vba_mcp.application.snapshot_service.SnapshotService`,
which runs the COM callable on the dedicated STA worker and then asks this store
to *finalize* the produced file.

Snapshot layout (spec §9.1)::

    <root>/<snapshot_id>/
        <file_name>          # the immutable .xlsm/.xlam copy
        manifest.json        # metadata + SHA-256 (written atomically)

Guarantees:

* **Immutable.** ``create`` / ``finalize_open_branch`` never mutate an existing
  snapshot directory. A snapshot id maps to exactly one directory; creating the
  same id twice raises :class:`~excel_vba_mcp.domain.errors.SnapshotFailedError`.
* **Atomic (crash-safe).** The snapshot is assembled in a hidden staging
  directory and then renamed into place with :func:`os.replace`. On any failure
  the staging directory is removed, so the target directory is *either absent or
  complete* — there is never a half-written ``manifest.json``. The manifest file
  itself is written via a temp file + :func:`os.replace`.
* **Content-addressed.** ``sha256`` is the SHA-256 of the stored snapshot file;
  ``restore`` / ``diff`` re-hash and compare against it.

The store is intentionally independent of the domain models (it defines its own
:class:`SnapshotManifest` schema for the on-disk manifest) so the storage layer
does not depend on the application/domain layers.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import os
import shutil
import tempfile
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.domain.errors import SnapshotFailedError

# Snapshot creation method persisted in the manifest (spec §9.1).
SnapshotMethod = Literal["SaveCopyAs", "file_copy"]

#: Name of the manifest file inside every snapshot directory.
_MANIFEST_NAME = "manifest.json"

#: In-memory schema version for the on-disk manifest.
_MANIFEST_SCHEMA_VERSION = 1

# Type alias for the optional VBA-export list. Defined at module scope so the
# builtin ``list`` is not shadowed by :meth:`SnapshotStore.list` inside the class.
_VbaExports = list[dict[str, Any]] | None


class SnapshotManifest(BaseModel):
    """On-disk snapshot manifest (spec §9.1).

    This is the storage-layer schema written to ``manifest.json``. It is a plain
    Pydantic model so it serializes deterministically via
    :meth:`model_dump` (``mode="json"``).
    """

    model_config = ConfigDict(extra="forbid")

    snapshot_id: str
    created_at: datetime = Field(description="UTC ISO-8601 creation timestamp.")
    source_path: str = Field(description="Canonical path of the source workbook.")
    file_name: str = Field(description="Snapshot file name inside the snapshot dir.")
    stored_path: str = Field(description="Absolute path of the stored snapshot file.")
    sha256: str = Field(description="SHA-256 hex digest of the stored snapshot file.")
    size: int = Field(description="Size of the stored snapshot file in bytes.")
    excel_version: str | None = None
    workbook_model: dict[str, Any] = Field(default_factory=dict)
    vba_exports: list[dict[str, Any]] | None = Field(
        default=None,
        description="Exported VBA components (path/sha256) or null when unavailable.",
    )
    snapshot_method: SnapshotMethod = "file_copy"
    external_side_effects_not_undone: bool = True
    model_version: int = _MANIFEST_SCHEMA_VERSION


def _utc_now() -> datetime:
    return datetime.now(UTC)


def _validate_snapshot_id(snapshot_id: str) -> str:
    """Reject ids that could escape / confuse the snapshot directory layout."""

    if not isinstance(snapshot_id, str) or not snapshot_id.strip():
        raise SnapshotFailedError(
            "snapshot_id must be a non-empty string.",
            context={"field": "snapshot_id"},
        )
    if "/" in snapshot_id or "\\" in snapshot_id or snapshot_id in {".", ".."}:
        raise SnapshotFailedError(
            "snapshot_id must not contain path separators.",
            context={"snapshot_id": snapshot_id},
        )
    return snapshot_id


def _sha256_file(path: Path) -> tuple[str, int]:
    """Return ``(sha256_hex, size_bytes)`` of ``path`` (streaming)."""

    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


class SnapshotStore:
    """Filesystem store for immutable workbook snapshots.

    Args:
        root: Directory that holds one sub-directory per snapshot. Created on
            first use; never contains repository sources.
    """

    def __init__(self, root: Path) -> None:
        self._root = Path(root)
        self._root.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    @property
    def root(self) -> Path:
        """The snapshot root directory."""

        return self._root

    def create(
        self,
        snapshot_id: str,
        *,
        source_path: str | Path,
        excel_version: str | None = None,
        workbook_model: dict[str, Any] | None = None,
        vba_exports: _VbaExports = None,
        method: SnapshotMethod = "file_copy",
    ) -> SnapshotManifest:
        """Create an immutable snapshot by copying the source file (closed branch).

        The source file (a *closed* workbook on disk) is copied into a fresh
        snapshot directory, hashed and recorded in an atomically-written
        ``manifest.json``. Never mutates an existing snapshot directory.

        Args:
            snapshot_id: Unique snapshot id (becomes the directory name).
            source_path: Path of the file to copy in.
            excel_version: Excel version string for the manifest.
            workbook_model: Compact workbook model metadata (sheet names, etc.).
            vba_exports: Exported VBA components (path/hash) or ``None``.
            method: Snapshot method recorded in the manifest (``file_copy``).

        Returns:
            The written :class:`SnapshotManifest`.

        Raises:
            SnapshotFailedError: If the snapshot id already exists or the copy
                / hashing fails (the partial directory is cleaned up).
        """

        _validate_snapshot_id(snapshot_id)
        src = Path(source_path)
        if not src.is_file():
            raise SnapshotFailedError(
                "snapshot source file does not exist.",
                context={"source_path": str(src)},
            )

        staging = self._make_staging(snapshot_id)
        try:
            file_name = src.name
            dest = staging / file_name
            shutil.copy2(src, dest)
            return self._write_manifest(
                staging=staging,
                snapshot_id=snapshot_id,
                source_path=src,
                file_name=file_name,
                excel_version=excel_version,
                workbook_model=workbook_model or {},
                vba_exports=vba_exports,
                method=method,
            )
        except Exception:
            self._cleanup_dir(staging)
            raise

    def write_copy(self, snapshot_id: str, src_path: str | Path) -> Path:
        """Copy ``src_path`` into a new snapshot directory (file-copy primitive).

        Low-level helper that only places the file; it does **not** write a
        manifest. Prefer :meth:`create` for the full closed-branch operation.

        Returns:
            The path of the copied file inside the snapshot directory.

        Raises:
            SnapshotFailedError: If the snapshot id already exists or the copy
                fails (the partial directory is cleaned up).
        """

        _validate_snapshot_id(snapshot_id)
        src = Path(src_path)
        if not src.is_file():
            raise SnapshotFailedError(
                "snapshot source file does not exist.",
                context={"source_path": str(src)},
            )
        staging = self._make_staging(snapshot_id)
        try:
            dest = staging / src.name
            shutil.copy2(src, dest)
            self._publish(staging, snapshot_id)
            return self.path_for(snapshot_id) / src.name
        except Exception:
            self._cleanup_dir(staging)
            raise

    def finalize_open_branch(
        self,
        snapshot_id: str,
        save_copy_as_fn: Callable[[Path], Any],
        *,
        source_path: str | Path,
        excel_version: str | None = None,
        workbook_model: dict[str, Any] | None = None,
        vba_exports: _VbaExports = None,
    ) -> SnapshotManifest:
        """Finalize the SaveCopyAs (open workbook) branch.

        The caller (``SnapshotService``) runs ``Workbook.SaveCopyAs`` on the STA
        worker *first*, producing the snapshot file in a temporary location; this
        method then places that produced file into a fresh immutable snapshot
        directory, hashes it and writes the manifest atomically.

        ``save_copy_as_fn`` is invoked with the destination file path inside the
        staging directory and must place the already-produced snapshot file there
        (a filesystem move/copy). It never runs COM — COM stays on the worker.

        Args:
            snapshot_id: Unique snapshot id.
            save_copy_as_fn: Callable ``(dest_path) -> Any`` that places the
                already-produced snapshot file at ``dest_path``.
            source_path: Canonical source path of the open workbook.
            excel_version: Excel version string.
            workbook_model: Compact workbook model metadata.
            vba_exports: Exported VBA components (path/hash) or ``None``.

        Returns:
            The written :class:`SnapshotManifest` (``snapshot_method="SaveCopyAs"``).

        Raises:
            SnapshotFailedError: If the snapshot id already exists or the
                placement / hashing fails (the partial directory is cleaned up).
        """

        _validate_snapshot_id(snapshot_id)
        src = Path(source_path)
        staging = self._make_staging(snapshot_id)
        try:
            file_name = src.name or "snapshot.xlsm"
            dest = staging / file_name
            save_copy_as_fn(dest)
            if not dest.is_file():
                raise SnapshotFailedError(
                    "save_copy_as_fn did not produce the snapshot file.",
                    context={"snapshot_id": snapshot_id, "dest": str(dest)},
                )
            return self._write_manifest(
                staging=staging,
                snapshot_id=snapshot_id,
                source_path=src,
                file_name=file_name,
                excel_version=excel_version,
                workbook_model=workbook_model or {},
                vba_exports=vba_exports,
                method="SaveCopyAs",
            )
        except Exception:
            self._cleanup_dir(staging)
            raise

    def list(self) -> list[SnapshotManifest]:
        """Return every snapshot manifest under the root (sorted by snapshot id).

        A directory without a readable ``manifest.json`` is ignored (it is not a
        complete snapshot).
        """

        manifests: list[SnapshotManifest] = []
        for entry in sorted(self._root.iterdir(), key=lambda p: p.name):
            if not entry.is_dir() or entry.name.startswith("."):
                continue
            manifest_path = entry / _MANIFEST_NAME
            if not manifest_path.is_file():
                continue
            try:
                manifests.append(self._read_manifest(entry))
            except (OSError, ValueError):
                continue
        return manifests

    def load(self, snapshot_id: str) -> SnapshotManifest:
        """Load and return the manifest for ``snapshot_id``.

        Raises:
            SnapshotFailedError: If the snapshot does not exist or its manifest
                cannot be read.
        """

        _validate_snapshot_id(snapshot_id)
        directory = self.path_for(snapshot_id)
        manifest_path = directory / _MANIFEST_NAME
        if not manifest_path.is_file():
            raise SnapshotFailedError(
                "snapshot does not exist.",
                context={"snapshot_id": snapshot_id, "path": str(directory)},
            )
        return self._read_manifest(directory)

    def path_for(self, snapshot_id: str) -> Path:
        """Return the directory path for ``snapshot_id`` (may not exist yet)."""

        _validate_snapshot_id(snapshot_id)
        return self._root / snapshot_id

    def file_path_for(self, snapshot_id: str) -> Path:
        """Return the stored snapshot file path for ``snapshot_id``."""

        manifest = self.load(snapshot_id)
        return self.path_for(snapshot_id) / manifest.file_name

    def sha256(self, path: str | Path) -> str:
        """Return the SHA-256 hex digest of ``path``."""

        digest, _ = _sha256_file(Path(path))
        return digest

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _make_staging(self, snapshot_id: str) -> Path:
        """Create a unique hidden staging directory and guard immutability."""

        final_dir = self.path_for(snapshot_id)
        if final_dir.exists():
            raise SnapshotFailedError(
                "snapshot already exists (immutable).",
                context={"snapshot_id": snapshot_id, "path": str(final_dir)},
            )
        staging = Path(tempfile.mkdtemp(prefix=f".{snapshot_id}.", dir=str(self._root)))
        return staging

    def _write_manifest(
        self,
        *,
        staging: Path,
        snapshot_id: str,
        source_path: Path,
        file_name: str,
        excel_version: str | None,
        workbook_model: dict[str, Any],
        vba_exports: _VbaExports,
        method: SnapshotMethod,
    ) -> SnapshotManifest:
        """Hash the staged file, write the manifest atomically and publish."""

        snapshot_file = staging / file_name
        if not snapshot_file.is_file():
            raise SnapshotFailedError(
                "snapshot file missing in staging directory.",
                context={"snapshot_id": snapshot_id, "file": str(snapshot_file)},
            )
        sha256, size = _sha256_file(snapshot_file)
        # ``stored_path`` is resolved after publish; record the final absolute path.
        final_dir = self._root / snapshot_id
        manifest = SnapshotManifest(
            snapshot_id=snapshot_id,
            created_at=_utc_now(),
            source_path=str(_canonical_source(source_path)),
            file_name=file_name,
            stored_path=str(final_dir / file_name),
            sha256=sha256,
            size=size,
            excel_version=excel_version,
            workbook_model=workbook_model,
            vba_exports=vba_exports,
            snapshot_method=method,
            external_side_effects_not_undone=True,
        )
        self._atomic_write_json(staging / _MANIFEST_NAME, manifest.model_dump(mode="json"))
        self._publish(staging, snapshot_id)
        return manifest

    def _publish(self, staging: Path, snapshot_id: str) -> None:
        """Atomically rename the staging directory to its final location."""

        final_dir = self._root / snapshot_id
        if final_dir.exists():  # defensive: never overwrite an immutable dir
            self._cleanup_dir(staging)
            raise SnapshotFailedError(
                "snapshot already exists (immutable).",
                context={"snapshot_id": snapshot_id},
            )
        os.replace(staging, final_dir)

    @staticmethod
    def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
        """Write ``payload`` as JSON to ``path`` atomically (temp + replace)."""

        text = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
        fd, tmp_name = tempfile.mkstemp(prefix=path.name + ".", dir=str(path.parent))
        tmp_path = Path(tmp_name)
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(text)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(tmp_path, path)
        except Exception:
            with contextlib.suppress(OSError):
                tmp_path.unlink(missing_ok=True)
            raise

    @staticmethod
    def _read_manifest(directory: Path) -> SnapshotManifest:
        manifest_path = directory / _MANIFEST_NAME
        raw = json.loads(manifest_path.read_text(encoding="utf-8"))
        return SnapshotManifest.model_validate(raw)

    @staticmethod
    def _cleanup_dir(directory: Path) -> None:
        with contextlib.suppress(OSError):
            shutil.rmtree(directory, ignore_errors=True)


def _canonical_source(source_path: Path) -> str:
    """Return an absolute path string for ``source_path`` (best effort)."""

    try:
        return str(source_path.resolve())
    except OSError:  # pragma: no cover - resolve rarely fails
        return str(source_path)


__all__ = ["SnapshotManifest", "SnapshotMethod", "SnapshotStore"]
