"""Unit tests for the immutable snapshot store (task_014).

These tests exercise :class:`SnapshotStore` purely on the filesystem — no Excel,
no COM. They verify manifest fields, SHA-256 correctness, atomic write
(no partial manifest on simulated failure), immutability (an existing snapshot
directory is never mutated) and list/load.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

import excel_vba_mcp.storage.snapshot_store as store_mod
from excel_vba_mcp.domain.errors import SnapshotFailedError
from excel_vba_mcp.storage.snapshot_store import SnapshotStore


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_source(tmp_path: Path, name: str = "Book.xlsm", data: bytes = b"workbook-bytes") -> Path:
    src = tmp_path / name
    src.write_bytes(data)
    return src


# ---------------------------------------------------------------------------
# Manifest fields + SHA-256 correctness (file-copy branch)
# ---------------------------------------------------------------------------
def test_create_file_copy_manifest_fields_and_hash(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    src = _make_source(tmp_path, data=b"immutable-content")

    manifest = store.create(
        "snap-1",
        source_path=src,
        excel_version="16.0",
        workbook_model={"name": "Book.xlsm", "sheet_count": 1},
        method="file_copy",
    )

    assert manifest.snapshot_id == "snap-1"
    assert manifest.file_name == "Book.xlsm"
    assert manifest.snapshot_method == "file_copy"
    assert manifest.external_side_effects_not_undone is True
    assert manifest.excel_version == "16.0"
    assert manifest.workbook_model == {"name": "Book.xlsm", "sheet_count": 1}
    assert manifest.size == len(b"immutable-content")
    assert manifest.sha256 == _sha256(src)
    # created_at is timezone-aware UTC.
    assert manifest.created_at.tzinfo is not None

    # The snapshot directory contains the copied file + a manifest.json.
    snapshot_dir = store.path_for("snap-1")
    assert (snapshot_dir / "Book.xlsm").is_file()
    manifest_on_disk = snapshot_dir / "manifest.json"
    assert manifest_on_disk.is_file()
    # The stored file hashes to the manifest digest.
    assert store.sha256(snapshot_dir / "Book.xlsm") == manifest.sha256
    # manifest.json round-trips to the same manifest.
    raw = json.loads(manifest_on_disk.read_text(encoding="utf-8"))
    assert raw["sha256"] == manifest.sha256
    assert raw["snapshot_method"] == "file_copy"


# ---------------------------------------------------------------------------
# SHA-256 matches a known digest
# ---------------------------------------------------------------------------
def test_sha256_matches_known_digest(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    src = _make_source(tmp_path, data=b"abc")
    manifest = store.create("snap-abc", source_path=src)
    assert manifest.sha256 == hashlib.sha256(b"abc").hexdigest()


# ---------------------------------------------------------------------------
# Atomic write: no partial manifest on simulated failure
# ---------------------------------------------------------------------------
def test_create_no_partial_manifest_on_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    src = _make_source(tmp_path)
    snapshot_root = store.root

    # Simulate a failure during hashing (after the file is staged but before the
    # manifest is written); the staging directory must be cleaned up entirely.
    def boom(_path: Path) -> tuple[str, int]:
        raise OSError("simulated hashing failure")

    monkeypatch.setattr(store_mod, "_sha256_file", boom)

    with pytest.raises(OSError, match="simulated hashing failure"):
        store.create("snap-fail", source_path=src)

    # No snapshot directory and no staging leftovers.
    assert not (snapshot_root / "snap-fail").exists()
    leftovers = [p.name for p in snapshot_root.iterdir() if p.name.startswith(".")]
    assert leftovers == [], f"staging directories left behind: {leftovers}"


def test_finalize_open_branch_cleans_up_on_failure(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    snapshot_root = store.root

    def failing_place(dest: Path) -> None:
        # Produce nothing -> finalize must raise and clean up.
        return None

    with pytest.raises(SnapshotFailedError):
        store.finalize_open_branch(
            "snap-open-fail",
            failing_place,
            source_path=tmp_path / "Book.xlsm",
        )

    assert not (snapshot_root / "snap-open-fail").exists()
    leftovers = [p.name for p in snapshot_root.iterdir() if p.name.startswith(".")]
    assert leftovers == []


def test_finalize_open_branch_success(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    produced = tmp_path / "produced.xlsm"
    produced.write_bytes(b"saved-by-savecopyas")

    def place(dest: Path) -> None:
        dest.write_bytes(produced.read_bytes())

    manifest = store.finalize_open_branch(
        "snap-open",
        place,
        source_path=tmp_path / "Book.xlsm",
        excel_version="16.0",
        workbook_model={"name": "Book.xlsm"},
    )

    assert manifest.snapshot_method == "SaveCopyAs"
    assert manifest.sha256 == _sha256(produced)
    assert manifest.external_side_effects_not_undone is True
    stored = store.path_for("snap-open") / "Book.xlsm"
    assert stored.is_file()
    assert store.sha256(stored) == manifest.sha256


# ---------------------------------------------------------------------------
# Immutability: existing snapshot dir is never mutated
# ---------------------------------------------------------------------------
def test_create_is_immutable_same_id_rejected(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    src = _make_source(tmp_path, data=b"v1")
    first = store.create("snap-fixed", source_path=src)
    manifest_before = (store.path_for("snap-fixed") / "manifest.json").read_text(encoding="utf-8")
    stored_before = (store.path_for("snap-fixed") / "Book.xlsm").read_bytes()

    # A second create with the same id must raise and NOT touch the existing dir.
    with pytest.raises(SnapshotFailedError):
        store.create("snap-fixed", source_path=_make_source(tmp_path, data=b"v2"))

    manifest_after = (store.path_for("snap-fixed") / "manifest.json").read_text(encoding="utf-8")
    stored_after = (store.path_for("snap-fixed") / "Book.xlsm").read_bytes()
    assert manifest_after == manifest_before
    assert stored_after == stored_before
    assert store.load("snap-fixed").sha256 == first.sha256


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def test_create_rejects_path_separator_id(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    with pytest.raises(SnapshotFailedError):
        store.create("../escape", source_path=_make_source(tmp_path))


def test_create_rejects_missing_source(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    with pytest.raises(SnapshotFailedError):
        store.create("snap-x", source_path=tmp_path / "does-not-exist.xlsm")


# ---------------------------------------------------------------------------
# list / load
# ---------------------------------------------------------------------------
def test_list_and_load(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    a = store.create("snap-a", source_path=_make_source(tmp_path, name="A.xlsm", data=b"a"))
    b = store.create("snap-b", source_path=_make_source(tmp_path, name="B.xlsm", data=b"b"))

    listed = store.list()
    assert [m.snapshot_id for m in listed] == ["snap-a", "snap-b"]

    loaded = store.load("snap-a")
    assert loaded.sha256 == a.sha256
    assert loaded.source_path == a.source_path
    assert store.load("snap-b").sha256 == b.sha256


def test_list_ignores_dirs_without_manifest(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    store.create("snap-ok", source_path=_make_source(tmp_path))
    # A stray directory with no manifest is ignored.
    (store.root / "not-a-snapshot").mkdir()

    ids = [m.snapshot_id for m in store.list()]
    assert ids == ["snap-ok"]


def test_load_missing_raises(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    with pytest.raises(SnapshotFailedError):
        store.load("nope")


def test_write_copy_primitive(tmp_path: Path) -> None:
    store = SnapshotStore(tmp_path / "snapshots")
    src = _make_source(tmp_path, name="Book.xlsm", data=b"copy-me")
    copied = store.write_copy("snap-copy", src)
    assert copied.is_file()
    assert _sha256(copied) == _sha256(src)
