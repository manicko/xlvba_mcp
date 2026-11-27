"""Unit tests for the streaming artifact writer (WP-F; spec §9, §10, §11).

The store is pure filesystem (no Excel / no COM). These tests prove the streaming
binding contract added for large-data file exports:

* ``open_stream`` writes incrementally and ``finalize`` produces a file whose
  ``sha256`` / ``size`` match the bytes written and the on-disk ``index.json``;
* an exception mid-stream (or an explicit ``abort``) leaves NO final file and
  cleans the temp file deterministically;
* the ``.csv`` / ``.jsonl`` artifact types resolve their canonical extensions;
* the path policy is still enforced (outside-root root, traversal / absolute name).
"""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterator
from pathlib import Path

import pytest

from excel_vba_mcp.domain.errors import PathNotAllowedError
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.artifact_store import ArtifactStore, ArtifactType

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class Rig:
    """A store rooted at ``<tmp>/artifacts`` with a matching allowlist."""

    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.artifacts = tmp_path / "artifacts"
        self.policy = PathPolicy({"artifacts": str(self.artifacts)})
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.store = ArtifactStore(self.artifacts, self.policy, self.audit)


@pytest.fixture
def rig(tmp_path: Path) -> Iterator[Rig]:
    instance = Rig(tmp_path)
    yield instance
    instance.audit.close()


def _index_entries(rig: Rig) -> list[dict[str, object]]:
    index_path = rig.artifacts / "index.json"
    payload = json.loads(index_path.read_text(encoding="utf-8"))
    return [entry for entry in payload if isinstance(entry, dict)]


# ---------------------------------------------------------------------------
# Incremental write + finalize
# ---------------------------------------------------------------------------
def test_stream_writes_incrementally_and_finalizes(rig: Rig) -> None:
    stream = rig.store.open_stream(ArtifactType.CSV, suggested_name="export")

    stream.write_text("a")
    stream.write_text("b")
    stream.write_text("c")
    info = stream.finalize(kind="data_export")

    expected = b"abc"
    assert info.name == "export.csv"
    assert info.size == len(expected)
    assert info.sha256 == hashlib.sha256(expected).hexdigest()
    stored = Path(info.stored_path)
    assert stored.read_bytes() == expected

    # The index records the exact size + sha of the written bytes.
    entries = _index_entries(rig)
    assert len(entries) == 1
    assert entries[0]["sha256"] == info.sha256
    assert entries[0]["size"] == info.size
    assert entries[0]["name"] == "export.csv"

    listed = rig.store.list_artifacts()
    assert [artifact.artifact_id for artifact in listed] == [info.artifact_id]


def test_stream_utf8_byte_count_and_digest(rig: Rig) -> None:
    stream = rig.store.open_stream(ArtifactType.JSONL, suggested_name="data")

    text = "Привет\n"  # non-ASCII to prove UTF-8 byte counting
    stream.write_text(text)
    info = stream.finalize(kind="data_export")

    expected = text.encode("utf-8")
    assert info.size == len(expected)
    assert info.sha256 == hashlib.sha256(expected).hexdigest()
    assert Path(info.stored_path).read_bytes() == expected


# ---------------------------------------------------------------------------
# Cleanup on failure
# ---------------------------------------------------------------------------
def test_exception_mid_stream_leaves_no_final_file(rig: Rig) -> None:
    stream = rig.store.open_stream(ArtifactType.CSV, suggested_name="export")
    stream.write_text("partial")

    with pytest.raises(RuntimeError), stream:
        stream.write_text("more")
        raise RuntimeError("boom")

    # No final artifact and no leftover temp file (only the dir may exist).
    assert list(rig.artifacts.iterdir()) == []
    assert rig.store.list_artifacts() == []


def test_abort_cleans_temp_without_finalizing(rig: Rig) -> None:
    stream = rig.store.open_stream(ArtifactType.CSV, suggested_name="export")
    stream.write_text("discard me")

    stream.abort()

    assert list(rig.artifacts.iterdir()) == []
    assert rig.store.list_artifacts() == []


def test_finalize_twice_is_rejected(rig: Rig) -> None:
    from excel_vba_mcp.domain.errors import ArtifactWriteError

    stream = rig.store.open_stream(ArtifactType.CSV, suggested_name="export")
    stream.write_text("x")
    stream.finalize(kind="data_export")

    with pytest.raises(ArtifactWriteError):
        stream.finalize(kind="data_export")


# ---------------------------------------------------------------------------
# Extension resolution
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    ("artifact_type", "expected_name"),
    [
        (ArtifactType.CSV, "report.csv"),
        (ArtifactType.JSONL, "report.jsonl"),
    ],
)
def test_stream_extension_resolves(
    rig: Rig, artifact_type: ArtifactType, expected_name: str
) -> None:
    stream = rig.store.open_stream(artifact_type, suggested_name="report")
    stream.write_text("x")
    info = stream.finalize(kind="data_export")

    assert info.name == expected_name
    assert Path(info.stored_path).name == expected_name


def test_stream_conflicting_extension_rejected(rig: Rig) -> None:
    with pytest.raises(PathNotAllowedError):
        rig.store.open_stream(ArtifactType.CSV, suggested_name="report.zip")


# ---------------------------------------------------------------------------
# Path policy still enforced
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("name", ["../evil", "..\\evil", "sub/dir", "/abs/path", ""])
def test_stream_traversal_names_rejected(rig: Rig, name: str) -> None:
    with pytest.raises(PathNotAllowedError):
        rig.store.open_stream(ArtifactType.CSV, suggested_name=name)


def test_stream_outside_root_is_rejected(tmp_path: Path) -> None:
    outside = tmp_path / "outside"
    policy = PathPolicy({"projects": str(tmp_path / "projects")})
    audit = AuditLogger(tmp_path / "audit.jsonl")
    try:
        store = ArtifactStore(outside, policy, audit)

        with pytest.raises(PathNotAllowedError):
            store.open_stream(ArtifactType.CSV, suggested_name="export")
        assert not outside.exists()
    finally:
        audit.close()