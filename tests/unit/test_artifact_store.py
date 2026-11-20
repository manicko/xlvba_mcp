"""Unit tests for the filesystem artifact store (task_034; spec §9, §10, §11).

The store is pure filesystem (no Excel / no COM). These tests prove the binding
contract:

* ``save_artifact`` / ``list_artifacts`` round-trip with the exact size and
  SHA-256 of the written bytes;
* the write is atomic (temp file + ``os.replace``) and a failed replace leaves
  neither a partial artifact nor a leftover temp file;
* traversal / absolute names, outside-root roots and disallowed artifact types /
  extensions are rejected with ``PATH_NOT_ALLOWED``;
* the audit entry redacts the artifact path to its basename and never contains
  the artifact bytes or any secret.
"""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterator
from pathlib import Path

import pytest

from excel_vba_mcp.domain.errors import ArtifactWriteError, PathNotAllowedError
from excel_vba_mcp.security.audit import AuditLogger, redact
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage import artifact_store as artifact_store_module
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

    def audit_records(self) -> list[dict[str, object]]:
        lines = self.audit.path.read_text(encoding="utf-8").splitlines()
        return [json.loads(line) for line in lines if line.strip()]


@pytest.fixture
def rig(tmp_path: Path) -> Iterator[Rig]:
    instance = Rig(tmp_path)
    yield instance
    instance.audit.close()


# ---------------------------------------------------------------------------
# save / list round-trip
# ---------------------------------------------------------------------------
def test_save_and_list_records_size_and_sha256(rig: Rig) -> None:
    payload = b"artifact-bytes\x00\x01\x02"

    info = rig.store.save_artifact(
        "shot", payload, artifact_type=ArtifactType.SCREENSHOT, workbook_id="wb-1"
    )

    assert info.name == "shot.png"
    assert info.size == len(payload)
    assert info.sha256 == hashlib.sha256(payload).hexdigest()
    assert Path(info.stored_path).is_file()
    assert Path(info.stored_path).read_bytes() == payload

    listed = rig.store.list_artifacts()
    assert [artifact.artifact_id for artifact in listed] == [info.artifact_id]
    assert listed[0].workbook_id == "wb-1"


def test_sha256_matches_known_digest(rig: Rig) -> None:
    payload = b"abc"

    info = rig.store.save_artifact("known", payload, artifact_type="log")

    assert info.sha256 == "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    assert info.size == 3


def test_list_filters_by_workbook(rig: Rig) -> None:
    rig.store.save_artifact("a", b"1", artifact_type="log", workbook_id="wb-1")
    rig.store.save_artifact("b", b"2", artifact_type="log", workbook_id="wb-2")
    rig.store.save_artifact("c", b"3", artifact_type="log")

    only_one = rig.store.list_artifacts(workbook_id="wb-1")

    assert [artifact.name for artifact in only_one] == ["a.log"]
    assert len(rig.store.list_artifacts()) == 3


def test_missing_index_lists_empty(tmp_path: Path) -> None:
    artifacts = tmp_path / "artifacts"
    store = ArtifactStore(artifacts, PathPolicy({"artifacts": str(artifacts)}), _NullAudit())

    assert store.list_artifacts() == []


class _NullAudit:
    """Minimal audit stand-in for the empty-index test (never records)."""

    def record(self, **_kwargs: object) -> dict[str, object]:  # noqa: ANN003
        return {}


# ---------------------------------------------------------------------------
# Atomic write
# ---------------------------------------------------------------------------
def test_atomic_write_leaves_no_temp_files(rig: Rig) -> None:
    rig.store.save_artifact("shot", b"payload", artifact_type="screenshot")

    assert list(rig.artifacts.iterdir()) == [
        rig.artifacts / "index.json",
        rig.artifacts / "shot.png",
    ]


def test_failed_replace_leaves_no_partial_artifact(
    rig: Rig, monkeypatch: pytest.MonkeyPatch
) -> None:
    def _boom(src: object, dst: object) -> None:
        raise OSError("simulated replace failure")

    monkeypatch.setattr(artifact_store_module.os, "replace", _boom)

    with pytest.raises(ArtifactWriteError):
        rig.store.save_artifact("shot", b"payload", artifact_type="screenshot")

    # No artifact, no leftover temp file.
    assert list(rig.artifacts.iterdir()) == []
    assert rig.store.list_artifacts() == []


def test_repeated_save_replaces_atomically_and_lists_latest(rig: Rig) -> None:
    first = rig.store.save_artifact("report", b"v1", artifact_type="diagnostic")
    second = rig.store.save_artifact("report", b"v2", artifact_type="diagnostic")

    listed = rig.store.list_artifacts()

    assert first.sha256 != second.sha256
    assert [artifact.sha256 for artifact in listed] == [second.sha256]
    assert (rig.artifacts / "report.json").read_bytes() == b"v2"


# ---------------------------------------------------------------------------
# Path policy / type rejections
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "name",
    ["../evil", "..\\evil", "sub/dir", "sub\\dir", "/abs/path", "", "   "],
)
def test_traversal_and_absolute_names_rejected(rig: Rig, name: str) -> None:
    with pytest.raises(PathNotAllowedError) as excinfo:
        rig.store.save_artifact(name, b"x", artifact_type="log")
    assert excinfo.value.code.value == "PATH_NOT_ALLOWED"


def test_path_for_rejects_traversal(rig: Rig) -> None:
    with pytest.raises(PathNotAllowedError):
        rig.store.path_for("../escape.png")


def test_outside_root_is_rejected(tmp_path: Path) -> None:
    outside = tmp_path / "outside"
    policy = PathPolicy({"projects": str(tmp_path / "projects")})
    audit = AuditLogger(tmp_path / "audit.jsonl")
    try:
        store = ArtifactStore(outside, policy, audit)

        with pytest.raises(PathNotAllowedError) as excinfo:
            store.save_artifact("shot", b"x", artifact_type="screenshot")
        assert excinfo.value.code.value == "PATH_NOT_ALLOWED"
        assert not (outside / "shot.png").exists()
    finally:
        audit.close()


def test_disallowed_artifact_type_rejected(rig: Rig) -> None:
    with pytest.raises(PathNotAllowedError) as excinfo:
        rig.store.save_artifact("evil.exe", b"x", artifact_type="exe")
    assert excinfo.value.code.value == "PATH_NOT_ALLOWED"


def test_conflicting_extension_rejected(rig: Rig) -> None:
    with pytest.raises(PathNotAllowedError) as excinfo:
        rig.store.save_artifact("shot.bmp", b"x", artifact_type="screenshot")
    assert excinfo.value.code.value == "PATH_NOT_ALLOWED"


def test_extension_is_appended_when_missing(rig: Rig) -> None:
    info = rig.store.save_artifact("report", b"x", artifact_type="diagnostic")

    assert info.name == "report.json"
    assert Path(info.stored_path).name == "report.json"


# ---------------------------------------------------------------------------
# Audit / redaction
# ---------------------------------------------------------------------------
def test_audit_redacts_artifact_path_and_omits_payload(rig: Rig) -> None:
    secret = b"password=hunter2"
    rig.store.save_artifact("shot", secret, artifact_type="screenshot", workbook_id="wb-9")

    records = [record for record in rig.audit_records() if record["tool"] == "artifact_save"]
    assert len(records) == 1
    record = records[0]
    params = record["params"]
    assert isinstance(params, dict)
    # The full path never reaches the log; only the basename does.
    assert params["artifact_path"] == "shot.png"
    assert str(rig.artifacts) not in json.dumps(record)
    # Metadata is present, payload/secret is not.
    assert params["sha256"] == hashlib.sha256(secret).hexdigest()
    assert params["size"] == len(secret)
    assert "hunter2" not in json.dumps(record)
    assert record["result"] == {
        "artifact_id": record["params"]["artifact_id"],
        "sha256": record["params"]["sha256"],
        "size": record["params"]["size"],
        "saved": True,
    }


def test_redact_covers_artifact_metadata_keys() -> None:
    redacted = redact(
        {
            "artifact_path": r"C:\data\artifacts\shot.png",
            "artifact_name": r"C:\data\artifacts\shot.png",
            "password": "hunter2",
            "size": 12,
        }
    )

    assert redacted["artifact_path"] == "shot.png"
    assert redacted["artifact_name"] == "shot.png"
    assert redacted["password"] == "<redacted>"
    assert redacted["size"] == 12
