"""Unit tests for the snapshot service (task_014).

These tests never open Excel. They inject a fake STA worker (runs the submitted
callable inline) and a real :class:`SessionService` whose :class:`ExcelAppAdapter`
``_app`` is a fake ``Application`` supporting ``SaveCopyAs``. A real
:class:`SnapshotStore` backed by a temp directory persists snapshots.

Covered:

* create (open branch via SaveCopyAs fn) — ``snapshot_method="SaveCopyAs"`` and a
  manifest with the produced file's hash; the open workbook is not replaced.
* create (closed branch via copy) — ``snapshot_method="file_copy"``.
* restore closes the current working copy without saving, then opens a fresh
  working copy from the snapshot baseline.
* diff reports hash + metadata differences.
* the pre-mutation gate: ``SNAPSHOT_FAILED`` blocks a mutation with no snapshot.
"""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

import excel_vba_mcp.storage.snapshot_store as store_mod
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.snapshot_service import (
    SnapshotService,
    _export_vba_components,
)
from excel_vba_mcp.domain.errors import (
    ApprovalDeniedError,
    ExcelVbaMcpError,
    SnapshotFailedError,
)
from excel_vba_mcp.domain.models import WorkbookInfo
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    OperationClass,
    SessionMode,
    WorkbookSessionState,
)
from excel_vba_mcp.security.approval import (
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

pytestmark = pytest.mark.anyio


# ---------------------------------------------------------------------------
# Fakes (no Excel / no COM)
# ---------------------------------------------------------------------------
class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeRange:
    def __init__(self, sheet: FakeSheet, address: str) -> None:
        self._sheet = sheet
        self.Address = address
        self.Value: Any = None


class FakeWorkbook:
    """Fake ``Workbook`` supporting SaveCopyAs + close/save + a cell grid."""

    def __init__(self, app: FakeApp, path: str, *, read_only: bool) -> None:
        self._app = app
        self.FullName = path
        self.Name = path.replace("\\", "/").rsplit("/", 1)[-1]
        self.ReadOnly = read_only
        self.Saved = True
        self.closed = False
        self.saved_on_close: bool | None = None
        self.save_count = 0
        self.save_copy_as_targets: list[str] = []
        self.Worksheets = [FakeSheet("Sheet1")]
        self._ranges: dict[str, FakeRange] = {}

    def Close(self, SaveChanges: bool = False) -> None:  # noqa: N802, N803 - COM casing
        self.closed = True
        self.saved_on_close = SaveChanges
        self.Saved = SaveChanges

    def Save(self) -> None:  # noqa: N802 - COM casing
        self.save_count += 1
        self.Saved = True

    def SaveCopyAs(self, filename: str) -> None:  # noqa: N802 - COM casing
        # Emulate Excel: write a *separate* file, leaving the source untouched.
        Path(filename).write_bytes(self._app.serialized_bytes())
        self.save_copy_as_targets.append(filename)

    def Range(self, address: str) -> FakeRange:  # noqa: N802 - COM casing
        key = address.upper()
        if key not in self._ranges:
            self._ranges[key] = FakeRange(self, address)
        return self._ranges[key]

    @property
    def Application(self) -> FakeApp:  # noqa: N802 - COM casing
        return self._app


class FakeWorkbooks:
    def __init__(self, app: FakeApp) -> None:
        self._app = app
        self.items: list[FakeWorkbook] = []
        self.fail_open = False
        self.open_calls: list[dict[str, Any]] = []

    def Open(self, path: str, **kwargs: Any) -> FakeWorkbook:  # noqa: N802 - COM casing
        if self.fail_open:
            raise RuntimeError("simulated open failure")
        self.open_calls.append({"path": path, **kwargs})
        wb = FakeWorkbook(self._app, path, read_only=bool(kwargs.get("ReadOnly", False)))
        self.items.append(wb)
        return wb

    def __iter__(self):
        return iter(self.items)

    def __len__(self) -> int:
        return len(self.items)


class FakeApp:
    """Fake ``Application`` with a mutable settings set + Workbooks collection."""

    def __init__(self) -> None:
        self.Visible = False
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.AutomationSecurity = 1
        self.Version = "16.0"
        self.Caption = "Excel"
        self.Hwnd = 0
        self.Workbooks = FakeWorkbooks(self)

    def serialized_bytes(self) -> bytes:
        # A deterministic "save" payload distinct from the raw source bytes.
        return b"SERIALIZED::" + self.Version.encode("ascii")


class FakeWorker:
    """Fake STA worker: runs the submitted callable inline, no thread/COM."""

    def __init__(self, status: WorkerStatus | None = None) -> None:
        self.submits: list[dict[str, Any]] = []
        self._status = status or WorkerStatus(state=WorkbookSessionState.RUNNING, accepting=True)

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.submits.append(
            {"name": getattr(fn, "__name__", "?"), "section": section, "op_class": op_class}
        )
        return fn()

    def get_status(self) -> WorkerStatus:
        return self._status


# ---------------------------------------------------------------------------
# Fixtures / builders
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


def _grant(approval: ApprovalPolicy, canonical: str, action: str) -> None:
    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, action=action),
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="tester",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


class Rig:
    def __init__(self, tmp_path: Path) -> None:
        self.tmp_path = tmp_path
        self.original = tmp_path / "Book.xlsm"
        self.original.write_bytes(b"ORIGINAL-BYTES")
        self.snapshot_root = tmp_path / "snapshots"
        self.roots = {str(tmp_path): str(tmp_path)}
        self.path_policy = PathPolicy(self.roots)
        self.canonical = str(
            self.path_policy.validate_path(str(self.original), OperationClass.MUTATING)
        )
        self.approval = ApprovalPolicy(mode=AccessMode.FULL)
        self.app = FakeApp()
        self.adapter = ExcelAppAdapter()
        self.adapter._app = self.app  # inject without DispatchEx
        self.worker = FakeWorker()
        self.audit = AuditLogger(tmp_path / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=tmp_path / "work",
        )
        self.store = SnapshotStore(self.snapshot_root)
        self.service = SnapshotService(
            self.store,
            self.session,
            self.worker,  # type: ignore[arg-type]
            self.path_policy,
            self.audit,
            approval=self.approval,
        )

    def grant(self, action: str, canonical: str | None = None) -> None:
        _grant(self.approval, canonical or self.canonical, action)

    async def open(self, *, read_only: bool = False):
        self.grant("excel_open")
        return await self.session.open_workbook(
            str(self.original), mode=SessionMode.INSPECT, read_only=read_only
        )


# ---------------------------------------------------------------------------
# Create — open branch (SaveCopyAs)
# ---------------------------------------------------------------------------
async def test_create_open_branch_uses_save_copy_as(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)

    manifest = await rig.service.create_snapshot(info.workbook_id)

    assert manifest.snapshot_method == "SaveCopyAs"
    # SaveCopyAs ran on the worker (diagnostic section, READ_ONLY).
    submit = next(s for s in rig.worker.submits if s["name"] == "save_copy_as")
    assert submit["section"] == "diagnostic"
    assert submit["op_class"] == OperationClass.READ_ONLY

    # The open workbook is untouched: it is still open, unsaved, with the same id.
    wb = rig.app.Workbooks.items[-1]
    assert wb.closed is False
    assert wb.save_copy_as_targets, "SaveCopyAs must have been called"

    # The stored snapshot hashes to the manifest and the source file is unchanged.
    stored = rig.store.path_for(manifest.snapshot_id) / manifest.file_name
    assert stored.is_file()
    assert rig.store.sha256(stored) == manifest.sha256
    assert rig.original.read_bytes() == b"ORIGINAL-BYTES"

    # The service now knows this workbook has a snapshot (pre-mutation gate).
    assert rig.service.has_snapshot(info.workbook_id)
    assert rig.service.require_snapshot(info.workbook_id).snapshot_id == manifest.snapshot_id


async def test_create_open_branch_records_source_and_excel_version(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)

    manifest = await rig.service.create_snapshot(info.workbook_id)

    assert manifest.excel_version == "16.0"
    assert Path(manifest.source_path) == rig.original.resolve()
    # The model name is the working-copy file name (a uuid.xlsm), not the source.
    assert manifest.workbook_model.get("name", "").endswith(".xlsm")
    assert Path(manifest.workbook_model.get("source_path", "")) == rig.original.resolve()
    assert manifest.external_side_effects_not_undone is True


# ---------------------------------------------------------------------------
# Create — closed branch (file copy)
# ---------------------------------------------------------------------------
async def test_create_closed_branch_uses_file_copy(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    # Register a known source for a never-opened (closed) workbook id.
    closed_id = "closed-wb"
    rig.service._known_sources[closed_id] = str(rig.original)

    manifest = await rig.service.create_snapshot(closed_id)

    assert manifest.snapshot_method == "file_copy"
    assert manifest.sha256 == rig.store.sha256(rig.original)
    # No COM was used for the closed branch.
    assert all(s["name"] != "save_copy_as" for s in rig.worker.submits)
    assert rig.service.require_snapshot(closed_id).snapshot_id == manifest.snapshot_id


async def test_create_unknown_workbook_raises(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(ExcelVbaMcpError):
        await rig.service.create_snapshot("ghost")


# ---------------------------------------------------------------------------
# Restore — close-without-save then open baseline
# ---------------------------------------------------------------------------
async def test_restore_closes_without_save_and_opens_baseline(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)
    manifest = await rig.service.create_snapshot(info.workbook_id)

    # Mutate the working copy in memory (do NOT save) before restore.
    wb = rig.app.Workbooks.items[-1]
    wb.Range("A1").Value = "MUTATED"

    # Restore opens a writable working copy from the snapshot file; grant approval
    # for that canonical path (writable opens are MUTATING and deny-by-default).
    snapshot_file = rig.store.path_for(manifest.snapshot_id) / manifest.file_name
    snapshot_canonical = str(
        rig.path_policy.validate_path(str(snapshot_file), OperationClass.MUTATING)
    )
    rig.grant("excel_open", canonical=snapshot_canonical)

    restored = await rig.service.restore_snapshot(
        manifest.snapshot_id, workbook_id=info.workbook_id
    )

    # The old working copy was closed WITHOUT saving.
    assert wb.closed is True
    assert wb.saved_on_close is False

    # The restored session is execute-mode and writable (FE-16): this is what lets
    # the Full-Edit workflow continue running macros after a restore instead of
    # failing with MACRO_BLOCKED on an inspect-mode session.
    assert isinstance(restored, WorkbookInfo)
    assert restored.mode == SessionMode.EXECUTE
    assert restored.read_only is False

    # A fresh working copy was opened from the snapshot baseline.
    assert restored.workbook_id != info.workbook_id
    new_wb = next(w for w in rig.app.Workbooks.items if w.FullName == restored.path)
    assert new_wb.Range("A1").Value is None  # baseline (unmutated)
    assert new_wb.closed is False

    # The immutable snapshot file is unchanged.
    assert rig.store.load(manifest.snapshot_id).sha256 == manifest.sha256


# ---------------------------------------------------------------------------
# Diff
# ---------------------------------------------------------------------------
async def test_diff_reports_hash_and_metadata(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)
    a = await rig.service.create_snapshot(info.workbook_id)

    # Close and take a closed-branch snapshot of the same source (file copy).
    await rig.session.close_workbook(info.workbook_id)
    rig.service._known_sources["closed"] = str(rig.original)
    b = await rig.service.create_snapshot("closed")

    diff = await rig.service.diff_snapshots(a.snapshot_id, b.snapshot_id)

    # SaveCopyAs re-serializes, so the file hashes differ.
    assert diff.file_identical is False
    assert diff.sha256_a == a.sha256
    assert diff.sha256_b == b.sha256
    # The snapshot_method metadata differs (SaveCopyAs vs file_copy).
    assert "snapshot_method" in diff.changed
    assert diff.changed["snapshot_method"]["a"] == "SaveCopyAs"
    assert diff.changed["snapshot_method"]["b"] == "file_copy"


async def test_diff_identical_files(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    src = tmp_path / "Book.xlsm"
    # Two closed-branch snapshots of byte-identical sources -> same hash.
    rig.store.create("s1", source_path=src)
    rig.store.create("s2", source_path=src)
    diff = await rig.service.diff_snapshots("s1", "s2")
    assert diff.file_identical is True


# ---------------------------------------------------------------------------
# List
# ---------------------------------------------------------------------------
async def test_list_snapshots(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)
    await rig.service.create_snapshot(info.workbook_id)
    listed = await rig.service.list_snapshots()
    assert len(listed) == 1
    assert listed[0].snapshot_method == "SaveCopyAs"


# ---------------------------------------------------------------------------
# Pre-mutation gate: SNAPSHOT_FAILED blocks the mutation
# ---------------------------------------------------------------------------
def test_require_snapshot_blocks_mutation_without_snapshot(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    with pytest.raises(SnapshotFailedError) as exc_info:
        rig.service.require_snapshot("never-snapshotted")
    assert exc_info.value.code.value == "SNAPSHOT_FAILED"


def test_require_snapshot_returns_latest(tmp_path: Path) -> None:
    rig = Rig(tmp_path)
    rig.store.create("s1", source_path=tmp_path / "Book.xlsm")
    rig.store.create("s2", source_path=tmp_path / "Book.xlsm")
    # Simulate the service having recorded s1 then s2 for the workbook.
    rig.service._latest_snapshot["wb"] = rig.store.load("s1")
    rig.service._latest_snapshot["wb"] = rig.store.load("s2")
    assert rig.service.require_snapshot("wb").snapshot_id == "s2"


# ---------------------------------------------------------------------------
# F-4 / F-9: approval gate + audit classification
# ---------------------------------------------------------------------------
async def test_create_restricted_mode_requires_approval(tmp_path: Path) -> None:
    """In restricted mode a snapshot create is denied by default (F-4)."""

    rig = Rig(tmp_path)
    restricted = ApprovalPolicy(mode=AccessMode.RESTRICTED)
    service = SnapshotService(
        rig.store,
        rig.session,
        rig.worker,  # type: ignore[arg-type]
        rig.path_policy,
        rig.audit,
        approval=restricted,
    )
    # A read-only inspect open is the only session restricted mode admits; a
    # snapshot create on it must still be denied (no grant over MCP).
    info = await rig.open(read_only=True)

    with pytest.raises(ApprovalDeniedError) as exc_info:
        await service.create_snapshot(info.workbook_id)
    assert exc_info.value.code.value == "APPROVAL_DENIED"


async def test_create_full_mode_succeeds_without_grant(tmp_path: Path) -> None:
    """In full mode a snapshot create is auto-allowed (no behavior change)."""

    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)

    manifest = await rig.service.create_snapshot(info.workbook_id)

    assert manifest.snapshot_id.startswith("snap-")


async def test_create_audit_records_mutating(tmp_path: Path) -> None:
    """The snapshot_create audit entry is classified ``mutating`` (F-9)."""

    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)

    await rig.service.create_snapshot(info.workbook_id)
    rig.audit.flush()

    records = [
        json.loads(line)
        for line in (tmp_path / "audit.jsonl").read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    create_events = [record for record in records if record["tool"] == "snapshot_create"]
    assert create_events, "snapshot_create must be audited"
    assert create_events[-1]["operation_class"] == "mutating"


# ---------------------------------------------------------------------------
# F7: export file hash renamed to export_file_sha256 + model_version bump
# ---------------------------------------------------------------------------
class _FakeComponent:
    """Minimal ``VBComponent`` that writes deterministic bytes on ``Export``."""

    def __init__(self, name: str, ctype: int, payload: bytes) -> None:
        self.Name = name
        self.Type = ctype
        self._payload = payload

    def Export(self, path: str) -> None:  # noqa: N802 - COM casing
        Path(path).write_bytes(self._payload)


class _FakeComponents:
    def __init__(self, items: list[_FakeComponent]) -> None:
        self._items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._items)

    def Item(self, index: int) -> _FakeComponent:  # noqa: N802 - COM casing
        return self._items[index - 1]


class _FakeProject:
    def __init__(self, items: list[_FakeComponent]) -> None:
        self.VBComponents = _FakeComponents(items)


class _FakeWorkbookWithVba:
    def __init__(self, items: list[_FakeComponent]) -> None:
        self.VBProject = _FakeProject(items)


def test_export_vba_components_uses_export_file_sha256(tmp_path: Path) -> None:
    """``_export_vba_components`` names the file hash ``export_file_sha256`` (F7)."""

    import hashlib

    payload = b"Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n"
    wb = _FakeWorkbookWithVba([_FakeComponent("Module1", 1, payload)])

    entries = _export_vba_components(wb, tmp_path / "exports")

    assert entries is not None
    assert len(entries) == 1
    entry = entries[0]
    # The exported-file byte hash is exposed under the unambiguous name ...
    assert entry["export_file_sha256"] == hashlib.sha256(payload).hexdigest()
    # ... and the old ambiguous ``sha256`` key is gone.
    assert "sha256" not in entry
    assert entry["component_name"] == "Module1"
    assert entry["type"] == 1
    assert entry["path"] == "exports/Module1.bas"


async def test_manifest_model_version_bumped(tmp_path: Path) -> None:
    """New manifests are stamped with the bumped ``model_version`` (F7)."""

    rig = Rig(tmp_path)
    info = await rig.open(read_only=False)

    manifest = await rig.service.create_snapshot(info.workbook_id)

    assert manifest.model_version == store_mod._MANIFEST_SCHEMA_VERSION
    assert manifest.model_version >= 2
    # The on-disk manifest.json carries the same bumped version.
    raw = json.loads(
        (rig.store.path_for(manifest.snapshot_id) / "manifest.json").read_text(
            encoding="utf-8"
        )
    )
    assert raw["model_version"] >= 2


async def test_snapshot_list_surfaces_export_file_sha256(tmp_path: Path) -> None:
    """``snapshot_list`` surfaces ``export_file_sha256`` on vba_exports (F7)."""

    rig = Rig(tmp_path)
    src = tmp_path / "Book.xlsm"
    src.write_bytes(b"workbook")
    entry = {
        "component_name": "Module1",
        "type": 1,
        "path": "exports/Module1.bas",
        "export_file_sha256": "a" * 64,
    }
    rig.store.create("snap-e", source_path=src, vba_exports=[entry])

    listed = await rig.service.list_snapshots()

    manifest = next(m for m in listed if m.snapshot_id == "snap-e")
    assert manifest.vba_exports is not None
    assert manifest.vba_exports[0]["export_file_sha256"] == "a" * 64
    assert "sha256" not in manifest.vba_exports[0]


async def test_diff_surfaces_export_file_sha256_change(tmp_path: Path) -> None:
    """``snapshot_diff`` still compares manifests by the renamed field (F7)."""

    rig = Rig(tmp_path)
    src = tmp_path / "Book.xlsm"
    src.write_bytes(b"workbook")
    rig.store.create(
        "s1",
        source_path=src,
        vba_exports=[
            {
                "component_name": "m",
                "type": 1,
                "path": "exports/m.bas",
                "export_file_sha256": "a" * 64,
            }
        ],
    )
    rig.store.create(
        "s2",
        source_path=src,
        vba_exports=[
            {
                "component_name": "m",
                "type": 1,
                "path": "exports/m.bas",
                "export_file_sha256": "b" * 64,
            }
        ],
    )

    diff = await rig.service.diff_snapshots("s1", "s2")

    # The manifests differ only in the exported-file hash, surfaced as a change.
    assert "vba_exports" in diff.changed
    assert diff.changed["vba_exports"]["a"][0]["export_file_sha256"] == "a" * 64
    assert diff.changed["vba_exports"]["b"][0]["export_file_sha256"] == "b" * 64


def test_load_migrates_legacy_vba_exports_sha256(tmp_path: Path) -> None:
    """An old-format manifest (v1 ``sha256``, no bumped version) migrates on load.

    PO4 = B chose a hard rename + ``model_version`` bump; loading an older on-disk
    manifest must not crash — it migrates ``sha256`` -> ``export_file_sha256`` in
    memory and stamps the current ``model_version`` (F7).
    """

    store = SnapshotStore(tmp_path / "snapshots")
    src = tmp_path / "Book.xlsm"
    src.write_bytes(b"legacy")
    store.create("snap-old", source_path=src)

    # Rewrite the manifest in the pre-F7 (v1) shape.
    manifest_path = store.path_for("snap-old") / "manifest.json"
    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    raw["model_version"] = 1
    raw["vba_exports"] = [
        {"component_name": "mod", "type": 1, "path": "exports/mod.bas", "sha256": "deadbeef"}
    ]
    manifest_path.write_text(json.dumps(raw), encoding="utf-8")

    loaded = store.load("snap-old")

    assert loaded.model_version == store_mod._MANIFEST_SCHEMA_VERSION
    assert loaded.vba_exports is not None
    migrated = loaded.vba_exports[0]
    assert migrated["export_file_sha256"] == "deadbeef"
    assert "sha256" not in migrated
    # The on-disk file is not rewritten (migration is in-memory only).
    on_disk = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert on_disk["model_version"] == 1
    assert on_disk["vba_exports"][0]["sha256"] == "deadbeef"


def test_load_without_model_version_migrates(tmp_path: Path) -> None:
    """A manifest missing ``model_version`` is treated as v1 and migrated (F7)."""

    store = SnapshotStore(tmp_path / "snapshots")
    src = tmp_path / "Book.xlsm"
    src.write_bytes(b"legacy")
    store.create("snap-nov", source_path=src)

    manifest_path = store.path_for("snap-nov") / "manifest.json"
    raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    raw.pop("model_version", None)
    raw["vba_exports"] = [
        {"component_name": "mod", "type": 1, "path": "exports/mod.bas", "sha256": "c0ffee"}
    ]
    manifest_path.write_text(json.dumps(raw), encoding="utf-8")

    loaded = store.load("snap-nov")

    assert loaded.model_version == store_mod._MANIFEST_SCHEMA_VERSION
    assert loaded.vba_exports is not None
    assert loaded.vba_exports[0]["export_file_sha256"] == "c0ffee"
    assert "sha256" not in loaded.vba_exports[0]
