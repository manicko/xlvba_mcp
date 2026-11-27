"""Live-Excel verification of ``sheet_copy(target_name=...)`` over the real service path (WP4).

This drives the production
:class:`~excel_vba_mcp.application.sheets_service.SheetsService` and its STA-local
:class:`~excel_vba_mcp.adapters.excel.sheets_adapter.SheetsAdapter` against an
**automation-created hidden Excel instance** (never the user's Excel), on a
disposable working copy of the committed ``CleanProject.xlsm`` fixture. It covers
the WP4 ``copy_sheet`` ``target_name`` contract:

* an explicit ``target_name`` renames the new sheet to exactly that name;
* omitting ``target_name`` keeps Excel's default ``"<source> (2)"`` name;
* a duplicate ``target_name`` is rejected as ``INVALID_RANGE`` / ``validation`` /
  not retryable (asserted via ``ExcelVbaMcpError.to_dict()``), the source sheet is
  untouched and no second sheet is created under the duplicate destination name;
* an invalid ``target_name`` (a ``:`` character, or longer than 31 chars) is
  rejected as ``INVALID_RANGE`` and never appears as a sheet; the workbook state
  stays known.

The original fixture on disk is byte-identical after the test (the session only
ever mutates its own working copy under ``work/`` and the worker is shut down
without saving), and the owned Excel PID is reaped after shutdown.

Run serially (no xdist) on a host with Excel + VBOM trust::

    .venv\\Scripts\\python.exe -m pytest \\
        tests/integration_excel/test_sheets_copy_target_name_live.py -v -s
"""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import sys
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

# Optional COM dependencies (None when unavailable so the module skips cleanly).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-untyped]
    import win32com  # type: ignore[import-untyped]
    import win32gui  # type: ignore[import-untyped]
except Exception:  # noqa: BLE001 - optional dependency probe
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import OperationClass, SessionMode
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

pytestmark = [pytest.mark.anyio, pytest.mark.integration_excel]

_CLEAN_PROJECT = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "CleanProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5


def _pywin32_available() -> bool:
    return (
        sys.platform == "win32"
        and pythoncom is not None
        and win32com is not None
        and win32gui is not None
    )


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


def _fixture_or_skip(path: Path) -> Path:
    if not path.is_file():
        pytest.skip(f"fixture not available: {path}")
    return path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _excel_pids() -> set[int]:
    try:  # pragma: no cover - depends on host
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except Exception:  # noqa: BLE001 - best effort
        return set()
    pids: set[int] = set()
    for line in out.stdout.splitlines():
        line = line.strip()
        if line.startswith('"EXCEL.EXE"'):
            parts = [p.strip('"') for p in line.split('","')]
            if len(parts) >= 2 and parts[1].isdigit():
                pids.add(int(parts[1]))
    return pids


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if pid not in _excel_pids():
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return pid not in _excel_pids()


# ---------------------------------------------------------------------------
# Envelope assertion: a rejected copy is a specific, non-retryable, corrective error.
# ---------------------------------------------------------------------------
def _assert_invalid_range_taxonomy(error: ExcelVbaMcpError) -> dict[str, Any]:
    payload = error.to_dict()
    assert payload["errorCode"] == "INVALID_RANGE"
    assert payload["errorCode"] not in {"COM_BUSY", "RPC_E_CALL_REJECTED"}
    assert payload["errorCategory"] == "validation"
    assert payload["errorCategory"] != "transient_com_failure"
    assert payload["isRetryable"] is False
    actions = payload["suggestedNextActions"]
    assert actions and all(isinstance(action, str) and action for action in actions)
    assert "retry" not in " ".join(actions).lower()
    return payload


# ---------------------------------------------------------------------------
# Fixtures / rig (mirrors test_sheets_service.py)
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session + snapshot + sheets service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.snapshots_dir = tmp_dir / "snapshots"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.snapshots_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {
                "work": str(self.work_dir),
                "projects": str(self.projects_dir),
                "snapshots": str(self.snapshots_dir),
            }
        )
        self.approval = ApprovalPolicy()
        self.app_adapter = ExcelAppAdapter()
        self.worker = ExcelComWorker(self.app_adapter, orphan_timeout_s=_ORPHAN_POLL_TIMEOUT_S)
        self.audit = AuditLogger(tmp_dir / "audit.jsonl")
        self.session = SessionService(
            self.worker,
            self.app_adapter,
            self.path_policy,
            self.approval,
            self.audit,
            work_dir=self.work_dir,
        )
        self.snapshot = SnapshotService(
            SnapshotStore(self.snapshots_dir),
            self.session,
            self.worker,
            self.path_policy,
            self.audit,
        )
        self.service = SheetsService(
            self.worker,
            self.session,
            self.snapshot,
            self.approval,
            self.path_policy,
            self.audit,
        )
        self._original: Path | None = None
        self._writable = False

    def bind_original(self, original: Path, *, writable: bool = False) -> None:
        self._original = original
        self._writable = writable

    @property
    def canonical(self) -> str:
        assert self._original is not None
        return str(self.path_policy.validate_path(str(self._original), OperationClass.MUTATING))

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        # Shut the worker down without saving: the disposable working copy is
        # discarded and the original fixture on disk is never touched.
        await self.worker.shutdown()
        self.audit.close()

    async def open(self) -> Any:
        assert self._original is not None
        _grant(
            self.approval,
            ApprovalKey(canonical_path=self.canonical, action="excel_open"),
        )
        return await self.session.open_workbook(
            str(self._original),
            mode=SessionMode.INSPECT,
            read_only=not self._writable,
        )

    def grant_mutation(self, action: str) -> None:
        _grant(
            self.approval,
            ApprovalKey(canonical_path=self.canonical, action=action),
            scope=ApprovalScope.SESSION,
        )


def _grant(
    approval: ApprovalPolicy,
    key: ApprovalKey,
    *,
    scope: ApprovalScope = ApprovalScope.SESSION,
) -> None:
    approval.grant(
        ApprovalGrant(
            key=key,
            scope=scope,
            granted_at=datetime.now(UTC),
            granted_by="wp4-sheet-copy-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


def _make_rig(tmp_path: Path, *, writable: bool = True) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    fixture = _fixture_or_skip(_CLEAN_PROJECT)
    original = rig.projects_dir / fixture.name
    shutil.copy2(fixture, original)
    rig.bind_original(original, writable=writable)
    return rig


# ---------------------------------------------------------------------------
# The WP4 sheet_copy(target_name) live test
# ---------------------------------------------------------------------------
@requires_excel
async def test_sheet_copy_target_name_live(tmp_path: Path) -> None:
    """Explicit target_name renames; default keeps ``(2)``; dup/invalid rejected."""

    rig = _make_rig(tmp_path, writable=True)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]

        info = await rig.open()
        wid = info.workbook_id
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        snapshot = await rig.snapshot.create_snapshot(wid)
        assert snapshot.snapshot_id

        rig.grant_mutation("sheet_add")
        rig.grant_mutation("sheet_copy")

        # A disposable source sheet to copy from.
        source = await rig.service.add_sheet(wid, "SrcA")
        assert source.name == "SrcA"

        # --- 1. copy WITH explicit target_name -> renamed to exactly target_name ---
        copied_named = await rig.service.copy_sheet(wid, "SrcA", target_name="CopyA")
        assert copied_named.name == "CopyA"
        names = {s.name for s in await rig.service.list_sheets(wid)}
        assert "CopyA" in names and "SrcA" in names

        # --- 2. copy WITHOUT target_name -> Excel default "<source> (2)" preserved ---
        copied_default = await rig.service.copy_sheet(wid, "SrcA")
        assert copied_default.name == "SrcA (2)", (
            f"default copy name must be 'SrcA (2)', got {copied_default.name!r}"
        )
        names = {s.name for s in await rig.service.list_sheets(wid)}
        assert "SrcA (2)" in names and "SrcA" in names

        # --- 3. duplicate target_name -> INVALID_RANGE, source/dest unchanged ---
        copya_count_before = sum(
            1 for s in await rig.service.list_sheets(wid) if s.name == "CopyA"
        )
        with pytest.raises(ExcelVbaMcpError) as dup_exc:
            await rig.service.copy_sheet(wid, "SrcA", target_name="CopyA")
        _assert_invalid_range_taxonomy(dup_exc.value)
        assert dup_exc.value.context.get("reason") == "duplicate_target_name"
        # Source untouched; no SECOND sheet created under the duplicate name.
        copya_count_after = sum(
            1 for s in await rig.service.list_sheets(wid) if s.name == "CopyA"
        )
        assert copya_count_after == copya_count_before == 1, (
            "duplicate target_name must not create another 'CopyA' sheet"
        )
        assert "SrcA" in {s.name for s in await rig.service.list_sheets(wid)}

        # --- 4. invalid target_name -> INVALID_RANGE, workbook state known ---
        names_before_invalid = {s.name for s in await rig.service.list_sheets(wid)}

        with pytest.raises(ExcelVbaMcpError) as colon_exc:
            await rig.service.copy_sheet(wid, "SrcA", target_name="Bad:Name")
        _assert_invalid_range_taxonomy(colon_exc.value)
        assert colon_exc.value.context.get("reason") == "invalid_target_name"

        overlong = "X" * 32
        with pytest.raises(ExcelVbaMcpError) as long_exc:
            await rig.service.copy_sheet(wid, "SrcA", target_name=overlong)
        _assert_invalid_range_taxonomy(long_exc.value)
        assert long_exc.value.context.get("reason") == "invalid_target_name"

        # Neither invalid name ever became a sheet; the source survives.
        names_after_invalid = {s.name for s in await rig.service.list_sheets(wid)}
        assert "Bad:Name" not in names_after_invalid
        assert overlong not in names_after_invalid
        assert "SrcA" in names_after_invalid
        # The only permitted growth is Excel's auto-named copy/copies (no invalid name).
        assert names_after_invalid - names_before_invalid <= {
            name for name in names_after_invalid if name.startswith("SrcA (")
        }, (
            "invalid target_name must not add a sheet under the requested name; "
            f"new={sorted(names_after_invalid - names_before_invalid)}"
        )

        # Original + committed fixture never modified.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        print(
            f"[sheet-copy] named={copied_named.name} default={copied_default.name} "
            f"owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"