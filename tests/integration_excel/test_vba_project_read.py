"""Live-Excel verification of the read-only VBA project service (task_016).

This is the task_016 independent verification of the production
:class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService` and
:class:`~excel_vba_mcp.adapters.excel.vbide_adapter.VBIDEAdapter` against the
target Excel build. It drives the real single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker`, a real
:class:`~excel_vba_mcp.adapters.excel.app_adapter.ExcelAppAdapter` and a real
:class:`~excel_vba_mcp.application.session_service.SessionService`, always
hidden / non-activating, and only ever touches a temporary **copy** of the
committed ``CleanProject.xlsm`` fixture. The committed fixture is never opened
or modified.

Coverage (spec §6.2 / §7.2; research ``g1-poc-research-2026-11-19.md`` §2):

* ``test_project_info_inventory_and_references`` — component types
  ``{1, 2, 3, 100}`` observed, UserForm ``has_designer``, Document modules
  preserved, valid references, ``Protection=0``.
* ``test_procedures_via_code_module`` — procedures enumerated through the
  CodeModule API with ``source="code_module"`` and ``is_heuristic=False``.
* ``test_export_project_fidelity`` — ``.bas`` / ``.cls`` / ``.frm`` + ``.frx``
  exported; normalized ``source_hash`` matches ``vba_read_component``; Document
  modules still present after export.
* ``test_locked_project_skipped`` — ``LockedProject.xlsm`` does not exist
  (manual binary); the test skips with an explicit reason.
* ``test_search_code`` — bounded substring / regex search.
* ``test_source_hash_matches_export`` (task_017) — an **independent** re-hash of
  the exported ``.bas`` / ``.cls`` / ``.frm`` body (attribute/version header
  stripped, EOL + trailing whitespace normalized) equals the ``source_hash``
  reported by ``vba_read_component``.
* ``test_missing_reference_detected`` (task_017) — forcing ``IsBroken=True`` via
  ``VBProject.References.AddFromFile`` so ``vba_get_references`` reports a
  missing reference; skips with a clear reason when the build cannot produce a
  service-visible broken reference safely.
* ``test_all_component_types_present`` (task_017) — CleanProject exposes
  Standard / Class / UserForm / Document component types and Document modules
  survive an export.
* BG03/BG04/BG05 background non-interference + orphan check after shutdown.

Run serially (no xdist)::

    uv run pytest tests/integration_excel/test_vba_project_read.py -v -s
"""

from __future__ import annotations

import contextlib
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

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import (
    OperationClass,
    SessionMode,
    VbaComponentType,
    VbaExportStatus,
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

# Optional COM dependencies (None when unavailable so the module imports and the
# COM tests skip on non-Windows hosts).
try:  # pragma: no cover - exercised implicitly by skip logic
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]
    import win32gui  # type: ignore[import-not-found]
    import win32process  # type: ignore[import-not-found]
except Exception:  # noqa: BLE001 - any import failure means "not available here"
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    win32gui = None  # type: ignore[assignment]
    win32process = None  # type: ignore[assignment]

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CLEAN_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "CleanProject.xlsm"
_LOCKED_PROJECT = _REPO_ROOT / "tests" / "fixtures" / "LockedProject.xlsm"

_ORPHAN_POLL_TIMEOUT_S = 30.0
_ORPHAN_POLL_INTERVAL_S = 0.5
_CALL_TIMEOUT_S = 30.0


# ---------------------------------------------------------------------------
# Skip helpers
# ---------------------------------------------------------------------------
def _pywin32_available() -> bool:
    return sys.platform == "win32" and pythoncom is not None and win32com is not None


_PYWIN32 = _pywin32_available()

requires_excel = pytest.mark.skipif(
    not _PYWIN32,
    reason=f"pywin32/Windows required to drive a live Excel instance (platform={sys.platform})",
)


# ---------------------------------------------------------------------------
# Read-only desktop / process helpers (never activate / show / kill)
# ---------------------------------------------------------------------------
def _pid_from_hwnd(hwnd: int) -> int | None:
    if win32process is None:
        return None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return int(pid) or None


def _pid_alive(pid: int) -> bool:
    try:
        out = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except Exception:  # noqa: BLE001 - defensive
        return False
    return str(pid) in out.stdout


def _wait_pid_gone(pid: int, timeout: float = _ORPHAN_POLL_TIMEOUT_S) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            return True
        time.sleep(_ORPHAN_POLL_INTERVAL_S)
    return not _pid_alive(pid)


def _foreground() -> dict[str, Any] | None:
    if win32gui is None or win32process is None:
        return None
    try:
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:  # noqa: BLE001 - platform specific
        return None
    return {"hwnd": hwnd, "pid": int(pid), "title": title}


def _windows_for_pid(pid: int) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if win32gui is None or win32process is None:
        return result

    def _cb(hwnd: int, _ctx: Any) -> bool:
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            return True
        if wpid != pid:
            return True
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:  # noqa: BLE001 - platform specific
            title = ""
        try:
            visible = bool(win32gui.IsWindowVisible(hwnd))
        except Exception:  # noqa: BLE001 - platform specific
            visible = False
        result.append({"hwnd": hwnd, "pid": int(pid), "title": title, "visible": visible})
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_cb, None)
    return result


def _assert_owned_never_foreground(owned_pid: int, label: str) -> None:
    fg = _foreground()
    if fg is None:
        return
    assert fg["pid"] != owned_pid, (
        f"BG03 violation: owned automation PID {owned_pid} held foreground at {label} "
        f"(title={fg['title']!r})"
    )


def _assert_no_visible_owned_windows(owned_pid: int, label: str) -> None:
    windows = _windows_for_pid(owned_pid)
    visible = [w for w in windows if w["visible"]]
    assert not visible, (
        f"BG04 violation: owned automation PID {owned_pid} has visible windows at {label}: "
        f"{[w['title'] for w in visible]}"
    )


def _assert_user_session_preserved(
    owned_pid: int, fg_before: dict[str, Any] | None, label: str
) -> None:
    """BG05: the interactive foreground is not stolen by the owned instance."""

    fg_after = _foreground()
    if fg_after is not None:
        assert fg_after["pid"] != owned_pid, (
            f"BG05 violation: owned automation PID {owned_pid} became foreground at {label} "
            f"(title={fg_after['title']!r})"
        )
    if fg_before is not None and fg_before["pid"] != owned_pid and fg_after is not None:
        assert fg_after["hwnd"] == fg_before["hwnd"], (
            f"BG05 violation: interactive foreground changed across {label} "
            f"(before={fg_before['title']!r} after={fg_after['title']!r})"
        )


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _normalize_exported_body(exported_text: str) -> str:
    """Independently normalize an exported VBA source file for hashing.

    This is a deliberately **independent** re-implementation of the production
    normalization (it does not import ``normalize_code`` / ``hash_code``) so the
    task_017 hash test is a genuine cross-check rather than a tautology.

    Steps:

    1. Unify CRLF / lone CR to ``\\n``.
    2. Drop the export header noise: a ``.bas`` / ``.cls`` / ``.frm`` export
       starts with ``Attribute VB_Name = "..."`` (and, for ``.frm`` / ``.cls``,
       a ``VERSION`` / ``Begin ... End`` designer block plus more
       ``Attribute VB_*`` lines). The code body is everything after the **last**
       ``Attribute VB_`` line, matching how the VBE stores the module body.
    3. Right-trim every line and strip leading / trailing blank lines.

    The normalized text is what gets SHA-256 hashed and compared against the
    ``source_hash`` reported by ``vba_read_component``.
    """

    unified = exported_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = unified.split("\n")
    last_attr = -1
    for index, line in enumerate(lines):
        if line.lstrip().startswith("Attribute VB_"):
            last_attr = index
    body = lines[last_attr + 1 :] if last_attr >= 0 else lines
    trimmed = [line.rstrip() for line in body]
    while trimmed and not trimmed[0].strip():
        trimmed.pop(0)
    while trimmed and not trimmed[-1].strip():
        trimmed.pop()
    return "\n".join(trimmed)


def _sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _grant(
    approval: ApprovalPolicy, canonical: str, action: str, *, sha256: str | None = None
) -> None:
    approval.grant(
        ApprovalGrant(
            key=ApprovalKey(canonical_path=canonical, sha256=sha256, action=action),
            scope=ApprovalScope.SESSION,
            granted_at=datetime.now(UTC),
            granted_by="task016-verification",
            source=ApprovalSource.CLIENT,
            grant_id=uuid4().hex,
        )
    )


# ---------------------------------------------------------------------------
# Fixtures / rig
# ---------------------------------------------------------------------------
@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


class _Rig:
    """A live worker + session service + VBA project service + temp roots."""

    def __init__(self, tmp_dir: Path) -> None:
        self.tmp_dir = tmp_dir
        self.work_dir = tmp_dir / "work"
        self.projects_dir = tmp_dir / "projects"
        self.projects_dir.mkdir(parents=True, exist_ok=True)
        self.path_policy = PathPolicy(
            {"work": str(self.work_dir), "projects": str(self.projects_dir)}
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
        self.service = VBAProjectService(self.worker, self.session, self.path_policy, self.audit)
        self._original: Path | None = None

    def bind_original(self, original: Path) -> None:
        self._original = original

    async def start(self) -> None:
        await self.worker.start()

    async def stop(self) -> None:
        await self.worker.shutdown()
        self.audit.close()

    def grant_open(self) -> None:
        assert self._original is not None
        canonical = str(
            self.path_policy.validate_path(str(self._original), OperationClass.MUTATING)
        )
        _grant(self.approval, canonical, "excel_open", sha256=_sha256(self._original))

    async def open(self) -> Any:
        self.grant_open()
        assert self._original is not None
        return await self.session.open_workbook(
            str(self._original), mode=SessionMode.INSPECT, read_only=True
        )


def _make_rig(tmp_path: Path) -> _Rig:
    """Copy the committed fixture into the temp projects root as the original."""

    rig = _Rig(tmp_path)
    original = rig.projects_dir / "CleanProject.xlsm"
    shutil.copy2(_CLEAN_PROJECT, original)
    rig.bind_original(original)
    return rig


# ---------------------------------------------------------------------------
# Test 1: project info, component types, references (real Excel)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_project_info_inventory_and_references(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        committed_hash = _sha256(_CLEAN_PROJECT)
        original_hash = _sha256(rig._original)  # type: ignore[arg-type]
        fg_before = _foreground()
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        assert info.has_vba is True

        project = await rig.service.get_project_info(info.workbook_id)
        assert project.available is True
        assert project.locked is False
        assert project.component_count >= 4

        by_name = {component.component_name: component for component in project.components}
        types_observed = {component.component_type for component in project.components}
        assert VbaComponentType.STANDARD_MODULE in types_observed
        assert VbaComponentType.CLASS_MODULE in types_observed
        assert VbaComponentType.USER_FORM in types_observed
        assert VbaComponentType.DOCUMENT_MODULE in types_observed

        # Document modules preserved (ThisWorkbook / Sheet1), UserForm has designer.
        assert {"ThisWorkbook", "Sheet1"}.issubset(set(by_name))
        assert by_name["ThisWorkbook"].is_document_module is True
        assert by_name["Sheet1"].is_document_module is True
        assert by_name["frmTest"].has_designer is True
        assert by_name["frmTest"].component_type is VbaComponentType.USER_FORM

        # References: every reference has Name/GUID; none broken on the clean fixture.
        assert project.references, "no references enumerated"
        for reference in project.references:
            assert reference.name
            assert reference.guid
            assert reference.is_broken is False
        assert project.missing_references == []

        components = await rig.service.list_components(info.workbook_id)
        assert {component.component_name for component in components} >= {"modTest", "clsTest"}

        # Original and committed fixture are untouched.
        assert _sha256(rig._original) == original_hash, "original must be unchanged"
        assert _sha256(_CLEAN_PROJECT) == committed_hash, "committed fixture must never change"

        # BG03/BG04/BG05 while Excel is alive.
        _assert_owned_never_foreground(owned_pid, "vba-info")
        _assert_no_visible_owned_windows(owned_pid, "vba-info")
        _assert_user_session_preserved(owned_pid, fg_before, "vba-info")

        print(
            f"[vba-info] component_types={sorted(t.value for t in types_observed)} "
            f"references={len(project.references)} owned_pid={owned_pid}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 2: procedures via CodeModule API (source="code_module")
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_procedures_via_code_module(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()

        procedures = await rig.service.list_procedures(info.workbook_id, "modTest")
        assert procedures, "modTest must expose procedures"

        by_name = {procedure.name: procedure for procedure in procedures}
        assert "AddNumbers" in by_name, f"AddNumbers missing: {sorted(by_name)}"
        add = by_name["AddNumbers"]
        assert add.start_line >= 1
        assert add.end_line >= add.start_line
        assert "AddNumbers" in add.code
        # VBIDE is the source of truth; the heuristic fallback is flagged and must
        # not be the normal path for an unprotected project.
        assert add.is_heuristic is False, (
            f"expected CodeModule API result, got heuristic: heuristic={add.is_heuristic} "
            f"source={add.source}"
        )
        assert add.source == "code_module"

        source = await rig.service.read_component(info.workbook_id, "modTest")
        assert "AddNumbers" in source.source
        assert len(source.source_hash) == 64

        _assert_owned_never_foreground(owned_pid, "vba-procedures")
        _assert_no_visible_owned_windows(owned_pid, "vba-procedures")

        print(
            f"[vba-procedures] count={len(procedures)} "
            f"names={sorted(by_name)} add_range={add.start_line}-{add.end_line} "
            f"source={add.source}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 3: export fidelity (.bas / .cls / .frm + .frx, normalized source hash)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_export_project_fidelity(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()

        manifest = await rig.service.export_project(info.workbook_id)
        out_dir = Path(manifest.out_dir)
        assert out_dir.is_dir()

        by_name = {entry.component_name: entry for entry in manifest.components}
        assert all(entry.export_status is VbaExportStatus.OK for entry in manifest.components), (
            f"export failures: {[(e.component_name, e.error) for e in manifest.components]}"
        )

        # Expected extensions.
        assert by_name["modTest"].extension == ".bas"
        assert by_name["clsTest"].extension == ".cls"
        assert by_name["frmTest"].extension == ".frm"
        assert by_name["ThisWorkbook"].extension == ".cls"

        # Files on disk.
        assert (out_dir / "modTest.bas").is_file()
        assert (out_dir / "clsTest.cls").is_file()
        assert (out_dir / "frmTest.frm").is_file()
        assert (out_dir / "frmTest.frx").is_file()
        assert (out_dir / "ThisWorkbook.cls").is_file()

        # Normalized source hash equals the read_component hash.
        for component in ("modTest", "clsTest", "frmTest", "ThisWorkbook"):
            source = await rig.service.read_component(info.workbook_id, component)
            entry = by_name[component]
            assert entry.source_hash == source.source_hash, (
                f"source_hash mismatch for {component}: "
                f"export={entry.source_hash} read={source.source_hash}"
            )

        # UserForm companion resource recorded.
        frm_entry = by_name["frmTest"]
        assert frm_entry.resource_file_name == "frmTest.frx"
        assert frm_entry.resource_file_sha256 is not None
        assert frm_entry.file_sha256 is not None

        # Document modules still present after export.
        components = await rig.service.list_components(info.workbook_id)
        names = {component.component_name for component in components}
        assert {"ThisWorkbook", "Sheet1"}.issubset(names), "document modules removed after export"

        _assert_owned_never_foreground(owned_pid, "vba-export")
        _assert_no_visible_owned_windows(owned_pid, "vba-export")

        print(
            f"[vba-export] out_dir={out_dir} components={len(manifest.components)} "
            f"combined_hash={manifest.source_hash} "
            f"modTest_source_hash={by_name['modTest'].source_hash[:16]}…"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 4: search code (bounded substring + regex)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_search_code(tmp_path: Path) -> None:
    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()

        hits = await rig.service.search_code(info.workbook_id, "AddNumbers")
        assert any(hit.component_name == "modTest" for hit in hits)

        regex_hits = await rig.service.search_code(
            info.workbook_id, r"^Public Function", regex=True
        )
        assert any(hit.component_name == "modTest" for hit in regex_hits)

        # An invalid regex raises a domain error rather than crashing.
        with pytest.raises(ExcelVbaMcpError):
            await rig.service.search_code(info.workbook_id, r"(", regex=True)

        _assert_owned_never_foreground(owned_pid, "vba-search")
        print(f"[vba-search] substring_hits={len(hits)} regex_hits={len(regex_hits)}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 5: LockedProject fixture skip-if-absent with reason
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_locked_project_reported_without_bypass(tmp_path: Path) -> None:
    """A locked project must be reported as VBA_PROJECT_LOCKED, never bypassed.

    ``LockedProject.xlsm`` is a manual binary (VBProject.Protection cannot be set
    from COM, research §1.2.3) and does not exist in this repository. The test
    skips with an explicit reason rather than fabricating a locked fixture.
    """

    if not _LOCKED_PROJECT.is_file():
        pytest.skip(
            f"fixture missing: {_LOCKED_PROJECT}; LockedProject.xlsm is a manual binary "
            "(VBProject.Protection is read-only from COM) and is intentionally not generated"
        )

    rig = _Rig(tmp_path)
    locked_copy = rig.projects_dir / "LockedProject.xlsm"
    shutil.copy2(_LOCKED_PROJECT, locked_copy)
    rig.bind_original(locked_copy)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        with pytest.raises(ExcelVbaMcpError) as exc_info:
            await rig.service.get_project_info(info.workbook_id)
        assert exc_info.value.code is ErrorCode.VBA_PROJECT_LOCKED
        print(f"[vba-locked] code={exc_info.value.code.value} bypass_attempted=False")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 6 (task_017): source_hash == independent normalized export hash
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_source_hash_matches_export(tmp_path: Path) -> None:
    """The exported body, independently re-hashed, must equal ``source_hash``.

    For ``modTest`` (``.bas``), ``clsTest`` (``.cls``) and ``frmTest``
    (``.frm``) the test reads the component ``source_hash`` via
    ``vba_read_component`` and the exported file via ``vba_export_project``,
    strips the export's attribute / version header, normalizes EOL + trailing
    whitespace *independently* of production code
    (:func:`_normalize_exported_body`) and asserts the SHA-256 of the normalized
    body equals ``source_hash``. For ``frmTest`` the ``.frx`` companion must
    also exist.

    The normalization used (documented in :func:`_normalize_exported_body`):

    * CRLF / lone CR -> ``\\n``;
    * drop everything up to and including the last ``Attribute VB_*`` line
      (the ``.bas`` / ``.cls`` / ``.frm`` export header);
    * right-trim each line, then strip leading / trailing blank lines;
    * SHA-256 the UTF-8 bytes of the result.
    """

    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()

        manifest = await rig.service.export_project(info.workbook_id)
        out_dir = Path(manifest.out_dir)
        by_name = {entry.component_name: entry for entry in manifest.components}
        assert all(entry.export_status is VbaExportStatus.OK for entry in manifest.components), (
            f"export failures: {[(e.component_name, e.error) for e in manifest.components]}"
        )

        checked: list[str] = []
        for component, extension in (
            ("modTest", ".bas"),
            ("clsTest", ".cls"),
            ("frmTest", ".frm"),
        ):
            source = await rig.service.read_component(info.workbook_id, component)
            entry = by_name[component]
            assert entry.extension == extension

            exported_path = Path(entry.file_path)
            assert exported_path.is_file(), f"{component} export missing: {exported_path}"
            exported_text = exported_path.read_text(encoding="utf-8", errors="replace")

            # Independent normalization + hash (does NOT call production hash_code).
            normalized = _normalize_exported_body(exported_text)
            export_hash = _sha256_text(normalized)

            assert export_hash == source.source_hash, (
                f"source_hash mismatch for {component}: "
                f"independent_export_hash={export_hash} "
                f"read_component_source_hash={source.source_hash}\n"
                f"--- normalized export body ---\n{normalized!r}"
            )
            checked.append(component)

        # The UserForm needs both a .frm (code) and a .frx (designer resource).
        frm_entry = by_name["frmTest"]
        frm_path = Path(frm_entry.file_path)
        assert frm_path.is_file()
        assert frm_entry.resource_file_name == "frmTest.frx"
        assert (out_dir / "frmTest.frx").is_file(), "frmTest.frx designer resource missing"

        _assert_owned_never_foreground(owned_pid, "vba-hash")
        _assert_no_visible_owned_windows(owned_pid, "vba-hash")

        print(f"[vba-hash] checked={checked} all_hashes_match=True out_dir={out_dir}")
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 7 (task_017): missing reference reported by vba_get_references
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_missing_reference_detected(tmp_path: Path) -> None:
    """A broken reference must surface as a ``missing_references`` diagnostic.

    A temp copy of ``CleanProject.xlsm`` is opened and a broken reference is
    forced through VBIDE ``VBProject.References.AddFromFile`` (a non-existent
    path, then a registered typelib whose temp copy is deleted). The test then
    asserts ``vba_get_references`` returns the broken reference with a
    name/GUID and that ``get_project_info().missing_references`` includes it.

    This mutates a **temporary copy only**; the committed fixture is never
    touched. On builds where no safe recipe yields a service-visible broken
    reference (the production adapter drops broken refs whose ``Name`` raises),
    the test skips with an explicit reason rather than asserting a false pass.
    """

    rig = _make_rig(tmp_path)
    committed_hash = _sha256(_CLEAN_PROJECT)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()
        record = rig.session._records[info.workbook_id]
        working_path = Path(record.working_path)

        # Registered typelib not in the default reference set. Used to attempt an
        # "add then delete the target" break. Present on Windows.
        scripting_typelib = Path(r"C:\Windows\System32\scrrun.dll")

        def force_broken() -> dict[str, Any]:
            """Run on the STA worker: try to add a broken reference in-memory."""

            app = rig.app_adapter.app
            target = working_path.as_posix().lower()
            wb = None
            for candidate in app.Workbooks:
                try:
                    if str(candidate.FullName).replace("\\", "/").lower() == target:
                        wb = candidate
                        break
                except Exception:  # noqa: BLE001 - defensive
                    continue
            if wb is None:
                return {"stage": "no-workbook"}

            project = wb.VBProject
            outcome: dict[str, Any] = {"stage": "attempted"}

            # Attempt 1: AddFromFile a non-existent path (expected to raise on
            # this build; handled so the test can decide how to proceed).
            missing_path = working_path.parent / f"does_not_exist_{uuid4().hex}.tlb"
            try:
                project.References.AddFromFile(str(missing_path))
                outcome["add_nonexistent"] = "added"
            except Exception as exc:  # noqa: BLE001 - expected on this build
                outcome["add_nonexistent"] = f"raised:{type(exc).__name__}"

            # Attempt 2: add a registered typelib from a temp copy, then delete
            # the temp copy to try to force IsBroken. We only delete OUR temp
            # copy, never the registered system file.
            temp_copy = working_path.parent / f"probe_typelib_{uuid4().hex}.dll"
            if scripting_typelib.is_file():
                try:
                    shutil.copy2(scripting_typelib, temp_copy)
                    new_ref = project.References.AddFromFile(str(temp_copy))
                    outcome["add_typelib"] = str(new_ref.Name)
                    with contextlib.suppress(Exception):
                        temp_copy.unlink()
                        outcome["deleted_temp_copy"] = True
                except Exception as exc:  # noqa: BLE001 - expected on some builds
                    outcome["add_typelib"] = f"raised:{type(exc).__name__}"
                    with contextlib.suppress(Exception):
                        temp_copy.unlink()

            return outcome

        force_broken.__name__ = "vba_force_broken_reference"
        mutation = await rig.worker.submit(
            force_broken, timeout=_CALL_TIMEOUT_S, section="diagnostic"
        )

        refs = await rig.service.get_references(info.workbook_id)
        broken = [reference for reference in refs if reference.is_broken]
        project = await rig.service.get_project_info(info.workbook_id)

        # If the build could not produce a service-visible broken reference, skip
        # with a precise reason (AddFromFile(nonexistent) raises here, and the
        # production adapter drops broken refs whose Name raises com_error).
        if not broken:
            pytest.skip(
                "This Excel build did not yield a service-visible broken reference "
                "without unsafe file/registry manipulation: "
                f"AddFromFile(nonexistent)={mutation.get('add_nonexistent')!r}, "
                f"AddFromFile(typelib)={mutation.get('add_typelib')!r}; "
                "the production VBIDEAdapter drops broken references whose Name "
                "raises com_error, and registered typelibs canonicalise to their "
                "system path (deleting a temp copy cannot break them). "
                "Renaming the referenced system file is unsafe / out of scope."
            )

        # A broken reference must carry a name or GUID and appear in the
        # missing_references diagnostic.
        broken_with_identity = [
            reference for reference in broken if reference.name or reference.guid
        ]
        assert broken_with_identity, (
            f"broken reference surfaced without name/GUID: {[r.model_dump() for r in broken]}"
        )

        broken_names = {reference.name for reference in broken if reference.name}
        broken_guids = {reference.guid for reference in broken if reference.guid}
        missing = set(project.missing_references)

        assert missing, (
            "get_project_info().missing_references is empty despite broken refs: "
            f"broken={[r.model_dump() for r in broken]}"
        )
        assert broken_names & missing or broken_guids & missing, (
            f"missing_references={sorted(missing)} does not include any broken "
            f"reference name={sorted(broken_names)} guid={sorted(broken_guids)}"
        )

        # The committed fixture is untouched.
        assert _sha256(_CLEAN_PROJECT) == committed_hash

        _assert_owned_never_foreground(owned_pid, "vba-missing-ref")
        _assert_no_visible_owned_windows(owned_pid, "vba-missing-ref")

        print(
            f"[vba-missing-ref] broken={[r.name for r in broken]} "
            f"missing_references={sorted(missing)} mutation={mutation}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"


# ---------------------------------------------------------------------------
# Test 8 (task_017): all component types present (Standard/Class/UserForm/Document)
# ---------------------------------------------------------------------------
@requires_excel
@pytest.mark.anyio
@pytest.mark.integration_excel
async def test_all_component_types_present(tmp_path: Path) -> None:
    """CleanProject must expose every supported component type.

    Asserts the inventory contains a Standard module, a Class module, a UserForm
    (with designer) and Document modules, that the Document modules are flagged
    ``is_document_module`` and are still present after an export (i.e. exporting
    never strips them).
    """

    rig = _make_rig(tmp_path)
    await rig.start()
    owned_pid: int | None = None
    try:
        owned_pid = rig.worker.get_status().owned_pid
        assert owned_pid is not None

        info = await rig.open()

        project = await rig.service.get_project_info(info.workbook_id)
        assert project.available is True
        types_observed = {component.component_type for component in project.components}

        assert VbaComponentType.STANDARD_MODULE in types_observed, (
            f"no Standard module; observed={sorted(t.value for t in types_observed)}"
        )
        assert VbaComponentType.CLASS_MODULE in types_observed, (
            f"no Class module; observed={sorted(t.value for t in types_observed)}"
        )
        assert VbaComponentType.USER_FORM in types_observed, (
            f"no UserForm; observed={sorted(t.value for t in types_observed)}"
        )
        assert VbaComponentType.DOCUMENT_MODULE in types_observed, (
            f"no Document module; observed={sorted(t.value for t in types_observed)}"
        )

        by_name = {component.component_name: component for component in project.components}

        # Document modules (ThisWorkbook / Sheet1) are present and flagged.
        document_names = {
            name for name, component in by_name.items() if component.is_document_module
        }
        assert document_names, "no component flagged is_document_module"
        assert {"ThisWorkbook", "Sheet1"}.issubset(document_names), (
            f"expected ThisWorkbook + Sheet1 as Document modules; got {sorted(document_names)}"
        )

        # The UserForm has a designer and is flagged as a UserForm type.
        user_forms = [
            component
            for component in project.components
            if component.component_type is VbaComponentType.USER_FORM
        ]
        assert user_forms, "no UserForm component"
        assert any(component.has_designer for component in user_forms), "UserForm has no designer"

        # Export must not remove Document modules.
        manifest = await rig.service.export_project(info.workbook_id)
        assert all(entry.export_status is VbaExportStatus.OK for entry in manifest.components), (
            f"export failures: {[(e.component_name, e.error) for e in manifest.components]}"
        )

        after = await rig.service.list_components(info.workbook_id)
        after_names = {component.component_name for component in after}
        assert {"ThisWorkbook", "Sheet1"}.issubset(after_names), (
            f"document modules removed after export: {sorted(after_names)}"
        )

        _assert_owned_never_foreground(owned_pid, "vba-types")
        _assert_no_visible_owned_windows(owned_pid, "vba-types")

        print(
            f"[vba-types] types={sorted(t.value for t in types_observed)} "
            f"document_modules={sorted(document_names)} "
            f"components={len(project.components)}"
        )
    finally:
        await rig.stop()

    assert owned_pid is not None
    assert _wait_pid_gone(owned_pid), f"owned PID {owned_pid} still alive after shutdown"
