"""Read-only VBA project service (task_016, spec §6.2 / §7.2).

:class:`VBAProjectService` is the application-layer orchestrator for the
read-only VBA project tools: ``vba_project_info``, ``vba_list_components``,
``vba_list_procedures``, ``vba_read_component``, ``vba_export_project``,
``vba_get_references`` and ``vba_search_code``.

It owns **no COM proxies**: every Excel/VBIDE access is delegated to the
single-STA :class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` via
``await worker.submit(...)`` and executed by the STA-local
:class:`~excel_vba_mcp.adapters.excel.vbide_adapter.VBIDEAdapter`. The service
resolves the open workbook through the :class:`SessionService` bookkeeping, then
submits a synchronous callable that re-locates the working copy inside the owned
instance and returns serializable DTOs. Nothing here touches COM (ARCH-009).

Safety contract (spec §7.2, §10; addendum BG03/BG04/BG05):

* **Read-only.** No VBA mutation in this task (task_018 owns mutations).
* **Copies only.** The working copy opened by the session service is read; the
  original workbook is never modified.
* **Hidden / non-activating.** All work runs on the owned, ``Visible=False``
  Excel instance via the adapter; the service never activates/selects anything.
* **Bounded export.** ``export_project`` writes under a temp directory by
  default; a caller-supplied ``out_dir`` must live inside a configured
  artifact root (``PathPolicy``), so a tool parameter can never redirect exports
  to an arbitrary location.
"""

from __future__ import annotations

import contextlib
import os
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any, TypeVar

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError, PathNotAllowedError
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaExportManifest,
    VbaProcedureInfo,
    VbaProjectInfo,
    VbaReferenceInfo,
    VbaSearchMatch,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds).
_DEFAULT_TIMEOUT_S = 60.0

# Logical artifact root under which a caller-supplied export directory must live
# (a subset of ``PathPolicy.ARTIFACT_ROOT_NAMES``: snapshots / artifacts).
_EXPORT_ROOT_NAMES = ("snapshots", "artifacts")

# Type of the synchronous callable handed to the STA worker.
_VbaCall = Callable[[Any], Any]

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class VBAProjectService:
    """Read-only VBA inventory / procedures / references / export.

    Args:
        worker: The single-STA COM worker (all Excel runs there).
        session_service: The workbook session lifecycle service (owns the worker
            and the ``workbook_id`` bookkeeping).
        path_policy: Allowlist / canonicalisation gate for caller-supplied export
            directories.
        audit: Append-only JSONL audit trail.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        path_policy: PathPolicy,
        audit: AuditLogger,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._path_policy = path_policy
        self._audit = audit
        self._adapter = VBIDEAdapter()

    # ------------------------------------------------------------------
    # Public API (async; no COM)
    # ------------------------------------------------------------------
    async def get_project_info(self, workbook_id: str) -> VbaProjectInfo:
        """Return the VBA project summary for ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_project_info",
            lambda wb: self._adapter.project_info(wb),
        )
        self._record_audit(
            "vba_project_info",
            workbook_id,
            {"component_count": result.component_count, "locked": result.locked},
        )
        return result

    async def list_components(self, workbook_id: str) -> list[VbaComponentInfo]:
        """Return the component inventory for ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_list_components",
            lambda wb: self._adapter.list_components(wb),
        )
        self._record_audit("vba_list_components", workbook_id, {"count": len(result)})
        return result

    async def list_procedures(self, workbook_id: str, component: str) -> list[VbaProcedureInfo]:
        """Return the procedures of ``component`` in ``workbook_id``."""

        result = await self._submit(
            workbook_id,
            "vba_list_procedures",
            lambda wb: self._adapter.list_procedures(wb, component),
        )
        self._record_audit(
            "vba_list_procedures",
            workbook_id,
            {"component": component, "count": len(result)},
        )
        return result

    async def read_component(self, workbook_id: str, component: str) -> VbaComponentSource:
        """Return the code text + normalized hash of ``component``."""

        result = await self._submit(
            workbook_id,
            "vba_read_component",
            lambda wb: self._adapter.read_component(wb, component),
        )
        self._record_audit(
            "vba_read_component",
            workbook_id,
            {"component": component, "source_hash": result.source_hash},
        )
        return result

    async def export_project(
        self, workbook_id: str, out_dir: str | None = None
    ) -> VbaExportManifest:
        """Export every VBA component of ``workbook_id`` and return the manifest.

        When ``out_dir`` is ``None`` a fresh temporary directory is used. A
        caller-supplied ``out_dir`` is validated by :class:`PathPolicy` against
        the configured artifact roots, so an export can never escape the
        allowlist.
        """

        if out_dir is None:
            target = Path(tempfile.mkdtemp(prefix="excel_vba_mcp_vba_export_"))
        else:
            target = self._validate_export_dir(out_dir)
            target.mkdir(parents=True, exist_ok=True)

        result = await self._submit(
            workbook_id,
            "vba_export_project",
            lambda wb: self._adapter.export_project(wb, target, workbook_id=workbook_id),
        )
        self._record_audit(
            "vba_export_project",
            workbook_id,
            {"out_dir": result.out_dir, "component_count": len(result.components)},
        )
        return result

    async def get_references(self, workbook_id: str) -> list[VbaReferenceInfo]:
        """Return the VBA project references (missing ones flagged)."""

        result = await self._submit(
            workbook_id,
            "vba_get_references",
            lambda wb: self._adapter.get_references(wb),
        )
        self._record_audit(
            "vba_get_references",
            workbook_id,
            {
                "count": len(result),
                "missing": [reference.name for reference in result if reference.is_broken],
            },
        )
        return result

    async def search_code(
        self, workbook_id: str, pattern: str, regex: bool = False
    ) -> list[VbaSearchMatch]:
        """Search every component's code in ``workbook_id`` (bounded)."""

        result = await self._submit(
            workbook_id,
            "vba_search_code",
            lambda wb: self._adapter.search_code(wb, pattern, regex=regex),
        )
        self._record_audit(
            "vba_search_code",
            workbook_id,
            {"pattern": pattern, "regex": regex, "count": len(result)},
        )
        return result

    # ------------------------------------------------------------------
    # Internal helpers (non-COM)
    # ------------------------------------------------------------------
    def _resolve_record(self, workbook_id: str) -> WorkbookInfo:
        """Resolve ``workbook_id`` to its open :class:`WorkbookInfo` (no COM)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise ExcelVbaMcpError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            context={"workbook_id": workbook_id},
        )

    async def _submit(
        self,
        workbook_id: str,
        tool: str,
        fn: Callable[[Any], _T],
    ) -> _T:
        """Run ``fn(wb)`` on the STA worker for the open working copy."""

        record = self._resolve_record(workbook_id)
        app_adapter = self._app_adapter()
        call = _bind_workbook(fn, app_adapter, record.path, workbook_id)
        call.__name__ = tool
        return await self._worker.submit(
            call,
            timeout=_DEFAULT_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )

    def _app_adapter(self) -> ExcelAppAdapter:
        """Return the owned hidden Excel application adapter (never a new one)."""

        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "vba_project"},
            )
        return adapter

    def _validate_export_dir(self, out_dir: str) -> Path:
        """Validate a caller-supplied export directory against the allowlist.

        The directory itself is validated inside a configured artifact root
        (``snapshots`` / ``artifacts``). ``PathPolicy.validate_path`` also
        enforces a file-extension allowlist, which is meant for files, not
        directories; the export files themselves (``.bas`` / ``.cls`` / ``.frm``
        / ``.frx``) are written by the adapter, not by the caller, so the
        directory check is done on the canonical directory path directly.
        """

        if not isinstance(out_dir, str) or not out_dir.strip():
            raise ExcelVbaMcpError(
                "out_dir must be a non-empty string.",
                context={"field": "out_dir"},
            )
        expected_root = _artifact_root_for(self._path_policy)
        candidate = Path(out_dir)
        if candidate.suffix:
            # A caller passed a file path; validate it as a file (extension gate
            # applies) so the rejection reason is precise.
            return self._path_policy.validate_path(
                out_dir, OperationClass.READ_ONLY, expected_root=expected_root
            )
        return _validate_dir(out_dir, self._path_policy, expected_root)

    def _record_audit(self, tool: str, workbook_id: str, result: dict[str, Any]) -> None:
        canonical = ""
        with contextlib.suppress(ExcelVbaMcpError):
            canonical = self._resolve_record(workbook_id).source_path or ""
        self._audit.record(
            tool=tool,
            operation_class="read_only",
            params={"workbook_id": workbook_id},
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": canonical},
            result=result,
        )


# ---------------------------------------------------------------------------
# Module-level helpers (no COM at import time)
# ---------------------------------------------------------------------------
def _bind_workbook(
    fn: _VbaCall,
    app_adapter: ExcelAppAdapter,
    working_path: str,
    workbook_id: str,
) -> Callable[[], Any]:
    """Wrap ``fn`` so it re-locates the working copy on the STA thread."""

    def run() -> Any:
        wb = _find_open_workbook(app_adapter, working_path)
        if wb is None:
            raise ExcelVbaMcpError(
                "The working copy is no longer open in the owned instance.",
                context={"workbook_id": workbook_id, "working_path": working_path},
            )
        return fn(wb)

    return run


def _find_open_workbook(app_adapter: ExcelAppAdapter, working_path: str) -> Any | None:
    """Return the open ``Workbook`` proxy matching ``working_path`` (STA thread)."""

    app = app_adapter.app
    if app is None:
        return None
    target = working_path.lower()
    try:
        for wb in app.Workbooks:
            try:
                if str(wb.FullName).lower() == target:
                    return wb
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return None


def _artifact_root_for(path_policy: PathPolicy) -> str | None:
    """Return the first configured artifact root name, if any."""

    roots = path_policy.roots
    for name in _EXPORT_ROOT_NAMES:
        if name in roots:
            return name
    return None


def _validate_dir(raw: str, path_policy: PathPolicy, expected_root: str | None) -> Path:
    """Validate a directory path against the allowlist (extension gate bypassed).

    ``PathPolicy.validate_path`` rejects a candidate by file extension, which is
    meaningful for files but not for directories. The directory is therefore
    canonicalized with the same ``resolve`` / ``realpath`` cross-check the policy
    uses and the containment / expected-root rules are re-applied here so an
    export can never escape the allowlist.
    """

    candidate = Path(raw)
    try:
        resolved = candidate.resolve()
    except OSError:  # pragma: no cover - resolve rarely fails
        resolved = Path(os.path.abspath(raw))
    real = Path(os.path.realpath(candidate))
    canonical = real if os.fspath(real) != os.fspath(resolved) else resolved

    roots = path_policy.roots
    matches = [name for name, root in roots.items() if _is_within(canonical, root)]
    if not matches:
        raise PathNotAllowedError(
            "Path not allowed (outside_roots).",
            context={"reason": "outside_roots", "operation": "read_only", "name": candidate.name},
        )
    if expected_root is not None and expected_root not in matches:
        raise PathNotAllowedError(
            "Path not allowed (root_mismatch).",
            context={
                "reason": "root_mismatch",
                "operation": "read_only",
                "name": candidate.name,
                "expected_root": expected_root,
            },
        )
    return canonical


def _is_within(child: str | Path, parent: str | Path) -> bool:
    """Return whether ``child`` is inside ``parent`` (no symlink resolution)."""

    try:
        common = os.path.commonpath([os.fspath(child), os.fspath(parent)])
    except ValueError:
        return False
    return os.path.normcase(common) == os.path.normcase(os.fspath(parent))


__all__ = ["VBAProjectService"]
