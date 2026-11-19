"""Production VBE compile service with fail-closed diagnostics (task_020).

On the validated build (Windows 10 LTSC 2019 17763 x64, Excel 16.0.14332 x64,
UI 1033, VBOM enabled) **non-activating compilation is infeasible**:

* ``VBE.CommandBars["Debug"]`` has no Compile control;
* ``FindControl(578)`` raises ``E_INVALIDARG``;
* the only route that reaches Compile (UIA ``Debug`` menu expansion) creates four
  visible ``MsoCommandBarPopup`` / shadow windows and makes the hidden VBE the
  foreground window, and a compile error shows a visible, non-dismissible
  ``#32770`` modal.

See ``.ai/audit/hidden-compile-research-2026-11-19.md`` and
``.ai/plans/plan_addendum_background_execution.md`` (BG03/BG04/BG07/BG09).

Therefore :meth:`CompileService.compile_project` **never** invokes a compile and
**never** returns ``compiled`` / ``compile_error``. It returns a truthful
:class:`CompileResult` with ``status=COMPILE_INDETERMINATE`` and a structured
``reason`` (``NON_ACTIVATING_COMPILE_UNAVAILABLE`` on this build), plus partial,
explicitly non-compile diagnostics:

* ``References.IsBroken`` (real, non-intrusive),
* ``VBProject.Protection`` (locked projects are reported, never bypassed),
* the component inventory,
* VBOM availability (``0x800A17B4`` -> ``ACCESS_VBOM_DISABLED``).

All Excel/VBIDE access is delegated to the single-STA worker via
``await worker.submit(...)``; the service owns no COM proxies (ARCH-009). The
submitted callable is tagged ``section="diagnostic"`` and
``op_class=EXECUTING`` because it touches the VBE command layer read-only; it is
still non-activating and never shows or focuses a window.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import UTC, datetime
from typing import Any, TypeVar

from pydantic import Field

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter
from excel_vba_mcp.adapters.windows_ui.vbe_compile_ui import (
    CommandDiscovery,
    VBECompileUI,
)
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ExcelVbaMcpError,
    VbaProjectLockedError,
)
from excel_vba_mcp.domain.models import (
    DomainModel,
    TimestampedModel,
    VbaComponentInfo,
    VbaReferenceInfo,
    WorkbookInfo,
)
from excel_vba_mcp.domain.statuses import CompileStatus, OperationClass
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Default per-call soft ceiling (seconds).
_DEFAULT_TIMEOUT_S = 60.0

# Structured, fail-closed reasons attached to an indeterminate result.
REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE = "NON_ACTIVATING_COMPILE_UNAVAILABLE"
REASON_ACCESS_VBOM_DISABLED = "ACCESS_VBOM_DISABLED"
REASON_VBA_PROJECT_LOCKED = "VBA_PROJECT_LOCKED"

# Type of the synchronous callable handed to the STA worker.
_ProbeCall = Callable[[], dict[str, Any]]

# Return type of the worker callables submitted by this service.
_T = TypeVar("_T")


class CompileDiagnostics(DomainModel):
    """Partial, explicitly non-compile diagnostics (research §Q5).

    ``is_compile`` is always ``False``: none of these fields proves that the VBA
    project compiled. They are non-intrusive observations that can explain a
    compile-blocking condition (locked project, missing reference, VBOM disabled).
    """

    is_compile: bool = False
    vbom_available: bool = True
    vbom_error: str | None = None
    protection: int | None = None
    locked: bool = False
    references: list[VbaReferenceInfo] = Field(default_factory=list)
    broken_references: list[str] = Field(default_factory=list)
    components: list[VbaComponentInfo] = Field(default_factory=list)
    component_count: int = 0
    notes: list[str] = Field(default_factory=list)


class CompileResult(TimestampedModel):
    """Fail-closed compile result.

    ``status`` is always :attr:`CompileStatus.INDETERMINATE` on this build. The
    ``reason`` explains why (``NON_ACTIVATING_COMPILE_UNAVAILABLE`` unless VBOM is
    disabled or the project is locked). ``diagnostics`` is explicitly
    ``is_compile=False`` and ``command_discovery`` records the read-only evidence
    that no Compile command route is available.
    """

    workbook_id: str
    status: CompileStatus
    reason: str
    diagnostics: CompileDiagnostics
    command_discovery: CommandDiscovery
    notes: list[str] = Field(default_factory=list)


class CompileService:
    """Non-activating compile service that fails closed to ``INDETERMINATE``.

    Args:
        worker: The single-STA COM worker (all Excel/VBE access runs there).
        session_service: The workbook session lifecycle service (owns the
            ``workbook_id`` bookkeeping).
        vba_project_service: The read-only VBA project service. Kept as a
            dependency for callers that already wire the VBA surface; the compile
            probe itself reads the project directly on the STA worker so the whole
            observation happens in one serialized call.
        path_policy: The allowlist / canonicalisation policy (kept for symmetry
            with the other application services).
        audit: Append-only JSONL audit trail.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        vba_project_service: VBAProjectService,
        path_policy: PathPolicy,
        audit: AuditLogger,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._vba = vba_project_service
        self._path_policy = path_policy
        self._audit = audit
        self._last: dict[str, CompileResult] = {}

    # ------------------------------------------------------------------
    # Public API (async; no COM)
    # ------------------------------------------------------------------
    async def compile_project(
        self, workbook_id: str, *, timeout_s: float = _DEFAULT_TIMEOUT_S
    ) -> CompileResult:
        """Return a fail-closed compile result for ``workbook_id``.

        The method never compiles and never returns ``compiled`` /
        ``compile_error``. It submits one read-only STA probe that collects
        non-intrusive diagnostics and command-discovery evidence, then returns
        ``status=COMPILE_INDETERMINATE`` with a structured ``reason``.

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            timeout_s: Soft per-call ceiling for the STA probe.

        Returns:
            A :class:`CompileResult` with ``status=COMPILE_INDETERMINATE``,
            ``reason`` and partial ``diagnostics`` (``is_compile=False``).

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown or the working copy
                is no longer open in the owned instance.
        """

        record = self._resolve_record(workbook_id)
        app_adapter = self._app_adapter()

        def compile_fn() -> dict[str, Any]:
            wb = _find_open_workbook(app_adapter, record.path)
            if wb is None:
                raise ExcelVbaMcpError(
                    "The working copy is no longer open in the owned instance.",
                    context={"workbook_id": workbook_id, "working_path": record.path},
                )
            return _collect_probe(wb, app_adapter)

        compile_fn.__name__ = "vba_compile"
        payload = await self._worker.submit(
            compile_fn,
            timeout=timeout_s,
            section="diagnostic",
            op_class=OperationClass.EXECUTING,
        )

        result = self._build_result(workbook_id, payload)
        self._last[workbook_id] = result
        self._record_audit(workbook_id, record, result)
        return result

    async def get_last_diagnostics(self, workbook_id: str) -> CompileDiagnostics:
        """Return the cached diagnostics for ``workbook_id`` without a forced re-run.

        When a previous :meth:`compile_project` result is cached it is returned
        as-is (no re-run). When no result is cached yet, one non-activating
        diagnostics probe is collected (still no compile) and cached.

        Args:
            workbook_id: The id returned by :func:`excel_open`.

        Returns:
            The cached (or freshly collected) :class:`CompileDiagnostics`
            (``is_compile=False``).
        """

        cached = self._last.get(workbook_id)
        if cached is not None:
            return cached.diagnostics
        result = await self.compile_project(workbook_id)
        return result.diagnostics

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

    def _app_adapter(self) -> ExcelAppAdapter:
        """Return the owned hidden Excel application adapter (never a new one)."""

        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "vba_compile"},
            )
        return adapter

    @staticmethod
    def _build_result(workbook_id: str, payload: dict[str, Any]) -> CompileResult:
        """Construct the fail-closed :class:`CompileResult` from the probe payload."""

        vbom_available = bool(payload.get("vbom_available", True))
        locked = bool(payload.get("locked", False))

        if not vbom_available:
            reason = REASON_ACCESS_VBOM_DISABLED
        elif locked:
            reason = REASON_VBA_PROJECT_LOCKED
        else:
            reason = REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE

        raw_references = payload.get("references", [])
        references = [VbaReferenceInfo(**item) for item in raw_references]
        raw_components = payload.get("components", [])
        components = [VbaComponentInfo(**item) for item in raw_components]
        broken_references = list(payload.get("broken_references", []))

        notes: list[str] = [
            "NOT a compile: non-activating VBA compilation is unavailable on this build.",
            f"status=indeterminate reason={reason}",
        ]
        notes.extend(str(note) for note in payload.get("notes", []))

        diagnostics = CompileDiagnostics(
            is_compile=False,
            vbom_available=vbom_available,
            vbom_error=payload.get("vbom_error"),
            protection=payload.get("protection"),
            locked=locked,
            references=references,
            broken_references=broken_references,
            components=components,
            component_count=len(components),
            notes=notes,
        )
        discovery = CommandDiscovery(**(payload.get("command_discovery") or {}))
        now = datetime.now(UTC)
        return CompileResult(
            workbook_id=workbook_id,
            status=CompileStatus.INDETERMINATE,
            reason=reason,
            diagnostics=diagnostics,
            command_discovery=discovery,
            notes=[
                "Compile was not invoked: no non-activating route exists on this build.",
                "No COM compile, no UIA menu expansion, no dialog interaction was performed.",
            ],
            created_at=now,
            updated_at=now,
        )

    def _record_audit(self, workbook_id: str, record: WorkbookInfo, result: CompileResult) -> None:
        """Record a truthful audit entry (never claims a compile ran)."""

        self._audit.record(
            tool="vba_compile",
            operation_class="executing",
            params={
                "workbook_id": workbook_id,
                "status": result.status.value,
                "reason": result.reason,
                "is_compile": False,
            },
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": record.source_path or record.path},
            result={
                "status": result.status.value,
                "reason": result.reason,
                "is_compile": False,
                "component_count": result.diagnostics.component_count,
                "broken_references": result.diagnostics.broken_references,
                "vbom_available": result.diagnostics.vbom_available,
            },
        )


# ---------------------------------------------------------------------------
# STA-local probe (runs on the worker thread; touches COM only)
# ---------------------------------------------------------------------------
def _collect_probe(wb: Any, app_adapter: ExcelAppAdapter) -> dict[str, Any]:
    """Collect partial, non-intrusive compile diagnostics (STA thread).

    This function never invokes Compile, never expands a UIA menu and never
    changes any application setting. It returns a JSON-serializable payload dict
    so it can cross the single-STA worker boundary.

    The payload keys are:

    ``vbom_available`` / ``vbom_error`` / ``protection`` / ``locked`` /
    ``references`` / ``broken_references`` / ``components`` / ``notes`` /
    ``command_discovery``.
    """

    adapter = VBIDEAdapter()
    payload: dict[str, Any] = {
        "vbom_available": True,
        "vbom_error": None,
        "protection": None,
        "locked": False,
        "references": [],
        "broken_references": [],
        "components": [],
        "notes": [],
        "command_discovery": None,
    }

    try:
        project = adapter._project(wb)  # noqa: SLF001 - STA-local access
    except AccessVbomDisabledError as exc:
        payload["vbom_available"] = False
        payload["vbom_error"] = exc.code.value
        payload["notes"].append(
            "VBOM access disabled (0x800A17B4): project diagnostics are unavailable."
        )
        payload["command_discovery"] = _safe_discover(app_adapter)
        return payload
    except Exception as exc:  # noqa: BLE001 - any other failure is diagnostic data
        payload["vbom_available"] = False
        payload["vbom_error"] = exc.__class__.__name__
        payload["notes"].append(f"VBProject access failed: {exc.__class__.__name__}: {exc}")
        payload["command_discovery"] = _safe_discover(app_adapter)
        return payload

    payload["protection"] = _safe_int(getattr(project, "Protection", 0))
    payload["locked"] = payload["protection"] != 0

    try:
        references = adapter.get_references(wb)
    except Exception as exc:  # noqa: BLE001 - defensive COM read
        payload["notes"].append(f"Reference enumeration failed: {exc.__class__.__name__}: {exc}")
    else:
        payload["references"] = [reference.model_dump(mode="json") for reference in references]
        payload["broken_references"] = [
            reference.name or reference.guid or "<unknown>"
            for reference in references
            if reference.is_broken
        ]

    try:
        components = adapter.list_components(wb)
    except VbaProjectLockedError:
        payload["notes"].append(
            "Component inventory unavailable: the VBA project is locked for viewing."
        )
    except AccessVbomDisabledError:
        payload["notes"].append(
            "Component inventory unavailable: VBOM access is disabled (0x800A17B4)."
        )
    except Exception as exc:  # noqa: BLE001 - defensive COM read
        payload["notes"].append(f"Component inventory failed: {exc.__class__.__name__}: {exc}")
    else:
        payload["components"] = [component.model_dump(mode="json") for component in components]

    payload["command_discovery"] = _safe_discover(app_adapter)
    return payload


def _safe_discover(app_adapter: ExcelAppAdapter) -> dict[str, Any]:
    """Run read-only command discovery, converting failures into evidence."""

    app = app_adapter.app
    if app is None:
        return CommandDiscovery(
            attempted=False,
            notes=["Excel application unavailable; command discovery skipped."],
        ).model_dump(mode="json")
    try:
        vbe = app.VBE
    except Exception as exc:  # noqa: BLE001 - VBOM disabled / VBE unavailable
        return CommandDiscovery(
            attempted=False,
            notes=[f"VBE unavailable: {exc.__class__.__name__}: {exc}"],
        ).model_dump(mode="json")
    try:
        discovery = VBECompileUI().discover_command(vbe)
    except Exception as exc:  # noqa: BLE001 - discovery must never crash the probe
        return CommandDiscovery(
            attempted=True,
            notes=[f"Command discovery failed: {exc.__class__.__name__}: {exc}"],
        ).model_dump(mode="json")
    return discovery.model_dump(mode="json")


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


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


__all__ = [
    "REASON_ACCESS_VBOM_DISABLED",
    "REASON_NON_ACTIVATING_COMPILE_UNAVAILABLE",
    "REASON_VBA_PROJECT_LOCKED",
    "CompileDiagnostics",
    "CompileResult",
    "CompileService",
]
