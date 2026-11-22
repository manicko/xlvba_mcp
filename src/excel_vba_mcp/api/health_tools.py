"""Excel health MCP tool (task_034; spec §5.6, §6.1, §11).

``excel_health`` reports the worker / session state **without forcing a re-run**:
it only reads the non-COM worker snapshot (``ExcelComWorker.get_status``) and the
session bookkeeping. It never submits work to the STA worker, so it stays
truthful while a call is hung or the session needs recovery — a health check that
forced a COM call would itself block on the hang it is meant to report.

It additionally performs a **read-only** liveness probe of the worker's owned
Excel pid (FE-18 / R-17): if the hidden ``EXCEL.EXE`` the worker owns has died
(killed externally or crashed) while the worker bookkeeping still says
``running``/``accepting``, the stale worker would otherwise look healthy until a
server restart. A dead owned pid therefore forces ``needs_recovery`` so the
verdict is truthful. The probe is a single ``OpenProcess`` + ``GetExitCodeProcess``
call (no COM, no kill, no recovery) and degrades to a no-op when the Win32 layer
is unavailable (non-Windows / no ctypes).

Like the other tool modules, this handler is thin and never touches COM
(ARCH-009).
"""

from __future__ import annotations

import sys

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.api.workbook_tools import get_session_service
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import HealthStatus, WorkbookSessionState

__all__ = ["ExcelHealth", "excel_health"]


class ExcelHealth(BaseModel):
    """Structured ``excel_health`` payload (no COM, no forced re-run)."""

    model_config = ConfigDict(extra="forbid")

    status: HealthStatus
    session_id: str | None = None
    session_state: WorkbookSessionState
    worker_state: WorkbookSessionState
    accepting: bool
    needs_recovery: bool
    owned_pid: int | None = None
    active_operation: str | None = None
    active_section: str | None = None
    last_hresult: int | None = None
    last_error: str | None = None
    open_workbooks: int = 0
    notes: list[str] = Field(default_factory=list)


def _overall_status(*, worker_state: WorkbookSessionState, needs_recovery: bool) -> HealthStatus:
    """Map the worker/session state to a single truthful health verdict."""

    if worker_state in (
        WorkbookSessionState.CLOSED,
        WorkbookSessionState.ERROR,
        WorkbookSessionState.STARTING,
    ):
        return HealthStatus.UNAVAILABLE
    if needs_recovery:
        return HealthStatus.NEEDS_RECOVERY
    return HealthStatus.OK


def _liveness_available() -> bool:
    """Whether the read-only Win32 pid-liveness probe is usable on this host.

    ``excel_health`` is a diagnostic and must degrade gracefully where the probe
    cannot run (non-Windows, or no ``ctypes``/``windll``). Returning ``False``
    here makes the caller skip the probe entirely and keep the pure state-based
    behaviour, so a missing Win32 layer never produces a false "dead pid" signal.
    """

    if sys.platform != "win32":
        return False
    try:
        import ctypes

        # Attribute access forces the Win32 binding to load; failure => unavailable.
        _ = ctypes.windll.kernel32
    except Exception:  # noqa: BLE001 - platform / binding availability probe
        return False
    return True


def _owned_pid_alive(pid: int) -> bool:
    """Read-only liveness probe for the worker's owned Excel pid (FE-18 / R-17).

    Delegates to :meth:`ExcelAppAdapter._pid_alive` (the single Win32
    ``OpenProcess`` + ``GetExitCodeProcess`` helper introduced by FE-17) so the
    liveness logic is never duplicated. Never touches COM, never kills, never
    recovers — ``excel_health`` stays a pure diagnostic.
    """

    return ExcelAppAdapter._pid_alive(pid)


async def excel_health() -> ExcelHealth:
    """Return worker/session health without touching COM or re-running work.

    Returns:
        An :class:`ExcelHealth` describing the worker state, whether the session
        is accepting new work, the owned Excel pid (if any), the active operation
        and whether recovery is required. When the session service is not
        configured (e.g. startup failed) a truthful ``unavailable`` payload is
        returned instead of raising.
    """

    try:
        session = get_session_service()
    except ExcelVbaMcpError:
        return ExcelHealth(
            status=HealthStatus.UNAVAILABLE,
            session_state=WorkbookSessionState.CLOSED,
            worker_state=WorkbookSessionState.CLOSED,
            accepting=False,
            needs_recovery=False,
            notes=["The Excel session service is not initialized."],
        )

    worker_status = session.worker_status()
    session_state = worker_status.state
    needs_recovery = (
        worker_status.state == WorkbookSessionState.NEEDS_RECOVERY
        or not worker_status.accepting
        or not session.accepting
    )

    # Read-only owned-pid liveness probe (FE-18 / R-17): a stale worker can report
    # running/accepting while its hidden Excel process is already dead, which would
    # otherwise look healthy until a server restart. Probe only when the worker
    # bookkeeping claims a live owned pid and the Win32 layer is available; the
    # probe is additive and never overrides a truthful state-based verdict.
    notes = list(worker_status.notes)
    owned_pid = worker_status.owned_pid
    if (
        owned_pid is not None
        and not needs_recovery
        and _liveness_available()
        and not _owned_pid_alive(owned_pid)
    ):
        needs_recovery = True
        notes.append(
            f"owned pid {owned_pid} is not alive; the hidden Excel process died "
            "(stale worker requires recovery)"
        )

    return ExcelHealth(
        status=_overall_status(worker_state=worker_status.state, needs_recovery=needs_recovery),
        session_id=session.session_id,
        session_state=session_state,
        worker_state=worker_status.state,
        accepting=bool(worker_status.accepting and session.accepting),
        needs_recovery=needs_recovery,
        owned_pid=owned_pid,
        active_operation=worker_status.active_operation,
        active_section=worker_status.active_section,
        last_hresult=worker_status.last_hresult,
        last_error=worker_status.last_error,
        open_workbooks=len(session.list_workbooks_sync()),
        notes=notes,
    )
