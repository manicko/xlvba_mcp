"""Agent-directed recovery MCP tool (Block A; report_07 §4.1, §5).

``excel_recover`` is the **agent-directed** recovery entry point. The server never
auto-terminates a live pid on a deadline; the agent observes the health/verdict
signals and decides to call ``excel_recover``. The tool is **thin**: it validates
arguments and delegates to
:meth:`excel_vba_mcp.application.session_service.SessionService.recover_owned_process`.
It never touches COM directly (ARCH-009).

Policy (allow-by-default RECOVERY class): recovery is auto-allowed in both FULL and
RESTRICTED mode (audited by the service), is **pid-scoped** (never a wildcard) and
discards **working copies only** — originals are never modified. A mutating op
stays denied in RESTRICTED. The response states it is destructive to working copies
and that originals are safe.

The active session service is published here via :func:`set_session_service` (called
from the server lifespan) and read via :func:`get_session_service`.
"""

from __future__ import annotations

from pydantic import BaseModel, ConfigDict

from excel_vba_mcp.api.workbook_tools import get_session_service

__all__ = ["RecoveryResult", "excel_recover"]


class RecoveryResult(BaseModel):
    """Structured ``excel_recover`` payload (no COM; originals safe)."""

    model_config = ConfigDict(extra="forbid")

    ok: bool
    terminated: bool
    confirmed: bool
    recovery_possible: bool
    process_restarted: bool
    prior_job: str
    what_lost: str
    reason: str | None = None
    owned_pid: int | None = None
    approval_source: str | None = None


async def excel_recover(reason: str = "agent_directed") -> RecoveryResult:
    """Recover the owned Excel process (agent-directed; fail-closed; originals safe).

    Terminates the wedged owned ``EXCEL.EXE`` **through the held process handle**
    (``TerminateProcess`` -> ``WaitForSingleObject`` == ``WAIT_OBJECT_0``), discards
    the working copies and rebuilds a fresh hidden instance. Recovery is **only**
    possible when a live, identity-verified handle is held; otherwise it is
    **refused** (fail closed) and the documented fallback is a server restart (U8).
    The server **never** auto-terminates a live pid on a deadline — this runs only
    when the agent calls it. An interrupted mutation is reported as
    ``prior_job="unknown"`` and is **never** auto-retried.

    Args:
        reason: Short label recorded in the audit entry / outcome.

    Returns:
        A :class:`RecoveryResult` with ``ok`` / ``terminated`` / ``confirmed`` /
        ``recovery_possible`` / ``process_restarted`` / ``prior_job`` and
        ``what_lost`` (destructive to working copies only; originals safe).
    """

    session = get_session_service()
    outcome = await session.recover_owned_process(reason=reason)
    return RecoveryResult.model_validate(outcome)
