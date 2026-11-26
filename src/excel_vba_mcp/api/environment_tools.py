"""Read-only environment / policy discovery MCP tool (U2; plan_03 §6.1).

``xlvba_environment`` is a **non-COM**, **side-effect-free**, read-only tool
(like ``server_health``): it never starts the worker, never touches COM, never
creates snapshots / grants / classifications and never authorizes a mutation. It
reports the server's access mode, policy, honest build (always ``null``) and
capabilities, plus a per-open-workbook summary whose ``has_snapshot`` reads the
**in-memory** snapshot gate (``SnapshotService.has_snapshot`` -> ``_latest_snapshot``),
never the on-disk store.

The handler is thin and delegates to non-COM bookkeeping only (ARCH-009): the
worker status snapshot, the session bookkeeping and the published configuration /
path policy. The active pieces are published by the server lifespan via
:func:`set_environment_runtime` and cleared via :func:`clear_environment_runtime`.
When the runtime is not fully wired (e.g. the worker is down) the tool still
returns a truthful report with ``excel.available=false`` and no workbooks.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from excel_vba_mcp import __version__
from excel_vba_mcp.config import ServerConfig
from excel_vba_mcp.domain.models import (
    EnvironmentApprovalPolicy,
    EnvironmentCapabilities,
    EnvironmentExcelInfo,
    EnvironmentMacroPolicy,
    EnvironmentPathPolicy,
    EnvironmentPolicy,
    EnvironmentReport,
    EnvironmentServerInfo,
    EnvironmentSnapshotPolicy,
    EnvironmentWorkbook,
)
from excel_vba_mcp.domain.statuses import (
    AccessMode,
    PivotFilterRequestKind,
    PivotSourceType,
    WorkbookSessionState,
)

if TYPE_CHECKING:
    from excel_vba_mcp.application.session_service import SessionService
    from excel_vba_mcp.application.snapshot_service import SnapshotService
    from excel_vba_mcp.security.path_policy import PathPolicy

__all__ = [
    "clear_environment_runtime",
    "set_environment_runtime",
    "xlvba_environment",
]

# Design-time Excel build label only (NOT a detected build; ``excel.build`` stays
# null). The validated target build is Excel 16.0 14332.
_TARGET_BUILD = "16.0.14332"

# Published runtime pieces (set by the server lifespan). All optional so the tool
# degrades truthfully when the worker is down / the runtime is not yet wired.
_config: ServerConfig | None = None
_session: SessionService | None = None
_snapshot: SnapshotService | None = None
_path_policy: PathPolicy | None = None


def set_environment_runtime(
    *,
    config: ServerConfig | None,
    session: SessionService | None,
    snapshot: SnapshotService | None,
    path_policy: PathPolicy | None,
) -> None:
    """Publish the runtime pieces ``xlvba_environment`` reads (non-COM).

    Called by the server lifespan. ``session`` / ``snapshot`` may be ``None`` when
    the worker is down; the tool then reports ``excel.available=false`` and no
    workbooks. The tool never mutates any of these objects.
    """

    global _config, _session, _snapshot, _path_policy
    _config = config
    _session = session
    _snapshot = snapshot
    _path_policy = path_policy


def clear_environment_runtime() -> None:
    """Clear the published runtime pieces (called on server shutdown)."""

    global _config, _session, _snapshot, _path_policy
    _config = None
    _session = None
    _snapshot = None
    _path_policy = None


def _server_info(config: ServerConfig) -> EnvironmentServerInfo:
    """Build the server block from typed configuration (no COM)."""

    return EnvironmentServerInfo(
        name="excel-vba-mcp",
        version=__version__,
        access_mode=config.access_mode,
        macro_timeout_s=config.macro_timeout_s,
        open_timeout_s=config.open_timeout_s,
        liveness_ceiling_s=config.liveness_ceiling_s,
        log_level=config.log_level,
    )


def _capabilities() -> EnvironmentCapabilities:
    """Build the capability block from registration + status enums (no COM)."""

    return EnvironmentCapabilities(
        vba=True,
        macros=True,
        pivot_tables=True,
        background_jobs=True,
        recovery=True,
        non_activating_compile=False,
        pivot_filter_kinds=[kind.value for kind in PivotFilterRequestKind],
        pivot_source_types=[PivotSourceType.DATABASE.value],
    )


def _policy(config: ServerConfig, path_policy: PathPolicy | None) -> EnvironmentPolicy:
    """Build the policy block from config + path policy (no COM)."""

    requires_grant = config.access_mode == AccessMode.RESTRICTED
    return EnvironmentPolicy(
        approval=EnvironmentApprovalPolicy(
            mode=config.access_mode,
            mutating_requires_grant=requires_grant,
            executing_requires_grant=requires_grant,
            recovery_auto_allowed=True,
        ),
        snapshot=EnvironmentSnapshotPolicy(
            required_for_mutations=True,
            created_by="agent:snapshot_create",
        ),
        path=EnvironmentPathPolicy(
            allowed_roots=(
                [str(root) for root in path_policy.roots.values()] if path_policy else []
            ),
            allow_unc=path_policy.allow_unc if path_policy else False,
        ),
        macro=EnvironmentMacroPolicy(
            background_safe_required=True,
            r13_preflight="veto",
            classification_is_authorization=False,
        ),
    )


def _excel_info(session: SessionService | None) -> EnvironmentExcelInfo:
    """Build the excel block from non-COM worker/session bookkeeping.

    ``available`` is true only when the worker is running and accepting.
    ``application_version`` is a non-COM best-effort value taken from an open
    workbook's cached ``Application.Version`` (or ``None``). ``build`` is always
    ``None`` (never fabricated); ``vbom_available`` is ``None`` at server level.
    """

    if session is None:
        return EnvironmentExcelInfo(available=False, build=None, target_build=_TARGET_BUILD)

    worker_status = session.worker_status()
    available = bool(
        worker_status.accepting
        and worker_status.state == WorkbookSessionState.RUNNING
    )
    open_infos = session.list_workbooks_sync()
    application_version = open_infos[0].excel_version if open_infos else None
    return EnvironmentExcelInfo(
        available=available,
        application_version=application_version,
        build=None,
        target_build=_TARGET_BUILD,
        owned_pid=worker_status.owned_pid,
        vbom_available=None,
    )


async def _workbooks(session: SessionService | None) -> list[EnvironmentWorkbook]:
    """Build the per-open-workbook summary (non-COM; has_snapshot via the gate).

    ``has_snapshot`` comes from the in-memory gate surfaced through
    ``excel_get_status``'s policy block (``SnapshotService.has_snapshot`` ->
    ``_latest_snapshot``), never the on-disk store.
    """

    if session is None:
        return []
    workbooks: list[EnvironmentWorkbook] = []
    for info in session.list_workbooks_sync():
        status = await session.get_status(info.workbook_id)
        workbooks.append(
            EnvironmentWorkbook(
                workbook_id=info.workbook_id,
                read_only=status.read_only,
                mode=status.mode,
                has_snapshot=bool(status.policy and status.policy.has_snapshot),
                dirty=status.dirty,
            )
        )
    return workbooks


async def xlvba_environment() -> EnvironmentReport:
    """Report access mode, policy, honest build and capabilities (non-COM, read-only).

    This tool never starts the worker, never touches COM, never creates snapshots /
    grants / classifications and never authorizes a mutation
    (``policy.macro.classification_is_authorization`` is always ``false``). The
    build is always ``null`` (no 4-part source exists); ``excel.available`` is
    ``false`` when the worker is down.

    Returns:
        An :class:`EnvironmentReport` with ``server``, ``excel``, ``capabilities``,
        ``policy``, ``glossary_uri`` and per-open-workbook ``workbooks``.
    """

    config = _config or ServerConfig()
    workbooks = await _workbooks(_session)
    return EnvironmentReport(
        schema_version=1,
        server=_server_info(config),
        excel=_excel_info(_session),
        capabilities=_capabilities(),
        policy=_policy(config, _path_policy),
        glossary_uri="xlvba://glossary",
        workbooks=workbooks,
    )