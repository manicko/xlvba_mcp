"""Job registry MCP tools (Block E — E/F boundary).

Thin, async tool handlers for the minimal in-memory job registry:
``job_get`` (poll one job's status/result) and ``job_cancel`` (record cancel
intent only). They are deliberately **thin**: each validates its arguments and
delegates to :class:`~excel_vba_mcp.application.job_registry.JobRegistry`. They
**never touch COM** (ARCH-009).

The active registry is published here via :func:`set_job_registry` (called from
the server lifespan) and read via :func:`get_job_registry`.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from excel_vba_mcp.application.job_registry import JobRegistry, JobStatus, RunningJob
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import OperationClass

# The active job registry, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_job_registry: JobRegistry | None = None


def set_job_registry(registry: JobRegistry | None) -> None:
    """Publish (or clear) the active job registry used by the tools."""

    global _job_registry
    _job_registry = registry


def get_job_registry() -> JobRegistry:
    """Return the active job registry or raise if it is not configured."""

    if _job_registry is None:
        raise ExcelVbaMcpError(
            "The job registry is not initialized.",
            context={"tool": "job_tools"},
        )
    return _job_registry


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


def start_long_op(
    fn: Callable[[], Any],
    *,
    timeout: float = 300.0,
    section: str = "diagnostic",
    op_class: OperationClass = OperationClass.READ_ONLY,
    idempotency_key: str | None = None,
) -> RunningJob:
    """Schedule a long op on the registry and return a prompt ``job_id`` handle.

    This is the Block F producer helper: a long op calls it instead of blocking,
    returning ``{status: "running", job_id, poll_after_s}`` well within the client
    wall-clock. The op runs on the existing single-STA ``submit`` (a sync callable
    is wrapped by the FIFO; an async callable is awaited as the registry task and
    itself uses ``submit``) — no second STA is created.

    Args:
        fn: A sync callable (wrapped by ``submit``) or an async callable (awaited
            as the registry-owned task).
        timeout: Soft per-call deadline hint (metadata for the async path).
        section: Work section tag (drives settings scope and recovery rules).
        op_class: Operation classification (drives retry eligibility).
        idempotency_key: Optional key; a repeated key while running dedupes to the
            existing ``job_id`` and a completed key returns the stored result.

    Returns:
        A :class:`RunningJob` (``status="running"`` for a new job; the terminal
        status and stored result when an idempotency key matches a completed job).
    """

    registry = get_job_registry()
    return registry.start(
        fn,
        timeout=timeout,
        section=section,
        op_class=op_class,
        idempotency_key=idempotency_key,
    )


async def job_get(job_id: str) -> JobStatus:
    """Return one job's status/result by ``job_id`` (poll; never touches COM).

    Args:
        job_id: The id returned when the long op was started.

    Returns:
        A :class:`JobStatus` with the job's ``status`` (``queued`` / ``running`` /
        ``succeeded`` / ``failed`` / ``expired`` / ``unknown``), the ``result`` once
        terminal, and a ``poll_after_s`` hint (clamped to ``[1, 30]``).
    """

    _require_non_empty(job_id, "job_id")
    registry = get_job_registry()
    return registry.status(job_id)


async def job_cancel(job_id: str) -> JobStatus:
    """Record a job's cancel **intent** (intent-only; never interrupts COM).

    The uninterruptible COM call on the single-STA thread runs to completion; this
    only records the caller's intent so a poller can observe it.

    Args:
        job_id: The id returned when the long op was started.

    Returns:
        The :class:`JobStatus` after the cancel intent was recorded.
    """

    _require_non_empty(job_id, "job_id")
    registry = get_job_registry()
    return registry.cancel(job_id)


__all__ = [
    "get_job_registry",
    "job_cancel",
    "job_get",
    "set_job_registry",
    "start_long_op",
]
