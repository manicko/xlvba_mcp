"""Unit tests for the job registry MCP tool wrappers (Block E).

These tests never open Excel and never import COM. They drive the thin
``job_get`` / ``job_cancel`` tool handlers against a real
:class:`~excel_vba_mcp.application.job_registry.JobRegistry` backed by a
pure-Python fake worker, and assert the server registers both tools.
"""

from __future__ import annotations

import asyncio
import threading
from collections.abc import Iterator
from typing import Any

import pytest

from excel_vba_mcp.api.job_tools import (
    get_job_registry,
    job_cancel,
    job_get,
    set_job_registry,
)
from excel_vba_mcp.application.job_registry import JobRegistry, JobState
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import OperationClass

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture(autouse=True)
def _reset_registry() -> Iterator[None]:
    """Ensure the published registry is cleared before and after each test."""

    set_job_registry(None)
    yield
    set_job_registry(None)


class ImmediateWorker:
    """Fake worker whose ``submit`` returns immediately."""

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        return fn()


class BlockingWorker:
    """Fake worker whose ``submit`` blocks until ``release`` is set."""

    def __init__(self) -> None:
        self.started = threading.Event()
        self.release = threading.Event()

    async def submit(
        self,
        fn: Any,
        *,
        timeout: float,
        section: str = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
    ) -> Any:
        self.started.set()
        while not self.release.is_set():
            await asyncio.sleep(0.005)
        return "done"


# ---------------------------------------------------------------------------
# get_job_registry lifecycle
# ---------------------------------------------------------------------------
def test_get_job_registry_raises_when_unset() -> None:
    set_job_registry(None)

    with pytest.raises(ExcelVbaMcpError):
        get_job_registry()


# ---------------------------------------------------------------------------
# job_get happy path + errors
# ---------------------------------------------------------------------------
async def test_job_get_returns_terminal_result() -> None:
    registry = JobRegistry(ImmediateWorker())
    set_job_registry(registry)

    handle = await _start_and_finish(registry)
    status = await job_get(handle)

    assert status.job_id == handle
    assert status.status is JobState.SUCCEEDED
    assert status.result == "x"
    assert 1.0 <= status.poll_after_s <= 30.0


async def test_job_get_empty_arg_rejected() -> None:
    set_job_registry(JobRegistry(ImmediateWorker()))

    with pytest.raises(ExcelVbaMcpError):
        await job_get("")


async def test_job_get_unknown_id_is_not_found() -> None:
    set_job_registry(JobRegistry(ImmediateWorker()))

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await job_get("does-not-exist")
    assert excinfo.value.code is ErrorCode.NOT_FOUND


# ---------------------------------------------------------------------------
# job_cancel happy path + errors
# ---------------------------------------------------------------------------
async def test_job_cancel_records_intent_only() -> None:
    worker = BlockingWorker()
    registry = JobRegistry(worker)
    set_job_registry(registry)

    handle = registry.start(lambda: "x", timeout=5.0)
    while not worker.started.is_set():
        await asyncio.sleep(0.005)

    status = await job_cancel(handle.job_id)
    assert status.cancel_requested is True
    assert status.status is JobState.RUNNING

    # The op still runs to completion (intent-only).
    worker.release.set()
    await registry.join(handle.job_id, timeout=2.0)


async def test_job_cancel_empty_arg_rejected() -> None:
    set_job_registry(JobRegistry(ImmediateWorker()))

    with pytest.raises(ExcelVbaMcpError):
        await job_cancel("")


async def test_job_cancel_unknown_id_is_not_found() -> None:
    set_job_registry(JobRegistry(ImmediateWorker()))

    with pytest.raises(ExcelVbaMcpError) as excinfo:
        await job_cancel("does-not-exist")
    assert excinfo.value.code is ErrorCode.NOT_FOUND


# ---------------------------------------------------------------------------
# Server registration (item 8: real assertion, not a prose claim)
# ---------------------------------------------------------------------------
async def test_server_registers_job_tools() -> None:
    import excel_vba_mcp.server as server

    tools = {tool.name for tool in await server.mcp.list_tools()}
    assert "job_get" in tools
    assert "job_cancel" in tools


async def _start_and_finish(registry: JobRegistry) -> str:
    handle = registry.start(lambda: "x", timeout=5.0)
    await registry.join(handle.job_id, timeout=2.0)
    return handle.job_id
