"""Unit tests for ``excel_health`` owned-pid liveness (FE-18 / R-17).

``excel_health`` is a read-only diagnostic: it must report a dead owned Excel
process truthfully even when the worker bookkeeping still says ``running`` /
``accepting``. These tests drive the tool with a fake session + worker and a
fake ``_owned_pid_alive`` probe (no COM, no real Excel):

* (a) owned pid alive + worker running -> ``ok`` / ``needs_recovery`` False;
* (b) owned pid dead + worker "running" -> ``needs_recovery`` True /
  ``NEEDS_RECOVERY`` (the R-17 defect);
* (c) worker already ``NEEDS_RECOVERY`` -> still ``NEEDS_RECOVERY`` (unchanged);
* (d) liveness layer unavailable -> graceful fallback, no exception.
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from excel_vba_mcp.adapters.excel.com_worker import WorkerStatus
from excel_vba_mcp.api import health_tools, workbook_tools
from excel_vba_mcp.api.health_tools import excel_health
from excel_vba_mcp.domain.statuses import HealthStatus, WorkbookSessionState

pytestmark = pytest.mark.anyio


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


# ---------------------------------------------------------------------------
# Fakes (no COM)
# ---------------------------------------------------------------------------
class _FakeSession:
    """Duck-typed session exposing only what ``excel_health`` reads."""

    def __init__(self, worker_status: WorkerStatus, *, accepting: bool) -> None:
        self._worker_status = worker_status
        self.session_id = "sess-health"
        self._accepting = accepting

    @property
    def accepting(self) -> bool:
        return self._accepting

    def worker_status(self) -> WorkerStatus:
        return self._worker_status

    def list_workbooks_sync(self) -> list[object]:
        return []


def _worker(
    *,
    state: WorkbookSessionState = WorkbookSessionState.RUNNING,
    accepting: bool = True,
    owned_pid: int | None = 4242,
    notes: list[str] | None = None,
) -> WorkerStatus:
    return WorkerStatus(
        state=state,
        accepting=accepting,
        owned_pid=owned_pid,
        notes=list(notes or []),
    )


@pytest.fixture
def session_service() -> Iterator[None]:
    """Clear the published session service after each test."""

    yield
    workbook_tools.set_session_service(None)


# ---------------------------------------------------------------------------
# (a) owned pid alive + worker running -> ok
# ---------------------------------------------------------------------------
async def test_owned_pid_alive_reports_ok(
    session_service: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(health_tools, "_liveness_available", lambda: True)
    monkeypatch.setattr(health_tools, "_owned_pid_alive", lambda _pid: True)
    workbook_tools.set_session_service(  # type: ignore[arg-type]
        _FakeSession(_worker(owned_pid=4242), accepting=True)
    )

    health = await excel_health()

    assert health.status is HealthStatus.OK
    assert health.needs_recovery is False
    assert health.owned_pid == 4242
    # No dead-pid note was injected.
    assert all("not alive" not in note for note in health.notes)


# ---------------------------------------------------------------------------
# (b) owned pid dead + worker "running" -> needs_recovery (the R-17 defect)
# ---------------------------------------------------------------------------
async def test_owned_pid_dead_reports_needs_recovery(
    session_service: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(health_tools, "_liveness_available", lambda: True)
    monkeypatch.setattr(health_tools, "_owned_pid_alive", lambda _pid: False)
    workbook_tools.set_session_service(  # type: ignore[arg-type]
        _FakeSession(_worker(owned_pid=4242, accepting=True), accepting=True)
    )

    health = await excel_health()

    assert health.needs_recovery is True
    assert health.status is HealthStatus.NEEDS_RECOVERY
    assert health.owned_pid == 4242
    # A clear, diagnostic note explains the dead owned pid.
    assert any("not alive" in note and "4242" in note for note in health.notes)


# ---------------------------------------------------------------------------
# (c) worker already NEEDS_RECOVERY -> still NEEDS_RECOVERY (unchanged)
# ---------------------------------------------------------------------------
async def test_worker_already_needs_recovery_unchanged(
    session_service: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Even if the probe would say "dead", the state-based verdict already wins;
    # the probe must not mask or duplicate the state signal.
    monkeypatch.setattr(health_tools, "_liveness_available", lambda: True)
    monkeypatch.setattr(health_tools, "_owned_pid_alive", lambda _pid: True)
    workbook_tools.set_session_service(  # type: ignore[arg-type]
        _FakeSession(
            _worker(
                state=WorkbookSessionState.NEEDS_RECOVERY,
                accepting=False,
                owned_pid=4242,
                notes=["EXCEL_HUNG: worker reset required"],
            ),
            accepting=False,
        )
    )

    health = await excel_health()

    assert health.status is HealthStatus.NEEDS_RECOVERY
    assert health.needs_recovery is True
    assert health.notes == ["EXCEL_HUNG: worker reset required"]


# ---------------------------------------------------------------------------
# (d) liveness layer unavailable -> graceful fallback (no exception)
# ---------------------------------------------------------------------------
async def test_liveness_unavailable_degrades_gracefully(
    session_service: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Liveness layer unavailable: the probe is skipped entirely and the pure
    # state-based behaviour is kept (a missing Win32 layer never flags "dead").
    monkeypatch.setattr(health_tools, "_liveness_available", lambda: False)

    def _fail_probe(_pid: int) -> bool:
        raise AssertionError("the pid probe must not run when liveness is unavailable")

    monkeypatch.setattr(health_tools, "_owned_pid_alive", _fail_probe)
    workbook_tools.set_session_service(  # type: ignore[arg-type]
        _FakeSession(_worker(owned_pid=4242, accepting=True), accepting=True)
    )

    health = await excel_health()

    assert health.status is HealthStatus.OK
    assert health.needs_recovery is False
    assert health.owned_pid == 4242