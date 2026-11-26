"""Loop-side job registry that OWNS the result (Block E; extended in Block F).

The :class:`JobRegistry` is a loop-side, lock-guarded, **non-COM** registry that
**owns the result** of a long-running operation so a client timeout or disconnect
never drops it. It never creates a second STA and never touches COM itself:

* a **sync** callable is wrapped by the existing single-STA
  ``ExcelComWorker.submit`` FIFO (Block E behaviour, unchanged);
* an **async** callable (a service coroutine that itself calls ``submit``) is
  awaited as the registry-owned task, so the same single-STA ``submit`` runs the
  op while the registry retains the result (Block F producer wiring).

:meth:`JobRegistry.start` returns a ``job_id`` promptly, so a long op responds
``{status: "running", job_id, poll_after_s}`` well within the client wall-clock.
The agent then **pulls/poll**s :meth:`JobRegistry.status` (MCP ``job_get``) by
``job_id``.

Lifecycle: ``queued -> running -> {succeeded | failed | expired | unknown |
interrupted_unknown}``.

* ``succeeded`` — the wrapped op returned a value (stored on the job).
* ``failed`` — the wrapped op raised a domain / unexpected error.
* ``expired`` — the wrapped op exceeded its execution deadline
  (``ExcelHungError``); the call is left running on the STA (never cancelled).
* ``unknown`` — reserved; never produced by this registry.
* ``interrupted_unknown`` — a persisted non-terminal job reconciled at startup
  (Block F); never ``succeeded``, never auto-retried.

``job_cancel`` is **intent-only**: it sets ``cancel_requested`` on the job but
**cannot** interrupt the uninterruptible COM call on the STA thread — the job
still runs to a terminal state.

Persistence (Block F): when a :class:`~excel_vba_mcp.storage.job_store.JobStore`
is supplied, a RUNNING marker is written at :meth:`start` and the record is
finalized on every terminal state. After a restart, :meth:`status` falls back to
the store so an already-completed result is retrievable by ``job_id``.
"""

from __future__ import annotations

import asyncio
import inspect
import threading
import time
from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from enum import StrEnum
from typing import Any, TypeVar, cast
from uuid import uuid4

from pydantic import BaseModel, Field

from excel_vba_mcp.domain.errors import ComBusyError, ErrorCode, ExcelHungError, ExcelVbaMcpError
from excel_vba_mcp.domain.models import DomainModel
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.storage.job_store import JobRecord, JobStore

T = TypeVar("T")

# ``poll_after_s`` guidance (report_07 §4.1): clamped to [1, 30], default 5 s
# (well below the 60 s client wall-clock).
_POLL_DEFAULT_S = 5.0
_POLL_MIN_S = 1.0
_POLL_MAX_S = 30.0

# Section type mirrors ``com_worker.Section`` without importing COM.
Section = str


class JobState(StrEnum):
    """Lifecycle state of a registry job."""

    QUEUED = "queued"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    EXPIRED = "expired"
    UNKNOWN = "unknown"
    # Block F (additive): a persisted non-terminal job reconciled at startup.
    INTERRUPTED_UNKNOWN = "interrupted_unknown"


_TERMINAL_STATES = frozenset(
    {
        JobState.SUCCEEDED,
        JobState.FAILED,
        JobState.EXPIRED,
        JobState.UNKNOWN,
        JobState.INTERRUPTED_UNKNOWN,
    }
)

# Result-payload statuses that mark a job as deadline-exceeded / interrupted even
# when the registry job itself already reached a terminal state (e.g. a macro that
# returned ``SUSPECTED_STALLED``). Such a job is finalized to INTERRUPTED_UNKNOWN
# by :meth:`JobRegistry.mark_interrupted` after a successful recovery.
_DEADLINE_RESULT_STATUSES = frozenset({"suspected_stalled", "needs_recovery", "timeout"})


def _clamp_poll(value: float) -> float:
    """Clamp a ``poll_after_s`` hint to ``[1, 30]`` (report_07 §4.1)."""

    return max(_POLL_MIN_S, min(_POLL_MAX_S, float(value)))


def _jsonable(value: Any) -> Any:
    """Return a JSON-serializable form of ``value`` for persistence."""

    if isinstance(value, BaseModel):
        return value.model_dump(mode="json")
    return value


def _result_status(value: Any) -> str | None:
    """Best-effort ``status`` of a job result payload (DTO or dict), else ``None``."""

    if isinstance(value, BaseModel):
        raw = getattr(value, "status", None)
    elif isinstance(value, dict):
        raw = value.get("status")
    else:
        raw = None
    return None if raw is None else str(raw)


def _is_interruptible(job: Job) -> bool:
    """True when a job must be finalized to INTERRUPTED_UNKNOWN on recovery."""

    if job.state in (JobState.QUEUED, JobState.RUNNING, JobState.EXPIRED):
        return True
    return _result_status(job.result) in _DEADLINE_RESULT_STATUSES


class JobEvent(DomainModel):
    """Additive history entry appended when a job is interrupted (never overwrites)."""

    at: datetime
    state: str
    reason: str
    result_status: str | None = None


@dataclass
class Job:
    """One registry-owned job record (loop-side, lock-guarded; no COM)."""

    job_id: str
    state: JobState
    created_at: float
    updated_at: float
    op_class: OperationClass
    section: Section
    poll_after_s: float = _POLL_DEFAULT_S
    started: bool = False
    cancel_requested: bool = False
    result: Any = None
    error_code: str | None = None
    error_text: str | None = None
    retry_after_s: float | None = None
    # Block F (additive): idempotency key + interrupted-job outcome metadata.
    idempotency_key: str | None = None
    outcome: str | None = None
    retryable: bool | None = None
    events: list[JobEvent] = field(default_factory=list)
    created_wall: datetime = field(default_factory=lambda: datetime.now(UTC))

    @property
    def terminal(self) -> bool:
        return self.state in _TERMINAL_STATES


class JobStatus(DomainModel):
    """Serializable ``job_get`` payload (status / result / poll hint)."""

    job_id: str
    status: JobState
    started: bool
    cancel_requested: bool
    poll_after_s: float
    result: Any | None = None
    error_code: str | None = None
    error_text: str | None = None
    retry_after_s: float | None = None
    # Block F (additive): lifecycle phase + interrupted-job outcome metadata.
    phase: str | None = None
    outcome: str | None = None
    retryable: bool | None = None
    events: list[JobEvent] = Field(default_factory=list)


class RunningJob(DomainModel):
    """Prompt long-op response: ``{status: "running", job_id, poll_after_s}``.

    When an idempotency key matches an already-completed job, ``status`` carries
    the terminal state and ``result`` the stored payload (additive fields).
    """

    status: str = "running"
    job_id: str
    poll_after_s: float
    result: Any | None = None
    error_code: str | None = None
    error_text: str | None = None
    retry_after_s: float | None = None
    outcome: str | None = None
    retryable: bool | None = None


class JobRegistry:
    """Loop-side, lock-guarded registry that OWNS the result.

    Args:
        worker: The single-STA COM worker whose ``submit`` FIFO is wrapped for
            sync callables. Only ``await worker.submit(fn, timeout=..., section=...,
            op_class=...)`` is used; the registry never touches COM directly.
        poll_default_s: Default ``poll_after_s`` hint (clamped to ``[1, 30]``).
        store: Optional :class:`JobStore` (Block F). When supplied, a RUNNING
            marker is persisted at start and the record is finalized on terminal;
            after a restart :meth:`status` falls back to the store.
    """

    def __init__(
        self,
        worker: Any,
        *,
        poll_default_s: float = _POLL_DEFAULT_S,
        store: JobStore | None = None,
    ) -> None:
        self._worker = worker
        self._lock = threading.Lock()
        self._jobs: dict[str, Job] = {}
        self._tasks: dict[str, asyncio.Task[None]] = {}
        self._by_key: dict[str, str] = {}
        self._poll_default_s = _clamp_poll(poll_default_s)
        self._store = store

    # ------------------------------------------------------------------
    # Public API (loop side)
    # ------------------------------------------------------------------
    def start(
        self,
        fn: Callable[[], T],
        *,
        timeout: float,
        section: Section = "diagnostic",
        op_class: OperationClass = OperationClass.READ_ONLY,
        idempotency_key: str | None = None,
    ) -> RunningJob:
        """Schedule ``fn`` and return a prompt ``job_id`` handle.

        The registry — not the caller — owns the running :class:`asyncio.Task`, so
        a client timeout / disconnect never drops the result: it is retained on the
        :class:`Job` and retrievable via :meth:`status` by ``job_id``.

        ``fn`` may be a sync callable (wrapped by the single-STA ``submit`` FIFO)
        or an async callable (awaited as the registry task; the coroutine itself
        uses the single-STA ``submit`` — no second STA).

        When ``idempotency_key`` is given and a job with the same key exists, the
        existing job is returned (dedupe): a running job yields its ``job_id``; a
        completed job yields its terminal status and stored result.
        """

        if idempotency_key is not None:
            existing = self._lookup_key(idempotency_key)
            if existing is not None:
                return self._running_from_job(existing)

        job_id = f"job-{uuid4().hex}"
        now = time.monotonic()
        job = Job(
            job_id=job_id,
            state=JobState.QUEUED,
            created_at=now,
            updated_at=now,
            op_class=OperationClass(op_class),
            section=section,
            poll_after_s=self._poll_default_s,
            idempotency_key=idempotency_key,
        )
        with self._lock:
            self._jobs[job_id] = job
            if idempotency_key is not None:
                self._by_key[idempotency_key] = job_id
        # Persist a RUNNING marker at start (Block F); reconcile turns any leftover
        # marker into interrupted_unknown after a restart.
        self._persist(job, JobState.RUNNING)
        if inspect.iscoroutinefunction(fn):
            task = asyncio.create_task(
                self._run_async(job, cast(Callable[[], Awaitable[T]], fn))
            )
        else:
            task = asyncio.create_task(self._run(job, fn, timeout, section, op_class))
        with self._lock:
            self._tasks[job_id] = task
        return RunningJob(job_id=job_id, poll_after_s=job.poll_after_s)

    def get(self, job_id: str) -> Job | None:
        """Return the live :class:`Job` record (or a store-backed one, or ``None``)."""

        with self._lock:
            job = self._jobs.get(job_id)
        if job is not None:
            return job
        record = self._store.get(job_id) if self._store is not None else None
        return self._record_to_job(record) if record is not None else None

    def status(self, job_id: str) -> JobStatus:
        """Return the serializable status/result DTO for ``job_id``.

        Live jobs are served from memory; after a restart a persisted terminal job
        is served from the :class:`JobStore`.

        Raises:
            ExcelVbaMcpError: ``NOT_FOUND`` when ``job_id`` is unknown.
        """

        job = self.get(job_id)
        if job is None:
            raise ExcelVbaMcpError(
                f"Unknown job_id {job_id!r}.",
                code=ErrorCode.NOT_FOUND,
                context={"job_id": job_id},
            )
        return self._to_status(job)

    def cancel(self, job_id: str) -> JobStatus:
        """Mark a job's cancel **intent** (intent-only; never interrupts COM).

        The uninterruptible COM call on the STA thread runs to completion; this only
        records the caller's intent so a poller can observe it. A store-backed
        (already terminal) job is a no-op that returns its persisted status.

        Raises:
            ExcelVbaMcpError: ``NOT_FOUND`` when ``job_id`` is unknown.
        """

        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None:
                if not job.terminal:
                    job.cancel_requested = True
                    job.updated_at = time.monotonic()
                return self._to_status(job)
        record = self._store.get(job_id) if self._store is not None else None
        if record is None:
            raise ExcelVbaMcpError(
                f"Unknown job_id {job_id!r}.",
                code=ErrorCode.NOT_FOUND,
                context={"job_id": job_id},
            )
        return self._record_to_status(record)

    def mark_interrupted(self, *, reason: str = "recovery") -> list[JobStatus]:
        """Finalize every interruptible job to ``INTERRUPTED_UNKNOWN`` (definitive).

        A job is interruptible when it is non-terminal (``queued`` / ``running``),
        already ``expired`` (deadline-exceeded), or carries a deadline-exceeded
        result payload (``suspected_stalled`` / ``needs_recovery`` / ``timeout``).
        The prior state is **appended** to an additive ``events`` history (never
        overwritten), then the job is set to ``interrupted_unknown`` with
        ``outcome='unknown'`` and ``retryable=False`` and persisted via the
        :class:`JobStore` (atomic).

        This is wired into the successful ``excel_recover`` path: after the owned
        process is confirmed restarted, the interrupted jobs are finalized so a
        later ``job_get`` reports a definitive unknown outcome (never completed,
        never replayed). A late completion from the still-running STA call is
        ignored by :meth:`_finalize` so it cannot overwrite the interrupted state.

        Returns:
            The :class:`JobStatus` of each job that was marked (empty when none).
        """

        now = datetime.now(UTC)
        marked: list[JobStatus] = []
        with self._lock:
            candidates = [job for job in self._jobs.values() if _is_interruptible(job)]
        for job in candidates:
            with self._lock:
                if job.state is JobState.INTERRUPTED_UNKNOWN or not _is_interruptible(job):
                    continue
                job.events = [
                    *job.events,
                    JobEvent(
                        at=now,
                        state=job.state.value,
                        reason=reason,
                        result_status=_result_status(job.result),
                    ),
                ]
                job.state = JobState.INTERRUPTED_UNKNOWN
                job.outcome = "unknown"
                job.retryable = False
                job.updated_at = time.monotonic()
            self._persist(job, JobState.INTERRUPTED_UNKNOWN)
            marked.append(self._to_status(job))
        return marked

    async def join(self, job_id: str, *, timeout: float = 5.0) -> JobStatus:
        """Await a job's background task (bounded); useful for tests / callers.

        Returns the terminal :class:`JobStatus`. A timeout returns the current
        (possibly still-running) status without raising.
        """

        task = self._tasks.get(job_id)
        if task is not None:
            try:
                await asyncio.wait_for(asyncio.shield(task), timeout=timeout)
            except TimeoutError:
                pass
            except Exception:  # noqa: BLE001 - the job records its own failure
                pass
        return self.status(job_id)

    # ------------------------------------------------------------------
    # Internal (loop side)
    # ------------------------------------------------------------------
    async def _run(
        self,
        job: Job,
        fn: Callable[[], T],
        timeout: float,
        section: Section,
        op_class: OperationClass,
    ) -> None:
        """Drive one sync job queued -> running -> terminal, owning the result."""

        self._transition(job, JobState.RUNNING, started=True)
        try:
            result = await self._worker.submit(
                fn, timeout=timeout, section=section, op_class=op_class
            )
        except ExcelHungError as exc:
            # Execution deadline exceeded: the call is left running on the STA.
            self._finalize(
                job,
                JobState.EXPIRED,
                error_code=exc.code.value,
                error_text=str(exc),
            )
            return
        except ComBusyError as exc:
            # Queue-wait timeout (never started): a transient busy; the job fails
            # and the agent may re-issue. ``retry_after_s`` is surfaced as a hint.
            self._finalize(
                job,
                JobState.FAILED,
                error_code=exc.code.value,
                error_text=str(exc),
                retry_after_s=_coerce_retry_after(exc),
            )
            return
        except ExcelVbaMcpError as exc:
            self._finalize(job, JobState.FAILED, error_code=exc.code.value, error_text=str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - any unexpected error fails the job
            self._finalize(
                job,
                JobState.FAILED,
                error_code=ErrorCode.COM_BUSY.value,
                error_text=f"{type(exc).__name__}: {exc}",
            )
            return
        self._finalize(job, JobState.SUCCEEDED, result=result)

    async def _run_async(self, job: Job, fn: Callable[[], Awaitable[T]]) -> None:
        """Drive one async job (a service coroutine) -> terminal, owning the result.

        The coroutine itself runs the op on the single-STA ``submit``; the registry
        only owns the task so a client timeout / disconnect never drops the result.
        """

        self._transition(job, JobState.RUNNING, started=True)
        try:
            result = await fn()
        except ExcelHungError as exc:
            self._finalize(
                job,
                JobState.EXPIRED,
                error_code=exc.code.value,
                error_text=str(exc),
            )
            return
        except ComBusyError as exc:
            self._finalize(
                job,
                JobState.FAILED,
                error_code=exc.code.value,
                error_text=str(exc),
                retry_after_s=_coerce_retry_after(exc),
            )
            return
        except ExcelVbaMcpError as exc:
            self._finalize(job, JobState.FAILED, error_code=exc.code.value, error_text=str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - any unexpected error fails the job
            self._finalize(
                job,
                JobState.FAILED,
                error_code=ErrorCode.COM_BUSY.value,
                error_text=f"{type(exc).__name__}: {exc}",
            )
            return
        self._finalize(job, JobState.SUCCEEDED, result=result)

    def _transition(self, job: Job, state: JobState, *, started: bool = False) -> None:
        with self._lock:
            # A recovery-interrupted job is definitive: a late task start must not
            # flip it back to running.
            if job.state is JobState.INTERRUPTED_UNKNOWN:
                return
            job.state = state
            job.updated_at = time.monotonic()
            if started:
                job.started = True

    def _finalize(
        self,
        job: Job,
        state: JobState,
        *,
        result: Any = None,
        error_code: str | None = None,
        error_text: str | None = None,
        retry_after_s: float | None = None,
    ) -> None:
        with self._lock:
            # Definitive interrupted outcome: a late completion from the still-running
            # STA call must never overwrite INTERRUPTED_UNKNOWN (never completed).
            if job.state is JobState.INTERRUPTED_UNKNOWN:
                return
            job.state = state
            job.updated_at = time.monotonic()
            job.result = result
            job.error_code = error_code
            job.error_text = error_text
            job.retry_after_s = retry_after_s
        self._persist(job, state)

    def _persist(self, job: Job, state: JobState) -> None:
        """Persist ``job`` as ``state`` (RUNNING marker or terminal record)."""

        store = self._store
        if store is None:
            return
        store.save(
            JobRecord(
                job_id=job.job_id,
                state=state.value,
                created_at=job.created_wall,
                updated_at=datetime.now(UTC),
                op_class=job.op_class.value,
                section=job.section,
                poll_after_s=job.poll_after_s,
                idempotency_key=job.idempotency_key,
                result=_jsonable(job.result),
                error_code=job.error_code,
                error_text=job.error_text,
                retry_after_s=job.retry_after_s,
                outcome=job.outcome,
                retryable=job.retryable,
                events=[event.model_dump(mode="json") for event in job.events],
            )
        )

    def _lookup_key(self, key: str) -> Job | None:
        """Return the job stored under ``key`` (memory first, then the store)."""

        with self._lock:
            job_id = self._by_key.get(key)
            job = self._jobs.get(job_id) if job_id is not None else None
        if job is not None:
            return job
        record = self._store.find_by_idempotency_key(key) if self._store is not None else None
        return self._record_to_job(record) if record is not None else None

    @staticmethod
    def _running_from_job(job: Job) -> RunningJob:
        return RunningJob(
            job_id=job.job_id,
            poll_after_s=job.poll_after_s,
            status=job.state.value,
            result=job.result,
            error_code=job.error_code,
            error_text=job.error_text,
            retry_after_s=job.retry_after_s,
            outcome=job.outcome,
            retryable=job.retryable,
        )

    @staticmethod
    def _record_to_job(record: JobRecord) -> Job:
        return Job(
            job_id=record.job_id,
            state=JobState(record.state),
            created_at=0.0,
            updated_at=0.0,
            op_class=OperationClass(record.op_class),
            section=record.section,
            poll_after_s=record.poll_after_s,
            started=record.state != JobState.QUEUED.value,
            result=record.result,
            error_code=record.error_code,
            error_text=record.error_text,
            retry_after_s=record.retry_after_s,
            idempotency_key=record.idempotency_key,
            outcome=record.outcome,
            retryable=record.retryable,
            events=[JobEvent(**event) for event in record.events],
            created_wall=record.created_at,
        )

    @staticmethod
    def _record_to_status(record: JobRecord) -> JobStatus:
        state = JobState(record.state)
        return JobStatus(
            job_id=record.job_id,
            status=state,
            started=record.state != JobState.QUEUED.value,
            cancel_requested=False,
            poll_after_s=record.poll_after_s,
            result=record.result,
            error_code=record.error_code,
            error_text=record.error_text,
            retry_after_s=record.retry_after_s,
            phase=_phase_for(state),
            outcome=record.outcome,
            retryable=record.retryable,
            events=[JobEvent(**event) for event in record.events],
        )

    @staticmethod
    def _to_status(job: Job) -> JobStatus:
        return JobStatus(
            job_id=job.job_id,
            status=job.state,
            started=job.started,
            cancel_requested=job.cancel_requested,
            poll_after_s=job.poll_after_s,
            result=job.result,
            error_code=job.error_code,
            error_text=job.error_text,
            retry_after_s=job.retry_after_s,
            phase=_phase_for(job.state),
            outcome=job.outcome,
            retryable=job.retryable,
            events=list(job.events),
        )


def _phase_for(state: JobState) -> str:
    """Return the lifecycle phase label for ``state`` (additive ``phase`` field)."""

    if state is JobState.INTERRUPTED_UNKNOWN:
        return "interrupted"
    return state.value


def _coerce_retry_after(exc: ComBusyError) -> float | None:
    """Extract ``retry_after_s`` from a busy error's context, if present."""

    raw = exc.context.get("retry_after_s") if exc.context else None
    if isinstance(raw, (int, float)):
        return float(raw)
    return None


__all__ = [
    "Job",
    "JobRegistry",
    "JobState",
    "JobStatus",
    "RunningJob",
]