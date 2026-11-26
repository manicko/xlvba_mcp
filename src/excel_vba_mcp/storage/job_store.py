"""Filesystem persistence for background jobs (Block F; report_07 §6.2, plan_02 D5).

:class:`JobStore` is a sibling of :class:`~excel_vba_mcp.storage.run_repository.RunRepository`
(it is **not** ``TestRunResult``-typed). It persists one JSON document per job under
an allowlisted root and is pure filesystem — it never touches Excel / COM and holds
no proxies.

Guarantees:

* **Atomic.** Every document is written with the extracted
  :func:`~excel_vba_mcp.storage.run_repository._atomic_write_json` helper
  (temp file + ``fsync`` + ``os.replace``); a crash never leaves a half-written
  job behind.
* **RUNNING marker + finalize.** The registry persists a ``running`` marker at job
  start and finalizes the record on a terminal state (``succeeded`` / ``failed`` /
  ``expired`` / ``interrupted_unknown``). Completed jobs keep their result payload
  so a later process can retrieve them by ``job_id``.
* **Startup reconcile.** Any non-terminal (``queued`` / ``running``) record found on
  disk is reconciled to ``interrupted_unknown`` — **never** ``succeeded`` and
  **never** retried. The reconciled records are returned by
  :meth:`JobStore.reconciled` for logging / audit.
* **Validated ids.** A job id becomes a file name, so the run-id validator is
  reused (``PATH_NOT_ALLOWED`` on an unsafe id).

The store keeps an in-memory index loaded at construction so ``get`` /
``find_by_idempotency_key`` are cheap and a freshly constructed store (a restart)
can serve already-persisted results.
"""

from __future__ import annotations

import threading
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pydantic import Field

from excel_vba_mcp.domain.models import DomainModel
from excel_vba_mcp.storage.run_repository import _atomic_write_json, _validate_run_id

__all__ = ["JobRecord", "JobStore"]

#: File suffix of a persisted job document.
_JOB_SUFFIX = ".json"

#: Non-terminal states that a startup reconcile must convert to ``interrupted_unknown``.
_NON_TERMINAL_STATES = frozenset({"queued", "running"})

#: Terminal state assigned to an interrupted (non-terminal) persisted job.
INTERRUPTED_UNKNOWN = "interrupted_unknown"


class JobRecord(DomainModel):
    """One persisted job document (JSON-serializable; no COM references).

    ``result`` holds the completed job's payload (already JSON-normalized by the
    registry). ``outcome`` / ``retryable`` are set for an interrupted job so a
    poller can distinguish ``interrupted_unknown`` (``outcome='unknown'``,
    ``retryable=False``) from a normal terminal state.
    """

    job_id: str
    state: str
    created_at: datetime
    updated_at: datetime
    op_class: str
    section: str
    poll_after_s: float
    idempotency_key: str | None = None
    result: Any | None = None
    error_code: str | None = None
    error_text: str | None = None
    retry_after_s: float | None = None
    outcome: str | None = None
    retryable: bool | None = None
    # Additive event history (e.g. a prior ``suspected_stalled`` appended when a
    # recovery finalizes the job to ``interrupted_unknown``). Stored as plain dicts
    # so the storage layer stays independent of the application DTOs.
    events: list[dict[str, Any]] = Field(default_factory=list)


class JobStore:
    """One JSON document per job under ``root`` (atomic; startup reconcile).

    Args:
        root: Directory that holds one ``<job_id>.json`` document per job. Created
            on first write; it never contains repository sources.
    """

    def __init__(self, root: Path) -> None:
        self._root = Path(root)
        self._root.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._records: dict[str, JobRecord] = {}
        self._by_key: dict[str, str] = {}
        self._reconciled: list[JobRecord] = self._load_and_reconcile()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    @property
    def root(self) -> Path:
        """The job store root directory."""

        return self._root

    @property
    def reconciled(self) -> list[JobRecord]:
        """Records converted to ``interrupted_unknown`` by the startup reconcile."""

        return list(self._reconciled)

    def path_for(self, job_id: str) -> Path:
        """Return the validated absolute path of the ``job_id`` document."""

        return self._root / f"{_validate_run_id(job_id)}{_JOB_SUFFIX}"

    def save(self, record: JobRecord) -> None:
        """Atomically persist ``record`` and update the in-memory index.

        Args:
            record: The job record to persist (RUNNING marker or terminal result).

        Raises:
            PathNotAllowedError: If the job id is not a safe file name.
            OSError: If the atomic write / replace fails (nothing is left behind).
        """

        _validate_run_id(record.job_id)
        payload = record.model_dump_json(indent=2).encode("utf-8")
        with self._lock:
            _atomic_write_json(self.path_for(record.job_id), payload)
            self._records[record.job_id] = record
            if record.idempotency_key:
                self._by_key[record.idempotency_key] = record.job_id

    def get(self, job_id: str) -> JobRecord | None:
        """Return the persisted record for ``job_id`` (or ``None`` when unknown)."""

        with self._lock:
            return self._records.get(job_id)

    def find_by_idempotency_key(self, key: str) -> JobRecord | None:
        """Return the most recent record stored under ``key`` (or ``None``)."""

        with self._lock:
            job_id = self._by_key.get(key)
            if job_id is None:
                return None
            return self._records.get(job_id)

    def all_records(self) -> list[JobRecord]:
        """Return every loaded record (unordered)."""

        with self._lock:
            return list(self._records.values())

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------
    def _load_and_reconcile(self) -> list[JobRecord]:
        """Load every document; reconcile non-terminal records to interrupted."""

        reconciled: list[JobRecord] = []
        for path in sorted(self._root.glob(f"*{_JOB_SUFFIX}")):
            if not path.is_file() or path.name.startswith("."):
                continue
            try:
                record = JobRecord.model_validate_json(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                # A corrupt document must not hide the readable jobs.
                continue
            if record.state in _NON_TERMINAL_STATES:
                record = _reconcile_record(record)
                _atomic_write_json(path, record.model_dump_json(indent=2).encode("utf-8"))
                reconciled.append(record)
            self._records[record.job_id] = record
            if record.idempotency_key:
                self._by_key[record.idempotency_key] = record.job_id
        return reconciled


def _reconcile_record(record: JobRecord) -> JobRecord:
    """Return ``record`` reconciled to ``interrupted_unknown`` (never completed)."""

    return record.model_copy(
        update={
            "state": INTERRUPTED_UNKNOWN,
            "updated_at": datetime.now(UTC),
            "outcome": "unknown",
            "retryable": False,
        }
    )