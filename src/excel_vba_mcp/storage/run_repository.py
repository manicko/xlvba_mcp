"""Filesystem persistence for declarative test runs (task_035; spec §9.1, §9.2).

:class:`RunRepository` stores the :class:`~excel_vba_mcp.domain.models.TestRunResult`
of every executed suite as one JSON document per run under an allowlisted root.
It is pure filesystem — it never touches Excel / COM and holds no proxies.

Guarantees:

* **Atomic.** The document is written to a temp file in the destination
  directory, flushed / ``fsync``-ed and then moved into place with
  ``os.replace``; a crash or error never leaves a half-written run behind.
* **Validated ids.** A run id becomes a file name, so separators, ``..``
  traversal, empty / dot names and control characters are rejected with
  :class:`~excel_vba_mcp.domain.errors.PathNotAllowedError` (the spec §11
  persistence code ``PATH_NOT_ALLOWED``).
* **Truthful reads.** ``get_run`` raises :class:`TestRunNotFoundError` for an
  unknown or unreadable id; it never fabricates a result.
"""

from __future__ import annotations

import contextlib
import json
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, PathNotAllowedError
from excel_vba_mcp.domain.models import TestRunResult
from excel_vba_mcp.domain.statuses import TestRunStatus
from excel_vba_mcp.security.path_policy import PathDenyReason

__all__ = ["RunRepository", "RunSummary", "TestRunNotFoundError"]

#: File suffix of a persisted run document.
_RUN_SUFFIX = ".json"

#: Upper bound for a run id (basename only; no directories).
_MAX_RUN_ID_LENGTH = 160

#: Run ids are server-generated (``run-<hex>``); keep the accepted shape tight so
#: a caller cannot smuggle a path into the repository root.
_RUN_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,159}$")


class TestRunNotFoundError(ExcelVbaMcpError):
    """No persisted test run exists for the requested id.

    The spec §11 error vocabulary is closed, so the persistence-failure code is
    reused: ``PATH_NOT_ALLOWED`` ("a persistence operation did not yield a
    result, nothing was mutated") rather than minting a new code.
    """

    default_code = ErrorCode.PATH_NOT_ALLOWED
    default_message = "The test run was not found."


class RunSummary(BaseModel):
    """Compact, bounded view of one persisted run (for ``list_runs``).

    ``list_runs`` deliberately does not return every full
    :class:`TestRunResult`: a long-lived server accumulates runs and a listing
    must stay cheap. Fetch the full document with
    :meth:`RunRepository.get_run` when the per-case detail is needed.
    """

    model_config = ConfigDict(extra="forbid")

    test_run_id: str
    suite: str
    status: TestRunStatus
    workbook_id: str | None = None
    baseline_snapshot_id: str | None = None
    case_count: int = 0
    started_at: datetime | None = None
    finished_at: datetime | None = None
    duration_ms: int | None = None
    stored_path: str = Field(description="Absolute path of the persisted run document.")


def _validate_run_id(run_id: str) -> str:
    """Return a safe run id or raise ``PATH_NOT_ALLOWED``."""

    if not isinstance(run_id, str) or not run_id.strip():
        raise PathNotAllowedError(
            "The test run id is empty.",
            context={"reason": PathDenyReason.EMPTY.value},
        )
    candidate = run_id.strip()
    if len(candidate) > _MAX_RUN_ID_LENGTH:
        raise PathNotAllowedError(
            "The test run id is too long.",
            context={"reason": "invalid_run_id"},
        )
    if candidate in {".", ".."} or "/" in candidate or "\\" in candidate:
        raise PathNotAllowedError(
            "The test run id must not contain a path separator.",
            context={"reason": PathDenyReason.TRAVERSAL.value},
        )
    if _RUN_ID_PATTERN.match(candidate) is None:
        raise PathNotAllowedError(
            "The test run id contains unsupported characters.",
            context={"reason": "invalid_run_id"},
        )
    return candidate


def _atomic_write_json(target: Path, payload: bytes) -> None:
    """Write ``payload`` to ``target`` atomically (temp file + ``os.replace``)."""

    target.parent.mkdir(parents=True, exist_ok=True)
    descriptor = -1
    temp_path: Path | None = None
    try:
        descriptor, temp_name = tempfile.mkstemp(
            prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent)
        )
        temp_path = Path(temp_name)
        with os.fdopen(descriptor, "wb") as handle:
            descriptor = -1  # ownership transferred to the file object
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_path, target)
        temp_path = None
    finally:
        if descriptor >= 0:
            with contextlib.suppress(OSError):
                os.close(descriptor)
        if temp_path is not None:
            with contextlib.suppress(OSError):
                temp_path.unlink(missing_ok=True)


class RunRepository:
    """Append-only-ish store of executed test runs (one JSON file per run).

    Args:
        root: Directory that holds one ``<run_id>.json`` document per run.
            Created on first write; it never contains repository sources.
    """

    def __init__(self, root: Path) -> None:
        self._root = Path(root)
        self._root.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    @property
    def root(self) -> Path:
        """The run repository root directory."""

        return self._root

    def path_for(self, run_id: str) -> Path:
        """Return the validated absolute path of the ``run_id`` document."""

        return self._root / f"{_validate_run_id(run_id)}{_RUN_SUFFIX}"

    def save_run(self, run: TestRunResult) -> Path:
        """Atomically persist ``run`` and return the written document path.

        Args:
            run: The completed run result to persist.

        Returns:
            The absolute path of the written JSON document.

        Raises:
            PathNotAllowedError: If the run id is not a safe file name.
            OSError: If the atomic write / replace fails (nothing is left behind).
        """

        run_id = _validate_run_id(run.test_run_id)
        target = self._root / f"{run_id}{_RUN_SUFFIX}"
        payload = run.model_dump_json(indent=2).encode("utf-8")
        _atomic_write_json(target, payload)
        return target

    def get_run(self, run_id: str) -> TestRunResult:
        """Return the persisted :class:`TestRunResult` for ``run_id``.

        Args:
            run_id: The id returned in :attr:`TestRunResult.test_run_id`.

        Raises:
            PathNotAllowedError: If the run id is not a safe file name.
            TestRunNotFoundError: If no readable run document exists for the id.
        """

        target = self.path_for(run_id)
        if not target.is_file():
            raise TestRunNotFoundError(
                f"No persisted test run exists for id {run_id!r}.",
                context={"test_run_id": run_id},
            )
        try:
            payload: Any = json.loads(target.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise TestRunNotFoundError(
                f"The persisted test run {run_id!r} could not be read.",
                context={"test_run_id": run_id, "error": exc.__class__.__name__},
            ) from exc
        return TestRunResult.model_validate(payload)

    def list_runs(self) -> list[RunSummary]:
        """Return compact summaries of every persisted run (newest first).

        A document that cannot be parsed is skipped rather than failing the whole
        listing (the repository is a convenience view; a corrupt file must not
        hide the readable runs).
        """

        summaries: list[tuple[datetime, RunSummary]] = []
        for path in sorted(self._root.glob(f"*{_RUN_SUFFIX}")):
            if not path.is_file() or path.name.startswith("."):
                continue
            try:
                run = TestRunResult.model_validate_json(path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                continue
            summaries.append(
                (
                    run.updated_at,
                    RunSummary(
                        test_run_id=run.test_run_id,
                        suite=run.suite,
                        status=run.status,
                        workbook_id=run.workbook_id,
                        baseline_snapshot_id=run.baseline_snapshot_id,
                        case_count=len(run.cases),
                        started_at=run.started_at,
                        finished_at=run.finished_at,
                        duration_ms=run.duration_ms,
                        stored_path=str(path),
                    ),
                )
            )
        summaries.sort(key=lambda item: item[0], reverse=True)
        return [summary for _, summary in summaries]
