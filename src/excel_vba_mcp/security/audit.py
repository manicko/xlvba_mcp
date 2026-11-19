"""Append-only JSONL audit trail (spec §10 "Аудит", §10 "MCP transport").

The audit log proves *what* happened and *that* it was approved — never the
secret values. It records the user/process, workbook identity (canonical path +
SHA-256), tool name, redacted parameters, approval decision, snapshot id, result
and HRESULT for every operation.

Hard requirements (research task_009 §4):

* **Never write to the MCP stdout stream.** Events go to a dedicated file handle
  opened in append mode; the STDIO JSON-RPC protocol is never corrupted (A01).
* **Never use the root ``logging`` StreamHandler** for audit output.
* **One JSON object per line (JSONL)** — append-only, tail-friendly, and a partial
  final line is the only corruption mode on a hard crash.
* **Serialized writes.** The STA worker thread and anyio worker threads may both
  emit events; a :class:`threading.Lock` guarantees lines never interleave.
* **Redaction is mandatory and lossy on purpose** (:func:`redact`): reserved
  secret keys become ``"<redacted>"``, paths are reduced to their basename, and
  macro arguments are logged as typed, length-capped summaries (no COM refs).

Audit path validation is the caller's responsibility: the file must live under an
allowlisted artifact root checked by ``PathPolicy`` (spec §10). :class:`AuditLogger`
only ensures the parent directory exists so the append handle can be opened; it
does not itself re-validate the path.
"""

from __future__ import annotations

import json
import threading
from collections.abc import Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

# Bump when the audit record schema changes incompatibly (research §4).
AUDIT_SCHEMA_VERSION = 1

# Case-insensitive parameter keys whose values must never be logged (spec §10).
_RESERVED_SECRET_KEYS: frozenset[str] = frozenset(
    {
        "password",
        "passwd",
        "pwd",
        "secret",
        "token",
        "connection_string",
        "conn_str",
        "api_key",
        "credential",
        "auth",
    }
)

_REDACTED = "<redacted>"

# Maximum number of macro arguments summarised (Excel allows 30 positional args).
_MAX_ARGS = 30
# Maximum characters kept from a single string argument's repr.
_MAX_ARG_STR = 80


def _utc_timestamp() -> str:
    """Current UTC time as ISO-8601 with a trailing ``Z``."""

    return datetime.now(UTC).isoformat().replace("+00:00", "Z")


def _basename(path: str) -> str:
    """Return the final path component for either Windows or POSIX separators."""

    trimmed = path.replace("\\", "/").rstrip("/")
    if not trimmed:
        return path
    return trimmed.rsplit("/", 1)[-1]


def _is_path_key(key_lower: str) -> bool:
    return (
        key_lower == "path"
        or key_lower.endswith("_path")
        or key_lower
        in {
            "filename",
            "file_name",
        }
    )


def _summarize_arg(value: Any) -> str:
    """Return a typed, length-capped summary of a single macro argument."""

    if value is None:
        return "null"
    if isinstance(value, bool):
        return f"bool({value})"
    if isinstance(value, (list, tuple)):
        return f"list[{len(value)}]"
    if isinstance(value, dict):
        return f"dict[{len(value)}]"
    if isinstance(value, (int, float)):
        return f"{type(value).__name__}({value})"
    if isinstance(value, str):
        text = value if len(value) <= _MAX_ARG_STR else value[:_MAX_ARG_STR] + "…"
        return f"str({text!r})"
    return f"{type(value).__name__}(<unserializable>)"


def _summarize_args(value: Any) -> dict[str, Any]:
    """Summarise a macro argument collection without leaking values or COM refs."""

    if isinstance(value, (list, tuple)):
        items = [_summarize_arg(item) for item in value[:_MAX_ARGS]]
        summary: dict[str, Any] = {"arg_count": len(value), "args": items}
        if len(value) > _MAX_ARGS:
            summary["truncated"] = len(value) - _MAX_ARGS
        return summary
    return {"arg_count": 0, "args": [], "note": "arguments were not a sequence"}


def _redact_value(key: str, value: Any) -> Any:
    key_lower = key.lower()
    if key_lower in _RESERVED_SECRET_KEYS:
        return _REDACTED
    if key_lower in {"args", "arguments"}:
        return _summarize_args(value)
    if _is_path_key(key_lower) and isinstance(value, str):
        return _basename(value)
    if isinstance(value, Mapping):
        return redact(value)
    if isinstance(value, (list, tuple)):
        return [_redact_value("", item) for item in value]
    return value


def redact(params: Mapping[str, Any]) -> dict[str, Any]:
    """Return a redacted, JSON-serializable copy of ``params``.

    * Reserved secret keys → ``"<redacted>"`` (case-insensitive).
    * Path-like string values → basename.
    * ``args``/``arguments`` → typed, length-capped summary.
    * Nested mappings/lists are scrubbed recursively.

    The function is pure and never mutates its input.
    """

    return {str(key): _redact_value(str(key), value) for key, value in params.items()}


class AuditLogger:
    """Thread-safe append-only JSONL audit writer.

    Args:
        path: Destination file. Created (with parents) if missing; opened in
            append mode so an existing log is extended, never truncated.
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        parent = path.parent
        if parent and not parent.exists():
            parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        # The handle is intentionally held open for the logger's lifetime (append
        # mode, closed by close()/__exit__); it must not be a scoped `with` here.
        self._handle = open(  # noqa: SIM115
            path, "a", encoding="utf-8", newline="\n"
        )

    @property
    def path(self) -> Path:
        """The destination file path."""

        return self._path

    def log(self, event: Mapping[str, Any]) -> None:
        """Write a single event as one JSON line (thread-safe, flushed).

        ``event`` is expected to be already shaped/redacted (see :meth:`record`).
        The write happens under a lock so concurrent writers never interleave.
        """

        line = json.dumps(event, ensure_ascii=False, default=str, sort_keys=False)
        with self._lock:
            self._handle.write(line + "\n")
            self._handle.flush()

    def record(
        self,
        *,
        tool: str,
        operation_class: str,
        params: Mapping[str, Any] | None = None,
        correlation_id: str | None = None,
        session_id: str | None = None,
        workbook_id: str | None = None,
        workbook: Mapping[str, Any] | None = None,
        approval: Mapping[str, Any] | None = None,
        snapshot_id: str | None = None,
        result: Mapping[str, Any] | None = None,
        user: str | None = None,
        process: str | None = None,
        excel_version: str | None = None,
    ) -> dict[str, Any]:
        """Assemble a schema-conformant event (redacting params) and log it.

        Returns the event dict that was written, which is convenient for callers
        that want to assert on the recorded shape.
        """

        workbook_payload = _normalize_workbook(workbook)
        event: dict[str, Any] = {
            "schema_version": AUDIT_SCHEMA_VERSION,
            "timestamp": _utc_timestamp(),
            "event_id": str(uuid4()),
            "correlation_id": correlation_id,
            "session_id": session_id,
            "workbook_id": workbook_id,
            "workbook": workbook_payload,
            "tool": tool,
            "operation_class": operation_class,
            "params": redact(params) if params is not None else {},
            "approval": dict(approval) if approval is not None else None,
            "snapshot_id": snapshot_id,
            "result": dict(result) if result is not None else None,
            "user": user,
            "process": process,
            "excel_version": excel_version,
        }
        self.log(event)
        return event

    def flush(self) -> None:
        """Flush the underlying handle."""

        with self._lock:
            self._handle.flush()

    def close(self) -> None:
        """Flush and close the file handle (safe to call more than once)."""

        with self._lock:
            if not self._handle.closed:
                self._handle.flush()
                self._handle.close()

    def __enter__(self) -> AuditLogger:
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.close()


def _normalize_workbook(workbook: Mapping[str, Any] | None) -> dict[str, Any] | None:
    if workbook is None:
        return None
    normalized = dict(workbook)
    canonical = normalized.get("canonical_path")
    if not normalized.get("file_name") and isinstance(canonical, str):
        normalized["file_name"] = _basename(canonical)
    return normalized
