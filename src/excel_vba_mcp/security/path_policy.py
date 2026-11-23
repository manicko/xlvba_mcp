"""Canonical path and workbook allowlist policy (spec §7.1, §10).

Centralises the filesystem boundary for workbook and artifact access. Every path
that reaches Excel COM must pass through :meth:`PathPolicy.validate_path` first;
arbitrary filesystem paths are never handed to Excel directly (spec §10
"MCP transport" / §7.1).

Policy rules (spec §10 "Доступ к файлам"):

* Only paths inside one of the configured allowlisted roots are accepted.
* Paths are canonicalised with :meth:`pathlib.Path.resolve` and cross-checked
  with :func:`os.path.realpath`; the resolved target must stay inside a resolved
  root. This blocks symlink / directory-junction / reparse-point escapes.
* ``..`` traversal that escapes a root is rejected.
* UNC paths (``\\\\server\\share``) are denied unless ``allow_unc=True``.
* Device paths (``\\\\?\\`` and ``\\\\.\\``) are denied unless explicitly allowed.
* Workbook files must use a permitted Excel extension; snapshot/artifact roots
  may additionally hold ``.json`` / ``.sha256`` files.

Rejections raise :class:`~excel_vba_mcp.domain.errors.PathNotAllowedError`
(code ``PATH_NOT_ALLOWED``) with a safe ``reason`` category and no secret
leakage. This module performs no filesystem writes.
"""

from __future__ import annotations

import os
from collections.abc import Mapping
from enum import StrEnum
from pathlib import Path
from typing import Any

from excel_vba_mcp.domain.errors import PathNotAllowedError
from excel_vba_mcp.domain.statuses import OperationClass

# Excel workbook / template / add-in formats accepted for workbook operations.
# Kept explicit and documented on purpose: never silently allow arbitrary files.
WORKBOOK_EXTENSIONS: frozenset[str] = frozenset(
    {
        ".xlsm",  # macro-enabled workbook
        ".xlsb",  # binary workbook (macro-capable)
        ".xlsa",  # Excel add-in
        ".xlam",  # macro-enabled add-in
        ".xltm",  # macro-enabled template
        ".xltx",  # template
        ".xlsx",  # standard workbook
        ".xls",  # legacy workbook
    }
)

# Snapshot sidecar artifacts (manifest, checksum) allowed only under artifact roots.
SNAPSHOT_ARTIFACT_EXTENSIONS: frozenset[str] = frozenset({".json", ".sha256"})

# Logical roots whose files may be snapshot/artifact sidecars in addition to
# workbook files.
ARTIFACT_ROOT_NAMES: frozenset[str] = frozenset({"snapshots", "artifacts"})


class PathDenyReason(StrEnum):
    """Reason categories attached to a :class:`PathNotAllowedError`."""

    EMPTY = "empty_path"
    DEVICE_DENIED = "device_path_denied"
    UNC_DENIED = "unc_denied"
    TRAVERSAL = "traversal"
    OUTSIDE_ROOTS = "outside_roots"
    SYMLINK_ESCAPE = "symlink_escape"
    ROOT_MISMATCH = "root_mismatch"
    EXTENSION_NOT_ALLOWED = "extension_not_allowed"


class PathPolicy:
    """Allowlist + canonicalisation gate for workbook and artifact paths.

    Args:
        roots: Mapping of logical root name (e.g. ``"projects"``) to an absolute
            directory. Roots are canonicalised on construction; no repo paths are
            hardcoded by this class.
        allow_unc: When ``False`` (default) UNC and device paths are rejected.
    """

    def __init__(
        self,
        roots: Mapping[str, str | Path],
        *,
        allow_unc: bool = False,
    ) -> None:
        if not roots:
            raise ValueError("PathPolicy requires at least one allowlisted root.")
        self._allow_unc = allow_unc
        resolved_roots: dict[str, Path] = {}
        for name, root in roots.items():
            if not isinstance(name, str) or not name.strip():
                raise ValueError("Root names must be non-empty strings.")
            resolved_roots[name] = _canonicalize(root)
        self._roots: dict[str, Path] = resolved_roots

    @property
    def roots(self) -> Mapping[str, Path]:
        """Read-only view of the canonicalised logical roots."""

        return dict(self._roots)

    @property
    def allow_unc(self) -> bool:
        """Whether UNC/device paths are permitted."""

        return self._allow_unc

    def validate_path(
        self,
        raw: str | Path,
        operation: OperationClass | str,
        *,
        expected_root: str | None = None,
    ) -> Path:
        """Validate ``raw`` against the allowlist and return the canonical path.

        Args:
            raw: Caller-supplied path (string or :class:`~pathlib.Path`).
            operation: Operation classification used for diagnostics and policy.
            expected_root: When given, the path must live inside this logical root.

        Returns:
            The canonical, symlink-resolved absolute :class:`~pathlib.Path`,
            suitable for downstream workbook locking / opening.

        Raises:
            PathNotAllowedError: For every rejection, with a ``reason`` category
                in ``error.context["reason"]`` and code ``PATH_NOT_ALLOWED``. A
                rejection because the path matches no root additionally lists the
                configured roots in ``error.context["allowed_roots"]``.
        """

        op = _coerce_operation(operation)

        raw_str = str(raw)
        probe = raw_str.strip()
        if not probe:
            raise self._reject(PathDenyReason.EMPTY, op, raw_str, expected_root)

        # Device paths (\\?\ and \\.\) bypass normal normalisation and can address
        # raw volumes; reject them unless explicitly allowed.
        if _is_device_path(probe) and not self._allow_unc:
            raise self._reject(PathDenyReason.DEVICE_DENIED, op, probe, expected_root)

        # UNC paths are denied by default (spec §10).
        if _is_unc_path(probe) and not self._allow_unc:
            raise self._reject(PathDenyReason.UNC_DENIED, op, probe, expected_root)

        candidate = Path(probe)
        lexical = Path(os.path.abspath(probe))
        resolved = _resolve_candidate(candidate)

        matches = [name for name, root in self._roots.items() if _is_within(resolved, root)]
        if not matches:
            # Distinguish a symlink/reparse escape (lexically inside a root but
            # resolving outside) from plain traversal / outside-root paths.
            if any(_is_within(lexical, root) for root in self._roots.values()):
                reason = PathDenyReason.SYMLINK_ESCAPE
            elif _has_parent_segment(probe):
                reason = PathDenyReason.TRAVERSAL
            else:
                reason = PathDenyReason.OUTSIDE_ROOTS
            # List the currently-allowed roots so a caller can self-diagnose a
            # missing allowlist entry (F5). Additive context only: the error
            # ``code`` (PATH_NOT_ALLOWED) and ``reason`` are unchanged.
            raise self._reject(
                reason,
                op,
                probe,
                expected_root,
                extra_context={"allowed_roots": [os.fspath(root) for root in self._roots.values()]},
            )

        if expected_root is not None and expected_root not in matches:
            raise self._reject(PathDenyReason.ROOT_MISMATCH, op, probe, expected_root)

        matched_root = (
            expected_root if expected_root is not None else _most_specific(matches, self._roots)
        )
        if candidate.suffix.lower() not in self._allowed_extensions(matched_root):
            raise self._reject(PathDenyReason.EXTENSION_NOT_ALLOWED, op, probe, expected_root)

        return resolved

    def _allowed_extensions(self, root_name: str) -> frozenset[str]:
        if root_name in ARTIFACT_ROOT_NAMES:
            return WORKBOOK_EXTENSIONS | SNAPSHOT_ARTIFACT_EXTENSIONS
        return WORKBOOK_EXTENSIONS

    def _reject(
        self,
        reason: PathDenyReason,
        operation: OperationClass,
        probe: str,
        expected_root: str | None,
        *,
        extra_context: Mapping[str, Any] | None = None,
    ) -> PathNotAllowedError:
        context: dict[str, Any] = {
            "reason": reason.value,
            "operation": operation.value,
            "name": _safe_name(probe),
        }
        if expected_root is not None:
            context["expected_root"] = expected_root
        if extra_context:
            context.update(extra_context)
        return PathNotAllowedError(
            f"Path not allowed ({reason.value}).",
            context=context,
        )


def _coerce_operation(operation: OperationClass | str) -> OperationClass:
    if isinstance(operation, OperationClass):
        return operation
    try:
        return OperationClass(operation)
    except ValueError as exc:  # programming error, not a path rejection
        raise ValueError(f"Unknown operation class: {operation!r}") from exc


def _is_device_path(path: str) -> bool:
    lowered = path.lower()
    return (
        lowered.startswith("\\\\.\\")
        or lowered.startswith("\\\\?\\")
        or lowered.startswith("//./")
        or lowered.startswith("//?/")
    )


def _is_unc_path(path: str) -> bool:
    return path.startswith("\\\\") or path.startswith("//")


def _has_parent_segment(path: str) -> bool:
    try:
        return ".." in Path(path).parts
    except (OSError, ValueError):  # pragma: no cover - defensive
        return ".." in path


def _canonicalize(path: str | Path) -> Path:
    try:
        return Path(os.path.realpath(path))
    except OSError:  # pragma: no cover - realpath rarely fails
        return Path(os.path.abspath(path))


def _resolve_candidate(candidate: Path) -> Path:
    try:
        resolved = candidate.resolve()
    except OSError:  # pragma: no cover - resolve rarely fails
        resolved = Path(os.path.abspath(candidate))
    real = Path(os.path.realpath(candidate))
    # Cross-check with realpath; prefer it when the two disagree (reparse points).
    return real if os.fspath(real) != os.fspath(resolved) else resolved


def _is_within(child: str | Path, parent: str | Path) -> bool:
    try:
        common = os.path.commonpath([os.fspath(child), os.fspath(parent)])
    except ValueError:
        # Different drives, or mixing absolute and relative paths.
        return False
    return _normalize(common) == _normalize(parent)


def _normalize(path: str | Path) -> str:
    return os.path.normcase(os.fspath(path)).rstrip("\\/")


def _most_specific(matches: list[str], roots: Mapping[str, Path]) -> str:
    return max(matches, key=lambda name: len(os.fspath(roots[name])))


def _safe_name(probe: str) -> str:
    """Return a non-sensitive basename for diagnostics (no directory / secret)."""

    trimmed = probe.rstrip("\\/")
    name = os.path.basename(trimmed) if trimmed else ""
    return name or "<empty>"
