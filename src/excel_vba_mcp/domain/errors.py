"""Typed error hierarchy with stable codes and HRESULT preservation (spec §11).

Every error carries a stable ``code`` (one of :class:`ErrorCode`) and, when a
COM/HRESULT failure is involved, both the raw ``hresult`` int and its
``0x%08X`` hex form. The original message and optional context are preserved so
that a caller can build a diagnostic DTO without losing information.

This module does not import ``pywin32``; HRESULT values are treated as plain
integers so the domain layer stays free of COM dependencies. The COM worker
(task_010) constructs these errors from real ``com_error`` values.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Any


class ErrorCode(StrEnum):
    """Stable, machine-readable error codes from spec §11 (Stage 1).

    The Stage 1 codes below are closed (spec §11). The Phase 2 codes appended
    afterwards (task_040) extend the vocabulary additively for the Ranges &
    Cells and PivotTables surface; none of the Stage 1 codes are removed or
    changed.
    """

    ACCESS_VBOM_DISABLED = "ACCESS_VBOM_DISABLED"
    VBA_PROJECT_LOCKED = "VBA_PROJECT_LOCKED"
    MISSING_REFERENCE = "MISSING_REFERENCE"
    COM_BUSY = "COM_BUSY"
    RPC_E_CALL_REJECTED = "RPC_E_CALL_REJECTED"
    COMPILE_ERROR = "COMPILE_ERROR"
    COMPILE_INDETERMINATE = "COMPILE_INDETERMINATE"
    MACRO_BLOCKED = "MACRO_BLOCKED"
    MACRO_NOT_FOUND = "MACRO_NOT_FOUND"
    MACRO_RUNTIME_ERROR = "MACRO_RUNTIME_ERROR"
    MACRO_TIMEOUT = "MACRO_TIMEOUT"
    EXCEL_HUNG = "EXCEL_HUNG"
    SNAPSHOT_FAILED = "SNAPSHOT_FAILED"
    WORKBOOK_READ_ONLY = "WORKBOOK_READ_ONLY"
    PATH_NOT_ALLOWED = "PATH_NOT_ALLOWED"
    # --- Phase 2 additions (task_040; additive, do not remove Stage 1 codes) ---
    INVALID_RANGE = "INVALID_RANGE"
    # WP-Range (additive): a syntactically valid Range whose cell count exceeds
    # ``MAX_RANGE_CELLS`` (100_000). Distinct from ``INVALID_RANGE`` (malformed /
    # ambiguous / wrong-shaped ranges) so a caller can tell "too large, page it"
    # from "bad address, fix it". Not retryable.
    RANGE_TOO_LARGE = "RANGE_TOO_LARGE"
    PROTECTED_TARGET = "PROTECTED_TARGET"
    UNSUPPORTED_PIVOT_SOURCE = "UNSUPPORTED_PIVOT_SOURCE"
    STALE_TARGET = "STALE_TARGET"
    AMBIGUOUS_FILTER = "AMBIGUOUS_FILTER"
    REFRESH_FAILED = "REFRESH_FAILED"
    MODAL_UI_BLOCKED = "MODAL_UI_BLOCKED"
    LISTOBJECT_INTERSECTION = "LISTOBJECT_INTERSECTION"
    APPROVAL_DENIED = "APPROVAL_DENIED"
    # Task_b5 (fix_005): macro background-safety result codes (additive; result
    # codes only — no exception classes, no ``_ERROR_BY_CODE`` entries).
    BACKGROUND_SAFETY_NOT_VERIFIED = "BACKGROUND_SAFETY_NOT_VERIFIED"
    INTERACTIVE_UI_NOT_SUPPORTED = "INTERACTIVE_UI_NOT_SUPPORTED"
    # Task FE-22 (finding F-5): caller-error codes (additive). These distinguish a
    # caller-supplied failure (unknown/lost ``workbook_id``) and an optimistic-
    # concurrency failure (stale ``expected_source_hash``) from genuine COM
    # contention, which the generic ``COM_BUSY`` fallback previously masked.
    WORKBOOK_NOT_FOUND = "WORKBOOK_NOT_FOUND"
    REVISION_CONFLICT = "REVISION_CONFLICT"
    # Block B7a (finding F6): a workbook object (shape / control) addressed by
    # name does not exist. Additive; distinct from ``WORKBOOK_NOT_FOUND`` /
    # ``MACRO_NOT_FOUND`` and from the generic ``COM_BUSY`` fallback so callers
    # can tell "no such object, do not retry" from "Excel busy, retry".
    NOT_FOUND = "NOT_FOUND"
    # U6-B (plan_05 §5.2): additive envelope codes. ``INTERNAL_ERROR`` is the
    # guaranteed structured fallback for an unhandled non-domain exception at the
    # ``call_tool`` boundary (so no bare "Error executing tool" leak reaches a
    # client); ``CONFIRMATION_REQUIRED`` is a caller-argument gate raised before
    # a destructive action when a required ``confirm`` flag is missing/false. Neither
    # is retryable.
    INTERNAL_ERROR = "INTERNAL_ERROR"
    CONFIRMATION_REQUIRED = "CONFIRMATION_REQUIRED"
    # WP1 (additive): a defined name (or table column header) collides with an
    # existing one in the same scope, or a same identifier already exists at
    # another scope (a cross-scope conflict). Distinct from the generic
    # ``COM_BUSY`` fallback so a caller can tell "choose a unique name" from
    # "Excel busy, retry". Not retryable; category ``validation``.
    NAME_CONFLICT = "NAME_CONFLICT"


# Well-known HRESULT values that map to domain codes (hex shown for readability).
# fmt: off
_HRESULT_TO_CODE: dict[int, ErrorCode] = {
    0x80010001: ErrorCode.RPC_E_CALL_REJECTED,        # RPC_E_CALL_REJECTED
    0x8001010A: ErrorCode.COM_BUSY,                   # RPC_E_SERVERCALL_RETRYLATER
    0x800A03EC: ErrorCode.MACRO_RUNTIME_ERROR,        # VBA runtime error (Excel)
    0x800A03ED: ErrorCode.MACRO_NOT_FOUND,            # unable to run / macro not found
    0x80004005: ErrorCode.MACRO_RUNTIME_ERROR,        # E_FAIL (generic macro failure)
    0x80040154: ErrorCode.MISSING_REFERENCE,          # REGDB_E_CLASSNOTREG (broken ref)
    0x8002801C: ErrorCode.ACCESS_VBOM_DISABLED,       # VBIDE access denied (best effort)
}
# fmt: on


def format_hresult(hresult: int) -> str:
    """Return the canonical ``0x%08X`` representation of an HRESULT."""

    return f"0x{hresult & 0xFFFFFFFF:08X}"


def code_from_hresult(hresult: int, default: ErrorCode | None = None) -> ErrorCode | None:
    """Map a raw HRESULT to a domain :class:`ErrorCode`.

    Args:
        hresult: The raw COM HRESULT as an integer.
        default: Code returned when the HRESULT is not in the known map and the
            caller wants a fallback; ``None`` means "unknown".

    Returns:
        The mapped code, or ``default`` when the HRESULT is unrecognized.
    """

    return _HRESULT_TO_CODE.get(hresult & 0xFFFFFFFF, default)


# ---------------------------------------------------------------------------
# Structured error envelope (U2; plan_03 §7): category / retry / next actions.
#
# These are PURE functions of the error code (and, for ``actions_for``, the
# context): no COM, no schema bump, no new ``ErrorCode`` members. ``to_dict``
# calls them so every error carries ``errorCategory`` / ``isRetryable`` /
# ``suggestedNextActions`` in addition to the original five keys.
#
# ``isRetryable`` is conservative (D-U10): only a transient COM busy/rejected
# result is retryable. ``STALE_TARGET`` / ``REVISION_CONFLICT`` are NOT
# retryable (a blind replay could re-apply a partially-applied mutation) and a
# job ``interrupted_unknown`` is never retryable.
# ---------------------------------------------------------------------------

# errorCategory taxonomy (plan_03 §7.2). Every ``ErrorCode`` member is mapped so
# ``category_for`` never needs a fallback for a real code.
_CATEGORY_BY_CODE: dict[ErrorCode, str] = {
    # permission_denied — policy/approval/trust denial; not fixable by retry.
    ErrorCode.ACCESS_VBOM_DISABLED: "permission_denied",
    ErrorCode.VBA_PROJECT_LOCKED: "permission_denied",
    ErrorCode.MACRO_BLOCKED: "permission_denied",
    ErrorCode.PATH_NOT_ALLOWED: "permission_denied",
    ErrorCode.APPROVAL_DENIED: "permission_denied",
    # missing_prerequisite — a required precondition is absent.
    ErrorCode.SNAPSHOT_FAILED: "missing_prerequisite",
    ErrorCode.WORKBOOK_READ_ONLY: "missing_prerequisite",
    ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED: "missing_prerequisite",
    ErrorCode.REVISION_CONFLICT: "missing_prerequisite",
    ErrorCode.STALE_TARGET: "missing_prerequisite",
    ErrorCode.MISSING_REFERENCE: "missing_prerequisite",
    # unsupported_feature — capability absent on this build/config.
    ErrorCode.UNSUPPORTED_PIVOT_SOURCE: "unsupported_feature",
    ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED: "unsupported_feature",
    ErrorCode.COMPILE_INDETERMINATE: "unsupported_feature",
    ErrorCode.COMPILE_ERROR: "unsupported_feature",
    ErrorCode.PROTECTED_TARGET: "unsupported_feature",
    ErrorCode.LISTOBJECT_INTERSECTION: "unsupported_feature",
    ErrorCode.AMBIGUOUS_FILTER: "unsupported_feature",
    ErrorCode.MACRO_RUNTIME_ERROR: "unsupported_feature",
    ErrorCode.REFRESH_FAILED: "unsupported_feature",
    # transient_com_failure — Excel/VBE busy; retry after backoff.
    ErrorCode.COM_BUSY: "transient_com_failure",
    ErrorCode.RPC_E_CALL_REJECTED: "transient_com_failure",
    # deadline — soft deadline exceeded; op left running.
    ErrorCode.MACRO_TIMEOUT: "deadline",
    # suspected_stall — deadline + possible modal; agent-directed recovery.
    ErrorCode.EXCEL_HUNG: "suspected_stall",
    ErrorCode.MODAL_UI_BLOCKED: "suspected_stall",
    # not_found — caller addressed a missing identity.
    ErrorCode.WORKBOOK_NOT_FOUND: "not_found",
    ErrorCode.MACRO_NOT_FOUND: "not_found",
    ErrorCode.NOT_FOUND: "not_found",
    # validation — caller argument/schema error (also the pydantic path).
    ErrorCode.INVALID_RANGE: "validation",
    ErrorCode.RANGE_TOO_LARGE: "validation",
    ErrorCode.CONFIRMATION_REQUIRED: "validation",
    ErrorCode.NAME_CONFLICT: "validation",
    # internal — unhandled non-domain failure at the tool boundary (U6-B).
    ErrorCode.INTERNAL_ERROR: "internal",
}

# Only a transient COM busy/rejected result is retryable (plan_03 §7.3, D-U10).
_RETRYABLE_CODES: frozenset[ErrorCode] = frozenset(
    {ErrorCode.COM_BUSY, ErrorCode.RPC_E_CALL_REJECTED}
)

# Job-state string keys (not ``ErrorCode`` members) that can appear in a job
# result and must still classify truthfully (plan_03 §7.2/§7.3).
_JOB_CATEGORY: dict[str, str] = {
    "expired": "deadline",
    "interrupted_unknown": "interrupted_unknown",
}
_JOB_ACTIONS: dict[str, list[str]] = {
    "expired": [
        "excel_recover if still stalled",
        "never auto-replay; inspect macro_get_last_result",
    ],
    "interrupted_unknown": ["never resubmit; report status"],
}

# suggestedNextActions per code (plan_03 §7.3). Static, conservative text; the
# ``context`` is accepted for signature symmetry / future context-aware hints.
_ACTIONS_BY_CODE: dict[ErrorCode, list[str]] = {
    ErrorCode.COM_BUSY: ["wait retry_after_s (if present) then retry the same read"],
    ErrorCode.RPC_E_CALL_REJECTED: ["wait retry_after_s (if present) then retry the same read"],
    ErrorCode.SNAPSHOT_FAILED: ["snapshot_create(workbook_id)", "retry mutation"],
    ErrorCode.WORKBOOK_READ_ONLY: ["re-open with read_only=false (mode=execute)"],
    ErrorCode.APPROVAL_DENIED: ["open in full mode or supply grant/allowlist entry", "retry"],
    ErrorCode.PATH_NOT_ALLOWED: ["use a path under details.allowed_roots"],
    ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED: [
        "macro_classify_background(...background_safe=true) after review",
        "retry macro_run",
    ],
    ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED: ["remove interactive-UI tokens (out of scope)"],
    ErrorCode.UNSUPPORTED_PIVOT_SOURCE: ["use a supported source (database)"],
    ErrorCode.COMPILE_INDETERMINATE: ["accept indeterminate compile"],
    ErrorCode.STALE_TARGET: [
        "re-read source/hash via vba_read_component",
        "the prior mutation MAY have partially applied — re-read state; do NOT blind-replay",
    ],
    ErrorCode.REVISION_CONFLICT: [
        "re-read source/hash via vba_read_component",
        "the prior mutation MAY have partially applied — re-read state; do NOT blind-replay",
    ],
    ErrorCode.EXCEL_HUNG: [
        "excel_recover",
        "re-open workbook",
        "do not replay the interrupted mutation",
    ],
    ErrorCode.MODAL_UI_BLOCKED: [
        "excel_recover",
        "re-open workbook",
        "do not replay the interrupted mutation",
    ],
    ErrorCode.MACRO_TIMEOUT: [
        "excel_recover if still stalled",
        "never auto-replay; inspect macro_get_last_result",
    ],
    ErrorCode.WORKBOOK_NOT_FOUND: ["verify workbook_id / object name / procedure"],
    ErrorCode.NOT_FOUND: ["verify workbook_id / object name / procedure"],
    ErrorCode.MACRO_NOT_FOUND: ["verify workbook_id / object name / procedure"],
    ErrorCode.INVALID_RANGE: ["fix arguments to match input schema"],
    ErrorCode.RANGE_TOO_LARGE: [
        "read/write in smaller ranges (offset/limit)",
        "split the target into blocks of at most 100000 cells",
    ],
    ErrorCode.ACCESS_VBOM_DISABLED: [
        "enable Trust access to the VBA project object model in Excel Trust Center"
    ],
    ErrorCode.VBA_PROJECT_LOCKED: ["unlock the VBA project (protection is never bypassed)"],
    ErrorCode.MISSING_REFERENCE: [
        "repair the broken VBA reference (VBA Editor > Tools > References)"
    ],
    ErrorCode.COMPILE_ERROR: ["fix the VBA compile error reported by vba_get_compile_diagnostics"],
    ErrorCode.MACRO_RUNTIME_ERROR: ["inspect macro_get_last_result and fix the macro error"],
    ErrorCode.REFRESH_FAILED: ["verify the pivot source is database; retry the refresh"],
    ErrorCode.PROTECTED_TARGET: ["use a writable target (the component/range is protected)"],
    ErrorCode.LISTOBJECT_INTERSECTION: ["use the table tools for ListObject targets"],
    ErrorCode.AMBIGUOUS_FILTER: ["supply a filter that resolves to exactly one target"],
    ErrorCode.MACRO_BLOCKED: ["open in execute mode / adjust macro policy"],
    ErrorCode.INTERNAL_ERROR: [
        "inspect server logs; the failure was not a domain error"
    ],
    ErrorCode.CONFIRMATION_REQUIRED: [
        "re-call with confirm=true after reviewing what is discarded"
    ],
    ErrorCode.NAME_CONFLICT: ["choose a unique name or delete the existing one"],
}

# Fallback category/actions for an unrecognized string key (never a real code).
_DEFAULT_CATEGORY = "validation"
_DEFAULT_ACTIONS = ["inspect the error details and correct the request"]


def _code_key(code: ErrorCode | str) -> str:
    """Normalize a code (``ErrorCode`` or job-state string) to its string key."""

    return code.value if isinstance(code, ErrorCode) else str(code)


def category_for(code: ErrorCode | str) -> str:
    """Return the ``errorCategory`` for ``code`` (plan_03 §7.2).

    Pure and total: every :class:`ErrorCode` member is mapped; the job-state
    strings ``expired`` / ``interrupted_unknown`` classify as ``deadline`` /
    ``interrupted_unknown``; anything unrecognized falls back to ``validation``.
    """

    key = _code_key(code)
    if key in _JOB_CATEGORY:
        return _JOB_CATEGORY[key]
    try:
        member = ErrorCode(key)
    except ValueError:
        return _DEFAULT_CATEGORY
    return _CATEGORY_BY_CODE.get(member, _DEFAULT_CATEGORY)


def retryable_for(code: ErrorCode | str) -> bool:
    """Return whether ``code`` is conservatively retryable (plan_03 §7.3, D-U10).

    Only ``COM_BUSY`` / ``RPC_E_CALL_REJECTED`` are retryable; every other code
    (including ``STALE_TARGET`` / ``REVISION_CONFLICT`` and any job state) is not.
    """

    key = _code_key(code)
    if key in _JOB_CATEGORY:
        return False
    try:
        member = ErrorCode(key)
    except ValueError:
        return False
    return member in _RETRYABLE_CODES


def actions_for(code: ErrorCode | str, context: dict[str, Any] | None = None) -> list[str]:
    """Return ``suggestedNextActions`` for ``code`` (plan_03 §7.3).

    Pure: the text is static per code (the ``context`` is accepted for signature
    symmetry and future context-aware hints, and never mutated).
    """

    del context  # accepted for signature symmetry; actions are static per code
    key = _code_key(code)
    if key in _JOB_ACTIONS:
        return list(_JOB_ACTIONS[key])
    try:
        member = ErrorCode(key)
    except ValueError:
        return list(_DEFAULT_ACTIONS)
    return list(_ACTIONS_BY_CODE.get(member, _DEFAULT_ACTIONS))


class ExcelVbaMcpError(Exception):
    """Base error carrying a stable ``code`` and optional HRESULT context.

    Subclasses set :attr:`default_code`; callers may still override ``code`` when
    construction requires it. The constructor signature is uniform across the
    hierarchy so the error can be rebuilt from a raw HRESULT without special
    casing.
    """

    default_code: ErrorCode = ErrorCode.COM_BUSY
    default_message: str = "An Excel VBA MCP error occurred."

    def __init__(
        self,
        message: str | None = None,
        *,
        code: ErrorCode | None = None,
        hresult: int | None = None,
        context: dict[str, Any] | None = None,
    ) -> None:
        resolved_message = message if message is not None else self.default_message
        super().__init__(resolved_message)
        self.code = code if code is not None else self.default_code
        self.message = resolved_message
        self.hresult = hresult
        self.context: dict[str, Any] = dict(context) if context else {}

    @property
    def hresult_hex(self) -> str | None:
        """Hex (``0x%08X``) form of :attr:`hresult`, or ``None`` when absent."""

        return format_hresult(self.hresult) if self.hresult is not None else None

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serializable representation (no COM objects).

        Final envelope (U6-B; plan_05 §5.2): the keys are emitted in the order
        ``errorCode, message, hresult, hresult_hex, details, errorCategory,
        isRetryable, suggestedNextActions``. ``errorCode`` ← ``code.value`` and
        ``details`` ← ``context`` (renamed from ``code``/``context``);
        ``hresult``/``hresult_hex`` stay additive; ``errorCategory`` /
        ``isRetryable`` / ``suggestedNextActions`` are unchanged pure functions.
        No schema bump; the new codes are additive.
        """

        return {
            "errorCode": self.code.value,
            "message": self.message,
            "hresult": self.hresult,
            "hresult_hex": self.hresult_hex,
            "details": self.context,
            "errorCategory": category_for(self.code),
            "isRetryable": retryable_for(self.code),
            "suggestedNextActions": actions_for(self.code, self.context),
        }

    def __str__(self) -> str:  # pragma: no cover - trivial
        if self.hresult is not None:
            return f"[{self.code.value}] {self.message} ({self.hresult_hex})"
        return f"[{self.code.value}] {self.message}"


class AccessVbomDisabledError(ExcelVbaMcpError):
    """Trust access to the VBA project object model is disabled (spec §7.2, §10)."""

    default_code = ErrorCode.ACCESS_VBOM_DISABLED
    default_message = "Access to the VBA project object model is disabled."


class VbaProjectLockedError(ExcelVbaMcpError):
    """The VBA project is locked/password-protected; protection must not be bypassed."""

    default_code = ErrorCode.VBA_PROJECT_LOCKED
    default_message = "The VBA project is locked."


class MissingReferenceError(ExcelVbaMcpError):
    """A VBA reference is broken or missing (spec §7.2)."""

    default_code = ErrorCode.MISSING_REFERENCE
    default_message = "A required VBA reference is missing or broken."


class ComBusyError(ExcelVbaMcpError):
    """Excel/VBE is busy or rejected a call; do not blindly replay mutations."""

    default_code = ErrorCode.COM_BUSY
    default_message = "Excel is busy or rejected the call."


class RpcCallRejectedError(ComBusyError):
    """RPC_E_CALL_REJECTED from a busy STA; retry only after checking busy state."""

    default_code = ErrorCode.RPC_E_CALL_REJECTED
    default_message = "The COM call was rejected (RPC_E_CALL_REJECTED)."


class CompileError(ExcelVbaMcpError):
    """VBE reported a compile error; carry dialog text and selection (spec §8.2)."""

    default_code = ErrorCode.COMPILE_ERROR
    default_message = "The VBA project failed to compile."


class CompileIndeterminateError(ExcelVbaMcpError):
    """Compile could not be confirmed (command missing, UIA failed, COM hung)."""

    default_code = ErrorCode.COMPILE_INDETERMINATE
    default_message = "The compile outcome could not be determined."


class MacroBlockedError(ExcelVbaMcpError):
    """The macro is not permitted by Office policy or server policy (spec §7.3)."""

    default_code = ErrorCode.MACRO_BLOCKED
    default_message = "The macro is blocked by policy."


class ApprovalDeniedError(ExcelVbaMcpError):
    """The operation was denied by the approval policy."""

    default_code = ErrorCode.APPROVAL_DENIED
    default_message = "The operation was denied by the approval policy."


class MacroNotFoundError(ExcelVbaMcpError):
    """The procedure does not exist or is not a callable entry point."""

    default_code = ErrorCode.MACRO_NOT_FOUND
    default_message = "The requested macro was not found."


class MacroRuntimeError(ExcelVbaMcpError):
    """The macro executed but ended in a runtime error (spec §7.3)."""

    default_code = ErrorCode.MACRO_RUNTIME_ERROR
    default_message = "The macro raised a runtime error."


class MacroTimeoutError(ExcelVbaMcpError):
    """The macro exceeded the soft timeout; the session enters recovery (spec §9.4)."""

    default_code = ErrorCode.MACRO_TIMEOUT
    default_message = "The macro exceeded the soft timeout."


class ExcelHungError(ExcelVbaMcpError):
    """Excel is hung or a modal dialog blocks COM; recovery required (spec §9.4)."""

    default_code = ErrorCode.EXCEL_HUNG
    default_message = "Excel is unresponsive."


class SnapshotFailedError(ExcelVbaMcpError):
    """A snapshot could not be created; mutating operations must not proceed."""

    default_code = ErrorCode.SNAPSHOT_FAILED
    default_message = "The snapshot operation failed."


class WorkbookReadOnlyError(ExcelVbaMcpError):
    """The workbook is not writable (read-only open, protection, or lock)."""

    default_code = ErrorCode.WORKBOOK_READ_ONLY
    default_message = "The workbook is read-only."


class PathNotAllowedError(ExcelVbaMcpError):
    """The path is outside the allowlisted roots (spec §10)."""

    default_code = ErrorCode.PATH_NOT_ALLOWED
    default_message = "The path is not allowed by policy."


class ArtifactWriteError(ExcelVbaMcpError):
    """An artifact could not be written under the allowed artifact root.

    Raised when the atomic temp-file write or the final ``os.replace`` fails.
    No partial artifact is persisted and no false success is reported
    (spec §9, §11).

    The spec §11 error vocabulary is closed, so the persistence-failure code is
    reused: ``SNAPSHOT_FAILED`` ("a persistence operation failed, nothing was
    mutated") rather than minting a new code.
    """

    default_code = ErrorCode.SNAPSHOT_FAILED
    default_message = "The artifact could not be written."


# ---------------------------------------------------------------------------
# Phase 2 typed errors (task_040; additive to the closed spec §11 set)
# ---------------------------------------------------------------------------


class InvalidRangeError(ExcelVbaMcpError):
    """A Range target is invalid, ambiguous, multi-area, or oversized (task_038 §7)."""

    default_code = ErrorCode.INVALID_RANGE
    default_message = "The range reference is invalid or out of bounds."


class RangeTooLargeError(ExcelVbaMcpError):
    """A syntactically valid Range exceeds ``MAX_RANGE_CELLS`` (WP-Range).

    Raised by the range read/write preflight when a well-formed target is larger
    than the 100_000-cell cap. Distinct from :class:`InvalidRangeError` (malformed
    / ambiguous / wrong-shaped ranges): the address is valid, the caller must page
    the work into smaller ranges. Not retryable.
    """

    default_code = ErrorCode.RANGE_TOO_LARGE
    default_message = "The range exceeds the maximum supported cell count."


class ProtectedTargetError(ExcelVbaMcpError):
    """The mutation target is protected, locked, or otherwise read-only."""

    default_code = ErrorCode.PROTECTED_TARGET
    default_message = "The target is protected or read-only."


class UnsupportedPivotSourceError(ExcelVbaMcpError):
    """The PivotTable/PivotCache source is external, OLAP, or Data Model (task_039 §7)."""

    default_code = ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    default_message = "The PivotTable source is not supported for this operation."


class StaleTargetError(ExcelVbaMcpError):
    """A PivotTable/PivotCache identity no longer matches the workbook state."""

    default_code = ErrorCode.STALE_TARGET
    default_message = "The target changed since it was read; retry with fresh identity."


class AmbiguousFilterError(ExcelVbaMcpError):
    """A field/caption/item filter resolved to zero or more than one target."""

    default_code = ErrorCode.AMBIGUOUS_FILTER
    default_message = "The filter target is missing or ambiguous."


class RefreshFailedError(ExcelVbaMcpError):
    """A PivotTable/PivotCache refresh returned no success signal (task_039 §4)."""

    default_code = ErrorCode.REFRESH_FAILED
    default_message = "The PivotTable refresh did not complete successfully."


class ModalUiBlockedError(ExcelVbaMcpError):
    """An unexpected modal dialog blocked the operation; recovery is required.

    The window is never surfaced or dismissed; the caller receives a truthful
    blocked/indeterminate outcome (background contract).
    """

    default_code = ErrorCode.MODAL_UI_BLOCKED
    default_message = "A modal dialog blocked the operation."


class ListObjectIntersectionError(ExcelVbaMcpError):
    """A structural Range operation intersects an Excel Table (task_038 §5.3).

    Structural row/column changes are owned by the Tables/ListObject services and
    must not be performed through the Range API.
    """

    default_code = ErrorCode.LISTOBJECT_INTERSECTION
    default_message = "The range intersects an Excel Table; use the table tools instead."


# ---------------------------------------------------------------------------
# Task FE-22 (finding F-5): caller-error typed errors (additive)
#
# These make a caller-supplied failure distinguishable from genuine COM
# contention. ``ExcelVbaMcpError.default_code = COM_BUSY`` remains the generic
# fallback for unclassified errors; these subclasses are raised only for the
# specific conditions below.
# ---------------------------------------------------------------------------


class WorkbookNotFoundError(ExcelVbaMcpError):
    """The ``workbook_id`` is unknown, not open, or its owned instance was lost.

    Raised for an unknown ``workbook_id`` (a caller error) and for a read/mutation
    path that discovers the working copy is no longer open in the owned instance
    (the lost-instance case). It is deliberately distinct from ``COM_BUSY`` so a
    client can tell "bad id / lost session" from "Excel is busy".

    Note: ``macro_run`` / ``vba_compile_diagnostics`` still report the lost owned instance as
    ``EXCEL_HUNG`` (``needs_recovery``) because macro execution requires recovery;
    the plain read/mutation paths report ``WORKBOOK_NOT_FOUND``.
    """

    default_code = ErrorCode.WORKBOOK_NOT_FOUND
    default_message = "The workbook is not open."


class RevisionConflictError(ExcelVbaMcpError):
    """The component source changed since it was read (stale ``expected_source_hash``).

    An optimistic-concurrency failure on a VBA component mutation, not COM
    contention.
    """

    default_code = ErrorCode.REVISION_CONFLICT
    default_message = "The component source changed since it was read."


class ObjectNotFoundError(ExcelVbaMcpError):
    """A workbook object (shape / control) addressed by name does not exist.

    Raised for a missing shape / control so a caller can distinguish "no such
    object, do not retry" from the generic ``COM_BUSY`` fallback ("Excel busy,
    retry"). Deliberately distinct from ``WORKBOOK_NOT_FOUND`` (unknown
    ``workbook_id``) and ``MACRO_NOT_FOUND`` (unknown procedure).
    """

    default_code = ErrorCode.NOT_FOUND
    default_message = "The requested object was not found."


class NameConflictError(ExcelVbaMcpError):
    """A defined name / table column header collides with an existing one.

    Raised for a duplicate identifier in the requested scope and for a
    cross-scope collision (the same identifier already exists at workbook level
    or on another worksheet), and for a duplicate table column header. It is
    deliberately distinct from the generic ``COM_BUSY`` fallback so a caller can
    tell "choose a unique name or delete the existing one" from "Excel busy,
    retry". Not retryable; category ``validation`` (WP1, additive).
    """

    default_code = ErrorCode.NAME_CONFLICT
    default_message = "The name conflicts with an existing one."


def error_from_hresult(
    hresult: int,
    message: str,
    *,
    context: dict[str, Any] | None = None,
) -> ExcelVbaMcpError:
    """Build the most specific typed error for a known HRESULT.

    Falls back to :class:`ExcelVbaMcpError` carrying ``COM_BUSY`` semantics when
    the HRESULT is unrecognized, so no HRESULT/message information is lost.
    """

    code = code_from_hresult(hresult, default=ErrorCode.COM_BUSY)
    assert code is not None  # default guarantees a code
    factory = _ERROR_BY_CODE.get(code, ExcelVbaMcpError)
    return factory(message, hresult=hresult, context=context)


# Mapping used by ``error_from_hresult`` to construct the typed subclass. The
# base ``ExcelVbaMcpError`` is the fallback.
_ERROR_BY_CODE: dict[ErrorCode, type[ExcelVbaMcpError]] = {
    ErrorCode.ACCESS_VBOM_DISABLED: AccessVbomDisabledError,
    ErrorCode.VBA_PROJECT_LOCKED: VbaProjectLockedError,
    ErrorCode.MISSING_REFERENCE: MissingReferenceError,
    ErrorCode.COM_BUSY: ComBusyError,
    ErrorCode.RPC_E_CALL_REJECTED: RpcCallRejectedError,
    ErrorCode.COMPILE_ERROR: CompileError,
    ErrorCode.COMPILE_INDETERMINATE: CompileIndeterminateError,
    ErrorCode.MACRO_BLOCKED: MacroBlockedError,
    ErrorCode.MACRO_NOT_FOUND: MacroNotFoundError,
    ErrorCode.MACRO_RUNTIME_ERROR: MacroRuntimeError,
    ErrorCode.MACRO_TIMEOUT: MacroTimeoutError,
    ErrorCode.EXCEL_HUNG: ExcelHungError,
    ErrorCode.SNAPSHOT_FAILED: SnapshotFailedError,
    ErrorCode.WORKBOOK_READ_ONLY: WorkbookReadOnlyError,
    ErrorCode.PATH_NOT_ALLOWED: PathNotAllowedError,
    # Phase 2 additions (task_040).
    ErrorCode.INVALID_RANGE: InvalidRangeError,
    ErrorCode.RANGE_TOO_LARGE: RangeTooLargeError,
    ErrorCode.PROTECTED_TARGET: ProtectedTargetError,
    ErrorCode.UNSUPPORTED_PIVOT_SOURCE: UnsupportedPivotSourceError,
    ErrorCode.STALE_TARGET: StaleTargetError,
    ErrorCode.AMBIGUOUS_FILTER: AmbiguousFilterError,
    ErrorCode.REFRESH_FAILED: RefreshFailedError,
    ErrorCode.MODAL_UI_BLOCKED: ModalUiBlockedError,
    ErrorCode.LISTOBJECT_INTERSECTION: ListObjectIntersectionError,
    # Task_b4 (fix_004): approval denial is a typed domain error, distinct from
    # COM contention. Inert for ``error_from_hresult`` (no HRESULT maps to it).
    ErrorCode.APPROVAL_DENIED: ApprovalDeniedError,
    # Task FE-22 (finding F-5): caller-error codes, inert for
    # ``error_from_hresult`` (no HRESULT maps to them).
    ErrorCode.WORKBOOK_NOT_FOUND: WorkbookNotFoundError,
    ErrorCode.REVISION_CONFLICT: RevisionConflictError,
    # Block B7a (finding F6): missing shape / control. Inert for
    # ``error_from_hresult`` (no HRESULT maps to it); registered so the
    # code -> exception mapping stays complete for callers that rebuild a typed
    # error from a code.
    ErrorCode.NOT_FOUND: ObjectNotFoundError,
    # WP1 (additive): duplicate / cross-scope name conflict. Inert for
    # ``error_from_hresult`` (no HRESULT maps to it); registered so the
    # code -> exception mapping stays complete.
    ErrorCode.NAME_CONFLICT: NameConflictError,
}
