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
        """Return a JSON-serializable representation (no COM objects)."""

        return {
            "code": self.code.value,
            "message": self.message,
            "hresult": self.hresult,
            "hresult_hex": self.hresult_hex,
            "context": self.context,
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

    Note: ``macro_run`` / ``vba_compile`` still report the lost owned instance as
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
}
