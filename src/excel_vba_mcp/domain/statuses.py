"""Operation and verification statuses for Phase 1 (spec §6, §8, §9, §11).

This module is the single source of truth for operation classification and for
the status vocabularies shared across layers (API, application services, COM
worker). Enums are ``str``-valued so they serialize directly in
``model_dump(mode="json")`` / ``model_dump_json()`` without custom encoders.

Statuses are closed vocabularies derived from the specification only; they do
not introduce Phase 2 concepts.
"""

from __future__ import annotations

from enum import StrEnum


class OperationClass(StrEnum):
    """Classification of an operation for policy and COM retry decisions.

    Shared by the approval policy (task_009) and the COM worker (task_010):
    read-only operations are safe to retry after a busy/rejected result, while
    mutating and executing operations must never be blindly replayed.
    """

    READ_ONLY = "read_only"
    MUTATING = "mutating"
    EXECUTING = "executing"


class WorkbookSessionState(StrEnum):
    """Lifecycle state of the server-owned Excel session (spec §7.1, §9.4)."""

    STARTING = "starting"
    RUNNING = "running"
    NEEDS_RECOVERY = "running/needs_recovery"
    BUSY = "busy"
    CLOSING = "closing"
    CLOSED = "closed"
    ERROR = "error"


class SessionMode(StrEnum):
    """Workbook open mode controlling macro trust (spec §7.1, §10.1)."""

    INSPECT = "inspect"
    EXECUTE = "execute"


class CompileStatus(StrEnum):
    """VBE compile outcome (spec §6.3, §8.1)."""

    COMPILED = "compiled"
    ALREADY_COMPILED = "already_compiled"
    COMPILE_ERROR = "compile_error"
    INDETERMINATE = "indeterminate"


class CompileMethod(StrEnum):
    """How a compile outcome was obtained (spec §8.2 diagnostic field)."""

    VBE_COMMANDBAR = "vbe_commandbar"
    UIA = "uia"
    UNKNOWN = "unknown"


class DiagnosticConfidence(StrEnum):
    """Confidence attached to a compile diagnostic (spec §8.2)."""

    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


class MacroRunStatus(StrEnum):
    """Outcome of an ``Application.Run`` macro invocation (spec §6.3, §7.3, §11)."""

    SUCCESS = "success"
    RUNTIME_ERROR = "runtime_error"
    BLOCKED = "blocked"
    NOT_FOUND = "not_found"
    TIMEOUT = "timeout"
    NEEDS_RECOVERY = "needs_recovery"


class TestRunStatus(StrEnum):
    """Overall status of a test suite run (spec §6.4, §9)."""

    RUNNING = "running"
    PASSED = "passed"
    FAILED = "failed"
    BLOCKED = "blocked"
    SKIPPED = "skipped"
    INDETERMINATE = "indeterminate"


class CaseStatus(StrEnum):
    """Status of a single test case (task_035: passed/failed/blocked/skipped/indeterminate)."""

    PASSED = "passed"
    FAILED = "failed"
    BLOCKED = "blocked"
    SKIPPED = "skipped"
    INDETERMINATE = "indeterminate"


class StepStatus(StrEnum):
    """Status of a single test step (spec §9.2)."""

    PASSED = "passed"
    FAILED = "failed"
    SKIPPED = "skipped"


class SnapshotStatus(StrEnum):
    """Outcome of a snapshot operation (spec §6.4, §9.1)."""

    CREATED = "created"
    FAILED = "failed"


class SheetObjectType(StrEnum):
    """Excel sheet object type (spec §7.5: Sheets vs Worksheets)."""

    WORKSHEET = "worksheet"
    CHART = "chart"
    MACRO = "macro"  # legacy Excel 4.0 macro sheet
    DIALOG = "dialog"  # legacy Excel 4.0 dialog sheet


class SheetVisibility(StrEnum):
    """Worksheet visibility (spec §7.5)."""

    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "very_hidden"


class VbaComponentType(StrEnum):
    """VBA component type (spec §7.2)."""

    STANDARD_MODULE = "standard_module"
    CLASS_MODULE = "class_module"
    USER_FORM = "user_form"
    DOCUMENT_MODULE = "document_module"


class ProcedureType(StrEnum):
    """VBA procedure type (spec §6.2)."""

    SUB = "sub"
    FUNCTION = "function"
    PROPERTY = "property"


class ProcedureVisibility(StrEnum):
    """VBA procedure visibility when reliably extracted (spec §6.2)."""

    PUBLIC = "public"
    PRIVATE = "private"
    FRIEND = "friend"
    UNKNOWN = "unknown"


class VbaProcedureKind(StrEnum):
    """VBA procedure kind (VBIDE ``ProcKind`` codes, spec §6.2)."""

    SUB = "sub"
    FUNCTION = "function"
    PROPERTY_GET = "property_get"
    PROPERTY_LET = "property_let"
    PROPERTY_SET = "property_set"
    UNKNOWN = "unknown"


class VbaExportStatus(StrEnum):
    """Outcome of exporting one VBA component (spec §7.2)."""

    OK = "ok"
    FAILED = "failed"
    SKIPPED = "skipped"


class ReferenceStatus(StrEnum):
    """VBA reference availability (spec §6.2, §7.2)."""

    AVAILABLE = "available"
    BROKEN = "broken"
    MISSING = "missing"


class NameScope(StrEnum):
    """Defined-name scope (spec §7.7)."""

    WORKBOOK = "workbook"
    WORKSHEET = "worksheet"


class NameTargetKind(StrEnum):
    """What a defined name refers to (spec §7.7: range vs non-range).

    ``RANGE`` names resolve through ``Name.RefersToRange``; every other kind is
    classified heuristically from the ``RefersTo`` string without evaluating the
    formula/expression (no ``Application.Evaluate``, no Python ``eval``).
    """

    RANGE = "range"
    FORMULA = "formula"
    CONSTANT = "constant"
    EXTERNAL = "external"
    ERROR = "error"
    UNKNOWN = "unknown"


class ControlKind(StrEnum):
    """Shape / control kind (spec §7.4)."""

    FORM_BUTTON = "form_button"
    FORM_CONTROL = "form_control"
    ACTIVEX = "activex"
    SHAPE = "shape"


class MacroMappingStatus(StrEnum):
    """Whether a control OnAction maps to a resolved VBA procedure (spec §7.4)."""

    RESOLVED = "resolved"
    ORPHANED = "orphaned"
    PROBABLE = "probable"
    UNKNOWN = "unknown"


class CellValueKind(StrEnum):
    """Cell value type for read-only output assertions (spec §7.6, §9.3)."""

    STRING = "string"
    NUMBER = "number"
    BOOLEAN = "boolean"
    DATE = "date"
    ERROR = "error"
    EMPTY = "empty"


class BaselinePolicy(StrEnum):
    """Baseline handling for a test suite (spec §9.1, Appendix B)."""

    CREATE = "create"
    USE_EXISTING = "use_existing"
