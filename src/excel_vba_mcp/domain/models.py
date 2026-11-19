"""Phase 1 serializable domain DTOs (spec §5.2, §6, §7, §8.2, §9, Appendix A).

All models are Pydantic v2 ``BaseModel`` subclasses that serialize to JSON via
``model_dump(mode="json")`` / ``model_dump_json()``. They carry only Python
values (strings, numbers, booleans, arrays, ISO-8601 timestamps, dict/list) —
never COM objects — so they can cross the STA-worker boundary and be emitted as
MCP structured output.

The models intentionally cover only Stage 1 entities. Phase 2 concepts
(full Range/Cells API, PivotTables, Charts, etc.) are out of scope here.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from excel_vba_mcp.domain.statuses import (
    CaseStatus,
    CellValueKind,
    CompileMethod,
    CompileStatus,
    ControlKind,
    DiagnosticConfidence,
    MacroMappingStatus,
    MacroRunStatus,
    NameScope,
    NameTargetKind,
    ProcedureType,
    ProcedureVisibility,
    ReferenceStatus,
    SessionMode,
    SheetObjectType,
    SheetVisibility,
    SnapshotStatus,
    StepStatus,
    TestRunStatus,
    VbaComponentType,
    VbaExportStatus,
    WorkbookSessionState,
)

# Pydantic model version for the Phase 1 contracts. Bump when the contract
# changes incompatibly.
SCHEMA_VERSION = 1


class DomainModel(BaseModel):
    """Base for all domain DTOs: strict, JSON-serializable, no extra fields."""

    model_config = ConfigDict(extra="forbid", use_enum_values=False)


class TimestampedModel(DomainModel):
    """Adds creation/update timestamps (ISO-8601) to a DTO."""

    created_at: datetime = Field(description="ISO-8601 creation timestamp (UTC).")
    updated_at: datetime = Field(description="ISO-8601 last-update timestamp (UTC).")


# ---------------------------------------------------------------------------
# Workbook / session identity
# ---------------------------------------------------------------------------


class WorkbookInfo(DomainModel):
    """Workbook identity and top-level metadata (spec Appendix A, §7.1)."""

    workbook_id: str = Field(description="Stable server-assigned workbook id.")
    name: str
    format: str = Field(description="File format, e.g. 'xlsm'.")
    path: str = Field(description="Actual working path of the opened workbook.")
    source_path: str | None = Field(default=None, description="Original allowlisted path.")
    read_only: bool
    has_vba: bool
    excel_version: str
    session_id: str
    mode: SessionMode
    model_version: int = Field(default=1)


class SessionStatus(TimestampedModel):
    """Excel session / workbook runtime status (spec §6.1 ``excel_get_status``)."""

    session_id: str
    workbook_id: str | None = None
    state: WorkbookSessionState
    mode: SessionMode
    busy: bool = False
    read_only: bool = False
    dirty: bool = False
    active_operation: str | None = None
    active_test_run_id: str | None = None
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# VBA project, components, procedures, references
# ---------------------------------------------------------------------------


class VbaReferenceInfo(DomainModel):
    """A VBA project reference (spec §6.2 ``vba_get_references``, §7.2)."""

    name: str
    guid: str | None = None
    version: str | None = None
    path: str | None = None
    status: ReferenceStatus
    is_broken: bool = False


class VbaComponentInfo(DomainModel):
    """Inventory entry for a VBComponent (spec §6.2, §7.2)."""

    component_name: str
    component_type: VbaComponentType
    line_count: int = 0
    declaration_line_count: int | None = None
    source_hash: str | None = None
    is_document_module: bool = False
    has_designer: bool = False
    is_exportable: bool = False


class VbaProcedureInfo(DomainModel):
    """A Sub/Function/Property within a component (spec §6.2)."""

    component_name: str
    name: str
    procedure_type: ProcedureType
    visibility: ProcedureVisibility = ProcedureVisibility.UNKNOWN
    start_line: int
    end_line: int
    code: str = ""
    is_heuristic: bool = Field(
        default=False,
        description="True when extracted via a textual fallback rather than VBIDE.",
    )
    source: str = Field(
        default="code_module",
        description="Extraction source: 'code_module' (VBIDE API) or 'heuristic'.",
    )


class VbaProjectInfo(DomainModel):
    """VBProject summary (spec §6.2 ``vba_project_info``)."""

    available: bool
    locked: bool = False
    name: str | None = None
    component_count: int = 0
    components: list[VbaComponentInfo] = Field(default_factory=list)
    references: list[VbaReferenceInfo] = Field(default_factory=list)
    missing_references: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class VbaComponentSource(TimestampedModel):
    """Read result for a single component's code module (spec §6.2)."""

    component_name: str
    component_type: VbaComponentType
    source: str
    line_count: int
    source_hash: str
    revision: int = Field(description="Monotonic revision of the component source.")


class VbaMutationResult(TimestampedModel):
    """Result of a guarded VBA source mutation (replace / insert, spec §7.2).

    Captures the source hash **before** and **after** the mutation together with
    the export paths written on the STA worker, so a caller can verify that the
    working copy changed and that the original fixture was never touched. The
    ``audit_id`` correlates this mutation with its audit-log entry.
    """

    workbook_id: str
    component_name: str
    action: str = Field(description="'replace' or 'insert'.")
    before_hash: str = Field(description="Normalized SHA-256 of the code body before mutation.")
    after_hash: str = Field(description="Normalized SHA-256 of the code body after mutation.")
    before_export_path: str = Field(
        description="Absolute path of the exported .bas/.cls/.frm file before mutation."
    )
    after_export_path: str = Field(
        description="Absolute path of the exported .bas/.cls/.frm file after mutation."
    )
    audit_id: str = Field(description="Audit log event id for this mutation.")
    procedure: str | None = Field(
        default=None,
        description="Affected procedure name for an insert, or None for a full replace.",
    )


class VbaExportEntry(DomainModel):
    """One exported component file (spec §7.2 export manifest).

    ``source_hash`` is the SHA-256 of the **normalized code body** (EOL unified
    to ``\\n``, per-line trailing whitespace trimmed, leading/trailing blank lines
    dropped). ``file_sha256`` is the SHA-256 of the exported file bytes; the two
    differ for UserForms because the exported ``.frm`` carries a designer header
    and ``.frx`` is not byte-reproducible.
    """

    component_name: str
    component_type: VbaComponentType
    file_name: str
    extension: str
    file_path: str
    file_sha256: str | None = None
    source_hash: str = Field(description="SHA-256 of the normalized code body.")
    resource_file_name: str | None = Field(
        default=None, description="Companion .frx name for UserForms."
    )
    resource_file_sha256: str | None = Field(
        default=None, description="SHA-256 of the .frx bytes (not reproducible)."
    )
    export_status: VbaExportStatus = VbaExportStatus.OK
    error: str | None = None


class VbaExportManifest(TimestampedModel):
    """Result of ``vba_export_project`` (spec §7.2)."""

    workbook_id: str
    out_dir: str
    components: list[VbaExportEntry] = Field(default_factory=list)
    source_hash: str | None = Field(
        default=None,
        description="Combined SHA-256 over all normalized component code bodies.",
    )


class VbaSearchMatch(DomainModel):
    """One bounded ``vba_search_code`` hit (spec §6.2)."""

    component_name: str
    component_type: VbaComponentType
    line_number: int
    line_text: str


# ---------------------------------------------------------------------------
# Compile diagnostics (spec §8.2)
# ---------------------------------------------------------------------------


class CompileDiagnostic(DomainModel):
    """Compile outcome diagnostic (spec §8.2 diagnostic format)."""

    status: CompileStatus
    project: str
    method: CompileMethod
    error_message: str | None = None
    module: str | None = None
    line_number: int | None = None
    selected_text: str | None = None
    dialog_title: str | None = None
    screenshot: str | None = Field(default=None, description="Path to the screenshot artifact.")
    confidence: DiagnosticConfidence = DiagnosticConfidence.LOW
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Macro execution (spec §6.3, §7.3)
# ---------------------------------------------------------------------------


class MacroRunResult(TimestampedModel):
    """Result of an ``Application.Run`` invocation (spec §6.3, §7.3)."""

    run_id: str
    macro_name: str
    status: MacroRunStatus
    started_at: datetime
    finished_at: datetime
    duration_ms: int
    return_value: Any | None = Field(default=None, description="Serializable Variant return value.")
    error_code: str | None = None
    error_text: str | None = None


# ---------------------------------------------------------------------------
# Snapshots (spec §6.4, §9.1)
# ---------------------------------------------------------------------------


class SnapshotManifest(DomainModel):
    """Snapshot manifest metadata (spec §9.1)."""

    snapshot_id: str
    workbook_id: str
    reason: str | None = None
    baseline_id: str | None = None
    test_id: str | None = None
    created_at: datetime
    source_path: str
    stored_path: str
    sha256: str
    excel_version: str | None = None
    has_vba: bool | None = None
    model_version: int = Field(default=1)
    vba_exports: list[str] = Field(default_factory=list)


class SnapshotInfo(DomainModel):
    """Snapshot record returned by list/restore (spec §6.4)."""

    snapshot_id: str
    workbook_id: str
    created_at: datetime
    reason: str | None = None
    baseline_id: str | None = None
    test_id: str | None = None
    stored_path: str
    sha256: str
    status: SnapshotStatus
    manifest: SnapshotManifest | None = None


# ---------------------------------------------------------------------------
# Workbook model (spec §6.1, §7.5, §7.6, §7.7, §7.4, Appendix A)
# ---------------------------------------------------------------------------


class TableColumnInfo(DomainModel):
    """A column of a ListObject table (spec §7.6)."""

    name: str
    index: int
    data_type: str | None = None
    has_totals_calculation: bool = False
    totals_function: str | None = None


class TableInfo(DomainModel):
    """Excel Table / ListObject summary (spec §7.6)."""

    name: str
    display_name: str | None = None
    sheet_name: str
    sheet_code_name: str | None = None
    range_address: str
    header_row_range: str | None = None
    data_body_range: str | None = None
    total_row_range: str | None = None
    show_totals: bool = False
    table_style: str | None = None
    row_count: int = 0
    columns: list[TableColumnInfo] = Field(default_factory=list)


class CellValue(DomainModel):
    """A typed cell value for read-only output assertions (spec §7.6, §9.3)."""

    kind: CellValueKind
    value: str | int | float | bool | None = None
    address: str | None = None


class SheetInfo(DomainModel):
    """Sheet inventory entry (spec §7.5)."""

    name: str
    code_name: str | None = None
    index: int
    object_type: SheetObjectType
    visibility: SheetVisibility
    protected: bool | None = Field(default=None, description="Diagnostic only.")
    table_count: int = 0
    used_range_address: str | None = None
    tables: list[TableInfo] = Field(default_factory=list)
    names: list[str] = Field(default_factory=list)
    controls: list[ControlInfo] = Field(default_factory=list)


class DefinedNameInfo(DomainModel):
    """A defined name (spec §7.7)."""

    name: str
    scope: NameScope
    name_local: str | None = None
    refers_to: str | None = None
    refers_to_local: str | None = None
    visible: bool = True
    sheet_name: str | None = Field(default=None, description="Set for worksheet-scoped names.")
    target_kind: NameTargetKind = NameTargetKind.UNKNOWN
    target_sheet: str | None = None
    target_address: str | None = None


class ControlInfo(DomainModel):
    """Shape / control inventory entry (spec §7.4)."""

    name: str
    kind: ControlKind
    shape_id: str | int | None = None
    prog_id: str | None = Field(default=None, description="ActiveX ProgID/ClassType.")
    caption: str | None = None
    alternative_text: str | None = None
    top_left_cell: str | None = None
    width: float | None = None
    height: float | None = None
    visible: bool | None = None
    on_action: str | None = None
    macro_mapping: MacroMappingStatus = MacroMappingStatus.UNKNOWN
    sheet_name: str | None = None


class WorkbookModel(TimestampedModel):
    """Compact, paginated Phase 1 workbook model (spec §6.1, Appendix A)."""

    schema_version: int = SCHEMA_VERSION
    workbook: WorkbookInfo
    sheets: list[SheetInfo] = Field(default_factory=list)
    vba_project: VbaProjectInfo | None = None
    names: list[DefinedNameInfo] = Field(default_factory=list)
    controls: list[ControlInfo] = Field(default_factory=list)
    model_version: int = Field(default=1)
    model_hash: str | None = None


# ---------------------------------------------------------------------------
# Test run / case results (spec §6.4, §9)
# ---------------------------------------------------------------------------


class AssertionResult(DomainModel):
    """Result of evaluating one assertion (spec §9.3)."""

    type: str
    passed: bool
    expected: Any | None = None
    actual: Any | None = None
    message: str | None = None


class TestStepResult(DomainModel):
    """Result of executing one test step (spec §9.2)."""

    action: str
    status: StepStatus
    error_code: str | None = None
    error_text: str | None = None
    details: dict[str, Any] = Field(default_factory=dict)


class TestCaseResult(TimestampedModel):
    """Result of a single declarative test case (spec §9)."""

    case_id: str
    status: CaseStatus
    steps: list[TestStepResult] = Field(default_factory=list)
    assertions: list[AssertionResult] = Field(default_factory=list)
    duration_ms: int = 0
    working_copy_path: str | None = None
    notes: list[str] = Field(default_factory=list)


class TestRunResult(TimestampedModel):
    """Overall result of a test suite run (spec §6.4, §9)."""

    test_run_id: str
    suite: str
    status: TestRunStatus
    workbook_id: str | None = None
    baseline_snapshot_id: str | None = None
    cases: list[TestCaseResult] = Field(default_factory=list)
    started_at: datetime
    finished_at: datetime | None = None
    duration_ms: int | None = None
    artifact_paths: list[str] = Field(default_factory=list)
