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
from typing import Any, Literal, Self

from pydantic import BaseModel, ConfigDict, Field, model_validator

from excel_vba_mcp.domain.statuses import (
    AccessMode,
    CaseStatus,
    CellValueKind,
    CompileMethod,
    CompileStatus,
    ConditionalFormatRequestType,
    ConditionalFormatType,
    ControlFormType,
    ControlKind,
    DataValidationType,
    DiagnosticConfidence,
    MacroMappingStatus,
    MacroRunStatus,
    ModelSectionStatus,
    NameScope,
    NameTargetKind,
    PivotAggregation,
    PivotFieldLayoutOrientation,
    PivotFieldOrientation,
    PivotFilterKind,
    PivotFilterRequestKind,
    PivotRefreshMethod,
    PivotRefreshStatus,
    PivotSourceType,
    ProcedureType,
    ProcedureVisibility,
    RangeShape,
    ReferenceStatus,
    SessionMode,
    ShapeKind,
    SheetObjectType,
    SheetVisibility,
    SnapshotStatus,
    StepStatus,
    TestRunStatus,
    ValidationOperator,
    VbaAccessStatus,
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
    """Workbook identity and top-level metadata (spec Appendix A, §7.1).

    Two-path model (U6-E; plan_05 §5.4):

    - ``path`` is the **working copy** actually opened by the server — the
      materialized file under the server ``work/`` directory (``wb.FullName``).
      All reads/writes and mutations target this copy; the original is never
      touched.
    - ``source_path`` is the **original, allowlisted** path the working copy was
      materialized from. It is ``None`` when the workbook has no on-disk source
      (for example a snapshot restored from a baseline that itself had no
      source).

    ``original_name`` is intentionally **absent**: it would be redundant with
    ``name`` (the file name is identical for the working copy and its source),
    so no separate field is exposed.

    Examples:

    - New open of ``C:/data/report.xlsm``: ``path`` =
      ``<temp>/excel_vba_mcp/work/<id>.xlsm``, ``source_path`` =
      ``C:/data/report.xlsm``, ``name`` = ``report.xlsm``.
    - Writable working copy: identical to the open above; edits land in ``path``
      and ``source_path`` still points at the untouched original.
    - Separate source/current: after ``excel_save`` the working copy at ``path``
      diverges from ``source_path`` (which remains the pristine original) until
      the change is deliberately persisted.
    """

    workbook_id: str = Field(description="Stable server-assigned workbook id.")
    name: str
    format: str = Field(description="File format, e.g. 'xlsm'.")
    path: str = Field(
        description=(
            "Actual working path of the opened workbook (the materialized "
            "working copy, wb.FullName); the original is never modified."
        )
    )
    source_path: str | None = Field(
        default=None,
        description=(
            "Original allowlisted path the working copy was copied from; None "
            "when the workbook has no on-disk source. There is no separate "
            "original_name (redundant with name)."
        ),
    )
    read_only: bool
    has_vba: bool
    excel_version: str
    session_id: str
    mode: SessionMode
    model_version: int = Field(default=1)


class SessionPolicy(DomainModel):
    """Per-workbook policy subset co-located on ``excel_get_status`` (U2; plan_03 §6.2).

    The authoritative server-level policy lives in ``xlvba_environment``; this
    block co-locates the per-workbook context on the status DTO. It is additive
    and optional (``SessionStatus.policy`` defaults to ``None``).

    ``classification_is_authorization`` is always ``False`` (a BG06 classification
    is bookkeeping only and never grants run permission) and ``r13_preflight`` is
    always ``"veto"`` (the R13 scan can only veto a dispatch, never certify one).
    ``has_snapshot`` reads the **in-memory** snapshot gate
    (``SnapshotService.has_snapshot`` -> ``_latest_snapshot``), not the on-disk
    store.
    """

    approval_mode: AccessMode
    mutating_requires_grant: bool
    executing_requires_grant: bool
    recovery_auto_allowed: bool
    snapshot_required_for_mutations: bool
    has_snapshot: bool = Field(
        default=False,
        description="Whether an in-memory snapshot gate exists for this workbook "
        "(SnapshotService.has_snapshot; the in-memory _latest_snapshot, not the store).",
    )
    classification_is_authorization: Literal[False] = False
    r13_preflight: Literal["veto"] = "veto"


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
    access_mode: AccessMode | None = None
    policy: SessionPolicy | None = None
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Environment / policy discovery (U2; plan_03 §6.1)
# ---------------------------------------------------------------------------
# The ``xlvba_environment`` payload is a strongly typed, additive DTO
# (``extra="forbid"``). It is non-COM and side-effect-free: ``excel.build`` is
# always ``None`` (no 4-part build source exists — never fabricated) and
# ``excel.available`` is ``False`` when the worker is down. ``has_snapshot`` in
# ``workbooks`` reads the in-memory snapshot gate, not the on-disk store.


class EnvironmentServerInfo(DomainModel):
    """Server-level identity and configuration (``xlvba_environment.server``)."""

    name: str
    version: str
    access_mode: AccessMode
    macro_timeout_s: float
    open_timeout_s: float
    liveness_ceiling_s: float
    log_level: str


class EnvironmentExcelInfo(DomainModel):
    """Owned Excel instance summary (``xlvba_environment.excel``).

    ``build`` is **always** ``None`` — there is no 4-part build source, so the
    field is honest ``null`` and can never be fabricated. ``target_build`` is a
    labelled design-time constant only (not a detected build).
    ``application_version`` is a non-COM best-effort value taken from an open
    workbook's cached ``Application.Version`` (or ``None`` when none is open).
    ``vbom_available`` is ``None`` at server level (no workbook model is built).
    """

    available: bool
    application_version: str | None = None
    build: None = Field(default=None, description="Always null; no 4-part build source exists.")
    target_build: str = Field(
        default="16.0.14332",
        description="Labelled design-time constant only; not a detected build.",
    )
    owned_pid: int | None = None
    vbom_available: bool | None = None


class EnvironmentCapabilities(DomainModel):
    """Capability flags derived from registration + status enums (plan_03 §6.1)."""

    vba: bool
    macros: bool
    pivot_tables: bool
    background_jobs: bool
    recovery: bool
    non_activating_compile: bool = Field(
        default=False,
        description="Always false on the validated build (compile never returns compiled/error).",
    )
    pivot_filter_kinds: list[str] = Field(default_factory=list)
    pivot_source_types: list[str] = Field(default_factory=list)


class EnvironmentApprovalPolicy(DomainModel):
    """Approval policy summary (``xlvba_environment.policy.approval``)."""

    mode: AccessMode
    mutating_requires_grant: bool
    executing_requires_grant: bool
    recovery_auto_allowed: bool


class EnvironmentSnapshotPolicy(DomainModel):
    """Snapshot policy summary (``xlvba_environment.policy.snapshot``)."""

    required_for_mutations: bool = True
    created_by: str = "agent:snapshot_create"


class EnvironmentPathPolicy(DomainModel):
    """Path allowlist summary (``xlvba_environment.policy.path``)."""

    allowed_roots: list[str] = Field(default_factory=list)
    allow_unc: bool = False


class EnvironmentMacroPolicy(DomainModel):
    """Macro gate summary (``xlvba_environment.policy.macro``).

    ``classification_is_authorization`` is always ``False`` and ``r13_preflight``
    is always ``"veto"`` — a classification never grants run permission and the
    R13 preflight can only veto.
    """

    background_safe_required: bool = True
    r13_preflight: Literal["veto"] = "veto"
    classification_is_authorization: Literal[False] = False


class EnvironmentPolicy(DomainModel):
    """Aggregate policy block for ``xlvba_environment`` (plan_03 §6.1)."""

    approval: EnvironmentApprovalPolicy
    snapshot: EnvironmentSnapshotPolicy = Field(default_factory=EnvironmentSnapshotPolicy)
    path: EnvironmentPathPolicy = Field(default_factory=EnvironmentPathPolicy)
    macro: EnvironmentMacroPolicy = Field(default_factory=EnvironmentMacroPolicy)


class EnvironmentWorkbook(DomainModel):
    """Per-open-workbook summary for ``xlvba_environment.workbooks`` (plan_03 §6.1).

    ``has_snapshot`` reads the in-memory snapshot gate
    (``SnapshotService.has_snapshot``), never the on-disk store.
    """

    workbook_id: str
    read_only: bool
    mode: SessionMode
    has_snapshot: bool
    dirty: bool


class EnvironmentReport(DomainModel):
    """Read-only ``xlvba_environment`` payload (plan_03 §6.1; ``extra="forbid"``).

    Non-COM and side-effect-free: it never starts the worker, never creates
    snapshots / grants / classifications and never authorizes a mutation.
    """

    schema_version: int = SCHEMA_VERSION
    server: EnvironmentServerInfo
    excel: EnvironmentExcelInfo
    capabilities: EnvironmentCapabilities
    policy: EnvironmentPolicy
    glossary_uri: str = "xlvba://glossary"
    workbooks: list[EnvironmentWorkbook] = Field(default_factory=list)


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
    session_recovery: str | None = Field(
        default=None,
        description="In-process session recovery outcome after a hang/timeout (F2b): "
        "'recovered' when the owned Excel instance was rebuilt and is running again, "
        "'needs_recovery' when recovery did not reach a running state, or None when the "
        "run did not hang. The prior workbook_id is invalid after recovery (open "
        "workbooks are discarded).",
    )


# ---------------------------------------------------------------------------
# Snapshots (spec §6.4, §9.1)
# ---------------------------------------------------------------------------


class SnapshotManifest(DomainModel):
    """Snapshot manifest metadata (spec §9.1).

    ``vba_exports`` is a list of per-component record dicts (matching the
    storage-layer manifest actually persisted in ``manifest.json``), each with
    ``component_name``, ``type``, ``path``, ``export_file_sha256`` and, for
    UserForms, ``frx``/``frx_sha256``. ``export_file_sha256`` is the SHA-256 of
    the **exported file bytes** — it is not the normalized ``source_hash`` used by
    the VBA mutation guards (F7). ``model_version`` is bumped to ``2`` for the
    ``sha256`` -> ``export_file_sha256`` field rename.
    """

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
    model_version: int = Field(default=2)
    vba_exports: list[dict[str, Any]] = Field(default_factory=list)


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


class SnapshotRestoreResult(DomainModel):
    """Result of ``snapshot_restore`` (plan_06 R1/R3; additive).

    An additive wrapper over the freshly opened restored :class:`WorkbookInfo`
    plus the disposition of the previous working copy's close (R1):

    * ``previous_close`` is ``"closed"`` when the previous working copy was
      closed without saving, ``"failed"`` when the close raised (the restore
      **still succeeds** and the restored session stays valid — the previous
      copy remains open and the agent is told via ``previous_close_error``), or
      ``"not_requested"`` when no previous ``workbook_id`` was supplied.
    * ``previous_close_error`` is safe, human-readable text (no secrets) when
      ``previous_close == "failed"``, else ``None``.

    The restored ``workbook.source_path`` is rebound to the snapshot manifest's
    ``source_path`` (the ORIGINAL), so identity / ``list_workbooks`` / reopen
    reflect the original, never the immutable snapshot file (plan_06 R3). The
    working copy itself stays ``work/<newid>.ext``; the snapshot file is never
    the session identity.
    """

    workbook: WorkbookInfo
    previous_close: Literal["closed", "failed", "not_requested"] = Field(
        description=(
            "Disposition of the previous working copy's close: 'closed' (closed "
            "without saving), 'failed' (close raised; restore still succeeds and "
            "the previous copy remains open), or 'not_requested' (no previous "
            "workbook_id supplied)."
        )
    )
    previous_close_error: str | None = Field(
        default=None,
        description="Safe, human-readable close-failure text when previous_close == 'failed'.",
    )


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
    pivot_tables: list[PivotTableInfo] = Field(
        default_factory=list,
        description="PivotTable inventory metadata scoped to this sheet (task_048); "
        "metadata only, never the PivotTable cell output.",
    )


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


class ShapeInfo(DomainModel):
    """A drawing-layer shape entry (spec §7.6 Shapes inventory).

    One entry per ``Worksheet.Shapes`` item. ``kind`` is derived from
    ``Shape.Type`` (``msoShapeType``); ``on_action`` is the shape's ``OnAction``
    text read verbatim — **the assigned macro is never executed**. Coordinates
    and size are Excel points; ``top``/``left``/``width``/``height`` are ``None``
    when the property is unavailable. ``text`` is a best-effort rendering of the
    shape's text frame (never ``None`` when a text frame exists but is empty:
    an empty string is meaningful). ``control_format`` is ``True`` only when the
    shape exposes an Excel ``ControlFormat`` object (Form Controls); ActiveX /
    plain shapes report ``None``.
    """

    name: str
    kind: ShapeKind
    shape_id: str | int | None = None
    sheet_name: str | None = None
    on_action: str | None = None
    visible: bool | None = None
    top: float | None = None
    left: float | None = None
    width: float | None = None
    height: float | None = None
    text: str | None = None
    alternative_text: str | None = None
    control_format: bool | None = Field(
        default=None,
        description="True when the shape exposes an Excel ControlFormat (Form Control).",
    )
    notes: list[str] = Field(
        default_factory=list,
        description="Diagnostics for properties that could not be read.",
    )


class ControlEntry(DomainModel):
    """A Form Control or ActiveX control entry (spec §7.6 Controls inventory).

    Form Controls and ActiveX controls live in different Excel collections and
    expose different metadata, so this DTO carries both shapes of information:

    * Form Controls (``ws.Shapes`` of type ``msoFormControl``) report
      ``control_type`` (button / checkbox / …), ``caption``, ``linked_cell`` and
      their ``on_action`` (the assigned macro, read verbatim, never executed).
    * ActiveX controls (``ws.OLEObjects``) report ``prog_id`` plus an **inferred,
      probable** event-procedure association. ``probable_procedure`` is a
      best-effort ``<sheet_code_name>_<control_name>_<event>`` guess (Excel's
      default document-module event handler naming); it is **not** a confirmed
      mapping and must never be reported as one. ``macro_mapping`` is
      ``PROBABLE`` for such an inferred association and ``UNKNOWN`` when no
      candidate can be formed.
    """

    name: str
    kind: ControlKind
    sheet_name: str | None = None
    shape_id: str | int | None = None
    prog_id: str | None = None
    control_type: ControlFormType | None = Field(
        default=None, description="Form Control subtype (button/checkbox/...)."
    )
    caption: str | None = None
    text: str | None = None
    linked_cell: str | None = None
    on_action: str | None = None
    visible: bool | None = None
    top: float | None = None
    left: float | None = None
    width: float | None = None
    height: float | None = None
    macro_mapping: MacroMappingStatus = MacroMappingStatus.UNKNOWN
    probable_procedure: str | None = Field(
        default=None,
        description="Inferred ActiveX event procedure (probable, not confirmed).",
    )
    notes: list[str] = Field(
        default_factory=list,
        description="Diagnostics for properties that could not be read.",
    )


class ModelSection(DomainModel):
    """Availability + diagnostics for one workbook-model section (spec §7.1, A03).

    The Stage 1 model is an aggregation of the existing inventory services. Each
    section records whether it was read (``status``), how many items it contains
    (``count``) and any per-section diagnostics (never a COM object).
    """

    name: str
    status: ModelSectionStatus
    count: int = 0
    diagnostics: list[str] = Field(default_factory=list)


class WorkbookModel(TimestampedModel):
    """Compact, paginated Phase 1 workbook model (spec §6.1, §7.1, Appendix A).

    This is an **aggregating** view over the Stage 1 inventory services (sheets,
    tables, names, shapes, controls and the VBA project); it is not a second
    source of object data and it never carries all-cells / Range data.

    ``vba_access`` reports the VBA project state without aborting the rest of the
    model: when VBProject is unavailable, locked or VBOM-trust is disabled the
    non-VBA sections are still returned. ``model_hash`` is a SHA-256 over the
    model content (excluding the timestamps and the hash itself) so two reads of
    an unchanged workbook produce the same hash.
    """

    schema_version: int = SCHEMA_VERSION
    workbook: WorkbookInfo
    sheets: list[SheetInfo] = Field(default_factory=list)
    tables: list[TableInfo] = Field(default_factory=list)
    vba_project: VbaProjectInfo | None = None
    vba_access: VbaAccessStatus = VbaAccessStatus.AVAILABLE
    vba_procedures: list[VbaProcedureInfo] = Field(default_factory=list)
    names: list[DefinedNameInfo] = Field(default_factory=list)
    shapes: list[ShapeInfo] = Field(default_factory=list)
    controls: list[ControlEntry] = Field(default_factory=list)
    sections: dict[str, ModelSection] = Field(default_factory=dict)
    model_version: int = Field(default=1)
    model_hash: str | None = None


class SheetModel(TimestampedModel):
    """Sheet-scoped Stage 1 model (spec §7.1, §7.5, A03).

    Sheet metadata plus the tables, names, shapes and controls scoped to that
    sheet, aggregated from the same inventory services as :class:`WorkbookModel`.
    Each scoped section carries its own status/diagnostics. The sheet model never
    carries all-cells / Range data.
    """

    schema_version: int = SCHEMA_VERSION
    workbook_id: str
    sheet: SheetInfo
    tables: list[TableInfo] = Field(default_factory=list)
    names: list[DefinedNameInfo] = Field(default_factory=list)
    shapes: list[ShapeInfo] = Field(default_factory=list)
    controls: list[ControlEntry] = Field(default_factory=list)
    pivot_tables: list[PivotTableInfo] = Field(
        default_factory=list,
        description="PivotTable inventory metadata scoped to this sheet (task_048); "
        "metadata only, never the PivotTable cell output.",
    )
    sections: dict[str, ModelSection] = Field(default_factory=dict)
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


# ===========================================================================
# Phase 2 — Ranges & Cells and PivotTables contracts (task_040)
# ===========================================================================
# These DTOs extend the Stage 1 domain additively. They are derived from the
# Phase 2 research decision records (task_038: docs/research/ranges_cells_api.md;
# task_039: docs/research/pivot_tables_api.md) and are JSON-serializable with no
# COM references. No Phase 2 cell matrix is ever unbounded: bulk requests are
# capped by the constants below and enforced by request validators.

# Bounded Range limits (task_038 §3.2). Cell/row/column/byte ceilings shared by
# the Range read/write requests and the STA-worker adapter.
MAX_RANGE_CELLS = 100_000
MAX_RANGE_ROWS = 100_000
MAX_RANGE_COLS = 512  # practical default; the hard Excel ceiling is XFD (16,384)
MAX_RANGE_COLS_HARD = 16_384
MAX_RESPONSE_CELLS = 50_000
MAX_RESPONSE_BYTES = 8 * 1024 * 1024
MAX_WRITE_PAYLOAD_CELLS = 100_000
MAX_VALIDATION_FORMULA_LENGTH = 255
# Upper bound on the cells a single PivotTable rendered-result page returns. The
# rendered output is read through the shared bounded Range reader, so a large
# PivotTable is paged rather than dumped across the worker boundary.
MAX_PIVOT_RESULT_CELLS = 50_000


# ---------------------------------------------------------------------------
# Range address / metadata
# ---------------------------------------------------------------------------


class RangeAddress(DomainModel):
    """Resolved target address of a Range (task_038 §8.2).

    ``rows``/``cols`` are the target dimensions; ``shape`` is the normalized
    value-block shape (:class:`RangeShape`). ``areas`` is the number of areas
    (multi-area targets are rejected upstream). ``is_merged`` / ``merged_area``
    report merged-cell geometry; ``in_table`` / ``table_name`` flag a target that
    intersects an Excel Table (ListObject). ``spilling`` / ``spill_parent`` carry
    dynamic-array spill diagnostics when available.
    """

    sheet: str
    address: str
    rows: int = Field(ge=0)
    cols: int = Field(ge=0)
    shape: RangeShape = RangeShape.BLOCK
    areas: int = Field(default=1, ge=1)
    is_merged: bool = False
    merged_area: str | None = None
    in_table: bool | None = None
    table_name: str | None = None
    spilling: bool | None = None
    spill_parent: str | None = None


class RangeMetadata(DomainModel):
    """Compact Range metadata (task_040 required surface).

    ``dimensions`` is a human-readable ``"<rows>x<cols>"`` string;
    ``cell_count`` is ``rows * cols``.
    """

    sheet: str
    address: str
    rows: int = Field(ge=0)
    cols: int = Field(ge=0)
    dimensions: str
    cell_count: int


class PageInfo(DomainModel):
    """Large-data completeness + version block (additive; large-data contract §2).

    Attached to paged reads so an agent never mistakes partial data for complete
    data. ``version`` is the workbook content version the page belongs to;
    ``source`` is ``"live"`` (open workbook), ``"snapshot"`` (immutable SaveCopyAs
    file) or ``"cache"`` (saved PivotCache records). ``reason`` is set only when
    the page is empty for a reason other than ``no_data``.
    """

    offset: int = Field(ge=0, description="0-based row/item offset of the first returned item.")
    limit: int | None = Field(default=None, description="Requested cap (None = server default).")
    returned: int = Field(ge=0, description="Items actually returned.")
    total: int | None = Field(
        default=None,
        description="Total available when safely known, else None.",
    )
    has_more: bool = Field(default=False, description="More items exist beyond this page.")
    complete: bool = Field(
        default=False,
        description="True iff every item has been delivered (returned == total).",
    )
    truncated: bool = Field(
        default=False,
        description="True iff a single response was cut by a limit.",
    )
    version: int = Field(default=0, ge=0, description="Workbook content version of this page.")
    source: Literal["live", "snapshot", "cache"] = Field(
        default="live",
        description="Data origin: live workbook, immutable snapshot, or saved PivotCache.",
    )
    reason: Literal[
        "no_data",
        "unsupported_source",
        "stale_cursor",
        "invalid_page",
        "read_failed",
    ] | None = Field(
        default=None,
        description="Set only when the page is empty for a reason other than 'no data'.",
    )


# ---------------------------------------------------------------------------
# Range values / formulas (2-D normalized)
# ---------------------------------------------------------------------------


class RangeValues(DomainModel):
    """Typed, 2-D-normalized Range values (task_038 §1.4).

    ``values`` holds one inner list per row, each aligned to ``cols``. Every
    cell is a typed :class:`CellValue` (never a lossy string coercion). The
    returned grid is bounded by :data:`MAX_RESPONSE_CELLS`; ``truncated`` marks
    a response cut short by an offset/limit window.
    """

    range: RangeAddress
    shape: RangeShape
    rows: int
    cols: int
    values: list[list[CellValue]]
    truncated: bool = False
    notes: list[str] = Field(default_factory=list)
    page: PageInfo | None = Field(
        default=None,
        description="Large-data completeness/version block (additive); None on legacy paths.",
    )

    @model_validator(mode="after")
    def _check_response_bounds(self) -> Self:
        if self.rows * self.cols > MAX_RESPONSE_CELLS:
            raise ValueError(
                f"Range response {self.rows}x{self.cols} exceeds "
                f"MAX_RESPONSE_CELLS ({MAX_RESPONSE_CELLS})."
            )
        if len(self.values) != self.rows:
            raise ValueError("RangeValues.values row count must equal rows.")
        return self


class FormulaCell(DomainModel):
    """One cell's formula data (task_038 §2.4).

    ``formula2`` is the primary Dynamic-Array formula-bar form;
    ``formula`` is the pre-DA (IIE) form when it differs. ``is_array`` /
    ``saved_as_array`` report legacy array participation.
    """

    address: str
    formula2: str
    formula: str | None = None
    is_array: bool | None = None
    saved_as_array: bool | None = None
    notes: list[str] = Field(default_factory=list)


class RangeFormulas(DomainModel):
    """Typed, 2-D-normalized Range formulas (task_038 §2.4)."""

    range: RangeAddress
    shape: RangeShape
    rows: int
    cols: int
    formulas: list[list[FormulaCell]]
    truncated: bool = False
    notes: list[str] = Field(default_factory=list)
    page: PageInfo | None = Field(
        default=None,
        description="Large-data completeness/version block (additive); None on legacy paths.",
    )

    @model_validator(mode="after")
    def _check_response_bounds(self) -> Self:
        if self.rows * self.cols > MAX_RESPONSE_CELLS:
            raise ValueError(
                f"Range formula response {self.rows}x{self.cols} exceeds "
                f"MAX_RESPONSE_CELLS ({MAX_RESPONSE_CELLS})."
            )
        if len(self.formulas) != self.rows:
            raise ValueError("RangeFormulas.formulas row count must equal rows.")
        return self


# ---------------------------------------------------------------------------
# Range formatting (bounded, tri-state)
# ---------------------------------------------------------------------------


class FontFormat(DomainModel):
    """Bounded font formatting (task_038 §4.2).

    Each field is tri-state: a value, ``None`` (mixed across the range — also
    listed in ``FormatInfo.mixed_properties``), or absent from the DTO (not
    represented). ``color`` is the OLE ``RGB`` integer.
    """

    name: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    size: float | None = None
    color: int | None = None


class FillFormat(DomainModel):
    """Bounded interior fill formatting (task_038 §4.2). ``color`` is an RGB int."""

    color: int | None = None
    pattern: str | None = None


class AlignmentFormat(DomainModel):
    """Bounded alignment formatting (task_038 §4.2)."""

    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool | None = None


class FormatInfo(DomainModel):
    """Typed, bounded Range formatting (task_038 §4.2).

    Tri-state on read: each value is ``None`` when cells disagree. The property
    names that returned mixed are listed in ``mixed_properties``; properties that
    exist but are not represented by this DTO are listed in ``unsupported``
    (never silently flattened).
    """

    range: RangeAddress
    number_format: str | None = None
    font: FontFormat = Field(default_factory=FontFormat)
    fill: FillFormat = Field(default_factory=FillFormat)
    alignment: AlignmentFormat = Field(default_factory=AlignmentFormat)
    mixed_properties: list[str] = Field(default_factory=list)
    unsupported: list[str] = Field(default_factory=list)


class RangeFormatPatch(DomainModel):
    """Sparse, partial Range format update (task_038 §4.3).

    Every field is optional; only fields explicitly set by the caller are
    written (use :meth:`applied_fields` to distinguish "set" from "left
    untouched" via Pydantic ``model_fields_set``).
    """

    number_format: str | None = None
    font_name: str | None = None
    font_bold: bool | None = None
    font_italic: bool | None = None
    font_size: float | None = None
    font_color: int | None = None
    fill_color: int | None = None
    fill_pattern: str | None = None
    alignment_horizontal: str | None = None
    alignment_vertical: str | None = None
    alignment_wrap_text: bool | None = None

    def applied_fields(self) -> list[str]:
        """Return the sorted names of the fields the caller explicitly set."""

        return sorted(self.model_fields_set)


class FormatMutationResult(DomainModel):
    """Result of ``range_set_format`` (task_038 §4.4)."""

    action: str = "set_format"
    range: RangeAddress
    applied: list[str] = Field(default_factory=list)
    ignored: list[str] = Field(default_factory=list)
    verified: bool = False
    result_format: FormatInfo | None = None
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Range data validation & conditional formatting (task_038 §6)
# ---------------------------------------------------------------------------


class ValidationRule(DomainModel):
    """Typed ``Range.Validation`` rule (task_038 §6.1).

    ``present`` is ``False`` when the range has no validation. Unsupported or
    mixed pieces are reported in ``unsupported`` rather than silently dropped.
    """

    present: bool = False
    type: DataValidationType = DataValidationType.NONE
    operator: ValidationOperator | None = None
    formula1: str | None = None
    formula2: str | None = None
    ignore_blank: bool | None = None
    in_cell_dropdown: bool | None = None
    apply_to: str | None = None
    alert_style: str | None = None
    error_message: str | None = None
    error_title: str | None = None
    input_message: str | None = None
    input_title: str | None = None
    show_error: bool | None = None
    show_input: bool | None = None
    unsupported: list[str] = Field(default_factory=list)


class ValidationSpec(DomainModel):
    """Request payload for ``range_set_validation`` (task_038 §6.1).

    Enforces the documented 255-character ``Formula1``/``Formula2`` limit before
    any COM mutation. Formulas are opaque strings passed to Excel; never
    evaluated in Python.
    """

    type: DataValidationType
    operator: ValidationOperator | None = None
    formula1: str | None = None
    formula2: str | None = None
    ignore_blank: bool | None = None
    in_cell_dropdown: bool | None = None
    alert_style: str | None = None
    error_message: str | None = None
    error_title: str | None = None
    input_message: str | None = None
    input_title: str | None = None
    show_error: bool | None = None
    show_input: bool | None = None

    @model_validator(mode="after")
    def _check_formula_lengths(self) -> Self:
        for name in ("formula1", "formula2"):
            value = getattr(self, name)
            if value is not None and len(value) > MAX_VALIDATION_FORMULA_LENGTH:
                raise ValueError(
                    f"{name} exceeds the {MAX_VALIDATION_FORMULA_LENGTH}-character limit."
                )
        return self


class ConditionalFormatRule(DomainModel):
    """Typed conditional-formatting rule (task_038 §6.3).

    ``supported`` is ``False`` for ``ColorScale`` / ``DataBar`` / ``IconSet``
    (distinct COM objects outside the bounded Phase 2 scope); ``unsupported``
    lists those types rather than silently omitting them.
    """

    priority: int = Field(ge=1)
    type: ConditionalFormatType
    operator: ValidationOperator | None = None
    formula1: str | None = None
    formula2: str | None = None
    stop_if_true: bool | None = None
    applies_to: str | None = None
    interior_color: int | None = None
    font: FontFormat | None = None
    supported: bool = True
    unsupported_reason: str | None = None


class ConditionalFormatSpec(DomainModel):
    """Request payload for one conditional-formatting rule (task_038 §6.3).

    ``type`` is the pruned request vocabulary (:class:`ConditionalFormatRequestType`):
    ``color_scale`` / ``data_bar`` / ``icon_set`` / ``unsupported`` are rejected
    here (pydantic) and are never sent to Excel. Read/classify uses the full
    :class:`ConditionalFormatType` (:class:`ConditionalFormatRule`).
    """

    type: ConditionalFormatRequestType
    operator: ValidationOperator | None = None
    formula1: str | None = None
    formula2: str | None = None
    stop_if_true: bool | None = None
    applies_to: str | None = None
    interior_color: int | None = None
    font: FontFormat | None = None
    priority: int | None = Field(default=None, ge=1)


class ConditionalFormatMutationResult(DomainModel):
    """Result of ``range_set_conditional_formatting`` (priority preserved)."""

    action: str
    range: RangeAddress
    rules: list[ConditionalFormatRule] = Field(default_factory=list)
    applied_priorities: list[int] = Field(default_factory=list)
    unsupported: list[str] = Field(default_factory=list)
    verified: bool = False
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# Range write request (bounded bulk) + mutation results
# ---------------------------------------------------------------------------


#: A raw JSON scalar accepted as a literal cell value by
#: :class:`RangeWriteRequest` (WP-Range ergonomics). Coercion to
#: :class:`CellValue` is **unambiguous** and driven only by the real JSON type:
#: a JSON string -> ``string`` (a leading ``=`` stays literal text; formula
#: interpretation is exclusive to :class:`FormulaWriteRequest`), a JSON number
#: (int/float, never bool) -> ``number``, a JSON bool -> ``boolean``, and JSON
#: ``null`` -> ``empty``. Strings such as ``"true"`` / ``"123"`` / ``"2024-01-01"``
#: are **not** coerced to bool/number/date — they stay strings.
ScalarValue = str | int | float | bool | None


def coerce_cell_value(raw: object) -> CellValue | dict[str, Any]:
    """Coerce one raw cell input to a :class:`CellValue` (unambiguous, WP-Range).

    A :class:`CellValue` passes through unchanged; a ``dict`` is returned as-is so
    Pydantic validates it as a :class:`CellValue` (a malformed dict surfaces as a
    normal validation error). A raw JSON scalar maps by its **real** type only:

    * ``None`` -> ``CellValue{kind: "empty"}`` (an empty cell);
    * ``bool`` -> ``CellValue{kind: "boolean"}`` (checked before ``int`` because
      ``bool`` is a subclass of ``int``);
    * ``int`` / ``float`` -> ``CellValue{kind: "number"}``;
    * ``str`` -> ``CellValue{kind: "string"}`` (a leading ``=`` stays text).

    Any other type (e.g. a nested ``list``) raises ``ValueError`` so the request
    is rejected rather than silently coerced.
    """

    if isinstance(raw, CellValue):
        return raw
    if raw is None:
        return CellValue(kind=CellValueKind.EMPTY)
    if isinstance(raw, bool):
        return CellValue(kind=CellValueKind.BOOLEAN, value=raw)
    if isinstance(raw, (int, float)):
        return CellValue(kind=CellValueKind.NUMBER, value=raw)
    if isinstance(raw, str):
        return CellValue(kind=CellValueKind.STRING, value=raw)
    if isinstance(raw, dict):
        return raw
    raise ValueError(
        "A cell value must be a CellValue object or a raw JSON scalar "
        "(string, number, boolean or null)."
    )


class RangeWriteRequest(DomainModel):
    """Bounded 2-D bulk write payload for values/formulas (task_038 §3.3).

    Rejected before any COM mutation when the grid exceeds the row/column/cell
    ceilings or when ``values`` shape does not match ``rows`` × ``cols``. A
    string beginning with ``=`` is **not** a formula: value writes store text;
    formula writes use :class:`FormulaWriteRequest`.

    Each cell accepts either a typed :class:`CellValue` or a raw JSON scalar
    (WP-Range ergonomics): a string, number, boolean or ``null`` is coerced to a
    literal :class:`CellValue` by real JSON type only (no ambiguous string
    coercion — see :data:`ScalarValue` / :func:`coerce_cell_value`). The
    normalization runs before field validation, so ``values`` always holds
    :class:`CellValue` instances after construction while the input schema still
    advertises the raw-scalar union.
    """

    rows: int = Field(ge=1)
    cols: int = Field(ge=1)
    values: list[list[CellValue | ScalarValue]]

    @model_validator(mode="before")
    @classmethod
    def _normalize_values(cls, data: object) -> object:
        """Coerce raw scalar cells to :class:`CellValue` before field validation."""

        if not isinstance(data, dict):
            return data
        raw_values = data.get("values")
        if not isinstance(raw_values, list):
            return data
        normalized: list[object] = []
        for row in raw_values:
            if isinstance(row, list):
                normalized.append([coerce_cell_value(cell) for cell in row])
            else:
                normalized.append(row)
        return {**data, "values": normalized}

    @model_validator(mode="after")
    def _check_bounds(self) -> Self:
        if self.rows > MAX_RANGE_ROWS:
            raise ValueError(f"rows exceeds MAX_RANGE_ROWS ({MAX_RANGE_ROWS}).")
        if self.cols > MAX_RANGE_COLS:
            raise ValueError(f"cols exceeds MAX_RANGE_COLS ({MAX_RANGE_COLS}).")
        cells = self.rows * self.cols
        if cells > MAX_WRITE_PAYLOAD_CELLS:
            raise ValueError(
                f"write payload {self.rows}x{self.cols}={cells} exceeds "
                f"MAX_WRITE_PAYLOAD_CELLS ({MAX_WRITE_PAYLOAD_CELLS})."
            )
        if len(self.values) != self.rows:
            raise ValueError("values row count must equal rows.")
        for index, row in enumerate(self.values):
            if len(row) != self.cols:
                raise ValueError(f"values[{index}] width must equal cols.")
        return self


class FormulaWriteRequest(DomainModel):
    """Bounded 2-D bulk formula-write payload (task_038 §2.4).

    Each entry must be an actual formula string beginning with ``=``; a bare
    constant belongs in :class:`RangeWriteRequest`.
    """

    rows: int = Field(ge=1)
    cols: int = Field(ge=1)
    formulas: list[list[str]]

    @model_validator(mode="after")
    def _check_bounds_and_formulas(self) -> Self:
        if self.rows > MAX_RANGE_ROWS:
            raise ValueError(f"rows exceeds MAX_RANGE_ROWS ({MAX_RANGE_ROWS}).")
        if self.cols > MAX_RANGE_COLS:
            raise ValueError(f"cols exceeds MAX_RANGE_COLS ({MAX_RANGE_COLS}).")
        cells = self.rows * self.cols
        if cells > MAX_WRITE_PAYLOAD_CELLS:
            raise ValueError(
                f"write payload {self.rows}x{self.cols}={cells} exceeds "
                f"MAX_WRITE_PAYLOAD_CELLS ({MAX_WRITE_PAYLOAD_CELLS})."
            )
        if len(self.formulas) != self.rows:
            raise ValueError("formulas row count must equal rows.")
        for index, row in enumerate(self.formulas):
            if len(row) != self.cols:
                raise ValueError(f"formulas[{index}] width must equal cols.")
        return self


class RangeMutationResult(DomainModel):
    """Before/after metadata + postcondition for a Range value/formula write.

    ``range_before`` / ``range_after`` are the resolved target addresses;
    ``postcondition_values`` holds the re-read values used to verify the write.
    """

    action: str
    range_before: RangeAddress | None = None
    range_after: RangeAddress | None = None
    verified: bool = False
    postcondition_values: RangeValues | None = None
    notes: list[str] = Field(default_factory=list)


class StructuralMutationResult(DomainModel):
    """Result of copy/insert/delete/merge/unmerge (task_038 §5.4).

    ``table_metadata_unchanged`` is ``True`` only when the operation did not
    intersect a ListObject. ``preconditions`` carries the best-effort preflight
    preview (merged cells, non-anchor data, referenced formulas).
    """

    action: str
    source: RangeAddress | None = None
    destination: RangeAddress | None = None
    shift: str | None = None
    range_before: RangeAddress | None = None
    range_after: RangeAddress | None = None
    affected_addresses: list[str] = Field(default_factory=list)
    table_metadata_unchanged: bool = True
    verified: bool = False
    preconditions: dict[str, Any] = Field(default_factory=dict)
    notes: list[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# PivotTables / PivotCaches / PivotFields (task_039 §8)
# ---------------------------------------------------------------------------


class PivotFieldInfo(DomainModel):
    """Bounded ``PivotField`` metadata (task_039 §3, §8).

    ``source_name`` is the stable identity; ``data_field_name`` (the caption) is
    set for data fields. ``orientation`` is the layout bucket; ``position`` is
    ``None`` when the field is hidden / unplaced. ``function`` is set for data
    fields only. ``filter_state`` summarizes the current filter kind.
    """

    source_name: str
    caption: str | None = None
    orientation: PivotFieldOrientation = PivotFieldOrientation.HIDDEN
    position: int | None = None
    function: PivotAggregation | None = None
    show_all_items: bool | None = None
    all_items_visible: bool | None = None
    filter_state: PivotFilterKind = PivotFilterKind.NONE
    is_data_field: bool = False
    data_field_name: str | None = None
    number_format: str | None = None


class PivotCacheInfo(DomainModel):
    """Bounded ``PivotCache`` metadata (task_039 §3, §8).

    ``source_data`` is the text form for database/Table sources;
    ``source_data_kind`` is ``"text" | "array" | "unavailable"``. ``olap`` and
    ``source_type`` drive the unsupported-source policy. ``shared_cache`` and
    ``affected_pivot_tables`` record the N:1 cache relationship.
    """

    cache_index: int
    source_type: PivotSourceType
    source_data: str | None = None
    source_data_kind: str = "text"
    refresh_date: datetime | None = None
    refresh_on_file_open: bool | None = None
    enable_refresh: bool | None = None
    record_count: int | None = None
    olap: bool = False
    shared_cache: bool = False
    affected_pivot_tables: list[str] = Field(default_factory=list)


class PivotTableInfo(DomainModel):
    """Bounded ``PivotTable`` metadata (task_039 §3, §8).

    ``name`` is the stable report key; ``cache_index`` links the table to its
    :class:`PivotCacheInfo`. ``source_status`` is ``"supported"`` or
    ``"unsupported_source"``. Fields are grouped by orientation; this never
    materializes the PivotTable cell output.
    """

    name: str
    sheet_name: str
    sheet_index: int
    index: int
    cache_index: int
    table_range1: str | None = None
    table_range2: str | None = None
    data_body_range: str | None = None
    row_grand: bool | None = None
    column_grand: bool | None = None
    refresh_date: datetime | None = None
    source_type: PivotSourceType = PivotSourceType.UNKNOWN
    source_status: str = "supported"
    data_fields: list[PivotFieldInfo] = Field(default_factory=list)
    row_fields: list[PivotFieldInfo] = Field(default_factory=list)
    column_fields: list[PivotFieldInfo] = Field(default_factory=list)
    page_fields: list[PivotFieldInfo] = Field(default_factory=list)
    hidden_fields: list[PivotFieldInfo] = Field(default_factory=list)


class PivotResultPage(DomainModel):
    """Bounded page of a PivotTable's **rendered** output (WP-C).

    Reads the values Excel actually displays in ``TableRange1`` — never the
    external source — so it works for every source type (database, external,
    OLAP / Data Model). The page is produced by the shared bounded Range reader,
    so ``rows`` / ``cols`` / ``values`` / ``page`` carry the same semantics as a
    paged :class:`RangeValues` read. ``header_row_count`` is the number of rows
    above ``DataBodyRange`` within ``TableRange1`` (``0`` when unknown) and
    ``fingerprint`` is a structural id that changes when the table range or
    refresh date changes.
    """

    pivot_table: str
    sheet_name: str
    source_type: PivotSourceType
    table_range: str
    data_body_range: str | None = None
    refresh_date: datetime | None = None
    header_row_count: int = 0
    rows: int
    cols: int
    values: list[list[CellValue]]
    page: PageInfo
    fingerprint: str

    @model_validator(mode="after")
    def _check_response_bounds(self) -> Self:
        if len(self.values) != self.rows:
            raise ValueError("PivotResultPage.values row count must equal rows.")
        if self.rows * self.cols > MAX_PIVOT_RESULT_CELLS:
            raise ValueError(
                f"Pivot result response {self.rows}x{self.cols} exceeds "
                f"MAX_PIVOT_RESULT_CELLS ({MAX_PIVOT_RESULT_CELLS})."
            )
        return self


class PivotInventory(DomainModel):
    """Full-workbook PivotTable + PivotCache inventory (task_039 §2, §8)."""

    workbook_id: str
    pivot_tables: list[PivotTableInfo] = Field(default_factory=list)
    caches: list[PivotCacheInfo] = Field(default_factory=list)


class PivotFieldLayout(DomainModel):
    """Target orientation/position for a field layout change (task_039 §5).

    ``orientation`` is the pruned request vocabulary
    (:class:`PivotFieldLayoutOrientation`): ``data`` is rejected here (pydantic)
    — use ``pivot_field_set_aggregation`` to place a field in the data area.
    """

    orientation: PivotFieldLayoutOrientation
    position: int | None = None


class PivotFieldProperties(DomainModel):
    """Bounded, sparse PivotField property update (task_039 §5)."""

    caption: str | None = None
    number_format: str | None = None
    show_all_items: bool | None = None


class PivotFilterSpec(DomainModel):
    """Manual (visible-items) filter request (task_039 §6; U3 pruned request).

    ``kind`` is the pruned request vocabulary (:class:`PivotFilterRequestKind`):
    only ``manual`` is accepted. Label / value / date / top-bottom requests are
    rejected here (pydantic) and never reach Excel. The current filter state is
    reported with the full :class:`PivotFilterKind` via :class:`PivotFilterState`.
    """

    kind: PivotFilterRequestKind = PivotFilterRequestKind.MANUAL
    visible_items: list[str] = Field(default_factory=list)
    supported: bool = True
    unsupported_reason: str | None = None


class PivotFilterState(DomainModel):
    """Current filter state of a PivotField (read/classify; full kind set).

    Mirrors :class:`PivotFilterSpec` but keeps the full :class:`PivotFilterKind`
    so a read can report ``label`` / ``value`` / ``date`` / ``top_bottom`` /
    ``unsupported`` / ``none`` filter kinds with ``supported=false`` rather than
    flattening them. Never used as a write request.
    """

    kind: PivotFilterKind = PivotFilterKind.NONE
    visible_items: list[str] = Field(default_factory=list)
    supported: bool = True
    unsupported_reason: str | None = None


class PivotFieldFilters(DomainModel):
    """Current filter state of one PivotField (task_039 §6)."""

    pivot_table: str
    field: str
    all_items_visible: bool | None = None
    visible_items: list[str] = Field(default_factory=list)
    hidden_items: list[str] = Field(default_factory=list)
    filter: PivotFilterState = Field(default_factory=PivotFilterState)
    unsupported: list[str] = Field(default_factory=list)


class PivotRefreshResult(DomainModel):
    """Result of a targeted refresh (task_039 §4).

    A refresh is successful only when ``refreshed`` is ``True`` (Excel returned a
    truth signal and/or ``refresh_date`` advanced) — absence of a COM exception
    alone is not success. ``affected_pivot_tables`` lists every table sharing the
    cache, regardless of which entry point triggered the refresh.
    """

    pivot_table: str
    cache_index: int
    method: PivotRefreshMethod
    affected_pivot_tables: list[str] = Field(default_factory=list)
    refreshed: bool = False
    refresh_return: bool | None = None
    started_at: datetime | None = None
    finished_at: datetime | None = None
    duration_ms: int | None = None
    refresh_date_before: datetime | None = None
    refresh_date_after: datetime | None = None
    status: PivotRefreshStatus = PivotRefreshStatus.FAILED
    error: dict[str, Any] | None = None


class SaveAsResult(TimestampedModel):
    """Result of ``excel_save_as`` (U7; plan_05 §5.5, §9).

    Fail-closed: ``outcome`` is ``"ok"`` only after the final size + SHA-256 of
    the target were verified. An interrupted or unprovable write reports
    ``"unknown"`` / ``"error"`` (never success) and, when a previous target was
    preserved, carries its ``backup_path`` (a single ``<target>.bak``). The
    session stays bound to its working copy (``wb.Save`` + a pure-filesystem
    copy, never ``wb.SaveAs``): ``workbook_id`` / working copy / source are
    unchanged by a Save As.
    """

    workbook_id: str
    target_path: str = Field(
        description="Canonical target path written (== source_path on a source-overwrite)."
    )
    source_path: str = Field(
        description="Canonical original source path of the open working copy."
    )
    source_overwritten: bool = Field(
        description=(
            "True when the original source file was overwritten "
            "(overwrite_source=True and the source hash matched)."
        )
    )
    overwrote_existing: bool = Field(
        description="True when a pre-existing target file was replaced (a .bak was written)."
    )
    backup_path: str | None = Field(
        default=None,
        description="Path of the single <target>.bak preserved before an overwrite, or None.",
    )
    sha256: str = Field(description="SHA-256 of the written target (verified).")
    size: int = Field(description="Size in bytes of the written target (verified).")
    outcome: Literal["ok", "unknown", "error"] = Field(
        description="'ok' only after the final size+SHA-256 verify; 'unknown'/'error' fail-closed."
    )
    mode: SessionMode = Field(description="Session mode of the bound working copy.")
    notes: list[str] = Field(default_factory=list)


# ===========================================================================
# Non-COM OOXML snapshot reads (WP-E1)
# ===========================================================================
# These DTOs describe a read-only, streaming read of worksheet rows / workbook
# metadata from an immutable snapshot file (a Workbook.SaveCopyAs image), without
# occupying the single-STA COM worker. They are JSON-serializable and carry no
# COM references. A page is bounded by ``MAX_RESPONSE_CELLS`` just like a Range
# read, and every result records the snapshot SHA-256 it was read from.


class OoxmlInputInfo(DomainModel):
    """Provenance of a non-COM OOXML read (WP-E1).

    ``input_sha256`` is the SHA-256 of the immutable snapshot file the data was
    parsed from; ``snapshot_id`` is the snapshot used (the freshly created one,
    or the caller-supplied one). ``source`` is always ``"snapshot"`` — the live
    working copy and the original are never read.
    """

    input_sha256: str
    snapshot_id: str | None = None
    source: Literal["snapshot"] = "snapshot"


class OxmlSheetMeta(DomainModel):
    """One worksheet's structural metadata from the snapshot (WP-E1).

    ``part`` is the package part path (e.g. ``xl/worksheets/sheet1.xml``);
    ``hidden`` is ``True`` for hidden / very-hidden sheets; ``rows_hint`` is an
    optional best-effort row count (``None`` when not cheaply known).
    """

    name: str
    part: str
    hidden: bool = False
    rows_hint: int | None = None


class OoxmlMetadata(DomainModel):
    """Workbook metadata read non-COM from an immutable snapshot (WP-E1).

    Lists the sheets (name/part/hidden/order via :class:`OxmlSheetMeta`), the
    workbook defined names and the table-part count. Never carries cell data.
    """

    workbook_id: str
    input: OoxmlInputInfo
    sheets: list[OxmlSheetMeta] = Field(default_factory=list)
    defined_names: list[str] = Field(default_factory=list)
    table_count: int = 0


class OoxmlSheetPage(DomainModel):
    """A bounded page of typed worksheet rows read from a snapshot (WP-E1).

    ``values`` holds one inner list per row, each aligned by column index; every
    cell is a typed :class:`CellValue`. The grid is bounded by
    :data:`MAX_RESPONSE_CELLS` and ``len(values) == rows``. ``page.source`` is
    ``"snapshot"`` and ``page.version`` is ``0`` (a snapshot is immutable).
    """

    workbook_id: str
    sheet: str
    input: OoxmlInputInfo
    rows: int
    cols: int
    values: list[list[CellValue]]
    page: PageInfo

    @model_validator(mode="after")
    def _check_response_bounds(self) -> Self:
        if len(self.values) != self.rows:
            raise ValueError("OoxmlSheetPage.values row count must equal rows.")
        if self.rows * self.cols > MAX_RESPONSE_CELLS:
            raise ValueError(
                f"OOXML sheet response {self.rows}x{self.cols} exceeds "
                f"MAX_RESPONSE_CELLS ({MAX_RESPONSE_CELLS})."
            )
        return self


class OoxmlPivotCacheMeta(DomainModel):
    """Freshness / completeness metadata of a saved PivotCache (WP-E2).

    Describes the **saved** cache (not the PivotTable's displayed result and not
    the live source): ``record_count`` is the definition's ``recordCount`` (when
    present), ``save_data`` is the ``saveData`` flag (``False`` means the cache
    stored no local records), ``refresh_on_load`` / ``refresh_date`` report the
    cache's refresh freshness, and ``has_records_part`` is ``True`` only when a
    ``pivotCacheRecords`` part is actually present in the snapshot.
    """

    cache_index: int
    record_count: int | None = None
    save_data: bool | None = None
    refresh_on_load: bool | None = None
    refresh_date: datetime | None = None
    has_records_part: bool = False


class OoxmlPivotCachePage(DomainModel):
    """A bounded page of typed **saved** PivotCache records (WP-E2).

    ``values`` holds one inner list per record, field-aligned; every cell is a
    typed :class:`CellValue` (shared-item indices resolved against the cache's
    shared-items table). The grid is bounded by :data:`MAX_RESPONSE_CELLS` and
    ``len(values) == rows``. ``page.source`` is ``"cache"`` and ``page.version`` is
    ``0`` (a snapshot is immutable). When the cache has no local records
    (``saveData="0"`` / OLAP / Data Model) the page is empty with
    ``page.reason == "unsupported_source"`` and ``meta.has_records_part`` is
    ``False`` — the records are never fabricated.
    """

    workbook_id: str
    cache_index: int
    input: OoxmlInputInfo
    meta: OoxmlPivotCacheMeta
    rows: int
    cols: int
    values: list[list[CellValue]]
    page: PageInfo

    @model_validator(mode="after")
    def _check_response_bounds(self) -> Self:
        if len(self.values) != self.rows:
            raise ValueError("OoxmlPivotCachePage.values row count must equal rows.")
        if self.rows * self.cols > MAX_RESPONSE_CELLS:
            raise ValueError(
                f"OOXML pivot-cache response {self.rows}x{self.cols} exceeds "
                f"MAX_RESPONSE_CELLS ({MAX_RESPONSE_CELLS})."
            )
        return self


# ===========================================================================
# Large-data file export (WP-F)
# ===========================================================================
# A data export streams a large dataset (worksheet Range, Excel Table or
# PivotTable rendered result) to a file artifact (CSV / JSONL) with bounded
# memory, returning only compact metadata — never the data itself. The DTO is
# JSON-serializable and carries no COM references.


class DataExportResult(DomainModel):
    """Compact result of a large-data file export (WP-F).

    ``path`` is the stored artifact file (CSV / JSONL); the data is never returned
    inline. ``rows`` is the number of data rows written (excluding a CSV header),
    ``bytes`` / ``sha256`` describe the written file, ``complete`` is ``True`` iff
    every source row was delivered, ``truncated`` is ``True`` iff the export was cut
    by ``limit_rows`` or the ``max_export_rows`` ceiling, ``source`` is the source
    kind (``"range"`` / ``"table"`` / ``"pivot-result"``) and ``source_version`` is
    the workbook content version of the last page read.
    """

    artifact_id: str
    path: str
    format: Literal["csv", "jsonl"]
    rows: int
    bytes: int
    sha256: str
    complete: bool
    source: str
    source_version: int
    truncated: bool = False
