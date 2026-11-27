"""MCP server entry point for the Excel VBA MCP server (STDIO transport).

The module builds the STDIO :class:`MCPServer` and registers the base
workbook/session tools (task_012). A single :class:`SessionService` runtime is
owned by the server lifespan (ARCH-005): the dedicated single-STA
:class:`ExcelComWorker` and its hidden :class:`ExcelAppAdapter` are created once,
the worker is started on lifespan entry, and ``await worker.shutdown()`` runs in
``finally`` on exit. All COM lives behind the worker; the tools never touch COM.

The runtime is published to :mod:`excel_vba_mcp.api.workbook_tools` via
``set_session_service`` so the (stateless) tool handlers can reach it.

Configuration (roots, audit log, work dir) is read from environment variables so
this module stays dependency-free; extending ``config.py`` with first-class
settings is deferred to the configuration task. See ``_build_runtime``.
"""

from __future__ import annotations

import dataclasses
import json
import logging
import os
import sys
import tempfile
from collections.abc import AsyncIterator, Callable
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import Context, MCPServer
from mcp.server.mcpserver.exceptions import ResourceError, ToolError
from mcp.types import CallToolResult, InputRequiredResult, ToolAnnotations
from pydantic import ValidationError

from excel_vba_mcp import __version__
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import (
    compile_tools,
    control_tools,
    environment_tools,
    job_tools,
    macro_tools,
    name_tools,
    pivot_tools,
    range_tools,
    shape_tools,
    sheet_tools,
    table_tools,
    test_tools,
    vba_tools,
    workbook_tools,
)
from excel_vba_mcp.api.compile_tools import (
    vba_compile_diagnostics,
)
from excel_vba_mcp.api.control_tools import (
    control,
)
from excel_vba_mcp.api.environment_tools import xlvba_environment
from excel_vba_mcp.api.health_tools import excel_health
from excel_vba_mcp.api.job_tools import job_get, job_request_cancel
from excel_vba_mcp.api.macro_tools import (
    macro,
    macro_run,
)
from excel_vba_mcp.api.name_tools import (
    name,
    name_delete,
    name_read,
)
from excel_vba_mcp.api.pivot_tools import (
    pivot_field,
    pivot_read,
    pivot_refresh,
)
from excel_vba_mcp.api.range_tools import (
    range_delete,
    range_merge,
    range_read,
    range_structural,
    range_write,
)
from excel_vba_mcp.api.recovery_tools import excel_recover
from excel_vba_mcp.api.shape_tools import (
    shape,
)
from excel_vba_mcp.api.sheet_tools import (
    sheet,
    sheet_delete,
    sheet_read,
)
from excel_vba_mcp.api.table_tools import (
    table,
    table_delete_column,
    table_delete_rows,
    table_read,
)
from excel_vba_mcp.api.test_tools import (
    snapshot_create,
    snapshot_diff,
    snapshot_list,
    snapshot_restore,
    test,
)
from excel_vba_mcp.api.vba_tools import (
    vba_export_project,
    vba_insert_code,
    vba_read,
    vba_replace_component_source,
)
from excel_vba_mcp.api.workbook_tools import (
    excel_save_as,
    workbook,
    workbook_read,
)
from excel_vba_mcp.application.compile_service import CompileService
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.job_registry import JobRegistry
from excel_vba_mcp.application.macro_service import BackgroundSafetyRegistry, MacroService
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.test_runner import TestRunner
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.application.workbook_model_service import WorkbookModelService
from excel_vba_mcp.config import Config, load_config
from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ErrorCode,
    ExcelVbaMcpError,
    VbaProjectLockedError,
)
from excel_vba_mcp.domain.statuses import VbaAccessStatus
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.artifact_store import ARTIFACT_ROOT_NAME, ArtifactStore
from excel_vba_mcp.storage.job_store import JobStore
from excel_vba_mcp.storage.owned_pid_store import (
    STATE_CREATING,
    OwnedPidStore,
    OwnedProcessRecord,
    build_creating_record,
    build_owned_record,
    default_owned_pid_dir,
)
from excel_vba_mcp.storage.run_repository import RunRepository
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

logger = logging.getLogger("excel_vba_mcp.server")

# Environment variable names for the runtime configuration.
_ENV_ROOTS = "EXCEL_VBA_MCP_ROOTS"
_ENV_AUDIT = "EXCEL_VBA_MCP_AUDIT_LOG"
_ENV_WORK_DIR = "EXCEL_VBA_MCP_WORK_DIR"
_ENV_SNAPSHOT_DIR = "EXCEL_VBA_MCP_SNAPSHOT_DIR"
_ENV_ARTIFACT_DIR = "EXCEL_VBA_MCP_ARTIFACT_DIR"
_ENV_RUN_DIR = "EXCEL_VBA_MCP_RUN_DIR"
_ENV_JOB_DIR = "EXCEL_VBA_MCP_JOB_DIR"
_ENV_OWNED_PID_DIR = "EXCEL_VBA_MCP_OWNED_PID_DIR"

# Tool descriptions surfaced over MCP (kept explicit for a stable contract).
_DESC_OPEN = (
    "Purpose: Open an allowlisted workbook as a hidden working copy. "
    "Use when: you need to read or mutate a workbook through any other tool. "
    "Do not use when: the workbook is already open — reuse its workbook_id "
    "(excel_list_workbooks); for read-only work prefer mode='inspect'. "
    "Preconditions: path must be under an allowlisted root (EXCEL_VBA_MCP_ROOTS, "
    "default cwd + temp); the original file is never modified. "
    "Parameters: path; mode='execute' (default; applies macro-by-UI) or "
    "'inspect' (macros disabled); read_only=False (default; writable unless "
    "read_only=true); background=False. "
    "Behavior: copies the file into work/ and opens it hidden; in restricted "
    "mode execute/writable needs approval; background=true returns a RunningJob "
    "handle. "
    "Result: WorkbookInfo, or RunningJob when background=true. "
    "Next actions: snapshot_create before any mutation, then the read/write tools. "
    "Example: {\"path\": \"C:/data/report.xlsm\", \"mode\": \"execute\", "
    "\"read_only\": false}"
)
_DESC_CLOSE = (
    "close refuses by default when the workbook has unsaved changes: a clean "
    "workbook closes normally, while a dirty workbook is refused unless "
    "save_changes=true (save then close) or discard_changes=true (discard) is "
    "passed; save_changes and discard_changes together are rejected as "
    "contradictory. "
)
_DESC_WORKBOOK = (
    "Purpose: Open, close or save an open workbook through one lifecycle tool. "
    "Use when: you need to open an allowlisted workbook as a hidden working "
    "copy, close a workbook you are done with, or persist the working copy in "
    "place. "
    "Do not use when: you need Save As to a new path — use excel_save_as; the "
    "workbook is already open — reuse its workbook_id (workbook_read "
    "action='list'). "
    "Preconditions: an open workbook_id for close/save; path under an allowlisted "
    "root (EXCEL_VBA_MCP_ROOTS, default cwd + temp) for open; the original file "
    "is never modified. "
    "Parameters: action='open'|'close'|'save'; path (open); workbook_id "
    "(close/save); mode='execute' (default) or 'inspect'; read_only; "
    "save_changes; discard_changes; background. "
    "Behavior: open copies the file into work/ and opens it hidden; save writes "
    "the working copy in place (no path argument); " + _DESC_CLOSE + "Result: "
    "WorkbookInfo (open/save), SessionStatus (close), or RunningJob when "
    "background=true. "
    "Next actions: workbook_read action='status' to confirm state; "
    "snapshot_create before further mutations. "
    "Example: {\"action\": \"open\", \"path\": \"C:/data/report.xlsm\"}"
)
_DESC_WORKBOOK_READ = (
    "Purpose: Read workbook/session state without mutating the workbook. "
    "Use when: you need session/workbook status, the list of open workbooks, or "
    "the aggregated Stage 1 workbook model. "
    "Do not use when: you need to open/close/save — use workbook; you need VBA "
    "source — use vba_read. "
    "Preconditions: an open workbook_id for model (required) and status "
    "(optional); none for list. "
    "Parameters: action='status'|'list'|'model'; workbook_id; include_vba. "
    "Behavior: read-only; status/list never touch COM and model never executes "
    "VBA — when VBA is unavailable the other sections are still returned and "
    "vba_access reports the state. "
    "Result: SessionStatus (status), list[WorkbookInfo] (list), WorkbookModel "
    "(model). "
    "Next actions: workbook action='open' to open a workbook; excel_health for "
    "worker state. "
    "Example: {\"action\": \"model\", \"workbook_id\": \"<id>\"}"
)
_DESC_SAVE = (
    "Purpose: Save an open workbook's working copy in place. "
    "Use when: you want to persist mutations to the working-copy file. "
    "Do not use when: you need a new output path — use excel_save_as instead "
    "(see xlvba://glossary / docs 09). "
    "Preconditions: an open, writable workbook_id. "
    "Behavior: writes the working copy in place (no path argument). "
    "Result: WorkbookInfo. "
    "Next actions: snapshot_create if you plan further mutations."
)
_DESC_SAVE_AS = (
    "Purpose: Save an open workbook's working copy to a (possibly new) target path without "
    "rebinding the session. Use when: you need the working copy written to a different/new "
    "file (Save As), including overwriting the original source. Do not use when: you only need "
    "to persist the working copy in place — use workbook action='save'; the target is the open "
    "working copy (rejected); the source changed since open (STALE_TARGET). Preconditions: an open "
    "workbook_id; MUTATING approval on the target Approval gate (see xlvba://glossary). "
    "Parameters: workbook_id, path, overwrite_existing=false, overwrite_source=false. "
    "Behavior: flushes the working copy with wb.Save on the STA worker, then a pure-filesystem "
    "atomic temp->replace copy (never wb.SaveAs, so the session stays bound to its working "
    "copy and workbook_id/working copy/source are unchanged); a single <target>.bak preserves "
    "any previous target; an existing target needs overwrite_existing=true, the original "
    "source needs overwrite_source=true plus a matching source hash; skips the snapshot gate "
    "(Save As does not mutate the open working copy). Result: SaveAsResult (outcome 'ok' only "
    "after a final size+SHA-256 verify; 'unknown'/'error' fail-closed with backup_path). Next "
    "actions: workbook action='save' to persist the working copy; snapshot_create before further "
    "mutations."
)
_DESC_STATUS = (
    "Purpose: Return session/workbook status (busy, read-only, dirty, active operation). "
    "Use when: you need to check a workbook's state before acting. "
    "Preconditions: an open workbook_id (or omit for the server session). "
    "Behavior: reports access_mode; the additive policy block (U2) co-locates "
    "per-workbook approval/snapshot context. "
    "Result: SessionStatus. "
    "Next actions: excel_health for worker state; xlvba_environment for access "
    "mode / policy / build."
)
_DESC_LIST = "List currently open workbooks from session bookkeeping."
_DESC_HEALTH = (
    "Return worker/session health (state, accepting, owned Excel pid, active "
    "operation, needs_recovery) without forcing a re-run or touching COM. "
    "Read-only: use it to check whether the session needs recovery."
)
_DESC_EXCEL_RECOVER = (
    "Purpose: Recover the owned Excel process (agent-directed). "
    "Use when: excel_health shows needs_recovery or a call is wedged. "
    "Do not use when: the server should not terminate a live pid on a deadline — "
    "it never auto-terminates; call this only when recovery is justified. "
    "Preconditions: a held, identity-verified owned process handle (RECOVERY class; "
    "see xlvba://glossary). "
    "Behavior: terminates the wedged owned EXCEL.EXE through the held handle "
    "(TerminateProcess + WaitForSingleObject), discards the working copies and "
    "rebuilds a fresh hidden instance; auto-allowed in full and restricted mode, "
    "audited, pid-scoped; fail-closed if no live handle is held. "
    "Discards the open working copies (originals are never modified) and does not "
    "roll back external side effects of any macro that already ran. "
    "Requires `confirm=true`; a missing/false confirm is rejected with "
    "CONFIRMATION_REQUIRED before any termination. "
    "Result: a recovery report; an interrupted operation is reported "
    "prior_job='unknown' and is never auto-retried. "
    "Next actions: re-open the workbook (the old workbook_id is invalid)."
)
_DESC_SNAPSHOT_CREATE = (
    "Create an immutable snapshot of a workbook (SaveCopyAs for open books, file "
    "copy for closed) with manifest and SHA-256."
)
_DESC_SNAPSHOT_LIST = "List snapshots and their metadata."
_DESC_SNAPSHOT_RESTORE = (
    "Purpose: Restore a working session from a snapshot. "
    "Use when: you want to roll the working copy back to a known baseline. "
    "Preconditions: an existing snapshot_id (and workbook_id when closed). "
    "Behavior: closes the current working copy without saving — unsaved changes "
    "are discarded — then opens the snapshot baseline; does NOT undo external "
    "macro side effects (database writes, emails, network calls, files, printing). "
    "Requires `confirm=true`; a missing/false confirm is rejected with "
    "CONFIRMATION_REQUIRED before the working copy is closed. "
    "Result: WorkbookInfo. "
    "Next actions: snapshot_create to re-baseline before further mutations."
)
_DESC_SNAPSHOT_DIFF = "Compare two snapshots by manifest metadata and file hash."
_DESC_TEST = (
    "Purpose: Validate and run a declarative test suite, or read a previously "
    "persisted run, through one dispatch tool. "
    "Use when: you have a suite object or a plain JSON-style mapping and want to "
    "run it (run-suite), or you need to read back a persisted run by id "
    "(get-run). "
    "Do not use when: you need to run a single macro — use macro_run; run-suite "
    "executes on disposable working copies and is not read-only. "
    "Preconditions: suite.workbook_id must be an open workbook (call workbook "
    "action='open' first) for run-suite; a baseline snapshot is created "
    "automatically unless baseline='use_existing'; run_id is required for "
    "get-run. "
    "Parameters: action='run-suite'|'get-run'; suite (run-suite): TestSuiteSpec "
    "or mapping; run_id (get-run); background (run-suite). The closed Range/Pivot "
    "action and assertion vocabularies are validated before any Excel call, so an "
    "unknown action/assertion or an arbitrary-code/unbounded payload is rejected "
    "up front. "
    "Behavior: run-suite uses one immutable baseline snapshot, one disposable "
    "working copy per case, external macro side effects are NOT rolled back, and "
    "background=true returns a RunningJob handle; get-run never touches Excel and "
    "never re-runs anything. "
    "Result: TestRunResult (run-suite / get-run), or RunningJob when "
    "background=true. "
    "Next actions: job_get to retrieve a background result; test action='get-run' "
    "to read a persisted run. "
    "Example: {\"action\": \"run-suite\", \"suite\": {\"suite\": \"smoke\", "
    "\"workbook_id\": \"<id>\", \"cases\": [{\"id\": \"c1\", \"steps\": "
    "[{\"action\": \"range_write_values\", \"sheet\": \"Sheet1\", "
    "\"range_address\": \"A1\", \"values_grid\": {\"rows\": 1, \"cols\": 1, "
    "\"values\": [[1]]}}]}]}, \"background\": false}"
)
_DESC_VBA_READ = (
    "Purpose: Read the VBA project surface of an open workbook through one "
    "read-only dispatcher. "
    "Use when: you need the project summary, component inventory, procedures, a "
    "component's source, references, a code search, or the cached compile "
    "diagnostics. "
    "Do not use when: you need to mutate VBA source — use "
    "vba_replace_component_source / vba_insert_code; you need to export the "
    "project — use vba_export_project; you need a fresh compile diagnostic — use "
    "vba_compile_diagnostics. "
    "Preconditions: an open workbook_id (required for every action). "
    "Parameters: action='project-info'|'list-components'|'list-procedures'|"
    "'read-component'|'get-references'|'search-code'|'get-compile-diagnostics'; "
    "workbook_id; component (list-procedures / read-component); pattern + regex "
    "(search-code). "
    "Behavior: read-only; routes to the existing per-verb handlers unchanged and "
    "forwards only validated arguments; a missing per-action required parameter "
    "surfaces as a structured ExcelVbaMcpError. "
    "Result: VbaProjectInfo (project-info), list[VbaComponentInfo] "
    "(list-components), list[VbaProcedureInfo] (list-procedures), "
    "VbaComponentSource (read-component), list[VbaReferenceInfo] "
    "(get-references), list[VbaSearchMatch] (search-code) or CompileDiagnostics "
    "(get-compile-diagnostics). "
    "Next actions: vba_replace_component_source / vba_insert_code to mutate; "
    "vba_export_project to export."
)
_DESC_VBA_EXPORT_PROJECT = (
    "Export every VBA component (.bas/.cls/.frm+.frx) and return the export "
    "manifest with per-file hashes; Document modules are never removed."
)
_DESC_VBA_REPLACE_COMPONENT_SOURCE = (
    "Purpose: Replace the entire code module of a Standard/Class VBA component. Use when: you "
    "need to rewrite a component's source wholesale. Do not use when: you only add code at a "
    "position — use vba_insert_code; Document modules and UserForms are rejected. "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read. Parameters: workbook_id; component; "
    "new_source; expected_source_hash — obtain it from vba_read_component / "
    "vba_list_components (the normalized source hash, not the snapshot export file hash). "
    "Behavior: the revision check runs before and inside the STA worker (see "
    "xlvba://glossary); a mismatch raises REVISION_CONFLICT. Result: VbaMutationResult "
    "(post-mutation component state). Next actions: vba_read_component to confirm the new "
    "source hash. Example: {\"workbook_id\": \"<id>\", \"component\": \"Module1\", "
    "\"new_source\": \"Sub Hello()\\n  MsgBox \\\"hi\\\"\\nEnd Sub\", "
    "\"expected_source_hash\": \"<hash from vba_read_component>\"}"
)
_DESC_VBA_INSERT_CODE = (
    "Purpose: Insert VBA code at a semantic position in a Standard/Class component. Use when: "
    "you add a procedure or declaration without rewriting the module. Do not use when: you "
    "replace the whole module — use vba_replace_component_source. Preconditions: writable "
    "session + snapshot + approval Approval gate (see xlvba://glossary) + postcondition "
    "re-read. Parameters: position is 'declarations', 'before_procedure:<name>' or "
    "'after_procedure:<name>'; expected_source_hash from vba_read_component / "
    "vba_list_components. Behavior: same revision guard as replace (REVISION_CONFLICT on "
    "mismatch). Result: VbaMutationResult. Next actions: vba_read_component to confirm the new "
    "source hash."
)
_DESC_VBA_COMPILE_DIAGNOSTICS = (
    "Purpose: Return a fail-closed compile diagnostic for an open workbook — "
    "diagnostic only, does NOT compile. "
    "Use when: you need the compile/protection/inventory diagnostic snapshot. "
    "Do not use when: you need a real compiled/compile_error result — none is "
    "available; use vba_get_compile_diagnostics to re-read the cache. "
    "Preconditions: an open workbook_id. "
    "Behavior: on the validated build non-activating compilation is unavailable, "
    "so this returns status='indeterminate' with "
    "reason='NON_ACTIVATING_COMPILE_UNAVAILABLE' and partial diagnostics "
    "(is_compile=false); it never returns compiled or compile_error and never "
    "shows or activates UI (see xlvba://glossary: indeterminate). "
    "Result: a compile diagnostic DTO (status indeterminate). "
    "Next actions: vba_get_compile_diagnostics for the cached result."
)
_DESC_MACRO_RUN = (
    "Purpose: Run one approved, background-safe macro of an execute-mode workbook. "
    "Use when: you need to execute a reviewed VBA procedure. "
    "Do not use when: the procedure is not yet classified — run "
    "macro action='classify-background' first; to re-read the last result use "
    "macro action='get-last-result'. "
    "Preconditions (gates, in order): (1) execute-mode + writable session; "
    "(2) procedure resolution; (3) approval — 'approved' is the call-time "
    "ApprovalPolicy.check(EXECUTING) outcome (allowed by default in full; "
    "allowlist/live grant in restricted); (4) background_safe for the current "
    "workbook hash (BG06); (5) interactive-UI preflight (R13; see "
    "xlvba://glossary); (6) dispatch. "
    "Parameters: workbook_id; macro_name; args (≤30 — the Excel Application.Run "
    "limit; a bare datetime.date is rejected, pass datetime.datetime); "
    "background=False. "
    "Behavior: otherwise a truthful blocked result is returned and the macro is "
    "NOT dispatched; never dismisses a dialog or reports a waiting macro as "
    "passed; background=true returns a RunningJob handle. "
    "Result: MacroRunResult, or RunningJob when background=true. "
    "Next actions: macro action='get-last-result'; job_get for a background run."
)
_DESC_MACRO = (
    "Purpose: List callable macros, classify a procedure's background_safe, or "
    "read the cached last macro run result through one dispatch tool. "
    "Use when: you need the macro candidate inventory of an open workbook, you "
    "have reviewed a procedure and want to record its background_safe "
    "classification, or you need to re-read the last macro run result. "
    "Do not use when: you need to execute a macro — use macro_run (standalone, "
    "distinct permission); classify-background is bookkeeping only and never "
    "grants run permission. "
    "Preconditions: an open workbook_id; procedure + background_safe for "
    "classify-background; review the code first (vba_read_component / "
    "vba_list_procedures). "
    "Parameters: action='list'|'classify-background'|'get-last-result'; "
    "workbook_id; procedure (classify-background); background_safe "
    "(classify-background); expected_project_source_hash (classify-background, "
    "optional optimistic-concurrency guard). "
    "Behavior: list and get-last-result are reads; classify-background is "
    "UNGATED bookkeeping bound to (workbook hash, source hash, procedure) "
    "(BG06) and invalidated on VBA source change; a heuristic or interactive-UI "
    "procedure cannot be classified background_safe (R13 veto); the "
    "classification never grants run permission — macro_run still enforces "
    "approval + BG06 + R13. "
    "Result: list[MacroInfo] (list), BackgroundSafetyClassification "
    "(classify-background) or MacroRunResult / None (get-last-result). "
    "Next actions: macro_run once classified; job_get for a background run. "
    "Example: {\"action\": \"classify-background\", \"workbook_id\": \"<id>\", "
    "\"procedure\": \"Module1.Recalculate\", \"background_safe\": true}"
)
_DESC_JOB_GET = (
    "Purpose: Return one job's status/result by job_id (poll; never touches COM). "
    "Use when: you hold a job_id from a background=true call and want its state. "
    "Preconditions: a job_id from a long-op handle. "
    "Parameters: job_id. "
    "Behavior: re-call no sooner than poll_after_s; status is one of "
    "queued/running/succeeded/failed/expired/unknown/interrupted_unknown; the "
    "result is available once terminal. interrupted_unknown (outcome='unknown', "
    "retryable=false) is a job left non-terminal by a prior process — never "
    "completed and never auto-retried; SUSPECTED_STALLED means a soft deadline "
    "was exceeded and the op was left running (not confirmed hung) — see "
    "xlvba://glossary. "
    "Result: JobStatus. "
    "Next actions: job_request_cancel to record cancel intent; excel_recover to actually "
    "stop a wedged op."
)
_DESC_JOB_REQUEST_CANCEL = (
    "Purpose: Record a job's cancel intent. "
    "Use when: you want to signal a long op to stop at its next checkpoint. "
    "Do not use when: you expect it to interrupt the running COM call — it is "
    "intent-only; never interrupts; observe cancel_requested via job_get and use "
    "excel_recover to actually stop a wedged process. "
    "Preconditions: a non-terminal job_id. "
    "Behavior: sets cancel_requested=True only if non-terminal; the single-STA "
    "COM call runs to completion and a poller observes cancel_requested. "
    "Result: JobStatus. "
    "Next actions: job_get to observe; excel_recover if it must be stopped."
)
_DESC_SHEETS_LIST = (
    "List every sheet of an open workbook (all types: Worksheet / Chart / Macro "
    "/ Dialog) with name, index, object_type, visibility and — for worksheets — "
    "code_name (read-only), protected, table_count and used_range."
)
_DESC_SHEET_GET = "Return one sheet's SheetInfo by name (Sheets vs Worksheets aware)."
_DESC_SHEET_ADD = (
    "Purpose: Add a worksheet. Use when: you need a new sheet in an open workbook. Do not use "
    "when: you need a non-worksheet sheet — only 'worksheet' is supported (others rejected). "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read. Parameters: workbook_id; name; "
    "sheet_type='worksheet' (default). Behavior: creates the sheet and returns its DTO. "
    "Result: SheetInfo. Next actions: range_write_values / table_create to populate it."
)
_DESC_SHEET_RENAME = (
    "Rename a sheet; CodeName is never changed. Guarded mutation (see xlvba://glossary)."
)
_DESC_SHEET_COPY = (
    "Purpose: Copy a sheet and return the new sheet's DTO. Use when: you need an in-workbook "
    "duplicate of an existing sheet. Do not use when: you need a cross-workbook copy or an "
    "arbitrary position — neither is supported. Preconditions: writable session + snapshot + "
    "approval Approval gate (see xlvba://glossary) + postcondition re-read. Parameters: "
    "workbook_id; name (source sheet). Behavior: the copy's name is generated by Excel "
    "(typically '<name> (2)'). Result: SheetInfo. Next actions: sheet_rename to give the copy "
    "a deliberate name."
)
_DESC_SHEET_DELETE = (
    "Delete a sheet (destructive). Guarded mutation (see xlvba://glossary). Rejected when it "
    "would leave no visible worksheet."
)
_DESC_SHEET_SET_VISIBILITY = (
    "Set a sheet's visibility (xlSheetVisible / xlSheetHidden). Guarded mutation (see "
    "xlvba://glossary). Hiding the last visible worksheet is rejected."
)
_DESC_SHEET = (
    "Purpose: Add, rename, copy or change the visibility of a sheet through one "
    "mutation tool. "
    "Use when: you need a new worksheet, an in-workbook duplicate, a new name, or "
    "to show/hide a sheet. "
    "Do not use when: you need to delete a sheet — use sheet_delete; you need to "
    "read the sheet inventory — use sheet_read. "
    "Preconditions: an open workbook_id; writable session + snapshot + approval "
    "Approval gate (see xlvba://glossary) + postcondition re-read. "
    "Parameters: action='add'|'rename'|'copy'|'set-visibility'; workbook_id; name; "
    "new_name (rename); sheet_type='worksheet' (add); visible (set-visibility). "
    "Behavior: add creates a worksheet and returns its DTO; rename never changes "
    "CodeName; copy's name is generated by Excel (typically '<name> (2)'); "
    "set-visibility rejects hiding the last visible worksheet. Guarded mutation "
    "(see xlvba://glossary). "
    "Result: SheetInfo. "
    "Next actions: sheet_read action='get' to confirm; range_write_values / "
    "table_create to populate a new sheet. "
    "Example: {\"action\": \"add\", \"workbook_id\": \"<id>\", \"name\": \"Data\"}"
)
_DESC_SHEET_READ = (
    "Purpose: Read the sheet inventory of an open workbook without mutating it. "
    "Use when: you need every sheet (all types) or one sheet's metadata by name. "
    "Do not use when: you need to add/rename/copy/hide a sheet — use sheet; you "
    "need to delete a sheet — use sheet_delete. "
    "Preconditions: an open workbook_id; name is required for get. "
    "Parameters: action='list'|'get'; workbook_id; name (get). "
    "Behavior: read-only; list returns one entry per Sheets item (Worksheet / "
    "Chart / Macro / Dialog) with name, index, object_type, visibility and — for "
    "worksheets — code_name (read-only), protected, table_count and used_range. "
    "Result: list[SheetInfo] (list) or SheetInfo (get). "
    "Next actions: sheet action='add' to create a sheet; range_read_values to read "
    "cells. "
    "Example: {\"action\": \"get\", \"workbook_id\": \"<id>\", \"name\": \"Sheet1\"}"
)
_DESC_TABLES_LIST = (
    "List every Excel Table (ListObject) of an open workbook with name, sheet, "
    "range address, HeaderRowRange, DataBodyRange, TotalRowRange, ShowTotals, "
    "TableStyle, row count and columns."
)
_DESC_TABLE_GET = (
    "Return one Excel Table's TableInfo by name/display name (range, header/body/"
    "total ranges, totals, style, rows, columns)."
)
_DESC_TABLE_READ_ROWS = (
    "Read a typed, bounded slice of a table's data rows. offset is a 0-based row "
    "index; limit caps the rows. Values are typed (str/int/float/bool/null/ISO "
    "date/Excel error) and the response is size-limited (truncated flag)."
)
_DESC_TABLE_APPEND_ROWS = (
    "Purpose: Append rows through ListRows.Add (table structure preserved). Use when: you want "
    "to add records to the end of an Excel Table. Do not use when: you need to update or "
    "delete existing rows — use table_update_rows / table_delete_rows. Preconditions: writable "
    "session + snapshot + approval Approval gate (see xlvba://glossary) + postcondition "
    "re-read. Parameters: workbook_id; name; rows is a non-empty list of {header: value} "
    "objects; headers must match the table columns (case-insensitive). Behavior: appends the "
    "rows and returns the new row count. Result: TableMutationResult. Next actions: "
    "table_read_rows to verify the appended data."
)
_DESC_TABLE_INSERT_ROW = (
    "Insert one data row at a 1-based position via ListRows.Add. Guarded mutation (see "
    "xlvba://glossary)."
)
_DESC_TABLE_UPDATE_ROWS = (
    "Purpose: Update rows by explicit index or an unambiguous where filter. Use when: you need "
    "to change values in existing Table rows. Do not use when: you need to add rows — use "
    "table_append_rows; to delete rows use table_delete_rows. Preconditions: writable session "
    "+ snapshot + approval Approval gate (see xlvba://glossary) + postcondition re-read. "
    "Parameters: updates=[{index|where, values}] — exactly one selector per entry; a where "
    "that matches zero or multiple rows is rejected with AMBIGUOUS_FILTER (no auto-guess). "
    "Behavior: applies each update and re-reads the postcondition. Result: "
    "TableMutationResult. Next actions: table_read_rows to verify. Example: {\"workbook_id\": "
    "\"<id>\", \"name\": \"Sales\", \"updates\": [{\"index\": 3, \"values\": {\"Amount\": "
    "250}}, {\"where\": {\"SKU\": \"A-1\"}, \"values\": {\"Amount\": 99}}]}"
)
_DESC_TABLE_DELETE_ROWS = (
    "Purpose: Delete rows by explicit indices or an unambiguous where filter. Use when: you "
    "need to remove rows from an Excel Table. Do not use when: you need to change row values — "
    "use table_update_rows. Preconditions: writable session + snapshot + approval Approval "
    "gate (see xlvba://glossary) + postcondition re-read. Parameters: indices (explicit "
    "1-based) or where (object) — exactly one selector; a where that matches zero or multiple "
    "rows is rejected with AMBIGUOUS_FILTER (no auto-guess). Behavior: deletes the matched "
    "rows and re-reads the postcondition. Result: TableMutationResult. Next actions: "
    "table_read_rows to verify."
)
_DESC_TABLE_ADD_COLUMN = (
    "Add a column via ListColumns.Add(Position) with a unique header. Guarded mutation (see "
    "xlvba://glossary)."
)
_DESC_TABLE_DELETE_COLUMN = (
    "Delete the column identified by a unique header. Guarded mutation (see xlvba://glossary)."
)
_DESC_TABLE_CREATE = (
    "Create a ListObject on a sheet over an A1 range or defined name. Guarded mutation (see "
    "xlvba://glossary)."
)
_DESC_TABLE_RESIZE = (
    "Resize a table to a new A1 range via ListObject.Resize. Guarded mutation (see "
    "xlvba://glossary)."
)
_DESC_TABLE = (
    "Purpose: Create, resize, append to, insert into, update or add a column to an "
    "Excel Table through one mutation tool. "
    "Use when: you need a new ListObject, a resized table, appended/inserted rows, "
    "updated row values, or a new column. "
    "Do not use when: you need to delete rows — use table_delete_rows; you need to "
    "delete a column — use table_delete_column; you need to read the table "
    "inventory/data — use table_read. "
    "Preconditions: an open workbook_id; writable session + snapshot + approval "
    "Approval gate (see xlvba://glossary) + postcondition re-read. "
    "Parameters: action='create'|'resize'|'append-rows'|'insert-row'|'update-rows'|"
    "'add-column'; workbook_id; sheet + range_or_name + style (create); name + "
    "new_range (resize); name + rows (append-rows); name + at + values (insert-row); "
    "name + updates (update-rows); name + header + position (add-column). "
    "Behavior: create makes a ListObject over an A1 range or defined name; resize "
    "requires the header row; append-rows preserves structure; insert-row inserts at "
    "a 1-based position; update-rows applies {index|where, values} with an "
    "ambiguous where rejected (no auto-guess); add-column requires a unique header. "
    "Guarded mutation (see xlvba://glossary). "
    "Result: TableMutationResult. "
    "Next actions: table_read action='read-rows' to verify. "
    "Example: {\"action\": \"append-rows\", \"workbook_id\": \"<id>\", \"name\": "
    "\"Sales\", \"rows\": [{\"Amount\": 100}]}"
)
_DESC_TABLE_READ = (
    "Purpose: Read the table inventory or data of an open workbook without mutating "
    "it. "
    "Use when: you need every Excel Table, one table's metadata by name, or a typed "
    "slice of a table's data rows. "
    "Do not use when: you need to create/resize/append/insert/update/add-column — use "
    "table; you need to delete rows — use table_delete_rows; you need to delete a "
    "column — use table_delete_column. "
    "Preconditions: an open workbook_id; name is required for get / read-rows. "
    "Parameters: action='list'|'get'|'read-rows'; workbook_id; name (get / "
    "read-rows); offset + limit (read-rows). "
    "Behavior: read-only; list returns one entry per ListObject with range/totals/"
    "style/row count/columns; get returns one TableInfo; read-rows returns typed "
    "values (str/int/float/bool/null/ISO date/Excel error) with a truncated flag. "
    "Result: list[TableInfo] (list), TableInfo (get) or TableRows (read-rows). "
    "Next actions: table action='append-rows' to add records; range_read_values to "
    "read cells. "
    "Example: {\"action\": \"read-rows\", \"workbook_id\": \"<id>\", \"name\": "
    "\"Sales\", \"offset\": 0, \"limit\": 50}"
)
_DESC_RANGE_GET = (
    "Resolve an explicit worksheet + A1-style address to Range metadata "
    "(address, rows, cols, shape, merged/table flags). Read-only; multi-area or "
    "invalid addresses are rejected with a structured error."
)
_DESC_RANGE_READ_VALUES = (
    "Purpose: Read typed cell values from a bounded Range window. "
    "Use when: you need cell values (not formulas) from a sheet. "
    "Do not use when: you need formulas — use range_read_formulas. "
    "Preconditions: an open workbook_id. "
    "Parameters: sheet; address; offset (0-based row index); limit. Caps: a "
    "request over 100000 cells is rejected; the response is capped at 50000 cells "
    "/ 8 MiB; truncated=true marks a cut-short window. "
    "Behavior: values are typed (empty/string/number/boolean/date/error) with no "
    "lossy string coercion. "
    "Result: RangeValues (rows, cols, values, truncated). "
    "Next actions: range_get for geometry; range_read_formulas for formulas."
)
_DESC_RANGE_READ_FORMULAS = (
    "Read typed formulas from a bounded Range window (Formula2 primary, Formula "
    "fallback). _xlfn/_xlws/_xlpm prefixes are returned verbatim and flagged. "
    "offset is a 0-based row index and limit caps the rows; oversized requests "
    "are rejected."
)
_DESC_RANGE_WRITE_VALUES = (
    "Purpose: Write typed cell values to a bounded Range. Use when: you need to set cell "
    "values (constants) in a sheet. Do not use when: you need formulas — a leading '=' is "
    "stored as text here; use range_write_formulas. Preconditions: this tool does NOT snapshot "
    "— call snapshot_create first; then writable session + snapshot + approval Approval gate "
    "(see xlvba://glossary) + postcondition re-read. Parameters: {rows, cols, values}; "
    "rows≤100000, cols≤512, cells≤100000; address is the top-left anchor; a string beginning "
    "with '=' is stored as text, not a formula. Behavior: the payload shape is validated "
    "against the target before any mutation (a wrong-sized payload is rejected with no "
    "change); cell writes inside an Excel Table are allowed only when the table row/column "
    "structure is unchanged. Result: RangeMutationResult. Next actions: range_read_values to "
    "verify. Example: {\"workbook_id\": \"<id>\", \"sheet\": \"Sheet1\", \"address\": \"A1\", "
    "\"payload\": {\"rows\": 1, \"cols\": 2, \"values\": [[1, 2]]}}"
)
_DESC_RANGE_WRITE_FORMULAS = (
    "Purpose: Write formulas (Formula2 primary, Formula fallback) to a bounded Range. Use "
    "when: you need to set formula strings beginning with '='. Do not use when: you need "
    "literal '=' text — use range_write_values (which stores a leading '=' as text). "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read. Parameters: every entry must be a formula "
    "string beginning with '='; a bare constant belongs in range_write_values. Behavior: cell "
    "writes inside an Excel Table are allowed only when the table row/column structure is "
    "unchanged. Result: RangeMutationResult. Next actions: range_read_formulas to verify."
)
_DESC_RANGE_GET_FORMAT = (
    "Read bounded, tri-state formatting for a Range: number_format, font "
    "(name/bold/italic/size/color), fill (color/pattern) and alignment "
    "(horizontal/vertical/wrap_text). Each property is a value, mixed (cells "
    "disagree; listed in mixed_properties) or unsupported (listed in "
    "unsupported); a mixed range is never flattened. Read-only."
)
_DESC_RANGE_SET_FORMAT = (
    "Apply a partial formatting update to a Range. Only the fields explicitly set in the patch "
    "are written; absent fields are left untouched. Values and enums are validated before any "
    "mutation; a null clears to automatic only where Excel allows. Guarded mutation (see "
    "xlvba://glossary)."
)
_DESC_RANGE_COPY = (
    "Copy a Range to an explicit destination (same workbook, no clipboard). The destination is "
    "mandatory; an overlapping source/destination, a shape-mismatched destination or a "
    "destination intersecting merged cells is rejected. Structural copies that intersect an "
    "Excel Table (ListObject) are rejected and routed to the table tools. Guarded mutation "
    "(see xlvba://glossary)."
)
_DESC_RANGE_INSERT = (
    "Insert cells at an explicit address and shift them in an explicit direction ('right' or "
    "'down'). An omitted/unknown shift is rejected (never guessed). Rejected for a protected "
    "sheet, an Excel Table (ListObject) intersection, a dynamic-array spill range, or merged "
    "cells without allow_merged. Guarded mutation (see xlvba://glossary)."
)
_DESC_RANGE_DELETE = (
    "Delete cells at an explicit address and shift the remainder in an explicit direction "
    "('left' or 'up'). An omitted/unknown shift is rejected (never guessed). Rejected for a "
    "protected sheet, an Excel Table (ListObject) intersection, a dynamic-array spill range, "
    "or merged cells without allow_merged. Guarded mutation (see xlvba://glossary)."
)
_DESC_RANGE_MERGE = (
    "Merge a Range into one cell (or one cell per row with across=true). Merging discards "
    "non-anchor data: the preflight reports the affected cells and the merge is blocked unless "
    "allow_data_loss=true is passed after review. Rejected for a protected sheet, an Excel "
    "Table (ListObject) intersection or a spill range. Guarded mutation (see xlvba://glossary)."
)
_DESC_RANGE_UNMERGE = (
    "Unmerge every merged area intersecting an explicit target. Rejected when no merged area "
    "intersects the target, for a protected sheet, or for an Excel Table (ListObject) "
    "intersection. Guarded mutation (see xlvba://glossary)."
)
_DESC_RANGE_GET_VALIDATION = (
    "Read the data-validation rule of a Range (type, Formula1/Formula2, operator, "
    "ignore_blank, in_cell_dropdown, alert_style and the message/title strings). "
    "A target without a rule returns present=false; a non-uniform target is "
    "reported with non_uniform_validation / mixed_validation in unsupported "
    "instead of being flattened. Read-only."
)
_DESC_RANGE_SET_VALIDATION = (
    "Add or modify the data-validation rule of a Range (list, whole number, decimal, date, "
    "time, text length, custom). Formula1/Formula2 are opaque strings bounded at 255 "
    "characters and are never evaluated; a missing operand, an over-long formula or an "
    "unsupported type is rejected before any mutation. An existing rule is modified so "
    "unspecified properties are preserved. Guarded mutation (see xlvba://glossary)."
)
_DESC_RANGE_GET_CONDITIONAL_FORMATTING = (
    "Read the conditional-formatting rules applying to a Range in ascending "
    "priority order (type, operator, Formula1/Formula2, stop_if_true, applies_to "
    "plus a bounded fill/font subset). ColorScale / DataBar / IconSet rules are "
    "reported with supported=false and an explicit unsupported_reason rather than "
    "omitted. Read-only."
)
_DESC_RANGE_SET_CONDITIONAL_FORMATTING = (
    "Add, modify or delete one supported conditional-format rule (mode: add, "
    "modify, delete). spec.priority is the only stable rule identity and is "
    "required for modify/delete; a rule type is never silently converted. "
    "Priorities are sheet-wide: assigning one renumbers the other rules. "
    "spec.type is the plain FormatCondition set only (cell_value, expression, "
    "text_string, blanks, no_blanks, errors, no_errors, time_period, top10, "
    "above_average, unique_values); color_scale, data_bar, icon_set and unsupported "
    "are not accepted — use range_get_conditional_formatting to read those rule "
    "types (reported supported=false). Guarded mutation: writable session + "
    "snapshot + approval + postcondition re-read."
)
_DESC_RANGE_READ = (
    "Purpose: Read Range metadata, values, formulas, formatting, validation or "
    "conditional formatting through one read-only dispatch tool. "
    "Use when: you need to resolve an address, read typed cell values/formulas, "
    "or inspect a Range's format/validation/conditional-format rules. "
    "Do not use when: you need to write values/formulas or change formatting — "
    "use range_write; you need to copy/insert/unmerge cells — use "
    "range_structural. "
    "Preconditions: an open workbook_id; sheet and address are required for "
    "every action. "
    "Parameters: action='get'|'read-values'|'read-formulas'|'get-format'|"
    "'get-validation'|'get-conditional-formatting'; workbook_id; sheet; address; "
    "offset + limit (read-values / read-formulas). "
    "Behavior: read-only; values are typed (empty/string/number/boolean/date/"
    "error) with no lossy string coercion; a mixed format/validation target is "
    "reported (never flattened) and unsupported conditional-format rule types "
    "are flagged rather than omitted. "
    "Result: RangeAddress (get), RangeValues (read-values), RangeFormulas "
    "(read-formulas), FormatInfo (get-format), ValidationRule (get-validation) "
    "or list[ConditionalFormatRule] (get-conditional-formatting). "
    "Next actions: range_write to change a value/format; range_structural to "
    "copy/insert/unmerge. "
    "Example: {\"action\": \"read-values\", \"workbook_id\": \"<id>\", "
    "\"sheet\": \"Sheet1\", \"address\": \"A1:C10\", \"offset\": 0, "
    "\"limit\": 50}"
)
_DESC_RANGE_WRITE = (
    "Purpose: Write values/formulas or set formatting/validation/conditional "
    "formatting through one mutation dispatch tool. "
    "Use when: you need to set cell values or formulas, apply a partial format "
    "update, or add/modify/delete a validation or conditional-format rule. "
    "Do not use when: you need to read a Range — use range_read; you need to "
    "copy/insert/unmerge cells — use range_structural; you need to delete cells "
    "— use range_delete. "
    "Preconditions: this tool does NOT snapshot — call snapshot_create first; "
    "then writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read. "
    "Parameters: action='write-values'|'write-formulas'|'set-format'|"
    "'set-validation'|'set-conditional-formatting'; workbook_id; sheet; address; "
    "payload (write-values / write-formulas); patch (set-format); spec "
    "(set-validation / set-conditional-formatting); mode (set-conditional-"
    "formatting: add/modify/delete). "
    "Behavior: write-values stores a leading '=' as text (use write-formulas for "
    "real formulas); the payload shape is validated against the target before "
    "any mutation; set-conditional-formatting mode='add' appends a rule and "
    "renumbers the sheet-wide priorities, so this tool is NOT idempotent. "
    "Guarded mutation (see xlvba://glossary). "
    "Result: RangeMutationResult (write-values / write-formulas), "
    "FormatMutationResult (set-format), ValidationRule (set-validation) or "
    "ConditionalFormatMutationResult (set-conditional-formatting). "
    "Next actions: range_read to verify. "
    "Example: {\"action\": \"write-values\", \"workbook_id\": \"<id>\", "
    "\"sheet\": \"Sheet1\", \"address\": \"A1\", \"payload\": {\"rows\": 1, "
    "\"cols\": 2, \"values\": [[1, 2]]}}"
)
_DESC_RANGE_STRUCTURAL = (
    "Purpose: Copy, insert or unmerge cells through one structural dispatch "
    "tool. "
    "Use when: you need to copy a Range to an explicit destination, insert cells "
    "with an explicit shift direction, or unmerge a merged area. "
    "Do not use when: you need to write values/formatting — use range_write; you "
    "need to delete cells — use range_delete; you need to merge cells — use "
    "range_merge. "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read. "
    "Parameters: action='copy'|'insert'|'unmerge'; workbook_id; sheet; "
    "source + destination_sheet + destination (copy); address + shift "
    "(insert: 'right' or 'down', required, never guessed) + allow_merged "
    "(insert); address (unmerge). "
    "Behavior: copy requires a disjoint, shape-matched destination; insert "
    "rejects an omitted/disallowed shift and any ListObject intersection; "
    "unmerge discards merges. This tool is NOT idempotent (re-insert duplicates "
    "cells; unmerge discards merges). "
    "Guarded mutation (see xlvba://glossary). "
    "Result: StructuralMutationResult. "
    "Next actions: range_read to verify the resulting address. "
    "Example: {\"action\": \"insert\", \"workbook_id\": \"<id>\", "
    "\"sheet\": \"Sheet1\", \"address\": \"A1\", \"shift\": \"down\"}"
)
_DESC_NAME = (
    "Purpose: Create or update a defined name (Named Range) in an explicit scope. "
    "Use when: you need a named range or constant in an open workbook, or you need "
    "to repoint an existing name's RefersTo. "
    "Do not use when: you need to delete a name — use name_delete; you need to "
    "read names — use name_read. "
    "Preconditions: an open workbook_id; writable session + snapshot + approval "
    "Approval gate (see xlvba://glossary) + postcondition re-read. "
    "Parameters: action='create'|'update'; workbook_id; name; refers_to (may be "
    "given with or without a leading '=', normalized); scope ('workbook' or "
    "'worksheet' + owning sheet). "
    "Behavior: create rejects duplicate and cross-scope conflicts; update preserves "
    "scope and only rewrites RefersTo. Guarded mutation (see xlvba://glossary). "
    "Result: DefinedNameInfo. "
    "Next actions: name_read action='get' to confirm. "
    "Example: {\"action\": \"create\", \"workbook_id\": \"<id>\", \"name\": "
    "\"Sales\", \"refers_to\": \"=Sheet1!$A$1\", \"scope\": \"workbook\"}"
)
_DESC_NAME_READ = (
    "Purpose: Read the defined names (Named Ranges) of an open workbook without "
    "mutating it. "
    "Use when: you need every defined name or one name's metadata by name + scope. "
    "Do not use when: you need to create/update a name — use name; you need to "
    "delete a name — use name_delete. "
    "Preconditions: an open workbook_id; name + scope are required for get (sheet "
    "also required for a worksheet scope). "
    "Parameters: action='list'|'get'; workbook_id; name (get); scope (get); sheet "
    "(get, worksheet scope). "
    "Behavior: read-only; list returns one entry per defined name across workbook "
    "and worksheet scopes with scope, NameLocal, Visible, RefersTo/RefersToLocal "
    "and target_kind (range with sheet+address, constant, formula, external, error "
    "or unknown); get returns one DefinedNameInfo. "
    "Result: list[DefinedNameInfo] (list) or DefinedNameInfo (get). "
    "Next actions: name action='create' to add a name; name_delete to remove one. "
    "Example: {\"action\": \"get\", \"workbook_id\": \"<id>\", \"name\": "
    "\"Sales\", \"scope\": \"workbook\"}"
)
_DESC_NAME_DELETE = (
    "Delete a defined name in an explicit scope. Guarded mutation (see xlvba://glossary). "
    "Returns the pre-delete name."
)
_DESC_SHAPE = (
    "Purpose: List, read or assign the OnAction of drawing shapes of an open "
    "workbook. "
    "Use when: you need the shape inventory, one shape's metadata, or you need "
    "to set/clear a shape's OnAction macro (the macro is NOT executed). "
    "Do not use when: you need Form/ActiveX controls — use control. "
    "Preconditions: an open workbook_id; name is required for get and "
    "set-on-action; macro_name is required for set-on-action; set-on-action is a "
    "Guarded mutation (see xlvba://glossary). "
    "Parameters: action='list'|'get'|'set-on-action'; workbook_id; name (get, "
    "set-on-action); macro_name (set-on-action; \"\" clears it). "
    "Behavior: list returns one entry per Worksheet.Shapes item (name, kind, "
    "on_action read verbatim, visibility, position/size, best-effort text); get "
    "returns one ShapeInfo; set-on-action writes Shape.OnAction only. "
    "Result: list[ShapeInfo] (list) or ShapeInfo (get / set-on-action). "
    "Next actions: shape action='get' to confirm; control for Form/ActiveX "
    "controls. "
    "Example: {\"action\": \"set-on-action\", \"workbook_id\": \"<id>\", "
    "\"name\": \"btnSave\", \"macro_name\": \"SaveData\"}"
)
_DESC_CONTROL = (
    "Purpose: List, read or assign the OnAction of Form Controls / ActiveX "
    "controls of an open workbook. "
    "Use when: you need the control inventory, one control's metadata, or you "
    "need to set/clear a Form Control's OnAction macro (the macro is NOT "
    "executed). "
    "Do not use when: you need drawing shapes — use shape. "
    "Preconditions: an open workbook_id; name is required for get and "
    "set-on-action; macro_name is required for set-on-action; set-on-action is a "
    "Guarded mutation (see xlvba://glossary). "
    "Parameters: action='list'|'get'|'set-on-action'; workbook_id; name (get, "
    "set-on-action); macro_name (set-on-action; \"\" clears it). "
    "Behavior: list reports Form Controls (control_type, caption, linked_cell, "
    "on_action) and ActiveX controls (prog_id, inferred/probable event "
    "procedure); get returns one ControlEntry; set-on-action writes OnAction on "
    "a Form Control only (ActiveX controls are rejected — no OnAction). "
    "Result: list[ControlEntry] (list) or ControlEntry (get / set-on-action). "
    "Next actions: control action='get' to confirm; shape for drawing shapes. "
    "Example: {\"action\": \"set-on-action\", \"workbook_id\": \"<id>\", "
    "\"name\": \"cmdRun\", \"macro_name\": \"RunReport\"}"
)
_DESC_WORKBOOK_GET_MODEL = (
    "Return the aggregated Stage 1 workbook model: sheets, tables, names, shapes, "
    "controls and (optionally) VBA components/procedures, each with a section "
    "status. Read-only: no all-cells/Range data and no VBA execution. When VBA is "
    "unavailable, locked or VBOM trust is disabled the other sections are still "
    "returned and vba_access reports the state."
)
_DESC_PIVOT_READ = (
    "Purpose: Read PivotTables, PivotCaches and PivotFields of an open workbook "
    "through one read-only dispatch tool. "
    "Use when: you need the PivotTable/Cache inventory, one PivotTable's or "
    "PivotCache's metadata, a PivotTable's field layout, or one field's current "
    "filter state. "
    "Do not use when: you need to refresh a PivotTable/Cache — use "
    "pivot_refresh; you need to change a field's layout/aggregation/filter — use "
    "pivot_field. "
    "Preconditions: an open workbook_id; pivot_table is required for table-get / "
    "fields-list / field-get-filters; cache_index is required for cache-get; "
    "field is required for field-get-filters. "
    "Parameters: action='tables-list'|'table-get'|'cache-get'|'fields-list'|"
    "'field-get-filters'; workbook_id; pivot_table; cache_index; field; "
    "background (tables-list only). "
    "Behavior: read-only; only worksheet/Table-backed caches are supported — "
    "external, OLAP and Data Model sources are reported as unsupported_source "
    "rather than refreshed; shared caches are reported once with their affected "
    "tables; a missing required argument surfaces as a structured "
    "ExcelVbaMcpError. "
    "Result: PivotInventory (tables-list), PivotTableInfo (table-get), "
    "PivotCacheInfo (cache-get), list[PivotFieldInfo] (fields-list), "
    "PivotFieldFilters (field-get-filters) or a RunningJob handle (tables-list "
    "with background=true). "
    "Next actions: pivot_refresh to refresh; pivot_field to reconfigure a field."
)
_DESC_PIVOT_REFRESH = (
    "Purpose: Refresh one PivotTable or one PivotCache through one dispatch "
    "tool (never Workbook.RefreshAll). "
    "Use when: you need to refresh a single PivotTable (table-scoped) or a "
    "shared PivotCache (cache-scoped, affects every table using it). "
    "Do not use when: you need to read pivot metadata — use pivot_read; you "
    "need to change a field — use pivot_field. "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read; pivot_table is required for "
    "refresh-table; cache_index is required for refresh-cache. "
    "Parameters: action='refresh-table'|'refresh-cache'; workbook_id; "
    "pivot_table (refresh-table); cache_index (refresh-cache). "
    "Behavior: the cache source type is checked first — external, OLAP and Data "
    "Model sources return unsupported_source and are not refreshed; success is "
    "confirmed by Excel's truth signal (RefreshTable Boolean and/or an advanced "
    "RefreshDate), never by the absence of a COM exception. This tool is NOT "
    "idempotent: a re-refresh may read different source data. "
    "Guarded mutation (see xlvba://glossary). "
    "Result: PivotRefreshResult (affected_pivot_tables, RefreshDate before/after, "
    "refreshed status). "
    "Next actions: pivot_read action='table-get' to verify the refreshed state."
)
_DESC_PIVOT_FIELD = (
    "Purpose: Configure one PivotField's layout, position, aggregation or "
    "properties, or apply/clear its filter, through one dispatch tool. "
    "Use when: you need to move/reorder a field, set its aggregation, patch its "
    "caption/number format/show-all-items, or apply/clear a manual filter. "
    "Do not use when: you need to read pivot metadata — use pivot_read; you "
    "need to refresh — use pivot_refresh. "
    "Preconditions: writable session + snapshot + approval Approval gate (see "
    "xlvba://glossary) + postcondition re-read; pivot_table and field are "
    "required for every action except field-clear-filters with scope='table' "
    "(whole-PivotTable clear). "
    "Parameters: action='field-set-layout'|'field-set-position'|"
    "'field-set-aggregation'|'field-set-properties'|'field-apply-filter'|"
    "'field-clear-filters'; workbook_id; pivot_table; field; orientation + "
    "delete (field-set-layout); position (field-set-position); function + "
    "caption (field-set-aggregation); properties (field-set-properties); spec "
    "(field-apply-filter); scope (field-clear-filters: 'field' or 'table'). "
    "Behavior: 'data' orientation is rejected (use field-set-aggregation); "
    "field-set-layout with delete=true is destructive (PivotField.Delete); "
    "unsupported sources are rejected; a missing required argument surfaces as a "
    "structured ExcelVbaMcpError. "
    "Guarded mutation (see xlvba://glossary). "
    "Result: PivotFieldInfo (field-set-layout / field-set-position / "
    "field-set-aggregation / field-set-properties) or PivotFieldFilters "
    "(field-apply-filter / field-clear-filters). "
    "Next actions: pivot_read action='fields-list' to verify the field state."
)

# Read-only MCP resource descriptions (spec §6.6). The URIs are per-workbook
# templates so a client can read the model / VBA inventory / one component's
# source without mutating the workbook and without running VBA.
_DESC_RES_MODEL = (
    "Read-only Stage 1 workbook model (JSON) for an open workbook. Aggregates "
    "sheets, tables, names, shapes, controls and (optionally) VBA, with section "
    "statuses and a content hash. Never reads all-cells/Range data."
)
_DESC_RES_VBA_COMPONENTS = (
    "Read-only VBA component inventory (JSON) for an open workbook: name, type, "
    "line count, normalized source hash and designer/export flags. Reports "
    "vba_access instead of failing when VBA is unavailable or locked."
)
_DESC_RES_VBA_SOURCE = (
    "Read-only VBA source code (text) of one component of an open workbook. The "
    "workbook is never modified and no VBA is executed."
)
_DESC_RES_TEST_RUN = (
    "Read-only summary of one declarative test run (JSON): suite, overall status, "
    "baseline snapshot id, per-case statuses/steps/assertions and the artifact "
    "references. Never touches Excel and never re-runs anything."
)
# Description for the read-only xlvba_environment tool (registered by U2). Kept
# here with the other descriptions so the contract is stable; it is non-COM and
# side-effect-free and never authorizes a mutation.
_DESC_ENVIRONMENT = (
    "Purpose: Report the server's access mode, policy, build and capabilities "
    "without touching COM. "
    "Use when: you need to know the access mode, whether mutations need a grant, "
    "the snapshot gate, path roots, macro gates or the honest build (always "
    "null). "
    "Do not use when: you need worker/session liveness — use excel_health; for a "
    "pure liveness probe use server_health. "
    "Preconditions: none (non-COM, side-effect-free, read-only). "
    "Behavior: never creates snapshots, grants or classifications; "
    "policy.classification_is_authorization=false is stated explicitly; build is "
    "null (no 4-part source); excel.available=false when the worker is down. "
    "Result: EnvironmentReport (server, excel, capabilities, policy, "
    "glossary_uri, workbooks). "
    "Next actions: workbook action='open' / snapshot_create as the policy requires. "
    "Capabilities (this server provides): VBA project access, macros, pivot "
    "tables (database-source; manual filters), background jobs, recovery, and "
    "non-activating compile diagnostics — NOT compilation (sourced from "
    "EnvironmentCapabilities). "
    "Routing is capability-based advice only: this server does NOT detect whether "
    "any other MCP server is installed or connected. "
    "Prefer xlvba when the task needs the working-copy lifecycle, session-aware "
    "operations, snapshot/guarded mutation, VBA/macro management, pivot/cache "
    "diagnostics, background jobs, Excel health/recovery, or policy-aware "
    "execution. "
    "Use excel-mcp when it exposes an equivalent supported operation that does "
    "NOT need xlvba's session/snapshot/working-copy/recovery guarantees. "
    "Matrix: working copy or session binding -> xlvba; snapshot + guarded "
    "mutation -> xlvba; VBA/macro/pivot-cache/background/recovery -> xlvba; "
    "plain cell/range/sheet op with no session guarantee needed -> excel-mcp. "
    "Examples: 'read one cell from a disposable file' -> excel-mcp; 'edit a "
    "cell, snapshot, then recover the owned Excel process' -> xlvba."
)
_DESC_RES_GLOSSARY = (
    "Read-only glossary (Markdown) of the server's jargon: R13, BG06, preflight, "
    "truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, "
    "SUSPECTED_STALLED, interrupted_unknown, indeterminate, working copy and the "
    "snapshot gate. Static; never touches Excel."
)


@dataclasses.dataclass
class _Runtime:
    """Container for the per-server Excel runtime (owned by the lifespan)."""

    worker: ExcelComWorker
    app_adapter: ExcelAppAdapter
    session: SessionService
    snapshot: SnapshotService
    vba: VBAProjectService
    sheets: SheetsService
    tables: TablesService
    names: NamesService
    ranges: RangesService
    pivots: PivotTablesService
    shapes: ShapesService
    controls: ControlsService
    compile: CompileService
    macro: MacroService
    model: WorkbookModelService
    artifacts: ArtifactStore
    runs: RunRepository
    runner: TestRunner
    jobs: JobRegistry
    job_store: JobStore
    path_policy: PathPolicy
    approval: ApprovalPolicy
    audit: AuditLogger


def _configured_roots() -> dict[str, str]:
    """Return PathPolicy roots from ``EXCEL_VBA_MCP_ROOTS`` or a safe default.

    ``EXCEL_VBA_MCP_ROOTS`` is an ``os.pathsep``-separated list of directories.
    When unset, the current working directory and the system temp directory are
    used so the server is usable out of the box; production deployments should
    set an explicit allowlist.

    The artifact root is always registered under the ``artifacts`` logical name
    (required by :class:`~excel_vba_mcp.storage.artifact_store.ArtifactStore`) in
    addition to the generic ``root_N`` entries.
    """

    raw = os.environ.get(_ENV_ROOTS, "").strip()
    if raw:
        dirs = [part.strip() for part in raw.split(os.pathsep) if part.strip()]
    else:
        dirs = [str(Path.cwd()), tempfile.gettempdir()]
    roots = {f"root_{idx}": d for idx, d in enumerate(dirs)}
    roots[ARTIFACT_ROOT_NAME] = str(
        _resolve_runtime_path(_ENV_ARTIFACT_DIR, _default_artifact_dir())
    )
    # Log the resolved allowlist at startup so an operator can confirm which roots
    # are actually in force — including the cwd+tempdir fallback when
    # ``EXCEL_VBA_MCP_ROOTS`` is unset (otherwise the fallback is silent).
    logger.info("Allowed roots: %s", roots)
    return roots


def _resolve_runtime_path(env_name: str, default: Path) -> Path:
    raw = os.environ.get(env_name, "").strip()
    return Path(raw) if raw else default


def _resolve_owned_pid_dir() -> Path:
    """Resolve the stable per-user directory for the ownership marker.

    Honors ``EXCEL_VBA_MCP_OWNED_PID_DIR`` when set (a non-volatile root); otherwise
    uses :func:`default_owned_pid_dir` (``%LOCALAPPDATA%\\excel-vba-mcp``). This is
    deliberately **not** the volatile per-work ``%TEMP%`` working-copy dir so the
    marker survives a server restart / logoff-logon and the next start can reconcile
    an orphaned owned Excel precisely.
    """

    raw = os.environ.get(_ENV_OWNED_PID_DIR, "").strip()
    return Path(raw) if raw else default_owned_pid_dir()


def _build_runtime(config: Config) -> _Runtime:
    """Construct the Excel runtime from environment + typed configuration."""

    path_policy = PathPolicy(_configured_roots())
    app_adapter = ExcelAppAdapter()
    # The server-wide liveness ceiling is the authority that clamps every per-call
    # COM deadline (Block E); open_timeout_s is the per-op-class open authority.
    worker = ExcelComWorker(app_adapter, liveness_ceiling_s=config.server.liveness_ceiling_s)
    # The sole production ApprovalPolicy: full-access by default (mutations and
    # executions authorized without per-action grants), restricted as an explicit
    # opt-in. Grants/allowlist are wired separately (task_014).
    approval = ApprovalPolicy(mode=config.server.access_mode)
    audit = AuditLogger(_resolve_runtime_path(_ENV_AUDIT, _default_audit_path()))
    session = SessionService(
        worker,
        app_adapter,
        path_policy,
        approval,
        audit,
        work_dir=_resolve_runtime_path(_ENV_WORK_DIR, _default_work_dir()),
        open_timeout_s=config.server.open_timeout_s,
        owned_pid_dir=_resolve_owned_pid_dir(),
    )
    snapshot = SnapshotService(
        SnapshotStore(_resolve_runtime_path(_ENV_SNAPSHOT_DIR, _default_snapshot_dir())),
        session,
        worker,
        path_policy,
        audit,
        approval=approval,
    )
    vba = VBAProjectService(
        worker,
        session,
        path_policy,
        audit,
        snapshot_service=snapshot,
        approval=approval,
    )
    sheets = SheetsService(worker, session, snapshot, approval, path_policy, audit)
    tables = TablesService(worker, session, snapshot, approval, path_policy, audit)
    ranges = RangesService(worker, session, snapshot, approval, path_policy, audit)
    names = NamesService(worker, session, snapshot, approval, path_policy, audit)
    pivots = PivotTablesService(
        worker, session, audit, snapshot_service=snapshot, approval=approval
    )
    shapes = ShapesService(worker, session, snapshot, approval, path_policy, audit)
    controls = ControlsService(worker, session, snapshot, approval, path_policy, audit)
    compile_service = CompileService(worker, session, vba, path_policy, audit)
    # background_safe starts empty (deny-by-default, BG06): no procedure is
    # eligible for background execution until it is explicitly classified for the
    # current workbook hash and verified on a disposable copy.
    macro_service = MacroService(
        worker,
        session,
        approval,
        BackgroundSafetyRegistry(),
        audit,
        macro_timeout_s=config.server.macro_timeout_s,
    )
    model_service = WorkbookModelService(
        worker,
        session,
        sheets,
        tables,
        names,
        shapes,
        controls,
        vba,
        path_policy,
        audit,
        pivot_tables_service=pivots,
    )
    # Evidence + test-run persistence (pure filesystem; no COM).
    artifacts = ArtifactStore(
        _resolve_runtime_path(_ENV_ARTIFACT_DIR, _default_artifact_dir()),
        path_policy,
        audit,
    )
    runs = RunRepository(_resolve_runtime_path(_ENV_RUN_DIR, _default_run_dir()))
    # Job persistence (Block F): a RUNNING marker is written at job start and the
    # record is finalized on terminal. The store reconciles any non-terminal record
    # left by a prior process to interrupted_unknown (never completed, never
    # retried). The in-memory registry (Block E) owns the live result; the store
    # makes a completed result retrievable by job_id after a restart.
    job_store = JobStore(_resolve_runtime_path(_ENV_JOB_DIR, _default_job_dir()))
    if job_store.reconciled:
        logger.warning(
            "Reconciled %d interrupted job(s) to interrupted_unknown at startup.",
            len(job_store.reconciled),
        )
    # Minimal in-memory job registry (Block E): wraps the single-STA submit FIFO
    # and owns the result so a client timeout/disconnect never drops it. The full
    # JobStore + startup reconcile are Block F.
    jobs = JobRegistry(worker, store=job_store)
    # Let a successful recovery finalize this session's in-flight jobs to
    # INTERRUPTED_UNKNOWN (Block F recovery wiring).
    session.set_job_registry(jobs)
    runner = TestRunner(
        worker,
        session,
        snapshot,
        compile_service,
        macro_service,
        sheets,
        tables,
        names,
        artifacts,
        runs,
        path_policy,
        audit,
        shapes_service=shapes,
        ranges_service=ranges,
        pivot_tables_service=pivots,
    )
    return _Runtime(
        worker=worker,
        app_adapter=app_adapter,
        session=session,
        snapshot=snapshot,
        vba=vba,
        sheets=sheets,
        tables=tables,
        names=names,
        ranges=ranges,
        pivots=pivots,
        shapes=shapes,
        controls=controls,
        compile=compile_service,
        macro=macro_service,
        model=model_service,
        artifacts=artifacts,
        runs=runs,
        runner=runner,
        jobs=jobs,
        job_store=job_store,
        path_policy=path_policy,
        approval=approval,
        audit=audit,
    )


def _default_audit_path() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "audit.jsonl"


def _default_work_dir() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "work"


def _default_snapshot_dir() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "snapshots"


def _default_artifact_dir() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "artifacts"


def _default_run_dir() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "runs"


def _default_job_dir() -> Path:
    return Path(tempfile.gettempdir()) / "excel_vba_mcp" / "jobs"


@asynccontextmanager
async def lifespan(_server: MCPServer) -> AsyncIterator[_Runtime]:
    """Own the Excel runtime: start the worker, publish it, tear it down.

    The worker (and its hidden Excel instance) is started on entry and shut down
    in ``finally``. If the worker cannot start (e.g. Excel/pywin32 unavailable),
    the failure is logged to ``stderr`` and the server still starts so that
    ``server_health`` and tool listing keep working; workbook tools then raise a
    truthful ``WorkerUnavailableError`` when called. This keeps the STDIO startup
    smoke test robust across hosts.
    """

    try:
        config = load_config()
    except ValidationError as exc:
        logger.error("Invalid server configuration: %s", exc)
        raise
    runtime = _build_runtime(config)
    # Publish the read-only environment runtime (U2) and the in-memory snapshot
    # gate unconditionally (before the worker start) so ``xlvba_environment`` and
    # ``excel_get_status`` report truthfully even when the worker is down
    # (``excel.available=false``). This is non-COM and never starts the worker.
    runtime.session.set_snapshot_gate(runtime.snapshot.has_snapshot)
    environment_tools.set_environment_runtime(
        config=config.server,
        session=runtime.session,
        snapshot=runtime.snapshot,
        path_policy=runtime.path_policy,
    )
    runtime.audit.record(
        tool="server_startup",
        operation_class="config",
        params={"access_mode": config.server.access_mode.value},
        result={"version": __version__},
    )
    # Startup reconcile (Block A #4 + lifecycle): if a prior owned Excel is still
    # alive AND owned by a live server (identity match) do NOT launch a second
    # Excel — surface it and require an operator/agent decision. If the owner is
    # dead but the owned Excel is a verified orphan, reap exactly that orphan
    # (pid + create-time + image verified) and proceed. Never a wildcard; never kill
    # on pid-alone.
    pid_store = OwnedPidStore(_resolve_owned_pid_dir())
    refusal: str | None = None
    start_error: Exception | None = None
    try:
        refusal = await _start_worker_guarded(runtime, pid_store)
    except Exception as exc:  # noqa: BLE001 - degrade gracefully, never crash stdio
        start_error = exc
    if refusal is not None:
        logger.error("Excel worker startup refused: %s", refusal)
        runtime.audit.record(
            tool="server_startup_recovery",
            operation_class="config",
            params={"reason": refusal},
            result={"started": False},
        )
    elif start_error is not None:
        logger.warning("Excel worker failed to start: %s", start_error)
        await _safe_shutdown(runtime.worker)
    else:
        workbook_tools.set_session_service(runtime.session)
        test_tools.set_snapshot_service(runtime.snapshot)
        test_tools.set_test_runner(runtime.runner)
        vba_tools.set_vba_project_service(runtime.vba)
        sheet_tools.set_sheets_service(runtime.sheets)
        table_tools.set_tables_service(runtime.tables)
        range_tools.set_ranges_service(runtime.ranges)
        pivot_tools.set_pivot_tables_service(runtime.pivots)
        name_tools.set_names_service(runtime.names)
        shape_tools.set_shapes_service(runtime.shapes)
        control_tools.set_controls_service(runtime.controls)
        compile_tools.set_compile_service(runtime.compile)
        macro_tools.set_macro_service(runtime.macro)
        job_tools.set_job_registry(runtime.jobs)
        workbook_tools.set_workbook_model_service(runtime.model)
    try:
        yield runtime
    finally:
        workbook_tools.set_session_service(None)
        workbook_tools.set_workbook_model_service(None)
        environment_tools.clear_environment_runtime()
        test_tools.set_snapshot_service(None)
        test_tools.set_test_runner(None)
        vba_tools.set_vba_project_service(None)
        sheet_tools.set_sheets_service(None)
        table_tools.set_tables_service(None)
        range_tools.set_ranges_service(None)
        pivot_tools.set_pivot_tables_service(None)
        name_tools.set_names_service(None)
        shape_tools.set_shapes_service(None)
        control_tools.set_controls_service(None)
        compile_tools.set_compile_service(None)
        macro_tools.set_macro_service(None)
        job_tools.set_job_registry(None)
        await _safe_shutdown(runtime.worker)
        runtime.audit.close()


async def _safe_shutdown(worker: ExcelComWorker) -> None:
    """Shut the worker down without masking the original error."""

    try:
        await worker.shutdown()
    except Exception as exc:  # noqa: BLE001 - shutdown is best-effort
        logger.warning("Excel worker shutdown error: %s", exc)


def _exe_basename(path: str) -> str:
    return path.replace("/", "\\").rsplit("\\", 1)[-1].strip().lower()


def _exe_matches(expected: str | None, actual: str | None) -> bool:
    """Whether two image paths refer to the same exe (basename, case-insensitive).

    An unknown ``expected`` (legacy record) or an unreadable ``actual`` never blocks
    the match — the create-time is the strong identity signal; the exe is a
    secondary confirmation.
    """

    if expected is None or actual is None:
        return True
    return _exe_basename(expected) == _exe_basename(actual)


def _reconcile_creating_record(
    record: OwnedProcessRecord,
    *,
    store: OwnedPidStore,
    pid_alive: Callable[[int], bool],
    create_time: Callable[[int], int | None],
    image_name: Callable[[int], str | None],
    terminate_verified: Callable[[int], bool],
    current_session_id: Callable[[], int | None] | None,
) -> str | None:
    """Reconcile a ``state="creating"`` marker (crash in the create->assign window).

    The owned Excel may not have existed / been assigned to the job when the server
    died, so the excel pid may be unknown (``0``). Apply the lifecycle conservatively:

    * **Owner alive + identity match** => a launch is in progress under a live owner;
      **refuse** a second launch.
    * **Owner dead + excel pid unknown (0)** => an unidentifiable orphan may exist;
      **refuse** a second launch (never pid-kill) and keep the marker.
    * **Owner dead + known excel pid not alive** => the launch left no orphan; clear
      and proceed.
    * **Owner dead + known excel alive + verified** (image ``EXCEL.EXE`` + create-time
      match + same session) => reap exactly that orphan, confirm, clear, proceed.
    * **Owner dead + known excel alive but unverified / cross-session** => **refuse**
      (never kill an unverified / cross-session process), keep the marker.
    """

    owner = record.owner
    excel = record.excel

    owner_pid_alive = owner.pid > 0 and pid_alive(owner.pid)
    owner_create_ok = (
        owner.create_time_filetime is None
        or create_time(owner.pid) == owner.create_time_filetime
    )
    owner_exe_ok = _exe_matches(owner.exe, image_name(owner.pid))
    if owner.pid > 0 and owner_pid_alive and owner_create_ok and owner_exe_ok:
        return (
            f"a server launch is in progress (creating) under live owner pid "
            f"{owner.pid}; refusing to start a second Excel instance — an "
            "operator/agent decision is required"
        )

    # Owner is dead. Unknown excel pid => an unidentifiable orphan may exist.
    if excel.pid <= 0:
        logger.warning(
            "creating marker found with a dead owner (pid %d) and no excel pid; an "
            "unidentifiable orphan may exist — refusing a second launch",
            owner.pid,
        )
        return (
            "a prior server crashed while creating its owned Excel (no verified excel "
            "pid was recorded); refusing to start a second Excel because an "
            "unidentifiable orphan may exist — an operator/agent decision is required"
        )

    if not pid_alive(excel.pid):
        # The known owned pid did not survive the crash: no orphan to reap.
        store.clear()
        return None

    excel_is_excel = _exe_basename(image_name(excel.pid) or "") == "excel.exe"
    if not excel_is_excel:
        logger.warning(
            "creating marker excel pid %d is alive but not EXCEL.EXE (pid reuse); "
            "clearing marker and proceeding",
            excel.pid,
        )
        store.clear()
        return None

    excel_create_ok = (
        excel.create_time_filetime is not None
        and create_time(excel.pid) == excel.create_time_filetime
    )
    if not excel_create_ok:
        return (
            f"prior owned Excel pid {excel.pid} is alive (EXCEL.EXE) but its identity "
            "could not be verified (create-time mismatch/unknown); refusing to start a "
            "second Excel and refusing to kill an unverified process — an "
            "operator/agent decision is required"
        )

    current_sess = current_session_id() if current_session_id is not None else None
    if (
        excel.session_id is not None
        and current_sess is not None
        and excel.session_id != current_sess
    ):
        return (
            f"prior owned Excel pid {excel.pid} is alive in Windows session "
            f"{excel.session_id} but this server is in session {current_sess}; "
            "refusing to kill a cross-session process — an operator/agent decision "
            "is required"
        )

    terminated = terminate_verified(excel.pid)
    still_alive = pid_alive(excel.pid)
    if terminated and not still_alive:
        logger.warning(
            "reaped verified orphaned owned Excel pid %d left by a creating-window "
            "crash (owner pid %d dead); clearing marker and proceeding",
            excel.pid,
            owner.pid,
        )
        store.clear()
        return None
    return (
        f"verified orphaned owned Excel pid {excel.pid} could not be terminated "
        f"(terminated={terminated}, still_alive={still_alive}); refusing to start a "
        "second Excel — an operator/agent decision is required"
    )


def _reconcile_prior_owned_pid(
    store: OwnedPidStore,
    *,
    pid_alive: Callable[[int], bool],
    create_time: Callable[[int], int | None],
    image_name: Callable[[int], str | None],
    terminate_verified: Callable[[int], bool],
    current_session_id: Callable[[], int | None] | None = None,
) -> str | None:
    """Decide whether it is safe to start the worker, reaping a verified orphan.

    Reads the persisted ownership record and applies the lifecycle reconcile:

    * **Owner alive + identity match** (pid + create-time + exe) => the prior server
      still owns its Excel; **refuse** a second launch (return a reason).
    * **Owner dead + owned Excel alive + identity verified** (image ``EXCEL.EXE`` +
      create-time match + **same Windows session**) => the prior server was
      tree-killed / crashed and left an orphan; **terminate exactly that orphan**
      (``terminate_verified`` — pid + create-time + image gated, never a wildcard),
      confirm death, clear the marker and proceed.
    * **Cross-session orphan** (owner dead, Excel verified, but the record's
      ``excel.session_id`` differs from ``current_session_id()``) => the orphan lives
      in a different Windows session; **refuse** (never kill a cross-session process,
      keep the marker) — an operator/agent decision is required.
    * **Pid reuse** (owner or excel create-time mismatch, or the excel pid is now a
      non-Excel process) => **never kill on pid-alone**; clear a stale marker and
      proceed, or refuse when an alive ``EXCEL.EXE`` cannot be identity-verified.

    Pure except for the injected probes, the orphan termination and clearing a dead
    marker, so it is unit tested without launching Excel.

    Args:
        current_session_id: Optional probe returning this server's current Windows
            session id. When ``None`` (or it returns ``None``) the session check is
            skipped (unknown session => fall back to the create-time/image gate). A
            known mismatch blocks the kill conservatively.

    Returns:
        ``None`` when startup may proceed, or a refusal reason string when a prior
        owned Excel is still owned by a live server (or an unverifiable / cross-session
        Excel is alive and must not be killed / double-launched).
    """

    record = store.read_record()
    if record is None:
        return None

    # A ``creating`` marker was written before DispatchEx (the owned Excel did not
    # exist / was not yet assigned to the job). A crash in that window leaves no
    # verified excel pid, so reconcile conservatively (refuse a second launch; reap
    # only a verified identity; never a pid-only kill).
    if record.state == STATE_CREATING:
        return _reconcile_creating_record(
            record,
            store=store,
            pid_alive=pid_alive,
            create_time=create_time,
            image_name=image_name,
            terminate_verified=terminate_verified,
            current_session_id=current_session_id,
        )

    owner = record.owner
    excel = record.excel

    # --- Owner liveness + identity (pid + create-time + exe) ---
    owner_pid_alive = owner.pid > 0 and pid_alive(owner.pid)
    owner_create_ok = (
        owner.create_time_filetime is None
        or create_time(owner.pid) == owner.create_time_filetime
    )
    owner_exe_ok = _exe_matches(owner.exe, image_name(owner.pid))
    if owner.pid > 0 and owner_pid_alive and owner_create_ok and owner_exe_ok:
        return (
            f"prior owned Excel pid {excel.pid} is still owned by live server pid "
            f"{owner.pid}; refusing to start a second Excel instance — an "
            "operator/agent decision is required (recover/adopt it, or restart)"
        )

    # --- Owner is dead (or its pid was reused by a different process) ---
    if not (excel.pid > 0 and pid_alive(excel.pid)):
        # Nothing left to reap: a dead owned pid must not block a fresh start.
        store.clear()
        return None

    # The owned Excel pid is alive but the owner is dead => likely orphan. Verify
    # identity before any kill (never kill on pid-alone).
    excel_is_excel = _exe_basename(image_name(excel.pid) or "") == "excel.exe"
    if not excel_is_excel:
        # The recorded pid is now a different (non-Excel) process: pid reuse. Never
        # kill it; clear the stale marker and proceed.
        logger.warning(
            "owned pid marker excel pid %d is alive but not EXCEL.EXE (pid reuse); "
            "clearing marker and proceeding",
            excel.pid,
        )
        store.clear()
        return None

    excel_create_ok = (
        excel.create_time_filetime is not None
        and create_time(excel.pid) == excel.create_time_filetime
    )
    if not excel_create_ok:
        # An EXCEL.EXE is alive on the recorded pid but its create-time does not
        # match the record (or is unknown): cannot verify it is our orphan. Never
        # kill an unverified process and never double-launch — refuse.
        return (
            f"prior owned Excel pid {excel.pid} is alive (EXCEL.EXE) but its "
            "identity could not be verified (create-time mismatch/unknown); "
            "refusing to start a second Excel and refusing to kill an unverified "
            "process — an operator/agent decision is required"
        )

    # Verified orphan: owner dead, image EXCEL.EXE, create-time matches. Before any
    # kill, enforce the session boundary: if the record's excel session id is known
    # and differs from this server's current session, the orphan lives in a different
    # Windows session — never kill cross-session; refuse conservatively (keep marker).
    current_sess = current_session_id() if current_session_id is not None else None
    if (
        excel.session_id is not None
        and current_sess is not None
        and excel.session_id != current_sess
    ):
        logger.warning(
            "owned Excel pid %d is in session %d but this server is in session %d; "
            "refusing to kill a cross-session process",
            excel.pid,
            excel.session_id,
            current_sess,
        )
        return (
            f"prior owned Excel pid {excel.pid} is alive in Windows session "
            f"{excel.session_id} but this server is in session {current_sess}; "
            "refusing to kill a cross-session process — an operator/agent decision "
            "is required"
        )

    # Verified orphan in the same session: terminate exactly that pid (image-gated)
    # and confirm death before clearing the marker.
    terminated = terminate_verified(excel.pid)
    still_alive = pid_alive(excel.pid)
    if terminated and not still_alive:
        logger.warning(
            "reaped verified orphaned owned Excel pid %d (owner pid %d dead); "
            "clearing marker and proceeding",
            excel.pid,
            owner.pid,
        )
        store.clear()
        return None
    return (
        f"verified orphaned owned Excel pid {excel.pid} could not be terminated "
        f"(terminated={terminated}, still_alive={still_alive}); refusing to start "
        "a second Excel — an operator/agent decision is required"
    )


async def _start_worker_guarded(runtime: _Runtime, store: OwnedPidStore) -> str | None:
    """Reconcile-then-start the worker, returning a refusal or ``None``.

    Reconciles the prior ownership record (refuse a live owner; reap a verified
    orphan; clear a stale marker). When startup may proceed a ``state="creating"``
    marker is persisted **before** the launch (before ``DispatchEx`` inside
    ``worker.start``) so a crash in the create->assign window is recoverable: the next
    start refuses a second launch (an unidentifiable orphan may exist) and reaps only a
    verified identity — never a pid-only kill. On success the marker is overwritten
    with the full ``running`` record (owned pid + owner identity); on a fail-closed
    start (no orphan by construction) it is cleared so a clean restart is not blocked.
    A ``worker.start()`` failure propagates to the caller (which degrades gracefully +
    safe-shuts the worker); it is not a reconcile refusal.
    """

    adapter = ExcelAppAdapter
    refusal = _reconcile_prior_owned_pid(
        store,
        pid_alive=adapter._pid_alive,
        create_time=adapter._process_create_time_filetime,
        image_name=adapter._image_name_of_pid,
        terminate_verified=adapter._terminate_pid_verified,
        current_session_id=lambda: adapter._capture_process_identity(os.getpid())[2],
    )
    if refusal is not None:
        return refusal
    # Persist a creating marker BEFORE the owned Excel exists (before DispatchEx),
    # so a crash in the create->assign window leaves a durable record to reconcile.
    owner_pid = os.getpid()
    o_ct, o_exe, o_sess = adapter._capture_process_identity(owner_pid)
    store.write_record(
        build_creating_record(
            owner_pid=owner_pid,
            owner_create_time=o_ct,
            owner_exe=o_exe,
            owner_session_id=o_sess,
        )
    )
    try:
        await runtime.worker.start()
    except Exception:
        # Fail-closed start (e.g. job-assign failure) leaves no orphan by
        # construction; clear the creating marker so a clean restart is not blocked.
        store.clear()
        raise
    new_pid = runtime.worker.get_status().owned_pid
    if new_pid is not None:
        e_ct, e_exe, e_sess = adapter._capture_process_identity(new_pid)
        store.write_record(
            build_owned_record(
                owner_pid=owner_pid,
                excel_pid=new_pid,
                owner_create_time=o_ct,
                owner_exe=o_exe,
                owner_session_id=o_sess,
                excel_create_time=e_ct,
                excel_exe=e_exe,
                excel_session_id=e_sess,
            )
        )
    else:
        # Started but no owned pid tracked: nothing to reconcile later.
        store.clear()
    return None


def _configure_logging() -> None:
    """Send logs to ``stderr`` only; ``stdout`` is reserved for the protocol."""

    logging.basicConfig(
        level=logging.INFO,
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )


# Truthful MCP ``ToolAnnotations`` (U3; plan_03 §11; B11 / WP10). The three
# frozensets encode the per-tool triple derived from the WP0 truth matrix
# (``tests/fixtures/wp0_baseline.json`` -> ``tool_annotations``) using MCP default
# semantics: readOnly = all actions read-only; destructive = any action
# destructive; idempotent = all actions idempotent. Every ``add_tool`` call emits
# all three hints as explicit booleans (never ``None`` — R-Q12 / A2), so an
# additive tool is not mis-advertised as destructive under the MCP
# ``destructiveHint=true`` default. These are hints only (client display is
# best-effort; Kilo approval is name-glob based) — they never change enforcement.
#
# ``range_structural`` is destructive-only (NOT idempotent — H4: re-insert
# duplicates; unmerge discards merges). ``vba_replace_component_source`` is
# destructive-only (NOT idempotent — R-Q3: re-call with the same hash fails the
# revision guard). The dual-hint set (destructive AND idempotent) is exactly
# ``{range_merge, pivot_field}`` (R-Q4).
#
# ``xlvba_environment`` and ``server_health`` are (read_only=True,
# destructive=False, idempotent=True) but carry INLINE annotations at their own
# call sites (A1 / R-Q10); they are intentionally NOT members of these frozensets
# and do not flow through ``_annotations``.
_ANN_READ_ONLY = frozenset(
    {
        "workbook_read",
        "excel_health",
        "snapshot_list",
        "snapshot_diff",
        "vba_read",
        "job_get",
        "sheet_read",
        "table_read",
        "pivot_read",
        "range_read",
        "name_read",
    }
)
_ANN_DESTRUCTIVE = frozenset(
    {
        "workbook",
        "excel_save_as",
        "sheet_delete",
        "range_merge",
        "range_structural",
        "excel_recover",
        "snapshot_restore",
        "table_delete_rows",
        "table_delete_column",
        "name_delete",
        "range_delete",
        "vba_replace_component_source",
        "vba_insert_code",
        "macro_run",
        "test",
        "pivot_field",
    }
)
_ANN_IDEMPOTENT = frozenset(
    {
        "workbook_read",
        "excel_health",
        "snapshot_list",
        "snapshot_diff",
        "vba_read",
        "job_get",
        "job_request_cancel",
        "sheet_read",
        "table_read",
        "pivot_read",
        "pivot_refresh",
        "pivot_field",
        "range_read",
        "range_merge",
        "name",
        "name_read",
        "shape",
        "control",
        "macro",
    }
)


def _annotations(name: str) -> ToolAnnotations:
    """Build truthful :class:`ToolAnnotations` for one tool by name (U3 §11).

    All three hints are emitted as explicit booleans derived from the WP0 truth
    matrix frozensets (R-Q12 / A2): an absent hint is ``False``, never ``None``,
    so the MCP ``destructiveHint=true`` default cannot mis-advertise an additive
    tool as destructive.
    """

    return ToolAnnotations(
        read_only_hint=name in _ANN_READ_ONLY,
        destructive_hint=name in _ANN_DESTRUCTIVE,
        idempotent_hint=name in _ANN_IDEMPOTENT,
    )


def _register_tools(server: MCPServer) -> None:
    """Register the base workbook/session tools on ``server``."""
    server.add_tool(
        workbook,
        name="workbook",
        description=_DESC_WORKBOOK,
        annotations=_annotations("workbook"),
    )
    server.add_tool(
        workbook_read,
        name="workbook_read",
        description=_DESC_WORKBOOK_READ,
        annotations=_annotations("workbook_read"),
    )
    server.add_tool(
        excel_save_as,
        name="excel_save_as",
        description=_DESC_SAVE_AS,
        annotations=_annotations("excel_save_as"),
    )
    server.add_tool(
        excel_health,
        name="excel_health",
        description=_DESC_HEALTH,
        annotations=_annotations("excel_health"),
    )
    server.add_tool(
        excel_recover,
        name="excel_recover",
        description=_DESC_EXCEL_RECOVER,
        annotations=_annotations("excel_recover"),
    )
    server.add_tool(
        snapshot_create,
        name="snapshot_create",
        description=_DESC_SNAPSHOT_CREATE,
        annotations=_annotations("snapshot_create"),
    )
    server.add_tool(
        snapshot_list,
        name="snapshot_list",
        description=_DESC_SNAPSHOT_LIST,
        annotations=_annotations("snapshot_list"),
    )
    server.add_tool(
        snapshot_restore,
        name="snapshot_restore",
        description=_DESC_SNAPSHOT_RESTORE,
        annotations=_annotations("snapshot_restore"),
    )
    server.add_tool(
        snapshot_diff,
        name="snapshot_diff",
        description=_DESC_SNAPSHOT_DIFF,
        annotations=_annotations("snapshot_diff"),
    )
    server.add_tool(
        test,
        name="test",
        description=_DESC_TEST,
        annotations=_annotations("test"),
    )
    server.add_tool(
        vba_read,
        name="vba_read",
        description=_DESC_VBA_READ,
        annotations=_annotations("vba_read"),
    )
    server.add_tool(
        vba_export_project,
        name="vba_export_project",
        description=_DESC_VBA_EXPORT_PROJECT,
        annotations=_annotations("vba_export_project"),
    )
    server.add_tool(
        vba_replace_component_source,
        name="vba_replace_component_source",
        description=_DESC_VBA_REPLACE_COMPONENT_SOURCE,
        annotations=_annotations("vba_replace_component_source"),
    )
    server.add_tool(
        vba_insert_code,
        name="vba_insert_code",
        description=_DESC_VBA_INSERT_CODE,
        annotations=_annotations("vba_insert_code"),
    )
    server.add_tool(
        vba_compile_diagnostics,
        name="vba_compile_diagnostics",
        description=_DESC_VBA_COMPILE_DIAGNOSTICS,
        annotations=_annotations("vba_compile_diagnostics"),
    )
    server.add_tool(
        macro,
        name="macro",
        description=_DESC_MACRO,
        annotations=_annotations("macro"),
    )
    server.add_tool(
        macro_run,
        name="macro_run",
        description=_DESC_MACRO_RUN,
        annotations=_annotations("macro_run"),
    )
    server.add_tool(
        job_get, name="job_get", description=_DESC_JOB_GET, annotations=_annotations("job_get")
    )
    server.add_tool(
        job_request_cancel,
        name="job_request_cancel",
        description=_DESC_JOB_REQUEST_CANCEL,
        annotations=_annotations("job_request_cancel"),
    )
    server.add_tool(
        sheet,
        name="sheet",
        description=_DESC_SHEET,
        annotations=_annotations("sheet"),
    )
    server.add_tool(
        sheet_read,
        name="sheet_read",
        description=_DESC_SHEET_READ,
        annotations=_annotations("sheet_read"),
    )
    server.add_tool(
        sheet_delete,
        name="sheet_delete",
        description=_DESC_SHEET_DELETE,
        annotations=_annotations("sheet_delete"),
    )
    server.add_tool(
        table,
        name="table",
        description=_DESC_TABLE,
        annotations=_annotations("table"),
    )
    server.add_tool(
        table_read,
        name="table_read",
        description=_DESC_TABLE_READ,
        annotations=_annotations("table_read"),
    )
    server.add_tool(
        table_delete_rows,
        name="table_delete_rows",
        description=_DESC_TABLE_DELETE_ROWS,
        annotations=_annotations("table_delete_rows"),
    )
    server.add_tool(
        table_delete_column,
        name="table_delete_column",
        description=_DESC_TABLE_DELETE_COLUMN,
        annotations=_annotations("table_delete_column"),
    )
    server.add_tool(
        pivot_read,
        name="pivot_read",
        description=_DESC_PIVOT_READ,
        annotations=_annotations("pivot_read"),
    )
    server.add_tool(
        pivot_refresh,
        name="pivot_refresh",
        description=_DESC_PIVOT_REFRESH,
        annotations=_annotations("pivot_refresh"),
    )
    server.add_tool(
        pivot_field,
        name="pivot_field",
        description=_DESC_PIVOT_FIELD,
        annotations=_annotations("pivot_field"),
    )
    server.add_tool(
        range_read,
        name="range_read",
        description=_DESC_RANGE_READ,
        annotations=_annotations("range_read"),
    )
    server.add_tool(
        range_write,
        name="range_write",
        description=_DESC_RANGE_WRITE,
        annotations=_annotations("range_write"),
    )
    server.add_tool(
        range_structural,
        name="range_structural",
        description=_DESC_RANGE_STRUCTURAL,
        annotations=_annotations("range_structural"),
    )
    server.add_tool(
        range_delete,
        name="range_delete",
        description=_DESC_RANGE_DELETE,
        annotations=_annotations("range_delete"),
    )
    server.add_tool(
        range_merge,
        name="range_merge",
        description=_DESC_RANGE_MERGE,
        annotations=_annotations("range_merge"),
    )
    server.add_tool(
        name,
        name="name",
        description=_DESC_NAME,
        annotations=_annotations("name"),
    )
    server.add_tool(
        name_read,
        name="name_read",
        description=_DESC_NAME_READ,
        annotations=_annotations("name_read"),
    )
    server.add_tool(
        name_delete,
        name="name_delete",
        description=_DESC_NAME_DELETE,
        annotations=_annotations("name_delete"),
    )
    server.add_tool(
        shape,
        name="shape",
        description=_DESC_SHAPE,
        annotations=_annotations("shape"),
    )
    server.add_tool(
        control,
        name="control",
        description=_DESC_CONTROL,
        annotations=_annotations("control"),
    )
    server.add_tool(
        xlvba_environment,
        name="xlvba_environment",
        description=_DESC_ENVIRONMENT,
        annotations=ToolAnnotations(
            read_only_hint=True, destructive_hint=False, idempotent_hint=True
        ),
    )


def _find_domain_cause(exc: BaseException) -> ExcelVbaMcpError | None:
    """Return the first :class:`ExcelVbaMcpError` in ``exc``'s cause chain.

    Cycle-safe (a seen-id set guards a self-referential ``__cause__`` chain).
    Returns ``None`` when no domain error is present so callers re-raise the
    original exception unchanged.
    """

    seen: set[int] = set()
    current: BaseException | None = exc
    while current is not None and id(current) not in seen:
        if isinstance(current, ExcelVbaMcpError):
            return current
        seen.add(id(current))
        current = current.__cause__
    return None


def _originating_cause(exc: BaseException) -> BaseException:
    """Return the deepest ``__cause__`` in ``exc``'s chain (cycle-safe).

    The MCP SDK wraps an unhandled tool exception in
    ``UnexpectedToolError("Error executing tool <name>")``; the real failure is
    the originating cause at the bottom of the ``__cause__`` chain. Used by the
    ``call_tool`` INTERNAL_ERROR fallback so the synthesized envelope reports the
    actual failure text and exception class, never the SDK wrapper's
    "Error executing tool" message (U6-B; plan_05 §5.2 / N1).
    """

    seen: set[int] = set()
    current: BaseException = exc
    while current.__cause__ is not None and id(current) not in seen:
        seen.add(id(current))
        current = current.__cause__
    return current


def _jsonable_validation_errors(error: ValidationError) -> list[dict[str, Any]]:
    """Return ``error.errors()`` as a JSON-serializable list (audit finding C-1).

    ``ValidationError.errors()`` can embed non-JSON-serializable objects:

    * ``ctx['error']`` holds the raw exception for errors raised by a
      ``model_validator`` (for example the range-size guards in
      ``RangeWriteRequest`` / ``FormulaWriteRequest``), and
    * ``input`` holds the raw input value, which for nested models is the
      instantiated model object (for example a ``CellValue``).

    Embedding either verbatim in the error ``context`` made ``json.dumps`` raise
    ``TypeError`` at the MCP boundary, so the client received an unstructured
    error and the structured-error contract was broken. Every entry is copied and
    any non-primitive value (at the top level and inside ``ctx``) is replaced by
    its ``str()`` form, so the payload is always serializable while keeping the
    diagnostic detail.
    """

    def _jsonable(value: Any) -> Any:
        if isinstance(value, (str, int, float, bool, type(None))):
            return value
        if isinstance(value, dict):
            return {str(key): _jsonable(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return [_jsonable(item) for item in value]
        return str(value)

    return [_jsonable(dict(entry)) for entry in error.errors()]


def _validation_error_as_domain(exc: BaseException) -> ExcelVbaMcpError | None:
    """Translate a pydantic ``ValidationError`` in ``exc``'s cause chain.

    The MCP SDK reports an input-schema mismatch as a bare
    ``ToolError("Error executing tool <name>: ...")`` whose ``__cause__`` is the
    ``ValidationError``; without this translation the client sees an unstructured
    error with no ``code`` (audit finding G-9 / F-5). When a ``ValidationError``
    is present it is wrapped into an :class:`ExcelVbaMcpError` so every tool
    failure is structured. The closed spec §11 vocabulary has no dedicated
    argument-validation code, so this reuses ``INVALID_RANGE`` — the same code the
    tool modules already use for malformed arguments (``_require_bool`` /
    ``_require_range_text`` / ``_require_shift`` / ``_parse_suite``) — rather than
    inventing a new one.
    """

    seen: set[int] = set()
    current: BaseException | None = exc
    while current is not None and id(current) not in seen:
        if isinstance(current, ValidationError):
            return ExcelVbaMcpError(
                "The tool arguments do not match the input schema.",
                code=ErrorCode.INVALID_RANGE,
                context={
                    "reason": "invalid_arguments",
                    "validation": _jsonable_validation_errors(current),
                },
            )
        seen.add(id(current))
        current = current.__cause__
    return None


class ExcelVbaMcpServer(MCPServer[_Runtime]):
    """MCPServer that surfaces domain errors as structured ``ToolError`` payloads.

    The MCP SDK wraps any non-``ToolError`` exception raised by a tool into
    ``UnexpectedToolError("Error executing tool <name>")``, discarding the
    :class:`ExcelVbaMcpError` ``code`` and ``context``. Argument-schema failures
    (``pydantic.ValidationError``) are likewise turned into a bare
    ``ToolError("Error executing tool <name>: ...")`` with no ``code``. This
    subclass catches that wrapper, locates the underlying domain error in the
    ``__cause__`` chain and re-raises a ``ToolError`` whose message is the JSON of
    ``to_dict()`` so real MCP clients can observe the stable ``errorCode`` and
    structured ``details``. When the cause is a pydantic ``ValidationError`` (the
    caller's arguments do not match the input schema) it is translated into an
    ``ExcelVbaMcpError`` (code ``INVALID_RANGE``, matching the tool modules'
    malformed-argument code) so **every** tool failure reaches the client
    structured (audit finding G-9 / F-5: no unstructured ``UnexpectedToolError``
    leak). Any other (non-domain, non-validation) exception is synthesized into a
    guaranteed ``INTERNAL_ERROR`` envelope from the originating cause (U6-B;
    plan_05 §5.2 / N1) so no bare "Error executing tool" text ever reaches a
    client.
    """

    async def call_tool(
        self,
        name: str,
        arguments: dict[str, Any],
        context: Context[_Runtime, Any] | None = None,
    ) -> CallToolResult | InputRequiredResult:
        try:
            return await super().call_tool(name, arguments, context)
        except ToolError as exc:
            cause = _find_domain_cause(exc)
            if cause is None:
                # No domain error in the chain: a pydantic ValidationError (argument
                # schema mismatch) is the one unanticipated-but-caller-facing failure
                # the SDK emits without a code; translate it so the client always
                # receives a structured payload.
                cause = _validation_error_as_domain(exc)
                if cause is None:
                    # Neither a domain nor a validation cause: an unhandled non-domain
                    # exception. Synthesize a guaranteed structured INTERNAL_ERROR
                    # envelope so no bare "Error executing tool" leak reaches the
                    # client (U6-B; plan_05 §5.2 / N1). The caught ``exc`` is the
                    # SDK's ``UnexpectedToolError`` whose own message is
                    # "Error executing tool <name>"; use the originating (deepest)
                    # ``__cause__`` so the message is the real failure text and the
                    # reported exception class is the actual one, never the wrapper.
                    origin = _originating_cause(exc)
                    fallback = ExcelVbaMcpError(
                        str(origin),
                        code=ErrorCode.INTERNAL_ERROR,
                        context={
                            "reason": "unhandled_exception",
                            "exception": origin.__class__.__name__,
                        },
                    )
                    raise ToolError(
                        json.dumps(fallback.to_dict(), ensure_ascii=False)
                    ) from exc
            raise ToolError(json.dumps(cause.to_dict(), ensure_ascii=False)) from cause


_configure_logging()

mcp: ExcelVbaMcpServer = ExcelVbaMcpServer("excel-vba-mcp", version=__version__, lifespan=lifespan)


@mcp.tool(
    annotations=ToolAnnotations(
        read_only_hint=True, destructive_hint=False, idempotent_hint=True
    )
)
def server_health() -> dict[str, str]:
    """Purpose: Read-only liveness probe (never touches Excel/COM).

    Do not use when: you need worker/session state (use excel_health) or access
    mode / policy / build (use xlvba_environment) — liveness is not policy.
    """

    return {"status": "ok", "service": "excel-vba-mcp", "version": __version__}


# Register the base workbook/session tools at import time so they are present
# regardless of whether the server is started via ``main()`` here or directly via
# ``excel_vba_mcp.__main__`` (which calls ``mcp.run`` on this module's ``mcp``).
_register_tools(mcp)


# ---------------------------------------------------------------------------
# Read-only MCP resources (spec §6.6; plan task_032)
# ---------------------------------------------------------------------------
# These resources expose the Stage 1 model and VBA inventory/source as
# read-only MCP resources. They are per-workbook URI templates so a client can
# read one workbook's model without mutating it and without running VBA:
#
#   excel://workbooks/{workbook_id}/model
#   excel://workbooks/{workbook_id}/vba/components
#   excel://workbooks/{workbook_id}/vba/components/{component_name}
#
# The first two return JSON; the third returns the component's VBA source text.
def _vba_access_from_error(exc: ExcelVbaMcpError) -> VbaAccessStatus:
    """Map a VBA service error to the model's ``vba_access`` vocabulary."""

    if isinstance(exc, AccessVbomDisabledError):
        return VbaAccessStatus.VBOM_DISABLED
    if isinstance(exc, VbaProjectLockedError):
        return VbaAccessStatus.LOCKED
    return VbaAccessStatus.UNAVAILABLE


@mcp.resource(
    "excel://workbooks/{workbook_id}/model",
    name="workbook_model",
    description=_DESC_RES_MODEL,
    mime_type="application/json",
)
async def workbook_model_resource(workbook_id: str) -> str:
    """Read-only Stage 1 workbook model as JSON (never mutates / runs VBA)."""

    service = workbook_tools.get_workbook_model_service()
    model = await service.get_workbook_model(workbook_id)
    return model.model_dump_json(indent=2)


@mcp.resource(
    "excel://workbooks/{workbook_id}/vba/components",
    name="workbook_vba_components",
    description=_DESC_RES_VBA_COMPONENTS,
    mime_type="application/json",
)
async def workbook_vba_components_resource(workbook_id: str) -> str:
    """Read-only VBA component inventory as JSON (reports vba_access on failure)."""

    vba_service = vba_tools.get_vba_project_service()
    try:
        components = await vba_service.list_components(workbook_id)
    except ExcelVbaMcpError as exc:
        payload: dict[str, object] = {
            "workbook_id": workbook_id,
            "vba_access": _vba_access_from_error(exc).value,
            "components": [],
            "diagnostics": [str(exc)],
        }
        return json.dumps(payload, ensure_ascii=False, indent=2)

    payload = {
        "workbook_id": workbook_id,
        "vba_access": VbaAccessStatus.AVAILABLE.value,
        "components": [component.model_dump(mode="json") for component in components],
        "diagnostics": [],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


@mcp.resource(
    "excel://workbooks/{workbook_id}/vba/components/{component_name}",
    name="workbook_vba_component_source",
    description=_DESC_RES_VBA_SOURCE,
    mime_type="text/plain",
)
async def workbook_vba_component_source_resource(workbook_id: str, component_name: str) -> str:
    """Read-only VBA source text of one component (never mutates / runs VBA)."""

    vba_service = vba_tools.get_vba_project_service()
    try:
        source = await vba_service.read_component(workbook_id, component_name)
    except ExcelVbaMcpError as exc:
        raise ResourceError(json.dumps(exc.to_dict())) from exc
    return source.source


@mcp.resource(
    "excel://test-runs/{run_id}",
    name="test_run_summary",
    description=_DESC_RES_TEST_RUN,
    mime_type="application/json",
)
async def test_run_resource(run_id: str) -> str:
    """Read-only summary of one declarative test run (never re-runs anything)."""

    try:
        run = test_tools.get_test_runner().get_run(run_id)
    except ExcelVbaMcpError as exc:
        raise ResourceError(json.dumps(exc.to_dict())) from exc
    return run.model_dump_json(indent=2)


# Static glossary (plan_03 §5.3, D-U8). The single authoritative term set; tool
# descriptions reference this resource by URI instead of re-explaining jargon.
_GLOSSARY_MD = """# xlvba glossary

- **R13** — heuristic interactive-UI token veto; a veto, never a certificate.
- **BG06** — background-safe classification bound to (workbook hash, source hash,
  procedure); invalidated on VBA source change.
- **preflight** — the R13 scan performed before dispatch.
- **truth signal** — Excel's own success evidence (the `RefreshTable` Boolean or an
  advanced `RefreshDate`); never "the absence of a COM exception".
- **fail-closed** — unknown/ambiguous states are non-success, never a fabricated
  success.
- **RECOVERY** — the auto-allowed, agent-directed recovery class (`excel_recover`).
- **STA** — single-threaded apartment: the one serialized Excel thread.
- **Stage 1** — the typed-DTO contract layer (`SCHEMA_VERSION=1`).
- **normalized source hash** — SHA-256 of the canonicalized component body; the
  `expected_source_hash` revision key.
- **SUSPECTED_STALLED** — a soft deadline was exceeded and the operation was left
  running; not a confirmed hang.
- **interrupted_unknown** — a job finalized by recovery; never replayed.
- **indeterminate** — the compile probe never returns compiled/compile_error.
- **working copy** — the opened, mutable copy; originals are never touched.
- **snapshot gate** — the in-memory `_latest_snapshot` checked by `require_snapshot`.
- **guarded mutation** — a write tool that requires a writable session, a snapshot,
  an approval gate, and a postcondition re-read before it reports success.
- **approval gate** — the policy check that auto-allows in full-access mode and
  requires an explicit grant in restricted mode; a denial is `APPROVAL_DENIED`.
"""


@mcp.resource(
    "xlvba://glossary",
    name="glossary",
    description=_DESC_RES_GLOSSARY,
    mime_type="text/markdown",
)
async def glossary_resource() -> str:
    """Read-only static glossary of server jargon (never touches Excel)."""

    return _GLOSSARY_MD


def main() -> None:
    """Run the MCP server over STDIO."""

    logger.info("Starting Excel VBA MCP server (STDIO)")
    mcp.run(transport="stdio")
