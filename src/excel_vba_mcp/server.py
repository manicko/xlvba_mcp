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
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from mcp.server.mcpserver import Context, MCPServer
from mcp.server.mcpserver.exceptions import ResourceError, ToolError
from mcp.types import CallToolResult, InputRequiredResult
from pydantic import ValidationError

from excel_vba_mcp import __version__
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import (
    compile_tools,
    control_tools,
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
from excel_vba_mcp.api.compile_tools import vba_compile, vba_get_compile_diagnostics
from excel_vba_mcp.api.control_tools import (
    control_get,
    control_set_on_action,
    controls_list,
)
from excel_vba_mcp.api.health_tools import excel_health
from excel_vba_mcp.api.macro_tools import (
    macro_classify_background,
    macro_get_last_result,
    macro_list,
    macro_run,
)
from excel_vba_mcp.api.name_tools import (
    name_create,
    name_delete,
    name_get,
    name_update,
    names_list,
)
from excel_vba_mcp.api.pivot_tools import (
    pivot_cache_get,
    pivot_cache_refresh,
    pivot_field_apply_filter,
    pivot_field_clear_filters,
    pivot_field_get_filters,
    pivot_field_set_aggregation,
    pivot_field_set_layout,
    pivot_field_set_position,
    pivot_field_set_properties,
    pivot_fields_list,
    pivot_table_get,
    pivot_table_refresh,
    pivot_tables_list,
)
from excel_vba_mcp.api.range_tools import (
    range_copy,
    range_delete,
    range_get,
    range_get_conditional_formatting,
    range_get_format,
    range_get_validation,
    range_insert,
    range_merge,
    range_read_formulas,
    range_read_values,
    range_set_conditional_formatting,
    range_set_format,
    range_set_validation,
    range_unmerge,
    range_write_formulas,
    range_write_values,
)
from excel_vba_mcp.api.shape_tools import (
    shape_get,
    shape_set_on_action,
    shapes_list,
)
from excel_vba_mcp.api.sheet_tools import (
    sheet_add,
    sheet_copy,
    sheet_delete,
    sheet_get,
    sheet_rename,
    sheet_set_visibility,
    sheets_list,
)
from excel_vba_mcp.api.table_tools import (
    table_add_column,
    table_append_rows,
    table_create,
    table_delete_column,
    table_delete_rows,
    table_get,
    table_insert_row,
    table_read_rows,
    table_resize,
    table_update_rows,
    tables_list,
)
from excel_vba_mcp.api.test_tools import (
    snapshot_create,
    snapshot_diff,
    snapshot_list,
    snapshot_restore,
    test_get_run,
    test_run_suite,
    test_suite_run,
)
from excel_vba_mcp.api.vba_tools import (
    vba_export_project,
    vba_get_references,
    vba_insert_code,
    vba_list_components,
    vba_list_procedures,
    vba_project_info,
    vba_read_component,
    vba_replace_component_source,
    vba_search_code,
)
from excel_vba_mcp.api.workbook_tools import (
    excel_close,
    excel_get_status,
    excel_list_workbooks,
    excel_open,
    excel_save,
    workbook_get_model,
)
from excel_vba_mcp.application.compile_service import CompileService
from excel_vba_mcp.application.controls_service import ControlsService
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

# Tool descriptions surfaced over MCP (kept explicit for a stable contract).
_DESC_OPEN = (
    "Open an allowlisted workbook as a hidden working copy. mode='inspect' "
    "disables macros; mode='execute' applies macro-by-UI (allowed by default; "
    "restricted mode requires approval). The original file is never modified."
)
_DESC_CLOSE = (
    "Close an open workbook. Defaults to NOT saving; set save_changes=true to persist changes."
)
_DESC_SAVE = "Explicitly save an open workbook's working copy."
_DESC_STATUS = "Return session/workbook status (busy, read-only, dirty, active operation)."
_DESC_LIST = "List currently open workbooks from session bookkeeping."
_DESC_HEALTH = (
    "Return worker/session health (state, accepting, owned Excel pid, active "
    "operation, needs_recovery) without forcing a re-run or touching COM. "
    "Read-only: use it to check whether the session needs recovery."
)
_DESC_SNAPSHOT_CREATE = (
    "Create an immutable snapshot of a workbook (SaveCopyAs for open books, file "
    "copy for closed) with manifest and SHA-256."
)
_DESC_SNAPSHOT_LIST = "List snapshots and their metadata."
_DESC_SNAPSHOT_RESTORE = (
    "Restore a working session from a snapshot: close the current working copy "
    "without saving, then open the snapshot baseline. Does NOT undo external "
    "macro side effects."
)
_DESC_SNAPSHOT_DIFF = "Compare two snapshots by manifest metadata and file hash."
_DESC_TEST_RUN_SUITE = (
    "Run a bounded declarative test suite on disposable copies of an open "
    "workbook. Actions and assertions come from closed allowlists, so no "
    "arbitrary Python, shell, COM or VBA can be expressed; an unknown action or "
    "assertion is rejected before any Excel call. Every case starts from its own "
    "disposable working copy of one immutable baseline snapshot (one case cannot "
    "contaminate another); a failed compile step blocks the case's macro steps. "
    "Returns per-case and per-step statuses; external macro side effects are NOT "
    "rolled back by the snapshot."
)
_DESC_TEST_GET_RUN = (
    "Return a previously persisted test run by run_id (statuses, per-case steps, "
    "assertions, artifact references). Read-only: it never touches Excel."
)
_DESC_TEST_SUITE_RUN = (
    "Validate and run a declarative test suite (task_053 explicit entry point). "
    "Accepts the suite as an object or a plain JSON-style mapping; the closed "
    "Range/Pivot action and assertion vocabularies are validated before any Excel "
    "call, so an unknown action/assertion or an arbitrary-code/unbounded payload "
    "is rejected up front. Behaves like test_run_suite: one immutable baseline "
    "snapshot, one disposable working copy per case, Range/Pivot steps dispatched "
    "to the existing services; external macro side effects are NOT rolled back."
)
_DESC_VBA_PROJECT_INFO = (
    "Return the VBA project summary: name, protection/lock state, component "
    "inventory and references (missing references are reported)."
)
_DESC_VBA_LIST_COMPONENTS = (
    "List VBA components with type (standard/class/user_form/document), line "
    "count, normalized source hash and designer/export flags."
)
_DESC_VBA_LIST_PROCEDURES = (
    "List the procedures of one VBA component via the VBIDE CodeModule API; "
    "heuristic fallbacks are flagged is_heuristic=true."
)
_DESC_VBA_READ_COMPONENT = (
    "Read the code module of one VBA component (code text + normalized source hash)."
)
_DESC_VBA_EXPORT_PROJECT = (
    "Export every VBA component (.bas/.cls/.frm+.frx) and return the export "
    "manifest with per-file hashes; Document modules are never removed."
)
_DESC_VBA_GET_REFERENCES = (
    "List VBA project references with Name/GUID/version/path; broken references "
    "are flagged is_broken=true."
)
_DESC_VBA_SEARCH_CODE = (
    "Search every VBA component's code for a substring or regular expression (bounded results)."
)
_DESC_VBA_REPLACE_COMPONENT_SOURCE = (
    "Replace the entire code module of a Standard/Class VBA component. Guarded "
    "mutation: requires a writable session, a successful snapshot and a matching "
    "expected_source_hash (revision check); allowed by default (full-access "
    "policy), with approval required only in restricted mode. Document modules "
    "and UserForms are rejected."
)
_DESC_VBA_INSERT_CODE = (
    "Insert VBA code at a semantic position in a Standard/Class component. "
    "position is 'declarations', 'before_procedure:<name>' or "
    "'after_procedure:<name>'. Same guard chain as replace."
)
_DESC_VBA_COMPILE = (
    "Return a fail-closed compile result for an open workbook. On the validated "
    "build non-activating compilation is unavailable, so this returns "
    "status='indeterminate' with reason='NON_ACTIVATING_COMPILE_UNAVAILABLE' and "
    "partial diagnostics (is_compile=false: broken references, protection, "
    "component inventory, VBOM availability). It never returns compiled or "
    "compile_error and never shows or activates UI."
)
_DESC_VBA_GET_COMPILE_DIAGNOSTICS = (
    "Return the last compile diagnostics (is_compile=false) without forcing a compile re-run."
)
_DESC_MACRO_LIST = (
    "List the callable macro candidates of an open workbook: each public "
    "Sub/Function of a standard/class module with its name, module, kind and "
    "background_safe classification (BG06)."
)
_DESC_MACRO_RUN = (
    "Run one approved, background-safe macro of an execute-mode workbook. The "
    "macro runs only when it is approved (allowed by default; restricted mode is "
    "opt-in), classified "
    "background_safe for the current workbook hash (BG06) and passes the "
    "interactive-UI preflight (R13); otherwise a truthful blocked result is "
    "returned and the macro is NOT dispatched. Never dismisses a dialog or reports "
    "a waiting macro as passed."
)
_DESC_MACRO_GET_LAST_RESULT = (
    "Return the cached last macro run result (status, duration, return value, "
    "structured error) without re-running the macro."
)
_DESC_MACRO_CLASSIFY_BACKGROUND = (
    "Classify a procedure's background_safe for the current workbook + VBA source "
    "revision. UNGATED bookkeeping (macro_run remains the security boundary): "
    "review a procedure (vba_read_component / vba_list_procedures) then assert "
    "background_safe here. The classification is bound to (workbook hash, source "
    "hash, procedure) and is invalidated on VBA source change; a heuristic or "
    "interactive-UI procedure cannot be classified background_safe (R13 veto)."
)
_DESC_SHEETS_LIST = (
    "List every sheet of an open workbook (all types: Worksheet / Chart / Macro "
    "/ Dialog) with name, index, object_type, visibility and — for worksheets — "
    "code_name (read-only), protected, table_count and used_range."
)
_DESC_SHEET_GET = "Return one sheet's SheetInfo by name (Sheets vs Worksheets aware)."
_DESC_SHEET_ADD = (
    "Add a worksheet (MVP: Worksheet only). Guarded mutation: requires a writable "
    "session and a snapshot; allowed by default (full-access policy), with approval "
    "required only in restricted mode."
)
_DESC_SHEET_RENAME = (
    "Rename a sheet; CodeName is never changed. Guarded mutation: requires a "
    "writable session and a snapshot; allowed by default (full-access policy), "
    "with approval required only in restricted mode."
)
_DESC_SHEET_COPY = (
    "Copy a sheet and return the new sheet's DTO. Guarded mutation: requires a "
    "writable session and a snapshot; allowed by default (full-access policy), "
    "with approval required only in restricted mode."
)
_DESC_SHEET_DELETE = (
    "Delete a sheet (destructive). Requires a writable session and a snapshot; "
    "allowed by default (full-access policy), with approval required only in "
    "restricted mode; rejected when it would leave no visible worksheet."
)
_DESC_SHEET_SET_VISIBILITY = (
    "Set a sheet's visibility (xlSheetVisible / xlSheetHidden). Guarded mutation: "
    "requires a writable session and a snapshot; allowed by default (full-access "
    "policy), with approval required only in restricted mode; hiding the last "
    "visible worksheet is rejected."
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
    "Append rows through ListRows.Add (table structure preserved). Guarded "
    "mutation: requires a writable session and a snapshot; allowed by default "
    "(full-access policy), with approval required only in restricted mode."
)
_DESC_TABLE_INSERT_ROW = (
    "Insert one data row at a 1-based position via ListRows.Add. Guarded mutation: "
    "requires a writable session and a snapshot; allowed by default (full-access "
    "policy), with approval required only in restricted mode."
)
_DESC_TABLE_UPDATE_ROWS = (
    "Update rows by explicit index or an unambiguous where filter. Ambiguous "
    "filters are rejected with a preview (no auto-guess). Guarded mutation: "
    "writable session + snapshot + approval."
)
_DESC_TABLE_DELETE_ROWS = (
    "Delete rows by explicit indices or an unambiguous where filter. Ambiguous "
    "filters are rejected with a preview. Guarded mutation: writable session + "
    "snapshot + approval."
)
_DESC_TABLE_ADD_COLUMN = (
    "Add a column via ListColumns.Add(Position) with a unique header. Guarded "
    "mutation: writable session + snapshot + approval."
)
_DESC_TABLE_DELETE_COLUMN = (
    "Delete the column identified by a unique header. Guarded mutation: writable "
    "session + snapshot + approval."
)
_DESC_TABLE_CREATE = (
    "Create a ListObject on a sheet over an A1 range or defined name. Guarded "
    "mutation: writable session + snapshot + approval."
)
_DESC_TABLE_RESIZE = (
    "Resize a table to a new A1 range via ListObject.Resize. Guarded mutation: "
    "writable session + snapshot + approval."
)
_DESC_RANGE_GET = (
    "Resolve an explicit worksheet + A1-style address to Range metadata "
    "(address, rows, cols, shape, merged/table flags). Read-only; multi-area or "
    "invalid addresses are rejected with a structured error."
)
_DESC_RANGE_READ_VALUES = (
    "Read typed cell values from a bounded Range window. Values are typed "
    "(empty/string/number/boolean/date/error) with no lossy string coercion; "
    "offset is a 0-based row index and limit caps the rows. The response is "
    "size-limited (truncated flag); oversized requests are rejected."
)
_DESC_RANGE_READ_FORMULAS = (
    "Read typed formulas from a bounded Range window (Formula2 primary, Formula "
    "fallback). _xlfn/_xlws/_xlpm prefixes are returned verbatim and flagged. "
    "offset is a 0-based row index and limit caps the rows; oversized requests "
    "are rejected."
)
_DESC_RANGE_WRITE_VALUES = (
    "Write typed cell values to a bounded Range. The payload shape is validated "
    "against the target before any mutation (a wrong-sized payload is rejected "
    "with no change). A string beginning with '=' is stored as text, not a "
    "formula. Guarded mutation: requires a writable working copy, a successful "
    "snapshot and a postcondition re-read; allowed by default (full-access "
    "policy), with approval required only in restricted mode. Cell writes inside "
    "an Excel Table are allowed only when the table row/column structure is "
    "unchanged."
)
_DESC_RANGE_WRITE_FORMULAS = (
    "Write formulas (Formula2 primary, Formula fallback) to a bounded Range. "
    "Every entry must be a formula string beginning with '='; a bare constant "
    "belongs in range_write_values. Guarded mutation: writable session + "
    "snapshot + approval + postcondition re-read. Cell writes inside an Excel "
    "Table are allowed only when the table row/column structure is unchanged."
)
_DESC_RANGE_GET_FORMAT = (
    "Read bounded, tri-state formatting for a Range: number_format, font "
    "(name/bold/italic/size/color), fill (color/pattern) and alignment "
    "(horizontal/vertical/wrap_text). Each property is a value, mixed (cells "
    "disagree; listed in mixed_properties) or unsupported (listed in "
    "unsupported); a mixed range is never flattened. Read-only."
)
_DESC_RANGE_SET_FORMAT = (
    "Apply a partial formatting update to a Range. Only the fields explicitly "
    "set in the patch are written; absent fields are left untouched. Values and "
    "enums are validated before any mutation; a null clears to automatic only "
    "where Excel allows. Guarded mutation: writable session + snapshot + "
    "approval + postcondition re-read."
)
_DESC_RANGE_COPY = (
    "Copy a Range to an explicit destination (same workbook, no clipboard). The "
    "destination is mandatory; an overlapping source/destination, a shape-"
    "mismatched destination or a destination intersecting merged cells is "
    "rejected. Structural copies that intersect an Excel Table (ListObject) are "
    "rejected and routed to the table tools. Guarded mutation: writable session "
    "+ snapshot + approval + postcondition re-read."
)
_DESC_RANGE_INSERT = (
    "Insert cells at an explicit address and shift them in an explicit direction "
    "('right' or 'down'). An omitted/unknown shift is rejected (never guessed). "
    "Rejected for a protected sheet, an Excel Table (ListObject) intersection, a "
    "dynamic-array spill range, or merged cells without allow_merged. Guarded "
    "mutation: writable session + snapshot + approval + postcondition re-read."
)
_DESC_RANGE_DELETE = (
    "Delete cells at an explicit address and shift the remainder in an explicit "
    "direction ('left' or 'up'). An omitted/unknown shift is rejected (never "
    "guessed). Rejected for a protected sheet, an Excel Table (ListObject) "
    "intersection, a dynamic-array spill range, or merged cells without "
    "allow_merged. Guarded mutation: writable session + snapshot + approval + "
    "postcondition re-read."
)
_DESC_RANGE_MERGE = (
    "Merge a Range into one cell (or one cell per row with across=true). Merging "
    "discards non-anchor data: the preflight reports the affected cells and the "
    "merge is blocked unless allow_data_loss=true is passed after review. "
    "Rejected for a protected sheet, an Excel Table (ListObject) intersection or "
    "a spill range. Guarded mutation: writable session + snapshot + approval."
)
_DESC_RANGE_UNMERGE = (
    "Unmerge every merged area intersecting an explicit target. Rejected when no "
    "merged area intersects the target, for a protected sheet, or for an Excel "
    "Table (ListObject) intersection. Guarded mutation: writable session + "
    "snapshot + approval + postcondition re-read."
)
_DESC_RANGE_GET_VALIDATION = (
    "Read the data-validation rule of a Range (type, Formula1/Formula2, operator, "
    "ignore_blank, in_cell_dropdown, alert_style and the message/title strings). "
    "A target without a rule returns present=false; a non-uniform target is "
    "reported with non_uniform_validation / mixed_validation in unsupported "
    "instead of being flattened. Read-only."
)
_DESC_RANGE_SET_VALIDATION = (
    "Add or modify the data-validation rule of a Range (list, whole number, "
    "decimal, date, time, text length, custom). Formula1/Formula2 are opaque "
    "strings bounded at 255 characters and are never evaluated; a missing operand, "
    "an over-long formula or an unsupported type is rejected before any mutation. "
    "An existing rule is modified so unspecified properties are preserved. Guarded "
    "mutation: writable session + snapshot + approval + postcondition re-read."
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
    "ColorScale / DataBar / IconSet are rejected explicitly. Guarded mutation: "
    "writable session + snapshot + approval + postcondition re-read."
)
_DESC_NAMES_LIST = (
    "List every defined name (Named Range) of an open workbook across workbook "
    "and worksheet scopes with scope, NameLocal, Visible, RefersTo/RefersToLocal "
    "and target_kind (range with sheet+address, constant, formula, external, "
    "error or unknown)."
)
_DESC_NAME_GET = (
    "Return one defined name by name and explicit scope ('workbook' or "
    "'worksheet'; worksheet scope also requires the owning sheet)."
)
_DESC_NAME_CREATE = (
    "Create a defined name in an explicit scope. Guarded mutation: requires a "
    "writable session and a snapshot; allowed by default (full-access policy), "
    "with approval required only in restricted mode; duplicate and cross-scope "
    "conflicts are rejected."
)
_DESC_NAME_UPDATE = (
    "Update a defined name's RefersTo in an explicit scope (scope is preserved). "
    "Guarded mutation: writable session + snapshot + approval."
)
_DESC_NAME_DELETE = (
    "Delete a defined name in an explicit scope. Guarded mutation: writable "
    "session + snapshot + approval. Returns the pre-delete name."
)
_DESC_SHAPES_LIST = (
    "List every drawing shape of an open workbook (all worksheets) with name, "
    "kind (AutoShape / TextBox / Group / Picture / Line / Form Control / ActiveX "
    "/ ...), on_action (the assigned macro, read verbatim and never executed), "
    "visibility, position/size and best-effort text."
)
_DESC_SHAPE_GET = "Return one shape's ShapeInfo by name (case-insensitive)."
_DESC_SHAPE_SET_ON_ACTION = (
    "Assign OnAction to a shape (the macro is NOT executed). Guarded mutation: "
    "requires a writable session and a snapshot; allowed by default (full-access "
    "policy), with approval required only in restricted mode. An empty "
    "macro_name clears the assignment."
)
_DESC_CONTROLS_LIST = (
    "List every Form Control and ActiveX control of an open workbook. Form "
    "Controls report control_type, caption, linked_cell and on_action; ActiveX "
    "controls report prog_id and an INFERRED/probable event procedure (never a "
    "confirmed mapping)."
)
_DESC_CONTROL_GET = "Return one control's ControlEntry by name (case-insensitive)."
_DESC_CONTROL_SET_ON_ACTION = (
    "Assign OnAction to a Form Control (the macro is NOT executed; ActiveX "
    "controls are rejected). Guarded mutation: requires a writable session and a "
    "snapshot; allowed by default (full-access policy), with approval required "
    "only in restricted mode."
)
_DESC_WORKBOOK_GET_MODEL = (
    "Return the aggregated Stage 1 workbook model: sheets, tables, names, shapes, "
    "controls and (optionally) VBA components/procedures, each with a section "
    "status. Read-only: no all-cells/Range data and no VBA execution. When VBA is "
    "unavailable, locked or VBOM trust is disabled the other sections are still "
    "returned and vba_access reports the state."
)
_DESC_PIVOT_TABLES_LIST = (
    "List every PivotTable and its cache across all worksheets of an open "
    "workbook. Enumeration is worksheet-scoped (a full-workbook sweep) because "
    "Workbook.PivotTables is not a usable inventory on the target build. Each "
    "PivotTable reports its name, sheet, cache index, TableRange1/TableRange2 "
    "addresses, source type and supported/unsupported status; shared caches are "
    "reported once with their affected tables. Read-only."
)
_DESC_PIVOT_TABLE_GET = (
    "Return one PivotTable's metadata by its stable Name: sheet, index, cache "
    "index, TableRange1/TableRange2/DataBodyRange addresses, row/column/page/"
    "hidden/data field layout, source type and supported/unsupported status. "
    "Read-only; never materializes the PivotTable cell output."
)
_DESC_PIVOT_CACHE_GET = (
    "Return one PivotCache's metadata by its 1-based CacheIndex: source type "
    "(database/external/consolidation/scenario/pivot_table/olap/data_model/"
    "unknown), SourceData, RefreshDate, RefreshOnFileOpen, EnableRefresh, OLAP "
    "indicator, shared-cache flag and the list of PivotTables that share this "
    "cache. Read-only."
)
_DESC_PIVOT_FIELDS_LIST = (
    "List a PivotTable's fields with source name, caption, orientation "
    "(hidden/row/column/page/data), 1-based position within the orientation "
    "(null when hidden), data-field aggregation function (from DataFields), "
    "ShowAllItems, manual filter state and data-field caption/number format. "
    "Read-only."
)
_DESC_PIVOT_TABLE_REFRESH = (
    "Refresh one PivotTable via PivotTable.RefreshTable (table-scoped, no "
    "Workbook.RefreshAll). The cache source type is checked first; external, "
    "OLAP and Data Model sources return unsupported_source and are not "
    "refreshed. The result reports every table sharing the cache "
    "(affected_pivot_tables), the RefreshDate before/after and a refreshed "
    "status confirmed by Excel's truth signal (the RefreshTable Boolean and/or "
    "an advanced RefreshDate), never by the absence of a COM exception. "
    "Guarded mutation: writable session + snapshot + approval + audit."
)
_DESC_PIVOT_CACHE_REFRESH = (
    "Refresh one PivotCache via PivotCache.Refresh (cache-scoped). A cache is "
    "shared N:1 with PivotTables, so affected_pivot_tables reports every table "
    "using the cache. Only a database/Table-backed cache is refreshed; "
    "external/OLAP/Data Model sources return unsupported_source. PivotCache."
    "Refresh returns no Boolean, so success is confirmed by the shared "
    "RefreshDate advancing. Never touches Workbook.RefreshAll or unrelated "
    "connections/queries. Guarded mutation: writable session + snapshot + "
    "approval + audit."
)
_DESC_PIVOT_FIELD_SET_LAYOUT = (
    "Set one PivotField's orientation bucket (hidden/row/column/page), or "
    "remove it. orientation='data' is rejected: assigning xlDataField does not "
    "create a data field, so use pivot_field_set_aggregation "
    "(PivotTable.AddDataField) for the data area. Removal is explicit: "
    "orientation='hidden' (default) is the non-destructive "
    "PivotField.Orientation=0 (the field and its items/filter state stay), "
    "while orientation='hidden' with delete=true is the destructive "
    "PivotField.Delete. External/OLAP/Data Model sources are rejected with "
    "unsupported_source. The returned PivotFieldInfo is the post-mutation state "
    "re-read from the PivotTable. Guarded mutation: writable session + snapshot "
    "+ approval + audit + postcondition."
)
_DESC_PIVOT_FIELD_SET_POSITION = (
    "Set one PivotField's 1-based position within its current orientation "
    "(row/column/page/data). Out-of-range positions and hidden (unplaced) "
    "fields are rejected before any COM mutation; the resulting orientation "
    "order is recorded in the audit entry. External/OLAP/Data Model sources are "
    "rejected with unsupported_source. The returned PivotFieldInfo is the "
    "post-mutation state re-read from the PivotTable. Guarded mutation: "
    "writable session + snapshot + approval + audit + postcondition."
)
_DESC_PIVOT_FIELD_SET_AGGREGATION = (
    "Add or update one data field's aggregation. An existing data field on the "
    "same source field is updated in place (DataFields[i].Function); a source "
    "field without a data field is added with PivotTable.AddDataField "
    "(assigning Orientation=xlDataField does not create one). Only the eleven "
    "accepted aggregation codes are allowed; legacy codes (9, -4114, -4153, "
    "-4154) are rejected. caption must be unique when supplied (duplicates are "
    "pre-validated here and rejected by Excel). External/OLAP/Data Model "
    "sources are rejected with unsupported_source. The returned PivotFieldInfo "
    "is the post-mutation state re-read from the PivotTable. Guarded mutation: "
    "writable session + snapshot + approval + audit + postcondition."
)
_DESC_PIVOT_FIELD_SET_PROPERTIES = (
    "Set bounded PivotField properties (caption / number_format / "
    "show_all_items). Only the properties explicitly set on the patch are "
    "written: caption and number_format target the data field (a source field "
    "without one is rejected) and show_all_items targets the source PivotField. "
    "An empty patch, a blank caption/number format or a duplicate caption is "
    "rejected before any COM mutation. External/OLAP/Data Model sources are "
    "rejected with unsupported_source. The returned PivotFieldInfo is the "
    "post-mutation state re-read from the PivotTable. Guarded mutation: "
    "writable session + snapshot + approval + audit + postcondition."
)

_DESC_PIVOT_FIELD_GET_FILTERS = (
    "Read one PivotField's current filter state: AllItemsVisible, the visible and "
    "hidden PivotItems (the manual PivotItems.Visible filter), the classified "
    "filter kind (none/manual) and any detected PivotFilters entries. Read-only; "
    "PivotFilters.Add is never called."
)
_DESC_PIVOT_FIELD_APPLY_FILTER = (
    "Apply a supported manual (visible-items) filter to one PivotField by setting "
    "PivotItems[i].Visible. Only the manual kind is supported: label, value, date "
    "and top/bottom filters (PivotFilters.Add) are unsupported on the target "
    "build, return an explicit unsupported result with filter.supported=false and "
    "a reason, and are never silently converted into a manual filter. The request "
    "(field identity, item names, duplicates, size) is validated before any "
    "mutation. External/OLAP/Data Model sources are rejected with "
    "unsupported_source. The returned PivotFieldFilters is the post-mutation state "
    "re-read from the PivotTable. Guarded mutation: writable session + snapshot + "
    "approval + audit + postcondition."
)
_DESC_PIVOT_FIELD_CLEAR_FILTERS = (
    "Clear PivotField filters with an explicit scope. scope='field' (default; "
    "requires field) calls PivotField.ClearAllFilters on that field; scope='table' "
    "calls PivotTable.ClearAllFilters for the whole PivotTable and returns a "
    "table-level result whose field is empty. External/OLAP/Data Model sources are "
    "rejected with unsupported_source. The result is confirmed by a post-mutation "
    "re-read. Guarded mutation: writable session + snapshot + approval + audit + "
    "postcondition."
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
    return roots


def _resolve_runtime_path(env_name: str, default: Path) -> Path:
    raw = os.environ.get(env_name, "").strip()
    return Path(raw) if raw else default


def _build_runtime(config: Config) -> _Runtime:
    """Construct the Excel runtime from environment + typed configuration."""

    path_policy = PathPolicy(_configured_roots())
    app_adapter = ExcelAppAdapter()
    worker = ExcelComWorker(app_adapter)
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
    macro_service = MacroService(worker, session, approval, BackgroundSafetyRegistry(), audit)
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
    runtime.audit.record(
        tool="server_startup",
        operation_class="config",
        params={"access_mode": config.server.access_mode.value},
        result={"version": __version__},
    )
    try:
        await runtime.worker.start()
    except Exception as exc:  # noqa: BLE001 - degrade gracefully, never crash stdio
        logger.warning("Excel worker failed to start: %s", exc)
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
        workbook_tools.set_workbook_model_service(runtime.model)
    try:
        yield runtime
    finally:
        workbook_tools.set_session_service(None)
        workbook_tools.set_workbook_model_service(None)
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
        await _safe_shutdown(runtime.worker)
        runtime.audit.close()


async def _safe_shutdown(worker: ExcelComWorker) -> None:
    """Shut the worker down without masking the original error."""

    try:
        await worker.shutdown()
    except Exception as exc:  # noqa: BLE001 - shutdown is best-effort
        logger.warning("Excel worker shutdown error: %s", exc)


def _configure_logging() -> None:
    """Send logs to ``stderr`` only; ``stdout`` is reserved for the protocol."""

    logging.basicConfig(
        level=logging.INFO,
        stream=sys.stderr,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )


def _register_tools(server: MCPServer) -> None:
    """Register the base workbook/session tools on ``server``."""

    server.add_tool(excel_open, name="excel_open", description=_DESC_OPEN)
    server.add_tool(excel_close, name="excel_close", description=_DESC_CLOSE)
    server.add_tool(excel_save, name="excel_save", description=_DESC_SAVE)
    server.add_tool(excel_get_status, name="excel_get_status", description=_DESC_STATUS)
    server.add_tool(excel_list_workbooks, name="excel_list_workbooks", description=_DESC_LIST)
    server.add_tool(excel_health, name="excel_health", description=_DESC_HEALTH)
    server.add_tool(snapshot_create, name="snapshot_create", description=_DESC_SNAPSHOT_CREATE)
    server.add_tool(snapshot_list, name="snapshot_list", description=_DESC_SNAPSHOT_LIST)
    server.add_tool(snapshot_restore, name="snapshot_restore", description=_DESC_SNAPSHOT_RESTORE)
    server.add_tool(snapshot_diff, name="snapshot_diff", description=_DESC_SNAPSHOT_DIFF)
    server.add_tool(test_run_suite, name="test_run_suite", description=_DESC_TEST_RUN_SUITE)
    server.add_tool(test_suite_run, name="test_suite_run", description=_DESC_TEST_SUITE_RUN)
    server.add_tool(test_get_run, name="test_get_run", description=_DESC_TEST_GET_RUN)
    server.add_tool(vba_project_info, name="vba_project_info", description=_DESC_VBA_PROJECT_INFO)
    server.add_tool(
        vba_list_components, name="vba_list_components", description=_DESC_VBA_LIST_COMPONENTS
    )
    server.add_tool(
        vba_list_procedures, name="vba_list_procedures", description=_DESC_VBA_LIST_PROCEDURES
    )
    server.add_tool(
        vba_read_component, name="vba_read_component", description=_DESC_VBA_READ_COMPONENT
    )
    server.add_tool(
        vba_export_project, name="vba_export_project", description=_DESC_VBA_EXPORT_PROJECT
    )
    server.add_tool(
        vba_get_references, name="vba_get_references", description=_DESC_VBA_GET_REFERENCES
    )
    server.add_tool(vba_search_code, name="vba_search_code", description=_DESC_VBA_SEARCH_CODE)
    server.add_tool(
        vba_replace_component_source,
        name="vba_replace_component_source",
        description=_DESC_VBA_REPLACE_COMPONENT_SOURCE,
    )
    server.add_tool(
        vba_insert_code,
        name="vba_insert_code",
        description=_DESC_VBA_INSERT_CODE,
    )
    server.add_tool(vba_compile, name="vba_compile", description=_DESC_VBA_COMPILE)
    server.add_tool(
        vba_get_compile_diagnostics,
        name="vba_get_compile_diagnostics",
        description=_DESC_VBA_GET_COMPILE_DIAGNOSTICS,
    )
    server.add_tool(macro_list, name="macro_list", description=_DESC_MACRO_LIST)
    server.add_tool(macro_run, name="macro_run", description=_DESC_MACRO_RUN)
    server.add_tool(
        macro_classify_background,
        name="macro_classify_background",
        description=_DESC_MACRO_CLASSIFY_BACKGROUND,
    )
    server.add_tool(
        macro_get_last_result,
        name="macro_get_last_result",
        description=_DESC_MACRO_GET_LAST_RESULT,
    )
    server.add_tool(sheets_list, name="sheets_list", description=_DESC_SHEETS_LIST)
    server.add_tool(sheet_get, name="sheet_get", description=_DESC_SHEET_GET)
    server.add_tool(sheet_add, name="sheet_add", description=_DESC_SHEET_ADD)
    server.add_tool(sheet_rename, name="sheet_rename", description=_DESC_SHEET_RENAME)
    server.add_tool(sheet_copy, name="sheet_copy", description=_DESC_SHEET_COPY)
    server.add_tool(sheet_delete, name="sheet_delete", description=_DESC_SHEET_DELETE)
    server.add_tool(
        sheet_set_visibility,
        name="sheet_set_visibility",
        description=_DESC_SHEET_SET_VISIBILITY,
    )
    server.add_tool(tables_list, name="tables_list", description=_DESC_TABLES_LIST)
    server.add_tool(table_get, name="table_get", description=_DESC_TABLE_GET)
    server.add_tool(table_read_rows, name="table_read_rows", description=_DESC_TABLE_READ_ROWS)
    server.add_tool(
        table_append_rows,
        name="table_append_rows",
        description=_DESC_TABLE_APPEND_ROWS,
    )
    server.add_tool(
        table_insert_row,
        name="table_insert_row",
        description=_DESC_TABLE_INSERT_ROW,
    )
    server.add_tool(
        table_update_rows,
        name="table_update_rows",
        description=_DESC_TABLE_UPDATE_ROWS,
    )
    server.add_tool(
        table_delete_rows,
        name="table_delete_rows",
        description=_DESC_TABLE_DELETE_ROWS,
    )
    server.add_tool(
        table_add_column,
        name="table_add_column",
        description=_DESC_TABLE_ADD_COLUMN,
    )
    server.add_tool(
        table_delete_column,
        name="table_delete_column",
        description=_DESC_TABLE_DELETE_COLUMN,
    )
    server.add_tool(table_create, name="table_create", description=_DESC_TABLE_CREATE)
    server.add_tool(table_resize, name="table_resize", description=_DESC_TABLE_RESIZE)
    server.add_tool(
        pivot_tables_list, name="pivot_tables_list", description=_DESC_PIVOT_TABLES_LIST
    )
    server.add_tool(pivot_table_get, name="pivot_table_get", description=_DESC_PIVOT_TABLE_GET)
    server.add_tool(pivot_cache_get, name="pivot_cache_get", description=_DESC_PIVOT_CACHE_GET)
    server.add_tool(
        pivot_fields_list, name="pivot_fields_list", description=_DESC_PIVOT_FIELDS_LIST
    )
    server.add_tool(
        pivot_table_refresh,
        name="pivot_table_refresh",
        description=_DESC_PIVOT_TABLE_REFRESH,
    )
    server.add_tool(
        pivot_cache_refresh,
        name="pivot_cache_refresh",
        description=_DESC_PIVOT_CACHE_REFRESH,
    )
    server.add_tool(
        pivot_field_set_layout,
        name="pivot_field_set_layout",
        description=_DESC_PIVOT_FIELD_SET_LAYOUT,
    )
    server.add_tool(
        pivot_field_set_position,
        name="pivot_field_set_position",
        description=_DESC_PIVOT_FIELD_SET_POSITION,
    )
    server.add_tool(
        pivot_field_set_aggregation,
        name="pivot_field_set_aggregation",
        description=_DESC_PIVOT_FIELD_SET_AGGREGATION,
    )
    server.add_tool(
        pivot_field_set_properties,
        name="pivot_field_set_properties",
        description=_DESC_PIVOT_FIELD_SET_PROPERTIES,
    )
    server.add_tool(
        pivot_field_get_filters,
        name="pivot_field_get_filters",
        description=_DESC_PIVOT_FIELD_GET_FILTERS,
    )
    server.add_tool(
        pivot_field_apply_filter,
        name="pivot_field_apply_filter",
        description=_DESC_PIVOT_FIELD_APPLY_FILTER,
    )
    server.add_tool(
        pivot_field_clear_filters,
        name="pivot_field_clear_filters",
        description=_DESC_PIVOT_FIELD_CLEAR_FILTERS,
    )
    server.add_tool(range_get, name="range_get", description=_DESC_RANGE_GET)
    server.add_tool(
        range_read_values, name="range_read_values", description=_DESC_RANGE_READ_VALUES
    )
    server.add_tool(
        range_read_formulas, name="range_read_formulas", description=_DESC_RANGE_READ_FORMULAS
    )
    server.add_tool(
        range_write_values, name="range_write_values", description=_DESC_RANGE_WRITE_VALUES
    )
    server.add_tool(
        range_write_formulas,
        name="range_write_formulas",
        description=_DESC_RANGE_WRITE_FORMULAS,
    )
    server.add_tool(range_get_format, name="range_get_format", description=_DESC_RANGE_GET_FORMAT)
    server.add_tool(range_set_format, name="range_set_format", description=_DESC_RANGE_SET_FORMAT)
    server.add_tool(range_copy, name="range_copy", description=_DESC_RANGE_COPY)
    server.add_tool(range_insert, name="range_insert", description=_DESC_RANGE_INSERT)
    server.add_tool(range_delete, name="range_delete", description=_DESC_RANGE_DELETE)
    server.add_tool(range_merge, name="range_merge", description=_DESC_RANGE_MERGE)
    server.add_tool(range_unmerge, name="range_unmerge", description=_DESC_RANGE_UNMERGE)
    server.add_tool(
        range_get_validation,
        name="range_get_validation",
        description=_DESC_RANGE_GET_VALIDATION,
    )
    server.add_tool(
        range_set_validation,
        name="range_set_validation",
        description=_DESC_RANGE_SET_VALIDATION,
    )
    server.add_tool(
        range_get_conditional_formatting,
        name="range_get_conditional_formatting",
        description=_DESC_RANGE_GET_CONDITIONAL_FORMATTING,
    )
    server.add_tool(
        range_set_conditional_formatting,
        name="range_set_conditional_formatting",
        description=_DESC_RANGE_SET_CONDITIONAL_FORMATTING,
    )
    server.add_tool(names_list, name="names_list", description=_DESC_NAMES_LIST)
    server.add_tool(name_get, name="name_get", description=_DESC_NAME_GET)
    server.add_tool(name_create, name="name_create", description=_DESC_NAME_CREATE)
    server.add_tool(name_update, name="name_update", description=_DESC_NAME_UPDATE)
    server.add_tool(name_delete, name="name_delete", description=_DESC_NAME_DELETE)
    server.add_tool(shapes_list, name="shapes_list", description=_DESC_SHAPES_LIST)
    server.add_tool(shape_get, name="shape_get", description=_DESC_SHAPE_GET)
    server.add_tool(
        shape_set_on_action,
        name="shape_set_on_action",
        description=_DESC_SHAPE_SET_ON_ACTION,
    )
    server.add_tool(controls_list, name="controls_list", description=_DESC_CONTROLS_LIST)
    server.add_tool(control_get, name="control_get", description=_DESC_CONTROL_GET)
    server.add_tool(
        control_set_on_action,
        name="control_set_on_action",
        description=_DESC_CONTROL_SET_ON_ACTION,
    )
    server.add_tool(
        workbook_get_model,
        name="workbook_get_model",
        description=_DESC_WORKBOOK_GET_MODEL,
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
    ``to_dict()`` so real MCP clients can observe the stable ``code`` and
    structured ``context``. When the cause is a pydantic ``ValidationError`` (the
    caller's arguments do not match the input schema) it is translated into an
    ``ExcelVbaMcpError`` (code ``INVALID_RANGE``, matching the tool modules'
    malformed-argument code) so **every** tool failure reaches the client
    structured (audit finding G-9 / F-5: no unstructured ``UnexpectedToolError``
    leak). Any other exception is re-raised unchanged (the SDK still reports it as
    an unexpected tool failure).
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
                    raise
            raise ToolError(json.dumps(cause.to_dict(), ensure_ascii=False)) from cause


_configure_logging()

mcp: ExcelVbaMcpServer = ExcelVbaMcpServer("excel-vba-mcp", version=__version__, lifespan=lifespan)


@mcp.tool()
def server_health() -> dict[str, str]:
    """Read-only liveness probe (never touches Excel/COM)."""

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
        raise ResourceError(str(exc)) from exc
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
        raise ResourceError(str(exc)) from exc
    return run.model_dump_json(indent=2)


def main() -> None:
    """Run the MCP server over STDIO."""

    logger.info("Starting Excel VBA MCP server (STDIO)")
    mcp.run(transport="stdio")
