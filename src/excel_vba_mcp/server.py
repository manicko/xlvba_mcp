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
import logging
import os
import sys
import tempfile
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from pathlib import Path

from mcp.server.mcpserver import MCPServer

from excel_vba_mcp import __version__
from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.api import (
    compile_tools,
    macro_tools,
    sheet_tools,
    table_tools,
    test_tools,
    vba_tools,
    workbook_tools,
)
from excel_vba_mcp.api.compile_tools import vba_compile, vba_get_compile_diagnostics
from excel_vba_mcp.api.macro_tools import macro_get_last_result, macro_list, macro_run
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
)
from excel_vba_mcp.application.compile_service import CompileService
from excel_vba_mcp.application.macro_service import BackgroundSafetyRegistry, MacroService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.snapshot_service import SnapshotService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.security.approval import ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy
from excel_vba_mcp.storage.snapshot_store import SnapshotStore

logger = logging.getLogger("excel_vba_mcp.server")

# Environment variable names for the runtime configuration.
_ENV_ROOTS = "EXCEL_VBA_MCP_ROOTS"
_ENV_AUDIT = "EXCEL_VBA_MCP_AUDIT_LOG"
_ENV_WORK_DIR = "EXCEL_VBA_MCP_WORK_DIR"
_ENV_SNAPSHOT_DIR = "EXCEL_VBA_MCP_SNAPSHOT_DIR"

# Tool descriptions surfaced over MCP (kept explicit for a stable contract).
_DESC_OPEN = (
    "Open an allowlisted workbook as a hidden working copy. mode='inspect' "
    "disables macros; mode='execute' applies macro-by-UI and requires approval. "
    "The original file is never modified."
)
_DESC_CLOSE = (
    "Close an open workbook. Defaults to NOT saving; set save_changes=true to persist changes."
)
_DESC_SAVE = "Explicitly save an open workbook's working copy."
_DESC_STATUS = "Return session/workbook status (busy, read-only, dirty, active operation)."
_DESC_LIST = "List currently open workbooks from session bookkeeping."
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
    "mutation: requires a writable session, a successful snapshot, a matching "
    "expected_source_hash (revision check) and an approval grant. Document "
    "modules and UserForms are rejected."
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
    "macro runs only when it is approved (deny-by-default), classified "
    "background_safe for the current workbook hash (BG06) and passes the "
    "interactive-UI preflight (R13); otherwise a truthful blocked result is "
    "returned and the macro is NOT dispatched. Never dismisses a dialog or reports "
    "a waiting macro as passed."
)
_DESC_MACRO_GET_LAST_RESULT = (
    "Return the cached last macro run result (status, duration, return value, "
    "structured error) without re-running the macro."
)
_DESC_SHEETS_LIST = (
    "List every sheet of an open workbook (all types: Worksheet / Chart / Macro "
    "/ Dialog) with name, index, object_type, visibility and — for worksheets — "
    "code_name (read-only), protected, table_count and used_range."
)
_DESC_SHEET_GET = "Return one sheet's SheetInfo by name (Sheets vs Worksheets aware)."
_DESC_SHEET_ADD = (
    "Add a worksheet (MVP: Worksheet only). Guarded mutation: requires a writable "
    "session, a successful snapshot and an approval grant (deny-by-default)."
)
_DESC_SHEET_RENAME = (
    "Rename a sheet; CodeName is never changed. Guarded mutation: requires a "
    "writable session, a snapshot and an approval grant."
)
_DESC_SHEET_COPY = (
    "Copy a sheet and return the new sheet's DTO. Guarded mutation: requires a "
    "writable session, a snapshot and an approval grant."
)
_DESC_SHEET_DELETE = (
    "Delete a sheet (destructive). Requires a writable session, a snapshot and an "
    "approval grant; rejected when it would leave no visible worksheet."
)
_DESC_SHEET_SET_VISIBILITY = (
    "Set a sheet's visibility (xlSheetVisible / xlSheetHidden). Guarded mutation: "
    "requires a writable session, a snapshot and an approval grant; hiding the "
    "last visible worksheet is rejected."
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
    "mutation: requires a writable session, a snapshot and an approval grant."
)
_DESC_TABLE_INSERT_ROW = (
    "Insert one data row at a 1-based position via ListRows.Add. Guarded mutation: "
    "requires a writable session, a snapshot and an approval grant."
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
    compile: CompileService
    macro: MacroService
    path_policy: PathPolicy
    approval: ApprovalPolicy
    audit: AuditLogger


def _configured_roots() -> dict[str, str]:
    """Return PathPolicy roots from ``EXCEL_VBA_MCP_ROOTS`` or a safe default.

    ``EXCEL_VBA_MCP_ROOTS`` is an ``os.pathsep``-separated list of directories.
    When unset, the current working directory and the system temp directory are
    used so the server is usable out of the box; production deployments should
    set an explicit allowlist.
    """

    raw = os.environ.get(_ENV_ROOTS, "").strip()
    if raw:
        dirs = [part.strip() for part in raw.split(os.pathsep) if part.strip()]
    else:
        dirs = [str(Path.cwd()), tempfile.gettempdir()]
    return {f"root_{idx}": d for idx, d in enumerate(dirs)}


def _resolve_runtime_path(env_name: str, default: Path) -> Path:
    raw = os.environ.get(env_name, "").strip()
    return Path(raw) if raw else default


def _build_runtime() -> _Runtime:
    """Construct the Excel runtime from environment configuration."""

    path_policy = PathPolicy(_configured_roots())
    app_adapter = ExcelAppAdapter()
    worker = ExcelComWorker(app_adapter)
    approval = ApprovalPolicy()  # deny-by-default; grants/allowlist: task_014
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
    compile_service = CompileService(worker, session, vba, path_policy, audit)
    # background_safe starts empty (deny-by-default, BG06): no procedure is
    # eligible for background execution until it is explicitly classified for the
    # current workbook hash and verified on a disposable copy.
    macro_service = MacroService(worker, session, approval, BackgroundSafetyRegistry(), audit)
    return _Runtime(
        worker=worker,
        app_adapter=app_adapter,
        session=session,
        snapshot=snapshot,
        vba=vba,
        sheets=sheets,
        tables=tables,
        compile=compile_service,
        macro=macro_service,
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

    runtime = _build_runtime()
    try:
        await runtime.worker.start()
    except Exception as exc:  # noqa: BLE001 - degrade gracefully, never crash stdio
        logger.warning("Excel worker failed to start: %s", exc)
        await _safe_shutdown(runtime.worker)
    else:
        workbook_tools.set_session_service(runtime.session)
        test_tools.set_snapshot_service(runtime.snapshot)
        vba_tools.set_vba_project_service(runtime.vba)
        sheet_tools.set_sheets_service(runtime.sheets)
        table_tools.set_tables_service(runtime.tables)
        compile_tools.set_compile_service(runtime.compile)
        macro_tools.set_macro_service(runtime.macro)
    try:
        yield runtime
    finally:
        workbook_tools.set_session_service(None)
        test_tools.set_snapshot_service(None)
        vba_tools.set_vba_project_service(None)
        sheet_tools.set_sheets_service(None)
        table_tools.set_tables_service(None)
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
    server.add_tool(snapshot_create, name="snapshot_create", description=_DESC_SNAPSHOT_CREATE)
    server.add_tool(snapshot_list, name="snapshot_list", description=_DESC_SNAPSHOT_LIST)
    server.add_tool(snapshot_restore, name="snapshot_restore", description=_DESC_SNAPSHOT_RESTORE)
    server.add_tool(snapshot_diff, name="snapshot_diff", description=_DESC_SNAPSHOT_DIFF)
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


_configure_logging()

mcp: MCPServer = MCPServer("excel-vba-mcp", version=__version__, lifespan=lifespan)


@mcp.tool()
def server_health() -> dict[str, str]:
    """Read-only liveness probe (never touches Excel/COM)."""

    return {"status": "ok", "service": "excel-vba-mcp", "version": __version__}


# Register the base workbook/session tools at import time so they are present
# regardless of whether the server is started via ``main()`` here or directly via
# ``excel_vba_mcp.__main__`` (which calls ``mcp.run`` on this module's ``mcp``).
_register_tools(mcp)


def main() -> None:
    """Run the MCP server over STDIO."""

    logger.info("Starting Excel VBA MCP server (STDIO)")
    mcp.run(transport="stdio")
