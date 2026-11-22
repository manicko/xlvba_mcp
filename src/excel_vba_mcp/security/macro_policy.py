"""Macro trust and operation-classification policy (spec §7.1-§7.3, §8, §10, §11).

This module groups two decision areas (plus small pure diagnostics):

1. **Operation classification** (:func:`classify_operation`) — an **illustrative,
   non-authoritative** table that maps a canonical tool name (plus its parameters)
   to an :class:`OperationClass`. It is **not** wired into the runtime enforcement
   path: each service sets its own operation class when it submits work to the
   worker, so the live classification is defined per-service (see
   `docs/05-mcp-tools-reference.md` and the service code), not by this function.
   The table is retained for documentation/teaching only and is **stale** — it
   predates the current tool catalogue. Do not rely on it for enforcement.
2. **Macro trust diagnostics** (:class:`MacroPolicy`) — decides the automation
   security mode for a session (inspect vs execute) and turns raw COM/state
   signals (VBOM availability, project protection, ``Application.Run`` HRESULT)
   into typed domain errors / :class:`MacroRunStatus` values.

Design constraints (spec §7.2, §10, §10.1):

* This module performs **no COM access, no registry access and no mutation**. It
  never reads or writes Trust Center / VBA object-model trust settings; it only
  *diagnoses* their state.
* It never sets the application automation security property. Choosing the value
  is a pure decision here; applying and restoring it around ``Workbooks.Open`` is
  the responsibility of the session service (task_012).
* Diagnostics are pure functions returning typed errors or ``None`` so they can
  be unit-tested without Excel.

Unknown-tool policy (illustrative table only)
---------------------------------------------
:func:`classify_operation` **fails closed**: an unrecognized tool name is
classified as :attr:`OperationClass.EXECUTING` (the most restrictive class). This
mirrors the deny-by-default stance of the real approval gate, but — because
``classify_operation`` is **not called at runtime** — it does not itself enforce
anything; each service sets its own operation class. A caller that wants to catch
programming errors should validate tool names against the API schema before
dispatch.
"""

from __future__ import annotations

from collections.abc import Mapping
from enum import IntEnum
from typing import Any

from excel_vba_mcp.domain.errors import AccessVbomDisabledError, VbaProjectLockedError
from excel_vba_mcp.domain.statuses import MacroRunStatus, OperationClass, SessionMode


class AutomationSecurityMode(IntEnum):
    """Excel ``MsoAutomationSecurity`` values (spec §10, §10.1).

    Mirrors the Office enumeration so the session service can apply the value
    directly. ``LOW`` exists for completeness but is never selected by
    :meth:`MacroPolicy.automation_security_for` — enabling all macros by default is
    forbidden by spec §10.
    """

    LOW = 1
    BY_UI = 2
    FORCE_DISABLE = 3


# vbext_pp_locked from the VBIDE ``vbext_ProjectProtection`` enum.
_VBEXT_PP_LOCKED = 1

# HRESULT values that carry macro-execution semantics (see domain/errors.py).
_HRESULT_MACRO_NOT_FOUND = 0x800A03ED
_HRESULT_MACRO_RUNTIME_ERROR = 0x800A03EC
_HRESULT_E_FAIL = 0x80004005

# ---------------------------------------------------------------------------
# Operation classification table
# ---------------------------------------------------------------------------
#
# Keys are canonical tool names (the union of the plan's API function names and
# the descriptive names used in the research/design). ``excel_close`` is handled
# separately because its class depends on the ``save_changes`` flag.

_READ_ONLY_TOOLS: frozenset[str] = frozenset(
    {
        # Workbook / session
        "excel_open",
        "excel_get_status",
        "excel_list_workbooks",
        "workbook_get_model",
        # Snapshots (SaveCopyAs / file copy do not mutate the open workbook)
        "snapshot_create",
        "snapshot_list",
        "snapshot_diff",
        # VBA project / source reads and exports
        "vba_project_info",
        "vba_list_components",
        "vba_get_components",
        "vba_read_component",
        "vba_get_component_source",
        "vba_list_procedures",
        "vba_get_references",
        "vba_search_code",
        "vba_get_compile_diagnostics",
        "vba_export",
        "vba_export_project",
        "vba_export_component",
        # Macro inventory
        "macro_list",
        "macro_get_last_result",
        # Sheets / tables / names / controls inventory
        "sheets_list",
        "sheet_get",
        "tables_list",
        "table_get",
        "table_read_rows",
        "names_list",
        "name_get",
        "controls_list",
        "control_get",
        "shapes_list",
        "shape_get",
    }
)

_MUTATING_TOOLS: frozenset[str] = frozenset(
    {
        # Workbook / snapshots
        "excel_save",
        "snapshot_restore",
        # VBA source / components
        "vba_write",
        "vba_insert_code",
        "vba_replace_component_source",
        "vba_add_component",
        "vba_remove_component",
        # Sheets
        "sheet_add",
        "sheet_rename",
        "sheet_copy",
        "sheet_delete",
        "sheet_set_visibility",
        # Tables (ListObjects)
        "table_create",
        "table_add_column",
        "table_append_rows",
        "table_insert_row",
        "table_update_rows",
        "table_delete_rows",
        "table_delete_column",
        "table_resize",
        "table_row_add",
        "table_row_delete",
        "table_column_delete",
        # Defined names
        "name_create",
        "name_update",
        "name_delete",
        # Control / shape macro wiring (latent execution path, not execution)
        "control_set_on_action",
        "shape_set_on_action",
    }
)

_EXECUTING_TOOLS: frozenset[str] = frozenset(
    {
        "macro_run",
        "compile",
        "vba_compile",
        "test_run",
        "test_run_suite",
    }
)


def _truthy(value: Any) -> bool:
    """Interpret a parameter flag defensively (accept bools and common strings)."""

    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


def classify_operation(
    tool_name: str,
    params: Mapping[str, Any] | None = None,
) -> OperationClass:
    """Classify a tool invocation as ``READ_ONLY``, ``MUTATING`` or ``EXECUTING``.

    **Non-authoritative / illustrative only.** This function is **not** called at
    runtime and is **not** wired into enforcement; the live operation class of each
    tool is defined per-service (each service passes its own ``OperationClass`` to
    the worker). The table is also stale — it predates the current tool catalogue.
    Do not rely on it for enforcement; use ``docs/05-mcp-tools-reference.md`` or the
    service code for the authoritative classification.

    Args:
        tool_name: Canonical tool/action name (e.g. ``"vba_write"``).
        params: Tool parameters; only consulted for conditional operations
            (currently ``excel_close`` / ``save_changes``).

    Returns:
        The operation class. Unknown tool names return
        :attr:`OperationClass.EXECUTING` (fail-closed; see module docstring).
    """

    resolved: Mapping[str, Any] = params if params is not None else {}

    # Conditional: close is read-only unless it persists changes (spec §7.1 —
    # "close без save_changes не должен молча записывать").
    if tool_name == "excel_close":
        if _truthy(resolved.get("save_changes")):
            return OperationClass.MUTATING
        return OperationClass.READ_ONLY

    if tool_name in _READ_ONLY_TOOLS:
        return OperationClass.READ_ONLY
    if tool_name in _MUTATING_TOOLS:
        return OperationClass.MUTATING
    if tool_name in _EXECUTING_TOOLS:
        return OperationClass.EXECUTING

    # Fail closed: an unknown tool is treated as the most restrictive class so it
    # cannot bypass the approval gate as a read-only operation.
    return OperationClass.EXECUTING


def automation_security_for(session_mode: SessionMode) -> AutomationSecurityMode:
    """Return the automation security mode for a session mode (spec §10.1).

    * ``INSPECT`` → ``FORCE_DISABLE`` (macros never run while inspecting).
    * ``EXECUTE`` → ``BY_UI`` (trusted copy; user/macro policy decides).

    ``LOW`` is never returned: enabling all macros by default is forbidden.
    """

    if session_mode == SessionMode.INSPECT:
        return AutomationSecurityMode.FORCE_DISABLE
    if session_mode == SessionMode.EXECUTE:
        return AutomationSecurityMode.BY_UI
    raise ValueError(f"Unknown session mode: {session_mode!r}")


def diagnose_vbom(is_accessible: bool) -> AccessVbomDisabledError | None:
    """Return :class:`AccessVbomDisabledError` when VBOM access is unavailable.

    Pure diagnostic: the caller probes ``VBProject`` accessibility and passes the
    boolean result. This function never touches the registry or Trust Center
    (spec §7.2, §10).
    """

    if is_accessible:
        return None
    return AccessVbomDisabledError(
        "Access to the VBA project object model is disabled. "
        "Enable 'Trust access to the VBA project object model' in the Excel "
        "Trust Center manually, then restart Excel."
    )


def diagnose_protection(protection: int) -> VbaProjectLockedError | None:
    """Return :class:`VbaProjectLockedError` when the VBA project is locked.

    ``protection`` is a ``VBProject.Protection`` value; ``1`` is
    ``vbext_pp_locked``. A locked project must never be bypassed (spec §7.2).
    """

    if protection == _VBEXT_PP_LOCKED:
        return VbaProjectLockedError(
            "The VBA project is locked/password-protected. Unlock it manually "
            "in the VBE (the owner must supply the password); protection is "
            "never bypassed automatically."
        )
    return None


def classify_macro_failure(
    hresult: int | None,
    *,
    approved: bool,
    evidence: str | None = None,
) -> MacroRunStatus:
    """Map an ``Application.Run`` outcome to a :class:`MacroRunStatus`.

    The mapping is evidence-based (spec §8.3, §11): a bare "no exception" is never
    reported as success, and a generic failure is not mislabeled ``NOT_FOUND``.

    Args:
        hresult: Raw COM HRESULT from ``Application.Run``, or ``None`` if the call
            did not return one (soft timeout / recovery path).
        approved: Whether policy approved the run. ``False`` short-circuits to
            :attr:`MacroRunStatus.BLOCKED` regardless of HRESULT.
        evidence: Optional human-readable corroboration (e.g. a VBIDE procedure
            lookup result). Does not override the HRESULT-based mapping; retained
            for diagnostics/audit.

    Returns:
        The macro run status. ``None``/unknown HRESULT maps to
        :attr:`MacroRunStatus.NEEDS_RECOVERY` (never a false success).
    """

    if not approved:
        return MacroRunStatus.BLOCKED

    if hresult is None:
        return MacroRunStatus.NEEDS_RECOVERY

    code = hresult & 0xFFFFFFFF
    if code == _HRESULT_MACRO_NOT_FOUND:
        return MacroRunStatus.NOT_FOUND
    if code in (_HRESULT_MACRO_RUNTIME_ERROR, _HRESULT_E_FAIL):
        return MacroRunStatus.RUNTIME_ERROR

    return MacroRunStatus.NEEDS_RECOVERY


class MacroPolicy:
    """Namespace grouping the pure macro-trust decisions (spec §7.2, §7.3, §10).

    All members are pure and stateless; the class exists so callers can depend on
    a single ``MacroPolicy`` object without importing free functions. It performs
    no COM access and mutates no Office/Trust Center state.
    """

    classify_operation = staticmethod(classify_operation)
    automation_security_for = staticmethod(automation_security_for)
    diagnose_vbom = staticmethod(diagnose_vbom)
    diagnose_protection = staticmethod(diagnose_protection)
    classify_macro_failure = staticmethod(classify_macro_failure)
