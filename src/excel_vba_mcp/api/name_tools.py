"""Defined-name (Named Range) MCP tools (task_028, spec §7.7 / §6.5).

Thin, async tool handlers for the defined-name surface: the ``name`` /
``name_read`` flat-signature dispatchers (WP4) over the per-verb handlers
``names_list``, ``name_get``, ``name_create``, ``name_update`` and
``name_delete``. They are deliberately **thin**: each validates its arguments,
delegates to :class:`~excel_vba_mcp.application.names_service.NamesService` and
returns a domain DTO. They **never touch COM** (ARCH-009): no ``pythoncom`` /
``win32com`` import lives in this module; all Excel access happens on the
single-STA worker inside the service.

A name is identified by ``scope`` (``"workbook"`` or ``"worksheet"``); a
worksheet-scoped name also requires ``sheet``. Mutating tools
(``name_create`` / ``name_update`` / ``name_delete``) only succeed when the
service guard chain passes: writable session, snapshot, approval
(deny-by-default).

Like the other tool modules, the active service is published here via
:func:`set_names_service` (called from the server lifespan) and read via
:func:`get_names_service`.
"""

from __future__ import annotations

from typing import Any, Literal

from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import DefinedNameInfo
from excel_vba_mcp.domain.statuses import NameScope

# The active names service, published by the server lifespan. ``None`` before
# startup and after shutdown; the tools raise a clear error if used out of band.
_names_service: NamesService | None = None


def set_names_service(service: NamesService | None) -> None:
    """Publish (or clear) the active names service used by the tools."""

    global _names_service
    _names_service = service


def get_names_service() -> NamesService:
    """Return the active names service or raise if it is not configured."""

    if _names_service is None:
        raise ExcelVbaMcpError(
            "The names service is not initialized.",
            context={"tool": "name_tools"},
        )
    return _names_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


def _require_scope(value: Any) -> NameScope:
    """Validate ``scope`` as ``"workbook"`` / ``"worksheet"``."""

    if isinstance(value, NameScope):
        return value
    if isinstance(value, str):
        try:
            return NameScope(value.strip().lower())
        except ValueError:
            pass
    raise ExcelVbaMcpError(
        "Argument 'scope' must be 'workbook' or 'worksheet'.",
        context={"field": "scope", "value": value},
    )


def _validate_scope_sheet(scope: NameScope, sheet: str | None) -> str | None:
    """Require ``sheet`` for a worksheet-scoped name; reject it otherwise."""

    if scope is NameScope.WORKBOOK:
        if sheet is not None:
            raise ExcelVbaMcpError(
                "Argument 'sheet' must be omitted for a workbook-scoped name.",
                context={"field": "sheet", "scope": scope.value},
            )
        return None
    if not isinstance(sheet, str) or not sheet.strip():
        raise ExcelVbaMcpError(
            "Argument 'sheet' is required for a worksheet-scoped name.",
            context={"field": "sheet", "scope": scope.value},
        )
    return sheet


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------
async def names_list(workbook_id: str) -> list[DefinedNameInfo]:
    """List every defined name of an open workbook (all scopes).

    Args:
        workbook_id: The id returned by :func:`excel_open`.

    Returns:
        One :class:`DefinedNameInfo` per defined name with scope, ``NameLocal``,
        ``Visible``, ``RefersTo`` / ``RefersToLocal`` and a ``target_kind`` of
        ``range`` (with sheet + address), ``constant``, ``formula``,
        ``external``, ``error`` or ``unknown``.
    """

    _require_non_empty(workbook_id, "workbook_id")
    service = get_names_service()
    return await service.list_names(workbook_id)


async def name_get(
    workbook_id: str,
    name: str,
    scope: str,
    sheet: str | None = None,
) -> DefinedNameInfo:
    """Return one defined name by name + explicit scope.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The defined-name identifier.
        scope: ``"workbook"`` or ``"worksheet"``.
        sheet: The owning worksheet (required for ``scope="worksheet"``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    resolved_scope = _require_scope(scope)
    resolved_sheet = _validate_scope_sheet(resolved_scope, sheet)
    service = get_names_service()
    return await service.get_name(workbook_id, name, resolved_scope, resolved_sheet)


# ---------------------------------------------------------------------------
# Mutating tools (guarded)
# ---------------------------------------------------------------------------
async def name_create(
    workbook_id: str,
    name: str,
    refers_to: str,
    scope: str,
    sheet: str | None = None,
) -> DefinedNameInfo:
    """Create a defined name in an explicit scope.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). A duplicate in the requested scope or a cross-scope
    name collision is rejected with a structured error.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The defined-name identifier.
        refers_to: The formula / reference (``=Sheet1!$A$1``, ``=42``, …).
        scope: ``"workbook"`` or ``"worksheet"``.
        sheet: The owning worksheet (required for ``scope="worksheet"``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(refers_to, "refers_to")
    resolved_scope = _require_scope(scope)
    resolved_sheet = _validate_scope_sheet(resolved_scope, sheet)
    service = get_names_service()
    return await service.create_name(
        workbook_id, name, refers_to, scope=resolved_scope, sheet=resolved_sheet
    )


async def name_update(
    workbook_id: str,
    name: str,
    refers_to: str,
    scope: str,
    sheet: str | None = None,
) -> DefinedNameInfo:
    """Update a name's ``RefersTo`` in an explicit scope (scope is preserved).

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default).

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The defined-name identifier.
        refers_to: The new formula / reference.
        scope: ``"workbook"`` or ``"worksheet"``.
        sheet: The owning worksheet (required for ``scope="worksheet"``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    _require_non_empty(refers_to, "refers_to")
    resolved_scope = _require_scope(scope)
    resolved_sheet = _validate_scope_sheet(resolved_scope, sheet)
    service = get_names_service()
    return await service.update_name(
        workbook_id, name, refers_to, scope=resolved_scope, sheet=resolved_sheet
    )


async def name_delete(
    workbook_id: str,
    name: str,
    scope: str,
    sheet: str | None = None,
) -> DefinedNameInfo:
    """Delete a name in an explicit scope.

    Guarded mutation: requires a writable session, a snapshot and an approval
    grant (deny-by-default). Returns the pre-delete :class:`DefinedNameInfo`.

    Args:
        workbook_id: The id returned by :func:`excel_open`.
        name: The defined-name identifier.
        scope: ``"workbook"`` or ``"worksheet"``.
        sheet: The owning worksheet (required for ``scope="worksheet"``).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(name, "name")
    resolved_scope = _require_scope(scope)
    resolved_sheet = _validate_scope_sheet(resolved_scope, sheet)
    service = get_names_service()
    return await service.delete_name(workbook_id, name, scope=resolved_scope, sheet=resolved_sheet)


# ---------------------------------------------------------------------------
# Flat-signature dispatchers (WP4, option (a))
# ---------------------------------------------------------------------------
async def name(
    action: Literal["create", "update"],
    workbook_id: str | None = None,
    name: str | None = None,
    refers_to: str | None = None,
    scope: str | None = None,
    sheet: str | None = None,
) -> DefinedNameInfo:
    """Defined-name mutation dispatch: ``create`` or ``update`` (WP4).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`name_create` / :func:`name_update`
    handlers unchanged and forwards only the validated arguments. Per-action
    required parameters are enforced by the target handler (``create`` /
    ``update`` both need ``name``, ``refers_to`` and ``scope``; a
    worksheet-scoped name also needs ``sheet``) so a missing required argument
    surfaces as a structured :class:`ExcelVbaMcpError` rather than a crash. An
    unknown ``action`` raises a structured :class:`ExcelVbaMcpError`.
    ``delete`` is intentionally **not** an action here — use the standalone
    :func:`name_delete`.

    Args:
        action: ``create`` (new defined name) or ``update`` (repoint ``RefersTo``).
        workbook_id: The id returned by ``open``; required for every action.
        name: The defined-name identifier; required for every action.
        refers_to: The formula / reference (``=Sheet1!$A$1``, ``=42``, …);
            required for every action.
        scope: ``"workbook"`` or ``"worksheet"``; required for every action.
        sheet: The owning worksheet; required for a worksheet-scoped name.

    Returns:
        The affected :class:`DefinedNameInfo`.
    """

    if action == "create":
        return await name_create(
            workbook_id or "", name or "", refers_to or "", scope or "", sheet
        )
    if action == "update":
        return await name_update(
            workbook_id or "", name or "", refers_to or "", scope or "", sheet
        )
    raise ExcelVbaMcpError(
        f"Unknown name action {action!r}; expected 'create' or 'update'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "defined_name", "action": action},
    )


async def name_read(
    action: Literal["list", "get"],
    workbook_id: str | None = None,
    name: str | None = None,
    scope: str | None = None,
    sheet: str | None = None,
) -> list[DefinedNameInfo] | DefinedNameInfo:
    """Defined-name read dispatch: ``list`` or ``get`` (WP4).

    Thin flat-signature dispatcher (option (a)): routes to the existing
    :func:`names_list` / :func:`name_get` handlers unchanged. ``list`` requires
    only ``workbook_id``; ``get`` also requires ``name`` and ``scope`` (and
    ``sheet`` for a worksheet scope), enforced by the target handler. An unknown
    ``action`` raises a structured :class:`ExcelVbaMcpError`.

    Args:
        action: ``list`` (every defined name) or ``get`` (one by name + scope).
        workbook_id: The id returned by ``open``; required for both actions.
        name: The defined-name identifier; required for ``get``.
        scope: ``"workbook"`` or ``"worksheet"``; required for ``get``.
        sheet: The owning worksheet; required for a worksheet-scoped ``get``.

    Returns:
        A list of :class:`DefinedNameInfo` (``list``) or a single
        :class:`DefinedNameInfo` (``get``).
    """

    if action == "list":
        return await names_list(workbook_id or "")
    if action == "get":
        return await name_get(workbook_id or "", name or "", scope or "", sheet)
    raise ExcelVbaMcpError(
        f"Unknown name_read action {action!r}; expected 'list' or 'get'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "defined_name_read", "action": action},
    )


__all__ = [
    "get_names_service",
    "name",
    "name_create",
    "name_delete",
    "name_get",
    "name_read",
    "name_update",
    "names_list",
    "set_names_service",
]
