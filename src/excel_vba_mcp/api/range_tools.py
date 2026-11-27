"""Range & Cells MCP tool contracts (task_040 schema group) and handlers.

This module defines the **signatures, Pydantic DTOs and docstrings** for the
Phase 2 Range/Cells surface and implements every tool as a thin async handler
(``task_042``–``task_046``): validate the arguments, delegate to
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and return a
domain DTO. This module deliberately imports **no Excel/COM** (``pythoncom`` /
``win32com``) — it only imports domain DTOs. All Excel access happens on the
single-STA worker inside the ranges service/adapter.

Tool groups:

* Read-only: ``range_get``, ``range_read_values``, ``range_read_formulas``,
  ``range_get_format``, ``range_get_validation``,
  ``range_get_conditional_formatting``.
* Mutating (snapshot + approval + audit + postcondition): ``range_write_values``,
  ``range_write_formulas``, ``range_set_format``, ``range_copy``,
  ``range_insert``, ``range_delete``, ``range_merge``, ``range_unmerge``,
  ``range_set_validation``, ``range_set_conditional_formatting``.

Structural operations that intersect an Excel Table (ListObject) are rejected
with :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError` and routed
to the table tools; the Range API never changes table row/column structure.
"""

from __future__ import annotations

from typing import Literal, cast

from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.models import (
    ConditionalFormatMutationResult,
    ConditionalFormatRule,
    ConditionalFormatSpec,
    FormatInfo,
    FormatMutationResult,
    FormulaWriteRequest,
    RangeAddress,
    RangeFormatPatch,
    RangeFormulas,
    RangeMutationResult,
    RangeValues,
    RangeWriteRequest,
    StructuralMutationResult,
    ValidationRule,
    ValidationSpec,
)

# ---------------------------------------------------------------------------
# Active service wiring (published by the server lifespan)
# ---------------------------------------------------------------------------
_ranges_service: RangesService | None = None


def set_ranges_service(service: RangesService | None) -> None:
    """Publish (or clear) the active ranges service used by the tools."""

    global _ranges_service
    _ranges_service = service


def get_ranges_service() -> RangesService:
    """Return the active ranges service or raise if it is not configured."""

    if _ranges_service is None:
        raise ExcelVbaMcpError(
            "The ranges service is not initialized.",
            context={"tool": "range_tools"},
        )
    return _ranges_service


def _require_non_empty(value: str, field: str) -> str:
    """Validate that a required string argument is non-empty."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a non-empty string.",
            context={"field": field},
        )
    return value


def _require_int(value: object, field: str, *, minimum: int | None = None) -> int:
    """Validate that an argument is an integer (optionally with a minimum)."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be an integer.",
            context={"field": field},
        )
    if minimum is not None and value < minimum:
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be >= {minimum}.",
            context={"field": field, "value": value},
        )
    return value


def _require_bool(value: object, field: str) -> bool:
    """Validate that an argument is a boolean flag (structural tools)."""

    if not isinstance(value, bool):
        raise ExcelVbaMcpError(
            f"Argument {field!r} must be a boolean.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "value": value},
        )
    return value


def _require_range_text(value: object, field: str) -> str:
    """Validate a sheet/address/destination argument of a structural tool."""

    if not isinstance(value, str) or not value.strip():
        raise ExcelVbaMcpError(
            f"Argument {field!r} is required and must be a non-empty string.",
            code=ErrorCode.INVALID_RANGE,
            context={"field": field, "value": value},
        )
    return value.strip()


#: Explicit shift directions accepted by ``range_insert`` (task_038 §5.2 item 2).
_INSERT_SHIFTS = ("right", "down")

#: Explicit shift directions accepted by ``range_delete`` (task_038 §5.2 item 2).
_DELETE_SHIFTS = ("left", "up")


def _require_shift(shift: object, allowed: tuple[str, ...], operation: str) -> str:
    """Validate an explicit shift name; an omitted/ambiguous shift is rejected."""

    if not isinstance(shift, str) or not shift.strip():
        raise ExcelVbaMcpError(
            f"Argument 'shift' is required for {operation} and must never be guessed.",
            code=ErrorCode.INVALID_RANGE,
            context={"shift": shift, "operation": operation},
        )
    normalized = shift.strip().lower()
    if normalized not in allowed:
        raise ExcelVbaMcpError(
            f"Argument 'shift' must be one of {list(allowed)} for {operation}.",
            code=ErrorCode.INVALID_RANGE,
            context={"shift": shift, "operation": operation, "allowed": list(allowed)},
        )
    return normalized


# ---------------------------------------------------------------------------
# Read-only tools
# ---------------------------------------------------------------------------


async def range_get(workbook_id: str, sheet: str, address: str) -> RangeAddress:
    """Resolve an explicit worksheet + A1 address to Range metadata.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name that owns the range.
        address: An A1-style address (e.g. ``"A1:C10"``); multi-area targets are
            rejected as invalid.

    Returns:
        A :class:`RangeAddress` with the resolved dimensions and shape.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.get_range(workbook_id, sheet, address)


async def range_read_values(
    workbook_id: str,
    sheet: str,
    address: str,
    offset: int = 0,
    limit: int | None = None,
) -> RangeValues:
    """Read typed cell values from a bounded Range window.

    Values are typed (``empty`` / ``string`` / ``number`` / ``boolean`` /
    ``date`` / ``error``) with no lossy string coercion. ``offset`` is a 0-based
    row index into the requested range; ``limit`` caps the returned rows. The
    response is bounded (``truncated`` marks a cut-short window).

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.
        offset: 0-based row index of the first row to return.
        limit: Optional cap on the number of rows returned.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    offset = _require_int(offset, "offset", minimum=0)
    if limit is not None:
        limit = _require_int(limit, "limit", minimum=1)
    service = get_ranges_service()
    return await service.read_values(workbook_id, sheet, address, offset=offset, limit=limit)


async def range_read_formulas(
    workbook_id: str,
    sheet: str,
    address: str,
    offset: int = 0,
    limit: int | None = None,
) -> RangeFormulas:
    """Read typed formulas from a bounded Range window (``Formula2`` primary).

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.
        offset: 0-based row index of the first row to return.
        limit: Optional cap on the number of rows returned.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    offset = _require_int(offset, "offset", minimum=0)
    if limit is not None:
        limit = _require_int(limit, "limit", minimum=1)
    service = get_ranges_service()
    return await service.read_formulas(workbook_id, sheet, address, offset=offset, limit=limit)


async def range_get_format(workbook_id: str, sheet: str, address: str) -> FormatInfo:
    """Read bounded, tri-state formatting for a Range.

    Each property is tri-state (``value`` / mixed / unsupported); mixed
    properties are listed in ``mixed_properties`` and unsupported ones in
    ``unsupported`` rather than being silently flattened.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.get_format(workbook_id, sheet, address)


async def range_get_validation(workbook_id: str, sheet: str, address: str) -> ValidationRule:
    """Read the data-validation rule of a Range (``present=false`` when none).

    A non-uniform target is not flattened to one cell's rule: the representative
    rule is returned with ``non_uniform_validation`` / ``mixed_validation`` in
    ``unsupported``.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.

    Read-only; no snapshot/approval.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.get_validation(workbook_id, sheet, address)


async def range_get_conditional_formatting(
    workbook_id: str, sheet: str, address: str
) -> list[ConditionalFormatRule]:
    """Read the conditional-formatting rules applying to a Range.

    Supported plain ``FormatCondition`` rules are returned with priority and
    ``stop_if_true``; ``ColorScale`` / ``DataBar`` / ``IconSet`` are flagged
    ``supported=false`` with an explicit ``unsupported_reason`` rather than omitted.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.

    Read-only; no snapshot/approval.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.get_conditional_formatting(workbook_id, sheet, address)


# ---------------------------------------------------------------------------
# Mutating tools (guarded: writable session + snapshot + approval + audit)
# ---------------------------------------------------------------------------


async def range_write_values(
    workbook_id: str,
    sheet: str,
    address: str,
    payload: RangeWriteRequest,
) -> RangeMutationResult:
    """Write typed cell values to a bounded Range.

    The payload shape is validated against the target **before** any COM
    mutation. A string beginning with ``=`` is stored as text, not a formula.
    Guarded mutation: requires a writable session, a snapshot, an approval grant
    and a postcondition re-read.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address (its dimensions must match payload).
        payload: The bounded 2-D value grid (:class:`RangeWriteRequest`).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.write_values(workbook_id, sheet, address, payload)


async def range_write_formulas(
    workbook_id: str,
    sheet: str,
    address: str,
    payload: FormulaWriteRequest,
) -> RangeMutationResult:
    """Write formulas (``Formula2``) to a bounded Range.

    Every entry must be an actual formula string beginning with ``=``; a bare
    constant belongs in :func:`range_write_values`. Guarded mutation: writable
    session + snapshot + approval + postcondition re-read.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address (its dimensions must match payload).
        payload: The bounded 2-D formula grid (:class:`FormulaWriteRequest`).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.write_formulas(workbook_id, sheet, address, payload)


async def range_set_format(
    workbook_id: str,
    sheet: str,
    address: str,
    patch: RangeFormatPatch,
) -> FormatMutationResult:
    """Apply a partial formatting update to a Range.

    Only fields explicitly set in ``patch`` are written; absent fields are left
    untouched. Guarded mutation: writable session + snapshot + approval +
    postcondition re-read. Properties outside the bounded set are rejected.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        patch: The sparse format update (:class:`RangeFormatPatch`).
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.set_format(workbook_id, sheet, address, patch)


async def range_copy(
    workbook_id: str,
    sheet: str,
    source: str,
    destination_sheet: str,
    destination: str,
) -> StructuralMutationResult:
    """Copy a Range to an explicit destination (no clipboard).

    ``destination_sheet`` and ``destination`` are mandatory; a missing
    destination is rejected. Structural copies that intersect a ListObject are
    rejected. Guarded mutation: writable session + snapshot + approval.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The source worksheet.
        source: The source A1-style address.
        destination_sheet: The destination worksheet (may differ from ``sheet``).
        destination: The destination A1-style address.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_range_text(source, "source")
    _require_range_text(destination_sheet, "destination_sheet")
    _require_range_text(destination, "destination")
    service = get_ranges_service()
    return await service.copy_range(workbook_id, sheet, source, destination_sheet, destination)


async def range_insert(
    workbook_id: str,
    sheet: str,
    address: str,
    shift: str,
    *,
    allow_merged: bool = False,
) -> StructuralMutationResult:
    """Insert cells and shift the range in an explicit direction.

    ``shift`` is mandatory (``"right"`` or ``"down"``); an omitted/ambiguous
    shift is rejected. Guarded mutation: writable session + snapshot + approval.
    Intersecting a ListObject is rejected.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        shift: ``"right"`` or ``"down"`` (required, never guessed).
        allow_merged: Permit inserting over merged cells after preview.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    shift_name = _require_shift(shift, _INSERT_SHIFTS, "insert")
    allow_merged = _require_bool(allow_merged, "allow_merged")
    service = get_ranges_service()
    return await service.insert_range(
        workbook_id, sheet, address, shift_name, allow_merged=allow_merged
    )


async def range_delete(
    workbook_id: str,
    sheet: str,
    address: str,
    shift: str,
    *,
    allow_merged: bool = False,
) -> StructuralMutationResult:
    """Delete cells and shift the range in an explicit direction.

    ``shift`` is mandatory (``"left"`` or ``"up"``). Guarded mutation: writable
    session + snapshot + approval. Intersecting a ListObject is rejected.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        shift: ``"left"`` or ``"up"`` (required, never guessed).
        allow_merged: Permit deleting over merged cells after preview.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    shift_name = _require_shift(shift, _DELETE_SHIFTS, "delete")
    allow_merged = _require_bool(allow_merged, "allow_merged")
    service = get_ranges_service()
    return await service.delete_range(
        workbook_id, sheet, address, shift_name, allow_merged=allow_merged
    )


async def range_merge(
    workbook_id: str,
    sheet: str,
    address: str,
    *,
    across: bool = False,
    allow_data_loss: bool = False,
) -> StructuralMutationResult:
    """Merge a Range into one cell.

    Non-anchor data is discarded on merge: the preflight reports affected cells
    and the operation is blocked unless ``allow_data_loss`` is set after preview.
    Guarded mutation: writable session + snapshot + approval.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        across: Merge each row separately (``Merge(Across=True)``).
        allow_data_loss: Permit merging when non-anchor cells hold data.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    across = _require_bool(across, "across")
    allow_data_loss = _require_bool(allow_data_loss, "allow_data_loss")
    service = get_ranges_service()
    return await service.merge_range(
        workbook_id, sheet, address, across=across, allow_data_loss=allow_data_loss
    )


async def range_unmerge(workbook_id: str, sheet: str, address: str) -> StructuralMutationResult:
    """Unmerge any merged area within a Range.

    Guarded mutation: writable session + snapshot + approval.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.unmerge_range(workbook_id, sheet, address)


async def range_set_validation(
    workbook_id: str,
    sheet: str,
    address: str,
    spec: ValidationSpec,
) -> ValidationRule:
    """Set (add/modify) the data-validation rule of a Range.

    Enforces the documented 255-character formula limit and per-type argument
    rules before any COM mutation. Guarded mutation: writable session + snapshot
    + approval + postcondition re-read.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        spec: The validation rule payload (:class:`ValidationSpec`).

    Guarded mutation: writable session + snapshot + approval + postcondition re-read.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.set_validation(workbook_id, sheet, address, spec)


async def range_set_conditional_formatting(
    workbook_id: str,
    sheet: str,
    address: str,
    spec: ConditionalFormatSpec,
    *,
    mode: str = "add",
) -> ConditionalFormatMutationResult:
    """Add / modify / delete a supported conditional-formatting rule.

    ``mode`` is ``"add"``, ``"modify"`` or ``"delete"``; ``spec.priority`` is the
    only stable rule identity and is required for ``modify`` / ``delete``. Rule
    order is preserved and priorities are sheet-wide: assigning a priority
    renumbers the other rules (reported through ``applied_priorities`` / ``notes``).
    Unsupported rule types (``xlColorScale`` / ``xlDataBar`` / ``xlIconSet``) are
    rejected explicitly; a rule type is never silently converted.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        spec: The conditional-formatting rule payload.
        mode: One of ``"add"``, ``"modify"``, ``"delete"``.

    Guarded mutation: writable session + snapshot + approval + postcondition re-read.
    """

    _require_non_empty(workbook_id, "workbook_id")
    _require_non_empty(sheet, "sheet")
    _require_non_empty(address, "address")
    service = get_ranges_service()
    return await service.set_conditional_formatting(workbook_id, sheet, address, spec, mode)


# ---------------------------------------------------------------------------
# Flat-signature dispatchers (WP7, option (a))
# ---------------------------------------------------------------------------
# The 16 range verbs are split into three cohesive dispatch tools plus two
# standalone destructive tools (``range_delete`` / ``range_merge``). Each
# dispatcher is a thin flat-signature router (option (a) — flat args + handler
# validation): it forwards only the validated arguments to the existing
# per-verb handlers unchanged (H5) and lets the target handler enforce its own
# required parameters, so a missing required argument surfaces as a structured
# :class:`ExcelVbaMcpError` rather than a crash. An unknown ``action`` raises a
# structured :class:`ExcelVbaMcpError`.


async def range_read(
    action: Literal[
        "get",
        "read-values",
        "read-formulas",
        "get-format",
        "get-validation",
        "get-conditional-formatting",
    ],
    workbook_id: str | None = None,
    sheet: str | None = None,
    address: str | None = None,
    offset: int = 0,
    limit: int | None = None,
) -> (
    RangeAddress
    | RangeValues
    | RangeFormulas
    | FormatInfo
    | ValidationRule
    | list[ConditionalFormatRule]
):
    """Range read dispatch: ``get``, ``read-values``, ``read-formulas``,
    ``get-format``, ``get-validation`` or ``get-conditional-formatting`` (WP7).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`range_get` /
    :func:`range_read_values` / :func:`range_read_formulas` /
    :func:`range_get_format` / :func:`range_get_validation` /
    :func:`range_get_conditional_formatting` handlers unchanged and forwards
    only the validated arguments. ``workbook_id``, ``sheet`` and ``address`` are
    required for every action and are enforced by the target handler, so a
    missing required argument surfaces as a structured
    :class:`ExcelVbaMcpError` rather than a crash. This tool is read-only.

    Args:
        action: The read verb to perform.
        workbook_id: The id returned by ``excel_open`` (required).
        sheet: The worksheet name (required).
        address: An A1-style address (required).
        offset: 0-based row index (``read-values`` / ``read-formulas``).
        limit: Optional cap on returned rows (``read-values`` / ``read-formulas``).

    Returns:
        A :class:`RangeAddress` (``get``), :class:`RangeValues`
        (``read-values``), :class:`RangeFormulas` (``read-formulas``),
        :class:`FormatInfo` (``get-format``), :class:`ValidationRule`
        (``get-validation``) or ``list[ConditionalFormatRule]``
        (``get-conditional-formatting``).
    """

    if action == "get":
        return await range_get(workbook_id or "", sheet or "", address or "")
    if action == "read-values":
        return await range_read_values(
            workbook_id or "", sheet or "", address or "", offset, limit
        )
    if action == "read-formulas":
        return await range_read_formulas(
            workbook_id or "", sheet or "", address or "", offset, limit
        )
    if action == "get-format":
        return await range_get_format(workbook_id or "", sheet or "", address or "")
    if action == "get-validation":
        return await range_get_validation(workbook_id or "", sheet or "", address or "")
    if action == "get-conditional-formatting":
        return await range_get_conditional_formatting(
            workbook_id or "", sheet or "", address or ""
        )
    raise ExcelVbaMcpError(
        f"Unknown range_read action {action!r}; expected 'get', 'read-values', "
        "'read-formulas', 'get-format', 'get-validation' or "
        "'get-conditional-formatting'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "range_read", "action": action},
    )


async def range_write(
    action: Literal[
        "write-values",
        "write-formulas",
        "set-format",
        "set-validation",
        "set-conditional-formatting",
    ],
    workbook_id: str | None = None,
    sheet: str | None = None,
    address: str | None = None,
    payload: RangeWriteRequest | FormulaWriteRequest | None = None,
    patch: RangeFormatPatch | None = None,
    spec: ValidationSpec | ConditionalFormatSpec | None = None,
    mode: str = "add",
) -> RangeMutationResult | FormatMutationResult | ValidationRule | ConditionalFormatMutationResult:
    """Range write dispatch: ``write-values``, ``write-formulas``,
    ``set-format``, ``set-validation`` or ``set-conditional-formatting`` (WP7).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`range_write_values` /
    :func:`range_write_formulas` / :func:`range_set_format` /
    :func:`range_set_validation` / :func:`range_set_conditional_formatting`
    handlers unchanged and forwards only the validated arguments. Per-action
    required parameters are enforced here and by the target handler:
    ``write-values`` / ``write-formulas`` require ``payload``;
    ``set-format`` requires ``address`` + ``patch``; ``set-validation`` and
    ``set-conditional-formatting`` require ``address`` + ``spec``. A missing
    required argument surfaces as a structured :class:`ExcelVbaMcpError`
    rather than a crash.

    This tool is **not** idempotent: ``set-conditional-formatting`` with
    ``mode='add'`` appends a rule and renumbers the sheet-wide priorities, so a
    repeated call is not a no-op.

    Args:
        action: The write verb to perform.
        workbook_id: The id returned by ``excel_open`` (required).
        sheet: The worksheet name (required).
        address: An A1-style target address (required).
        payload: The bounded value/formula grid (``write-values`` /
            ``write-formulas``).
        patch: The sparse format update (``set-format``).
        spec: The validation / conditional-formatting rule (``set-validation`` /
            ``set-conditional-formatting``).
        mode: ``add`` / ``modify`` / ``delete`` (``set-conditional-formatting``).

    Returns:
        A :class:`RangeMutationResult` (``write-values`` / ``write-formulas``),
        :class:`FormatMutationResult` (``set-format``), :class:`ValidationRule`
        (``set-validation``) or :class:`ConditionalFormatMutationResult`
        (``set-conditional-formatting``).
    """

    if action == "write-values":
        if payload is None:
            raise ExcelVbaMcpError(
                "Argument 'payload' is required for action 'write-values'.",
                code=ErrorCode.INVALID_RANGE,
                context={"tool": "range_write", "action": action, "field": "payload"},
            )
        return await range_write_values(
            workbook_id or "",
            sheet or "",
            address or "",
            cast(RangeWriteRequest, payload),
        )
    if action == "write-formulas":
        if payload is None:
            raise ExcelVbaMcpError(
                "Argument 'payload' is required for action 'write-formulas'.",
                code=ErrorCode.INVALID_RANGE,
                context={"tool": "range_write", "action": action, "field": "payload"},
            )
        return await range_write_formulas(
            workbook_id or "",
            sheet or "",
            address or "",
            cast(FormulaWriteRequest, payload),
        )
    if action == "set-format":
        if patch is None:
            raise ExcelVbaMcpError(
                "Argument 'patch' is required for action 'set-format'.",
                code=ErrorCode.INVALID_RANGE,
                context={"tool": "range_write", "action": action, "field": "patch"},
            )
        return await range_set_format(
            workbook_id or "",
            sheet or "",
            address or "",
            patch,
        )
    if action == "set-validation":
        if spec is None:
            raise ExcelVbaMcpError(
                "Argument 'spec' is required for action 'set-validation'.",
                code=ErrorCode.INVALID_RANGE,
                context={"tool": "range_write", "action": action, "field": "spec"},
            )
        return await range_set_validation(
            workbook_id or "",
            sheet or "",
            address or "",
            cast(ValidationSpec, spec),
        )
    if action == "set-conditional-formatting":
        if spec is None:
            raise ExcelVbaMcpError(
                "Argument 'spec' is required for action "
                "'set-conditional-formatting'.",
                code=ErrorCode.INVALID_RANGE,
                context={"tool": "range_write", "action": action, "field": "spec"},
            )
        return await range_set_conditional_formatting(
            workbook_id or "",
            sheet or "",
            address or "",
            cast(ConditionalFormatSpec, spec),
            mode=mode,
        )
    raise ExcelVbaMcpError(
        f"Unknown range_write action {action!r}; expected 'write-values', "
        "'write-formulas', 'set-format', 'set-validation' or "
        "'set-conditional-formatting'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "range_write", "action": action},
    )


async def range_structural(
    action: Literal["copy", "insert", "unmerge"],
    workbook_id: str | None = None,
    sheet: str | None = None,
    address: str | None = None,
    source: str | None = None,
    destination_sheet: str | None = None,
    destination: str | None = None,
    shift: str | None = None,
    *,
    allow_merged: bool = False,
) -> StructuralMutationResult:
    """Range structural dispatch: ``copy``, ``insert`` or ``unmerge`` (WP7).

    Thin flat-signature dispatcher (option (a) — flat args + handler
    validation): routes to the existing :func:`range_copy` /
    :func:`range_insert` / :func:`range_unmerge` handlers unchanged and forwards
    only the validated arguments. Per-action required parameters are enforced by
    the target handler: ``copy`` requires ``source`` / ``destination_sheet`` /
    ``destination``; ``insert`` requires ``address`` + an explicit ``shift`` in
    ``{right, down}`` (an omitted or disallowed shift is rejected, never
    guessed); ``unmerge`` requires ``address``. A missing required argument
    surfaces as a structured :class:`ExcelVbaMcpError` rather than a crash.

    This tool is **not** idempotent (H4): re-``insert`` duplicates cells and
    ``unmerge`` discards merges.

    Args:
        action: The structural verb to perform.
        workbook_id: The id returned by ``excel_open`` (required).
        sheet: The source worksheet name (required).
        address: An A1-style target address (``insert`` / ``unmerge``).
        source: The source A1-style address (``copy``).
        destination_sheet: The destination worksheet (``copy``).
        destination: The destination A1-style address (``copy``).
        shift: ``right`` or ``down`` (``insert``; required, never guessed).
        allow_merged: Permit inserting over merged cells after preview.

    Returns:
        A :class:`StructuralMutationResult`.
    """

    if action == "copy":
        return await range_copy(
            workbook_id or "",
            sheet or "",
            source or "",
            destination_sheet or "",
            destination or "",
        )
    if action == "insert":
        return await range_insert(
            workbook_id or "",
            sheet or "",
            address or "",
            shift or "",
            allow_merged=allow_merged,
        )
    if action == "unmerge":
        return await range_unmerge(workbook_id or "", sheet or "", address or "")
    raise ExcelVbaMcpError(
        f"Unknown range_structural action {action!r}; expected 'copy', 'insert' "
        "or 'unmerge'.",
        code=ErrorCode.INVALID_RANGE,
        context={"tool": "range_structural", "action": action},
    )


__all__ = [
    "range_copy",
    "range_delete",
    "range_get",
    "range_get_conditional_formatting",
    "range_get_format",
    "range_get_validation",
    "range_insert",
    "range_merge",
    "range_read",
    "range_read_formulas",
    "range_read_values",
    "range_set_conditional_formatting",
    "range_set_format",
    "range_set_validation",
    "range_structural",
    "range_unmerge",
    "range_write",
    "range_write_formulas",
    "range_write_values",
]
