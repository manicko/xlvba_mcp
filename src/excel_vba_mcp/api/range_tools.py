"""Range & Cells MCP tool contracts (task_040 schema group).

This module defines the **signatures, Pydantic DTOs and docstrings** for the
Phase 2 Range/Cells surface. It is a frozen contract only (task_040): the bodies
raise :class:`NotImplementedError` and are implemented by later tasks
(``task_042``–``task_046``). This module deliberately imports **no Excel/COM**
(``pythoncom`` / ``win32com``) — it only imports domain DTOs. All Excel access
happens on the single-STA worker inside the ranges service/adapter.

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

The read-only tools (``range_get``, ``range_read_values``,
``range_read_formulas``, task_042) are implemented here as thin async handlers
that validate their arguments, delegate to
:class:`~excel_vba_mcp.application.ranges_service.RangesService` and return a
domain DTO. The remaining stubs are implemented by later tasks
(``task_043``–``task_046``).
"""

from __future__ import annotations

from excel_vba_mcp.application.ranges_service import RangesService
from excel_vba_mcp.domain.errors import ExcelVbaMcpError
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

_NOT_IMPLEMENTED = (
    "Range/Cells tool {name} is a frozen contract (task_040) and is implemented "
    "in a later task; it performs no Excel/COM access here."
)


def _not_implemented(name: str) -> NotImplementedError:
    """Return the canonical not-yet-implemented error for a contract stub."""

    return NotImplementedError(_NOT_IMPLEMENTED.format(name=name))


# ---------------------------------------------------------------------------
# Active service wiring (read tools; published by the server lifespan)
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

    raise _not_implemented("range_get_format")


async def range_get_validation(workbook_id: str, sheet: str, address: str) -> ValidationRule:
    """Read the data-validation rule of a Range (``present=false`` when none).

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.
    """

    raise _not_implemented("range_get_validation")


async def range_get_conditional_formatting(
    workbook_id: str, sheet: str, address: str
) -> list[ConditionalFormatRule]:
    """Read the conditional-formatting rules applying to a Range.

    Supported plain ``FormatCondition`` rules are returned with priority and
    ``stop_if_true``; ``ColorScale`` / ``DataBar`` / ``IconSet`` are flagged
    ``supported=false`` rather than omitted.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: An A1-style address for the target block.
    """

    raise _not_implemented("range_get_conditional_formatting")


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

    raise _not_implemented("range_set_format")


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

    raise _not_implemented("range_copy")


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

    raise _not_implemented("range_insert")


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

    raise _not_implemented("range_delete")


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

    raise _not_implemented("range_merge")


async def range_unmerge(workbook_id: str, sheet: str, address: str) -> StructuralMutationResult:
    """Unmerge any merged area within a Range.

    Guarded mutation: writable session + snapshot + approval.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
    """

    raise _not_implemented("range_unmerge")


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
    """

    raise _not_implemented("range_set_validation")


async def range_set_conditional_formatting(
    workbook_id: str,
    sheet: str,
    address: str,
    spec: ConditionalFormatSpec,
    *,
    mode: str = "add",
) -> ConditionalFormatMutationResult:
    """Add / modify / delete a supported conditional-formatting rule.

    ``mode`` is ``"add"``, ``"modify"`` or ``"delete"``; priority ordering is
    preserved. Unsupported rule types are rejected explicitly. Guarded mutation:
    writable session + snapshot + approval + postcondition re-read.

    Args:
        workbook_id: The id returned by ``excel_open``.
        sheet: The worksheet name.
        address: The A1-style target address.
        spec: The conditional-formatting rule payload.
        mode: One of ``"add"``, ``"modify"``, ``"delete"``.
    """

    raise _not_implemented("range_set_conditional_formatting")


__all__ = [
    "range_copy",
    "range_delete",
    "range_get",
    "range_get_conditional_formatting",
    "range_get_format",
    "range_get_validation",
    "range_insert",
    "range_merge",
    "range_read_formulas",
    "range_read_values",
    "range_set_conditional_formatting",
    "range_set_format",
    "range_set_validation",
    "range_unmerge",
    "range_write_formulas",
    "range_write_values",
]
