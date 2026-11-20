"""STA-local Excel Range/Cells adapter (task_042 read, task_043 write, task_044
format, task_045 structural, task_046 validation + conditional formatting).

Phase 2 — Range/Cells.

:class:`RangesAdapter` is the single place where the Excel ``Worksheet.Range`` /
``Range.Value2`` / ``Range.Formula2`` object model is touched for the Phase-2
surface (``range_get``, ``range_read_values``, ``range_read_formulas``,
``range_write_values``, ``range_write_formulas``, ``range_get_format``,
``range_set_format``, ``range_copy``, ``range_insert``, ``range_delete``,
``range_merge``, ``range_unmerge``, ``range_get_validation``,
``range_set_validation``, ``range_get_conditional_formatting``,
``range_set_conditional_formatting``).
It is a **pure STA-local helper**: every method is synchronous, receives the live
``Workbook`` COM proxy (``Any``) and returns only Pydantic DTOs / primitives, so
results cross the single-STA worker boundary safely
(:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never imports
``pywin32`` at module import time and never calls ``.Activate`` / ``.Select`` on
any object (addendum BG03/BG04/BG05).

Read semantics (task_038 decision record, ``docs/research/ranges_cells_api.md``):

* **Address.** An explicit ``sheet`` + A1-style ``address``. Whole-row (``1:3``)
  and whole-column (``A:C``) targets are supported **only within the configured
  bounds**; an invalid, ambiguous (multi-area) or oversized target is rejected
  with a typed :class:`~excel_vba_mcp.domain.errors.InvalidRangeError`.
* **Values.** ``Range.Value2`` is the raw channel (scalar for one cell, 2-D for
  any multi-cell block). Dates are typed **only** from ``Range.NumberFormat``
  (never guessed from a bare serial). Classification reuses the Stage-1
  :func:`~excel_vba_mcp.adapters.excel.cell_values.classify_cell_value`.
* **Formulas.** ``Range.Formula2`` is the primary channel on build 14332;
  ``Range.Formula`` is the fallback. ``_xlfn`` / ``_xlws`` / ``_xlpm`` prefixes
  are passed through verbatim and flagged in ``notes``.

Write semantics (task_043):

* ``write_values`` accepts a scalar or a 2-D block; the payload shape is
  validated against the resolved target **before** any COM mutation.
* ``write_formulas`` writes through ``Formula2`` (primary) with a ``Formula``
  fallback on a build without dynamic arrays.
* ListObject targets are allowed only for cell-level writes that do **not** change
  the table's row/column structure; the adapter captures before/after
  ``ListObject`` geometry and rejects a structural change with
  :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError`.
* The adapter never evaluates a formula; strings are passed verbatim to Excel.

Format semantics (task_044, ``docs/research/ranges_cells_api.md`` §4):

* ``get_format`` reads a **bounded, tri-state** format DTO (``number_format``;
  ``font.{name,bold,italic,size,color}``; ``fill.{color,pattern}``;
  ``alignment.{horizontal,vertical,wrap_text}``). Each property is ``value`` /
  ``mixed`` (Excel returns ``Null`` when the cells disagree — the property is
  ``None`` and named in ``mixed_properties``) / ``unsupported`` (aspects the DTO
  deliberately does not model, always reported in ``unsupported``). Mixed
  formatting is never flattened to one cell's value.
* ``set_format`` applies **only** the fields the caller explicitly set
  (``RangeFormatPatch.model_fields_set``); absent fields are left untouched.
  Values/enums are validated before any COM mutation; a present ``None`` clears
  to automatic only where Excel allows (``number_format`` -> ``General``,
  ``font_color`` -> automatic, ``fill_color`` -> no fill). The targeted
  properties are re-read after the write as a postcondition.

Bounds (task_038 §3.2/§3.3): ``MAX_RANGE_CELLS`` / ``MAX_RESPONSE_CELLS`` are
enforced **before** materializing a response; an oversized request is rejected
and an oversized response is truncated (``truncated=True``). The whole
``UsedRange`` is never read unless it is explicitly requested within limits.

Structural semantics (task_045, ``docs/research/ranges_cells_api.md`` §5):

* ``copy_range`` always takes an **explicit destination** (the clipboard is never
  used) and rejects an overlapping source/destination, a shape-mismatched
  multi-cell destination and a destination that intersects merged cells.
* ``insert_range`` / ``delete_range`` always take an **explicit shift**
  (``right`` / ``down`` and ``left`` / ``up``); an omitted or unknown shift is
  rejected rather than left to Excel's shape-based inference.
* ``merge_range`` previews the non-anchor cells that hold data and blocks the
  merge with ``INVALID_RANGE`` unless ``allow_data_loss`` is set; ``across``
  merges one cell per row.
* ``unmerge_range`` only acts when a merged area actually intersects the target.
* **Hard ListObject guard:** any structural target that intersects a
  ``ListObject.Range`` (``Application.Intersect``) is rejected with
  :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError` and routed to
  the table tools; a table's structure is never changed here.
* Merged-cell targets, protected sheets and dynamic-array spill ranges are
  guarded per task_038 §7 (reject, or preview + ``allow_merged``).

Validation / conditional-format semantics (task_046, ``docs/research/ranges_cells_api.md`` §6):

* ``get_validation`` reads ``Range.Validation`` (type, ``Formula1``/``Formula2``,
  ``IgnoreBlank``, ``InCellDropdown``, ``Operator``, ``AlertStyle`` and the
  message/title strings). Excel answers the range-level read only when the rule
  is uniform, so a non-uniform target is probed in a bounded way and reported
  with ``non_uniform_validation`` / ``mixed_validation`` in ``unsupported``.
* ``set_validation`` adds (``Validation.Add``) or modifies
  (``Validation.Modify``) the rule; ``Formula1``/``Formula2`` are opaque strings
  bounded at 255 characters and never evaluated.
* ``get_conditional_formatting`` / ``set_conditional_formatting`` handle the
  plain ``FormatCondition`` types only; ``xlColorScale`` / ``xlDataBar`` /
  ``xlIconSet`` are detected and reported as unsupported instead of being
  dropped. Rule priorities are sheet-wide and unique; ``modify`` / ``delete``
  require the priority as the rule identity and a rule type is never silently
  converted.
"""

from __future__ import annotations

from collections.abc import Callable, Collection, Iterator
from contextlib import contextmanager, suppress
from datetime import datetime, timedelta
from typing import Any

from excel_vba_mcp.adapters.excel.cell_values import (
    EXCEL_ERROR_CODE_TO_LITERAL,
    classify_cell_value,
)
from excel_vba_mcp.domain.errors import (
    InvalidRangeError,
    ListObjectIntersectionError,
    ProtectedTargetError,
)
from excel_vba_mcp.domain.models import (
    MAX_RANGE_CELLS,
    MAX_RESPONSE_CELLS,
    MAX_VALIDATION_FORMULA_LENGTH,
    AlignmentFormat,
    CellValue,
    ConditionalFormatMutationResult,
    ConditionalFormatRule,
    ConditionalFormatSpec,
    FillFormat,
    FontFormat,
    FormatInfo,
    FormatMutationResult,
    FormulaCell,
    RangeAddress,
    RangeFormatPatch,
    RangeFormulas,
    RangeMutationResult,
    RangeValues,
    StructuralMutationResult,
    ValidationRule,
    ValidationSpec,
)
from excel_vba_mcp.domain.statuses import (
    CellValueKind,
    ConditionalFormatType,
    DataValidationType,
    RangeShape,
    ValidationOperator,
)

#: Formula prefixes that mark a function newer than the reading build
#: (task_038 §2.1). Presence is surfaced in ``notes``; the string itself is
#: passed through verbatim and never stripped or rewritten.
_FORMULA_PREFIXES: tuple[str, ...] = (
    "_xlfn._xlws.",
    "_xlfn._xlpm.",
    "_xlfn.",
    "_xlws.",
    "_xlpm.",
)

#: ``Range.NumberFormat`` token detection for date/time formats (task_038 §1.4).
#: A numeric value is typed as a date only when its number format carries one of
#: these tokens; ``General`` / numeric formats never produce a date (no guessing).
_DATE_FORMAT_TOKENS: tuple[str, ...] = ("yy", "dd", "mmm")

#: Excel's 1900 date system epoch (``serial 0`` renders as 1899-12-30).
_EXCEL_EPOCH = datetime(1899, 12, 30)


# ---------------------------------------------------------------------------
# Formatting enums (task_044, task_038 §4.2)
# ---------------------------------------------------------------------------
# These are the documented Excel ``XlHAlign`` / ``XlVAlign`` / ``XlPattern``
# enumeration values (Microsoft Learn). They are resolved here — the only layer
# that touches Excel constants — so the service/API validate names without COM.

#: ``Range.HorizontalAlignment`` (``XlHAlign``) name -> value.
_HORIZONTAL_ALIGNMENT_VALUES: dict[str, int] = {
    "xlGeneral": 1,  # xlHAlignGeneral
    "xlLeft": -4131,  # xlHAlignLeft
    "xlCenter": -4108,  # xlHAlignCenter
    "xlRight": -4152,  # xlHAlignRight
    "xlFill": 5,  # xlHAlignFill
    "xlJustify": -4130,  # xlHAlignJustify
    "xlCenterAcrossSelection": 7,  # xlHAlignCenterAcrossSelection
    "xlDistributed": -4117,  # xlHAlignDistributed
}

#: ``Range.VerticalAlignment`` (``XlVAlign``) name -> value.
_VERTICAL_ALIGNMENT_VALUES: dict[str, int] = {
    "xlTop": -4160,  # xlVAlignTop
    "xlCenter": -4108,  # xlVAlignCenter
    "xlBottom": -4107,  # xlVAlignBottom
    "xlJustify": -4130,  # xlVAlignJustify
    "xlDistributed": -4117,  # xlVAlignDistributed
}

#: ``Range.Interior.Pattern`` (``XlPattern``) name -> value.
_FILL_PATTERN_VALUES: dict[str, int] = {
    "xlNone": -4142,  # xlPatternNone
    "xlSolid": 1,  # xlPatternSolid
    "xlAutomatic": -4105,  # xlPatternAutomatic
    "xlGray25": -4124,  # xlPatternGray25
    "xlGray50": -4125,  # xlPatternGray50
    "xlGray75": -4126,  # xlPatternGray75
    "xlGray8": 18,  # xlPatternGray8
    "xlGray16": 17,  # xlPatternGray16
    "xlHorizontal": -4128,  # xlPatternHorizontal
    "xlVertical": -4166,  # xlPatternVertical
    "xlDown": -4121,  # xlPatternDown
    "xlUp": -4162,  # xlPatternUp
    "xlChecker": 9,  # xlPatternChecker
    "xlGrid": 15,  # xlPatternGrid
    "xlCrissCross": 16,  # xlPatternCrissCross
    "xlLightHorizontal": 11,  # xlPatternLightHorizontal
    "xlLightVertical": 12,  # xlPatternLightVertical
    "xlLightDown": 13,  # xlPatternLightDown
    "xlLightUp": 14,  # xlPatternLightUp
    "xlSemiGray75": 10,  # xlPatternSemiGray75
    "xlLinearGradient": 4000,  # xlPatternLinearGradient
    "xlRectangularGradient": 4001,  # xlPatternRectangularGradient
}

#: Reverse lookups (value -> name) for reading a format property back.
_HORIZONTAL_ALIGNMENT_NAMES: dict[int, str] = {
    value: name for name, value in _HORIZONTAL_ALIGNMENT_VALUES.items()
}
_VERTICAL_ALIGNMENT_NAMES: dict[int, str] = {
    value: name for name, value in _VERTICAL_ALIGNMENT_VALUES.items()
}
_FILL_PATTERN_NAMES: dict[int, str] = {value: name for name, value in _FILL_PATTERN_VALUES.items()}

#: ``XlColorIndex`` sentinels used to detect an automatic / absent font/fill color.
_XL_COLOR_INDEX_AUTOMATIC = -4105  # xlColorIndexAutomatic
_XL_COLOR_INDEX_NONE = -4142  # xlColorIndexNone

#: RGB color ceiling (``0xFFFFFF``) and font-size bounds accepted by Excel.
_MAX_RGB_COLOR = 0xFFFFFF
_MIN_FONT_SIZE = 1.0
_MAX_FONT_SIZE = 1638.0

#: ``Range.NumberFormat`` / ``Font.Name`` string length ceiling (Excel UI limit).
_MAX_FORMAT_STRING_LENGTH = 255

#: Range formatting aspects that exist but are intentionally not represented by
#: :class:`~excel_vba_mcp.domain.models.FormatInfo` (task_038 §4.2). Reported on
#: every read so a caller never assumes the bounded DTO is the full format state.
_UNSUPPORTED_FORMAT_PROPERTIES: tuple[str, ...] = (
    "borders",
    "gradient",
    "theme_color",
    "number_format_local",
    "protect_contents",
)

#: Format fields a caller may clear to automatic by passing ``None`` (task_038
#: §4.3, "null explicitly clears to automatic where Excel allows"). Every other
#: field rejects an explicit ``None`` because it has no automatic state.
_CLEARABLE_FORMAT_FIELDS: frozenset[str] = frozenset({"number_format", "font_color", "fill_color"})

# ---------------------------------------------------------------------------
# Structural operations (task_045, task_038 §5)
# ---------------------------------------------------------------------------
# These are the documented Excel ``XlInsertShiftDirection`` /
# ``XlDeleteShiftDirection`` / ``XlInsertCopyOrigin`` enumeration values
# (Microsoft Learn). They are resolved here — the only layer that touches Excel
# constants — so the service/API validate the shift *name* without COM. The
# ``Shift`` argument is always explicit: an omitted/ambiguous shift is rejected
# and Excel is never allowed to guess it from the range shape (task_038 §7).

#: ``Range.Insert`` shift direction (``XlInsertShiftDirection``).
_XL_SHIFT_TO_RIGHT = -4161  # xlShiftToRight
_XL_SHIFT_DOWN = -4121  # xlShiftDown

#: ``Range.Delete`` shift direction (``XlDeleteShiftDirection``).
_XL_SHIFT_TO_LEFT = -4159  # xlShiftToLeft
_XL_SHIFT_UP = -4162  # xlShiftUp

#: ``Range.Insert`` copy origin (``XlInsertCopyOrigin``): format from the cells
#: above/left (the Excel default), never guessed at random.
_XL_FORMAT_FROM_LEFT_OR_ABOVE = 0  # xlFormatFromLeftOrAbove

#: Shift name -> ``Range.Insert`` enum value.
_INSERT_SHIFT_VALUES: dict[str, int] = {
    "right": _XL_SHIFT_TO_RIGHT,
    "down": _XL_SHIFT_DOWN,
}

#: Shift name -> ``Range.Delete`` enum value.
_DELETE_SHIFT_VALUES: dict[str, int] = {
    "left": _XL_SHIFT_TO_LEFT,
    "up": _XL_SHIFT_UP,
}

#: Upper bound on the cells scanned when locating merged areas / non-anchor data
#: inside a partially merged target. A larger target is reported through
#: ``preconditions`` rather than scanned cell-by-cell over COM.
_MAX_MERGE_SCAN_CELLS = 4096

#: Upper bound on the number of referencing cells reported by the best-effort
#: same-sheet formula scan (task_038 §5.2 item 5).
_MAX_REFERENCE_HITS = 50


# ---------------------------------------------------------------------------
# Data validation (task_046, task_038 §6.1)
# ---------------------------------------------------------------------------
# Documented ``XlDVType`` / ``XlFormatConditionOperator`` / ``XlAlertStyle``
# values (Microsoft Learn). They are resolved here — the only layer that touches
# Excel constants — so the service/API can validate a rule name without COM.

XL_VALIDATE_INPUT_ONLY = 0  # xlValidateInputOnly
XL_VALIDATE_WHOLE_NUMBER = 1  # xlValidateWholeNumber
XL_VALIDATE_DECIMAL = 2  # xlValidateDecimal
XL_VALIDATE_LIST = 3  # xlValidateList
XL_VALIDATE_DATE = 4  # xlValidateDate
XL_VALIDATE_TIME = 5  # xlValidateTime
XL_VALIDATE_TEXT_LENGTH = 6  # xlValidateTextLength
XL_VALIDATE_CUSTOM = 7  # xlValidateCustom

XL_COND_BETWEEN = 1  # xlBetween
XL_COND_EQUAL = 2  # xlEqual
XL_COND_GREATER = 3  # xlGreater
XL_COND_GREATER_EQUAL = 4  # xlGreaterEqual
XL_COND_LESS = 5  # xlLess
XL_COND_LESS_EQUAL = 6  # xlLessEqual
XL_COND_NOT_BETWEEN = 7  # xlNotBetween
XL_COND_NOT_EQUAL = 8  # xlNotEqual

XL_ALERT_INFORMATION = 1  # xlValidAlertInformation
XL_ALERT_STOP = 2  # xlValidAlertStop
XL_ALERT_WARNING = 3  # xlValidAlertWarning

#: ``Range.Validation.Type`` -> :class:`DataValidationType` (``NONE`` has no int).
_VALIDATION_TYPE_BY_INT: dict[int, DataValidationType] = {
    XL_VALIDATE_INPUT_ONLY: DataValidationType.INPUT_ONLY,
    XL_VALIDATE_WHOLE_NUMBER: DataValidationType.WHOLE_NUMBER,
    XL_VALIDATE_DECIMAL: DataValidationType.DECIMAL,
    XL_VALIDATE_LIST: DataValidationType.LIST,
    XL_VALIDATE_DATE: DataValidationType.DATE,
    XL_VALIDATE_TIME: DataValidationType.TIME,
    XL_VALIDATE_TEXT_LENGTH: DataValidationType.TEXT_LENGTH,
    XL_VALIDATE_CUSTOM: DataValidationType.CUSTOM,
}
_VALIDATION_TYPE_VALUES: dict[DataValidationType, int] = {
    value: key for key, value in _VALIDATION_TYPE_BY_INT.items()
}

_OPERATOR_BY_INT: dict[int, ValidationOperator] = {
    XL_COND_BETWEEN: ValidationOperator.BETWEEN,
    XL_COND_EQUAL: ValidationOperator.EQUAL,
    XL_COND_GREATER: ValidationOperator.GREATER,
    XL_COND_GREATER_EQUAL: ValidationOperator.GREATER_EQUAL,
    XL_COND_LESS: ValidationOperator.LESS,
    XL_COND_LESS_EQUAL: ValidationOperator.LESS_EQUAL,
    XL_COND_NOT_BETWEEN: ValidationOperator.NOT_BETWEEN,
    XL_COND_NOT_EQUAL: ValidationOperator.NOT_EQUAL,
}
_OPERATOR_VALUES: dict[ValidationOperator, int] = {
    value: key for key, value in _OPERATOR_BY_INT.items()
}

_ALERT_STYLE_BY_INT: dict[int, str] = {
    XL_ALERT_INFORMATION: "xlValidAlertInformation",
    XL_ALERT_STOP: "xlValidAlertStop",
    XL_ALERT_WARNING: "xlValidAlertWarning",
}
_ALERT_STYLE_VALUES: dict[str, int] = {value: key for key, value in _ALERT_STYLE_BY_INT.items()}

#: Validation types that use only ``Formula1``; Excel ignores ``Formula2``.
_VALIDATION_SINGLE_FORMULA_TYPES: frozenset[DataValidationType] = frozenset(
    {DataValidationType.LIST, DataValidationType.CUSTOM}
)

#: Operators that require **both** operands (task_038 §6.1).
_DUAL_FORMULA_OPERATORS: frozenset[ValidationOperator] = frozenset(
    {ValidationOperator.BETWEEN, ValidationOperator.NOT_BETWEEN}
)

#: Excel UI limit for a validation message / title string (task_038 §6.3).
MAX_VALIDATION_TEXT_LENGTH = 255

#: Upper bound on the cells probed when Excel refuses to read a range-level
#: validation rule (a mixed or partially covered target). The probe is
#: deterministic (first, last and evenly spaced cells) and bounded.
_MAX_VALIDATION_PROBE_CELLS = 32


# ---------------------------------------------------------------------------
# Conditional formatting (task_046, task_038 §6.3)
# ---------------------------------------------------------------------------
# Documented ``XlFormatConditionType`` values (Microsoft Learn). Only the plain
# ``FormatCondition`` types are supported; ``xlColorScale`` / ``xlDataBar`` /
# ``xlIconSet`` are distinct COM objects and are reported as unsupported rather
# than silently dropped.

XL_CF_CELL_VALUE = 1  # xlCellValue
XL_CF_EXPRESSION = 2  # xlExpression
XL_CF_COLOR_SCALE = 3  # xlColorScale (distinct COM object)
XL_CF_DATA_BAR = 4  # xlDataBar (distinct COM object)
XL_CF_TOP10 = 5  # xlTop10
XL_CF_ICON_SET = 6  # xlIconSet (distinct COM object)
XL_CF_DUPLICATE_VALUES = 7  # xlDuplicateValues (not modelled)
XL_CF_UNIQUE_VALUES = 8  # xlUniqueValues
XL_CF_TEXT_STRING = 9  # xlTextString
XL_CF_BLANKS = 10  # xlBlanksCondition
XL_CF_TIME_PERIOD = 11  # xlTimePeriod
XL_CF_ABOVE_AVERAGE = 12  # xlAboveAverageCondition
XL_CF_NO_BLANKS = 13  # xlNoBlanksCondition
XL_CF_ERRORS = 16  # xlErrorsCondition
XL_CF_NO_ERRORS = 17  # xlNoErrorsCondition

_CF_TYPE_BY_INT: dict[int, ConditionalFormatType] = {
    XL_CF_CELL_VALUE: ConditionalFormatType.CELL_VALUE,
    XL_CF_EXPRESSION: ConditionalFormatType.EXPRESSION,
    XL_CF_COLOR_SCALE: ConditionalFormatType.COLOR_SCALE,
    XL_CF_DATA_BAR: ConditionalFormatType.DATA_BAR,
    XL_CF_TOP10: ConditionalFormatType.TOP10,
    XL_CF_ICON_SET: ConditionalFormatType.ICON_SET,
    XL_CF_DUPLICATE_VALUES: ConditionalFormatType.UNSUPPORTED,
    XL_CF_UNIQUE_VALUES: ConditionalFormatType.UNIQUE_VALUES,
    XL_CF_TEXT_STRING: ConditionalFormatType.TEXT_STRING,
    XL_CF_BLANKS: ConditionalFormatType.BLANKS,
    XL_CF_TIME_PERIOD: ConditionalFormatType.TIME_PERIOD,
    XL_CF_ABOVE_AVERAGE: ConditionalFormatType.ABOVE_AVERAGE,
    XL_CF_NO_BLANKS: ConditionalFormatType.NO_BLANKS,
    XL_CF_ERRORS: ConditionalFormatType.ERRORS,
    XL_CF_NO_ERRORS: ConditionalFormatType.NO_ERRORS,
}
_CF_TYPE_VALUES: dict[ConditionalFormatType, int] = {
    value: key
    for key, value in _CF_TYPE_BY_INT.items()
    if value is not ConditionalFormatType.UNSUPPORTED
}

#: Plain ``FormatCondition`` types supported by ``range_set_conditional_formatting``.
_SUPPORTED_CF_TYPES: frozenset[ConditionalFormatType] = frozenset(
    {
        ConditionalFormatType.CELL_VALUE,
        ConditionalFormatType.EXPRESSION,
        ConditionalFormatType.TEXT_STRING,
        ConditionalFormatType.BLANKS,
        ConditionalFormatType.NO_BLANKS,
        ConditionalFormatType.ERRORS,
        ConditionalFormatType.NO_ERRORS,
        ConditionalFormatType.TIME_PERIOD,
        ConditionalFormatType.TOP10,
        ConditionalFormatType.ABOVE_AVERAGE,
        ConditionalFormatType.UNIQUE_VALUES,
    }
)

#: Rule types that require an ``Operator`` argument to ``FormatConditions.Add``.
_CF_OPERATOR_REQUIRED_TYPES: frozenset[ConditionalFormatType] = frozenset(
    {ConditionalFormatType.CELL_VALUE, ConditionalFormatType.TEXT_STRING}
)

#: Rule types that require a ``Formula1`` operand.
_CF_FORMULA_REQUIRED_TYPES: frozenset[ConditionalFormatType] = frozenset(
    {
        ConditionalFormatType.CELL_VALUE,
        ConditionalFormatType.EXPRESSION,
        ConditionalFormatType.TEXT_STRING,
        ConditionalFormatType.TOP10,
        ConditionalFormatType.TIME_PERIOD,
    }
)

#: Explicit reasons for the rule types that are detected but not supported.
_CF_UNSUPPORTED_REASONS: dict[int, str] = {
    XL_CF_COLOR_SCALE: (
        "xlColorScale is a distinct COM object (ColorScale); outside the bounded Phase 2 scope."
    ),
    XL_CF_DATA_BAR: (
        "xlDataBar is a distinct COM object (Databar); outside the bounded Phase 2 scope."
    ),
    XL_CF_ICON_SET: (
        "xlIconSet is a distinct COM object (IconSetCondition); outside the bounded Phase 2 scope."
    ),
    XL_CF_DUPLICATE_VALUES: (
        "duplicate-values rules (type 7) are not modelled by the bounded Phase 2 scope."
    ),
}

#: Conditional-formatting mutation modes accepted by ``range_set_conditional_formatting``.
CF_MODES: frozenset[str] = frozenset({"add", "modify", "delete"})


# ---------------------------------------------------------------------------
# Safe COM readers (never raise; no proxy leaks)
# ---------------------------------------------------------------------------
def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively (never raises)."""

    if value is None:
        return default
    try:
        text = str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    return text if text else default


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively (never raises)."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


# ---------------------------------------------------------------------------
# Address parsing / normalization (no COM)
# ---------------------------------------------------------------------------
def _col_index(letters: str) -> int:
    """Return the 1-based column index for Excel column ``letters`` (e.g. ``A`` -> 1)."""

    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


def _col_letter(index: int) -> str:
    """Return the Excel column letter(s) for a 1-based column ``index``."""

    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _cell_address(row: int, col: int) -> str:
    """Return an A1-style address (no ``$``) for a 1-based ``(row, col)`` pair."""

    return f"{_col_letter(col)}{row}"


def _require_sheet(wb: Any, name: str) -> Any:
    """Resolve ``name`` to a worksheet proxy or raise a structured error."""

    if not isinstance(name, str) or not name.strip():
        raise InvalidRangeError(
            "Argument 'sheet' must be a non-empty string.", context={"sheet": name}
        )
    try:
        return wb.Worksheets.Item(name)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            f"Worksheet {name!r} was not found in the workbook.", context={"sheet": name}
        ) from exc


def _resolve_range(worksheet: Any, address: str, *, sheet_name: str) -> Any:
    """Resolve an A1 ``address`` to a single-area Range on ``worksheet``.

    Rejects multi-area and unparseable addresses. Cell-count bounds are enforced
    by the callers (:meth:`RangesAdapter.read_values` / ``read_formulas``) before
    any cell is read.
    """

    if not isinstance(address, str) or not address.strip():
        raise InvalidRangeError(
            "Argument 'address' must be a non-empty A1-style string.",
            context={"sheet": sheet_name, "address": address},
        )
    try:
        range_obj = worksheet.Range(address.strip())
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            f"Range {address!r} is not valid on sheet {sheet_name!r}.",
            context={"sheet": sheet_name, "address": address},
        ) from exc

    # Multi-area guard (task_038 §7): reject rather than guess an area.
    try:
        area_count = int(range_obj.Areas.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        area_count = 1
    if area_count > 1:
        raise InvalidRangeError(
            f"Range {address!r} spans {area_count} areas (ambiguous); a single area is required.",
            context={"sheet": sheet_name, "address": address, "areas": area_count},
        )
    return range_obj


def _describe_range(range_obj: Any, *, sheet: str) -> RangeAddress:
    """Build a :class:`RangeAddress` from a resolved Range."""

    rows = _safe_int(range_obj.Rows.Count, default=0)
    cols = _safe_int(range_obj.Columns.Count, default=0)
    address = _safe_str(range_obj.Address)
    return RangeAddress(
        sheet=sheet,
        address=address,
        rows=rows,
        cols=cols,
        shape=_shape_for(rows, cols),
        areas=1,
    )


def _shape_for(rows: int, cols: int) -> RangeShape:
    """Return the normalized :class:`RangeShape` tag for a dimensions pair."""

    if rows <= 1 and cols <= 1:
        return RangeShape.SCALAR
    if rows <= 1:
        return RangeShape.ROW
    if cols <= 1:
        return RangeShape.COLUMN
    return RangeShape.BLOCK


# ---------------------------------------------------------------------------
# Value / formula reading (STA-local)
# ---------------------------------------------------------------------------
def _normalize_grid(raw: Any) -> tuple[tuple[Any, ...], ...]:
    """Normalize a ``Range.Value2`` / ``Formula2`` read into a 2-D grid.

    Excel returns a scalar for a single-cell range and a 2-D array for any
    multi-cell range (task_038 §1.4); both are normalized to 2-D here.
    """

    if not isinstance(raw, tuple):
        return ((raw,),)
    if not raw:
        return ((),)
    if isinstance(raw[0], tuple):
        return tuple(tuple(row) for row in raw)
    return (tuple(raw),)


def _grid_cell(grid: tuple[tuple[Any, ...], ...], row: int, col: int) -> Any:
    """Return ``grid[row][col]`` defensively (``None`` when out of range)."""

    if row < len(grid):
        line = grid[row]
        if col < len(line):
            return line[col]
    return None


def _is_date_format(number_format: Any) -> bool:
    """Return ``True`` only when ``number_format`` is a date/time format.

    ``None`` (mixed) and ``General`` / numeric formats are never dates, so a
    bare numeric serial is never guessed into a date (task_038 §1.4).
    """

    if number_format is None:
        return False
    lowered = str(number_format).lower()
    if lowered in ("general", ""):
        return False
    return any(token in lowered for token in _DATE_FORMAT_TOKENS)


def _classify_cell(value: Any, *, address: str, number_format: Any) -> CellValue:
    """Classify a raw ``Value2`` value into a typed :class:`CellValue`.

    A number is promoted to :attr:`CellValueKind.DATE` (ISO-8601 string) **only**
    when its number format is a date/time format and the value is a plausible
    Excel serial; otherwise the Stage-1 ``classify_cell_value`` order applies
    (bool before int, OLE error codes -> error, etc.).
    """

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and _is_date_format(number_format)
    ):
        moment = _EXCEL_EPOCH + timedelta(days=float(value))
        return CellValue(kind=CellValueKind.DATE, value=moment.isoformat(), address=address)
    return classify_cell_value(value, address=address)


def _classify_grid(
    grid: tuple[tuple[Any, ...], ...],
    *,
    sub_range: Any,
    top_row: int,
    left_col: int,
) -> list[list[CellValue]]:
    """Classify a whole ``Value2`` grid using a single ``NumberFormat`` read.

    ``Range.NumberFormat`` is read **once** for the bounded sub-range. When it is
    ``Null`` (the cells disagree, task_038 §1.1) only then does the adapter fall
    back to per-cell ``NumberFormat`` reads for the returned cells, so a uniform
    date/number column costs one format read instead of one per cell.
    """

    rows = len(grid)
    try:
        range_format = sub_range.NumberFormat
    except Exception:  # noqa: BLE001 - defensive COM read
        range_format = None

    if range_format is not None:
        return [
            [
                _classify_cell(
                    _grid_cell(grid, r, c),
                    address=_cell_address(top_row + r, left_col + c),
                    number_format=range_format,
                )
                for c in range(len(grid[r]) if r < rows else 0)
            ]
            for r in range(rows)
        ]

    # Mixed formats: read per-cell only for the cells being returned (bounded).
    result: list[list[CellValue]] = []
    for r in range(rows):
        row_values: list[CellValue] = []
        for c in range(len(grid[r])):
            try:
                cell_format = sub_range.Cells(r + 1, c + 1).NumberFormat
            except Exception:  # noqa: BLE001 - defensive per-cell COM read
                cell_format = None
            row_values.append(
                _classify_cell(
                    _grid_cell(grid, r, c),
                    address=_cell_address(top_row + r, left_col + c),
                    number_format=cell_format,
                )
            )
        result.append(row_values)
    return result


def _sub_range(worksheet: Any, range_obj: Any, start_row: int, row_count: int, cols: int) -> Any:
    """Return the bounded sub-range for a window starting at ``start_row``."""

    top_row = _safe_int(range_obj.Row) + start_row
    left_col = _safe_int(range_obj.Column)
    return worksheet.Range(
        worksheet.Cells(top_row, left_col),
        worksheet.Cells(top_row + row_count - 1, left_col + cols - 1),
    )


def _read_formula2_grid(range_obj: Any) -> tuple[tuple[Any, ...], ...]:
    """Read ``range_obj.Formula2`` as a 2-D grid, falling back to ``Formula``.

    On the target build (14332) ``Formula2`` is available; a build without it
    (or an OLAP source) raises and the adapter transparently uses ``Formula``.
    """

    try:
        raw = range_obj.Formula2
    except Exception:  # noqa: BLE001 - Formula2 unavailable / OLAP -> fallback
        raw = range_obj.Formula
    return _normalize_grid(raw)


def _read_formula_grid(range_obj: Any) -> tuple[tuple[Any, ...], ...]:
    """Read the secondary ``range_obj.Formula`` (IIE) grid (empty on failure)."""

    try:
        raw = range_obj.Formula
    except Exception:  # noqa: BLE001 - Formula unavailable (e.g. OLAP)
        return ((),)
    return _normalize_grid(raw)


def _build_formula_cell(address: str, raw2: Any, raw1: Any) -> FormulaCell:
    """Assemble one :class:`FormulaCell` (formula2 primary, formula when different)."""

    formula2 = "" if raw2 is None else str(raw2)
    formula1 = "" if raw1 is None else str(raw1)
    formula_secondary = formula1 if formula1 != formula2 else None
    return FormulaCell(address=address, formula2=formula2, formula=formula_secondary)


def _flag_formula_prefix(cell: FormulaCell, notes: list[str]) -> None:
    """Append a note when a formula carries a newer-function prefix (verbatim)."""

    for prefix in _FORMULA_PREFIXES:
        if prefix in cell.formula2 or (cell.formula and prefix in cell.formula):
            notes.append(f"{cell.address}: formula uses the {prefix!r} prefix (newer function).")
            return


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class RangesAdapter:
    """STA-local Excel Range read/metadata operations (no COM at import)."""

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------
    def resolve_range(self, wb: Any, sheet: str, address: str) -> RangeAddress:
        """Resolve an explicit ``sheet`` + A1 ``address`` to :class:`RangeAddress`.

        Validates the sheet exists, the address parses, the target is a single
        area and within the configured bounds; rejects invalid/ambiguous/
        oversized targets with a typed :class:`InvalidRangeError`.
        """

        worksheet = _require_sheet(wb, sheet)
        range_obj = _resolve_range(worksheet, address, sheet_name=_safe_str(worksheet.Name))
        return _describe_range(range_obj, sheet=_safe_str(worksheet.Name))

    # ------------------------------------------------------------------
    # Values (read-only, typed, bounded)
    # ------------------------------------------------------------------
    def read_values(
        self,
        wb: Any,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeValues:
        """Read typed cell values from a bounded ``Range.Value2`` window.

        ``offset`` is a 0-based row index into the requested range; ``limit``
        caps the returned rows. The response is bounded by
        :data:`~excel_vba_mcp.domain.models.MAX_RESPONSE_CELLS` (an oversized
        window is truncated and ``truncated`` is set) and by
        :data:`~excel_vba_mcp.domain.models.MAX_RANGE_CELLS` (an oversized
        request is rejected before any cell is read).
        """

        _validate_window(offset, limit)
        worksheet, range_obj, resolved = self._open(wb, sheet, address)

        top_row = _safe_int(range_obj.Row)
        left_col = _safe_int(range_obj.Column)
        total_rows = resolved.rows
        total_cols = resolved.cols

        start, end, truncated = _window(total_rows, total_cols, offset, limit)
        row_count = end - start
        shape = _shape_for(total_rows, total_cols)

        if row_count <= 0 or total_cols <= 0:
            return RangeValues(
                range=resolved,
                shape=shape,
                rows=0,
                cols=total_cols,
                values=[],
                truncated=truncated or start >= total_rows,
            )

        sub_range = _sub_range(worksheet, range_obj, start, row_count, total_cols)
        grid = _normalize_grid(sub_range.Value2)
        values = _classify_grid(
            grid, sub_range=sub_range, top_row=top_row + start, left_col=left_col
        )

        return RangeValues(
            range=resolved,
            shape=shape,
            rows=row_count,
            cols=total_cols,
            values=values,
            truncated=truncated,
        )

    # ------------------------------------------------------------------
    # Formulas (read-only, 2-D normalized)
    # ------------------------------------------------------------------
    def read_formulas(
        self,
        wb: Any,
        sheet: str,
        address: str,
        *,
        offset: int = 0,
        limit: int | None = None,
    ) -> RangeFormulas:
        """Read typed formulas from a bounded window (``Formula2`` primary).

        ``Formula2`` is read first (Dynamic-Arrays form, build 14332); when the
        property is unavailable the adapter falls back to ``Formula``. A
        ``formula`` secondary field is included only when it differs from
        ``formula2`` (an ``@`` was added/removed by IIE). ``_xlfn`` / ``_xlws``
        / ``_xlpm`` prefixes are surfaced in ``notes``; the formula strings are
        never stripped or rewritten.
        """

        _validate_window(offset, limit)
        worksheet, range_obj, resolved = self._open(wb, sheet, address)

        top_row = _safe_int(range_obj.Row)
        left_col = _safe_int(range_obj.Column)
        total_rows = resolved.rows
        total_cols = resolved.cols

        start, end, truncated = _window(total_rows, total_cols, offset, limit)
        row_count = end - start
        shape = _shape_for(total_rows, total_cols)
        notes: list[str] = []

        if row_count <= 0 or total_cols <= 0:
            return RangeFormulas(
                range=resolved,
                shape=shape,
                rows=0,
                cols=total_cols,
                formulas=[],
                truncated=truncated or start >= total_rows,
                notes=notes,
            )

        sub_range = _sub_range(worksheet, range_obj, start, row_count, total_cols)
        formula2_grid = _read_formula2_grid(sub_range)
        formula_grid = _read_formula_grid(sub_range)

        formulas: list[list[FormulaCell]] = []
        for row_offset in range(row_count):
            excel_row = top_row + start + row_offset
            row_cells: list[FormulaCell] = []
            for col_offset in range(total_cols):
                cell_address = _cell_address(excel_row, left_col + col_offset)
                raw2 = _grid_cell(formula2_grid, row_offset, col_offset)
                raw1 = _grid_cell(formula_grid, row_offset, col_offset)
                cell = _build_formula_cell(cell_address, raw2, raw1)
                _flag_formula_prefix(cell, notes)
                row_cells.append(cell)
            formulas.append(row_cells)

        return RangeFormulas(
            range=resolved,
            shape=shape,
            rows=row_count,
            cols=total_cols,
            formulas=formulas,
            truncated=truncated,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _open(self, wb: Any, sheet: str, address: str) -> tuple[Any, Any, RangeAddress]:
        """Resolve ``sheet``/``address`` to ``(worksheet, range, RangeAddress)``."""

        worksheet = _require_sheet(wb, sheet)
        range_obj = _resolve_range(worksheet, address, sheet_name=_safe_str(worksheet.Name))
        return worksheet, range_obj, _describe_range(range_obj, sheet=_safe_str(worksheet.Name))

    # ------------------------------------------------------------------
    # Values (write, typed, bounded)
    # ------------------------------------------------------------------
    def write_values(
        self,
        wb: Any,
        sheet: str,
        address: str,
        rows: int,
        cols: int,
        values: list[list[CellValue]],
    ) -> RangeMutationResult:
        """Write typed cell values to a bounded Range and verify the result.

        The payload shape (``rows`` x ``cols``) is validated against the resolved
        target **before** any COM mutation; a wrong-sized payload raises a typed
        :class:`~excel_vba_mcp.domain.errors.InvalidRangeError` and leaves the
        workbook untouched. ``None`` clears a cell; ``datetime`` values are
        converted to an Excel serial; Excel error literals are written as error
        values; a string beginning with ``=`` is stored as text (never a formula).

        When the target intersects a ListObject, the adapter captures the table
        geometry before and after the write and rejects a structural change with
        :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError`.
        """

        worksheet, range_obj, resolved = self._open(wb, sheet, address)
        _validate_payload_shape(resolved.rows, resolved.cols, rows, cols, len(values))
        _validate_target_cells(resolved.rows * resolved.cols)

        table_before = _capture_tables(worksheet, range_obj)
        raw_grid = _values_to_raw_grid(values, rows, cols)
        _assign_value2(range_obj, raw_grid, rows, cols)
        _guard_table_structure(worksheet, range_obj, table_before)

        post = self.read_values(wb, sheet, address)
        verified = _values_match(values, post)
        notes: list[str] = []
        if not verified:
            notes.append("Postcondition mismatch: the re-read values differ from the payload.")
        return RangeMutationResult(
            action="write_values",
            range_before=resolved,
            range_after=post.range,
            verified=verified,
            postcondition_values=post,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Formulas (write, 2-D normalized)
    # ------------------------------------------------------------------
    def write_formulas(
        self,
        wb: Any,
        sheet: str,
        address: str,
        rows: int,
        cols: int,
        formulas: list[list[str]],
    ) -> RangeMutationResult:
        """Write formulas (``Formula2``) to a bounded Range and verify the result.

        The payload shape (``rows`` x ``cols``) is validated against the resolved
        target **before** any COM mutation; a wrong-sized payload raises a typed
        :class:`~excel_vba_mcp.domain.errors.InvalidRangeError` and leaves the
        workbook untouched. Every entry must be an actual formula string beginning
        with ``=`` (a bare constant belongs in :meth:`write_values`). Formula
        strings are passed verbatim to Excel (``_xlfn`` / ``_xlws`` / ``_xlpm``
        prefixes are never stripped or rewritten); no formula is evaluated in
        Python.

        When the target intersects a ListObject, the adapter captures the table
        geometry before and after the write and rejects a structural change with
        :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError`.
        """

        worksheet, range_obj, resolved = self._open(wb, sheet, address)
        _validate_payload_shape(resolved.rows, resolved.cols, rows, cols, len(formulas))
        _validate_target_cells(resolved.rows * resolved.cols)
        _require_formula_strings(formulas)

        table_before = _capture_tables(worksheet, range_obj)
        raw_grid = _formulas_to_raw_grid(formulas, rows, cols)
        _assign_formula2(range_obj, raw_grid, rows, cols)
        _guard_table_structure(worksheet, range_obj, table_before)

        post_formulas = self.read_formulas(wb, sheet, address)
        post_values = self.read_values(wb, sheet, address)
        verified = _formulas_match(formulas, post_formulas)
        notes: list[str] = []
        if not verified:
            notes.append("Postcondition mismatch: the re-read formulas differ from the payload.")
        return RangeMutationResult(
            action="write_formulas",
            range_before=resolved,
            range_after=post_values.range,
            verified=verified,
            postcondition_values=post_values,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Formatting (read-only tri-state; guarded partial write; task_044)
    # ------------------------------------------------------------------
    def get_format(self, wb: Any, sheet: str, address: str) -> FormatInfo:
        """Read the bounded, tri-state formatting of a Range (read-only).

        Each supported property is reported as a value, ``mixed`` (Excel returns
        ``Null`` when the cells disagree — the value is ``None`` and the property
        is named in ``mixed_properties``) or, for aspects outside the DTO,
        ``unsupported`` (always listed in ``unsupported``). A mixed range is never
        flattened to a single cell's value (task_038 §4.2).
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)

        mixed: list[str] = []
        number_format = _read_number_format(range_obj, mixed)
        font = _read_font(range_obj, mixed)
        fill = _read_fill(range_obj, mixed)
        alignment = _read_alignment(range_obj, mixed)

        return FormatInfo(
            range=resolved,
            number_format=number_format,
            font=font,
            fill=fill,
            alignment=alignment,
            mixed_properties=mixed,
            unsupported=list(_UNSUPPORTED_FORMAT_PROPERTIES),
        )

    def set_format(
        self, wb: Any, sheet: str, address: str, patch: RangeFormatPatch
    ) -> FormatMutationResult:
        """Apply a partial formatting update and verify the targeted properties.

        Only the fields the caller explicitly set (``patch.model_fields_set``)
        are written; absent fields are left untouched (task_038 §4.3). The patch
        is validated before any COM mutation (a bad enum/value leaves the workbook
        unchanged). After the writes the same properties are re-read as a
        postcondition and ``verified`` reflects whether every targeted property
        matches the request.
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)

        # Validate the whole patch before touching Excel (defense in depth; the
        # service already validated before submitting to the STA worker).
        validate_format_patch(patch)

        applied = patch.applied_fields()
        _apply_format(range_obj, patch)

        post = self.get_format(wb, sheet, address)
        verified = _format_matches_patch(range_obj, patch, post)
        notes: list[str] = []
        if not verified:
            notes.append("Postcondition mismatch: the re-read formatting differs from the patch.")
        return FormatMutationResult(
            action="set_format",
            range=resolved,
            applied=applied,
            ignored=[],
            verified=verified,
            result_format=post,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Data validation (read-only; guarded write; task_046)
    # ------------------------------------------------------------------
    def get_validation(self, wb: Any, sheet: str, address: str) -> ValidationRule:
        """Read the data-validation rule of a Range (``present=false`` when none).

        ``Range.Validation`` is read at the range level first: Excel answers that
        read only when every cell of the target shares the same rule. When it
        refuses (a mixed or only partially covered target) the adapter falls back
        to a bounded, deterministic cell probe (first, last and evenly spaced
        cells, at most ``_MAX_VALIDATION_PROBE_CELLS``) and reports the
        representative rule together with ``non_uniform_validation`` /
        ``mixed_validation`` in ``unsupported`` — a mixed target is never
        silently flattened to one cell's rule (task_038 §6.4).
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)
        return _read_validation_rule(range_obj, apply_to=resolved.address)

    def set_validation(
        self, wb: Any, sheet: str, address: str, spec: ValidationSpec
    ) -> ValidationRule:
        """Add or modify the data-validation rule of a Range and verify the result.

        The rule is validated before any COM mutation (a bad type, an over-long
        ``Formula1``, a missing operand or an explicit ``None`` on a rule-bearing
        type leaves the workbook untouched). An existing rule is **modified**
        (``Validation.Modify``) so the properties the caller did not set keep
        their current values; a target without a rule is **added**
        (``Validation.Add``). Formulas are opaque strings passed to Excel and are
        never evaluated here. The postcondition is a re-read of the same target;
        a mismatch is reported through ``unsupported`` (``postcondition_mismatch``)
        rather than claimed as success.
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)

        # Validate the whole rule before touching Excel (defense in depth; the
        # service already validated before submitting to the STA worker).
        validate_validation_spec(spec)
        _require_unprotected(worksheet)
        _apply_validation(range_obj, spec)

        post = _read_validation_rule(range_obj, apply_to=resolved.address)
        if not _validation_matches_spec(post, spec):
            post.unsupported.append("postcondition_mismatch")
        return post

    # ------------------------------------------------------------------
    # Conditional formatting (read-only; guarded write; task_046)
    # ------------------------------------------------------------------
    def get_conditional_formatting(
        self, wb: Any, sheet: str, address: str
    ) -> list[ConditionalFormatRule]:
        """Read the conditional-formatting rules that apply to a Range.

        Every rule is returned in ascending ``priority`` order. Plain
        ``FormatCondition`` rules carry ``type`` / ``operator`` / ``formula1`` /
        ``formula2`` / ``stop_if_true`` / ``applies_to`` plus the bounded
        ``interior_color`` and ``font`` subset. ``xlColorScale`` / ``xlDataBar`` /
        ``xlIconSet`` (distinct COM objects) are reported with ``supported=false``
        and an explicit ``unsupported_reason`` instead of being dropped.
        """

        worksheet = _require_sheet(wb, sheet)
        range_obj = _resolve_range(worksheet, address, sheet_name=_safe_str(worksheet.Name))
        return _read_format_conditions(range_obj.FormatConditions)

    def set_conditional_formatting(
        self,
        wb: Any,
        sheet: str,
        address: str,
        spec: ConditionalFormatSpec,
        mode: str = "add",
    ) -> ConditionalFormatMutationResult:
        """Add, modify or delete one supported conditional-format rule.

        ``add`` appends a rule (``FormatConditions.Add``), ``modify`` replaces the
        rule identified by ``spec.priority`` (the rule type is never silently
        converted: a type mismatch is rejected) and ``delete`` removes that rule.
        ``spec.priority`` is the only stable rule identity and is mandatory for
        ``modify`` / ``delete``. Priorities are sheet-wide and unique: assigning a
        ``Priority`` renumbers the other rules (reported through
        ``applied_priorities`` / ``notes``), and deleting a rule leaves a gap
        rather than renumbering. The postcondition re-reads the whole collection;
        a rule formula is compared after normalizing Excel's leading ``=`` so an
        operand written as ``10`` is not reported as a mismatch for Excel's ``=10``.
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)

        validate_conditional_format_spec(spec, mode)
        _require_unprotected(worksheet)

        conditions = range_obj.FormatConditions
        before = _read_format_conditions(conditions)

        applied: list[int] = []
        if mode == "delete":
            _delete_cf_rule(conditions, before, spec)
        else:
            applied = _apply_cf_rule(conditions, spec, mode, before)

        after = _read_format_conditions(conditions)
        verified = _cf_postcondition(before, after, spec, mode, applied)
        unsupported = _unsupported_cf_types(after)
        notes: list[str] = []
        if mode in ("add", "modify") and spec.priority is not None:
            notes.append(
                "Conditional-format priorities are sheet-wide: assigning a priority "
                "renumbers the other rules on the sheet."
            )
        if mode == "delete":
            notes.append(
                "Deleting a conditional-format rule does not renumber the remaining "
                "rules; priorities may contain gaps."
            )
        if not verified:
            notes.append(
                "Postcondition mismatch: the re-read rule collection does not match "
                "the requested operation."
            )
        return ConditionalFormatMutationResult(
            action=mode,
            range=resolved,
            rules=after,
            applied_priorities=applied,
            unsupported=unsupported,
            verified=verified,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Structural operations (copy / insert / delete / merge / unmerge)
    # ------------------------------------------------------------------
    # Every structural operation is **explicit** (destination / shift / target),
    # **guarded** (protected sheet, merged cells, spill ranges, ListObject
    # intersection) and **postcondition-verified** (resulting address, dimensions
    # and table metadata). No clipboard, no ``Select`` / ``Activate``, no UI.

    def copy_range(
        self,
        wb: Any,
        sheet: str,
        source: str,
        destination_sheet: str,
        destination: str,
    ) -> StructuralMutationResult:
        """Copy ``source`` to an explicit ``destination`` in the same workbook.

        The destination is mandatory (a clipboard copy is never performed). The
        copy is rejected when source and destination overlap (ambiguous), when a
        multi-cell destination does not match the source shape, when the
        destination intersects merged cells, when the destination sheet is
        protected, or when either endpoint intersects an Excel Table (ListObject).
        The postcondition re-reads the pasted block and verifies that its values
        and dimensions match the source.
        """

        source_ws = _require_sheet(wb, sheet)
        source_name = _safe_str(source_ws.Name)
        source_range = _resolve_range(source_ws, source, sheet_name=source_name)
        source_addr = _describe_range(source_range, sheet=source_name)
        _validate_target_cells(source_addr.rows * source_addr.cols)

        destination_ws = _require_sheet(wb, destination_sheet)
        destination_name = _safe_str(destination_ws.Name)
        destination_range = _resolve_range(destination_ws, destination, sheet_name=destination_name)
        destination_addr = _describe_range(destination_range, sheet=destination_name)
        _validate_target_cells(destination_addr.rows * destination_addr.cols)

        if source_name == destination_name and _ranges_intersect(source_range, destination_range):
            raise InvalidRangeError(
                "The copy source and destination overlap; an unambiguous copy needs a "
                "disjoint destination.",
                context={
                    "source": source_addr.address,
                    "destination": destination_addr.address,
                    "sheet": source_name,
                },
            )

        anchor_only = destination_addr.rows == 1 and destination_addr.cols == 1
        if not anchor_only and (
            destination_addr.rows != source_addr.rows or destination_addr.cols != source_addr.cols
        ):
            raise InvalidRangeError(
                "The copy destination shape must match the source (or be a single anchor cell).",
                context={
                    "source_rows": source_addr.rows,
                    "source_cols": source_addr.cols,
                    "destination_rows": destination_addr.rows,
                    "destination_cols": destination_addr.cols,
                },
            )

        _require_unprotected(destination_ws)
        _guard_structural_table_intersection(wb, source_ws, source_range, "copy source")
        _guard_structural_table_intersection(
            wb, destination_ws, destination_range, "copy destination"
        )

        merged_areas = _merged_areas(destination_range)
        if merged_areas:
            raise InvalidRangeError(
                "The copy destination intersects merged cells; range_copy has no "
                "allow_merged opt-in, so the copy is rejected.",
                context={"merged_areas": merged_areas, "destination": destination_addr.address},
            )

        source_values = self.read_values(wb, source_name, source)
        referenced = _referenced_formulas(source_ws, source_addr)
        preconditions: dict[str, Any] = {
            "source": {"sheet": source_name, "address": source_addr.address},
            "destination": {"sheet": destination_name, "address": destination_addr.address},
            "destination_anchor_only": anchor_only,
            "merged_areas": merged_areas,
            "referenced": referenced["referenced"],
            "referencing_cells": referenced["cells"],
            "reference_scan": referenced["note"],
        }

        try:
            source_range.Copy(destination_range)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                "Excel rejected the copy operation.",
                context={
                    "source": source_addr.address,
                    "destination": destination_addr.address,
                    "error": str(exc),
                },
            ) from exc

        post_address = _expanded_destination(destination_range, source_addr.rows, source_addr.cols)
        post = self.read_values(wb, destination_name, post_address)
        verified = _range_values_equal(source_values, post)

        notes: list[str] = []
        if not verified:
            notes.append(
                "Postcondition mismatch: the pasted values/dimensions differ from the source."
            )
        return StructuralMutationResult(
            action="copy",
            source=source_addr,
            destination=destination_addr,
            range_before=source_addr,
            range_after=post.range,
            affected_addresses=[source_addr.address, post.range.address],
            table_metadata_unchanged=True,
            verified=verified,
            preconditions=preconditions,
            notes=notes,
        )

    def insert_range(
        self,
        wb: Any,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        """Insert cells at ``address`` and shift them in an explicit direction.

        ``shift`` is mandatory (``"right"`` / ``"down"``); an omitted or unknown
        shift is rejected and Excel is never allowed to infer it from the range
        shape. The insert is rejected when the target is protected, intersects a
        ListObject, lies inside a dynamic-array spill range, or intersects merged
        cells unless ``allow_merged`` is set.
        """

        return self._shift_cells(
            wb,
            sheet,
            address,
            shift,
            allowed_shifts=_INSERT_SHIFT_VALUES,
            operation="insert",
            allow_merged=allow_merged,
            mutate=lambda range_obj, shift_value: range_obj.Insert(
                Shift=shift_value, CopyOrigin=_XL_FORMAT_FROM_LEFT_OR_ABOVE
            ),
        )

    def delete_range(
        self,
        wb: Any,
        sheet: str,
        address: str,
        shift: str,
        *,
        allow_merged: bool = False,
    ) -> StructuralMutationResult:
        """Delete cells at ``address`` and shift the remainder explicitly.

        ``shift`` is mandatory (``"left"`` / ``"up"``); an omitted or unknown
        shift is rejected and Excel is never allowed to infer it from the range
        shape. The delete is rejected when the target is protected, intersects a
        ListObject, lies inside a dynamic-array spill range, or intersects merged
        cells unless ``allow_merged`` is set.
        """

        return self._shift_cells(
            wb,
            sheet,
            address,
            shift,
            allowed_shifts=_DELETE_SHIFT_VALUES,
            operation="delete",
            allow_merged=allow_merged,
            mutate=lambda range_obj, shift_value: range_obj.Delete(Shift=shift_value),
        )

    def merge_range(
        self,
        wb: Any,
        sheet: str,
        address: str,
        *,
        across: bool = False,
        allow_data_loss: bool = False,
    ) -> StructuralMutationResult:
        """Merge the cells of ``address`` into one (one per row when ``across``).

        Merging discards every non-anchor value. The preflight reports the
        affected ``cells_with_data``; the merge is blocked with ``INVALID_RANGE``
        unless the caller passes ``allow_data_loss=True`` after reviewing that
        preview. The merge is also rejected when the target is protected,
        intersects a ListObject or is a dynamic-array spill range.
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)
        _validate_target_cells(resolved.rows * resolved.cols)
        _require_unprotected(worksheet)
        _guard_structural_table_intersection(wb, worksheet, range_obj, "merge")

        spill = _spill_state(range_obj)
        if spill["spilling"] or spill["non_anchor"]:
            raise InvalidRangeError(
                "The merge target is a dynamic-array spill range; merging would break the spill.",
                context={"address": resolved.address, "spill": spill},
            )

        grid = _read_value_grid(range_obj)
        anchor_row = _safe_int(_read_attr(range_obj, "Row"))
        anchor_col = _safe_int(_read_attr(range_obj, "Column"))
        cells_with_data = _non_anchor_addresses(
            grid, across=across, top=anchor_row, left=anchor_col
        )
        existing_merges = _merged_areas(range_obj)
        referenced = _referenced_formulas(worksheet, resolved)
        preconditions: dict[str, Any] = {
            "across": across,
            "cells_with_data": cells_with_data,
            "merged_areas": existing_merges,
            "referenced": referenced["referenced"],
            "referencing_cells": referenced["cells"],
            "reference_scan": referenced["note"],
            "spill": spill,
        }

        if cells_with_data and not allow_data_loss:
            raise InvalidRangeError(
                "Merging discards non-anchor data; review the preview and pass "
                "allow_data_loss=true to confirm.",
                context={
                    "address": resolved.address,
                    "anchor": _cell_address(anchor_row, anchor_col),
                    "cells_with_data": cells_with_data,
                },
            )

        anchor_before = _read_attr(range_obj.Cells(1, 1), "Value2")
        try:
            # ``Range.Merge`` over non-anchor data raises Excel's own
            # "keep only the upper-left value?" alert when DisplayAlerts is on.
            # The merge is already explicitly approved by the caller
            # (``allow_data_loss``), so the alert is suppressed for this one call
            # and the prior setting is restored in ``finally`` — no UI is shown.
            with _alerts_suppressed(_read_attr(wb, "Application")):
                range_obj.Merge(Across=bool(across))
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                "Excel rejected the merge operation.",
                context={"address": resolved.address, "error": str(exc)},
            ) from exc

        merged_after = _merged_areas(range_obj)
        anchor_after = _read_attr(range_obj.Cells(1, 1), "Value2")
        verified = bool(merged_after) and anchor_after == anchor_before

        notes: list[str] = []
        if cells_with_data:
            notes.append(
                f"Merged with allow_data_loss=true; {len(cells_with_data)} non-anchor "
                "cell(s) holding data were discarded."
            )
        if not verified:
            notes.append("Postcondition mismatch: the merged geometry/value was not confirmed.")
        return StructuralMutationResult(
            action="merge",
            source=resolved,
            range_before=resolved,
            range_after=_describe_range(range_obj, sheet=sheet_name),
            affected_addresses=cells_with_data,
            table_metadata_unchanged=True,
            verified=verified,
            preconditions=preconditions,
            notes=notes,
        )

    def unmerge_range(self, wb: Any, sheet: str, address: str) -> StructuralMutationResult:
        """Unmerge every merged area that intersects ``address``.

        The unmerge is rejected when no merged area intersects the target (there
        is nothing to unmerge), when the sheet is protected, or when the target
        intersects a ListObject. The postcondition confirms that no merged area
        remains inside the target.
        """

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)
        _validate_target_cells(resolved.rows * resolved.cols)
        _require_unprotected(worksheet)
        _guard_structural_table_intersection(wb, worksheet, range_obj, "unmerge")

        merged_areas = _merged_areas(range_obj)
        if not merged_areas:
            raise InvalidRangeError(
                "No merged area intersects the target; there is nothing to unmerge.",
                context={"address": resolved.address},
            )

        preconditions: dict[str, Any] = {"merged_areas": merged_areas}
        try:
            range_obj.UnMerge()
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                "Excel rejected the unmerge operation.",
                context={"address": resolved.address, "error": str(exc)},
            ) from exc

        remaining = _merged_areas(range_obj)
        verified = not remaining
        notes: list[str] = []
        if not verified:
            notes.append(
                "Postcondition mismatch: merged areas remain inside the target "
                f"({', '.join(remaining)})."
            )
        return StructuralMutationResult(
            action="unmerge",
            source=resolved,
            range_before=resolved,
            range_after=_describe_range(range_obj, sheet=sheet_name),
            affected_addresses=list(merged_areas),
            table_metadata_unchanged=True,
            verified=verified,
            preconditions=preconditions,
            notes=notes,
        )

    def _shift_cells(
        self,
        wb: Any,
        sheet: str,
        address: str,
        shift: str,
        *,
        allowed_shifts: dict[str, int],
        operation: str,
        allow_merged: bool,
        mutate: Callable[[Any, int], Any],
    ) -> StructuralMutationResult:
        """Shared insert/delete path: validate, preflight, mutate, verify."""

        worksheet = _require_sheet(wb, sheet)
        sheet_name = _safe_str(worksheet.Name)
        range_obj = _resolve_range(worksheet, address, sheet_name=sheet_name)
        resolved = _describe_range(range_obj, sheet=sheet_name)
        _validate_target_cells(resolved.rows * resolved.cols)
        shift_value = _resolve_shift(shift, allowed_shifts, operation)
        _require_unprotected(worksheet)
        _guard_structural_table_intersection(wb, worksheet, range_obj, operation)

        merged_areas = _merged_areas(range_obj)
        spill = _spill_state(range_obj)
        grid = _read_value_grid(range_obj)
        non_empty = _count_non_empty(grid)
        target_top = _safe_int(_read_attr(range_obj, "Row"))
        target_left = _safe_int(_read_attr(range_obj, "Column"))
        referenced = _referenced_formulas(worksheet, resolved)
        preconditions: dict[str, Any] = {
            "shift": shift.strip().lower(),
            "merged_areas": merged_areas,
            "non_empty_cells": non_empty,
            "referenced": referenced["referenced"],
            "referencing_cells": referenced["cells"],
            "reference_scan": referenced["note"],
            "spill": spill,
        }

        if merged_areas and not allow_merged:
            raise InvalidRangeError(
                f"The {operation} target intersects merged cells; pass allow_merged=true "
                "to confirm the merge-geometry change.",
                context={"address": resolved.address, "merged_areas": merged_areas},
            )
        if spill["non_anchor"]:
            raise InvalidRangeError(
                f"The {operation} target lies inside a dynamic-array spill range.",
                context={"address": resolved.address, "spill": spill},
            )

        try:
            mutate(range_obj, shift_value)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                f"Excel rejected the {operation} operation.",
                context={
                    "address": resolved.address,
                    "shift": shift.strip().lower(),
                    "error": str(exc),
                },
            ) from exc

        # ``Range.Insert`` re-targets the Range object at the shifted cells and
        # ``Range.Delete`` invalidates it, so the postcondition re-resolves the
        # original target address from the worksheet instead of trusting the
        # mutated object.
        target_range = _resolve_range(worksheet, resolved.address, sheet_name=sheet_name)
        post = _describe_range(target_range, sheet=sheet_name)
        post_values = self.read_values(wb, sheet, resolved.address)
        shift_name = shift.strip().lower()
        if operation == "insert":
            if shift_name == "right":
                shifted_values = _read_block(
                    worksheet,
                    target_top,
                    target_left + resolved.cols,
                    resolved.rows,
                    resolved.cols,
                )
            else:
                shifted_values = _read_block(
                    worksheet,
                    target_top + resolved.rows,
                    target_left,
                    resolved.rows,
                    resolved.cols,
                )
            verified = _insert_verified(
                resolved, post, post_values, shifted_values=shifted_values, pre_grid=grid
            )
        else:
            verified = _delete_verified(resolved, post, post_values, before_non_empty=non_empty)

        notes: list[str] = []
        if not verified:
            notes.append(
                f"Postcondition mismatch: the {operation} did not produce the expected "
                "address/dimensions/content."
            )
        return StructuralMutationResult(
            action=operation,
            source=resolved,
            shift=shift.strip().lower(),
            range_before=resolved,
            range_after=post,
            affected_addresses=[resolved.address],
            table_metadata_unchanged=True,
            verified=verified,
            preconditions=preconditions,
            notes=notes,
        )


# ---------------------------------------------------------------------------
# Formatting helpers (task_044)
# ---------------------------------------------------------------------------
def validate_format_patch(patch: RangeFormatPatch) -> None:
    """Validate a :class:`RangeFormatPatch` before any COM mutation (pure).

    Rejects a bad enum name, an out-of-range numeric value, an over-long string,
    or an explicit ``None`` on a field that has no automatic state. Only the
    fields the caller set (``model_fields_set``) are validated.
    """

    fields = patch.model_fields_set

    if "number_format" in fields and patch.number_format is not None:
        _require_format_string(patch.number_format, "number_format")

    if "font_name" in fields:
        _reject_null(patch.font_name, "font_name")
        if patch.font_name is not None:
            _require_format_string(patch.font_name, "font_name")

    if "font_bold" in fields:
        _reject_null(patch.font_bold, "font_bold")
        if patch.font_bold is not None:
            _require_bool(patch.font_bold, "font_bold")

    if "font_italic" in fields:
        _reject_null(patch.font_italic, "font_italic")
        if patch.font_italic is not None:
            _require_bool(patch.font_italic, "font_italic")

    if "font_size" in fields:
        _reject_null(patch.font_size, "font_size")
        if patch.font_size is not None:
            size = patch.font_size
            if isinstance(size, bool) or not isinstance(size, (int, float)):
                raise InvalidRangeError(
                    "Argument 'font_size' must be a number.", context={"font_size": size}
                )
            if size < _MIN_FONT_SIZE or size > _MAX_FONT_SIZE:
                raise InvalidRangeError(
                    f"Argument 'font_size' must be within {_MIN_FONT_SIZE}..{_MAX_FONT_SIZE}.",
                    context={"font_size": size},
                )

    if "font_color" in fields and patch.font_color is not None:
        _require_color(patch.font_color, "font_color")

    if "fill_color" in fields and patch.fill_color is not None:
        _require_color(patch.fill_color, "fill_color")

    if "fill_pattern" in fields:
        _reject_null(patch.fill_pattern, "fill_pattern")
        if patch.fill_pattern is not None:
            _require_enum(patch.fill_pattern, _FILL_PATTERN_VALUES, "fill_pattern")

    if "alignment_horizontal" in fields:
        _reject_null(patch.alignment_horizontal, "alignment_horizontal")
        if patch.alignment_horizontal is not None:
            _require_enum(
                patch.alignment_horizontal,
                _HORIZONTAL_ALIGNMENT_VALUES,
                "alignment_horizontal",
            )

    if "alignment_vertical" in fields:
        _reject_null(patch.alignment_vertical, "alignment_vertical")
        if patch.alignment_vertical is not None:
            _require_enum(
                patch.alignment_vertical, _VERTICAL_ALIGNMENT_VALUES, "alignment_vertical"
            )

    if "alignment_wrap_text" in fields:
        _reject_null(patch.alignment_wrap_text, "alignment_wrap_text")
        if patch.alignment_wrap_text is not None:
            _require_bool(patch.alignment_wrap_text, "alignment_wrap_text")


def _require_format_string(value: str, field: str) -> None:
    """Reject an empty or over-long format/name string."""

    if not isinstance(value, str) or not value.strip():
        raise InvalidRangeError(
            f"Argument {field!r} must be a non-empty string.", context={field: value}
        )
    if len(value) > _MAX_FORMAT_STRING_LENGTH:
        raise InvalidRangeError(
            f"Argument {field!r} exceeds the {_MAX_FORMAT_STRING_LENGTH}-character limit.",
            context={field: value},
        )


def _reject_null(value: object, field: str) -> None:
    """Reject an explicit ``None`` on a field that has no automatic state.

    Only the clearable fields (``number_format``, ``font_color``,
    ``fill_color``) accept ``None`` to mean "clear to automatic"; every other
    field must omit the value to leave it untouched (task_038 §4.3).
    """

    if value is None and field not in _CLEARABLE_FORMAT_FIELDS:
        raise InvalidRangeError(
            f"Argument {field!r} does not accept null; omit it to leave the property unchanged.",
            context={field: None},
        )


def _require_bool(value: object, field: str) -> None:
    """Reject a non-boolean flag."""

    if not isinstance(value, bool):
        raise InvalidRangeError(f"Argument {field!r} must be a boolean.", context={field: value})


def _require_color(value: object, field: str) -> None:
    """Reject a non-integer or out-of-range RGB color."""

    if isinstance(value, bool) or not isinstance(value, int):
        raise InvalidRangeError(
            f"Argument {field!r} must be an integer RGB color (0x000000..0x{_MAX_RGB_COLOR:06X}).",
            context={field: value},
        )
    if value < 0 or value > _MAX_RGB_COLOR:
        raise InvalidRangeError(
            f"Argument {field!r} must be within 0..0x{_MAX_RGB_COLOR:06X}.",
            context={field: value},
        )


def _require_enum(value: object, allowed: Collection[str], field: str) -> None:
    """Reject an enum name outside the supported set."""

    if not isinstance(value, str) or value not in allowed:
        raise InvalidRangeError(
            f"Argument {field!r} is not a supported value; expected one of {sorted(allowed)}.",
            context={field: value, "supported": sorted(allowed)},
        )


def _apply_format(range_obj: Any, patch: RangeFormatPatch) -> None:
    """Write only the fields the caller set on ``range_obj`` (STA-local).

    A present ``None`` clears to automatic only for the clearable fields
    (``number_format``, ``font_color``, ``fill_color``); every other field was
    rejected by :func:`validate_format_patch`.
    """

    fields = patch.model_fields_set

    if "number_format" in fields:
        _set_com(
            range_obj,
            "NumberFormat",
            "General" if patch.number_format is None else patch.number_format,
        )

    if "font_name" in fields:
        _set_com(range_obj.Font, "Name", patch.font_name)

    if "font_bold" in fields:
        _set_com(range_obj.Font, "Bold", patch.font_bold)

    if "font_italic" in fields:
        _set_com(range_obj.Font, "Italic", patch.font_italic)

    if "font_size" in fields:
        _set_com(range_obj.Font, "Size", patch.font_size)

    if "font_color" in fields:
        if patch.font_color is None:
            _set_com(range_obj.Font, "ColorIndex", _XL_COLOR_INDEX_AUTOMATIC)
        else:
            _set_com(range_obj.Font, "Color", patch.font_color)

    if "fill_color" in fields:
        if patch.fill_color is None:
            _set_com(range_obj.Interior, "ColorIndex", _XL_COLOR_INDEX_NONE)
        else:
            _set_com(range_obj.Interior, "Color", patch.fill_color)

    if "fill_pattern" in fields:
        _set_com(range_obj.Interior, "Pattern", _FILL_PATTERN_VALUES[str(patch.fill_pattern)])

    if "alignment_horizontal" in fields:
        _set_com(
            range_obj,
            "HorizontalAlignment",
            _HORIZONTAL_ALIGNMENT_VALUES[str(patch.alignment_horizontal)],
        )

    if "alignment_vertical" in fields:
        _set_com(
            range_obj,
            "VerticalAlignment",
            _VERTICAL_ALIGNMENT_VALUES[str(patch.alignment_vertical)],
        )

    if "alignment_wrap_text" in fields:
        _set_com(range_obj, "WrapText", patch.alignment_wrap_text)


def _set_com(obj: Any, attribute: str, value: Any) -> None:
    """Assign one COM property, converting a rejection into a typed error."""

    try:
        setattr(obj, attribute, value)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            f"Excel rejected setting {attribute!r}.",
            context={"attribute": attribute, "error": str(exc)},
        ) from exc


def _format_matches_patch(range_obj: Any, patch: RangeFormatPatch, post: FormatInfo) -> bool:
    """Return whether every targeted property was applied (postcondition)."""

    fields = patch.model_fields_set

    if "number_format" in fields:
        expected = "General" if patch.number_format is None else patch.number_format
        if post.number_format != expected:
            return False

    if "font_name" in fields and post.font.name != patch.font_name:
        return False

    if "font_bold" in fields and post.font.bold != patch.font_bold:
        return False

    if "font_italic" in fields and post.font.italic != patch.font_italic:
        return False

    if "font_size" in fields:
        if patch.font_size is None or post.font.size is None:
            return False
        if abs(float(post.font.size) - float(patch.font_size)) > 1e-6:
            return False

    if "font_color" in fields:
        if patch.font_color is None:
            if _read_int_or_none(range_obj.Font, "ColorIndex") != _XL_COLOR_INDEX_AUTOMATIC:
                return False
        elif post.font.color != patch.font_color:
            return False

    if "fill_color" in fields:
        if patch.fill_color is None:
            if _read_int_or_none(range_obj.Interior, "ColorIndex") != _XL_COLOR_INDEX_NONE:
                return False
        elif post.fill.color != patch.fill_color:
            return False

    if "fill_pattern" in fields and post.fill.pattern != patch.fill_pattern:
        return False

    if "alignment_horizontal" in fields and post.alignment.horizontal != patch.alignment_horizontal:
        return False

    if "alignment_vertical" in fields and post.alignment.vertical != patch.alignment_vertical:
        return False

    if "alignment_wrap_text" in fields:
        return post.alignment.wrap_text == patch.alignment_wrap_text

    return True


def _read_number_format(range_obj: Any, mixed: list[str]) -> str | None:
    """Read ``Range.NumberFormat`` (``None`` = mixed across the range)."""

    value = _read_attr(range_obj, "NumberFormat")
    if value is None:
        mixed.append("number_format")
        return None
    return str(value)


def _read_font(range_obj: Any, mixed: list[str]) -> FontFormat:
    """Read the bounded ``Range.Font`` subset (tri-state)."""

    font = range_obj.Font
    name = _read_str_prop(font, "Name", "font.name", mixed)
    bold = _read_bool_prop(font, "Bold", "font.bold", mixed)
    italic = _read_bool_prop(font, "Italic", "font.italic", mixed)
    size = _read_float_prop(font, "Size", "font.size", mixed)
    color = _read_int_prop(font, "Color", "font.color", mixed)
    return FontFormat(name=name, bold=bold, italic=italic, size=size, color=color)


def _read_fill(range_obj: Any, mixed: list[str]) -> FillFormat:
    """Read the bounded ``Range.Interior`` subset (tri-state, no-fill aware)."""

    interior = range_obj.Interior

    color_index = _read_int_or_none(interior, "ColorIndex")
    if color_index is not None and color_index == _XL_COLOR_INDEX_NONE:
        color: int | None = None  # no fill, not mixed
    else:
        raw_color = _read_int_or_none(interior, "Color")
        if raw_color is None:
            mixed.append("fill.color")
            color = None
        else:
            color = raw_color

    pattern_value = _read_int_or_none(interior, "Pattern")
    if pattern_value is None:
        mixed.append("fill.pattern")
        pattern: str | None = None
    else:
        pattern = _FILL_PATTERN_NAMES.get(pattern_value, str(pattern_value))

    return FillFormat(color=color, pattern=pattern)


def _read_alignment(range_obj: Any, mixed: list[str]) -> AlignmentFormat:
    """Read the bounded ``Range`` alignment subset (tri-state)."""

    horizontal = _read_enum_prop(
        range_obj, "HorizontalAlignment", _HORIZONTAL_ALIGNMENT_NAMES, "alignment.horizontal", mixed
    )
    vertical = _read_enum_prop(
        range_obj, "VerticalAlignment", _VERTICAL_ALIGNMENT_NAMES, "alignment.vertical", mixed
    )
    wrap_text = _read_bool_prop(range_obj, "WrapText", "alignment.wrap_text", mixed)
    return AlignmentFormat(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)


def _read_attr(obj: Any, attribute: str) -> Any:
    """Read a COM property; return ``None`` for ``Null`` or any failure."""

    try:
        value = getattr(obj, attribute)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return value


def _read_int_or_none(obj: Any, attribute: str) -> int | None:
    """Read an integer COM property; ``None`` for ``Null`` or failure."""

    value = _read_attr(obj, attribute)
    if value is None:
        return None
    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _read_str_prop(obj: Any, attribute: str, path: str, mixed: list[str]) -> str | None:
    """Read a string property; ``None`` (mixed) when Excel returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        mixed.append(path)
        return None
    return str(value)


def _read_bool_prop(obj: Any, attribute: str, path: str, mixed: list[str]) -> bool | None:
    """Read a boolean property; ``None`` (mixed) when Excel returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        mixed.append(path)
        return None
    return bool(value)


def _read_float_prop(obj: Any, attribute: str, path: str, mixed: list[str]) -> float | None:
    """Read a float property; ``None`` (mixed) when Excel returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        mixed.append(path)
        return None
    try:
        return float(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        mixed.append(path)
        return None


def _read_int_prop(obj: Any, attribute: str, path: str, mixed: list[str]) -> int | None:
    """Read an integer property; ``None`` (mixed) when Excel returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        mixed.append(path)
        return None
    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        mixed.append(path)
        return None


def _read_enum_prop(
    obj: Any, attribute: str, name_map: dict[int, str], path: str, mixed: list[str]
) -> str | None:
    """Read an enum property; ``None`` (mixed) when Excel returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        mixed.append(path)
        return None
    try:
        int_value = int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        mixed.append(path)
        return None
    return name_map.get(int_value, str(int_value))


# ---------------------------------------------------------------------------
# Data validation helpers (task_046, task_038 §6.1)
# ---------------------------------------------------------------------------
def validate_validation_spec(spec: ValidationSpec) -> None:
    """Validate a ``range_set_validation`` payload before any COM mutation (pure).

    Rejects a rule type without rule semantics (``none`` / ``input_only``), a
    missing or over-long ``Formula1`` / ``Formula2``, a ``Formula2`` on a
    single-formula type, a missing second operand for ``between`` /
    ``not_between``, an unknown operator / alert style and an over-long message.
    Formulas are never evaluated — only their length and presence are checked.
    """

    if spec.type in (DataValidationType.NONE, DataValidationType.INPUT_ONLY):
        raise InvalidRangeError(
            "A data-validation rule requires a rule-bearing type; "
            f"{spec.type.value!r} carries no rule semantics.",
            context={"type": spec.type.value},
        )

    _require_rule_formula(spec.formula1, "formula1", required=True)
    if spec.formula2 is not None:
        _require_rule_formula(spec.formula2, "formula2", required=False)

    if spec.type in _VALIDATION_SINGLE_FORMULA_TYPES:
        if spec.formula2:
            raise InvalidRangeError(
                f"Validation type {spec.type.value!r} uses only Formula1; Formula2 is ignored.",
                context={"type": spec.type.value, "formula2": spec.formula2},
            )
    elif spec.operator in _DUAL_FORMULA_OPERATORS and not spec.formula2:
        raise InvalidRangeError(
            f"Operator {spec.operator.value!r} requires both Formula1 and Formula2.",
            context={"operator": spec.operator.value},
        )

    if spec.operator is not None:
        _require_enum(spec.operator, _OPERATOR_VALUES, "operator")
    if spec.alert_style is not None:
        _require_enum(spec.alert_style, _ALERT_STYLE_VALUES, "alert_style")

    for name in ("error_title", "error_message", "input_title", "input_message"):
        value = getattr(spec, name)
        if value is not None:
            _require_validation_text(value, name)

    for name in ("ignore_blank", "in_cell_dropdown", "show_error", "show_input"):
        value = getattr(spec, name)
        if value is not None and not isinstance(value, bool):
            raise InvalidRangeError(f"Argument {name!r} must be a boolean.", context={name: value})


def validate_conditional_format_spec(spec: ConditionalFormatSpec, mode: str) -> None:
    """Validate a ``range_set_conditional_formatting`` payload (pure, pre-COM).

    Rejects an unknown ``mode``, an unsupported rule type (``xlColorScale`` /
    ``xlDataBar`` / ``xlIconSet`` and unknown types), a missing rule identity for
    ``modify`` / ``delete``, a missing operator/formula for a type that requires
    it, an over-long formula and an out-of-range colour/font value.
    """

    if mode not in CF_MODES:
        raise InvalidRangeError(
            f"Argument 'mode' must be one of {sorted(CF_MODES)}.", context={"mode": mode}
        )

    if spec.type not in _SUPPORTED_CF_TYPES:
        reason = _cf_unsupported_reason(spec.type)
        raise InvalidRangeError(
            f"Conditional-format type {spec.type.value!r} is not supported: {reason}",
            context={"type": spec.type.value, "reason": reason},
        )

    if mode in ("modify", "delete") and spec.priority is None:
        raise InvalidRangeError(
            f"Argument 'priority' is required to {mode} a conditional-format rule; "
            "it is the only stable rule identity.",
            context={"mode": mode},
        )

    if spec.operator is not None:
        _require_enum(spec.operator, _OPERATOR_VALUES, "operator")
    elif spec.type in _CF_OPERATOR_REQUIRED_TYPES:
        raise InvalidRangeError(
            f"Argument 'operator' is required for conditional-format type {spec.type.value!r}.",
            context={"type": spec.type.value},
        )

    if spec.formula1 is not None:
        _require_rule_formula(
            spec.formula1, "formula1", required=spec.type in _CF_FORMULA_REQUIRED_TYPES
        )
    elif spec.type in _CF_FORMULA_REQUIRED_TYPES:
        raise InvalidRangeError(
            f"Argument 'formula1' is required for conditional-format type {spec.type.value!r}.",
            context={"type": spec.type.value},
        )
    if spec.formula2 is not None:
        _require_rule_formula(spec.formula2, "formula2", required=False)

    if spec.interior_color is not None:
        _require_color(spec.interior_color, "interior_color")
    if spec.font is not None:
        _validate_cf_font(spec.font)


def _validate_cf_font(font: FontFormat) -> None:
    """Validate the bounded ``FormatCondition.Font`` subset before any COM write."""

    if font.name is not None:
        _require_format_string(font.name, "font.name")
    if font.bold is not None and not isinstance(font.bold, bool):
        raise InvalidRangeError(
            "Argument 'font.bold' must be a boolean.", context={"font.bold": font.bold}
        )
    if font.italic is not None and not isinstance(font.italic, bool):
        raise InvalidRangeError(
            "Argument 'font.italic' must be a boolean.", context={"font.italic": font.italic}
        )
    if font.size is not None:
        if isinstance(font.size, bool) or not isinstance(font.size, (int, float)):
            raise InvalidRangeError(
                "Argument 'font.size' must be a number.", context={"font.size": font.size}
            )
        if font.size < _MIN_FONT_SIZE or font.size > _MAX_FONT_SIZE:
            raise InvalidRangeError(
                f"Argument 'font.size' must be within {_MIN_FONT_SIZE}..{_MAX_FONT_SIZE}.",
                context={"font.size": font.size},
            )
    if font.color is not None:
        _require_color(font.color, "font.color")


def _require_rule_formula(value: object, field: str, *, required: bool) -> None:
    """Validate one rule operand: presence (optionally) and the 255-char limit."""

    if value is None or (isinstance(value, str) and not value.strip()):
        if required:
            raise InvalidRangeError(
                f"Argument {field!r} is required for this rule.",
                context={field: value},
            )
        return
    if not isinstance(value, str):
        raise InvalidRangeError(f"Argument {field!r} must be a string.", context={field: value})
    if len(value) > MAX_VALIDATION_FORMULA_LENGTH:
        raise InvalidRangeError(
            f"Argument {field!r} exceeds the {MAX_VALIDATION_FORMULA_LENGTH}-character limit.",
            context={field: value, "length": len(value)},
        )


def _require_validation_text(value: str, field: str) -> None:
    """Validate a validation message/title string against Excel's UI limit."""

    if not isinstance(value, str):
        raise InvalidRangeError(f"Argument {field!r} must be a string.", context={field: value})
    if len(value) > MAX_VALIDATION_TEXT_LENGTH:
        raise InvalidRangeError(
            f"Argument {field!r} exceeds the {MAX_VALIDATION_TEXT_LENGTH}-character limit.",
            context={field: value, "length": len(value)},
        )


def _read_validation_rule(range_obj: Any, *, apply_to: str) -> ValidationRule:
    """Read the validation rule of a resolved Range (never raises, never guesses).

    The range-level read is authoritative when Excel answers it (every cell of
    the target shares one rule). Otherwise a bounded cell probe produces the
    representative rule and marks the result ``non_uniform_validation`` (one
    distinct rule, coverage differs) or ``mixed_validation`` (several distinct
    rules); an exhausted probe on a large target is flagged
    ``validation_probe_limited``.
    """

    direct = _read_validation(_read_attr(range_obj, "Validation"))
    if direct is not None:
        return direct.model_copy(update={"apply_to": apply_to})

    rows, cols = _declare(range_obj)
    if rows <= 0 or cols <= 0:
        return ValidationRule(present=False, apply_to=apply_to)

    probed, limited = _probe_validation(range_obj, rows, cols)
    if not probed:
        rule = ValidationRule(present=False, apply_to=apply_to)
        if limited:
            rule.unsupported.append("validation_probe_limited")
        return rule

    signatures = {_validation_signature(item) for item in probed}
    rule = probed[0].model_copy(update={"apply_to": apply_to})
    if len(signatures) > 1:
        rule.unsupported.extend(["mixed_validation", f"distinct_rules={len(signatures)}"])
    else:
        rule.unsupported.append("non_uniform_validation")
    if limited:
        rule.unsupported.append("validation_probe_limited")
    return rule


def _read_validation(validation: Any) -> ValidationRule | None:
    """Read a ``Validation`` object; ``None`` when Excel has no rule to read.

    Excel raises ``1004`` when the range carries no validation (or when the cells
    of the range disagree), so an unreadable ``Type`` is the "no rule" signal and
    is never reported as a fabricated ``none`` rule.
    """

    if validation is None:
        return None
    type_value = _read_int_or_none(validation, "Type")
    if type_value is None:
        return None

    unsupported: list[str] = []
    validation_type = _VALIDATION_TYPE_BY_INT.get(type_value)
    if validation_type is None:
        unsupported.append(f"validation_type={type_value}")
        validation_type = DataValidationType.NONE

    operator_value = _read_int_or_none(validation, "Operator")
    operator = _OPERATOR_BY_INT.get(operator_value) if operator_value is not None else None
    if operator_value is not None and operator is None:
        unsupported.append(f"operator={operator_value}")

    alert_value = _read_int_or_none(validation, "AlertStyle")
    alert_style = _ALERT_STYLE_BY_INT.get(alert_value) if alert_value is not None else None
    if alert_value is not None and alert_style is None:
        unsupported.append(f"alert_style={alert_value}")

    return ValidationRule(
        present=True,
        type=validation_type,
        operator=operator,
        formula1=_read_optional_str(validation, "Formula1"),
        formula2=_read_optional_str(validation, "Formula2"),
        ignore_blank=_read_optional_bool(validation, "IgnoreBlank"),
        in_cell_dropdown=_read_optional_bool(validation, "InCellDropdown"),
        alert_style=alert_style,
        error_message=_read_optional_str(validation, "ErrorMessage"),
        error_title=_read_optional_str(validation, "ErrorTitle"),
        input_message=_read_optional_str(validation, "InputMessage"),
        input_title=_read_optional_str(validation, "InputTitle"),
        show_error=_read_optional_bool(validation, "ShowError"),
        show_input=_read_optional_bool(validation, "ShowInput"),
        unsupported=unsupported,
    )


def _validation_signature(rule: ValidationRule) -> tuple[Any, ...]:
    """Return the comparable identity of a validation rule."""

    return (
        rule.type,
        rule.operator,
        rule.formula1,
        rule.formula2,
        rule.ignore_blank,
        rule.in_cell_dropdown,
        rule.alert_style,
    )


def _validation_probe_coordinates(rows: int, cols: int) -> list[tuple[int, int]]:
    """Return the 1-based ``(row, col)`` coordinates of the bounded probe."""

    total = rows * cols
    if total <= _MAX_VALIDATION_PROBE_CELLS:
        indices = list(range(total))
    else:
        last = total - 1
        count = _MAX_VALIDATION_PROBE_CELLS
        sampled = {0, last}
        sampled.update(round(index * last / (count - 1)) for index in range(count))
        indices = sorted(sampled)
    return [(index // cols + 1, index % cols + 1) for index in indices]


def _probe_validation(range_obj: Any, rows: int, cols: int) -> tuple[list[ValidationRule], bool]:
    """Probe up to ``_MAX_VALIDATION_PROBE_CELLS`` cells for a validation rule."""

    rules: list[ValidationRule] = []
    for row, col in _validation_probe_coordinates(rows, cols):
        cell = range_obj.Cells(row, col)
        rule = _read_validation(_read_attr(cell, "Validation"))
        if rule is not None:
            rules.append(rule)
    return rules, rows * cols > _MAX_VALIDATION_PROBE_CELLS


def _apply_validation(range_obj: Any, spec: ValidationSpec) -> None:
    """Add or modify the range's validation rule (STA-local, one COM path)."""

    validation = range_obj.Validation
    top_left = range_obj.Cells(1, 1)
    exists = (
        _read_validation(validation) is not None
        or _read_validation(_read_attr(top_left, "Validation")) is not None
    )

    type_value = _VALIDATION_TYPE_VALUES[spec.type]
    operator_value = (
        _OPERATOR_VALUES[spec.operator] if spec.operator is not None else XL_COND_BETWEEN
    )
    alert_value = (
        _ALERT_STYLE_VALUES[str(spec.alert_style)]
        if spec.alert_style is not None
        else XL_ALERT_STOP
    )
    arguments = (
        type_value,
        alert_value,
        operator_value,
        spec.formula1 or "",
        spec.formula2 or "",
    )
    try:
        if exists:
            validation.Modify(*arguments)
        else:
            validation.Add(*arguments)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            "Excel rejected the data-validation rule.",
            context={
                "type": spec.type.value,
                "operator": spec.operator.value if spec.operator is not None else None,
                "formula1": spec.formula1,
                "formula2": spec.formula2,
                "existing": exists,
                "error": str(exc),
            },
        ) from exc

    fields = spec.model_fields_set
    if "ignore_blank" in fields and spec.ignore_blank is not None:
        _set_com(validation, "IgnoreBlank", bool(spec.ignore_blank))
    if "in_cell_dropdown" in fields and spec.in_cell_dropdown is not None:
        _set_com(validation, "InCellDropdown", bool(spec.in_cell_dropdown))
    if "show_error" in fields and spec.show_error is not None:
        _set_com(validation, "ShowError", bool(spec.show_error))
    if "show_input" in fields and spec.show_input is not None:
        _set_com(validation, "ShowInput", bool(spec.show_input))
    if "error_title" in fields and spec.error_title is not None:
        _set_com(validation, "ErrorTitle", spec.error_title)
    if "error_message" in fields and spec.error_message is not None:
        _set_com(validation, "ErrorMessage", spec.error_message)
    if "input_title" in fields and spec.input_title is not None:
        _set_com(validation, "InputTitle", spec.input_title)
    if "input_message" in fields and spec.input_message is not None:
        _set_com(validation, "InputMessage", spec.input_message)


def _validation_matches_spec(rule: ValidationRule, spec: ValidationSpec) -> bool:
    """Return whether a re-read validation rule matches the requested spec."""

    if not rule.present or rule.type is not spec.type:
        return False
    if spec.formula1 is not None and rule.formula1 != spec.formula1:
        return False
    if spec.formula2 is not None and (rule.formula2 or "") != spec.formula2:
        return False
    if (
        spec.operator is not None
        and spec.type not in _VALIDATION_SINGLE_FORMULA_TYPES
        and rule.operator is not spec.operator
    ):
        return False

    fields = spec.model_fields_set
    expectations: tuple[tuple[str, Any, Any], ...] = (
        ("ignore_blank", "ignore_blank", spec.ignore_blank),
        ("in_cell_dropdown", "in_cell_dropdown", spec.in_cell_dropdown),
        ("alert_style", "alert_style", spec.alert_style),
        ("error_message", "error_message", spec.error_message),
        ("error_title", "error_title", spec.error_title),
        ("input_message", "input_message", spec.input_message),
        ("input_title", "input_title", spec.input_title),
        ("show_error", "show_error", spec.show_error),
        ("show_input", "show_input", spec.show_input),
    )
    for field, attribute, expected in expectations:
        if field in fields and expected is not None and getattr(rule, attribute) != expected:
            return False
    return True


# ---------------------------------------------------------------------------
# Conditional-formatting helpers (task_046, task_038 §6.3)
# ---------------------------------------------------------------------------
def _classify_cf_type(type_value: int | None) -> tuple[ConditionalFormatType, bool, str | None]:
    """Return ``(type, supported, unsupported_reason)`` for a raw rule type."""

    if type_value is None:
        return ConditionalFormatType.UNSUPPORTED, False, "the rule type could not be read."
    cf_type = _CF_TYPE_BY_INT.get(type_value)
    if cf_type is None:
        return (
            ConditionalFormatType.UNSUPPORTED,
            False,
            f"unknown conditional-format type {type_value}.",
        )
    if cf_type in _SUPPORTED_CF_TYPES:
        return cf_type, True, None
    reason = _CF_UNSUPPORTED_REASONS.get(type_value)
    if reason is None:
        reason = f"{cf_type.value!r} is not a supported plain FormatCondition type."
    return cf_type, False, reason


def _cf_unsupported_reason(cf_type: ConditionalFormatType) -> str:
    """Return the explicit unsupported reason for a rule type."""

    for type_value, reason in _CF_UNSUPPORTED_REASONS.items():
        if _CF_TYPE_BY_INT.get(type_value) is cf_type:
            return reason
    return f"{cf_type.value!r} is not a supported plain FormatCondition type."


def _read_format_conditions(conditions: Any) -> list[ConditionalFormatRule]:
    """Read every rule of a ``FormatConditions`` collection in priority order."""

    rules: list[ConditionalFormatRule] = []
    count = _read_int_or_none(conditions, "Count") or 0
    for index in range(1, count + 1):
        try:
            rule_obj = conditions.Item(index)
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        rules.append(_read_format_condition(rule_obj))
    rules.sort(key=lambda item: item.priority)
    return rules


def _read_format_condition(rule_obj: Any) -> ConditionalFormatRule:
    """Assemble one :class:`ConditionalFormatRule` from a COM rule object."""

    type_value = _read_int_or_none(rule_obj, "Type")
    cf_type, supported, reason = _classify_cf_type(type_value)
    priority = _read_int_or_none(rule_obj, "Priority")
    operator_value = _read_int_or_none(rule_obj, "Operator")
    return ConditionalFormatRule(
        priority=priority if priority is not None and priority >= 1 else 1,
        type=cf_type,
        operator=_OPERATOR_BY_INT.get(operator_value) if operator_value is not None else None,
        formula1=_read_optional_str(rule_obj, "Formula1"),
        formula2=_read_optional_str(rule_obj, "Formula2"),
        stop_if_true=_read_optional_bool(rule_obj, "StopIfTrue"),
        applies_to=_read_applies_to(rule_obj),
        interior_color=_read_cf_interior_color(rule_obj) if supported else None,
        font=_read_cf_font(rule_obj) if supported else None,
        supported=supported,
        unsupported_reason=reason,
    )


def _read_applies_to(rule_obj: Any) -> str | None:
    """Read ``FormatCondition.AppliesTo.Address`` (``None`` when unavailable)."""

    applies_to = _read_attr(rule_obj, "AppliesTo")
    if applies_to is None:
        return None
    return _safe_str(_read_attr(applies_to, "Address")) or None


def _read_cf_interior_color(rule_obj: Any) -> int | None:
    """Read the bounded ``FormatCondition.Interior`` colour.

    ``Interior.ColorIndex`` is ``Null`` for a rule without an explicit fill, so
    the colour is reported as ``None`` instead of Excel's default black.
    """

    interior = _read_attr(rule_obj, "Interior")
    if interior is None:
        return None
    color_index = _read_int_or_none(interior, "ColorIndex")
    if color_index is None:
        return None
    return _read_int_or_none(interior, "Color")


def _read_cf_font(rule_obj: Any) -> FontFormat | None:
    """Read the bounded ``FormatCondition.Font`` subset (``None`` when unset).

    Excel reports ``Name = "None"``, ``Size = Null`` and ``Color = Null`` for a
    rule without a font override. A bold/italic-only rule has no name/size/colour
    but is still reported (the non-default flags are the evidence of a font
    override), so no rule detail is silently dropped.
    """

    font_obj = _read_attr(rule_obj, "Font")
    if font_obj is None:
        return None
    name = _read_optional_str(font_obj, "Name")
    if name == "None":  # Excel reports the string "None" for an unset font name
        name = None
    size = _read_optional_float(font_obj, "Size")
    color = _read_optional_int(font_obj, "Color")
    bold = _read_optional_bool(font_obj, "Bold")
    italic = _read_optional_bool(font_obj, "Italic")
    if name is None and size is None and color is None and not bold and not italic:
        return None
    return FontFormat(name=name, bold=bold, italic=italic, size=size, color=color)


def _read_optional_str(obj: Any, attribute: str) -> str | None:
    """Read a string property; ``None`` when Excel raises (property not applicable)."""

    value = _read_attr(obj, attribute)
    if value is None:
        return None
    return str(value)


def _read_optional_bool(obj: Any, attribute: str) -> bool | None:
    """Read a boolean property; ``None`` when Excel raises."""

    value = _read_attr(obj, attribute)
    if value is None:
        return None
    return bool(value)


def _read_optional_float(obj: Any, attribute: str) -> float | None:
    """Read a float property; ``None`` when Excel raises or returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _read_optional_int(obj: Any, attribute: str) -> int | None:
    """Read an integer property; ``None`` when Excel raises or returns ``Null``."""

    value = _read_attr(obj, attribute)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _find_cf_rule(
    conditions: Any, before: list[ConditionalFormatRule], spec: ConditionalFormatSpec
) -> Any:
    """Return the COM rule object identified by ``spec.priority``."""

    priority = spec.priority
    if priority is None:  # pragma: no cover - validated before the call
        raise InvalidRangeError("A rule priority is required.", context={"priority": None})
    match = next((rule for rule in before if rule.priority == priority), None)
    if match is None:
        raise InvalidRangeError(
            f"No conditional-format rule has priority {priority}.",
            context={"priority": priority, "priorities": [rule.priority for rule in before]},
        )
    if match.type is not spec.type:
        raise InvalidRangeError(
            "The rule at the requested priority has a different type; a conditional-format "
            "rule type is never silently converted.",
            context={
                "priority": priority,
                "existing_type": match.type.value,
                "requested_type": spec.type.value,
            },
        )

    count = _read_int_or_none(conditions, "Count") or 0
    for index in range(1, count + 1):
        rule_obj = conditions.Item(index)
        if _read_int_or_none(rule_obj, "Priority") == priority:
            return rule_obj
    raise InvalidRangeError(  # pragma: no cover - the collection changed under us
        f"The rule with priority {priority} disappeared before it could be updated.",
        context={"priority": priority},
    )


def _delete_cf_rule(
    conditions: Any, before: list[ConditionalFormatRule], spec: ConditionalFormatSpec
) -> None:
    """Delete the rule identified by ``spec.priority`` (no renumbering)."""

    rule_obj = _find_cf_rule(conditions, before, spec)
    try:
        rule_obj.Delete()
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            "Excel rejected the conditional-format delete.",
            context={"priority": spec.priority, "error": str(exc)},
        ) from exc


def _apply_cf_rule(
    conditions: Any,
    spec: ConditionalFormatSpec,
    mode: str,
    before: list[ConditionalFormatRule],
) -> list[int]:
    """Add (or replace) one rule and return the resulting priorities."""

    if mode == "modify":
        # ``FormatCondition`` has no in-place Modify and ``Formula1`` is
        # read-only, so a modify is a delete + add with the priority restored.
        _find_cf_rule(conditions, before, spec).Delete()

    arguments: list[Any] = [
        _CF_TYPE_VALUES[spec.type],
        _OPERATOR_VALUES[spec.operator] if spec.operator is not None else None,
        spec.formula1,
    ]
    if spec.formula2 is not None:
        arguments.append(spec.formula2)

    try:
        rule_obj = conditions.Add(*arguments)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            "Excel rejected the conditional-format rule.",
            context={
                "type": spec.type.value,
                "operator": spec.operator.value if spec.operator is not None else None,
                "formula1": spec.formula1,
                "formula2": spec.formula2,
                "mode": mode,
                "error": str(exc),
            },
        ) from exc

    _apply_cf_rule_format(rule_obj, spec)

    applied: list[int] = []
    if spec.priority is not None:
        try:
            rule_obj.Priority = spec.priority
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                "Excel rejected the requested conditional-format priority.",
                context={"priority": spec.priority, "error": str(exc)},
            ) from exc
        applied.append(spec.priority)
    else:
        priority = _read_int_or_none(rule_obj, "Priority")
        if priority is not None:
            applied.append(priority)
    return applied


def _apply_cf_rule_format(rule_obj: Any, spec: ConditionalFormatSpec) -> None:
    """Apply the bounded ``StopIfTrue`` / ``Interior`` / ``Font`` subset."""

    if spec.stop_if_true is not None:
        _set_com(rule_obj, "StopIfTrue", bool(spec.stop_if_true))

    if spec.interior_color is not None:
        interior = _read_attr(rule_obj, "Interior")
        if interior is None:  # pragma: no cover - defensive COM read
            raise InvalidRangeError(
                "The conditional-format rule does not expose an Interior format.",
                context={"type": spec.type.value},
            )
        _set_com(interior, "Color", int(spec.interior_color))

    if spec.font is not None:
        font_obj = _read_attr(rule_obj, "Font")
        if font_obj is None:  # pragma: no cover - defensive COM read
            raise InvalidRangeError(
                "The conditional-format rule does not expose a Font format.",
                context={"type": spec.type.value},
            )
        if spec.font.name is not None:
            _set_com(font_obj, "Name", spec.font.name)
        if spec.font.bold is not None:
            _set_com(font_obj, "Bold", bool(spec.font.bold))
        if spec.font.italic is not None:
            _set_com(font_obj, "Italic", bool(spec.font.italic))
        if spec.font.size is not None:
            _set_com(font_obj, "Size", spec.font.size)
        if spec.font.color is not None:
            _set_com(font_obj, "Color", int(spec.font.color))


def _normalize_cf_formula(value: str | None) -> str | None:
    """Normalize a rule formula for comparison (Excel stores it with a leading ``=``).

    Excel accepts ``FormatConditions.Add`` operands with or without the leading
    ``=`` and always stores them with it, so the postcondition comparison must not
    report a mismatch for the two equivalent spellings. The returned rule keeps
    Excel's stored form; only the comparison is normalized.
    """

    if value is None:
        return None
    return value[1:] if value.startswith("=") else value


def _cf_rule_matches(rule: ConditionalFormatRule, spec: ConditionalFormatSpec) -> bool:
    """Return whether a re-read rule matches the requested specification."""

    if rule.type is not spec.type:
        return False
    if spec.formula1 is not None and _normalize_cf_formula(rule.formula1) != (
        _normalize_cf_formula(spec.formula1)
    ):
        return False
    if spec.formula2 is not None and _normalize_cf_formula(rule.formula2) != (
        _normalize_cf_formula(spec.formula2)
    ):
        return False
    if spec.operator is not None and rule.operator is not spec.operator:
        return False
    if spec.stop_if_true is not None and rule.stop_if_true is not spec.stop_if_true:
        return False
    if spec.interior_color is not None and rule.interior_color != spec.interior_color:
        return False
    if spec.font is not None and rule.font is not None:
        for field in ("name", "bold", "italic", "size", "color"):
            expected = getattr(spec.font, field)
            if expected is not None and getattr(rule.font, field) != expected:
                return False
    return True


def _cf_postcondition(
    before: list[ConditionalFormatRule],
    after: list[ConditionalFormatRule],
    spec: ConditionalFormatSpec,
    mode: str,
    applied: list[int],
) -> bool:
    """Verify the re-read rule collection against the requested operation."""

    if mode == "delete":
        return len(after) == len(before) - 1 and not any(
            rule.priority == spec.priority for rule in after
        )

    expected_count = len(before) + 1 if mode == "add" else len(before)
    if len(after) != expected_count:
        return False
    if not any(_cf_rule_matches(rule, spec) for rule in after):
        return False
    if spec.priority is not None:
        return any(
            rule.priority == spec.priority and _cf_rule_matches(rule, spec) for rule in after
        )
    return not applied or any(rule.priority in applied for rule in after)


def _unsupported_cf_types(rules: list[ConditionalFormatRule]) -> list[str]:
    """Return the sorted, de-duplicated unsupported rule types of a collection."""

    found: set[str] = set()
    for rule in rules:
        if rule.supported:
            continue
        found.add(rule.type.value)
    return sorted(found)


# ---------------------------------------------------------------------------
# Write helpers (task_043)
# ---------------------------------------------------------------------------
def _validate_payload_shape(
    target_rows: int, target_cols: int, rows: int, cols: int, width: int
) -> None:
    """Validate the write payload shape against the resolved target (no COM).

    Rejects a wrong-sized payload, an empty grid or a ragged row **before** any
    workbook mutation.
    """

    if width != rows:
        raise InvalidRangeError(
            f"Payload row count {width} does not match rows={rows}.",
            context={"rows": rows, "payload_rows": width},
        )
    if rows != target_rows or cols != target_cols:
        raise InvalidRangeError(
            f"Payload shape {rows}x{cols} does not match the target {target_rows}x{target_cols}.",
            context={
                "target_rows": target_rows,
                "target_cols": target_cols,
                "payload_rows": rows,
                "payload_cols": cols,
            },
        )


def _validate_target_cells(cells: int) -> None:
    """Reject a target larger than ``MAX_RANGE_CELLS`` before any mutation."""

    if cells > MAX_RANGE_CELLS:
        raise InvalidRangeError(
            f"Target {cells} cells exceeds MAX_RANGE_CELLS ({MAX_RANGE_CELLS}).",
            context={"cells": cells, "limit": MAX_RANGE_CELLS},
        )


def _values_to_raw_grid(
    values: list[list[CellValue]], rows: int, cols: int
) -> tuple[tuple[Any, ...], ...]:
    """Convert a typed :class:`CellValue` grid into a raw Excel value grid."""

    grid: list[tuple[Any, ...]] = []
    for r in range(rows):
        row_values: list[Any] = []
        for c in range(cols):
            row_values.append(_cell_to_raw(values[r][c]))
        grid.append(tuple(row_values))
    return tuple(grid)


def _cell_to_raw(cell: CellValue) -> Any:
    """Convert one typed :class:`CellValue` into the raw value Excel should store."""

    if cell.kind is CellValueKind.EMPTY or cell.value is None:
        return None
    if cell.kind is CellValueKind.DATE:
        text = str(cell.value)
        try:
            moment = datetime.fromisoformat(text)
        except ValueError:
            return text
        return (moment - _EXCEL_EPOCH).total_seconds() / 86400.0
    if cell.kind is CellValueKind.ERROR:
        text = str(cell.value)
        code = _ERROR_LITERAL_TO_CODE.get(text)
        return code if code is not None else text
    if cell.kind is CellValueKind.BOOLEAN:
        return bool(cell.value)
    if cell.kind is CellValueKind.NUMBER:
        return cell.value
    return cell.value


#: Inverse of :data:`excel_vba_mcp.adapters.excel.cell_values.EXCEL_ERROR_CODE_TO_LITERAL`
#: for writing an Excel error value back through ``Range.Value2`` (task_038 §7).
_ERROR_LITERAL_TO_CODE: dict[str, int] = {
    literal: code for code, literal in EXCEL_ERROR_CODE_TO_LITERAL.items()
}


def _formulas_to_raw_grid(
    formulas: list[list[str]], rows: int, cols: int
) -> tuple[tuple[Any, ...], ...]:
    """Convert a formula-string grid into a raw Excel value grid (verbatim)."""

    grid: list[tuple[Any, ...]] = []
    for r in range(rows):
        grid.append(tuple(formulas[r][c] for c in range(cols)))
    return tuple(grid)


def _require_formula_strings(formulas: list[list[str]]) -> None:
    """Reject a formula payload that carries a non-formula / non-string entry."""

    for r, row in enumerate(formulas):
        for c, value in enumerate(row):
            if not isinstance(value, str) or not value.startswith("="):
                raise InvalidRangeError(
                    "A formula write requires every entry to be a formula string "
                    "beginning with '='; use write_values for constants.",
                    context={"row": r, "col": c, "value": value},
                )


def _assign_value2(
    range_obj: Any, raw_grid: tuple[tuple[Any, ...], ...], rows: int, cols: int
) -> None:
    """Assign a raw value grid to ``range_obj.Value2`` (single COM mutation)."""

    payload: Any = raw_grid
    if rows == 1 and cols == 1:
        payload = raw_grid[0][0]
    try:
        range_obj.Value2 = payload
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise InvalidRangeError(
            "Excel rejected the value write.",
            context={"rows": rows, "cols": cols, "error": str(exc)},
        ) from exc


def _assign_formula2(
    range_obj: Any, raw_grid: tuple[tuple[Any, ...], ...], rows: int, cols: int
) -> None:
    """Assign a raw formula grid through ``Formula2`` (fallback ``Formula``)."""

    payload: Any = raw_grid
    if rows == 1 and cols == 1:
        payload = raw_grid[0][0]
    try:
        range_obj.Formula2 = payload
    except Exception:  # noqa: BLE001 - Formula2 unavailable -> fallback to Formula
        try:
            range_obj.Formula = payload
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise InvalidRangeError(
                "Excel rejected the formula write.",
                context={"rows": rows, "cols": cols, "error": str(exc)},
            ) from exc


# ---------------------------------------------------------------------------
# ListObject structure guard (task_043)
# ---------------------------------------------------------------------------
def _capture_tables(worksheet: Any, range_obj: Any) -> dict[str, str]:
    """Capture the geometry of every ListObject intersecting ``range_obj``.

    Returns a mapping ``table name -> range address`` for the tables whose
    ``Range`` intersects the target; an empty mapping means the target is outside
    every table.
    """

    captured: dict[str, str] = {}
    try:
        list_objects = worksheet.ListObjects
        count = _safe_int(list_objects.Count)
    except Exception:  # noqa: BLE001 - sheet without ListObjects
        return captured
    for index in range(1, count + 1):
        try:
            list_object = list_objects.Item(index)
            table_range = list_object.Range
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        if not _ranges_intersect(range_obj, table_range):
            continue
        name = _safe_str(list_object.Name) or f"ListObject{index}"
        captured[name] = _safe_str(table_range.Address)
    return captured


def _ranges_intersect(a: Any, b: Any) -> bool:
    """Return whether two resolved single-area ranges intersect (best-effort)."""

    try:
        a_top = _safe_int(a.Top)
        a_left = _safe_int(a.Left)
        a_height = _safe_int(a.Height, default=1)
        a_width = _safe_int(a.Width, default=1)
        b_top = _safe_int(b.Top)
        b_left = _safe_int(b.Left)
        b_height = _safe_int(b.Height, default=1)
        b_width = _safe_int(b.Width, default=1)
    except Exception:  # noqa: BLE001 - defensive COM read
        return False
    a_bottom = a_top + a_height - 1
    a_right = a_left + a_width - 1
    b_bottom = b_top + b_height - 1
    b_right = b_left + b_width - 1
    return a_top <= b_bottom and b_top <= a_bottom and a_left <= b_right and b_left <= a_right


def _guard_table_structure(worksheet: Any, range_obj: Any, table_before: dict[str, str]) -> None:
    """Reject a write that changed an intersecting ListObject's row/column structure.

    Cell-level writes inside a table are allowed; a change to the table geometry
    (row count / column count / address) is a structural change owned by the
    Tables service and is rejected with
    :class:`~excel_vba_mcp.domain.errors.ListObjectIntersectionError`.
    """

    if not table_before:
        return
    table_after = _capture_tables(worksheet, range_obj)
    for name, address_before in table_before.items():
        address_after = table_after.get(name)
        if address_after != address_before:
            raise ListObjectIntersectionError(
                "The write changed the structure of an Excel Table; "
                "structural changes are owned by the table tools.",
                context={
                    "table": name,
                    "table_range": address_before,
                    "table_range_after": address_after,
                },
            )


# ---------------------------------------------------------------------------
# Postcondition verification (task_043)
# ---------------------------------------------------------------------------
def _values_match(expected: list[list[CellValue]], post: RangeValues) -> bool:
    """Compare the written value payload against the re-read typed values."""

    if post.rows != len(expected):
        return False
    for r, row in enumerate(expected):
        if len(post.values[r]) != len(row):
            return False
        for c, cell in enumerate(row):
            if not _cell_matches(cell, post.values[r][c]):
                return False
    return True


def _cell_matches(expected: CellValue, actual: CellValue) -> bool:
    """Return whether a re-read cell value matches the written payload cell."""

    # A declared EMPTY and a re-read None both mean "no value"; a declared None
    # value with any kind is likewise treated as an empty write.
    if expected.kind is CellValueKind.EMPTY or expected.value is None:
        return actual.kind is CellValueKind.EMPTY or actual.value is None
    if actual.value is None:
        return False
    if expected.kind is CellValueKind.ERROR:
        return actual.kind is CellValueKind.ERROR and actual.value == expected.value
    if expected.kind is CellValueKind.DATE:
        if actual.kind is not CellValueKind.DATE:
            return False
        return _dates_equal(str(expected.value), str(actual.value))
    if expected.kind is CellValueKind.BOOLEAN:
        return actual.kind is CellValueKind.BOOLEAN and bool(actual.value) == bool(expected.value)
    if expected.kind is CellValueKind.NUMBER:
        if actual.kind is not CellValueKind.NUMBER:
            return False
        try:
            return float(actual.value) == float(expected.value)  # type: ignore[arg-type,unused-ignore]
        except (TypeError, ValueError):
            return False
    return actual.kind is CellValueKind.STRING and actual.value == expected.value


def _dates_equal(expected: str, actual: str) -> bool:
    """Compare two ISO-8601 date strings by their date/time prefix (second precision)."""

    return expected[:19] == actual[:19]


def _formulas_match(expected: list[list[str]], post: RangeFormulas) -> bool:
    """Compare the written formula payload against the re-read ``Formula2`` grid."""

    if post.rows != len(expected):
        return False
    for r, row in enumerate(expected):
        if len(post.formulas[r]) != len(row):
            return False
        for c, value in enumerate(row):
            if post.formulas[r][c].formula2 != value:
                return False
    return True


# ---------------------------------------------------------------------------
# Structural helpers (task_045, task_038 §5)
# ---------------------------------------------------------------------------
def _resolve_shift(shift: str, allowed: dict[str, int], operation: str) -> int:
    """Resolve an explicit shift *name* to its Excel enum value (never guessed).

    An omitted / blank / unknown shift is rejected with a typed
    :class:`~excel_vba_mcp.domain.errors.InvalidRangeError`; Excel is never
    allowed to decide the shift direction from the range shape (task_038 §7).
    """

    if not isinstance(shift, str) or not shift.strip():
        raise InvalidRangeError(
            f"Argument 'shift' is required for {operation} and must never be guessed.",
            context={"shift": shift, "operation": operation},
        )
    normalized = shift.strip().lower()
    if normalized not in allowed:
        raise InvalidRangeError(
            f"Argument 'shift' must be one of {sorted(allowed)} for {operation}.",
            context={"shift": shift, "operation": operation, "allowed": sorted(allowed)},
        )
    return allowed[normalized]


def _require_unprotected(worksheet: Any) -> None:
    """Reject a structural mutation on a protected worksheet (task_038 §7)."""

    if bool(_read_attr(worksheet, "ProtectContents")):
        raise ProtectedTargetError(
            "The worksheet is protected; structural range operations are rejected.",
            context={"sheet": _safe_str(worksheet.Name)},
        )


@contextmanager
def _alerts_suppressed(app: Any) -> Iterator[None]:
    """Suppress Excel's built-in alert prompts for one call, restoring the prior value.

    ``Range.Merge`` over data raises Excel's own "keep only the upper-left value?"
    alert; the caller has already approved the data loss, so the alert would be a
    stray modal dialog. ``DisplayAlerts`` is restored in ``finally`` on every path
    (task_038 §8.4, addendum BG03). A missing ``Application`` (unit-test fakes)
    yields without touching any setting.
    """

    if app is None:
        yield
        return

    prior = _read_attr(app, "DisplayAlerts")
    try:
        app.DisplayAlerts = False
    except Exception:  # noqa: BLE001 - setting unavailable -> run without suppression
        yield
        return
    try:
        yield
    finally:
        with suppress(Exception):
            app.DisplayAlerts = True if prior is None else prior


def _table_intersections(wb: Any, worksheet: Any, range_obj: Any) -> list[dict[str, str]]:
    """Return every ListObject on ``worksheet`` whose ``Range`` hits ``range_obj``.

    Detection uses ``Application.Intersect`` when available (task_038 §5.3) and
    falls back to a rectangle intersection computed from ``Top`` / ``Left`` /
    ``Height`` / ``Width``. The intersection address is included so the caller
    can see exactly which cells overlap the table.
    """

    intersections: list[dict[str, str]] = []
    try:
        list_objects = worksheet.ListObjects
        count = _safe_int(list_objects.Count)
    except Exception:  # noqa: BLE001 - sheet without ListObjects
        return intersections

    app = _read_attr(wb, "Application")
    for index in range(1, count + 1):
        try:
            list_object = list_objects.Item(index)
            table_range = list_object.Range
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
        intersection = _intersection_address(app, range_obj, table_range)
        if intersection is None:
            continue
        intersections.append(
            {
                "table": _safe_str(list_object.Name) or f"ListObject{index}",
                "table_range": _safe_str(table_range.Address),
                "intersection_address": intersection,
            }
        )
    return intersections


def _guard_structural_table_intersection(
    wb: Any, worksheet: Any, range_obj: Any, operation: str
) -> None:
    """Reject a structural operation whose target intersects a ListObject.

    Structural row/column changes are owned by the Tables/ListObject services and
    must never be performed through the Range API (task_038 §5.3). The rejection
    names the owning table and its range so the caller can route to the table
    tools instead.
    """

    intersections = _table_intersections(wb, worksheet, range_obj)
    if not intersections:
        return
    first = intersections[0]
    raise ListObjectIntersectionError(
        f"The {operation} target intersects Excel Table {first['table']!r}; "
        "structural changes are owned by the table tools.",
        context={
            "operation": operation,
            "table": first["table"],
            "table_range": first["table_range"],
            "intersection_address": first["intersection_address"],
            "tables": intersections,
        },
    )


def _intersection_address(app: Any, a: Any, b: Any) -> str | None:
    """Return the address of ``a`` ∩ ``b``, or ``None`` when they do not intersect."""

    if app is not None:
        try:
            result = app.Intersect(a, b)
        except Exception:  # noqa: BLE001 - different sheets / COM failure -> fallback
            result = None
        if result is not None:
            return _safe_str(result.Address) or None

    try:
        a_top = _safe_int(a.Top)
        a_left = _safe_int(a.Left)
        a_height = _safe_int(a.Height, default=1)
        a_width = _safe_int(a.Width, default=1)
        b_top = _safe_int(b.Top)
        b_left = _safe_int(b.Left)
        b_height = _safe_int(b.Height, default=1)
        b_width = _safe_int(b.Width, default=1)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None

    top = max(a_top, b_top)
    left = max(a_left, b_left)
    bottom = min(a_top + a_height - 1, b_top + b_height - 1)
    right = min(a_left + a_width - 1, b_left + b_width - 1)
    if top > bottom or left > right:
        return None
    if top == bottom and left == right:
        return _cell_address(top, left)
    return f"{_cell_address(top, left)}:{_cell_address(bottom, right)}"


def _declare(obj: Any) -> tuple[int, int]:
    """Return the declared ``(rows, cols)`` of a range (0, 0 when unavailable)."""

    rows_obj = _read_attr(obj, "Rows")
    cols_obj = _read_attr(obj, "Columns")
    rows = _safe_int(_read_attr(rows_obj, "Count") if rows_obj is not None else 0)
    cols = _safe_int(_read_attr(cols_obj, "Count") if cols_obj is not None else 0)
    return rows, cols


def _read_value_grid(range_obj: Any) -> tuple[tuple[Any, ...], ...]:
    """Read ``range_obj.Value2`` reshaped to the range's declared dimensions.

    Excel returns a scalar for a single cell, a flat tuple for a single row or
    column and a nested tuple for a block. The declared ``Rows.Count`` /
    ``Columns.Count`` disambiguate the flat cases so a single-column range is
    never mistaken for a single row.
    """

    raw = _read_attr(range_obj, "Value2")
    rows, cols = _declare(range_obj)
    if rows <= 0 or cols <= 0:
        return _normalize_grid(raw)
    if rows * cols == 1:
        return ((raw,),)
    if isinstance(raw, tuple) and raw and isinstance(raw[0], tuple):
        return tuple(tuple(row) for row in raw)
    if rows == 1:
        return (tuple(raw),)
    if cols == 1:
        return tuple((value,) for value in raw)
    return _normalize_grid(raw)


def _count_non_empty(grid: tuple[tuple[Any, ...], ...]) -> int:
    """Count the cells of a raw value grid that hold a value."""

    total = 0
    for row in grid:
        for value in row:
            if value is not None and value != "":
                total += 1
    return total


def _non_anchor_addresses(
    grid: tuple[tuple[Any, ...], ...], *, across: bool, top: int, left: int
) -> list[str]:
    """Return the sheet addresses of non-anchor cells that hold data (merge preview).

    The anchor is the top-left cell of the range, or — when ``across`` is set —
    the first cell of every row, because ``Merge(Across=True)`` keeps one value
    per row.
    """

    rows = len(grid)
    cols = max((len(row) for row in grid), default=0)
    addresses: list[str] = []
    for r in range(rows):
        for c in range(cols):
            is_anchor = c == 0 if across else (r == 0 and c == 0)
            if is_anchor:
                continue
            value = _grid_cell(grid, r, c)
            if value is not None and value != "":
                addresses.append(_cell_address(top + r, left + c))
    return addresses


def _merged_areas(range_obj: Any) -> list[str]:
    """Return the addresses of the merged areas that intersect ``range_obj``.

    ``Range.MergeCells`` is ``True`` for a fully merged range, ``False`` for an
    unmerged one and ``Null`` when the range mixes merged and unmerged cells. The
    mixed case is scanned cell-by-cell, bounded by ``_MAX_MERGE_SCAN_CELLS``; the
    caller reports the affected ``MergeArea`` addresses rather than guessing.
    """

    state = _read_attr(range_obj, "MergeCells")
    if state is False:
        return []
    if state is True:
        address = _merge_area_address(range_obj.Cells(1, 1))
        return [address] if address else []

    rows, cols = _declare(range_obj)
    if rows <= 0 or cols <= 0:
        return []

    areas: list[str] = []
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            if (r - 1) * cols + (c - 1) >= _MAX_MERGE_SCAN_CELLS:
                return areas
            cell = range_obj.Cells(r, c)
            if _read_attr(cell, "MergeCells") is not True:
                continue
            address = _merge_area_address(cell)
            if address and address not in areas:
                areas.append(address)
    return areas


def _merge_area_address(cell: Any) -> str | None:
    """Return the ``MergeArea`` address of one cell, or ``None`` when unavailable."""

    merge_area = _read_attr(cell, "MergeArea")
    if merge_area is None:
        return None
    return _safe_str(_read_attr(merge_area, "Address")) or None


def _spill_state(range_obj: Any) -> dict[str, Any]:
    """Report the dynamic-array spill state of a target (best-effort, no guessing).

    ``spilling`` is ``True`` when the range itself is a spill anchor and
    ``non_anchor`` is ``True`` when the range resolves to cells inside somebody
    else's spill range (``SpillParent`` differs from the range address). A build
    without the dynamic-array properties yields ``None`` / ``False`` rather than
    an invented state.
    """

    spilling = _read_attr(range_obj, "Spilling")
    parent = _read_attr(range_obj, "SpillParent")
    parent_address = _safe_str(_read_attr(parent, "Address")) if parent is not None else ""
    own_address = _safe_str(_read_attr(range_obj, "Address"))
    non_anchor = bool(parent_address) and parent_address.replace("$", "") != own_address.replace(
        "$", ""
    )
    return {
        "spilling": None if spilling is None else bool(spilling),
        "spill_parent": parent_address or None,
        "non_anchor": non_anchor,
    }


def _referenced_formulas(worksheet: Any, target: RangeAddress) -> dict[str, Any]:
    """Best-effort same-sheet scan for formulas that reference ``target``.

    The scan reads the sheet's ``UsedRange.Formula`` once and looks for the
    target's address text (with and without ``$``) inside formula strings. It is
    deliberately **not** presented as complete: other sheets, defined names and
    external references are not inspected (task_038 §5.2 item 5).
    """

    note = (
        "best-effort same-sheet UsedRange scan; other sheets, names and external "
        "references are not inspected"
    )
    used = _read_attr(worksheet, "UsedRange")
    if used is None:
        return {"referenced": False, "cells": [], "note": f"{note} (used range unavailable)"}

    row_count, col_count = _declare(used)
    if row_count <= 0 or col_count <= 0:
        return {"referenced": False, "cells": [], "note": f"{note} (empty used range)"}
    if row_count * col_count > MAX_RANGE_CELLS:
        return {
            "referenced": False,
            "cells": [],
            "note": f"{note} (used range exceeds MAX_RANGE_CELLS; scan skipped)",
        }

    top = _safe_int(_read_attr(used, "Row"), default=1)
    left = _safe_int(_read_attr(used, "Column"), default=1)
    try:
        grid = _normalize_grid(_read_attr(used, "Formula"))
    except Exception:  # noqa: BLE001 - defensive COM read
        return {"referenced": False, "cells": [], "note": f"{note} (formula read failed)"}

    tokens = _reference_tokens(target)
    hits: list[str] = []
    for r in range(row_count):
        for c in range(col_count):
            text = _grid_cell(grid, r, c)
            if not isinstance(text, str) or not text.startswith("="):
                continue
            upper = text.upper()
            if any(token in upper for token in tokens):
                hits.append(_cell_address(top + r, left + c))
                if len(hits) >= _MAX_REFERENCE_HITS:
                    return {"referenced": True, "cells": hits, "note": note}
    return {"referenced": bool(hits), "cells": hits, "note": note}


def _reference_tokens(target: RangeAddress) -> tuple[str, ...]:
    """Return the uppercase address tokens a formula may use for ``target``."""

    normalized = target.address.replace("$", "").upper()
    if target.rows <= 1 and target.cols <= 1:
        return (normalized,)
    tokens = {normalized, normalized.split(":")[0]}
    return tuple(sorted(tokens))


def _expanded_destination(destination_range: Any, rows: int, cols: int) -> str:
    """Return the pasted block address for a copy onto ``destination_range``.

    A single anchor cell receives the whole source block; a destination that
    already matches the source shape keeps its own address.
    """

    top = _safe_int(_read_attr(destination_range, "Row"), default=1)
    left = _safe_int(_read_attr(destination_range, "Column"), default=1)
    return f"{_cell_address(top, left)}:{_cell_address(top + rows - 1, left + cols - 1)}"


def _range_values_equal(expected: RangeValues, post: RangeValues) -> bool:
    """Compare two typed value grids cell-by-cell (postcondition helper)."""

    if post.rows != expected.rows or post.cols != expected.cols:
        return False
    for r in range(expected.rows):
        for c in range(expected.cols):
            if not _cell_matches(expected.values[r][c], post.values[r][c]):
                return False
    return True


def _all_cells_empty(post: RangeValues) -> bool:
    """Return whether every cell of a re-read block is empty."""

    return all(
        cell.kind is CellValueKind.EMPTY or cell.value is None
        for row in post.values
        for cell in row
    )


def _read_block(
    worksheet: Any, top: int, left: int, rows: int, cols: int
) -> tuple[tuple[Any, ...], ...]:
    """Read a raw ``rows`` x ``cols`` block anchored at ``(top, left)``."""

    block = worksheet.Range(
        worksheet.Cells(top, left), worksheet.Cells(top + rows - 1, left + cols - 1)
    )
    return _read_value_grid(block)


def _insert_verified(
    before: RangeAddress,
    post: RangeAddress,
    post_values: RangeValues,
    *,
    shifted_values: tuple[tuple[Any, ...], ...],
    pre_grid: tuple[tuple[Any, ...], ...],
) -> bool:
    """Postcondition for an insert.

    Excel moves the target content to the shifted block (``right`` -> the same
    rows one block-width to the right, ``down`` -> the same columns one
    block-height down) and leaves the inserted cells blank. The postcondition
    verifies the target address/dimensions, that the inserted cells are empty and
    that the pre-insert content reappears at the shifted address.
    """

    return (
        _same_geometry(before, post)
        and _all_cells_empty(post_values)
        and shifted_values == pre_grid
    )


def _delete_verified(
    target: RangeAddress,
    post: RangeAddress,
    post_values: RangeValues,
    *,
    before_non_empty: int,
) -> bool:
    """Postcondition for a delete: same address/geometry, content not increased."""

    after_non_empty = sum(
        1
        for row in post_values.values
        for cell in row
        if cell.kind is not CellValueKind.EMPTY and cell.value is not None
    )
    return _same_geometry(target, post) and after_non_empty <= before_non_empty


def _same_geometry(before: RangeAddress, post: RangeAddress) -> bool:
    """Return whether the re-read range kept its address and dimensions."""

    return (
        before.address.replace("$", "") == post.address.replace("$", "")
        and before.rows == post.rows
        and before.cols == post.cols
    )


__all__ = [
    "RangesAdapter",
    "validate_format_patch",
    "validate_validation_spec",
    "validate_conditional_format_spec",
]


def _validate_window(offset: int, limit: int | None) -> None:
    """Reject an invalid ``offset`` / ``limit`` before any COM access."""

    if offset < 0:
        raise InvalidRangeError("Argument 'offset' must be >= 0.", context={"offset": offset})
    if limit is not None and limit < 1:
        raise InvalidRangeError(
            "Argument 'limit' must be >= 1 when provided.", context={"limit": limit}
        )


def _window(
    total_rows: int, total_cols: int, offset: int, limit: int | None
) -> tuple[int, int, bool]:
    """Compute the bounded ``(start, end, truncated)`` row window.

    Rejects a request whose full size exceeds
    :data:`~excel_vba_mcp.domain.models.MAX_RANGE_CELLS` before any cell is read,
    then bounds the returned rows by
    :data:`~excel_vba_mcp.domain.models.MAX_RESPONSE_CELLS` (setting
    ``truncated`` when the response is cut short).
    """

    if total_rows * total_cols > MAX_RANGE_CELLS:
        raise InvalidRangeError(
            f"Range {total_rows}x{total_cols}={total_rows * total_cols} exceeds "
            f"MAX_RANGE_CELLS ({MAX_RANGE_CELLS}).",
            context={
                "rows": total_rows,
                "cols": total_cols,
                "cells": total_rows * total_cols,
                "limit": MAX_RANGE_CELLS,
            },
        )

    start = min(offset, total_rows)
    end = total_rows if limit is None else min(total_rows, offset + limit)
    requested = end - start

    truncated = requested * total_cols > MAX_RESPONSE_CELLS
    if truncated:
        max_rows = max(0, MAX_RESPONSE_CELLS // max(1, total_cols))
        end = start + min(requested, max_rows)
    return start, end, truncated
