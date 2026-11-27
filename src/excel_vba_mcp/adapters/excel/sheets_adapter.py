"""STA-local Sheets adapter (task_024, spec §7.5 / §6.5).

:class:`SheetsAdapter` is the single place where the Excel ``Sheets`` /
``Worksheets`` object model is touched. It is a **pure STA-local helper**: every
method is synchronous, receives the live ``Workbook`` COM proxy (``Any``) and
returns only Pydantic DTOs / primitives, so results cross the single-STA worker
boundary safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`).
It never imports ``pywin32`` at module import time and never calls
``.Activate`` / ``.Select`` on any object.

Sheets vs Worksheets (spec §7.5):

* ``Workbook.Sheets`` enumerates **all** sheet types (Worksheet, Chart, Macro,
  Dialog); ``Workbook.Worksheets`` enumerates only ``Worksheet`` objects. The
  adapter distinguishes the two collections explicitly: ``object_type`` is
  derived from membership in ``wb.Worksheets`` (a ``Worksheet`` proxy) versus
  ``wb.Sheets`` (anything else).
* ``CodeName`` is a **read-only** identifier and is only read from a
  ``Worksheet``; it is never written (renaming a sheet never touches it).
* ``Visible`` is one of ``xlSheetVisible`` (0) / ``xlSheetHidden`` (1) /
  ``xlSheetVeryHidden`` (2); the adapter never hides the last visible worksheet
  and never mutates a sheet's code name.

All mutating methods assume the caller already validated the session
(writable), a snapshot exists and approval was granted; the adapter only
executes the Excel object-model call on the STA thread.
"""

from __future__ import annotations

from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError, InvalidRangeError, ObjectNotFoundError
from excel_vba_mcp.domain.models import SheetInfo
from excel_vba_mcp.domain.statuses import SheetObjectType, SheetVisibility

# Excel ``XlspSheetVisibility`` codes (note: xlSheetVisible is -1, not 0;
# 0 is xlSheetHidden). Getting this wrong makes the last-visible guard misfire.
_XL_SHEET_VISIBLE = -1
_XL_SHEET_HIDDEN = 0
_XL_SHEET_VERY_HIDDEN = 2

# ``WshSheetType`` codes for ``Sheets.Add(Type=...)`` (MVP supports Worksheet).
_XL_WORKSHEET = -4167

# Excel worksheet-name constraints (validation for an explicit copy target name).
_MAX_SHEET_NAME_LEN = 31
_INVALID_SHEET_NAME_CHARS = frozenset(":\\/?*[]")


def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively (never raises)."""

    try:
        text = value
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    if text is None:
        return default
    if isinstance(text, str):
        return text
    return str(text)


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively (never raises)."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _visibility_from_code(code: int) -> SheetVisibility:
    if code == _XL_SHEET_VISIBLE:
        return SheetVisibility.VISIBLE
    if code == _XL_SHEET_HIDDEN:
        return SheetVisibility.HIDDEN
    if code == _XL_SHEET_VERY_HIDDEN:
        return SheetVisibility.VERY_HIDDEN
    # Unknown codes are treated as visible rather than silently hidden.
    return SheetVisibility.VISIBLE


def _is_worksheet(wb: Any, sheet: Any) -> bool:
    """Return whether ``sheet`` is a ``Worksheet`` (member of ``wb.Worksheets``).

    The authoritative test is membership in ``wb.Worksheets`` (which enumerates
    only ``Worksheet`` objects); a Chart / Macro / Dialog sheet in ``wb.Sheets``
    is never in ``wb.Worksheets`` and is therefore a non-worksheet. The
    membership check is by name (COM proxies compare by identity unreliably
    across calls).
    """

    sheet_name = _safe_str(sheet.Name)
    try:
        worksheets = wb.Worksheets
        for ws in worksheets:
            try:
                if _safe_str(ws.Name) == sheet_name:
                    return True
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
    except Exception:  # noqa: BLE001 - defensive COM read
        pass
    return False


def _read_code_name(sheet: Any) -> str | None:
    """Read ``CodeName`` from a ``Worksheet`` (read-only; never written)."""

    try:
        value = sheet.CodeName
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    if value is None:
        return None
    text = str(value)
    return text or None


def _read_protected(sheet: Any) -> bool | None:
    """Read ``Protected`` (diagnostic only; never mutates protection)."""

    try:
        return bool(sheet.ProtectContents)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _read_table_count(sheet: Any) -> int:
    """Return the number of ``ListObjects`` on a worksheet (0 when absent)."""

    try:
        return _safe_int(sheet.ListObjects.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        return 0


def _read_used_range_address(sheet: Any) -> str | None:
    """Return the ``UsedRange.Address`` string, or ``None`` when unavailable."""

    try:
        address = sheet.UsedRange.Address
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    if address is None:
        return None
    text = str(address)
    return text or None


def describe_sheet(wb: Any, sheet: Any) -> SheetInfo:
    """Build a :class:`SheetInfo` DTO for one sheet proxy (STA-local).

    Args:
        wb: The owning ``Workbook`` proxy (used to distinguish
            ``Sheets`` vs ``Worksheets`` membership).
        sheet: One ``Sheets`` collection item proxy.

    Returns:
        The populated :class:`SheetInfo` (name, index, object_type, visibility,
        and Worksheet-only code_name / protected / table_count / used_range).
    """

    name = _safe_str(sheet.Name)
    index = _safe_int(sheet.Index)
    object_type_code = _safe_int(sheet.Type, default=-1)
    worksheet = _is_worksheet(wb, sheet)
    visibility = _visibility_from_code(_safe_int(sheet.Visible, _XL_SHEET_VISIBLE))
    object_type = (
        SheetObjectType.WORKSHEET if worksheet else _classify_non_worksheet(object_type_code)
    )
    code_name = _read_code_name(sheet) if worksheet else None
    protected = _read_protected(sheet) if worksheet else None
    table_count = _read_table_count(sheet) if worksheet else 0
    used_range = _read_used_range_address(sheet) if worksheet else None
    return SheetInfo(
        name=name,
        code_name=code_name,
        index=index,
        object_type=object_type,
        visibility=visibility,
        protected=protected,
        table_count=table_count,
        used_range_address=used_range,
    )


def _classify_non_worksheet(object_type: int) -> SheetObjectType:
    """Classify a non-worksheet sheet from its ``Sheets``-collection ``Type``.

    ``object_type`` is the raw ``sheet.Type`` value read once in
    :func:`describe_sheet`: ``xlChart`` (-4109) -> Chart,
    ``xlExcel4IntlMacroSheet`` (3) / ``xlExcel4MacroSheet`` (4) -> Macro,
    ``xlDialogSheet`` (1) -> Dialog. Anything else is treated as a Chart sheet
    (a non-worksheet sheet), never as a Worksheet.
    """

    if object_type == 4:  # xlExcel4MacroSheet
        return SheetObjectType.MACRO
    if object_type == 1:  # xlDialogSheet (on the Sheets collection)
        return SheetObjectType.DIALOG
    # xlChart (-4109), legacy xlChart (3), or unknown -> a Chart (non-worksheet) sheet.
    return SheetObjectType.CHART


def _require_sheet(sheets: Any, name: str) -> Any:
    """Resolve ``name`` to a sheet proxy or raise a structured error."""

    try:
        return sheets.Item(name)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        # A missing sheet is a caller-addressed missing identity, not COM
        # contention: surface NOT_FOUND (not the generic COM_BUSY fallback).
        raise ObjectNotFoundError(
            f"Sheet {name!r} was not found in the workbook.",
            context={"sheet": name, "reason": "not_found"},
        ) from exc


def _sheet_count(sheets: Any) -> int:
    """Return ``Sheets.Count`` defensively (0 when unavailable)."""

    try:
        return _safe_int(sheets.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        return 0


def _visible_worksheet_count(wb: Any) -> int:
    """Count visible worksheets in ``wb`` (the last-visible guard input)."""

    count = 0
    try:
        worksheets = wb.Worksheets
        for ws in worksheets:
            try:
                code = _safe_int(ws.Visible, _XL_SHEET_VISIBLE)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
            if code == _XL_SHEET_VISIBLE:
                count += 1
    except Exception:  # noqa: BLE001 - defensive COM read
        return count
    return count


class SheetsAdapter:
    """STA-local sheet inventory + worksheet operations (no COM at import)."""

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def list_sheets(self, wb: Any) -> list[SheetInfo]:
        """Return one :class:`SheetInfo` per ``wb.Sheets`` item (all types)."""

        sheets = wb.Sheets
        result: list[SheetInfo] = []
        for sheet in sheets:
            result.append(describe_sheet(wb, sheet))
        return result

    def get_sheet(self, wb: Any, name: str) -> SheetInfo:
        """Return the :class:`SheetInfo` for a single sheet by name."""

        sheet = _require_sheet(wb.Sheets, name)
        return describe_sheet(wb, sheet)

    # ------------------------------------------------------------------
    # Mutations (executed on the STA thread)
    # ------------------------------------------------------------------
    def add_sheet(self, wb: Any, name: str, *, sheet_type: str = "worksheet") -> SheetInfo:
        """Add a new worksheet named ``name`` and return its DTO.

        MVP supports ``sheet_type == "worksheet"`` only; any other type is
        rejected with a structured error (no silent fallback).
        """

        if sheet_type != SheetObjectType.WORKSHEET.value:
            raise ExcelVbaMcpError(
                f"Unsupported sheet type {sheet_type!r}; only 'worksheet' is supported.",
                context={"sheet_type": sheet_type, "sheet": name},
            )
        sheets = wb.Sheets
        try:
            added = sheets.Add(Type=_XL_WORKSHEET)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not add sheet {name!r}.",
                context={"sheet": name, "sheet_type": sheet_type},
            ) from exc
        try:
            added.Name = name
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not name the new sheet {name!r} (duplicate or invalid name?).",
                context={"sheet": name, "sheet_type": sheet_type},
            ) from exc
        return describe_sheet(wb, added)

    def rename_sheet(self, wb: Any, old_name: str, new_name: str) -> SheetInfo:
        """Rename ``old_name`` to ``new_name`` and return the updated DTO.

        ``CodeName`` is never modified; only the user-visible ``Name`` changes.
        """

        sheets = wb.Sheets
        sheet = _require_sheet(sheets, old_name)
        try:
            sheet.Name = new_name
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not rename sheet {old_name!r} to {new_name!r} "
                "(duplicate or invalid name?).",
                context={"old_name": old_name, "new_name": new_name},
            ) from exc
        return describe_sheet(wb, sheet)

    def copy_sheet(
        self,
        wb: Any,
        name: str,
        *,
        to_end: bool = True,
        target_name: str | None = None,
    ) -> SheetInfo:
        """Copy ``name`` and return the DTO of the new sheet.

        ``to_end=True`` places the copy after the last sheet; otherwise Excel
        places it before the source. ``CodeName`` is never touched on the source
        (the copy receives its own Excel-assigned code name).

        When ``target_name`` is given the new sheet is renamed to it (validated
        non-empty, ``<= 31`` chars and unique; an invalid value raises
        ``INVALID_RANGE``). When ``target_name`` is ``None`` the copy keeps the
        Excel-generated name (unchanged behaviour).
        """

        sheets = wb.Sheets
        source = _require_sheet(sheets, name)
        before_names = {_safe_str(sheet.Name) for sheet in sheets}
        try:
            # Copy as an in-workbook copy. ``Copy()`` with no arguments creates a
            # NEW workbook (not an in-workbook copy) on some Excel builds, so we
            # anchor the copy with ``Before`` to guarantee a same-workbook copy.
            anchor = sheets.Item(1)
            source.Copy(Before=anchor)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not copy sheet {name!r}.",
                context={"sheet": name, "to_end": to_end},
            ) from exc
        # Re-fetch the collection: the pre-copy COM object can be stale after Copy.
        sheets = wb.Sheets
        new_sheet = _find_copied_sheet(sheets, source, before_names, name=name, to_end=to_end)
        if target_name is not None:
            existing_names = {_safe_str(sheet.Name) for sheet in sheets}
            candidate = _validate_target_name(target_name, existing_names)
            try:
                new_sheet.Name = candidate
            except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
                raise InvalidRangeError(
                    f"Could not set the copied sheet name to {candidate!r}.",
                    context={"target_name": target_name, "reason": "invalid_target_name"},
                ) from exc
        return describe_sheet(wb, new_sheet)

    def delete_sheet(self, wb: Any, name: str) -> SheetInfo:
        """Delete ``name`` and return the DTO captured *before* deletion.

        The caller is responsible for the snapshot / approval / visible-sheet
        guards; the adapter executes the deletion only. The returned DTO reflects
        the pre-delete collection state so the caller can verify the postcondition
        (the sheet is gone).
        """

        sheets = wb.Sheets
        sheet = _require_sheet(sheets, name)
        info = describe_sheet(wb, sheet)
        try:
            sheet.Delete()
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not delete sheet {name!r}.",
                context={"sheet": name},
            ) from exc
        return info

    def set_visibility(self, wb: Any, name: str, visible: bool) -> SheetInfo:
        """Set a sheet's visibility and return the updated DTO.

        ``visible=True`` -> ``xlSheetVisible`` (0); ``visible=False`` ->
        ``xlSheetHidden`` (1). Hiding the last visible worksheet is rejected with
        a structured error (Excel requires at least one visible sheet).
        """

        sheets = wb.Sheets
        sheet = _require_sheet(sheets, name)
        if not visible and _visible_worksheet_count(wb) <= 1:
            # Only when the target is itself the single visible worksheet.
            target_code = _safe_int(sheet.Visible, _XL_SHEET_VISIBLE)
            if target_code == _XL_SHEET_VISIBLE:
                raise ExcelVbaMcpError(
                    "Cannot hide the last visible worksheet; at least one "
                    "worksheet must remain visible.",
                    context={"sheet": name},
                )
        sheet.Visible = _XL_SHEET_VISIBLE if visible else _XL_SHEET_HIDDEN
        return describe_sheet(wb, sheet)


def visible_worksheet_count(wb: Any) -> int:
    """Public helper: count visible worksheets in ``wb`` (guard input)."""

    return _visible_worksheet_count(wb)


def _validate_target_name(target_name: str, existing_names: set[str]) -> str:
    """Validate an explicit copy target name (non-empty, <=31 chars, unique).

    Rejects an empty / over-long / Excel-illegal / duplicate name with
    ``INVALID_RANGE`` (a caller argument error) before the rename is attempted, so
    a bad ``target_name`` never reaches the COM ``Name`` setter.
    """

    candidate = target_name.strip()
    if not candidate:
        raise InvalidRangeError(
            "Argument 'target_name' must be a non-empty string.",
            context={"target_name": target_name, "reason": "invalid_target_name"},
        )
    if len(candidate) > _MAX_SHEET_NAME_LEN:
        raise InvalidRangeError(
            f"Argument 'target_name' must be at most {_MAX_SHEET_NAME_LEN} characters.",
            context={"target_name": target_name, "reason": "invalid_target_name"},
        )
    if any(char in _INVALID_SHEET_NAME_CHARS for char in candidate):
        raise InvalidRangeError(
            "Argument 'target_name' contains a character Excel does not allow in a "
            "sheet name (: \\ / ? * [ ]).",
            context={"target_name": target_name, "reason": "invalid_target_name"},
        )
    if candidate.casefold() in {name.casefold() for name in existing_names}:
        raise InvalidRangeError(
            f"A sheet named {candidate!r} already exists; 'target_name' must be unique.",
            context={"target_name": target_name, "reason": "duplicate_target_name"},
        )
    return candidate


def _find_copied_sheet(
    sheets: Any,
    source: Any,
    before_names: set[str],
    *,
    name: str,
    to_end: bool,
) -> Any:
    """Return the newly created sheet proxy after ``Copy()`` (STA-local).

    The copy is the newly appeared sheet (Excel names it ``"<name> (2)"`` by
    default); it is resolved by diffing the pre-copy name set, with a fallback to
    the source name pattern. A copy that cannot be identified raises a structured
    error rather than guessing.
    """

    for sheet in sheets:
        sheet_name = _safe_str(sheet.Name)
        if sheet_name not in before_names:
            return sheet
    # Fallback: the copy shares the source name pattern and is not the source.
    source_name = _safe_str(source.Name)
    for sheet in sheets:
        sheet_name = _safe_str(sheet.Name)
        if sheet_name != source_name and sheet_name.startswith(source_name):
            return sheet
    raise ExcelVbaMcpError(
        "Could not determine the copied sheet after Copy().",
        context={"sheet": name, "to_end": to_end},
    )


__all__ = [
    "SheetsAdapter",
    "describe_sheet",
    "visible_worksheet_count",
]
