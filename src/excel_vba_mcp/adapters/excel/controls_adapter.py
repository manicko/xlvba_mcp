"""STA-local Controls adapter (task_030, spec §7.6 / A10).

:class:`ControlsAdapter` is the single place where Form Controls and ActiveX
controls are enumerated. It is a **pure STA-local helper**: every method is
synchronous, receives the live ``Workbook`` COM proxy (``Any``) and returns only
Pydantic DTOs / primitives, so results cross the single-STA worker boundary
safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never
imports ``pywin32`` at module import time and never calls ``.Activate`` /
``.Select`` on any object.

Form Controls vs ActiveX (spec §7.6, R09):

* **Form Controls** are shapes of type ``msoFormControl`` (``Shape.Type == 8`` in
  the ``MsoShapeType`` enumeration). They expose ``Shape.ControlFormat`` (the
  ``ControlFormat`` object) and ``Shape.OnAction`` — the assigned macro text,
  read **verbatim**. ``FormControl.Type`` (``xlFormControl``) distinguishes
  button / checkbox / option button / list / combo / edit / spin / scroll / group.
  The control's ``TextFrame`` carries its caption.
* **ActiveX controls** live in ``Worksheet.OLEObjects`` as ``OLEObject`` items
  with a ``ProgId`` (``Forms.CommandButton.1``, …). They do **not** have an
  ``OnAction`` — their behavior is implemented by event procedures in the
  worksheet's document module (e.g. ``CommandButton1_Click``). The adapter
  therefore reports a **probable** event-procedure association
  (``probable_procedure`` / ``macro_mapping=PROBABLE``) derived from the sheet
  code name + control name, and **never** claims an exact mapping or a confirmed
  click behaviour.

Diagnostics (spec acceptance criterion): every property read goes through a
``_safe_*`` helper that swallows the COM error, records a diagnostic note on the
returned DTO and leaves the field ``None``. A failure to read one property never
aborts the enumeration of the remaining controls.

All mutating methods assume the caller already validated the session (writable),
a snapshot exists and approval was granted; the adapter only executes the Excel
object-model calls on the STA thread.
"""

from __future__ import annotations

from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError, ObjectNotFoundError
from excel_vba_mcp.domain.models import ControlEntry
from excel_vba_mcp.domain.statuses import ControlFormType, ControlKind, MacroMappingStatus

# ``MsoShapeType.msoFormControl`` — the only shape type that is a Form Control.
_MSO_FORM_CONTROL = 8

# ``xlFormControl`` codes -> :class:`ControlFormType`.
_FORM_CONTROL_TYPES: dict[int, ControlFormType] = {
    0: ControlFormType.BUTTON,
    1: ControlFormType.CHECKBOX,
    2: ControlFormType.OPTION_BUTTON,
    3: ControlFormType.SPIN_BUTTON,
    4: ControlFormType.LIST_BOX,
    5: ControlFormType.GROUP_BOX,
    6: ControlFormType.SCROLL_BAR,
    7: ControlFormType.EDIT_BOX,
    8: ControlFormType.COMBO_BOX,
}

# Name hints used only when ``ControlFormat.Type`` / ``FormControlType`` are both
# unavailable (diagnostic fallback, recorded in ``notes``).
_FORM_CONTROL_NAME_HINTS: tuple[tuple[str, ControlFormType], ...] = (
    ("butto", ControlFormType.BUTTON),
    ("check", ControlFormType.CHECKBOX),
    ("option", ControlFormType.OPTION_BUTTON),
    ("radio", ControlFormType.OPTION_BUTTON),
    ("spin", ControlFormType.SPIN_BUTTON),
    ("list", ControlFormType.LIST_BOX),
    ("group", ControlFormType.GROUP_BOX),
    ("scroll", ControlFormType.SCROLL_BAR),
    ("edit", ControlFormType.EDIT_BOX),
    ("combo", ControlFormType.COMBO_BOX),
)

# ActiveX event procedure suffixes used to build the *probable* association.
# Excel's default document-module handler naming is
# ``<sheet_code_name>_<control_name>_<Event>``; only a small, well-known subset
# is listed here — this is a heuristic, never a confirmed mapping.
_ACTIVE_X_EVENT_SUFFIXES: tuple[str, ...] = (
    "click",
    "dblclick",
    "change",
    "gotfocus",
    "lostfocus",
    "keydown",
    "keyup",
    "keypress",
    "mousedown",
    "mouseup",
    "mousemove",
)


# ---------------------------------------------------------------------------
# Defensive COM readers (never raise; record a diagnostic note on failure)
# ---------------------------------------------------------------------------
def _safe_raw(obj: Any, name: str, default: Any = None) -> Any:
    """Read ``obj.name`` defensively (never raises, returns ``default``)."""

    if obj is None:
        return default
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_str(obj: Any, name: str, default: str = "") -> str:
    """Read a string property defensively (never raises)."""

    value = _safe_raw(obj, name, None)
    if value is None:
        return default
    if isinstance(value, str):
        return value
    try:
        return str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_float(obj: Any, name: str) -> float | None:
    """Read a numeric property defensively (``None`` when unavailable)."""

    value = _safe_raw(obj, name, None)
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_visible(shape: Any) -> bool | None:
    """Read ``Shape.Visible`` (``None`` when unavailable)."""

    value = _safe_raw(shape, "Visible", None)
    if value is None:
        return None
    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _safe_visible_flag(obj: Any, name: str) -> bool | None:
    """Read a boolean property defensively (``None`` when unavailable)."""

    value = _safe_raw(obj, name, None)
    if value is None:
        return None
    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _shape_type_code(shape: Any) -> int | None:
    """Read ``Shape.Type`` as an int (``None`` when unavailable)."""

    value = _safe_raw(shape, "Type", None)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _shape_id(shape: Any) -> str | int | None:
    """Read ``Shape.Id`` (an int when available, else ``None``)."""

    value = _safe_raw(shape, "Id", None)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _shape_text(shape: Any) -> str | None:
    """Best-effort read of a shape's text (never raises)."""

    frame2 = _safe_raw(shape, "TextFrame2", None)
    if frame2 is not None:
        text_range = _safe_raw(frame2, "TextRange", None)
        if text_range is not None:
            value = _safe_raw(text_range, "Text", None)
            if value is not None:
                return str(value)
    frame1 = _safe_raw(shape, "TextFrame", None)
    if frame1 is not None:
        characters = _safe_raw(frame1, "Characters", None)
        if characters is not None:
            value = _safe_raw(characters, "Text", None)
            if value is not None:
                return str(value)
    return None


def _form_control_type(shape: Any, control_format: Any, notes: list[str]) -> ControlFormType | None:
    """Resolve a Form Control's ``xlFormControl`` subtype (best effort).

    ``ControlFormat.Type`` is the documented source, but some hosts do not expose
    it reliably on a late-bound proxy. The shape's ``FormControlType`` property is
    tried as a fallback, and finally the shape ``Name`` is used as a hint
    (``Button 1`` -> button, ``Check 1`` -> checkbox, …). ``None`` means the
    subtype could not be determined (recorded as a note).
    """

    code = _read_int(control_format, "Type")
    if code is None:
        code = _read_int(shape, "FormControlType")
    if code is not None:
        if code in _FORM_CONTROL_TYPES:
            return _FORM_CONTROL_TYPES[code]
        notes.append(f"unknown FormControlType code {code}")
        return ControlFormType.UNKNOWN

    # Name-based fallback (diagnostic only).
    name = _safe_str(shape, "Name").strip().lower()
    for hint, form_type in _FORM_CONTROL_NAME_HINTS:
        if hint in name:
            notes.append(f"control_type inferred from name {name!r}")
            return form_type

    notes.append("ControlFormat.Type unavailable")
    return None


def _read_int(obj: Any, name: str) -> int | None:
    """Read an integer property defensively (``None`` when unavailable)."""

    value = _safe_raw(obj, name, None)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _probable_activex_procedure(sheet: Any, control_name: str) -> str | None:
    """Build a *probable* ActiveX event procedure name (heuristic, not confirmed).

    Excel's default document-module handler naming is
    ``<sheet_code_name>_<control_name>_<Event>``. This helper returns the first
    candidate (``..._Click``) built from the sheet's ``CodeName`` when it is
    readable; otherwise ``None``. It is intentionally a **guess**: the result is
    surfaced as ``macro_mapping=PROBABLE`` and must never be treated as a
    confirmed event mapping.
    """

    code_name = _safe_str(sheet, "CodeName").strip()
    if not code_name or not control_name.strip():
        return None
    return f"{code_name}_{control_name}_{_ACTIVE_X_EVENT_SUFFIXES[0].capitalize()}"


# ---------------------------------------------------------------------------
# DTO builders
# ---------------------------------------------------------------------------
def _read_linked_cell(shape: Any, notes: list[str]) -> str | None:
    """Read a Form Control's ``LinkedCell`` (best effort; ``None`` if unavailable)."""

    control_format = _safe_raw(shape, "ControlFormat", None)
    if control_format is None:
        notes.append("ControlFormat unavailable")
        return None
    return _safe_str(control_format, "LinkedCell") or None


def _describe_form_control(shape: Any, sheet_name: str | None) -> ControlEntry:
    """Build a :class:`ControlEntry` for one Form Control shape (never raises)."""

    notes: list[str] = []
    control_type = _form_control_type(shape, _safe_raw(shape, "ControlFormat", None), notes)
    linked_cell = _read_linked_cell(shape, notes)
    text = _shape_text(shape)
    return ControlEntry(
        name=_safe_str(shape, "Name") or "<unnamed>",
        kind=ControlKind.FORM_CONTROL,
        sheet_name=sheet_name,
        shape_id=_shape_id(shape),
        control_type=control_type,
        caption=text,
        text=text,
        linked_cell=linked_cell,
        on_action=_safe_str(shape, "OnAction") or None,
        visible=_safe_visible(shape),
        top=_safe_float(shape, "Top"),
        left=_safe_float(shape, "Left"),
        width=_safe_float(shape, "Width"),
        height=_safe_float(shape, "Height"),
        macro_mapping=MacroMappingStatus.UNKNOWN,
        notes=notes,
    )


def _describe_activex(ole_object: Any, sheet: Any, sheet_name: str | None) -> ControlEntry:
    """Build a :class:`ControlEntry` for one ActiveX ``OLEObject`` (never raises)."""

    notes: list[str] = []
    name = _safe_str(ole_object, "Name") or "<unnamed>"
    prog_id = _safe_str(ole_object, "ProgId") or None
    probable = _probable_activex_procedure(sheet, name)
    mapping = MacroMappingStatus.PROBABLE if probable is not None else MacroMappingStatus.UNKNOWN
    # ``OLEObject`` has no ``Visible`` / coordinates of its own; the shape that
    # hosts it does (``OLEObject.Shape``). Fall back to the OLEObject's own
    # ``Visible`` / ``Top`` / ``Left`` / ``Width`` / ``Height`` (the
    # ``OLEObject`` proxy exposes them) when the host shape is unavailable.
    shape = _safe_raw(ole_object, "Shape", None)
    visible = (
        _safe_visible(shape) if shape is not None else _safe_visible_flag(ole_object, "Visible")
    )
    if shape is not None:
        top = _safe_float(shape, "Top")
        left = _safe_float(shape, "Left")
        width = _safe_float(shape, "Width")
        height = _safe_float(shape, "Height")
        shape_id = _shape_id(shape)
    else:
        top = _safe_float(ole_object, "Top")
        left = _safe_float(ole_object, "Left")
        width = _safe_float(ole_object, "Width")
        height = _safe_float(ole_object, "Height")
        shape_id = None
        notes.append("host shape unavailable (coordinates read from OLEObject)")
    if prog_id is None:
        notes.append("ProgId unavailable")
    if probable is None:
        notes.append("probable event procedure unavailable (no sheet CodeName)")
    return ControlEntry(
        name=name,
        kind=ControlKind.ACTIVEX,
        sheet_name=sheet_name,
        shape_id=shape_id,
        prog_id=prog_id,
        caption=None,
        text=None,
        linked_cell=None,
        on_action=None,
        visible=visible,
        top=top,
        left=left,
        width=width,
        height=height,
        macro_mapping=mapping,
        probable_procedure=probable,
        notes=notes,
    )


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class ControlsAdapter:
    """STA-local Form Control + ActiveX inventory (no COM at import)."""

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def list_controls(self, wb: Any) -> list[ControlEntry]:
        """Return one :class:`ControlEntry` per Form Control and ActiveX control."""

        result: list[ControlEntry] = []
        for sheet in _iter_worksheets(wb):
            sheet_name = _safe_str(sheet, "Name") or None
            for shape in _iter_shapes(sheet):
                if _shape_type_code(shape) == _MSO_FORM_CONTROL:
                    result.append(_describe_form_control(shape, sheet_name))
            for ole_object in _iter_ole_objects(sheet):
                result.append(_describe_activex(ole_object, sheet, sheet_name))
        return result

    def get_control(self, wb: Any, name: str) -> ControlEntry:
        """Return the :class:`ControlEntry` for one control by name.

        The name is matched case-insensitively across Form Controls and ActiveX
        controls. A missing name or an ambiguous (duplicate) name raises a
        structured :class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError`.
        """

        target = _validate_name(name)
        matches = [
            entry for entry in self.list_controls(wb) if entry.name.strip().lower() == target
        ]
        if not matches:
            raise ObjectNotFoundError(
                f"Control {name!r} was not found in the workbook.",
                context={"name": name},
            )
        if len(matches) > 1:
            raise ExcelVbaMcpError(
                f"Control name {name!r} is ambiguous ({len(matches)} matches); refusing to guess.",
                context={"name": name, "reason": "ambiguous", "matches": len(matches)},
            )
        return matches[0]

    # ------------------------------------------------------------------
    # Mutation (executed on the STA thread)
    # ------------------------------------------------------------------
    def set_on_action(self, wb: Any, name: str, macro_name: str) -> ControlEntry:
        """Assign ``OnAction`` to a **Form Control** shape and return the re-read DTO.

        ``OnAction`` exists only on Form Controls (and plain shapes); an ActiveX
        ``OLEObject`` has no ``OnAction`` property, so setting one is rejected with
        a structured error instead of silently writing to a non-existent member.
        The macro is **not executed** — only the property is written.
        """

        target = _validate_name(name)
        shape = self._resolve_form_control(wb, target)
        try:
            shape.OnAction = macro_name
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not set OnAction for control {name!r}.",
                context={"name": name, "macro_name": macro_name},
            ) from exc
        return self.get_control(wb, name)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _resolve_form_control(self, wb: Any, target_name: str) -> Any:
        """Resolve a name to a Form Control shape proxy (ActiveX is rejected)."""

        ole_names: set[str] = set()
        for sheet in _iter_worksheets(wb):
            for ole_object in _iter_ole_objects(sheet):
                ole_names.add(_safe_str(ole_object, "Name").strip().lower())
            for shape in _iter_shapes(sheet):
                if _safe_str(shape, "Name").strip().lower() != target_name:
                    continue
                if _shape_type_code(shape) == _MSO_FORM_CONTROL:
                    return shape
                if target_name in ole_names:
                    raise ExcelVbaMcpError(
                        f"Control {target_name!r} is an ActiveX control; it has no OnAction "
                        "property (event procedures live in the document module).",
                        context={"name": target_name, "kind": "activex"},
                    )
                raise ExcelVbaMcpError(
                    f"Shape {target_name!r} is not a Form Control and has no OnAction property.",
                    context={"name": target_name, "kind": "shape"},
                )
        raise ObjectNotFoundError(
            f"Control {target_name!r} was not found in the workbook.",
            context={"name": target_name},
        )


# ---------------------------------------------------------------------------
# Module-level COM collection helpers (defensive; never raise)
# ---------------------------------------------------------------------------
def _iter_worksheets(wb: Any) -> list[Any]:
    """Return the workbook's worksheets as a list (best effort)."""

    worksheets = _safe_raw(wb, "Worksheets", None)
    if worksheets is None:
        return []
    try:
        count = int(worksheets.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        return []
    result: list[Any] = []
    for index in range(1, count + 1):
        try:
            result.append(worksheets.Item(index))
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
    return result


def _iter_shapes(sheet: Any) -> list[Any]:
    """Return a worksheet's ``Shapes`` collection as a list (best effort)."""

    shapes = _safe_raw(sheet, "Shapes", None)
    if shapes is None:
        return []
    try:
        count = int(shapes.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        return []
    result: list[Any] = []
    for index in range(1, count + 1):
        try:
            result.append(shapes.Item(index))
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
    return result


def _iter_ole_objects(sheet: Any) -> list[Any]:
    """Return a worksheet's ``OLEObjects`` collection as a list (best effort)."""

    ole_objects = _safe_raw(sheet, "OLEObjects", None)
    if ole_objects is None:
        return []
    try:
        count = int(ole_objects.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        return []
    result: list[Any] = []
    for index in range(1, count + 1):
        try:
            result.append(ole_objects.Item(index))
        except Exception:  # noqa: BLE001 - defensive per-item COM read
            continue
    return result


def _validate_name(name: str) -> str:
    """Validate a control name argument."""

    if not isinstance(name, str) or not name.strip():
        raise ExcelVbaMcpError(
            "Argument 'name' must be a non-empty string.",
            context={"name": name},
        )
    return name.strip().lower()


__all__ = ["ControlsAdapter"]
