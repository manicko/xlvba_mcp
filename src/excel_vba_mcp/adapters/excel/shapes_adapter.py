"""STA-local Shapes adapter (task_030, spec §7.6 / A10).

:class:`ShapesAdapter` is the single place where the Excel ``Worksheet.Shapes``
collection is touched. It is a **pure STA-local helper**: every method is
synchronous, receives the live ``Workbook`` COM proxy (``Any``) and returns only
Pydantic DTOs / primitives, so results cross the single-STA worker boundary
safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never
imports ``pywin32`` at module import time and never calls ``.Activate`` /
``.Select`` on any object.

What is read (spec §7.6 Shapes):

* ``Shape.Name`` / ``Shape.Type`` (``msoShapeType``) / ``Shape.Visible`` /
  ``Shape.Top`` / ``Shape.Left`` / ``Shape.Width`` / ``Shape.Height`` /
  ``Shape.AlternativeText``;
* ``Shape.OnAction`` — the assigned macro name, read **verbatim**. The macro is
  never executed and no ``Application.Run`` happens here;
* ``Shape.TextFrame2.TextRange.Text`` (best effort) for the shape's text;
* ``Shape.ControlFormat`` is probed (never dereferenced) so a Form Control shape
  is reported with ``control_format=True`` while plain shapes / ActiveX report
  ``None``.

Defensive reads (spec acceptance criterion "missing properties are diagnosable,
not unhandled COM errors"): every property read goes through a ``_safe_*`` helper
that swallows the COM error, records a diagnostic note on the returned DTO and
leaves the field ``None``. A failure to read one property never aborts the
enumeration of the remaining shapes.

All mutating methods assume the caller already validated the session (writable),
a snapshot exists and approval was granted; the adapter only executes the Excel
object-model calls on the STA thread.
"""

from __future__ import annotations

from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError, ObjectNotFoundError
from excel_vba_mcp.domain.models import ShapeInfo
from excel_vba_mcp.domain.statuses import ShapeKind

# ---------------------------------------------------------------------------
# ``MsoShapeType`` codes (subset relevant to a Shape inventory).
# ---------------------------------------------------------------------------
# Authoritative ``MsoShapeType`` values used by Excel:
#   msoAutoShape=1, msoCallout=2, msoChart=3, msoComment=4, msoFreeform=5,
#   msoGroup=6, msoEmbeddedOLEObject=7, msoFormControl=8, msoOleControlObject=9,
#   msoLine=10, msoTextBox=17, msoPicture=13, msoPlaceholder=11, ...
# The table maps each relevant code to its :class:`ShapeKind`; codes that this
# inventory does not distinguish collapse to ``SHAPE``.
_MSO_SHAPE_TYPE: dict[int, ShapeKind] = {
    1: ShapeKind.AUTOSHAPE,  # msoAutoShape
    2: ShapeKind.AUTOSHAPE,  # msoCallout
    3: ShapeKind.CHART,  # msoChart
    5: ShapeKind.FREEFORM,  # msoFreeform
    6: ShapeKind.GROUP,  # msoGroup
    7: ShapeKind.SHAPE,  # msoEmbeddedOLEObject (embedded document host)
    8: ShapeKind.FORM_CONTROL,  # msoFormControl
    9: ShapeKind.ACTIVEX,  # msoOleControlObject (ActiveX host shape)
    10: ShapeKind.LINE,  # msoLine
    11: ShapeKind.PLACEHOLDER,  # msoPlaceholder
    13: ShapeKind.PICTURE,  # msoPicture
    17: ShapeKind.TEXT_BOX,  # msoTextBox
    21: ShapeKind.TEXT_BOX,  # msoTextEffect (rendered as text)
}

# Authoritative ``MsoShapeType`` names -> kind. Late-binding proxies expose the
# numeric code (``Shape.Type``), which the table above maps; when a build reports
# an unexpected code the adapter falls back to these names.
_MSO_SHAPE_NAME_KINDS: dict[str, ShapeKind] = {
    "msoautoshape": ShapeKind.AUTOSHAPE,
    "msocallout": ShapeKind.AUTOSHAPE,
    "msochart": ShapeKind.CHART,
    "msofreeform": ShapeKind.FREEFORM,
    "msogroup": ShapeKind.GROUP,
    "msopccontrol": ShapeKind.ACTIVEX,
    "msoolecontrolobject": ShapeKind.ACTIVEX,
    "msoformcontrol": ShapeKind.FORM_CONTROL,
    "msoline": ShapeKind.LINE,
    "msoplaceholder": ShapeKind.PLACEHOLDER,
    "msopicture": ShapeKind.PICTURE,
    "msotextbox": ShapeKind.TEXT_BOX,
}

# ``Shape.Name`` hints used only when neither the code nor the type name resolves
# the kind (e.g. a bare late-bound proxy with no ``Type``).
_MSO_NAME_HINTS: tuple[tuple[str, ShapeKind], ...] = (
    ("textbox", ShapeKind.TEXT_BOX),
    ("formcontrol", ShapeKind.FORM_CONTROL),
    ("olecontrol", ShapeKind.ACTIVEX),
    ("autoshape", ShapeKind.AUTOSHAPE),
    ("group", ShapeKind.GROUP),
    ("picture", ShapeKind.PICTURE),
    ("line", ShapeKind.LINE),
    ("chart", ShapeKind.CHART),
    ("placeholder", ShapeKind.PLACEHOLDER),
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


def _shape_type_code(shape: Any) -> int | None:
    """Read ``Shape.Type`` as an int (``None`` when unavailable)."""

    value = _safe_raw(shape, "Type", None)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _classify_shape(shape: Any, code: int | None) -> ShapeKind:
    """Map a shape to a :class:`ShapeKind` from its ``Type`` code."""

    if code is not None and code in _MSO_SHAPE_TYPE:
        return _MSO_SHAPE_TYPE[code]
    # Late-binding fallback 1: the type's *name* (``Shape.Type`` string form) is
    # unambiguous even when a build reports a code outside the table above.
    type_name = _safe_str(shape, "Type").strip().lower()
    for name_key, kind in _MSO_SHAPE_NAME_KINDS.items():
        if type_name == name_key:
            return kind
    # Late-binding fallback 2: classify by the shape's ``Name`` (ActiveX hosts and
    # Form Controls typically embed their object type in the name).
    name = _safe_str(shape, "Name").strip().lower()
    for hint, kind in _MSO_NAME_HINTS:
        if hint in name:
            return kind
    return ShapeKind.UNKNOWN


def _shape_id(shape: Any) -> str | int | None:
    """Read ``Shape.Id`` (an int when available, else ``None``)."""

    value = _safe_raw(shape, "Id", None)
    if value is None:
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _read_text(shape: Any, notes: list[str]) -> str | None:
    """Best-effort read of a shape's text (never raises).

    ``TextFrame2.TextRange.Text`` is tried first (works for AutoShapes and text
    boxes); a legacy ``TextFrame.Characters.Text`` fallback covers older hosts.
    ``None`` means "no text frame / text not available"; an empty string means a
    text frame exists but is empty (meaningful, preserved).
    """

    frame2 = _safe_raw(shape, "TextFrame2", None)
    if frame2 is not None:
        text_range = _safe_raw(frame2, "TextRange", None)
        if text_range is not None:
            value = _safe_raw(text_range, "Text", None)
            if value is not None:
                return str(value)
        notes.append("text unavailable (TextFrame2 has no readable TextRange)")

    frame1 = _safe_raw(shape, "TextFrame", None)
    if frame1 is not None:
        characters = _safe_raw(frame1, "Characters", None)
        if characters is not None:
            value = _safe_raw(characters, "Text", None)
            if value is not None:
                return str(value)
        notes.append("text unavailable (TextFrame has no readable Characters)")

    return None


def _has_control_format(shape: Any) -> bool | None:
    """Return whether ``Shape.ControlFormat`` is exposed (never dereferenced).

    Probing the property is enough: a Form Control shape exposes a
    ``ControlFormat`` object, a plain shape / ActiveX shape raises. ``None`` is
    returned when the probe cannot decide.
    """

    try:
        value = shape.ControlFormat
    except Exception:  # noqa: BLE001 - not a Form Control / unavailable
        return None
    return value is not None


def _safe_top_left_cell(shape: Any) -> str | None:
    """Read ``Shape.TopLeftCell.Address`` (diagnostic; ``None`` if unavailable)."""

    cell = _safe_raw(shape, "TopLeftCell", None)
    if cell is None:
        return None
    address = _safe_raw(cell, "Address", None)
    if address is None:
        return None
    return str(address)


def _describe_shape(shape: Any, sheet_name: str | None) -> ShapeInfo:
    """Build a :class:`ShapeInfo` for one ``Shape`` (STA-local, never raises)."""

    notes: list[str] = []
    code = _shape_type_code(shape)
    kind = _classify_shape(shape, code)
    text = _read_text(shape, notes)
    return ShapeInfo(
        name=_safe_str(shape, "Name") or "<unnamed>",
        kind=kind,
        shape_id=_shape_id(shape),
        sheet_name=sheet_name,
        on_action=_safe_str(shape, "OnAction") or None,
        visible=_safe_visible(shape),
        top=_safe_float(shape, "Top"),
        left=_safe_float(shape, "Left"),
        width=_safe_float(shape, "Width"),
        height=_safe_float(shape, "Height"),
        text=text,
        alternative_text=_safe_str(shape, "AlternativeText") or None,
        control_format=_has_control_format(shape),
        notes=notes,
    )


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class ShapesAdapter:
    """STA-local Shapes inventory + guarded ``OnAction`` (no COM at import)."""

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def list_shapes(self, wb: Any) -> list[ShapeInfo]:
        """Return one :class:`ShapeInfo` per shape of every worksheet."""

        result: list[ShapeInfo] = []
        for sheet in _iter_worksheets(wb):
            sheet_name = _safe_str(sheet, "Name") or None
            for shape in _iter_shapes(sheet):
                result.append(_describe_shape(shape, sheet_name))
        return result

    def get_shape(self, wb: Any, name: str) -> ShapeInfo:
        """Return the :class:`ShapeInfo` for one shape by name.

        The name is matched case-insensitively across every worksheet. A missing
        name or an ambiguous (duplicate) name raises a structured
        :class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError`.
        """

        target = _validate_name(name)
        matches: list[ShapeInfo] = []
        for sheet in _iter_worksheets(wb):
            sheet_name = _safe_str(sheet, "Name") or None
            for shape in _iter_shapes(sheet):
                info = _describe_shape(shape, sheet_name)
                if info.name.strip().lower() == target:
                    matches.append(info)
        if not matches:
            raise ObjectNotFoundError(
                f"Shape {name!r} was not found in the workbook.",
                context={"name": name},
            )
        if len(matches) > 1:
            raise ExcelVbaMcpError(
                f"Shape name {name!r} is ambiguous ({len(matches)} matches); refusing to guess.",
                context={"name": name, "reason": "ambiguous", "matches": len(matches)},
            )
        return matches[0]

    # ------------------------------------------------------------------
    # Mutation (executed on the STA thread)
    # ------------------------------------------------------------------
    def set_on_action(self, wb: Any, name: str, macro_name: str) -> ShapeInfo:
        """Assign ``shape.OnAction = macro_name`` and return the re-read DTO.

        The macro is **not executed**: only the ``OnAction`` property is written.
        An empty / ``None`` ``macro_name`` clears the assignment (Excel stores the
        empty string).
        """

        target = _validate_name(name)
        shape = self._resolve_shape(wb, target)
        try:
            shape.OnAction = macro_name
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not set OnAction for shape {name!r}.",
                context={"name": name, "macro_name": macro_name},
            ) from exc
        return self.get_shape(wb, name)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _resolve_shape(self, wb: Any, target_name: str) -> Any:
        """Resolve a shape name to its live ``Shape`` proxy or raise."""

        matches: list[Any] = []
        for sheet in _iter_worksheets(wb):
            for shape in _iter_shapes(sheet):
                if _safe_str(shape, "Name").strip().lower() == target_name:
                    matches.append(shape)
        if not matches:
            raise ObjectNotFoundError(
                f"Shape {target_name!r} was not found in the workbook.",
                context={"name": target_name},
            )
        if len(matches) > 1:
            raise ExcelVbaMcpError(
                f"Shape name {target_name!r} is ambiguous ({len(matches)} matches); "
                "refusing to guess.",
                context={"name": target_name, "reason": "ambiguous", "matches": len(matches)},
            )
        return matches[0]


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


def _validate_name(name: str) -> str:
    """Validate a shape name argument."""

    if not isinstance(name, str) or not name.strip():
        raise ExcelVbaMcpError(
            "Argument 'name' must be a non-empty string.",
            context={"name": name},
        )
    return name.strip().lower()


__all__ = ["ShapesAdapter"]
