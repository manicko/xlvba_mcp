"""STA-local defined-name (Named Range) adapter (task_028, spec §7.7 / §6.5).

:class:`NamesAdapter` is the single place where the Excel ``Names`` /
``Name`` object model is touched. It is a **pure STA-local helper**: every method
is synchronous, receives the live ``Workbook`` COM proxy (``Any``) and returns
only Pydantic DTOs / primitives, so results cross the single-STA worker boundary
safely (:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never
imports ``pywin32`` at module import time and never calls ``.Activate`` /
``.Select`` on any object.

Scope model (spec §7.7):

* ``Workbook.Names`` enumerates **all** defined names (workbook-scoped and
  worksheet-scoped); ``Worksheet.Names`` enumerates only the names scoped to
  that worksheet. A name is identified by ``(name, scope)`` — and for worksheet
  scope also by the owning sheet — so the same identifier may legally exist once
  at workbook level and once per worksheet.
* ``Name.Parent`` is the ``Workbook`` for a workbook-scoped name and the
  ``Worksheet`` for a worksheet-scoped name. The adapter distinguishes the two
  by probing ``Parent.FullName`` (only a ``Workbook`` exposes it).

Non-range targets (spec §7.7): not every name refers to a range. The adapter
first reads ``RefersTo`` / ``RefersToLocal`` / ``Visible`` / ``NameLocal`` and
only then attempts ``Name.RefersToRange`` inside a **separate error-handling
branch**. When that succeeds the name is a ``RANGE`` target and the sheet +
address are reported; when it raises, the ``RefersTo`` string is classified
heuristically (constant / formula / external / error / unknown) **without
evaluating** the formula (no ``Application.Evaluate`` and no Python ``eval``).

Conflicts (spec §7.7): creating or updating a name detects a duplicate in the
requested scope and a cross-scope collision (the same identifier already exists
at workbook level or on another worksheet), raising a structured
:class:`~excel_vba_mcp.domain.errors.ExcelVbaMcpError` whose ``context`` carries
``reason`` and the existing scopes.

All mutating methods assume the caller already validated the session
(writable), a snapshot exists and approval was granted; the adapter only executes
the Excel object-model calls on the STA thread.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import Any

from excel_vba_mcp.domain.errors import ExcelVbaMcpError
from excel_vba_mcp.domain.models import DefinedNameInfo
from excel_vba_mcp.domain.statuses import NameScope, NameTargetKind

# Excel error literals that can appear inside a name's ``RefersTo`` (spec §7.7:
# a name may have a broken reference). Kept as plain text; never evaluated.
_ERROR_LITERALS: frozenset[str] = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
    }
)

# A numeric literal (integer / decimal / scientific), optionally signed.
_NUMERIC_LITERAL = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$")

# Boolean literals accepted as defined-name constants.
_BOOLEAN_LITERALS: frozenset[str] = frozenset({"TRUE", "FALSE"})


# ---------------------------------------------------------------------------
# Defensive COM readers (never raise)
# ---------------------------------------------------------------------------
def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively (never raises)."""

    if value is None:
        return default
    try:
        text = value
    except Exception:  # noqa: BLE001 - defensive COM read
        return default
    if text is None:
        return default
    if isinstance(text, str):
        return text
    try:
        return str(text)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_prop(obj: Any, name: str, default: Any = None) -> Any:
    """Read ``obj.name`` defensively (never raises, returns ``default``)."""

    if obj is None:
        return default
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_visible(name_obj: Any) -> bool:
    """Read ``Name.Visible`` (default ``True`` when unavailable)."""

    value = _safe_prop(name_obj, "Visible", True)
    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return True


def _iter_collection(collection: Any) -> Iterator[Any]:
    """Yield the items of a COM collection via ``Count`` / ``Item`` (defensive).

    Falls back to Python iteration when ``Count`` is unavailable. Never raises:
    a partially-initialized or non-collection object simply yields nothing.
    """

    if collection is None:
        return
    try:
        count = int(collection.Count)
    except Exception:  # noqa: BLE001 - defensive COM read
        count = 0
    if count > 0:
        for index in range(1, count + 1):
            try:
                yield collection.Item(index)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
        return
    try:
        yield from collection
    except Exception:  # noqa: BLE001 - defensive COM read
        return


def _parent_is_workbook(parent: Any) -> bool:
    """Return whether ``Name.Parent`` is the owning ``Workbook``.

    Only a ``Workbook`` exposes ``FullName``; a ``Worksheet`` does not, so a
    successful read is the authoritative scope test (COM proxies are compared by
    property, never by identity).
    """

    if parent is None:
        return False
    try:
        _ = parent.FullName
        return True
    except Exception:  # noqa: BLE001 - Worksheet has no FullName
        return False


# ---------------------------------------------------------------------------
# Non-range classification (no evaluation)
# ---------------------------------------------------------------------------
def _strip_equals(text: str) -> str:
    """Drop the leading ``=`` Excel uses for a name's ``RefersTo`` formula."""

    stripped = text.strip()
    return stripped[1:] if stripped.startswith("=") else stripped


def _is_constant(body: str) -> bool:
    """Return whether a non-range ``RefersTo`` body is a literal constant."""

    if not body:
        return False
    if len(body) >= 2 and body.startswith('"') and body.endswith('"'):
        return True
    if body.upper() in _BOOLEAN_LITERALS:
        return True
    return bool(_NUMERIC_LITERAL.match(body))


def classify_target_kind(refers_to: str | None) -> NameTargetKind:
    """Classify a non-range ``RefersTo`` string (spec §7.7, no evaluation).

    Heuristics only — the formula/expression is never evaluated:

    * ``ERROR`` — the reference contains an Excel error literal (``#REF!`` …);
    * ``EXTERNAL`` — the reference points into another workbook (``[Book]Sheet!…``);
    * ``CONSTANT`` — a numeric / string / boolean literal;
    * ``FORMULA`` — any other expression / function call;
    * ``UNKNOWN`` — an empty or unreadable reference.
    """

    if not refers_to:
        return NameTargetKind.UNKNOWN
    body = _strip_equals(refers_to)
    if not body:
        return NameTargetKind.UNKNOWN
    if any(literal in body for literal in _ERROR_LITERALS) or body.startswith("#"):
        return NameTargetKind.ERROR
    if "[" in body and "]" in body:
        return NameTargetKind.EXTERNAL
    if _is_constant(body):
        return NameTargetKind.CONSTANT
    return NameTargetKind.FORMULA


# ---------------------------------------------------------------------------
# Scope helpers
# ---------------------------------------------------------------------------
def _coerce_scope(scope: NameScope | str) -> NameScope:
    """Coerce ``scope`` to :class:`NameScope` (accepts its string value)."""

    if isinstance(scope, NameScope):
        return scope
    try:
        return NameScope(str(scope).strip().lower())
    except ValueError as exc:
        raise ExcelVbaMcpError(
            f"Unknown name scope {scope!r}; expected 'workbook' or 'worksheet'.",
            context={"scope": str(scope)},
        ) from exc


def _require_sheet_for_scope(scope: NameScope, sheet: str | None) -> str | None:
    """Validate the sheet argument for the requested scope."""

    if scope is NameScope.WORKBOOK:
        return None
    if not isinstance(sheet, str) or not sheet.strip():
        raise ExcelVbaMcpError(
            "Argument 'sheet' is required for a worksheet-scoped name.",
            context={"scope": scope.value, "sheet": sheet},
        )
    return sheet


def _validate_name(name: str) -> str:
    """Validate a defined-name identifier."""

    if not isinstance(name, str) or not name.strip():
        raise ExcelVbaMcpError(
            "Argument 'name' must be a non-empty string.",
            context={"name": name},
        )
    return name


def _normalize_refers_to(refers_to: str) -> str:
    """Validate and normalize a name's ``RefersTo`` formula."""

    if not isinstance(refers_to, str) or not refers_to.strip():
        raise ExcelVbaMcpError(
            "Argument 'refers_to' must be a non-empty formula string.",
            context={"refers_to": refers_to},
        )
    text = refers_to.strip()
    # Excel stores a name's RefersTo with a leading '='; accept both forms.
    return text if text.startswith("=") else f"={text}"


def _require_worksheet(wb: Any, sheet: str) -> Any:
    """Resolve ``sheet`` to a worksheet proxy or raise a structured error."""

    try:
        return wb.Worksheets.Item(sheet)
    except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
        raise ExcelVbaMcpError(
            f"Worksheet {sheet!r} was not found in the workbook.",
            context={"sheet": sheet},
        ) from exc


def _range_target_info(rng: Any) -> tuple[str | None, str | None]:
    """Return ``(sheet_name, address)`` for a resolved ``RefersToRange``."""

    sheet_name: str | None = None
    worksheet = _safe_prop(rng, "Worksheet", None)
    if worksheet is not None:
        sheet_name = _safe_str(_safe_prop(worksheet, "Name", None)) or None
    address = _safe_str(_safe_prop(rng, "Address", None)) or None
    return sheet_name, address


def _strip_sheet_prefix(raw: str, sheet_name: str | None) -> str:
    """Strip Excel's ``Sheet!`` qualifier from a worksheet-scoped name.

    Excel reports the ``Name`` / ``NameLocal`` of a worksheet-scoped name as
    ``SheetName!Name`` (the sheet name is quoted when it contains spaces). The
    adapter exposes the unqualified identifier plus the scope's ``sheet_name``.
    """

    if not raw:
        return raw
    text = raw.strip()
    if sheet_name:
        prefix = f"{sheet_name.strip()}!"
        if text.lower().startswith(prefix.lower()):
            return text[len(prefix) :]
    if text.startswith("'"):
        quote_end = text.find("'", 1)
        if quote_end > 0 and text[quote_end + 1 : quote_end + 2] == "!":
            return text[quote_end + 2 :]
    if "!" in text:
        return text.rsplit("!", 1)[-1]
    return text


def _name_identifier(name_obj: Any, scope: NameScope, sheet_name: str | None) -> str:
    """Return the unqualified identifier of a ``Name`` object."""

    raw = _safe_str(_safe_prop(name_obj, "Name", None))
    if scope is NameScope.WORKSHEET:
        return _strip_sheet_prefix(raw, sheet_name)
    return raw


def _describe_name(
    name_obj: Any,
    *,
    scope: NameScope,
    sheet_name: str | None,
) -> DefinedNameInfo:
    """Build a :class:`DefinedNameInfo` for one ``Name`` object (STA-local).

    ``RefersToRange`` is attempted in its own error-handling branch: success
    yields a ``RANGE`` target with sheet + address; failure falls back to a
    heuristic classification of the ``RefersTo`` string (no evaluation).
    """

    refers_to = _safe_str(_safe_prop(name_obj, "RefersTo", None)) or None
    refers_to_local = _safe_str(_safe_prop(name_obj, "RefersToLocal", None)) or None
    name_local = _safe_str(_safe_prop(name_obj, "NameLocal", None)) or None
    if scope is NameScope.WORKSHEET and name_local is not None:
        name_local = _strip_sheet_prefix(name_local, sheet_name) or None

    target_kind = NameTargetKind.UNKNOWN
    target_sheet: str | None = None
    target_address: str | None = None
    try:
        rng = name_obj.RefersToRange
    except Exception:  # noqa: BLE001 - not a range: classify without evaluating
        rng = None
    if rng is not None:
        target_kind = NameTargetKind.RANGE
        target_sheet, target_address = _range_target_info(rng)
    else:
        target_kind = classify_target_kind(refers_to)

    return DefinedNameInfo(
        name=_name_identifier(name_obj, scope, sheet_name),
        scope=scope,
        name_local=name_local,
        refers_to=refers_to,
        refers_to_local=refers_to_local,
        visible=_safe_visible(name_obj),
        sheet_name=sheet_name,
        target_kind=target_kind,
        target_sheet=target_sheet,
        target_address=target_address,
    )


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------
class NamesAdapter:
    """STA-local defined-name inventory + scoped CRUD (no COM at import)."""

    # ------------------------------------------------------------------
    # Inventory (read-only)
    # ------------------------------------------------------------------
    def list_names(self, wb: Any) -> list[DefinedNameInfo]:
        """Return one :class:`DefinedNameInfo` per defined name (all scopes).

        Workbook-scoped names come from ``wb.Names`` (filtered by
        ``Parent``); worksheet-scoped names come from each ``ws.Names``. The two
        passes are de-duplicated by ``(name, scope, sheet)``.
        """

        result: list[DefinedNameInfo] = []
        for name_obj, scope, sheet_name in self._iter_entries(wb):
            result.append(_describe_name(name_obj, scope=scope, sheet_name=sheet_name))
        return result

    def get_name(
        self,
        wb: Any,
        name: str,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Return the :class:`DefinedNameInfo` for one name in ``scope``."""

        resolved_scope = _coerce_scope(scope)
        sheet_name = _require_sheet_for_scope(resolved_scope, sheet)
        name_obj = self._resolve_name(wb, name, resolved_scope, sheet_name)
        return _describe_name(name_obj, scope=resolved_scope, sheet_name=sheet_name)

    # ------------------------------------------------------------------
    # Mutations (executed on the STA thread)
    # ------------------------------------------------------------------
    def create_name(
        self,
        wb: Any,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Create a defined name in ``scope`` and return its DTO.

        A duplicate in the requested scope and a cross-scope collision (the same
        identifier already exists at another scope / sheet) are rejected with a
        structured error before anything is added (spec §7.7 conflict checks).
        """

        resolved_scope = _coerce_scope(scope)
        sheet_name = _require_sheet_for_scope(resolved_scope, sheet)
        _validate_name(name)
        normalized = _normalize_refers_to(refers_to)
        self._guard_no_conflict(wb, name, resolved_scope, sheet_name)

        collection = self._names_collection(wb, resolved_scope, sheet_name)
        try:
            collection.Add(Name=name, RefersTo=normalized)
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not create defined name {name!r} in {resolved_scope.value} scope.",
                context={
                    "name": name,
                    "scope": resolved_scope.value,
                    "sheet": sheet_name,
                    "refers_to": normalized,
                },
            ) from exc
        return self.get_name(wb, name, resolved_scope, sheet_name)

    def update_name(
        self,
        wb: Any,
        name: str,
        refers_to: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Update an existing name's ``RefersTo`` in ``scope`` and return its DTO.

        The name's scope is preserved (the update never moves a name between
        scopes). ``RefersTo`` is written directly; no formula is evaluated.
        """

        resolved_scope = _coerce_scope(scope)
        sheet_name = _require_sheet_for_scope(resolved_scope, sheet)
        normalized = _normalize_refers_to(refers_to)
        name_obj = self._resolve_name(wb, name, resolved_scope, sheet_name)
        try:
            name_obj.RefersTo = normalized
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not update defined name {name!r} in {resolved_scope.value} scope.",
                context={
                    "name": name,
                    "scope": resolved_scope.value,
                    "sheet": sheet_name,
                    "refers_to": normalized,
                },
            ) from exc
        return self.get_name(wb, name, resolved_scope, sheet_name)

    def delete_name(
        self,
        wb: Any,
        name: str,
        *,
        scope: NameScope | str,
        sheet: str | None = None,
    ) -> DefinedNameInfo:
        """Delete a name in ``scope`` and return its pre-delete DTO."""

        resolved_scope = _coerce_scope(scope)
        sheet_name = _require_sheet_for_scope(resolved_scope, sheet)
        name_obj = self._resolve_name(wb, name, resolved_scope, sheet_name)
        info = _describe_name(name_obj, scope=resolved_scope, sheet_name=sheet_name)
        try:
            name_obj.Delete()
        except Exception as exc:  # noqa: BLE001 - convert COM error to domain error
            raise ExcelVbaMcpError(
                f"Could not delete defined name {name!r} in {resolved_scope.value} scope.",
                context={"name": name, "scope": resolved_scope.value, "sheet": sheet_name},
            ) from exc
        return info

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _iter_entries(self, wb: Any) -> Iterator[tuple[Any, NameScope, str | None]]:
        """Yield ``(name_obj, scope, sheet_name)`` for every defined name.

        Workbook-scoped names are taken from ``wb.Names`` (filtered by
        ``Parent``); worksheet-scoped names from each ``ws.Names``. Duplicates
        are removed by ``(name, scope, sheet)``.
        """

        seen: set[tuple[str, str, str]] = set()

        for name_obj in _iter_collection(_safe_prop(wb, "Names", None)):
            if not _parent_is_workbook(_safe_prop(name_obj, "Parent", None)):
                continue
            key = (
                _name_identifier(name_obj, NameScope.WORKBOOK, None).strip().lower(),
                NameScope.WORKBOOK.value,
                "",
            )
            if key in seen:
                continue
            seen.add(key)
            yield name_obj, NameScope.WORKBOOK, None

        worksheets = _safe_prop(wb, "Worksheets", None)
        for sheet in _iter_collection(worksheets):
            sheet_name = _safe_str(_safe_prop(sheet, "Name", None))
            for name_obj in _iter_collection(_safe_prop(sheet, "Names", None)):
                key = (
                    _name_identifier(name_obj, NameScope.WORKSHEET, sheet_name).strip().lower(),
                    NameScope.WORKSHEET.value,
                    sheet_name.strip().lower(),
                )
                if key in seen:
                    continue
                seen.add(key)
                yield name_obj, NameScope.WORKSHEET, sheet_name or None

    def _resolve_name(
        self,
        wb: Any,
        name: str,
        scope: NameScope,
        sheet: str | None,
    ) -> Any:
        """Resolve ``(name, scope, sheet)`` to one ``Name`` object or raise."""

        target = name.strip().lower()
        matches: list[Any] = []
        for name_obj, entry_scope, entry_sheet in self._iter_entries(wb):
            if entry_scope is not scope:
                continue
            if _name_identifier(name_obj, entry_scope, entry_sheet).strip().lower() != target:
                continue
            if (
                scope is NameScope.WORKSHEET
                and (entry_sheet or "").strip().lower() != (sheet or "").strip().lower()
            ):
                continue
            matches.append(name_obj)

        if not matches:
            raise ExcelVbaMcpError(
                f"Defined name {name!r} was not found in {scope.value} scope.",
                context={"name": name, "scope": scope.value, "sheet": sheet},
            )
        if len(matches) > 1:
            raise ExcelVbaMcpError(
                f"Defined name {name!r} is ambiguous in {scope.value} scope "
                f"(matched {len(matches)} entries); refusing to guess.",
                context={
                    "name": name,
                    "scope": scope.value,
                    "sheet": sheet,
                    "reason": "ambiguous",
                },
            )
        return matches[0]

    def _guard_no_conflict(
        self,
        wb: Any,
        name: str,
        scope: NameScope,
        sheet: str | None,
    ) -> None:
        """Reject a duplicate (same scope) or a cross-scope name collision."""

        target = name.strip().lower()
        existing: list[tuple[str, str | None]] = []
        for info in self.list_names(wb):
            if info.name.strip().lower() != target:
                continue
            existing.append((info.scope.value, info.sheet_name))

        same_scope = any(
            entry_scope == scope.value and (entry_sheet or "") == (sheet or "")
            for entry_scope, entry_sheet in existing
        )
        if same_scope:
            raise ExcelVbaMcpError(
                f"Defined name {name!r} already exists in {scope.value} scope.",
                context={
                    "name": name,
                    "scope": scope.value,
                    "sheet": sheet,
                    "reason": "duplicate",
                    "existing": [{"scope": s, "sheet": sh} for s, sh in existing],
                },
            )
        if existing:
            raise ExcelVbaMcpError(
                f"Defined name {name!r} already exists at another scope; "
                "refusing to create an ambiguous cross-scope name.",
                context={
                    "name": name,
                    "scope": scope.value,
                    "sheet": sheet,
                    "reason": "scope_conflict",
                    "existing": [{"scope": s, "sheet": sh} for s, sh in existing],
                },
            )

    def _names_collection(self, wb: Any, scope: NameScope, sheet: str | None) -> Any:
        """Return the ``Names`` collection for the requested scope."""

        if scope is NameScope.WORKBOOK:
            return wb.Names
        assert sheet is not None  # guaranteed by _require_sheet_for_scope
        return _require_worksheet(wb, sheet).Names


__all__ = ["NamesAdapter", "classify_target_kind"]
