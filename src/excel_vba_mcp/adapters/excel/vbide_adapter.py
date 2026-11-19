"""STA-local VBIDE adapter (task_016, spec §6.2 / §7.2).

:class:`VBIDEAdapter` is the single place where the VBA project object model is
touched. It is a **pure STA-local helper**: every method is synchronous, receives
the live ``Workbook`` COM proxy (``Any``) and returns only Pydantic DTOs /
primitives, so results cross the single-STA worker boundary safely
(:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never
imports ``pywin32`` at module import time, never activates/selects anything and
never mutates the workbook (read-only inventory, procedure discovery and export
only — task_018 owns mutations).

Binding VBIDE facts (research ``g1-poc-research-2026-11-19.md`` §2):

* ``Workbook.VBProject`` → ``VBComponents``; ``Type``: ``1`` = Standard,
  ``2`` = Class, ``3`` = UserForm (``Designer`` is not ``None``), ``100`` =
  Document module (``ThisWorkbook`` / ``Sheet1`` — never removed/recreated).
* Procedures are enumerated with the ``CodeModule`` API (``CountOfLines``,
  ``Lines``, ``ProcOfLine``, ``ProcStartLine``, ``ProcCountLines``); the
  ``Find``-based textual fallback is explicitly marked ``heuristic`` on every
  returned procedure.
* ``Component.Export(path)`` does **not** append an extension — the caller must
  pass ``<name>.bas`` / ``.cls`` / ``.frm``. Exporting a UserForm also writes a
  companion ``<name>.frx``; ``.frx`` bytes are not reproducible, so only the
  normalized code body is hashed.
* ``VBProject.References`` expose ``Name`` / ``GUID`` / ``Major`` / ``Minor`` /
  ``FullPath`` / ``IsBroken``; any ``IsBroken`` reference is reported in
  ``VbaProjectInfo.missing_references``.
* ``VBProject.Protection`` (0 / 1) → :class:`VbaProjectLockedError`
  (``VBA_PROJECT_LOCKED``), read-only, never bypassed.
* ``Workbook.VBProject`` raises HRESULT ``0x800A17B4`` when VBOM trust is off →
  :class:`AccessVbomDisabledError` (``ACCESS_VBOM_DISABLED``); the registry is
  never written.
"""

from __future__ import annotations

import contextlib
import hashlib
import re
from collections.abc import Callable, Iterator
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ErrorCode,
    ExcelVbaMcpError,
    VbaProjectLockedError,
)
from excel_vba_mcp.domain.models import (
    VbaComponentInfo,
    VbaComponentSource,
    VbaExportEntry,
    VbaExportManifest,
    VbaProcedureInfo,
    VbaProjectInfo,
    VbaReferenceInfo,
    VbaSearchMatch,
)
from excel_vba_mcp.domain.statuses import (
    ProcedureType,
    ProcedureVisibility,
    ReferenceStatus,
    VbaComponentType,
    VbaExportStatus,
)

# VBIDE ``VBComponent.Type`` codes → domain component types (research §2.1).
_COMPONENT_TYPES: dict[int, VbaComponentType] = {
    1: VbaComponentType.STANDARD_MODULE,
    2: VbaComponentType.CLASS_MODULE,
    3: VbaComponentType.USER_FORM,
    100: VbaComponentType.DOCUMENT_MODULE,
}

# VBIDE ``vbext_ProcKind`` codes → domain procedure kinds.
_PROCEDURE_KINDS: dict[int, ProcedureType] = {
    0: ProcedureType.SUB,
    1: ProcedureType.FUNCTION,
    2: ProcedureType.PROPERTY,
    3: ProcedureType.PROPERTY,
    4: ProcedureType.PROPERTY,
}

# Export extension per VBIDE component type (research §2.2: pass the extension).
_EXPORT_EXTENSIONS: dict[int, str] = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}

# HRESULT for "Programmatic access to Visual Basic Project is not trusted".
ACCESS_VBOM_DISABLED_HRESULT = 0x800A17B4

# Extraction source markers for ``VbaProcedureInfo.source``.
_PROCEDURE_SOURCE = "code_module"
_PROCEDURE_SOURCE_HEURISTIC = "heuristic"

# Guard against pathological modules: the maximum number of code lines scanned by
# the procedure scans and the maximum search matches returned.
_MAX_SCAN_LINES = 10_000
_MAX_SEARCH_MATCHES = 500

# A procedure declaration line, anchored and scoped to one line. The attribute
# prefix (``Attribute VB_Name = ...``) is excluded so it never matches.
_PROCEDURE_HEADER_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*"
    r"(?P<kind>Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)",
    re.IGNORECASE,
)

# Detects a procedure *header* (declaration) line; used to validate a
# ``ProcOfLine`` result before trusting the VBIDE procedure APIs.
_PROCEDURE_DECLARATION_RE = re.compile(
    r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*"
    r"(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+[A-Za-z_]",
    re.IGNORECASE,
)

# VBIDE ``CodeModule.ProcOfLine(line, kind)`` returns a ``(name, kind)`` tuple on
# the target build; ``ProcStartLine(name, kind)`` / ``ProcCountLines(name, kind)``
# take the procedure *name*. ``kind`` is ``vbext_pk_Proc`` (0) to match any kind.
_VBEXT_PK_PROC = 0

# Default number of ``Find`` matches requested from a CodeModule (0 = all).
_FIND_ALL_MATCHES = 0


def utc_now() -> datetime:
    """Current UTC time (used for the ``TimestampedModel`` DTOs)."""

    return datetime.now(UTC)


def normalize_code(text: str) -> str:
    """Normalize EOL / trailing whitespace so the same source hashes identically.

    CRLF and CR are unified to ``\\n``, each line is right-trimmed and
    leading/trailing blank lines are dropped (research §2.6: the ``.frx`` is not
    byte-reproducible, so the normalized code body is the only stable identity).
    """

    unified = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.rstrip() for line in unified.split("\n")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(lines)


def hash_code(text: str) -> str:
    """Return the SHA-256 hex digest of the normalized code body."""

    return hashlib.sha256(normalize_code(text).encode("utf-8")).hexdigest()


def extract_frm_code_body(frm_text: str) -> str:
    """Return the VBA code body of an exported ``.frm``.

    A ``.frm`` is a designer header (``VERSION`` / ``Begin ... End``) followed by
    an ``Attribute VB_*`` block and then the code; the body is everything after
    the last ``Attribute VB_`` line (research §2.2).
    """

    unified = frm_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = unified.split("\n")
    last_attr = -1
    for index, line in enumerate(lines):
        if line.startswith("Attribute VB_"):
            last_attr = index
    body = lines[last_attr + 1 :] if last_attr >= 0 else lines
    return "\n".join(body)


def _sha256_file(path: Path) -> str | None:
    """Return the SHA-256 hex digest of ``path`` or ``None`` when absent."""

    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _outer_hresult(exc: BaseException) -> int | None:
    """Return the outer HRESULT carried on a COM error, if any."""

    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        return hresult & 0xFFFFFFFF
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0] & 0xFFFFFFFF
    return None


def _inner_hresult(exc: BaseException) -> int | None:
    """Return the EXCEPINFO inner HRESULT from a ``com_error``, if present."""

    args = getattr(exc, "args", None)
    if not isinstance(args, tuple) or len(args) < 3:
        return None
    excepinfo = args[2]
    if not isinstance(excepinfo, tuple) or len(excepinfo) < 6:
        return None
    candidate = excepinfo[5]
    if isinstance(candidate, int):
        return candidate & 0xFFFFFFFF
    return None


def _classify_vbproject_error(exc: BaseException) -> AccessVbomDisabledError | ExcelVbaMcpError:
    """Classify a ``Workbook.VBProject`` access failure into a domain error.

    Returns an :class:`AccessVbomDisabledError` when the VBOM-disabled HRESULT
    (``0x800A17B4``) is present on the exception (outer or inner); otherwise a
    generic :class:`ExcelVbaMcpError`.
    """

    if (
        _outer_hresult(exc) == ACCESS_VBOM_DISABLED_HRESULT
        or _inner_hresult(exc) == ACCESS_VBOM_DISABLED_HRESULT
    ):
        return AccessVbomDisabledError(
            "Programmatic access to the VBA project is not trusted.",
            hresult=ACCESS_VBOM_DISABLED_HRESULT,
            context={
                "outer_hresult": _outer_hresult(exc),
                "inner_hresult": _inner_hresult(exc),
            },
        )
    return ExcelVbaMcpError(
        str(exc) or exc.__class__.__name__,
        code=ErrorCode.COM_BUSY,
        context={"stage": "Workbook.VBProject"},
    )


def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively."""

    try:
        return str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_int(value: Any, default: int = 0) -> int:
    """Read a COM integer property defensively."""

    try:
        return int(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_bool(value: Any, default: bool = False) -> bool:
    """Read a COM boolean property defensively."""

    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


class VBIDEAdapter:
    """Stateless, STA-local reader/exporter over a live ``Workbook`` VBProject.

    The adapter holds no references between calls. Every method receives the
    STA-local ``wb`` proxy, touches only the VBIDE object model and returns
    DTOs / primitives. ``Workbook.VBProject`` access is funnelled through
    :meth:`_project`, which maps the VBOM-disabled HRESULT to
    :class:`AccessVbomDisabledError`; protection is checked up-front so a locked
    project fails with :class:`VbaProjectLockedError` and never attempts a
    bypass.
    """

    # ------------------------------------------------------------------
    # Project info
    # ------------------------------------------------------------------
    def project_info(self, wb: Any) -> VbaProjectInfo:
        """Return a :class:`VbaProjectInfo` summary of ``wb``'s VBA project."""

        project = self._project(wb)
        self._ensure_unlocked(project)

        components = self._components(project)
        references = self._references(project)
        missing = [reference.name for reference in references if reference.is_broken]

        warnings: list[str] = []
        if missing:
            warnings.append("missing_references: " + ", ".join(missing))

        return VbaProjectInfo(
            available=True,
            locked=False,
            name=_safe_str(getattr(project, "Name", "")) or None,
            component_count=len(components),
            components=components,
            references=references,
            missing_references=missing,
            warnings=warnings,
        )

    # ------------------------------------------------------------------
    # Components
    # ------------------------------------------------------------------
    def list_components(self, wb: Any) -> list[VbaComponentInfo]:
        """Return the component inventory of ``wb``'s VBA project."""

        project = self._project(wb)
        self._ensure_unlocked(project)
        return self._components(project)

    # ------------------------------------------------------------------
    # Procedures
    # ------------------------------------------------------------------
    def list_procedures(self, wb: Any, component_name: str) -> list[VbaProcedureInfo]:
        """Return the procedures of ``component_name`` (CodeModule API first).

        Procedures are enumerated with the VBIDE ``CodeModule`` API
        (``CountOfLines`` / ``Lines`` / ``ProcOfLine`` / ``ProcStartLine`` /
        ``ProcCountLines``). When the CodeModule API is unavailable (locked or
        partial project) a ``Find``-based / textual fallback is used and every
        returned procedure is flagged ``is_heuristic=True``.
        """

        project = self._project(wb)
        self._ensure_unlocked(project)
        component = self._require_component(project, component_name)
        code_module = self._code_module(component)

        procedures = self._procedures_from_code_module(component_name, code_module)
        if procedures is None:
            text = self._read_code_text(code_module)
            procedures = self._procedures_heuristic(component_name, text)
        return procedures

    # ------------------------------------------------------------------
    # Source read
    # ------------------------------------------------------------------
    def read_component(self, wb: Any, component_name: str) -> VbaComponentSource:
        """Return the code text + normalized hash of ``component_name``."""

        project = self._project(wb)
        self._ensure_unlocked(project)
        component = self._require_component(project, component_name)
        text = self._read_code_text(self._code_module(component))
        now = utc_now()
        return VbaComponentSource(
            component_name=_safe_str(component.Name, component_name),
            component_type=_component_type(component),
            source=text,
            line_count=len(normalize_code(text).split("\n")) if text.strip() else 0,
            source_hash=hash_code(text),
            revision=1,
            created_at=now,
            updated_at=now,
        )

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------
    def export_project(
        self, wb: Any, out_dir: str | Path, *, workbook_id: str = ""
    ) -> VbaExportManifest:
        """Export every component of ``wb``'s VBA project into ``out_dir``.

        ``Component.Export`` does not append an extension, so the adapter always
        passes ``<name>.bas`` / ``.cls`` / ``.frm``. UserForms additionally
        produce a ``.frx`` resource (not byte-reproducible; recorded but not
        hashed for identity). Document modules are exported as ``.cls`` and are
        never removed or recreated.
        """

        project = self._project(wb)
        self._ensure_unlocked(project)
        target = Path(out_dir)
        target.mkdir(parents=True, exist_ok=True)

        entries: list[VbaExportEntry] = []
        for component in self._iter_components(project):
            entries.append(self._export_component(component, target))

        combined = hashlib.sha256()
        for entry in entries:
            combined.update(entry.source_hash.encode("utf-8"))
        now = utc_now()
        return VbaExportManifest(
            workbook_id=workbook_id,
            out_dir=str(target),
            components=entries,
            source_hash=combined.hexdigest(),
            created_at=now,
            updated_at=now,
        )

    # ------------------------------------------------------------------
    # References
    # ------------------------------------------------------------------
    def get_references(self, wb: Any) -> list[VbaReferenceInfo]:
        """Return ``wb``'s VBA project references (``IsBroken`` preserved)."""

        project = self._project(wb)
        return self._references(project)

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------
    def search_code(self, wb: Any, pattern: str, *, regex: bool = False) -> list[VbaSearchMatch]:
        """Search every component's code for ``pattern`` (bounded results).

        With ``regex=False`` the match is a case-insensitive substring test;
        with ``regex=True`` ``pattern`` is compiled as a regular expression and
        any invalid pattern raises a domain error. Results are bounded by
        :data:`_MAX_SEARCH_MATCHES` so a pathological module cannot flood the
        response.
        """

        project = self._project(wb)
        self._ensure_unlocked(project)

        matcher: Callable[[str], bool]
        if regex:
            try:
                compiled = re.compile(pattern)
            except re.error as exc:
                raise ExcelVbaMcpError(
                    f"Invalid regular expression: {exc}",
                    code=ErrorCode.COM_BUSY,
                    context={"pattern": pattern},
                ) from exc
            compiled_search = compiled.search
            matcher = _regex_matcher(compiled_search)
        else:
            matcher = _substring_matcher(pattern.casefold())

        matches: list[VbaSearchMatch] = []
        for component in self._iter_components(project):
            component_type = _component_type(component)
            text = self._read_code_text(self._code_module(component))
            for index, line in enumerate(text.split("\n"), start=1):
                if len(matches) >= _MAX_SEARCH_MATCHES:
                    return matches
                if matcher(line):
                    matches.append(
                        VbaSearchMatch(
                            component_name=_safe_str(component.Name),
                            component_type=component_type,
                            line_number=index,
                            line_text=line.strip(),
                        )
                    )
        return matches

    # ------------------------------------------------------------------
    # Project / protection helpers
    # ------------------------------------------------------------------
    def _project(self, wb: Any) -> Any:
        """Return ``wb.VBProject``, mapping the VBOM-disabled HRESULT.

        Raises:
            AccessVbomDisabledError: When VBOM trust is disabled
                (``0x800A17B4``, outer or inner HRESULT).
            ExcelVbaMcpError: For any other COM failure.
        """

        try:
            return wb.VBProject
        except Exception as exc:  # noqa: BLE001 - classified below
            raise _classify_vbproject_error(exc) from exc

    def _ensure_unlocked(self, project: Any) -> None:
        """Raise :class:`VbaProjectLockedError` when the project is protected."""

        if _safe_int(getattr(project, "Protection", 0)) != 0:
            raise VbaProjectLockedError(
                "The VBA project is locked for viewing; protection is never bypassed.",
                context={"protection": _safe_int(getattr(project, "Protection", 0))},
            )

    def _components(self, project: Any) -> list[VbaComponentInfo]:
        """Build the component inventory from ``VBComponents``."""

        return [self._describe_component(component) for component in self._iter_components(project)]

    def _describe_component(self, component: Any) -> VbaComponentInfo:
        """Build a :class:`VbaComponentInfo` for one ``VBComponent``."""

        component_type = _component_type(component)
        raw_type = _safe_int(getattr(component, "Type", 0))
        has_designer = component_type is VbaComponentType.USER_FORM and self._has_designer(
            component
        )
        line_count = _safe_int(getattr(self._code_module(component), "CountOfLines", 0))
        text = self._read_code_text(self._code_module(component))
        return VbaComponentInfo(
            component_name=_safe_str(component.Name),
            component_type=component_type,
            line_count=line_count,
            declaration_line_count=None,
            source_hash=hash_code(text) if text.strip() else None,
            is_document_module=component_type is VbaComponentType.DOCUMENT_MODULE,
            has_designer=has_designer,
            is_exportable=raw_type in _EXPORT_EXTENSIONS,
        )

    def _export_component(self, component: Any, target: Path) -> VbaExportEntry:
        """Export one component (``.bas`` / ``.cls`` / ``.frm`` + ``.frx``)."""

        raw_type = _safe_int(getattr(component, "Type", 0))
        component_type = _component_type(component)
        name = _safe_str(component.Name)
        extension = _EXPORT_EXTENSIONS.get(raw_type, ".txt")
        file_path = target / f"{name}{extension}"
        text = self._read_code_text(self._code_module(component))
        source_hash = hash_code(text) if text.strip() else hash_code("")
        entry = VbaExportEntry(
            component_name=name,
            component_type=component_type,
            file_name=file_path.name,
            extension=extension,
            file_path=str(file_path),
            file_sha256=None,
            source_hash=source_hash,
            resource_file_name=None,
            resource_file_sha256=None,
            export_status=VbaExportStatus.OK,
            error=None,
        )
        try:
            component.Export(str(file_path))
        except Exception as exc:  # noqa: BLE001 - export failure is data, not a crash
            entry.export_status = VbaExportStatus.FAILED
            entry.error = str(exc) or exc.__class__.__name__
            return entry

        if not file_path.is_file():
            entry.export_status = VbaExportStatus.FAILED
            entry.error = "export did not produce a file"
            return entry

        entry.file_sha256 = _sha256_file(file_path)
        resource_path = target / f"{name}.frx"
        if resource_path.is_file():
            entry.resource_file_name = resource_path.name
            entry.resource_file_sha256 = _sha256_file(resource_path)
        return entry

    def _references(self, project: Any) -> list[VbaReferenceInfo]:
        """Read ``VBProject.References`` into DTOs (``IsBroken`` preserved)."""

        references: list[VbaReferenceInfo] = []
        try:
            items = list(project.References)
        except Exception:  # noqa: BLE001 - defensive COM read
            return references
        for reference in items:
            with contextlib.suppress(Exception):
                references.append(
                    VbaReferenceInfo(
                        name=_safe_str(getattr(reference, "Name", "")),
                        guid=_safe_str(getattr(reference, "GUID", "")) or None,
                        version=_format_version(reference),
                        path=_safe_str(getattr(reference, "FullPath", "")) or None,
                        status=(
                            ReferenceStatus.BROKEN
                            if _safe_bool(getattr(reference, "IsBroken", False))
                            else ReferenceStatus.AVAILABLE
                        ),
                        is_broken=_safe_bool(getattr(reference, "IsBroken", False)),
                    )
                )
        return references

    # ------------------------------------------------------------------
    # Component / code-module accessors
    # ------------------------------------------------------------------
    @staticmethod
    def _iter_components(project: Any) -> Iterator[Any]:
        """Yield every ``VBComponent`` of ``project`` (index-based, defensive)."""

        try:
            components = project.VBComponents
            count = int(components.Count)
        except Exception:  # noqa: BLE001 - defensive COM read
            return
        for index in range(1, count + 1):
            try:
                yield components.Item(index)
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue

    @staticmethod
    def _require_component(project: Any, component_name: str) -> Any:
        """Return the named ``VBComponent`` or raise a domain error."""

        if not component_name or not component_name.strip():
            raise ExcelVbaMcpError(
                "component_name must be a non-empty string.",
                code=ErrorCode.COM_BUSY,
                context={"field": "component_name"},
            )
        try:
            return project.VBComponents(component_name)
        except Exception:  # noqa: BLE001 - absent component
            known = [
                _safe_str(item.Name)
                for item in VBIDEAdapter._iter_components(project)
                if _safe_str(item.Name)
            ]
            raise ExcelVbaMcpError(
                f"VBA component {component_name!r} not found.",
                code=ErrorCode.COM_BUSY,
                context={"component_name": component_name, "known": known},
            ) from None

    @staticmethod
    def _code_module(component: Any) -> Any:
        """Return ``component.CodeModule`` (STA-local proxy)."""

        return component.CodeModule

    @staticmethod
    def _read_code_text(code_module: Any) -> str:
        """Return the full code-module text via ``Lines(1, CountOfLines)``."""

        count = 0
        with contextlib.suppress(Exception):
            count = int(code_module.CountOfLines)
        if count <= 0:
            return ""
        try:
            return _safe_str(code_module.Lines(1, count))
        except Exception:  # noqa: BLE001 - defensive COM read
            return ""

    @staticmethod
    def _has_designer(component: Any) -> bool:
        """Return whether a UserForm component exposes a ``Designer``."""

        try:
            return component.Designer is not None
        except Exception:  # noqa: BLE001 - defensive COM read
            return False

    # ------------------------------------------------------------------
    # Procedure discovery
    # ------------------------------------------------------------------
    def _procedures_from_code_module(
        self, component_name: str, code_module: Any
    ) -> list[VbaProcedureInfo] | None:
        """Enumerate procedures with the VBIDE CodeModule API.

        Returns ``None`` when the CodeModule API is not usable (e.g. a locked
        project), so the caller can fall back to the heuristic scan.
        """

        lines: list[str] = []
        try:
            total = int(code_module.CountOfLines)
            if total > 0:
                lines = _safe_str(code_module.Lines(1, total)).split("\n")
        except Exception:  # noqa: BLE001 - CodeModule API unavailable
            return None

        procedures: list[VbaProcedureInfo] = []
        for line_number, raw_line in enumerate(lines[:_MAX_SCAN_LINES], start=1):
            if _PROCEDURE_DECLARATION_RE.match(raw_line.rstrip("\r")) is None:
                continue
            header = _PROCEDURE_HEADER_RE.match(raw_line.rstrip("\r"))
            if header is None:  # pragma: no cover - defensive (declaration implies header)
                continue

            proc_name, kind_code, name_ok = _proc_of_line(
                code_module, line_number, header.group("name")
            )
            start_ok, proc_start = _try_call(code_module.ProcStartLine, proc_name, _VBEXT_PK_PROC)
            count_ok, proc_count = _try_call(code_module.ProcCountLines, proc_name, _VBEXT_PK_PROC)
            if not (name_ok and start_ok and count_ok):
                # A VBIDE procedure API is unavailable or returned an unusable
                # value: emit a heuristic record so callers see the flagged
                # fallback rather than a false "code_module" result.
                procedures.append(
                    VbaProcedureInfo(
                        component_name=component_name,
                        name=header.group("name"),
                        procedure_type=_procedure_type(-1, header.group("kind")),
                        visibility=_procedure_visibility(raw_line),
                        start_line=line_number,
                        end_line=line_number,
                        code=raw_line,
                        is_heuristic=True,
                        source=_PROCEDURE_SOURCE_HEURISTIC,
                    )
                )
                continue

            proc_start_line = _safe_int(proc_start, line_number)
            proc_count_line = _safe_int(proc_count, 1)
            proc_end = proc_start_line + max(proc_count_line - 1, 0)
            start_line = _find_procedure_start(code_module, proc_start_line, proc_end, line_number)
            procedures.append(
                VbaProcedureInfo(
                    component_name=component_name,
                    name=header.group("name"),
                    procedure_type=_procedure_type(_safe_int(kind_code, -1), header.group("kind")),
                    visibility=_procedure_visibility(raw_line),
                    start_line=start_line,
                    end_line=proc_end,
                    code=_safe_str(_call_or_none(code_module.Lines, start_line, proc_end)),
                    is_heuristic=False,
                    source=_PROCEDURE_SOURCE,
                )
            )
        return procedures

    def _procedures_heuristic(self, component_name: str, text: str) -> list[VbaProcedureInfo]:
        """Fallback procedure scan marked ``is_heuristic=True`` (research §7.2).

        Used only when the CodeModule API cannot enumerate procedures. The result
        is explicitly flagged as heuristic on every procedure so callers never
        treat it as authoritative VBIDE output.
        """

        procedures: list[VbaProcedureInfo] = []
        lines = text.split("\n")
        for line_number, raw_line in enumerate(lines[:_MAX_SCAN_LINES], start=1):
            header = _PROCEDURE_HEADER_RE.match(raw_line.rstrip("\r"))
            if header is None:
                continue
            procedures.append(
                VbaProcedureInfo(
                    component_name=component_name,
                    name=header.group("name"),
                    procedure_type=_procedure_type(-1, header.group("kind")),
                    visibility=_procedure_visibility(raw_line),
                    start_line=line_number,
                    end_line=line_number,
                    code=raw_line,
                    is_heuristic=True,
                    source=_PROCEDURE_SOURCE_HEURISTIC,
                )
            )
        return procedures


# ---------------------------------------------------------------------------
# Module-level helpers (pure, no COM)
# ---------------------------------------------------------------------------
def _format_version(reference: Any) -> str | None:
    """Format a reference's ``Major.Minor`` version defensively."""

    major = _safe_int(getattr(reference, "Major", 0))
    minor = _safe_int(getattr(reference, "Minor", 0))
    if major == 0 and minor == 0:
        return None
    return f"{major}.{minor}"


def _substring_matcher(needle: str) -> Callable[[str], bool]:
    """Return a case-insensitive substring matcher bound to ``needle``."""

    def match(line: str) -> bool:
        return needle in line.casefold()

    return match


def _regex_matcher(search: Callable[[str], Any]) -> Callable[[str], bool]:
    """Return a matcher backed by a compiled pattern's ``search`` method."""

    def match(line: str) -> bool:
        return search(line) is not None

    return match


def _component_type(component: Any) -> VbaComponentType:
    """Map a ``VBComponent.Type`` code to a :class:`VbaComponentType`."""

    raw_type = _safe_int(getattr(component, "Type", 0))
    return _COMPONENT_TYPES.get(raw_type, VbaComponentType.STANDARD_MODULE)


def _procedure_type(kind_code: int, keyword: str) -> ProcedureType:
    """Resolve a procedure kind from the VBIDE code (or the header keyword)."""

    if kind_code in _PROCEDURE_KINDS:
        return _PROCEDURE_KINDS[kind_code]
    lowered = keyword.casefold()
    if lowered == "sub":
        return ProcedureType.SUB
    if lowered == "function":
        return ProcedureType.FUNCTION
    if lowered.startswith("property"):
        return ProcedureType.PROPERTY
    return ProcedureType.PROPERTY


def _procedure_visibility(declaration_line: str) -> ProcedureVisibility:
    """Infer procedure visibility from its declaration line."""

    tokens = declaration_line.strip().split()
    for token in tokens:
        lowered = token.casefold()
        if lowered == "public":
            return ProcedureVisibility.PUBLIC
        if lowered == "private":
            return ProcedureVisibility.PRIVATE
        if lowered == "friend":
            return ProcedureVisibility.FRIEND
    return ProcedureVisibility.UNKNOWN


def _call_or_none(func: Any, *args: Any) -> Any:
    """Call ``func`` defensively, returning ``None`` on any COM failure."""

    try:
        return func(*args)
    except Exception:  # noqa: BLE001 - defensive COM read
        return None


def _try_call(func: Any, *args: Any) -> tuple[bool, Any]:
    """Call ``func``; return ``(True, value)`` on success or ``(False, None)``."""

    try:
        return True, func(*args)
    except Exception:  # noqa: BLE001 - defensive COM read
        return False, None


def _proc_of_line(code_module: Any, line_number: int, fallback_name: str) -> tuple[str, int, bool]:
    """Resolve ``CodeModule.ProcOfLine`` to ``(name, kind, ok)``.

    On the target build ``ProcOfLine(line, vbext_pk_Proc)`` returns a
    ``(name, kind)`` tuple. Older/other builds may return a bare kind; both are
    accepted. ``ok`` is ``False`` when the API is unavailable or returns an
    unusable value, which makes the caller emit a heuristic record.
    """

    ok, value = _try_call(code_module.ProcOfLine, line_number, _VBEXT_PK_PROC)
    if not ok:
        return fallback_name, -1, False
    if isinstance(value, tuple) and len(value) >= 2:
        name = _safe_str(value[0], fallback_name) or fallback_name
        return name, _safe_int(value[1], -1), True
    if isinstance(value, int):
        return fallback_name, value, True
    return fallback_name, -1, False


def _end_line(code_module: Any, proc_name: str, start_line: int, fallback: int) -> int:
    """Return the last line of the named procedure that starts at ``start_line``."""

    ok, count = _try_call(code_module.ProcCountLines, proc_name, _VBEXT_PK_PROC)
    if ok and isinstance(count, int) and count > 0:
        return start_line + count - 1
    return fallback


def _find_procedure_start(code_module: Any, start_line: int, end_line: int, fallback: int) -> int:
    """Return the procedure header line within ``[start_line, end_line]``.

    VBIDE's ``ProcStartLine`` is authoritative when it points at a declaration
    header; otherwise the first declaration line in the range is used so the
    reported start is truthful.
    """

    ok, first = _try_call(code_module.Lines, start_line, 1)
    if ok and _is_declaration(_safe_str(first)):
        return start_line
    span_ok, span = _try_call(code_module.Lines, start_line, max(end_line - start_line, 0) + 1)
    if span_ok:
        for offset, line in enumerate(_safe_str(span).split("\n")):
            if _is_declaration(line):
                return start_line + offset
    return fallback


def _is_declaration(text: str) -> bool:
    """Return whether ``text`` is a single VBA procedure declaration line."""

    return _PROCEDURE_DECLARATION_RE.match(text.rstrip("\r")) is not None


__all__ = [
    "ACCESS_VBOM_DISABLED_HRESULT",
    "VBIDEAdapter",
    "extract_frm_code_body",
    "hash_code",
    "normalize_code",
]
