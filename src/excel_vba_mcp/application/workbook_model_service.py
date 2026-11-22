"""Stage-1 workbook model aggregation service (task_032, spec §7.1 / §6.6 / A03).

:class:`WorkbookModelService` builds a single, structured, serializable model of
an open workbook by **aggregating the existing Stage 1 inventory services**:

* sheets — :class:`~excel_vba_mcp.application.sheets_service.SheetsService`
* tables — :class:`~excel_vba_mcp.application.tables_service.TablesService`
* names — :class:`~excel_vba_mcp.application.names_service.NamesService`
* shapes — :class:`~excel_vba_mcp.application.shapes_service.ShapesService`
* controls — :class:`~excel_vba_mcp.application.controls_service.ControlsService`
* VBA components / procedures —
  :class:`~excel_vba_mcp.application.vba_project_service.VBAProjectService`

Design rules (plan task_032, spec §7.1, A03):

* **Aggregate, never re-traverse.** This service performs **no COM traversal of
  its own**: every fact comes from one of the injected services. It only resolves
  the ``workbook_id`` through the session bookkeeping and awaits the services.
* **No all-cells / Range data.** The model contains object inventory and
  metadata only; it never reads ``Range`` / ``Cells`` values and exposes no
  Range/Cells API (that is Phase 2).
* **VBA is optional and non-fatal.** When ``include_vba`` is true the VBA project
  is read; if VBProject is unavailable, locked or VBOM trust is disabled the
  other sections are still returned and ``WorkbookModel.vba_access`` reports the
  state. A failure never aborts the whole model.
* **Per-section status.** Every section records ``available`` / ``unavailable`` /
  ``partial`` plus diagnostics, so a caller can tell exactly what was read.
* **Content hash.** ``model_hash`` is a SHA-256 over the model content (excluding
  the timestamps and the hash itself), so two reads of an unchanged workbook hash
  identically.

Safety (BG03/BG04/BG05): the service only awaits the read-only service methods;
it never activates/selects anything, never runs VBA and never mutates the
workbook.
"""

from __future__ import annotations

import hashlib
import json
from collections.abc import Awaitable, Callable
from datetime import UTC, datetime
from typing import Any, TypeVar

from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.application.controls_service import ControlsService
from excel_vba_mcp.application.names_service import NamesService
from excel_vba_mcp.application.pivot_tables_service import PivotTablesService
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.application.shapes_service import ShapesService
from excel_vba_mcp.application.sheets_service import SheetsService
from excel_vba_mcp.application.tables_service import TablesService
from excel_vba_mcp.application.vba_project_service import VBAProjectService
from excel_vba_mcp.domain.errors import (
    AccessVbomDisabledError,
    ExcelVbaMcpError,
    VbaProjectLockedError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import (
    DefinedNameInfo,
    ModelSection,
    PivotTableInfo,
    SheetModel,
    VbaProcedureInfo,
    VbaProjectInfo,
    WorkbookInfo,
    WorkbookModel,
)
from excel_vba_mcp.domain.statuses import (
    ModelSectionStatus,
    NameScope,
    VbaAccessStatus,
)
from excel_vba_mcp.security.audit import AuditLogger
from excel_vba_mcp.security.path_policy import PathPolicy

# Section names (stable keys in ``WorkbookModel.sections`` / ``SheetModel.sections``).
SECTION_SHEETS = "sheets"
SECTION_TABLES = "tables"
SECTION_NAMES = "names"
SECTION_SHAPES = "shapes"
SECTION_CONTROLS = "controls"
SECTION_VBA = "vba"
SECTION_PIVOTS = "pivots"

# Result of a worker call submitted by this service.
_T = TypeVar("_T")


class WorkbookModelService:
    """Aggregate the Stage 1 inventory services into a workbook / sheet model.

    Args:
        worker: The single-STA COM worker (kept for signature symmetry; this
            service delegates all Excel access to the injected services).
        session_service: Workbook lifecycle service (resolves ``workbook_id``).
        sheets_service: Sheets inventory service.
        tables_service: Excel Table inventory service.
        names_service: Defined-name inventory service.
        shapes_service: Drawing-shape inventory service.
        controls_service: Form Control / ActiveX inventory service.
        vba_project_service: Read-only VBA project service (optional section).
        path_policy: Allowlist policy (kept for signature symmetry / audit).
        audit: Append-only JSONL audit trail.
        pivot_tables_service: Read-only PivotTables inventory service
            (task_048, optional keyword section). When provided,
            ``get_sheet_model`` adds PivotTable metadata scoped to the sheet
            without materializing the PivotTable cell output.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        sheets_service: SheetsService,
        tables_service: TablesService,
        names_service: NamesService,
        shapes_service: ShapesService,
        controls_service: ControlsService,
        vba_project_service: VBAProjectService,
        path_policy: PathPolicy,
        audit: AuditLogger,
        *,
        pivot_tables_service: PivotTablesService | None = None,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._sheets = sheets_service
        self._tables = tables_service
        self._names = names_service
        self._shapes = shapes_service
        self._controls = controls_service
        self._vba = vba_project_service
        self._path_policy = path_policy
        self._audit = audit
        self._pivots = pivot_tables_service

    # ------------------------------------------------------------------
    # Public API (async; no COM here — every fact comes from a service)
    # ------------------------------------------------------------------
    async def get_workbook_model(
        self, workbook_id: str, *, include_vba: bool = True
    ) -> WorkbookModel:
        """Return the aggregated Stage 1 model for ``workbook_id``.

        The model aggregates sheets, tables, names, shapes, controls and (when
        ``include_vba`` is true) the VBA project. Each section carries a status;
        a failing section never aborts the others. ``vba_access`` reports the VBA
        project state (``available`` / ``unavailable`` / ``locked`` /
        ``vbom_disabled`` / ``skipped``).

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            include_vba: When ``True`` (default) the VBA project is aggregated.

        Returns:
            A :class:`WorkbookModel` with a content ``model_hash``.
        """

        info = self._resolve_record(workbook_id)
        diagnostics: list[str] = []
        sections: dict[str, ModelSection] = {}

        sheets, section = await self._collect(
            SECTION_SHEETS, lambda: self._sheets.list_sheets(workbook_id), diagnostics
        )
        sections[SECTION_SHEETS] = section

        tables, section = await self._collect(
            SECTION_TABLES, lambda: self._tables.list_tables(workbook_id), diagnostics
        )
        sections[SECTION_TABLES] = section

        names, section = await self._collect(
            SECTION_NAMES, lambda: self._names.list_names(workbook_id), diagnostics
        )
        sections[SECTION_NAMES] = section

        shapes, section = await self._collect(
            SECTION_SHAPES, lambda: self._shapes.list_shapes(workbook_id), diagnostics
        )
        sections[SECTION_SHAPES] = section

        controls, section = await self._collect(
            SECTION_CONTROLS, lambda: self._controls.list_controls(workbook_id), diagnostics
        )
        sections[SECTION_CONTROLS] = section

        vba_project, vba_procedures, vba_access, vba_section = await self._collect_vba(
            workbook_id, include_vba
        )
        sections[SECTION_VBA] = vba_section
        if vba_section.diagnostics:
            diagnostics.extend(vba_section.diagnostics)

        now = _utc_now()
        model = WorkbookModel(
            workbook=info,
            sheets=sheets,
            tables=tables,
            vba_project=vba_project,
            vba_access=vba_access,
            vba_procedures=vba_procedures,
            names=names,
            shapes=shapes,
            controls=controls,
            sections=sections,
            created_at=now,
            updated_at=now,
        )
        model.model_hash = content_hash(model)
        self._record_audit(
            "workbook_get_model",
            {"workbook_id": workbook_id, "include_vba": include_vba},
            {
                "sheets": len(sheets),
                "tables": len(tables),
                "names": len(names),
                "shapes": len(shapes),
                "controls": len(controls),
                "vba_access": vba_access.value,
                "model_hash": model.model_hash,
            },
            info,
        )
        return model

    async def get_sheet_model(self, workbook_id: str, sheet_name: str) -> SheetModel:
        """Return the sheet-scoped Stage 1 model for ``sheet_name``.

        Aggregates the sheet metadata plus the tables, names, shapes, controls and
        (when the pivot service is wired) PivotTables scoped to that sheet from the
        same inventory services. Each scoped section carries a status; a failing
        section never aborts the others. PivotTable metadata only is added — never
        the PivotTable cell output (task_048).

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            sheet_name: The sheet name (as shown in Excel).

        Returns:
            A :class:`SheetModel` with a content ``model_hash``.
        """

        info = self._resolve_record(workbook_id)
        diagnostics: list[str] = []

        sheet = await self._sheets.get_sheet(workbook_id, sheet_name)

        all_tables, table_section = await self._collect(
            SECTION_TABLES, lambda: self._tables.list_tables(workbook_id), diagnostics
        )
        all_names, name_section = await self._collect(
            SECTION_NAMES, lambda: self._names.list_names(workbook_id), diagnostics
        )
        all_shapes, shape_section = await self._collect(
            SECTION_SHAPES, lambda: self._shapes.list_shapes(workbook_id), diagnostics
        )
        all_controls, control_section = await self._collect(
            SECTION_CONTROLS, lambda: self._controls.list_controls(workbook_id), diagnostics
        )

        all_pivots, pivot_section = await self._collect_pivots(workbook_id, diagnostics)

        tables = [table for table in all_tables if table.sheet_name == sheet.name]
        names = [name for name in all_names if _name_on_sheet(name, sheet.name)]
        shapes = [shape for shape in all_shapes if shape.sheet_name == sheet.name]
        controls = [control for control in all_controls if control.sheet_name == sheet.name]
        pivots = [pivot for pivot in all_pivots if pivot.sheet_name == sheet.name]

        sections = {
            SECTION_TABLES: _rescope(table_section, len(tables)),
            SECTION_NAMES: _rescope(name_section, len(names)),
            SECTION_SHAPES: _rescope(shape_section, len(shapes)),
            SECTION_CONTROLS: _rescope(control_section, len(controls)),
            SECTION_PIVOTS: _rescope(pivot_section, len(pivots)),
        }

        now = _utc_now()
        model = SheetModel(
            workbook_id=info.workbook_id,
            sheet=sheet,
            tables=tables,
            names=names,
            shapes=shapes,
            controls=controls,
            pivot_tables=pivots,
            sections=sections,
            created_at=now,
            updated_at=now,
        )
        model.model_hash = content_hash(model)
        self._record_audit(
            "sheet_get_model",
            {"workbook_id": workbook_id, "sheet": sheet.name},
            {
                "tables": len(tables),
                "names": len(names),
                "shapes": len(shapes),
                "controls": len(controls),
                "pivot_tables": len(pivots),
                "model_hash": model.model_hash,
            },
            info,
        )
        return model

    # ------------------------------------------------------------------
    # Aggregation helpers (non-COM)
    # ------------------------------------------------------------------
    async def _collect(
        self,
        name: str,
        call: Callable[[], Awaitable[list[_T]]],
        diagnostics: list[str],
    ) -> tuple[list[_T], ModelSection]:
        """Await one section read, isolating failures into a section status.

        A failing section yields an empty list and an ``unavailable`` status with
        a diagnostic; it never propagates and never aborts the model.
        """

        local: list[str] = []
        try:
            items = await call()
        except Exception as exc:  # noqa: BLE001 - section isolation is intentional
            local.append(f"{name}: {_describe_error(exc)}")
            diagnostics.extend(local)
            return [], ModelSection(
                name=name,
                status=ModelSectionStatus.UNAVAILABLE,
                count=0,
                diagnostics=local,
            )
        return items, ModelSection(
            name=name,
            status=ModelSectionStatus.AVAILABLE,
            count=len(items),
        )

    async def _collect_pivots(
        self, workbook_id: str, diagnostics: list[str]
    ) -> tuple[list[PivotTableInfo], ModelSection]:
        """Aggregate the optional PivotTable section without aborting the model.

        When the pivot service is wired (``pivot_tables_service`` was provided),
        the full PivotTable inventory is read and returned for sheet-scoping by
        the caller. When it is not wired, or the read fails, an empty list and a
        ``unavailable`` / ``skipped`` section status are returned so the rest of
        the sheet model still builds (task_048).
        """

        if self._pivots is None:
            return [], ModelSection(
                name=SECTION_PIVOTS,
                status=ModelSectionStatus.UNAVAILABLE,
                count=0,
                diagnostics=[f"{SECTION_PIVOTS}: pivot service not configured"],
            )
        local: list[str] = []
        try:
            inventory = await self._pivots.list_pivot_tables(workbook_id)
        except Exception as exc:  # noqa: BLE001 - section isolation is intentional
            local.append(f"{SECTION_PIVOTS}: {_describe_error(exc)}")
            diagnostics.extend(local)
            return [], ModelSection(
                name=SECTION_PIVOTS,
                status=ModelSectionStatus.UNAVAILABLE,
                count=0,
                diagnostics=local,
            )
        return inventory.pivot_tables, ModelSection(
            name=SECTION_PIVOTS,
            status=ModelSectionStatus.AVAILABLE,
            count=len(inventory.pivot_tables),
        )

    async def _collect_vba(
        self, workbook_id: str, include_vba: bool
    ) -> tuple[VbaProjectInfo | None, list[VbaProcedureInfo], VbaAccessStatus, ModelSection]:
        """Aggregate the optional VBA section without ever aborting the model."""

        if not include_vba:
            return (
                None,
                [],
                VbaAccessStatus.SKIPPED,
                ModelSection(name=SECTION_VBA, status=ModelSectionStatus.AVAILABLE, count=0),
            )

        try:
            project = await self._vba.get_project_info(workbook_id)
        except AccessVbomDisabledError as exc:
            return _vba_failure(VbaAccessStatus.VBOM_DISABLED, SECTION_VBA, exc)
        except VbaProjectLockedError as exc:
            return _vba_failure(VbaAccessStatus.LOCKED, SECTION_VBA, exc)
        except ExcelVbaMcpError as exc:
            return _vba_failure(VbaAccessStatus.UNAVAILABLE, SECTION_VBA, exc)
        except Exception as exc:  # noqa: BLE001 - never abort the model on VBA
            return _vba_failure(VbaAccessStatus.UNAVAILABLE, SECTION_VBA, exc)

        components = project.components
        procedures: list[VbaProcedureInfo] = []
        procedure_diagnostics: list[str] = []
        for component in components:
            try:
                component_procedures = await self._vba.list_procedures(
                    workbook_id, component.component_name
                )
            except Exception as exc:  # noqa: BLE001 - per-component isolation
                procedure_diagnostics.append(
                    f"{SECTION_VBA}: procedures for {component.component_name!r}: "
                    f"{_describe_error(exc)}"
                )
                continue
            procedures.extend(component_procedures)

        if procedure_diagnostics:
            section = ModelSection(
                name=SECTION_VBA,
                status=ModelSectionStatus.PARTIAL,
                count=len(components),
                diagnostics=procedure_diagnostics,
            )
        else:
            section = ModelSection(
                name=SECTION_VBA,
                status=ModelSectionStatus.AVAILABLE,
                count=len(components),
            )
        return project, procedures, VbaAccessStatus.AVAILABLE, section

    # ------------------------------------------------------------------
    # Bookkeeping / audit (non-COM)
    # ------------------------------------------------------------------
    def _resolve_record(self, workbook_id: str) -> WorkbookInfo:
        """Resolve ``workbook_id`` to its open :class:`WorkbookInfo` (no COM)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise WorkbookNotFoundError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            context={"workbook_id": workbook_id},
        )

    def _record_audit(
        self,
        tool: str,
        params: dict[str, Any],
        result: dict[str, Any],
        record: WorkbookInfo,
    ) -> None:
        self._audit.record(
            tool=tool,
            operation_class="read_only",
            params=params,
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": record.source_path or ""},
            result=result,
        )


# ---------------------------------------------------------------------------
# Module-level helpers (pure, no COM)
# ---------------------------------------------------------------------------
def _utc_now() -> datetime:
    return datetime.now(UTC)


def _describe_error(exc: BaseException) -> str:
    """Return a compact, non-secret description of an exception."""

    message = str(exc).strip()
    return message or exc.__class__.__name__


def _vba_failure(
    access: VbaAccessStatus, section_name: str, exc: BaseException
) -> tuple[None, list[VbaProcedureInfo], VbaAccessStatus, ModelSection]:
    """Build the VBA failure tuple (no project, no procedures, a diagnostic)."""

    diagnostic = f"{section_name}: {_describe_error(exc)}"
    section = ModelSection(
        name=section_name,
        status=ModelSectionStatus.UNAVAILABLE,
        count=0,
        diagnostics=[diagnostic],
    )
    return None, [], access, section


def _name_on_sheet(name: DefinedNameInfo, sheet_name: str) -> bool:
    """Return whether a defined name is scoped to (or targets) ``sheet_name``."""

    if name.sheet_name == sheet_name:
        return True
    return name.scope is NameScope.WORKBOOK and name.target_sheet == sheet_name


def _rescope(section: ModelSection, count: int) -> ModelSection:
    """Return a copy of ``section`` with the sheet-scoped item count."""

    return ModelSection(
        name=section.name,
        status=section.status,
        count=count,
        diagnostics=list(section.diagnostics),
    )


def content_hash(model: WorkbookModel | SheetModel) -> str:
    """Return the SHA-256 of the model content (excluding timestamps + hash).

    The hash is computed over the JSON dump with ``model_hash`` / ``created_at`` /
    ``updated_at`` excluded, so an unchanged workbook always yields the same hash
    while the timestamps still change per read.
    """

    payload = model.model_dump(mode="json", exclude={"model_hash", "created_at", "updated_at"})
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


__all__ = [
    "SECTION_CONTROLS",
    "SECTION_NAMES",
    "SECTION_PIVOTS",
    "SECTION_SHAPES",
    "SECTION_SHEETS",
    "SECTION_TABLES",
    "SECTION_VBA",
    "WorkbookModelService",
    "content_hash",
]
