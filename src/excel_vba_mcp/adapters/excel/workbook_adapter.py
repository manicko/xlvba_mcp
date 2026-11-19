"""STA-local workbook description helpers (task_012).

:class:`WorkbookAdapter` turns a live ``Workbook`` COM proxy into serializable
domain DTOs. It runs **only on the dedicated COM STA worker thread** (inside a
submitted callable); it never touches the event loop and never crosses the worker
boundary with a proxy. All methods accept the STA-local ``wb`` proxy as ``Any``
and return Pydantic DTOs / primitives so the worker's return-boundary guard
(``_ensure_dto``) accepts them.

Design constraints (ARCH-009, addendum §2.2):

* No ``.Activate`` / ``.Select`` and no visibility changes — only direct
  object-model reads.
* No COM is imported at module import time; the adapter is a pure helper over a
  proxy supplied by the caller.
* Attribute reads are defensive: a partially-initialized or protected object
  yields a safe default rather than leaking a ``com_error``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_vba_mcp.domain.models import WorkbookInfo
from excel_vba_mcp.domain.statuses import SessionMode

# Workbook extensions that can carry a VBA project (spec §7.2, Appendix A).
_VBA_EXTENSIONS: frozenset[str] = frozenset({".xlsm", ".xlsb", ".xlsa", ".xlam", ".xltm"})

# Fallback Excel version when the Application object cannot be read.
_UNKNOWN_EXCEL_VERSION = "unknown"


def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively."""

    try:
        return str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _safe_bool(value: Any, default: bool = False) -> bool:
    """Read a COM boolean property defensively."""

    try:
        return bool(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _format_from_name(name: str) -> str:
    """Return the lowercase extension (without dot) used as ``WorkbookInfo.format``."""

    suffix = Path(name).suffix.lower()
    return suffix[1:] if suffix else ""


def _has_vba(name: str) -> bool:
    """Return whether the workbook extension can carry a VBA project."""

    return Path(name).suffix.lower() in _VBA_EXTENSIONS


class WorkbookAdapter:
    """Stateless, STA-local reader that describes an open ``Workbook`` proxy.

    The adapter holds no references between calls; every method receives the
    STA-local ``wb`` proxy and returns only DTOs / primitives.
    """

    def describe(
        self,
        wb: Any,
        *,
        workbook_id: str,
        session_id: str,
        mode: SessionMode,
        source_path: str | None = None,
    ) -> WorkbookInfo:
        """Build a :class:`WorkbookInfo` from a live ``Workbook`` proxy.

        Args:
            wb: The STA-local ``Workbook`` COM proxy.
            workbook_id: Server-assigned stable workbook id.
            session_id: Owning session id.
            mode: Session open mode (inspect / execute).
            source_path: Original allowlisted path (the working copy path is read
                from ``wb.FullName``).

        Returns:
            A serializable :class:`WorkbookInfo`. ``path`` is the actual working
            path (``wb.FullName``); ``read_only`` mirrors ``wb.ReadOnly``.
        """

        name = _safe_str(getattr(wb, "Name", ""))
        full_name = _safe_str(getattr(wb, "FullName", "")) or name
        read_only = _safe_bool(getattr(wb, "ReadOnly", False))
        excel_version = (
            _safe_str(
                getattr(getattr(wb, "Application", None), "Version", ""),
                _UNKNOWN_EXCEL_VERSION,
            )
            or _UNKNOWN_EXCEL_VERSION
        )

        return WorkbookInfo(
            workbook_id=workbook_id,
            name=name,
            format=_format_from_name(name or full_name),
            path=full_name,
            source_path=source_path,
            read_only=read_only,
            has_vba=_has_vba(name or full_name),
            excel_version=excel_version,
            session_id=session_id,
            mode=mode,
        )

    def sheet_names(self, wb: Any) -> list[str]:
        """Return the worksheet names of ``wb`` (direct object-model read)."""

        try:
            return [str(sheet.Name) for sheet in wb.Worksheets]
        except Exception:  # noqa: BLE001 - defensive COM read
            return []

    def is_dirty(self, wb: Any) -> bool:
        """Return whether ``wb`` has unsaved changes (``not wb.Saved``)."""

        return not _safe_bool(getattr(wb, "Saved", True), default=True)


__all__ = ["WorkbookAdapter"]
