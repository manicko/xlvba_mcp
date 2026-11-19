"""Read-only VBE command discovery and dialog observation (task_020).

This module is the **only** place that looks at the VBE command layer. It is
strictly non-activating:

* :meth:`VBECompileUI.discover_command` enumerates ``VBE.CommandBars`` and probes
  ``FindControl`` **read-only**. It never calls ``CommandBarControl.Execute`` and
  never expands a UIA menu. On the validated build (Excel 16.0.14332) the Debug
  bar has no Compile entry, ``FindControl(578)`` raises ``E_INVALIDARG`` and the
  only route that reaches Compile (UIA ``Debug`` menu expansion) creates visible
  popup windows and steals foreground. Discovery therefore reports
  ``route_usable = False`` and the production compile service fails closed to
  ``COMPILE_INDETERMINATE`` (see
  ``.ai/audit/hidden-compile-research-2026-11-19.md``).
* :meth:`VBECompileUI.observe_dialog` only **reads** window text when a dialog
  already exists. It never calls ``SetFocus``, click, type, ``SetForegroundWindow``
  or dismisses anything. It is **not** used by the background compile path.
* :class:`DialogReader` reads a dialog's text via UI Automation
  (``CurrentName`` / ``TextPattern``) without activating the dialog.

Every method is synchronous and intended to run on the single-STA worker thread
when it touches COM. UIA is imported lazily so importing this module never
requires ``comtypes`` / ``pywin32``.
"""

from __future__ import annotations

import contextlib
from typing import Any

from pydantic import Field

from excel_vba_mcp.domain.models import DomainModel

# VBE command-bar facts (research §3.1, R03).
DEBUG_MENU_CAPTION = "Debug"
COMPILE_CONTROL_ID = 578
_COMPILE_CAPTION_MARKER = "compil"

# Win32 / UIA constants used by the read-only dialog observer.
_DIALOG_CLASS = "#32770"
_UIA_TREE_SCOPE_SUBTREE = 7
_UIA_TEXT_PATTERN_ID = 10014
_CONTROL_TYPE_TEXT = 50020


class CommandBarControlInfo(DomainModel):
    """One ``CommandBarControl`` observed read-only (no ``Execute``)."""

    bar: str
    id: int
    caption: str
    enabled: bool


class FindControlProbe(DomainModel):
    """Outcome of one ``CommandBar.FindControl`` probe."""

    args: list[str]
    result: str | None = None
    error: str | None = None


class CommandDiscovery(DomainModel):
    """Read-only VBE command discovery evidence (never an execution).

    ``executed`` is always ``False`` and ``uia_expand_attempted`` is always
    ``False``: the production service must not invoke a control or expand a UIA
    menu. ``route_usable`` is ``True`` only when a Compile ``CommandBarControl``
    was actually found, which does not happen on the validated build.
    """

    attempted: bool = False
    executed: bool = False
    execute_reason: str | None = None
    bars_count: int | None = None
    bars_error: str | None = None
    debug_controls_count: int = 0
    debug_controls: list[CommandBarControlInfo] = Field(default_factory=list)
    has_compile_entry: bool = False
    find_control: list[FindControlProbe] = Field(default_factory=list)
    compile_caption_scan: list[CommandBarControlInfo] = Field(default_factory=list)
    route_usable: bool = False
    uia_expand_attempted: bool = False
    notes: list[str] = Field(default_factory=list)


class DialogWindowInfo(DomainModel):
    """A dialog window observed read-only (never activated)."""

    hwnd: int
    title: str
    class_name: str
    text: str | None = None


class DialogObservation(DomainModel):
    """Read-only dialog observation (no focus / click / type / dismiss)."""

    present: bool = False
    read_only: bool = True
    windows: list[DialogWindowInfo] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)


def _safe_str(value: Any, default: str = "") -> str:
    """Read a COM string property defensively."""

    try:
        return str(value)
    except Exception:  # noqa: BLE001 - defensive COM read
        return default


def _error_text(exc: BaseException) -> str:
    """Return a compact ``Type: message`` string for an exception."""

    message = str(exc) or exc.__class__.__name__
    return f"{exc.__class__.__name__}: {message[:300]}"


class VBECompileUI:
    """Read-only observer over the VBE command layer and dialogs.

    The class holds no COM references between calls. It never executes a control,
    never expands a menu, never focuses/clicks/types and never dismisses a dialog.
    """

    def discover_command(self, vbe: Any) -> CommandDiscovery:
        """Enumerate the VBE command bars read-only and report Compile availability.

        Args:
            vbe: The live ``VBE`` COM proxy (STA thread).

        Returns:
            A :class:`CommandDiscovery` describing what was observed. On the
            validated build it records no Compile entry, ``FindControl`` errors and
            ``route_usable = False``.
        """

        discovery = CommandDiscovery(attempted=True)
        discovery.notes.append(
            "Read-only discovery only: no CommandBarControl.Execute and no UIA menu expansion."
        )

        try:
            bars = vbe.CommandBars
            discovery.bars_count = int(bars.Count)
        except Exception as exc:  # noqa: BLE001 - defensive COM read
            discovery.bars_error = _error_text(exc)
            discovery.execute_reason = "VBE CommandBars unavailable"
            return discovery

        debug_bar = self._get_debug_bar(bars, discovery)
        if debug_bar is not None:
            self._inspect_debug_bar(debug_bar, discovery)
            self._probe_find_control(debug_bar, discovery)

        self._scan_all_bars(bars, discovery)

        if not discovery.has_compile_entry and discovery.compile_caption_scan:
            discovery.has_compile_entry = True
        discovery.route_usable = bool(discovery.compile_caption_scan)
        if discovery.route_usable:
            discovery.execute_reason = None
        else:
            discovery.execute_reason = (
                "no Compile CommandBarControl found on this build; "
                "the only route (UIA Debug-menu expansion) is non-activating-unsafe"
            )
        return discovery

    def observe_dialog(self, read_only: bool = True) -> DialogObservation:
        """Read any existing ``#32770`` dialog text without activating it.

        This is **not** used by the background compile path. It exists only for
        read-only diagnostics on an already-visible dialog. It never calls
        ``SetFocus``, click, type, ``SetForegroundWindow``, ``ShowWindow`` or
        ``PostMessage``.

        Args:
            read_only: Must be ``True``. ``False`` is rejected because interactive
                dialog handling is not supported.

        Returns:
            A :class:`DialogObservation` with the visible dialog windows and their
            text (when readable).

        Raises:
            ValueError: If ``read_only`` is ``False``.
        """

        if not read_only:
            raise ValueError(
                "VBECompileUI.observe_dialog is read-only; interactive dialog handling "
                "is not supported and would violate the background execution contract."
            )

        observation = DialogObservation(read_only=True)
        observation.notes.append(
            "Read-only observation only: never SetFocus/click/type/dismiss. "
            "Not used by the background compile path."
        )
        try:
            import win32gui  # type: ignore[import-untyped]  # noqa: PLC0415
        except Exception as exc:  # noqa: BLE001 - non-Windows / pywin32 absent
            observation.notes.append(f"win32gui unavailable: {_error_text(exc)}")
            return observation

        reader = DialogReader()
        windows: list[DialogWindowInfo] = []

        def _callback(hwnd: int, _ctx: Any) -> bool:
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if win32gui.GetClassName(hwnd) != _DIALOG_CLASS:
                    return True
                title = win32gui.GetWindowText(hwnd)
                text = reader.read_dialog_text(int(hwnd))
            except Exception:  # noqa: BLE001 - defensive per-window read
                return True
            windows.append(
                DialogWindowInfo(hwnd=int(hwnd), title=title, class_name=_DIALOG_CLASS, text=text)
            )
            return True

        with contextlib.suppress(Exception):
            win32gui.EnumWindows(_callback, None)

        observation.windows = windows
        observation.present = bool(windows)
        return observation

    # ------------------------------------------------------------------
    # Internal helpers (STA thread; read-only COM access)
    # ------------------------------------------------------------------
    @staticmethod
    def _get_debug_bar(bars: Any, discovery: CommandDiscovery) -> Any | None:
        """Return the ``Debug`` command bar, or ``None`` after recording the error."""

        try:
            return bars[DEBUG_MENU_CAPTION]
        except Exception as exc:  # noqa: BLE001 - defensive COM read
            discovery.notes.append(f"Debug bar lookup failed: {_error_text(exc)}")
            return None

    def _inspect_debug_bar(self, debug_bar: Any, discovery: CommandDiscovery) -> None:
        """Enumerate the Debug bar controls read-only."""

        controls: list[CommandBarControlInfo] = []
        try:
            for control in debug_bar.Controls:
                described = self._describe_control(DEBUG_MENU_CAPTION, control)
                if described is not None:
                    controls.append(described)
        except Exception as exc:  # noqa: BLE001 - defensive COM read
            discovery.notes.append(f"Debug bar enumeration failed: {_error_text(exc)}")
        discovery.debug_controls = controls
        discovery.debug_controls_count = len(controls)
        discovery.has_compile_entry = any(
            _COMPILE_CAPTION_MARKER in control.caption.casefold() for control in controls
        )

    def _probe_find_control(self, debug_bar: Any, discovery: CommandDiscovery) -> None:
        """Probe ``FindControl`` read-only (no ``Execute``)."""

        for args in (
            (COMPILE_CONTROL_ID,),
            (COMPILE_CONTROL_ID, "Compile"),
            (COMPILE_CONTROL_ID, "", True),
            (COMPILE_CONTROL_ID, "", True, 1),
        ):
            probe = FindControlProbe(args=[str(arg) for arg in args])
            try:
                found = debug_bar.FindControl(*args)
            except Exception as exc:  # noqa: BLE001 - build-specific failure is evidence
                probe.error = _error_text(exc)
            else:
                probe.result = (
                    None if found is None else _safe_str(getattr(found, "Caption", found))
                )
            discovery.find_control.append(probe)

    def _scan_all_bars(self, bars: Any, discovery: CommandDiscovery) -> None:
        """Scan every command bar for a Compile caption (read-only)."""

        hits: list[CommandBarControlInfo] = []
        try:
            iterator = iter(bars)
        except Exception as exc:  # noqa: BLE001 - defensive COM read
            discovery.notes.append(f"CommandBars iteration failed: {_error_text(exc)}")
            return

        for bar in iterator:
            try:
                bar_name = _safe_str(getattr(bar, "Name", ""))
                for control in bar.Controls:
                    caption = _safe_str(getattr(control, "Caption", ""))
                    if _COMPILE_CAPTION_MARKER not in caption.casefold():
                        continue
                    described = self._describe_control(bar_name, control)
                    if described is not None:
                        hits.append(described)
            except Exception:  # noqa: BLE001 - a single bad bar must not abort discovery
                continue

        discovery.compile_caption_scan = hits

    @staticmethod
    def _describe_control(bar_name: str, control: Any) -> CommandBarControlInfo | None:
        """Build a :class:`CommandBarControlInfo`, or ``None`` when unreadable."""

        try:
            return CommandBarControlInfo(
                bar=bar_name,
                id=int(control.Id),
                caption=_safe_str(control.Caption),
                enabled=bool(control.Enabled),
            )
        except Exception:  # noqa: BLE001 - defensive COM read
            return None


class DialogReader:
    """Read-only UI Automation text reader for one dialog window handle.

    The reader never activates or focuses the dialog. It uses
    ``IUIAutomation.ElementFromHandle`` + ``CurrentName`` and, when available,
    the ``TextPattern`` ``DocumentRange.GetText``. ``comtypes`` is imported
    lazily so importing this module never requires it.
    """

    def read_dialog_text(self, hwnd: int) -> str | None:
        """Return the concatenated visible text of the dialog, or ``None``.

        Args:
            hwnd: The dialog window handle.

        Returns:
            The dialog text (deduplicated, newline-joined) or ``None`` when no
            text can be read non-intrusively.
        """

        try:
            import comtypes  # noqa: PLC0415
            import comtypes.client  # noqa: PLC0415
        except Exception:  # noqa: BLE001 - comtypes absent
            return None

        try:
            module = comtypes.client.GetModule("UIAutomationCore.dll")
            automation = comtypes.CoCreateInstance(
                module.CUIAutomation().IPersist_GetClassID(),
                interface=module.IUIAutomation,
                clsctx=comtypes.CLSCTX_INPROC_SERVER,
            )
            root = automation.ElementFromHandle(int(hwnd))
            if root is None:
                return None

            texts: list[str] = []
            root_name = _safe_str(getattr(root, "CurrentName", ""))
            if root_name:
                texts.append(root_name)

            condition = automation.CreateTrueCondition()
            found = root.FindAll(_UIA_TREE_SCOPE_SUBTREE, condition)
            for index in range(int(found.Length)):
                try:
                    element = found.GetElement(index)
                    control_type = int(element.CurrentControlType)
                except Exception:  # noqa: BLE001 - defensive per-element read
                    continue
                if control_type != _CONTROL_TYPE_TEXT:
                    continue
                value = _safe_str(getattr(element, "CurrentName", ""))
                if value:
                    texts.append(value)

            if len(texts) <= 1:
                document_text = self._read_text_pattern(root, module)
                if document_text:
                    texts.append(document_text)

            joined = "\n".join(dict.fromkeys(texts))
            return joined or None
        except Exception:  # noqa: BLE001 - UIA unavailable / non-Windows
            return None

    @staticmethod
    def _read_text_pattern(root: Any, module: Any) -> str | None:
        """Best-effort ``TextPattern`` read (never activates the dialog)."""

        try:
            pattern = root.GetCurrentPattern(_UIA_TEXT_PATTERN_ID)
            text_pattern = pattern.QueryInterface(module.IUIAutomationTextPattern)
            value = text_pattern.DocumentRange.GetText(-1)
        except Exception:  # noqa: BLE001 - TextPattern is optional
            return None
        text = _safe_str(value)
        return text or None


__all__ = [
    "CommandBarControlInfo",
    "CommandDiscovery",
    "DialogObservation",
    "DialogReader",
    "DialogWindowInfo",
    "FindControlProbe",
    "VBECompileUI",
]
