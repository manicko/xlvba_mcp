"""Read-only, non-activating Windows UI adapters (task_020).

These adapters only observe: they never call ``SetFocus``, click, type,
``ShowWindow``, ``SetForegroundWindow`` or dismiss a dialog. The background
compile path uses :class:`VBECompileUI.discover_command` read-only and never
expands a UIA menu.
"""

from excel_vba_mcp.adapters.windows_ui.vbe_compile_ui import (
    CommandBarControlInfo,
    CommandDiscovery,
    DialogObservation,
    DialogReader,
    DialogWindowInfo,
    FindControlProbe,
    VBECompileUI,
)

__all__ = [
    "CommandBarControlInfo",
    "CommandDiscovery",
    "DialogObservation",
    "DialogReader",
    "DialogWindowInfo",
    "FindControlProbe",
    "VBECompileUI",
]
