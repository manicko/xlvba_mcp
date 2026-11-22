"""Generate the full-access MCP end-to-end workflow fixture (task_b7, fix_007).

This is a committed, idempotent fixture-authoring tool (not production code). It
builds one macro-enabled workbook through a **hidden, non-activating**
``Excel.Application`` created with ``DispatchEx`` and adds a small deterministic
standard module through VBIDE:

* ``FullAccessWorkflow.xlsm``
  ``Sheet1`` with a known value in ``E1`` and a workbook-scoped defined name
  ``SettingsKey`` that refers to ``Sheet1!$E$1``. A standard module ``modWorkflow``
  carries three public procedures used by the real MCP-protocol e2e test
  (``tests/integration_excel/test_mcp_full_access_workflow.py``):

  - ``WriteCell``  — sets ``Sheet1!A1 = 42`` (background-safe).
  - ``WriteCell2`` — sets ``Sheet1!A2 = 7``  (background-safe).
  - ``ShowMessage``— shows a ``MsgBox`` (interactive; the R13 preflight must
    veto it, never run it).

  ``WriteCell`` / ``WriteCell2`` deliberately avoid every interactive-UI token in
  ``security/macro_policy`` / ``MacroService.INTERACTIVE_TOKENS`` (no
  ``MsgBox`` / ``.Show`` / ``.Run`` / ``DoEvents`` / ``DisplayAlerts`` / …), so
  they can be legitimately classified ``background_safe``.

The generator is idempotent: it rebuilds the fixture from scratch inside a
temporary directory, reopens it read-only to verify the sheet, defined name,
``E1`` value and component inventory, and only then copies it into
``tests/fixtures`` — so an interrupted or mismatched run never corrupts an
existing fixture and never leaves temp files behind.

It requires Windows + Excel + VBOM trust and must run serially::

    uv run python tests/fixtures/build_full_access_fixture.py
    uv run python tests/fixtures/build_full_access_fixture.py --verify

No ``.Activate`` / ``.Select`` / ``SendKeys`` is used anywhere; the owned Excel
instance stays ``Visible = False`` for its whole lifetime (BG03/BG04).
"""

from __future__ import annotations

import argparse
import contextlib
import shutil
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

REPO_ROOT = Path(__file__).resolve().parents[2]
FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures"

#: The single fixture this generator produces.
WORKFLOW_FIXTURE = "FullAccessWorkflow.xlsm"

#: Excel ``FileFormat`` for a macro-enabled workbook (``.xlsm``).
XLSM_FILE_FORMAT = 52

#: VBIDE ``VBComponents.Add`` type code for a standard module.
VB_COMPONENT_STANDARD = 1

#: The known value written to ``Sheet1!E1`` (the target of ``SettingsKey``).
SETTINGS_KEY_VALUE = "configured"

#: VBA source for ``modWorkflow``. The two writer subs avoid every interactive-UI
#: token so they are eligible for ``background_safe`` classification; the message
#: sub intentionally triggers the R13 interactive-UI veto.
MOD_WORKFLOW_SOURCE = """\
Option Explicit

Public Sub WriteCell()
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = 42
End Sub

Public Sub WriteCell2()
    ThisWorkbook.Worksheets("Sheet1").Range("A2").Value = 7
End Sub

Public Sub ShowMessage()
    MsgBox "interactive", vbInformation, "probe"
End Sub
"""


def _normalize_single_sheet(excel: object, wb: object) -> None:
    """Reduce the new workbook to exactly one worksheet named ``Sheet1``.

    A new workbook's default sheet count/name depends on the host's
    ``Application.SheetsInNewWorkbook``; renaming the first sheet and deleting any
    extras makes the fixture deterministic.
    """

    worksheets = wb.Worksheets
    worksheets(1).Name = "Sheet1"
    excel.DisplayAlerts = False
    for index in range(int(worksheets.Count), 1, -1):
        worksheets(index).Delete()


def _build_workflow_fixture(app: object, target: Path) -> None:
    """Build ``FullAccessWorkflow.xlsm`` and save it to ``target``."""

    excel = app
    wb = excel.Workbooks.Add()
    try:
        _normalize_single_sheet(excel, wb)
        sheet = wb.Worksheets("Sheet1")
        sheet.Range("E1").Value = SETTINGS_KEY_VALUE

        # Workbook-scoped defined name -> Sheet1!$E$1.
        wb.Names.Add("SettingsKey", "=Sheet1!$E$1")

        proj = wb.VBProject
        mod = proj.VBComponents.Add(VB_COMPONENT_STANDARD)
        mod.Name = "modWorkflow"
        mod.CodeModule.AddFromString(MOD_WORKFLOW_SOURCE)

        wb.SaveAs(str(target), XLSM_FILE_FORMAT)
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def _verify(app: object, path: Path) -> dict[str, object]:
    """Reopen a built fixture read-only and return a structural summary."""

    wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, Notify=False, AddToMru=False)
    try:
        proj = wb.VBProject
        components: list[dict[str, object]] = []
        for comp in proj.VBComponents:
            components.append(
                {
                    "name": str(comp.Name),
                    "type": int(comp.Type),
                    "lines": int(comp.CodeModule.CountOfLines),
                }
            )
        names = {c["name"] for c in components}
        defined_names = [str(item.Name) for item in wb.Names]
        sheet_names = [str(item.Name) for item in wb.Worksheets]
        module_source = ""
        if "modWorkflow" in names:
            code_module = proj.VBComponents("modWorkflow").CodeModule
            total = int(code_module.CountOfLines)
            if total > 0:
                module_source = str(code_module.Lines(1, total))
        e1_value = wb.Worksheets("Sheet1").Range("E1").Value
        return {
            "path": str(path),
            "protection": int(proj.Protection),
            "components": components,
            "sheet_names": sheet_names,
            "defined_names": defined_names,
            "e1_value": e1_value,
            "has_modWorkflow": "modWorkflow" in names,
            "has_write_cell": "WriteCell" in module_source,
            "has_write_cell2": "WriteCell2" in module_source,
            "has_show_message": "ShowMessage" in module_source,
        }
    finally:
        wb.Close(SaveChanges=False)


def _verify_ok(info: dict[str, object]) -> list[str]:
    """Return a list of unmet expectations for a built fixture (empty == OK)."""

    problems: list[str] = []
    if not info.get("has_modWorkflow"):
        problems.append("missing modWorkflow component")
    if not info.get("has_write_cell"):
        problems.append("modWorkflow missing WriteCell")
    if not info.get("has_write_cell2"):
        problems.append("modWorkflow missing WriteCell2")
    if not info.get("has_show_message"):
        problems.append("modWorkflow missing ShowMessage")
    if "Sheet1" not in (info.get("sheet_names") or []):
        problems.append("missing Sheet1 worksheet")
    if "SettingsKey" not in (info.get("defined_names") or []):
        problems.append("missing SettingsKey defined name")
    if info.get("e1_value") != SETTINGS_KEY_VALUE:
        problems.append(f"E1 value {info.get('e1_value')!r} != {SETTINGS_KEY_VALUE!r}")
    return problems


def _run(dest_dir: Path, verify: bool) -> int:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    exit_code = 0
    try:
        tmp_dir = Path(tempfile.mkdtemp(prefix="build_full_access_"))
        try:
            tmp_path = tmp_dir / WORKFLOW_FIXTURE
            _build_workflow_fixture(app, tmp_path)

            if verify:
                info = _verify(app, tmp_path)
                problems = _verify_ok(info)
                if problems:
                    print(f"[verify] FAILED {WORKFLOW_FIXTURE}: {problems}")
                    print(f"[verify] info: {info}")
                    return 1
                print(
                    f"[verify] {WORKFLOW_FIXTURE}: OK "
                    f"({info['sheet_names']}, {info['defined_names']})"
                )

            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_path = dest_dir / WORKFLOW_FIXTURE
            shutil.copy2(tmp_path, dest_path)
            print(f"[built] {WORKFLOW_FIXTURE} -> {dest_path} ({dest_path.stat().st_size} bytes)")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    finally:
        with contextlib.suppress(Exception):
            app.Quit()
        del app
        pythoncom.CoUninitialize()

    print("[done] fixture written to", dest_dir)
    return exit_code


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--verify",
        action="store_true",
        help="reopen the built fixture and assert its sheet / name / component inventory",
    )
    args = parser.parse_args(argv)
    return _run(FIXTURES_DIR, args.verify)


if __name__ == "__main__":
    sys.exit(main())