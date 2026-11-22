"""Generate the full-access MCP end-to-end workflow fixture (task_b7, fix_007).

This is a committed, idempotent fixture-authoring tool (not production code). It
builds one macro-enabled workbook through a **hidden, non-activating**
``Excel.Application`` created with ``DispatchEx`` and adds a small deterministic
standard module through VBIDE:

* ``FullAccessWorkflow.xlsm``

  - ``Sheet1`` carries a real Excel Table (ListObject) ``tblSettings`` over
    ``A1:B3`` — headers ``SettingKey`` | ``SettingValue`` and two data rows
    ``factor`` | ``5`` and ``label`` | ``alpha`` — plus ``Sheet1!E1`` with a known
    value (``configured``) and a workbook-scoped defined name ``SettingsKey`` that
    refers to ``Sheet1!$E$1``.
  - ``Result`` is a second worksheet initialised to ``A1 = 0`` and ``A2 = ""``.

  A standard module ``modWorkflow`` carries two public procedures used by the real
  MCP-protocol e2e test (``tests/integration_excel/test_mcp_full_access_workflow.py``):

  - ``Recalculate`` — reads the ``factor`` value from ``tblSettings`` and writes an
    observable result (``Result!A1 = factor * 2`` and ``Result!A2 = "factor=<n>"``).
    This is the **table-driven** background-safe macro: two consecutive
    ``table_update_rows`` + ``macro_run`` cycles must produce different
    ``Result!A1`` values, proving the macro reads the live table state.
  - ``ShowMessage`` — shows a ``MsgBox`` (interactive; the R13 preflight must veto
    it, never run it).

  ``Recalculate`` deliberately avoids every interactive-UI token in
  ``security/macro_policy`` / ``MacroService.INTERACTIVE_TOKENS`` (no ``MsgBox`` /
  ``.Show`` / ``.Run`` / ``DoEvents`` / ``DisplayAlerts`` / ``Shell`` / …) and
  never calls ``.Save``, so it is eligible for ``background_safe`` classification
  and safe to run repeatedly in the background.

The generator is idempotent: it rebuilds the fixture from scratch inside a
temporary directory, reopens it read-only to verify the table, ``Result`` sheet,
defined name, ``E1`` value and component inventory, and only then copies it into
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

#: Excel ``ListObjects.Add`` ``SourceType`` code for a range source.
XL_SRC_RANGE = 1

#: Excel ``ListObjects.Add`` ``XlListObjectHasHeaders`` code (``xlYes``).
XL_LIST_OBJECT_HAS_HEADERS = 1

#: The name of the ListObject the ``Recalculate`` macro reads.
SETTINGS_TABLE_NAME = "tblSettings"

#: The table headers (``SettingKey`` | ``SettingValue``).
SETTINGS_KEY_HEADER = "SettingKey"
SETTINGS_VALUE_HEADER = "SettingValue"

#: The initial ``factor`` / ``label`` data rows of ``tblSettings``.
SETTINGS_TABLE_ROWS: tuple[tuple[str, object], ...] = (("factor", 5), ("label", "alpha"))

#: The known value written to ``Sheet1!E1`` (the target of ``SettingsKey``).
SETTINGS_KEY_VALUE = "configured"

#: Name of the observable output worksheet written by ``Recalculate``.
RESULT_SHEET_NAME = "Result"

#: VBA source for ``modWorkflow``. ``Recalculate`` reads the live ``factor`` value
#: from ``tblSettings`` and writes an observable result; it avoids every
#: interactive-UI token so it is eligible for ``background_safe`` classification.
#: ``ShowMessage`` intentionally triggers the R13 interactive-UI veto.
MOD_WORKFLOW_SOURCE = """\
Option Explicit

Public Sub Recalculate()
    Dim ws As Worksheet, tbl As ListObject, out As Worksheet
    Dim factor As Double
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set tbl = ws.ListObjects("tblSettings")
    factor = CDbl(tbl.DataBodyRange.Cells(1, 2).Value)
    Set out = ThisWorkbook.Worksheets("Result")
    out.Range("A1").Value = factor * 2
    out.Range("A2").Value = "factor=" & CStr(factor)
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

        # Real Excel Table (ListObject) ``tblSettings`` over A1:B3. Write the
        # header + data block first, then create the ListObject with a header row.
        sheet.Range("A1").Value = SETTINGS_KEY_HEADER
        sheet.Range("B1").Value = SETTINGS_VALUE_HEADER
        for offset, (key, value) in enumerate(SETTINGS_TABLE_ROWS, start=2):
            sheet.Range(f"A{offset}").Value = key
            sheet.Range(f"B{offset}").Value = value
        table = sheet.ListObjects.Add(
            XL_SRC_RANGE, sheet.Range("A1:B3"), False, XL_LIST_OBJECT_HAS_HEADERS
        )
        table.Name = SETTINGS_TABLE_NAME

        # Observable output sheet written by ``Recalculate``.
        result_sheet = wb.Worksheets.Add()
        result_sheet.Name = RESULT_SHEET_NAME
        result_sheet.Range("A1").Value = 0
        result_sheet.Range("A2").Value = ""

        # Known value in E1 plus the workbook-scoped defined name -> Sheet1!$E$1.
        sheet.Range("E1").Value = SETTINGS_KEY_VALUE
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

        # Table (ListObject) structure on Sheet1.
        table_info: dict[str, object] = {}
        sheet1 = wb.Worksheets("Sheet1")
        if int(sheet1.ListObjects.Count) > 0:
            table = sheet1.ListObjects(SETTINGS_TABLE_NAME)
            headers = [
                str(table.HeaderRowRange.Cells(1, col).Value2)
                for col in range(1, int(table.ListColumns.Count) + 1)
            ]
            body = table.DataBodyRange
            table_info = {
                "name": str(table.Name),
                "headers": headers,
                "row_count": int(body.Rows.Count),
                "first_row": [body.Cells(1, 1).Value, body.Cells(1, 2).Value],
            }

        # Result sheet initial state.
        result_a1: object = None
        result_a2: object = None
        has_result = RESULT_SHEET_NAME in sheet_names
        if has_result:
            result_sheet = wb.Worksheets(RESULT_SHEET_NAME)
            result_a1 = result_sheet.Range("A1").Value
            result_a2 = result_sheet.Range("A2").Value

        return {
            "path": str(path),
            "protection": int(proj.Protection),
            "components": components,
            "sheet_names": sheet_names,
            "defined_names": defined_names,
            "e1_value": e1_value,
            "table": table_info,
            "has_result_sheet": has_result,
            "result_a1": result_a1,
            "result_a2": result_a2,
            "has_modWorkflow": "modWorkflow" in names,
            "has_recalculate": "Recalculate" in module_source,
            "has_show_message": "ShowMessage" in module_source,
        }
    finally:
        wb.Close(SaveChanges=False)


def _verify_ok(info: dict[str, object]) -> list[str]:
    """Return a list of unmet expectations for a built fixture (empty == OK)."""

    problems: list[str] = []
    if not info.get("has_modWorkflow"):
        problems.append("missing modWorkflow component")
    if not info.get("has_recalculate"):
        problems.append("modWorkflow missing Recalculate")
    if not info.get("has_show_message"):
        problems.append("modWorkflow missing ShowMessage")
    if "Sheet1" not in (info.get("sheet_names") or []):
        problems.append("missing Sheet1 worksheet")
    if not info.get("has_result_sheet"):
        problems.append("missing Result worksheet")
    if info.get("result_a1") != 0:
        problems.append(f"Result!A1 initial value {info.get('result_a1')!r} != 0")
    if "SettingsKey" not in (info.get("defined_names") or []):
        problems.append("missing SettingsKey defined name")
    if info.get("e1_value") != SETTINGS_KEY_VALUE:
        problems.append(f"E1 value {info.get('e1_value')!r} != {SETTINGS_KEY_VALUE!r}")

    table = info.get("table") or {}
    if not table:
        problems.append("missing tblSettings ListObject")
    else:
        expected_headers = [SETTINGS_KEY_HEADER, SETTINGS_VALUE_HEADER]
        if table.get("headers") != expected_headers:
            problems.append(f"table headers {table.get('headers')!r} != {expected_headers!r}")
        if table.get("row_count") != len(SETTINGS_TABLE_ROWS):
            problems.append(
                f"table row count {table.get('row_count')!r} != {len(SETTINGS_TABLE_ROWS)}"
            )
        first_key, first_value = SETTINGS_TABLE_ROWS[0]
        if list(table.get("first_row") or []) != [first_key, first_value]:
            problems.append(
                f"table first data row {table.get('first_row')!r} != {[first_key, first_value]!r}"
            )
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