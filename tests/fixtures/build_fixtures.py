"""Generate Excel VBA test fixtures via COM + VBIDE (task_003, hybrid strategy).

This is a committed, idempotent generator. It creates the two COM/VBIDE-buildable
fixtures the G1 proof-of-concept gates need:

* ``CleanProject.xlsm``       - standard module, class module, UserForm with a
  control (``.frm`` + ``.frx``), document modules, valid references.
* ``CompileErrorProject.xlsm``- a project with a guaranteed compile error
  (``Option Explicit`` + an undefined identifier).

A locked project cannot be produced through VBIDE (``VBProject.Protection`` is
read-only from COM); ``LockedProject.xlsm`` is therefore a committed manual binary
and is intentionally not generated here.

The generator is idempotent: it rebuilds both fixtures from scratch and overwrites
the committed binaries. Each workbook is built inside a temporary directory and
copied into place only after a successful save, so an interrupted run never
corrupts an existing fixture and no temp files are left behind.

It requires Windows + Excel + VBOM trust and must run serially::

    uv run python tests/fixtures/build_fixtures.py
    uv run python tests/fixtures/build_fixtures.py --verify   # rebuild + reopen

This script is a fixture-authoring tool, not production code. It deliberately
mirrors the COM usage proven in ``.ai/audit/g1-poc-research-2026-11-19.md``:
``DispatchEx`` on an STA thread, ``VBComponents.Add`` type codes (1=standard,
2=class, 3=UserForm), ``Designer.Controls.Add`` for the button, ``Save`` before
``Export`` so the ``.frx`` exists.
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

# Excel ``FileFormat`` for a macro-enabled workbook (``.xlsm``).
XLSM_FILE_FORMAT = 52

# VBIDE ``VBComponents.Add`` type codes.
VB_COMPONENT_STANDARD = 1
VB_COMPONENT_CLASS = 2
VB_COMPONENT_USER_FORM = 3

# VBA source for each component. Fixed strings keep the generated source
# deterministic; only the surrounding binary container is opaque.
MOD_TEST_SOURCE = """\
Option Explicit

Public Function AddNumbers(ByVal a As Double, ByVal b As Double) As Double
    AddNumbers = a + b
End Function

Public Function NoArgs() As String
    NoArgs = "no-args"
End Function

Public Sub Echo(ByVal s As String)
    Debug.Print s
End Sub
"""

# ``Name`` is a VBA statement keyword, so the class exposes ``FullName`` instead
# of a bare ``Name`` field to avoid an identifier collision (project rule 9).
CLS_TEST_SOURCE = """\
Option Explicit

Public FullName As String
"""

FRM_TEST_SOURCE = """\
Option Explicit
"""

# Compile-error module. ``x`` is declared so the *only* undefined token is
# ``UndefinedIdentifier`` -> VBE reports "Variable not defined" reliably.
MOD_BAD_SOURCE = """\
Option Explicit

Public Sub Bad()
    Dim x As Long
    x = UndefinedIdentifier
End Sub
"""


def _build_clean_project(app: object, target: Path) -> None:
    """Build ``CleanProject.xlsm`` and save it to ``target``."""

    excel = app
    wb = excel.Workbooks.Add()
    try:
        proj = wb.VBProject

        mod = proj.VBComponents.Add(VB_COMPONENT_STANDARD)
        mod.Name = "modTest"
        mod.CodeModule.AddFromString(MOD_TEST_SOURCE)

        cls = proj.VBComponents.Add(VB_COMPONENT_CLASS)
        cls.Name = "clsTest"
        cls.CodeModule.AddFromString(CLS_TEST_SOURCE)

        frm = proj.VBComponents.Add(VB_COMPONENT_USER_FORM)
        frm.Name = "frmTest"
        frm.CodeModule.AddFromString(FRM_TEST_SOURCE)
        button = frm.Designer.Controls.Add("Forms.CommandButton.1", "btnOK", True)
        button.Caption = "OK"
        button.Left = 10
        button.Top = 10

        # SaveAs persists the UserForm designer resource (``.frx``) into the file.
        wb.SaveAs(str(target), XLSM_FILE_FORMAT)
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def _build_compile_error_project(app: object, target: Path) -> None:
    """Build ``CompileErrorProject.xlsm`` and save it to ``target``."""

    excel = app
    wb = excel.Workbooks.Add()
    try:
        proj = wb.VBProject
        mod = proj.VBComponents.Add(VB_COMPONENT_STANDARD)
        mod.Name = "modBad"
        mod.CodeModule.AddFromString(MOD_BAD_SOURCE)
        wb.SaveAs(str(target), XLSM_FILE_FORMAT)
        wb.Save()
    finally:
        wb.Close(SaveChanges=False)


def _build_all(app: object, dest_dir: Path) -> dict[str, Path]:
    """Build every fixture into ``dest_dir``; return name -> path."""

    built: dict[str, Path] = {}
    specs = {
        "CleanProject.xlsm": _build_clean_project,
        "CompileErrorProject.xlsm": _build_compile_error_project,
    }
    for name, builder in specs.items():
        path = dest_dir / name
        builder(app, path)
        built[name] = path
    return built


def _verify(app: object, path: Path) -> dict[str, object]:
    """Reopen a built fixture and return a component inventory summary."""

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
        return {
            "path": str(path),
            "protection": int(proj.Protection),
            "components": components,
            "has_modTest": "modTest" in names,
            "has_clsTest": "clsTest" in names,
            "has_frmTest": "frmTest" in names,
            "has_modBad": "modBad" in names,
            "has_thisworkbook": "ThisWorkbook" in names,
            "has_sheet1": "Sheet1" in names,
        }
    finally:
        wb.Close(SaveChanges=False)


def _run(dest_dir: Path, verify: bool) -> int:
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    exit_code = 0
    try:
        tmp_dir = Path(tempfile.mkdtemp(prefix="build_fixtures_"))
        try:
            built = _build_all(app, tmp_dir)
            dest_dir.mkdir(parents=True, exist_ok=True)
            for name, src_path in built.items():
                dest_path = dest_dir / name
                shutil.copy2(src_path, dest_path)
                size = dest_path.stat().st_size
                print(f"[built] {name} -> {dest_path} ({size} bytes)")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        if verify:
            for name in ("CleanProject.xlsm", "CompileErrorProject.xlsm"):
                info = _verify(app, dest_dir / name)
                print(f"[verify] {name}: {info}")
                expected = (
                    {"has_modTest", "has_clsTest", "has_frmTest", "has_thisworkbook", "has_sheet1"}
                    if name == "CleanProject.xlsm"
                    else {"has_modBad", "has_thisworkbook", "has_sheet1"}
                )
                missing = [key for key in expected if not info.get(key)]
                if missing:
                    print(f"[verify] FAILED {name}: missing {missing}")
                    exit_code = 1
    finally:
        with contextlib.suppress(Exception):
            app.Quit()
        del app
        pythoncom.CoUninitialize()

    print("[done] fixtures written to", dest_dir)
    return exit_code


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--verify",
        action="store_true",
        help="reopen each built fixture and assert its component inventory",
    )
    args = parser.parse_args(argv)
    return _run(FIXTURES_DIR, args.verify)


if __name__ == "__main__":
    sys.exit(main())
