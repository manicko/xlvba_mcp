# Agent instructions (VBA / Excel) — English

## Context: "cannot compile" when Excel is installed
- **Excel installed does not mean** an external process can compile/check VBA. Common blockers: **Trust Center** must allow **Trust access to the VBA project object model (VBOM)**; macro/security policies (Protected View, blocked macros from internet, disabled content); Office edition/policies differ by machine; **no stable, universal COM API** to run "Compile VBAProject" identically across Office builds/locales (scripts may try VBE CommandBars/menu/reflection and still be best-effort).
- Many "I cannot compile" messages come from: missing VBOM trust, running without an interactive Excel session that permits the action, or confusing **VBA syntax compile** (L2) with **runtime/functional** checks (L3). Distinguish these in every report.

## Repository facts (this project)
- VBA export sources (edit these; source of truth for VBA code): `mko_treport/src/vbs/export/*.bas`, `*.cls`
- Tracked working workbook (live settings/templates + VBA applied by build): `mko_treport/src/xls/treport.xlsm` (single workbook, no separate canonical copy)
- Derived/runtime data copy (not tracked as build source/sync source): `data/theiss/treport.xlsm` — can be replaced explicitly from the tracked workbook via script deploy option
- Build script: `mko_treport/src/scripts/build_treport_vba.ps1` (see `docs/02-build-and-deployment/vba-build.md`)
- Build reports/backups: `.ai/tmp/build_treport_vba_*.txt`, `.ai/tmp/treport_backup_*.xlsm`

## Required workflow when you change VBA exports
1. **Backup / safe verification.** Never edit/verify workbooks blindly.
   - **Default build (applies current VBA to the tracked working workbook):**
     ```powershell
     .\mko_treport\src\scripts\build_treport_vba.ps1
     ```
     Builds `mko_treport/src/vbs/export` into `mko_treport/src/xls/treport.xlsm` in place, creates a timestamped backup and report in `.ai/tmp`. By default it does **not** update `mko_treport/src/vbs/export/*` or `mko_treport/src/vbs/vbaProject.bin` and does **not** modify `data\theiss\treport.xlsm`.
   - **Verification without modifying the tracked workbook (preferred when you only need to check/build):** copy the tracked workbook to `.ai/tmp` and build that copy:
     ```powershell
     Copy-Item -LiteralPath "mko_treport\src\xls\treport.xlsm" -Destination ".ai\tmp\treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm" -Force
     .\mko_treport\src\scripts\build_treport_vba.ps1 `
       -TargetWorkbook ".ai\tmp\treport_vba_check_<YYYYMMDD_HHMMSS>.xlsm" `
       -BackupDir ".ai\tmp" `
       -ReportPath ".ai\tmp\build_treport_vba_<YYYYMMDD_HHMMSS>.txt"
     ```
   Always record the exact command and report path.

2. **L1 — automated build/import (script):** review the generated report for import/legacy/duplicate issues and the **Compile/validation** note (best-effort). If it reports missing VBOM access, state clearly: enable **Trust access to the VBA project object model** in Excel Trust Center (Macro Settings), close Excel, retry. Note corporate policies may block this.

3. **L2 — manual compile in Excel (required for "compiles" claim):** open the **built/checked** `.xlsm` in Excel → VBA Editor → **Debug → Compile VBAProject**.
   - If clean: record file path, date/time, and "L2 passed (manual VBE Compile)".
   - If errors: fix in `mko_treport/src/vbs/export/`, rebuild (on a new working copy or tracked workbook as appropriate), repeat until resolved (or report blockers precisely).

4. **L3 — functional checks (when logic/run behavior changes):** test `Refresh_data` / `Create_report` on **copies** of data/workbooks only (do not corrupt originals). Verify expected outputs and that source files remain unchanged per project rules.

5. **Explicit post-verification actions (opt-in only, never implicit):**
   - **Sync exports/bin (`-SyncExports`)**: only after successful L2 (and L3 if applicable), and only when you intentionally want to refresh tracked `mko_treport/src/vbs/export/*` and/or `mko_treport/src/vbs/vbaProject.bin` from a single, verified built workbook. Run with `-SyncExports` against that unambiguous target. Note: `VBComponent.Export` output is not post-normalized (format/EOL/encoding/case may vary by VBE environment) and `vbaProject.bin` is a volatile binary artifact — review diffs, avoid committing machine churn.
   - **Deploy data copy (`-DeployToDataTheiss`)**: only when explicitly needed; replaces `data\theiss\treport.xlsm` with a copy of the tracked `mko_treport/src/xls/treport.xlsm` (derived/runtime copy, never a sync source). Use after a successful, verified build/save of the tracked workbook.

## What to report after VBA edits (always)
- List changed files under `mko_treport/src/vbs/export/` (and note if exports/bin/workbooks were intentionally updated).
- Exact `build_treport_vba.ps1` command used (or why not run).
- Path to build report and short summary (SUCCESS/ISSUES/ERROR; note VBOM issues if any).
- L2 result: done (with file/time) / not done (state this is a limitation, not proof of compilation).
- L3 result if performed, or "not performed".

## Do not
- Claim "compiled successfully" based only on the script's best-effort `VBProject.Compile` call without L2 when a user asks for verification.
- Run or request running macros on production/original data without working copies and explicit need.
- Run `-SyncExports` or `-DeployToDataTheiss` implicitly; these are explicit opt-in actions.

## Reference
- Build/deploy behavior and options: `docs/02-build-and-deployment/vba-build.md`
- Full analysis report (created 2026-11-15): `.ai/reports/vba-compile-and-verification-report-2026-11-15.md`
