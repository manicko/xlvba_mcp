# task_002 — Planner output (Tech Lead session, 2026-11-19)

_Source: `task` subagent_type=planner (bounded PoC execution plan)._

## Prerequisites / context
- Windows 11 x64 target, interactive desktop required, no Service/RDP-without-session assumptions.
- Excel installed (bitness aligned with Python per project target); no Trust Center/registry changes.
- Serial execution; no `xdist` for Office UI/COM PoC. Do not touch user Excel beyond read-only observation required to prove isolation.

## Sequencing (exact)
1. Manual pre-run checklist.
2. Env/deps check (`uv sync --locked`, imports).
3. Baseline capture before any `DispatchEx`: env log fields + user-session isolation baseline (PIDs, visible Excel windows, ROT/GetActiveObject read-only probe as specified).
4. Run serially: `test_dispatch_ex_creates_isolated_excel_instance`, then `test_excel_workbook_open_close_lifecycle`.
5. Post-run: re-collect isolation baseline; confirm only baseline user state remains and no orphaned owned PID.
6. On failure/hang: capture HRESULT/exception/COM details, modal/screenshot rules (screenshot only if dialog appears; no SendKeys/click automation), terminate only owned PID if last resort and document.
7. Finalize evidence artifact and explicit Go/No-Go for R01/task_002.

## Minimal fixtures / isolation
- No pre-existing repo fixtures required; create runtime temp workbook in temp dir for lifecycle test (prefer `.xlsx`, save/open/close sequence as planned). No files opened from `projects/`/`work/`/`snapshots/`.
- Baseline: `EXCEL.EXE` PID list + window titles before/after; optional read-only ROT snapshot before/after (never mutate). Owned PID must be new (not in baseline). Post-quit: owned PID gone, user baseline unchanged (or documented external change).

## Harness location and skeleton (outline)
- `tests/integration_excel/test_excel_instance_poc.py`
  - Windows/Excel/pytest guards; markers as appropriate.
  - Helpers: logging to `.ai/tmp/excel_instance_poc_<ts>.log` (+ optional console capture), env collection (Windows build/Office version/build/bitness/VBE language best-effort), isolation snapshots, COM error formatting, STA init/uninit, `DispatchEx` owned Excel, temp workbook lifecycle, safe quit with bounded wait.
  - Fixtures: logger, STA COM context, owned Excel with baseline/teardown (function scope, serial).
  - Tests: `test_dispatch_ex_creates_isolated_excel_instance`, `test_excel_workbook_open_close_lifecycle` as named in plan.

## Pass/fail mapping (from planner)
Criteria mapped to ACs: separate instance + identity; user session unaffected; no orphan after open/close/Quit; COM proxies on creating STA thread (documented design/evidence); message pump behavior and limitations documented (including RPC/busy/modal observations or “none observed”). Evidence template fields listed (env, isolation baselines, test results, COM observations, go_no_go_r01_task002).

## Manual run checklist
(before/while/after/hang procedures as written in planner output; emphasizes interactive desktop, no blind dismissal of dialogs, own-PID-only termination if required.)
