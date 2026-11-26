# Stage 1 — Live Integration Evidence (excel_save_as) + Candidate Real-Agent Run

- **Candidate revision:** `ef3e6cb` (working tree; Save As fix applied post-test, see §3).
- **Baseline revision:** `40c8b4d` (pre-plan_05).
- **Live env:** Excel 16.0, owned PID 27704, access_mode full, single-STA worker, 0 open workbooks at start.
- **Test dir:** `C:\py_exp\xlvba_mcp\.ai\tmp\saveas_live` (isolated; no production data).
- **Source fixture:** `source.xlsx` (minimal valid xlsx, SHA256 `17A576DB…`).

## 1. Scenario matrix (live MCP tool path; files independently hashed)

| # | Scenario | Input | Expected | Actual | Verdict |
|---|---|---|---|---|---|
| 1 | New target | save_as out_new.xlsx (new) | ok, file created | ok; out_new `45F024AF…`; source unchanged `17A576DB…` | PASS |
| 8 | Reopen artifact | open out_new, read A1 | valid, A1=edited | opened clean (no repair); A1=`edited-via-mcp` | PASS |
| 2 | Existing target, overwrite off | save_as out_new (exists) | reject, no change | `PROTECTED_TARGET/target_exists`; out_new unchanged | PASS |
| 3 | Existing target, overwrite on | save_as out_new overwrite_existing=true | ok + .bak=pre-op | ok; new `FA6ACC…`; `.bak`=`45F024AF…` (exact pre-op) | PASS |
| 4a | Same as source, no flag | save_as source.xlsx | reject | `PROTECTED_TARGET/source_protected` | PASS |
| 4b | Same as source, overwrite_source=true | save_as source.xlsx | ok, source replaced, .bak=orig | ok; source `FD35D8…`; `.bak`=`17A576DB…` (original) | PASS |
| 5 | External source mod | modify source2 after open, save_as | detect drift | `STALE_TARGET/source_modified_since_open` (expected vs actual sha) | PASS |
| 6 | Pre-existing .bak | target + pre-existing .bak, overwrite | deterministic conflict | **DEFECT**: returned `outcome="unknown"` (fixed, §3) | FIXED |
| 7 | Subsequent edit+save | save_as out_s7, edit A1, excel_save | target unchanged | out_s7 `2C8A9F…` unchanged; working copy A1=`after-save-edit` | PASS |
| 9 | Failure/interruption | temp/backup/replace/verify | no false success; unknown if unprovable | covered by 15 unit cases (incl interruption→unknown) + code review; live injection unsafe | COVERED (unit) |
| 10 | Repeated op | save twice | conflict handling | covered by #3/#6 | PASS |

Structured error envelope (U6-B) observed live on every rejection (`errorCode/errorCategory/isRetryable/suggestedNextActions/details`). Snapshot precondition enforced live (`SNAPSHOT_FAILED` → `snapshot_create`).

## 2. Candidate real-agent run (model-in-the-loop, candidate only)

A general LLM subagent autonomously executed 3 tasks via the live MCP server (no human intervention, no source reading):
- T1 environment discovery: `xlvba_xlvba_environment` → access_mode=full, non_activating_compile=false, routing=`range_read_values`. Correct.
- T2 read: `excel_open`(inspect,writable) → `range_read_values`(A1="source") → `excel_close`(no save). Correct; no snapshot needed for read.
- T3 save as: `excel_open`(execute) → `snapshot_create` → `range_write_values`(typed CellValue "updated_by_agent") → `excel_save_as`(new path, overwrite_existing=false) → outcome ok, sha `EA7425F9…` → `excel_close`. Correct typed args, correct tool selection, no errors.

This is a genuine LLM-agent-through-MCP run (candidate), not a deterministic proxy.

## 3. Stage 1 defect + fix

- **Defect (scenario 6):** pre-existing `<target>.bak` returned `outcome="unknown"` (reserved for interrupted/indeterminate) instead of a deterministic conflict.
- **Fix:** `save_as_workbook` now rejects a pre-existing `.bak` in the decision phase (pre-mutation) with `PROTECTED_TARGET` / `reason=backup_exists` (isError). Late `unknown` return removed; genuine interruption→unknown path kept. Regression tests updated (cases 12,15). Unit `test_save_as.py` 15 passed; full unit green; ruff clean.
- **Live re-test of the fix:** requires MCP server restart (running server holds pre-fix code). Deferred to Stage 2 post-restart health check.

## 4. Comparative baseline benchmark — BLOCKED

The required baseline-vs-candidate model-in-the-loop comparison (≥3 trials/workflow, same model/Kilo/MCP) is **not executable in this environment**:
- Only one MCP server (candidate) is connected; the Kilo global config (`~/.config/kilo/kilo.jsonc`) is read-restricted and cannot be re-pointed to a baseline server.
- No second server process is reachable by a subagent; the baseline revision `40c8b4d` is identifiable but not runnable as a routed variant here.
- Per assignment rules, proxy results are NOT substituted for the real comparison.

Candidate-only real-agent evidence is provided as supplementary (not a comparison).