# xlvba Usage Audit — Agent Reports (Auditor + Researcher)

- Date: 2026-11-25
- Scope: xlvba_* tool usage by the two agents launched in the parent session (Auditor, Researcher) during the `report_w` "second sheet copies only the top row" investigation.
- Basis: each agent reported ONLY the xlvba_* tools it actually called in its own session. No calls are invented. Non-xlvba actions (PowerShell file copies, OS `Stop-Process` to kill a hung worker) are mentioned only as recovery context and are explicitly flagged as NOT xlvba calls.
- Source agent reports (temp): `.ai/tmp/xlvba_usage_auditor.md`, `.ai/tmp/xlvba_usage_researcher.md`.

---

## 1. Did the agents use xlvba tools?

| Agent | Used xlvba? |
|-------|-------------|
| Auditor | **YES** |
| Researcher | **YES** |

Both agents used xlvba tools. Neither agent executed macros or modified originals; both worked on staged copies in `inspect` mode.

---

## 2. Tools and functions used (per agent)

### 2.1 Auditor — tools called

`xlvba_server_health`, `xlvba_excel_open`, `xlvba_sheets_list`, `xlvba_tables_list`, `xlvba_table_read_rows`, `xlvba_range_read_values`, `xlvba_pivot_tables_list`, `xlvba_pivot_table_get`, `xlvba_excel_health`, `xlvba_excel_list_workbooks`, `xlvba_excel_close`.

| Function | Purpose | Result |
|----------|---------|--------|
| `xlvba_server_health` | Liveness pre-flight | WORKED (`status:ok`, v0.1.0) |
| `xlvba_excel_open` | Open copies in `mode=inspect` | WORKED on allowed paths; **FAILED** on a disallowed path (`PATH_NOT_ALLOWED`) |
| `xlvba_sheets_list` | Enumerate sheets, used range, counts | WORKED reliably (found `dashboard_data` = `$A$1:$B$566386`, ~566k rows) |
| `xlvba_tables_list` | List ListObjects + columns | WORKED (`general`, `reports`, `placeholders`, `data_files`) |
| `xlvba_table_read_rows` | Read settings rows (typed) | WORKED (exact filters/params for both `report_w` rows) |
| `xlvba_range_read_values` | Bounded range reads (output + source page-field block) | WORKED reliably, incl. 566k-row workbook (no pivot COM) |
| `xlvba_pivot_tables_list` | Enumerate all PivotTables | **FAILED** — `MCP error -32001: Request timed out`; hung the worker |
| `xlvba_pivot_table_get` | Read `PivotTable1` metadata | **FAILED** — first `EXCEL_HUNG`, then `-32001` timeout on a fresh worker |
| `xlvba_excel_health` | Diagnose worker state | WORKED (read-only diagnostic; `needs_recovery`, `owned_pid`, `active_operation`, `notes`) |
| `xlvba_excel_list_workbooks` | List open workbooks (bookkeeping) | WORKED (no COM); can be **stale** after worker death |
| `xlvba_excel_close` | Close copies without saving | MIXED — light copies WORKED; heavy `report_w` copy **timed out** while worker hung |

Not called (deliberately): `mode=execute` opens, `xlvba_macro_run`, `xlvba_vba_compile`/`vba_*`, `xlvba_workbook_get_model`, `xlvba_pivot_fields_list`, `xlvba_pivot_cache_get`, `xlvba_names_list`. VBOM/Trust Center was therefore not exercised.

### 2.2 Researcher — tools called

`xlvba_server_health`, `xlvba_excel_health`, `xlvba_excel_open`, `xlvba_tables_list`, `xlvba_table_read_rows`, `xlvba_range_read_values`, `xlvba_excel_close` (7 distinct tools).

| Function | Purpose | Result |
|----------|---------|--------|
| `xlvba_server_health` | Liveness pre-flight | WORKED (`status:ok`, v0.1.0) |
| `xlvba_excel_health` | Confirm worker accepting before work | WORKED (`accepting:true`, `needs_recovery:false`); `notes` carried noisy historical incidents |
| `xlvba_excel_open` (`inspect`) | Open 4 staged copies (incl. 129 MB `report_wclip`) | WORKED for all 4 |
| `xlvba_tables_list` | Enumerate settings ListObjects + columns | WORKED (all-optional columns reported `data_type:"empty"`) |
| `xlvba_table_read_rows` | Read `data_files`/`reports`/`general`/`placeholders` | WORKED (typed cells with addresses) |
| `xlvba_range_read_values` | Read page-field block `clean_data_full!A1:B12` on 3 sources | WORKED incl. 129 MB file (deliberate safe alternative to pivot COM) |
| `xlvba_excel_close` | Close copies (no save) | WORKED; first close returned transient `busy:true` |

Not called: `macro_run`, `vba_*`, `pivot_*`, `sheet_*`, `range_write_*`, `snapshot_*`, `test_*`. No macros run.

---

## 3. Detailed problem reports (per function)

### P1 — `xlvba_excel_open`: path allowlist rejection (Auditor)
- Symptom: opening a copy under the project temp dir was rejected.
- Error: `{"code":"PATH_NOT_ALLOWED","message":"Path not allowed (outside_roots)...","context":{"reason":"outside_roots","operation":"mutating","allowed_roots":["C:\\py_exp\\xlvba_mcp","C:\\Users\\MatveichevN\\AppData\\Local\\Temp","C:\\Users\\MatveichevN\\AppData\\Local\\Temp\\excel_vba_mcp\\artifacts"]}}`.
- Impact: had to re-stage all working copies under an allowed artifacts root. Originals never touched.
- Note: Researcher avoided this by copying originals into the allowed root first.

### P2 — `xlvba_pivot_tables_list`: worker hang / timeout on large workbook (Auditor)
- Symptom: full-workbook PivotTable enumeration on `report_w.xlsx` (with ~566k-row `dashboard_data`) timed out and hung the single COM worker.
- Error: `MCP error -32001: Request timed out`. `xlvba_excel_health` then showed `needs_recovery`, `active_operation:"pivot_tables_list"`, `owned_pid:24020`, notes `["submit timeout section=diagnostic token=9"]`.
- Impact: the single worker was blocked; all subsequent COM calls failed until recovery. Root cause: very large data sheet makes full pivot enumeration too slow.

### P3 — `xlvba_pivot_table_get`: timeout / EXCEL_HUNG (Auditor)
- Symptom: reading `PivotTable1` metadata on `clean_data_full` failed twice.
- Errors: first `{"code":"EXCEL_HUNG","message":"Excel COM worker needs recovery; refusing new work"...}` (worker already hung); after killing the worker and opening a fresh copy, a targeted `pivot_table_get PivotTable1` again failed with `MCP error -32001: Request timed out`.
- Impact: could not read filtered `TableRange1`/`DataBodyRange` directly; root cause had to be established from the produced output workbook + source page-field block (via `range_read_values`) + static analysis.

### P4 — `xlvba_excel_close`: timeout while worker hung (Auditor)
- Symptom: closing the heavy `report_w` copy while the worker was hung timed out (`-32001`, queued behind the stuck pivot call). Closing light copies worked.
- Impact: could not cleanly close during a hang; recovery required an out-of-band worker kill.

### P5 — `xlvba_excel_health`: noisy historical `notes` (Researcher)
- Symptom: `notes` array carried a rolling history of prior incidents not caused by the session (`submit timeout ... token=9/13`, `active call returned after NEEDS_RECOVERY`, `drain_close complete`, `reset rebuilt owned instance`), present from the first call.
- Impact: alarming/noisy; current vs historical state not clearly separated. Worked around by reading `session_state`/`worker_state`/`accepting` instead of `notes`.

### P6 — `xlvba_excel_close`: transient `busy:true` (Researcher)
- Symptom: first close after a large workbook returned `busy:true`, `active_operation:"close_workbook"`; later closes `busy:false`.
- Impact: minor friction; no failure.

### P7 — `xlvba_excel_list_workbooks`: stale bookkeeping (Auditor)
- Symptom: listed workbooks whose worker had already been killed.
- Impact: misleading during recovery; it is bookkeeping-only (no COM) and not proof a workbook is actually open.

### Cross-cutting — single-worker head-of-line blocking + manual recovery
- A hung COM call blocks ALL subsequent COM calls (open/close/pivot) until recovery. There is no xlvba "restart worker" tool; recovery was manual via OS `Stop-Process` on `owned_pid` (NOT an xlvba call), after which a fresh `xlvba_excel_open` respawned a worker (`owned_pid` changed 24020 → 26548). `xlvba_excel_close` cannot succeed while the worker is hung.

---

## 4. What is convenient

- `inspect` mode + hidden `work\` working copy: originals are never modified — safe for read-only auditing.
- `xlvba_sheets_list`, `xlvba_tables_list`, `xlvba_table_read_rows`, `xlvba_range_read_values` are fast and reliable, even on a ~566k-row / 129 MB workbook, because they avoid pivot COM.
- Typed reads (`kind`/`value`/`address`) and rich table/column metadata make settings and output verification precise (dates ISO, booleans real, blanks explicit `kind:"empty"`).
- `tables_list` reports per-column `data_type` (`empty` for all-blank optional columns) — useful to detect optional config columns.
- Automatic staging protects originals; `excel_close` defaults to `save_changes:false`.
- `xlvba_server_health` / `xlvba_excel_health` (`accepting`, `needs_recovery`, `open_workbooks`, `owned_pid`) give a clear liveness/diagnostic surface.
- Small bounded range reads (e.g. `A1:B12`) are cheap even on huge files — a safe way to read pivot page-field state without pivot COM.

---

## 5. What is NOT convenient / painful

- Path allowlist is restrictive: `xlvba_excel_open` rejects the project tree (`.ai/tmp`) with `PATH_NOT_ALLOWED (outside_roots)`. Copies must be staged under an allowed root, then evidence copied out.
- Pivot enumeration/metadata on large workbooks hangs the single worker (`pivot_tables_list`, `pivot_table_get` → `-32001`; first call may return `EXCEL_HUNG`).
- A hung worker blocks all subsequent COM calls; even `excel_close` fails during a hang.
- Recovery is manual and out-of-band (OS kill of `owned_pid`); no xlvba restart call.
- `excel_list_workbooks` can show stale bookkeeping after a worker death.
- No per-call timeout / partial tolerance on heavy pivot calls — they fail hard and poison the session instead of returning a bounded "too large" result.
- Identity indirection: the working file is renamed to a GUID in `excel_vba_mcp\work\`; you must track `workbook_id` (original name only in `source_path`).
- "`inspect` is not read-only": it returns `read_only:false` (writable copy); non-persistence relies on closing without save — a footgun if you assume inspect == read-only.
- Noisy historical `notes` in `excel_health`; transient `busy:true` on first close after a large workbook.
- Copy-heavy for big files: originals staged into the allowed root, then the server copies again (double I/O for 129 MB).

---

## 6. Everything an xlvba developer should know

**Allowed roots / paths**
- Observed allowed roots: `C:\py_exp\xlvba_mcp`, `C:\Users\<user>\AppData\Local\Temp`, `C:\Users\<user>\AppData\Local\Temp\excel_vba_mcp\artifacts`. The project tree is rejected (`PATH_NOT_ALLOWED (outside_roots)`).
- Always copy originals into an allowed root with a distinct name and open that copy; keep originals untouched; copy results/evidence out afterwards.

**Modes / persistence**
- `mode:"inspect"` disables macros but opens a **writable** hidden working copy (`read_only:false`); `mode:"execute"` additionally allows macro-by-UI execution. Inspect ≠ read-only.
- Originals are never modified (server works on a staged copy under `...\excel_vba_mcp\work\<GUID>`); changes persist only via `excel_save` or `excel_close(save_changes:true)`. Default close = no save.
- `excel_open` returns `path` (server working copy, GUID-named) and `source_path` (your input). Track `workbook_id`.

**Read-only inspection pattern that works**
- `server_health` → `excel_health` (check `accepting`/`needs_recovery`) → `excel_open(inspect)` → `tables_list` / `table_read_rows` / `range_read_values` → `excel_close` (no save).

**Large-pivot hang avoidance (critical)**
- Do NOT call `pivot_tables_list`, `pivot_table_get`, or full-workbook model enumeration on huge sources (`report_w` ~566k-row `dashboard_data`; `report_wclip` 129 MB) — known EXCEL_HUNG / submit-timeout triggers.
- Instead range-read the page-field block (e.g. `clean_data_full!A1:B12`) with `range_read_values`. Small block reads are cheap even on 129 MB files.

**Hang signature & recovery**
- Hang: `MCP error -32001: Request timed out`, then `excel_health` shows `needs_recovery` with `active_operation` and `notes` like `submit timeout section=diagnostic token=N`; a follow-up COM call may return `EXCEL_HUNG ("Excel COM worker needs recovery; refusing new work")`.
- No xlvba restart tool. Recover by killing the owned Excel worker PID (`excel_health.owned_pid`) via the OS, then a fresh `excel_open` respawns a worker (`owned_pid` changes). `excel_close` will not succeed while hung.
- `excel_list_workbooks` reads bookkeeping only (no COM) and can be stale after a worker death — not proof a workbook is open.
- Verify health via `session_state`/`worker_state`/`accepting`, not the rolling `notes` log.

**Typed reads**
- Reads return `kind` (string/number/boolean/date/empty/error) + `address`; use directly, don't string-coerce. `tables_list` reports per-column `data_type` (`empty` for all-blank columns).

**Macros / compile (not exercised this session; relevant context)**
- No macros were run. `macro_run` enforces background-safety vetoes (R13 interactive-UI, BG06) and blocks `MsgBox`-ending macros — end-to-end `Create_report`/`Refresh_data` would need `MsgBox` neutralized on a copy first.
- `vba_compile` returns `indeterminate` without VBOM/Trust Center; an L2 manual VBE compile is required for a "compiles" claim. VBOM/Trust Center was not exercised here.

**Hygiene**
- Close workbooks promptly (frees the worker, avoids memory buildup). Stage copies with clear names; never open originals directly. Budget for double I/O on huge workbooks.

---

## 7. One-line summary for the xlvba maintainer

The non-COM read paths (`server_health`, `excel_health`, `excel_open(inspect)`, `sheets_list`, `tables_list`, `table_read_rows`, `range_read_values`, `excel_close`) are fast, safe, and sufficient for most auditing — including on very large workbooks. The COM-heavy pivot paths (`pivot_tables_list`, `pivot_table_get`) hang the single worker on large data and poison the session with no built-in recovery; add a per-call timeout / bounded "too large" result and an in-tool worker restart, reduce `notes` noise, clarify that `inspect` is writable, and document the path allowlist and GUID working-copy indirection.