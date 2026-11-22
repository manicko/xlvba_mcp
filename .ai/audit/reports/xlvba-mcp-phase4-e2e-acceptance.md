# XLVBA MCP — Phase 4 Real Edit/Macro/Edit/Macro E2E Acceptance

**Task:** `FE-13` / `FULL-EDIT-MODE-ORCHESTRATION-001` (Phase 4 independent Validator acceptance)
**Date:** 2026-11-22
**Validator role:** Independent functional verification through the real MCP STDIO boundary.
**Production code modified:** NONE. No production defect found; no fix required.
**Test files modified:** NONE (edit tool is globally denied in this environment; a direct `table_get`
post-restore assertion was considered optional — see step 16 note — but coverage is already genuine
and transitively complete, so no test change was required).

---

## 1. Environment

| Item | Value |
|---|---|
| OS | Microsoft Windows 10 Корпоративная LTSC (Windows 10 Enterprise LTSC), NT 10.0.17763.0 |
| Python | 3.14.7 (`.venv\Scripts\python.exe`) |
| Excel | 16.0 (COM `Application.Version=16.0`, `Build=14332.0`) |
| `mcp` package | 2.3.0 |
| `excel-vba-mcp` | 0.1.0 |
| Access mode | `EXCEL_VBA_MCP_ACCESS_MODE` **unset** → default `full` |
| VBOM trust | Assumed available; verified indirectly — `macro_classify_background`, `vba_list_components` and `macro_run` (all VBIDE/VBProject-backed) operated successfully, which requires "Trust access to the VBA project object model". |
| Repo / branch | `C:\py_exp\xlvba_mcp` @ `master`, HEAD `db91dbd` |
| Serial execution | `-p no:xdist` (no parallel Excel/COM) |

---

## 2. Test command and result

```
.venv\Scripts\python.exe -m pytest tests/integration_excel/test_mcp_full_access_workflow.py tests/integration_excel/test_mcp_error_surfacing_matrix.py -o addopts="" -q -p no:xdist -s
```

**Result: 5 passed in 33.59 s** (verbose rerun: 5 passed in 30.83 s). **0 failed, 0 skipped.**

| Test | Result |
|---|---|
| `test_mcp_full_access_workflow.py::test_full_access_workflow_e2e` | PASSED |
| `test_mcp_full_access_workflow.py::test_persistence_roundtrip` | PASSED |
| `test_mcp_full_access_workflow.py::test_read_only_session_denies_mutation` | PASSED |
| `test_mcp_full_access_workflow.py::test_restricted_mode_denies_execute_open` | PASSED |
| `test_mcp_error_surfacing_matrix.py::test_error_surfacing_matrix` | PASSED |

All five drive the real server (`python -m excel_vba_mcp` over STDIO via `stdio_client` +
`ClientSession`); none are skipped (Excel + VBOM available).

---

## 3. Originals-unchanged confirmation

`tests/fixtures/FullAccessWorkflow.xlsm` SHA-256:

```
before : 6B905837B8DEEB549505F9A8A51C4F55507D377E4160BB5D9D6363A19AC1D4D9
after  : 6B905837B8DEEB549505F9A8A51C4F55507D377E4160BB5D9D6363A19AC1D4D9
```

* Identical to `git show HEAD:tests/fixtures/FullAccessWorkflow.xlsm` (committed blob).
* `git status --porcelain` on the fixture is clean after both runs.
* Each e2e test copies the fixture into a per-test temp root and points all server roots
  (`EXCEL_VBA_MCP_ROOTS/WORK_DIR/SNAPSHOT_DIR/ARTIFACT_DIR/RUN_DIR/AUDIT_LOG`) at that root;
  the tests assert both the repo fixture and the temp copy stay byte-identical (PO-6).
* An independent throwaway probe (`.ai/tmp`, deleted after use) re-confirmed
  `fixture_unchanged = true` across a full open→edit→macro→edit→macro→save→restore cycle.

---

## 4. Real-MCP-boundary statement

Every result below was driven through the **real MCP STDIO boundary**:
`StdioServerParameters(command=sys.executable, args=["-m","excel_vba_mcp"])` →
`stdio_client` → `mcp.ClientSession` → `session.call_tool(...)`. No result relies on direct
in-process service calls. Both the committed e2e harness and the independent Validator probe use
this same boundary.

---

## 5. Phase-4 steps 1–17 → observed evidence

Key observed values (committed-test assertions that passed **and** independently reproduced by the
Validator probe): default `excel_open` → `read_only=false, mode="execute"`; settings columns
`[SettingKey, SettingValue]`; cycle 1 `Sheet1!B2=10` → `Result!A1=20`; cycle 2 `Sheet1!B2=20` →
`Result!A1=40` (≠20); step-14 `range_write_values` `Result!A3` verified, read-back `"e2e-probe"`;
post-restore `mode="execute", read_only=false`, `Result!A1=40`, `Sheet1!B2=20`.

| Step | Requirement | Observed evidence (value/code) | Result |
|---|---|---|---|
| 1 | Start server, ACCESS_MODE unset | `_running_server(access_mode=None)`; `server_env` omits `EXCEL_VBA_MCP_ACCESS_MODE`; server log "Starting Excel VBA MCP server (STDIO)"; `excel_health` status `ok`, `worker_state running` | PASS |
| 2 | `excel_open` default → writable execute | `excel_open{path}` → `{read_only:false, mode:"execute"}`; `excel_get_status` → `{mode:"execute", read_only:false}` | PASS |
| 3 | Create initial snapshot before mutation | `snapshot_create` **before** `table_update_rows` → `snapshot_id` returned (PO-3) | PASS |
| 4 | Read the settings table | `table_read_rows tblSettings` → `columns=[SettingKey, SettingValue]`, factor row=5, label row=alpha; initial `Result!A1=0` | PASS |
| 5 | Change first setting via table mutation | `table_update_rows tblSettings {index:1, SettingValue:10}` | PASS |
| 6 | Read it back | `range_read_values Sheet1!B2` → **10** (cycle 1) | PASS |
| 7 | Classify deterministic macro | `macro_classify_background modWorkflow.Recalculate true` → `{background_safe:true}` | PASS |
| 8 | Run the macro | `macro_run modWorkflow.Recalculate` → `status:"success"` | PASS |
| 9 | Output reflects setting 1 | `range_read_values Result!A1` → **20** | PASS |
| 10 | Change the setting again | `table_update_rows tblSettings {index:1, SettingValue:20}` | PASS |
| 11 | Read it back | `range_read_values Sheet1!B2` → **20** (cycle 2) | PASS |
| 12 | Run the macro again (same session) | second `macro_run modWorkflow.Recalculate` on the **same** `workbook_id` → `status:"success"` | PASS |
| 13 | Output reflects setting 2 (≠ step 9) | `range_read_values Result!A1` → **40**, asserted `!= 20` (live table read, not cached) | PASS |
| 14 | ≥1 additional non-VBA mutation, read back | `range_write_values Result!A3 "e2e-probe"` → `verified:true`; `range_read_values Result!A3` → `"e2e-probe"` | PASS |
| 15 | Save + verify persistence (working-copy/snapshot lifecycle) | `excel_save` → ok; `test_persistence_roundtrip`: snapshot → `excel_close(save_changes=True)` → `snapshot_restore` | PASS |
| 16 | Restore snapshot; verify settings/output/table/VBA/names | post-restore `excel_get_status` → `{mode:"execute", read_only:false}` (FE-16); `Result!A1=40`; `Sheet1!B2=20`; `vba_list_components` contains `modWorkflow`; `macro_list` contains `modWorkflow.Recalculate` & `modWorkflow.ShowMessage`; `names_list` contains `SettingsKey`; **table structure** verified (see note) | PASS |
| 17 | Original repo/source fixtures byte-identical | fixture SHA-256 unchanged before/after (repo + temp copy asserted in-test; independent probe confirms) | PASS |

**Step 16 table-structure note (evidence-based):** the `tblSettings` ListObject is verified to
survive the restore **transitively but reliably**: the post-restore `macro_run Recalculate`
succeeded, and `Recalculate` executes `Set tbl = ws.ListObjects("tblSettings")` then reads
`tbl.DataBodyRange.Cells(1,2).Value` — if the ListObject or its DataBodyRange had not survived, that
call would raise and the macro would fail. The direct cell read-back `Sheet1!B2 == 20` (B2 is inside
the table body) corroborates. A direct `table_get tblSettings` assertion would make the table
structure explicit rather than transitive; the `edit` tool is globally denied in this environment so
it could not be added, and coverage is already genuine, so this is recorded as a minor advisory only
(see §8), not a FAIL.

---

## 6. "Also verify" items

| Item | Observed evidence (code/value) | Result |
|---|---|---|
| Unclassified macro → `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_run Recalculate` **before** classification → `{status:"blocked", error_code:"BACKGROUND_SAFETY_NOT_VERIFIED"}` | PASS |
| Interactive macro → `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_classify_background modWorkflow.ShowMessage true` → `error_code:"INTERACTIVE_UI_NOT_SUPPORTED"` (R13 veto; no dialog observed) | PASS |
| Safe classified macro runs successfully | `macro_classify_background Recalculate true` → `background_safe:true`; cycle 1 / cycle 2 / post-restore `macro_run` all `status:"success"` | PASS |
| Explicit read-only session → `WORKBOOK_READ_ONLY` | read-only `excel_open` then `range_write_values` / `table_update_rows` / `macro_run` → each `WORKBOOK_READ_ONLY` | PASS |
| Restricted mode preserved → `APPROVAL_DENIED` | restricted-mode server (`access_mode="restricted"`) default `excel_open` → `APPROVAL_DENIED` | PASS |
| Full mode needs no grants / host allowlist | default `excel_open` succeeds writable/execute with **no grant tool** present or used; no per-op approval required in full mode | PASS |

Supplementary structured-error coverage (from `test_error_surfacing_matrix`): every representative
non-read-only tool returns a structured `{code, context}` error for both a real denial
(`WORKBOOK_READ_ONLY` / `SNAPSHOT_FAILED`) and a malformed argument (`INVALID_RANGE`), with **no**
unstructured `UnexpectedToolError` leak.

---

## 7. Verdict

All 17 Phase-4 steps are exercised by committed, passing, real-MCP integration tests; the six
"Also verify" items are all confirmed with concrete error codes/values; the source fixture is
byte-identical before and after. No production defect was found, so no production change was made.

**PHASE 4: PASS**

(17/17 steps PASS; 6/6 "Also verify" PASS; 5/5 tests pass, 0 skip; fixture SHA-256 unchanged.
One minor advisory: step 16 table-structure is verified transitively via the post-restore macro read
rather than by a direct `table_get` assertion — non-blocking.)