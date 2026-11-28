# Independent Auditor 3 — Safety, Operations, Limitations, Large-Data, Competitive Positioning

- **Role:** Independent Auditor 3 (READ-ONLY). No code or docs modified.
- **HEAD:** `f79faf2b8797a897429754d71e89f8fb9d031317` (branch `master`), verified clean except this report's output directory (`.ai/audit/docs/`, untracked).
- **Date:** 2026-11-28.
- **Scope:** docs 02/03/08/09/10 + README + ru mirrors; source under `src/excel_vba_mcp/`; existing evidence in `.ai/large-data/*`, `.ai/audit/reports/*`, `.ai/tmp/xlvba-*`.
- **Method:** source reading, `mcp.list_tools()` runtime enumeration, `pytest` unit run, and independent web verification of the two competitor repositories. Every assertion is graded below.

## Verification grade legend
1 = Source (read in code) · 2 = Unit (pytest) · 3 = Integration (service/dispatch) · 4 = Real-workbook (live Excel on a real file) · 5 = MCP-runtime (tools/list or live tool call) · 6 = Unverified.

---

## Baseline verification (independent)

| Claim | Result | Grade |
|---|---|---|
| HEAD `f79faf2`, clean | Confirmed `git rev-parse HEAD` = `f79faf2…`; `git status` clean except `.ai/audit/docs/` | 1 |
| 43 MCP tools | `await mcp.list_tools()` returned **43** names (control, data_export, defined_name, defined_name_delete, defined_name_read, excel_health, excel_recover, job_get, job_request_cancel, macro, macro_run, pivot_field, pivot_read, pivot_refresh, range_delete, range_merge, range_read, range_structural, range_write, server_health, shape, sheet, sheet_delete, sheet_read, snapshot_create, snapshot_diff, snapshot_list, snapshot_restore, table, table_delete_column, table_delete_rows, table_read, test_suite, vba_compile_diagnostics, vba_export_project, vba_insert_code, vba_read, vba_replace_component_source, workbook, workbook_read, workbook_save_copy, workbook_snapshot_read, xlvba_environment) | 5 |
| 4 read-only resources | `workbook_model`, `workbook_vba_components`, `workbook_vba_component_source`, `test_run_summary` registered via `@mcp.resource` | 1/5 |
| Python `>=3.11,<3.15` | `pyproject.toml` `requires-python`; `.python-version` pins 3.14 | 1 |
| Unit suite green | `pytest tests/unit -m "not integration_excel"` → **1615 passed, 0 failed, 0 errors, 0 skipped** (exit 0); full collection = 2012 tests (≈397 live-Excel integration) | 2 |

---

## A. Installation & environment

| Assertion | Evidence | Grade |
|---|---|---|
| Windows-only (COM + VBIDE) | `pyproject.toml` deps `pywin32>=306`, `comtypes>=1.2.0`; docs 02/09 state Windows-only | 1 |
| Validated OS: Win10 LTSC 2019 build 17763; Excel 16.0.14332 x64; Python 3.14 x64 | README requirements table; docs 02 System Requirements; `xlvba_environment` `target_build="16.0.14332"` with `build=None` (never fabricated) — `environment_tools.py:55-57,178` | 1 |
| README requirements table matches code | README Python `>=3.11,<3.15` == `pyproject.toml`; Excel build 16.0.14332 == `_TARGET_BUILD`; OS matches docs 02/09 | 1 |
| VBOM trust required, manual per-machine | `macro_policy.diagnose_vbom` → `AccessVbomDisabledError`; docs 02/08/09; server never sets registry | 1 |
| Allowed roots via `EXCEL_VBA_MCP_ROOTS` (os.pathsep; default CWD + temp) | `server.py` `_ENV_ROOTS`; `path_policy.PathPolicy`; docs 08 "Filesystem access" | 1 |
| Transport = MCP STDIO (stdout protocol, stderr logs) | `server.py` `mcp.run(transport="stdio")` via `__main__`; docs 02/08 | 1 |
| Clients: Kilo + Claude Desktop | README Step 5 + docs 02 both give Kilo `type: local` and Claude `mcpServers` configs | 1 |
| Startup/health checks | `server_health` (non-COM liveness), `excel_health` (worker/session, non-COM), `excel_list_workbooks`; docs 02 "Health Checks" | 1/5 |
| `EXCEL_VBA_MCP_LOG_LEVEL` parsed but NOT applied | `config.py` parses it; docs 09/08 explicitly state "parsed but not applied"; logging fixed at INFO at import (`server._configure_logging`) | 1 |

**Discrepancies (A):** none material. Minor: README says "43 MCP tools and 4 read-only MCP resources" — confirmed accurate.

---

## B. Workbook safety & data integrity (user-visible contract)

| Control | User meaning | Evidence | Grade |
|---|---|---|---|
| **Original never modified** | Open copies the allowlisted original into `work/` and opens only the copy | `session_service.open_workbook` → `_materialize_working_copy`; `_sha256_file(canonical)` captured pre-copy | 1 |
| **Working copies** | `WorkbookInfo.path` = working copy (`wb.FullName`); `source_path` = original | `session_service._SessionRecord`; docs 09 "Two paths per workbook" | 1 |
| **Snapshots (immutable, SHA-256)** | `SaveCopyAs` for open books, file copy for closed; manifest + hash | `snapshot_service.create_snapshot` open/closed branches; `SnapshotStore` | 1 |
| **Restore** | Opens fresh working copy from snapshot baseline, validates it, then closes previous WITHOUT saving; rebinds identity to original; never reverts external side effects | `snapshot_service.restore_snapshot` (WP3 order + R1/R3); `previous_close` enum | 1 |
| **Save vs independent copy** | `workbook action="save"` writes working copy in place; `workbook_save_copy` = atomic temp→replace to a (new/original) path, `.bak`, size+SHA verify, never `wb.SaveAs`, session stays bound | `session_service.save_workbook` vs `save_as_workbook` / `_atomic_save_as_copy` | 1 |
| **Overwrite/backup behavior** | Existing target needs `overwrite_existing`; original needs `overwrite_source` + matching source hash; single `<target>.bak`; pre-existing `.bak` rejected (`backup_exists`) | `save_as_workbook` decision matrix; `_atomic_save_as_copy` | 1 |
| **Mutation preconditions** | Writable session + successful snapshot + approval + postcondition re-read | `ranges_service._require_writable/_require_snapshot/_require_approval`; same in tables/pivot/sheets | 1 |
| **Approval / confirmation** | `full` (default) auto-allows MUTATING/EXECUTING; `restricted` deny-by-default; destructive ops need `confirm=true` | `approval.check`; `recovery_tools.excel_recover` + `snapshot_restore` `CONFIRMATION_REQUIRED` | 1 |
| **Path restrictions** | Canonicalise + realpath cross-check; block traversal/symlink/UNC/device; extension allowlist | `path_policy.validate_path` (full reason enum) | 1 |
| **Error reporting** | Every failure → structured `ToolError` JSON with stable `code`/`context`; pydantic `ValidationError`→`INVALID_RANGE` | `ExcelVbaMcpServer.call_tool`, `_validation_error_as_domain` | 1 |
| **Interrupted operations** | `interrupted_unknown` job (`outcome=unknown`, `retryable=false`), never auto-replayed; recovery reports `prior_job="unknown"` | `job_get` desc; `recover_owned_process` `prior_job` logic (session_service:1301-1307) | 1 |
| **Save refuses dirty close** | `close` refuses unsaved changes unless `save_changes`/`discard_changes` (contradictory flags rejected) | `session_service.close_workbook` | 1 |

**Assessment:** The safety contract is coherent, implemented, and matches docs 08/09. This is the strongest, most defensible part of the product. No safety discrepancies found beyond the already-documented `macro_run`-has-no-snapshot deviation (see F).

---

## C. Excel & VBA behavior

| Assertion | Evidence | Grade |
|---|---|---|
| Macro security honored, never weakened permanently | `AutomationSecurity` set only inside `automation_security(mode)` context, restored in `finally`; `inspect`→ForceDisable(3), `execute`→ByUI(2); LOW never selected | 1 |
| VBOM required; locked projects never bypassed | `diagnose_vbom`/`diagnose_protection`; `VBA_PROJECT_LOCKED` | 1 |
| **Compile diagnostics are fail-closed indeterminate — NEVER claims compiled** | `CompileService.compile_project` returns `status=COMPILE_INDETERMINATE`, `reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`, `diagnostics.is_compile=False`; `_build_result` notes "NOT a compile"; `vba_compile_diagnostics` docstring "never returns compiled or compile_error" | 1/2 |
| Macro execution limits (≤30 args; soft deadline; no auto-kill) | `macro_run` desc (≤30, `Application.Run` limit); `macro_timeout_s` default 90s; `_handle_hung`→`SUSPECTED_STALLED`, `owned_pid_terminated=False` | 1 |
| Interactive-UI (R13) veto | `INTERACTIVE_TOKENS` scan (case-insensitive), heuristic/heuristic-body → `INTERACTIVE_UI_NOT_SUPPORTED` before dispatch; documented as veto not certificate | 1/2 |
| background_safe classification (BG06) | `BackgroundSafetyRegistry` keyed by `(workbook_sha256, project_source_hash, procedure)`; unclassified→`BACKGROUND_SAFETY_NOT_VERIFIED`; source-change invalidates | 1/2 |
| Recovery after Excel failure | `excel_recover` (agent-directed, `confirm=true`, pid-scoped, fail-closed, originals-safe) + macro-hang auto-recovery (F2b) | 1/2/4 |
| `xlvba_environment` honest build | `build=None` always; `non_activating_compile=False`; `classification_is_authorization=False` | 1/5 |

**Assessment:** Docs are careful and accurate here. The compile story is described truthfully everywhere I checked (README note, docs 09, tool description, service). **No overstatement of compilation found.** This is a notable honesty strength vs. competitors that imply macro execution.

---

## D. Large data & PivotTables

| Capability | Status | Evidence | Grade |
|---|---|---|---|
| Large range paging | **Implemented + real-data-tested** | `ranges_service.read_values` offset/limit + `PageInfo`; benchmark: 200k×20 read in 105 pages, complete=true | 1/4 |
| Table paging | **Implemented** | `tables_service.read_table_rows` offset/limit + `max_table_cells` | 1/2 |
| Size limits (`LargeDataConfig`) | **Implemented** | `max_range_cells=100k`, `max_response_cells=50k`, `max_response_bytes=8MiB`, `max_table_cells=50k`, `max_pivot_result_cells=50k`, `max_pivot_members=10k`, `max_export_rows=10M`, `max_file_jobs=2`, ZIP/XML guards | 1 |
| Large workbook open | **Implemented (copy-based)** | working-copy model; no size cap on open itself | 1 |
| Sheet discovery | **Implemented** | `sheet_read list`; OOXML `read_metadata` | 1 |
| Pivot metadata | **Implemented** | `pivot_read tables-list/table-get/cache-get`; single worksheet sweep + per-sheet cost budget | 1/4 |
| **Pivot result reading (`pivot_read result-read`)** | **Implemented + real-data-tested** | `pivot_tables_adapter.read_pivot_result` reads `TableRange1` via shared bounded range reader for ALL source types; benchmark: external-source pivot 50×20 read | 1/4 |
| Member enumeration | **Partially supported** | DATABASE only, ≤10k, truncated flag; external/OLAP/Data Model fail-closed | 1 |
| External/OLAP/Data Model | **Partially supported** | displayed result read = supported (all types); field/member/filter/refresh = `unsupported_source` fail-closed | 1/4 |
| PivotCache reading (OOXML) | **Implemented** | `ooxml_service.read_pivot_cache_records`; `saveData=false`/no records → `reason="unsupported_source"`, never fabricated | 1/2 |
| OOXML streaming (`workbook_snapshot_read`) | **Implemented, non-COM** | `ooxml_service` + `asyncio.to_thread`; `zip_safety.open_safe` (member count, uncompressed, ratio, traversal, depth) | 1/2 |
| Export (`data_export`) | **Implemented + real-data-tested** | `data_export_service` one page in memory, atomic finalize, CSV/JSONL, returns path/rows/bytes/sha/completeness; benchmark peak 33.9 MB for 5000 rows | 1/4 |
| File-size / memory limits | **Implemented** | `max_file_mb=100`, `max_uncompressed_mb=1024`, `max_compression_ratio=1000`, `max_zip_members=10000` | 1 |
| Completeness/truncation | **Implemented** | `complete`/`truncated`/`has_more` on every page DTO; `STALE_CURSOR` on version drift | 1/4 |

**Not implemented / unsupported (confirmed):** CSV import, server-side filtering/aggregation, Parquet export, `.xlsb` (research path absent). Pivot **create/delete**, charts, slicers (docs 09).

**Honesty check:** `.ai/large-data/large_data_results.md` and the benchmark are careful to label MCP-transport validation of the 2 new tools as in-process (limitation 2), and mark OLAP member enumeration as fail-closed/inferred. My source read confirms those caveats are real, not overstated. The one stale item: the consolidated real-data report (2026-11-27) lists the external-pivot `fields-list` wedge as a Major open defect, but the final comparison report and current source show it **FIXED** (`_gate_supported_source_read` at `pivot_tables_adapter.py:977-995` raises `UnsupportedPivotSourceError` before any `PivotFields` handle). The wedge is resolved in HEAD.

---

## E. Product differentiation (independently verified)

Competitor facts verified live from each repo's README (2026-11-28):

- **haris-musa/excel-mcp-server** (4.3k★, MIT, Python/openpyxl): "No Excel required", cross-platform; charts, PivotTables (create), slicers, sparklines, 479-function formula engine, Goal Seek, trace precedents; large-file streaming/paging with `--max-file-mb 100`, `--max-cells 100000`, `--max-pivot-cells 3000000`, `--max-loaded-cells 3000000`; VBA **read** only, VBA write opt-in (`--allow-vba-write`), **never runs macros**; folder confinement, formula safety check, read-only mode, atomic saves.
- **sbroenne/mcp-server-excel** (971★, MIT, C#/.NET): drives **real Excel via COM**; "64 MCP tools across 33 feature areas, 437 operations"; Power Query, DAX, Power Pivot, Data Model, VBA, Python-in-Excel, Goal Seek, scenarios, data tables, charts, slicers; **runs VBA**; multi-process worker pool (blocked op doesn't block controls); attach/discover an open workbook; Windows + Excel 2016+ + interactive desktop.

### Comparison table

| User requirement | xlvba | haris-musa (openpyxl) | sbroenne (COM/.NET) | Which fits better |
|---|---|---|---|---|
| Real Excel + VBA **execution** | Yes (gated `macro_run`, R13/BG06) | No (never runs macros) | Yes (runs VBA) | sbroenne or xlvba (xlvba adds gating) |
| Original file provably untouched | Yes (working copy + SHA) | Atomic saves, opens file directly | Opens live file (Excel decides read-only) | **xlvba** |
| Snapshots / restore / diff | Yes | No | No | **xlvba** |
| Controlled mutations (snapshot+approval+postcondition) | Yes | Read-only mode + formula safety | Worker isolation, no snapshot gate | **xlvba** |
| Macro safeguards (background_safe, R13 veto) | Yes | N/A (no execution) | Runs VBA, fewer gates | **xlvba** (for guarded exec) |
| VBA edit + compile diagnostics | Yes (edit + fail-closed diagnostics) | VBA write opt-in; no compile | VBA edit/run; no non-activating compile | xlvba (diagnostics honesty) |
| Structured testing (declarative suite) | Yes (`test_suite`) | No | No (has llm-tests, not a tool) | **xlvba** |
| Health / recovery surface | Yes (`server_health`/`excel_health`/`excel_recover`) | No | Worker pool health (no agent recover verb) | **xlvba** |
| Large-data paging / export | Yes (range/table/pivot + CSV/JSONL) | Yes (streaming + CSV) | GetData unbounded (no cap) | xlvba or haris-musa |
| External / Data-Model pivot read | No (fail-closed; displayed result only) | Reads via openpyxl | Reads fully (COM) | **sbroenne / haris-musa** |
| Capability breadth (charts, slicers, DAX, PQ, what-if) | Narrow | Broad (charts/slicers/pivots/formulas) | Broadest (DAX/PQ/Data Model/scenarios) | **sbroenne / haris-musa** |
| Cross-platform / no-Excel | No (Windows+Excel only) | Yes | No (Windows+Excel) | **haris-musa** |
| Fewest steps for a one-off read | More (snapshot/typed) | Fewest | Few | haris-musa |

**Positioning verdict (objective):** xlvba is **not universally better**. It is the strongest for *agent-driven, safety-first automation on real files where originals must stay untouched, writes must be verified, and macros must be gated*. haris-musa wins for *no-Excel, cross-platform, broad file-feature coverage*; sbroenne wins for *breadth of live-Excel features (DAX/PQ/Data Model) and fewer safety gates*. Competitor claims above are taken verbatim from their live READMEs; no unverified competitor claim is asserted as fact.

---

## F. Limitations (consolidated)

**Hard limitations (architectural, by design)**
- Windows + desktop Excel only; no Linux/macOS, no server-side batch. *Effect:* cannot run in containers/CI without Windows+Excel. *Alt:* haris-musa (cross-platform, no Excel).
- Single owned hidden Excel instance, single STA, FIFO. *Effect:* one long op blocks others; ~36 s for 200k×20 full extraction. *Alt:* sbroenne multi-worker pool.
- Originals never opened; edits live in `work/`. *Effect:* must `workbook_save_copy`/`save` to persist; close discards. *Alt:* none needed (safety feature).

**Unsupported (deliberate scope)**
- No workbook creation, no PivotTable/Cache create-delete, no charts/slicers, no row/col sizing/freeze/page setup, no external links/PQ, no import/print/PDF/email, no events, no component add/remove, no reference changes, no form designer. *Effect:* not a general Excel authoring tool. *Alt:* haris-musa / sbroenne for those features.
- Compile always `indeterminate`; `macro_run` needs `background_safe`; macros showing UI cannot run headless. *Effect:* VBA that shows a dialog is refused by design. *Alt:* sbroenne (runs VBA, interactive desktop).

**Environment-dependent**
- VBOM trust (manual, may be blocked by Group Policy); macro/Protected-View/Mark-of-the-Web; non-`.xlsm` formats allowed but unvalidated; `EXCEL_VBA_MCP_LOG_LEVEL` parsed but inert. *Effect:* VBA features unavailable without trust; behavior varies by machine.

**Untested / not verified (Grade 6)**
- Compile behavior on builds other than 16.0.14332; non-English Office UI; external/OLAP/Data-Model **member/CubeField enumeration** (fail-closed, hang evidence inferred, not executed on this runtime); MCP-transport run of `workbook_snapshot_read`/`data_export` (validated in-process only per large_data_results limitation 2); additive-enum tolerance of external clients (U4, assumed not verified).

**Known risks / trade-offs**
- Default `full` mode = no per-action approval (usability over sandbox). *Alt:* set `restricted` (read-only over MCP).
- R13 token scan is heuristic (misses indirect/transitive UI). *Alt:* human review before `background_safe`.
- `macro_run` has no snapshot precondition (documented spec deviation). *Alt:* snapshot manually first.
- Audit log append-only by convention, not tamper-evident.
- `snapshot_restore` caller contract footgun: must pass old `workbook_id` and switch to returned `workbook.id` (else two copies). *Alt:* follow docs 09 operational note.

---

## Doc discrepancies (with severity)

| # | Severity | Location | Discrepancy | Recommendation |
|---|---|---|---|---|
| D1 | **Major** | `.ai/tmp/xlvba-realdata-consolidated-2026-11-27.md` §Cross-cutting #1 / Scoreboard (S8.2 BLOCKED) | Describes external-pivot `fields-list` worker wedge as an OPEN Major blocker; **FIXED in HEAD** (`pivot_tables_adapter.py:977-995` `_gate_supported_source_read`; final comparison report confirms). A reader of the consolidated report alone would wrongly believe the blocker is open. | Update/annotate the consolidated report to mark the wedge RESOLVED at HEAD; point to the fix. |
| D2 | Minor | `docs/09-limitations-and-compatibility.md:98` | Notes `RANGE_TOO_LARGE` "limit-specific code pending" in the older comparison; current docs 09 line 98 already describes it as delivered — verify no residual "pending" text remains in ru mirror. | Confirm ru mirror matches EN (delivered). |
| D3 | Minor | `docs/09` line 149 / `docs/08` line 256 | Both state "table row reads are additionally capped at **2,000** cells per response", but `LargeDataConfig.max_table_cells` default is **50,000** (config.py:97-100; `tables_service` passes `max_table_cells`). The 2,000 figure is stale. | Update both docs to 50,000 (config-driven) or state "config-driven, default 50,000". |
| D4 | Minor | README "Tool catalog" / docs 02 | Accurate (43/4). No change — recorded to confirm verification. | None. |
| D5 | Advisory | `docs/08`/`docs/09` | `macro_policy.classify_operation` documented as non-authoritative/stale — accurate and well-flagged. No change. | None. |

**Top safety/limits doc discrepancies by severity:** D1 (Major, stale open-blocker), D3 (Minor, wrong table cell cap 2,000 vs 50,000), D2 (Minor, ru-mirror parity check).

---

## Recommended doc changes (advisory, not mandatory)

1. **[DOC-UPDATE] (D1, small):** Annotate `.ai/tmp/xlvba-realdata-consolidated-2026-11-27.md` S8.2 as RESOLVED at HEAD with the `_gate_supported_source_read` citation; this is the single highest-value correction because the stale report contradicts current behavior.
2. **[DOC-UPDATE] (D3, trivial):** Correct "2,000 cells per response" → "50,000 (config-driven, `max_table_cells`)" in docs 09 line 149 and docs 08 line 256 (and ru mirrors).
3. **[BEST-PRACTICE] (small):** Add a one-line "Safety contract summary" to docs 09 pointing to the working-copy + snapshot + verified-save + fail-closed-recovery chain as the product's differentiator; it is currently spread across docs 08/09 without a single canonical statement.
4. **[BEST-PRACTICE] (small):** In docs 09 "Verified vs suspected", explicitly elevate "external/OLAP/Data-Model displayed-result read = Verified (real external pivot)" and keep "member/CubeField enumeration = Suspected" — the current table lumps all pivot source behavior ambiguously.
5. **[DOC-UPDATE] (trivial):** README/docs already state compile is diagnostic-only; no change needed — recorded as a verified honesty strength to preserve.

---

## Verification-grade summary (requested)

| Item | Grade | Note |
|---|---|---|
| Compile diagnostics (never claims compiled) | **1–2** | Source + unit; fail-closed `indeterminate`/`is_compile=false` confirmed; not re-run live by me (existing Grade-4 evidence in `.ai/audit/hidden-compile-research-2026-11-19.md`) |
| Macro execution (gating) | **1–2** | Source + unit; R13/BG06/approval gating verified; actual successful macro run not demonstrated end-to-end in reviewed real-data (runs were R13-blocked, N/A) → execution-success is Grade 4/6 |
| Large-data paging / export | **1–2 + 4** | Source + unit (1615 pass) + real 30.8 MB workbook (benchmark/large_data_results, not re-run by me) |
| Pivot result-read | **1–2 + 4** | Source + unit + real external-source pivot read (benchmark) |
| Recovery | **1–2 + 4** | Source + unit + real `excel_recover` pid changes in real-data runs |

No files were modified.