# Documentation Fact Ledger (sole writing source for the Docs Specialist)

**Repo HEAD:** `f79faf2b8797a897429754d71e89f8fb9d031317` (master, clean) · **Date:** 2026-11-28 · **Owner:** Tech Lead
**Sources:** `.ai/audit/docs/audit_1_product_capabilities.md`, `audit_2_tools_and_actions.md`, `audit_3_safety_limits_positioning.md` (all independent, evidence-graded).
**Verification levels:** 1 Source · 2 Unit-tested · 3 Integration · 4 Real-workbook · 5 MCP-runtime · 6 Unverified.
**Rule for Docs Specialist:** Use ONLY statements below. If a fact is not here, do not invent it — return it to the Tech Lead. Preserve qualifications/limitations verbatim in meaning. Do NOT modify code (including `_DESC_*` strings). English-only; do not create new non-English content.

---

## A. Canonical surface (do not contradict)

| Statement | Confirmed wording | Evidence | Level |
|---|---|---|---|
| Tool count | **43 MCP tools** | `mcp.list_tools()`=43 (audits 1/2/3) | 5 |
| Composition | 20 dispatch tools (with `action`) + 23 standalone | audit_2 | 5 |
| Action count | 74 total action values across the 20 dispatch tools | audit_2 | 5 |
| Annotations | readOnly **14** / destructive **16** / idempotent **23** | audit_2/3 | 5 |
| Resources | **4 data resource templates** + **1 static glossary** (`xlvba://glossary`) = 5 total | `list_resources`+`list_resource_templates` (audit_2/3) | 5 |
| Canonical renames (no aliases) | `workbook_save_copy`, `test_suite`, `defined_name`, `defined_name_read`, `defined_name_delete`, `workbook_snapshot_read`; old names absent | audits 1/2/3 | 5 |
| `test_suite` actions | `run`, `get-run` (`run-suite` rejected) | audit_2 | 5 |

**Doc action (README, LOW):** README says "43 MCP tools and 4 read-only MCP resources". Clarify to "4 data resources plus a glossary resource (5 total)" OR keep "4 read-only resources" + the existing glossary note — either is acceptable; make the count unambiguous. Keep the 43-tool catalog table exactly as-is (already canonical).

---

## B. Per-domain capabilities (what a user can do) — all 19 implemented

For each: user problem → operations → verification → key limitation. (Full per-action detail: audit_2.)

1. **Workbook/session** — open an existing workbook as a hidden working copy; close/save without touching original. `workbook`(open/close/save), `workbook_read`(status/list/model). L2/3/5. Limit: original never modified; no new-workbook creation; non-`.xlsm` unverified (6).
2. **Save-as copy** — produce a new/again-saved copy without rebinding session. `workbook_save_copy` (atomic temp→replace, `.bak`, size+SHA-256 verify, fail-closed). L2/3/4. Limit: existing target needs `overwrite_existing`; overwriting source needs `overwrite_source`+matching hash.
3. **VBA read** — inspect components/procedures/references/search. `vba_read`(7 actions), `vba_export_project`. L2/3/4. Limit: needs VBOM trust + unlocked project.
4. **VBA edit** — replace module / insert code at semantic position (revision-guarded). `vba_replace_component_source`, `vba_insert_code`. L2/3/4. Limit: Standard/Class modules only; `REVISION_CONFLICT` on stale hash.
5. **Compile diagnostics** — `vba_compile_diagnostics`, `vba_read(get-compile-diagnostics)`. **ALWAYS `indeterminate` / `NON_ACTIVATING_COMPILE_UNAVAILABLE`; `is_compile=false`; NEVER returns `compiled`/`compile_error`.** L1/2. (Honesty strength — preserve.)
6. **Macro execution** — run an approved non-interactive macro, typed result. `macro_run`; `macro`(list/classify-background/get-last-result). Gating: approval + `background_safe`(BG06) + R13 veto; 90 s soft deadline; hang⇒self-recovery. L1/2 (gating); successful run e2e = L4/6 (reviewed runs were R13-blocked/N/A). Limit: deny-by-default; no dialog dismissal; VBA showing UI is refused.
7. **Jobs/cancel** — async long ops; poll; record cancel intent. `job_get`, `job_request_cancel`. L2. Limit: cancel is intent-only, never interrupts a running call.
8. **Snapshots** — immutable baseline; list/diff/restore. `snapshot_create/list/restore/diff`. L2/3/4. Limit: restore destructive to working copy, needs `confirm=true`+`workbook_id`; does not undo external macro side effects.
9. **Declarative test runner** — bounded edit→macro→assert suites on disposable copies. `test_suite`(run/get-run); closed vocab (37 actions, 29 assertions). L2/3/4. Limit: `extra=forbid`; VBA replace/insert steps rejected; external side effects not rolled back.
10. **Sheets** — list/get/add/rename/copy/delete/visibility. `sheet_read`, `sheet`, `sheet_delete`. L2/3/4. Limit: add=worksheet-only; copy in-workbook; delete refuses zero visible; no move/reorder.
11. **Tables** — list/get/read rows; append/insert/update/delete rows; add/delete column; create/resize. `table_read`, `table`, `table_delete_rows`, `table_delete_column`. L2/3/4. Limit: bounded reads (`truncated`); ambiguous `where` rejected.
12. **Ranges/cells** — read/write values/formulas/format/validation/CF; copy/insert/delete/merge/unmerge. `range_read`(6), `range_write`(5), `range_structural`(3), `range_delete`, `range_merge`. L2/3/4. Limit: 100k cells/target, 50k/response; `=`-string stored as text; ListObject intersection rejected; explicit shift required.
13. **Defined names** — list/get/create/update/delete (workbook & worksheet scope). `defined_name_read`, `defined_name`, `defined_name_delete`. L2/3/4. Limit: explicit scope; duplicates/cross-scope rejected; `RefersTo` opaque (no eval).
14. **PivotTables** — inventory; refresh table/cache; field layout/position/aggregation/properties; manual filters; **read rendered result**. `pivot_read`(6 incl. `result-read`), `pivot_refresh`(2), `pivot_field`(6). L2/3/4. Limit: database/Table sources for field ops; external/OLAP/Data-Model field/member/filter/refresh = `unsupported_source`; displayed-result read works for all source types; manual filters only; no create/delete; never `RefreshAll`.
15. **Shapes** — list/get; assign/clear `OnAction`. `shape`. L2/3/4. Limit: inventory+OnAction only; macro never executed; no create/delete/format.
16. **Controls** — list/get; assign `OnAction` to Form Controls. `control`. L2/3/4. Limit: ActiveX rejected for OnAction; no create/delete.
17. **Large-data export** — stream range/table/pivot result to CSV/JSONL artifact (bounded memory). `data_export` (returns path/rows/bytes/sha256/completeness). L2/4. Limit: metadata only; `limit_rows` may truncate.
18. **Non-COM snapshot read (OOXML)** — metadata / sheet rows / saved PivotCache records from an immutable snapshot without the Excel worker. `workbook_snapshot_read`(3). L2. Limit: reads saved cache records, not displayed result; `unsupported_source` when no local records.
19. **Environment/health/recovery** — `xlvba_environment`(non-COM), `excel_health`(non-COM+pid), `excel_recover`(confirm, pid-scoped, originals-safe), `server_health`. L1/2/4. Limit: `excel.build` always null; recovery needs `confirm=true`; no auto-terminate.

**NOT implemented (verified absent — do NOT document as features):** new-workbook creation; PivotTable/PivotCache create/delete; charts/PivotCharts/slicers; row/column sizing/freeze/page setup; protection changes; forced recalculation; external links/connections/PQ; import/export/print/PDF/email; events; component add/remove; reference changes; form designer; GUI automation. (L1/6.)

---

## C. Verified end-to-end workflows (use these in docs/07; label verification)

**Verified (real-workbook integration test composes the chain, L3/4):** open→inspect→close; VBA insert/replace (revision-guarded); classify→run→get-last-result; test-suite debug loop (isolation, compile-error-blocks-macro); snapshot→mutate→restore; large-table edit; paged range read; pivot read/refresh/result-read; structure inspect before edit; save-as copy; hung-macro recovery.
**Source-composed (unit-tested links, no single e2e — label "illustrative", L2):** pivot-result→`data_export`; multi-session; non-COM snapshot read e2e; large-data export to artifact.

12 priority use cases (verified-first): audit unfamiliar workbook; add macro to existing `.xlsm`; run reviewed macro + capture result; regression-test with disposable copies; roll back a bad change; bulk-edit large table; page a huge range; refresh+read PivotTable; export large dataset; produce independent copy (Save As); recover wedged Excel; discover policy/capabilities.

---

## D. Safety & data-integrity contract (docs/08, docs/09 — user terms)

- **Original never modified:** open copies the allowlisted original into `work/`; SHA-256 captured pre-copy. (L1)
- **Working copy vs source:** edits live only in the working copy; `source_path`=original. (L1)
- **Snapshots:** immutable, SHA-256 manifest. **Restore:** fresh working copy from baseline, previous closed WITHOUT saving, identity rebound to original, external side effects NOT undone. (L1)
- **Save vs copy:** `workbook save` writes working copy in place; `workbook_save_copy` = atomic temp→replace to a (new/original) path + `.bak` + size/SHA verify, session stays bound. (L1)
- **Overwrite/backup:** existing target needs `overwrite_existing`; original needs `overwrite_source`+matching hash; single `.bak`; pre-existing `.bak` rejected. (L1)
- **Mutation preconditions:** writable session + successful snapshot + approval + postcondition re-read. (L1)
- **Approval:** `full` (default) auto-allows; `restricted` deny-by-default; destructive ops need `confirm=true`. (L1)
- **Path restrictions:** canonicalise+realpath; block traversal/symlink/UNC/device; extension allowlist; roots via `EXCEL_VBA_MCP_ROOTS`. (L1)
- **Errors:** structured JSON with stable `code`/`context`. **Interrupted ops:** `interrupted_unknown` (`outcome=unknown`, `retryable=false`), never auto-replayed. (L1)
- **Macro security:** honored, never weakened permanently. **Compile:** diagnostic-only, never claims compiled. **R13:** heuristic veto before dispatch. **BG06:** classification keyed by (workbook hash, source hash, procedure). (L1/2)

---

## E. Large data & PivotTables (qualified — do NOT overstate)

Implemented + real-data-tested (L1/4): large range paging (200k×20 in 105 pages), pivot result-read (external 50×20), `data_export` (peak 33.9 MB / 5000 rows). Implemented (L1/2): table paging, OOXML streaming, PivotCache read, completeness/truncation (`STALE_CURSOR`).
Limits (L1): `max_range_cells=100k`, `max_response_cells=50k`, `max_response_bytes=8MiB`, `max_table_cells=50k`, `max_pivot_result_cells=50k`, `max_pivot_members=10k`, `max_export_rows=10M`, `max_file_jobs=2`; file `max_file_mb=100`, `max_uncompressed_mb=1024`.
Partially supported (L1): member enumeration = DATABASE only ≤10k (external/OLAP/Data-Model fail-closed); external/OLAP/Data-Model = displayed-result read supported, field/member/filter/refresh `unsupported_source`.
Not implemented: CSV import, server-side filter/aggregation, Parquet, `.xlsb`; pivot create/delete, charts, slicers.
**MCP-transport run of `workbook_snapshot_read`/`data_export` validated in-process only (limitation 2) — label accordingly (L6 for transport).**

---

## F. Limitations (consolidated; docs/09)

- **Hard:** Windows+desktop Excel only; single owned hidden Excel (FIFO — one long op blocks others, ~36 s for 200k×20); originals never opened (must save/copy to persist).
- **Unsupported (scope):** see §B not-implemented; compile always indeterminate; `macro_run` needs `background_safe`; UI macros refused.
- **Environment-dependent:** VBOM trust (may be Group-Policy blocked); macro/Protected-View/Mark-of-the-Web; non-`.xlsm` allowed but unvalidated; `EXCEL_VBA_MCP_LOG_LEVEL` parsed but inert.
- **Untested (6):** non-16.0.14332 builds; non-English Office UI; external member/CubeField enumeration; MCP-transport run of the 2 large-data tools; additive-enum tolerance.
- **Risks/trade-offs:** default `full` = no per-action approval; R13 heuristic; `macro_run` has no snapshot precondition; audit log append-only by convention; `snapshot_restore` caller must switch to returned `workbook.id`.

---

## G. Positioning (objective; docs/01 + README optional section)

Verified live (2026-11-28) from competitor READMEs. xlvba is **not universally better**.
- **xlvba strongest for:** agent-driven, safety-first automation on real files where originals stay untouched, writes are verified, macros are gated; snapshots/restore; structured testing; health/recovery.
- **haris-musa (openpyxl, no Excel, cross-platform):** broad file-feature coverage (charts/slicers/pivot create/formula engine); VBA read-only, never runs macros. Better for no-Excel/cross-platform.
- **sbroenne (COM/.NET, real Excel):** broadest live-Excel features (DAX/PQ/Data Model/scenarios), runs VBA, multi-worker. Better for breadth / fewer safety gates.
Comparison table (use as-is):

| User requirement | xlvba | haris-musa | sbroenne | Fits better |
|---|---|---|---|---|
| Real Excel + VBA execution | Yes (gated) | No | Yes | sbroenne/xlvba |
| Original provably untouched | Yes (copy+SHA) | Atomic, opens file | Opens live file | **xlvba** |
| Snapshots/restore/diff | Yes | No | No | **xlvba** |
| Controlled mutations | Yes | Read-only+formula check | Worker isolation | **xlvba** |
| Macro safeguards | Yes | N/A | Fewer gates | **xlvba** |
| VBA edit + compile diagnostics | Yes (fail-closed) | Write opt-in | Edit/run | xlvba (honesty) |
| Structured testing | Yes (`test_suite`) | No | No | **xlvba** |
| Health/recovery | Yes | No | Worker pool | **xlvba** |
| Large-data paging/export | Yes | Yes | Unbounded | xlvba/haris |
| External/Data-Model pivot read | Displayed result only | Yes | Yes (full) | sbroenne/haris |
| Capability breadth | Narrow | Broad | Broadest | sbroenne/haris |
| Cross-platform/no-Excel | No | Yes | No | **haris** |

---

## H. Per-document actions (authoritative)

- **README.md** — PRESERVE structure/tone/length. Keep 43-tool catalog table (canonical). Clarify resource count (§A). Keep compile note. Optionally add a short, factual differentiation pointer to docs/01. No old names.
- **docs/01-overview.md** — ensure 43-tool count; add user-oriented "who it's for / when to use / when an alternative fits" + pointer to comparison (§G). No code architecture.
- **docs/02-installation-and-setup.md** — verify steps, requirements, Kilo/Claude configs, `EXCEL_VBA_MCP_ROOTS`, health check. Keep migration table (historical old→new is allowed there).
- **docs/03-configuration.md** — verify env vars/defaults against config.py (`EXCEL_VBA_MCP_ROOTS`, `EXCEL_VBA_MCP_ACCESS_MODE`, `EXCEL_VBA_MCP_LOG_LEVEL` [parsed-but-inert], timeouts, large-data limits). User terms.
- **docs/04-architecture.md** — keep but reframe to user-visible consequences (single sequential Excel session; long op blocks others; recovery). Minimal internal detail.
- **docs/05-mcp-tools-reference.md** — add `pivot_read` action **`result-read`** (EN; RU is out of scope this task). Keep comprehensive; human terms; valid examples. (Do NOT edit `_DESC_PIVOT_READ` — that is code; record as a separate code follow-up.)
- **docs/06-excel-capabilities.md** — MAJOR: replace all pre-consolidation per-verb names (66) with the 43-tool surface; add rows for `data_export`, `workbook_snapshot_read`, `job_get`/`job_request_cancel`, `xlvba_environment`, `workbook_save_copy`; remove `excel_list_workbooks`/`excel_get_status`; restate coverage counts vs 43.
- **docs/07-workflows-and-examples.md** — MAJOR: rewrite recipes with canonical names; `xlvba_environment` is implemented (not "planned"); provide the §C verified workflows (goal → setup → steps → result → verify → limitations); label tested vs illustrative.
- **docs/08-security.md** — fix table read cap **2,000 → 50,000 (config-driven `max_table_cells`)** (line ~256); keep safety contract; user terms.
- **docs/09-limitations-and-compatibility.md** — fix **2,000 → 50,000** (line ~149); separate Implemented/Tested/Untested; external pivot displayed-result = Verified, member enumeration = Suspected; add a one-line safety-contract summary; keep compile honesty.
- **docs/10-troubleshooting.md** — lead with user symptoms (Excel not responding; workbook already open; unsupported pivot source; range too large; macro needs interactive UI); error codes secondary.
- **docs/11-development-and-maintenance.md** — developer-facing; update any tool-count/name references only; otherwise keep.
- **docs/ru/** — OUT OF SCOPE this task (English-only policy). Do NOT create/modify non-English content. Note in results that ru mirrors are now stale and need a separate sync decision.
- **`.ai/tmp/xlvba-realdata-consolidated-2026-11-27.md`** (D1 Major) — annotate S8.2 external-pivot wedge as RESOLVED at HEAD (`pivot_tables_adapter.py:977-995`). This is an `.ai/tmp` working note, not published docs; fix if trivial, else record disposition.

**Code follow-ups (NOT this task — record only):** `_DESC_PIVOT_READ` omits `result-read` (D-02); `inputSchema.title` cosmetic internal names (prior task).

---

## I. Reconciled conflicts
- Resource count: **5 total = 4 templates + 1 glossary** (README "4" = data templates; clarify). 
- Table read cap: **50,000** (not 2,000) — docs 08/09 wrong.
- External-pivot wedge: **FIXED** at HEAD (stale `.ai/tmp` report only).
- README + docs/05 are CURRENT; docs/06 & docs/07 are STALE (HIGH) — primary rewrite targets.