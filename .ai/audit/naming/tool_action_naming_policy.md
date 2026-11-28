# WP2 — Tool & Action Naming Policy

**Project:** `xlvba_mcp` · **Repo HEAD:** `0261683` · **Date:** 2026-11-28
**Purpose:** Define tool-naming and action-naming rules, benchmark against two real Excel MCP servers, and list ADDITIONAL rename proposals (require Tech-Lead approval). Research only.

---

## 1. Reference points (actual current structures)

### 1.1 `haris-musa/excel-mcp-server` (file-parser, no Excel required)

Source: `TOOLS.md` (generated from schemas) + README, verified 2026-11-28. ~46 tools, grouped by area. **Flat, verb-first, one purpose per tool**; no `action` dispatch — each tool is a single operation.

- Workbooks: `create_workbook`, `describe_workbook`, `set_workbook_settings`, `list_workbooks`, `export_workbook`, `import_workbook`
- CSV: `import_csv`, `export_csv`
- Sheets: `describe_sheet`, `create_sheet`, `rename_sheet`, `copy_sheet`, `delete_sheet`, `insert_rows_or_columns`, `delete_rows_or_columns`
- Cells: `read_range`, `write_range`, `clear_range`, `copy_range`, `sort_range`, `transform_range`, `find_cells`, `replace_cells`
- Formulas: `trace_formula`, `goal_seek`
- Formatting: `format_range`, `merge_cells`, `set_sheet_layout`, `add_conditional_format`, `add_data_validation`
- Objects: `set_table`, `create_chart`, `delete_chart`, `create_pivot_table`, `delete_pivot_table`, `add_slicer`, `delete_slicer`, `insert_image`, `delete_image`, `add_sparklines`, `delete_sparklines`
- Names/notes: `set_defined_name`, `delete_defined_name`, `set_note`, `delete_note`
- Macros: `read_vba`; (opt-in) `write_vba_module`, `delete_vba_module`

**Naming traits:** `snake_case`; verb-first (`create_`, `read_`, `write_`, `delete_`, `set_`, `add_`, `rename_`, `copy_`, `move_`); domain noun is the **object** of the verb (`_workbook`, `_sheet`, `_range`, `_table`, `_defined_name`, `_pivot_table`); multiword verbs use `_or_` / `_` joins (`insert_rows_or_columns`, `set_workbook_settings`). **No tool carries an `action` enum** — the verb IS the tool name. Read-only tools are a distinct prefix family (`describe_`, `read_`, `list_`, `find_`, `trace_`).

### 1.2 `sbroenne/mcp-server-excel` (live COM automation, Windows + Excel)

Source: README + `everymcpserver` tool dump + `IWorkbookCommands.cs`, verified 2026-11-28. **64 tools across 33 feature areas, 437 operations.** Uses a **consolidated tool + `action` enum** model — the same shape as `xlvba`.

- Each area is one tool with an `action` enum; most areas have a `_read` companion for read-only actions (e.g. `file` / `file_read`, `worksheet` / `worksheet_read`, `range` / `range_edit` / `range_format` / `range_link`).
- Action values are **kebab-case verbs** that do NOT repeat the parent noun: `file` actions include `open`/`create`/`attach`/`save`/`close`/`approve`; `worksheet` actions `create`/`rename`/`delete`/`move`/`copy`/`copy-to-file`/`move-to-file`; `window` actions `show`/`hide`/`bring-to-front`/`set-state`/`set-position`/`arrange`/`freeze-panes`/`set-zoom`; `workbook` actions `refresh-all`/`save-as`/`save-copy-as`/`export-fixed-format`/`set-protection`/`get-protection`/`inspect-document`/`list-external-links`/`break-external-link`/`apply-theme`.
- Read-only actions are marked `[McpReadOnlyActions(...)]` in source; destructive tools carry `Destructive = true`.
- Newer `batch` tool runs a list of commands in one call.

**Naming traits:** `snake_case` tool names; **kebab-case action values**; action verbs are concise and rely on the parent tool for domain context (no `worksheet_create`, just `create`); compound actions use kebab (`save-copy-as`, `refresh-all`, `freeze-panes`, `break-external-link`); `_read` suffix cleanly separates read surfaces.

### 1.3 Benchmark takeaways for `xlvba`

| Concern | haris-musa | sbroenne | xlvba (current) |
|---|---|---|---|
| Tool case | snake_case | snake_case | snake_case |
| Action case | (no actions) | kebab-case | kebab-case |
| Consolidation | one op per tool | tool + action enum | tool + action enum (like sbroenne) |
| Read split | prefix family | `_read` companion | `_read` companion |
| Domain word in action | n/a | omitted (parent supplies it) | mostly omitted (good) |
| Redundant domain in tool | none | none | a few (`excel_save_as`, `workbook_file_read`) |

`xlvba` already follows the sbroenne pattern closely. The approved renames move it further toward that model (drop redundant `excel_`/`file_` prefixes, disambiguate `test`→`test_suite`, `name`→`defined_name`).

---

## 2. Tool-naming rules (normative)

1. **Case:** `snake_case`, lowercase ASCII. (Matches both references and the existing surface.)
2. **Verb-first for mutations, noun-first for reads/lifecycle:** mutating tools lead with the domain noun + verb intent is carried by actions (`sheet`, `table`, `range_write`); pure readers use `_read` (`sheet_read`, `table_read`, `range_read`, `defined_name_read`, `workbook_snapshot_read`).
3. **Concise when the parent supplies the domain; compound only to disambiguate.** `workbook_save_copy` (not `excel_save_as_workbook_copy`); `defined_name` (not `name`).
4. **No redundant domain words.** Drop the server/product prefix (`excel_`) and the file-format word (`file`/`workbook_file`) when the meaning is already clear. This is the basis for `excel_save_as→workbook_save_copy` and `workbook_file_read→workbook_snapshot_read`.
5. **No shortening that creates ambiguity.** `name`→`defined_name` (Excel "Name" is ambiguous with a cell's name / a named range); `test`→`test_suite` (a bare `test` collides with the test-runner concept and reads as a verb).
6. **One tool = one cohesive surface.** A tool groups actions that share a resource and guard profile; do not merge unrelated resources.
7. **Standalone high-risk tools keep their own name** (own permission/confirmation), even when a dispatcher exists (`macro_run`, `sheet_delete`, `range_delete`, `range_merge`, `table_delete_rows`, `table_delete_column`, `defined_name_delete`, `excel_recover`, `workbook_save_copy`).
8. **`_read` suffix is the read/mutate boundary marker** for a resource that has both surfaces.
9. **Lifecycle/capability tools are named by role:** `workbook` (lifecycle), `server_health`, `excel_health`, `xlvba_environment`, `job_get`, `job_request_cancel`.
10. **Preserve the `xlvba_` client prefix contract** — renames change the published name only; the Kilo glob `xlvba_*` continues to match (see §6).

---

## 3. Action-naming rules (normative)

1. **Case:** `kebab-case`, lowercase (matches sbroenne; already used: `set-visibility`, `append-rows`, `field-set-layout`, `get-last-result`).
2. **Concise verb when the parent supplies the domain.** Under `sheet`, use `add`/`rename`/`copy`/`set-visibility` (not `sheet-add`). Under `workbook_snapshot_read`, use `metadata`/`sheet-rows`/`pivot-cache-records`.
3. **Compound (kebab) when it disambiguates.** `refresh-table` vs `refresh-cache`; `field-set-layout`; `set-on-action`; `get-last-result`; `classify-background`; `run` (was `run-suite`).
4. **No action value may redundantly repeat a word already in the parent tool name.** (Basis for `run-suite`→`run` under `test_suite`: "suite" is already in the parent.)
5. **Read actions are a closed, consistent family:** `list`, `get`, `read-*`, `*-list`, `*-get` (e.g. `tables-list`, `table-get`, `read-rows`, `read-values`). Keep the `read-` prefix for row/value reads and `get`/`list` for metadata.
6. **Mutation verbs are `create`/`update`/`write-*`/`set-*`/`add-*`/`insert-*`/`append-*`/`delete`/`clear-*`/`apply-*`/`refresh-*`.** Reserve `delete` for the destructive standalone tools; dispatchers that intentionally exclude a destructive verb say so (`defined_name` excludes `delete`; `table` excludes `delete-column`).
7. **`get-run` is retained and means "retrieve a persisted run by id"** (see §5). It is a read, distinct from `run` (execute).
8. **Unknown action ⇒ structured `INVALID_RANGE`, never a silent default.** (Existing dispatcher contract; keep it.)

---

## 4. Recommendation gate — ADDITIONAL renames (PROPOSALS, need Tech-Lead approval)

> The following are **proposals only**. They are NOT approved and must not be implemented without explicit Tech-Lead sign-off. Each is bounded and justified against the rules above and the two references.

| # | Current | Proposed | Rule | Rationale | Risk |
|---|---------|----------|------|-----------|------|
| P1 | `excel_health` | `worker_health` | R4 (drop `excel_` prefix) | It reports the **worker/session** state (worker_state, accepting, owned_pid), not Excel application health; `server_health` already covers server liveness. `worker_health` disambiguates the two health tools. | Low — but `excel_` prefix is a recognizable family; some clients may expect it. |
| P2 | `excel_recover` | `worker_recover` | R4 / symmetry with P1 | Recovers the owned **Excel process / worker** (terminates owned `EXCEL.EXE`, discards working copies). Pairs with `worker_health`. | Low–Med — "recover" is high-risk; a rename touches approval docs/globs. |
| P3 | `snapshot_*` (4 tools) | keep | — | No change recommended. `snapshot_create/list/restore/diff` are already concise and unambiguous. | None |
| P4 | `vba_read` action `get-compile-diagnostics` | `compile-diagnostics` | R4 (drop redundant `get-` when parent is `vba_read`) | The `get-` prefix is redundant under a `_read` parent; sbroenne uses bare nouns for reads (`get-info` is the exception, not the rule). | Low — purely cosmetic; changes an action value. |
| P5 | `workbook` action `save` vs `workbook_save_copy` | keep both | R7 | `save` (in-place) and `workbook_save_copy` (Save As / copy) are distinct and correctly separated. No change. | None |
| P6 | `data_export` | keep | R7 | Standalone, own permission, clear. No change. | None |
| P7 | `pivot_read` action `result-read` | keep | R5 | `result-read` is a read of the rendered result; consistent with `read-*` family. No change. | None |
| P8 | `macro` action `classify-background` | keep | R3 | Compound is warranted (classifies the `background_safe` flag); not redundant with `macro`. No change. | None |

**Net additional-rename proposals requiring approval: P1 (`excel_health`→`worker_health`), P2 (`excel_recover`→`worker_recover`), P4 (`vba_read` action `get-compile-diagnostics`→`compile-diagnostics`).** All others are explicitly "keep" (no change). These are bounded, low-blast-radius, and improve consistency with the reference model; they are **not** part of the approved set and must be reviewed separately.

---

## 5. Verdicts requested by Tech Lead

### 5.1 `workbook_file_read` snapshot-semantics verdict → **RENAME TO `workbook_snapshot_read` IS JUSTIFIED**

Evidence (source of truth):
- `src/excel_vba_mcp/api/ooxml_tools.py` docstring: "reads worksheet rows or workbook metadata from an **immutable snapshot file** (a `Workbook.SaveCopyAs` image) without occupying the single-STA COM worker".
- `src/excel_vba_mcp/application/ooxml_service.py` `_resolve_input()`: when `snapshot_id` is given it loads that snapshot from the store; **otherwise it creates a fresh snapshot** (`self._snapshot.create_snapshot(workbook_id)`) and reads the snapshot's `stored_path`. The only input is the snapshot file — it "never touches the live working copy or the original".
- The read path parses the snapshot file via `open_safe(path, ...)` (pure file I/O + `xml.etree` in a worker thread). It does **not** read a live file representation independent of snapshots.

**Conclusion:** the implementation reads immutable snapshots exclusively. `workbook_snapshot_read` is semantically accurate and the approved rename is justified. No exception needed.

### 5.2 `get-run` meaning verdict → **RETAIN; it means "retrieve a persisted run by id"**

Evidence: `src/excel_vba_mcp/api/test_tools.py` `test_get_run(run_id)` → "Return a previously persisted test run" via `runner.get_run(run_id)`; the `test` dispatcher routes `get-run` to `test_get_run(run_id)`. It is a pure read of a persisted `TestRunResult` by `run_id`, distinct from `run-suite` (execute). Under the renamed parent `test_suite`, `run` (execute) + `get-run` (read by id) is a clean, non-redundant pair. **Keep `get-run` unchanged.**

---

## 6. Kilo MCP client configuration — location & glob

- **Documented pattern (authoritative for this project):** server-name glob `xlvba_*`, set in the client `permission` block as `"xlvba_*": "ask"`. Source: `docs/02-installation-and-setup.md` lines 108-137; `docs/05-mcp-tools-reference.md` line 81; `docs/03-configuration.md` (approval globs note).
- **Live config file on this machine:** the environment reports the global Kilo config directory `C:\Users\MatveichevN\.config\kilo\` (same structure as the project `.kilo/`). The expected file is `C:\Users\MatveichevN\.config\kilo\kilo.json` (schema `https://app.kilo.ai/config.json`).
- **Read status:** this research session is **blocked by a global `external_directory` deny rule** (`{"permission":"external_directory","pattern":"*","action":"deny","source":"global"}`) from reading any path outside the workspace, so the live `kilo.json` glob value could **not** be read directly. The glob is therefore reported from in-repo documentation (`xlvba_*`), which is the project's authoritative statement of the pattern.
- **No `kilo.json` exists inside the repo** (glob `**/kilo.json` → none); the project uses `.kilo/` for agents/commands/rules only.
- **Effect of renames on the glob:** none — `xlvba_*` matches any published tool name with the `xlvba_` client prefix, so `xlvba_workbook_save_copy`, `xlvba_test_suite`, `xlvba_defined_name`, `xlvba_workbook_snapshot_read` all continue to match. No glob change is required for the approved renames. (If P1/P2 were approved, `xlvba_worker_health` / `xlvba_worker_recover` would also still match `xlvba_*`.)

---

## 7. Migration note (no aliases)

Per the approved decision there are **no backward-compat aliases**: this is a single-switchover rename of the published names. The fixture `test_description_contract.py::test_single_switchover_no_legacy_names_or_aliases` already enforces that no retired old name is registered and no `meta.deprecated` alias is exposed. Renames must update: the `name=` argument in `server._register_tools`, the `_ANN_*` frozensets, the `_DESC_*` constants' cross-references, `docs/05-mcp-tools-reference.md`, and the fixture ledger targets — but those are implementation tasks, out of scope for this research.