# Tool & Action Naming — Decision Log

**Project:** `xlvba_mcp` · **Date:** 2026-11-28 · **Owner:** Tech-Lead (naming refactor)
**Scope:** Documentation/config alignment to the canonical MCP interface. The source rename is DONE and verified.
**Authoritative runtime truth:** `.ai/audit/artifacts/tools_list_after.json` — **43 tools**, annotations **14 readOnly / 16 destructive / 23 idempotent**; the 6 old names are absent.

---

## 1. Approved & implemented renames (canonical, no aliases)

Single-switchover rename of the published tool names. There are **no** backward-compatibility aliases and no `meta.deprecated` entries; the old names are removed from active docs/examples.

| # | Old name | New canonical name | Kind |
|---|----------|--------------------|------|
| 1 | `excel_save_as` | `workbook_save_copy` | tool |
| 2 | `test` | `test_suite` | tool |
| 3 | `name` | `defined_name` | tool |
| 4 | `name_read` | `defined_name_read` | tool |
| 5 | `name_delete` | `defined_name_delete` | tool |
| 6 | `workbook_file_read` | `workbook_snapshot_read` | tool |
| 7 | `test` action `run-suite` | `test_suite` action `run` | action (keep `get-run`) |

Notes:
- `get-run` is retained and means "retrieve a persisted run by id" (a read), distinct from `run` (execute).
- The Kilo permission glob `xlvba_*` matches every published name, so all renamed tools stay covered with no config change (documented in `README.md`, `docs/02-installation-and-setup.md`, `docs/ru/02-installation-and-setup.md`).

## 2. Count reconciliation (stale → runtime)

The documentation carried several stale tool counts. All were reconciled to the live `tools/list` (43 tools):

| Location | Stale | Corrected |
|----------|-------|-----------|
| `docs/05-mcp-tools-reference.md` (catalog) | 41 tools | **43 tools** |
| `docs/05` annotation distribution | 13 / 16 / 21 | **14 / 16 / 23** |
| `docs/05` group summary total | 41 | **43** |
| `README.md` tool catalog | 88 tools | **43 tools** |
| `docs/01-overview.md` | 91 tools | **43 tools** |
| `docs/02-installation-and-setup.md` migration | 92 → 41 | **92 → 43** |
| `docs/ru/05`, `docs/ru/02`, `docs/ru/01` | 41 / 91 | **43** |

The catalog was extended from 41 to 43 by adding the two non-COM tools that the earlier catalog omitted:
`workbook_snapshot_read` (read-only, non-COM streaming OOXML) and `data_export` (writes an artifact file). Both are now listed in the group summary, the per-group tables, and the annotation distribution.

### Stale fixture correction (recorded, not edited here)

A test fixture (`tests/fixtures/wp0_baseline.json`, consumed by `tests/unit/test_wp0_baseline.py`) carried a stale **94**-tool baseline. The runtime truth is **43** tools; the fixture correction 94 → 43 is a source/test-side change performed by the implementation workstream. This log records the correction only — the fixture is under `tests/` and was **not** modified by the documentation task.

## 3. DEFERRED proposals — NOT implemented (require separate Tech-Lead approval)

These additional renames were analyzed in `.ai/audit/naming/tool_action_naming_policy.md` §4 but are **not** part of the approved set. They are **not** implemented in source or docs and must be reviewed separately before any change:

| # | Current | Proposed | Rule | Status |
|---|---------|----------|------|--------|
| P1 | `excel_health` | `worker_health` | R4 (drop `excel_` prefix) | **DEFERRED** — not implemented |
| P2 | `excel_recover` | `worker_recover` | R4 / symmetry with P1 | **DEFERRED** — not implemented |
| P4 | `vba_read` action `get-compile-diagnostics` | `compile-diagnostics` | R4 (drop redundant `get-`) | **DEFERRED** — not implemented |

All other analyzed items (P3 `snapshot_*`, P5 `workbook save` vs `workbook_save_copy`, P6 `data_export`, P7 `pivot_read result-read`, P8 `macro classify-background`) were explicitly "keep / no change".

## 4. Out of scope

- **Inherited large-data WIP.** The unrelated large-data execution-contract work (`.ai/large-data/`, `data_export_service`, `ooxml` streaming, zip-safety, value-bounds) is a separate, inherited work-in-progress and is **out of scope** for this naming/documentation task. It is neither reverted nor advanced here.
- **Pre-existing 88-tool-era names in prose.** Several docs still reference the older consolidated-era per-verb names (e.g. `excel_open`, `excel_close`, `excel_save`, `excel_get_status`, `excel_list_workbooks`, `sheets_list`, `range_read_values`, `name_get`, `test_get_run`, `test_run_suite`) in narrative/examples. These predate this refactor, are **not** among the approved renames, and are **not** touched by this task (the verify scope covers only the 6 renames + `run-suite`→`run`). They remain as known documentation debt for a separate follow-up. The `docs/02` migration table intentionally keeps the historical old→new mapping.

## 5. Files touched by the documentation task

`README.md`, `docs/01-overview.md`, `docs/02-installation-and-setup.md`, `docs/03-configuration.md`, `docs/05-mcp-tools-reference.md`, `docs/06-excel-capabilities.md`, `docs/07-workflows-and-examples.md`, `docs/09-limitations-and-compatibility.md`, `docs/10-troubleshooting.md`, and the `docs/ru/` mirrors of the same. `src/` and `tests/` were **not** modified by this task.