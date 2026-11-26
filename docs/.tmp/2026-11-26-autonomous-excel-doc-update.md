# Documentation update 2026-11-26 — notes

## ru/ mirror not updated (translation deferred)

The English SSOT docs under `docs/` were updated for the recently shipped autonomous-Excel work
(process lifecycle / Job Object, ownership record + startup reconcile, in-process JobRegistry +
JobStore, `job_get` / `job_cancel` / `excel_recover`, `background=true` long ops, pivot single-sweep
inventory, RECOVERY security class, new env knobs).

The Russian mirror under `docs/ru/` was **not** updated in this pass. An accurate translation of the
new technical material (job lifecycle, Job Object, reconcile policy, new tools) was not attempted to
avoid introducing inaccurate Russian text. The `docs/ru/*` files are therefore stale relative to the
English SSOT until a qualified translation pass is done.

Action for a maintainer: translate the updated sections of the following English files into their
`docs/ru/` counterparts (keep the same frontmatter `id`/`domain`, translate `tags`/`related` values
only where a Russian equivalent already exists, keep relative cross-links pointing at the ru/
siblings):

- `docs/05-mcp-tools-reference.md` → `docs/ru/05-mcp-tools-reference.md`
- `docs/04-architecture.md` → `docs/ru/04-architecture.md`
- `docs/03-configuration.md` → `docs/ru/03-configuration.md`
- `docs/06-excel-capabilities.md` → `docs/ru/06-excel-capabilities.md`
- `docs/08-security.md` → `docs/ru/08-security.md`
- `docs/09-limitations-and-compatibility.md` → `docs/ru/09-limitations-and-compatibility.md`
- `docs/10-troubleshooting.md` → `docs/ru/10-troubleshooting.md`
- `docs/11-development-and-maintenance.md` → `docs/ru/11-development-and-maintenance.md`
- `docs/01-overview.md` (tool count 88 → 91) → `docs/ru/01-overview.md`

## Source-of-truth reports not present

The task referenced `.ai/audit/reports/report_07 — Autonomous Excel Observability Architecture.md`,
`rollout_report.md`, `post_reload_report.md` and `excel_process_lifecycle_report.md`. The
`.ai/audit/reports/` directory is empty in this checkout, so those files could not be read. All facts
in the doc updates were verified directly against the code under `src/excel_vba_mcp/` (config,
statuses, com_worker, app_adapter, session_service, job_registry, job_store, owned_pid_store,
pivot_tables_service/adapter, server registrations) instead.