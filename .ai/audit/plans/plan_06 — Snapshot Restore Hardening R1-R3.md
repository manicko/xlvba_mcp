# plan_06 — Snapshot Restore Hardening (R1–R3)

- **Status:** PROPOSED — review-gated. **No production code change authorized until approved.**
- **Revision:** `24201f4`. Research-only origin; see `research_D_snapshot_restore_lifecycle.md`, `report_12`, `research_B_D_decisions.md`.
- **Scope:** additive hardening ONLY. No rewrite of the WP3 lifecycle. Preserve the distinction: **original workbook** (never opened/modified), **immutable snapshot** (stored baseline), **mutable working copy** (per-session, `work/<id>.ext`).
- **Non-negotiables:** single hidden Excel, single-STA FIFO, Job Object lifecycle, working-copy model, `confirm` gating, no auto-replay of an unknown restore, SCHEMA_VERSION=1.

## R1 — Surface `previous_close` status (never fail restore on close failure)
- **Problem:** `snapshot_service.py:374-381` swallows a failed previous-copy close; the restore returns success while two copies remain open (G1).
- **Contract:** `restore_snapshot` returns the restored `WorkbookInfo` plus an additive `previous_close: Literal["closed","failed","not_requested"]` and, on failure, `previous_close_error` (safe text, no secrets). A close failure does NOT change the restore outcome (still success) and does NOT invalidate the restored session; the agent is told the previous copy remains open.
- **Changed files:** `application/snapshot_service.py` (return DTO), `api/test_tools.py` (surface in response), `domain/models.py` (additive fields), tests.
- **Tests:** close-success → `closed`; injected close failure → `failed` + restored active + previous still open; `workbook_id=None` → `not_requested`.

## R2 — Require/document `workbook_id` for restore
- **Problem:** `workbook_id=None` never closes the old copy (G6); ambiguous.
- **Contract:** `snapshot_restore` requires `workbook_id` (reject missing with `INVALID_RANGE`/validation, isRetryable=false) OR explicitly documents `not_requested` semantics if optional is retained. Chosen: **require `workbook_id`** for a deterministic single-active-copy contract. Update tool schema/description + docs (EN+RU).
- **Changed files:** `api/test_tools.py`, `server.py` `_DESC_SNAPSHOT_RESTORE`, docs `05` (+RU), tests.
- **Tests:** missing `workbook_id` → validation error; with id → close-previous path.

## R3 — Rebind restored canonical to `manifest.source_path`
- **Problem:** restored record `canonical` = snapshot file path (G5) → `list_workbooks` shows snapshot path; `excel_open(original)` duplicates; same-snapshot repeat → `COM_BUSY`.
- **Contract:** after a successful restore, set the restored record's `canonical` to `manifest.source_path` (the original), so identity/`list_workbooks`/reopen reflect the original, not the snapshot file. The working copy remains `work/<newid>.ext`; the snapshot file is immutable and not the session identity. Add a clearer same-snapshot error (`SAME_SNAPSHOT` reason) if a repeat is attempted.
- **Changed files:** `application/snapshot_service.py` (rebind + `set_snapshot_gate`), `application/session_service.py` (record canonical update), tests.
- **Tests:** after restore `canonical == manifest.source_path`; `list_workbooks` shows original; repeat same snapshot → clear error; different snapshot works.

## Dependency / sequencing
R1 and R2 are independent; R3 depends on R1 (same DTO/record touch). Order: R1 → R2 → R3 (or R1+R3 together, R2 separate). Keep each independently reviewable.

## Acceptance
- Unit tests for each R; full `pytest tests/unit -q` green; ruff clean; live integration (disposable workbook + snapshot) verifying one open copy (restored, active), original untouched, one owned Excel.
- No change to WP3 order; no new ErrorCode beyond a possible `SAME_SNAPSHOT` reason (reuse `COM_BUSY`/validation if preferred); SCHEMA_VERSION=1.

## Rollback
Pure additive; revert the R1–R3 commits to return to the accepted WP3 behavior.

## Deferred (separate, NOT in this plan)
- B3 background execution for slow supported pivots (separate future WP; keep distinct from D hardening).
- G4/G7–G10 (orphan cleanup, in-flight job finalization, stale maps, idempotency key, `_state` init) — document only unless a defect is observed.