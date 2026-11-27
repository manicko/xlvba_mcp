# Snapshot Restore — Final Caller Contract (R1–R3)

- **Status:** R1–R3 implemented and verified (Auditor `report_13` ACCEPT; Validator `report_14a` APPROVE). This document is the authoritative caller contract; no rewrite of the verified lifecycle.

## R1 — previous-copy cleanup status: COMPLETE
`SnapshotRestoreResult.previous_close ∈ {closed, failed, not_requested}` (+ safe `previous_close_error`). A previous-copy close failure is caught, surfaced as `failed`, and NEVER fails the restore nor invalidates the restored session; the previous copy remains open and the caller is told. `workbook_id=None` → `not_requested`.

## R2 — workbook identity: COMPLETE (workbook_id required)
`snapshot_restore` REQUIRES `workbook_id` (the OLD/previous working-copy id to close). Missing/empty/stale → `INVALID_RANGE` (`isRetryable=false`, reason `workbook_id_required`) BEFORE any open/close. Identity is resolved by stable id, never by filename.

## R3 — canonical path & repeated restore: COMPLETE
After open+validate, the restored record is rebound: `canonical = manifest.source_path` (ORIGINAL), `source_sha256 = manifest.sha256`; working copy stays `work/<newid>.ext`; the snapshot file is immutable and never the session identity. Same-snapshot repeat while active → deterministic `same_snapshot` isError (non-retryable); different snapshot → allowed (new working copy, previous closed). Original never modified (copy-not-move).

## Caller contract (agent-facing)
1. Pass `workbook_id` = the id of the working copy you want replaced/closed.
2. On success, use the RETURNED `workbook.id` (NEW id) for all subsequent calls; the old id is closed.
3. Check `previous_close`: `closed` (old closed), `failed` (old still open — close it explicitly or continue with the new id), `not_requested` (no id supplied).
4. Repeating the SAME snapshot → `same_snapshot` error (use a different snapshot or continue editing).
5. A no-op restore is allowed only when the active content matches the snapshot; if the working copy changed, restore actually rolls it back to the snapshot.

## Lifecycle (verified order)
load manifest → open snapshot as new working copy → validate restored → close previous by stable id (discard) → rebind identity to original → audit/return. Invariants: never close source/restored/unrelated; don't close previous before replacement valid+active; create/validate/activate failure never closes previous; close-previous failure never invalidates restored; cleanup by id.

## Repeated-restore semantics
- First restore: old→new, old closed, new canonical=original.
- Repeat same snapshot (no-op or after edits): `same_snapshot` deterministic error (no duplicate open).
- Restore different snapshot: allowed; previous (current) closed, new active.
- Interruption between activation and cleanup: restored active + previous open, indeterminate, NO auto-replay; recovery explicit.

## Tests
Unit (restore/confirm) + live `test_snapshot_restore_hardening.py` (5 passed): success, R1 close-failure, R2 missing-id, R3 canonical/same/different/no-op/rollback, interruption-before-cleanup. Original hash unchanged; one owned Excel.