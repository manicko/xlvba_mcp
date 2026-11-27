# R1–R3 Snapshot Restore Hardening — Implementation Report

- **Revision:** `24201f4` + R1–R3 (uncommitted). Authorized by Product Owner; B3 deferred (not implemented).
- **Auditor:** `report_13 — R1-R3 Audit.md` → **ACCEPT** (all 6 items PASS).
- **Design prerequisite:** `plan_06a — R3 Path and Identity Lifecycle.md`.

## What shipped (additive; WP3 order preserved: load→open→validate→close-previous→rebind→audit)
- **R1 — `previous_close` status:** `SnapshotRestoreResult` DTO adds `previous_close ∈ {closed,failed,not_requested}` + safe `previous_close_error`. A previous-copy close failure is caught broadly (`snapshot_service.py:466-498`), surfaced as `failed`, and NEVER fails the restore nor invalidates the restored session. `workbook_id=None`→`not_requested`.
- **R2 — require `workbook_id`:** `snapshot_restore` rejects missing/empty `workbook_id` with `INVALID_RANGE` (isRetryable=false) before any open/close (`test_tools.py:157-167`). Schema/description/docs (EN+RU) updated.
- **R3 — canonical/session mapping:** after open+validate, `rebind_restored_identity` sets restored `canonical=manifest.source_path` (original) + `info.source_path` + `source_sha256=manifest.sha256`, migrates per-path active/lock maps; working copy stays `work/<newid>.ext`; snapshot file immutable. Same-snapshot repeat → deterministic `same_snapshot` isError before open; different snapshot allowed. Original never modified (copy-not-move).

## Files changed
`domain/models.py` (SnapshotRestoreResult), `application/snapshot_service.py` (R1/R3), `application/session_service.py` (`rebind_restored_identity`), `api/test_tools.py` (R2 + DTO), `server.py` (`_DESC_SNAPSHOT_RESTORE`), docs `05` (+RU), tests (unit + `tests/integration_excel/test_snapshot_restore_hardening.py`).

## Verification (reported separately)
- **Unit:** `pytest tests/unit -q` → **1462 passed**, ruff clean.
- **Live Excel integration** (`test_snapshot_restore_hardening.py`, real single-STA Excel, disposable workbook + store under `.ai/tmp`): **5 passed in ~9–10s**, exit 0 — success (one active restored, previous closed, original hash unchanged, one owned Excel), R1 close-failure (`failed`, previous open, restored valid), R2 missing-id rejected, R3 canonical=original + same-snapshot error + different-snapshot ok + no-op valid + rollback replaces content/original untouched, interruption-before-cleanup (restored active, previous open, no auto-replay).
- **Interruption/recovery:** covered by the interruption live case (KeyboardInterrupt bypasses the R1 `except Exception`, leaving restored active + previous open; recovery explicit, no auto-replay) and the transient-close case (`failed`, restored valid).

## Non-blocking advisories (Auditor, LOW)
- `rebind_restored_identity` leaves `record.source_size` at the snapshot size (consistent in normal flow).
- On a failed-close path the rebind overwrites the original-path lock-map entry while the orphaned previous copy still holds it (contracts hold; lock bookkeeping only).

## Status
Implementation complete and independently **ACCEPT**ed. Awaiting final acceptance. No B3/background work; no rewrite of the verified normal restore lifecycle; SCHEMA_VERSION=1; no new ErrorCode. Changes uncommitted (no commit requested).