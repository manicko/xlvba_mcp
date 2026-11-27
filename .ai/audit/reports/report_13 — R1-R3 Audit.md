# report_13 — Independent Audit: Snapshot Restore Hardening R1–R3

- **Auditor role:** Independent architecture auditor (read-only; no source modified).
- **Repo / revision:** `C:\py_exp\xlvba_mcp`, HEAD `24201f4` + R1–R3 in working tree.
- **Scope:** R1 (`previous_close` surface, never fail restore on close failure), R2 (require `workbook_id`), R3 (rebind identity to original + deterministic same-snapshot error).
- **Method:** Source review against the plan (`plan_06`, `plan_06a`) + unit and live-Excel test execution. Evidence is file:line. Tests were run, not trusted from the plan.
- **Date:** 2026-11-27.

## Verdict summary

| # | Contract | Result |
|---|----------|--------|
| 1 | R1 `_close_previous_copy` broad-catch → `failed`, never re-raises; restore success; restored active; `None`→`not_requested` | **PASS** |
| 2 | R2 handler rejects missing/empty `workbook_id` with `INVALID_RANGE`, `isRetryable=false`, before any action | **PASS** |
| 3 | R3 `rebind_restored_identity` sets canonical + `info.source_path` + `source_sha256`; migrates active/lock maps; same-snapshot guard before open; snapshot immutable; original hash unchanged | **PASS** |
| 4 | WP3 order preserved; no B3/background; no new worker; `SCHEMA_VERSION=1`; no new `ErrorCode` | **PASS** |
| 5 | Tests present and green (unit + live) | **PASS** |
| 6 | Interruption between activation and cleanup → restored active + previous open, no auto-replay; transient close COM → `failed`, restored valid | **PASS** |

**Final recommendation: ACCEPT** (2 advisory LOW observations below; neither blocks acceptance).

---

## Detailed evidence

### 1 — R1: `previous_close` never fails the restore — PASS

- `_close_previous_copy` (`src/excel_vba_mcp/application/snapshot_service.py:466-498`):
  - `workbook_id is None or workbook_id == restored_id` → returns `("not_requested", None)` (`:487-488`).
  - Clean close → `("closed", None)` (`:498`).
  - **Broad** `except Exception as exc` → returns `("failed", _safe_error_text(exc))` (`:496-497`). It does **not** re-raise.
- `_safe_error_text` (`snapshot_service.py:729-738`) uses the domain `message` (never the context/paths), falls back to `str(exc)`/class name, caps at 500 chars — safe, no secrets.
- Restore still succeeds and the restored session stays valid: after `_close_previous_copy` returns, `restore_snapshot` unconditionally rebinds (`:437-441`), records the audit, and returns `SnapshotRestoreResult(...)` (`:460-464`). The close result is only *reported*, never used to branch the outcome.
- `workbook_id=None` → `not_requested`: `_close_previous_copy` short-circuits (`:487`) and the previous copy is left open.
- DTO: `SnapshotRestoreResult.previous_close: Literal["closed","failed","not_requested"]` + `previous_close_error: str|None` (`src/excel_vba_mcp/domain/models.py:565-598`).
- Surfaced to the client unchanged: the API handler returns the service result directly (`src/excel_vba_mcp/api/test_tools.py:181`).
- Unit proof: `tests/unit/test_snapshot_service.py:492` (`..._failed_surfaces_but_succeeds`) and `:525` (`..._not_requested`).

### 2 — R2: `workbook_id` required before any action — PASS

- Handler `snapshot_restore` (`src/excel_vba_mcp/api/test_tools.py:119-181`):
  - Order: `_require_non_empty(snapshot_id)` (`:155`) → **missing/empty `workbook_id` rejected** (`:157-167`) → `_require_non_empty(workbook_id)` (`:168`) → `confirm` gate (`:169-179`) → only then `service.restore_snapshot(...)` (`:180-181`).
  - Missing/empty id raises `ExcelVbaMcpError(code=ErrorCode.INVALID_RANGE, context={tool, field:"workbook_id", reason:"workbook_id_required"})` (`:158-167`) — before any open/close.
- `isRetryable=false`: `INVALID_RANGE` is not in `_RETRYABLE_CODES` (only `COM_BUSY`/`RPC_E_CALL_REJECTED`) → `retryable_for` returns `False` (`src/excel_vba_mcp/domain/errors.py:179-181, 294-308`); category = `validation` (`errors.py:171`); emitted in `to_dict()` as `isRetryable` (`errors.py:381`).
- Server description/doc updated (`src/excel_vba_mcp/server.py:303-317`): "current workbook_id (required) … a missing workbook_id is rejected with INVALID_RANGE (isRetryable=false)".
- Unit proof: `tests/unit/test_confirm_flags.py:127` asserts `INVALID_RANGE`, `field=="workbook_id"`, `isRetryable is False`, and `fake.restore_calls == []` (delegate not called).

### 3 — R3: identity rebind + same-snapshot guard + immutability — PASS

- `rebind_restored_identity` (`src/excel_vba_mcp/application/session_service.py:245-285`):
  - `record.canonical = canonical` (`:277`), `record.source_sha256 = source_sha256` (`:278`), `record.info = record.info.model_copy(update={"source_path": canonical})` (`:279`).
  - Migrates per-path maps when the canonical changed: pops old `_active_by_canonical`/`_locks`, sets them under the rebound canonical with the same `record.lock` (`:280-285`).
- Called with the ORIGINAL + baseline hash: `rebind_restored_identity(info.workbook_id, canonical=manifest.source_path, source_sha256=manifest.sha256)` (`snapshot_service.py:437-441`), then `_restored_snapshot_by_canonical[manifest.source_path] = snapshot_id` (`:442`).
- Working copy stays `work/<newid>.ext`: `open_workbook` copies the snapshot file via `_materialize_working_copy` → `shutil.copy2(canonical, work/<id><suffix>)` (`session_service.py:1689-1695`). The snapshot file is **copied, never moved/mutated**.
- Same-snapshot guard **before open**: `restore_snapshot` computes `previous_canonical = _canonical_for(workbook_id)` and, if `_restored_snapshot_by_canonical.get(previous_canonical) == snapshot_id`, raises `INVALID_RANGE` with `context.reason="same_snapshot"` (`snapshot_service.py:373-392`) — this block precedes the `open_workbook` call (`:397`). No duplicate working copy is created.
- Different snapshot allowed: the guard only triggers on id equality; a different id falls through to open (`:397`).
- Original never modified / snapshot immutable: open copies; unit `test_restore_closes_without_save_and_opens_baseline` asserts the stored snapshot hash is unchanged (`tests/unit/test_snapshot_service.py:370`); `test_restore_noop_...` and `test_restore_rollback_...` assert the original bytes are unchanged (`:647, :676`).
- Unit proof: `tests/unit/test_snapshot_service.py:547` (canonical rebind), `:576` (same-snapshot rejected, `isRetryable False`), `:604` (different allowed), `:628` (no-op), `:650` (rollback).

### 4 — Lifecycle / invariants — PASS

- WP3 order preserved in `restore_snapshot` (`snapshot_service.py`): load manifest (`:370`) → same-snapshot guard (`:373-392`) → **open** restored (`:397`) → **validate** (`:410-420`) → **close-previous** (`:427`) → **rebind** (`:437`) → audit (`:446`). The R3 guard is additive and sits before open; the WP3 sequence (open→validate→close-previous) is unchanged.
- No B3 / no background: `restore_snapshot` is a plain `async` method; no `start_long_op`/`background` path. (Background exists only for `test`/`pivot` per `server.py`, untouched here.)
- No new worker: restore reuses `self._session.open_workbook` and the existing `self._worker`; `rebind_restored_identity` is pure bookkeeping (no COM, no new instance).
- `SCHEMA_VERSION = 1` unchanged: `src/excel_vba_mcp/domain/models.py:65`; `git diff 24201f4 -- models.py` shows no `SCHEMA_VERSION` change.
- No new `ErrorCode`: `src/excel_vba_mcp/domain/errors.py` is **not** in the diff (`git diff --name-only` empty). Same-snapshot reuses `INVALID_RANGE` + `reason="same_snapshot"`.
- Diff is additive: 5 src files, +261/−35 (the deletions are the replaced restore body/docstrings).

### 5 — Tests — PASS (executed)

Environment: `.venv` Python 3.14.7, pytest 9.2.0, Windows (live Excel available).

Unit (`pytest tests/unit`): **1462 collected, exit 0 — all passed.**
- R1: `test_snapshot_service.py::test_restore_previous_close_failed_surfaces_but_succeeds`, `::test_restore_without_workbook_id_reports_not_requested`.
- R2: `test_confirm_flags.py::test_snapshot_restore_missing_workbook_id_rejected` (+ `..._false_confirm_rejects_before_action`).
- R3: `test_snapshot_service.py::test_restore_rebinds_canonical_to_original`, `::test_restore_same_snapshot_repeat_rejected`, `::test_restore_different_snapshot_allowed`, `::test_restore_noop_yields_valid_copy_bound_to_original`, `::test_restore_rollback_replaces_working_content`, plus WP3 `::test_restore_validation_failure_keeps_previous_open`.

Live (`pytest tests/integration_excel/test_snapshot_restore_hardening.py -p no:faulthandler -rN`): **5 passed in 10.28s, exit 0.**
- `test_restore_success_single_active_copy`, `test_restore_close_failure_reports_failed`, `test_restore_missing_workbook_id_rejected`, `test_restore_same_snapshot_and_different`, `test_restore_interruption_before_cleanup`.

### 6 — Interruption / recovery — PASS

- Transient close COM failure → `failed`, restored valid: a raised `ExcelVbaMcpError` is an `Exception`, caught by `_close_previous_copy` (`snapshot_service.py:496`) → `previous_close="failed"`; restore still returns success and the restored copy is active. Live proof: `test_restore_close_failure_reports_failed` (asserts `failed`, safe text, restored active/valid, previous still open).
- Interruption between activation and cleanup → restored active + previous open, no auto-replay: the live test raises `KeyboardInterrupt` (a `BaseException`, **not** caught by `except Exception`) at the close-previous step (`test_snapshot_restore_hardening.py:396-405`). It propagates out of `restore_snapshot` (no replay/retry), leaving the restored copy opened/validated/registered and the previous copy still open; rebind has not yet run (indeterminate state, recovery explicit). The test then closes the previous copy explicitly and keeps the restored one as the single active session (`:430-436`).

---

## Advisory observations (LOW — do not block acceptance)

- **[BEST-PRACTICE] `rebind_restored_identity` does not update `record.source_size`.** It sets `source_sha256` but leaves `source_size` at the snapshot file's size (`session_service.py:275-279`). In the normal flow the snapshot is a byte copy of the original, so sizes match and `save_as_workbook`'s size+hash re-verify stays consistent. Effort: trivial. Priority: recommended (set `source_size` from the original at rebind for symmetry). Why: avoids a latent size/hash mismatch if a future caller supplies a baseline whose size differs from the current original.
- **[BEST-PRACTICE] Failed-close + rebind lock-map edge.** When the previous close fails, the previous record stays in `_records` and still holds the original-path lock, but `rebind_restored_identity` then overwrites `_active_by_canonical[original]` and `_locks[original]` with the restored id/lock (`session_service.py:280-285`). Contracts still hold (restored active; previous open and closeable by id via `_require_record`), but the original-path lock bookkeeping becomes internally inconsistent until the orphaned previous copy is closed. Effort: small. Priority: recommended (document or reconcile the orphaned lock on the failed-close path). Why: prevents a confusing lock leak in the rare transient-close-failure case; not a correctness defect for R1–R3.

## Recommendation

**ACCEPT.** All six verification items PASS against source and executed tests. R1/R2/R3 are additive, preserve the WP3 lifecycle, keep `SCHEMA_VERSION=1`, add no new `ErrorCode`, and introduce no background/worker. The two advisory LOW items are optional hardening and do not affect the contracts.