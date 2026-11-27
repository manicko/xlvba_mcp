# Research D — `snapshot_restore` lifecycle verification

**Role:** Researcher D (research only; no production changes).
**Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `24201f4`.
**Date:** 2026-11-27.
**Scope:** verify the WP3 `snapshot_restore` lifecycle against the mandatory invariants and the 10-case failure/interruption matrix; decide whether further change is required.
**Method:** source trace (file:line) + controlled integration PoC against the real single-STA `ExcelComWorker` with disposable workbooks/snapshots under `.ai/tmp/research_D`. No production file was modified (`git status` clean for `src/`; PoC artifacts are under `.ai/tmp` only).

---

## 1. Current lifecycle (D1)

### 1.1 Call path

| Layer | Entry point | Behaviour |
|---|---|---|
| MCP tool | `api/test_tools.py::snapshot_restore` (119–165) | validates args; `confirm=False` → `CONFIRMATION_REQUIRED` **before** any service call (150–163); delegates to `SnapshotService.restore_snapshot` (165). |
| Application | `application/snapshot_service.py::restore_snapshot` (284–397) | orchestrates open → validate → close-previous (details below). |
| Application | `application/session_service.py::open_workbook` (249–382) | copies the snapshot file to `work/<new-uuid>.xlsm` (1647–1654), opens **that** copy (384–426), registers a new `_SessionRecord` (356–372). |
| Application | `application/session_service.py::close_workbook` (432–555) | resolves the record by `workbook_id` (472), reads the real dirty flag (487), closes the **working copy** by `working_path` (511–523), pops bookkeeping (527–531). |
| Storage | `storage/snapshot_store.py::load` (331–347), `path_for` (349–353) | read-only manifest load; no dependency on any open workbook. |

### 1.2 Exact order in `restore_snapshot`

| Step | Lines | Action |
|---|---|---|
| 0 | 340–341 | `store.load(snapshot_id)` → `SnapshotFailedError` if absent/corrupt; resolve `snapshot_file`. |
| 1 | 346–351 | `session.open_workbook(snapshot_file, mode=EXECUTE, read_only=False)` — **restored copy opened first**, while the previous copy is untouched. |
| 1b | 353–354 | best-effort `info.model_copy(source_path=manifest.source_path)` (returned DTO only). |
| 2 | 359–360 | `_validate_restored_copy(info)` — STA inspect by working path. |
| 2b | 361–369 | on validation failure: best-effort close **the new (invalid) restored copy** (`discard_changes=True`), re-raise. Previous copy is never closed. |
| 3 | 374–381 | validation OK → close **the obsolete previous** copy by `workbook_id` (`save_changes=False, discard_changes=True`), wrapped in `suppress(ExcelVbaMcpError)`. |
| 4 | 383–384 | record `_latest_snapshot[new_id]`, `_known_sources[new_id]`. |
| 5 | 385–396 | audit (`restored_workbook_id`), return `WorkbookInfo`. |

`_validate_restored_copy` (399–448) runs a read-only STA callable (414–429) that locates the workbook by `info.path` via `_find_open_workbook` (1684–1700, case-insensitive `wb.FullName` match) and touches `wb.FullName`; `False` or any exception → `SnapshotFailedError`.

### 1.3 Identity model (stable IDs, not filenames)

| Entity | Old working copy | Restored working copy |
|---|---|---|
| `workbook_id` | `uuid4` at `session_service.py:331` (e.g. `225422ae-…`) | **new** `uuid4` (e.g. `5686fba9-…`) |
| `working_path` | `work/<old-uuid>.xlsm` (1647–1654) | `work/<new-uuid>.xlsm` |
| `_SessionRecord.canonical` | original source path | **snapshot file path** (`…/snapshots/snap-…/CleanProject.xlsm`) — `open_workbook` is called with the snapshot file (snapshot_service 346–351) |
| `_active_by_canonical` key | original source path | snapshot file path |
| Returned `WorkbookInfo.source_path` | original | original (patched at 353–354) |
| `record.info.source_path` (used by `list_workbooks`) | original | **snapshot file path** (unpatched — see gap G5) |
| Registered in same session? | yes (`_records`) | yes (`_records`) — both are briefly registered between steps 1 and 3 |
| Subsequent MCP calls use | old id becomes invalid after step 3 (`WorkbookNotFoundError`) | **new id** returned by restore |
| Pending job / registry ref | jobs are keyed by `job_id`, not `workbook_id` (`job_registry.py`); no per-workbook job index | same |
| Old copy unsaved changes | possible; step 3 discards them (`discard_changes=True`) | none (fresh copy, `Saved=True`) |
| Original source open separately? | **no** — only working copies are ever opened; the original is streamed for hashing only (`_sha256_file`, 115–124) | no |
| Snapshot depends on either open workbook? | **no** — `store.load` + file copy from disk; old close is independent | no |

### 1.4 Session/registration facts

- `open_workbook` registers `_records[workbook_id]` and `_active_by_canonical[canonical] = workbook_id` **after** the COM open (371–372); the per-path lock is held for the workbook lifetime (320–321, released by close 530–531).
- Fast duplicate reject: if the canonical is already mapped, open raises before any COM (309–314, re-checked 324–329).
- `set_snapshot_gate` (234–243) publishes `SnapshotService.has_snapshot`; wired at `server.py:1381`. `has_snapshot` reads the in-memory `_latest_snapshot` only.
- `set_job_registry` (224–232) is wired at `server.py:1290`; recovery finalizes jobs, restore does not.
- There is **no single "active workbook" pointer**. "Active" means the id the caller must pass; after a successful restore the old id is removed from `_records` and the new id is the only valid handle.

### 1.5 Root cause of the historical "extra open workbook"

WP3 **fixed** the original defect (close-both / zero-copies). Verified happy path: exactly one record, one open workbook in Excel, old closed, restored active (evidence §4.1). The only remaining ways to end with two open copies are:

- **G1** old-close failure is swallowed (375) → restore returns success with two copies (evidence §4.3);
- **G6** caller omits `workbook_id` → old is never closed (evidence §4.5).

Both are explicit, non-default paths; the default `workbook_id`-supplied happy path is correct.

---

## 2. Contract verification (D2)

| Invariant | Verdict | Evidence |
|---|---|---|
| Never close the source | **PASS** | source is never opened; close resolves `record.working_path` (session 511–523). |
| Never close the restored copy | **PASS** (except deliberate invalid-copy cleanup) | step 3 targets `workbook_id` and is guarded `workbook_id != info.workbook_id` (374). |
| Never close unrelated workbooks | **PASS** | close is by stable `workbook_id` → that record's `working_path` only. |
| Don't close old before replacement valid + active | **PASS** | close at 374 follows open+register (346–372) and validate (360). |
| Create/validate/activate failure must not close old | **PASS** | open failure propagates at 346; validation failure closes the **new** copy (362–368) and re-raises; registration failure propagates from `open_workbook`. |
| Failure to close old must not invalidate restored | **PASS** | `suppress(ExcelVbaMcpError)` (375); restored stays registered/active. |
| Cleanup by identity, not filename | **PASS** | `workbook_id` → `_SessionRecord.working_path` (stable uuid path). |
| Exactly one copy after success | **PASS on happy path; FAIL on G1/G6** | §4.1 vs §4.3/§4.5. |
| Explicit "verify restored active" post-condition | **IMPLICIT** | no re-read; relies on `open_workbook` returning after 371–372. Gap G2. |

**Atomicity gaps (no mandatory invariant is violated on the default path):**

- **G1 (MEDIUM)** — old-close failure is silent: two records/two Excel workbooks, `snapshot_restore` returns success, no field tells the caller. Recoverable by `excel_close(old_id, discard_changes=true)`.
- **G2 (LOW)** — no explicit post-open assertion that the new id is in `_records` / `_active_by_canonical`.
- **G3 (LOW–MEDIUM)** — failed-validation cleanup is best-effort (362); if that close fails, the invalid restored copy stays registered/open while `SNAPSHOT_FAILED` is raised.
- **G4 (LOW)** — if registration itself fails after the COM open (371–372), the new workbook is open in Excel but has no id (orphan); old remains.
- **G5 (MEDIUM)** — restored `_SessionRecord.canonical` is the **snapshot file**, not `manifest.source_path`. Consequences: (a) `list_workbooks`/`excel_get_status` report the snapshot path while the restore result reports the original; (b) `excel_open(original)` after restore succeeds → duplicate logical copy (evidence §4.6); (c) `excel_save_as`/`snapshot_create` lineage points at the snapshot file; (d) restoring the **same** snapshot twice is blocked because the snapshot canonical is still mapped to the previous restored copy (evidence §4.4).
- **G6 (LOW–MEDIUM)** — `workbook_id=None` never closes the old copy (documented, but silently accumulates copies).
- **G7 (LOW–MEDIUM)** — in-flight background jobs for the old id are not finalized on restore (`mark_interrupted` is recovery-only, `job_registry.py:355–399`); a running macro can fail against the closed old copy.
- **G8 (LOW)** — `_latest_snapshot` / `_known_sources` keep old ids forever; `has_snapshot(old_id)` stays `True` after close.
- **G9 (MEDIUM, no-auto-replay clarity)** — restore has no idempotency key. The system never auto-replays, but a client retry after a timeout creates a second restored copy. The new id is recoverable from the `excel_open` audit record (session 375–381) even when the restore call itself never returned.
- **G10 (LOW, out of scope)** — `SessionService.__init__` never sets `_state`; `_require_accepting_new_work` (1466–1479) reads `self._state.value`. All current paths set `_state` before `_accepting=False`, so it is latent, not live.

---

## 3. Failure / interruption matrix (D3)

"Active id" = the id a subsequent MCP call must use. "Indeterminate" = the call did not return a trustworthy outcome.

| # | Case | Open workbooks after | Active id | Recoverable? | Outcome | Recovery needed |
|---|---|---|---|---|---|---|
| 1 | Snapshot validation fail (`store.load`, 340) | old only | old | yes | **failure** `SNAPSHOT_FAILED`, deterministic | none; fix `snapshot_id` |
| 2 | Restored creation fail (open raises: path/approval/hung/duplicate) | old only (working file cleaned at session 350) | old | yes | **failure** (PATH_NOT_ALLOWED / APPROVAL_DENIED / EXCEL_HUNG / COM_BUSY) | fix cause, retry |
| 3 | Open OK, validation fail (360) | old only (new closed best-effort 362–368) | old | yes | **failure** `SNAPSHOT_FAILED` | none; retry. Residual G3 if cleanup close fails |
| 4 | Session registration fail (371–372) | old + **unregistered** new (orphan in Excel) | old | partial | **indeterminate** | close orphan by path / `excel_recover`; no id to target |
| 5 | Active-switch fail (same statement as 4) | same as 4 | old | partial | **indeterminate** | same as 4 |
| 6 | Old close fail (375 suppressed) | **old + restored** | restored (old id still valid but stale) | yes | **success returned**, invariant broken (G1) | `excel_close(old_id, discard_changes=true)` |
| 7 | Interrupt between activation and cleanup | Excel alive: old + restored; Excel dead: none | none after restart | yes | **indeterminate** (no return) | `excel_recover` / server restart, then reopen; **do not auto-replay**; new id recoverable from `excel_open` audit |
| 8 | Transient COM error on close (`ComBusyError`/`RPC_E_CALL_REJECTED`) | old + restored | restored | yes | **success returned** (same as 6) | retry close after backoff |
| 9 | Old copy dirty | restored only | restored | n/a | **success**; old discarded without save | none; original untouched |
| 10 | Repeated restores | different snapshots: one at a time; same snapshot: previous stays | restored | yes | different: **success**; same: **failure** `COM_BUSY` "already open" (G5) | close previous restored id first, or use another snapshot |

Notes:

- **No auto-replay of an unknown restore:** confirmed. Restore is caller-driven, synchronous, and never retried internally. A client retry is not deduplicated (G9); the audit trail (`snapshot_restore` result 391–395, plus the `excel_open` record 375–381) is the recovery source of truth.
- Case 7 is the only genuinely "indeterminate" state; it is fail-closed on restart (owned-pid reconcile refuses a second Excel; `recover_owned_process` terminates the owned pid and discards working copies, originals untouched).
- Cases 6/8 are "success with residue", not data loss: the old copy is disposable and was already slated for discard.

---

## 4. Controlled integration evidence (D4)

PoC: `.ai/tmp/research_D/poc_snapshot_restore.py` + `poc_supplementary.py` (real `ExcelComWorker`, `SessionMode.EXECUTE`, `ApprovalPolicy(FULL)`, disposable `CleanProject.xlsm` per scenario). Raw lines are `EVIDENCE:` JSON in the run output.

### 4.1 Successful restore (`success`)

Before:
```
owned_pid 22920, state running
records: [225422ae-…  path work/225422ae-….xlsm  source CleanProject.xlsm  execute rw]
active_by_canonical: {CleanProject.xlsm: 225422ae-…}
open_in_excel: [work/225422ae-….xlsm  Saved=false]   # dirty after mutation
original_sha256 2a7c3c0b…  snapshot_sha256 fbca54c3…
```
After:
```
owned_pid 22920 (unchanged), state running
records: [5686fba9-…  path work/5686fba9-….xlsm  source snapshots/snap-…/CleanProject.xlsm]
active_by_canonical: {snapshots/snap-…/CleanProject.xlsm: 5686fba9-…}
open_in_excel: [work/5686fba9-….xlsm  Saved=true]    # exactly one
original_sha256 2a7c3c0b… (unchanged)  snapshot_sha256 fbca54c3… (unchanged)
restored: id 5686fba9-…, path work/5686fba9-….xlsm, source_path CleanProject.xlsm, execute, rw
restored_a1 = null (baseline; mutation discarded)
```
→ restored open + active, previous closed, source byte-identical, snapshot immutable, one owned Excel, PID stable.

### 4.2 Validation failure (`validation_fail`)

`_validate_restored_copy` forced to raise. After: records = `[old]` only; `open_in_excel` = 1 (old); `work_dir` contains only the old file; error surfaced (`SNAPSHOT_FAILED` in the real path; the PoC's generic `ExcelVbaMcpError` reported `COM_BUSY`). Old untouched, invalid restored copy removed.

### 4.3 Old-close failure (`old_close_fail`)

`session.close_workbook(old_id)` forced to raise. After:
```
records: [old 726addd1-…, new 499e7f00-…]
active_by_canonical: {CleanProject.xlsm: old, snapshots/snap-…: new}
open_in_excel: [old, new]   # TWO
restore returned success (no error)
```
→ confirms G1: restored active, old still open, caller not informed.

### 4.4 Repeated restores

- Same snapshot twice: second call fails at `open_workbook` with `COM_BUSY` "Workbook is already open for 'CleanProject.xlsm'" (fast reject 309–314; canonical is the snapshot file). Previous restored copy remains open/active. → G5.
- Different snapshots (`diff_snapshots`): chain `ddcabde8 → f552d7a0 → d9a64151`, records = `[d9a64151]` only, one open in Excel, `active_by_canonical` keyed by the latest snapshot path. → repeated restore works when snapshots differ.

### 4.5 `workbook_id=None` (`no_workbook_id`)

Records = `[old 7619e905…, new d61c64d9…]`, `open_in_excel` = 2, both canonical keys present. → G6.

### 4.6 Reopen original after restore (`reopen_original`)

`excel_open(original)` after a restore succeeds: records = `[restored 9b02ffea…, duplicate 4c8fabbb…]`, two open workbooks, two canonical keys (snapshot path + original). → G5 duplicate-logical-copy consequence.

### 4.7 Process hygiene

- The pre-existing user Excel (PID 26964, start 14:53) was **not** touched and is still the only Excel after the runs.
- Owned PIDs 22920 / 20508 / 21116 were reaped by `worker.shutdown()` (all "gone").
- Exactly one owned Excel per scenario; restore never restarts Excel (PID stable across 4.1).
- No unrelated workbook was opened or closed in any scenario.

---

## 5. Recommendation

**No further change is required for the mandatory invariants.** WP3's order (open → validate → register/switch → close previous with discard → verify-by-construction) satisfies every must-have: source/restored/unrelated are never closed, the old copy is never closed before the replacement is valid and registered, create/validate/activate failures leave the old copy open, and an old-close failure does not invalidate the restored copy. The historical "extra open workbook" defect is fixed on the default path.

Recommended **additive / documentation** changes (non-blocking, no contract break):

1. **R1 (addresses G1/G8, highest value):** surface the previous-close result additively on `WorkbookInfo`/a restore result (`previous_workbook_id`, `previous_close: "closed" | "failed" | "not_requested"`, `previous_close_error`). Do **not** turn a close failure into a restore failure — that would violate the "close-old failure must not invalidate restored" invariant.
2. **R2 (G6):** document that `workbook_id` must be supplied for the exactly-one guarantee; optionally default it to the sole open workbook when exactly one exists.
3. **R3 (G5):** rebind the restored record's canonical to `manifest.source_path` (or add a distinct `source_canonical` field) so `list_workbooks`/`save_as`/lineage agree with the restore result, and so the snapshot file is never a session canonical. This also unblocks same-snapshot repeated restore. If rebinding is out of scope, document the current semantics explicitly.
4. **R4 (G5 case 10):** document that restoring the same snapshot twice requires closing the previous restored copy first; optionally return a clearer `SNAPSHOT_FAILED`/`STALE_TARGET` instead of the generic `COM_BUSY` "already open".
5. **R5 (G7):** finalize in-flight jobs for the old id on restore (reuse `JobRegistry.mark_interrupted` semantics) or document that restore does not affect running jobs.
6. **R6 (G9):** add an idempotency key to `snapshot_restore` if client retries must be deduplicated; otherwise document "restore is not idempotent; a retry creates a new working copy".
7. **R7 (G8):** prune `_latest_snapshot` / `_known_sources` entries when a workbook id is closed.
8. **R8 (G10):** initialize `_state` in `SessionService.__init__` (defensive).

**Contract changes required: none.** R1–R8 are additive fields, documentation, or defensive cleanup; they do not alter the verified lifecycle order.

---

## Appendix A — evidence artifacts

- PoC scripts: `.ai/tmp/research_D/poc_snapshot_restore.py`, `.ai/tmp/research_D/poc_supplementary.py`.
- Per-scenario dirs: `.ai/tmp/research_D/{success,validation_fail,old_close_fail,repeated,diff_snapshots,no_workbook_id,reopen_original}` (disposable workbooks, snapshots, `work/`, `audit.jsonl`).
- Run command: `.\.venv\Scripts\python.exe .ai\tmp\research_D\poc_snapshot_restore.py` (and `poc_supplementary.py`).

## Appendix B — limitations

- The validation-failure PoC injected a generic `ExcelVbaMcpError` (reported `COM_BUSY`); the production `_validate_restored_copy` raises `SnapshotFailedError` (`SNAPSHOT_FAILED`). Lifecycle behaviour is identical.
- Interruption (case 7) was analysed, not executed (it requires killing the host process); the restart/reconcile path is covered by existing recovery code (`session_service.py:1059–1411`) and was not re-run here.
- "Active id" is a caller-facing notion; the service has no single active pointer (see §1.4).
- Integration used `ApprovalPolicy(FULL)`; in `RESTRICTED` mode case 2 additionally fails with `APPROVAL_DENIED` before any open (old untouched).