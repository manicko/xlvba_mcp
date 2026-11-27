# Research B & D — Tech Lead Decisions

- **Revision:** `24201f4`. Research-only; no production changes.
- **Auditor:** `report_12 — Research B_D Audit.md` — all claims CONFIRMED; both investigations sound (two non-substantive citation typos noted, conclusions unaffected).
- **Reports:** `research_B_pivot_background_execution.md`, `research_D_snapshot_restore_lifecycle.md`.

| Topic | Finding (evidence) | Recommendation | Implementation required? |
|---|---|---|---|
| **B — shorter soft timeout (B2)** | Timeout is client-enforced (`com_worker.py` `asyncio.wait_for`); on expiry COM continues and the single-STA stays `DEADLINE_EXCEEDED`/accepting; a shorter timeout gives **no starvation relief** and risks false `expired` on slow supported pivots. | **No** — do not implement B2. | **No.** |
| **B — background job (B3)** | `start_long_op` returns a `job_id` in ~0 s without changing single-STA/FIFO (`job_tools.py`, `job_registry.py`); prevents client `-32001`; adds `job_get`/`job_request_cancel`/interrupted semantics. It does **not** remove FIFO starvation (a blocked COM call still occupies the worker); it makes the wait observable/recoverable. B1 gate (shipped) already rejects external/OLAP in ~0.05 s before `PivotFields`. | **Yes (optional, additive)** — wire `background` into `fields-list`/`field-get-filters` for slow *supported* pivots. Keep B1 as the primary fix. | **Optional, separate WP** (not part of the hang fix). Reconsider only if benchmarks show supported-source enumeration is regularly slow. |
| **D — close obsolete working copy** | WP3 already implements open→validate→close-previous(discard); verified: success leaves exactly 1 open (restored, active), previous closed, original untouched, one owned Excel. Default path violates no invariant. | **Yes — already implemented (WP3); verified.** | **No further change for the mandatory invariants.** |
| **D — restore lifecycle changes** | Additive gaps only (no invariant breach): G1 old-close failure swallowed → 2 open but success returned; G5 restored canonical = snapshot path (same-snapshot repeat → `COM_BUSY`); G6 `workbook_id=None` never closes old; plus G3/G4/G7–G10. | **Recommended additive hardening (R1–R8), prioritized:** R1 surface `previous_close` status (never fail restore on close failure); R2 require/document `workbook_id`; R3 rebind restored canonical to `manifest.source_path`; R4 clearer same-snapshot error; R5 finalize old-id jobs; R6 idempotency key/doc; R7 prune stale maps; R8 init `_state`. | **Optional additive WP** (R1–R3 first). No change to the accepted single-STA/working-copy architecture. |

## Notes
- B and D are kept separate (different subsystems, separate acceptance criteria).
- No B2/B3 or D lifecycle modification is implemented as part of this research.
- The source-type gate (B1) and WP3 lifecycle remain the accepted remediations; this addendum only decides the optional follow-ups.