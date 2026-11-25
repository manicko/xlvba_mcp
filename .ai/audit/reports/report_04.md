# report_04 — Consolidated Priority Decisions (Blocks A–F)

- **Status:** Recommended, subject to PoC
- **Date:** 2026-11-25
- **Scope:** Advisory selection of one priority solution per audit block (A–F) for `xlvba_mcp`, with rationale, rejected alternatives, preserved non-negotiables, and a cross-block roadmap. No production code is changed by this report.
- **Method:** Auditor verification (report_03 evidence re-checked against current source) + validator selection. Validator re-opened the decision-critical anchors read-only; all cited `file:line` anchors were confirmed accurate, including the corrected macro_service citation (`recover_session` invoked at `macro_service.py:960` inside `_recover_after_hung`, 949–966).

## Executive decision

Adopt a single coherent program built on one principle — **slow ≠ hung** — and one foundation. **Block E** (timeout-semantics split + liveness watchdog) is the keystone; **Block A** (hung-state recovery) and **Block F** (job-seam) are sequenced on top of it. **Block D** ships first as low-risk cleanup that makes A/E observability truthful. **Block B** (pivot cost guard) ships in parallel with E. **Block C** is independent documentation plus one additive DTO field. **Block F** is deferred until a real >90 s disconnect need is observed. Every choice preserves the non-negotiables: single-STA (ARCH-009), originals-never-modified, deny-by-default approval, pid-scoped kill only, external/OLAP/Data-Model pivots stay `unsupported_source`, never materialize pivot output, never `Workbook.RefreshAll`, and `SCHEMA_VERSION` stays 1 (all new DTO fields additive/defaulted, `models.py:62-68`).

## Priority decisions

| Block | Chosen solution | Priority rank | Effort | Depends on |
|-------|-----------------|---------------|--------|------------|
| D | **D1** — reorder `_clear_active` before `_complete_result`; epoch-scoped notes; add `list_workbooks_reconciled` (leave `_sync` untouched) | 1 (ship first) | S | none |
| E | **E1** (target) with **E2** as phased fallback — split per-op-class deadline (queue-wait never poisons) + free-thread liveness watchdog; only CONFIRMED_HUNG reaps | 2 (foundation) | M | none |
| A | **A1** — SLOW/SUSPECTED/CONFIRMED_HUNG state machine + off-thread pid-scoped reap + thin `excel_recover` last-resort tool | 3 | M | E |
| B | **B1** — per-sheet cost guard returning bounded `too_large` (offending sheet named) + sheet-hint early-exit; steer to `range_read_values` | 4 (parallel w/ E) | M | E (soft — per-class budget) |
| C | **C1** (+ C2 docs folded in) — log resolved roots at startup; additive `original_name` on `WorkbookInfo`; document inspect-is-writable and double-I/O; defer inspect⇒read_only contract change | 5 | S | none |
| F | **Defer** — no job-seam now; rely on E + progress notifications; build internal start/status/cancel seam only if a real >90 s op disconnect is observed | 6 (last/deferred) | S now / L later | E → A |

---

## Block A — Worker hang / single-STA head-of-line blocking / no in-tool recovery

**Selected solution: A1.** Introduce an explicit `SLOW → SUSPECTED → CONFIRMED_HUNG` state machine, an off-thread pid-scoped reap (never on the STA, never wildcard), and a thin `excel_recover` last-resort tool that delegates to the existing `recover_session` (`session_service.py:552`) / worker `reset` path. A1 is chosen because the defect is not merely "no restart button": a single FIFO (`com_worker.py:267`) + single STA thread (`:300`) means one wedged native call blocks every subsequent op, and `close_workbook` currently enqueues through the same FIFO (`session_service.py:381-386`) rather than the internal drain (`is_close_drain` `:145-147`, `_make_internal_item` `:715-726`). A1 makes the head-of-line state observable and recoverable without weakening single-STA; the `excel_recover` tool is the explicit, deny-by-default-gated operator escape hatch.

**Rejected alternatives:**
- **A2 (tool only):** necessary but insufficient — exposes recovery without fixing observability or the head-of-line state model; the tool would fire blind. A2 is absorbed as the tool facet of A1.
- **Per-op workers / worker pool:** violates ARCH-009 single-STA.
- **Auto-kill-on-timeout:** a slow op is not a dead process; killing on deadline corrupts the "slow ≠ hung" principle.
- **Wildcard `taskkill /IM`:** kills unrelated Excel instances; violates pid-scoped-kill-only.

**Non-negotiables preserved:** single-STA (reap is off-thread, pid-scoped); originals-never-modified (recovery rebuilds the hidden working copy only); deny-by-default approval on `excel_recover`; pid-scoped kill only.

**Verification/PoC:** see §PoC A.

**Anchors:** `com_worker.py:267,300,391,395-396,401-409`; `session_service.py:145-147,381-386,552,715-726`; `macro_service.py:960`; `server.py:1004-1224+` (no restart tool today).

---

## Block B — Pivot enumeration/metadata perf + no bounded "too large" reads

**Selected solution: B1.** Add a per-sheet cost guard that returns a bounded `too_large` result naming the offending sheet, plus a sheet-hint early-exit so callers can target one sheet, and steer large inventories to `range_read_values`. B1 is chosen because it deterministically bounds the actual hazard (the 566k-row sheet) and returns an actionable, structured result instead of burning the 60 s budget (`pivot_tables_service.py:98`) and timing out. The full `ws.PivotTables()` sweep (`:36,158-161`) and per-call re-enumeration in `_find_pivot_table` (`:906-916`) are exactly what the guard short-circuits.

**Rejected alternatives:**
- **B2 (larger per-class budget B):** masks the symptom; a bigger ceiling still leaves the 566k-row sheet unbounded and merely delays the timeout. The per-class budget is still adopted *from Block E* as the correct home for the ceiling — but not as the fix for B.
- **File-size trigger:** unreliable — the 5.5 MB `report_w` carries the 566k-row hazard, so size does not predict cost.
- **Auto-kill on timeout:** violates "slow ≠ hung".

**Non-negotiables preserved:** external/OLAP/Data-Model pivots remain `unsupported_source`; pivot output is never materialized; `too_large` is a bounded read result, not a license to dump cells.

**Verification/PoC:** see §PoC B.

**Anchors:** `pivot_tables_service.py:36,98,158-161,850-869,906-916`.

---

## Block C — Path allowlist ergonomics, GUID indirection, double I/O, inspect-is-writable

**Selected solution: C1 (with C2 documentation folded in).** (1) Log the resolved allowlist roots at startup so the CWD/temp fallback (`server.py:765-770`) is never silent; (2) add an additive, defaulted `original_name` field to `WorkbookInfo` so `describe` is not GUID-only (`workbook_adapter.py:97-98,108-119`; `session_service.py:246`); (3) document that a writable inspect open is a mutation (`session_service.py:211-216`) and that `excel_open` defaults to `mode=execute, read_only=False` (`workbook_tools.py:84-88`) — the inspect-is-writable footgun; (4) document/budget the `shutil.copy2` double I/O (`session_service.py:867-874`) rather than change copy semantics. The inspect⇒read_only contract change is **deferred** (it is a behavior change, not a doc fix). C2 is not a separate track — its "document the double I/O" action is executed inside C1.

**Rejected alternatives:**
- **Hardlink/move originals:** breaks originals-never-modified.
- **Runtime root add:** weakens deny-by-default.
- **Silently flipping default `read_only=True`:** contract change; deferred, not smuggled into a docs fix.

**Non-negotiables preserved:** originals-never-modified; deny-by-default approval; `SCHEMA_VERSION` stays 1 (`original_name` is additive/defaulted).

**Verification/PoC:** see §PoC C.

**Anchors:** `server.py:765-770`; `session_service.py:211-216,246,867-874`; `workbook_adapter.py:97-98,108-119`; `workbook_tools.py:84-88`; `models.py:62-68`.

---

## Block D — Diagnostics / bookkeeping hygiene

**Selected solution: D1.** (1) Reorder so `_clear_active` runs **before** `_complete_result` (`com_worker.py:632,640`) — currently `_complete_result` at `:632` precedes `finally: _clear_active()` at `:640`, a truthfulness gap; (2) epoch-scope `_note` output so diagnostics show current-epoch notes only (the rolling-50 buffer at `:854-859` is retained, but health surfaces the current epoch, not a verbatim dump — `health_tools.py:139,166`); (3) add `list_workbooks_reconciled` as a **new** tool that performs the COM reconcile, leaving `list_workbooks_sync` bookkeeping-only (`session_service.py:527-541`). D1 is chosen because it is low-risk, makes A/E observability truthful, and introduces no contract break. D2 (make reconciled the default) is deferred as a contract change.

**Rejected alternatives:**
- **D2 (reconciled as default):** contract change; defer.
- **Clearing notes globally each call:** loses history needed to diagnose A/E.
- **Making `list_workbooks_sync` reconcile:** forces COM into the sync non-COM path.

**Non-negotiables preserved:** single-STA (reconcile stays on the async/worker path); no contract break (`SCHEMA_VERSION` stays 1; new tool + new field additive).

**Verification/PoC:** see §PoC D.

**Anchors:** `com_worker.py:632,640,854-859`; `health_tools.py:139,165,166`; `session_service.py:527-541`.

---

## Block E — Timeout semantics conflation (deadline vs liveness) — cross-cutting root cause

**Selected solution: E1 (target), with E2 as a phased fallback.** E1 splits the single deadline (`com_worker.py:381`) into a **per-op-class deadline B** where queue-wait must **not** set `NEEDS_RECOVERY` (only an exec-breach marks `SLOW`), plus a **free-thread liveness watchdog C** that polls `_pid_alive` (`app_adapter.py:434-467`) and promotes `SLOW → SUSPECTED → CONFIRMED_HUNG`; **only CONFIRMED_HUNG** triggers the off-thread pid-scoped reap + `reset`. Env knobs: `EXCEL_VBA_MCP_OPEN_TIMEOUT_S=300`, `EXCEL_VBA_MCP_LIVENESS_CEILING_S=600`. This is the keystone because the same `timeout` currently bounds both queue-wait (`:391`) and exec deadline (`:602`), a `TimeoutError` unconditionally poisons to `NEEDS_RECOVERY` (`:395-396`), `_maybe_auto_recover` is dead-pid-only (`session_service.py:729`, `_pid_alive` semantics), and `reset()` queues `_do_reset` on the same FIFO (`:489-521`, body `:767-784`) so it cannot preempt. **E2 is the phase-1 fallback** if the watchdog proves too invasive for the first cut: decouple queue-wait from exec deadline (queue-wait returns `busy`/`retry_after_s` and never poisons) and raise the open budget via the env knob, deferring the watchdog. E1 is stated as the target; E2 de-risks delivery. The uncancellable native call is correctly **not** cancelled (`:36-40,392-409`) — E1 preserves that.

**Rejected alternatives:**
- **Auto-kill on timeout:** kills a slow-but-live process; violates "slow ≠ hung".
- **File-size trigger:** unreliable predictor of cost.
- **Per-op workers:** violates ARCH-009.
- **Hard exec limit as a generic replacement for the soft threshold:** loses the slow-vs-hung distinction that the liveness watchdog exists to make.

**Non-negotiables preserved:** single-STA (watchdog is free-thread, read-only `_pid_alive`; reap is off-thread, pid-scoped); originals-never-modified; deny-by-default approval; pid-scoped kill only.

**Verification/PoC:** see §PoC E.

**Anchors:** `com_worker.py:36-40,381,391,395-396,489-521,602,767-784`; `session_service.py:68,329-334,729`; `app_adapter.py:434-467`.

---

## Block F — Long-running op handling / Operation Manager job-seam

**Selected solution: Defer.** No `operation_*` tools are built now. open/save/test_suite_run/macro_run remain synchronous through `worker.submit` with the (post-E) per-class budget. Rely on Block E + progress notifications. Build an internal start/status/cancel seam — shaped 1:1 to the Tasks extension (SEP-2663, SDK-absent) — **only** if a genuine >90 s operation disconnect is observed in practice, and only after E → A. The job-seam is not the root-cause fix; the root cause is E. A blanket async surface is rejected because small-model polling of many job handles fails in practice.

**Rejected alternatives:**
- **Blanket async for all tools:** small-model polling failure mode.
- **Depending on the experimental Tasks extension today:** SDK-absent; premature.
- **Multiple workers / VM / containers:** violates ARCH-009 and operational simplicity.

**Non-negotiables preserved:** single-STA; originals-never-modified; deny-by-default; pid-scoped kill only. (Deferred work must re-verify all of these at build time.)

**Verification/PoC:** see §PoC F.

**Anchors:** `server.py` tool surface (no `operation_*`); `com_worker.py` single `submit` timeout path (post-E per-class).

---

## Sequencing / roadmap

Recommended implementation order (S/M/L effort noted):

1. **D — D1 (S).** Ship first. Pure cleanup; no contract break; makes A/E observability truthful (correct `_clear_active`/`_complete_result` order, epoch-scoped notes, reconciled listing). Low risk, unblocks honest diagnostics for everything after.
2. **E — E1 (M), E2 (S) as fallback.** Foundation. Split queue-wait from exec deadline, add liveness watchdog, gate reap on CONFIRMED_HUNG. Everything in A/F depends on this. If E1's watchdog is too large for the first PR, land E2 first (decouple + open budget) and schedule the watchdog as a fast follow.
3. **A — A1 (M).** On top of E. Adds the explicit hung-state machine surface and the `excel_recover` last-resort tool; consumes E's SLOW/SUSPECTED/CONFIRMED_HUNG states and off-thread reap.
4. **B — B1 (M), parallel with E.** Mostly independent; benefits from E's per-class budget. Can be developed and merged alongside E without ordering conflict.
5. **C — C1 (S), independent.** Docs + startup root log + additive `original_name`; schedule anywhere; no dependency.
6. **F — Defer (S now / L later).** Last and conditional. Revisit only after a real >90 s disconnect need is observed, and only after E → A.

Justification: the dependency spine is **D → E → A → F**, with **B** parallel to E and **C** free-floating. D first because it is the cheapest truthful-observability win and de-risks validating E/A. E before A because A's state machine and reap are defined in terms of E's states. F last because it is the largest, least-certain, and explicitly not the root-cause fix.

---

## PoC / acceptance criteria

Concrete, per-block (complements report_02 §8 and report_03 §9):

- **A:** Simulate a wedged native call (blocking stub). Confirm: (1) worker enters SLOW, then SUSPECTED, then CONFIRMED_HUNG via the watchdog; (2) a second op enqueued behind it is rejected fast with `busy`/`retry_after_s` (post-E) rather than hanging; (3) `excel_recover` requires approval, reaps only the owned pid (verify no other Excel pid touched), and returns the session to RUNNING; (4) originals file hash unchanged before/after.
- **B:** Open the 5.5 MB `report_w` (566k-row sheet). Confirm: inventory returns bounded `too_large` naming the sheet within budget (no 60 s burn); sheet-hint early-exit returns only the targeted sheet; an external/OLAP/Data-Model pivot still returns `unsupported_source`; no pivot output cells are materialized.
- **C:** Start with no roots env; confirm startup log shows the resolved CWD/temp roots; `describe` returns `original_name` (additive) alongside the GUID; docs state inspect-is-writable and default `mode=execute, read_only=False`; `SCHEMA_VERSION` still 1.
- **D:** Force a timeout; confirm health shows `_clear_active` already run (active cleared) and notes are current-epoch only; `list_workbooks_reconciled` performs COM reconcile while `list_workbooks_sync` stays bookkeeping-only.
- **E:** With a slow-but-alive op, confirm queue-wait timeout returns `busy`/`retry_after_s` and does **not** set NEEDS_RECOVERY; an exec breach marks SLOW (not hung); only a genuinely dead pid (or CONFIRMED_HUNG) triggers reap; env knobs `EXCEL_VBA_MCP_OPEN_TIMEOUT_S=300` / `EXCEL_VBA_MCP_LIVENESS_CEILING_S=600` are honored; the uncancellable native call is never cancelled.
- **F:** No new tools in this cycle; confirm open/save/test_suite_run/macro_run still succeed under the post-E per-class budget and surface progress notifications. Re-evaluate job-seam only on a documented >90 s disconnect.

All PoCs run on copies only; originals never modified; single-STA maintained; pid-scoped kill verified.

---

## Risks & open questions

- **E1 watchdog complexity (M):** the free-thread liveness poller is the largest new surface. Mitigation: E2 phased fallback; watchdog behind env knob, default conservative.
- **Threshold tuning:** `OPEN_TIMEOUT_S=300` / `LIVENESS_CEILING_S=600` are starting points; real workloads may need per-class tuning. Open question: should the per-class budget B be operator-configurable per op-class, or fixed?
- **B cost-guard heuristic:** the per-sheet cost model must not false-positive on legitimately large-but-bounded sheets. Open question: base the guard on row/used-range count, pivot count, or a measured time sample?
- **C contract deferral:** inspect-is-writable remains a footgun until the read_only default change is scheduled. Tracked as deferred, not forgotten.
- **F uncertainty:** deferral assumes E + progress notifications suffice for real workloads; if a >90 s disconnect appears, the job-seam (L) re-enters scope and must re-verify every non-negotiable.
- **D2 deferral:** reconciled listing is not the default; consumers must opt in until a future contract bump.
- **Stale-citation hygiene:** report_03's macro_service citation is corrected here (`macro_service.py:960`); no other report_03 evidence was found stale in the validator's spot-checks.