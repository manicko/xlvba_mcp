# Report 11b — Stage 1 Technical Auditor (Save As `.bak` fix)

- **Role:** Independent Technical Auditor (Stage 1).
- **Date:** 2026-11-26
- **Repo:** `C:\py_exp\xlvba_mcp`
- **Candidate:** `ef3e6cb` + uncommitted Save As `.bak` fix (working tree).
- **Scope:** AUDIT ONLY. No source files were modified. The only file written is this report.
- **Verdict:** **SAFE TO SHIP** (all mandatory checks PASS; three LOW advisories, none blocking).

---

## 1. PASS/FAIL table

| # | Check | Result | Primary evidence |
|---|---|---|---|
| 1a | Single hidden Excel + single-STA FIFO preserved | **PASS** | `com_worker.py` not in diff; flush still `self._worker.submit(...)` `session_service.py:765` |
| 1b | Job Object lifecycle preserved | **PASS** | No worker/job/lifecycle file in `git diff --name-only` |
| 1c | Working-copy model preserved | **PASS** | `_materialize_working_copy`/`open_workbook` unchanged; Save As refuses working copy `session_service.py:630` |
| 1d | Originals protected (overwrite_source + hash match) | **PASS** | `session_service.py:657-676` (re-hash + `source_modified_since_open`/`source_replaced`) |
| 1e | `INTERRUPTED_UNKNOWN` still first-class | **PASS** | `job_store.py:51`, `job_registry.py:81,393`; genuine interruption→unknown `session_service.py:858` |
| 1f | No auto-retry | **PASS** | `PROTECTED_TARGET` ∉ `_RETRYABLE_CODES` `errors.py:173-175`; guard raises, never retries |
| 1g | No `Workbook.RefreshAll` | **PASS** | Only docstrings state "never used"; no call site |
| 1h | No pivot materialization | **PASS** | `pivot_tables_adapter.py:532` metadata-only |
| 1i | `SCHEMA_VERSION=1` | **PASS** | `models.py:65` |
| 1j | No new `ErrorCode` beyond accepted set | **PASS** | Reuses `PROTECTED_TARGET` `errors.py:45` + pre-existing `ExcelVbaMcpError` import `session_service.py:55` |
| 2a | Early `PROTECTED_TARGET/backup_exists` guard in decision phase | **PASS** | `session_service.py:683-699` (after decision matrix, before flush) |
| 2b | Late `unknown` return removed | **PASS** | Old in-`_atomic_save_as_copy` block deleted (diff) |
| 2c | Genuine interruption path intact | **PASS** | `session_service.py:854-865` (`outcome="unknown" if backup_created else "error"`) |
| 2d | No regression to atomic temp→replace→verify | **PASS** | `session_service.py:800-872` unchanged (mkstemp→copy→fsync→verify→backup→`os.replace`→final verify→rollback) |
| 3 | No stray edits outside `session_service.py` + `test_save_as.py` | **PASS** | `git diff --name-only` = 2 files; untracked = `.ai/audit/**` docs only |
| 4a | `pytest tests/unit -q` green | **PASS** | exit code `0`; `test_save_as.py` 15 passed |
| 4b | ruff clean on changed files | **PASS** | `ruff check …` → "All checks passed!" |
| 5 | Comparative-benchmark blocker is real | **PASS (confirmed real)** | Global config read blocked by `external_directory: deny`; only candidate xlvba server reachable |

---

## 2. Evidence detail

### 2.1 The fix is minimal and correctly placed
`git diff HEAD --stat` = **2 files, +49 / −38**. No import churn: `ExcelVbaMcpError` was already imported (`session_service.py:55`) and used for the "already open" paths (`:311`, `:326`).

New guard (`session_service.py:683-699`), placed **after** the decision matrix (`:671-681`) and **before** the flush (`:702`):
```python
if exists and Path(f"{canonical_t}.bak").exists():
    raise ExcelVbaMcpError(
        "A backup already exists for the target; refusing to overwrite it (no auto-resume).",
        code=ErrorCode.PROTECTED_TARGET,
        context={"reason": "backup_exists", "target_path": str(canonical_t),
                 "backup_path": f"{canonical_t}.bak"},
    )
```
The late `outcome="unknown"` block was removed from `_atomic_save_as_copy`; the docstring (`:786-790`) now points to the early guard. The genuine interruption path (`except OSError` → `unknown` when `backup_created`) is untouched, and the atomic publish algorithm is byte-for-byte unchanged.

### 2.2 Behavior is fail-closed and non-retryable
`PROTECTED_TARGET` maps to `errorCategory="unsupported_feature"`, `isRetryable=False` (`errors.py:148,173-175`). The caller receives an actionable `isError` (not an indeterminate `unknown`), with `.bak` and destination both left untouched — "no auto-resume" is honored.

### 2.3 Test coherence
`tests/unit/test_save_as.py` (15 cases) updated cases 12 & 15 to expect `PROTECTED_TARGET/backup_exists`; case 11 (`test_interruption_after_backup_is_unknown`) still asserts the genuine post-backup interruption → `unknown`; case 7 explicitly unlinks the `.bak` before the second save so the new guard does not mask path-spelling recognition. No other test in the suite asserts the old Save As `.bak`→`unknown` behavior (the `unknown` matches in `test_job_store.py` / `test_job_recovery_interrupt.py` are job-recovery, unrelated).

### 2.4 Test / lint execution
- `pytest tests/unit` → **exit 0** (green).
- `pytest tests/unit/test_save_as.py -v` → **15 passed in 1.71s**.
- `ruff check src/excel_vba_mcp/application/session_service.py tests/unit/test_save_as.py` → **All checks passed!**

### 2.5 Benchmark blocker is genuine (not fabricated)
Reproduced directly: reading the global Kilo config `C:\Users\MatveichevN\.config\kilo\kilo.jsonc` is refused by the active permission rule `{"permission":"external_directory","pattern":"*","action":"deny","source":"global"}`. Only the candidate `xlvba` MCP server is connected (no baseline server is routable to a subagent), and the single-Excel / single-STA invariant forbids a concurrent second server. A true paired baseline-vs-candidate model-in-the-loop run is therefore infeasible here. Documented consistently in `agent_usability_benchmark_results.md §2` and `stage1_integration_evidence.md §4`.

---

## 3. Advisory findings (LOW — none mandatory, none blocking)

- **[BEST-PRACTICE][LOW] A1 — Use the typed subclass.** The guard raises the base `ExcelVbaMcpError(code=PROTECTED_TARGET)` while the same method raises the typed `ProtectedTargetError` at `:631/:673/:678`. Envelope output is identical, so this is a consistency nit only. *What:* raise `ProtectedTargetError(...)`. *Why:* uniform typed-error surface. *Effort:* trivial. *Priority:* recommended.
- **[BEST-PRACTICE][LOW] A2 — TOCTOU window.** `.bak` existence is now checked in the decision phase (before the `await`ed flush) rather than immediately before `_copy_and_fsync`, slightly widening the window in which an external actor could create a `.bak` that would then be overwritten. Under the single-writer working-copy model this is negligible. *What:* optionally re-check inside `_atomic_save_as_copy`. *Why:* defense-in-depth. *Effort:* trivial. *Priority:* optional.
- **[BEST-PRACTICE][LOW] A3 — Action text fit.** `PROTECTED_TARGET`'s `suggestedNextActions` ("use a writable target …") is an imperfect remedy for `backup_exists` (real remedy: remove/rename the `.bak` or choose a new target). Consistent with the existing mapping; not a regression. *What:* context-aware action text. *Why:* clearer agent guidance. *Effort:* trivial. *Priority:* optional.

---

## 4. Ship statement

The `.bak` fix is **minimal, correctly placed, and fail-closed**. It moves a deterministic conflict (pre-existing `<target>.bak`) from an indeterminate `outcome="unknown"` to a pre-mutation `PROTECTED_TARGET/backup_exists` error, without weakening any accepted Stage 1 invariant: single hidden Excel, single-STA FIFO, Job Object lifecycle, working-copy model, source-overwrite protection (`overwrite_source` + hash), first-class `INTERRUPTED_UNKNOWN`, no auto-retry, no `RefreshAll`, no pivot materialization, `SCHEMA_VERSION=1`, and no new `ErrorCode`. The genuine interruption→`unknown` path and the atomic temp→replace→verify algorithm are intact. Unit tests are green and ruff is clean. The comparative-benchmark blocker is real and correctly disclosed.

**The fix is safe to ship.** The three LOW advisories are optional polish and do not gate release.