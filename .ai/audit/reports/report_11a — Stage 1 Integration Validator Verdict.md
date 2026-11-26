# Report 11a — Stage 1 Integration Validator Verdict (excel_save_as)

- **Role:** Independent Integration Validator (Stage 1 Save As rollout)
- **Candidate:** `ef3e6cb` (HEAD, confirmed `ef3e6cb11c7a33900a11855d604053ea6729c568`) + uncommitted `.bak` fix
- **Working tree:** `M src/excel_vba_mcp/application/session_service.py`, `M tests/unit/test_save_as.py` (the `.bak` fix; `git diff --stat` = 2 files, +49/−38)
- **Baseline:** `40c8b4d`
- **Mode:** VALIDATION ONLY — no source modified
- **Method:** code inspection + live-file hash re-computation + unit execution (`pytest`, `ruff`) + config/environment inspection

## Verdict summary

| # | Item | Verdict |
|---|------|---------|
| 1 | Live scenario matrix credibility (hash-backed, flags, session binding) | **APPROVE** |
| 2 | Scenario-6 `.bak` defect fix (deterministic pre-mutation, no false success) | **APPROVE** |
| 3 | No `wb.SaveAs`; `_find_open_workbook` intact; single-STA/Job unchanged | **APPROVE** |
| 4 | Candidate real-agent run genuine + correct tool selection/typed args | **APPROVE** (testimonial caveat) |
| 5 | Comparative baseline benchmark genuinely blocked (no proxy) | **APPROVE** |

**Stage 1 Save-As integration verdict: PASS** — with one required documentation fix (contract §5.3 is stale) and one deferred live re-test (scenario 6 via MCP, Stage 2). No code blocker.

---

## Item 1 — Live scenario matrix credibility — APPROVE

Each PASS is backed by an independent SHA-256, and the on-disk artifacts in `.ai/tmp/saveas_live/` re-compute to **exactly** the matrix values (independently hashed by this validator):

| Matrix claim | File | Re-computed SHA-256 (actual) | Match |
|---|---|---|---|
| source original `17A576DB…` | `source.xlsx.bak` | `17A576DB00D1EE3B77E80F6B39C42D95871B4A40973F06F3C2E6A21A71290541` | ✓ |
| scenario 4b source replaced `FD35D8…` | `source.xlsx` | `FD35D88405BDCB13DFF19EC051D3D4EE180ABF7FA729B1F1F87A3C7F86D3E09B` | ✓ |
| scenario 1 out_new `45F024AF…` | `out_new.xlsx.bak` | `45F024AF5908DEBE315B5FF043F98CA3AB7EA9CEFD9529C0411EB0E898606205` | ✓ |
| scenario 3 out_new `FA6ACC…` | `out_new.xlsx` | `FA6ACC941CD1113DB2BDD92429435067A0065C3433F3EDF2896D1677E0839905` | ✓ |
| scenario 7 out_s7 `2C8A9F…` (unchanged) | `out_s7.xlsx` | `2C8A9F141CC93184C9709CD8DE0F4CF2FDD6913619CA1243CC1841C56402D62D` | ✓ |
| agent T3 output `EA7425F9…` | `bench_out.xlsx` | `EA7425F9AE8657EC0E6CB40E9A63EE0F2E152FD795EF2C980CCA486E26A22543` | ✓ |

**Internal hash-chain coherence (strong anti-fabrication signal):** `out_new.xlsx.bak` (the pre-op backup captured by scenario 3) equals the scenario-1 output hash `45F024AF…`; `source.xlsx.bak` equals the original source hash `17A576DB…`. The backups and live files form a consistent sequential chain rather than independent asserted strings.

**Flags behave per contract** (verified against `session_service.py`):
- Overwrite-off on existing target → `PROTECTED_TARGET/target_exists` (L677–681); target untouched.
- `T==S` no flag → `PROTECTED_TARGET/source_protected` (L672–676).
- `T==S` + `overwrite_source` → source replaced, `.bak`=original, gated by source re-hash (L657–666, L668, L672).
- Target↔source compared by `normcase` (L630, L648, L668) — alias/case safe.

**Session stays bound (scenario 7):** `out_s7.xlsx`=`2C8A9F…` unchanged after a subsequent `excel_save`; the later save writes the working copy via `save_workbook`→`wb.Save()` (L537–552) and never the Save-As target. Unit case 13 (`test_session_identity_unchanged_and_later_save`, L432–454) asserts `workbook_id`/`path`/`source_path` unchanged and target hash unchanged — **passes**.

**Caveat (non-blocking):** on-disk files are the *end state* after the sequential run, not per-scenario snapshots; credibility rests on the coherent hash chain + code path, which is sufficient.

## Item 2 — Scenario-6 `.bak` defect fix — APPROVE

**Fix location & ordering (pre-mutation confirmed):** `session_service.py` L689–699 rejects a pre-existing `<target>.bak` *before* `_flush_working_copy` (L702) and `_atomic_save_as_copy` (L705). The temp file is created only inside `_atomic_save_as_copy` via `tempfile.mkstemp` (L802), i.e. **after** the guard — so the rejection creates no temp and performs no `os.replace`. The existing `.bak` and destination are never touched.

**Deterministic error (not `unknown`):** raises `ExcelVbaMcpError(code=ErrorCode.PROTECTED_TARGET, context.reason="backup_exists")` (L690–699) → surfaces as an `isError`, not an `outcome="unknown"` success-shaped result.

**Genuine interruption→`unknown` preserved:** `_atomic_save_as_copy` still returns `outcome = "unknown" if backup_created else "error"` on `OSError` (L858), and the final-verify-failure branch returns `"unknown"` with `backup_path` (L835–845). Unit case 11 (`test_interruption_after_backup_is_unknown`, L374–400) still asserts `outcome=="unknown"` + `backup_path` — **passes**.

**No false-success path:** the only `outcome="ok"` return is L847–853, guarded by the final size+SHA-256 verify of the published target (L834–835). Success is never inferred from "no exception."

**No silent source overwrite:** source overwrite requires `same and overwrite_source` (L672) **and** a matching re-hashed source (L657–666); drift → `STALE_TARGET` before any write.

**Regression tests assert the new contract:**
- Case 12 (L406–426): second save with existing `.bak` → `pytest.raises(ExcelVbaMcpError)`, `code is PROTECTED_TARGET`, `reason=="backup_exists"`, `.bak`=`b"version-1"` and target=`b"original-working-bytes"` both unchanged.
- Case 15 (L487–507): pre-existing `.bak` → deterministic `PROTECTED_TARGET`/`backup_exists`, `backup_path`/`target_path` asserted, `.bak`=`b"prior-interrupted-backup"` and target=`b"current-target"` unchanged.

**Execution evidence (this validator):** `pytest tests/unit/test_save_as.py` → **15 passed in 1.09s**; `ruff check` on both changed files → **All checks passed!**

**WARN (non-blocking):** the *live MCP* re-test of the fix was **not** performed — evidence §3 L40 defers it to a Stage 2 post-restart health check (the running server holds pre-fix code). The on-disk `out_bak.xlsx` (`2D1732D0…`, 1589 B) + `out_bak.xlsx.bak` (28 B) are *consistent* with the fix (target unchanged, sentinel `.bak` untouched) but their provenance through the fixed MCP path is not proven. The "FIXED" verdict therefore rests on code review + passing unit cases 12/15 + ruff — acceptable for Stage 1 because the guard is deterministic and unit-proven; live confirmation is genuinely outstanding.

## Item 3 — No `wb.SaveAs`; `_find_open_workbook` intact; single-STA/Job unchanged — APPROVE

- **`wb.SaveAs` never called:** a full-tree search for `SaveAs` yields 15 hits, **all** comments/docstrings or the identifiers `SaveAsResult` / `GetSaveAsFilename` (e.g. `session_service.py` L591/701 comments, `models.py` L1485 `class SaveAsResult`, `macro_service.py` L137 `"GetSaveAsFilename"`). No executable `.SaveAs(` call exists. The flush uses `wb.Save()` (L760) and the publish is pure-filesystem `os.replace` (L831).
- **`_find_open_workbook` intact:** L1604–1620 matches the open `Workbook` by `FullName.lower() == working_path.lower()`; because Save As never rebinds `FullName`, resolution is unaffected. Unchanged by the diff.
- **Single-STA / Job Object unchanged:** the entire uncommitted diff is confined to `save_as_workbook` (event-loop guard) and `_atomic_save_as_copy` (pure filesystem) plus docstrings. No new thread, no COM, no Job/process change is introduced — the single-STA worker + Job Object architecture is unchanged by construction.
- Tool is wired: `excel_save_as` registered in `server.py` (L166 import; L1799–1802 registration with annotations).

## Item 4 — Candidate real-agent run genuine + correct tool selection — APPROVE (testimonial caveat)

- **Artifact corroboration:** the claimed T3 output hash `EA7425F9…` re-computes exactly from `bench_out.xlsx` (`EA7425F9AE86…`); the input fixture `bench_fixture.xlsx` (`5E31FDB6…`, 1589 B) is present. The run therefore produced a real file matching the claimed result — not a fabricated hash.
- **Tool selection / typed args (per evidence §2):** T1 `xlvba_xlvba_environment`; T2 `excel_open(inspect,writable)`→`range_read_values`(A1)→`excel_close`(no save); T3 `excel_open(execute)`→`snapshot_create`→`range_write_values`(typed `CellValue "updated_by_agent"`)→`excel_save_as(new path, overwrite_existing=false)`→`excel_close`. This is the correct sequence (snapshot before mutation; typed value write; save-as with safe flag; no macro execution), with no errors.
- **Caveat:** "genuine LLM subagent, no human intervention, no source reading" is a *process* claim that cannot be falsified from artifacts alone. The matching output hash is strong corroboration that the run occurred and produced the claimed artifact; genuineness of the agent loop is accepted as testimonial evidence, consistent with the artifacts.

## Item 5 — Comparative baseline benchmark genuinely blocked (no proxy) — APPROVE

- **No baseline server is routable:** this validator's own connected MCP inventory is `xlvba` (candidate), `context7`, `playwright` — there is **no** baseline Excel server present. No project-side baseline server is configured (`.kilo` "baseline" matches are unrelated audit-methodology docs; no `mcpServers`/baseline entry).
- **Global config is agent-read-restricted (confirmed live):** the agent-facing `read` tool on `C:\Users\MatveichevN\.config\kilo\kilo.jsonc` is rejected by the `{"permission":"external_directory","pattern":"*","action":"deny","source":"global"}` rule. Re-pointing to a baseline would require an external-directory **write**, also covered by that deny. (Note: a shell-level read reported the file length, but the agent's file tool is genuinely blocked — the "read-restricted" claim holds at the agent/tool layer.)
- **No proxy substituted:** evidence §4 states proxy results are not substituted, and the candidate-only run is explicitly labeled *supplementary, not a comparison*. The block is structural (no second server, no route, no write access), not mere inconvenience.

---

## Required Fixes

1. **DOC-UPDATE — `save_as_export_contract.md` §5 step 3 (line 45) is stale.** It still states: *"if `<target>.bak` already exists (prior interrupted save) → **stop**, `outcome="unknown"`, preserve `.bak`, do not replace destination."* After the fix the behavior is a deterministic `PROTECTED_TARGET` / `reason=backup_exists` **isError** raised pre-mutation — **not** `outcome="unknown"`. This also contradicts the same document's header (line 4: *"Validator ACCEPT (after `.bak` guard fix)"*). Update §5.3 (and cross-check §7 line 55, which remains correct for *genuine* post-backup interruption) so the contract matches the implementation.

## Advisory Recommendations (non-blocking)

1. **Complete the deferred live re-test (Stage 2):** after an MCP server restart, run scenario 6 through the live `excel_save_as` path and confirm the deterministic `PROTECTED_TARGET/backup_exists` envelope plus untouched `.bak`/target. This closes the only gap between unit-proven and live-proven behavior.
2. **Pre-existing TOCTOU (out of scope, not introduced here):** there is no per-target lock; two concurrent `excel_save_as` calls on the same target could race between the guard (L689) and backup creation (L825). This predates the fix (the original `_atomic_save_as_copy` had the same window) and is not part of the scenario matrix — noted for future hardening only.
3. **Orphan `.bak` with absent target:** the guard requires `exists` (L689); if the target is absent but an orphan `.bak` remains, the save proceeds and `_atomic_save_as_copy` (with `target_exists=False`) never touches the `.bak` (L825) — safe, no overwrite, no false success. No action needed; documented for completeness.

## Warnings

- **Documentation inconsistency:** contract §5.3 vs implementation (see Required Fix 1).
- **Rollout:** the fix is uncommitted in the working tree; Stage 2 must ensure the candidate server is restarted so the live path runs the fixed code (the current server holds pre-fix code per evidence §3 L40).
- **Dependency/architecture:** none — no dependency, schema (`SCHEMA_VERSION=1`, additive `SaveAsResult`), or concurrency-model change.

## Execution Validation

- Targets exist and are current: `save_as_workbook` (L577), `_atomic_save_as_copy` (L772), `_find_open_workbook` (L1604), `excel_save_as` (workbook_tools L167; server L1799) all present at `ef3e6cb` + working tree.
- Plan not stale: fix matches plan_05 §5.5/§9 intent (fail-closed, session stays bound, deterministic backup conflict).
- Architecture consistent and not degraded: pure-filesystem publish + STA flush preserved; boundaries clean (api → application → STA worker).
- **Execution readiness: READY**, conditional on Required Fix 1 (doc) and the Stage 2 live re-test (Advisory 1).

**Final Stage 1 Save-As integration verdict: PASS.**