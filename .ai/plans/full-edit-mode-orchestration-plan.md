# Full Edit Mode Orchestration — Dedicated Execution Plan

**Plan ID:** `FULL-EDIT-MODE-ORCHESTRATION-001` / plan file `FE`
**Owner:** Planner (maintained live by the Planner through Phases 0–6)
**Accountable:** Tech Lead (orchestrator)
**Repo / branch:** `C:\py_exp\xlvba_mcp` @ `master`
**Baseline commit:** `8c22f11` (`docs: align with full-access default, macro classification, and error taxonomy (task_b8_docs)`); prior remediation `B1–B8` complete: `46d9b1b`(B2) `c246edd`(B3) `fe63af2`(B4) `47713f7`(B5) `035edea`(B6) `71e9d39`(B7) `8c22f11`(B8). **B9 (independent final verification) is the open remediation item.**

**Relationship to other documents (additive only — never overwrite):**
- Source orchestration spec: `.ai/plans/task_tech_lead_full_edit_mode_orchestrator.md`
- Corrective plan: `.ai/plans/plan_corrective_full_access.md` (`fix_001`–`fix_009`)
- Block decomposition + live ledger: `.ai/plans/01-plan-corrective-full-access-execution.md` (B1–B9)
- Discrepancy report: `.ai/audit/problems/doc-vs-impl-discrepancies-2026-11-21.md` (D-001…D-005)
- Existing real-MCP e2e harness: `tests/integration_excel/test_mcp_full_access_workflow.py`

This plan does **not** re-implement B1–B8. It (1) establishes a baseline, (2) confirms which prior
findings still apply at the current commit, (3) decomposes any **confirmed remaining** gaps into small
tasks (`FE-*`), (4) drives them through Implementor, then (5) executes Phase 3 function-by-function MCP
verification, Phase 4 real edit/macro/edit/macro E2E, Phase 5 hidden-execution checks, Phase 6 final
audit + capability report.

**Planning-only constraint honoured:** no source, no existing plan, no audit file is modified by this
document.

---

## 1. Sources of truth and hard constraints (inherited, binding)

Priority order (spec §4): approved spec/addendum → corrective plan + this orchestration spec → current
code/tests/docs/audit evidence → Researcher findings. **Prior reports are not proof of current behavior;
every task's Implementor/Auditor re-confirms the finding at the current commit before acting.**

Non-negotiable constraints (spec §4/§10, corrective plan §1):
- Full access is default when `EXCEL_VBA_MCP_ACCESS_MODE` is unset; default `excel_open` is writable/execute.
- No per-action grants / host allowlist in normal Full mode; restricted/read-only modes preserved as explicit controls.
- Preserve background-safety mechanism (`macro_classify_background` + R13 veto + hash-bound registry); never blanket-allow.
- Excel stays hidden in the **current interactive Windows user session**; no VM/container/separate session/service/SYSTEM/remote; no foreground clicks / `SendKeys` / focus theft.
- Use working-copy + snapshot architecture; originals must remain byte-identical; never mutate a repo fixture directly.
- Do not weaken Trust Center / VBOM / macro security / path policy / audit logging / static guards (`tests/unit/test_security_policy_static.py`).
- Max **2** parallel subagents; **never** two Implementors on the same file; **all** Excel/STA integration serial through the single `ExcelComWorker`.
- Each completed implementation task = its own Git commit with tests green; record commit hash in §5 ledger.

---

## 2. Baseline confirmation (Phase 0 — resolves "does the prior finding still apply?")

The Auditor baseline report (`.ai/audit/reports/xlvba-mcp-full-edit-mode-baseline.md`, **dependency**)
is authoritative for the live tool count and current behavior. The Planner's pre-read at `8c22f11`
records the following **evidence** so tasks below are targeted, not speculative:

| # | Candidate gap | Evidence at `8c22f11` | Still applies? | Disposition |
|---|---|---|---|---|
| G-1 | **Tool count is 88, not 87** | `server.py` has **87** `server.add_tool(...)` calls + **1** `@mcp.tool()` (`server_health`) = **88** live tools; **4** `@mcp.resource(...)` (`workbook_model`, `workbook_vba_components`, `workbook_vba_component_source`, `test_run_summary`); **0** prompts. Docs `01-overview.md:39`, `04-architecture.md:60`, `05-mcp-tools-reference.md:21` say **88**. The "87" figure in the corrective plan/audit (C-audit-14/23) predates `macro_classify_background`. | **YES** — stale 87 count in older plan docs | **FE-01** (docs alignment; verify 88). Auditor confirms live count. |
| G-2 | **`classify_operation` stale + not wired** | `security/macro_policy.py::classify_operation` lists nonexistent tools (`compile`,`vba_write`,`test_run`) and omits all pivot/range/names/shapes/controls/sheet/table/workbook/vba tools. `grep classify_operation(` in `src/` → **1 hit** (its own def). Referenced only by `tests/unit/test_macro_policy.py`. Runtime gate is each service's hardcoded `OperationClass`. | **YES** — still stale, still unwired | **FE-06** (decide: document unwired vs correct table). Not an enforcement path; low risk. |
| G-3 | **`pywinauto` unused dependency** | `pyproject.toml:13` `"pywinauto>=0.6.8"`; `grep pywinauto src/` → **no import** (UI automation is `comtypes`). `adapters/windows_ui/vbe_compile_ui.py` uses comtypes. | **YES** | **FE-07** (remove or mark reserved + fix release-notes row). |
| G-4 | **Residual interactive tokens** `AppActivate` / `Application.Visible` | `INTERACTIVE_TOKENS` in `application/macro_service.py` includes `MsgBox,InputBox,.Show,FileDialog,Shell,CreateObject,FollowHyperlink,SendKeys,DoEvents,GetOpenFilename,GetSaveAsFilename,CallByName,ExecuteGlobal,Execute,Eval,OnTime,DisplayAlerts,.Run,CommandBars` — **no `AppActivate`**, and `Application.Visible` only matches `.Show` if written as such (a literal `Application.Visible = True` is **not** caught). `docs/05` line ~586 already documents this residual. | **YES** — documented residual (B5 W2) | **FE-04** (decision: add tokens vs document residual). Low. |
| G-5 | **`macro_run` snapshot precondition (CT-B7 / PO-3)** | `application/macro_service.py::MacroService.__init__` has **no** `SnapshotService`; `run_macro` never calls `require_snapshot`. `application/snapshot_service.py` docstring lists "macro execution" among gated entry points. Mutation tools DO require a snapshot; `macro_run` does not. | **YES** — real spec deviation | **CT-B7** (decision gate). **Tech Lead decision needed (§8 D-1).** |
| G-6 | **Stale worker / dead-owned-Excel `excel_health` (R-17)** | `api/health_tools.py::excel_health` reports `worker_state/accepting/needs_recovery/owned_pid` purely from `ExcelComWorker.get_status()`; worker only enters `NEEDS_RECOVERY` on timeout, **not** COM disconnect (`owned_pid` dead, `last_hresult=0x800706BA`, health `ok`). No liveness probe of `owned_pid`. | **YES** — environment/hardening gap | **FE-08** (decision: harden vs document). Prior decision = document only. **Tech Lead decision needed (§8 D-2).** |
| G-7 | **`vba_export_project` / `snapshot_create` file writes ungated** | `application/vba_project_service.py::export_project` writes to `tempfile.mkdtemp(...)` or a `PathPolicy`-validated artifact dir with **no** `_require_writable`/`_require_snapshot`/`_require_approval` (mutation helpers start at `replace_component_source`). `snapshot_service.py::restore` has **no** approval gate (documented R-19). | **YES** — by design, under-documented | **FE-09** (document side effects; confirm no weakening). Low. |
| G-8 | **`recover_session` / `close_session` not MCP-exposed (D-002)** | `application/session_service.py::recover_session` / `close_session` exist; **not** registered in `server.py::_register_tools`. `excel_health` reports `needs_recovery` but a client cannot trigger recovery (remedy = close/reopen). | **YES** | **FE-11** (decision: expose `excel_recover_session` vs document). **Tech Lead decision needed (§8 D-4).** |
| G-9 | **Structured error surfacing for all tools** | `server.py::ExcelVbaMcpServer.call_tool` maps `ExcelVbaMcpError` → `ToolError(json)` (absorbs B1 R-16). But any **non**-`ExcelVbaMcpError` exception (e.g. `pydantic.ValidationError`, `KeyError`, bare `ValueError`) still becomes SDK `UnexpectedToolError("Error executing tool <name>")` — no `code`/`context`. Needs per-tool verification. | **PARTIAL** — domain errors covered; non-domain unclear | **FE-03** (test matrix + add wrapper only if a tool leaks). |
| G-10 | **D-001 docs reflect `macro_classify_background`** | `docs/05` documents `macro_classify_background` (rows/section), `background_safe` gate, `INTERACTIVE_UI_NOT_SUPPORTED`, `STALE_TARGET`; docs updated by B8. D-001 ("not settable over MCP") is **stale** as a current statement. | **RESOLVED by B8** — confirm in FE-01 | **FE-01** confirms; no separate task. |
| G-11 | **Doc-vs-impl D-002/D-003/D-004/D-005** | D-004/D-005 marked doc-only corrected (but those docs are deleted in the working tree, `git status` shows `D`). D-002 → G-8/FE-11; D-003 → G-3/FE-07. | **YES** (residual) | folded into FE-07 / FE-11. |

> If the Auditor's baseline reports a **different** live count or contradicts any row above, the Planner
> amends this section and the affected `FE-*` task before implementation (Gate 1).

---

## 3. Task catalogue (FE-*)

Legend — Risk: **L**ow / **M**edium / **H**igh. Agent set per spec §6 risk table.
All `FE-*` implementation tasks are **single-Implementor, one commit each**, unless a Research/Decision
gate precedes them. Excel/STA-touching tests are always serial.

| ID | Title | Scope (semantic) | Depends on | Risk | Agents (required) | Parallel? |
|---|---|---|---|---|---|---|
| **FE-00** | Baseline acceptance & live tool census | Read-only. Produce/confirm baseline report; capture **live** tool catalogue (88 tools + 4 resources + 0 prompts) via real `initialize`/`list_tools` on a disposable workbook; classify each tool READ_ONLY/MUTATING/EXECUTING/MIXED/ADMIN; record effective access mode with env unset; confirm G-1…G-11 at current commit. **No code.** | — | **H** (audit; live Excel read-only) | **Auditor** (req), **Validator** (req for live census), Planner (this plan), Researcher only if uncertainty | **No** — single serial live probe (STA) |
| **FE-01** | Tool-count & macro-classification doc alignment (G-1, G-10) | Docs-only: ensure tool count is **88** everywhere (`README.md`, `docs/01..07`, `docs/ru/01..05`), `macro_classify_background` present, `background_safe` workflow accurate; fix broken `08-security.md` links to real docs; no new doc files; follow `docs/00-overview/doc-maintenance-rules.md`. | FE-00 | **M** | **Implementor** (docs), **Validator** (spot-check) — per 01-plan B8 doc path | **Yes** (docs; disjoint from code) |
| **FE-02** | Verify structured error surfacing across all mutating/executing/admin tools (G-9) | Test-first: through real MCP, call a representative of every non-read-only tool with (a) a mutating denial and (b) a malformed arg; assert each returns structured `code` (`APPROVAL_DENIED`/`WORKBOOK_READ_ONLY`/`SNAPSHOT_FAILED`/`BACKGROUND_SAFETY_NOT_VERIFIED`/`INTERACTIVE_UI_NOT_SUPPORTED`/`COM_BUSY`/`STALE_TARGET`). If any tool leaks `UnexpectedToolError` (no code), add the **smallest** wrapper/translation at the `ExcelVbaMcpServer.call_tool` boundary or in the offending `api/*_tools.py` (no domain error weakening). | FE-00 | **M** | **Implementor** (req), **Auditor** (boundary review), **Validator** (independent matrix) | **No** — serial MCP |
| **FE-03** | Expand real-MCP e2e harness with a second non-VBA mutation + explicit restricted-mode macro denial (G-5 test-side) | Extend `tests/integration_excel/test_mcp_full_access_workflow.py` (TEST-ONLY): add (a) a non-VBA mutation besides `table_update_rows` (e.g. `range_write_values`/`name_create`) read back per Phase 4 step 14; (b) an explicit restricted-mode server case asserting `macro_run` → `APPROVAL_DENIED` (or `excel_open` execute denial already covered — keep both); preserve PO-1..PO-6. | FE-00 | **M** | **Implementor** (req), **Validator** (reproduce), **Auditor** (fixture determinism) | **No** — serial Excel |
| **FE-04** | Residual interactive tokens `AppActivate` / `Application.Visible` (G-4) | Decision-gated (see §8 D-3). If approved: add `AppActivate`, `Application.Visible` (and consider `.Activate`, `WindowState`) to `INTERACTIVE_TOKENS` in `application/macro_service.py`, update `tests/unit/test_macro_service.py` + docs residual note; re-run B5 MUST-fix tests. If deferred: docs-only residual note already present → mark NOT APPLICABLE. | FE-00 | **L** | **Implementor** (if approved), **Auditor** (false-safe check), **Validator** (case-insensitive/heuristic tests) | **Yes** (isolated file; not concurrent with FE-10) |
| **FE-05** | Verify restricted/read-only safeguards remain intact (regression) | Test/verify only: restricted server denies execute open (`APPROVAL_DENIED`); read-only session mutation → `WORKBOOK_READ_ONLY`; macro security block surfaces `MACRO_BLOCKED`/`MACRO_NOT_FOUND` per B5 note; `tests/unit/test_security_policy_static.py` green. No behavior change expected. | FE-00 | **M** | **Validator** (req), **Auditor** (trace) | **No** — serial Excel |
| **FE-06** | `classify_operation` stale table — correct or document unwired (G-2) | Decision-gated (§8 D-5). Option A: correct `_READ_ONLY_TOOLS`/`_MUTATING_TOOLS`/`_EXECUTING_TOOLS` additively to match the 88-tool catalogue + keep `MacroPolicy.classify_operation` delegate, update `tests/unit/test_macro_policy.py`; keep it **non-runtime** (do not wire as enforcement). Option B: add explicit "not used at runtime; owned by services" docstring + leave table. Low risk. | FE-00 | **L** | **Implementor**, **Auditor** (confirm not wired into path) | **Yes** (isolated file) |
| **FE-07** | `pywinauto` dependency disposition (G-3, D-003) | Decision-gated (§8 D-6). Remove `"pywinauto>=0.6.8"` from `pyproject.toml` (+ `uv.lock` refresh) OR mark `# reserved: UI automation via comtypes` with a release-notes row. Update any release-notes "Key deps". Verify `tests/unit/test_security_policy_static.py` and import graph unaffected. | FE-00 | **L** | **Implementor**, **Validator** (smoke import) | **Yes** (isolated file) |
| **FE-08** | Stale worker / dead-owned-Excel health (G-6, R-17) — harden or document | Decision-gated (§8 D-2). **Document path (default):** add explicit known-limitation + recovery runbook (close/reopen; restart server) to `docs/06`/`docs/05`; no code. **Harden path (only if approved):** add a bounded owned-PID liveness check to `ExcelComWorker.get_status()`/`excel_health` that flips `needs_recovery=true` on `0x800706BA` without touching COM, plus unit tests. High-touch (worker lifecycle) if hardened. | FE-00 | **M** (doc) / **H** (harden) | doc: **Implementor**+`Validator`; harden: **Researcher**→**Auditor**→**Implementor**→**Validator** | **No** (worker serial) if hardened |
| **FE-09** | Filesystem-write side effects of `vba_export_project` / `snapshot_create` — document + confirm no ungated escape (G-7) | Verify `export_project` uses temp dir or `PathPolicy`-validated artifact root; `snapshot_create` writes under snapshot root. Add explicit side-effect + preconditions to `docs/05`/`docs/06` and tool docstrings; **do not** add approval gates to read-only tools (R-19). If a path escape is found → corrective task. | FE-00 | **L** | **Implementor** (doc), **Auditor** (path-policy confirm), **Validator** (call + path check) | **Yes** (docs) |
| **FE-10** | `macro_run` snapshot precondition — CT-B7 (G-5) | Decision-gated (§8 D-1). **If approved:** inject `SnapshotService` into `MacroService` and call `require_snapshot` before dispatch (→ `SNAPSHOT_FAILED` when absent); update ~15 macro tests + e2e to snapshot-first (e2e already does); update docs. **If deferred:** record as known limitation in docs; no code. High-touch (macro lifecycle + constructor wiring). | FE-00; coordinate with FE-03 | **H** if approved; **L** if doc-only | **Researcher**→**Auditor**→**Implementor**→**Validator** (if code); doc-only: **Implementor**+**Validator** | **No** — serial macro/STA |
| **FE-11** | Surface `recover_session` / `close_session` (G-8, D-002) — expose tool or document | Decision-gated (§8 D-4). Expose `excel_recover_session` MCP tool (thin handler → `SessionService.recover_session`) with tests + docs; OR document "recovery = close/reopen" limitation. | FE-00 | **M** (expose) / **L** (doc) | expose: **Implementor**→**Auditor**→**Validator**; doc: **Implementor**+**Validator** | **Yes** if doc; **No** if code (serial STA) |
| **FE-12** | Phase 3 full function-by-function MCP verification + capability matrix | Execute §6 Phase 3: real server, real client, every live tool tested or classified; produce the capability report matrix. **May create corrective tasks** if FAILs found. | FE-01..FE-11 (all prior FE complete or deferred) | **H** | **Auditor** (req), **Validator** (req), Implementor only for approved fixtures/correctives | **No** — serial Excel |
| **FE-13** | Phase 4 real edit/macro/edit/macro E2E proof | Execute §7 Phase 4 steps 1–17 against the harness; independent Validator acceptance; Auditor call-path confirmation. | FE-12 | **H** | **Validator** (req), **Auditor** (req), Implementor (fixtures only) | **No** — serial Excel |
| **FE-14** | Phase 5 hidden-execution / desktop non-interference checks | Execute §8 Phase 5 checklist in current interactive Windows session. | FE-13 | **H** | **Auditor** (req), **Validator** (req), Researcher only if technical issue | **No** — serial |
| **FE-15** | Phase 6 final independent audit + capability report + ledger close | Final Auditor architecture/security review; Validator product acceptance; Planner closes §5 ledger; Tech Lead issues verdict. | FE-14 | **H** | **Auditor** (req), **Validator** (req), Planner (ledger) | **No** — release gate |

---

## 4. Dependencies, execution order, parallel groups

### 4.1 DAG

```
FE-00 ──▶ ┌─ FE-01 (docs) ─────────────────────────────┐
          ├─ FE-02 (error surfacing) ──────────────────┤
          ├─ FE-03 (e2e extend) ──┐                    │
          ├─ FE-04 (tokens)       │ (if FE-10 code)    │
          ├─ FE-05 (safeguards)   │                    │
          ├─ FE-06 (classify tbl) │                    │
          ├─ FE-07 (pywinauto)    │                    │
          ├─ FE-08 (worker/doc)   │                    │
          ├─ FE-09 (fs side fx)   │                    │
          ├─ FE-10 (CT-B7) ───────┘                    │
          └─ FE-11 (recover/doc) ──────────────────────┤
                                                       ▼
                              FE-12 (Phase 3) ──▶ FE-13 (Phase 4) ──▶ FE-14 (Phase 5) ──▶ FE-15 (Phase 6)
```

### 4.2 Strict execution order

1. **FE-00** (baseline gate; no code). *Serial.*
2. **Decision gates** (Tech Lead, §8): resolve D-1 (FE-10), D-2 (FE-08), D-3 (FE-04), D-4 (FE-11), D-5 (FE-06), D-6 (FE-07) — record outcome in §5 before the affected FE starts implementation. Doc-only dispositions unblock immediately.
3. **Parallel Group A (max 2, disjoint files, no Excel):** `FE-01` (docs) ∥ one of {`FE-06`, `FE-07`, `FE-09`} (each isolated single-file). Rotate the second slot; never two writers on one file.
4. **Parallel Group B (after Group A, no Excel):** `FE-02` (error surfacing — touches `server.py` boundary or `api/*`) ∥ `FE-04`/`FE-06`/`FE-07`/`FE-09` remainders (disjoint). `FE-02` must not overlap another Implementor on `server.py`.
5. **FE-10** (CT-B7) and **FE-11** (recover) — **serial**, STA/COM-touching, after decision recorded. If FE-10 code approved, it precedes **FE-03** (e2e must snapshot-first and reflect the new precondition).
6. **FE-03** (e2e extend) then **FE-05** (safeguards regression) — serial Excel.
7. **FE-08** (worker/doc) — after decision; if hardened, serial and before FE-12 (worker stability affects all live checks).
8. **FE-12 → FE-13 → FE-14 → FE-15** — strictly serial verification/release gates.

### 4.3 Parallelism hard rules (restated)

- Max **2** concurrent subagents; **never** two Implementors on the same file.
- **No** concurrent Excel/COM: all `tests/integration_excel/*` run `-p no:xdist` (serial).
- Researcher + Auditor may investigate independent topics in parallel (read-only).
- Validator inspects **completed, committed** work only (except a pre-implementation design review).
- Do not merge/commit until the task's tests + required reviews pass.

### 4.4 Critical path

**FE-00 → (decision gates) → FE-02 → FE-10? → FE-03 → FE-12 → FE-13 → FE-14 → FE-15.**
FE-01/04/06/07/08/09/11 are side branches that must land before FE-12 (or be explicitly deferred with a
documented limitation) but do not extend the critical path if parallelized per §4.2.

---

## 5. Status ledger (maintained live by Planner)

Status values: `Not started` / `In progress` / `Blocked` / `Done+commit <hash>` / `Deferred (doc)`.

| ID | Status | Commit | Notes / decision ref |
|---|---|---|---|
| FE-00 | Not started | — | awaiting Auditor baseline report |
| FE-01 | Not started | — | confirm 88 count; D-10 resolved by B8 |
| FE-02 | Not started | — | G-9 |
| FE-03 | Not started | — | extends B7 harness; PO-1..PO-6 preserved |
| FE-04 | Not started | — | D-3 (tokens) |
| FE-05 | Not started | — | regression |
| FE-06 | Not started | — | D-5 (classify_operation) |
| FE-07 | Not started | — | D-6 (pywinauto) |
| FE-08 | Not started | — | D-2 (worker/health) |
| FE-09 | Not started | — | G-7 |
| FE-10 | Not started | — | **D-1 (CT-B7)** |
| FE-11 | Not started | — | D-4 (recover) |
| FE-12 | Not started | — | Phase 3 |
| FE-13 | Not started | — | Phase 4 |
| FE-14 | Not started | — | Phase 5 |
| FE-15 | Not started | — | Phase 6 / release gate |

Decision records (Tech Lead fills; Planner transcribes): D-1..D-6 outcomes appended here as recorded.

### Tech Lead decision log (recorded 2026-11-22, Gate 0)

Gate 0 (Auditor baseline + Tech Lead independent code verification) ACCEPTED. Live surface confirmed at **88 tools** (87 `server.add_tool` + `server_health`), **4 resource templates**, **0 prompts**; default access mode **full**; originals byte-identical. Two HIGH findings were re-confirmed in source by the Tech Lead (not taken on Auditor report alone): F-1 (`snapshot_service.restore_snapshot` hardcodes `SessionMode.INSPECT` → `macro_run`→`MACRO_BLOCKED` after restore) and F-2 (health has no PID liveness probe).

| Decision | Finding | Options | **Tech Lead choice** | Affected task |
|---|---|---|---|---|
| **D-1 (NEW, F-1)** | `snapshot_restore` drops session to `inspect` → macros disabled after restore (HIGH, SPEC-DEVIATION) | (A) restore in execute mode; (B) leave inspect + document | **(A) CODE FIX.** A restore that disables macros breaks the required "save/snapshot/restore then keep editing+running macros" capability. Minimal, spec-faithful: `restore_snapshot` opens the restored working copy with `SessionMode.EXECUTE` (writable). Verify existing restore tests + e2e updated to assert `macro_run` succeeds post-restore. | **FE-16** (new; replaces the Planner's non-existent FE for F-1) |
| D-2 (R-17/F-2) | stale worker / untruthful `excel_health` | (A) document + runbook; (B) harden health with PID liveness | **(A) DOCUMENT.** Worker-lifecycle hardening is High risk and out of this release's scope; no COM-touching liveness probe in `get_status`. Add known-limitation + recovery runbook (close/reopen; restart server) to docs/05/06. | FE-08 (doc path) |
| D-3 (G-4/F-4 tokens) | residual interactive tokens | (A) add tokens; (B) document | **(A) ADD.** Case-insensitive substring veto; `Recalculate` fixture uses none of them, so strictly tightens safety with no regression. Add `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`; update unit tests + docs residual note. | FE-04 (code path) |
| D-4 (D-002/G-8) | recover_session/close_session not exposed | (A) expose `excel_recover_session`; (B) document close/reopen | **(B) DOCUMENT.** Exposing recovery is scope creep beyond Full-Edit acceptance and pairs with the deferred F-2 hardening; document the close/reopen remedy. | FE-11 (doc path) |
| D-5 (F-7/classify_operation) | stale, unwired table | (A) correct table; (B) docstring-only "not authoritative" | **(B) DOCUMENT unwired.** Correcting the table would imply authority it does not have; keep it non-runtime and mark clearly non-authoritative. | FE-06 (doc path) |
| D-6 (F-10/pywinauto) | unused dependency | (A) remove; (B) mark reserved | **(A) REMOVE** from `pyproject.toml`, refresh lock, fix release-notes row; verify static guards + import graph. | FE-07 (code path) |
| D-7 (F-3/CT-B7) | `macro_run` snapshot precondition absent | (A) require snapshot; (B) document deviation | **(B) DOCUMENT** as a known limitation (Planner's conservative path). Core edit/macro/edit/macro flow already snapshots first; constructor wiring + ~15 macro-test churn is not justified for this release. | FE-10 (doc path) |

**Revised execution order (Tech Lead):** FE-16 (F-1 fix, HIGH, serial Excel) → FE-04 (tokens, isolated) + FE-06 (classify doc) + FE-07 (pywinauto) + FE-09 (fs side-effects doc) in ≤2 parallel groups → FE-01 (tool-count/docs alignment) → FE-03 (e2e extension incl. post-restore macro) → FE-02 (error surfacing) → FE-05 (safeguard regression) → FE-08/FE-10/FE-11 (doc dispositions) → FE-12 (Phase 3) → FE-13 (Phase 4) → FE-14 (Phase 5) → FE-15 (Phase 6). Max 2 agents; never two Implementors on one file; all Excel/STA serial.

### Revised status ledger (Tech Lead)

| ID | Status | Commit | Notes / decision ref |
|---|---|---|---|
| FE-00 | **Done** | — | Gate 0 accepted; baseline `.ai/audit/reports/xlvba-mcp-full-edit-mode-baseline.md` |
| FE-16 | **Done+commit** | `0993370` | D-1: restore execute-mode fix; 867 unit + 4 live integration green; reviewed diff = minimal mode flip + docstrings |
| FE-04 | **Done+commit** | `5d649a6` | D-3: 4 tokens added; +3 unit tests; 870 unit green; ruff/mypy clean; Recalculate unaffected |
| FE-06 | **Done+commit** | `119c303` | D-5: classify_operation documented non-authoritative (item 14) |
| FE-07 | **Done+commit** | `c28b3dd` | D-6: pywinauto removed (pyproject+uv.lock+README+2 docs); no imports; 870 unit green; smoke import ok |
| FE-09 | **Done+commit** | `119c303` | G-7: snapshot_create/vba_export_project file writes documented (by design, ungated full+restricted, R-19) |
| FE-01 | **Done** | — | Docs already say **88** everywhere (EN+RU); no change needed; verified grep |
| FE-03 | **Done+commit** | `8be7245` | extend harness steps 4/6/11/14; 4 live integration pass; fixture byte-identical |
| FE-02 | **Done+commit** | `cc7abd0` | G-9: pydantic ValidationError→structured code at server boundary; 18-tool matrix + e2e 5 passed; ruff/mypy clean |
| FE-05 | **Done+commit** | `e5ebcc8` | read-only denial for range/table/macro asserted; FE-02 matrix now asserts REAL per-tool denial codes (not schema INVALID_RANGE); 5 live tests pass; fixture byte-identical |
| FE-08 | **Done+commit** | `119c303` | D-2: stale worker/health documented (item 11) |
| FE-10 | **Done+commit** | `119c303` | D-7: macro_run snapshot deviation documented (item 12) |
| FE-11 | **Done+commit** | `119c303` | D-4: recovery close/reopen documented (item 13) |
| FE-12 | **Done+commit** | `db91dbd` | Phase 3: 84/88 PASS, 0 FAIL, 4 NOT VERIFIED (shape/control tools — no fixture surface; covered at service layer by test_shapes_controls.py). Report `.ai/audit/reports/xlvba-mcp-phase3-function-matrix.md`; fixtures byte-identical |
| FE-13 | **Done+commit** | `dece703` | Phase 4: 17/17 PASS, 6/6 also-verify, 5/5 live tests via real MCP; fixture byte-identical; report `.ai/audit/reports/xlvba-mcp-phase4-e2e-acceptance.md` |
| FE-14 | **Done+commit** | `b9c13b0` | Phase 5: 9/10 PASS, 1 NOT VERIFIED (focus-theft — point-sampling only; zero forbidden-API call sites; owned Excel never visible/foreground; evidence needed = continuous foreground hook). Report `.ai/audit/reports/xlvba-mcp-phase5-hidden-execution.md` |
| FE-15 | **Done+commit** | `b08d638`, `1633115` | Phase 6: final code audit found 1 BLOCKING (C-1: ValidationError ctx not JSON-serializable → unstructured error) → **fixed** + regression tests (872 unit, 5 e2e both orders, 0 orphans). A-2 (comtypes undeclared) + D-2 (stale veto docs) + T-3 (matrix gap) fixed. Capability verdict **PARTIALLY VERIFIED** (accepted; 2 non-blocking residuals: 4 shape/control MCP tools, focus-theft continuous obs). Reports: `-capability-audit.md`, `-capability-report.md` |

---

## 6. Phase 3 — Function-by-function MCP verification matrix (template)

**Method:** launch the real server (`python -m excel_vba_mcp`, stdio) with `EXCEL_VBA_MCP_ACCESS_MODE`
**unset**; real `ClientSession` handshake + `list_tools`; call each tool against a **disposable copy**
(`FullAccessWorkflow.xlsm` and/or `RangesCellsProject`/`PivotTablesProject` fixtures) in a temp root;
record evidence. Do **not** call services directly as a substitute. Status is `PASS` only when the MCP
call was observed, not merely registered/unit-tested.

Columns (one row per live tool — **88 tools**; seed from Auditor baseline inventory):

| # | tool | category | input schema (required/optional) | preconditions | side effects | tested args | actual result / error (`code`) | status (PASS/PARTIAL/FAIL/NOT VERIFIED/NOT APPLICABLE) | call path (api→service→adapter→COM) | semantic code refs (module.class.fn) | evidence (req id / output hash / audit id) | known limitation |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | excel_open | MIXED | path; mode?; read_only? | path allowlist | creates work copy; owns Excel | `{path:<copy>}` | e.g. `read_only:false, mode:execute` | | workbook_tools.excel_open→SessionService.open_workbook→… | | | |
| 2 | excel_close | MIXED | workbook_id; save_changes? | open session | closes/destroys work copy | … | | | | | | working copy deleted (F1) |
| … | … | … | … | … | … | … | … | … | … | … | … |
| 88 | server_health | ADMIN | — | — | none | `{}` | `{status:ok}` | | server_health | | | |

Coverage target: every one of the 88 rows has a status; `NOT VERIFIED`/`NOT APPLICABLE` require a precise
reason. Error-semantics rows (Phase 3 step 8) assert the distinctions: `APPROVAL_DENIED` (restricted) vs
`WORKBOOK_READ_ONLY` (read-only) vs `SNAPSHOT_FAILED` (missing snapshot) vs `BACKGROUND_SAFETY_NOT_VERIFIED`
/ `INTERACTIVE_UI_NOT_SUPPORTED` (macro) vs `COM_BUSY` (contention) vs `PATH_NOT_ALLOWED`.

Non-tool surfaces (separate inventory, not tools): 4 resources —
`excel://workbooks/{workbook_id}/model`, `.../vba/components`, `.../vba/components/{component_name}`,
`excel://test-runs/{run_id}`; **0** prompts.

---

## 7. Phase 4 — Real edit / macro / edit / macro E2E proof checklist

Maps spec §5 Phase 4 steps 1–17 to the existing harness `tests/integration_excel/test_mcp_full_access_workflow.py`
(`test_full_access_workflow_e2e` + `test_persistence_roundtrip` + denial tests) and notes gaps (FE-03).

| Step | Spec requirement | Harness coverage | Gap / action |
|---|---|---|---|
| 1 | Start server, ACCESS_MODE unset | `_running_server(access_mode=None)` | ✅ |
| 2 | `excel_open` default → writable execute | asserts `read_only:false`, `mode:execute` | ✅ |
| 3 | Create initial snapshot before mutation | `snapshot_create` before `table_update_rows` | ✅ (PO-3) |
| 4 | Read the settings table | `table_read_rows`/`range_read_values` (implicit via update) | ⚠ add explicit `table_read_rows`/`table_get` read (FE-03) |
| 5 | Change first setting via table mutation | `table_update_rows(tblSettings, index1, SettingValue=10)` | ✅ |
| 6 | Read it back | `range_read_values` Sheet1!B2 (in persistence test) | ⚠ add read-back in main e2e (FE-03) |
| 7 | Classify deterministic macro | `macro_classify_background(Recalculate,true)` once | ✅ |
| 8 | Run macro | `macro_run modWorkflow.Recalculate` | ✅ |
| 9 | Read output reflects setting 1 | `Result!A1 == 20` | ✅ |
| 10 | Change setting again | `table_update_rows(SettingValue=20)` | ✅ |
| 11 | Read it back | (add explicit read, FE-03) | ⚠ |
| 12 | Run macro again (same session) | second `macro_run` same `workbook_id` | ✅ |
| 13 | Output reflects setting 2 (≠ step 9) | `Result!A1 == 40`, `!=20` | ✅ |
| 14 | ≥1 additional non-VBA mutation read back | — | ❌ **FE-03** (add e.g. `range_write_values` or `name_create` + read) |
| 15 | Save + verify persistence (working-copy/snapshot lifecycle) | `excel_save`; `test_persistence_roundtrip` snapshot→close(save)→restore | ✅ (F1 semantics) |
| 16 | Restore snapshot; verify settings/output/table/VBA/names | `test_persistence_roundtrip` checks Result!A1, Sheet1!B2, modWorkflow, names | ✅ (add table-structure check if cheap, FE-03) |
| 17 | Original repo/source fixtures byte-identical | repo hash + temp copy hash unchanged | ✅ |

Also verify (spec Phase 4 "Also verify"): unclassified macro → `BACKGROUND_SAFETY_NOT_VERIFIED` ✅;
interactive macro (`ShowMessage`) → `INTERACTIVE_UI_NOT_SUPPORTED` ✅; safe classified runs ✅;
read-only session → `WORKBOOK_READ_ONLY` ✅; restricted mode preserved → `APPROVAL_DENIED` ✅;
full mode needs no grants ✅ (no grant tool used).

---

## 8. Phase 5 — Hidden execution & desktop non-interference checklist

Current interactive Windows session; record objective evidence, mark unobservable `NOT VERIFIED`.

- [ ] Excel launched **hidden**, owned instance (owned PID ∉ pre-existing user PIDs) — harness asserts.
- [ ] No VM/container/separate Windows session/service/SYSTEM/remote host (host is interactive session).
- [ ] Existing user Excel instances not reused/controlled/closed — baseline PID set survives (harness asserts).
- [ ] No owned visible window (per-PID `EnumWindows`+`IsWindowVisible` empty) — harness asserts at start/open/cycles.
- [ ] No `SetForegroundWindow` / `ShowWindow` / `SendKeys` / foreground clicks in code path (grep audit).
- [ ] Interactive VBA rejected without displaying UI — `ShowMessage` classify veto (no dialog observed).
- [ ] No foreground UI automation to operate VBE — code inspection (comtypes UIAutomationCore only).
- [ ] Worker health/recovery correct after repeated mutation + macro execution — `excel_health` ok pre/post.
- [ ] Owned PID reaped after server shutdown; user PIDs alive — harness asserts.
- [ ] Focus-theft: mark `NOT VERIFIED` if no reliable observation tool; identify evidence needed (e.g. foreground-window logger).

---

## 9. Tech Lead decisions needed

Each item: evidence + options + Planner recommendation. **No implementation on the affected FE until recorded.**

**D-1 — CT-B7: should `macro_run` require a snapshot? (`FE-10`)**
Evidence: `MacroService` has no `SnapshotService`; `run_macro` never calls `require_snapshot`; `snapshot_service` docstring lists macro execution as a gated entry point; mutation tools DO gate; B7 e2e snapshots first out of necessity. Trade-off: rollback baseline + doc-conformance vs constructor wiring + ~15 macro test updates + behavior change.
**Options:** (A) require snapshot → `SNAPSHOT_FAILED` before dispatch; (B) keep ungated, document deviation.
**Recommendation:** (A) is the spec-faithful choice and low-risk if e2e already snapshots first; approve only if Tech Lead accepts the ~15-test update + one commit. Otherwise (B) with a docs limitation.

**D-2 — R-17 stale worker / dead-owned-Excel health (`FE-08`).**
Evidence: `excel_health` trusts `get_status()`; worker only `NEEDS_RECOVERY` on timeout, not COM disconnect; prior B1 decision = document/hygiene (fresh server per e2e).
**Options:** (A) document + recovery runbook (close/reopen, restart server); (B) harden worker/`excel_health` with a bounded owned-PID liveness probe flipping `needs_recovery` (no COM touch) + unit tests.
**Recommendation:** (A) for this release (out of prior scope; worker lifecycle is High risk). (B) only if Tech Lead wants a code change and accepts H-risk worker edits + re-verification.

**D-3 — Residual interactive tokens `AppActivate` / `Application.Visible` (`FE-04`).**
Evidence: `INTERACTIVE_TOKENS` lacks `AppActivate`; `Application.Visible = True` not matched by `.Show`; docs already flag the residual.
**Options:** (A) add `AppActivate`, `Application.Visible`, and consider `.Activate`/`WindowState`; update tests/docs. (B) document residual only (current).
**Recommendation:** (A) — cheap, strictly tightens the veto (defense-in-depth), consistent with B5 MUST-fix direction. Low risk.

**D-4 — Expose `recover_session` / `close_session` as MCP tools? (`FE-11`, D-002).**
Evidence: `SessionService.recover_session`/`close_session` exist but are unregistered; `excel_health` reports `needs_recovery` with no client remedy except close/reopen.
**Options:** (A) add `excel_recover_session` tool (thin handler + tests + docs); (B) document close/reopen remedy.
**Recommendation:** (A) materially improves Full-Edit usability after a hang and is low/medium risk; otherwise (B). Either is acceptable; do not leave undocumented.

**D-5 — `classify_operation` stale/unwired table (`FE-06`).**
Evidence: table lists nonexistent tools, omits most; called only by unit tests; runtime gate is per-service `OperationClass`.
**Options:** (A) correct table additively to match 88-tool catalogue, keep non-runtime, update `test_macro_policy.py`; (B) add "not used at runtime" docstring, leave table.
**Recommendation:** (A) if low-effort (removes a misleading artifact); else (B). Never wire it as enforcement.

**D-6 — `pywinauto` dependency (`FE-07`, D-003).**
Evidence: declared in `pyproject.toml`, imported nowhere; UI automation via `comtypes`.
**Options:** (A) remove dependency + refresh `uv.lock` + fix release-notes row; (B) mark `# reserved` comment + release-notes note.
**Recommendation:** (A) — removes an unused dep (supply-chain hygiene); confirm no transitive reliance first.

**D-7 (scope guard) — tool count 88 vs plan's 87 (G-1).**
The corrective plan/audit (C-audit-14/23) says 87; current code is **88** (87 `add_tool` + `server_health`). Docs already say 88. **No product change**; FE-01 only aligns any residual "87" mentions in `.ai/` plan prose if Tech Lead wants plan-doc consistency (do **not** rewrite existing plans per constraints — record in this plan only). Confirm Auditor's live count matches 88 at Gate 0.

---

## 10. Risk register (orchestration-level)

| ID | Risk | Impact | Mitigation |
|---|---|---|---|
| R-FE-01 | Baseline report contradicts Planner pre-read (count/behavior) | Mis-scoped tasks | FE-00 Gate 0 reconciles; amend §2/§3 before any FE code |
| R-FE-02 | Two Implementors on one file (`server.py` in FE-02 vs others) | Merge conflict | §4.3 hard rule; FE-02 owns `server.py` boundary alone |
| R-FE-03 | Concurrent Excel/COM across FE tasks | Flaky/hung desktop | All integration serial `-p no:xdist`; one STA worker |
| R-FE-04 | FE-10 (snapshot) expands into macro-test churn | Schedule slip | Decision gate; ~15 tests scoped; e2e already snapshot-first |
| R-FE-05 | FE-08 harden touches worker lifecycle | High-risk regression | Default doc-only; harden only with Researcher+Auditor+Validator |
| R-FE-06 | Docs claim "verified" without live evidence | False FULL-EDIT-VERIFIED | Phase 3 observed-call rule; Auditor+Validator sign-off |
| R-FE-07 | `docs/03-Operations/*` + `docs/research/*` are deleted in working tree | Docs reference dead files | FE-01 fixes links to live docs; never restore deleted files |
| R-FE-08 | Unobservable focus-theft in Phase 5 | Over-claim | Mark `NOT VERIFIED` + evidence needed (spec §5 Phase 5) |
| R-FE-09 | pywinauto removal breaks lock/other dep | Build break | FE-07 refreshes `uv.lock`, runs smoke import + full unit subset |

---

## 11. Definition of Done (this orchestration)

- [ ] FE-00 baseline accepted (Gate 0); Auditor live count = 88 (+4 resources, 0 prompts) reconciled with G-1.
- [ ] D-1..D-6 recorded; every FE either `Done+commit` or `Deferred (doc)` with evidence.
- [ ] FE-12 matrix: every one of the 88 tools has an evidence-backed status; non-tool surfaces inventoried.
- [ ] FE-13: two edit/macro/edit/macro cycles on one `workbook_id` observed through real MCP; step 14 non-VBA mutation added; originals byte-identical.
- [ ] FE-14: hidden/no-focus/no-interference checks pass or explicitly `NOT VERIFIED` with evidence needed.
- [ ] FE-15: final Auditor + Validator pass; verdict recorded; capability report at `.ai/audit/reports/xlvba-mcp-full-edit-mode-capability-audit.md`.
- [ ] No existing plan/audit file overwritten; all new artifacts under `.ai/plans/` (this file) and `.ai/audit/reports/`.

**Final verdict vocabulary (spec §11):** `FULL EDIT MODE VERIFIED` / `PARTIALLY VERIFIED` / `NOT VERIFIED` / `FAILED`.
Never `FULL EDIT MODE VERIFIED` on registered tools / unit tests / service calls / Implementor assertion alone.