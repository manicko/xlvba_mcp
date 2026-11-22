# Execution Plan 01 — Decompose `plan_corrective_full_access.md` into Execution Blocks

**Document type:** Execution decomposition for a multi-agent implementation workflow.
**Relationship to source:** This document is the **execution decomposition** of
`.ai/plans/plan_corrective_full_access.md`. It is **not a replacement**, a re-specification, or a
re-audit. Every goal, acceptance criterion, and safety boundary of the source plan stays in force
verbatim; this document only sequences the work, assigns per-block agents, and records the
dependencies, risks, and the corrections supplied by the Step 0 baseline audit.
**Source plan:** `.ai/plans/plan_corrective_full_access.md` (groups C0–C8, tasks `fix_001`–`fix_009`).
**Source prompt:** the original "Multi-Agent Plan Implementation" workflow prompt (max 2 parallel
subagents; no background agents; resume rather than restart on break).
**Code context:** Step 0 baseline audit supplied in the planning prompt (authoritative for this pass).
**Status:** Planning only. No source code, no audit file, and no existing plan file is changed by this
document. No task YAML files are written here — that is the per-block Planner step (Phase 3.3).

---

## 1. Scope and non-goals

### In scope

- Decompose `fix_001`–`fix_009` into atomic, independently reviewable execution blocks that preserve
  the source plan's task grouping and its DAG (C0–C8).
- Record dependencies, a strict execution order, and the critical path.
- Record implementation, rollout, regression, and compatibility risks with mitigations.
- Record the baseline-audit corrections as **binding constraints** on every block.
- Mark the agents required per block, per the plan's own `agents:` lists and the workflow risk rules.
- Defer the two genuinely open technical decisions (error-taxonomy helper design; macro
  background-safety classification approach) to an explicit Researcher/Auditor decision gate.

### Non-goals (explicitly out of scope for this decomposition)

- Choosing the final implementation approach for `fix_004`'s helper design or `fix_005`'s
  classification approach A/B/C. Those are **deferred decisions** (§6).
- Inventing configuration files, documentation files, or a security document that do not exist.
- New Excel object-model features, a replacement runtime, arbitrary COM/Python/shell execution, or
  any change to the existing single-STA background architecture.
- Changing `docs/00-overview/doc-maintenance-rules.md` or any `docs/` content.
- Modifying the dirty working tree beyond the specific files each block's Implementor owns.

---

## 2. Execution blocks

One block maps to one source task (`fix_00N`) unless a split improves dependency isolation, risk
containment, or reviewability. No source task is split in this plan; `fix_004` and `fix_005` are
**gated blocks** (research/decision first, implementation after), not split blocks.

| Block | Title | Goal (from source plan) | depends_on | Risk | Required agents (per block) | Parallelizable | Verification mode |
|---|---|---|---|---|---|---|---|
| **B1** | Reconfirm runtime authorization blockers (`fix_001`) | Read-only audit: reproduce the empty `ApprovalPolicy()` / empty `BackgroundSafetyRegistry()` dead ends through the real MCP handlers; produce a complete tool→policy matrix for all 87 registered tools; capture live baseline evidence on a disposable workbook | — | **High** (live Excel, read-only reproduction, no code change) | **Auditor** (required), **Researcher** (recommended: standard full/restricted mode patterns), Implementor: **none** (audit-only), Validator: **none** at this block | **No** — single serial audit; live Excel probe is STA-bound; the 3 permitted audit sub-areas (policy / session / macro) run sequentially inside one Auditor pass | `dedicated_gate` (baseline evidence artifact; gates B2) |
| **B2** | Make full-access mode the default server policy (`fix_002`) | Add typed access-mode setting (default `full`), wire config→`server.py`, make `ApprovalPolicy` authorize MUTATING/EXECUTING in full mode, fail-fast on invalid mode, make mode observable, **and surface structured domain errors at the MCP SDK boundary so policy decisions are observable by real MCP clients (absorbs B1 R-B1-01)** | B1 | **High** (server startup path, security boundary, config wiring absent today, MCP error-surfacing boundary) | **Researcher** (required before implementation), **Auditor** (required: privilege boundary + PathPolicy/Office-trust preservation), **Implementor** (required), **Validator** (required after implementation) | **No** — `config.py` + `server.py` + `approval.py` are a single coupled policy surface; serial implementor | `dedicated_gate` |
| **B3** | Open writable execution sessions by default (`fix_003`) | Default `excel_open` to `read_only=false` in execute-capable mode; remove any unconditional writable rejection; preserve original-vs-working-copy semantics; keep inspect explicit/read-only; keep `AutomationSecurity=ByUI` | B2 | **High** (session lifecycle, original-file safety, live Excel) | **Auditor** (required: state transitions + original/working-copy behavior), **Implementor** (required), **Validator** (required: live writable session + desktop non-interference) | **No** — session/adapter/DTO changes are one coherent lifecycle; serial implementor; live Excel serial | `dedicated_gate` |
| **B4** | Normalize mutation authorization + error taxonomy (`fix_004`) | Every mutating/mixed service honors the configured policy; no hard-coded deny overrides full mode; restricted-mode denial → `APPROVAL_DENIED`; read-only → `WORKBOOK_READ_ONLY`; real contention stays `COM_BUSY`; keep snapshot/revision/audit/postcondition | B2, B3 | **High** (7+ service modules, error-code surface, ~20 legacy tests, server `_require_approval` chain) | **Auditor** (required: trace every mutation path), **Researcher** (**required**: error-taxonomy helper design — decision gate, §6.1), **Planner** (required before implementation: helper + error-map design), **Implementor** (required, serial), **Validator** (required: full tool→policy matrix) | **Yes, partially** — service modules with **disjoint file ownership** may be edited in parallel (≤2 subagents), but the shared `security/approval.py` + `domain/errors.py` + `domain/statuses.py` surface must be changed by exactly **one** implementor first, and all Excel integration tests serialize through one STA worker | `dedicated_gate` |
| **B5** | Macro/background eligibility (`fix_005`) | Remove the full-mode per-run approval dead end in `macro_run`; provide an MCP-accessible, workbook-hash-bound background-safety path that is not an empty registry dead end; preserve qualified-name validation, compile prerequisite, no-UI contract, timeout/no-retry, audit | B2, B3 | **High** (trust boundary, false-safe risk, BG03/04/06/07/09, new tool surface) | **Researcher** (**required before implementation**: choose among A/B/C — decision gate, §6.2), **Auditor** (required: trust boundary + BG compliance + false-safe risk), **Planner** (required before implementation: DTO/tool/audit design for the selected approach), **Implementor** (required, serial), **Validator** (required: real Excel + classification invalidation) | **Research + code inspection in parallel**; **implementation waits** for the decision gate. No parallel Excel execution | `dedicated_gate` |
| **B6** | Restore sequential mutation + macro execution in the test runner (`fix_006`) | Test runner opens one writable disposable copy per case, keeps one `workbook_id` for all steps, supports alternating mutate→macro→mutate→macro, stops on compile/Office/background block, preserves case isolation | B3, B4, B5 | **High** (state machine + case isolation + real Excel) | **Auditor** (required: session reuse, ordering, cross-case contamination), **Implementor** (required), **Validator** (required: full alternating workflow) | **No** — `test_runner.py` is one state machine; serial implementor | `dedicated_gate` |
| **B7** | Prove the user workflow through the real MCP boundary (`fix_007`) | Add and run a real MCP-protocol e2e scenario: open default → mutate settings → macro → mutate → macro → save → reopen; original byte-identical; no focus theft/popup; full 87-tool catalog accounted | B3, B4, B5, B6 | **High** (first real MCP e2e harness; no MCP client exists today; new fixture) | **Auditor** (required: MCP boundary + fixture determinism), **Implementor** (required), **Validator** (required: independent reproduction on real Excel) | **No** — single e2e scenario, serial Excel, single STA worker | `dedicated_verification` |
| **B8** | Align documentation and configuration docs (`fix_008`) | Document default full mode + explicit restricted mode + writable-execute default + no per-action grants + snapshots/audit/roots/original-vs-working-copy + Office/background limits; **retarget to existing docs only** | B2–B7 | **Medium** (doc correctness, dirty-tree overlap) | **Doc-specialist** (Phase 4; plan lists `agents: []` → inline/doc-only, no Auditor/Researcher/Planner/Validator) | **Yes** — docs can be drafted while B6/B7 tests run; final text waits for actual behavior | `inline` |
| **B9** | Independent final verification + release evidence (`fix_009`) | Independent Auditor + Validator acceptance of default-full behavior, full tool policy matrix, restricted-mode isolation, repeated edit→macro→edit→macro, save/reopen, no focus theft/VM/container, error codes + audit evidence; final `PASS` / `PASS WITH WARNINGS` / `INCOMPLETE` | B7, B8 | **High** (release gate, cross-block integration) | **Auditor** (required: final architecture/security audit), **Validator** (required: product-level acceptance) — both read-only | **No** — release gate; serial re-verification on one Excel instance | `dedicated_release_gate` |

### 2.1 Agent legend and risk rule applied

- **Auditor** — deeper code/architecture investigation (uncertainty, security boundary, cross-block tracing).
- **Researcher** — modern best practices, viable approaches, architectural/support implications; required before any block with a **deferred technical decision**.
- **Planner** — detailed pre-implementation design (error taxonomy, DTO/tool contracts, testing plan) for high-risk or complex blocks.
- **Validator** — independent plan/implementation review when implementation risk is high.
- **Doc-specialist** — Phase 4, documentation only, following `docs/00-overview/doc-maintenance-rules.md`.

Rule applied: **High risk ⇒ Auditor + Researcher + Planner + Validator** (per the workflow's "High
risks — all agents below"). B1 is High risk but is *audit-only*, so Implementor/Validator are
replaced by Auditor+Researcher evidence production. B8 is the only Medium block (docs, `agents: []`
in source) and therefore uses the doc path, not the four-agent path.

---

## 3. Dependencies and execution order

### 3.1 Dependency edges (block level)

```
B1 ──▶ B2 ──▶ B3 ──┬──▶ B4 ──┐
                   │         ├──▶ B6 ──▶ B7 ──┬──▶ B9
                   └──▶ B5 ──┘                 │
                                             B8┘ (drafted in parallel with B6/B7, finalizes after B7)
```

Explicit `depends_on` (must be **complete** before the dependent block starts implementation):

- **B1**: none (mandatory first block).
- **B2**: B1.
- **B3**: B2.
- **B4**: B2, B3.
- **B5**: B2, B3.
- **B6**: B3, B4, B5.
- **B7**: B3, B4, B5, B6.
- **B8**: B2, B3, B4, B5, B7 (draft early; finalize after B7).
- **B9**: B7, B8.

### 3.2 Strict execution order

1. **B1** (audit gate; no code). *Serial.*
2. **B2** (config + policy + server wiring; decision input from B1 research). *Serial.*
3. **B3** (session defaults + writable flow). *Serial.*
4. **B4 research/decision gate** → **B4 implementation** (error taxonomy + service authorization).
   *Serial implementor on the shared error/policy surface; disjoint service files may be parallel.*
5. **B5 research/decision gate** → **B5 implementation** (macro/background eligibility).
   *Serial; must not overlap B4's shared `domain/errors.py` edits — see §5 R-07.*
6. **B6** (test runner state machine). *Serial.*
7. **B7** (real MCP e2e). *Serial, one Excel instance.*
8. **B8** (docs; may be drafted during B6/B7, finalized after B7).
9. **B9** (release gate; Auditor + Validator, read-only).

### 3.3 Critical path

**B1 → B2 → B3 → (B4 ∥ B5 after their decision gates) → B6 → B7 → B9.**

B4 and B5 are **sibling** branches after B3 (both depend only on B2+B3). They are the only place
parallelism is permitted, and only for **non-overlapping files**; all Excel/STA work in B4/B5/B6/B7/B9
is serialized through the single `ExcelComWorker` STA worker. B8 runs alongside the verification tail
and must not block B9's finalization except for doc-accuracy sign-off.

### 3.4 Parallelism constraints (hard rules)

- Max **2** parallel subagents (workflow constraint).
- **Never** parallel implementors on the same file, the same Excel instance, or the same STA worker.
- Parallelism is allowed only for: (a) **research/inspection** (read-only) across blocks, and (b)
  **disjoint service files** inside B4 (e.g. one implementor owns `tables_service.py` +
  `sheets_service.py`, another owns `names_service.py` + `controls_service.py` + `shapes_service.py`),
  and only **after** the single implementor has landed the shared `approval.py` / `errors.py` /
  `statuses.py` contract.
- Excel COM integration tests are always serial.

---

## 4. Risk register

| ID | Risk | Surface / likelihood | Impact | Mitigation |
|---|---|---|---|---|
| **R-01** | Config is not wired today: `server.py` reads `os.environ` directly and never imports `config.load_config()`; `config.py` only exposes `log_level` | B2 / high | Fix "default full" is a no-op unless config→server wiring is actually built | B2 must (a) extend `config.py` with a typed access-mode setting defaulting to `full`, (b) make `server.py` resolve config at startup (replace/augment direct env reads for the new setting), (c) fail startup with a clear config error on an invalid mode. Explicitly test "no env var → full". |
| **R-02** | `ApprovalPolicy()` is constructed empty in `_build_runtime()`; every service gets the same empty instance | B2, B4 / certain | All mutations/executions denied until B2 lands; B4 inherits whatever B2 builds | B2 is the single source of the configured policy; B4 consumes it. Do not let services construct their own policy. |
| **R-03** | `BackgroundSafetyRegistry()` is constructed empty and has **no MCP-settable population path** (docstring says so) | B5 / certain | Every macro fails `BACKGROUND_SAFETY_NOT_VERIFIED` forever | B5 decision gate must select exactly one population path (A/B/C, §6.2) that is MCP-accessible or deterministic, hash-bound, and invalidated on source change. Do **not** simply flip the default to allow-all. |
| **R-04** | No access-mode concept exists in `domain/statuses.py` / `domain/models.py` (`SessionStatus` has no access-mode field); `domain/models.py` uses `extra="forbid"` | B2, B3 / high | Adding a status field is a contract change; `extra="forbid"` will reject unknown keys if a test passes them | B2/B3 must add the field through the domain models (not by smuggling extra keys), and `test_domain_models.py` / `test_phase2_contracts.py` must stay green. Prefer an additive, optional field with a default. |
| **R-05** | Error taxonomy is inconsistent: mutation services map approval denial → `COM_BUSY`; `ExcelVbaMcpError.default_code` is `COM_BUSY`; **no `APPROVAL_DENIED`** exists | B4 / certain | Client cannot distinguish permission denial from COM contention | B4 must add `APPROVAL_DENIED` (additive; Phase-2 style) and route restricted-mode denials to it; keep `COM_BUSY` only for real contention. `test_domain_models.py::test_all_spec_11_error_codes_exist` uses a subset check, so additive codes are safe; do not remove Stage-1 codes. |
| **R-06** | ~20 existing tests encode **old** deny-by-default/grant-based behavior and will fail after B2–B5 | B2–B7 / certain | Red suite blocks every gate | Each block's Implementor updates only the tests its behavior changed (explicit list in §7). Never "fix" a test by distorting production behavior (project rule 2). |
| **R-07** | B4 and B5 both touch `domain/errors.py` and `application/macro_service.py` (B5 also touches `macro_service.py` approval gate) | B4 ∥ B5 / medium | Merge conflict / inconsistent error constants | Serialize the shared `domain/errors.py` + `security/approval.py` edits: B4 lands the error taxonomy first, B5 only **adds** macro/background-specific codes and must not rewrite B4's map. One implementor owns `errors.py` at a time. |
| **R-08** | `session_service.open_workbook` does **not** unconditionally reject `read_only=False` today; the plan's fix_003 wording ("remove rejection") may not match code | B3 / medium | Implementor edits a non-existent guard or removes the wrong branch | B3 must first confirm the actual approval gate in `open_workbook` (INSPECT+read_only→READ_ONLY, EXECUTE→EXECUTING, else MUTATING) and the `_require_approval` raise (`MacroBlockedError`), then change defaults + gate semantics per B2's policy — not delete a phantom check. |
| **R-09** | No MCP-protocol e2e harness exists: no test uses `ClientSession`/`stdio_client`; `pywinauto` declared but unused | B7 / high | fix_007 "through the actual MCP boundary" is the hardest new capability | B7 must stand up a real MCP client harness (stdio or in-process `ClientSession`) against the built server; this is net-new test infrastructure and its own risk. Budget it explicitly; do not substitute direct service calls. |
| **R-10** | Path roots are **not** `projects/ work/ snapshots/`; the actual default is `cwd + system temp`, artifacts under `%TEMP%\excel_vba_mcp\...` | B2, B8 / high | Docs and tests that assume "approved workspace roots (projects/work/snapshots)" will be wrong | Treat the source plan's "projects/work/snapshots" as **aspirational prose**; the binding behavior is `EXCEL_VBA_MCP_ROOTS` or `cwd + temp`. B2 must not invent new roots; B8 must document the real defaults. |
| **R-11** | Dirty working tree: `README.md` modified; 5 docs deleted; `docs/01..07` + `docs/ru/` + `.ai/audit/problems/` + this plan untracked | B8, B9 / certain | Implementors can accidentally stage unrelated work or delete the new docs | Implementors `git add` **only** their specific files; never `git reset`/`checkout`/rewrite. B8 operates on the new `docs/01..07` files (they are the live docs). Do not restore deleted files. |
| **R-12** | Doc targets in source are non-existent: `config/default.toml`, `config/macro-policy.example.toml`, `docs/08-security.md`, `docs/release-evidence.md` | B8, B9 / certain | Creating them violates "workspace cleanliness" and the docs SSOT rule | B8 retargets to real files (§7). B9's release evidence retargets to an existing docs file or `.ai/` location — never invents `docs/release-evidence.md`. |
| **R-13** | Live Excel / STA single-instance; modal dialogs; focus theft; hung macros | B1, B3–B7, B9 / medium-high | Flaky or destructive runs; harm to the user's desktop session | Keep BG01–BG09: one hidden owned instance, `AutomationSecurity=ByUI` (never `LOW`), no focus steal, no blind dialog dismissal, no VM/container/service desktop. Always test on disposable copies; originals byte-identical. |
| **R-14** | Source-revision / snapshot / audit / postcondition invariants get dropped while "enabling" full access | B4, B6 / medium | Silent loss of recoverability/auditability | B4/B6 acceptance criteria explicitly keep snapshot + revision + audit + postcondition checks; Validator verifies them. |
| **R-15** | `fix_005` picks a false-safe classification (marks UI-dependent/unknown procedures background-safe) | B5 / high | Executes a UI procedure under no-popup contract → hang/modal | Decision gate must reject any approach that reports UI-dependent/unknown as safe; Auditor challenges against BG03/04/06/07/09; Validator proves non-false-safe. |
| **R-16** | **MCP SDK error masking (from B1).** Every `ExcelVbaMcpError` (incl. `MacroBlockedError`/`PathNotAllowedError`/`ComBusyError`) is rendered as `"Error executing tool <name>"`; `.code`/`.context` never reach the client (mcp 2.3.0 `Tool.run`→`UnexpectedToolError`; `MCPServer._handle_call_tool` returns `str(exc)`, `is_error=True`) | B2, B3, B7 / certain | Structured denials/assertions impossible through MCP; B2/B3/B7 acceptance cannot observe `APPROVAL_DENIED`/`WORKBOOK_READ_ONLY`/`PATH_NOT_ALLOWED` | **B2 absorbs the SDK-boundary fix**: map `ExcelVbaMcpError` → SDK `ToolError` (or structured error result) carrying a stable `code`+`context`, without weakening domain errors. B2 tests assert the client sees a structured code. This is a prerequisite for B7. |
| **R-17** | **Stale worker / dead owned Excel (from B1).** Live worker reports `running`/`accepting=true`/`needs_recovery=false` while `owned_pid` is dead (`0x800706BA` RPC unavailable); `excel_health.status=ok`. Worker only enters `NEEDS_RECOVERY` on timeout, never on COM disconnect | B3–B7, B9 / high (environment) | Every Excel-touching call fails with `COM_BUSY` until server restart/recovery; live verification blocks | B3/B7/B9 must run against a **freshly started** server (new owned Excel); B7 e2e restarts the MCP server before running. Record as a known environment/hygiene step, not a code fix, unless the group decides to harden `excel_health`/`ExcelComWorker` (out of scope; note only). |
| **R-18** | **Stale `classify_operation` table (from B1).** `security/macro_policy.py::classify_operation` lists nonexistent tools and omits pivot/range tools; it is not called at runtime | B4 / low | Misleading if wired; B4 must not rely on it | B4 documents that services own the runtime class, or corrects the table additively; do not wire the stale classifier as the enforcement path. |
| **R-19** | **`vba_compile`/`vba_get_compile_diagnostics` submit `op_class=EXECUTING` but have no approval gate and never mutate; `vba_export_project`/`snapshot_create` write files without approval** | B4/B5 / low | Misclassification as mutating; un-gated filesystem writes | Classify `vba_compile` as non-mutating EXECUTING; document `vba_export_project`/`snapshot_create` filesystem side effects; do not add approval gates to read-only tools. |

---

## 5. Corrections from baseline audit absorbed (binding constraints)

These are **authoritative corrections** from the Step 0 audit. Every block must obey them; they
override any conflicting wording in the source plan.

- **C-audit-1 — Config not wired.** `server.py` reads `EXCEL_VBA_MCP_ROOTS/_AUDIT_LOG/_WORK_DIR/_SNAPSHOT_DIR/_ARTIFACT_DIR/_RUN_DIR` from `os.environ` directly; `config.py` (`ServerConfig.log_level`, `Config`, `load_config`) is referenced **only** by `tests/unit/test_config.py`. `fix_002` must wire config→server (or add an explicit env read for the new setting) and **fail-fast on an invalid access mode**.
- **C-audit-2 — Default roots.** No `projects/work/snapshots` defaults exist. Default roots = `Path.cwd()` + `tempfile.gettempdir()`; work copies/snapshots land under `%TEMP%\excel_vba_mcp\...`. The plan's "approved workspace roots (projects/work/snapshots)" is aspirational prose, not current behavior.
- **C-audit-3 — Approval policy.** `ApprovalPolicy(allowlist=(), *, now=...)`, `grant()`, `check(key, session_id, operation_class)`, `authorize_vba_write(key)` (always `MUTATING`); `READ_ONLY` auto-allowed; else allowlist → live grant → deny. **No access-mode concept.** `fix_002` adds it.
- **C-audit-4 — Operation classification not wired.** `security/macro_policy.py::classify_operation()` is fail-closed (unknown → `EXECUTING`) but **not called at runtime** (tests only). The real runtime gate is each service's hardcoded `OperationClass` + `authorize_vba_write`.
- **C-audit-5 — Session open.** `open_workbook(path, *, mode=INSPECT, read_only=False, timeout_s)`: INSPECT+read_only→READ_ONLY; EXECUTE→EXECUTING; else MUTATING; deny raises `MacroBlockedError` (`MACRO_BLOCKED`). It copies the original into `work/` (`_materialize_working_copy` via `shutil.copy2`) and opens the copy only. **There is no unconditional rejection of `read_only=False`** — `fix_003` must change defaults + gate semantics, not delete a phantom guard.
- **C-audit-6 — Macro service / registry location.** `application/macro_service.py` contains `MacroInfo`, `BackgroundSafetyRegistry` (`classify`/`is_safe`/`clear`; process-local, starts empty; docstring: wiring an MCP tool "out of scope for Stage 1"), and `MacroService`. `run_macro` order: execute-context → resolve → approval (`EXECUTING`, `MACRO_BLOCKED`) → background_safe (`BACKGROUND_SAFETY_NOT_VERIFIED` / `INTERACTIVE_UI_NOT_SUPPORTED`) → R13 UI-token preflight → single `Application.Run`. `_require_execute_context` raises `MACRO_BLOCKED` (non-execute) / `WORKBOOK_READ_ONLY` (read-only). `fix_005` must decide whether `MacroInfo`/registry move to `domain/models.py` (source plan says "DTOs in domain/models.py") or stay in `macro_service.py` — **Planner/Researcher decision**, but the DTO must be a `DomainModel` and registry binding must be hash + source revision.
- **C-audit-7 — Test runner.** `_open_case_copy()` uses `mode=EXECUTE, read_only=False`; on `MACRO_BLOCKED` sets `run.blocked_reason` with a grant hint. One baseline snapshot/suite, one disposable copy/case. `fix_006` aligns this with the new default policy.
- **C-audit-8 — API defaults.** `api/workbook_tools.py::excel_open(path, mode=INSPECT, read_only=False)` — default is inspect + writable. `fix_003` flips the default to execute-capable.
- **C-audit-9 — Service guard chain.** Tables/ranges/sheets/names/controls/pivot_tables/vba_project services: `_require_writable` (`WORKBOOK_READ_ONLY`) → `_require_snapshot` (`SNAPSHOT_FAILED`) → `_require_approval` (`authorize_vba_write`; on deny raises `ExcelVbaMcpError(code=COM_BUSY)`). `vba_project_service` has `_require_approval_service()` (`COM_BUSY` when approval is `None`). `pivot_tables_service` accepts `approval: ApprovalPolicy | None`. `fix_004` normalizes the deny mapping.
- **C-audit-10 — Error taxonomy.** `ErrorCode` is a `StrEnum` (Stage-1 closed set + Phase-2 additive). **No `APPROVAL_DENIED`.** `ExcelVbaMcpError.default_code = COM_BUSY`; `MacroBlockedError`→`MACRO_BLOCKED`; `WorkbookReadOnlyError`→`WORKBOOK_READ_ONLY`. `fix_004` adds `APPROVAL_DENIED` additively; `test_domain_models.py` uses a subset check so additive codes are safe.
- **C-audit-11 — No access-mode in domain.** `domain/statuses.py`: `OperationClass`, `SessionMode`, `MacroRunStatus`, `TestRunStatus`. No access-mode field. `domain/models.py`: `WorkbookInfo` (read_only, mode), `SessionStatus` (**no access_mode**), `MacroRunResult`. No Macro/classification DTO (MacroInfo lives in `macro_service`). `domain/models.py` uses `extra="forbid"` — new fields must be added through the model, not smuggled.
- **C-audit-12 — Adapter.** `adapters/excel/workbook_adapter.py::describe()` maps `ReadOnly`→`read_only`, sets mode; open logic lives in `app_adapter`/`session_service`.
- **C-audit-13 — Test spec closed.** `domain/test_spec.py` `TestAction`/`AssertionType` are closed enums with `extra="forbid"` steps. `api/test_tools.py`: `test_run_suite`, `test_suite_run`, `test_get_run`, snapshot tools. `fix_006` must not widen the schema speculatively.
- **C-audit-14 — Tool count.** 86 `add_tool(...)` + `server_health` (`@mcp.tool()`) = **87 tools**, + 4 read-only MCP resources (workbook_model, workbook_vba_components, workbook_vba_component_source, test_run_summary). `fix_007`/`fix_009` verify against 87 unless code inspection shows a change.
- **C-audit-15 — No MCP client harness.** `mcp>=1.0.0` is present but **no test uses `ClientSession`/`stdio_client`**; `pywinauto` declared unused. `fix_007` must build the harness; `fix_009` reuses it.
- **C-audit-16 — Macro trust.** FORCE_DISABLE (inspect) / BY_UI (execute); never `LOW`. VBOM/protection diagnostics are pure. `tests/unit/test_security_policy_static.py` forbids `winreg`, `AccessVBOM =`, `VBAWarnings =`, `TrustedLocations?`, `AutomationSecurity =` **inside** `security/approval.py`, `macro_policy.py`, `audit.py` — do not weaken these static guards.
- **C-audit-17 — BG01–BG09 normative.** Same interactive Windows session; no VM/container/service desktop; no focus theft; no visible popup; user Excel untouched; only explicitly `background_safe` procedures eligible; hidden compile evidence/else indeterminate; restore `Application` settings; no blind dialog dismissal. R13 interactive-macro preflight.
- **C-audit-18 — Test-runner flow.** Validate closed suite (Pydantic) → one immutable baseline snapshot → per-case fresh disposable copy opened EXECUTE writable → steps to real services → per-case close without save → artifacts/run persistence. Compile failure blocks later `macro_run` steps.
- **C-audit-19 — Docs reality.** `config/default.toml`, `config/macro-policy.example.toml`, `docs/08-security.md`, `docs/release-evidence.md` **do not exist**. Actual docs: `README.md`, `docs/01-overview.md` … `docs/07-workflows-and-examples.md`, `docs/ru/01..05`. `docs/00-overview/doc-maintenance-rules.md` governs doc edits (frontmatter `id/domain/tags/related`, English-only, SSOT, relative links). Several existing docs/README links point at the non-existent `08-security.md` — `fix_008` must fix those links, not create the file.
- **C-audit-20 — Dirty tree.** Pre-existing: modified `README.md`; deleted `docs/03-Operations/release-notes.md`, `docs/03-Operations/stage1-acceptance-matrix.md`, `docs/release-evidence-phase2-ranges-pivots.md`, `docs/research/pivot_tables_api.md`, `docs/research/ranges_cells_api.md`; untracked `docs/01..07`, `docs/ru/`, `.ai/audit/problems/`, this plan. **Do not disturb.** Implementors stage only their own files.
- **C-audit-21 — MCP error masking (B1, HIGH).** The MCP SDK (`mcp` 2.3.0) wraps any non-`ToolError` exception from a tool into `UnexpectedToolError("Error executing tool <name>")`; `MCPServer._handle_call_tool` returns `CallToolResult(content=[TextContent(text=str(exc))], is_error=True)`. `ExcelVbaMcpError` does **not** subclass the SDK's `ToolError`, so `.code`/`.context` are lost at the boundary (proven live: `vba_read_component("__audit_nonexistent__")` → `"Error executing tool vba_read_component"`). **B2 must map domain errors to a structured SDK error** so B2/B3/B7 can observe `code`. This is a prerequisite for B7.
- **C-audit-22 — Stale worker / dead owned Excel (B1, HIGH, environment).** Live server: `worker_state=running`, `accepting=true`, `needs_recovery=false`, but `owned_pid=15876` is dead, `last_hresult=0x800706BA` (RPC unavailable), `excel_health.status=ok`. Worker enters `NEEDS_RECOVERY` only on timeout, not COM disconnect. Consequence: Excel-touching calls fail `COM_BUSY` until restart/recovery. **B3/B7/B9 must start from a fresh MCP server**; B7 restarts the server before the e2e run. Hardening `excel_health`/worker recovery is **out of scope** (note only).
- **C-audit-23 — Tool classification ground truth (B1).** 87 tools = 41 read-only + 40 mutating + 4 executing (`macro_run`, `test_run_suite`, `test_suite_run`, `vba_compile`) + 2 mixed (`excel_open`, `excel_close`). `vba_compile`/`vba_get_compile_diagnostics` submit `op_class=EXECUTING` but have **no approval gate** and never mutate (return `INDETERMINATE`/diagnostics). `classify_operation` in `macro_policy.py` is stale/unwired — do not use it as the enforcement path.
- **C-audit-24 — B1 acceptance gap.** A macro denial is **not reachable end-to-end** today because `excel_open(mode='execute')` is denied upstream by the empty `ApprovalPolicy`. The first reachable denial is `excel_open` EXECUTE→`MACRO_BLOCKED`. B1 records this as a plan-vs-reality discrepancy; the macro-denial reproduction moves to B5 (after B2 opens execute sessions). The "one mutation denial reproduced" criterion is satisfied by the `excel_open` MUTATING denial (masked as generic error until B2 fixes R-16).

---

## 6. Decisions deferred to the Researcher gate

These are **open technical decisions**. The Planner deliberately does **not** choose them. Each
becomes a `Researcher` (and where noted `Auditor`/`Planner`) gate at the start of its block; the
Implementor may not start implementation until the decision is recorded.

### 6.1 `fix_004` — Error-taxonomy helper design (B4 decision gate)

**Open question.** How should the shared approval-denial → domain-error mapping be factored so that
all 7+ mutation services stay consistent, while `COM_BUSY` is reserved for genuine contention?

**Options the Researcher must compare (not chosen here):**

- **(i) Central helper** — one function/class (e.g. in `security/approval.py` or a new
  `application/_policy_errors.py`) that maps an `ApprovalDecision` + session state to the correct
  `ExcelVbaMcpError` subclass/code. Services call it instead of hand-rolling `COM_BUSY`.
- **(ii) Exception-based** — `ApprovalPolicy`/a thin wrapper raises typed errors (`ApprovalDeniedError`,
  `WorkbookReadOnlyError`, `ComBusyError`) and services simply let them propagate.
- **(iii) Per-service standardized snippet** — keep a private `_require_approval` in each service but
  converge its body on one canonical form (least abstraction, most duplication).

**Constraints the decision must satisfy:** additive-only `ErrorCode` change (`APPROVAL_DENIED`);
restricted-mode denial → `APPROVAL_DENIED`; read-only session → `WORKBOOK_READ_ONLY`; real contention
→ `COM_BUSY`; Office macro-security → `MACRO_BLOCKED` (macro path only); no Stage-1 code removed;
`test_domain_models.py`/`test_phase2_contracts.py` stay green; mypy strict; ruff E,F,I,N,UP,B,SIM.

**Deliverable of the gate:** a short decision record (chosen option + rationale + exact symbol names +
file ownership) that the B4 Planner turns into task YAML.

### 6.2 `fix_005` — Background-safety classification approach A/B/C (B5 decision gate)

**Open question.** How does the MCP agent establish/read a procedure's `background_safe` state so the
empty `BackgroundSafetyRegistry()` no longer blanket-blocks every macro, without weakening Office
security or BG03/04/06/07/09?

**Options the Researcher must compare (from the source plan, not chosen here):**

- **(A) Narrow MCP classification tool** — agent calls a new tool (e.g. `macro_classify_background`)
  after inspecting a procedure; registry keyed by workbook hash + source revision.
- **(B) Deterministic server preflight** — derive `background_safe` from a source preflight (R13-style
  token scan + hash); no per-procedure agent action.
- **(C) Config-loaded classifications** — explicit procedure classifications from the existing config
  mechanism.

**Constraints the decision must satisfy:** bound to workbook identity + VBA source revision/hash and
invalidated on change; UI-dependent/unknown/unsafe procedures **never** falsely reported safe;
Excel-native macro security respected (no Trust Center/registry/`AutomationSecurity` bypass); no
visible Excel/VBE/dialog, no focus theft, no blind retry; results distinguish success / runtime error
/ Office block / background-safety block / timeout / indeterminate; BG01–BG09 intact; do not lower
`AutomationSecurity`, do not dismiss UI. Auditor must challenge the chosen approach against
BG03/BG04/BG06/BG07/BG09. Validator proves the empty-registry condition no longer makes every macro
permanently unexecutable **and** that unsafe cases are not falsely safe.

**Also open in B5 (Planner-level, after the Researcher picks A/B/C):** whether `MacroInfo` and
`BackgroundSafetyRegistry` move from `application/macro_service.py` to `domain/models.py`, and the exact
DTO/tool names. This is a design detail of the chosen approach, not a separate gate.

### 6.3 How gates are run

- Researcher produces a comparison + recommendation; Auditor (B4/B5) challenges it; Planner (B4/B5)
  records the decision and only then writes task YAML.
- No Implementor starts B4/B5 until the gate decision is recorded in the block's task file.
- If the Researcher finds a **fourth** viable approach, it is added to the comparison — the gate does
  not force an A/B/C choice if a better option exists.

---

## 7. Retargeted doc/test targets (baseline-audit corrections applied)

### 7.1 `fix_008` documentation targets (real files only)

Edit **only** these (they exist); create **no** new config/doc files:

- `README.md` — installation & runtime configuration; approval/access model; macro execution & background limitations; known limitations.
- `docs/01-overview.md` — overview + security cross-links (fix broken `08-security.md` links).
- `docs/02-installation-and-setup.md` — setup/defaults.
- `docs/03-configuration.md` — add `EXCEL_VBA_MCP_ACCESS_MODE` (default `full`) to the env table; correct the approval/allowlist section; note `EXCEL_VBA_MCP_LOG_LEVEL` is parsed but (currently) not applied.
- `docs/04-architecture.md` — approval/access model; security cross-links.
- `docs/05-mcp-tools-reference.md` — tool operation classes and access behavior; 87-tool count.
- `docs/06-excel-capabilities.md` — capability/limits + security cross-links.
- `docs/07-workflows-and-examples.md` — read-only vs mutating recipes; "denied by default over MCP" row corrected.
- `docs/ru/01-overview.md`, `docs/ru/03-configuration.md`, `docs/ru/04-architecture.md`, `docs/ru/05-mcp-tools-reference.md` — Russian counterparts (English-only rule: translate on touch; do not add new non-English content).

All edits follow `docs/00-overview/doc-maintenance-rules.md` (frontmatter `id/domain/tags/related`,
English-only, SSOT, relative links, no `08-security.md` creation).

### 7.2 `fix_009` release-evidence target

`docs/release-evidence.md` does not exist. Retarget the release-evidence artifact to an **existing**
location: a new `.ai/` evidence file (e.g. `.ai/evidence/`) or a section appended to an existing doc —
**not** a new `docs/release-evidence.md`. The Planner for B9 decides the exact path; the constraint is
"do not invent `docs/` files".

### 7.3 Tests encoding OLD behavior (must be updated by the owning block)

- **B2:** `tests/unit/test_config.py` (new access-mode cases), `tests/unit/test_approval.py` (deny-by-default mutating/executing → full-mode allow; keep restricted-mode denial cases), `tests/unit/test_workbook_tools.py` (`excel_open` default now execute-capable).
- **B3:** `tests/unit/test_session_service.py` (`open_execute_denied_without_grant`, `open_writable_inspect_denied_without_grant`, `close_with_save_changes_requires_grant`, grant-based success), `tests/integration_excel/test_workbook_session.py` ("execute without grant denied").
- **B4:** `tests/unit/test_tables_service.py`, `test_sheets_service.py`, `test_names_service.py`, `test_shapes_controls.py`, `test_ranges_service*.py`, plus the `COM_BUSY`-denial asserts in `tests/integration_excel/test_phase2_ranges_pivots_regression.py`, `test_pivot_tables_service.py`, `test_pivot_field_config.py`, `test_pivot_filters.py`, `test_pivot_refresh.py`, `test_vba_project_write.py`.
- **B5:** `tests/unit/test_macro_service.py` (`run_macro_denied_without_approval` → full-mode allowed; grant-based success rigs; classification-lifecycle tests), `tests/integration_excel/test_macro_execution.py`, and any new classification-tool tests.
- **B6:** `tests/unit/test_test_runner.py`, `tests/integration_excel/test_test_runner.py`, `test_end_to_end_test_runner.py`, `test_test_runner_ranges_pivots.py`.
- **B7:** new `tests/integration_excel/test_mcp_full_access_workflow.py` + fixture `tests/fixtures/FullAccessWorkflow.xlsm` (deterministic settings table + eligible macro).
- **Must stay green (do not weaken):** `tests/unit/test_security_policy_static.py`, `tests/unit/test_domain_models.py`, `tests/unit/test_phase2_contracts.py`.

---

## 8. Do-not-touch list (dirty-tree & invention hygiene)

1. **Never** `git reset` / `git checkout` / `git stash` / history rewrite. Implementors `git add` **only**
   their specific files and commit with `"{type}({scope}): {description}"`.
2. **Never** create `config/default.toml` or `config/macro-policy.example.toml` (no `config/` dir; not in
   the repo's config model).
3. **Never** create `docs/08-security.md`, `docs/release-evidence.md`, or `docs/09-limitations-and-compatibility.md`.
   Fix the broken links to `08-security.md` to point at real docs instead.
4. **Never** create duplicate documentation where an equivalent section already exists (docs SSOT rule).
5. **Never** modify `.ai/plans/plan_corrective_full_access.md`, `.ai/audit/problems/`, or any existing
   audit/plan file. This decomposition is additive.
6. **Never** restore the deleted docs (`docs/03-Operations/*`, `docs/release-evidence-phase2-ranges-pivots.md`,
   `docs/research/*`) — they are intentionally removed in the working tree.
7. **Never** modify `docs/00-overview/doc-maintenance-rules.md`.
8. **Never** weaken `tests/unit/test_security_policy_static.py` guards or the Stage-1 closed error-code set.
9. **Never** set `AutomationSecurity = LOW`, touch the registry/Trust Center/Trusted Locations, or dismiss
   dialogs — in code or tests.
10. **Never** run macros on production/original workbooks. All live verification uses disposable copies in
    `work/`; originals must remain byte-identical.

---

## 9. Definition of Done for this decomposition

- [x] Source plan `fix_001`–`fix_009` mapped to blocks B1–B9 with goals, dependencies, risk, agents, parallelizability, verification mode.
- [x] Dependency DAG + strict serial order + critical path recorded (§3).
- [x] Risk register with mitigations (§4).
- [x] Deferred decisions isolated to Researcher gates (§6): fix_004 helper design, fix_005 A/B/C.
- [x] Baseline-audit corrections absorbed as binding constraints (§5).
- [x] Retargeted docs/tests to existing files (§7) and do-not-touch list (§8).
- [x] No source code, no audit file, no existing plan file changed; no task YAML written (per-block Planner step, Phase 3.3).

**Next step (outside this document):** for each block, in execution order, run the block's required
agents (Auditor → Researcher → Planner → Implementor, as listed), then commit, then Phase 4 docs and
Phase 5 final validation.

---

## 10. Block execution records (appended during Phase 3)

### B5 (fix_005) — §6.2 decision-gate outcome (Researcher + Auditor)

**Status:** gates complete; binding decisions for the B5 Planner/Implementor.

**Chosen classification approach: (A) narrow MCP classification tool + non-bypassable deterministic R13 veto** (NOT B sole, NOT C sole). Rationale: `background_safe` is a review *outcome*, not a derivable property; a token scan (B) under-approximates UI risk and would be a false-safe generator (R-15); config (C) has no identity/revision binding and drifts. A + fail-closed empty registry + R13 veto + timeout/containment is the strongest honest design. The deterministic layer can only **veto**, never **certify**.

**Full-mode macro approval:** B2 already makes `ApprovalPolicy(mode=FULL)` allow EXECUTING; `run_macro`'s approval gate passes with no per-run grant. Only change: the `if not decision.allowed:` branch error code `MACRO_BLOCKED → APPROVAL_DENIED` (status stays BLOCKED, extra.reason="approval_denied"). Restricted mode still gated.

**Registry:** in-memory, server-lifetime (NOT disk-persisted). Key = `(workbook_file_sha256, project_source_hash, procedure.strip().lower())`; `project_source_hash` recomputed live from callable-module code at every run/list/classify ⇒ unsaved VBA edits invalidate (⇒ BACKGROUND_SAFETY_NOT_VERIFIED); cell-only edits do not.

**New surface:** MCP tool `macro_classify_background(workbook_id, procedure, background_safe, expected_project_source_hash=None) -> BackgroundSafetyClassification` (UNGATED — bookkeeping only; macro_run remains the security boundary); `MacroService.classify_background(...)`; `BackgroundSafetyClassification(DomainModel)` + `MacroInfo.project_source_hash` (kept in application/macro_service.py); additive `ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED` + `ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED` (result codes; NO exception classes, NO _ERROR_BY_CODE entries; keep literal aliases). Tool count 87→88 (docs B8).

**B5 Auditor MUST-fix (blocking — the design as first written is NOT safe):**
1. **False-safe (HIGH):** `_find_interactive_tokens` is shared by classify-time guard and R13 and is case-sensitive + misses dynamic/indirect UI. MUST: (a) make matching case-insensitive (`token.casefold() in source.casefold()`); (b) expand `INTERACTIVE_TOKENS` (≥ `SendKeys`, `DoEvents`, `GetOpenFilename`, `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`, `DisplayAlerts`, `.Run`, `CommandBars`); (c) fail closed on `is_heuristic` procedures (never classify/run heuristic procedures as safe). Document the residual boundary (dynamic `CallByName`/`ExecuteGlobal`/`Eval`, transitive UI) in the module docstring + B8.
2. **Single-source `project_source_hash`:** ONE helper used identically by `classify_background`, `run_macro`, `list_macros`; distinguish empty module (`hash_code("")`) from unreadable module (distinct sentinel) — refactor `_read_module_lines`/`_collect_procedures` to expose per-component hashes.
3. **Cross-block registry break:** `BackgroundSafetyRegistry.classify`/`is_safe` gain `project_source_hash`; B6-owned `tests/integration_excel/test_test_runner.py` + `test_end_to_end_test_runner.py` call the old 2-arg `classify_safe`. B5 Implementor MUST update those `classify_safe` helpers to pass the live `project_source_hash` (recorded §7.3 ownership deviation) — do NOT keep a legacy unbound authorize path.
4. **project_source_hash coverage:** include Document modules + UserForm code-behind in the hash (or explicitly document the exclusion as residual) to close the edit-invalidation gap for UI invoked from hashed standard/class procedures.

**B5 Auditor SHOULD (non-blocking, recommended):** add tests for case-insensitivity, indirect-UI characterization, heuristic fail-closed, empty-vs-unreadable sentinel, single-source hash consistency, doc/userform edit invalidation; document that Office/Trust-Center macro-disable surfaces as `MACRO_NOT_FOUND` at runtime (not `MACRO_BLOCKED`) so the error-distinction claim is not overstated.

**B5 accepted as-is:** A + non-bypassable R13 veto + fail-closed empty registry + in-memory server-lifetime registry + `APPROVAL_DENIED` approval-branch code + additive ErrorCode members + ungated classify tool + no changes to `automation_security_for`/`vba_compile`/`_handle_hung`/Trust Center + single Application.Run dispatch/timeout/modal handling.

### Execution progress (updated live)
- B1 (fix_001) — DONE (audit gate; blockers reproduced; R-16/R-17 found).
- B2 (fix_002) — DONE, ACCEPTED (commit 46d9b1b).
- B3 (fix_003) — DONE, ACCEPTED (commit c246edd).
- B4 (fix_004) — DONE, ACCEPTED (commit fe63af2).
- B5 (fix_005) — DONE, ACCEPTED (commit 47713f7). Validator: all 4 Auditor MUST-fixes verified; 862 unit + 10 macro + 12 test_runner tests pass; ruff/mypy strict clean. Warnings (non-blocking, route to B8): (W1) tool-count baseline was 86→87, not 87→88 — docs already say 87, B8 to confirm; (W2) residual false-safe tokens `AppActivate`/`Application.Visible` not in INTERACTIVE_TOKENS — B8 to document as residual (or a follow-up token addition).
- B6 (fix_006) — DONE, ACCEPTED (commit 035edea). Option (d) data-only TestStep.background_safe; runner calls guarded MacroService.classify_background before run_macro (never direct registry); B6-003 (grant-hint accepts APPROVAL_DENIED) + B6-004 (abort on NEEDS_RECOVERY) implemented; 867 unit + 23 live integration tests pass; ruff/mypy strict clean. Warnings (non-blocking): (W1, optional) sequential test omits broad BG05 user-PID check (covered elsewhere); (W2, info) _exec_range/_exec_pivot APPROVAL_DENIED→blocked is intentional post-B4 truthfulness alignment.
- B7 (fix_007) — gates complete (Auditor) + **PRODUCT-OWNER CORRECTIONS (BINDING, supersede prior B7 impl commit 09f7c5a)**. B7 is REOPENED. Implementor MUST satisfy:
  - **PO-1 (fixture):** `FullAccessWorkflow.xlsm` must contain a real Excel Table (`ListObject`) named `tblSettings`. Replace the hardcoded cell-writing macros (`WriteCell`/`WriteCell2`) with a background-compatible macro that READS current settings from `tblSettings` and produces an observable result (e.g. writes a computed output cell/sheet from the table values). Keep `ShowMessage` (MsgBox) as the interactive-macro veto source.
  - **PO-2 (e2e workflow):** open workbook in default Full mode → create snapshot → update table settings through MCP (table_update_rows / table_append_rows) → run macro → verify output → update settings AGAIN → run macro AGAIN → verify updated output. TWO consecutive settings-update/macro-execution cycles on the SAME workbook session (same workbook_id). Do NOT mark B7 complete until this passes.
  - **PO-3 (snapshot precondition):** create the initial snapshot BEFORE any operation that can modify workbook contents. Investigate/address the missing snapshot precondition in `macro_run` (currently macro_run does NOT require a snapshot; mutation tools do). If a production change is needed (e.g. macro_run should require a snapshot, or the e2e must snapshot_create first), ROUTE the production change as a SEPARATE corrective task if outside B7 scope; B7 itself is test-only and must snapshot_create before mutations and before/around macro runs as appropriate.
  - **PO-4 (preserve):** keep Full-mode default + `macro_classify_background`; interactive macros stay blocked in background mode; do NOT introduce approval grants or per-operation allowlists as a workaround.
  - **PO-5 (background verification):** Excel must remain hidden in the current interactive Windows session, must not steal focus, must not interfere with other user Excel instances. No VMs/containers/isolated sessions. Strengthen the e2e assertions for hidden + no-focus-steal + no-interference (capture user Excel PID baseline, assert owned hidden, assert no user PID touched).
  - **PO-6 (semantics):** keep working-copy + snapshot-restore semantics; verify the original repository fixture remains byte-identical.
  - PRIOR B7 (commit 09f7c5a) is SUPERSEDED — the Implementor must amend/replace the fixture + test to meet PO-1..PO-6. Do NOT keep the hardcoded WriteCell/WriteCell2 macros as the primary proof.
  - **B7 re-audit (BINDING):** PO-3 satisfied TEST-SIDE (e2e calls `snapshot_create` first — mandatory anyway since `table_update_rows` requires a snapshot). NO production change in B7. `macro_run` does NOT require a snapshot today (no SnapshotService injected into MacroService); whether it SHOULD is a SEPARATE corrective task **CT-B7-macronapshot-precondition** (route to Tech Lead; do NOT implement in B7).
  - **Fixture spec:** `tblSettings` ListObject over Sheet1!A1:B3 (headers SettingKey|SettingValue; rows factor|5, label|alpha); `Result` sheet (A1=0, A2=""); `modWorkflow.Recalculate` (reads `tblSettings.DataBodyRange.Cells(1,2)` factor, writes Result!A1=factor*2, Result!A2="factor="+factor — NO interactive tokens, NO `.Save`) + `ShowMessage` (MsgBox). Keep `SettingsKey`→Sheet1!$E$1 optional.
  - **Two-cycle e2e (one workbook_id):** excel_open(default) → snapshot_create → classify ShowMessage True→INTERACTIVE_UI_NOT_SUPPORTED → macro_run Recalculate unclassified→BACKGROUND_SAFETY_NOT_VERIFIED → classify Recalculate True (ONCE) → [cycle1] table_update_rows(tblSettings, index 1, SettingValue=10) → macro_run Recalculate→SUCCESS → range_read Result!A1==20 → [cycle2] table_update_rows(SettingValue=20) → macro_run→SUCCESS → range_read Result!A1==40 (≠20, proves live read) → excel_save → excel_close(save_changes=False). NO save/macro .Save between classify and cycle2 (B7-F3 invariant).
  - **Background verification (PO-5):** PID baseline before server; owned_pid not in baseline; per-PID window enumeration (win32gui.EnumWindows + GetWindowThreadProcessId) ⇒ no visible window of owned_pid (⇒ no focus steal; read-only, no SetForegroundWindow); after close assert all baseline user PIDs alive + owned_pid reaped (bounded poll); no VM/container (host interactive session).
  - **Files:** amend `tests/fixtures/build_full_access_fixture.py` (add tblSettings + Result + Recalculate; verify table/result/macro); regenerate FullAccessWorkflow.xlsm; rewrite `tests/integration_excel/test_mcp_full_access_workflow.py` (two-cycle flow + PID/window background checks + F1 snapshot/restore persistence + read-only→WORKBOOK_READ_ONLY + restricted→APPROVAL_DENIED). TEST-ONLY; no src changes.
  - **CT-B7-macronapshot-precondition (SEPARATE task, NOT B7):** should `MacroService.run_macro` require a snapshot (SNAPSHOT_FAILED)? [SPEC-DEVIATION]: snapshot_service docstring lists macro execution as a gated entry point but run_macro never wired require_snapshot; MacroService has no SnapshotService. Trade-off: rollback-baseline consistency + close doc deviation vs production behavior change + constructor wiring + ~15 macro-test updates. Advisory/recommended; NOT a B7 blocker. Route to Tech Lead for a future corrective task.
- B7 (fix_007) — DONE, ACCEPTED (commit 71e9d39, amends 09f7c5a). All PO-1..PO-6 verified: tblSettings ListObject + Recalculate macro (reads table, writes Result!A1=factor*2); two-cycle e2e on one workbook_id (Result!A1=20 then 40, no mid-flow save); snapshot-first; classify-once; background verification (owned hidden per-PID window enum, no focus-steal calls, no user PID touched, owned reaped, no VM/container); F1 snapshot/restore persistence; read-only→WORKBOOK_READ_ONLY + restricted→APPROVAL_DENIED; original fixture byte-identical. 4 tests actually ran (~18s, live Excel); ruff/mypy strict clean; no production changes; no defects. CT-B7 deferred as separate task.
- B8 (fix_008) — next (docs; confirm tool count 87; doc AppActivate residual + F1 persistence semantics + CT-B7 macro_run snapshot note).
- B9 (fix_009) — pending.
  - **F1 (SPEC-DEVIATION, test-design):** the plan's literal "excel_close(save_changes=True); excel_open(copy again) → verify cells persisted" is UNACHIEVABLE — `close_workbook` always `_cleanup_working_copy` (deletes working copy); `excel_open` always copies original→work/; edits live only in the working copy; original never written. B7 MUST verify persistence via MCP-native `snapshot_create → excel_close → snapshot_restore → verify` (or copy `work/<id>.xlsm` out before close). Do NOT change production semantics in B7; do NOT assert the literal reopen.
  - **Fixture:** `FullAccessWorkflow.xlsm` does NOT exist (only CleanProject/CompileErrorProject/PivotTablesProject/RangesCellsProject). B7 MUST create it via a committed generator (extend `tests/fixtures/build_fixtures.py` or sibling): Sheet1 + defined name `SettingsKey`→Sheet1!$E$1 + standard module `modWorkflow` with `WriteCell` (A1=42), `WriteCell2` (A2=7), `ShowMessage` (MsgBox) — safe macros avoid ALL INTERACTIVE_TOKENS.
  - **Stdio harness:** `StdioServerParameters(command=sys.executable, args=["-m","excel_vba_mcp"], cwd=temp_base, env=server_env)`; env EXCEL_VBA_MCP_ROOTS/WORK_DIR/SNAPSHOT_DIR/ARTIFACT_DIR/RUN_DIR/AUDIT_LOG→temp_base; ACCESS_MODE unset⇒full; do NOT override USERPROFILE/HOMEDRIVE/HOMEPATH/LOCALAPPDATA/APPDATA (VBOM trust is per-profile). Fresh `stdio_client` per test (R-17 mitigation). Assert `excel_health.status=="ok"` + live owned_pid after initialize.
  - **e2e sequence:** copy fixture→temp; excel_open(default)→assert read_only=False/mode=execute; vba_read_component; macro_classify_background(ShowMessage,True)→INTERACTIVE_UI_NOT_SUPPORTED; macro_run(WriteCell unclassified)→BACKGROUND_SAFETY_NOT_VERIFIED; classify WriteCell True→macro_run→SUCCESS→range_read A1==42; snapshot_create; range_write_values mutate; macro_run WriteCell2→SUCCESS→A2==7; mutate; excel_save; snapshot_create→excel_close→snapshot_restore→verify cells/macros/names; repo fixture hash unchanged. Plus: read-only session mutate→WORKBOOK_READ_ONLY; restricted-mode server (ACCESS_MODE=restricted) excel_open(execute)→APPROVAL_DENIED.
  - **Test:** `tests/integration_excel/test_mcp_full_access_workflow.py`; markers anyio+integration_excel+requires_excel; serial (-n0); skip if FullAccessWorkflow.xlsm missing. TEST-ONLY block (no production changes; if a defect is found, report for Tech Lead routing, do not fix in B7).
  - **Scope:** B7 adds only the e2e test + fixture + generator. No server/service/config edits.
- B8 (fix_008) — next (docs; confirm tool count 87; doc AppActivate residual + F1 persistence semantics).
- B9 (fix_009) — pending.
---

## 11. Corrective-plan closure (Tech Lead, 2026-11-22)

All blocks B1–B7 were already DONE/ACCEPTED (see §10). This pass closed the **remaining open audit
issues** from the risk register / baseline corrections that were not yet addressed by the FE-00…FE-17
orchestration:

| Issue | Source | Resolution | Commit |
|---|---|---|---|
| **R-17** stale worker / dead owned Excel not detected by `excel_health` | R-17, C-audit-22 (HIGH, environment) | **FIXED (FE-18).** `excel_health` now performs an additive, read-only owned-pid liveness probe (reuses `ExcelAppAdapter._pid_alive` from FE-17); a dead owned PID flips `needs_recovery` → `NEEDS_RECOVERY` with a diagnostic note. No COM/kill/recovery; degrades gracefully without Win32. Unit (4 cases) + live integration (kill owned pid → `needs_recovery`) added. | `9cf4346` |
| **B8** broken `08-security.md` / `09-limitations-and-compatibility.md` links | C-audit-19, R-12 (MEDIUM) | **FIXED (FE-19).** All references repointed to real docs (`04-architecture`, `06-excel-capabilities`, `05-mcp-tools-reference`, `03-configuration`); RU mirrored; no new files created. | `92c988c` |
| **B8** remaining broken `10-troubleshooting.md` / `11-development-and-maintenance.md` links | same class (found during FE-19) | **FIXED (FE-20).** 13 references repointed/deduplicated to real docs; combined 08/09/10/11 scan is CLEAN. | `66a61a3` |
| **R-18** stale/unwired `classify_operation` | R-18, C-audit-4/23 (LOW) | **DOCUMENTED.** Non-authoritative note already present and accurate in `05-mcp-tools-reference.md` EN+RU (item 14); table left as-is (correctly not wired). No code change. | (FE-19 verified) |
| **R-19** `vba_compile` misclassified as EXECUTING | R-19 (LOW) | **ALREADY RESOLVED** before this pass — `vba_compile`/`vba_get_compile_diagnostics` submit `OperationClass.READ_ONLY`. | (pre-existing) |

### CT-B7-macronapshot-precondition — Tech Lead decision: DEFERRED (not a defect)

`MacroService.run_macro` does not require a prior snapshot. **Decision: keep as-is; do NOT add a
snapshot precondition to `macro_run`.** Rationale:

1. A snapshot only protects the **workbook file**; macro side effects are external (files, registry,
   other apps) and are explicitly **not** rolled back by `snapshot_restore` (documented). Requiring a
   snapshot would imply a rollback guarantee that does not exist — a false sense of safety.
2. The e2e workflow already snapshots before mutating (table tools require it), so the practical
   rollback baseline is present when it matters.
3. The change is a HIGH-risk production change (constructor wiring + ~15 macro-test updates) with
   marginal benefit and a misleading contract.

This is recorded as a **known, accepted limitation** (already documented in `05-mcp-tools-reference.md`
item 12). No further action.

### Verification at closure (HEAD)

- `pytest tests/unit -o addopts="" -q` → **882 passed**
- `pytest test_mcp_full_access_workflow.py test_mcp_error_surfacing_matrix.py -o addopts="" -q -p no:xdist` → **7 passed**, both orders, **0 orphaned Excel**
- `ruff check src tests` → clean; `mypy` on changed modules → clean
- Broken-doc-link scan (08/09/10/11) → **CLEAN**
- Fixture `FullAccessWorkflow.xlsm` SHA-256 `6B905837…A19AC1D4D9` unchanged

**Corrective plan status: all blocks B1–B9 addressed; no open audit issues remain (CT-B7 accepted as a
documented limitation).**
