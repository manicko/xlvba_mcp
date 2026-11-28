# Execution Plan — `xlvba_*` Tool-Surface Consolidation (92 → ~41)

Source plan: `.ai/audit/reports/final-remediation-report-tool-consolidation.md` (WP0–WP14, §11 gates).
Template: `.ai/tasks/templates/task_template.yaml`.
Scope of this document: **decomposition only**. No source/test/doc/audit file is modified by the
Planner. Each block below is an independently executable, reviewable unit of work for one
Implementor, with explicit agent requirements, semantic scope, tests, and dependencies.

---

## 1. Overview

The remediation converts a 92-tool MCP surface into a ~41-tool surface (target range ~38–42, not a
hard number) by consolidating **non-destructive verbs per domain** into `action`-dispatch tools
(with `_read` companions where reads exist) while keeping **destructive / high-risk / lifecycle /
approval-sensitive operations standalone** with unchanged names and signatures.

This plan decomposes WP0–WP14 into **16 ordered execution blocks (B0–B15)**. Implementation is
**strictly sequential** (one Implementor at a time); the block order is the safe execution order.
`B13` (WP13 toolsets) is **DEFERRED**. `B14` is the documentation step; `B15` is final validation.

### Execution model (binding)

- **Tech Lead orchestrates only** — never implements.
- **One Implementor at a time.** Blocks are executed in the listed order; a block is committed and
  validated before the next begins.
- **Max 2 parallel subagents**, and only for *analysis/verification* within a block (e.g. Auditor +
  Researcher). No background agents. Implementation is never parallelized.
- **Resume old sessions on break** rather than starting fresh, to preserve context.
- **Throwaway files only in `.ai/tmp`** (project rule 17). Measurement scripts and PoC variants are
  never committed.
- **No dual-run, no temporary aliases** (prior no-backward-compatibility decision). Production
  `tools/list` goes directly 92 → final (~41); the ~121–130 window is a development-branch artifact
  only and must never reach production.

### Risk → agent-set legend

| Risk | Agent set (analysis phase) | Implementor | Validator |
|---|---|---|---|
| HIGH | Auditor + Researcher + Planner (≤2 parallel) | 1 (sequential) | required |
| MEDIUM | Planner (light) + Auditor spot-check | 1 (sequential) | required |
| LOW | Planner only | 1 (sequential) | required (inline) |
| GATE | Researcher + Auditor + Planner | none (decision artifact) | required |

---

## 2. Global constraints (apply to every block)

### 2.1 Project rules (`.kilo/rules/project.md`)
1. **English only** — all code, comments, logs, docstrings, docs.
2. **Production code is king** — if tests conflict with architecture/business logic, fix/remove tests.
3. **Strict separation of concerns** — UI/API/business/data boundaries preserved (thin async handlers
   → application services → single-STA `ExcelComWorker`).
4. **Single responsibility** — one tool = one cohesive domain; no mega `call_tool`.
5. **Avoid overengineering** — reuse existing handlers verbatim; dispatch only routes/forwards.
6. **Composition over inheritance.**
7. **Follow existing patterns** — `_require_*` guards, `set_*_service` injection, `add_tool`
   registration, existing DTOs (`RangeWriteRequest`, `FormulaWriteRequest`, `PivotFilterSpec`, …).
8. **Meaningful, consistent naming** — keep domain prefixes (`range_`, `pivot_`, `vba_`, `table_`,
   `sheet_`, `name_`).
9. **VBA identifier safety** — no VBA identifiers are introduced by this work; if any VBA export is
   touched (it is not, per scope), verify non-collision.
10. **VBA declaration placement** — N/A (no VBA edits).
17. **Workspace cleanliness** — throwaway scripts/variants only in `.ai/tmp`.

### 2.2 Hard constraints (non-negotiable, from the plan + Auditor)
- **H1 — Approval granularity preserved.** `macro_run`, `excel_save_as`, `table_delete_column` stay
  **standalone with unchanged names/signatures** (Kilo permissions bind to `{server}_{tool}`, not to
  an `action` value). `workbook` `close` gets an explicit, tested disposition (standalone or
  protected action). No consolidation may merge two distinct permission profiles.
- **H2 — No dual-run.** Old names are not registered alongside new ones; no `meta.deprecated` alias
  trick; single switchover at B12.
- **H3 — Truthful annotations from the truth matrix.** Hints computed from the WP0 per-action truth
  matrix using MCP default semantics (`readOnly=false`, `destructive=true`, `idempotent=false`).
  Absence is **not** neutral.
- **H4 — `range_structural` is NOT idempotent** (re-`insert` duplicates; `unmerge` discards merges).
- **H5 — Handlers reused verbatim.** Consolidation is registration + dispatch + validation only;
  runtime behavior unchanged.
- **H6 — No general `xlvba_batch`** (rollback-less footgun). Reuse `test_suite_run` isolation if
  batch is ever needed.

### 2.3 Auditor corrections folded in
- **A1 — `xlvba_environment` inline annotation.** It is registered via `add_tool` with an **inline**
  `ToolAnnotations(read_only_hint=True)` and is **NOT** a member of the `_annotations(name)`
  frozenset resolution. B11 (WP10) must preserve/verify the inline handling and the contract test
  must reflect that it does not flow through `_ANN_READ_ONLY`.
- **A2 — `None`-vs-`False` default semantics.** `_annotations` currently emits `None` for absent
  hints (clients then apply MCP default `destructiveHint=true`). B11 must **decide and document**
  whether to emit explicit `False` for additive/read-only tools or rely on `None`→default; the truth
  matrix drives the choice. Recorded as a WP10 decision (not a WP14-style gate).
- **A3 — Docs `ru/` mirror + stale count.** `docs/05-mcp-tools-reference.md` is stale (says 91) and a
  full Russian mirror `docs/ru/05-mcp-tools-reference.md` must be kept in sync. B14 reconciles
  91 → 92 → final count in both.
- **A4 — Docs path mismatch.** `docs/00-overview/doc-maintenance-rules.md` references a non-existent
  `docs/02-build-and-deployment/vba-build.md`; actual docs are flat `docs/01-overview.md` …
  `docs/11-development-and-maintenance.md` plus `docs/ru/`. B14 records/corrects this reference
  defect within doc scope (no audit-file edits).

### 2.4 Product Owner decisions (binding — recorded 2026-11-27)
- **D1 — `workbook` `close` unsaved-change disposition = Option A (refuse by default).** Resolves
  WP1 "verify". A clean workbook closes normally. If there are unsaved changes, `close` **refuses by
  default**; discarding requires explicit `discard_changes=true`. Preserve an explicit save-on-close
  option if the existing implementation supports it (e.g. `save_changes=true`); **reject contradictory
  flags** (e.g. `discard_changes=true` together with `save_changes=true`) with a structured
  `ExcelVbaMcpError`. Enforced in the handler; covered by per-action tests in B2. Recorded in the B0
  security/approval matrix disposition for `close`.
- **D2 — WP13 toolset layer = DEFERRED (Option A).** Ship the complete ~41-tool default surface only.
  B13 remains deferred; revisit only when a second client profile with genuinely different domain
  requirements appears. No `--toolsets` / `XLVBA_TOOLSETS` work now.

### 2.5 B0 resolution decisions (Tech Lead, binding for B2–B12)
Recorded from the B0 Researcher findings (Q1–Q12). These refine the plan's acceptance criteria where
the matrix proved a different truth; the user's Product Owner choices (D1/D2) take precedence.
- **R-Q1 — `workbook` `close` stays inside `workbook` (PO chose Option A, not extraction).** The
  truthful tool triple for `workbook` is **`(readOnly=false, destructive=true, idempotent=false)`**
  because the `discard_changes=true` path is destructive. B2 acceptance "workbook not destructive" is
  **amended** to: "workbook carries a destructive `close` action (discard path) but `close` refuses by
  default; `open`/`save` are non-destructive." Do NOT extract `workbook_close` (PO rejected Option C).
- **R-Q2 — `test` stays a single tool (plan WP6).** `test(run-suite|get-run)` merges executing
  (`P_EXEC`) and read (`P_READ`) server-side classes; this is an **accepted, documented H1 deviation**
  (run-suite's real approval gate is `macro_run`'s `EXECUTING` class + the `xlvba_*` name glob). Record
  it in the B0 security matrix as an intentional mixed-profile tool. Do NOT extract `test_get_run`.
- **R-Q3 — `vba_replace_component_source` is `(F,T,F)` (NOT idempotent).** Re-call with the same
  `expected_source_hash` fails the revision guard ⇒ not idempotent in effect. The plan §4.4 dual-hint
  claim for this tool is **corrected**: it is destructive-only.
- **R-Q4 — Dual-hint tools are exactly `{range_merge, pivot_field}`.** `range_merge` `(F,T,T)`
  intentional; `pivot_field` `(F,T,T)` matrix-proven via `field-set-layout(delete=True)`. No other tool
  carries both. B11 exact-triple test asserts exactly this set.
- **R-Q5 — `shape`/`control` are `(F,F,T)`** (all actions idempotent per matrix). B6's "default false"
  parenthetical is corrected; the matrix is authoritative (H3).
- **R-Q6 — `range_write` idempotency verified in B8.** Default `(F,F,F)` because
  `set-conditional-formatting(mode="add")` appends a rule (renumbers priorities) ⇒ not idempotent.
  B8 verifies against `ranges_service.set_conditional_formatting`; B11 locks the verified triple.
- **R-Q7 — `test` is `(F,T,F)`** (run-suite can dispatch `macro_run` steps ⇒ external side effects ⇒
  destructive).
- **R-Q8 — B2 MUST wire dirty detection.** `record.dirty` is never set `True` today; `WorkbookAdapter.is_dirty`
  (`not wb.Saved`) exists but is unused. B2 must use `is_dirty(wb)`/`wb.Saved` (not `record.dirty`) to
  implement refuse-by-default `close`.
- **R-Q9 — B2 MUST rewrite `_DESC_CLOSE`** (currently "unsaved edits are lost") to the D1 refuse-by-default
  semantics, preserving the §5.1 template markers asserted by `test_description_contract.py`.
- **R-Q10 — Keep `xlvba_environment` inline** (`ToolAnnotations(read_only_hint=True,...)` at its `add_tool`
  call); do NOT route it through `_ANN_READ_ONLY`. `server_health` stays inline. B11 preserves this (A1).
- **R-Q11 — Final tool count = 41** (Q1 kept close in `workbook`, Q2 kept `test`); B12
  `test_tool_count_is_final_exposed` asserts 41; docs reconcile 91 → 92 → 41.
- **R-Q12 — Emit explicit booleans, never `None`.** `_annotations` and the inline annotations emit explicit
  `True`/`False` for all three hints (A2). This fixes 20+ additive tools currently mis-advertised as
  destructive under MCP `destructiveHint=true` default.

### 2.6 B1 gate decision (WP14 schema PoC — recorded 2026-11-27)
- **Decision = option (a): flat args + handler validation.** Evidence (mcp 2.3.0, real in-memory
  `ClientSession`↔`MCPServer` round-trip, `.ai/tmp` only): public `add_tool` has no `inputSchema` param;
  variant (b) `oneOf` is reachable only via unsupported low-level `Tool`/`ToolManager` injection, and
  `Tool.run` validates via the flat `fn_metadata` (not `parameters`) ⇒ advertised `oneOf` diverges from
  enforced contract. Real-protocol arg-error rate and first-call success were **identical** for (a) and (b).
  Therefore per-action required-param enforcement stays **handler-side** in B2–B10, backed by per-action
  validation tests + `action` enum/description docs. Residual risk (schema shows per-action fields optional)
  is accepted and mitigated as above. This is specific to mcp 2.3.0 and does not generalize to MCP/other SDKs.

---

## 3. Ordered execution blocks

> Blocks are executed **in order**. `depends_on` encodes both logical dependency and the
> one-Implementor sequencing. Disjoint-module note: B2–B10 touch disjoint `api/*_tools.py` modules,
> but remain sequential because only one Implementor runs at a time.

### B0 — WP0: Baseline, mapping ledger, security/approval matrix, annotation truth matrix
- **Risk:** HIGH (foundational; all later blocks verify against these artifacts).
- **Agents:** Auditor + Researcher + Planner (analysis, ≤2 parallel); Implementor (artifact author);
  Validator (completeness).
- **Depends on:** none.

```yaml
id: B0_wp0_baseline
title: WP0 — Freeze baseline numbers and produce the four baseline artifacts
priority: high
depends_on: []
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B0 — WP0 Baseline
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP0 — Baseline: measurement, mapping ledger, security/approval matrix, annotation truth matrix"
description: >
  Produce the four "before" artifacts that every later block is verified against:
  (1) serialized tools/list token measurement, (2) total old→new mapping ledger (92 rows),
  (3) per-tool security/approval matrix (92 tools), (4) per-action annotation truth matrix.
  No production behavior change. Reconcile 92 = 91 add_tool + 1 @mcp.tool (server_health).
goals:
  - record baseline count (92) and serialized tools/list byte size in .ai/tmp
  - map every one of the 92 current tool names to exactly one defined target
  - record risk + Kilo approval policy + new name + permission profile for all 92 tools
  - record actual (readOnly, destructive, idempotent) per action of every new tool
  - prove no consolidation merges two distinct permission profiles
files:
  - path: .ai/tmp/  (throwaway measurement script + baseline output; NOT committed)
  - path: tests/unit/test_description_contract.py
    targets:
      - type: module
        name: test_description_contract
    semantic_anchors:
      insert_after:
        type: function
        value: test_tool_count_is_92_exposed
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
      - type: function
        name: server_health
    semantic_anchors:
      insert_before:
        type: function
        value: _register_tools
changes:
  - action: add_throwaway_measurement
    description: >
      Add a .ai/tmp script that serializes server.mcp.list_tools() (model_dump by_alias) and
      prints count + byte size; run before any surface change.
    code_hint: |
      # .ai/tmp only — never committed
      # count == 92 ; bytes == len(json.dumps(data, default=str).encode())
  - action: add_baseline_artifacts
    description: >
      Commit data fixtures (or test fixtures) for: mapping ledger (92 rows), security/approval
      matrix (92 tools), annotation truth matrix (every action of every new tool).
    code_hint: |
      ledger row: {old_name -> target: {tool, action} | standalone: name}
      matrix row: {tool, risk, kilo_policy, new_name, permission_profile}
      truth row:  {tool, action, read_only, destructive, idempotent}
acceptance_criteria:
  - baseline count 92 and serialized byte size recorded in .ai/tmp
  - mapping ledger has 92 rows; every row resolves; no duplicates/orphans
  - security/approval matrix covers all 92 tools; the four high-risk tools
    (macro_run, excel_save_as, table_delete_column, workbook close) have explicit disposition
  - annotation truth matrix covers every action of every new tool; range_structural marked not idempotent
  - tests assert ledger totalness (len == 92) and matrix completeness
```

- **Architectural constraints:** read-only w.r.t. production code; artifacts are data/fixtures only.
- **Required tests:** ledger totalness (`len == 92`), matrix completeness, truth-matrix coverage.
- **Acceptance:** as YAML. This block **closes gates §11.1–§11.3** (approval matrix, final map, truth matrix) except the schema PoC (§11.4 → B1) and migration/benchmark (§11.5–§11.6 → B12/B15).

---

### B1 — WP14: Schema PoC decision gate (flat args vs action-specific `inputSchema`)
- **Risk:** GATE (blocks B2; must not be foreclosed).
- **Agents:** Researcher + Auditor + Planner (no production Implementor).
- **Depends on:** B0.

```yaml
id: B1_wp14_schema_poc_gate
title: WP14 — Empirically decide per-action required-param mechanism before WP1
priority: high
depends_on:
  - B0_wp0_baseline
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B1 — WP14 Schema PoC Gate
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP14 — Schema PoC gate (flat args vs action-specific JSON Schema)"
  - "3. Per-action required-param strategy (the crux)"
description: >
  Build two throwaway variants in .ai/tmp (never committed) and compare them through
  SDK -> MCP -> Kilo -> a real LLM: (a) flat args + handler validation;
  (b) action-specific inputSchema via JSON Schema 2020-12 oneOf (or the Python SDK
  low-level custom-schema hook). Decide on evidence. Do not generalize one SDK's
  limits to MCP. The decision is recorded, not assumed.
goals:
  - measure argument-error rate and first-call success for (a) and (b)
  - record whether (b) is achievable through the PUBLIC add_tool surface or requires unsupported internals
  - document the decision with evidence
files:
  - path: .ai/tmp/  (two throwaway PoC variants + comparison output; NOT committed)
changes:
  - action: build_poc_variants
    description: >
      Variant (a): flat signature dispatch tool. Variant (b): action-specific oneOf schema.
      Run both through the real client path; capture argument errors + first-call success.
    code_hint: |
      # decision matrix: {variant, arg_error_rate, first_call_success, client_support, uses_unsupported_internals}
  - action: record_decision
    description: >
      Write a short comparison + decision into the plan/PR notes. If (b) yields no measurable
      gain or poor client support, keep (a).
acceptance_criteria:
  - comparison of argument-error rate and first-call success for both variants recorded
  - documented decision; if (b) not viable via public API, (a) is confirmed with rationale
  - decision referenced by B2..B10 (per-action enforcement stays handler-side under (a))
```

- **Decision-gate framing (do not foreclose):** The Auditor's verified finding is that public
  `MCPServer.add_tool` exposes **no** `inputSchema`/custom-schema parameter, so option (b) `oneOf`
  would require unsupported SDK internals — making (a) the only *fully-supported* path today. This is
  a strong input, **not** the decision. B1 must still run the PoC and record the evidence-based
  choice. **The Planner does not select a/b here.**
- **Required tests:** none in production; PoC output is the artifact.
- **Acceptance:** as YAML. **Closes gate §11.4.**

---

### B2 — WP1: `workbook` + `workbook_read`; standalone `excel_save_as`; `close` disposition
- **Risk:** HIGH (unsaved-change loss; `excel_save_as` permission).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B0, B1.

```yaml
id: B2_wp1_workbook
title: WP1 — Collapse non-destructive workbook verbs; keep save-as standalone; resolve close
priority: high
depends_on:
  - B0_wp0_baseline
  - B1_wp14_schema_poc_gate
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B2 — WP1 workbook
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP1 — workbook + workbook_read; standalone excel_save_as"
description: >
  Add workbook(action: open|close|save) and workbook_read(action: status|list|model)
  dispatching to the existing handlers unchanged. Keep excel_save_as standalone (own
  permission/overwrite policy). Verify excel_close unsaved-change behavior and add an
  explicit disposition (discard_changes flag enforced in handler, default = refuse, or
  extract workbook_close standalone if a distinct permission is required).
goals:
  - register workbook, workbook_read, excel_save_as
  - workbook is not destructive and does not carry save-as
  - excel_save_as keeps destructive_hint=True and its overwrite policy
  - close disposition documented and tested
files:
  - path: src/excel_vba_mcp/api/workbook_tools.py
    targets:
      - type: function
        name: excel_open
      - type: function
        name: excel_close
      - type: function
        name: excel_save
      - type: function
        name: excel_save_as
      - type: function
        name: excel_get_status
      - type: function
        name: excel_list_workbooks
      - type: function
        name: workbook_get_model
    semantic_anchors:
      insert_after:
        type: function
        value: workbook_get_model
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      Register workbook(action Literal["open","close","save"]) and
      workbook_read(action Literal["status","list","model"]) via add_tool; route to existing
      handlers; forward validated args only.
    code_hint: |
      workbook(action, ...) -> excel_open|excel_close|excel_save
      workbook_read(action, ...) -> excel_get_status|excel_list_workbooks|workbook_get_model
  - action: close_disposition
    description: >
      Confirm excel_close can discard unsaved changes (default save_changes=False). Add explicit
      discard_changes flag (default refuse) enforced in handler + per-action test, OR extract
      workbook_close standalone if a distinct permission is required.
    code_hint: |
      # disposition recorded in security/approval matrix (B0)
acceptance_criteria:
  - workbook, workbook_read, excel_save_as registered
  - workbook not destructive; no save-as in workbook action set
  - excel_save_as keeps destructive_hint=True + overwrite policy + unchanged signature
  - close disposition documented and tested (refuse-by-default or standalone)
```

- **Architectural constraints:** H1 (save-as + close), H5 (handlers reused).
- **Required tests:** per-action validation (missing required arg → structured `ExcelVbaMcpError`);
  `close` refuse-by-default test; `excel_save_as` overwrite-policy test.
- **Acceptance:** as YAML.

---

### B3 — WP2: `sheet` + `sheet_read`; standalone `sheet_delete`
- **Risk:** MEDIUM (one destructive extraction).
- **Agents:** Planner (light) + Auditor spot-check; Implementor; Validator.
- **Depends on:** B2 (sequential).

```yaml
id: B3_wp2_sheet
title: WP2 — Collapse 6 non-destructive sheet verbs; extract destructive sheet_delete
priority: medium
depends_on:
  - B2_wp1_workbook
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B3 — WP2 sheet
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP2 — sheet + sheet_read; standalone sheet_delete"
description: >
  Add sheet(action: add|rename|copy|set-visibility) and sheet_read(action: list|get).
  Keep sheet_delete standalone with unchanged signature.
goals:
  - register sheet, sheet_read, sheet_delete
  - sheet not destructive; sheet_delete destructive
files:
  - path: src/excel_vba_mcp/api/sheet_tools.py
    targets:
      - type: function
        name: sheets_list
      - type: function
        name: sheet_get
      - type: function
        name: sheet_add
      - type: function
        name: sheet_rename
      - type: function
        name: sheet_copy
      - type: function
        name: sheet_delete
      - type: function
        name: sheet_set_visibility
    semantic_anchors:
      insert_after:
        type: function
        value: sheet_set_visibility
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      sheet(action Literal["add","rename","copy","set-visibility"]) +
      sheet_read(action Literal["list","get"]); sheet_delete standalone unchanged.
    code_hint: |
      sheet(action, ...) -> sheet_add|sheet_rename|sheet_copy|sheet_set_visibility
      sheet_read(action, ...) -> sheets_list|sheet_get
acceptance_criteria:
  - sheet, sheet_read, sheet_delete registered
  - sheet not destructive; sheet_delete keeps destructive_hint=True + unchanged signature
```

- **Required tests:** per-action validation for `sheet` required args; `sheet_delete` destructive hint.
- **Acceptance:** as YAML.

---

### B4 — WP3: `table` + `table_read`; standalone `table_delete_rows` **and** `table_delete_column`
- **Risk:** HIGH (two standalone destructive deletions; XOR guard; `table_delete_column` permission).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B3 (sequential).

```yaml
id: B4_wp3_table
title: WP3 — Collapse non-destructive table verbs; keep both deletions standalone
priority: high
depends_on:
  - B3_wp2_sheet
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B4 — WP3 table
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP3 — table + table_read; standalone table_delete_rows and table_delete_column"
description: >
  Add table(action: create|resize|append-rows|insert-row|update-rows|add-column) and
  table_read(action: list|get|read-rows). Keep table_delete_rows standalone (XOR guard
  unchanged) and table_delete_column standalone (own permission). delete-column is removed
  from the table action set.
goals:
  - register table, table_read, table_delete_rows, table_delete_column
  - table not destructive and has no delete-column
  - both deletions destructive with unchanged signatures + guards
files:
  - path: src/excel_vba_mcp/api/table_tools.py
    targets:
      - type: function
        name: tables_list
      - type: function
        name: table_get
      - type: function
        name: table_read_rows
      - type: function
        name: table_append_rows
      - type: function
        name: table_insert_row
      - type: function
        name: table_update_rows
      - type: function
        name: table_delete_rows
      - type: function
        name: table_add_column
      - type: function
        name: table_delete_column
      - type: function
        name: table_create
      - type: function
        name: table_resize
    semantic_anchors:
      insert_after:
        type: function
        value: table_resize
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      table(action Literal["create","resize","append-rows","insert-row","update-rows","add-column"])
      + table_read(action Literal["list","get","read-rows"]); both deletions standalone.
    code_hint: |
      table(action, ...) -> table_create|table_resize|table_append_rows|table_insert_row|table_update_rows|table_add_column
      table_read(action, ...) -> tables_list|table_get|table_read_rows
      # table_delete_rows XOR (indices vs where) unchanged; table_delete_column unchanged
acceptance_criteria:
  - table, table_read, table_delete_rows, table_delete_column registered
  - table not destructive; no delete-column action
  - both deletions destructive, signatures unchanged
  - per-action tests for XOR guard and column delete
```

- **Architectural constraints:** H1 (`table_delete_column`), H5; XOR guard (`indices is None) == (where is None)`) preserved verbatim.
- **Required tests:** XOR-guard per-action test; `table_delete_column` `header` required test; `table` not destructive.
- **Acceptance:** as YAML.

---

### B5 — WP4: `name` + `name_read`; standalone `name_delete`
- **Risk:** MEDIUM.
- **Agents:** Planner (light); Implementor; Validator.
- **Depends on:** B4 (sequential).

```yaml
id: B5_wp4_name
title: WP4 — Collapse name create/update; extract delete
priority: medium
depends_on:
  - B4_wp3_table
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B5 — WP4 name
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP4 — name + name_read; standalone name_delete"
description: >
  Add name(action: create|update) and name_read(action: list|get). Keep name_delete standalone.
goals:
  - register name, name_read, name_delete
  - name not destructive; name_delete destructive
files:
  - path: src/excel_vba_mcp/api/name_tools.py
    targets:
      - type: function
        name: names_list
      - type: function
        name: name_get
      - type: function
        name: name_create
      - type: function
        name: name_update
      - type: function
        name: name_delete
    semantic_anchors:
      insert_after:
        type: function
        value: name_update
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      name(action Literal["create","update"]) + name_read(action Literal["list","get"]);
      name_delete standalone unchanged.
    code_hint: |
      name(action, ...) -> name_create|name_update
      name_read(action, ...) -> names_list|name_get
acceptance_criteria:
  - name, name_read, name_delete registered
  - name not destructive; name_delete destructive
```

- **Required tests:** per-action validation for `name`; `name_delete` destructive.
- **Acceptance:** as YAML.

---

### B6 — WP5: `shape`, `control` (single mixed tools, no `_read` split)
- **Risk:** MEDIUM (mixed annotation; action-dependent idempotency).
- **Agents:** Planner (light, annotation) + Auditor spot-check; Implementor; Validator.
- **Depends on:** B5 (sequential).

```yaml
id: B6_wp5_shape_control
title: WP5 — Collapse each 3-verb shape/control group into one mixed tool
priority: medium
depends_on:
  - B5_wp4_name
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B6 — WP5 shape/control
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP5 — shape, control (single mixed tools, no _read split)"
description: >
  Add shape(action: list|get|set-on-action) and control(action: list|get|set-on-action).
  set-on-action is a guarded mutation; list/get are reads. Neither tool is read-only nor
  destructive (mixed); idempotent_hint computed from the truth matrix (action-dependent => default false).
goals:
  - register shape, control
  - neither read-only nor destructive; idempotency from matrix
files:
  - path: src/excel_vba_mcp/api/shape_tools.py
    targets:
      - type: module
        name: shape_tools
    semantic_anchors:
      insert_after:
        type: function
        value: shape_set_on_action
  - path: src/excel_vba_mcp/api/control_tools.py
    targets:
      - type: module
        name: control_tools
    semantic_anchors:
      insert_after:
        type: function
        value: control_set_on_action
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      shape(action Literal["list","get","set-on-action"]) and
      control(action Literal["list","get","set-on-action"]); route to existing handlers.
    code_hint: |
      shape(action, ...) -> shape list|get|set_on_action
      control(action, ...) -> control list|get|set_on_action
acceptance_criteria:
  - shape, control registered
  - neither read-only nor destructive (mixed); idempotent_hint from matrix (default false)
```

- **Required tests:** per-action validation for `set-on-action`; annotation triple from matrix.
- **Acceptance:** as YAML.

---

### B7 — WP6: `macro` (reads/bookkeeping) + standalone `macro_run`; `test`
- **Risk:** HIGH (`macro_run` approval granularity; truthful phrases).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B6 (sequential).

```yaml
id: B7_wp6_macro_test
title: WP6 — Keep VBA execution standalone; consolidate macro reads + test verbs
priority: high
depends_on:
  - B6_wp5_shape_control
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B7 — WP6 macro/test
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP6 — macro (reads/bookkeeping) + standalone macro_run; test"
description: >
  Add macro(action: list|classify-background|get-last-result); keep macro_run standalone with
  unchanged signature and permission profile; add test(action: run-suite|get-run). run-suite is
  not read-only; classify-background is bookkeeping only (preserve truthful phrase).
  snapshot_* stay standalone (B10/WP8 scope for snapshot tools is unchanged).
goals:
  - register macro, macro_run, test
  - macro_run keeps its own name (distinct Kilo permission)
  - macro/test not read-only; truthful phrases preserved
files:
  - path: src/excel_vba_mcp/api/macro_tools.py
    targets:
      - type: function
        name: macro_list
      - type: function
        name: macro_run
      - type: function
        name: macro_classify_background
      - type: function
        name: macro_get_last_result
    semantic_anchors:
      insert_after:
        type: function
        value: macro_get_last_result
  - path: src/excel_vba_mcp/api/test_tools.py
    targets:
      - type: function
        name: test_suite_run
      - type: function
        name: test_get_run
    semantic_anchors:
      insert_after:
        type: function
        value: test_get_run
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_dispatch_tools
    description: >
      macro(action Literal["list","classify-background","get-last-result"]) +
      test(action Literal["run-suite","get-run"]); macro_run standalone unchanged.
    code_hint: |
      macro(action, ...) -> macro_list|macro_classify_background|macro_get_last_result
      test(action, ...) -> test_suite_run|test_get_run
      # macro_run unchanged (own permission)
acceptance_criteria:
  - macro, macro_run, test registered
  - macro_run keeps its own name + distinct permission profile
  - macro/test not read-only; truthful phrases preserved (classify-background bookkeeping)
```

- **Architectural constraints:** H1 (`macro_run`), H5; truthful description phrases preserved.
- **Required tests:** `macro_run` standalone/permission test; `macro`/`test` not read-only; per-action validation.
- **Acceptance:** as YAML.

---

### B8 — WP7: `range_read` / `range_write` / `range_structural` + standalone `range_delete`, `range_merge`
- **Risk:** HIGH (3-way split; `range_structural` not idempotent; guards).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B7 (sequential).

```yaml
id: B8_wp7_range
title: WP7 — Split 16 range verbs into read/write/structural + 2 standalone destructive
priority: high
depends_on:
  - B7_wp6_macro_test
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B8 — WP7 range
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP7 — Range split into range_read / range_write / range_structural"
description: >
  range_read(action: get|read-values|read-formulas|get-format|get-validation|
  get-conditional-formatting) read-only. range_write(action: write-values|write-formulas|
  set-format|set-validation|set-conditional-formatting) with per-action payload enforced in
  handler. range_structural(action: copy|insert|unmerge) with shift/payload guards.
  Standalone unchanged: range_delete (shift in {left,up}), range_merge.
goals:
  - register range_read, range_write, range_structural + standalone range_delete, range_merge
  - annotations from truth matrix; range_structural NOT idempotent
  - insert rejects missing shift; range_write rejects missing payload
files:
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - type: module
        name: range_tools
      - type: function
        name: _require_shift
    semantic_anchors:
      insert_after:
        type: function
        value: range_set_conditional_formatting
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_three_way_split
    description: >
      range_read / range_write / range_structural dispatch to existing read/write/structural
      handlers; per-action payload validated in handler (option (a) per B1). range_delete and
      range_merge remain standalone and destructive.
    code_hint: |
      range_read(action, ...) -> get|read_values|read_formulas|get_format|get_validation|get_conditional_formatting
      range_write(action, payload, ...) -> write_values|write_formulas|set_format|set_validation|set_conditional_formatting
      range_structural(action, ...) -> copy|insert|unmerge  # _require_shift for insert
      # range_delete, range_merge standalone unchanged
acceptance_criteria:
  - 3 range tools + 2 standalone destructive registered
  - range_structural not idempotent_hint=true (re-insert duplicates)
  - insert rejects missing shift; range_write rejects missing payload (per-action tests)
```

- **Architectural constraints:** H3/H4 (`range_structural` not idempotent), H5; `_require_shift` / `_INSERT_SHIFTS` / `_DELETE_SHIFTS` preserved; no mega range.
- **Required tests:** `insert` missing-`shift` → error; `range_write` missing-`payload` → error; `range_structural` not idempotent; `range_delete` shift ∈ {left,up}; `range_merge` dual-hint.
- **Acceptance:** as YAML.

---

### B9 — WP8: `pivot_read` / `pivot_refresh` / `pivot_field`
- **Risk:** HIGH (3-way split; `cache_index` guards; idempotency not assumed).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B8 (sequential).

```yaml
id: B9_wp8_pivot
title: WP8 — Split 14 pivot verbs into read/refresh/field
priority: high
depends_on:
  - B8_wp7_range
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B9 — WP8 pivot
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP8 — Pivot split into pivot_read / pivot_refresh / pivot_field"
description: >
  pivot_read(action: tables-list|table-get|cache-get|fields-list|field-get-filters) read-only;
  pivot_refresh(action: refresh-table|refresh-cache); pivot_field(action: field-set-layout|
  field-set-position|field-set-aggregation|field-set-properties|field-apply-filter|
  field-clear-filters). cache-get/refresh-cache require cache_index; field actions require
  pivot_table+field (field-clear-filters also accepts table-scope, preserve behavior).
goals:
  - register pivot_read, pivot_refresh, pivot_field
  - read-only only on pivot_read
  - cache_index / pivot_table / field guards enforced
  - idempotency of pivot_refresh/pivot_field from matrix (not assumed true)
files:
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - type: module
        name: pivot_tools
      - type: function
        name: pivot_cache_get
      - type: function
        name: pivot_field_get_filters
      - type: function
        name: pivot_table_refresh
      - type: function
        name: pivot_cache_refresh
    semantic_anchors:
      insert_after:
        type: function
        value: pivot_field_clear_filters
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_three_way_split
    description: >
      pivot_read / pivot_refresh / pivot_field dispatch to existing handlers; per-action
      required-param guards (cache_index, pivot_table, field) enforced in handler.
    code_hint: |
      pivot_read(action, ...) -> tables_list|table_get|cache_get|fields_list|field_get_filters
      pivot_refresh(action, ...) -> refresh_table|refresh_cache
      pivot_field(action, ...) -> field_set_layout|field_set_position|field_set_aggregation|field_set_properties|field_apply_filter|field_clear_filters
acceptance_criteria:
  - 3 pivot tools registered; read-only only on pivot_read
  - cache-get/refresh-cache reject missing cache_index; field actions reject missing pivot_table/field
  - idempotency of pivot_refresh/pivot_field set from matrix (not assumed true)
```

- **Architectural constraints:** H3 (idempotency from matrix), H5; `_require_int`/`_require_non_empty`/pivot guards preserved; no mega pivot.
- **Required tests:** `cache_index` missing → error; `pivot_table`/`field` missing → error; `pivot_refresh` not assumed idempotent.
- **Acceptance:** as YAML.

---

### B10 — WP9: `vba_read` (+ standalone `vba_replace_component_source`, `vba_insert_code`, `vba_export_project`, `vba_compile_diagnostics`)
- **Risk:** HIGH (mutations standalone; hash guards; compile not read-only).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B9 (sequential). **Completes WP1–WP9.**

```yaml
id: B10_wp9_vba_read
title: WP9 — Collapse 7 read/diagnostic VBA verbs; keep mutations/export/compile standalone
priority: high
depends_on:
  - B9_wp8_pivot
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B10 — WP9 vba_read
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP9 — vba_read (+ standalone vba_replace_component_source, vba_insert_code, vba_export_project, vba_compile_diagnostics)"
description: >
  Add vba_read(action: project-info|list-components|list-procedures|read-component|
  get-references|search-code|get-compile-diagnostics) read-only. Standalone unchanged:
  vba_replace_component_source, vba_insert_code (expected_source_hash), vba_export_project,
  vba_compile_diagnostics (not read-only). vba_get_compile_diagnostics folded into vba_read.
goals:
  - register vba_read + standalone mutations/export/compile
  - vba_read read-only; mutations keep required expected_source_hash + destructive hint
  - vba_compile_diagnostics keeps "does NOT compile" and is not read-only
files:
  - path: src/excel_vba_mcp/api/vba_tools.py
    targets:
      - type: module
        name: vba_tools
      - type: function
        name: vba_replace_component_source
      - type: function
        name: vba_insert_code
    semantic_anchors:
      insert_after:
        type: function
        value: vba_search_code
  - path: src/excel_vba_mcp/api/compile_tools.py
    targets:
      - type: function
        name: vba_compile_diagnostics
      - type: function
        name: vba_get_compile_diagnostics
    semantic_anchors:
      insert_after:
        type: function
        value: vba_get_compile_diagnostics
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
changes:
  - action: add_vba_read
    description: >
      vba_read(action Literal[...]) read-only; fold vba_get_compile_diagnostics as
      get-compile-diagnostics. Keep mutations/export/compile standalone unchanged.
    code_hint: |
      vba_read(action, ...) -> project_info|list_components|list_procedures|read_component|get_references|search_code|get_compile_diagnostics
      # vba_replace_component_source / vba_insert_code / vba_export_project / vba_compile_diagnostics standalone
acceptance_criteria:
  - vba_read read-only
  - two mutation tools keep required expected_source_hash + destructive hint
  - vba_compile_diagnostics keeps "does NOT compile" and is not read-only
  - vba_replace_component_source idempotency verified (re-call same hash fails revision guard => not idempotent)
```

- **Architectural constraints:** H1 (hash-guarded mutations standalone), H5.
- **Required tests:** `expected_source_hash` required test; `vba_compile_diagnostics` not read-only; `vba_read` read-only.
- **Acceptance:** as YAML.

---

### B11 — WP10 + WP11(lockstep part 1): Annotation rewrite from truth matrix + annotation/matrix contract tests
- **Risk:** HIGH (annotations are security/UX hints; defaults semantics).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B10 (WP10 after WP1–WP9).

```yaml
id: B11_wp10_wp11_annotations
title: WP10 (+WP11 lockstep) — Rewrite _ANN_* from truth matrix; add matrix-based contract tests
priority: high
depends_on:
  - B10_wp9_vba_read
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B11 — WP10 + WP11 annotations
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP10 — Annotation rewrite (_ANN_*) from the truth matrix"
  - "WP11 — Contract & new contract tests (annotation/matrix parts)"
  - "4. Annotations plan (MCP default semantics + per-action truth matrix)"
description: >
  Replace name-keyed _ANN_READ_ONLY/_ANN_DESTRUCTIVE/_ANN_IDEMPOTENT with the new-name sets,
  computed from the WP0 truth matrix using MCP default semantics. Resolve dual-hint
  (range_merge, vba_replace_component_source intentional). Preserve xlvba_environment inline
  annotation (not via frozenset). Decide and document None-vs-False emission. Add
  matrix-based annotation-truthfulness contract tests in lockstep.
goals:
  - every new tool has an exact, tested annotation triple derived from the matrix
  - range_structural not idempotent; dual-hint tools intentional and asserted
  - xlvba_environment inline read_only annotation preserved; None-vs-False decision documented
files:
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _annotations
      - type: data
        name: _ANN_READ_ONLY
      - type: data
        name: _ANN_DESTRUCTIVE
      - type: data
        name: _ANN_IDEMPOTENT
    semantic_anchors:
      insert_before:
        type: function
        value: _register_tools
  - path: tests/unit/test_description_contract.py
    targets:
      - type: module
        name: test_description_contract
    semantic_anchors:
      insert_after:
        type: data
        value: _TRUTH
changes:
  - action: rewrite_annotation_sets
    description: >
      Recompute the three frozensets from the truth matrix (MCP defaults). Keep
      xlvba_environment inline read_only_hint=True at its add_tool call (do NOT add to
      _ANN_READ_ONLY). Document None-vs-False choice.
    code_hint: |
      readOnly = all actions read-only
      destructive = any action destructive (false only if all additive)
      idempotent = all actions idempotent (range_structural excluded)
  - action: add_annotation_contract_tests
    description: >
      Assert exact triple per new tool from the matrix; range_structural not idempotent;
      only range_merge/vba_replace_component_source (and matrix-proven) carry both.
    code_hint: |
      for tool in final_names: assert annotations(tool) == matrix[tool]
acceptance_criteria:
  - exact tested annotation triple per new tool from matrix
  - range_structural not idempotent; dual-hint intentional + asserted
  - no tool over-flagged relative to matrix
  - xlvba_environment inline annotation intact; None-vs-False decision recorded
```

- **Architectural constraints:** H3 (truth matrix), A1 (inline env), A2 (None-vs-False decision).
- **Required tests:** matrix-based annotation truthfulness; dual-hint exact-triple; `range_structural` not idempotent.
- **Acceptance:** as YAML. **Closes gate §11.3 (annotation truth).**

---

### B12 — WP12 + WP11(lockstep part 2): Single switchover + count/name contract tests + mapping/approval tests
- **Risk:** HIGH (production switchover; approval globs; count/name lockstep).
- **Agents:** Auditor + Researcher + Planner; Implementor; Validator.
- **Depends on:** B11.

```yaml
id: B12_wp12_wp11_switchover
title: WP12 (+WP11 lockstep) — Single switchover to final names; update count/name contracts
priority: high
depends_on:
  - B11_wp10_wp11_annotations
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B12 — WP12 + WP11 switchover
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "WP12 — Client migration (single switchover, no dual-run)"
  - "WP11 — Contract & new contract tests (count/name/mapping/approval parts)"
  - "6. Migration & compatibility (single switchover)"
description: >
  Validate the new surface on an isolated branch/test server; update Kilo config + approval
  globs; then single switchover so production tools/list goes 92 -> final (~41) with no
  dual-run window. Update test_description_contract.py in lockstep: count test, _CRITICAL,
  _EXAMPLE_TOOLS, _TRUTH, guard list; add mapping-totalness and security/approval-matrix tests.
goals:
  - no dual-run; production tools/list never ~121-130
  - server.mcp.list_tools() returns exactly the final names
  - contract file updated in lockstep; mapping + approval tests green
  - migration note documents new approval globs + single-switchover procedure
files:
  - path: src/excel_vba_mcp/server.py
    targets:
      - type: function
        name: _register_tools
    semantic_anchors:
      insert_before:
        type: function
        value: server_health
  - path: tests/unit/test_description_contract.py
    targets:
      - type: function
        name: test_tool_count_is_92_exposed
      - type: data
        name: _CRITICAL
      - type: data
        name: _EXAMPLE_TOOLS
      - type: data
        name: _TRUTH
    semantic_anchors:
      insert_after:
        type: data
        value: _TRUTH
changes:
  - action: single_switchover
    description: >
      Register only final names; remove old names (no aliases). Verify list_tools returns final set.
    code_hint: |
      # _register_tools registers final names only
  - action: update_contract_lockstep
    description: >
      test_tool_count_is_final_exposed; replace _CRITICAL old names with new consolidated/standalone
      names; keep _EXAMPLE_TOOLS; keep _TRUTH phrases; update guard list; add mapping-totalness +
      security/approval-matrix completeness tests.
    code_hint: |
      assert len(list_tools()) == final_count
      assert registered_names == ledger.target_set
acceptance_criteria:
  - no dual-run; list_tools returns exactly final names
  - count/_CRITICAL/_EXAMPLE_TOOLS/_TRUTH/guard-list updated + green
  - mapping-totalness + approval-matrix tests present + green
  - migration note documents new globs + single switchover
```

- **Architectural constraints:** H1 (approval globs), H2 (no dual-run), A3 (count reconcile 91→92→final).
- **Required tests:** `test_tool_count_is_final_exposed`; mapping totalness (92→targets); approval-matrix completeness (four high-risk standalone/protected; no merged profiles).
- **Acceptance:** as YAML. **Closes gates §11.5 (migration) and §11.2 (final map).**

---

### B13 — WP13: Toolset layer — **DEFERRED (documented optional; no implementation)**
- **Risk:** LOW (deferred).
- **Agents:** none now (Planner note only).
- **Depends on:** B12 (would depend on final names).
- **Decision:** **DEFER.** Per plan §10, implement only if a second client profile with a disjoint
  domain set appears; the single default surface (~41) is the end state (rule 5: avoid
  overengineering). Recorded here so the option is not lost.
- **If ever approved (future, not now):** new `src/excel_vba_mcp/toolsets.py` (`ToolsetSpec`:
  name → frozenset of final names), `--toolsets` CLI / `XLVBA_TOOLSETS` env parsing, registration
  gating in `_register_tools`, `core` always included, unknown name fails fast, membership over
  **final** names; unit tests: every final name in ≥1 toolset, `⋃ toolsets == final set`,
  `core ⊆ every resolved selection`. Six toolsets: `core`, `vba`, `macro-test`, `data`,
  `analysis`, `objects` (see plan §10.2). **No stubs now.**

---

### B14 — Documentation update (per `docs/00-overview/doc-maintenance-rules.md`)
- **Risk:** MEDIUM (doc contract; stale count; ru mirror).
- **Agents:** Planner (light) + Auditor spot-check; Implementor; Validator.
- **Depends on:** B12 (docs reflect final surface).

```yaml
id: B14_documentation
title: Documentation — reconcile tool catalogue to final surface (+ ru mirror)
priority: medium
depends_on:
  - B12_wp12_wp11_switchover
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B14 — Documentation
source_file: docs/00-overview/doc-maintenance-rules.md
description: >
  Update the tool catalogue and related docs to the final surface, following
  doc-maintenance-rules (frontmatter contract, English-only canonical, relative cross-links,
  single source of truth). Reconcile stale count 91 -> 92 -> final. Keep the Russian mirror
  docs/ru/ in sync. Record the doc path-mismatch defect (rules reference non-existent
  docs/02-build-and-deployment/vba-build.md; actual docs are flat 01..11 + docs/ru/).
goals:
  - docs/05-mcp-tools-reference.md reflects final ~41 tools (count + names + annotations)
  - docs/ru/05-mcp-tools-reference.md mirror updated in lockstep
  - stale 91 reconciled to 92 then final count
  - path-mismatch defect recorded/corrected within doc scope
files:
  - path: docs/05-mcp-tools-reference.md
    targets:
      - type: document
        name: mcp-tools-reference
  - path: docs/ru/05-mcp-tools-reference.md
    targets:
      - type: document
        name: mcp-tools-reference-ru
  - path: docs/00-overview/doc-maintenance-rules.md
    targets:
      - type: document
        name: doc-maintenance-rules
    semantic_anchors:
      insert_after:
        type: section
        value: File Placement Rules
changes:
  - action: update_tool_catalogue
    description: >
      Rewrite the tool table to final names/annotations; update count; preserve frontmatter;
      use relative cross-links; keep English canonical.
    code_hint: |
      | tool | actions | readOnly | destructive | idempotent |
  - action: sync_ru_mirror
    description: >
      Mirror the catalogue change into docs/ru/05-mcp-tools-reference.md (keep in sync).
  - action: record_path_mismatch
    description: >
      Note the non-existent docs/02-build-and-deployment/vba-build.md reference and the actual
      flat docs layout; correct the reference if within doc scope.
acceptance_criteria:
  - catalogue reflects final surface; count reconciled 91->92->final
  - ru mirror in sync; frontmatter valid; English canonical; relative links valid
  - path-mismatch recorded/corrected; no audit-file edits
```

- **Architectural constraints:** doc-maintenance-rules (frontmatter, SSOT, relative links, English-only canonical, split thresholds); A3/A4.
- **Required tests:** doc lint (frontmatter present, links valid, line count < 800 soft / 1000 hard).
- **Acceptance:** as YAML.

---

### B15 — Final validation (Validator last)
- **Risk:** HIGH (release gate).
- **Agents:** Validator (lead) + Auditor (sign-off); no Implementor unless defects found.
- **Depends on:** B11, B12, B14.

```yaml
id: B15_final_validation
title: Final validation — full suite, measurement, definition-of-done
priority: high
depends_on:
  - B11_wp10_wp11_annotations
  - B12_wp12_wp11_switchover
  - B14_documentation
source_reference: .ai/plans/01-tool-surface-consolidation.md
source_section: B15 — Final validation
source_file: .ai/audit/reports/final-remediation-report-tool-consolidation.md
source_blocks:
  - "9. Definition of done (checklist)"
  - "5. Token & selection measurement plan"
description: >
  Run the full unit suite + integration_excel ranges/pivots/tables/vba/test-runner regression
  (handlers reused). Measure before/after tools/list bytes (target ~35-55%, measured not asserted)
  and report selection-quality benchmark. Walk the definition-of-done checklist. No production
  data macros.
goals:
  - full tests/unit green; integration_excel regression green
  - token reduction measured; selection-quality reported
  - definition-of-done fully satisfied
files:
  - path: tests/unit/
  - path: tests/integration_excel/
  - path: .ai/tmp/  (after-measurement output)
changes:
  - action: run_full_verification
    description: >
      Run unit + integration regression; measure after bytes; record numbers for PR.
    code_hint: |
      # same list_tools() serialization command as B0, after switchover
acceptance_criteria:
  - full tests/unit green; integration_excel ranges/pivots/tables/vba/test-runner green
  - token reduction measured within ~35-55% (or restated); selection benchmark reported
  - definition-of-done checklist all checked
```

- **Acceptance:** as YAML. **Closes gate §11.6 (benchmark).**

---

## 4. WP14 decision gate (explicit framing — not a foregone conclusion)

- **Question:** per-action required-param mechanism — **(a)** flat args + handler validation, or
  **(b)** action-specific `inputSchema` via JSON Schema 2020-12 `oneOf` / SDK low-level custom-schema hook.
- **Evidence input (Auditor, verified):** public `MCPServer.add_tool` has **no** `inputSchema`/
  custom-schema parameter; the SDK derives a **flat** arg model from the function signature. So (b)
  likely requires unsupported internals today. This makes (a) the only *fully-supported* path — but
  it is **not** the Planner's decision to make.
- **Gate procedure (B1):** build both throwaway variants in `.ai/tmp`; run SDK → MCP → Kilo → real
  LLM; compare argument-error rate + first-call success + client support + whether (b) needs
  unsupported internals.
- **Decision rule:** if (b) shows no measurable gain or poor client support (or needs unsupported
  internals) → **keep (a)** and rely on handler guards + per-action validation tests + `action`
  enum/description docs. Otherwise adopt (b) and update B2–B10 accordingly.
- **Residual risk under (a):** consolidated schemas show per-action fields as optional; mitigated by
  handler guards (unchanged code), per-action validation contract tests, and documented required
  params. Tracked as R1/R14.
- **The final choice is recorded in B1 and referenced by B2–B10; it is not assumed anywhere in this plan.**

---

## 5. WP13 deferral note

WP13 (toolset layer) is **DEFERRED** (plan §10). Rationale: MCP `tools/list` must not vary
per-connection; most clients cannot select arbitrary toolsets for a third-party server; the
`--toolsets` precedent is a launch-process flag (deployment convenience, not a protocol feature).
A single client profile is the current reality, so the single default surface (~41) is the end state.
Revisit only when a second client profile with a disjoint domain set appears. If approved later, the
design constraints in B13 apply (default = all; `core` always included; membership over final names;
unknown name fails fast; not per-session/authorization).

---

## 6. Dependency graph (DAG)

```
B0 ──► B1 ──► B2 ──► B3 ──► B4 ──► B5 ──► B6 ──► B7 ──► B8 ──► B9 ──► B10 ──► B11 ──► B12 ──► B14 ──► B15
        (gate)                                                                              │
                                                                                              ▼
                                                                                            B13 (DEFERRED)
```

- B2–B10 touch disjoint `api/*_tools.py` modules but run **sequentially** (one Implementor).
- B11 (WP10) after B10 (WP1–WP9). B12 (WP12) after B11; WP11 contract updates are embedded in B11
  and B12 (lockstep). B14 after B12. B15 last. B13 deferred (would depend on B12).

---

## 7. Traceability (plan WP → block)

| Plan WP | Block | Risk | Agents (analysis) |
|---|---|---|---|
| WP0 | B0 | HIGH | Auditor + Researcher + Planner |
| WP14 | B1 | GATE | Researcher + Auditor + Planner |
| WP1 | B2 | HIGH | Auditor + Researcher + Planner |
| WP2 | B3 | MEDIUM | Planner + Auditor spot-check |
| WP3 | B4 | HIGH | Auditor + Researcher + Planner |
| WP4 | B5 | MEDIUM | Planner |
| WP5 | B6 | MEDIUM | Planner + Auditor spot-check |
| WP6 | B7 | HIGH | Auditor + Researcher + Planner |
| WP7 | B8 | HIGH | Auditor + Researcher + Planner |
| WP8 | B9 | HIGH | Auditor + Researcher + Planner |
| WP9 | B10 | HIGH | Auditor + Researcher + Planner |
| WP10 + WP11(part) | B11 | HIGH | Auditor + Researcher + Planner |
| WP12 + WP11(part) | B12 | HIGH | Auditor + Researcher + Planner |
| WP13 | B13 | DEFERRED | none |
| Doc step | B14 | MEDIUM | Planner + Auditor spot-check |
| Final | B15 | HIGH | Validator + Auditor |

---

## 8. Final-validation checklist (definition of done)

- [ ] B0 baseline in `.ai/tmp`: count 92, serialized bytes, mapping ledger (92 rows), security/approval matrix (92 tools), annotation truth matrix (every action).
- [ ] §11 gates resolved: approval matrix (B0), final map (B0/B12), truth matrix (B0/B11), schema PoC (B1), migration (B12), benchmark (B15).
- [ ] Consolidated + standalone high-risk/destructive + lifecycle tools registered (B2–B10); no mega range/pivot; working target ~41 (range ~38–42).
- [ ] `macro_run`, `excel_save_as`, `table_delete_column` standalone; `workbook` `close` disposition implemented/tested; no merged permission profiles.
- [ ] Final `server.mcp.list_tools()` returns exactly the final names (single switchover; no `meta.deprecated`, no dual-run).
- [ ] Per-action required params enforced (handlers) with passing per-action validation tests; WP14 decision recorded.
- [ ] `_ANN_*` computed from truth matrix with MCP default semantics; exact triple asserted; `range_structural` not idempotent; only matrix-proven dual-hint tools carry both; `xlvba_environment` inline annotation intact; None-vs-False decision documented.
- [ ] `test_description_contract.py` updated (count, `_CRITICAL`, `_EXAMPLE_TOOLS`, `_TRUTH`, guard list) and green; matrix-based + mapping-totalness + per-action tests added and green.
- [ ] Full `tests/unit` green; `tests/integration_excel` ranges/pivots/tables/vba/test-runner regression green (handlers reused).
- [ ] Token reduction **measured** before/after within ~35–55% (or restated); selection-quality benchmark reported; numbers in PR.
- [ ] Migration note documents single switchover + new client approval globs (no dual-run).
- [ ] No general `xlvba_batch`; `range_delete`/`range_merge`/`sheet_delete`/`table_delete_rows`/`table_delete_column`/`name_delete`/`macro_run`/`excel_save_as`/`vba_replace_component_source`/`vba_insert_code` remain standalone and destructive.
- [ ] Docs updated: `docs/05-mcp-tools-reference.md` + `docs/ru/05-mcp-tools-reference.md` reflect final surface; count reconciled 91→92→final; doc path-mismatch recorded (B14).
- [ ] (Deferred) WP13 `--toolsets` only if approved; default registers all; union of toolsets == final set; unknown name fails fast.