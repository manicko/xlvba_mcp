# Final Remediation Report — `xlvba_*` Tool-Surface Consolidation

Baseline: AUDITOR verdict **ACCEPT-WITH-CHANGES**; independent review verdict
**ACCEPT WITH CHANGES** (9 review points). This report converts both into an ordered,
implementation-ready remediation plan. **No code is modified by this task.** All paths are
relative to `C:\py_exp\xlvba_mcp`. Line references were verified against the working tree on
2026-11-26; items that could not be fully pinned are marked **verify**.

Review-driven corrections applied (see §11 for the binding pre-implementation gates):
- Target count softened to **~38–42** (a target, not a hard number); +3 mandatory standalone
  tools (`macro_run`, `excel_save_as`, `table_delete_column`) ⇒ working target **~41**.
- **Approval granularity is a first-class constraint**: Kilo permissions bind to
  `{server}_{tool}` names, not to an `action` argument value; annotations are hints, not
  enforcement. Consolidation must not merge operations that require distinct permissions.
- MCP annotation **default semantics** corrected (`readOnly=false`, `destructive=true`,
  `idempotent=false`); `range_structural` is **not** idempotent.
- Discriminated-union rejection softened to a **PoC gate** (flat args vs action-specific
  JSON Schema via `oneOf`).
- **Dual-run / temporary aliases removed** (prior no-backward-compatibility decision).
- **Progressive Tool Discovery** reclassified from "rejected" to **roadmap** (distinct from a
  mega `call_tool`); not implemented now only because Kilo does not yet confirm per-tool
  progressive loading.

---

## 1. Executive summary

**Verdict: PROCEED WITH CHANGES.** The naive "collapse to ~29 `action`-dispatch tools"
proposal is rejected; the revised **~41-tool** surface (consolidate non-destructive verbs per
domain + `_read` companions, keep destructive/high-risk/lifecycle operations standalone) is
approved subject to the work packages and the §11 gates.

End-state (3 bullets):

- **~41 tools** (from 92; target range ~38–42, not a hard number): consolidated per-domain
  `action`-dispatch tools (each with a `_read` companion where reads exist), **standalone
  high-risk/destructive tools** (`sheet_delete`, `table_delete_rows`, `table_delete_column`,
  `name_delete`, `range_delete`, `range_merge`, `macro_run`, `excel_save_as`,
  `vba_replace_component_source`, `vba_insert_code`, `excel_recover`, `snapshot_restore`), and
  **standalone lifecycle/health/recovery/snapshot/job tools**.
- **Approval granularity preserved by design.** Operations that require distinct Kilo
  permissions (notably `macro_run` vs macro reads, `excel_save_as` vs `excel_save`,
  `table_delete_column` vs `table`) stay on **standalone tools with unchanged names/signatures**.
  Where a consolidated tool would merge distinct permission profiles, the operation is
  extracted instead. `_require_*` guards remain **validation**, not authorization; any
  server-side action-level policy is out of scope unless separately specified and tested (§11.1).
- **Truthful annotations** computed from a **per-action truth matrix** (§4) using MCP default
  semantics; `range_structural` is **not** idempotent; dual-hint (`range_merge`,
  `vba_replace_component_source`) is intentional and test-locked. Old names are removed by a
  **single switchover** (no dual-run, no `meta.deprecated`).

Expected outcome: materially smaller `tools/list` payload (target **~35–55%**, to be
**measured**, not asserted), unchanged runtime behavior (handlers reused verbatim), an explicit
test-backed handler contract, a documented client-migration (single switchover), and a
**measured** tool-selection-quality check. The 35–55% figure is a target, not a proven result.

---

## 2. Work breakdown (ordered work packages)

> Ordering rule: WP0 establishes the measurement, mapping, **security/approval matrix**, and
> **annotation truth matrix** baseline; WP1–WP9 build the new surface behind existing handlers;
> WP10 flips annotations from the truth matrix; WP11 locks contracts; WP12 performs the single
> switchover; WP14 is the schema PoC gate. Do not start WP10 before WP1–WP9; do not start WP12
> before WP9 sign-off; resolve §11 gates before WP1.

### WP0 — Baseline: measurement, mapping ledger, security/approval matrix, annotation truth matrix (no behavior change)
- **Goal:** freeze the "before" numbers and produce the four baseline artifacts that later
  steps are verified against: (1) token measurement, (2) total old→new mapping, (3) per-tool
  security/approval matrix, (4) per-action annotation truth matrix.
- **Files/areas:** `src/excel_vba_mcp/server.py` (`_register_tools`, lines 1784–2328;
  `server_health` decorator at 2498–2499); `tests/unit/test_description_contract.py`.
- **Change:** add a throwaway measurement script under `.ai/tmp/` (per project rule 17, not
  committed) and committed artifacts (data files or test fixtures):
  - **Mapping ledger:** every one of the 92 current tool names → exactly one target
    (consolidated tool + action, or standalone name). Reconcile: **92 = 91 `add_tool` + 1
    `@mcp.tool` (`server_health`)**.
  - **Security/approval matrix (§11.1):** for each of the 92 tools, record risk level and the
    current Kilo approval policy, and the intended new name + its permission profile. Prove
    that no consolidation merges two distinct permission profiles.
  - **Annotation truth matrix (§11.3 / §4):** for each action of each new tool, record actual
    `readOnly` / `destructive` / `idempotent` behavior; tool-level hints are computed from it.
- **Acceptance criteria:**
  - Baseline serialized `tools/list` byte size and tool count (92) recorded in `.ai/tmp`.
  - Mapping ledger has 92 rows; every row resolves to a defined target; no duplicates/orphans.
  - Security/approval matrix covers all 92 tools; the four high-risk tools
    (`macro_run`, `excel_save_as`, `table_delete_column`, `workbook` `close`) have an explicit
    disposition (standalone or protected action).
  - Annotation truth matrix covers every action of every new tool.
  - Tests assert ledger totalness (`len == 92`) and matrix completeness (see WP11).

### WP1 — `workbook` + `workbook_read`; standalone `excel_save_as`
- **Goal:** collapse the non-destructive workbook verbs; keep `save-as` standalone (distinct
  overwrite/file policy); verify `close` unsaved-change behavior.
- **Files/areas:** `src/excel_vba_mcp/api/workbook_tools.py`
  (`excel_open` 93, `excel_close` 135, `excel_save` 152, `excel_save_as` 167,
  `excel_get_status` 205, `excel_list_workbooks` 222, `workbook_get_model` 229);
  registration in `server.py` 1786–1864, 2317–2322.
- **Change:** new `workbook(action: Literal["open","close","save"], ...)` dispatching to the
  existing functions; new `workbook_read(action: Literal["status","list","model"], ...)`.
  `excel_save_as` **stays standalone** (own permission/overwrite policy). Handlers reused
  unchanged; dispatch only routes and forwards validated args.
- **`close` protection (verify):** confirm whether `excel_close` can discard unsaved changes.
  If yes, add an explicit `discard_changes` flag enforced in the handler (default = refuse) with
  a per-action test; if a distinct permission is required, extract `workbook_close` standalone.
- **Acceptance criteria:** `workbook`, `workbook_read`, `excel_save_as` registered; `workbook`
  is **not** destructive and does **not** carry `save-as`; `excel_save_as` keeps
  `destructive_hint=True` and its overwrite policy; `close` disposition documented and tested.

### WP2 — `sheet` + `sheet_read`; standalone `sheet_delete`
- **Goal:** collapse 6 non-destructive sheet verbs; extract the destructive one.
- **Files/areas:** `api/sheet_tools.py` (`sheets_list` 62, `sheet_get` 79, `sheet_add` 93,
  `sheet_rename` 111, `sheet_copy` 130, `sheet_delete` 147, `sheet_set_visibility` 165).
- **Change:** `sheet(action: Literal["add","rename","copy","set-visibility"], ...)` +
  `sheet_read(action: Literal["list","get"], ...)`. `sheet_delete` **standalone** (B3), current
  signature.
- **Acceptance criteria:** `sheet`, `sheet_read`, `sheet_delete` registered; `sheet` not
  destructive; `sheet_delete` keeps `destructive_hint=True`.

### WP3 — `table` + `table_read`; standalone `table_delete_rows` **and** `table_delete_column`
- **Goal:** collapse non-destructive table verbs; keep **both** destructive deletions standalone.
- **Files/areas:** `api/table_tools.py` (`tables_list` 105, `table_get` 121,
  `table_read_rows` 135, `table_append_rows` 166, `table_insert_row` 187,
  `table_update_rows` 220, `table_delete_rows` 245, `table_add_column` 293,
  `table_delete_column` 320, `table_create` 339, `table_resize` 364).
- **Change:** `table(action: Literal["create","resize","append-rows","insert-row",
  "update-rows","add-column"], ...)` + `table_read(action: Literal["list","get","read-rows"],
  ...)`. `table_delete_rows` **standalone** (XOR guard at `table_tools.py:268` unchanged);
  `table_delete_column` **standalone** (own permission policy). `delete-column` is **removed**
  from the `table` action set.
- **Acceptance criteria:** `table`/`table_read`/`table_delete_rows`/`table_delete_column`
  registered; `table` not destructive and has no `delete-column`; both deletions destructive
  with signatures unchanged; per-action tests for the XOR guard and column delete.

### WP4 — `name` + `name_read`; standalone `name_delete`
- **Goal:** collapse name create/update; extract delete.
- **Files/areas:** `api/name_tools.py` (`names_list` 102, `name_get` 120, `name_create` 146,
  `name_update` 178, `name_delete` 209).
- **Change:** `name(action: Literal["create","update"], ...)` +
  `name_read(action: Literal["list","get"], ...)`. `name_delete` standalone.
- **Acceptance criteria:** three tools registered; `name` not destructive; `name_delete`
  destructive.

### WP5 — `shape`, `control` (single mixed tools, no `_read` split)
- **Goal:** collapse each 3-verb group into one tool (list/get/set-on-action).
- **Files/areas:** `api/shape_tools.py` (63/81/98); `api/control_tools.py` (64/83/100).
- **Change:** `shape(action: Literal["list","get","set-on-action"], ...)` and
  `control(action: Literal["list","get","set-on-action"], ...)`. `set-on-action` is a guarded
  mutation; `list`/`get` are reads.
- **Acceptance criteria:** both registered; **neither** read-only nor destructive (mixed);
  `idempotent_hint` computed from the truth matrix (action-dependent ⇒ default `false`), §4.

### WP6 — `macro` (reads/bookkeeping) + standalone `macro_run`; `test`
- **Goal:** keep VBA execution on its own tool; consolidate the rest.
- **Files/areas:** `api/macro_tools.py` (`macro_list` 70, `macro_run` 87,
  `macro_classify_background` 132, `macro_get_last_result` 171); `api/test_tools.py`
  (`test_suite_run` 214, `test_get_run` 276; snapshot_* 97–168 stay standalone, WP8).
- **Change:** `macro(action: Literal["list","classify-background","get-last-result"], ...)`;
  `macro_run` **standalone, unchanged signature and permission profile**; `test(action:
  Literal["run-suite","get-run"], ...)`. `test run-suite` is **not** read-only;
  `classify-background` is bookkeeping only (preserve truthful phrase, §7).
- **Acceptance criteria:** `macro`, `macro_run`, `test` registered; `macro_run` keeps its own
  name (distinct Kilo permission); `macro`/`test` not read-only; truthful phrases preserved.

### WP7 — Range split into `range_read` / `range_write` / `range_structural`
- **Goal:** collapse 16 range verbs into 3 cohesive tools (B2: **no single mega range**).
- **Files/areas:** `api/range_tools.py` (read 154–303; write 304–382, 529–580;
  structural 383–528). Guards: `_require_shift` 130–146, `_INSERT_SHIFTS`/`_DELETE_SHIFTS`
  123–127.
- **Change:**
  - `range_read(action: Literal["get","read-values","read-formulas","get-format",
    "get-validation","get-conditional-formatting"], ...)` — read-only.
  - `range_write(action: Literal["write-values","write-formulas","set-format",
    "set-validation","set-conditional-formatting"], ...)` — per-action `payload` enforced in
    handler (§3).
  - `range_structural(action: Literal["copy","insert","unmerge"], ...)` — `insert` requires
    `shift ∈ {right,down}` via `_require_shift`; `copy` requires source/destination; `unmerge`
    requires `address`.
  - **Standalone (unchanged):** `range_delete` (shift ∈ {left,up}), `range_merge`.
- **Acceptance criteria:** 3 range tools + 2 standalone destructive registered; annotations
  computed from the truth matrix (§4) — **`range_structural` is not `idempotent_hint=true`**
  (re-`insert` duplicates); `insert` rejects missing `shift`; `range_write` rejects missing
  `payload` (per-action tests).

### WP8 — Pivot split into `pivot_read` / `pivot_refresh` / `pivot_field`
- **Goal:** collapse 14 pivot verbs into 3 cohesive tools (B2: **no single mega pivot**).
- **Files/areas:** `api/pivot_tools.py` (read 211–303; refresh 309–345; field 347–571).
  Anchors: `pivot_cache_get` `cache_index` 256; `pivot_field_get_filters`
  `pivot_table`+`field` 284; `pivot_table_refresh` 309; `pivot_cache_refresh` 328.
- **Change:**
  - `pivot_read(action: Literal["tables-list","table-get","cache-get","fields-list",
    "field-get-filters"], ...)` — read-only; `cache-get` requires `cache_index`, others
    require `pivot_table` (+`field` for `field-get-filters`) (§3).
  - `pivot_refresh(action: Literal["refresh-table","refresh-cache"], ...)` —
    `refresh-table` requires `pivot_table`, `refresh-cache` requires `cache_index`.
  - `pivot_field(action: Literal["field-set-layout","field-set-position",
    "field-set-aggregation","field-set-properties","field-apply-filter",
    "field-clear-filters"], ...)` — all require `pivot_table`+`field` (except
    `field-clear-filters` which also accepts table-scope; preserve existing behavior).
- **Acceptance criteria:** 3 pivot tools registered; read-only only on `pivot_read`;
  `cache-get`/`refresh-cache` reject missing `cache_index`; field actions reject missing
  `pivot_table`/`field`; idempotency of `pivot_refresh`/`pivot_field` set from the truth matrix
  (**not** assumed true — re-refresh may yield different data).

### WP9 — `vba_read` (+ standalone `vba_replace_component_source`, `vba_insert_code`, `vba_export_project`, `vba_compile_diagnostics`)
- **Goal:** collapse 7 read/diagnostic VBA verbs; keep mutations + export + compile standalone.
- **Files/areas:** `api/vba_tools.py` (reads 68–188; mutations 191–264);
  `api/compile_tools.py` (`vba_compile_diagnostics` 63, `vba_get_compile_diagnostics` 84).
- **Change:** `vba_read(action: Literal["project-info","list-components","list-procedures",
  "read-component","get-references","search-code","get-compile-diagnostics"], ...)` —
  read-only. **Standalone unchanged:** `vba_replace_component_source`
  (expected_source_hash 195/217), `vba_insert_code` (expected_source_hash 232/256),
  `vba_export_project`, `vba_compile_diagnostics` (not read-only); `vba_get_compile_diagnostics`
  folded into `vba_read` as `get-compile-diagnostics`.
- **Acceptance criteria:** `vba_read` read-only; the two mutation tools keep required
  `expected_source_hash` + destructive hint; `vba_compile_diagnostics` keeps "does NOT compile"
  and is not read-only; `vba_replace_component_source` idempotency verified (re-call with same
  hash likely fails the revision guard ⇒ not idempotent).

### WP10 — Annotation rewrite (`_ANN_*`) from the truth matrix — see §4
- **Goal:** replace name-keyed sets with the new-name sets; compute hints from the WP0 truth
  matrix using MCP default semantics; resolve dual-hint.
- **Files/areas:** `server.py:1686–1781` (`_ANN_READ_ONLY`, `_ANN_DESTRUCTIVE`,
  `_ANN_IDEMPOTENT`, `_annotations`).
- **Acceptance criteria:** every new tool has an **exact, tested** annotation triple derived
  from the truth matrix; `range_structural` is not idempotent; dual-hint tools are intentional
  and asserted; no tool over-flagged relative to the matrix.

### WP11 — Contract & new contract tests — see §7
- **Goal:** update existing name-keyed contracts in lockstep; add totalness, matrix-based
  annotation-truthfulness, and per-action validation tests.
- **Acceptance criteria:** full unit suite green; new tests present and passing.

### WP12 — Client migration (single switchover, no dual-run)
- **Goal:** replace old names with new names in one switchover; no temporary aliases.
- **Files/areas:** `server.py` `_register_tools` (final names only); Kilo config + approval
  globs (external, documented); test server/branch for pre-switch validation.
- **Change:** validate the new surface on an isolated branch/test server; update Kilo config and
  approval globs; verify; then switch to the new names. Old names are **not** registered during
  the window (prior no-backward-compatibility decision). `list_tools` returns exactly the final
  ~41 names.
- **Acceptance criteria:** no dual-run; `tools/list` never grows to ~121–130 in production;
  `server.mcp.list_tools()` returns exactly the final names; migration note documents the new
  approval globs and the single-switchover procedure.

### WP13 — Toolset layer (optional, deployment-time filter) — see §10
- **Goal:** add an opt-in `--toolsets` filter restricting the *registered* set to named domain
  groups, with **no change to the default** (no flag ⇒ all ~41 tools).
- **Files/areas:** `server.py` `_register_tools` (registration gating); new
  `src/excel_vba_mcp/toolsets.py` (`ToolsetSpec`: name → frozenset of final tool names);
  entrypoint CLI/env parsing; `docs/` migration note.
- **Change:** parse `--toolsets` (CLI) / `XLVBA_TOOLSETS` (env); resolve to a tool-name
  allow-set; skip `add_tool` for names not in the set. `core` is **always** included. Membership
  defined over **final post-WP12 names**.
- **Design constraints:** default registers all → zero behavior change; process-level filtering
  (GitHub MCP model), **not** per-connection/authorization; unknown name ⇒ fail fast.
- **Acceptance criteria:** `--toolsets data,analysis` registers exactly that union **plus
  `core`**; default registers all; unit test asserts every final name is in ≥1 toolset and the
  union of all toolsets == final set; unknown name exits non-zero.

### WP14 — Schema PoC gate (flat args vs action-specific JSON Schema) — see §3
- **Goal:** empirically decide the per-action required-param mechanism before WP1.
- **Change:** build two throwaway variants (in `.ai/tmp`, not committed): (a) flat args +
  handler validation; (b) action-specific `inputSchema` via JSON Schema 2020-12 `oneOf` (or the
  Python SDK low-level custom-schema hook). Run both through SDK → MCP → Kilo → a real LLM.
- **Acceptance criteria:** a short comparison of argument-error rate and first-call success for
  both variants; a documented decision. If (b) gives no measurable gain or poor client support,
  keep (a). Do not generalize one SDK's limits to MCP.

> WP ordering note: WP1–WP9 may proceed in parallel after WP0 (disjoint modules). WP10 depends
> on WP1–WP9. WP12 depends on WP9 sign-off. WP14 gates WP1 (resolve before building). WP13 is
> **optional** and depends on WP12; defer if a single client profile holds (rule 5).

---

## 3. Per-action required-param strategy (the crux)

**Decision: option (a) — documented handler-enforcement (excel-mcp precedent) — is the working
default, but the rejection of option (b) is no longer categorical.** A **PoC (WP14)** must
compare (a) against (b) before WP1; the final choice is evidence-based.

### Justification (working default)
- The MCP Python SDK builds a **flat** Pydantic arg model from the function signature
  (`.venv/.../mcp/server/mcpserver/tools/base.py:107`:
  `parameters = func_arg_metadata.arg_model.model_json_schema(by_alias=True)`). A single `action`
  enum cannot mark per-action fields "required", so a flat signature moves required-ness into
  the handler.
- The codebase **already** enforces these invariants in handlers/services (`_require_shift`
  130–146; `if (indices is None) == (where is None)` 268;
  `_require_non_empty(expected_source_hash)` 217/256; `_require_int(cache_index)` 265).
  Handler-enforcement reuses this single source of truth verbatim.
- Project rules: *Avoid Overengineering* (5), *Follow Existing Patterns* (7). A nested
  `request` union would force `{action, request:{...}}`, break flat ergonomics, and churn DTOs.

### Why (b) is no longer rejected outright
- MCP permits **JSON Schema 2020-12** for `inputSchema`, so `oneOf` can express per-action
  required fields at the protocol level.
- The Python SDK exposes a **low-level custom-schema** path (not only signature-derived).
- Therefore the flat-signature limitation is an **SDK default**, not an MCP limit. WP14 must
  measure whether action-specific schemas reduce argument errors and improve first-call success
  through a real LLM in Kilo. If not, keep (a).

### Residual risk & mitigation (if (a) is chosen)
Consolidated tools' schemas show per-action fields as **optional**. Mitigation:
1. Handler guard is the contract (unchanged code).
2. **Per-action validation contract test** (§7) omits each required param and asserts a
   structured `ExcelVbaMcpError`.
3. `action` enum + the description's `Parameters` section document required params per action.

### Affected tools — concrete handling
| Tool | Action | Required param(s) | Enforcement (existing anchor) |
|---|---|---|---|
| `range_structural` | `insert` | `shift ∈ {right,down}` | `_require_shift(... _INSERT_SHIFTS)` `range_tools.py:130–146` |
| `range_structural` | `copy` | `source_range`,`destination`(+sheet) | existing copy preflight |
| `range_structural` | `unmerge` | `address` | existing unmerge guard |
| `range_write` | `write-values` | `payload` (`RangeWriteRequest`) | model + handler 304 |
| `range_write` | `write-formulas` | `payload` (`FormulaWriteRequest`) | model + handler 331 |
| `range_write` | `set-format`/`set-validation`/`set-conditional-formatting` | `address`+spec | existing handlers 357/529/557 |
| `table_delete_rows` *(standalone)* | — | `indices` XOR `where` | **unchanged** `table_tools.py:268` |
| `table_delete_column` *(standalone)* | — | `header` | **unchanged** `table_tools.py:320+` |
| `vba_replace_component_source` *(standalone)* | — | `expected_source_hash` | **unchanged** 195/217 |
| `vba_insert_code` *(standalone)* | — | `expected_source_hash` | **unchanged** 232/256 |
| `excel_save_as` *(standalone)* | — | `path` (+ overwrite flags) | **unchanged** `workbook_tools.py:167+` |
| `pivot_read` | `cache-get` | `cache_index` | `_require_int(cache_index)` 265 |
| `pivot_read` | `table-get`/`fields-list` | `pivot_table` | `_require_non_empty` 251/279 |
| `pivot_read` | `field-get-filters` | `pivot_table`+`field` | 299–301 |
| `pivot_refresh` | `refresh-table` | `pivot_table` | 322–323 |
| `pivot_refresh` | `refresh-cache` | `cache_index` | 328+ |
| `pivot_field` | all | `pivot_table`+`field` | per-handler 347–571 |

**Key architectural win from B3:** the hardest invariants (`table_delete_rows` XOR,
`table_delete_column`, `vba_*` hash, `excel_save_as` overwrite, `range_delete` shift) live on
**standalone tools whose signatures are unchanged** — consolidation never degrades them.

---

## 4. Annotations plan (MCP default semantics + per-action truth matrix)

`_annotations(name)` (`server.py:1774–1781`) reads three frozensets. Rewrite them to the new
names. **Do not hand-author hints:** compute them from the WP0 **per-action truth matrix**.

### 4.1 MCP default semantics (binding)
The spec defines defaults; **absence is not neutral**:
- `readOnlyHint` default **false**.
- `destructiveHint` default **true**.
- `idempotentHint` default **false**.

Correct computation rule per tool:
- `readOnlyHint=true` **iff all** actions are read-only.
- `destructiveHint=true` **if at least one** action can perform destructive changes;
  `destructiveHint=false` **only if all** actions are exclusively additive.
- `idempotentHint=true` **only if all** actions are truly idempotent (re-call with same args
  causes no additional state change).

### 4.2 Per-action truth matrix (WP0 artifact, then compute)
Before rewriting `_ANN_*`, record for each action its actual `(readOnly, destructive,
idempotent)`. Known corrections vs the prior draft:
- **`range_structural` is NOT idempotent** — re-`insert` duplicates rows/columns; `unmerge`
  discards merges. So `idempotent_hint` must be the default `false` (or omitted), and
  `destructive_hint` follows `unmerge`.
- **`pivot_refresh` not assumed idempotent** — re-refresh may read different source data.
- **`pivot_field` per action** — verify each (layout/position/aggregation/properties/filter).
- **`range_write` verify** — confirm set-format/set-validation/set-conditional-formatting do
  not accumulate effects before marking idempotent.
- **`vba_replace_component_source` verify** — re-call with same `expected_source_hash` likely
  fails the revision guard ⇒ not idempotent in effect.
- **`table` internal inconsistency resolved** — `table` is mixed (create vs update/append);
  it is **not** idempotent and **not** read-only; `delete-column`/`delete-rows` are extracted.

### 4.3 Resulting sets (to be finalized from the matrix)
- `_ANN_READ_ONLY`: `workbook_read`, `sheet_read`, `table_read`, `name_read`, `range_read`,
  `pivot_read`, `vba_read`, `excel_health`, `xlvba_environment`, `server_health` (inline at
  2498 — keep), `snapshot_list`, `snapshot_diff`, `job_get`.
- `_ANN_DESTRUCTIVE` (standalone + any tool with ≥1 destructive action): `sheet_delete`,
  `table_delete_rows`, `table_delete_column`, `name_delete`, `range_delete`, `range_merge`,
  `macro_run`, `excel_save_as`, `vba_replace_component_source`, `vba_insert_code`,
  `excel_recover`, `snapshot_restore`, and `range_structural` (via `unmerge`) — **verify**.
- `_ANN_IDEMPOTENT` (only if all actions idempotent, per matrix): candidates `range_write`,
  `name`, `pivot_field`, `shape`/`control` (set-on-action) — each **gated by the matrix**;
  `range_structural` is excluded.

### 4.4 Dual-hint resolution (`range_merge`, `vba_replace_component_source`)
Both are legitimately **destructive AND idempotent** (re-applying identical args converges).
Intentional; lock with an exact-triple contract test (§7). No other tool may carry both unless
the matrix proves it.

---

## 5. Token & selection measurement plan

**Approach (measured, not asserted — M2):** serialize `tools/list` before/after via the
contract-test code path (`server.mcp.list_tools()`, `test_description_contract.py:82,87`).

Baseline (in `.venv`, before WP1):
```powershell
& .\.venv\Scripts\python.exe -c "import asyncio,json,excel_vba_mcp.server as s; t=asyncio.run(s.mcp.list_tools()); data=[x.model_dump(by_alias=True) for x in t]; print('count',len(t)); print('bytes',len(json.dumps(data,default=str).encode()))"
```
After (after WP12): same command. **No dual-run** — measure only the final surface.

**Target (M2, not proven):** **~35–55%** reduction in serialized `tools/list` bytes
(92 → ~41). Record actual before/after numbers in `.ai/tmp` and reference them in the PR.
**Selection quality:** run a small benchmark (a fixed set of representative tasks) to confirm
tool-selection accuracy does not regress; report first-call success and argument errors. Do
**not** claim a number that was not measured.

---

## 6. Migration & compatibility (single switchover)

- **No dual-run / no temporary aliases (M1):** per the prior no-backward-compatibility
  decision, old names are **not** registered alongside new ones. `meta.deprecated` does not hide
  tools in mcp 2.x anyway (`.../mcpserver/server.py:516–531`, `_meta=info.meta`), so aliases
  would only inflate `tools/list`.
- **Procedure:** validate the new surface on an isolated branch/test server → update Kilo config
  and approval globs → verify → single switchover to the new names. Production `tools/list` goes
  directly from 92 to the final ~41 (no ~121–130 window).
- **Contract/test lockstep (M3):** update `test_description_contract.py` in the same change as
  the surface (see §7); `_ANN_*` (`server.py:1686–1771`) change with WP10.
- **Client-side approval globs:** Kilo approval is **name-glob based** and annotations are hints
  only (`server.py:1682–1685`). New names change the globs clients match on; document the new
  names/globs. **Approval granularity must be preserved** — see §11.1; do not merge operations
  with distinct permission profiles.
- **No general `xlvba_batch` (M4):** do not add a rollback-less batch; reuse
  `test_suite_run`'s disposable-working-copy isolation (`test_tools.py:214`,
  `_run_suite_blocking` 185) if batch is needed.

---

## 7. Verification / test plan

### Update existing (name-keyed) contracts — `tests/unit/test_description_contract.py`
- `test_tool_count_is_92_exposed` (86–89) → `test_tool_count_is_final_exposed` (assert the
  final count; **verify** exact ~41 after WP12; no dual-run assertion).
- `_CRITICAL` (18–44): replace old names with new consolidated/standalone names carrying the
  §5.1 template (`workbook`, `range_write`, `range_structural`, `table`, `vba_read`, `macro`,
  `test`, plus standalone `sheet_delete`, `table_delete_rows`, `table_delete_column`,
  `macro_run`, `excel_save_as`, `vba_replace_component_source`, `vba_insert_code`,
  `excel_recover`, `snapshot_restore`, `job_request_cancel`, `job_get`, `vba_compile_diagnostics`).
- `_EXAMPLE_TOOLS` (47–52): keep `range_write`, `table`, `vba_replace_component_source`, `test`.
- `_TRUTH` (55–59): keep `vba_compile_diagnostics`→"does NOT compile",
  `macro`→"never grants run permission" (classify-background), `job_request_cancel`→
  "intent-only; never interrupts".
- Guard list (117–137): new guarded tools (`sheet`, `table`, `name`, `range_structural`,
  `range_write`, `pivot_refresh`, `pivot_field`, `shape`, `control`) + standalone destructive
  (`sheet_delete`, `range_delete`, `range_merge`, `table_delete_rows`, `table_delete_column`,
  `name_delete`, `macro_run`, `excel_save_as`).

### New contract tests
1. **Mapping totalness:** load the WP0 ledger; assert 92 old names each map to exactly one
   target; assert the registered new-name set equals the ledger's target set.
2. **Security/approval matrix completeness:** every new tool has a documented permission
   profile; the four high-risk tools are standalone or protected; no consolidation merges two
   distinct profiles (assert against the matrix).
3. **Annotation truthfulness (matrix-based):** for each new tool assert the **exact** triple
   computed from the truth matrix; `range_structural` is not idempotent; only
   `range_merge`/`vba_replace_component_source` (and any matrix-proven tool) carry both
   destructive+idempotent.
4. **Per-action validation:** for each consolidated action with a required param (§3 table),
   omit the param and assert a structured `ExcelVbaMcpError`.

### Behavior (reuse, no change expected)
- Full `tests/unit` green.
- `tests/integration_excel/test_phase2_ranges_pivots_regression.py`, `test_ranges_service_*`,
  `test_pivot_*`, `test_tables_service.py`, `test_vba_mutation.py`, `test_test_runner*.py` —
  handlers reused, behavior unchanged. **Do not** run macros on production data.

---

## 8. Risk register

| # | Risk | Impact | Mitigation |
|---|---|---|---|
| R1 | Schema no longer marks per-action fields required (B1) | Client sends incomplete args; failure moves to runtime | Handler guards unchanged + per-action validation tests (§7.4); WP14 PoC may adopt `oneOf` |
| R2 | Over-flagged `destructive_hint` on mixed tools (B3) | Wrong approval UX; untruthful hints | Compute from truth matrix; extract destructive verbs; exact-triple test |
| R3 | Dual-hint drift on `range_merge`/`vba_replace` | Annotation inconsistency | Intentional + exact-triple test |
| R4 | `meta.deprecated` assumed to hide old tools (M1) | `tools/list` not reduced; confusion | No aliases; single switchover; measure |
| R5 | Token/selection claim unmeasured (M2) | Credibility; regression unnoticed | Measure before/after + selection benchmark; restate 35–55% as target |
| R6 | Contract tests not updated in lockstep (M3) | Red CI; stale assertions | Same-change update of `test_description_contract.py` + `_ANN_*` |
| R7 | Accidental mega range/pivot (B2) | Loss of cohesion; worse schema | Enforce 3-way split; review gate |
| R8 | Rollback-less batch footgun (M4) | Partial destructive batch, no undo | No general batch; reuse `test_suite_run` isolation |
| R9 | Count drift (92 baseline) | Off-by-one in contracts | WP0 reconciliation (91 `add_tool` + 1 `@mcp.tool`) |
| R10 | Client approval globs stale after rename | Denied/re-prompted calls | Document new globs; single switchover |
| **R11** | **Consolidation merges distinct permission profiles** | **Loss of approval granularity (security)** | **§11.1 matrix; keep `macro_run`/`excel_save_as`/`table_delete_column` (+ `close` check) standalone** |
| **R12** | **Annotation defaults misread as neutral** | **Wrong `destructiveHint`/`idempotentHint`** | **§4.1 default semantics + truth matrix; `range_structural` not idempotent** |
| **R13** | **Dual-run inflates production `tools/list`** | **Temporary 121–130 bloat; contradicts prior decision** | **Removed; single switchover (§6)** |
| **R14** | **SDK limit mistaken for MCP limit** | **Premature rejection of `oneOf` schemas** | **WP14 PoC gates the decision (§3)** |

---

## 9. Definition of done (checklist)

- [ ] WP0 baseline recorded in `.ai/tmp`: count 92, serialized bytes, mapping ledger (92 rows),
      security/approval matrix (92 tools), annotation truth matrix (every action).
- [ ] §11 pre-implementation gates resolved (approval matrix, final map, truth matrix, schema PoC,
      migration plan, benchmark).
- [ ] Consolidated tools + standalone high-risk/destructive + lifecycle tools registered
      (WP1–WP9); **no** mega range/pivot; working target ~41 (range ~38–42, not hard).
- [ ] `macro_run`, `excel_save_as`, `table_delete_column` standalone; `workbook` `close`
      disposition implemented/tested; no consolidation merges distinct permission profiles.
- [ ] Final `server.mcp.list_tools()` returns exactly the final names (single switchover; no
      `meta.deprecated`, no dual-run).
- [ ] Per-action required params enforced (handlers) with a passing per-action validation test;
      WP14 decision recorded.
- [ ] `_ANN_*` computed from the truth matrix with MCP default semantics; exact triple asserted;
      `range_structural` not idempotent; only matrix-proven dual-hint tools carry both.
- [ ] `test_description_contract.py` updated (count, `_CRITICAL`, `_EXAMPLE_TOOLS`, `_TRUTH`,
      guard list) and green; matrix-based + mapping-totalness + per-action tests added and green.
- [ ] Full `tests/unit` green; `tests/integration_excel` ranges/pivots/tables/vba/test-runner
      regression green (handlers reused).
- [ ] Token reduction **measured** (before/after) within target ~35–55% (or restated);
      selection-quality benchmark reported; numbers in PR.
- [ ] Migration note documents the single switchover and new client approval globs (no dual-run).
- [ ] No general `xlvba_batch`; `range_delete`/`range_merge`/`sheet_delete`/`table_delete_rows`/
      `table_delete_column`/`name_delete`/`macro_run`/`excel_save_as`/
      `vba_replace_component_source`/`vba_insert_code` remain standalone and destructive.
- [ ] (Optional, only if WP13 approved) `--toolsets` filter per §10; default registers all;
      union of toolsets == final set; unknown name fails fast.

---

## 10. Toolset layer (optional deployment-time filter)

> Decision status: **optional / deferred**. Implement WP13 **only if** a second client profile
> with a disjoint domain set appears; otherwise the single default surface (~41) is the end state.

### 10.1 Decision & rationale
Layered strategy, in priority order:

1. **Domain consolidation (WP1–WP9) — mandatory backbone.** Always-on, client-agnostic,
   preserves the guarded-mutation + snapshot + approval + truthful-annotation model. 92 → ~41.
2. **Namespacing — already satisfied.** Existing prefixes (`range_`, `pivot_`, `vba_`,
   `table_`, `sheet_`, `name_`) provide grouping and collision safety; the spec recommends
   server-id prefixing at aggregation. Keep prefixes on new names.
3. **Toolset filter (this section) — optional.** A process-level `--toolsets` allow-list for
   multi-profile deployments. Low priority.
4. **Rejected: single mega `call_tool`.** One external identity, one shared schema, one shared
   annotation set ⇒ cannot separate per-operation permissions/behavior. Unacceptable here.
5. **Roadmap (not rejected): Progressive Tool Discovery.** The host fetches the full
   `tools/list` but does **not** inject all definitions; it exposes a light `search_tools`, then
   loads the selected tool with its **original name, schema, and annotations** and executes that
   original tool. This does **not** destroy per-tool annotations/approval (it is distinct from a
   mega `call_tool`). **Not implemented now** only because Kilo's public docs do not yet confirm
   built-in per-tool progressive loading; revisit when the client supports it.

**Why toolsets are not the primary lever (honest constraints):**
- MCP spec (2026-07-28 Tools): `tools/list` **MUST NOT** vary per-connection; variation only by
  authorization scope. A local single-user server gains nothing from scope filtering.
- Most clients cannot select arbitrary toolsets for a third-party server; the GitHub-MCP
  `--toolsets` precedent is a **launch-process** flag (deployment choice).
- Therefore toolsets are a deployment convenience, **not** a runtime/protocol feature, and must
  never be presented as per-session filtering.

### 10.2 Toolset taxonomy (membership over final post-WP12 names) — **six toolsets**
There are **six** toolsets: `core`, `vba`, `macro-test`, `data`, `analysis`, `objects`.

| Toolset | Domain(s) | Tools (final names) |
|---|---|---|
| **core** *(always on)* | Workbook/Session, Snapshots, Jobs, Health/Env | `workbook`, `workbook_read`, `excel_save_as`, `excel_health`, `excel_recover`, `xlvba_environment`, `server_health`, `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`, `job_get`, `job_request_cancel` |
| **vba** | VBA Project | `vba_read`, `vba_replace_component_source`, `vba_insert_code`, `vba_export_project`, `vba_compile_diagnostics` |
| **macro-test** | Macro Execution & Compile, Testing | `macro`, `macro_run`, `test` |
| **data** | Sheets, Tables, Names, Ranges & Cells | `sheet`, `sheet_read`, `sheet_delete`, `table`, `table_read`, `table_delete_rows`, `table_delete_column`, `name`, `name_read`, `name_delete`, `range_read`, `range_write`, `range_structural`, `range_delete`, `range_merge` |
| **analysis** | PivotTables (future: Charts & Slicers, Data Model) | `pivot_read`, `pivot_refresh`, `pivot_field` |
| **objects** | Shapes & Controls (future: Connections & PQ, Additional Excel Objects) | `shape`, `control` |

Notes:
- `core` is the union baseline; every `--toolsets` selection implicitly includes it.
- The four **future domains** (Charts & Slicers, Connections & Power Query, Data Model,
  Additional Excel Objects) are not yet implemented; when added they join `analysis`/`objects`
  via the same `X` / `X_read` + standalone-destructive template. Do not create stubs now.
- Membership is defined over **final** names only (post-WP12).

### 10.3 Behavior & invariants
- No flag / empty value ⇒ register **all** ~41 tools (identical to today's default).
- `--toolsets data,analysis` ⇒ register `core ∪ data ∪ analysis` and nothing else.
- Pure set operation at registration; handlers reused verbatim; runtime behavior, approval gates,
  snapshots, audit unchanged.
- Unknown/misspelled toolset ⇒ startup failure listing valid names (fail fast).
- Unit test: (a) every final name in ≥1 toolset; (b) `⋃ toolsets == final set`; (c)
  `core ⊆ every resolved selection`.

### 10.4 Out of scope
- Per-connection/per-session tool variation (spec-prohibited).
- Authorization/scope-based filtering (no multi-tenant model in a local server).
- Mega `call_tool` dispatcher (rejected) and a self-hosted `search_tools` (roadmap, §10.1.5).

---

## 11. Pre-implementation gates (binding — resolve before WP1)

These six gates must be closed before implementation starts. They are the conditions of the
"accept with changes" verdict.

1. **Security/Approval matrix (WP0).** For each of the 92 tools: risk level + current Kilo
   approval policy + new name + permission profile. Prove no consolidation merges two distinct
   profiles. Special attention: `macro_run`, `excel_save_as`, `table_delete_column`, and
   `workbook` `close` (unsaved-change loss). If server-side action-level policy is desired, it
   must be independent of Kilo permissions and have its own tests — `_require_*` is validation,
   not authorization.
2. **Final consolidation map (WP0 ledger update).** Mapping ledger with standalone tools
   included. Goal is **no loss of functions or safeguards**, not a fixed count (~41 target).
3. **Annotation truth matrix (WP0 → WP10).** Per-action actual properties; fix the
   `range_structural` idempotency error; align tool-level hints with MCP default semantics (§4).
4. **Schema PoC (WP14).** Compare flat args + handler validation vs action-specific JSON Schema
   (`oneOf`) through SDK → MCP → Kilo → real LLM; decide on evidence.
5. **Migration plan (WP12).** No dual-run, no temporary aliases; isolated validation, then single
   switchover; update Kilo config and approval globs.
6. **Benchmark (§5).** Measure real token reduction and tool-selection quality; 35–55% is a
   target, not a proven result.

---

### Verification caveats (evidence)
- All cited `src/excel_vba_mcp/api/*_tools.py` line anchors and `server.py` annotation/
  registration anchors were read directly on 2026-11-26 and match the auditor's references
  (minor: auditor cited `server.py:1682–1742/1739/1769`; actual sets span **1686–1771**, with
  `range_merge` at 1733/1749 and `vba_replace_component_source` at 1739/1769 — dual-membership
  confirmed).
- MCP SDK anchors verified in `.venv/.../mcp/server/mcpserver/tools/base.py:107` and
  `.../mcpserver/server.py:516–531,618–667`.
- **Verify during WP0:** exact final tool count (target ~41, range ~38–42), the precise
  `server_health` registration path (decorator at `server.py:2498–2499`; annotations inline), and
  the `excel_close` unsaved-change behavior.