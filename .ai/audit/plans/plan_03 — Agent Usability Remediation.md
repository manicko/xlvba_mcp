# Plan 03 — Agent Usability Remediation

- **Role:** Planner (dependency-aware planning). PLAN ONLY — no production/test/config changes.
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca01fd10136716a598cad1db6b6845fed73`.
- **Binding inputs:** `report_08` (auditor), `report_08a` (validator — binding carve-outs), `research_1a/1b/2a/2b`, `agent_usability_baseline.md`, `.ai/tmp/benchmark_tasks.yaml`, `decision_log.md` §"Agent Usability Remediation (plan_03)" (D-U1..D-U10).
- **Authority order when sources conflict:** `decision_log` D-U* > `report_08a` carve-outs > `report_08` > research > baseline. Source at HEAD wins over any prose.
- **Glossary:** terms defined in §5.3 and shipped as the static MCP resource `xlvba://glossary` (D-U8) with a compact mirror inside `xlvba_environment`. All sections reference these terms; they are not redefined per section.

---

## 1. Executive summary

The `xlvba` MCP server exposes **91 tools** whose combined `tools/list` payload is **451,108 chars** (~113K tokens), of which **outputSchema is 75.05%**, inputSchema 17.9%, and descriptions only 5.4%. Source-verified usability defects are real but narrow: a small set of preconditions are undocumented (`snapshot_create` must precede guarded mutations; VBA hash source; test workbook-open), three request enums advertise values the implementation rejects, the error envelope lacks machine-readable `errorCategory`/`isRetryable`/`suggestedNextActions`, no tool reports access mode / snapshot-gate / build, and there is no glossary for jargon (R13, BG06, STA, truth signal, `SUSPECTED_STALLED`, `INTERRUPTED_UNKNOWN`).

This plan delivers five independently reviewable work packages:

- **U1 — Descriptions + glossary** (description-only + documentation-resource): rewrite the workflow-critical `_DESC_*` constants to a fixed standard, add neighbor-routing, publish the glossary resource + compact mirror. No schema change.
- **U2 — Environment/policy discovery + structured error envelope** (runtime-API + new-capability): add read-only, non-COM `xlvba_environment`; add an additive `policy` block to `excel_get_status`; extend `ExcelVbaMcpError.to_dict()` additively with `errorCategory`/`isRetryable`/`suggestedNextActions`; expose the in-memory `has_snapshot` gate. Text-JSON envelope only (no SDK change).
- **U3 — Schema/enum consistency + truthful annotations** (schema + annotations): split request vs read enums so `range_set_conditional_formatting`, `pivot_field_set_layout`, `pivot_field_apply_filter` cannot advertise rejected values; emit truthful `ToolAnnotations` from existing access-class metadata. Ships **with** the matching description edits (D-U6).
- **U4 — Token reduction via outputSchema slimming** (schema; D-U1): slim the oversized return DTOs (`WorkbookModel`, `SheetInfo`, test-suite DSL) while **keeping** `structuredContent`; document why a search tool / `defer_loading` / `structured_output=False` are rejected. No new tool.
- **U5 — Deferred product decisions** (investigation only, no implementation): produce a separate product contract outline for `save_as`/export and `xlvba_job_list` (D-U9).

**Explicitly out of scope:** any SDK change; error `structuredContent`; a custom `xlvba_search_tools`; `defer_loading`; `structured_output=False`; renaming `job_cancel`, `vba_compile`, or the test tools; collapsing the test tools; implementing `save_as`/`job_list`.

**Sequencing:** U1 ∥ U3 → U2 → U4; U5 runs in parallel and produces documents only. Full DAG in §13.

---

## 2. Verified baseline

All figures independently reproduced by the auditor at HEAD (`.ai/tmp/measure_tools_2a.py`, imports the server, no Excel started).

- **Tool count:** **91 exposed** = 90 `server.add_tool(...)` in `server.py::_register_tools` + 1 `@mcp.tool()` `server_health` (`server.py::server_health`). **4** resource templates (`server.py::workbook_model_resource`, `workbook_vba_components_resource`, `workbook_vba_component_source_resource`, and the test-run resource), all URI-templated (0 static). **0** prompts.
- **Token cost:** `tools/list` JSON = **451,108 chars** ≈ **~113K tokens** (4:1; ~129K at 3.5:1). Breakdown: **outputSchema 338,547 (75.05%)**, inputSchema 80,656 (17.9%), descriptions 24,272 (5.4%). Largest definitions: `workbook_get_model` (~28.8K output), `test_suite_run`/`test_run_suite` (~22.1K/22.3K **input** — the suite DSL), `sheet_*` (~10.9K output ×7). **Descriptions are the smallest lever; outputSchema is the dominant lever.**
- **SDK:** `mcp` 2.3.0, `mcp-types` 2.3.0, `pydantic` 2.14.4, Python 3.14.7. `add_tool(fn, name, title, description, annotations, icons, meta, structured_output)` is supported (`mcp/.../server.py::add_tool`); `ToolAnnotations{title, read_only_hint, destructive_hint, idempotent_hint, open_world_hint}` is supported but **0/91 emitted**; `outputSchema` auto-derived and emitted for **91/91**; `structuredContent` produced on **success only**; `ToolError` → `is_error=true` **text-only** (`mcp/.../server.py` `call_tool`); `defer_loading`/`deferLoading`/`search_tools` = **0 hits** in the installed stack.
- **Access model:** `AccessMode ∈ {FULL, RESTRICTED}`, default **FULL** (`config.py::ServerConfig.access_mode`). `READ_ONLY`→auto-allowed; `FULL`→MUTATING/EXECUTING allowed by policy (no per-action grant); `RESTRICTED`→deny-by-default (allowlist/live grant); `RECOVERY`→auto-allowed in both, pid-scoped (`security/approval.py::ApprovalPolicy.check`/`check_recovery`). Snapshot / writability / PathPolicy / Office-trust / audit run **independently** of mode. `access_mode` is discoverable via `excel_get_status` (`session_service.py` `SessionStatus.access_mode` populated), **not** via `server_health`, **not** in `_DESC_STATUS`.
- **Error payload today:** `errors.py::ExcelVbaMcpError.to_dict` = `{code, message, hresult, hresult_hex, context}`; surfaced as a JSON string inside `ToolError` by `server.py::ExcelVbaMcpServer.call_tool`. No `errorCategory`/`isRetryable`/`suggestedNextActions` exist at the error level; only job-level `retryable` exists.
- **Snapshot gate:** `snapshot_service.py::SnapshotService.require_snapshot` reads the in-memory `_latest_snapshot`; `has_snapshot` exists but is **not exposed**; `snapshot_list` lists the on-disk **store**, not the gate; guarded mutations do **not** auto-snapshot.

---

## 3. Findings and severity

Classification from `report_08` §2 (source-verified at HEAD). "Severity" is the audit's; the plan acts on CONFIRMED/PARTIAL items only.

| id | tool / symbol | classification | severity | one-line reality | handled by |
|---|---|---|---|---|---|
| A01 | `excel_open` | PARTIALLY_CONFIRMED (shape part REFUTED) | Blocker→low | Shape is in `WorkbookInfo` + auto outputSchema; gap = prose omits defaults + writability truth | U1 |
| A02 | `range_write_values` | CONFIRMED | Blocker→low | Never says "call `snapshot_create` first"; payload shape only in schema | U1 |
| A03 | `vba_compile` | CONFIRMED (naming, not doc) | Blocker→low | Description honest; **name** misleads; keep name (D-U5) | U1 |
| A04 | `test_run_suite`/`test_suite_run` | CONFIRMED (+doc deviation) | Blocker→low | Near-duplicate; docs wrongly claim `test_run_suite` accepts `background`; keep both (D-U5) | U1 |
| A05 | `range_set_conditional_formatting` | CONFIRMED | Blocker→med | Request enum advertises rejected `color_scale/data_bar/icon_set/unsupported` | U3 (+U1) |
| A06 | `pivot_field_set_layout` | CONFIRMED (overload documented) | Major→med | `orientation='data'` rejected; enum advertises it | U3 (+U1) |
| A07 | `pivot_field_apply_filter` | CONFIRMED | Major→med | Only `manual` applied; enum advertises label/value/date/top_bottom | U3 (+U1) |
| A08 | `job_cancel` | CONFIRMED (accurate) | Major→low | Description accurate; name over-promises; keep name (D-U4) | U1 |
| A09 | `snapshot_restore`/`excel_close`/`excel_recover` | PARTIALLY_CONFIRMED | Major→low | Descriptions disclose no-save/discard; "silent" overstated | U1 |
| A10 | `macro_run` | PARTIALLY_CONFIRMED | Major→low | Gates described but not numbered; "approved" undefined | U1 |
| A11 | `macro_classify_background` | PARTIAL→leaning REFUTED | Major→low | Already says "UNGATED bookkeeping"; add "never grants run permission" | U1 |
| A12 | `table_update_rows`/`table_delete_rows` | CONFIRMED (minor) | Major→low | `where`/`updates` grammar + ambiguity not in description | U1 |
| A13 | `name_create`/`name_update` | CONFIRMED (minor) | Major→low | `=` optional; "cross-scope conflict" undefined | U1 |
| A14 | `range_read_values` | PARTIAL (offset REFUTED; caps CONFIRMED) | Major→low | 0-based stated; three numeric caps not stated | U1 |
| A15 | `vba_replace_component_source`/`vba_insert_code` | CONFIRMED (minor) | Major→low | Hash source not named | U1 |
| A16 | `sheet_add`/`sheet_copy` | CONFIRMED (minor) | Minor | `sheet_type` default + generated copy name not stated | U1 |
| A17 | `server_health` vs `excel_health` | PARTIALLY_CONFIRMED | Minor | Distinct; `server_health` terse, no redirect | U1 |
| A18 | mode/policy discoverability | PARTIAL (core REFUTED) | systemic | `excel_get_status` **does** report mode; missing in `server_health`/`_DESC_STATUS` | U2 |
| A19 | snapshot precondition | CONFIRMED | systemic | `has_snapshot` unexposed; mutation doesn't auto-snapshot | U2 (+U1) |
| A20 | build/version opacity | CONFIRMED | systemic | No 4-part build source; `build` must be honest `null` | U2 |
| A21 | undefined jargon | CONFIRMED | systemic | No glossary | U1 |
| A22 | duplication / no routing | CONFIRMED | systemic | No "prefer X" guidance | U1 |
| A23 | persistence gap (`save_as`) | CONFIRMED | systemic | No save-as/export tool | U5 (defer) |
| A24 | destructive-by-default | PARTIALLY_CONFIRMED | systemic | Disclosed; no `confirm` flag | U1 |
| A25 | cancellation trap | CONFIRMED (accurate) | systemic | Name over-promises; keep name (D-U4) | U1 |

**Scoreboard (validator §1):** PASS 2 · PARTIAL 5 · FAIL 6 · N/A 1 — remediation justified (6 hard failures + 5 partials, all source-confirmed).

---

## 4. Non-negotiable architecture and safety constraints

These are invariants from `decision_log` (both sections) and `report_08` §5.2. **No work package may weaken them.**

1. **Single hidden Excel + single-STA FIFO worker.** No new COM path; `xlvba_environment` is non-COM (like `server_health`).
2. **Originals never modified; working-copy model.** All mutations target a working copy; originals are hashed before/after.
3. **Agent-directed `excel_recover` only; server never auto-terminates a live pid.** `RECOVERY` auto-allowed in both modes, pid-scoped, audited, discards working copies only.
4. **`INTERRUPTED_UNKNOWN` is never auto-retried**; `outcome='unknown'`, `retryable=false`, never completed, never replayed.
5. **No pivot output materialization; no `Workbook.RefreshAll`.** Pivot refresh is table/cache-scoped, truth-signal confirmed.
6. **No mandatory human approval in normal flows.** FULL is default; the remediation only *reports* policy.
7. **`SCHEMA_VERSION = 1` additive.** All DTO additions are new optional fields on **new or existing model classes** (respect `models.py::DomainModel` `extra="forbid"`); no ad-hoc payloads. No new `ErrorCode` members.
8. **Fail-closed everywhere.** Unknown/stale/indeterminate states are explicit non-success; `xlvba_environment` never fabricates a build.
9. **Snapshot gate, approval policy, PathPolicy, Office-trust, BG06 + R13 gates, audit** are unchanged and run independently of access mode.
10. **No SDK change; no error `structuredContent`; no search tool; no `defer_loading`; no `structured_output=False`** (D-U1/D-U2, validator §2.2).
11. **Truthful descriptions/annotations only.** No claim the implementation does not provide; `vba_compile` never described as compiling; `macro_classify_background` never described as granting execution; `job_cancel` never described as interrupting.
12. **`isRetryable` is conservative and per-code** (D-U10). `STALE_TARGET`/`REVISION_CONFLICT` are not blanket-retryable; `INTERRUPTED_UNKNOWN` never retryable.

---

## 5. Target tool-description standard

### 5.1 Required section order (workflow-critical tools)

Every `_DESC_*` constant for a Blocker/Major/workflow-critical tool MUST follow this fixed template (sections omitted only when genuinely empty, never reordered):

1. **Purpose** — one sentence, verb-first, what the tool does.
2. **Use when** — the concrete situations that make this the right tool.
3. **Do not use when** — the neighbor tool(s) to prefer instead (see §9 routing).
4. **Preconditions** — writable session / snapshot / approval / open workbook / execute-mode, stated explicitly (e.g. "does NOT snapshot — call `snapshot_create` first").
5. **Parameters** — the payload object shape and numeric ceilings where the schema alone is insufficient (e.g. `{rows, cols, values}`, `rows≤100000`, `cols≤512`, cells≤100000).
6. **Behavior** — ordering, renumbering, rejection rules, truth-signal confirmation.
7. **Result** — the returned DTO name and key fields.
8. **Next actions** — the follow-up tool(s) to verify or continue.

### 5.2 Rules

- **R-D1 (no fabrication):** never state a capability the code lacks; `build` is `null`; `vba_compile` "diagnostic only — does not compile".
- **R-D2 (consistency with schema):** descriptions MUST match the **pruned** request enums from U3; the three enum tools' prose ships in U3, not U1.
- **R-D3 (one guard sentence):** standardize the guard phrasing to: *"Guarded mutation: writable session + snapshot + approval (full-access default; restricted requires approval) + postcondition re-read."* Replaces the current mix of "requires a writable working copy, a successful snapshot".
- **R-D4 (define "approved"):** `macro_run` "approved" = the call-time `ApprovalPolicy.check(EXECUTING)` outcome (allowed by default in FULL; allowlist/live grant in RESTRICTED). It is **not** the BG06 classification.
- **R-D5 (glossary by reference):** descriptions reference `xlvba://glossary` for jargon (R13, BG06, STA, truth signal, fail-closed, `SUSPECTED_STALLED`, `INTERRUPTED_UNKNOWN`, `indeterminate`, normalized source hash) rather than re-explaining.
- **R-D6 (token discipline):** descriptions are ~5.4% of payload; keep additions terse. Do not bloat prose to compensate for schema (that is U4's job).
- **R-D7 (neighbor routing):** every "Do not use when" names the concrete alternative (§9 table).
- **R-D8 (examples):** add one request example to the workflow-critical tools (write-values, table-update, vba-replace, test-suite) — the only place examples are added.

### 5.3 Glossary (authoritative term set — shipped as `xlvba://glossary`)

The glossary resource defines, at minimum: **R13** (heuristic interactive-UI token veto; a veto, never a certificate), **BG06** (background-safe classification bound to `(workbook hash, source hash, procedure)`; invalidated on VBA source change), **preflight** (the R13 scan before dispatch), **truth signal** (Excel's own success evidence: `RefreshTable` Boolean / advanced `RefreshDate`; never "absence of a COM exception"), **fail-closed** (unknown/ambiguous ⇒ non-success, never fabricated success), **RECOVERY** (auto-allowed agent-directed recovery class), **STA** (single-threaded apartment — the one serialized Excel thread), **Stage 1** (the typed-DTO contract layer; `SCHEMA_VERSION=1`), **normalized source hash** (SHA-256 of the canonicalized component body; the `expected_source_hash` revision key), **`SUSPECTED_STALLED`** (soft-deadline exceeded; op left running; not confirmed hung), **`interrupted_unknown`** (job finalized by recovery; never replayed), **`indeterminate`** (compile probe never returns compiled/compile_error), **working copy** (the opened, mutable copy; originals are never touched), **snapshot gate** (in-memory `_latest_snapshot` checked by `require_snapshot`).

---

## 6. Target agent-facing environment/policy contract

### 6.1 New read-only tool `xlvba_environment` (authoritative server-level source)

- **Type:** new-capability, **non-COM**, **side-effect-free**, `readOnlyHint:true`. Must NOT create snapshots, grants, or classifications; must NOT imply it authorizes mutations. `policy.classification_is_authorization=false` is stated explicitly.
- **Registration:** `server.py::_register_tools` via `server.add_tool(xlvba_environment, name="xlvba_environment", description=_DESC_ENVIRONMENT, annotations=ToolAnnotations(read_only_hint=True))`. Handler in a new `api/environment_tools.py` (composition; no logic in `server.py`).
- **Payload (new `EnvironmentReport` DTO in `models.py`, `extra="forbid"`, all additive; `SCHEMA_VERSION` stays 1):**

```jsonc
{
  "schema_version": 1,
  "server": {
    "name": "excel-vba-mcp",
    "version": "0.1.0",                 // __version__
    "access_mode": "full",              // AccessMode
    "macro_timeout_s": 90.0, "open_timeout_s": 300.0,
    "liveness_ceiling_s": 600.0, "log_level": "INFO"
  },
  "excel": {                            // null when worker not started
    "available": true,
    "application_version": "16.0",      // Application.Version (NOT the build)
    "build": null,                      // honest null; no 4-part source exists
    "target_build": "16.0.14332",       // labelled design-time constant only
    "owned_pid": 12345,
    "vbom_available": true              // from model vba_access, else null
  },
  "capabilities": {
    "vba": true, "macros": true, "pivot_tables": true,
    "background_jobs": true, "recovery": true,
    "non_activating_compile": false,    // always false on validated build
    "pivot_filter_kinds": ["manual"],
    "pivot_source_types": ["database"]
  },
  "policy": {
    "approval": { "mode": "full", "mutating_requires_grant": false,
                  "executing_requires_grant": false, "recovery_auto_allowed": true },
    "snapshot": { "required_for_mutations": true, "created_by": "agent:snapshot_create" },
    "path": { "allowed_roots": ["<root1>"], "allow_unc": false },
    "macro": { "background_safe_required": true, "r13_preflight": "veto",
               "classification_is_authorization": false }
  },
  "glossary_uri": "xlvba://glossary",
  "workbooks": [                         // per-open-workbook summary
    { "workbook_id": "...", "read_only": false, "mode": "execute",
      "has_snapshot": false, "dirty": false }
  ]
}
```

- **Field provenance (all HIGH):** `access_mode` `config.py::ServerConfig`; timeouts/log_level `config.py::ServerConfig`; version `excel_vba_mcp/__init__.py::__version__`; `application_version` `workbook_adapter.py` (`Application.Version`); `owned_pid` worker status / `health_tools.py`; capabilities from tool registration + `statuses.py` enums; compile `compile_service.py`; path roots `path_policy.py`; macro gates `macro_service.py`; snapshot `snapshot_service.py`.
- **`has_snapshot` MUST read the in-memory gate** `SnapshotService.has_snapshot(workbook_id)` (i.e. `_latest_snapshot`), **not** `SnapshotStore.list()` (D-U3, validator §2.2(5)). Reporting the store is a correctness bug.
- **`build` MUST be `null`**, never fabricated `14332`; `target_build` is a labelled design constant.

### 6.2 Additive `policy` block on `excel_get_status`

- Add one **optional** field to `models.py::SessionStatus` (class already has `access_mode`; `extra="forbid"` ⇒ the field must be added to the class, not passed ad-hoc): `policy: SessionPolicy | None = None`.
- `SessionPolicy` (new DTO) carries the **per-workbook** subset: `approval_mode`, `mutating_requires_grant`, `executing_requires_grant`, `recovery_auto_allowed`, `snapshot_required_for_mutations`, `has_snapshot` (in-memory gate), `classification_is_authorization=false`, `r13_preflight="veto"`.
- `session_service.py` populates it where `SessionStatus` is built (alongside `access_mode`). `xlvba_environment` remains the authoritative server-level source; this block co-locates per-workbook context.

### 6.3 `server_health` (unchanged behavior, description only)

`server_health` stays a pure liveness probe (`{status, service, version}`); its description is updated (U1) to say "liveness only; use `excel_health` for worker/session state and `xlvba_environment` for access mode / policy / build."

---

## 7. Standard structured error contract

### 7.1 Form (D-U2; validator carve-out)

- On `isError=true`, the server returns a **text-content JSON envelope** (the existing `server.py::ExcelVbaMcpServer.call_tool` already `json.dumps(cause.to_dict())`). **No `structuredContent` on error** (SDK `ToolError` is text-only; infeasible without an SDK change — out of scope).
- `errors.py::ExcelVbaMcpError.to_dict` is extended **additively** (existing keys kept; `SCHEMA_VERSION` stays 1; no new `ErrorCode` members):

```jsonc
{
  "code": "SNAPSHOT_FAILED",            // existing ErrorCode
  "errorCategory": "missing_prerequisite",
  "message": "A successful snapshot is required before mutating this workbook.",
  "isRetryable": false,
  "suggestedNextActions": ["snapshot_create(workbook_id) then retry the mutation"],
  "details": { "workbook_id": "..." },  // = existing context
  "hresult": null,
  "hresult_hex": null
}
```

- Implementation: three **pure functions** in `errors.py` — `category_for(code)`, `retryable_for(code)`, `actions_for(code, context)` — called inside `to_dict`. No COM, no schema bump.

### 7.2 Category taxonomy (mutually exclusive, additive)

| errorCategory | meaning | example codes |
|---|---|---|
| `permission_denied` | policy/approval/trust denial; not fixable by retry | `APPROVAL_DENIED`, `MACRO_BLOCKED`, `PATH_NOT_ALLOWED`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` |
| `missing_prerequisite` | required precondition absent | `SNAPSHOT_FAILED`, `WORKBOOK_READ_ONLY`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `REVISION_CONFLICT`, `STALE_TARGET` |
| `unsupported_feature` | capability absent on this build/config | `UNSUPPORTED_PIVOT_SOURCE`, `INTERACTIVE_UI_NOT_SUPPORTED`, `COMPILE_INDETERMINATE`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `AMBIGUOUS_FILTER`, `PROTECTED_TARGET` |
| `transient_com_failure` | Excel/VBE busy; retry after backoff | `COM_BUSY`, `RPC_E_CALL_REJECTED` |
| `deadline` | soft deadline exceeded; op left running | `MACRO_TIMEOUT` (SUSPECTED_STALLED), job `expired` |
| `suspected_stall` | deadline + possible modal; agent-directed recovery | `EXCEL_HUNG`, `MODAL_UI_BLOCKED` |
| `interrupted_unknown` | job finalized by recovery; never replay | job `interrupted_unknown` |
| `not_found` | caller addressed a missing identity | `WORKBOOK_NOT_FOUND`, `MACRO_NOT_FOUND`, `NOT_FOUND` |
| `validation` | caller argument/schema error | pydantic `ValidationError`→`INVALID_RANGE` |

### 7.3 `isRetryable` + `suggestedNextActions` per code (D-U10 conservative)

| code | isRetryable | suggestedNextActions |
|---|---|---|
| `COM_BUSY` / `RPC_E_CALL_REJECTED` | **true** | `["wait retry_after_s (if present) then retry the same read"]` |
| `SNAPSHOT_FAILED` | false | `["snapshot_create(workbook_id)", "retry mutation"]` |
| `WORKBOOK_READ_ONLY` | false | `["re-open with read_only=false (mode=execute)"]` |
| `APPROVAL_DENIED` | false | `["open in full mode or supply grant/allowlist entry", "retry"]` |
| `PATH_NOT_ALLOWED` | false | `["use a path under details.allowed_roots"]` |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | false | `["macro_classify_background(...background_safe=true) after review", "retry macro_run"]` |
| `INTERACTIVE_UI_NOT_SUPPORTED` | false | `["remove interactive-UI tokens (out of scope)"]` |
| `UNSUPPORTED_PIVOT_SOURCE` / `COMPILE_INDETERMINATE` | false | `["use a supported source (database) / accept indeterminate compile"]` |
| **`STALE_TARGET` / `REVISION_CONFLICT`** | **false** | `["re-read source/hash via vba_read_component", "the prior mutation MAY have partially applied — re-read state; do NOT blind-replay"]` |
| `EXCEL_HUNG` / `MODAL_UI_BLOCKED` | false | `["excel_recover", "re-open workbook", "do not replay the interrupted mutation"]` |
| `MACRO_TIMEOUT` (SUSPECTED_STALLED) | false | `["excel_recover if still stalled", "never auto-replay; inspect macro_get_last_result"]` |
| `WORKBOOK_NOT_FOUND` / `NOT_FOUND` / `MACRO_NOT_FOUND` | false | `["verify workbook_id / object name / procedure"]` |
| `INVALID_RANGE` (validation) | false | `["fix arguments to match input schema"]` |
| job `interrupted_unknown` | **false** | `["never resubmit; report status"]` |

> **D-U10 resolution:** `STALE_TARGET`/`REVISION_CONFLICT` are `isRetryable=false` **with an explicit "re-read; prior mutation may have applied" caveat** — never a blanket `true` that invites blind replay of a partially-applied mutation.

---

## 8. Job-state and cancellation contract

- **States:** `JobState ∈ {queued, running, succeeded, failed, expired, unknown, interrupted_unknown}` (`job_registry.py::JobState`); terminal set = `{succeeded, failed, expired, unknown, interrupted_unknown}`.
- **`job_cancel` is intent-only (D-U4; name unchanged):** sets `cancel_requested=True` only if non-terminal; the single-STA COM call runs to completion; a poller observes `cancel_requested` via `job_get`. Description states "intent-only; never interrupts"; to actually stop a wedged process use `excel_recover` (agent-directed). **No rename.**
- **`SUSPECTED_STALLED` ≠ confirmed hung:** soft-deadline exceeded; owned pid left running; never terminated; never auto-replayed. Distinct from a confirmed exit.
- **`INTERRUPTED_UNKNOWN`:** produced by `mark_interrupted` on a successful `excel_recover`; definitive, `outcome='unknown'`, `retryable=false`, never completed, never replayed; a late STA completion cannot overwrite it. **Never auto-retried.**
- **Job handle contract:** long ops return `{status:"running", job_id, poll_after_s}` with `poll_after_s ∈ [1,30]`; callers re-call `job_get` no sooner than `poll_after_s`. Busy-spin faster than `poll_after_s` is a benchmark UC/UAA violation.
- **Descriptions updated (U1):** `job_get` (state semantics + interrupted_unknown), `job_cancel` (intent-only + observe via `job_get` + `excel_recover`), `macro_get_last_result` (SUSPECTED_STALLED semantics). Glossary defines the states.

---

## 9. Tool naming and alias strategy

- **No renames.** Keep `job_cancel` (D-U4), `vba_compile` (D-U5), `test_run_suite` **and** `test_suite_run` (D-U5). No aliases are introduced (none exist today; none needed).
- **Test tools routing (description-only):** canonical for **new** use = `test_suite_run` (accepts `TestSuiteSpec | dict` + `background`); `test_run_suite` = legacy typed-spec blocking runner. Each description's "Do not use when" points at the other. Fix docs that wrongly attribute `background` to `test_run_suite`.
- **Health tools routing:** `server_health` = liveness only; `excel_health` = worker/session state; `xlvba_environment` = access mode / policy / build / capabilities.
- **`xlvba_*` vs `excel-mcp_*`:** descriptions add a neighbor line where overlap exists (ranges/sheets/tables/names): "prefer the `xlvba_*` tool for typed DTOs + structured errors; `excel-mcp_*` is the alternate client." No behavioral change.
- **`vba_compile`:** keep name; first line sharpened to "diagnostic only — does not compile"; "Do not use when you need a real compile result (none available)."
- **Neighbor-routing table (authoritative; used in "Do not use when"):**

| Instead of | Use | Why |
|---|---|---|
| `range_write_values` (formula) | `range_write_formulas` | `=...` in write_values is stored as text |
| `range_set_conditional_formatting` (color/data/icon) | `range_get_conditional_formatting` (read, `supported=false`) | those rule types are rejected on write |
| `pivot_field_set_layout` (`orientation="data"`) | `pivot_field_set_aggregation` | data area is `AddDataField` |
| `pivot_field_set_layout` (reorder) | `pivot_field_set_position` | layout sets bucket; position sets order |
| `pivot_field_apply_filter` (label/value/date/top) | `pivot_field_get_filters` (read) | only `manual` is applied |
| `test_run_suite` (raw JSON / background) | `test_suite_run` | only `test_suite_run` accepts a mapping + `background` |
| `test_suite_run` (typed spec, blocking) | `test_run_suite` | canonical blocking runner |
| `vba_compile` (expecting compiled/compile_error) | `vba_get_compile_diagnostics` | compile never returns compiled/error |
| `job_cancel` (expecting interruption) | `job_get` (observe) + `excel_recover` | cancel is intent-only |
| `sheet_add` (non-worksheet) | — (unsupported) | only `"worksheet"` supported |
| `range_insert`/`range_delete` over a Table | table tools | ListObject intersection is rejected |
| `server_health` (need mode/policy) | `xlvba_environment` / `excel_health` | liveness ≠ policy |

---

## 10. Progressive tool discovery strategy

- **Decision (D-U1; validator U4 REJECT):** do **not** add a custom `xlvba_search_tools`, do **not** rely on `defer_loading`, do **not** use `structured_output=False`.
  - `defer_loading`/`search_tools` are **absent** from `mcp`/`mcp-types` 2.3.0 (0 hits); Kilo loads all tools at once.
  - A standalone search tool **adds a 92nd always-loaded definition** and removes nothing → **net context increase**.
  - `structured_output=False` also removes `structuredContent` on success (`func_metadata.convert_result` returns content-only) — rejected because clients may parse the typed DTO from `structuredContent`.
- **Real lever = outputSchema slimming (U4):** slim the oversized return DTOs while **keeping** `structuredContent`:
  - `WorkbookModel` (~28.8K output) — reduce embedded `$defs` duplication / drop unused nested detail.
  - `SheetInfo` (~10.9K ×7) — the repeated per-sheet schema is the single biggest repeated cost; factor shared `$defs`.
  - test-suite DSL input (~22K ×2) — dedupe the `TestSuiteSpec`/`TestStep`/`Assertion` schema between the two test tools where the SDK permits (shared `$ref`), without changing the closed vocabulary.
  - **Constraint:** additive-safe; verify no consumer relies on removed schema fields (inventory is unknown — see §18); do not remove fields that carry `structuredContent` semantics.
- **Resources:** keep the 4 existing per-workbook resources; add the static glossary resource `xlvba://glossary` (D-U8). Resource-based progressive disclosure is **doc-supported but not confirmed** in the live Kilo build (meta-tools not observed) — the compact glossary mirror inside `xlvba_environment` is the reliable channel; the resource is the forward-compatible channel.
- **Why not a dispatcher:** a full `xlvba_call(tool_name, arguments)` collapsing 90→2 would reduce context but erodes per-tool `{server}_{tool}` permissions and client-side input-schema validation; out of scope unless explicitly mandated (it is not).

---

## 11. MCP annotation strategy

- **Emit truthful `ToolAnnotations` (D-U7; currently 0/91)** derived from existing access-class metadata, at `add_tool(... annotations=...)`:
  - `readOnlyHint: true` — read-only tools (reads, lists, `server_health`, `excel_health`, `xlvba_environment`, `job_get`, `snapshot_list`, `test_get_run`, `macro_get_last_result`, `vba_read_component`, `vba_list_*`, `pivot_*_list/get`, `range_read_*`, `range_get_*`, `table_read_rows`, `sheets_list`, `sheet_get`, `names_list`, `name_get`, `shapes_list`, `shape_get`, `controls_list`, `control_get`, `workbook_get_model`, `excel_get_status`, `excel_list_workbooks`).
  - `destructiveHint: true` — destructive ops (`sheet_delete`, `range_merge` [data loss], `excel_recover` [discards working copies], `snapshot_restore` [discards unsaved], `table_delete_rows`, `name_delete`, `range_delete`, `vba_replace_component_source`, `vba_insert_code`).
  - `idempotentHint: true` where a re-read verifies the same end state (most guarded mutations with postcondition re-read; `range_write_values`, `table_update_rows`, `pivot_field_set_*`, `name_create/update`, `sheet_rename`, `shape_set_on_action`, `control_set_on_action`).
  - **Never** annotate `macro_run` or `vba_compile` as read-only; `macro_run` is `destructiveHint` unset (execution, not environment-destructive) but **not** read-only.
- Annotations are **hints** (SDK doc): client display is best-effort; Kilo approval is name-glob based, so annotations do not change enforcement — they are truthful metadata only.
- **Non-goal:** do not add `title=` (0 today; not required by any finding; avoid churn).

---

## 12. Description/schema changes by tool

Change types: **D**=description-only, **R**=documentation-resource, **S**=schema, **A**=annotations, **N**=runtime-API/new-capability, **L**=tool-lifecycle (deferred). "Exact change" is the semantic delta; implementers edit the named symbol.

### 12.1 Description-only (U1)

| tool | type | exact change |
|---|---|---|
| `excel_open` | D | Add: "Defaults: `mode='execute'`, `read_only=False` (writable unless `read_only=true`). Path must be under an allowlisted root (`EXCEL_VBA_MCP_ROOTS`, default cwd+temp)." Add example. |
| `excel_close` | D | Bold "Defaults to NOT saving — unsaved edits are lost unless `save_changes=true`." |
| `excel_save` | D | "Saves the working copy **in place** (no path argument); there is no save-as (see `xlvba://glossary` / §18)." |
| `range_write_values` | D | "This tool does **NOT** snapshot — call `snapshot_create` first." Add payload `{rows, cols, values}` + ceilings `rows≤100000`, `cols≤512`, cells≤100000; `address`=top-left anchor; leading `=` stored as text. Add example. |
| `range_read_values` | D | Publish caps: request rejected `>100000` cells; response capped `50000` cells / `8 MiB`; `truncated=true` marks a cut-short window. |
| `range_write_formulas` | D | Neighbor line: "for literal `=` text use `range_write_values`." |
| `range_set_conditional_formatting` | D (in U3) | Align prose with pruned request enum; "color_scale/data_bar/icon_set/unsupported are not accepted." |
| `pivot_field_set_layout` | D (in U3) | Align prose with pruned settable enum (`data` rejected). |
| `pivot_field_apply_filter` | D (in U3) | Align prose with pruned request enum (only `manual`). |
| `macro_run` | D | Numbered gate list (1 execute-mode+writable, 2 resolve, 3 approval `check(EXECUTING)`, 4 BG06, 5 R13 preflight, 6 dispatch); define "approved" (R-D4); add "≤30 args; bare `datetime.date` rejected — pass `datetime.datetime`." |
| `macro_classify_background` | D | Add "Classification is bookkeeping only and **never grants run permission**; `macro_run` still enforces approval + BG06 + R13." |
| `table_update_rows` | D | Publish grammar: `updates=[{index\|where, values}]` (exactly one selector); `where` must match exactly one row else `AMBIGUOUS_FILTER`; add example. |
| `table_delete_rows` | D | Same grammar/ambiguity note. |
| `table_append_rows` | D | "`rows` is a non-empty list of `{header: value}`; headers must match table columns (case-insensitive)." |
| `name_create` | D | "`refers_to` may be given with or without a leading `=` (normalized); duplicate/cross-scope conflicts are rejected." |
| `name_update` | D | Same `=`-optional note. |
| `sheet_add` | D | "`sheet_type` defaults to `"worksheet"`; only `"worksheet"` is supported (others rejected)." |
| `sheet_copy` | D | "The copy's name is generated by Excel (typically `<name> (2)`)." |
| `snapshot_restore` | D | "Unsaved changes in the current working copy are **discarded** (closed without saving)." |
| `excel_recover` | D | Bold "discards working copies; originals never modified." |
| `job_cancel` | D | "Intent-only; never interrupts; observe `cancel_requested` via `job_get`; use `excel_recover` to actually stop." |
| `job_get` | D | State semantics incl. `interrupted_unknown` (never auto-retried) and `SUSPECTED_STALLED` (see glossary). |
| `vba_compile` | D | First line "diagnostic only — does **not** compile"; keep name (D-U5). |
| `vba_replace_component_source` | D | "Obtain `expected_source_hash` from `vba_read_component` / `vba_list_components`." |
| `vba_insert_code` | D | Same hash-source note. |
| `server_health` | D | "Liveness only; use `excel_health` for worker/session state and `xlvba_environment` for access mode / policy / build." |
| `test_run_suite` | D | Routing + precondition: "`suite.workbook_id` must be an **open** workbook (call `excel_open` first); a baseline snapshot is created automatically unless `baseline='use_existing'`. Legacy typed-spec blocking runner; for raw JSON/`background` use `test_suite_run`." |
| `test_suite_run` | D | Routing + precondition (same); "canonical for new use; accepts spec or mapping + `background`." |
| `excel_get_status` | D | Mention the additive `policy` block (U2) and that `access_mode` is reported. |
| all guarded-mutation tools | D | Apply R-D3 one-sentence guard phrasing consistently (F19). |
| workflow-critical tools | D/R | Add one request example (R-D8) and glossary reference (R-D5). |

### 12.2 Schema / enums (U3; ships with the matching description edits)

| tool | type | exact change |
|---|---|---|
| `range_set_conditional_formatting` | S | In `statuses.py`, add request enum `ConditionalFormatRequestType` = plain `FormatCondition` set only (`cell_value, expression, text_string, blanks, no_blanks, errors, no_errors, time_period, top10, above_average, unique_values`); type the request `spec.type` against it in `models.py`. **Keep** `ConditionalFormatType` (full) for read/classify. |
| `pivot_field_set_layout` | S | Add request enum `PivotFieldLayoutOrientation` = `hidden, row, column, page`; type the request `orientation` against it. Keep `PivotFieldOrientation` (full, incl. `data`) for read. |
| `pivot_field_apply_filter` | S | Add request enum `PivotFilterRequestKind` = `manual` (and `none` where the request needs it); type `PivotFilterSpec.kind` against it. Keep `PivotFilterKind` (full) for read/classify. |
| all tools | A | Emit truthful `ToolAnnotations` per §11. |

> Pruned values are still rejected at runtime, but the **advertised** request contract narrows (validator Q14). This is a **versioned request-contract change**, documented in §17. Both old/new reject the pruned values; the difference is the schema no longer advertises them.

### 12.3 Runtime-API / new-capability (U2)

| symbol | type | exact change |
|---|---|---|
| `xlvba_environment` (new) | N | New non-COM read-only tool + `api/environment_tools.py` + `EnvironmentReport` DTO; payload §6.1. |
| `models.py::SessionStatus` | N | Add optional `policy: SessionPolicy \| None`; new `SessionPolicy` DTO (§6.2). |
| `session_service.py` (SessionStatus builder) | N | Populate `policy` (incl. `has_snapshot` via `SnapshotService.has_snapshot`). |
| `errors.py::ExcelVbaMcpError.to_dict` | N | Additive `errorCategory`/`isRetryable`/`suggestedNextActions` via pure functions (§7). |
| `snapshot_service.py::has_snapshot` | N | Expose the in-memory gate to the environment/status builders (no behavior change to the gate). |

### 12.4 Documentation-resource (U1/U4)

| symbol | type | exact change |
|---|---|---|
| `xlvba://glossary` (new static resource) | R | New `@mcp.resource("xlvba://glossary", ...)` returning the §5.3 glossary (markdown). |
| `docs/05-mcp-tools-reference.md`, `docs/07-workflows-and-examples.md`, `docs/01-overview.md`, `docs/ru/*` | R | Fix `test_run_suite` `background` mis-attribution; add test-tool routing; reflect pruned enums; add `xlvba_environment`; keep tool-count language accurate (91 exposed, `>=87` assert). |

### 12.5 Tool-lifecycle (U5 — DEFERRED, no implementation)

| symbol | type | exact change |
|---|---|---|
| `save_as` / export | L | Investigation + separate product contract only (§18). No tool added. |
| `xlvba_job_list` | L | Investigation + contract outline only (§18). No tool added. |

---

## 13. Dependency DAG

```
U1 (descriptions + glossary resource)  ──┐
U3 (enum schema prune + 3 enum desc +    ──┼──► U2 (xlvba_environment + policy block
     annotations) [parallel with U1]      │      + error envelope + has_snapshot)
                                          │
U1 ──────────────────────────────────────┘   (U2 needs the glossary mirror/uri from U1)

U1 + U2 + U3 ──► U4 (outputSchema slimming; needs final descriptions + schemas + env DTO)

U5 (deferred investigation + product contract)  ── independent, parallel, docs only
```

- **U1 ∥ U3** (no cross-dependency; the 3 enum-tool descriptions live in U3 so schema+prose ship together, D-U6).
- **U2 after U1** (glossary mirror + `glossary_uri`).
- **U4 after U1+U2+U3** (slim DTOs only once descriptions/schemas/environment are final; verify no `structuredContent` consumer breaks).
- **U5 anywhere** (produces documents only).
- **Hard ordering rule:** never ship a description that contradicts a not-yet-pruned enum, and never prune an enum without its description (D-U6).

---

## 14. Work packages and implementation order

Each package is independently reviewable and atomic. "Changed files" are targets; implementers use semantic anchors, not line numbers.

### U1 — Descriptions + glossary (description-only + documentation-resource)

- **Scope:** rewrite the `_DESC_*` constants in §12.1 to the §5 standard; add neighbor-routing (§9); add examples (R-D8); publish `xlvba://glossary` (§5.3) + compact mirror pointer; update docs (`docs/05,07,01`, `docs/ru/*`) for test-tool routing and the `background` mis-attribution.
- **Non-scope:** any schema/enum change (U3); any new tool (U2); outputSchema slimming (U4); `save_as`/`job_list` (U5).
- **Changed files:** `src/excel_vba_mcp/server.py` (`_DESC_*` constants only); `src/excel_vba_mcp/server.py` (new `_DESC_ENVIRONMENT`, `_DESC_RES_GLOSSARY`, glossary resource); `docs/05-mcp-tools-reference.md`, `docs/07-workflows-and-examples.md`, `docs/01-overview.md`, `docs/ru/*`.
- **Acceptance criteria:** every workflow-critical `_DESC_*` follows §5.1 order; R-D3 guard sentence applied uniformly; no claim the code lacks (R-D1); `vba_compile` "diagnostic only"; `macro_classify_background` "never grants run permission"; `job_cancel` "intent-only"; test-tool routing present; glossary resource returns the §5.3 terms; docs no longer attribute `background` to `test_run_suite`.
- **Tests:** unit — description-contract lint (each critical tool contains Purpose/Use-when/Preconditions/Result/Next-actions markers; forbidden phrases absent: "compile" as a success claim for `vba_compile`, "grants run" for classify, "interrupt" for cancel); integration — `list_tools()` descriptions non-empty and contain the required markers; resource `xlvba://glossary` readable.
- **Compatibility impact:** none (descriptions + docs + one static resource). Tool count 91→91 (resource is not a tool).
- **Rollback:** revert `_DESC_*` constants and the glossary resource; no state.

### U3 — Schema/enum consistency + truthful annotations (schema + annotations)

- **Scope:** add the three request enums (§12.2) and retype the request fields; ship the 3 matching description edits; emit `ToolAnnotations` (§11) on all `add_tool` calls.
- **Non-scope:** description rewrites beyond the 3 enum tools (U1); new tools (U2); DTO slimming (U4).
- **Changed files:** `src/excel_vba_mcp/domain/statuses.py` (new request enums); `src/excel_vba_mcp/domain/models.py` (retype `spec.type`, `orientation`, `PivotFilterSpec.kind`); `src/excel_vba_mcp/server.py` (`_register_tools` add `annotations=`; 3 enum `_DESC_*`); `src/excel_vba_mcp/api/range_tools.py`/`pivot_tools.py` (only if a type import changes).
- **Acceptance criteria:** request schemas no longer list `color_scale/data_bar/icon_set/unsupported` (CF), `data` (pivot orientation), `label/value/date/top_bottom` (pivot filter); read/classify enums unchanged; annotations emitted for all tools and truthful per §11 (`macro_run`/`vba_compile` not read-only); pruned values still rejected at runtime (now as `INVALID_RANGE` validation, documented).
- **Tests:** unit — pydantic request models reject pruned values pre-COM; `list_tools()` shows no pruned enum values in request schemas and shows `annotations`; integration — sending a pruned value returns a validation error (not a domain `unsupported`), read tools still report `supported=false` for those types.
- **Compatibility impact:** **versioned request-contract narrowing** (validator Q14): clients sending pruned values move from domain rejection to pydantic `INVALID_RANGE`; both reject. Document in §17. Annotations additive; client display best-effort.
- **Rollback:** revert enum additions + retypes + annotations; no state.

### U2 — Environment/policy discovery + structured error envelope (runtime-API + new-capability)

- **Scope:** new `xlvba_environment` (§6.1); additive `SessionStatus.policy` (§6.2); additive `to_dict` envelope (§7); expose `has_snapshot` gate (§6.1). Non-COM, side-effect-free.
- **Non-scope:** descriptions (U1) except `_DESC_ENVIRONMENT`; enum/annotations (U3); slimming (U4); any SDK change; error `structuredContent`.
- **Changed files:** `src/excel_vba_mcp/api/environment_tools.py` (new); `src/excel_vba_mcp/domain/models.py` (`EnvironmentReport`, `SessionPolicy`, `SessionStatus.policy`); `src/excel_vba_mcp/application/session_service.py` (populate policy); `src/excel_vba_mcp/domain/errors.py` (envelope + pure functions); `src/excel_vba_mcp/server.py` (register `xlvba_environment`; `_DESC_ENVIRONMENT`).
- **Acceptance criteria:** `xlvba_environment` is `readOnlyHint:true`, non-COM (no worker start), `excel.available=false` when worker down, `build:null`, `target_build` labelled, `has_snapshot` = in-memory gate, `policy.classification_is_authorization=false`; `excel_get_status` returns the `policy` block; every error `to_dict` carries `errorCategory`/`isRetryable`/`suggestedNextActions` per §7.3 (`STALE_TARGET`/`REVISION_CONFLICT` false + caveat; `interrupted_unknown` false).
- **Tests:** unit — table-driven `category_for`/`retryable_for`/`actions_for` per code (incl. the D-U10 cases); `EnvironmentReport` rejects extra fields; `has_snapshot` reads `_latest_snapshot` (not store) — assert by creating a snapshot and checking true, and by a store-only entry and checking false. Integration — `xlvba_environment` over live server returns honest `build:null` and correct `access_mode`; a forced `SNAPSHOT_FAILED` (mutation without snapshot) yields the envelope with `isRetryable:false` + `snapshot_create` action; `excel_get_status.policy.has_snapshot` matches the gate.
- **Compatibility impact:** additive (new tool 91→92; new optional DTO fields; additive error keys). `>=87` assert still passes. Strict error consumers see new keys (tolerant clients unaffected).
- **Rollback:** remove the tool + DTOs + envelope; `to_dict` reverts to 5 keys; no state.

### U4 — Token reduction via outputSchema slimming (schema; D-U1)

- **Scope:** slim `WorkbookModel`, `SheetInfo`, and test-suite DSL output/input schemas (shared `$ref`, drop unused nested detail) while keeping `structuredContent`; document the rejected alternatives (search tool / `defer_loading` / `structured_output=False`).
- **Non-scope:** any new tool; any description change (U1); any enum change (U3); removing fields that carry `structuredContent` semantics.
- **Changed files:** `src/excel_vba_mcp/domain/models.py` (DTO field slimming / shared refs); possibly `src/excel_vba_mcp/domain/test_spec.py` (DSL dedupe); `docs/09-limitations-and-compatibility.md` (rationale).
- **Acceptance criteria:** `tools/list` outputSchema char share reduced vs the 451,108 / 75.05% baseline (measured by the Phase-7 harness `tc_tool_desc`); `structuredContent` still emitted on success for slimmed tools; no DTO field removed that a consumer relies on (inventory caveat §18); descriptions unchanged.
- **Tests:** unit — slimmed DTOs still serialize the fields the integration suite reads; integration — `list_tools()` outputSchema smaller; a success call still returns `structuredContent`; benchmark `tc_tool_desc` lower (Phase 7).
- **Compatibility impact:** additive-safe; **risk** if an external consumer parses a removed nested field (unknown inventory — §18). Prefer `$ref` dedupe over field removal.
- **Rollback:** revert DTO slimming; schemas return to baseline; no state.

### U5 — Deferred product decisions (investigation only; no implementation)

- **Scope:** produce a separate product contract outline for `save_as`/export and `xlvba_job_list` (§18); confirm the client-forwarding unknowns. **No production/test/config change.**
- **Non-scope:** implementing any tool; changing `excel_save`.
- **Changed files:** `.ai/audit/plans/` (new product-contract document only).
- **Acceptance criteria:** contract outline covers original protection, output path policy, overwrite, snapshot interaction, result retrieval for `save_as`; `job_list` read-only scope; unknowns list. No code.
- **Tests:** n/a (document).
- **Compatibility impact:** none.
- **Rollback:** delete the document.

**Implementation order:** U1 → U3 (parallel) → U2 → U4; U5 in parallel. See §13.

---

## 15. Tests and acceptance criteria

### 15.1 Unit tests (fast, no Excel)

- **Error envelope (U2):** table-driven test over every `ErrorCode` asserting `category_for`, `retryable_for`, `actions_for`; explicit D-U10 cases (`STALE_TARGET`/`REVISION_CONFLICT` → `isRetryable=false` + "may have partially applied" caveat; job `interrupted_unknown` → false); `to_dict` contains all 5 original keys plus the 3 new keys.
- **Request enums (U3):** `ConditionalFormatRequestType` rejects `color_scale/data_bar/icon_set/unsupported`; `PivotFieldLayoutOrientation` rejects `data`; `PivotFilterRequestKind` rejects `label/value/date/top_bottom`; read enums still accept them.
- **Environment DTO (U2):** `EnvironmentReport` with `extra="forbid"` rejects unknown keys; `build` is `None`; `has_snapshot` unit uses a fake store to prove the gate (`_latest_snapshot`) is read, not the store.
- **Annotations (U3):** `list_tools()` returns `readOnlyHint`/`destructiveHint`/`idempotentHint` per §11; `macro_run`/`vba_compile` are not read-only.
- **Description contract (U1):** lint each critical `_DESC_*` for the §5.1 markers and forbidden phrases.

### 15.2 Integration tests (live Excel; marker `integration_excel`)

- **Tool catalog:** `len(tool_names) >= 87` still holds (now 92 with `xlvba_environment`); do not assert "all-90".
- **Environment:** `xlvba_environment` returns `build:null`, correct `access_mode`, `excel.available` reflects worker; `excel_get_status.policy.has_snapshot` matches the in-memory gate (create snapshot → true; store-only entry → false).
- **Error envelope:** force `SNAPSHOT_FAILED` (guarded mutation without `snapshot_create`) → envelope with `errorCategory=missing_prerequisite`, `isRetryable=false`, `suggestedNextActions` includes `snapshot_create`; force `STALE_TARGET` (wrong `expected_source_hash`) → `isRetryable=false` + caveat, source unchanged.
- **Enum rejection:** send a pruned CF/pivot value → validation error (documented), not a silent accept; read tool still reports `supported=false`.
- **Annotations present** on a representative read + destructive + idempotent tool.
- **Originals-safe:** fixtures SHA-256 unchanged after each run (existing harness pattern).

### 15.3 Acceptance (per work package)

A work package is accepted only when **all** its §14 acceptance criteria pass and its §15.1/§15.2 tests are green. The remediation as a whole is accepted only when the Phase-7 benchmark (§16) shows no regression and the targeted improvements.

---

## 16. Agent Usability Benchmark specification

- **Reference:** `.ai/tmp/benchmark_tasks.yaml` (`BENCHMARK_TASKS`, 28 tasks) + `research_2b_usability_benchmark.md`. The YAML is authoritative for prompts, expected sequences, `expected_args`, `setup`, `access_mode`, `must_not`, and `failure_taxonomy`.
- **What it measures:** actual agent task execution against the live server (tool selection, first-call validity, completion, unnecessary calls, recovery correctness, unsafe attempts, token cost, completion time) — not LLM ratings of prose.
- **Metrics (per task, aggregated mean/median + 95% CI, N≥3 fresh-server runs):**
  - **CTS** correct-tool-selection fraction; **FCV** first-call validity; **WC** workflow completion (deterministic); **UC** unnecessary calls; **RC** recovery correctness (T23/T24); **UAA** unsafe action attempts (guardrail, must not increase); **TC** token consumption (`tc_tool_desc` + `tc_context`); **CT** completion time.
  - Recording: one JSONL line per task-run (§2.9 of research_2b).
- **Harness:** reuse `tests/integration_excel/test_mcp_full_access_workflow.py` primitives (`server_env`, `_running_server`, `_copy_fixture`, `structured`/`error_code`, `_wait_healthy`); wrap `ClientSession.call_tool` to capture the real wire trace; fixtures copied to a temp root, SHA-256 asserted unchanged.
- **Before/after protocol:** run the 28 tasks on **baseline** (`e2289ca`, pre-change) and on **final** (post-U1..U4), holding constant model, agent config, Kilo version, harness version, access mode per task, Excel build, fixtures, N runs, prompt text; the only independent variable is the change under evaluation.
- **Targets:** **set only after the baseline run** (no numeric target asserted before baseline exists). Acceptance = no WC/RC regression, UAA not higher (ideally 0), CTS/FCV improved, UC lower, and `tc_tool_desc`/`tc_task` lower for U4. Target deltas chosen from baseline variance (a target the baseline 95% CI does not already satisfy).
- **Tool-count note:** harness asserts `>=87` (or 91/92), never "all-90".

---

## 17. Compatibility and migration plan

- **Aliases:** none introduced; none required (no renames — D-U4/D-U5). Existing callers of `job_cancel`, `vba_compile`, `test_run_suite`, `test_suite_run` are unaffected.
- **Additive schema (`SCHEMA_VERSION=1`):** new `EnvironmentReport`/`SessionPolicy` DTOs and optional `SessionStatus.policy`; additive `to_dict` keys; new `xlvba_environment` tool (91→92). No existing field removed or retyped incompatibly (U3 retypes **request** enums only; read enums intact).
- **Enum prune = versioned request-contract narrowing (validator Q14):** clients sending `color_scale`/`data`/`label`/… move from domain rejection to pydantic `INVALID_RANGE`; both reject, but the **advertised** input contract narrows. Document in `docs/05` and `docs/09`; treat as a minor-version contract note.
- **Consumers:** tolerant JSON clients unaffected by additive error keys; strict clients must tolerate new keys (noted). `structuredContent` preserved (U4 keeps it); `structured_output=False` not used.
- **Tool-count assert:** keep `>=87` (`test_mcp_full_access_workflow.py`); do not claim "all-90". Docs updated to "91 exposed (92 after U2)".
- **Docs (2 languages):** `docs/05`, `docs/07`, `docs/01`, `docs/ru/*` updated for test-tool routing, pruned enums, `xlvba_environment`, and the `background` mis-attribution. `docs/09` gets the enum-narrowing note.
- **Rollout:** U1/U3 are doc+schema only (immediately reversible); U2 adds a read-only tool + additive fields; U4 is additive-safe with a consumer-inventory caveat. Each package independently revertible (§14 Rollback).

---

## 18. Deferred items and product decisions

These are **investigation + a separate product contract only** (D-U9). **No implementation in this remediation.**

### 18.1 `save_as` / export contract outline (U5)

A future `save_as(path)` tool must decide, before build:
- **Original protection:** writes a **new** file; never mutates the original; PathPolicy allowlist + Office-trust apply to the output path.
- **Snapshot interaction:** state explicitly whether the snapshot gate is required/skipped for a new-file write (it is not an in-place mutation) — do not silently require/skip.
- **Overwrite policy:** reject existing targets unless an explicit `overwrite` flag is passed (mirror `range` overwrite discipline).
- **Result retrieval:** return the new path + a `WorkbookInfo`-like DTO for the saved copy; define whether the working copy switches to the new path.
- **Export:** `vba_export_project` already covers VBA-only export; a fixed-format/PDF export is a separate decision.
- **Do not** change `excel_save` in-place semantics.

### 18.2 `xlvba_job_list` outline (U5)

- Read-only; lists **current-process** jobs only (job registry is in-memory, server-lifetime); no secrets/paths beyond what `job_get` returns; additive DTO; `SCHEMA_VERSION=1`.

### 18.3 Client-forwarding unknowns (must be confirmed against the target Kilo build)

- Whether Kilo forwards `outputSchema` to the model (determines whether schema-shape claims matter) — **unverifiable from inside the session**.
- Whether Kilo surfaces resource meta-tools (`list_mcp_resources`/`read_mcp_resource`) — docs say yes; **not observed live**; the compact glossary mirror in `xlvba_environment` is the reliable channel regardless.
- Whether any external consumer relies on `structuredContent` / specific nested DTO fields (no in-repo inventory) — U4 must prefer `$ref` dedupe over field removal until confirmed.
- Exact tokenization (char→token ratios are estimates, ~4:1 / 3.5:1 lower bound).
- A real 4-part Excel build source (only `Application.Version` = "16.0" is read) — `build` stays `null`.

---

*End of plan_03. Anchors @`e2289ca`. PLAN ONLY — authorizes no code change; execution is per work package §14.*