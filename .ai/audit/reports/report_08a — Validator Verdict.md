# Report 08a — Validator Verdict (Independent Validation)

- **Role:** Independent Validator (usability-remediation). VALIDATION ONLY — no production/test/config changes made.
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca01fd10136716a598cad1db6b6845fed73`.
- **Inputs reconciled:** `report_08 — Agent Usability Reconciliation.md`; `research_1a_tool_contracts.md`; `research_1b_runtime_policy.md`; `research_2a_mcp_discovery.md`; `research_2b_usability_benchmark.md`; `agent_usability_baseline.md`.
- **Method:** Every load-bearing claim was re-derived from source at HEAD (files cited below). Installed SDK (`mcp`/`mcp-types` 2.3.0) inspected for the error/structuredContent/outputSchema contract. No Excel process started; no files under `src/`, `tests/`, or config modified. Where the auditor/researcher and source disagree, source wins; disagreements are flagged.
- **Stance:** Adversarial. A proposal is approved only if its promised behavior is actually provided by the implementation and its schemas do not advertise values the implementation rejects.

---

## 1. Per-question verdict table (1–15)

Legend: **PASS** = agent can do it today from tool surface; **PARTIAL** = possible but with a real gap (undocumented / jargon / contradictory); **FAIL** = agent cannot do it today without reading source; "→ proposed" = verdict under the U1–U5 remediation if the stated conditions are met.

| # | Question | Verdict (today) | Under proposal | Evidence (file:line, HEAD) |
|---|----------|-----------------|----------------|-----------------------------|
| 1 | Determine current access mode before a mutation | **PARTIAL** | PASS | `SessionStatus.access_mode` is populated server-wide (`session_service.py:1104`, `models.py:111`) so it IS discoverable pre-mutation, but `_DESC_STATUS` omits it (`server.py:227`) and `server_health` does not report it (`server.py:1838`). `xlvba_environment`/policy block closes the discoverability gap. |
| 2 | Determine whether a snapshot is needed and who creates it | **FAIL** | PASS | Gate reads in-memory `_latest_snapshot` (`snapshot_service.py:400-425`); `has_snapshot` exists but is **not exposed** (`snapshot_service.py:427-430`); `snapshot_list` lists the on-disk **store**, not the gate (`snapshot_service.py:385-395`); mutation does **not** auto-snapshot. `_DESC_RANGE_WRITE_VALUES` says "requires … a successful snapshot" but never "call `snapshot_create` first" (`server.py:471-480`). |
| 3 | Determine whether a workbook is writable | **PASS** | PASS | `WorkbookInfo.read_only` (`models.py:91`) via `excel_get_status` / `excel_list_workbooks`; `_require_writable` rejects with `WORKBOOK_READ_ONLY`. |
| 4 | Identify actual Excel build and server build | **PARTIAL** | PARTIAL | Server package version via `server_health` → `"0.1.0"` (`server.py:1838`, `__init__.py:5`). Excel `Application.Version` = `"16.0"` (not 4-part) (`workbook_adapter.py:100-106`, `models.py:93`). **No tool returns the 4-part build (14332)**; it is a docstring constant only. Proposal honestly returns `build:null` + labelled `target_build` → still PARTIAL (real build unavailable). |
| 5 | Select correct tool among overlapping tools | **FAIL** | PASS | `test_run_suite`≈`test_suite_run` (`server.py:257-281`; `test_tools.py:226-234` — `test_suite_run` just calls `test_run_suite`); `server_health` terse, no redirect to `excel_health` (`server.py:229-233`); `xlvba_*` vs `excel-mcp_*` no routing. No selection rule anywhere. |
| 6 | Distinguish `vba_compile` diagnostics from verified compilation | **PARTIAL** | PASS | Description is honest: "never returns compiled or compile_error", `status='indeterminate'`, `is_compile=false` (`server.py:320-327`); code always `INDETERMINATE` (`compile_service.py:334-346`). Gap is the **name** misleads, not the prose. |
| 7 | Understand background classification ≠ execution authorization | **PARTIAL** | PASS | Description already states "UNGATED bookkeeping (macro_run remains the security boundary)" (`server.py:351-358`), but "UNGATED"/"security boundary" are undefined jargon (A21). Needs plainer "classification never grants run permission". |
| 8 | Understand cancellation intent vs actual COM interruption | **PASS** | PASS | `job_cancel` description accurate: "intent-only; never interrupts the running COM call … runs to completion" (`server.py:367-371`; `job_registry.py:328-353`). Only the name `cancel` over-promises; rename optional. |
| 9 | Understand `SUSPECTED_STALLED` ≠ confirmed hung | **PARTIAL** | PARTIAL | `SUSPECTED_STALLED="suspected_stalled"` exists with an accurate code comment (`statuses.py:102-105`) and is returned in results, but **no tool description explains it** vs confirmed hung; `job_get`/`macro_get_last_result` list statuses without semantics (`server.py:359-366, 347-350`). Glossary (U1) fixes; no schema change needed. |
| 10 | Identify correct action after every major error category | **FAIL** | PASS | `to_dict()` = `{code,message,hresult,hresult_hex,context}` only (`errors.py:138-147`); **no** `errorCategory`/`isRetryable`/`suggestedNextActions` (grep: none). Only job-level `retryable` exists. Additive envelope (U2) closes it. |
| 11 | Complete a workflow without reading server source | **FAIL** | PASS | Mutation workflows fail on: snapshot precondition (Q2), `test_*.workbook_id`-must-be-open + auto-baseline (not in description, `test_runner.py:378`), VBA `expected_source_hash` source (not stated, `server.py:308-319`), enum rejections (Q15). Note: `inputSchema` **is** always forwarded to the model, so payload shapes (`{rows,cols,values}`) are available — the failures are precondition/contradiction, not shape. |
| 12 | Does deferred loading reduce tool-context tokens in Kilo? | **FAIL (N/A)** | FAIL | `defer_loading`/`deferLoading`/`search_tools` = **0 hits** in `mcp`+`mcp_types` 2.3.0; Kilo has no deferral (2A §3–§4). A standalone search tool adds a 92nd always-loaded definition → **net increase**. Only a full `xlvba_call` dispatcher reduces context (out of scope; erodes per-tool permissions/validation). Real lever is `outputSchema` (see §5). |
| 13 | Can existing agents continue using deprecated aliases? | **N/A (none exist)** | PASS (conditional) | No deprecated aliases exist today. If U3 keeps `test_run_suite` canonical **and** retains `test_suite_run` as a thin alias (or vice-versa), existing callers survive. If either is hard-removed, callers break (docs in 2 languages + task_053 traceability). |
| 14 | Can schema changes be introduced without breaking consumers? | **PARTIAL** | PARTIAL | Additive optional fields on model classes are safe under `extra="forbid"` (`models.py:62,68`) **only if added to the class**, not ad-hoc. **Enum pruning is technically a schema-contract break**: a client sending `color_scale`/`data`/`label` would move from domain rejection to pydantic `INVALID_RANGE` (`server.py:1799-1804`). Both reject, but the advertised contract changes — must be versioned/documented. |
| 15 | Are descriptions/examples consistent with live schemas? | **FAIL** | PASS | Schema advertises values the implementation rejects: `ConditionalFormatType` includes `color_scale/data_bar/icon_set/unsupported` (`statuses.py:437-440`) but adapter rejects non-plain types pre-COM; `PivotFieldOrientation` includes `data` (`statuses.py:470`) rejected; `PivotFilterKind` includes `label/value/date/top_bottom` (`statuses.py:520-524`) never applied. Descriptions contradict the enums. Pruning request enums (U3) fixes. |

**Scoreboard:** PASS 2 (Q3, Q8) · PARTIAL 5 (Q1, Q4, Q6, Q7, Q9, Q14) · FAIL 6 (Q2, Q5, Q10, Q11, Q12, Q15) · N/A 1 (Q13). The remediation is justified: 6 hard failures + 5 partials, all source-confirmed.

---

## 2. Unsafe / misleading description or schema violations

### 2.1 Existing violations the plan MUST fix (schema advertises rejected values)
These are **current** contradictions. A description-only remediation that does NOT prune the request enums leaves the violation in place — that is a blocking defect, not a cosmetic one.

1. **`range_set_conditional_formatting`** — request enum `ConditionalFormatType` advertises `color_scale`/`data_bar`/`icon_set`/`unsupported` (`statuses.py:437-440`, surfaced via `models.py` `spec.type`), but `validate_conditional_format_spec` rejects any type ∉ `_SUPPORTED_CF_TYPES` **pre-COM**. Schema promises values the implementation rejects. **Must prune the request enum** (or the schema remains a lie).
2. **`pivot_field_set_layout`** — `orientation` enum advertises `data` (`statuses.py:470`); `_require_layout_orientation` rejects `data` pre-STA. **Must prune `data` from the settable enum.**
3. **`pivot_field_apply_filter`** — `PivotFilterSpec.kind` advertises `label/value/date/top_bottom` (`statuses.py:520-524`); only `manual` is applied, others return `supported=false`. **Must prune or the schema over-promises.**

### 2.2 Unsafe proposals to block or caveat (do NOT let the plan ship these as written)
1. **Error `structuredContent` on failure is NOT feasible without an SDK change.** The SDK converts a raised `ToolError` into `CallToolResult(content=[TextContent(text=str(exc))], is_error=True)` — **text only, no structuredContent** (`.venv/.../mcp/server/mcpserver/server.py:456`). Success carries `structured_content` (`func_metadata.py:226`). The server re-raises `ToolError(json.dumps(cause.to_dict()))` (`server.py:1808-1826`). **Any proposal that says errors will also emit `structuredContent` is misleading and must be dropped or explicitly deferred.** The additive `to_dict()` envelope (text JSON) IS feasible with no SDK change.
2. **`isRetryable=true` for `STALE_TARGET`/`REVISION_CONFLICT` is unsafe as a blanket value.** These are optimistic-concurrency failures on a mutation that may have **partially applied** before the hash check. A naive `isRetryable:true` + "retry with fresh hash" encourages blind replay. Must be `isRetryable:false` **or** carry an explicit "re-read state; do NOT assume the prior mutation did not apply" caveat. (Researcher 1B §6 Q5 flags this; the plan must resolve it, not inherit it.)
3. **`xlvba_environment` must be described as read-only/side-effect-free.** It must NOT be described (or implemented) as creating snapshots, grants, or classifications, and must not imply it authorizes mutations. `classification_is_authorization=false` must be stated. Any wording like "environment check enables X" is a violation.
4. **`build` must be `null`, never fabricated `14332`.** Describing `target_build` as "the Excel build" is misleading; it is a design-time constant (`compile_service.py:3`, `statuses.py:476`). Only `Application.Version` ("16.0") is real (`workbook_adapter.py:100-106`).
5. **`has_snapshot` must report the in-memory gate (`_latest_snapshot`), not the store.** Reporting the store (`snapshot_service.py:385-395`) would mispredict `require_snapshot` (`:400-425`) and is a correctness bug, not a doc nuance.
6. **`structured_output=False` removes `structuredContent` too.** `convert_result` with `output_model is None` returns `CallToolResult(content=unstructured_content)` — **no structured_content** (`func_metadata.py:211-212`). Any description/claim that "structured output is still available" after `structured_output=False` is false. See §5.
7. **`vba_compile` must never be described as compiling.** Description is honest today (`server.py:320-327`); a rewrite that softens "does not compile" would be a regression. Rename is optional; if renamed, an alias must remain (Q13).
8. **`job_cancel` rename must not imply interruption.** If renamed to `job_request_cancel`, the description must still say "intent-only; never interrupts" (`server.py:367-371`). A name like `job_abort`/`job_kill` would be a violation.
9. **`macro_classify_background` must not be described as granting execution.** "UNGATED bookkeeping" is correct (`server.py:351-358`); any "classify to enable run" phrasing is a violation (the BG06 gate + approval + R13 still run, `macro_service.py:477-580`).

---

## 3. Approve / Reject per remediation group (U1–U5)

### U1 — Descriptions + glossary — **APPROVE**
- **Rationale:** Highest value / lowest risk. Descriptions are only ~5.4% of `tools/list` (2A §5) but carry the precondition/neighbor-routing load that drives Q2/Q5/Q10/Q11. Source confirms the gaps are real (snapshot precondition, caps, hash source, gate list, enum rejections).
- **Conditions:**
  1. Must add, verbatim where specified: `range_write_values` "does NOT snapshot — call `snapshot_create` first" + ceilings (`rows≤100000`, `cols≤512`, cells≤100000) (`models.py:666-672`); `range_read_values` caps (100k request / 50k·8MiB response); `vba_replace_component_source` hash source (`vba_read_component`/`vba_list_components`); `macro_run` **numbered** gate list + define "approved"; table `where`/`updates` grammar + ambiguity definition; `name_*` `=`-optional; `sheet_add` `sheet_type`; `test_*` workbook-open + auto-baseline precondition; `excel_open` defaults (`mode=execute`, `read_only=False`).
  2. Glossary MUST define: R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, `SUSPECTED_STALLED` vs `interrupted_unknown` vs `indeterminate`.
  3. Must NOT introduce any claim the implementation does not provide (see §2.2). Descriptions must stay consistent with the pruned schemas from U3.
- **Blocking concerns:** none. This group is safe to start immediately.

### U2 — Runtime discovery + structured errors — **APPROVE, with one hard carve-out**
- **Rationale:** Closes Q1/Q2/Q4/Q9/Q10. `xlvba_environment` (non-COM, like `server_health` `server.py:1834-1838`) + additive `to_dict()` envelope are both feasible on the installed SDK.
- **Conditions:**
  1. `xlvba_environment` MUST be read-only, non-COM, side-effect-free; `excel.available=false` when worker down; `build:null` honest; `has_snapshot` = in-memory gate; `policy.classification_is_authorization=false`.
  2. Additive error envelope (`errorCategory`, `isRetryable`, `suggestedNextActions`) is **feasible as text-content JSON with NO SDK change** (`to_dict()` extension is additive; `server.py:1826` already JSON-encodes it). **APPROVED on that basis.**
  3. **CARVE-OUT / REJECT the sub-proposal to emit error `structuredContent`:** not possible without an SDK change (`server.py:456` text-only). Drop it or mark explicitly out-of-scope. Do not describe errors as "structured content".
  4. Resolve `isRetryable` for `STALE_TARGET`/`REVISION_CONFLICT` per §2.2(2) before implementation.
  5. Respect `extra="forbid"` (`models.py:68`): new fields go on new/optional DTO classes (`EnvironmentReport`), not ad-hoc payloads; keep `SCHEMA_VERSION=1`.
- **Blocking concerns:** the error-`structuredContent` carve-out (must be removed from scope); `isRetryable` semantics for concurrency errors.

### U3 — Schema/tool-surface consistency + annotations + cancellation rename — **APPROVE (schema+annotations); REJECT the rename as a blocker unless aliased**
- **Rationale:** Pruning request enums is the **correct** fix for Q15/§2.1 (description-only leaves the schema lying). `add_tool(..., annotations=ToolAnnotations|None, structured_output=...)` is supported (`server.py:618-627` SDK) and currently unused (0/91) — emitting `readOnlyHint`/`destructiveHint`/`idempotentHint` from existing access-class metadata is feasible and truthful.
- **Conditions:**
  1. Prune **request** enums (CF drop `color_scale/data_bar/icon_set/unsupported`; pivot orientation drop `data`; pivot filter drop `label/value/date/top_bottom`). Keep read/classify enums intact for read tools. Treat as a **versioned schema-contract change** (Q14) and document it; both old/new reject the pruned values, but the advertised input contract narrows.
  2. Annotations MUST be derived from real access class (READ_ONLY→`readOnlyHint:true`; destructive ops→`destructiveHint:true`; idempotent where re-read verifies). Do **not** annotate `macro_run`/`vba_compile` as read-only.
  3. **Cancellation rename: REJECT as a hard rename.** The description is already accurate (Q8 PASS). A rename is high-compat (docs `05-mcp-tools-reference.md`, tests) and the benefit is cosmetic. **Approve only if a backward-compatible alias is retained** and the description still says "intent-only; never interrupts". Otherwise drop the rename.
- **Blocking concerns:** rename compat; enum-prune must be versioned, not silent.

### U4 — Progressive discovery — **REJECT (as a token-reduction lever in Kilo)**
- **Rationale:** Source-confirmed dead end for this client. `defer_loading` absent from `mcp`/`mcp_types` 2.3.0 (0 hits); Kilo loads all tools at once (2A §3–§4). A standalone `xlvba_search_tools` **adds** a 92nd always-loaded definition and removes nothing → **net context increase** (Q12 FAIL). Resources are the only no-dispatcher mechanism but the resource meta-tools were **not observed** in the live session (2A §4.4/§7).
- **Conditions to reconsider (all required):** (a) a **full** `xlvba_call(tool_name, arguments)` dispatcher that collapses 90→2 definitions — with explicit sign-off on the loss of per-tool `{server}_{tool}` permissions and client-side input-schema validation; (b) confirmed Kilo build that surfaces `list_mcp_resources`/`read_mcp_resource`; (c) benchmark evidence (2B) that `tc_tool_desc` actually drops. None of these hold today.
- **Blocking concerns:** shipping U4 as "progressive discovery" would **increase** context and is therefore worse than the baseline. The real token lever is outputSchema reduction (§5), which belongs to U3/§5, not U4.

### U5 — Product capability gaps (`save_as` / `job_list`) — **APPROVE (both), with conditions**
- **Rationale:** `save_as`/`save_copy_as`/`export_fixed_format` **do not exist** as tools (grep: only internal `save_copy_as_fn` in snapshot store). This is a real gap (Q11-adjacent, T28). `job_list` is a read-only additive gap (job registry is in-memory/server-lifetime, `job_registry.py`).
- **Conditions:**
  1. `save_as(path)` MUST be a **new** tool; do not change `excel_save` in-place semantics (`workbook_tools.py:147-159`). Must enforce PathPolicy allowlist + Office-trust; decide snapshot interaction explicitly (it writes a new file, not in-place mutation — do not silently require/skip the snapshot gate without stating which).
  2. `job_list` MUST be read-only, list only current-process jobs, and not expose secrets/paths beyond what `job_get` already returns.
  3. Both are additive; keep `SCHEMA_VERSION=1` (new DTOs).
- **Blocking concerns:** `save_as` path-policy/snapshot semantics must be decided before build; otherwise it is an unreviewed write primitive.

---

## 4. Prioritized blocking concerns the Planner MUST address

1. **[BLOCKER] Error `structuredContent` is not feasible without an SDK change.** Remove it from U2 scope; the text-content JSON envelope is the only no-SDK option (`server.py:456`, `server.py:1826`). Any plan text promising structured error content is wrong.
2. **[BLOCKER] U4 as proposed increases context.** Do not ship a standalone search tool or "progressive discovery" that adds a 92nd definition (Q12). Either drop U4 or convert to a full dispatcher with explicit permission/validation sign-off.
3. **[BLOCKER] Enum pruning must accompany (not replace) description fixes.** Description-only leaves the schema advertising rejected values (§2.1, Q15). Prune request enums and version the contract change (Q14).
4. **[BLOCKER] `isRetryable` for `STALE_TARGET`/`REVISION_CONFLICT` must not encourage blind replay** of a partially-applied mutation (§2.2(2)). Resolve to `false` or add an explicit "re-read; prior mutation may have applied" caveat.
5. **[BLOCKER] `has_snapshot` must read the in-memory gate, not the store** (§2.2(5)); otherwise the discovery tool mispredicts `require_snapshot` and is a correctness bug.
6. **[MAJOR] `structured_output=False` also removes `structuredContent`** (`func_metadata.py:211-212`). Any token-reduction plan must state this and confirm no consumer relies on the structured channel (§5). Prefer slimming oversized DTOs over blanket `structured_output=False`.
7. **[MAJOR] Cancellation rename is high-compat / low-value.** Retain an alias or drop it; never ship a name that implies interruption (§2.2(8), Q8).
8. **[MAJOR] `xlvba_environment` must be side-effect-free and honest about `build:null`** (§2.2(3)(4)); `classification_is_authorization=false` must be explicit.
9. **[MAJOR] `save_as` path-policy/snapshot semantics undecided** (U5). Decide before build.
10. **[MINOR] Tool-count reconciliation:** exposed catalog is **91** (90 `add_tool` + `server_health`), not 90. Harness currently asserts `>=87` (`test_mcp_full_access_workflow.py:499`) — keep `>=87` or update to 91, but do not claim "all-90". Docs (2 languages) reference both test tools; collapsing has doc cost in `docs/01,05,06,07` + `docs/ru/*`.

---

## 5. Token-reduction claim — verification

- **Is cutting `outputSchema` the real lever? YES.** Measured `tools/list` = 451,108 chars; `outputSchema` = 338,547 (**75.05%**), inputSchema 80,656 (17.9%), descriptions 24,272 (5.4%) (auditor §1.2, reproduced; 2A §5). Descriptions are the **smallest** lever; any plan that frames prose-shortening as the token fix is wrong.
- **Is `structured_output=False` effective? YES, mechanically.** `func_metadata.py:364-365` returns metadata with no output schema; `Tool.output_schema` (`base.py:61-62`) then yields `None`. This is a server-side change needing no Kilo feature.
- **Risk (must be stated):** `structured_output=False` also drops `structuredContent` on success (`func_metadata.py:211-212` returns content-only). Clients that parse the typed DTO from `structuredContent` (rather than text) would break. Kilo's `structuredContent` handling is **unknown/unverified** (2A §4.2/§7). **Safer alternative:** slim the oversized return DTOs (`WorkbookModel` 28.8K, `SheetInfo` ~10.9K×7, test-suite DSL 22K input) — this cuts the 75% driver **while preserving** `structuredContent`. Recommend DTO-slimming first; reserve `structured_output=False` for tools whose structured channel is genuinely redundant with text, and only after confirming no consumer depends on it.

---

## 6. What cannot be validated (and why)

1. **Whether Kilo forwards `outputSchema` to the model / renders `structuredContent`.** Kilo does not document internal tool serialization; global `~/.config/kilo/kilo.jsonc` is read-blocked (`external_directory: deny`). This is the load-bearing unknown for the "shape undefined" family — the shapes ARE in `tools/list` outputSchema, but if Kilo drops outputSchema, only prose is visible. **Unverifiable from inside the session.**
2. **Whether Kilo surfaces resource meta-tools** (`list_mcp_resources`/`read_mcp_resource`). Docs say yes; **not observed** in the live session (2A §4.4). Resource-based progressive disclosure is therefore doc-supported but unconfirmed.
3. **Exact tokenization.** Char→token ratios (~4:1, 3.5:1 lower bound) are estimates, not a tokenizer measurement.
4. **Whether any external consumer relies on `structuredContent`.** No consumer inventory is available in-repo; the risk in §5 is asserted as a risk, not quantified.
5. **A real 4-part Excel build source.** Only `Application.Version` ("16.0") is read (`workbook_adapter.py:100-106`); no confirmed read-only source for 14332. `build` must stay `null`.
6. **`macro_policy.classify_operation` non-authoritative status** (1B §1.2) — plausible from source, not independently re-derived here; low risk, flagged for the Planner to confirm if U2 touches policy reporting.

---

## 7. Validator determination

The remediation is **substantially justified** (6 FAIL + 5 PARTIAL, all source-confirmed) but **not approvable as a blanket package**. Approve **U1** and **U5** (with conditions); approve **U2** only after removing the error-`structuredContent` sub-proposal and fixing `isRetryable`; approve **U3** schema+annotations but **reject the cancellation rename** unless aliased; **reject U4** as a token lever. The Planner must resolve the 10 blocking concerns in §4 — especially the SDK error constraint (#1), the U4 context increase (#2), enum-prune-vs-description (#3), and the `structuredContent` risk of `structured_output=False` (#6) — before any implementation is considered execution-ready.