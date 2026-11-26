# Report 08 — Agent Usability Reconciliation (Independent Auditor)

- **Role:** Independent Auditor (Phase 2 source-code audit). AUDIT ONLY — no production/test/config changes made.
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca01fd10136716a598cad1db6b6845fed73` (verified via `git rev-parse HEAD`).
- **Inputs reconciled:**
  - Original audit: `.ai/tmp/xlvba-mcp-usability-audit-2026-11-26.md`
  - `research_1a_tool_contracts.md`, `research_1b_runtime_policy.md`, `research_2a_mcp_discovery.md`, `research_2b_usability_benchmark.md`
  - Baseline: `agent_usability_baseline.md`
- **Method:** Every Blocker/Major (and the cross-cutting) claim was re-derived from source at HEAD. Where the audit/researcher and source disagree, source wins. Token figures were independently reproduced by re-running the read-only measurement (`.ai/tmp/measure_tools_2a.py`, imports the server, does not start Excel). No Excel process was started; no files under `src/`, `tests/`, or config were modified.

---

## 1. Methodology & verification notes

1. **Registration truth:** `server.py` has exactly **90** `server.add_tool(...)` calls (`server.py:1479–1704`) plus **1** `@mcp.tool()` `server_health` (`server.py:1834–1838`) = **91 tools exposed**, and **4** `@mcp.resource(...)` templates (`server.py:1869, 1883, 1913, 1930`). Confirmed by `Select-String` counts and by live `list_tools()`.
2. **Token measurement (independent):** re-ran `.ai/tmp/measure_tools_2a.py`. Output reproduced researcher 2A **exactly**: `TOOL_COUNT=91`, `TOOLS_JSON_CHARS=451108`, `TOTAL_DESC_CHARS=24272`, `TOTAL_INPUT_SCHEMA_CHARS=80656`, `TOTAL_OUTPUT_SCHEMA_CHARS=338547`, `TOOLS_WITH_OUTPUT_SCHEMA=91`, `TOOLS_WITH_ANNOTATIONS=0`, `TOOLS_WITH_TITLE=0`.
3. **SDK truth:** `mcp` 2.3.0, `mcp-types` 2.3.0, `pydantic` 2.14.4, Python 3.14.7 (verified via `importlib.metadata`). `defer_loading`/`deferLoading`/`search_tools` = **0 hits** across `mcp`+`mcp_types` (recursive grep).
4. **Error/structuredContent truth:** SDK `MCPServer.call_tool` converts a raised `ToolError` into `CallToolResult(content=[TextContent(text=str(exc))], is_error=True)` (`.venv/.../mcp/server/mcpserver/server.py:456`) — **text content only, no `structuredContent`**. Success results (typed DTO returns) go through `FuncMetadata.convert_result` → `CallToolResult(content=..., structured_content=...)` (`func_metadata.py:186–226`). The server's `ExcelVbaMcpServer.call_tool` re-raises `ToolError(json.dumps(cause.to_dict()))` (`server.py:1808–1826`), so **every tool error is `isError=true` with a JSON string in text content, never structuredContent**.
5. **Output schema is auto-derived and rich:** every tool emits `outputSchema` (top-level `object`); for DTO returns it embeds full `$defs` (e.g. `excel_open` outputSchema contains `$defs.WorkbookInfo` and `$defs.RunningJob`; `range_read_values` embeds `$defs.CellValue`; `vba_compile` embeds the compile DTO). This is decisive for several "undefined shape" claims (see §3).

---

## 2. Findings table

Legend — **Classification:** CONFIRMED / PARTIALLY_CONFIRMED / REFUTED / ALREADY_FIXED / UNVERIFIABLE. Severity in brackets is the *audit's* severity; the "Discrepancy" column states the source-verified reality.

| id | source | tool / symbol | file:line (HEAD) | current behavior (source) | discrepancy vs claim | classification | recommended change (advisory) | compatibility risk |
|---|---|---|---|---|---|---|---|---|
| A01 | audit Blocker | `excel_open` | desc `server.py:216–222`; handler `api/workbook_tools.py:88–127`; DTO `domain/models.py:83–96`; outputSchema embeds `$defs.WorkbookInfo` | defaults `mode=EXECUTE`, `read_only=False`, `background=False`; returns `WorkbookInfo \| RunningJob` | "WorkbookInfo shape undefined" is wrong — shape is in `models.py` **and** in the auto `outputSchema`. Real gap: prose omits defaults + writability truth table | PARTIALLY_CONFIRMED (shape part REFUTED) | add one line: defaults + "writable unless read_only=true"; no schema change | low (description-only) |
| A02 | audit Blocker | `range_write_values` | desc `server.py:471–480`; handler `api/range_tools.py:304–328`; payload `models.py:1000–1030`; gate `snapshot_service.py:400–425` | requires a **pre-existing** snapshot (`require_snapshot` reads in-memory `_latest_snapshot`; mutation does **not** auto-snapshot); payload `{rows,cols,values}`, `address`=top-left anchor | description says "requires … a successful snapshot" but never says the agent must call `snapshot_create` first; payload shape only in schema | CONFIRMED | "This tool does NOT snapshot — call `snapshot_create` first"; restate `{rows,cols,values}` + ceilings | low |
| A03 | audit Blocker | `vba_compile` | desc `server.py:320–327`; `compile_service.py:156–217, 334–346`; handler `compile_tools.py:63–81` | **always** `status=INDETERMINATE`, reason `NON_ACTIVATING_COMPILE_UNAVAILABLE`; never `compiled`/`compile_error`; description already says so | description is **honest**; only the **name** misleads. "Blocker" severity overstated | CONFIRMED (naming/affordance, not doc contradiction) | sharpen first line to "diagnostic only — does not compile"; rename is optional | rename = HIGH (tests/docs use `vba_compile`); desc = low |
| A04 | audit Blocker | `test_run_suite` / `test_suite_run` | desc `server.py:257–266, 271–281`; `api/test_tools.py:163–186, 189–234` | `test_suite_run` accepts `TestSuiteSpec \| dict` + `background`, parses dict, then **calls `test_run_suite(spec)`** (`:234`); `test_run_suite` is typed-spec, blocking only | genuine near-duplicate; no selection rule; **docs wrongly claim `test_run_suite` accepts `background`** (`docs/05-mcp-tools-reference.md:173,837`) | CONFIRMED (+ doc-vs-code deviation) | keep one canonical; deprecate the other; fix docs | HIGH — tests (`test_mcp_error_surfacing_matrix.py:357,416,438,501`, `test_macro_policy.py:94`) + docs (`01-overview.md`, `05-mcp-tools-reference.md`, `07-workflows-and-examples.md`) reference both; tool-count assert `>=87` |
| A05 | audit Blocker / 1A Major | `range_set_conditional_formatting` | desc `server.py:560–567`; enum `statuses.py:418–440`; spec `models.py:969–980`; `_SUPPORTED_CF_TYPES` `ranges_adapter.py:439+`; reject `ranges_adapter.py:2247–2252` | request enum `ConditionalFormatType` includes `color_scale/data_bar/icon_set/unsupported`, but `validate_conditional_format_spec` rejects any type ∉ `_SUPPORTED_CF_TYPES` (plain `FormatCondition` set) **pre-COM** | enum and description disagree; agent sending `color_scale` is rejected | CONFIRMED | split request vs read enum (prune request enum) **or** state rejected values in prose | MEDIUM (schema change; check tests) |
| A06 | audit Major / 1A | `pivot_field_set_layout` | desc `server.py:675–687`; enum `statuses.py:463–470`; spec `models.py:1186–1190`; `_LAYOUT_ORIENTATIONS` `pivot_tables_service.py:114`; reject `:944–964` | `orientation='data'` rejected pre-STA; `hidden`+`delete=false`=non-destructive `Orientation=0`, `hidden`+`delete=true`=destructive `PivotField.Delete` | enum advertises `data` (rejected); `hidden` overloaded — **but the description documents both** | CONFIRMED (overload documented) | prune `data` from settable enum; keep prose | MEDIUM |
| A07 | audit Major / 1A | `pivot_field_apply_filter` | desc `server.py:727–738`; enum `statuses.py:510–524`; spec `models.py:1201–1211`; handler `pivot_tools.py:478–505` | only `manual` (`PivotItems[i].Visible`) applied; `label/value/date/top_bottom` return `filter.supported=false` + reason, never converted | enum advertises kinds never applied; description states manual-only | CONFIRMED | prune request enum or mark unsupported in prose | MEDIUM |
| A08 | audit Major / 1A Minor | `job_cancel` | desc `server.py:367–371`; `job_registry.py:328–353` | sets `cancel_requested=True` only if non-terminal; the STA COM call runs to completion | description is **accurate** ("intent-only; never interrupts"); "Major" overstated | CONFIRMED (accurate; severity overstated) | optional: "observe via `job_get`; use `excel_recover` to actually stop" | low |
| A09 | audit Major | `snapshot_restore` / `excel_close` / `excel_recover` | desc `server.py:251–255, 223–225, 234–245`; `recovery_tools.py:46–68`; `snapshot_service.restore` | close default `save_changes=False`; restore closes working copy **without saving**; recover discards working copies (originals safe) | "silent loss" is overstated — all three descriptions **do** disclose no-save/discard; no `confirm` flag exists | PARTIALLY_CONFIRMED | bold "unsaved edits lost"; optional `confirm` flag | low–medium (flag = schema) |
| A10 | audit Major | `macro_run` | desc `server.py:336–346`; `macro_service.py:383–618` | gate order: execute-mode+writable (`:423`) → resolve (`:426–473`) → deny-by-default approval `check(EXECUTING)` (`:484`) → BG06 (`:517/523`) → R13 preflight (`:535/576`) → single `Application.Run` (`:582`) | gates described but **not numbered**; "approved" (allowlist/grant) never defined | PARTIALLY_CONFIRMED | numbered checklist + define "approved" | low |
| A11 | audit Major | `macro_classify_background` | desc `server.py:351–358`; `macro_service.py:721–826` | UNGATED bookkeeping; records review outcome; **does not authorize execution**; `run_macro` still enforces approval+R13 | description **already says** "UNGATED bookkeeping (macro_run remains the security boundary)" | PARTIALLY_CONFIRMED → leaning REFUTED | optional: add "classification never grants run permission" | low |
| A12 | audit Major / 1A Minor | `table_update_rows` / `table_delete_rows` | desc `server.py:428–437`; `table_tools.py:220–242`; `tables_adapter.py:892–1011` | `updates` = `[{index\|where, values}]` (exactly one selector); `where` must match exactly one row else `AMBIGUOUS_FILTER`; grammar enforced in **adapter**, not tool layer | grammar not in description; "ambiguous" undefined (= matched ≠ 1 row) | CONFIRMED (minor) | publish grammar + example + define ambiguity | low |
| A13 | audit Major / 1A Minor | `name_create` / `name_update` | desc `server.py:578–587`; `name_tools.py:146–206`; `names_adapter.py:248–258` | `refers_to` accepted **with or without** leading `=` (normalized); duplicate/cross-scope conflicts rejected | description examples imply `=` required; "cross-scope conflict" undefined | CONFIRMED (minor) | state `=` optional; define conflict | low |
| A14 | audit Major / 1A Minor | `range_read_values` | desc `server.py:459–464`; `range_tools.py:174–203`; caps `models.py:666–672` | `offset` 0-based (stated); response capped at `MAX_RESPONSE_CELLS=50_000` / `MAX_RESPONSE_BYTES=8 MiB`; request rejected > `MAX_RANGE_CELLS=100_000` | "offset base unclear" is **REFUTED** (0-based stated); caps **not stated** | PARTIALLY_CONFIRMED (offset REFUTED; caps CONFIRMED) | publish the three numeric caps | low |
| A15 | audit Major / 1A Minor | `vba_replace_component_source` / `vba_insert_code` | desc `server.py:308–319`; `vba_tools.py:191–264` | requires `expected_source_hash` = SHA-256 of current normalized body (stale→`REVISION_CONFLICT`/`STALE_TARGET`) | 2-step (read hash first) not stated; hash source not named | CONFIRMED (minor) | "obtain `expected_source_hash` from `vba_read_component`/`vba_list_components`" | low |
| A16 | audit Minor / 1A | `sheet_add` / `sheet_copy` | desc `server.py:378–392`; `sheets_adapter.py:267–345` | `sheet_add` default `sheet_type="worksheet"`, others rejected; `sheet_copy` name Excel-generated (`<name> (2)`) | defaults/behavior not in description | CONFIRMED (minor) | enumerate `sheet_type`; note generated copy name | low |
| A17 | audit Minor | `server_health` vs `excel_health` | `server.py:229–233, 1834–1838`; `health_tools.py:103–167` | `server_health`→`{status,service,version}` (no COM); `excel_health`→`ExcelHealth{status,session_state,worker_state,accepting,needs_recovery,owned_pid,…}` | tools **are** distinct; `server_health` description is terse and doesn't redirect to `excel_health` | PARTIALLY_CONFIRMED | "liveness only; use `excel_health` for worker/session state" | low |
| A18 | audit systemic #1 | mode/policy discoverability | `session_service.py:1104`; `models.py:111`; desc `server.py:227`; `server.py:1834–1838` | `excel_get_status`→`SessionStatus.access_mode` **is populated**; `server_health` does **not** report mode; `_DESC_STATUS` omits it | "no tool reports the current mode" is **REFUTED** (`excel_get_status` does); "not in server_health / not in description" CONFIRMED | PARTIALLY_CONFIRMED (core claim REFUTED) | add mode/build/snapshot to `server_health` or a new `xlvba_environment`; mention in `_DESC_STATUS` | low–medium |
| A19 | audit systemic #2 | snapshot precondition | `snapshot_service.py:400–430, 385–395, 176, 348` | `require_snapshot` reads in-memory `_latest_snapshot`; populated only by `create_snapshot`/`restore_snapshot`; `has_snapshot` (`:427–430`) **not exposed**; `snapshot_list` lists the **store**, not the gate | agent cannot predict whether the gate passes; mutation does not auto-snapshot | CONFIRMED | expose `has_snapshot` (or report in status/environment); document "call `snapshot_create` first" | low–medium |
| A20 | audit systemic #3 | build/version opacity | `server.py:1838`; `workbook_adapter.py:99–116`; `models.py:93`; `compile_service.py` (docstring only) | `server_health`→package `__version__="0.1.0"`; `WorkbookInfo.excel_version`=`Application.Version` ("16.0", not 4-part); "target build 16.0.14332" is a **docstring constant only**, no runtime gate, no tool returns it | no tool returns the 4-part build or a capability flag | CONFIRMED | disclose target build once (environment/health); return `build:null` honestly if unavailable | low |
| A21 | audit systemic #4 | undefined jargon | `macro_service.py:34–50, 126–150`; `statuses.py`; descriptions | R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash used with no glossary | no glossary resource/tool | CONFIRMED | ship a glossary resource / Terms block | low |
| A22 | audit systemic #5 | duplication / no routing | `server.py` (all); `excel-mcp_*` overlap | `test_run_suite`≈`test_suite_run`; `xlvba_*` overlaps `excel-mcp_*`; two health tools | no "prefer X" guidance | CONFIRMED | neighbor-routing lines in descriptions | low |
| A23 | audit systemic #6 | persistence gap | `server.py` (no `save_as`/`save_copy_as`/`export_fixed_format`); `excel_save` `workbook_tools.py:147–159`; `vba_export_project` | `excel_save` saves the working copy **in place** (no path arg); no save-as/export; `vba_export_project` exports VBA only | no described way to write results to a **new** file | CONFIRMED | add `save_as(path)`/export (product decision) | medium (new tool) |
| A24 | audit systemic #7 | destructive-by-default | `server.py:223–225, 251–255, 234–245` | close/restore/recover discard unsaved; descriptions disclose it; only `sheet_delete`/`range_merge` carry `allow_*` | "silent" overstated; no `confirm` flag | PARTIALLY_CONFIRMED | bold warnings; optional confirm | low–medium |
| A25 | audit systemic #8 | cancellation trap | `server.py:367–371`; `job_registry.py:328–353` | intent-only; runs to completion | "cancel" implies interruption it cannot do; description accurate | CONFIRMED (accurate) | rename `job_request_cancel` (optional) | medium (rename) |

---

## 3. Reconciliation of researcher reports (key claims)

### 3.1 Researcher 2A — MCP discovery & token cost — **AGREE (all key claims CONFIRMED)**
- "outputSchema emitted for all 91 tools; ~75% of `tools/list`" → **CONFIRMED**, reproduced exactly (338,547 / 451,108 = **75.05%**; desc 5.4%, input 17.9%).
- "annotations/title: none" → **CONFIRMED** (0/91 each; no `annotations=`/`title=` on any `add_tool`).
- "`defer_loading` absent from mcp/mcp_types 2.3.0" → **CONFIRMED** (0 recursive hits).
- "SDK supports outputSchema/structuredContent/annotations/resources" → **CONFIRMED** (`Tool.output_schema` `base.py:61`; `ToolAnnotations` `_types.py:1365–1401`; `CallToolResult.structured_content` `_types.py:1474`; `add_tool(..., annotations, structured_output)` `server.py:618–655`).
- "Baseline's 'none emitted' is wrong for outputSchema" → **CONFIRMED** (baseline §2 line 23 is wrong; annotations part was right).
- **One caveat I add:** 2A marks "Kilo forwards outputSchema to the model" as *unknown/likely*. This is the load-bearing unknown for the whole "shape undefined" family: the shapes **are** in `tools/list` outputSchema, but if the client does not forward outputSchema to the model, only the prose descriptions are visible. **UNVERIFIABLE from inside the session** — must be confirmed against the target Kilo build.
- Minor correction: 2A says baseline "90 add_tool"; the **exposed** tool count is **91** (90 `add_tool` + `server_health`). 2A states this correctly; researcher 2B's "90" undercounts by one.

### 3.2 Researcher 1A — tool contracts — **AGREE (all six contradictions CONFIRMED)**
- F1 CF enum vs rejection → CONFIRMED (`statuses.py:437–439`; `ranges_adapter.py:2247–2252`).
- F2 pivot `data` rejected / `hidden` overloaded → CONFIRMED (`statuses.py:470`; `pivot_tables_service.py:958–964`; overload documented `server.py:675–687`).
- F3 `vba_compile` never compiles → CONFIRMED (`compile_service.py:334–346`); description honest → naming issue.
- F5 `test_run_suite`≈`test_suite_run` → CONFIRMED (`test_tools.py:226–234`).
- F8 `pivot_field_apply_filter` enum vs manual-only → CONFIRMED (`statuses.py:510–524`; `pivot_tools.py:484–491`).
- F4 `job_cancel` accurate → CONFIRMED; F6 caps not stated → CONFIRMED (offset 0-based **is** stated, so 1A's own nuance is right).
- 1A's line-number corrections (descriptions at `server.py:216–770`, registration `1476–1704`) → **CONFIRMED**.

### 3.3 Researcher 1B — runtime policy — **AGREE (core claims CONFIRMED)**
- Access modes FULL/RESTRICTED, default FULL, RECOVERY auto in both → CONFIRMED (`statuses.py:54–65`; `config.py:49–53`; `approval.py:189–317`).
- Macro gate order + BG06/R13 → CONFIRMED (`macro_service.py:383–618`).
- `macro_classify_background` ungated, not authorization → CONFIRMED (`macro_service.py:721–826`).
- Snapshot gate reads in-memory `_latest_snapshot`; agent must create; `has_snapshot` unexposed; `snapshot_list` lists store → CONFIRMED (`snapshot_service.py:400–430, 385–395`).
- Build not returned; `excel_version`="16.0" → CONFIRMED (`workbook_adapter.py:99–116`).
- `to_dict()` lacks `isRetryable`/`suggestedNextActions`/`errorCategory` → CONFIRMED (`errors.py:138–147`; grep: none).
- Discoverability matrix row (a) "access mode: Yes (partial) via `excel_get_status`" → CONFIRMED (`session_service.py:1104`), and correctly notes it is **not** in `server_health` and not in `_DESC_STATUS`.
- Proposed `xlvba_environment` + additive error envelope → **AGREE in principle** (design, not a source claim); the `extra="forbid"` caution (`models.py:68`) is real and correctly flagged.
- Not independently re-derived (low priority, plausible from source): `macro_policy.classify_operation` being non-authoritative. Treat as UNVERIFIABLE-by-me / low-risk.

### 3.4 Researcher 2B — benchmark — **AGREE (design sound); one count correction**
- Grounding facts (no `save_as`; `excel_open` defaults; `range_write_values` payload; `table_update_rows` grammar; macro gates; `job_cancel` intent-only; `excel_recover` `prior_job="unknown"`; `vba_replace` hash; `test_suite_run` closed spec; fixtures) → all **CONFIRMED** against source.
- Tool-count assertion `>= 87` at `test_mcp_full_access_workflow.py:499` → CONFIRMED.
- **Correction:** 2B uses "90 tools registered / all-90". The exposed catalog is **91** (add `server_health`). Harness should assert against 91 (or `>=87` as today) and not claim "all-90".
- The benchmark itself is a plan (metrics/CTS/FCV/WC/UC/RC/UAA/TC); nothing to falsify. Its "exact numeric targets set after baseline" stance is methodologically sound.

---

## 4. Verified baseline (restatement)

- **Tool count:** **91 tools exposed** = 90 `server.add_tool(...)` (`server.py:1479–1704`) + 1 `@mcp.tool()` `server_health` (`server.py:1834–1838`). **4** resource templates (`server.py:1869–1943`), all URI-templated (0 static). **0** prompts.
- **Token-cost driver:** `tools/list` JSON = **451,108 chars** ≈ **~113K tokens** (4:1). Breakdown: **outputSchema 338,547 (75.05%)**, inputSchema 80,656 (17.9%), descriptions 24,272 (5.4%). Largest: `workbook_get_model` (28.8K output), `test_suite_run`/`test_run_suite` (22.1K/22.3K **input** — the suite DSL), `sheet_*` (~10.9K output ×7). **Descriptions are the smallest lever; outputSchema is the dominant lever.** (Reproduced independently.)
- **SDK capabilities (mcp 2.3.0 / mcp-types 2.3.0):** outputSchema auto-derived from return annotations and emitted for all tools; `structuredContent` produced on **success** only; `ToolError`→`is_error=true` **text-only**; `ToolAnnotations` (readOnly/destructive/idempotent/openWorld) supported but **not emitted**; resources with URI templates supported and used; **`defer_loading`/search meta-tool absent**.
- **Access model:** `AccessMode` ∈ {FULL, RESTRICTED}, default **FULL** (`config.py:49–53`). `READ_ONLY`→auto-allowed; `FULL`→MUTATING/EXECUTING allowed by policy (no per-action grant); `RESTRICTED`→deny-by-default (allowlist/live grant); `RECOVERY`→auto-allowed in both, pid-scoped (`approval.py:189–317`). Snapshot / writability / PathPolicy / Office-trust / audit run **independently** of mode. `access_mode` is discoverable via `excel_get_status` (`session_service.py:1104`), **not** via `server_health` and **not** mentioned in `_DESC_STATUS`.

---

## 5. Confirmed target agent-facing architecture

### 5.1 What MUST change (agent-facing)
1. **Descriptions (highest value, lowest risk):** for every Blocker/Major tool lead with Purpose / Use-when / Use-the-neighbor-instead / Preconditions / Parameters / Behavior / Result / Next actions. Add: `excel_open` defaults + writability; `range_write_values` "does not snapshot — call `snapshot_create` first" + `{rows,cols,values}`; `range_read_values` caps (100k request / 50k·8MiB response); `vba_replace_component_source` hash source; `macro_run` numbered gate list + define "approved"; table `where`/`updates` grammar; `name_*` `=`-optional; `sheet_add` `sheet_type`; `vba_compile` "diagnostic only".
2. **Request-enum pruning (schema):** split request vs read/classify enums so `ConditionalFormatType` (drop `color_scale/data_bar/icon_set/unsupported` from the **request**), `PivotFieldOrientation` (drop `data` from settable), `PivotFilterKind` (drop `label/value/date/top_bottom` from request) cannot advertise rejected values.
3. **Structured error envelope:** extend `to_dict()` **additively** with `errorCategory`, `isRetryable`, `suggestedNextActions` (derived by pure functions from `ErrorCode` + `context`); no new `ErrorCode` members, `SCHEMA_VERSION` stays 1. Note errors are text-content today — decide whether to also emit `structuredContent` on error (SDK currently does not).
4. **Mode/build/snapshot discovery:** add a read-only `xlvba_environment` (no COM) and/or extend `server_health`/`excel_get_status` with `access_mode`, `snapshot_required`, `created_by`, `build` (honest `null` if unavailable), `target_build`, and per-workbook `has_snapshot` (expose `SnapshotService.has_snapshot`). This unblocks ~8 tools.
5. **Collapse `test_run_suite`/`test_suite_run`** to one canonical tool (keep `background` + dict/spec union) and fix the docs that wrongly attribute `background` to `test_run_suite`.
6. **Reduce outputSchema footprint:** apply `structured_output=False` where the structured channel is redundant, and/or slim the oversized return DTOs (`WorkbookModel`, `SheetInfo`, the test-suite DSL). This attacks the 75% driver with **no client feature required**.
7. **Truthful MCP annotations:** emit `readOnlyHint`/`destructiveHint`/`idempotentHint` from existing access-class metadata (currently 0 emitted).
8. **Glossary + neighbor routing + examples** for the ~10 workflow-critical tools; disclose the validated target build once.
9. **Persistence:** decide on `save_as(path)`/export to close the gap (product decision).

### 5.2 What must NOT change
- Typed DTOs and the stable `ErrorCode` taxonomy (`errors.py:19–67`); fail-closed design; explicit `unknown`/`interrupted_unknown`/`indeterminate`/`suspected_stalled` non-success states.
- The snapshot gate (`require_snapshot`), approval policy (FULL default; RESTRICTED deny-by-default), RECOVERY auto-allow, PathPolicy/Office-trust, originals-safe guarantee, single-STA worker, BG06 + R13 gates.
- `SCHEMA_VERSION = 1` — all additions must be **additive** (new optional DTO fields / new models), respecting `extra="forbid"` (`models.py:68`).
- Do **not** add a standalone `xlvba_search_tools` as a context reducer (in Kilo it adds a 92nd always-loaded tool and removes nothing); only a full dispatcher (`xlvba_call`) would reduce context, which erodes per-tool permissions and client-side validation — out of scope unless explicitly mandated.
- Do **not** fabricate the 4-part build; return `build:null` unless a real read-only source exists.

---

## 6. Unresolved decisions needing Tech Lead / product input

1. **Rename `vba_compile`?** (compat HIGH: tests/docs). Or keep name + sharpen description only.
2. **Collapse the two test tools?** (compat HIGH: tests + docs + task_053 traceability). Which survives?
3. **Enum pruning vs description-only** for CF / pivot orientation / pivot filter kinds (schema change vs prose).
4. **outputSchema reduction strategy:** `structured_output=False` per tool vs slimming DTOs — tradeoff against any consumer relying on `structuredContent`.
5. **Discovery surface:** new `xlvba_environment` tool vs additive `SessionStatus.policy` (requires `extra="forbid"` DTO change without SCHEMA_VERSION bump) vs extending `server_health`.
6. **Error `structuredContent` on failure:** keep text-only JSON (SDK default) or also emit structured error content.
7. **`isRetryable` for `STALE_TARGET`/`REVISION_CONFLICT`** — `true` risks encouraging blind replay of a partially-applied mutation.
8. **Persistence gap:** add `save_as`/export (new tool) or accept in-place `excel_save` + `vba_export_project` as the only persistence.
9. **Client behavior confirmation:** does the target Kilo build forward `outputSchema` to the model (determines whether schema-shape claims matter) and surface resource meta-tools (`list_mcp_resources`/`read_mcp_resource`)? Observed absent in this session; docs say present.
10. **Canonical tool count** for harness/assertions: 91 (exposed) vs 90 (`add_tool`) vs `>=87` (current test).

---

## 7. Evidence index (primary)
- Registration/descriptions: `server.py:216–770` (desc), `1476–1704` (add_tool), `1808–1826` (call_tool error wrapping), `1834–1838` (server_health), `1869–1943` (resources).
- Enums/limits/errors: `statuses.py:17–27,54–65,418–440,463–470,510–524`; `models.py:65–68,83–96,99–112,390–396,666–672,969–980,1000–1030,1186–1211`; `errors.py:19–67,138–147`.
- Behavior: `compile_service.py:156–217,334–346`; `ranges_adapter.py:439+,2247–2252`; `pivot_tables_service.py:114,944–964`; `pivot_tools.py:478–505`; `job_registry.py:71–92,328–353,355–364`; `snapshot_service.py:385–430`; `approval.py:189–317`; `macro_service.py:383–618,721–826`; `test_tools.py:163–234`; `test_runner.py:378–401,1746–1763`; `session_service.py:1104`; `workbook_adapter.py:99–116`; `names_adapter.py:248–258`; `sheets_adapter.py:267–345`; `tables_adapter.py:892–1011`; `recovery_tools.py:46–68`; `health_tools.py:103–167`.
- SDK: `.venv/.../mcp/server/mcpserver/server.py:456,516–526,618–655`; `tools/base.py:35–127`; `utilities/func_metadata.py:122–131,186–226,364`; `exceptions.py:43`; `mcp_types/_types.py:1365–1408,1424,1474`.
- Consumers: `tests/integration_excel/test_mcp_full_access_workflow.py:499`; `tests/integration_excel/test_mcp_error_surfacing_matrix.py:357,416,438,501`; `tests/unit/test_macro_policy.py:94`; `docs/05-mcp-tools-reference.md:173–174,837–838`.
- Measurement: `.ai/tmp/measure_tools_2a.py` (re-run; figures in §1.2/§4).