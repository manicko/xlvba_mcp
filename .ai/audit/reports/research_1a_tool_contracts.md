# Research 1A — Tool Description Quality & Workflow Preconditions (xlvba MCP)

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca01fd10136716a598cad1db6b6845fed73`
- **Scope:** description constants + registration in `src/excel_vba_mcp/server.py`; focus-tool handlers in `src/excel_vba_mcp/api/*.py`; behavior in `src/excel_vba_mcp/application/*.py` and `src/excel_vba_mcp/adapters/excel/*.py`.
- **Method:** read-only. No production/test/config files modified.
- **Note on the brief's line estimates:** the `_DESC_*` constants actually live at **server.py:216–770** (not 749–1475); the read-only *resource* descriptions are 752–770. Tool registration is **server.py:1476–1704** (matches the brief). All line citations below are verified against HEAD.

Severity legend: **Blocker** (agent cannot complete the task / will reliably fail), **Major** (likely wrong tool or wrong payload, wasted round-trips), **Minor** (friction, missing example, incomplete result).

---

## 1. Summary table of findings

| # | Tool | Severity | One-line issue |
|---|------|----------|----------------|
| F1 | `range_set_conditional_formatting` | Major | Input enum `ConditionalFormatType` advertises `color_scale`/`data_bar`/`icon_set`/`unsupported`, but every non-plain type is rejected pre-COM. |
| F2 | `pivot_field_set_layout` | Major | `orientation` enum includes `data`, but `data` is rejected; `hidden` is overloaded (non-destructive hide vs `delete=true` destructive). |
| F3 | `vba_compile` | Major | Named "compile" but never compiles; always `indeterminate`/`NON_ACTIVATING_COMPILE_UNAVAILABLE`. Description is honest; the **name** misleads. |
| F4 | `job_cancel` | Minor | Accurately documented as intent-only ("never interrupts the running COM call"); limitation is real, not a doc bug. |
| F5 | `test_run_suite` vs `test_suite_run` | Major | Near-duplicate; `test_suite_run` just parses a dict then calls `test_run_suite`. Neighbor-selection confusion. |
| F6 | `range_read_values` | Minor | Offset base (0-based) is stated, but the numeric output cap (`MAX_RESPONSE_CELLS`=50 000, `MAX_RESPONSE_BYTES`=8 MiB) and request cap (`MAX_RANGE_CELLS`=100 000) are not. |
| F7 | `range_write_values` | Major | Description never states the payload object shape (`rows`/`cols`/`values`) nor the numeric ceilings; agent must guess the grid schema. |
| F8 | `pivot_field_apply_filter` | Major | `PivotFilterSpec.kind` enum advertises `label`/`value`/`date`/`top_bottom`/`unsupported`, but only `manual` is applied; others return `supported=false`. |
| F9 | `test_run_suite` / `test_suite_run` | Major | Missing precondition: `suite.workbook_id` **must be an open workbook** (and baseline snapshot is auto-created). Not stated. |
| F10 | `table_update_rows` | Minor | `where`/`updates` grammar is documented, but only enforced in the adapter, not the tool layer; no request example. |
| F11 | `table_append_rows` | Minor | `rows` shape (`[{header: value}]`) only in docstring, not in the tool description. |
| F12 | `sheet_add` | Minor | `sheet_type` default `"worksheet"` and "others rejected" not stated in description. |
| F13 | `name_create` / `name_update` | Minor | Leading `=` on `refers_to` is optional (normalized) but the description implies it is required. |
| F14 | `macro_run` | Minor | 30-arg `Application.Run` ceiling and `datetime.date` rejection not stated. |
| F15 | `vba_replace_component_source` / `vba_insert_code` | Minor | Do not say where `expected_source_hash` comes from (`vba_read_component` / `vba_list_components`). |
| F16 | `snapshot_restore` | Minor | Does not state unsaved changes to the current working copy are discarded. |
| F17 | `sheet_copy` | Minor | New sheet name is Excel-generated; not stated. |
| F18 | `excel_open` | Minor | Allowlist source (`EXCEL_VBA_MCP_ROOTS`) and default `read_only=False` not explicit; no example. |
| F19 | Cross-cutting | Minor | "Guarded mutation" / "writable session + snapshot + approval" phrasing is inconsistent across descriptions; approval default (full-access) stated only on some. |
| F20 | Cross-cutting | Minor | No request/response examples anywhere in the description constants. |

---

## 2. Detailed findings (focus tools) with evidence and proposed rewrites

Template used for rewrites: **Purpose / Use when / Do not use when / Preconditions / Parameters / Behavior / Result / Next actions.**

### F1 — `range_set_conditional_formatting` (Major)

- **Description:** server.py:560–567. Says "ColorScale / DataBar / IconSet are rejected explicitly."
- **Input enum:** `ConditionalFormatType` (statuses.py:418–440) includes `COLOR_SCALE="color_scale"`, `DATA_BAR="data_bar"`, `ICON_SET="icon_set"`, `UNSUPPORTED="unsupported"`. The tool's `spec.type` is typed `ConditionalFormatType` (models.py:969–980), so the JSON schema advertises all of them.
- **Actual behavior:** `validate_conditional_format_spec` (ranges_adapter.py:2233–2252) rejects any `spec.type not in _SUPPORTED_CF_TYPES` **before** the guard chain; `_SUPPORTED_CF_TYPES` is the plain `FormatCondition` set only. Service: ranges_service.py:515–570.
- **Ambiguity:** an agent reading the enum will send `color_scale` and be rejected; the description text and the enum disagree.
- **Proposed rewrite:**
  > **Purpose:** Add/modify/delete one *plain* conditional-format rule on a Range.
  > **Use when:** the rule is a `FormatCondition` (cell_value, expression, text_string, blanks, no_blanks, errors, no_errors, time_period, top10, above_average, unique_values).
  > **Do not use when:** the rule is a ColorScale / DataBar / IconSet — those COM objects are out of scope and are **rejected** (`color_scale`/`data_bar`/`icon_set` in the enum are not accepted). Use `range_get_conditional_formatting` to read them (reported `supported=false`).
  > **Preconditions:** writable session + snapshot + approval (full-access default; restricted requires approval).
  > **Parameters:** `mode` ∈ {`add`,`modify`,`delete`} (default `add`); `spec.priority` (1-based, sheet-wide) is the only stable identity and is **required** for `modify`/`delete`; `spec.type` one of the plain types above; `operator`/`formula1`/`formula2` required per type; `interior_color`, `font` optional.
  > **Behavior:** assigning a priority renumbers the other sheet-wide rules; never silently converts a rule type.
  > **Result:** `ConditionalFormatMutationResult` (action, rules, applied_priorities, unsupported, verified).
  > **Next actions:** `range_get_conditional_formatting` to confirm.

### F2 — `pivot_field_set_layout` (Major)

- **Description:** server.py:675–687 (accurate prose).
- **Enum:** `PivotFieldOrientation` (statuses.py:463–470) = `hidden/row/column/page/data`.
- **Actual behavior:** `_require_layout_orientation` (pivot_tables_service.py:944–969) rejects `data` (not in `_LAYOUT_ORIENTATIONS`, defined at :114) before STA dispatch; `set_pivot_field_layout` (:398–457) implements `hidden` as non-destructive `Orientation=0` and `hidden+delete=true` as destructive `PivotField.Delete`.
- **Ambiguity:** enum advertises `data` (rejected); `hidden` carries two meanings keyed off the separate `delete` flag.
- **Proposed rewrite:**
  > **Purpose:** Move a PivotField into a layout bucket (row/column/page/hidden) or remove it.
  > **Use when:** placing/hiding/removing a field.
  > **Do not use when:** you want the field in the **data** area — `orientation="data"` is **rejected**; use `pivot_field_set_aggregation` (`AddDataField`). To reorder within a bucket use `pivot_field_set_position`.
  > **Preconditions:** writable + snapshot + approval; worksheet/Table-backed cache only (else `unsupported_source`).
  > **Parameters:** `orientation` ∈ {`hidden`,`row`,`column`,`page`} (the `data` enum value is rejected); `delete` (default `false`): `hidden`+`false`=non-destructive `Orientation=0` (items/filter state kept); `hidden`+`true`=destructive `PivotField.Delete`.
  > **Behavior:** post-mutation re-read; `unsupported_source` for external/OLAP/Data Model.
  > **Result:** `PivotFieldInfo` (post-mutation; last-known for destructive delete).
  > **Next actions:** `pivot_fields_list` to confirm layout.

### F3 — `vba_compile` (Major — naming)

- **Description:** server.py:320–327 (honest: "never returns compiled or compile_error").
- **Actual behavior:** `compile_project` (compile_service.py:156–217) submits a read-only probe and `_build_result` (:334–346) **always** returns `CompileStatus.INDETERMINATE` with reason `NON_ACTIVATING_COMPILE_UNAVAILABLE` (or VBOM/locked reason). `is_compile` is always `False` (:92).
- **Assessment:** description matches code — **not** a doc contradiction. The problem is the tool **name** implies compilation. This is a discoverability/affordance issue, not a false description.
- **Proposed rewrite (keep name, sharpen first line):**
  > **Purpose:** Return a **fail-closed, non-activating** compile *diagnostic* for an open workbook. This tool does **not** compile VBA.
  > **Use when:** you need broken-reference / protection / component-inventory / VBOM evidence.
  > **Do not use when:** you need a real compile result — none is available on this build; do not expect `compiled`/`compile_error`.
  > **Preconditions:** open workbook.
  > **Result:** `CompileResult{status=indeterminate, reason, diagnostics(is_compile=false), command_discovery}`.
  > **Next actions:** `vba_get_compile_diagnostics` for the cached copy; fix references and re-probe.

### F4 — `job_cancel` (Minor — accurate)

- **Description:** server.py:367–371; handler job_tools.py:113–128; registry `cancel` job_registry.py:328–353 sets `cancel_requested=True` only; the STA COM call runs to completion (module docstring :29–30).
- **Assessment:** description is accurate. The limitation (cancel cannot interrupt) is real and correctly disclosed. No rewrite required beyond optionally adding "use `job_get` to observe `cancel_requested`."

### F5 — `test_run_suite` vs `test_suite_run` (Major — near-duplicate)

- **Descriptions:** server.py:257–266 and 271–281.
- **Actual behavior:** `test_suite_run` (test_tools.py:189–234) parses a dict via `_parse_suite` then **calls `test_run_suite(spec)`** (:234); `test_run_suite` (:163–186) is the canonical runner. Difference: `test_suite_run` accepts `dict | TestSuiteSpec` and a `background` flag; `test_run_suite` accepts only `TestSuiteSpec`, no background.
- **Ambiguity:** two tools with near-identical descriptions; an agent cannot tell which to call.
- **Proposed rewrite (differentiate):**
  > `test_run_suite` — **Purpose:** run a declarative suite given as a `TestSuiteSpec`. **Use when:** you already hold a typed spec and want a blocking run. **Do not use when:** you have a raw JSON mapping or need `background=true` → use `test_suite_run`.
  > `test_suite_run` — **Purpose:** validate+run a suite from a raw JSON mapping (or spec), optionally in background. **Use when:** an MCP caller sends a plain object or needs `background=true`. **Do not use when:** you already have a `TestSuiteSpec` and want the simplest blocking call → `test_run_suite`.

### F6 — `range_read_values` (Minor)

- **Description:** server.py:459–464; handler range_tools.py:174–203; adapter read_values ranges_adapter.py:856–910; window `_window` :3704–3736.
- **Stated:** offset is 0-based; limit caps rows; `truncated` flag; oversized requests rejected.
- **Not stated (limits):** request rejected when `rows*cols > MAX_RANGE_CELLS` (100 000, models.py:666); response truncated to `MAX_RESPONSE_CELLS` (50 000, :670) and `MAX_RESPONSE_BYTES` (8 MiB, :671).
- **Proposed addition:** append to description: "The target block is rejected if it exceeds 100 000 cells; the returned grid is capped at 50 000 cells / 8 MiB and `truncated=true` marks a cut-short window."

### F7 — `range_write_values` (Major)

- **Description:** server.py:471–480; handler range_tools.py:304–328; payload `RangeWriteRequest` models.py:1000–1030.
- **Gap:** description says "bounded 2-D value grid" but never gives the object shape. The payload is `{rows:int≥1, cols:int≥1, values:[[CellValue,...],...]}` with `rows*cols ≤ 100 000`, `rows ≤ 100 000`, `cols ≤ 512` (models.py:667–672, 1009–1030). `address` is the **top-left anchor** whose dimensions must equal `rows×cols` (range_tools.py:320). A leading `=` is stored as text (:474).
- **Proposed rewrite:**
  > **Purpose:** Write a rectangular block of typed values to a Range.
  > **Use when:** writing constants/text (a leading `=` is stored as **text**, not a formula — use `range_write_formulas` for formulas).
  > **Do not use when:** writing formulas → `range_write_formulas`; writing inside a Table's structure → table tools.
  > **Preconditions:** writable working copy + snapshot + approval; inside a Table only if row/column structure is unchanged.
  > **Parameters:** `address` = top-left A1 anchor (block spans `payload.rows × payload.cols`); `payload` = `{rows, cols, values}` with `values` a `rows×cols` grid; ceilings `rows≤100000`, `cols≤512`, cells≤100000.
  > **Behavior:** shape validated pre-COM (wrong size rejected, no change); postcondition re-read.
  > **Result:** `RangeMutationResult{range_before, range_after, verified, postcondition_values}`.

### F8 — `pivot_field_apply_filter` (Major)

- **Description:** server.py:727–738 (accurate prose).
- **Enum:** `PivotFilterSpec.kind` defaults `MANUAL` (models.py:1201–1211); `PivotFilterKind` (statuses.py:510–524) includes `label/value/date/top_bottom/unsupported`.
- **Actual behavior:** only `manual` (`PivotItems[i].Visible`) is applied; other kinds return `filter.supported=false` with a reason (pivot_tools.py:478–505; service apply_pivot_field_filter).
- **Ambiguity:** enum advertises filter kinds that are never applied.
- **Proposed rewrite:** add "Only `kind="manual"` (visible_items) is applied. `label`/`value`/`date`/`top_bottom` are advertised by the enum but **unsupported** on this build and return `supported=false` — they are never converted to a manual filter."

### F9 — `test_run_suite` / `test_suite_run` missing precondition (Major)

- **Evidence:** `TestSuiteSpec.workbook_id` is required (test_spec.py:457); `run_suite` docstring "The suite's `workbook_id` must be an open workbook" (test_runner.py:378) and `_resolve_workbook` (:401). Baseline is auto-created (`BaselinePolicy.CREATE`, test_spec.py:458; test_runner.py:379–380).
- **Gap:** neither description states the workbook must be opened with `excel_open` first, nor that a baseline snapshot is created automatically.
- **Proposed addition:** "Precondition: `suite.workbook_id` must be an **open** workbook (call `excel_open` first). A baseline snapshot is created automatically unless `baseline="use_existing"`."

### F10 — `table_update_rows` (Minor)

- **Description:** server.py:428–432; handler table_tools.py:220–242 (only checks list-of-dicts via `_require_row_dicts`); grammar enforced in adapter `_resolve_updates` (tables_adapter.py:895–928) and `_resolve_where` (:954+).
- **Gap:** "exactly one selector" is enforced in the adapter, not the tool; no example. Add example: `updates=[{"index":2,"values":{"Qty":5}},{"where":{"ID":"A-1"},"values":{"Qty":9}}]`.

### F11 — `table_append_rows` (Minor)

- **Description:** server.py:418–422; handler table_tools.py:166–184. `rows` shape only in docstring (:177). Add to description: "`rows` is a non-empty list of `{header: value}` objects; headers must match table columns (case-insensitive)."

### F12 — `sheet_add` (Minor)

- **Description:** server.py:378–382; handler sheet_tools.py:93–108; adapter add_sheet rejects non-`worksheet` (sheets_adapter.py:267–277).
- **Gap:** default `sheet_type="worksheet"` and "other values rejected" not in description. Add: "`sheet_type` defaults to `"worksheet"`; only `"worksheet"` is supported (others rejected)."

### F13 — `name_create` / `name_update` (Minor)

- **Description:** server.py:578–587; handler name_tools.py:146–206; adapter `_normalize_refers_to` accepts with/without leading `=` (names_adapter.py:248–258).
- **Gap:** description examples use `=...` implying required. Add: "`refers_to` may be given with or without a leading `=` (normalized)."

### F14 — `macro_run` (Minor)

- **Description:** server.py:336–346; handler macro_tools.py:87–129; coercion macro_runner.py:83–126.
- **Gap:** `Application.Run` 30-arg ceiling and `datetime.date` rejection not stated. Add: "args: str/int/float/bool/None/datetime and 1-D/2-D lists (≤30 args); a bare `datetime.date` is rejected — pass `datetime.datetime`."

### F15 — `vba_replace_component_source` / `vba_insert_code` (Minor)

- **Description:** server.py:308–319; handlers vba_tools.py:191–264. Add: "Obtain `expected_source_hash` from `vba_read_component` (or `vba_list_components`)."

### F16 — `snapshot_restore` (Minor)

- **Description:** server.py:251–255; handler test_tools.py:122–143. Add: "Unsaved changes in the current working copy are discarded (closed without saving)."

### F17 — `sheet_copy` (Minor)

- **Description:** server.py:388–392; handler sheet_tools.py:130–144. Add: "The copy's name is generated by Excel (typically `<name> (2)`)."

### F18 — `excel_open` (Minor)

- **Description:** server.py:216–222; handler workbook_tools.py:88–127. Add: "Path must be under an allowlisted root (`EXCEL_VBA_MCP_ROOTS`, default cwd+temp). Defaults: `mode="execute"`, `read_only=False`."

### F19/F20 — Cross-cutting (Minor)

- Inconsistent guard phrasing ("Guarded mutation: writable session + snapshot + approval" vs "requires a writable working copy, a successful snapshot"). Standardize one sentence.
- No request/response examples in any `_DESC_*` constant.

---

## 3. Neighbor routing (use X instead of Y)

| Instead of | Use | Why |
|------------|-----|-----|
| `range_write_values` (formula) | `range_write_formulas` | `=...` in write_values is stored as text. |
| `range_set_conditional_formatting` (color/data/icon) | `range_get_conditional_formatting` (read-only, `supported=false`) | those rule types are rejected on write. |
| `pivot_field_set_layout` (`orientation="data"`) | `pivot_field_set_aggregation` | `data` layout is rejected; data area is `AddDataField`. |
| `pivot_field_set_layout` (reorder) | `pivot_field_set_position` | layout sets bucket; position sets order. |
| `pivot_field_apply_filter` (label/value/date/top) | `pivot_field_get_filters` (read) | only `manual` filter is applied. |
| `test_run_suite` (raw JSON / background) | `test_suite_run` | only `test_suite_run` accepts a mapping + `background`. |
| `test_suite_run` (typed spec, blocking) | `test_run_suite` | canonical blocking runner. |
| `vba_compile` (expecting compiled/compile_error) | `vba_get_compile_diagnostics` | compile never returns compiled/error; diagnostics only. |
| `job_cancel` (expecting interruption) | `job_get` (observe `cancel_requested`) + `excel_recover` | cancel is intent-only; recovery is the real stop. |
| `sheet_add` (non-worksheet) | — (unsupported) | only `"worksheet"` supported. |
| `range_insert`/`range_delete` over a Table | table tools | ListObject intersection is rejected. |

---

## 4. Verified contradictions (CONFIRMED / REFUTED / UNCLEAR)

| Claim | Verdict | Evidence |
|-------|---------|----------|
| `range_set_conditional_formatting` enum advertises color_scale/data_bar/icon_set but they are rejected | **CONFIRMED** | enum statuses.py:437–439; rejection ranges_adapter.py:2247–2252 (`_SUPPORTED_CF_TYPES`). |
| `pivot_field_set_layout` enum includes `data` but it is rejected | **CONFIRMED** | enum statuses.py:470; rejection pivot_tables_service.py:958–968 (`_LAYOUT_ORIENTATIONS` :114). |
| `pivot_field_set_layout` `hidden` overloaded (hide vs delete) | **CONFIRMED (documented)** | service :414–416; handler pivot_tools.py:348–365. Overload exists but the description documents it. |
| `vba_compile` named "compile" but never returns compiled/compile_error (always indeterminate) | **CONFIRMED** | compile_service.py:334–346 always `INDETERMINATE`; :161–164 "never returns compiled/compile_error". Description (server.py:320–327) is honest → naming issue, not doc contradiction. |
| `job_cancel` "never interrupts the running COM call" | **CONFIRMED (accurate)** | job_registry.py:328–345; module doc :29–30. Description matches code. |
| `test_run_suite` vs `test_suite_run` near-duplicate | **CONFIRMED** | test_tools.py:226–234 (`test_suite_run` → `test_run_suite`). |
| `range_read_values` offset base and max output size stated or not | **Offset base: STATED** (0-based, server.py:462). **Max output size: NOT STATED** (limits exist at models.py:666–671 but absent from description). |

Additional verified (not in brief): `pivot_field_apply_filter` enum advertises label/value/date/top_bottom but only manual is applied — **CONFIRMED** (statuses.py:510–524; pivot_tools.py:484–491).

---

## 5. Open questions for the Auditor / Validator

1. **Naming vs behavior:** Is renaming `vba_compile` (e.g. `vba_compile_diagnostics`) in scope, or must the honest description carry the full burden? (F3)
2. **Enum pruning:** Should `ConditionalFormatType` / `PivotFieldOrientation` / `PivotFilterKind` be split into "request" vs "read/classify" enums so the request schema cannot advertise rejected values? (F1, F2, F8)
3. **Tool consolidation:** Is merging `test_run_suite`/`test_suite_run` (single tool with `background` + dict/spec union) permitted, or must both remain for task_053 traceability? (F5)
4. **Limits disclosure:** Confirm the canonical numbers to publish (MAX_RANGE_CELLS 100 000, MAX_RESPONSE_CELLS 50 000, MAX_RESPONSE_BYTES 8 MiB, MAX_WRITE_PAYLOAD_CELLS 100 000, MAX_RANGE_COLS 512) — are these stable API guarantees or internal knobs? (F6, F7)
5. **`range_write_values` payload schema:** Should the description embed the `{rows,cols,values}` shape, or is the JSON-schema `RangeWriteRequest` considered sufficient (the description never restates it)? (F7)
6. **Precondition disclosure for test tools:** Confirm `workbook_id`-must-be-open is a hard precondition (test_runner.py:378) worth surfacing in the description. (F9)
7. **Approval wording:** Should every mutating description uniformly state "full-access default; restricted requires approval", or is the current per-tool variation acceptable? (F19)