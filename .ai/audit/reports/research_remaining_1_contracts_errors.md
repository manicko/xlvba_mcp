# Research 1 — Remaining-Usability Contracts, Errors, WorkbookInfo, Names, Confirm, Descriptions

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d`
- **Scope:** RESEARCH ONLY. No code/test/config changes. Backward compatibility NOT required (old names may be removed). Lazy loading excluded.
- **Method:** Direct source reading (primary source of truth). SDK behaviour verified by inspecting `mcp.server.mcpserver.exceptions` in `.venv`. Runtime counts verified by importing the server (`list_tools()`).
- **Confidence:** HIGH for all file:line facts below (read directly). MEDIUM where a design choice is proposed (marked "proposal").

---

## 1. Canonical naming / duplicate consolidation

### 1.1 `vba_compile` → `vba_compile_diagnostics` (diagnostics-only semantics)

**Verified facts**
- Handler: `api/compile_tools.py:63` `async def vba_compile(workbook_id: str) -> CompileResult`. Docstring `:64-77` already states it "never compiles and never returns compiled / compile_error".
- Registered: `server.py:1979-1984` (`name="vba_compile"`, `description=_DESC_VBA_COMPILE`, `annotations=_annotations("vba_compile")`).
- Description constant: `_DESC_VBA_COMPILE` `server.py:407-421` (contains the truthful phrase "does NOT compile", enforced by `tests/unit/test_description_contract.py:56`).
- Import into server: `server.py:55`.
- Re-export: `api/__init__.py:10` and `:59`.
- `__all__`: `api/compile_tools.py:102`.
- Annotation status: **not** in `_ANN_READ_ONLY` (`server.py:1747-1790`) — comment `server.py:1746` "macro_run and vba_compile are intentionally NOT read-only". Not in `_ANN_DESTRUCTIVE`/`_ANN_IDEMPOTENT`.

**Internal callers of the literal `"vba_compile"` (NOT the tool registration):**
- `application/compile_service.py:198` `compile_fn.__name__ = "vba_compile"` (STA-worker call label for audit/retry).
- `application/compile_service.py:261` `context={"tool": "vba_compile"}`; `:276` `tool="vba_compile"`; `:352` `tool="vba_compile"` (audit records).
- `domain/test_spec.py:73` `VBA_COMPILE = "vba_compile"` — a **TestAction step enum**, semantically distinct from the tool; dispatched by `application/test_runner.py:622` (`if action is TestAction.VBA_COMPILE`) → `_exec_compile` → `compile_service.compile_project` (`test_runner.py:773`). **Do not rename the test action** as part of the tool rename (different contract surface); note the shared string for the Auditor.
- `security/macro_policy.py:165` `_EXECUTING_TOOLS` (explicitly **stale / illustrative / not wired** — see `macro_policy.py:186-191`).
- `domain/errors.py:576` docstring mention.
- `server.py:1746` comment.

**Tests referencing the tool name / action:**
- `tests/unit/test_compile_service.py:391` `assert submit["name"] == "vba_compile"`; `:517`.
- `tests/unit/test_description_contract.py:39` (`_CRITICAL`), `:56` (`_TRUTH`).
- `tests/unit/test_macro_policy.py:92`.
- `tests/unit/test_domain_models.py:144,184` (TestAction — keep).
- `tests/unit/test_test_runner.py:891,928,958` (action — keep).
- `tests/integration_excel/test_end_to_end_test_runner.py:298,870,883` (action — keep).
- `tests/integration_excel/test_vba_compile.py` (filename + docstring `:32`).

**Docs:** `docs/01-overview.md:79,129,143`; `docs/05-mcp-tools-reference.md:63,92,96,98-100,124,238,241,808-809,814,849,895,922,942`; `docs/06-excel-capabilities.md:111,274,321`; `docs/07-workflows-and-examples.md:525,547,569,586,610,627,689`; `docs/09-limitations-and-compatibility.md:86,112`; `docs/11-development-and-maintenance.md:179`; `docs/04-architecture.md:510`; Russian mirrors `docs/ru/*` (01,05,06,07,08,09,10,11).

**Benchmark:** `vba_compile` not present in `.ai/tmp/benchmark_tasks.yaml` (only `job_cancel`, `test_suite_run`). No benchmark change needed for this rename (but add a task — see §7).

### 1.2 `job_cancel` → `job_request_cancel` (intent-only semantics)

**Verified facts**
- Handler: `api/job_tools.py:113` `async def job_cancel(job_id: str) -> JobStatus`. Docstring `:114-124` states "Record a job's cancel **intent** (intent-only; never interrupts COM)".
- Registered: `server.py:2015-2020`.
- Description: `_DESC_JOB_CANCEL` `server.py:489-500` (truthful phrase "intent-only; never interrupts", enforced `test_description_contract.py:58`).
- Import: `server.py:63`. `__all__`: `job_tools.py:133`.
- Annotation: not in any `_ANN_*` set (read-only by convention; docs list it read-only).
- Semantics source: `application/job_registry.py:29-31` ("job_cancel is intent-only … cannot interrupt the uninterruptible COM call").

**Internal callers:** none via `start_long_op(partial(job_cancel,…))`. Only docstrings (`job_registry.py:29`).

**Tests:** `tests/unit/test_job_tools.py:20,122-165` (`test_job_cancel_records_intent_only`, `_empty_arg_rejected`, `_unknown_id_is_not_found`, `:165` `assert "job_cancel" in tools`); `tests/unit/test_job_registry.py:11,160-162`; `tests/unit/test_description_contract.py:37,58`.

**Docs:** `docs/05-mcp-tools-reference.md:122,202,204,880`; `docs/04-architecture.md:282`.

**Benchmark:** `.ai/tmp/benchmark_tasks.yaml:471-477` (`expected_first_tool: job_cancel`, `expected_sequence: [job_cancel]`, `job_cancel: {job_id: …}`, `token_measure: tool_description_tokens(job_cancel)+…`). **Must update** to `job_request_cancel`.

### 1.3 `test_run_suite` vs `test_suite_run` — consolidate to `test_suite_run` (superset)

**Verified facts**
- `test_run_suite`: `api/test_tools.py:163` `async def test_run_suite(suite: TestSuiteSpec) -> TestRunResult` — typed-spec, blocking only. Registered `server.py:1907-1912`.
- `test_suite_run`: `api/test_tools.py:189` `async def test_suite_run(suite: TestSuiteSpec | dict[str, Any], *, background: bool = False) -> TestRunResult | RunningJob` — **superset** (accepts mapping, supports `background`). Registered `server.py:1913-1918`.
- **Internal coupling:** `test_suite_run` calls `test_run_suite` directly (`test_tools.py:234`) and via `start_long_op(partial(test_run_suite, spec), …)` (`test_tools.py:228-233`). Consolidation must keep **one request schema + one execution contract**; the `partial(test_run_suite, …)` reference is the only `start_long_op` coupling.
- Descriptions: `_DESC_TEST_RUN_SUITE` `server.py:303-319`; `_DESC_TEST_SUITE_RUN` `server.py:324-348` (each "Do not use when" points at the other; `test_suite_run` is `_EXAMPLE_TOOLS` member).
- Imports: `server.py:144-145`. `__all__`: `test_tools.py:277-278`.
- Audit context string: `application/test_runner.py:1763` `context={"workbook_id": …, "tool": "test_run_suite"}`.
- Stale policy set: `security/macro_policy.py:167` `_EXECUTING_TOOLS` (illustrative only).

**Tests:** `tests/unit/test_description_contract.py:42,43,51`; `tests/integration_excel/test_mcp_error_surfacing_matrix.py:357,416,438,501-503`; `tests/unit/test_macro_policy.py:94`.

**Docs:** `docs/01-overview.md:88-90,127`; `docs/05-mcp-tools-reference.md:81,99,121,180-195,765,849-850,900`; `docs/06-excel-capabilities.md:114`; `docs/07-workflows-and-examples.md:455,539-540`; Russian mirrors.

**Benchmark:** `.ai/tmp/benchmark_tasks.yaml:554,556,569,570` (`test_suite_run`); `must_not` references `test_run_suite` (`:569`). Update `must_not` wording after removal.

**Decision (HIGH):** canonical = **`test_suite_run`**. Remove `test_run_suite` outright (no alias). Keep the internal helper (rename the typed-spec core to a private `_run_suite_blocking` or inline it) so `test_suite_run` retains one execution path.

### 1.4 Full migration surface (summary checklist for the implementer)

| Surface | `vba_compile`→`vba_compile_diagnostics` | `job_cancel`→`job_request_cancel` | remove `test_run_suite` |
|---|---|---|---|
| Handler def + `__all__` | compile_tools.py:63,102 | job_tools.py:113,133 | test_tools.py:163,277 |
| Server import | server.py:55 | server.py:63 | server.py:144 |
| `add_tool` name/desc/annotations | server.py:1979-1984 | server.py:2015-2020 | server.py:1907-1912 |
| `_DESC_*` constant | 407-421 | 489-500 | 303-319 (delete) |
| `_ANN_*` sets | (none) | (none) | (none) |
| `api/__init__.py` | 10,59 | — | — |
| Service audit labels | compile_service.py:198,261,276,352 | — | test_runner.py:1763 |
| Stale policy set | macro_policy.py:165 | — | macro_policy.py:167 |
| Unit tests | test_compile_service, test_description_contract, test_macro_policy | test_job_tools, test_job_registry, test_description_contract | test_description_contract, test_mcp_error_surfacing_matrix, test_macro_policy |
| Docs (EN+RU) | many (see §1.1) | 05,04 | many (see §1.3) |
| Benchmark yaml | (add task) | 471-477 | 554-570 |

**Risk:** the `TestAction.VBA_COMPILE = "vba_compile"` enum (`test_spec.py:73`) and the `_EXECUTING_TOOLS` set (`macro_policy.py:161-169`) are **not** the tool name; renaming them is out of scope and would break the declarative-suite contract. The Auditor must confirm the implementer does not touch them. The `macro_policy` sets are documented as stale/non-wired (`:186-191`) — updating them is optional cleanup, not required.

---

## 2. Guaranteed error contract (every `isError=true` → full envelope)

### 2.1 Current state (verified)

- `ExcelVbaMcpError.to_dict()` `domain/errors.py:335-353` emits **8 keys**: `code, message, hresult, hresult_hex, context, errorCategory, isRetryable, suggestedNextActions`. The three added keys come from pure functions `category_for` (`:249`), `retryable_for` (`:267`), `actions_for` (`:284`).
- `category_for`/`retryable_for`/`actions_for` are **total** over every `ErrorCode` (29 members, `:28-67`) and the job-state strings `expired`/`interrupted_unknown` (`:168-178`); unknown → `validation` / not-retryable / default actions (`:239-240`). Proven by `tests/unit/test_error_envelope.py:47-62`.
- Boundary wrapper: `ExcelVbaMcpServer.call_tool` `server.py:2484-2502`:
  1. `try: return await super().call_tool(...)`.
  2. `except ToolError as exc:` → `cause = _find_domain_cause(exc)` (`server.py:2383-2398`, walks `__cause__` chain for an `ExcelVbaMcpError`).
  3. If no domain cause → `cause = _validation_error_as_domain(exc)` (`server.py:2432-2461`, walks chain for `pydantic.ValidationError`, wraps as `INVALID_RANGE` with `context.reason="invalid_arguments"` + sanitized `validation`).
  4. If still `None` → `raise` (re-raise the SDK `UnexpectedToolError`).
  5. Else `raise ToolError(json.dumps(cause.to_dict(), ensure_ascii=False)) from cause`.
- SDK behaviour (verified from `.venv/.../mcp/server/mcpserver/exceptions.py`): `ToolError` → client gets `is_error=True` with the message in `content` (text). `UnexpectedToolError` (subclass) → message is only `"Error executing tool <name>"`, original withheld. The SDK itself raises `UnexpectedToolError` for an **unknown tool** and for **argument-schema failures**; both are caught by `except ToolError`.
- COM errors: the STA worker converts `com_error` via `error_from_hresult` (`errors.py:609-624`) into a typed `ExcelVbaMcpError`, so COM failures on the worker path **are** domain errors (structured). HRESULT→code map `:72-80`.

### 2.2 Gaps vs the target contract

Target: **every** `isError=true` response contains `errorCode, errorCategory, message, isRetryable, suggestedNextActions, details`.

| Gap | Evidence | Effect |
|---|---|---|
| Key is `code` not `errorCode` | `errors.py:345` | Field-name mismatch |
| Key is `context` not `details` | `errors.py:349` | Field-name mismatch |
| Non-domain, non-validation exceptions escape unstructured | `server.py:2500-2501` `raise` | Client gets bare `"Error executing tool <name>"` (no envelope). Covers: output-conversion failures, a raw `RuntimeError` not wrapped by the worker, any future non-`ExcelVbaMcpError`. Proven-possible by `test_mcp_error_surfacing.py:75-80` (`test_non_domain_error_surfaced_as_unexpected_tool_error`). |
| `structuredContent` cannot carry errors | SDK `ToolError` is text-only (`exceptions.py` docstring) | Must use additive **text-JSON** envelope, never `structuredContent` |

### 2.3 Design — boundary wrapping + field mapping (proposal, HIGH feasibility)

1. **Single envelope builder** `_error_envelope(exc: BaseException) -> dict[str, Any]` in `server.py` (or `domain/errors.py`), producing exactly the 6 target keys:
   - `errorCode` ← `domain.code.value`
   - `errorCategory` ← `category_for(domain.code)`
   - `message` ← `domain.message`
   - `isRetryable` ← `retryable_for(domain.code)`
   - `suggestedNextActions` ← `actions_for(domain.code, domain.context)`
   - `details` ← `domain.context` (rename of `context`)
   - Optionally keep `hresult`/`hresult_hex` **additively** (allowed; target says "contains", not "only").
2. **Rename in `to_dict`** (`errors.py:344-353`): `code`→`errorCode`, `context`→`details` (backward compat not required). Update `tests/unit/test_error_envelope.py:30-31,41-42` accordingly.
3. **Close the escape hatch** in `call_tool` (`server.py:2492-2502`): when `cause is None` after both `_find_domain_cause` and `_validation_error_as_domain`, **do not `raise`** — synthesize a fallback envelope from the caught `UnexpectedToolError` (see code choice below) and `raise ToolError(json.dumps(fallback, …)) from exc`. This guarantees the 6 keys for *every* `isError` path.
4. **Fallback code choice (open for Auditor):** the closed spec §11 vocabulary has no "internal/unexpected" code. Options:
   - (a) reuse `ErrorCode.COM_BUSY` (generic fallback already used by `error_from_hresult` `:621`) — but semantically wrong for a non-COM crash;
   - (b) add an **additive** `ErrorCode.INTERNAL_ERROR` (category `validation`/new `internal`, not retryable) — cleanest, spec §11 Stage-1 set is closed but Phase-2 codes were appended additively (`errors.py:43-67`), so an additive code is consistent with precedent;
   - (c) reuse `INVALID_RANGE` (already the argument-validation code) — wrong for a crash.
   - **Recommendation:** (b) additive `INTERNAL_ERROR`, category `internal` (or `validation`), `isRetryable=False`, actions `["inspect server logs; the failure was not a domain error"]`.
5. **`details` must be JSON-serializable** for the fallback too (the caught exception's `str()` only — never embed the live object). Reuse the `_jsonable` sanitizer pattern (`server.py:2420-2427`).
6. **Do not** attempt `structuredContent` (SDK text-only). The text-JSON envelope is the contract.

**Net:** the Pydantic class alone is insufficient; the guarantee lives at `call_tool`. Every `isError` then carries the 6 keys.

---

## 3. `WorkbookInfo` field reference (`domain/models.py:86-99`)

| Field | Type | Null? | Default | Meaning (verified) |
|---|---|---|---|---|
| `workbook_id` | `str` | no | — | Stable server-assigned id (`:89`). |
| `name` | `str` | no | — | Workbook name (`wb.Name`, `workbook_adapter.py:97,110`). |
| `format` | `str` | no | — | File format e.g. `'xlsm'` (`:91`; `_format_from_name`). |
| `path` | `str` | no | — | **Actual working path** = `wb.FullName` (`workbook_adapter.py:98,112`). The mutable working copy under `work/`. |
| `source_path` | `str \| None` | yes | `None` | **Original allowlisted (canonical) path** the working copy was materialized from (`:93`; set on open `session_service.py:384`, on save `:518`). `None` when not applicable (e.g. a freshly restored snapshot baseline, closed book). |
| `read_only` | `bool` | no | — | Mirrors `wb.ReadOnly` (`workbook_adapter.py:99,114`). |
| `has_vba` | `bool` | no | — | `_has_vba(name or full_name)` (`:115`). |
| `excel_version` | `str` | no | — | `Application.Version` best-effort (`:100-106,116`). |
| `session_id` | `str` | no | — | Owning session id (`:97`). |
| `mode` | `SessionMode` | no | — | `inspect` / `execute` (`:98`). |
| `model_version` | `int` | no | `1` | DTO contract version (`:99`). |

**path vs source_path (working vs source):** `path` is the live working copy (`work/<...>/report.xlsm`); `source_path` is the untouched original (`C:/data/report.xlsm`). The original is **never** modified (`session_service` copies via `shutil.copy2`; baseline `:45`). On `excel_save` the working copy is written in place; `path` and `source_path` are unchanged.

**Examples (proposal):**
- New open (writable working copy): `{workbook_id:"wb-1", name:"report.xlsm", format:"xlsm", path:"…/work/wb-1/report.xlsm", source_path:"C:/data/report.xlsm", read_only:false, has_vba:true, excel_version:"16.0", session_id:"s1", mode:"execute", model_version:1}`.
- Read-only inspect: `read_only:true, mode:"inspect"`, `source_path` still the original.
- After save: identical `path`/`source_path`; only `dirty` (on `SessionStatus`, not here) changes.
- Snapshot restore: `path` = new working copy materialized from the snapshot baseline; `source_path` = the snapshot's stored path (the allowlisted source it was copied from), **not** the pre-restore original — the pre-restore working copy was discarded.
- **Do NOT add `original_name`** — redundant with `name` (the original file name equals `name`; the original *path* is `source_path`).

---

## 4. Cross-scope defined names — conflict spec (`adapters/excel/names_adapter.py`, `application/names_service.py`)

### 4.1 Scope model (verified)
- A name is identified by `(name, scope)`, plus the owning sheet for worksheet scope (`names_adapter.py:16-17`). Excel permits the same identifier once at workbook level and once per worksheet.
- `Workbook.Names` enumerates all; `Worksheet.Names` enumerates that sheet's (`:13-17`). Scope is derived from `Name.Parent` via `Parent.FullName` probe (`:142-156`).
- Inventory de-dups by `(name, scope, sheet)` (`:498-533`).

### 4.2 Conflict rules (verified `_guard_no_conflict` `:576-618`, called only by `create_name` `:422`)
- **Duplicate (same scope):** an existing name with the same `scope.value` **and** same sheet → reject `reason="duplicate"` (`:592-606`).
- **Cross-scope conflict:** any existing name with the same identifier at a **different** `(scope, sheet)` → reject `reason="scope_conflict"`, message "refusing to create an ambiguous cross-scope name" (`:607-618`). `context.existing` lists `[{scope, sheet}]`.
- **Net effect:** creating a name whose identifier already exists **anywhere else** (workbook scope, or a *different* worksheet) is **rejected** — stricter than raw Excel (which would allow workbook+worksheet coexistence). The server refuses to create an ambiguous cross-scope name.
- **Resolution order** (`_resolve_name` `:535-574`): filter entries by `scope` + case-insensitive `name` + (worksheet) sheet; 0 matches → not-found error (`:558`); >1 → `reason="ambiguous"` refuse (`:563`); exactly 1 → resolved.
- **Update cannot change scope** (`update_name` `:439-470`): resolves the name in the given scope, writes `RefersTo` directly, **scope preserved** (never moves between scopes). No conflict guard on update (it only repoints an existing name).
- Service guard chain for mutations (`names_service.py:266-289`): writable (`WORKBOOK_READ_ONLY`) → snapshot (`SNAPSHOT_FAILED`) → approval (`APPROVAL_DENIED`).

### 4.3 Success / reject examples (match actual code)
- **SUCCESS** — create `Sales` workbook scope when no `Sales` exists anywhere → `Names.Add` on `wb.Names`.
- **REJECT duplicate** — create `Sales` workbook scope when `Sales` workbook scope already exists → `reason="duplicate"`.
- **REJECT scope_conflict** — create `Sales` worksheet scope on `Sheet1` when `Sales` exists at workbook scope (or on `Sheet2`) → `reason="scope_conflict"`; `existing=[{"scope":"workbook","sheet":null}]` (or `{"scope":"worksheet","sheet":"Sheet2"}`).
- **SUCCESS update** — update `Sales` (workbook) `refers_to` → `RefersTo` rewritten; `scope` stays `workbook` (cannot become worksheet).
- **REJECT ambiguous resolve** — if two entries somehow match (defensive) → `reason="ambiguous"`, no guess.

---

## 5. Confirmation semantics (`snapshot_restore`, `excel_recover`)

### 5.1 Current state (verified — no `confirm`)
- `snapshot_restore(snapshot_id, workbook_id=None)` `api/test_tools.py:122-143` → `snapshot_service.restore_snapshot`. Closes the current working copy **without saving** (discards unsaved changes) and opens a fresh working copy from the snapshot baseline (`snapshot_service.py:18-22,299-310`). **Does NOT undo external macro side effects** (file-only; `external_side_effects_not_undone` on every manifest).
- `excel_recover(reason="agent_directed")` `api/recovery_tools.py:46-69` → `session_service.recover_owned_process`. Terminates the wedged owned `EXCEL.EXE` via the held handle, **discards working copies only** (originals never modified), rebuilds a hidden instance (`session_service.py:726-877`, `what_lost` `:865`). Recovery is process-level; it does **not** roll back external side effects of any macro that ran.
- Both registered destructive: `_ANN_DESTRUCTIVE` includes `excel_recover`, `snapshot_restore` (`server.py:1795-1796`).

### 5.2 Design — explicit agent-intent `confirm` flag (proposal)
- Add `confirm: bool = False` to both tool signatures (agent intent, **not** human approval — distinct from `ApprovalPolicy`).
- **Missing/false → reject before any destructive action**, with a clear structured error. Code choice (Auditor): additive `ErrorCode.CONFIRMATION_REQUIRED` (category `validation`, not retryable, actions `["re-call with confirm=true after reviewing what is discarded"]`) is cleanest; reuse of `APPROVAL_DENIED` is semantically wrong (that is policy denial). Additive codes have precedent (`errors.py:43-67`).
- Place the check **before** `service.restore_snapshot` / `session.recover_owned_process` so no COM/termination occurs.
- **Descriptions must state exactly what is discarded vs preserved and must NOT imply external macro side effects are rolled back:**
  - `snapshot_restore`: "Discards the current working copy (unsaved changes lost). Preserves the original file and the immutable snapshot baseline. Does **not** undo external macro side effects (file-only restore). Requires `confirm=true`."
  - `excel_recover`: "Terminates the owned Excel process and discards all open working copies. Originals are never modified. Does **not** roll back external side effects of any macro that already ran. Requires `confirm=true`."
- Keep the truthful `RecoveryResult.what_lost` (`session_service.py:865,947`) and the snapshot manifest `external_side_effects_not_undone` as the machine-readable echo.

---

## 6. Description overhead inventory (92 `_DESC_*`; 37,125 chars total)

**Verified counts** (runtime `list_tools()` + source scan):
- Tools: **92**; total description chars: **37,125**.
- `_DESC_*` constants in `server.py`: **96** (92 tools + `_DESC_ENVIRONMENT` + 5 resource descs; some constants shared).
- Tools using the full 8-part template (Purpose/Use when/Do not use when/Preconditions/Parameters/Behavior/Result/Next actions): **27**.
- Marker frequencies: `Purpose:` 27, `Use when:` 26, `Preconditions` 26, `Behavior:` 26, `Result:` 27, `Next actions:` 26, `Do not use when:` 23, `Parameters:` 19, `Example:` 5.
- **Repeated boilerplate:**
  - Full guard sentence (133 chars) on **27** tools ≈ **3,591 chars**.
  - `writable session + snapshot + approval` ×26; `postcondition re-read` ×34; `full-access default; restricted` ×20; `Guarded mutation:` ×23.
  - `xlvba://glossary` referenced ×7 (jargon already externalized — good).

**Move to glossary / shared docs (safe, high value):**
- The uniform 133-char guard sentence → define once in `xlvba://glossary` (e.g. term **guarded mutation**) and replace in each description with a short token `"Guarded mutation (see xlvba://glossary)."`. Saves ≈ 27×~110 ≈ **~3,000 chars**. (Update `test_description_contract.py:110-137` which asserts the full sentence — the test must change to assert the token.)
- The clause `(full-access default; restricted requires approval)` (×20) → glossary term **approval gate**.
- R13/BG06/preflight/truth-signal explanations already in glossary (`server.py:2628-2649`); descriptions already reference the URI — keep.

**Must stay in-description (do NOT move):**
- Safety-critical truthful phrases enforced by `test_description_contract.py:55-59`: `vba_compile` "does NOT compile", `job_cancel` "intent-only; never interrupts", `macro_classify_background` "never grants run permission". Plus `snapshot_restore`/`excel_recover` "does not undo external macro side effects" (§5).
- Per-tool routing ("Do not use when → other tool") — the cross-references between `test_run_suite`/`test_suite_run`, `range_write_values`/`range_write_formulas`, etc.
- Tool-specific side effects and the one `Example:` per `_EXAMPLE_TOOLS` (R-D8).
- Tool-specific preconditions (e.g. `range_write_values` "does NOT snapshot — call snapshot_create first").

**Risk:** the description-contract lint (`test_description_contract.py`) hard-codes the full guard sentence and the `_CRITICAL`/`_TRUTH`/`_EXAMPLE_TOOLS` name lists; any rename (§1) or boilerplate move (§6) **must** update this test in lockstep, or it fails. This is the single highest-coupling test for this work package.

---

## 7. Unresolved questions for the Auditor / Validator

1. **Fallback error code (§2.3-4):** approve additive `ErrorCode.INTERNAL_ERROR` (recommended) vs reuse `COM_BUSY`/`INVALID_RANGE`? Validator must confirm the closed-vocabulary precedent allows an additive code and that `category_for`/`retryable_for`/`actions_for` are extended for it (the totality test `test_error_envelope.py:47` will then require the new member to be mapped).
2. **`to_dict` field rename (§2.3-2):** confirm renaming `code`→`errorCode` and `context`→`details` outright (no alias) is acceptable given "backward compat not required"; this touches `test_error_envelope.py:30-31,41-42` and any consumer reading `payload["code"]`.
3. **`TestAction.VBA_COMPILE` (§1.4):** confirm the tool rename must **not** touch the declarative-suite action enum `test_spec.py:73` / `test_runner.py:622` (different contract). If the Auditor wants the action renamed too, that is a separate, larger change (suite schema + docs + many tests).
4. **`macro_policy._EXECUTING_TOOLS` (§1.4):** update the stale illustrative set (`macro_policy.py:161-169`) or leave (documented non-wired)? Validator to decide if leaving stale names is acceptable.
5. **Name scope strictness (§4.2):** the server rejects cross-scope names that raw Excel would allow. Confirm this stricter behaviour is intended product semantics (not a bug to relax) before documenting examples.
6. **`snapshot_restore` `source_path` (§3):** confirm the restored `WorkbookInfo.source_path` should be the snapshot baseline path (not the pre-restore original). Validator to verify against `snapshot_service.restore_snapshot` actual value.
7. **Confirm-flag error code (§5.2):** additive `CONFIRMATION_REQUIRED` vs reuse existing; and whether `confirm` should also gate `excel_close(save_changes=true)` / `excel_save` (out of current scope but adjacent).
8. **Benchmark tasks (§1.2/§1.3/§7):** `.ai/tmp/benchmark_tasks.yaml` references `job_cancel` (471-477) and `test_run_suite` (569). Confirm the implementer updates these and adds tasks for the renamed tools, canonical test tool, structured errors, WorkbookInfo, name conflicts, confirm flags (per baseline §9).
9. **Description-contract test coupling (§6):** the lint hard-codes the full guard sentence and name lists; confirm the boilerplate move + renames update it together (single commit) to avoid a red suite mid-migration.