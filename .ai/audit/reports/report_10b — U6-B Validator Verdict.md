# Report 10b — U6-B Validator Verdict (Guaranteed Structured Error Envelope)

- **Role:** Independent Validator (did **not** implement U6-B). **Validation only — no source/test/config change made.**
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d` (U6-B changes present in the working tree, uncommitted).
- **Contract:** `.ai/audit/plans/plan_05 — Remaining Usability & Save As Implementation.md` §5.2, §4 N1/N2/N7, §8 U6-B.
- **Method:** read source at working tree; read SDK at `.venv` (mcp 2.3.0); `git diff` scope review; ran `tests/unit` + the two focused files.
- **Verdict up front:** **ACCEPT** (all 6 items APPROVE). One non-blocking rollout WARNING (stale old-key consumers in `tests/integration_excel`, out of U6-B's declared scope).

---

## Per-item verdicts

### Item 1 — `domain/errors.py` envelope + new codes — **APPROVE**

- `ErrorCode.INTERNAL_ERROR` + `ErrorCode.CONFIRMATION_REQUIRED` added: `errors.py:74-75` (additive block, comment `:68-73`). Diff confirms pure addition; no Stage-1/Phase-2 member removed or reordered.
- `_CATEGORY_BY_CODE` extended: `CONFIRMATION_REQUIRED → "validation"` `errors.py:167`; `INTERNAL_ERROR → "internal"` `errors.py:169`.
- `_ACTIONS_BY_CODE` extended: `INTERNAL_ERROR → ["inspect server logs; the failure was not a domain error"]` `errors.py:247-249`; `CONFIRMATION_REQUIRED → ["re-call with confirm=true after reviewing what is discarded"]` `errors.py:250-252`. Text matches §5.2 verbatim.
- `_RETRYABLE_CODES` **unchanged**: `errors.py:173-175` still `{COM_BUSY, RPC_E_CALL_REJECTED}`; neither new code is retryable (confirmed by `retryable_for` totality test + explicit per-code tests).
- `to_dict` key order: `errorCode, message, hresult, hresult_hex, details, errorCategory, isRetryable, suggestedNextActions` — `errors.py:364-373`. Mapping `errorCode ← code.value` `:365`, `details ← context` `:369`; `hresult`/`hresult_hex` additive `:367-368`. Exact order asserted by `test_error_envelope.py:44-53`.
- Note (non-issue): `_ERROR_BY_CODE` (`errors.py:649-686`) is **not** extended with the two new codes. This is correct — that map only serves `error_from_hresult` (HRESULT→typed subclass); neither new code has an HRESULT or a dedicated subclass, and no test requires completeness there. `category_for`/`retryable_for`/`actions_for` remain total via `_CATEGORY_BY_CODE.get(..., _DEFAULT_CATEGORY)`.

### Item 2 — `server.py` hatch fallback (`~:2500`) — **APPROVE**

- **No bare re-raise remains.** The former bare `raise` is replaced by the synthesized fallback (`server.py:2521-2541`); the only other raise is the structured domain/validation raise `server.py:2542`. `git diff` confirms `- raise` → fallback block.
- Non-domain path builds `ExcelVbaMcpError(str(origin), code=INTERNAL_ERROR, context={"reason":"unhandled_exception","exception": origin.__class__.__name__})` `server.py:2531-2538` and `raise ToolError(json.dumps(fallback.to_dict(), ensure_ascii=False)) from exc` `:2539-2541`.
  - `errorCode=INTERNAL_ERROR` ✓; all 6 keys via `to_dict()` ✓; `details.exception` = class name `:2536` ✓.
  - `message = str(origin)` where `origin = _originating_cause(exc)` `:2530` — the **deepest** `__cause__`, i.e. the real failure, **not** the SDK wrapper's `"Error executing tool <name>"`. `_originating_cause` is cycle-safe (seen-id set) `server.py:2401-2417`.
  - No live exception embedded: context holds only two strings (`reason`, `exception` class name); `message` is `str(...)`. Cycle-safe.
- Verified against SDK: `UnexpectedToolError(ToolError)` (`.venv/.../mcp/server/mcpserver/exceptions.py:61`) so `except ToolError` catches it; `MCPServer.call_tool` raises (does not return isError) `server.py:533-548`; `_handle_call_tool` turns a caught `ToolError` into `CallToolResult(is_error=True, text=str(exc))` `server.py:445-456`. Therefore the hatch's JSON `ToolError` is what the client receives as isError text.
- Test proof: `test_mcp_error_surfacing.py:75-102` asserts `message == "boom"`, `details.exception == "RuntimeError"`, `errorCode == "INTERNAL_ERROR"`, no `UnexpectedToolError`, no `"Error executing tool"`.
- The plan §5.2 shorthand `message=<str(exc) only>` is superseded by the binding task brief + N1 ("real originating cause, NOT 'Error executing tool'"); using `str(origin)` is the correct, stricter reading. **Not a deviation.**

### Item 3 — every server isError path yields the 6 keys — **APPROVE**

Three and only three hatch branches, all funnel through `to_dict()`:
1. Domain cause: `_find_domain_cause` `server.py:2383-2398` → `raise ToolError(json.dumps(cause.to_dict())) from cause` `:2542`.
2. Validation cause: `_validation_error_as_domain` `server.py:2451-2480` (reuses `INVALID_RANGE`, `_jsonable_validation_errors` sanitizer `:2420-2448`) → same `:2542` raise.
3. Non-domain fallback: INTERNAL_ERROR envelope `:2531-2541`.

No branch returns a bare string or partial dict (grep for `CallToolResult(`/`is_error`/`isError` in `src` finds no tool-side isError producer; the only `raise ToolError` are `:2539` and `:2542`, both JSON). Resource handlers `server.py:2646,2662` also emit `exc.to_dict()` JSON (consistent, not a tool isError path). `MCPError` raised by a tool is re-raised by the SDK as a JSON-RPC **protocol** error (`server.py:443-444`), not an `is_error` tool result — out of N1's isError scope (noted, not a leak).

### Item 4 — tests updated in the SAME change — **APPROVE**

- `test_error_envelope.py`: `_ORIGINAL_KEYS` renamed to `{errorCode, message, hresult, hresult_hex, details}` `:30`; `_NEW_KEYS` `:31`; fixed order assertion `:44-53`; `errorCode`/`details` asserts `:41-42`; **new-code** tests `test_internal_error_is_internal_not_retryable_with_action` `:113-120` and `test_confirmation_required_is_validation_not_retryable_with_action` `:123-130`; totality parametrize `list(ErrorCode)` `:58` now covers 31 codes.
- `test_mcp_error_surfacing.py`: leak test **rewritten** to `test_non_domain_error_surfaced_as_structured_internal_error` `:75-102` (structured INTERNAL_ERROR, 6 keys, no `UnexpectedToolError`, no `"Error executing tool"`); domain test key rename `:71-72`; model-validator test `:132-133`.
- Lockstep consumer edits (allowed by item 4) are pure renames: `test_domain_models.py:217,219`, `test_errors.py:77,81,88,171-172,187-188`, `test_phase2_contracts.py:544,623` (`code→errorCode`, `context→details`). No logic change.
- All U6-B edits are in one uncommitted working-tree change (`git status`: `errors.py`, `server.py`, `test_error_envelope.py`, `test_mcp_error_surfacing.py` + 3 consumer test files). Satisfies N7 (single change).

### Item 5 — no SDK change / no structuredContent / SCHEMA_VERSION=1 / no extra codes — **APPROVE**

- SDK untouched: `git status` shows no `.venv`/mcp change (and `.venv` is not tracked).
- No error `structuredContent`: grep `structuredContent|structured_content` in `src` → 0 hits.
- `SCHEMA_VERSION = 1`: `domain/models.py:65` (unchanged; not in diff).
- No extra `ErrorCode` members: `len(list(ErrorCode)) == 31` (29 baseline + 2), both new names present, no others added (runtime check).

### Item 6 — tests green — **APPROVE**

- Focused: `pytest tests/unit/test_error_envelope.py tests/unit/test_mcp_error_surfacing.py` → **48 passed**.
- Full unit: `pytest tests/unit` → **1131 passed**, exit 0 (`~27s`).

---

## Paths that can still leak a bare error

- **None within U6-B's scope.** Every tool `is_error` path (domain / validation / non-domain) emits the full JSON envelope; no bare `"Error executing tool"` reaches a client (proven by `test_mcp_error_surfacing.py` and SDK trace).
- Theoretical residual (not a current leak): if a future SDK emitted `UnexpectedToolError` with **no** `__cause__`, `_originating_cause` would return the wrapper and `message` would be `"Error executing tool <name>"`. Current SDK 2.3.0 always sets `__cause__` to the original (`.venv/.../exceptions.py:68`), so this cannot occur today.
- `MCPError` from a tool → JSON-RPC protocol error (not `is_error`), by design outside N1.

---

## Warnings (non-blocking)

- **W1 — Stale old-key consumers in `tests/integration_excel` (rollout/maintainability).** The breaking `code→errorCode` / `context→details` rename was **not** propagated to Excel-requiring integration tests that parse the envelope:
  - `tests/integration_excel/test_mcp_error_surfacing_matrix.py:523,525,527,532` (`payload["code"]`).
  - `tests/integration_excel/test_mcp_full_access_workflow.py:311` (`error_code()` helper, `payload["code"]`).
  - `tests/integration_excel/test_vba_project_write.py:556-557,612,686,751-754` (`err.to_dict()["code"]`, `["context"]`, `diagnostics["code"]/["context"]`).
  These will fail when the integration suite is run against the U6-B envelope. This is **out of U6-B's declared changed-files** (plan §8 U6-B lists only the two unit test files) and outside the unit gate, so it does **not** block U6-B acceptance; but it must be fixed (in U6-A, which already touches `test_mcp_error_surfacing_matrix.py`, or a dedicated follow-up) before integration sign-off. Flagging per N7's intent (envelope consistency).
  - Note: `tests/unit/test_pivot_refresh.py:569,582,600,615` `result.error["code"]` is a **different** surface (`PivotRefreshResult.error` DTO field, not `to_dict()`); unaffected — unit suite green confirms.
- **W2 — Sequencing.** U6-B is foundational; U6-E (`CONFIRMATION_REQUIRED`) and U7 (envelope) depend on it. Do not revert U6-B in isolation after they land (plan §11). No action for this verdict.

---

## Required fixes

- None blocking U6-B.

## Advisory recommendations

- Update the W1 integration_excel consumers to `errorCode`/`details` in the same commit that next touches them (U6-A) to keep the integration suite runnable.
- Optional: add a unit-level guard that asserts the INTERNAL_ERROR `message` is never the literal `"Error executing tool"` substring for a wrapped non-domain cause (defensive against a future SDK that omits `__cause__`). Current tests already assert the positive case (`message == "boom"`).

---

## Final U6-B verdict: **ACCEPT**

All six verification items APPROVE with cited evidence. The guaranteed structured envelope is real: `to_dict` emits the 8-key (6 required + additive hresult/hresult_hex) envelope in the fixed order; the `call_tool` hatch synthesizes a cycle-safe `INTERNAL_ERROR` fallback from the originating cause with no bare re-raise; every isError path is structured; the two new codes are additive and non-retryable; `SCHEMA_VERSION` stays 1; no SDK/`structuredContent` change; unit suite (1131) and focused files (48) green. The only follow-up is the out-of-scope integration_excel key staleness (W1), which does not affect the unit-scoped U6-B contract.