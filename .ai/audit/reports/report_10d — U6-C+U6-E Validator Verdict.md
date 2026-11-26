# Report 10d — U6-C + U6-E Validator Verdict

- **Role:** Independent Validator (did **not** implement U6-C / U6-E).
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` (working tree contains U6-A + U6-B + U6-C + U6-E).
- **Contract:** `.ai/audit/plans/plan_05 — Remaining Usability & Save As Implementation.md` §5.3 (routing), §5.4 (confirm), §8 U6-C, §8 U6-E, §4 N5/N6.
- **Method:** read source at working tree; inspected handler order; ran `tests/unit` (1137 passed, exit 0); confirmed `list_tools()` count and tool membership programmatically.
- **Verdict:** **U6-C APPROVE**, **U6-E APPROVE** (one advisory wording note, non-blocking).

---

## U6-C — Capability-based routing guidance

### Item 1 — `_DESC_ENVIRONMENT` routing section — **APPROVE**

Evidence (`src/excel_vba_mcp/server.py:993-1025`):

| Requirement | Status | Evidence |
|---|---|---|
| Capability-based routing section present | PASS | `server.py:1008-1024` — dedicated "Capabilities (this server provides)" + "Routing is capability-based advice only" block |
| States what THIS server provides | PASS | `server.py:1008-1011` — "VBA project access, macros, pivot tables (database-source; manual filters), background jobs, recovery, and non-activating compile diagnostics — NOT compilation" |
| "prefer xlvba / use excel-mcp" matrix + examples | PASS | `server.py:1014-1019` ("Prefer xlvba when…", "Use excel-mcp when…"); matrix `server.py:1020-1022`; examples `server.py:1023-1024` |
| Explicitly states it does NOT detect other servers | PASS | `server.py:1012-1013` — "this server does NOT detect whether any other MCP server is installed or connected." |
| No `excel-mcp_*` presence claim | PASS | `grep excel-mcp_\|excel_mcp_` over `server.py` → **0 hits**. "excel-mcp" appears only as conditional routing advice ("Use excel-mcp when it exposes an equivalent supported operation…", `:1018`), never as a presence/absence assertion. |
| Glossary not duplicated | PASS | No glossary term is re-defined in `_DESC_ENVIRONMENT`; jargon is referenced by capability name only. `_GLOSSARY_MD` (`server.py:2668`) is untouched by U6-C. |
| Existing "Do not use when" routing kept | PASS | `server.py:999-1000` retained (worker/session liveness → excel_health / server_health). |

N6 compliance: guidance is phrased "this server provides capability X" and makes no claim about any other server's presence — **compliant**.

**Advisory (non-blocking):** `server.py:1010` appends "(sourced from EnvironmentCapabilities)" to the phrase "non-activating compile diagnostics". The literal flag in `EnvironmentCapabilities` is `non_activating_compile=False` (`api/environment_tools.py:121`). The wording is internally consistent — the server provides *diagnostics* (`vba_compile_diagnostics`, which returns `indeterminate`), **not** compilation, and the `False` flag encodes exactly that "compilation is not available" — and it matches the plan §5.3 example verbatim. The parenthetical "sourced from EnvironmentCapabilities" is slightly loose (the capability *list* is not a 1:1 field copy), but it is not misleading and does not assert a false capability. No change required.

**U6-C verdict: APPROVE.**

---

## U6-E — Semantics docs + confirm flags

### Item 2 — `confirm` rejects BEFORE any delegate call — **APPROVE**

`snapshot_restore` (`src/excel_vba_mcp/api/test_tools.py:119-165`):
- Signature `*, confirm: bool = False` (`:122-123`).
- Guard `if not confirm: raise ExcelVbaMcpError(..., code=ErrorCode.CONFIRMATION_REQUIRED, ...)` at `:150-163`.
- Delegate `get_snapshot_service()` + `service.restore_snapshot(...)` only at `:164-165` — **after** the guard. On reject the service is never fetched/invoked. **Reject-before-action confirmed.**

`excel_recover` (`src/excel_vba_mcp/api/recovery_tools.py:47-97`):
- Signature `*, confirm: bool = False` (`:49-50`).
- Guard `if not confirm: raise ... CONFIRMATION_REQUIRED` at `:81-94`.
- Delegate `get_session_service()` + `session.recover_owned_process(...)` only at `:95-96` — **after** the guard. **Reject-before-action confirmed.**

`RecoveryResult.confirmed` unchanged as an **output echo**: field still present (`recovery_tools.py:37`), populated from the service outcome dict (`:97` `RecoveryResult.model_validate(outcome)`); it is **not** consulted as an input gate. **Compliant with plan §5.4 / N5.**

Both tools remain in `_ANN_DESTRUCTIVE` (`server.py:1801-1802`).

### Item 3 — Descriptions: discard vs preserve + external-side-effect caveat + "Requires confirm=true" — **APPROVE**

`_DESC_SNAPSHOT_RESTORE` (`server.py:298-309`):
- Discard: "unsaved changes are discarded" (`:302-303`).
- Preserve: snapshot baseline + untouched original (implied by "opens the snapshot baseline"; original never modified per model).
- External-side-effect caveat: "does NOT undo external macro side effects (database writes, emails, network calls, files, printing)" (`:303-304`).
- "Requires `confirm=true`; a missing/false confirm is rejected with CONFIRMATION_REQUIRED before the working copy is closed" (`:305-306`).
- No implication of external rollback — explicitly negated.

`_DESC_EXCEL_RECOVER` (`server.py:274-292`):
- Discard: "Discards the open working copies (originals are never modified)" (`:285`).
- External-side-effect caveat: "does not roll back external side effects of any macro that already ran" (`:285-286`).
- "Requires `confirm=true`; a missing/false confirm is rejected with CONFIRMATION_REQUIRED before any termination" (`:287-288`).
- No implication of external rollback — explicitly negated.

Both descriptions state discard vs preserve, carry the external-side-effect caveat, and state "Requires `confirm=true`". **No misleading text; no reject-after-action bug.**

### Item 4 — WorkbookInfo docstring + docs/05 two-path & name-scope — **APPROVE**

`WorkbookInfo` docstring (`src/excel_vba_mcp/domain/models.py:86-114`):
- `path` documented as the **working copy** actually opened (`:91-94`).
- `source_path` documented as the **original, allowlisted** path, **nullable** (`:95-98`); field `source_path: str | None = Field(default=None, ...)` (`:125-132`).
- `original_name` documented as **intentionally absent** (redundant with `name`) (`:100-102`, and field description `:129-130`).
- **No field added**: `grep original_name` over `models.py` → only doc/comment hits (`:100`, `:130`); no `original_name` attribute. `SCHEMA_VERSION = 1` (`models.py:65`) unchanged. No `SaveAsResult` present (U7 not landed).

docs/05 two-path table (`docs/05-mcp-tools-reference.md:503-522`): `path` vs `source_path` vs `name` table with new-open / writable / post-edit scenarios; `original_name` documented absent (`:513-514`).

docs/05 name-scope rules (`docs/05-mcp-tools-reference.md:383-419`) vs `names_adapter.py`:
| Rule | docs/05 | names_adapter.py | Match |
|---|---|---|---|
| Duplicate (same scope) → `reason="duplicate"` | `:388-389` | `_guard_no_conflict` `:592-606` | PASS |
| Cross-scope / other-sheet → `reason="scope_conflict"`, stricter than Excel | `:390-392` | `:607-618` | PASS |
| Identical name in different worksheet scopes → conflict | `:390` ("or on another sheet") | `same_scope` includes `(entry_sheet==sheet)` `:592-595`; cross-sheet falls to `existing` `:607` | PASS |
| Update preserves scope (never re-scopes) | `:408-411` | `update_name` `:439-470` (writes `RefersTo` only) | PASS |
| Resolution order, ambiguous → `reason="ambiguous"` | `:413-419` | `_resolve_name` `:535-574` | PASS |
| Create on conflict rejected before add | `:385-406` | `create_name` calls `_guard_no_conflict` `:422` before `collection.Add` `:426` | PASS |

docs/09 also consistent (`docs/09-limitations-and-compatibility.md:232-249`): two-path, no `original_name`, confirm-before-termination, snapshots restore files not external effects.

### Item 5 — `tests/unit/test_confirm_flags.py` proves reject-before-action — **APPROVE**

`tests/unit/test_confirm_flags.py`:
- `snapshot_restore` missing confirm (`:86-98`) and false confirm (`:101-111`): assert `CONFIRMATION_REQUIRED`, `context["tool"]=="snapshot_restore"`, `"discard" in context`, and **`fake.restore_calls == []`** (delegate not called).
- `excel_recover` missing confirm (`:129-141`) and false confirm (`:144-154`): assert `CONFIRMATION_REQUIRED`, `context["tool"]=="excel_recover"`, `"discard" in context`, and **`fake.recover_calls == []`**.
- Positive paths (`:114-123`, `:157-167`) confirm delegation occurs only with `confirm=True`.
- Fakes installed via `set_snapshot_service` / `set_session_service`; no COM. Tests genuinely prove reject-before-action for both tools.

### Item 6 — `list_tools()` == 91, no Save As — **APPROVE**

Programmatic check (`asyncio.run(server.mcp.list_tools())`):
- `COUNT 91`.
- `has_save_as False` (no `excel_save_as`).
- `has_vba_compile_diagnostics True`, `has_job_request_cancel True`, `has_test_run_suite False` (U6-A landed).
- `test_description_contract.py:83-86` asserts `len(tools) == 91`.
- `_DESC_SAVE` (`server.py:251-252`) still says "there is no save-as" — correct, U7 not landed.

### Item 7 — `pytest tests/unit -q` green — **APPROVE**

`& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" -m pytest tests/unit -q` → **`1137 passed in 40.00s`, exit code 0**. No failures/errors.

---

## Rollout Analysis

- U6-C is description-only on `server.py` `_DESC_ENVIRONMENT`; no tool/capability logic change; independently revertible. No dependency risk.
- U6-E-C (confirm) depends on U6-B `CONFIRMATION_REQUIRED` (present, `domain/errors.py`) — dependency satisfied. Additive `confirm=False` default preserves backward behavior for any caller that omits it (but such a caller now correctly gets `CONFIRMATION_REQUIRED`, which is the intended safety gate).
- U6-E-A/B are doc-only; revert independently.
- Count coupling (91) is consistent with U6-A landed and U7 not yet landed. No Save As surface present.
- No schema bump (`SCHEMA_VERSION=1`), no `SaveAsResult`, no `original_name` field — additive-only discipline held.

## Execution Validation

- All change targets exist and are non-stale: `_DESC_ENVIRONMENT` (`server.py:993`), `_DESC_SNAPSHOT_RESTORE` (`:298`), `_DESC_EXCEL_RECOVER` (`:274`), `snapshot_restore` (`test_tools.py:119`), `excel_recover` (`recovery_tools.py:47`), `WorkbookInfo` (`models.py:86`), `names_adapter.py` `_guard_no_conflict`/`update_name`/`_resolve_name`.
- Semantic anchors (functions/handlers/DTO docstrings) used — no fragile line-only targeting in the implementation.
- Architecture consistent: handlers stay thin, guards precede delegation, docs match adapter behavior.

## Warnings

- **(Advisory, non-blocking)** `_DESC_ENVIRONMENT:1010` "(sourced from EnvironmentCapabilities)" is a loose attribution for the "non-activating compile diagnostics" phrase (the `non_activating_compile` flag is `False`). Wording is not misleading and matches plan §5.3 verbatim; no action required.
- No reject-after-action bug found. No misleading discard/preserve text found. No external-rollback implication found.

## Required Fixes

- **None.**

## Advisory Recommendations

- Optional: tighten `_DESC_ENVIRONMENT:1010` parenthetical to avoid implying a 1:1 field copy from `EnvironmentCapabilities` (e.g. drop "sourced from EnvironmentCapabilities" or rephrase to "see `xlvba_environment` capabilities"). Purely cosmetic; not required for acceptance.

---

## Final Verdict

- **U6-C: APPROVE** — capability-based routing present, this-server-only, prefer-xlvba/excel-mcp matrix + examples, explicit no-other-server-detection, no `excel-mcp_*` presence claim, glossary not duplicated.
- **U6-E: APPROVE** — confirm gates reject before any delegate call (proven by `test_confirm_flags.py`), `RecoveryResult.confirmed` unchanged as output echo, descriptions state discard vs preserve + external-side-effect caveat + "Requires confirm=true" with no rollback implication, `WorkbookInfo` two-path docstring with `original_name` intentionally absent and no field added, docs/05 + docs/09 name-scope and two-path tables match `names_adapter.py`, `list_tools()==91`, no Save As, `tests/unit` green (1137 passed).