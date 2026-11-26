# Report 08d — U2 Validator Verdict (Independent)

**Work package:** U2 — Environment/policy discovery + structured error envelope
**Plan:** `.ai/audit/plans/plan_03 — Agent Usability Remediation.md` §6, §7 (incl. §7.3 + D-U10), §12.3, §14 U2, §4
**Validator role:** Independent (did not implement U2). VALIDATION ONLY — no code/test/config modified.
**Working tree:** accepted U1 + U3 + U2 changes.
**Date:** 2026-11-26

## Method

Read the contract sections, then the implementation (`environment_tools.py`, `errors.py`, `models.py`, `session_service.py`, `snapshot_service.py`, `server.py`, `workbook_adapter.py`), the two new test files, and the working-tree diff. Ran the unit suite and two empirical probes (tool inventory + per-code envelope mapping). Every verdict below cites `file:line`.

## Per-item verdict

| # | Item | Verdict |
|---|------|---------|
| 1 | `xlvba_environment` non-COM, side-effect-free, `readOnlyHint:true`, registered (92), `excel.available=false` when down | **APPROVE** |
| 2 | `build` always null; `target_build` labelled; `has_snapshot` reads in-memory gate (not store) | **APPROVE** |
| 3 | `classification_is_authorization=false`; `r13_preflight="veto"`; approval/snapshot/path/macro blocks truthful | **APPROVE** |
| 4 | `SessionStatus.policy` additive optional, populated at build, `extra="forbid"` respected | **APPROVE** |
| 5 | `to_dict` keeps 5 keys + adds 3 via pure functions; mapping matches §7.3 exactly; no new ErrorCode; SCHEMA_VERSION=1 | **APPROVE** (one plan-internal doc inconsistency noted, not an impl defect) |
| 6 | No SDK change; no error `structuredContent`; no enum/annotation change; no DTO slimming; no `save_as`/`job_list` | **APPROVE** |
| 7 | `pytest tests/unit` green; tests are real, not superficial | **APPROVE** (1127 passed) |
| 8 | Null `application_version`/`vbom_available` limitation not misleading, no acceptance violation | **APPROVE** |

## Evidence

### Item 1 — non-COM, side-effect-free, read-only, registered, truthful down-state — APPROVE

- Non-COM by construction: the handler reads only module globals and non-COM session methods (`environment_tools.py:210-234`). `_excel_info` calls `session.worker_status()` (`environment_tools.py:168`) → `SessionService.worker_status` returns `self._worker.get_status()` with no submit (`session_service.py:606-609`); `_workbooks` calls `session.get_status()` (`environment_tools.py:197`) which is explicitly non-COM ("never submits to the worker", `session_service.py:544-548`, body `session_service.py:555-584`). No worker start, no snapshot/grant/classification anywhere in the module (`environment_tools.py:1-17` docstring; no `submit`, no `create_snapshot`, no approval call).
- Side-effect-free: `set_environment_runtime`/`clear_environment_runtime` only assign module globals (`environment_tools.py:67-95`); the tool never mutates the published objects (`environment_tools.py:74-79`).
- `readOnlyHint:true`: registered with `annotations=ToolAnnotations(read_only_hint=True)` (`server.py:2375-2380`) and listed in `_ANN_READ_ONLY` (`server.py:1752`).
- Registered, tool count 92: empirical `mcp.list_tools()` → `COUNT 92`, `HAS_ENV True`, `ENV_READONLY True`.
- `excel.available=false` when worker down: `environment_tools.py:165-166` returns `available=False` when `session is None`; `environment_tools.py:169-172` requires `accepting and state == RUNNING`. Covered by `test_environment_tools.py:173-189`.

### Item 2 — honest build, labelled target, in-memory gate — APPROVE

- `build` always null: DTO field `build: None = Field(default=None, ...)` (`models.py:183`); both construction paths pass `build=None` (`environment_tools.py:166`, `:178`). Never fabricated.
- `target_build` labelled design constant: `_TARGET_BUILD = "16.0.14332"` with comment "Design-time … NOT a detected build" (`environment_tools.py:55-57`); field description "Labelled design-time constant only; not a detected build" (`models.py:184-187`).
- `has_snapshot` reads the in-memory gate, not the store:
  - `environment_tools.py:203` takes `has_snapshot` from `status.policy.has_snapshot`.
  - `session_service.py:1140-1144` `_build_policy` reads `self._snapshot_gate(workbook_id)`.
  - Gate published as `SnapshotService.has_snapshot` (`server.py:1293`; setter `session_service.py:208-217`).
  - `SnapshotService.has_snapshot` returns `workbook_id in self._latest_snapshot` (in-memory) — `snapshot_service.py:427-430`. The on-disk `SnapshotStore` is never consulted.
  - Proven by `test_environment_tools.py:238-250` (in-memory id → True; store-only id → False) and `:268-294` (SessionStatus.policy gate).

### Item 3 — macro/approval/snapshot/path policy truthful — APPROVE

- `classification_is_authorization=False` and `r13_preflight="veto"` set in `_policy` (`environment_tools.py:148-152`) and enforced by `Literal` types (`models.py:240-241`, `models.py:127-128`).
- Approval block derived from `AccessMode.RESTRICTED` (`environment_tools.py:130-137`); snapshot block `required_for_mutations=True`, `created_by="agent:snapshot_create"` (`environment_tools.py:138-141`); path block from `PathPolicy.roots`/`allow_unc` (`environment_tools.py:142-147`). All match §6.1.
- Asserted by `test_environment_tools.py:188-189` (down) and `:209-212` (restricted → grants required).

### Item 4 — SessionStatus.policy additive optional, populated, strict — APPROVE

- Field added to the class (not ad-hoc): `policy: SessionPolicy | None = None` (`models.py:144`); `SessionPolicy` is a `DomainModel` (`models.py:102-128`) and `DomainModel` is `extra="forbid"` (`models.py:71`). Diff confirms `+class SessionPolicy` and `+policy: SessionPolicy | None = None` are additive.
- Populated where `SessionStatus` is built: `_build_status` passes `policy=self._build_policy(workbook_id)` (`session_service.py:1123`), so every status path (open/close/save/recover/get_status) carries it.
- Covered by `test_environment_tools.py:268-294`.

### Item 5 — additive error envelope, §7.3 mapping exact — APPROVE

- `to_dict` keeps the original five keys and adds three (`errors.py:335-353`): `code`, `message`, `hresult`, `hresult_hex`, `context` + `errorCategory`, `isRetryable`, `suggestedNextActions`. Pure functions `category_for`/`retryable_for`/`actions_for` (`errors.py:249-299`), no COM, no schema bump.
- `isRetryable=true` ONLY for `COM_BUSY`/`RPC_E_CALL_REJECTED`: `_RETRYABLE_CODES = frozenset({COM_BUSY, RPC_E_CALL_REJECTED})` (`errors.py:162-164`); `retryable_for` returns membership (`errors.py:267-281`). Empirical: `RETRYABLE_TRUE ['COM_BUSY', 'RPC_E_CALL_REJECTED']`.
- `STALE_TARGET`/`REVISION_CONFLICT` false + caveat: `errors.py:196-203` ("the prior mutation MAY have partially applied — re-read state; do NOT blind-replay"); `retryable_for` false (not in retryable set). Empirical: `STALE False [...partially applied... do NOT blind-replay]`, `REV False`.
- Job `interrupted_unknown` false: `_JOB_CATEGORY`/`retryable_for` short-circuit to False (`errors.py:168-171`, `:274-276`); actions `["never resubmit; report status"]` (`errors.py:177`). Empirical: `INT False ['never resubmit; report status']`.
- `SNAPSHOT_FAILED` → `missing_prerequisite` + `snapshot_create` action (`errors.py:129`, `:185`). Empirical: `SNAP ['snapshot_create(workbook_id)', 'retry mutation']`.
- `APPROVAL_DENIED` → `permission_denied` + full-mode/grant action (`errors.py:127`, `:187`). Empirical: `APPROVAL ['open in full mode or supply grant/allowlist entry', 'retry']`.
- No new `ErrorCode` members: diff of `errors.py` shows no added/removed enum member (the `+208` lines are the envelope block, `_CATEGORY_BY_CODE`, `_RETRYABLE_CODES`, `_JOB_*`, `_ACTIONS_BY_CODE`, the three pure functions and the `to_dict` body). Membership `errors.py:19-67` unchanged.
- `SCHEMA_VERSION = 1` (`models.py:65`); `EnvironmentReport.schema_version` defaults to it (`models.py:274`) and the handler passes `schema_version=1` (`environment_tools.py:227`).
- Tests `test_error_envelope.py:47-118` are table-driven over `list(ErrorCode)` (`:47-62`) plus explicit D-U10 cases (`:68-106`).

**Advisory (not a defect):** Plan §7.2 lists `INVALID_RANGE` under `unsupported_feature` (plan line 222) while §7.3 and §7.1 classify it as `validation` (plan lines 228, 246). The implementation follows §7.3/§7.1 → `validation` (`errors.py:158`, empirical `INVALID_RANGE_CAT validation`). This is an internal plan-table inconsistency, not an implementation deviation; the chosen category is the one §7.3 (the per-code authority) and the `to_dict` example both use.

### Item 6 — no SDK change, no structuredContent, no enum/annotation/DTO-slimming, no save_as/job_list — APPROVE

- No `structuredContent`/`structured_content`/`structured_output`/`defer_loading` anywhere in `src` (grep: no files). Error path is text-only JSON: `raise ToolError(json.dumps(cause.to_dict(), ...))` (`server.py:2502`), matching §7.1.
- No `save_as`/`xlvba_job_list` tool: empirical inventory has neither (`False False`); only the pre-existing `excel_save` base tool exists. U5 is deferred (plan §12.5/§14 U5).
- No U2 enum/annotation change: the `statuses.py` diff adds `ConditionalFormatRequestType`, `PivotFieldLayoutOrientation`, and prunes `PivotFilterRequestKind` to `MANUAL` — these are U3 scope (schema/enum consistency), consumed read-only by `_capabilities()` (`environment_tools.py:122-123`). `_DESC_ENVIRONMENT` is the only description U2 adds (`server.py:1004-1019`), as scoped by §14 U2 non-scope. No DTO slimming (U4 deferred).

### Item 7 — tests green and substantive — APPROVE

- `& "C:\py_exp\xlvba_mcp\.venv\Scripts\python.exe" -m pytest tests/unit -q` → **1127 passed in 27.53s**.
- `test_environment_tools.py` is real: strict-DTO rejection (`:134-167`), worker-down truthfulness (`:173-189`), worker-up capabilities + gate propagation (`:195-218`), full-mode no-grant (`:221-233`), **in-memory-vs-store gate** (`:238-250`), SessionStatus.policy gate + restricted grants (`:268-294`). The in-memory-vs-store test constructs a `SnapshotService` with `_latest_snapshot={"w-inmem"}` and a fake store listing only `"w-store-only"`, asserting True/False respectively (`:238-250`) — exactly the §14 U2 required assertion.
- `test_error_envelope.py` is real: additive-keys + JSON-serializability (`:34-44`), totality over every `ErrorCode` (`:47-62`), D-U10 `STALE_TARGET`/`REVISION_CONFLICT` false + caveat (`:68-80`), only-COM-busy retryable (`:83-86`), `SNAPSHOT_FAILED` action (`:89-93`), `APPROVAL_DENIED` category (`:96-99`), job `interrupted_unknown`/`expired` (`:102-112`), unknown-string fallback (`:115-118`).

### Item 8 — null application_version / vbom_available limitation — APPROVE (not misleading)

- `application_version` is a non-COM best-effort value taken from the first open workbook's cached `Application.Version` (captured at open by `workbook_adapter.py:100-106`), or `None` when none is open (`environment_tools.py:173-174`). `vbom_available` is always `None` at server level because no workbook model is built (`environment_tools.py:181`). Both are documented honestly in the DTO (`models.py:176-189`).
- No acceptance criterion requires these to be non-null. §14 U2 acceptance (plan line 446) lists `build:null`, `target_build` labelled, `has_snapshot` gate, `classification_is_authorization=false`, `excel.available=false` when down — all met. Plan §6.1 explicitly allows `vbom_available` "from model vba_access, **else null**" (plan line 153); the server has no model, so `null` is the correct branch. `application_version` provenance is `workbook_adapter.py` (`Application.Version`) per plan line 178 — satisfied via the cached workbook value.
- **Minor honesty nuance (advisory, not blocking):** when a workbook is open but its `Application.Version` read failed at open time, `WorkbookInfo.excel_version` holds the sentinel `"unknown"` (`workbook_adapter.py:32`, `:103-105`), so `application_version` could surface `"unknown"` rather than `null`. This is truthful (an explicit unknown, never a fabricated version) and does not violate any criterion; recorded only for completeness.

## Misleading-field review

No misleading field found. `build` is structurally `None` and cannot be fabricated (`models.py:183`). `target_build` is labelled in both code and field description. `classification_is_authorization`/`r13_preflight` are `Literal`-pinned to the safe values. `has_snapshot` is sourced from the in-memory gate. The only nuance is the `"unknown"` sentinel possibility for `application_version` (advisory above), which is explicit, not fabricated.

## Scope concerns

- U2 stayed within scope: new `environment_tools.py`; additive `EnvironmentReport`/`SessionPolicy`/`SessionStatus.policy`; `session_service` policy population + `set_snapshot_gate`; `errors.py` envelope + pure functions; `server.py` registration + `_DESC_ENVIRONMENT` + lifespan wiring (`server.py:1289-1299`, `:1351`). `snapshot_service.has_snapshot` is exposed unchanged (`snapshot_service.py:427-430`).
- The `statuses.py` enum pruning in the same working tree is U3, not U2; U2 only *reads* the pruned `PivotFilterRequestKind`/`PivotSourceType` for capabilities. No cross-contamination.
- Non-negotiables (§4) intact: single hidden STA worker untouched (U2 adds no COM path, §4.1); `SCHEMA_VERSION=1` additive (§4.7); fail-closed honest build (§4.8); no SDK/structuredContent (§4.10); truthful annotations (§4.11); conservative per-code retry (§4.12).

## Final U2 verdict

**ACCEPT.**

All eight verification items pass with cited evidence. `xlvba_environment` is non-COM, side-effect-free, `readOnlyHint:true`, registered (tool count 92), and reports `excel.available=false` when the worker is down. `build` is always null with a labelled `target_build`; `has_snapshot` reads the in-memory `SnapshotService.has_snapshot` gate, never the store. The policy blocks are truthful and `Literal`-pinned. `SessionStatus.policy` is additive/optional, populated at build, under `extra="forbid"`. `to_dict` is additive (5 + 3 keys) via pure functions, and the per-code mapping matches §7.3 exactly (only `COM_BUSY`/`RPC_E_CALL_REJECTED` retryable; `STALE_TARGET`/`REVISION_CONFLICT` false with the partial-apply caveat; `interrupted_unknown` false; `SNAPSHOT_FAILED`→`snapshot_create`; `APPROVAL_DENIED`→full-mode/grant), with no new `ErrorCode` members and `SCHEMA_VERSION=1`. No SDK change, no error `structuredContent`, no enum/annotation/DTO-slimming, no `save_as`/`job_list`. Unit suite green (1127 passed) with substantive coverage.

**Blocking reasons:** none.

**Advisory (non-blocking):** (1) Reconcile the plan §7.2 `INVALID_RANGE` category (listed under `unsupported_feature`) with §7.3/implementation (`validation`) in a future doc edit — implementation is correct per §7.3. (2) `application_version` may surface the `"unknown"` sentinel when a workbook's version read failed at open; acceptable and truthful, but if a strict "version-or-null" contract is later desired, normalize the sentinel to `None` at the DTO boundary.