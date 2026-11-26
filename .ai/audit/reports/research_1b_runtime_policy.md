# Research 1B — Runtime Policy & State Transparency

**Repo:** `C:\py_exp\xlvba_mcp` · **Git HEAD:** `e2289ca` · **Role:** Researcher 1B (research only, no code changed)
**Scope:** What an agent can and cannot discover about the runtime environment/policy today, and a proposed agent-readable **environment/policy contract** + **structured error contract**.

Confidence tags: **HIGH** = directly read in source; **MEDIUM** = inferred from source with a small gap; **LOW** = assumption flagged as open question.

---

## 1. Current policy & state model (evidence-based)

### 1.1 Access modes — FULL vs RESTRICTED

- `AccessMode` is a closed `StrEnum` with exactly `FULL="full"` and `RESTRICTED="restricted"` (`src/excel_vba_mcp/domain/statuses.py:54-65`).
- Config default is **FULL** (`src/excel_vba_mcp/config.py:49-53`), loaded from `EXCEL_VBA_MCP_ACCESS_MODE` (`config.py:39`, `config.py:102-104`); unset → FULL.
- The single production `ApprovalPolicy` is constructed with `mode=config.server.access_mode` (`src/excel_vba_mcp/server.py:861`).
- Semantics in `ApprovalPolicy.check` (`src/excel_vba_mcp/security/approval.py:189-264`):
  - `READ_ONLY` → auto-allowed, `source=None` (`approval.py:213-219`).
  - `FULL` → `MUTATING`/`EXECUTING` allowed by policy, `source=POLICY`, no grant (`approval.py:224-230`).
  - `RESTRICTED` → allowlist match or live grant required; otherwise **deny-by-default** (`approval.py:232-264`).
- Neither mode bypasses PathPolicy / Office-trust / snapshot / audit (`statuses.py:54-62`, `approval.py:221-223`).

**Key point:** FULL is *not* "no policy" — it only removes the per-action approval grant requirement. Snapshot, writability, path allowlist, macro BG06/R13 gates still apply in FULL.

### 1.2 Operation classes — READ_ONLY / MUTATING / EXECUTING (+ RECOVERY)

- `OperationClass` = `READ_ONLY`, `MUTATING`, `EXECUTING` (`statuses.py:17-27`). Used for approval and COM-retry decisions.
- The per-service operation class is set by each service when submitting to the worker; `security/macro_policy.classify_operation` is explicitly **non-authoritative / illustrative / stale** and **not wired into enforcement** (`src/excel_vba_mcp/security/macro_policy.py:1-12`, `180-221`). Do not rely on it (HIGH).
- **RECOVERY** is a distinct access class handled only by `ApprovalPolicy.check_recovery` (`approval.py:289-317`): auto-allowed in **both** FULL and RESTRICTED, pid-scoped, audited, discards working copies only. It is *not* an `OperationClass` member — it is a separate approval path (`session_service.py:757-763`).

### 1.3 Macro approval + `macro_classify_background` (BG06) — does classification authorize execution?

- `macro_run` gate order (`src/excel_vba_mcp/application/macro_service.py:383-647`):
  1. execute-mode + writable session (`macro_service.py:422-423`, `_require_execute_context` `666-681`);
  2. resolve procedure from inventory (`macro_service.py:426-473`);
  3. **deny-by-default approval** `check(key, EXECUTING)` with key = canonical_path + sha256 + `"macro_run"` + procedure (`macro_service.py:477-503`); denial → `BLOCKED`/`APPROVAL_DENIED`;
  4. **BG06 background_safe gate** bound to `(workbook_sha256, project_source_hash, procedure)` (`macro_service.py:505-542`); unclassified → `BACKGROUND_SAFETY_NOT_VERIFIED`; classified unsafe → `INTERACTIVE_UI_NOT_SUPPORTED`;
  5. **R13 interactive-UI preflight veto** (`macro_service.py:544-580`);
  6. single `Application.Run` dispatch (`macro_service.py:582-618`).
- `macro_classify_background` is **UNGATED bookkeeping** (`macro_service.py:721-826`, `api/macro_tools.py:132-168`). It records a review outcome in the in-memory `BackgroundSafetyRegistry`; it is **not** a security boundary and **does not authorize execution** — `run_macro` still enforces approval + R13 (`macro_service.py:731-733`, `macro_tools.py:140-145`).
- Classification is invalidated by a workbook-hash change or a `project_source_hash` change (source edit); cell-only edits do not invalidate (`macro_service.py:229-245`, `505-506`, `1133-1150`).
- Registry is in-memory, server-lifetime, starts empty, fail-closed (`macro_service.py:229-286`).

### 1.4 Snapshot prerequisites — who creates it?

- Guarded mutations call `snapshot_service.require_snapshot(workbook_id)` **before** mutating. This is wired into ranges, tables, sheets, names, controls, shapes, pivot_tables, and vba_project services (grep evidence: `ranges_service.py:286`, `tables_service.py:297`, `sheets_service.py:145`, `names_service.py:280`, `controls_service.py:147`, `shapes_service.py:141`, `pivot_tables_service.py:331`, `vba_project_service.py:324`).
- `require_snapshot` reads the **in-memory** `_latest_snapshot[workbook_id]` and raises `SNAPSHOT_FAILED` if absent (`src/excel_vba_mcp/application/snapshot_service.py:400-425`).
- `_latest_snapshot` is populated **only** by `create_snapshot` (`snapshot_service.py:176`) or `restore_snapshot` (`snapshot_service.py:348`).
- **Therefore the agent must call `snapshot_create` itself; a guarded mutation does NOT take its own snapshot** (HIGH). `create_snapshot` itself is gated by approval (MUTATING) but cannot require a pre-existing snapshot (`snapshot_service.py:161-166`, `456-486`).
- **Discoverability gap:** the gate uses the in-memory `_latest_snapshot` dict, but the only listing tool `snapshot_list` returns `self._store.list()` (the on-disk store), not `_latest_snapshot` (`snapshot_service.py:385-395`). `has_snapshot` exists (`snapshot_service.py:427-430`) but is **not exposed as any MCP tool**. So an agent cannot reliably predict whether `require_snapshot` will pass for a given `workbook_id` (MEDIUM-HIGH).

### 1.5 Read-only vs writable workbook semantics

- Writability is `WorkbookInfo.read_only`, set at open from the `read_only` argument and mirrored from `wb.ReadOnly` (`src/excel_vba_mcp/adapters/excel/workbook_adapter.py:99`, `114`; `domain/models.py:91`).
- `_require_writable` rejects a read-only working copy with `WORKBOOK_READ_ONLY` (`ranges_service.py:806-816`; same pattern in every mutation service).
- `excel_open` defaults to `mode=execute`, `read_only=False` (writable) (`api/workbook_tools.py:88-127`).
- Macro execution additionally requires execute-mode **and** writable (`macro_service.py:666-681`).
- Writability is discoverable per workbook via `excel_get_status(workbook_id)` / `excel_list_workbooks` / `WorkbookInfo.read_only` (HIGH).

### 1.6 Build / version gates (target build, Excel 16.0 14332)

- The "target build" (Excel 16.0.14332) is a **design-time constant referenced in docstrings/comments only** — e.g. `compile_service.py:3`, `statuses.py:476`, `pivot_tables_adapter.py:1286`, `ranges_adapter.py:31`. It is **not enforced by a runtime gate** and **not returned by any tool** (HIGH).
- The only version-ish values exposed:
  - `WorkbookInfo.excel_version` = `Application.Version` (typically `"16.0"`, **not** the 4-part build) (`workbook_adapter.py:100-106`, `models.py:93`).
  - `server_health` returns server package `__version__ = "0.1.0"` (`server.py:1834-1838`, `src/excel_vba_mcp/__init__.py:5`).
- No tool returns the actual Excel build number (14332) or a server "build" beyond the package version (HIGH).

### 1.7 Macro preflight (R13 veto) + postcondition checks

- R13 is a **heuristic token scan** (case-insensitive) over the single procedure body; any hit blocks with `INTERACTIVE_UI_NOT_SUPPORTED`; a heuristic (imprecisely enumerated) procedure is also blocked fail-closed (`macro_service.py:544-580`, tokens `126-150`, `_find_interactive_tokens` `1177-1188`). It is a **veto, never a certificate** (`macro_service.py:34-50`).
- Postcondition: the only post-dispatch verification in the macro path is the honest status mapping (`classify_macro_failure` `macro_policy.py:273-310`; `_handle_hung` `macro_service.py:828-912`). Range mutations verify by re-reading written values (`ranges_service.py:281-283` docstring). There is no generic "postcondition" framework beyond per-service re-reads (MEDIUM).

### 1.8 Job lifecycle, `job_cancel`, `SUSPECTED_STALLED`, `INTERRUPTED_UNKNOWN`

- `JobState` = `queued, running, succeeded, failed, expired, unknown, interrupted_unknown` (`src/excel_vba_mcp/application/job_registry.py:71-81`); terminal set `84-92`.
- Transitions: sync job `queued→running→terminal` (`job_registry.py:421-467`); deadline → `EXPIRED` (call left running, never cancelled) `436-444`; busy → `FAILED` with `retry_after_s` `445-455`.
- `job_cancel` is **intent-only**: sets `cancel_requested=True`, cannot interrupt the uninterruptible STA COM call; the job still runs to a terminal state (`job_registry.py:328-353`; `api/job_tools.py:113-128`).
- `SUSPECTED_STALLED` is a `MacroRunStatus` (`statuses.py:105`) produced by `_handle_hung` on soft-deadline exceed; owned pid left running, never terminated, never auto-replayed (`macro_service.py:828-912`). A job whose result status ∈ `{suspected_stalled, needs_recovery, timeout}` is interruptible (`job_registry.py:98`, `127-132`).
- `INTERRUPTED_UNKNOWN` is produced by `mark_interrupted`, wired into a **successful** `excel_recover` (`session_service.py:827-835`, `868-880`; `job_registry.py:355-399`): definitive, `outcome='unknown'`, `retryable=False`, never completed, never replayed. A late STA completion cannot overwrite it (`job_registry.py:509-534`).
- `prior_job` in recovery is `"unknown"` when an op was in flight / deadline-exceeded / needs-recovery, else `"completed"` (`session_service.py:765-773`).

### 1.9 Error taxonomy (`domain/errors.py`)

`ErrorCode` closed set (`errors.py:19-67`):
`ACCESS_VBOM_DISABLED, VBA_PROJECT_LOCKED, MISSING_REFERENCE, COM_BUSY, RPC_E_CALL_REJECTED, COMPILE_ERROR, COMPILE_INDETERMINATE, MACRO_BLOCKED, MACRO_NOT_FOUND, MACRO_RUNTIME_ERROR, MACRO_TIMEOUT, EXCEL_HUNG, SNAPSHOT_FAILED, WORKBOOK_READ_ONLY, PATH_NOT_ALLOWED, INVALID_RANGE, PROTECTED_TARGET, UNSUPPORTED_PIVOT_SOURCE, STALE_TARGET, AMBIGUOUS_FILTER, REFRESH_FAILED, MODAL_UI_BLOCKED, LISTOBJECT_INTERSECTION, APPROVAL_DENIED, BACKGROUND_SAFETY_NOT_VERIFIED, INTERACTIVE_UI_NOT_SUPPORTED, WORKBOOK_NOT_FOUND, REVISION_CONFLICT, NOT_FOUND`.

- HRESULT→code map: `errors.py:72-80`; `code_from_hresult` `90-102`; `error_from_hresult` `403-418`; code→exception `_ERROR_BY_CODE` `423-459`.
- Error payload today = `to_dict()` → `{code, message, hresult, hresult_hex, context}` (`errors.py:138-147`). Surfaced to clients as a JSON string inside `ToolError` by `ExcelVbaMcpServer.call_tool` (`server.py:1808-1826`); pydantic `ValidationError` is coerced to `INVALID_RANGE` (`server.py:1799-1804`, `1819-1825`).
- **`isRetryable`: does not exist at the error level** (HIGH). Only job-level `retryable` exists (`job_registry.py:164`, `188`, `207`; set `False` only for interrupted_unknown `395`).
- **`suggestedNextActions`: does not exist anywhere** (HIGH; grep found none).
- `context` is a free-form `dict[str, Any]` (`errors.py:130`), so structure varies per raise site (e.g. `allowed_roots` only on PATH_NOT_ALLOWED `path_policy.py:168-174`; `reason` on path denials `path_policy.py:201-213`).

---

## 2. Discoverability matrix (agent question → discoverable today? → gap)

| # | Agent question | Discoverable today? | Evidence / gap |
|---|---|---|---|
| a | Current access mode (FULL/RESTRICTED) | **Yes (partial)** | `excel_get_status` → `SessionStatus.access_mode` (`session_service.py:1104`, `models.py:111`). Gap: only after session service initialized; not in `server_health`; description doesn't mention it (`server.py:227`). |
| b | Whether a mutation is allowed | **No (only by trying)** | No preflight tool. Approval is evaluated at call time (`approval.py:189-264`); outcome only via error `APPROVAL_DENIED`. |
| c | Whether approval is required / how to satisfy it | **No** | No tool exposes mode→requirement mapping or grant/allowlist state. `ApprovalPolicy` has no read tool; grants are client-supplied at runtime (`approval.py:24-29`, `176-187`). |
| d | Whether a snapshot is needed / who creates it | **No (partial)** | Gate is `require_snapshot` on in-memory `_latest_snapshot` (`snapshot_service.py:400-425`); `has_snapshot` not exposed; `snapshot_list` lists the store, not the gate (`snapshot_service.py:385-395`). Agent must call `snapshot_create` itself. |
| e | Whether a workbook is writable | **Yes** | `WorkbookInfo.read_only` via `excel_get_status`/`excel_list_workbooks` (`workbook_adapter.py:99`; `models.py:91`). |
| f | Actual Excel build + server build/version | **Partial** | Server version via `server_health` (`server.py:1838`); Excel `Application.Version` (e.g. "16.0") via `WorkbookInfo.excel_version` (`workbook_adapter.py:100-106`). **Gap:** no 4-part build (14332); no target-build capability flag. |
| g | Feature capabilities (VBA, macros, pivot, background jobs, recovery) | **No** | No capability tool. Capabilities are implicit in tool registration (`server.py:1479-1523+`) and per-result `unsupported` statuses (e.g. `PivotSourceType.UNSUPPORTED`, `statuses.py:443-460`). Compile is always `INDETERMINATE` (`compile_service.py:294-346`). |
| h | Correct next action after each major error category | **No** | `to_dict()` has no `isRetryable`/`suggestedNextActions` (`errors.py:138-147`). Only job `retry_after_s`/`retryable` exist (`job_registry.py:184-188`). |

---

## 3. Proposed agent-readable ENVIRONMENT & POLICY contract

**Placement recommendation:** add a **new read-only tool `xlvba_environment`** (no `workbook_id`, no COM) rather than overloading `excel_get_status`/`server_health`.
- Rationale: `excel_get_status` is session/worker-state oriented and requires the session service; `server_health` is a pure liveness probe returning `dict[str,str]`. A dedicated tool keeps SCHEMA_VERSION=1 DTOs untouched and gives a stable, single discovery call (see §5 risk notes).
- Also add a **compact `policy` block to `excel_get_status`** (additive, optional) so per-workbook context (writable/snapshot) is co-located; keep `xlvba_environment` as the authoritative server-level source.

Proposed `xlvba_environment` payload (all additive; `SCHEMA_VERSION` stays 1):

```jsonc
{
  "schema_version": 1,
  "server": {
    "name": "excel-vba-mcp",
    "version": "0.1.0",                 // __version__
    "access_mode": "full",              // AccessMode
    "macro_timeout_s": 90.0,            // ServerConfig
    "open_timeout_s": 300.0,
    "liveness_ceiling_s": 600.0,
    "log_level": "INFO"
  },
  "excel": {                            // null when worker not started
    "available": true,
    "application_version": "16.0",      // Application.Version (not build)
    "build": null,                      // 4-part build if obtainable, else null (honest)
    "target_build": "16.0.14332",       // design-time constant, labelled as target
    "owned_pid": 12345,
    "vbom_available": true              // from a read-only probe / model vba_access
  },
  "capabilities": {
    "vba": true, "macros": true, "pivot_tables": true,
    "background_jobs": true, "recovery": true,
    "non_activating_compile": false,    // always false on validated build
    "pivot_filter_kinds": ["manual"],   // supported subset
    "pivot_source_types": ["database"]
  },
  "policy": {
    "approval": { "mode": "full", "mutating_requires_grant": false,
                  "executing_requires_grant": false, "recovery_auto_allowed": true },
    "snapshot": { "required_for_mutations": true, "created_by": "agent:snapshot_create" },
    "path": { "allowed_roots": ["<root1>", "..."], "allow_unc": false },
    "macro": { "background_safe_required": true, "r13_preflight": "veto",
               "classification_is_authorization": false }
  },
  "workbooks": [                        // optional per-open-workbook summary
    { "workbook_id": "...", "read_only": false, "mode": "execute",
      "has_snapshot": false, "dirty": false }
  ]
}
```

Field provenance (all HIGH): access_mode `config.py:49-53`/`approval.py:171-174`; timeouts `config.py:54-73`; log_level `config.py:48`; server version `__init__.py:5`; excel version `workbook_adapter.py:100-106`; owned_pid `health_tools.py:47`; capabilities from tool registration `server.py:1479-1523+` and enums `statuses.py:443-524`; compile `compile_service.py:294-346`; path roots `path_policy.py:98-102`; macro gates `macro_service.py:477-580`; snapshot `snapshot_service.py:400-430`; per-workbook `has_snapshot` requires exposing `SnapshotService.has_snapshot` (`snapshot_service.py:427-430`).

**Discoverability deltas this closes:** (a) full, (b) via `policy.approval`, (c) via `policy.approval`, (d) via `policy.snapshot` + `workbooks[].has_snapshot`, (f) partial (build honestly null), (g) full, (h) via §4.

---

## 4. Proposed STRUCTURED ERROR contract (mapped to existing taxonomy)

Extend `ExcelVbaMcpError.to_dict()` **additively** (keep existing keys; add new ones) so every tool failure carries a uniform envelope. No new `ErrorCode` members required; `errorCategory` is derived.

```jsonc
{
  "code": "SNAPSHOT_FAILED",            // existing ErrorCode (errors.py:19-67)
  "errorCategory": "missing_prerequisite", // see mapping below
  "message": "A successful snapshot is required before mutating this workbook.",
  "isRetryable": false,
  "suggestedNextActions": [
    "snapshot_create(workbook_id) then retry the mutation"
  ],
  "details": { "workbook_id": "..." },  // = existing context
  "hresult": null,
  "hresult_hex": null
}
```

### 4.1 Category taxonomy (mutually exclusive; additive)

| errorCategory | Meaning | Example codes |
|---|---|---|
| `permission_denied` | Policy/approval/trust denial; not fixable by retry | `APPROVAL_DENIED`, `MACRO_BLOCKED`, `PATH_NOT_ALLOWED`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` |
| `missing_prerequisite` | A required precondition is absent (snapshot, writable, execute-mode) | `SNAPSHOT_FAILED`, `WORKBOOK_READ_ONLY`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `REVISION_CONFLICT`, `STALE_TARGET` |
| `unsupported_feature` | Capability absent on this build/config; do not retry | `UNSUPPORTED_PIVOT_SOURCE`, `INTERACTIVE_UI_NOT_SUPPORTED`, `COMPILE_INDETERMINATE`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `AMBIGUOUS_FILTER`, `PROTECTED_TARGET` |
| `transient_com_failure` | Excel/VBE busy; retry after backoff | `COM_BUSY`, `RPC_E_CALL_REJECTED` |
| `deadline` | Soft deadline exceeded; op left running | `MACRO_TIMEOUT` (SUSPECTED_STALLED), job `expired` |
| `suspected_stall` | Deadline + possible modal; agent-directed recovery | `EXCEL_HUNG`, `MODAL_UI_BLOCKED`, `MacroRunStatus.SUSPECTED_STALLED` |
| `interrupted_unknown` | Job finalized by recovery; never replay | job `interrupted_unknown` (`outcome=unknown`, `retryable=false`) |
| `not_found` | Caller addressed a missing identity | `WORKBOOK_NOT_FOUND`, `MACRO_NOT_FOUND`, `NOT_FOUND` |
| `validation` | Caller argument/schema error | pydantic `ValidationError`→`INVALID_RANGE` (`server.py:1799-1804`) |

### 4.2 `isRetryable` + `suggestedNextActions` per code (proposed)

| Code | isRetryable | suggestedNextActions |
|---|---|---|
| COM_BUSY / RPC_E_CALL_REJECTED | true | `["wait retry_after_s (if present) then retry the same read"]` |
| EXCEL_HUNG / MODAL_UI_BLOCKED | false | `["excel_recover", "re-open workbook", "do not replay the interrupted mutation"]` |
| MACRO_TIMEOUT (SUSPECTED_STALLED) | false | `["excel_recover if still stalled", "never auto-replay; inspect macro_get_last_result"]` |
| SNAPSHOT_FAILED | false | `["snapshot_create(workbook_id)", "retry mutation"]` |
| WORKBOOK_READ_ONLY | false | `["re-open with read_only=false (mode=execute)"]` |
| APPROVAL_DENIED | false | `["open in full mode or supply grant/allowlist entry", "retry"]` |
| PATH_NOT_ALLOWED | false | `["use a path under details.allowed_roots"]` |
| BACKGROUND_SAFETY_NOT_VERIFIED | false | `["macro_classify_background(...background_safe=true) after review", "retry macro_run"]` |
| INTERACTIVE_UI_NOT_SUPPORTED | false | `["remove interactive-UI tokens or run interactively (out of scope)"]` |
| UNSUPPORTED_PIVOT_SOURCE / COMPILE_INDETERMINATE | false | `["use a supported source (database) / accept indeterminate compile"]` |
| STALE_TARGET / REVISION_CONFLICT | true | `["re-read source/hash", "retry with fresh expected_*"]` |
| WORKBOOK_NOT_FOUND / NOT_FOUND / MACRO_NOT_FOUND | false | `["verify workbook_id / object name / procedure"]` |
| INVALID_RANGE (validation) | false | `["fix arguments to match input schema"]` |

Implementation note: mapping is a pure function `category_for(code)` + `retryable_for(code)` + `actions_for(code, context)`; no new ErrorCode members, no schema bump. `context` already carries e.g. `allowed_roots` (`path_policy.py:168-174`) and `reason` (`path_policy.py:201-213`) to populate `details`.

---

## 5. Compatibility / risk notes

- **Do not weaken safety.** The contract is read-only/additive. `xlvba_environment` must not create snapshots, grants, or classifications, and must not expose secrets (path roots are already non-secret; `path_policy._safe_name` redacts). Macro `classification_is_authorization=false` must be stated explicitly to avoid implying BG06 grants execution.
- **No mandatory human approval added.** FULL remains default (`config.py:49-53`); the contract only *reports* policy. RECOVERY stays auto-allowed (`approval.py:289-317`).
- **Keep `SCHEMA_VERSION=1`.** All additions are optional fields on new/optional DTOs; existing `DomainModel` uses `extra="forbid"` (`models.py:68`) — so new fields must be added to the model classes, not passed ad-hoc. Prefer a **new** `EnvironmentReport` model and an **additive** `policy` block on `SessionStatus` (already has `access_mode` `models.py:111`) to avoid breaking `extra="forbid"` consumers. `to_dict()` extension is additive and safe (`errors.py:138-147`).
- **Honesty about build.** `Application.Version` yields "16.0", not the 4-part build; the contract must return `build: null` unless a real build source is added (do not fabricate 14332). `target_build` is a labelled design constant only.
- **Snapshot discoverability caveat.** Exposing `has_snapshot` requires a new tool/field reading `SnapshotService.has_snapshot` (in-memory). `snapshot_list` (store) must not be conflated with the gate.
- **Error envelope change is client-visible.** `call_tool` already JSON-encodes `to_dict()` (`server.py:1826`); adding keys is backward-compatible for tolerant clients but strict clients should be noted.
- **No new COM path.** `xlvba_environment` is non-COM (like `server_health` `server.py:1834-1838`); `excel.available=false` when the worker is down.

---

## 6. Open questions for Auditor / Validator

1. **Build number source (LOW/MEDIUM):** is there any supported COM/registry source for the 4-part Excel build (e.g. `Application.Build` or `Version` on some hosts) that is safe to call read-only, or should `build` stay `null`? (Today only `Application.Version` is read — `workbook_adapter.py:100-106`.)
2. **`has_snapshot` semantics (MEDIUM):** should the environment contract report the in-memory gate (`_latest_snapshot`) or the persisted store? They differ (`snapshot_service.py:176` vs `385-395`). Recommend the gate, since that is what mutations check.
3. **Capability detection vs static declaration (MEDIUM):** should `capabilities` be *probed* (e.g. VBOM via a read-only model call) or *statically declared* from build constants? Static is simpler and matches current design (`compile_service.py:3`), but may misreport on non-validated hosts.
4. **Error `errorCategory` for `COMPILE_INDETERMINATE` (MEDIUM):** classify as `unsupported_feature` (current placement) vs a dedicated `indeterminate` category — confirm no client depends on the current bare code.
5. **`isRetryable` for `STALE_TARGET`/`REVISION_CONFLICT` (MEDIUM):** proposed `true` (re-read + retry). Confirm this does not encourage blind replay of a mutation that already partially applied.
6. **Where to expose `policy` (MEDIUM):** new `xlvba_environment` only, vs also additive `SessionStatus.policy`. Confirm the extra="forbid" DTO change is acceptable without a SCHEMA_VERSION bump.
7. **Validation of "no new mandatory approval" (HIGH):** confirm the Auditor agrees the environment tool must be non-mutating and side-effect-free.