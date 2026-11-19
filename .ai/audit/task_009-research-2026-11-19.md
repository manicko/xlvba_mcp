# task_009 — Office/Excel macro trust and security boundaries (Research)

- **Date:** 2026-11-19
- **Role:** Researcher (investigation only — **no production code modified**)
- **Task:** `task_009` — Implement approval policy, macro policy and operation audit logging
- **Spec sections:** §7.1 (Workbook / Excel Session), §7.2 (VBA Project), §7.3 (Macro Execution), §10 (Security & trust model), §10.1 (Safe open order), §11 (errors), A01/A08/A15
- **Plan sections:** task_009 (lines 457–508), R10 risk, G2 gate (researcher required)

## Scope

This report answers the five researcher questions for `task_009` inside the
**already-approved architecture** (no new runtime/transport, no new abstractions).
Inputs verified from the repository:

- `src/excel_vba_mcp/domain/statuses.py` — `OperationClass` = `READ_ONLY` /
  `MUTATING` / `EXECUTING`; `SessionMode` = `INSPECT` / `EXECUTE`; `MacroRunStatus`.
- `src/excel_vba_mcp/domain/errors.py` — `ErrorCode` (incl. `MACRO_BLOCKED`,
  `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `ACCESS_VBOM_DISABLED`,
  `VBA_PROJECT_LOCKED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED`,
  `WORKBOOK_READ_ONLY`), HRESULT map, `error_from_hresult`.
- `src/excel_vba_mcp/domain/models.py` — Pydantic v2 DTO style
  (`DomainModel(extra="forbid")`, `TimestampedModel`, `MacroRunResult`,
  `CompileDiagnostic`, `SnapshotManifest`).
- `src/excel_vba_mcp/security/path_policy.py` — `PathPolicy.validate_path()`
  returns a **canonical** `Path` (realpath cross-check), used as workbook
  identity anchor. `ARTIFACT_ROOT_NAMES = {"snapshots","artifacts"}` already
  governs where audit/snapshot sidecars may be written.

**No approval/audit/macro-policy code exists yet** — confirmed absent:
`security/approval.py`, `security/audit.py`, `security/macro_policy.py` do not
exist. This report recommends contracts and policy logic only; it does not
write them.

Confidence key: **[HIGH]** = verified against primary source (project code or
Microsoft Learn); **[MED]** = reasoned inference; **[LOW]** = assumption to be
proven in task_009 unit tests.

---

## 1. Operation → OperationClass table

Classification drives two independent decisions:

1. **Approval gate** (task_009): `READ_ONLY` → allowed without approval;
   `MUTATING`/`EXECUTING` → require approval or allowlist entry.
2. **COM retry** (task_010): only `READ_ONLY` may be safely retried after
   `COM_BUSY`/`RPC_E_CALL_REJECTED`; `MUTATING`/`EXECUTING` are never blindly
   replayed (spec §11, statuses.py docstring).

| # | Stage-1 operation | Tool / API surface | Class | Rationale / ambiguity |
|---|---|---|---|---|
| 1 | Open workbook | `excel_open` → `Workbooks.Open` | **READ_ONLY** (on disk) / state-changing (on session) | Opens an isolated instance; does not mutate the file. Session state changes, but no *workbook* mutation and no macro execution. Macro trust is decided by mode (inspect=ForceDisable, execute=ByUI), not by class. **[MED]** |
| 2 | Close workbook (no save) | `excel_close` → `Workbook.Close(SaveChanges:=False)` | **READ_ONLY** (file unchanged) / **MUTATING** if `save_changes=True` | **AMBIGUOUS.** Spec §7.1: "Close без save_changes не должен молча записывать". Close-without-save must be classified READ_ONLY for approval; close-with-save is MUTATING. Policy must key on the `save_changes` flag, not the verb. **[HIGH rationale, MED impl]** |
| 3 | Save workbook | `excel_save` → `Workbook.Save` | **MUTATING** | Persists in-memory changes to disk. **[HIGH]** |
| 4 | Snapshot create | `snapshot_create` → `SaveCopyAs` (open) / file copy (closed) | **READ_ONLY** on the source (SaveCopyAs does not alter the in-memory/open workbook; spec §9.1) | Produces a new artifact; source unchanged. Required precondition *before* mutations, so it must not itself require approval — otherwise deadlock (can't snapshot to get approval to mutate). **[MED]** |
| 5 | Snapshot restore | `snapshot_restore` → close copy w/o save + open snapshot | **MUTATING** (replaces working copy state) | Discards working state and reopens baseline. Treat as MUTATING (destructive to working copy). **[MED]** |
| 6 | VBA component read (inventory) | `VBProject.VBComponents` enumeration | **READ_ONLY** | No file/code change. **[HIGH]** |
| 7 | VBA source read (CodeModule) | `CodeModule.Lines[...]` | **READ_ONLY** | **[HIGH]** |
| 8 | VBA source export (.bas/.cls/.frm+.frx) | VBIDE `Export` | **READ_ONLY** (writes sidecar to artifact root) | Writes export artifacts but not to the workbook. The artifact write is governed by PathPolicy artifact roots, not approval. **[MED]** |
| 9 | VBA source write (edit/insert) | `CodeModule.AddFromString` / `DeleteLines` / `ReplaceLine` | **MUTATING** | **AMBIGUOUS only in appearance:** `AddFromString` *looks* additive but changes the code module. Spec §7.2: code change only on a working copy, after snapshot, with before/after hash. Always MUTATING. **[HIGH]** |
| 10 | Add VBA component | `VBComponents.Add(...)` | **MUTATING** | Structural code change. **[HIGH]** |
| 11 | Remove VBA component | `VBComponents.Remove(...)` | **MUTATING** (destructive) | Must NOT remove Document Modules (spec §7.2). Destructive → snapshot + approval. **[HIGH]** |
| 12 | Sheet add | `Sheets.Add` / `Worksheets.Add` | **MUTATING** | Structural. **[HIGH]** |
| 13 | Sheet delete | `Worksheet.Delete` | **MUTATING** (destructive) | Spec §7.5: destructive → snapshot + explicit confirmation + ≥1 visible sheet remains. **[HIGH]** |
| 14 | ListObject row add | `ListRows.Add` | **MUTATING** | Spec §7.6, R07. **[HIGH]** |
| 15 | ListObject row delete/update | `ListRows(i).Delete` / value write | **MUTATING** (destructive if ambiguous) | Ambiguous filter → preview + require clarification, never auto-update "similar" rows (§7.6). **[HIGH]** |
| 16 | ListObject column add | `ListColumns.Add(Position)` | **MUTATING** | Spec §7.6. **[HIGH]** |
| 17 | ListObject column delete | `ListColumns(i).Delete` | **MUTATING** (destructive) | **[HIGH]** |
| 18 | Defined-name add | `Names.Add(...)` | **MUTATING** | Spec §7.7. **[HIGH]** |
| 19 | Defined-name delete | `Names(name).Delete` | **MUTATING** (destructive) | **[HIGH]** |
| 20 | OnAction set | `control_set_on_action` → `Shape/ControlFormat.OnAction = ...` | **MUTATING** | **AMBIGUOUS:** setting `OnAction` does not *run* a macro, but it wires a control to a macro (a latent execution path) and is explicitly gated by approval in the spec tool list ("control_set_on_action (последний — только после approval)", line 300). Classify MUTATING, not EXECUTING. **[HIGH]** |
| 21 | Compile VBA | VBE `Debug → Compile` via CommandBars/UIA | **EXECUTING** | Runs the VBA compiler inside VBE; does not run a macro but executes code-generation and can raise modal dialogs. Spec §8 treats compile as a distinct execution-class technical mechanism. Not a file mutation. **[MED]** — see note below. |
| 22 | Run macro | `Application.Run` | **EXECUTING** | Spec §7.3: "потенциально опасное действие", require approval/allowlist by path/hash + procedure, log each run. **[HIGH]** |
| 23 | Test run | declarative runner (compile + macro_run + mutations + assertions) | **EXECUTING** (composite) | Contains macro execution and mutations. Classify at the *suite* level as EXECUTING so approval covers the highest-risk constituent; individual steps still carry their own class for retry. **[MED]** |

### Ambiguous cases — explicit resolutions

- **`SaveCopyAs` (snapshot create) → READ_ONLY.** Excel docs: SaveCopyAs
  "saves a copy of the workbook to a file but does not modify the open workbook
  in memory." It must be READ_ONLY so that the *mandatory pre-mutation snapshot*
  (spec §9.1, task_014) is not itself blocked behind the approval it is meant to
  protect. The artifact *write* is a separate concern handled by PathPolicy
  artifact-root allowlisting.
- **`Workbook.Close` → class depends on `save_changes`.** Default (no
  `save_changes`) = READ_ONLY (never silently saves, §7.1). `save_changes=True`
  = MUTATING (persists changes). Recommend the policy gate receive the
  *effective* class after resolving the flag, not the bare verb.
- **`CodeModule.AddFromString` → MUTATING (never READ_ONLY).** "Add" is
  additive-sounding but mutates the module; §7.2 requires snapshot + source
  revision check + before/after hash. Same for `ReplaceLine`/`DeleteLines`.
- **`OnAction` set → MUTATING, not EXECUTING.** It does not invoke a macro now;
  it creates a persistent binding to one. EXECUTING is reserved for actions that
  actually run code (compile, macro run, test run). This matches the spec's own
  gate ("только после approval") while keeping the retry semantics clean (an
  OnAction set can be safely reasoned about; it is not a blind-replay hazard).
- **Compile → EXECUTING.** No public `Workbook.CompileProject()` exists (spec
  §8); compilation is the VBE `Debug → Compile` command. It executes the VBA
  compiler and can surface modal dialogs, so it belongs with execution-class
  operations for retry/modal handling, even though it is not a macro run.

> **Implementation note (task_009):** classification should be a single
> pure function `classify(tool_name, params) -> OperationClass` (or a static
> table keyed by `tool_name`) co-located with the policy, so API/adapters and
> the COM worker share one source of truth. The `PathPolicy.validate_path`
> signature already takes an `OperationClass`, so the class must be computed
> *before* path validation.

---

## 2. Approval / allowlist model

### Requirements from spec

- §10 Approval: read = READ_ONLY (no approval); VBA write, table/Name/Sheet
  change, object delete, macro run ⇒ local approval **or** allowlist by
  **workbook hash + action**.
- §7.3: macro approval by **workbook path/hash + macro procedure**.
- Plan R10: "approval or allowlist by workbook path/hash + procedure".
- Deny-by-default is mandatory: "unapproved mutation/execution is denied"
  (acceptance criteria), and "deny-by-default" is in the validator's brief.

### Recommended schema (minimal, fits Pydantic v2 DTO style)

Two layers: a **policy document** (static allowlist, config-supplied) and a
**grant** (one-time / session approval supplied by the MCP client at runtime).
Both share one `ApprovalKey`.

```text
ApprovalScope   = "one_time" | "session"        # lifetime of a grant
ApprovalSource  = "client" | "allowlist" | "config"

ApprovalKey:
    canonical_path : str          # from PathPolicy.validate_path() — NEVER raw input
    sha256         : str | None   # content hash of the workbook file (preferred key)
    action         : str          # canonical tool/action name, e.g. "vba_write",
                                  #   "macro_run", "sheet_delete", "table_row_delete",
                                  #   "name_add", "control_set_on_action"
    procedure      : str | None   # for macro_run: "Module.Proc" (qualified);
                                  #   None = action applies to whole workbook

ApprovalGrant:
    key            : ApprovalKey
    scope          : ApprovalScope
    granted_at     : datetime (UTC)
    expires_at     : datetime (UTC) | None   # None = no expiry (session end still applies)
    granted_by     : str          # client identity / operator label (NOT a secret)
    source         : ApprovalSource
    grant_id       : str          # correlation id for audit
```

Allowlist entry (static, config/`macro-policy.example.toml`):

```text
AllowlistEntry:
    key            : ApprovalKey
    note           : str | None   # operator rationale, no secrets
```

### Identity rule (ties to PathPolicy — anti-bypass)

- `canonical_path` **must** come from `PathPolicy.validate_path()`, which
  resolves symlinks/reparse points and rejects traversal/UNC (spec §10, A15).
  Never key an approval on caller-supplied text.
- `sha256` is the **strongest** key: an approved grant for hash `abc…` does not
  authorize a different file that happens to share a basename. When the file
  content changes, the hash changes and a fresh approval is required. Recommend
  **hash is authoritative when present**; `canonical_path` is the fallback for
  not-yet-hashed or mid-open states.
- Approval check must run **before dispatch to the STA worker** (in the API/
  application layer), because the worker must never receive an unapproved
  MUTATING/EXECUTING request (COM proxies never leave the worker; policy is a
  pre-worker concern).

### Decision flow (deny-by-default)

1. Compute `OperationClass` for the tool+params (pure function, §1).
2. If `READ_ONLY` → allow; still audit.
3. Else (`MUTATING`/`EXECUTING`):
   a. If an `AllowlistEntry` matches `(sha256|canonical_path, action[, procedure])`
      → allow (source=`allowlist`).
   b. Else if a live `ApprovalGrant` matches and `now <= expires_at` and
      scope/sid satisfies → allow (source=`client`/`config`).
   c. Else → **DENY** with `MacroBlockedError`/`PathNotAllowedError`-style typed
      error (see §3 for macro vs generic). Default deny is the terminal branch.
4. Audit the decision either way (allow *and* deny).

### One-time vs session approval

- **one_time:** consumed on first successful match (delete the grant after use);
  required for the most dangerous actions (macro_run, vba_write,
  destructive deletes) unless the operator explicitly opted into session scope.
- **session:** valid until `expires_at` or session teardown, whichever first;
  suitable for iterative work on one working copy. Recommend session grants be
  bound to `workbook_id`/`canonical_path` so a grant cannot drift to another
  workbook.
- **Interaction with snapshots:** a snapshot is READ_ONLY (§1 #4) and therefore
  never needs approval — this avoids the "must approve before snapshot but must
  snapshot before mutate" deadlock.

### No interactive UI

Per the task brief, approval is supplied by the **MCP-layer client** (e.g. an
elicitation/approval object or an out-of-band allowlist file). The server exposes
a single `ApprovalPolicy.check(key) -> ApprovalDecision` and never blocks on
stdin/UI. This keeps the STDIO event loop non-blocking (the server must not
invent an interactive prompt).

---

## 3. Macro policy — diagnose, never mutate

### What the spec permits / forbids (verbatim basis)

- §7.2: VBOM access is a manual Trust Center setting; server **only diagnoses**
  absence and must not change registry/Trust Center. `[S3, S7]`
- §10 Macro security: **do not** set `AutomationSecurity=Low` as a blanket
  rule; **do not** enable all macros; use Office policy/Trusted Document/
  Trusted Publisher/Trusted Location only if the user trusts the source; **do
  not** remove MOTW (Mark-of-the-Web) automatically. `[S4]`
- §10.1 / §7.1: inspect ⇒ `ForceDisable`; execute ⇒ `ByUI` immediately before
  `Workbooks.Open`; restore original `AutomationSecurity` in `finally`. Never
  change `AccessVBOM`, `MacroSettings`, Trusted Locations automatically.
- §10 VBA project lock: do **not** bypass password/protection; return
  `VBA_PROJECT_LOCKED` and suggest manual unlock by owner.
- Plan task_009 extra_context restates all of the above; acceptance criteria:
  "inspect mode does not enable macros", "Trust Center/registry never modified",
  "VBA project protection never bypassed".

### Verified Microsoft facts

- **`Application.AutomationSecurity`** (Excel) — read/write `MsoAutomationSecurity`
  (`msoAutomationSecurityLow=1`, `ByUI=2`, `ForceDisable=3`). Defaults to
  `Low` at app start; MS explicitly warns to **set it immediately before/after
  `Workbooks.Open`** "to avoid malicious subversion", and to restore the
  original value. `DisplayAlerts` does **not** suppress security warnings. `[HIGH — MS Learn]`
- **`Workbook.VBProject`** — read-only `VBProject` object; **`VBProject.Protection`**
  is read-only and returns `vbext_pp_none=0` / `vbext_pp_locked=1`. Accessing
  `VBProject` when VBOM trust is off throws (programmatic access "not
  trusted"); the server must map this to `ACCESS_VBOM_DISABLED` and must not
  write the `AccessVBOM` registry value (which Office reverts on exit anyway).
  `[HIGH — MS Learn + MS blog]`
- **`Workbook.HasPassword`** — read-only Boolean; `True` if the workbook has a
  protection password. Note: a workbook must be opened to read it, and a
  password-protected file cannot be opened without the password; therefore
  `HasPassword` diagnoses an *already-open* workbook. `[HIGH — MS Learn]`
- **Trust Center / AccessVBOM** lives at
  `HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM`; Office
  caches it at startup and **reverts out-of-band edits on exit**, so a server
  cannot reliably toggle it anyway — an *additional* reason the spec forbids
  it. `[HIGH — MS Learn blog / Renenyffenegger]`

### Diagnosis-only recommendations (no mutation)

`MacroPolicy` (task_009) must be a **read-only diagnostic + pure decision**
component. Concretely:

1. **VBOM availability** — attempt a cheap `VBProject.VBComponents.Count`;
   on failure (`0x800A17B4` / "not trusted") emit `ACCESS_VBOM_DISABLED` with a
   user-facing instruction. Never write the registry. `[HIGH]`
2. **Project protection** — read `VBProject.Protection`; if `vbext_pp_locked`
   (1) → `VBA_PROJECT_LOCKED`. Never attempt SendKeys/registry bypass. `[HIGH]`
3. **Workbook password state** — read `Workbook.HasPassword` as a diagnostic
   flag only; never store or echo the password (§10 Secrets). `[HIGH]`
4. **Current automation security** — *read* `Application.AutomationSecurity`
   to record the **pre-existing** value in diagnostics/audit so the inspect/
   execute context manager can restore it exactly (spec §10.1 point 1:
   "запомнить исходное AutomationSecurity … немедленно вернуть прежнее
   значение"). Reading is allowed; *setting* is restricted to the narrow
   open/close window described in §10.1. `[HIGH]`
5. **Macro-enabled?** — `Workbook.HasVBProject` / code-module presence (already
   in `WorkbookInfo.has_vba`). `[MED]`

### MACRO_BLOCKED vs MACRO_NOT_FOUND vs MACRO_RUNTIME_ERROR

These three codes already exist in `errors.py`; the policy/adapters must
disambiguate them by **evidence**, not by a single generic COM error:

| Code | When to emit | Evidence source | Confidence |
|---|---|---|---|
| `MACRO_BLOCKED` | Server policy denied the run (no approval / not in allowlist) **or** Office blocked it (inspect-mode `ForceDisable`, Trust Center disabled macros for the source) | Pre-dispatch policy decision, or `Application.Run` throwing because macros are disabled for that file | Policy decision = [HIGH]; Office-block detection = [MED] |
| `MACRO_NOT_FOUND` | Procedure absent or not a callable entry point (private/exposed issues) | `Application.Run` HRESULT `0x800A03ED` maps to `MACRO_NOT_FOUND` in `errors.py`; corroborate with VBIDE `CodeModule` procedure lookup (§7.2 `ProcOfLine`) | [HIGH for HRESULT mapping; MED for private-proc nuance] |
| `MACRO_RUNTIME_ERROR` | Macro was found and started, then raised an exception | HRESULT `0x800A03EC` / `0x80004005` (E_FAIL) map to `MACRO_RUNTIME_ERROR` | [HIGH] |
| `MACRO_TIMEOUT` | Soft timeout exceeded; session → `running/needs_recovery` | Own timing, not COM | [HIGH — design] |

**Critical anti-pattern to avoid (spec §8.3, §11.1):** do **not** report a
compile success or macro success merely because `Application.Run` returned
without a COM exception. `Application.Run` can return the generic "macro may
not be available" message when the project has a **compile error**; the same
generic failure must not be mislabeled `MACRO_NOT_FOUND`. task_005 must prove
which HRESULT Excel actually returns per condition, and task_009's classifier
must map by verified evidence. `[MED — PoC-dependent]`

**Modal-dialog caution:** if a macro raises a modal error dialog, `Application.Run`
blocks and the next COM call must not start until the dialog is handled by UIA
(spec §11, A14). The macro policy does not own this; it just must not classify
the resulting `EXCEL_HUNG`/`needs_recovery` as a clean macro status.

---

## 4. Audit trail

### Requirements (spec §10 Audit; plan task_009)

"Логировать пользователя/процесс, путь/хэш книги, tool name, параметры (с
редактированием секретов), approval, snapshot, время, результат и HRESULT."
Never to MCP stdout (§10 MCP transport: "Не писать логи в stdout").

### Recommended JSON event schema (append-only JSONL)

One JSON object per line (JSONL) is the simplest append-only, tail-friendly,
corruption-resistant format. Written to a **configurable artifact path**
(default under `artifacts/`, which `path_policy.ARTIFACT_ROOT_NAMES` already
permits for `.json`-class sidecars; recommend a dedicated
`artifacts/audit/audit-<date>.jsonl` or a configured `audit_path`).

```jsonc
{
  "schema_version": 1,
  "timestamp": "2026-11-19T14:59:31.123456Z",   // UTC, ISO-8601
  "event_id": "b1e...",                          // unique per event (uuid4)
  "correlation_id": "...",                        // per MCP request/ op (spec §6: Correlation ID)
  "session_id": "...",
  "workbook_id": "...",                           // stable server id (WorkbookInfo.workbook_id)
  "workbook": {
    "canonical_path": "C:\\repo\\work\\Report-copy.xlsm", // canonical, from PathPolicy
    "file_name": "Report-copy.xlsm",              // basename for readability
    "sha256": "9f86d081884c7d65..."               // content hash (authoritative identity)
  },
  "tool": "vba_write",                            // canonical tool/action name
  "operation_class": "mutating",                  // READ_ONLY | MUTATING | EXECUTING
  "params": { /* redacted — see policy below */ },
  "approval": {
    "required": true,
    "decision": "allow",                          // allow | deny
    "source": "client",                           // client | allowlist | config
    "grant_id": "...",                            // correlation to grant
    "scope": "one_time",
    "expires_at": "2026-11-19T15:30:00Z"
  },
  "snapshot_id": "...",                           // pre-mutation baseline, if any (SnapshotManifest.snapshot_id)
  "result": {
    "status": "success",                          // success | error | denied
    "error_code": null,                           // ErrorCode.value when failed
    "hresult": null,                              // raw int when COM involved
    "hresult_hex": "0x00000000",                  // format_hresult() form
    "duration_ms": 1234
  },
  "user": "corp\\m.matveichev",                   // process user (NOT a secret)
  "process": "excel_vba_mcp",                     // server identity
  "excel_version": "16.0.14332"
}
```

Field mapping to existing DTOs so task_009 stays DRY:

- `workbook_id`, `session_id` ← `WorkbookInfo` / `SessionStatus`.
- `snapshot_id` ← `SnapshotManifest.snapshot_id`.
- `result.error_code`, `hresult`, `hresult_hex` ← `ExcelVbaMcpError.to_dict()`
  (`code`, `hresult`, `hresult_hex`); reuse `format_hresult`.
- `duration_ms` ← same convention as `MacroRunResult.duration_ms`.
- `operation_class` ← the shared `OperationClass` enum from `statuses.py`.

### Redaction policy (mandatory)

Define one pure `redact(params) -> dict` used everywhere; never log raw params.

1. **Paths:** log `canonical_path` + `file_name` (basename). The canonical path
   is already the *resolved, allowlisted* path, so it is safe to record; still
   store only basename in any free-text field to minimize directory leakage.
2. **Secrets (deny by default):** never log workbook open **passwords**,
   database credentials, connection strings, API tokens, or any value under a
   reserved key. Reserved sensitive key names to scrub (case-insensitive):
   `password`, `passwd`, `pwd`, `secret`, `token`, `connection_string`,
   `conn_str`, `api_key`, `credential`, `auth`. Replace value with
   `"<redacted>"` and record the key name only.
3. **Macro arguments:** arguments are caller data. Log a **typed, length-capped**
   summary (type + truncated repr), not full arbitrary objects; never log COM
   references (they are non-serializable anyway). Arrays → element count +
   first-N element types.
4. **No environment secrets:** never read/echo env vars into audit records.
5. **Redaction is lossy-on-purpose:** the audit proves *what* happened and
   *that* it was approved, not the secret values.

### Not corrupting the STDIO protocol

- Audit writes go to a **dedicated file handle**, opened in append mode
  (`"a"`, encoding `utf-8`), **never** to `sys.stdout`/`stdout.buffer`.
- Do **not** use the `logging` root StreamHandler (defaults to stderr, which is
  acceptable, but stdout is forbidden). Use a dedicated
  `logging.FileHandler` → JSON formatter, **or** a small purpose-built
  `AuditLogger` that serializes a `BaseModel`/dict and writes one line per
  event with an explicit lock (the STA worker and anyio worker threads can both
  emit, so guard with a `threading.Lock` or a single-writer queue).
- **Prefer async-safe append:** since audit may be emitted from the STA worker
  thread and from API worker threads, funnel through a queue + single writer, or
  guard the append with a lock, so concurrent lines never interleave.
- File location must pass PathPolicy artifact-root checks (never an arbitrary
  path), so an attacker-controlled tool parameter cannot redirect audit to an
  arbitrary location.

---

## 5. Threat-model check (spec invariants as pure policy)

| Invariant (spec) | Implementable as pure policy? | Notes / pitfalls |
|---|---|---|
| Inspect mode never enables macros | **Yes** | `SessionService` sets `AutomationSecurity=ForceDisable` *immediately before* `Workbooks.Open` and restores original in `finally` (spec §10.1.1). MacroPolicy must only *record* the prior value. Pitfall: a naive impl leaves `AutomationSecurity=Low` (the app default) → silent macro enable. Must explicitly set ForceDisable per open. `[HIGH]` |
| Approval cannot be bypassed by path aliases | **Yes, if tied to canonical path + hash** | `PathPolicy.validate_path` already resolves symlink/reparse/UNC/traversal. Approval key MUST use its returned canonical path (and hash), never the raw string. Pitfall: keying on raw input lets `Report.xlsm` vs `.\sub\..\Report.xlsm` vs a junction resolve to different keys. `[HIGH]` |
| No Trust Center / registry writes | **Yes** | Pure prohibition: MacroPolicy exposes read-only diagnostics; a test asserts *no* call path reaches `AccessVBOM`/`VBAWarnings`/Trusted Locations. Pitfall: a "helpful" auto-enable of VBOM. `[HIGH]` |
| Deny-by-default | **Yes** | Default branch of `ApprovalPolicy.check` returns deny; allowlist/ grant are explicit allow cases. `[HIGH]` |
| Never bypass VBA project protection | **Yes** | `VBProject.Protection == vbext_pp_locked` → `VBA_PROJECT_LOCKED`, no SendKeys/registry. `[HIGH]` |
| Close without save_changes never saves | **Yes** | `excel_close` defaults `SaveChanges:=False`. Classify by effective flag (§1 #2). `[HIGH]` |
| Snapshot before mutation | **Yes** | Snapshot is READ_ONLY (§1 #4) so it is never gated behind the approval it protects; `SNAPSHOT_FAILED` blocks the mutation (task_014). `[MED]` |
| No arbitrary COM dispatch / shell / eval | **Yes (out of task_009 scope, but policy-adjacent)** | Macro name validated/normalized (§7.3), not free text; declarative test spec forbids Python/shell (§9.2). `[HIGH]` |

### Where a naive implementation would violate the spec

1. **Classifying `Close`/`AddFromString` as READ_ONLY** → silent mutation
   without approval. (Resolved in §1.)
2. **Leaving `AutomationSecurity` at its `Low` default** → inspect mode leaks
   macro execution. (Must set ForceDisable per open and restore in finally.)
3. **Keying approval on raw path text** → symlink/junction/`..` bypass.
   (Must key on `PathPolicy` canonical path + SHA-256.)
4. **Mapping every `Application.Run` failure to `MACRO_NOT_FOUND`** → hides
   compile errors and runtime errors. (Map by verified HRESULT; run task_005
   PoC to confirm the per-condition codes.)
5. **Auto-enabling VBOM / editing `AccessVBOM`** → violates §10/§7.2 and is
   reverted by Office anyway. (Diagnose only.)
6. **Logging raw params including passwords / connection strings** → violates
   §10 Secrets. (Mandatory redaction, §4.)
7. **Writing audit to stdout** → corrupts the STDIO JSON-RPC stream (A01).
   (Dedicated file handle + lock, §4.)
8. **Treating snapshot as a sandbox** → macros can still write DB/email/network
   (spec §9.4). The audit record must not imply side effects were undone.

### How to test deny-by-default (task_009 unit-test matrix)

- `classify` returns `READ_ONLY` for the read operations and the correct
  class for every ambiguous case in §1 (table-driven).
- `ApprovalPolicy.check` with **no** allowlist and **no** grant → `deny` for
  every MUTATING/EXECUTING action; **allow** for READ_ONLY.
- Expired grant (`now > expires_at`) → `deny`; session grant after
  `session_id` change → `deny`.
- Path-alias test: two distinct string spellings that resolve to the **same**
  canonical path → same approval key (one approval covers both); two files with
  the **same basename** but **different SHA-256** → distinct keys (approval for
  one does not cover the other).
- `redact` unit test: reserved keys replaced with `"<redacted>"`; non-secret
  params preserved; arrays length-capped.
- Audit smoke: events append to a temp file, are valid JSONL, and **stdout is
  byte-identical** before/after emitting an audit event (proves no stdout leak).
- Policy-invariant tests (static): assert no module references `AccessVBOM`,
  `VBAWarnings`, `TrustedLocations`, or writes to `HKCU/HKLM`; assert
  `AutomationSecurity` is set to `ForceDisable` before every `Open` in
  inspect path (can be checked via a fake adapter recording call order).

---

## 6. Test strategy (summary)

1. **Unit (task_009 acceptance):** allow/deny matrix, allowlist vs one-time vs
   session grant, expiry, deny-by-default, path-alias equivalence, same-basename
   different-hash separation, audit redaction, audit-to-file-no-stdout,
   `classify` table completeness.
2. **Static policy assertions:** no Trust Center / registry / MOTW writes
   anywhere in `security/`; `AutomationSecurity` restore in `finally`.
3. **Integration (task_012/013/023):** inspect open ⇒ macros disabled; execute
   open ⇒ approved context; close-without-save leaves file unchanged; macro
   run denied when unapproved, allowed + audited when approved; snapshot
   precedes mutation.
4. **Macro status disambiguation (task_005/022):** per-condition HRESULT capture
   → verify `MACRO_BLOCKED` / `MACRO_NOT_FOUND` / `MACRO_RUNTIME_ERROR`
   mapping is evidence-based, not generic.

---

## 7. References

**Project (verified primary source):**
- `.ai/plans/plan.md` task_009 (lines 457–508), R10 risk (line 38), A01/A08/A15.
- `.ai/tmp/spec_text.txt` §7.1 (318–323), §7.2 (324–331), §7.3 (332–338),
  §8.3 (378–383), §9.1 (386–391), §9.4 (410–412), §10 (413–433), §10.1 (435–439),
  §11 (440–466), §11.1 (468–474), A01/A08/A15 (479–523).
- `src/excel_vba_mcp/domain/statuses.py` — `OperationClass`, `SessionMode`,
  `MacroRunStatus`.
- `src/excel_vba_mcp/domain/errors.py` — `ErrorCode`, HRESULT map,
  `error_from_hresult`, `format_hresult`.
- `src/excel_vba_mcp/domain/models.py` — `DomainModel`, `TimestampedModel`,
  `WorkbookInfo`, `MacroRunResult`, `CompileDiagnostic`, `SnapshotManifest`.
- `src/excel_vba_mcp/security/path_policy.py` — `PathPolicy.validate_path`,
  `ARTIFACT_ROOT_NAMES`.

**Microsoft Learn (verified):**
- Application.AutomationSecurity (Excel) —
  https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity
- MsoAutomationSecurity enumeration —
  https://learn.microsoft.com/en-us/office/vba/api/office.msoautomationsecurity
- Workbook.VBProject (Excel) —
  https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbproject
- vbext_ProjectProtection enum (vbext_pp_none=0 / vbext_pp_locked=1) —
  https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.vbext_projectprotection
- Workbook.HasPassword (Excel) —
  https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.haspassword
- Programmatic access to VBA project object model (AccessVBOM, HKCU\...\Security,
  reverted on exit) — MS Learn archive:
  https://learn.microsoft.com/en-us/archive/blogs/cristib/vba-how-to-programmatically-enable-access-to-the-vba-object-model-using-macros
- "Programmatic access to Visual Basic Project is not trusted" (AccessVBOM note) —
  https://www.renenyffenegger.ch/notes/development/languages/VBA/Programmatic-access-to-Visual-Basic-Project-is-not-trusted

**Confidence statement:** The AutomationSecurity, VBProject.Protection,
Workbook.HasPassword, and AccessVBOM-registry-revert behaviors are **HIGH**
(primary Microsoft Learn). The operation-class table is **HIGH** for unambiguous
cases and **MED** for the five flagged ambiguities (Close/save_changes,
SaveCopyAs, AddFromString, OnAction, compile/test-run), each with an explicit
rationale and a task_009 test hook to lock the decision.