# task_009 — Independent Validator Report

- **Date:** 2026-11-19
- **Validator role:** read-only verification (source NOT modified; tests executed)
- **Commit under review:** `72ddf93` — `feat(security): add approval policy, macro policy and audit trail (task_009)`
- **Binding design:** `.ai/audit/task_009-research-2026-11-19.md`
- **Plan:** `.ai/plans/plan.md` task_009 (acceptance criteria)
- **Repo:** `C:\py_exp\xlvba_mcp`

---

## 1. Scope

Independent validation of the three new security modules and their tests against
the research design and plan acceptance criteria:

- `src/excel_vba_mcp/security/macro_policy.py` (316 lines)
- `src/excel_vba_mcp/security/approval.py` (256 lines)
- `src/excel_vba_mcp/security/audit.py` (262 lines)
- `tests/unit/test_macro_policy.py`, `tests/unit/test_approval.py`,
  `tests/unit/test_audit.py`, `tests/unit/test_security_policy_static.py`

Commit `72ddf93` adds **only** these 7 files (verified via `git show --stat`);
no unrelated source was touched. Working tree is clean at commit `3e9ccc3`
(child of `72ddf93`), so the reviewed content equals the committed content.

Checks performed: deny-by-default, approval identity (canonical path + sha256
authority, procedure wildcard/exact), grant lifecycle (one-time/session/expiry),
macro-policy purity, macro-failure classification, operation classification of
the research §1 ambiguous cases, audit redaction / JSONL validity / stdout
isolation / thread-safety, static policy test efficacy, anti-gaming, and spec §10
invariants.

---

## 2. Actual results — commands and outcomes

| Command | Outcome |
|---|---|
| `uv run pytest tests/unit -q` | **145 passed** in 0.43s |
| `uv run ruff check src tests/unit` | **All checks passed!** |
| `uv run ruff format --check src tests/unit` | **27 files already formatted** |
| `uv run mypy src` | **Success: no issues found in 20 source files** |

Targeted behavioral probes (read-only, run via `uv run python -c`):

```
SESSION grant, check session_id=None -> True          # WARNING W1
SESSION grant, check session_id=s2   -> False
ONE_TIME first -> True | second -> False              # consumption OK
unknown tool class -> executing                       # fail-closed OK
```

Audit write-path probe (temp file) confirmed `params.password -> "<redacted>"`
and `params.args -> {arg_count,args}`, while `approval` / `result` / `workbook`
payloads are written verbatim (see W2). stdout isolation is proven by
`test_audit_writes_nothing_to_stdout` (MARKER before/after unchanged) and by code
inspection: `AuditLogger` writes only to its own append-mode file handle, never
`sys.stdout`.

Static-test efficacy proof (no source modified): applying the exact regexes from
`test_security_policy_static.py` to the real `macro_policy.py` yields
`winreg=False, AutomationSecurity==False`; appending `import winreg` +
`Application.AutomationSecurity = 1` to an in-memory copy yields both `True`. The
test reads files fresh from disk (`path.read_text()`), so it genuinely scans the
source and would fail if a forbidden pattern were introduced.

---

## 3. Verification detail (per required check)

### 3.1 Deny-by-default — PASS
`ApprovalPolicy.check` (`approval.py:177-237`): `READ_ONLY` short-circuits to
`allowed=True` (`:197-203`); `MUTATING`/`EXECUTING` with no allowlist and no live
grant reach the **terminal** `return ApprovalDecision(allowed=False, ... "denied
by default")` (`:232-237`). Tests `test_deny_by_default_mutating`,
`test_deny_by_default_executing` confirm. `operation_class=None` also falls
through to deny (fail-closed). This is the real terminal branch, not a decoy.

### 3.2 Approval identity — PASS (see W3)
`_identity_matches` (`approval.py:111-116`): hash is authoritative when **both**
sides carry it; else canonical-path fallback.
- Same basename, different sha256 → distinct keys (`test_same_basename_different_sha256_distinct_keys`).
- Same canonical path, different textual spelling → same key
  (`test_same_canonical_path_different_spellings_same_key`, uses `Path.resolve()`).
- Procedure wildcard (`None` entry) vs exact (`_procedure_matches`, `:119-124`;
  `test_procedure_wildcard_and_exact`).
- `canonical_path` is a caller contract anchored on `PathPolicy.validate_path`
  (never raw input) — enforced by contract, not re-canonicalized here (W3).

### 3.3 Grant lifecycle — PASS
- One-time consumed on first match (`approval.py:222-224`; `test_one_time_grant_consumed`:
  first True, second False — reproduced).
- Session grant valid within same session, denied after session change
  (`_grant_is_live` `:240-256`; `test_session_grant_valid_within_session`,
  `test_session_grant_denied_after_session_change`).
- Session and one-time expiry both deny (`test_session_grant_expiry_denies`,
  `test_one_time_grant_expiry_denies`). **Caveat W1** below.

### 3.4 Macro-policy purity — PASS
`macro_policy.py` imports only `domain.errors` and `domain.statuses`; no `win32*`,
`pythoncom`, `winreg`, `subprocess`, `Shell`, `SendKeys`, or `AutomationSecurity`
assignment (grep clean; static test enforces). `automation_security_for`
(`:215-228`): `INSPECT→FORCE_DISABLE(3)`, `EXECUTE→BY_UI(2)`, `LOW` never
returned (`test_automation_security_never_low`). Value selection is pure;
application/restore is deferred to the session service (task_012) per design.

### 3.5 Macro-failure classification — PASS (see W4)
`classify_macro_failure` (`:264-301`): `approved=False → BLOCKED`;
`0x800A03ED → NOT_FOUND`; `0x800A03EC`/`0x80004005 → RUNTIME_ERROR`; `None` and
unknown codes → `NEEDS_RECOVERY`. **Never returns `SUCCESS`** (verified:
`test_classify_macro_failure_never_success` over `None,0x0,0x800A03EC,0x800A03ED,
0x80004005,0xDEADBEEF`). No false-success path. Office-side block vs not-found is
a known PoC-dependent nuance (W4).

### 3.6 Operation classification (research §1) — PASS
`classify_operation` (`:178-212`) + tables (`:78-167`):
- `excel_close` no `save_changes` → `READ_ONLY`; truthy `save_changes` (bool or
  `"true"/"1"/"yes"/"on"`) → `MUTATING` (`test_excel_close_*`). ✔
- `snapshot_create` → `READ_ONLY` (SaveCopyAs) ✔
- `vba_write`/`AddFromString` family → `MUTATING` ✔
- `control_set_on_action`/`shape_set_on_action` → `MUTATING` (not EXECUTING) ✔
- `compile`/`vba_compile`/`macro_run`/`test_run`/`test_run_suite` → `EXECUTING` ✔
- Unknown/empty tool → `EXECUTING` fail-closed (`test_unknown_tool_fails_closed_as_executing`). ✔

### 3.7 Audit — PASS (see W2)
- Redaction (`audit.py:41-148`): reserved keys case-insensitive → `"<redacted>"`,
  path-like → basename, `args`/`arguments` → typed length-capped summary, nested
  mappings/lists scrubbed recursively. On the **actual write path**
  (`record` → `redact(params)` → `log`), not a dead helper. Tests
  `test_redact_*` confirm.
- JSONL validity: `test_jsonl_lines_parse`, `test_append_mode_preserves_existing_lines`.
- **stdout never written:** dedicated append-mode file handle, no `sys.stdout`,
  no `logging` StreamHandler; `test_audit_writes_nothing_to_stdout` proves it.
- Thread-safe: `threading.Lock` around write+flush; 16-thread/64-event test
  confirms no interleaving and no loss (`test_concurrent_writes_do_not_interleave`).

### 3.8 Anti-gaming — PASS
- No `assert True` / no-op tests in the task_009 suite (grep; only unrelated
  `pass` in `test_path_policy.py` exception body).
- Deny-by-default is the real terminal branch (§3.1).
- Redaction is on the write path (§3.7).
- Audit logger truly avoids stdout (§3.7).
- Static policy test scans live source and demonstrably fails on injected
  forbidden patterns (§2).

### 3.9 Spec §10 invariants — PASS
- Inspect mode cannot enable macros: `automation_security_for(INSPECT)=FORCE_DISABLE`,
  never `LOW`; no `AutomationSecurity` assignment possible in the module (static test).
- No path-alias approval bypass at the policy layer: keyed on canonical path +
  sha256; alias equivalence delegated to `PathPolicy.validate_path` contract (W3).
- No project-protection bypass: `diagnose_protection(1) → VbaProjectLockedError`;
  no SendKeys/registry path anywhere.

---

## 4. Defects

No Critical or High defects. Low / advisory items (non-blocking):

- **W1 (Low, defense-in-depth) — `approval.py:240-256` (`_grant_is_live`).**
  A `SESSION`-scoped grant bound to `session_id="s1"` is still accepted when
  `check(...)` is called with `session_id=None`
  (probe: `-> True`). The session binding is only enforced when *both*
  `bound_session` and the incoming `session_id` are non-`None`. A stricter
  deny-by-default posture would reject a bound session grant when no session id
  is presented. Repro:
  ```python
  p=ApprovalPolicy(now=lambda:T0)
  p.grant(grant(scope=SESSION), session_id="s1")
  p.check(key, session_id=None, operation_class=MUTATING).allowed  # -> True
  ```
  Impact depends on the API layer (task_012) always passing `session_id`;
  session binding is not in the task_009 acceptance criteria. Recommend the
  session service always supply `session_id` and/or tighten `_grant_is_live` to
  treat `bound_session set + session_id None` as a mismatch.

- **W2 (Low, defense-in-depth) — `audit.py:224-229` (`AuditLogger.record`).**
  Only `params` is passed through `redact()`; `approval`, `result`, `workbook`
  (and `user`/`process`/`excel_version`) are serialized verbatim
  (`dict(approval)`, `dict(result)`, `_normalize_workbook`). Probe confirmed a
  `secret` key placed in the `approval` dict is written in clear
  (`"approval": {"secret": "leak?"}`). Per the research §4 schema these payloads
  are meant to carry only non-secret fields (grant_id/scope/expires_at,
  error_code/hresult, canonical_path/file_name/sha256), so this is not a schema
  violation; it is an unguarded-by-construction gap. Recommend applying
  `redact()` (or a shared `_scrub`) to these sub-objects too for defense-in-depth.

- **W3 (Advisory) — approval canonical-path trust boundary.**
  `approval.py` does not re-canonicalize `canonical_path`; it relies on callers
  passing `PathPolicy.validate_path` output (documented at `:14-20`). There is no
  runtime assertion that the supplied path is canonical. Path-alias equivalence
  and anti-traversal guarantees therefore hold only if the API layer (task_012)
  always routes through `PathPolicy`. This matches the design; noted as a
  cross-task integration obligation, not a task_009 defect.

- **W4 (Advisory, deferred by design) — Office-side macro block vs NOT_FOUND.**
  When policy approves (`approved=True`) but Office blocks the run (inspect-mode
  `ForceDisable`, Trust Center), `Application.Run` can surface the same
  `0x800A03ED` that also means "macro not found", so `classify_macro_failure`
  would return `NOT_FOUND` rather than `BLOCKED`. Research §3 flags this as
  MED/PoC-dependent and assigns per-condition HRESULT proof to task_005. Policy
  denials (unapproved) are correctly `BLOCKED`. Acceptable within task_009 scope.

- **Note (not a defect) — `classify_macro_failure` never returns `SUCCESS`.**
  By design (`:264-301`): a bare no-exception return is never a success. Success
  is determined by the caller from other evidence. Verified no input produces a
  false success.

- **Note (not a defect) — unhashed request matches hash-keyed entry by path.**
  `_identity_matches` (`:111-116`): if the request omits `sha256` but the entry
  has one, matching falls back to `canonical_path`. This is the documented
  "path is fallback" behavior; it means the caller must always compute the file
  hash to get the strong content-identity guarantee. Callers that omit the hash
  weaken the check to path-only.

---

## 5. Unmet acceptance criteria

Plan task_009 acceptance criteria — **all met**:

| Criterion | Status | Evidence |
|---|---|---|
| unapproved mutation/execution is denied | MET | `approval.py:232-237`; `test_deny_by_default_*` |
| approved action is auditable with workbook identity and result | MET | `audit.py:189-232`; `test_jsonl_lines_parse` (workbook+result) |
| inspect mode does not enable macros | MET | `macro_policy.py:215-228` (FORCE_DISABLE, never LOW); static test |
| Trust Center/registry are never modified | MET | purity grep + `test_security_policy_static.py` (winreg/AccessVBOM/VBAWarnings/TrustedLocations) |
| VBA project protection is never bypassed | MET | `diagnose_protection(1) → VbaProjectLockedError`; no bypass path |

Tests-to-run (`allow/deny, audit redaction, workbook hash matching`) present and
passing.

---

## 6. Warnings

- **W1/W2** are genuine low-severity defense-in-depth gaps (session-binding when
  `session_id` omitted; non-redaction of `approval`/`result`/`workbook`
  sub-objects). Neither violates a documented schema or acceptance criterion.
- **W3/W4** are cross-task obligations (path canonicalization in API layer;
  per-condition HRESULT PoC in task_005), correctly deferred by the design.
- Architectural note: macro trust *enforcement* (setting ForceDisable before
  `Open`, restoring in `finally`) is intentionally out of task_009 and belongs to
  the session service (task_012). task_009 provides only the pure decision; this
  is consistent with the approved architecture and the static test prevents the
  security modules from ever applying the setting themselves.

---

## 7. Required fixes

None blocking. Recommended (non-mandatory) hardening:
1. Treat a `SESSION` grant with a bound `session_id` as invalid when
   `check()` is called without a `session_id` (W1).
2. Run `approval`/`result`/`workbook` sub-objects through redaction before write
   (W2).

---

## 8. Advisory recommendations

- Add a unit test asserting a bound session grant is denied when `session_id` is
  omitted (locks W1 behavior).
- Add a test asserting reserved keys nested inside `approval`/`result` are
  redacted (locks W2).
- Document in the API layer (task_012) the mandatory
  `PathPolicy.validate_path` → `ApprovalKey.canonical_path` hand-off so W3 cannot
  regress.

---

## Final result: **PASS WITH WARNINGS**

All five plan acceptance criteria are met; the full unit suite (145 tests),
ruff lint, ruff format, and mypy all pass. No Critical or High defects were
found. The two warnings (W1 session-binding when `session_id` omitted; W2
non-redaction of non-`params` sub-objects) are low-severity defense-in-depth gaps
that do not violate any documented schema or acceptance criterion, plus two
advisory cross-task items (W3, W4) correctly deferred by the design.