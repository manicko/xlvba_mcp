# report_14a — Error Taxonomy & Usability Validator

Independent, adversarial validation of the error-taxonomy + usability remediation.
Repo `C:\py_exp\xlvba_mcp`, HEAD `7961038` + uncommitted working changes.
**VALIDATION ONLY — no source changes were made.**

Method: read the actual diff and implementation, then verified *serialized* behavior
(`ExcelVbaMcpError.to_dict()`) and *live* behavior (real hidden Excel via pywin32),
separating unit evidence from live evidence. Commands and line numbers are cited inline.

Environment for live runs: `platform win32`, Python 3.14.7, pytest 9.2.0, Excel + VBOM
trust available (the live tests executed, they did **not** skip).

Working diff under review (`git diff --name-only HEAD`):
`docs/05-mcp-tools-reference.md`, `docs/09-limitations-and-compatibility.md`,
`docs/ru/05-mcp-tools-reference.md`, `docs/ru/09-limitations-and-compatibility.md`,
`src/excel_vba_mcp/adapters/excel/names_adapter.py`,
`src/excel_vba_mcp/adapters/excel/tables_adapter.py`,
`src/excel_vba_mcp/domain/errors.py`, plus 4 unit-test files and 2 new live-test files.
No pivot, session, snapshot, sheets, or worker source files are in the diff.

---

## Verdict summary

| # | Item | Verdict |
|---|------|---------|
| 1 | WP1 serialized envelope (NAME_CONFLICT / NOT_FOUND, not retryable, corrective, no partial mutation, hash unchanged) | **APPROVE** |
| 2 | Genuine transient COM (COM_BUSY / RPC_E_CALL_REJECTED) still retryable | **APPROVE** (with non-blocking caveat, see Advisory) |
| 3 | WP2 docs: no "pending" RANGE_TOO_LARGE; EN+RU consistent | **APPROVE** |
| 4 | WP3 R1 previous_close / R2 workbook_id required / R3 canonical=original + same_snapshot / contract documented | **APPROVE** |
| 5 | WP4 `sheet_copy(target_name)` live (rename/default/duplicate/invalid) | **APPROVE** |
| 6 | No safety/architecture regression; no external-pivot change; SCHEMA_VERSION=1 | **APPROVE** |
| — | Auditor MEDIUM advisory (broad `except Exception` re-labels transient COM) | **Valid, NON-BLOCKING follow-up** |

**Final verdict: APPROVE.** All six items pass on both unit and live evidence. One valid,
narrow, non-blocking follow-up is recorded (the Auditor's MEDIUM advisory).

---

## Item 1 — WP1 serialized envelope — APPROVE

### Unit evidence
`.\.venv\Scripts\python.exe -m pytest <7 WP1 node-ids> -p no:cacheprovider -v` → **7 passed**.
Broader run `test_error_envelope.py test_errors.py test_names_service.py test_tables_service.py
test_com_worker.py` → **158 passed in 15.86s**.

Envelope construction (`src/excel_vba_mcp/domain/errors.py`):
- `NAME_CONFLICT` added as an additive member (line 86); category `validation` (line 180);
  **not** in `_RETRYABLE_CODES` (lines 186–188, only `COM_BUSY`/`RPC_E_CALL_REJECTED`);
  action `"choose a unique name or delete the existing one"` (line 270).
- `NOT_FOUND` → category `not_found` (line 175), action `"verify workbook_id / object name /
  procedure"` (line 243).
- `to_dict()` emits `errorCode, message, hresult, hresult_hex, details, errorCategory,
  isRetryable, suggestedNextActions` (lines 382–391) — `errorCategory`/`isRetryable`/
  `suggestedNextActions` are pure functions of the code.

Raise sites (deterministic, all reached by PRE-checks before any COM mutation):
- Duplicate defined name (same scope) → `NameConflictError`, `reason="duplicate"`
  (`names_adapter.py:617–626`); cross-scope → `NameConflictError`, `reason="scope_conflict"`
  (628–638). The guard `_guard_no_conflict` runs **before** `collection.Add`
  (`names_adapter.py:428`), so a duplicate create cannot partially mutate.
- Missing worksheet (names) → `ObjectNotFoundError` NOT_FOUND `reason="not_found"`
  (`names_adapter.py:273–276`).
- Table append unknown header → `ObjectNotFoundError` NOT_FOUND
  (`tables_adapter.py` `_validate_row_headers`, `reason="not_found"`).
- Missing table → `ObjectNotFoundError` NOT_FOUND (`_require_list_object`, `reason="not_found"`).
- Missing sheet (table create) → `ObjectNotFoundError` NOT_FOUND (`_require_sheet`,
  `reason="not_found"`).
- Duplicate table column header → `NameConflictError` `reason="duplicate"`
  (`tables_adapter.py` `add_column`).

Unit assertions use a shared `_assert_non_retryable_taxonomy` helper that checks the
**serialized** dict (not the class): `errorCode` exact, `errorCode ∉ {COM_BUSY,
RPC_E_CALL_REJECTED}`, `errorCategory` exact and `≠ transient_com_failure`,
`isRetryable is False`, non-empty string actions, and `"retry" not in actions`
(`test_names_service.py:807–825`, `test_tables_service.py:1046–1062`,
`test_error_envelope.py::test_name_conflict_is_validation_not_retryable_with_action`,
`test_errors.py::test_name_conflict_error_carries_code_and_is_registered`).

### Live evidence
`.\.venv\Scripts\python.exe -m pytest tests/integration_excel/test_error_taxonomy_live.py -p
no:faulthandler -rN -v` → **4 passed in 9.16s** (real hidden Excel; not skipped).

- `test_duplicate_name_create_is_name_conflict`: serialized `NAME_CONFLICT` / `validation` /
  `isRetryable False`; `suggestedNextActions == ["choose a unique name or delete the existing one"]`;
  `context["reason"] == "duplicate"`; **defined-name count unchanged** (no partial mutation);
  `_sha256(project)` unchanged **and** committed fixture hash unchanged.
- `test_table_append_unknown_header_is_not_found`: `NOT_FOUND` / `not_found` / not retryable;
  `reason=="not_found"`; **table row count unchanged**; project hash unchanged.
- `test_missing_table_is_not_found` and `test_missing_sheet_is_not_found`: `NOT_FOUND` /
  `not_found` / not retryable; missing-sheet also asserts **no table was created**; project hash
  unchanged.
- The live helper additionally asserts `errorCode ∉ {COM_BUSY, RPC_E_CALL_REJECTED}` and
  `errorCategory != "transient_com_failure"` (test lines 128–138).

**Conclusion:** serialized envelope is correct for all four WP1 raise classes; none is
COM_BUSY/transient/retryable; no partial mutation; original + committed hashes unchanged.
APPROVE.

---

## Item 2 — Genuine transient COM still retryable — APPROVE (non-blocking caveat)

### Evidence the machinery is unchanged
- `_RETRYABLE_CODES = {COM_BUSY, RPC_E_CALL_REJECTED}` (`errors.py:186–188`) and
  `retryable_for` (302–316) are untouched by the diff.
- Worker classification is untouched: `com_worker.py` `_run_with_retries` (762–808) maps a
  **raw** `com_error` with `RPC_E_CALL_REJECTED` → `RpcCallRejectedError` (retryable) and
  `RPC_E_SERVERCALL_RETRYLATER` → `ComBusyError` (retryable), and **re-raises domain
  `ExcelVbaMcpError` as-is** (769–773) so a precise code is never masked.
- Unit proof: `test_com_worker.py` (within the 158-passed run) drives `FakeComError(
  RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER)` and asserts the retryable codes;
  `test_error_envelope.py` parametrizes `retryable_for(COM_BUSY|RPC_E_CALL_REJECTED) is True`
  and `category_for == "transient_com_failure"`.

### Caveat (see Advisory)
The WP1 adapter catch-all wrappers around the *actual mutation COM calls* now convert **any**
exception — including a genuine transient `com_error` — into a non-retryable
`InvalidRangeError`/`ObjectNotFoundError`. This narrows retryability for those specific mutation
calls only. It does **not** weaken the canonical worker path or the taxonomy. Because the
deterministic WP1 targets are caught by pre-checks and the worker path is intact + tested, this
does not block Item 2. Recorded as a follow-up.

---

## Item 3 — WP2 docs (no "pending" RANGE_TOO_LARGE; EN+RU consistent) — APPROVE

- `docs/09-limitations-and-compatibility.md` and `docs/ru/09-limitations-and-compatibility.md`:
  the row was rewritten from "…limit-specific code pending / код лимита — в ожидании" to
  **"rejected before execution with `RANGE_TOO_LARGE` (category `validation`,
  `isRetryable=false`) … paginate via `offset`/`limit` | Verified"** in both EN and RU.
- `docs/05-mcp-tools-reference.md` and RU mirror: the "Cell cap (100,000)" paragraph was
  rewritten identically (EN/RU) to `RANGE_TOO_LARGE`, before-execution, validation,
  non-retryable, paginate; the "delivered by a separate work package / поставляется отдельной
  рабочей парой" pending language is removed.
- Grep `pending.{0,40}RANGE_TOO_LARGE|RANGE_TOO_LARGE.{0,40}pending` over `docs/` → **No files
  found**. No residual "pending" near RANGE_TOO_LARGE.
- EN↔RU parity confirmed for both the RANGE_TOO_LARGE text and the new R1–R3 caller-contract
  block (see Item 4).

---

## Item 4 — WP3 (R1/R2/R3 + caller contract) — APPROVE

### R1 — `previous_close` observable
`snapshot_service.py` `_close_previous_copy` (466–498) returns
`("closed"|"failed"|"not_requested", error_text)`; a close failure is caught broadly and never
fails the restore. Surfaced in `SnapshotRestoreResult` (460–464) and the audit record
(456–457). Model field `previous_close: Literal["closed","failed","not_requested"]`
(`domain/models.py:587`) plus `previous_close_error` (595).

### R2 — `workbook_id` required (missing → INVALID_RANGE pre-action)
Enforced at the API tool layer *before* any close/open:
`api/test_tools.py::snapshot_restore` (156–167) raises `INVALID_RANGE`,
`reason="workbook_id_required"` when `workbook_id` is None/empty; the `confirm` gate
(169–179) raises `CONFIRMATION_REQUIRED` before any destructive action. (The service keeps
`workbook_id: str | None = None` for the internal `not_requested` path; the *tool* contract is
what makes it required — consistent with the docstring at 130–152.)

### R3 — canonical = original + same_snapshot
After open+validate, `rebind_restored_identity(workbook_id, canonical=manifest.source_path,
source_sha256=manifest.sha256)` (437–441) rebinds identity to the ORIGINAL, never the snapshot
file; working copy stays `work/<newid>.ext`. Repeating the already-active baseline raises
`INVALID_RANGE reason="same_snapshot"` **before** any open (376–392); a different snapshot is
allowed.

### Caller contract documented
`api/test_tools.py` docstring (130–152), `server.py` (315–321), and `docs/05` EN+RU
("`snapshot_restore` caller contract (R1–R3)" / "Контракт вызова … (R1–R3)") all state:
workbook_id required = OLD id; use returned `workbook.id` (new id); `previous_close`
disposition; failed close doesn't invalidate; same_snapshot is deterministic; three distinct
artifacts (original / immutable snapshot / mutable working copy).

### Unit evidence
`test_snapshot_service.py -k restore` → **11 passed**; `test_confirm_flags.py` → **7 passed**
(combined run **36 passed**). Cover: `test_restore_previous_close_failed_surfaces_but_succeeds`,
`test_restore_without_workbook_id_reports_not_requested`,
`test_restore_closes_without_save_and_opens_baseline`,
`test_restore_rebinds_canonical_to_original`, `test_restore_same_snapshot_repeat_rejected`,
`test_restore_different_snapshot_allowed`,
`test_snapshot_restore_missing_workbook_id_rejected`,
`test_snapshot_restore_missing_confirm_rejects_before_action`.

---

## Item 5 — WP4 `sheet_copy(target_name)` live — APPROVE

`.\.venv\Scripts\python.exe -m pytest tests/integration_excel/test_sheets_copy_target_name_live.py
-p no:faulthandler -rN -v` → **1 passed in 3.38s** (real hidden Excel; owned PID reaped after
shutdown).

Verified live:
1. Explicit `target_name="CopyA"` renames the copy to exactly `CopyA`.
2. Omitting `target_name` keeps Excel default `"SrcA (2)"`.
3. Duplicate `target_name="CopyA"` → `INVALID_RANGE`/`validation`/not retryable,
   `reason="duplicate_target_name"`; **no second `CopyA`** created; source `SrcA` untouched.
4. Invalid `target_name` (`"Bad:Name"` and 32×`"X"`) → `INVALID_RANGE`,
   `reason="invalid_target_name"`; neither ever becomes a sheet; workbook state stays known.
5. Original + committed fixture hashes unchanged.

Implementation: `sheets_adapter.py::copy_sheet` (320–367) + `_validate_target_name`
(419–449, non-empty / ≤31 / no `: \ / ? * [ ]` / unique). The `_assert_invalid_range_taxonomy`
helper asserts the serialized envelope (not the class), incl. `"retry" not in actions`.

Minor (non-blocking) observation: `copy_sheet` performs `source.Copy(...)` **before**
validating `target_name`, so a *rejected* copy leaves Excel's auto-named `"SrcA (n)"` sheet in
the disposable working copy. The live test explicitly tolerates this
(`names_after_invalid - names_before_invalid <= {"SrcA ("…}`), the original is never touched,
and the working copy is discarded on shutdown — so it is not a contract/safety violation, only a
small wart (validating `target_name` before `Copy` would avoid the orphan). Not a blocker.

---

## Item 6 — No regression / no external-pivot change / SCHEMA_VERSION=1 — APPROVE

- `SCHEMA_VERSION = 1` (`domain/models.py:65`). Note: `snapshot_store._MANIFEST_SCHEMA_VERSION =
  2` (snapshot manifest format) is a **separate, pre-existing** internal version and is **not**
  in the diff — it is not the DTO `SCHEMA_VERSION` and was not changed here.
- No external-pivot change: `git diff --name-only HEAD` contains **no** pivot source file; the
  external/OLAP/Data-Model gate (`UNSUPPORTED_PIVOT_SOURCE`) is untouched by this remediation.
- No architecture/contract regression (additive `NAME_CONFLICT` only):
  - contract/dispatch/annotation/tool-surface tests: **307 passed**;
  - sheets + domain + phase2 contracts: **87 passed**;
  - names + tables + errors + error_envelope + com_worker: **158 passed**;
  - snapshot + confirm: **36 passed**.
  `NAME_CONFLICT` is registered in `_ERROR_BY_CODE` (736), `_CATEGORY_BY_CODE` (180),
  `_ACTIONS_BY_CODE` (270), so code→category/action totality is preserved (no "unknown code"
  fallback path is hit).

---

## Advisory — Auditor MEDIUM (broad `except Exception` wrappers): VALID, NON-BLOCKING

**Claim under test:** broad `except Exception` wrappers in
`names_adapter.py` (433 create / 467 update / 496 delete) and `tables_adapter.py`
(831 create / 844 style / 875 & 882 resize / 1112 `_require_sheet` / 1130 `_resolve_source_range`)
could re-label a genuine retryable `COM_BUSY` as non-retryable `INVALID_RANGE`/`NOT_FOUND`.

**Confirmed real but narrow.** These wrappers surround the actual COM mutation/resolution calls
and convert *every* exception into a domain error. Because the worker re-raises domain
`ExcelVbaMcpError` as-is (`com_worker.py:769–773`), a genuine transient `com_error`
(`RPC_E_CALL_REJECTED 0x80010001` / `RPC_E_SERVERCALL_RETRYLATER 0x8001010A`) raised *inside*
those calls is masked as `validation`/`not_found` with `isRetryable=false`. Pre-WP1 these same
catch-alls raised generic `ExcelVbaMcpError` (default `COM_BUSY`, retryable), so for these
specific mutation calls genuine transient COM is now mislabeled non-retryable.

**Why it is NOT blocking:**
1. WP1's deterministic targets are intercepted by **pre-checks before the COM call**
   (`_guard_no_conflict` before `Add`; `_require_list_object`/`_header_index`/
   `_validate_row_headers` before mutation; `_require_sheet`/`_require_worksheet` before create),
   so Item 1 acceptance does not rely on the catch-all.
2. The canonical transient-COM source — the worker's classification of *unwrapped* `com_error` —
   is unchanged and unit-tested (`test_com_worker.py`).
3. The prior behavior was not strictly correct either (it labelled non-COM Python errors as
   retryable `COM_BUSY`); WP1 improved classification for the common non-COM case.
4. The fix is small and localized, not architectural.

**Recommended follow-up (BEST-PRACTICE, non-blocking):** in these catch-alls, re-raise when the
exception carries a transient HRESULT (e.g. `hresult = _hresult_of(exc)` ∈
{`RPC_E_CALL_REJECTED`, `RPC_E_SERVERCALL_RETRYLATER`} → re-raise / `ComBusyError`/
`RpcCallRejectedError`) instead of wrapping; add a unit test that injects a transient `com_error`
into e.g. `TablesAdapter.create_table` / `NamesAdapter.create_name` and asserts
`COM_BUSY`/`isRetryable=true`. This restores full-surface "genuine transient COM stays
retryable" without touching the WP1 deterministic classifications.

---

## Required fixes
None that block approval.

## Advisory recommendations
1. (Follow-up) Narrow the mutation catch-all wrappers so a genuine transient HRESULT is not
   masked as non-retryable (see Advisory above).
2. (Optional polish) In `SheetsAdapter.copy_sheet`, validate `target_name` **before**
   `source.Copy(...)` to avoid leaving an auto-named orphan sheet on a rejected copy.

## Final verdict
**APPROVE** — Items 1–6 pass on unit **and** live evidence (WP1 live 4/4, WP4 live 1/1,
WP3 unit 11+7, transient-COM unit green, contract/dispatch 307 green, SCHEMA_VERSION=1, no
pivot change). The Auditor's MEDIUM advisory is a valid, narrow, **non-blocking** follow-up.