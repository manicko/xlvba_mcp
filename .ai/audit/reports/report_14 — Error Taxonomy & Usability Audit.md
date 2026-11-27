# Report 14 — Error Taxonomy & Usability Audit

**Auditor role:** Independent architecture auditor (read-only).
**Repo:** `C:\py_exp\xlvba_mcp`
**HEAD:** `7961038` (`docs(xlvba): R1-R3 plan/report, research B/D + decisions, snapshot hardening docs`) + WP1 / WP2 / WP3-doc / WP4 in the working tree.
**Date:** 2026-11-27
**Scope:** WP1 (error taxonomy in `names_adapter.py` / `tables_adapter.py`), WP2 (docs/09 + RU `RANGE_TOO_LARGE`), WP3 (R1/R2/R3 snapshot-restore hardening + caller contract docs), WP4 (`sheet_copy(target_name)`), plus test execution and "no unrelated change" checks.
**Constraint honored:** No source files modified. Only this report was written.

---

## Verdict summary

| # | Item | Result |
|---|------|--------|
| 1 | WP1 — no bare `ExcelVbaMcpError(` on deterministic paths; correct non-retryable codes; `NAME_CONFLICT` additive; transient COM unchanged; taxonomy/default/`SCHEMA_VERSION` unchanged | **PASS** (1 MEDIUM advisory) |
| 2 | WP2 — docs/09 (+RU) `RANGE_TOO_LARGE` no longer "pending"; consistent with 05/08; reject-before-execution, non-retryable, split/paginate, no silent truncation | **PASS** |
| 3 | WP3 — R1/R2/R3 implemented (`previous_close`, required `workbook_id`, canonical rebind, `same_snapshot`) + caller contract in docs/05 (+RU) | **PASS** |
| 4 | WP4 — `sheet_copy(target_name)` forwards + validates (source-verified) + live test added | **PASS** |
| 5 | Tests — `tests/unit` and the two live files; no deterministic case is `COM_BUSY`/retryable | **PASS** (1483 unit, 5 live) |
| 6 | No unrelated changes; no new worker; no external-pivot change | **PASS** |

**Final: ACCEPT** (one non-blocking MEDIUM advisory recorded under WP1).

---

## WP1 — Error taxonomy in `names_adapter.py` / `tables_adapter.py` — PASS

### 1a. No bare `ExcelVbaMcpError(` without `code=` remains (grep)
`grep "ExcelVbaMcpError(" src/excel_vba_mcp/adapters/excel/names_adapter.py src/excel_vba_mcp/adapters/excel/tables_adapter.py` → **0 matches**. Both files now import and raise only typed subclasses:
- `names_adapter.py:47-51` imports `InvalidRangeError, NameConflictError, ObjectNotFoundError`.
- `tables_adapter.py:44-48` imports the same three.

All 12 raise sites in `names_adapter.py` and all 35 raise sites in `tables_adapter.py` use a typed subclass (verified by enumerating `raise \w+Error(`). Neither file contains any `code=` argument, so no deterministic site can resolve to the base default `COM_BUSY`.

### 1b. Semantically correct, non-retryable codes on deterministic paths
Independent probe (not relying on the test-suite alone):
```
NameConflictError   -> NAME_CONFLICT  validation   isRetryable=False  ['choose a unique name or delete the existing one']
ObjectNotFoundError -> NOT_FOUND      not_found    isRetryable=False  ['verify workbook_id / object name / procedure']
InvalidRangeError   -> INVALID_RANGE  validation   isRetryable=False  ['fix arguments to match input schema']
base ExcelVbaMcpError -> COM_BUSY    (retryable=True)  # unchanged default
retryable set = ['COM_BUSY', 'RPC_E_CALL_REJECTED']    # unchanged
```
Representative deterministic sites (file:line):
- Missing object → `NOT_FOUND`: `names_adapter.py:273,574`; `tables_adapter.py:356,374,762,908,1008,1115`.
- Invalid / ambiguous argument → `INVALID_RANGE`: `names_adapter.py:222,234,245,256,584`; `tables_adapter.py:435,440,532,574,610,647,656,695,714,748,772,813,818,832,845,867,876,883,931,943,955,973,983,999,1021,1051,1069,1133`.
- Duplicate / cross-scope name → `NAME_CONFLICT`: `names_adapter.py:617,628`; `tables_adapter.py:703` (duplicate column header).

### 1c. `NAME_CONFLICT` is a new additive code (category `validation`, `isRetryable=false`, corrective action)
- Enum member: `errors.py:86` (`NAME_CONFLICT = "NAME_CONFLICT"`), with rationale comment `errors.py:81-85`.
- Category map: `errors.py:180` → `"validation"`.
- Actions: `errors.py:270` → `["choose a unique name or delete the existing one"]`.
- Exception class: `errors.py:660-674` (`NameConflictError`, `default_code = NAME_CONFLICT`).
- Code→exception registration: `errors.py:736` (`ErrorCode.NAME_CONFLICT: NameConflictError`); inert for `error_from_hresult` (no HRESULT maps to it), as documented.
- Additive only: `tests/unit/test_errors.py` asserts the Stage-1 closed set is untouched (subset check) and that `NAME_CONFLICT` is a new member.

### 1d. Genuine transient COM paths still retryable (`COM_BUSY` / `RPC_E_CALL_REJECTED` unchanged)
- `com_worker.py` is **not modified** (not in `git diff --name-only`). The HRESULT→code map (`errors.py:91-100`), `_RETRYABLE_CODES` (`errors.py:186-188`), and the worker classification/retry logic (`com_worker.py:769-808`, `810-828`) are unchanged.
- No deterministic raise in the two adapters carries `code=COM_BUSY`/`RPC_E_CALL_REJECTED` (grep: only prose comments mention COM_BUSY).

### 1e. Taxonomy map + global default + `SCHEMA_VERSION`
- `_DEFAULT_CATEGORY = "validation"` (`errors.py:274`) and `_DEFAULT_ACTIONS` (`errors.py:275`) unchanged; base `default_code = ErrorCode.COM_BUSY` (`errors.py:346`) unchanged.
- `SCHEMA_VERSION = 1` (`domain/models.py:65`) unchanged — `models.py` is not in the diff.
- `_MANIFEST_SCHEMA_VERSION = 2` (snapshot store) is a separate, unchanged constant.

### WP1 advisory finding (MEDIUM, non-blocking) — `[BEST-PRACTICE]`
**What:** In the broad `except Exception as exc` wrappers around single COM calls, a *genuine* transient COM failure is now re-labelled as non-retryable `INVALID_RANGE`, because the wrapper does not inspect the original HRESULT and the worker re-raises domain errors as-is (`com_worker.py:769-773`, `_hresult_of` at `243-258` inspects only the top-level exception, not `__cause__`). Affected wrappers:
- `names_adapter.py:433-443` (`collection.Add`), `467-477` (`RefersTo` set), `496-505` (`Delete`).
- `tables_adapter.py:831-839` (`ListObjects.Add`), `844-848` (`TableStyle`), `884-890` (`Resize`), `1130-1140` (`Range`/`Names` resolution).

**Why it matters:** If Excel returns `RPC_E_SERVERCALL_RETRYLATER` (`0x8001010A`, `COM_BUSY`) inside one of these narrow windows, the caller receives `INVALID_RANGE` / `isRetryable=false` ("fix arguments") instead of a retryable `COM_BUSY`. Before WP1 these same sites defaulted to `COM_BUSY` (retryable) — so WP1 is a net improvement for the common deterministic case but narrows the rare transient case. Impact is low-frequency and window-limited; the primary WP1 goal (deterministic raises correctly classified) is fully met and verified by unit + live tests.

**Recommendation (advisory):** In those `except` blocks, re-raise a retryable error when `_hresult_of(exc)` maps to `COM_BUSY`/`RPC_E_CALL_REJECTED` (or `code_from_hresult(...)` is retryable), and only fall back to `InvalidRangeError` otherwise. **Effort:** small. **Priority:** recommended (not mandatory).

---

## WP2 — docs/09 (+RU) `RANGE_TOO_LARGE` — PASS
`git diff docs/09-limitations-and-compatibility.md docs/ru/09-limitations-and-compatibility.md`: the "limit-specific code pending" row is replaced.
- EN `docs/09-limitations-and-compatibility.md:98`: "rejected **before execution** with an explicit `RANGE_TOO_LARGE` error (category `validation`, `isRetryable=false`) — never silently truncated … split/paginate … via `offset` / `limit`." Status `Verified`.
- RU `docs/ru/09-limitations-and-compatibility.md:97`: same semantics ("отклоняется **до выполнения** … `RANGE_TOO_LARGE` … `isRetryable=false` … разбить … через `offset` / `limit`").
- No remaining "pending"/"limit-specific code" text anywhere in `docs/` (grep for `pending|limit-specific` → none).
- Consistent with siblings: `docs/05-mcp-tools-reference.md:327-332` (+RU `:340-345`), `docs/06-excel-capabilities.md:87`, `docs/08-security.md:245,254` all state reject-before-execution / not silently clipped / `RANGE_TOO_LARGE` distinct from `truncated=true`.

---

## WP3 — R1/R2/R3 snapshot-restore hardening + caller contract — PASS
All four required behaviors are implemented in code (from commit `2a90e7a`, present at HEAD) and documented.

- **`previous_close` (R1):** `snapshot_service.py:427,466-498` (`_close_previous_copy` returns `("closed"|"failed"|"not_requested")`, a close failure never fails the restore); result model `models.py:587,595`; surfaced in `SnapshotRestoreResult` (`snapshot_service.py:460-464`).
- **Required `workbook_id` (R2):** enforced at the tool handler **before** any close/open — `api/test_tools.py:156-168` (`INVALID_RANGE`, `isRetryable=false`, `reason="workbook_id_required"`); `confirm` gate at `:169-179`.
- **Canonical rebind (R3):** `snapshot_service.py:437-442` (`rebind_restored_identity(canonical=manifest.source_path, source_sha256=manifest.sha256)`); the working copy stays `work/<newid>.ext`, the immutable snapshot is never the session identity.
- **`same_snapshot` (R3):** `snapshot_service.py:377-392` — deterministic `INVALID_RANGE` / `reason="same_snapshot"` raised **before** any open.
- **Caller contract documented:** `docs/05-mcp-tools-reference.md:218-240` (new "caller contract (R1–R3)" block) and `:579-585` (tool row); RU mirror `docs/ru/05-mcp-tools-reference.md:224-247, 594-600`.

---

## WP4 — `sheet_copy(target_name)` forwarding + validation + live test — PASS
- **Forwarding (source-verified):** `api/sheet_tools.py:133-153` (`sheet_copy(..., *, target_name=None)` → `service.copy_sheet(..., target_name=target_name)`); dispatch `sheet(...)` forwards `target_name` for `action=="copy"` at `:237-238`; tool param declared `:200,223-225`.
- **Validation (source-verified):** `sheets_adapter.py:320-367` (`copy_sheet` renames only when `target_name is not None`); `_validate_target_name` at `:419-449` rejects empty / `>31` chars / illegal chars (`: \ / ? * [ ]`) / duplicate (case-folded) with `INVALID_RANGE` **before** the COM `Name` setter; a rename COM failure is wrapped as `INVALID_RANGE` (`:362-366`).
- **Live test added:** `tests/integration_excel/test_sheets_copy_target_name_live.py` (drives real `SheetsService`/`SheetsAdapter` on a hidden owned Excel over a disposable copy; asserts named copy, default `"<source> (2)"`, duplicate → `INVALID_RANGE`/`duplicate_target_name` with no second sheet, invalid `:` and 32-char names → `INVALID_RANGE`/`invalid_target_name`, fixture byte-identical, owned PID reaped).

---

## 5 — Test execution

Environment: `.venv` (Python 3.14.7, pytest 9.2.0, anyio 4.16.0), Windows host with Excel + pywin32.

**Unit** — `pytest tests/unit -q`:
```
1483 passed in 28.31s
```
(0 failed, 0 errors.)

**Live** — `pytest tests/integration_excel/test_error_taxonomy_live.py tests/integration_excel/test_sheets_copy_target_name_live.py -p no:faulthandler -rN`:
```
collected 5 items
test_error_taxonomy_live.py ....            [ 80%]
test_sheets_copy_target_name_live.py .      [100%]
5 passed in 11.59s
```
All 5 live tests executed (not skipped) against a real hidden Excel instance.

**No deterministic case is `COM_BUSY`/retryable:** both live files and the new unit tests assert `to_dict()` with `errorCode ∉ {COM_BUSY, RPC_E_CALL_REJECTED}`, `errorCategory ≠ transient_com_failure`, `isRetryable is False`, and corrective (non-"retry") `suggestedNextActions` — for `NAME_CONFLICT`, `NOT_FOUND`, and `INVALID_RANGE` cases (`test_error_taxonomy_live.py:117-138`; `test_names_service.py` / `test_tables_service.py` `_assert_non_retryable_taxonomy`). Independent probe above confirms the same. Live tests additionally assert **no partial mutation** (name/table counts unchanged) and byte-identical fixtures.

---

## 6 — No unrelated changes / no new worker / no external-pivot change — PASS
`git diff --name-only` (tracked modifications) is exactly the expected set:
```
docs/05-mcp-tools-reference.md
docs/09-limitations-and-compatibility.md
docs/ru/05-mcp-tools-reference.md
docs/ru/09-limitations-and-compatibility.md
src/excel_vba_mcp/adapters/excel/names_adapter.py
src/excel_vba_mcp/adapters/excel/tables_adapter.py
src/excel_vba_mcp/domain/errors.py
tests/unit/test_error_envelope.py
tests/unit/test_errors.py
tests/unit/test_names_service.py
tests/unit/test_tables_service.py
```
Untracked additions are only the two new live test files (`test_error_taxonomy_live.py`, `test_sheets_copy_target_name_live.py`) plus pre-existing `.ai/` audit artifacts.
- **No new worker:** `git ls-files --others --exclude-standard -- src/` → empty; `com_worker.py` unmodified.
- **No external-pivot change:** `pivot_tables_adapter.py` unmodified (not in diff).
- **No `models.py` / `SCHEMA_VERSION` change.**

---

## Final decision

**ACCEPT.**

All six verification items pass. WP1's deterministic raise sites are correctly and non-retryably classified (`NOT_FOUND` / `INVALID_RANGE` / `NAME_CONFLICT`), the `NAME_CONFLICT` code is additive with `validation` category, `isRetryable=false` and a corrective action, the transient-COM classification machinery and global taxonomy defaults are unchanged, and `SCHEMA_VERSION` remains `1`. WP2/WP3/WP4 are complete and consistent across EN/RU docs and code, with passing unit (1483) and live (5) tests. One non-blocking MEDIUM advisory is recorded for WP1 (broad `except` wrappers can mask a genuine transient `COM_BUSY` as non-retryable `INVALID_RANGE` in a narrow window); it does not affect the acceptance decision.