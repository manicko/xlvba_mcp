# Error Taxonomy Remediation Report (WP1)

- **Revision:** `7961038` + WP1. No taxonomy-map change; no global default change; additive `NAME_CONFLICT` only; `SCHEMA_VERSION=1`.

## Root cause
Adapters raised bare `ExcelVbaMcpError(...)` (no `code=`) → default `COM_BUSY` → `transient_com_failure`/retryable. The taxonomy map (`errors.py`) was already correct; the defect was at the raise sites.

## Decisions
- Missing object (table/sheet/defined name/unknown column) → `ObjectNotFoundError` (`NOT_FOUND`, `not_found`, isRetryable=false).
- Invalid/ambiguous argument (scope, name syntax, `RefersTo`, offset/limit, out-of-range, ambiguous selector, invalid row/header data, pagination) → `InvalidRangeError` (`INVALID_RANGE`, `validation`, non-retryable) with specific message + `context.reason`.
- Duplicate same-scope / cross-scope name → **new** `ErrorCode.NAME_CONFLICT` (`validation`, non-retryable, action "choose a unique name or delete the existing one") + `NameConflictError`. Additive; no version bump.
- Ambiguous name resolution → `INVALID_RANGE` reason `ambiguous_name`.
- Genuine transient COM (`ComBusyError`/`RpcCallRejectedError`/retryable HRESULT) → re-raised unchanged (retryable), never wrapped.

## Affected sites (before → after)
- `names_adapter.py`: 218 invalid_scope, 230 missing_sheet, 241 invalid_name, 252 invalid_refers_to, 430/463/491 create/update/delete COM wrappers (now INVALID_RANGE + transient re-raise), 573 ambiguous_name, 606 duplicate→NAME_CONFLICT, 617 scope_conflict→NAME_CONFLICT (all were COM_BUSY).
- `tables_adapter.py`: NOT_FOUND at 352/370/748/885/967/1066; NAME_CONFLICT at 699; INVALID_RANGE at the remaining 28 (431/436/528/570/606/643/652/691/705/734/753/794/799/813/822/844/853/860/904/912/923/937/943/958/976/1006/1024/1084) — all were COM_BUSY.
- Transient-COM guard added (`adapters/excel/com_transient.py`) at names 435/471/502 and tables 833/848/881/890/1124/1144 so a genuine retryable HRESULT is re-raised, not relabeled.

## Rationale
`INVALID_RANGE` is the established "invalid/ambiguous argument" validation code (non-retryable); `NAME_CONFLICT` is a precise additive code for duplicate/cross-scope (the taxonomy had no conflict code). Missing objects use the established `NOT_FOUND`. No deterministic case is `COM_BUSY`/retryable; genuine transient COM remains retryable.

## Test evidence
- Unit: `pytest tests/unit -q` green (incl. `test_names_service.py`, `test_tables_service.py`, `test_errors.py`, `test_error_envelope.py`, `test_transient_com_wrappers.py` — 13 transient-COM tests). Each asserts the serialized `to_dict()` (errorCode/category/isRetryable/actions), not just the class.
- Live (`test_error_taxonomy_live.py`, real Excel): 4 passed — duplicate name→NAME_CONFLICT (name count unchanged, original hash unchanged), unknown header→NOT_FOUND (row count unchanged), missing table/sheet→NOT_FOUND; none COM_BUSY/retryable.
- Independent Auditor `report_14` ACCEPT; Validator `report_14a` APPROVE; MEDIUM advisory (transient-COM masking) closed by the guard + tests.