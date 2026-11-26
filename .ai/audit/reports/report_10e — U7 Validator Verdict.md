# Report 10e — U7 (Save As / export) Validator Verdict

- **Role:** Independent Validator, Work Package U7. **VALIDATION ONLY — no source/test/config change made.**
- **Repo / tree:** `C:\py_exp\xlvba_mcp` (working tree has U6-A..E + U7).
- **Contract:** `.ai/audit/plans/plan_05 — Remaining Usability & Save As Implementation.md` §5.5 (decision matrix), §9 (safety model), §8 (U7), §4 N4.
- **Mandate:** REJECT any variant that can (a) silently overwrite a source file or (b) ambiguously report a failed save as successful. Be adversarial; cite lines.
- **Evidence base:** direct read of `src/excel_vba_mcp/{application/session_service.py, api/workbook_tools.py, domain/models.py, domain/errors.py, server.py}`, `tests/unit/test_save_as.py`, `tests/unit/test_description_contract.py`; `pytest tests/unit/test_save_as.py` (14 passed) and full `pytest tests/unit` (green, no F/E).

---

## Per-item verdicts

### Item 1 — Registration, tool count 92, `excel_save` unchanged, session stays bound — **APPROVE**

- `excel_save_as` registered: `server.py:1889-1894` (`name="excel_save_as"`, `description=_DESC_SAVE_AS`, `annotations=_annotations("excel_save_as")`); imported `server.py:166`; handler `api/workbook_tools.py:167-202` (thin, no COM; delegates to `service.save_as_workbook`).
- Tool count **92**: `tests/unit/test_description_contract.py:84-87` (`assert len(tools) == 92`) — **passes** (run confirmed). `_CRITICAL` includes `excel_save_as` (`test_description_contract.py:22`).
- `excel_save` **unchanged**: handler `workbook_tools.py:152-164` (no `path` arg, calls `save_workbook`); service `session_service.py:507-571` still `wb.Save()` in place (`:545`). `_DESC_SAVE` now cross-refs `excel_save_as` (`server.py:249-258`) — description-only, behaviour intact.
- **No `wb.SaveAs` anywhere** in `src` (grep: only doc/comment mentions at `session_service.py:591,683`, `models.py:1493`, `server.py:271`).
- Session stays bound: `save_as_workbook` never mutates `record.canonical` / `working_path` / `workbook_id` (`session_service.py:577-727`); `_find_open_workbook` still resolves by `working_path` (`.lower()` FullName match, `:1582-1598`). Test 13 (`test_save_as.py:417-439`) asserts identity unchanged and a later `excel_save` writes the working copy with the Save-As target hash unchanged.

### Item 2 — Source identity capture / re-hash / STALE_TARGET — **APPROVE**

- Fields added to `_SessionRecord`: `source_sha256: str`, `source_size: int` (`session_service.py:107-108`).
- Captured **after** `validate_path` (`:290`) and **before** `_materialize_working_copy` (`:337`): `source_sha256, source_size = _sha256_file(canonical)` at `:336`. Streaming SHA-256 (`_sha256_file`, `:115-124`).
- Re-validated + re-hashed at Save As: re-`validate_path(record.canonical, READ_ONLY)` `:647`; **normcase of realpath'd canonicals** compare `:648` (`os.path.normcase(str(re_canonical)) != os.path.normcase(record.canonical)`) → `STALE_TARGET/source_replaced` `:649-656`; re-hash `:657` → mismatch → `STALE_TARGET/source_modified_since_open` `:658-666`. Target/source equality also normcase (`:668`) — never string/`Path` equality.
- Tests: case 8 external modification → `STALE_TARGET/source_modified_since_open` (`test_save_as.py:308-319`); case 14 source replaced (re-canonical differs) → `STALE_TARGET/source_replaced` (`:445-466`).

### Item 3 — Decision matrix exactly per §5.5 — **APPROVE**

Implementation `session_service.py:668-681` + working-copy guard `:630-637`:
- T==W → `PROTECTED_TARGET/target_is_working_copy` (`:630-637`) — matrix row 4.
- same(T==S) & `not overwrite_source` → `PROTECTED_TARGET/source_protected` (`:672-676`) — row 5.
- `not same` & exists & `not overwrite_existing` → `PROTECTED_TARGET/target_exists` (`:677-681`) — row 2.
- `not same` & not exists → **write** new file — row 1.
- `not same` & exists & `overwrite_existing` → **write + .bak** — row 3.
- same & `overwrite_source` & hash match → **write + .bak of source** — row 6 (hash mismatch already rejected at `:658`, row 7).
- Missing flags raise typed `PROTECTED_TARGET` (not a vague confirm). Default `overwrite_existing=False`/`overwrite_source=False` (`workbook_tools.py:171-172`, `session_service.py:582-583`).
- Tests 1,3,4,5,6,7 (`test_save_as.py:177-302`) cover each branch.
- **Observation (not a defect):** the source re-hash at `:647-666` runs **unconditionally**, so a drifted source rejects even a brand-new-target Save As. This is stricter than a literal "source hash = any" reading of the matrix but is exactly what §9 mandates ("At Save As: … re-hash source; if != → STALE_TARGET") and is fail-closed. No silent-overwrite / false-success path.

### Item 4 — Atomic + backup + verify + fail-closed outcome — **APPROVE (core) / one §9 DEVIATION (Required Fix)**

Core algorithm `_atomic_save_as_copy` (`session_service.py:754-851`) — all required steps present and correct:
1. `wb.Save()` flush on STA before any FS copy (`_flush_working_copy` `:729-752`, awaited FIFO `:747`).
2. Temp on **target FS**: `mkstemp(prefix=f".{name}.", suffix=".tmp", dir=canonical_t.parent)` `:780-782`; copy working→tmp + `flush`+`fsync` `:785-790`.
3. **Staged verify** size+sha vs working `:792-793`; mismatch → `outcome="error"`, destination untouched `:794-801`.
4. Single `<target>.bak` when target exists `:803-807`.
5. `os.replace(tmp, canonical_t)` same-volume atomic `:809`; `tmp=None` `:810`.
6. **Final verify** size+sha `:812-813`; `outcome="ok"` returned **only** after it passes `:824-831`.
7. `finally` unlinks leftover temp `:844-850`.
- Failed pre-backup write → `outcome="error"`, existing dest preserved, no `.bak` (`:832-843`, `backup_created` gate). Test 10 (`test_save_as.py:337-360`).
- Interrupted post-backup → `outcome="unknown"` + `backup_path`, **no auto-resume** (no resume/retry code anywhere) `:832-843`. Test 11 (`:366-392`).
- **Never `ok` from "no exception":** `ok` is emitted only at `:825` after the final verify; the `except OSError` path returns `unknown`/`error` `:832-843`.
- Final-verify mismatch → best-effort rollback from `.bak` + `unknown` `:813-823` (`_rollback_from_backup` `:861-884`).

**DEVIATION (Required Fix) — §9 retry/backup-preservation clause unimplemented:**
§9 mandates *"a retry must not orphan a second backup (detect existing `.bak` for the same target and stop)."* `_atomic_save_as_copy` **unconditionally overwrites** any existing `.bak` (`:802-805`, `_copy_and_fsync` opens dest `"wb"`), with **no detection of a pre-existing `.bak`**. Concrete data-loss path:
1. Existing target `T`=O; Save As `overwrite_existing=True` writes `T.bak`=O, `os.replace` publishes N, then the process dies before/at final verify → state `T`=N, `T.bak`=O (only copy of O), outcome would be `unknown`.
2. A retry (fresh session reopened from source S) Save As to `T` with `overwrite_existing=True`: `target_exists=True` (T=N) → `_copy_and_fsync(T=N → T.bak)` **overwrites `T.bak`=O with N**. O is permanently lost.
This is **not** a silent *source* overwrite (source is hash-protected) and **not** a false success (outcome is `unknown`), so it does **not** trip the two named REJECT triggers — but it violates the binding §9 safety model and can destroy the only recovery copy. **Required fix:** before overwriting, if `f"{canonical_t}.bak"` already exists, stop and return `outcome="unknown"` with `backup_path` and a "manual reconcile required; existing .bak preserved" note (do not overwrite the recovery copy).

### Item 5 — `SaveAsResult` / SCHEMA_VERSION=1 / PathPolicy / approval — **APPROVE**

- `SaveAsResult(TimestampedModel)` `models.py:1485-1523`; `extra="forbid"` inherited via `DomainModel` (`:68-71`). All §5.5 fields present: `workbook_id, target_path, source_path, source_overwritten, overwrote_existing, backup_path (str|None), sha256, size, outcome Literal["ok","unknown","error"], mode, notes, created_at, updated_at`.
- `SCHEMA_VERSION = 1` (`models.py:65`) — unchanged.
- **No new ErrorCode beyond U6-B:** `ProtectedTargetError`→`PROTECTED_TARGET` (`errors.py:521-524`), `StaleTargetError`→`STALE_TARGET` (`:535-538`), `PathNotAllowedError`→`PATH_NOT_ALLOWED` (`:486-489`) are all pre-existing Phase 2 codes (`errors.py:45,47,42`). U6-B's `INTERNAL_ERROR`/`CONFIRMATION_REQUIRED` (`:74-75`) are the only recent additions; U7 adds none.
- PathPolicy applied to target `validate_path(path, MUTATING)` `:628` and source re-validate `READ_ONLY` `:647` — **no bypass**.
- Approval **MUTATING** on target canonical (== source when T==S) `:641-644`; FULL auto-allows, RESTRICTED requires grant → `APPROVAL_DENIED` (`_require_approval` `:1514-1543`). Skip-snapshot honoured (no snapshot gate call in `save_as_workbook`).
- Annotations: `excel_save_as` absent from `_ANN_READ_ONLY`/`_ANN_DESTRUCTIVE`/`_ANN_IDEMPOTENT` (`server.py:1777-1862`) → `read_only_hint=None, destructive_hint=None, idempotent_hint=None` (mutating; destructive off) per §8.

### Item 6 — Tests cover the 14 cases — **APPROVE**

`tests/unit/test_save_as.py` (14 tests, all pass) maps 1:1 to the §9/§8 matrix:
1 new path `:177-196`; 2 save-as copies working bytes `:199-215`; 3 unauthorized overwrite reject + target untouched `:221-231`; 4 authorized overwrite + `.bak` `:237-249`; 5 source w/o flag reject + **source hash unchanged** `:255-263`; 6 source + flag + hash match + `.bak` `:269-280`; 7 case/alias recognised `:286-302`; 8 external change → STALE_TARGET + **no output written** `:308-319`; 9 outside roots → PATH_NOT_ALLOWED `:325-331`; 10 permission/disk failure → error, **dest untouched, no .bak** `:337-360`; 11 interruption → unknown + backup_path, dest not replaced `:366-392`; 12 single `.bak` overwritten on 2nd save `:398-411`; 13 session identity + later `excel_save` writes working copy, **target hash unchanged** `:417-439`; 14 source replaced → STALE_TARGET `:445-466`.
- **Source hash unchanged after failed save** explicitly asserted: case 5 (`:263`), case 3 target untouched (`:231`), case 8 no output (`:319`), case 10 dest untouched (`:359-360`).

### Item 7 — Run tests — **APPROVE**

- `pytest tests/unit/test_save_as.py -q` → **14 passed**.
- `pytest tests/unit -q` → **green** (no `F`/`E`; ~1150 tests, 100%).
- `test_description_contract.py::test_tool_count_is_92_exposed` → **passes**.

---

## Silent-overwrite / false-success analysis

- **Silent source overwrite:** none found. Source is protected by (a) `overwrite_source` default-False + typed `PROTECTED_TARGET/source_protected`, and (b) unconditional re-hash + re-canonical → `STALE_TARGET` on any drift/replacement before any write (`session_service.py:646-676`). The working copy is never the target (`:630-637`). `wb.SaveAs` is never used.
- **False success:** none found. `outcome="ok"` is reachable only after staged verify **and** final size+sha verify (`:792-793, 812-831`); every failure/interruption path returns `error`/`unknown` (`:794-801, 813-823, 832-843`). No code path returns `ok` from "no exception".
- **Residual data-loss (not source, not false-success):** the §9 `.bak`-retry gap in Item 4 can destroy the previous-target recovery copy on a retry after an interrupted overwrite. Flagged as Required Fix.

---

## Required Fixes

1. **(§9 conformance, Item 4)** In `_atomic_save_as_copy` (`session_service.py:802-807`), before overwriting `f"{canonical_t}.bak"`, detect a pre-existing `.bak` for the same target and **stop** with `outcome="unknown"` + `backup_path` + a "manual reconcile required; existing .bak preserved" note, instead of overwriting the only recovery copy. Add a regression test (retry-after-interruption must not clobber a pre-existing `.bak`).

## Advisory Recommendations

- Document (in `_DESC_SAVE_AS` / docs 09) that the source re-hash is unconditional, so a source changed since open blocks **any** Save As target (not only source-overwrite) — matches §9 but is stricter than the matrix's "source hash = any" wording.
- Consider recording `overwrote_existing` distinctly for the source-overwrite case (currently `True` because the source exists) to avoid operator ambiguity; cosmetic only.

## Warnings

- **Maintainability/rollout:** the `.bak`-retry gap is a narrow crash-boundary + retry edge case, but it is the one place U7 can lose data; treat the Required Fix as release-relevant if §9 is considered strictly binding.
- No dependency/rollout ordering issue found: U7 is additive (new tool + additive DTO), `SCHEMA_VERSION=1`, count 91→92 coordinated, `excel_save`/session binding untouched.

---

## Final U7 verdict: **ACCEPT (with one Required Fix against §9)**

Both binding safety non-negotiables are satisfied — there is **no silent source-overwrite path** and **no false-success path**; the decision matrix, source-identity re-verification, atomic temp→replace + single `.bak` + staged/final verify, fail-closed `outcome`, session-binding, DTO/`SCHEMA_VERSION=1`/PathPolicy/MUTATING-approval, and the 14-case test surface all conform, and the full unit suite is green. The single outstanding item is the §9 "detect existing `.bak` and stop on retry" clause (Item 4), which does not create either named REJECT trigger but can destroy the previous-target recovery copy; it is logged as a Required Fix to complete §9 conformance.