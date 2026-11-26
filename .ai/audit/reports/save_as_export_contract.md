# Save As / Export Contract (U7)

- **Tool:** `excel_save_as` (single tool; `excel_save` unchanged).
- **Status:** implemented + unit-tested (15 cases) + Validator ACCEPT (after `.bak` guard fix).
- **SCHEMA_VERSION:** 1 (additive `SaveAsResult`).

## 1. Semantics (answers to §5.1)
1. **Is the working copy saved to the target path?** Yes — the current working copy is flushed (`wb.Save()`) and atomically copied to the target.
2. **Does the session remain bound to the working copy?** Yes. The session `workbook_id`/`working_path`/`canonical` are unchanged.
3. **Does Save As change the session's current path/identity?** No. `wb.SaveAs` is never used (it would rebind `FullName` and break `_find_open_workbook` + original protection).
4. **Is export separate from the session?** Yes — export is a filesystem copy; the session is untouched.
5. **Can the agent continue editing after save?** Yes; subsequent `excel_save` writes the working copy (target hash unchanged).
6. **What does a later `excel_save` save to?** The working copy (in place), never the Save As target.

One tool suffices; a separate `excel_export` is unnecessary (same operation, different flags). PDF/XPS fixed-format export is out of scope (PathPolicy rejects those extensions).

## 2. Signature
```
excel_save_as(workbook_id: str, path: str, *,
              overwrite_existing: bool = False,
              overwrite_source: bool = False) -> SaveAsResult
```

## 3. Overwrite rules (default safe)
| Target T vs Source S | T exists | overwrite_source | overwrite_existing | Result |
|---|---|---|---|---|
| T≠S | no | any | any | write (new file) |
| T≠S | yes | any | false | reject `PROTECTED_TARGET/target_exists` |
| T≠S | yes | any | true | write + `<T>.bak` |
| T==working copy | — | any | any | reject `PROTECTED_TARGET/target_is_working_copy` |
| T==S | — | false | any | reject `PROTECTED_TARGET/source_protected` |
| T==S | — | true | (implied) | write source-overwrite + `<S>.bak` (hash must match) |

`overwrite_source` is the ONLY sanctioned path to modify an original; it is distinct from `overwrite_existing` (not a vague `confirm`).

## 4. Path & identity safety
- Target validated by `PathPolicy.validate_path` (allowlisted roots, extension, symlink/UNC/device rejection).
- Target↔source compared by `normcase` + realpath (Windows case-insensitive, alias-safe) — never string/`Path` equality.
- Source identity captured at open (`source_sha256` + `source_size` before materialize); re-validated + re-hashed at Save As.
- External source modification → `STALE_TARGET/source_modified_since_open`; source replaced → `STALE_TARGET/source_replaced`.

## 5. Atomic write + backup
1. `wb.Save()` on the single-STA worker (flush working copy).
2. Stage temp on target FS (`mkstemp(dir=target.parent)`) → copy → flush+fsync → verify staged size+sha == working.
3. If target exists: if `<target>.bak` already exists (prior interrupted save) → **deterministic pre-mutation rejection** `PROTECTED_TARGET` / `reason=backup_exists` (isError), `.bak` and destination left untouched (no auto-resume). The backup is created/overwritten only when no prior `.bak` exists.
4. `os.replace(tmp, target)` (same volume → atomic).
5. Verify final size+sha; only then `outcome="ok"`.
6. `finally` unlink leftover temp.

## 6. Result (`SaveAsResult`)
`workbook_id, target_path, source_path, source_overwritten, overwrote_existing, backup_path (str|None), sha256, size, outcome ("ok"|"unknown"|"error"), mode, notes, created_at, updated_at`.

## 7. Errors / interruption
- Failed write preserves existing destination; `outcome="error"` (typed error) or `"unknown"` if final state unprovable.
- Interrupted after backup → `outcome="unknown"` + `backup_path`; never success; no auto-resume.
- Never `ok` from "no exception"; `ok` only after final size+hash verify.
- Approval MUTATING (FULL auto-allows; RESTRICTED requires grant → `APPROVAL_DENIED`). Snapshot gate skipped (Save As does not mutate the open working copy).