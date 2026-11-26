# Research 2 — Save As / Export (remaining-usability remediation, `xlvba` MCP)

- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `40c8b4d7a1e5a7406bbd9f074319e112711e60fc`
- **Role:** Researcher 2. **RESEARCH ONLY** — no code/test/config change. Backward compatibility NOT required.
- **Method:** read-only source inspection at HEAD. Every claim below is cited `file:line` and verified against the working tree.
- **Invariants that MUST be preserved:** single hidden Excel (one `DispatchEx`, `Visible=False`), single-STA FIFO worker, Job Object `KILL_ON_JOB_CLOSE`, working-copy model (originals never modified), originals protected by default. **Save As source-overwrite is a deliberate, narrow, opt-in exception** — the only sanctioned way to write back to an allowlisted source, and only behind an explicit flag + stale-source guard.
- **Authority order:** the task brief > `plan_03a` contract outline (`.ai/audit/plans/plan_03a — Deferred Product Contracts (save_as, job_list).md`) > source at HEAD. Where the brief and `plan_03a` differ, the brief's "session stays bound to the working copy" decision governs (§6).

Severity/confidence legend: **HIGH** = directly read in source; **MEDIUM** = inferred from multiple reads; **LOW** = design choice / unverifiable here.

---

## 1. Verified facts — current working-copy & save (HIGH)

### 1.1 `open_workbook` (session_service.py:223)

- Signature `open_workbook(path, *, mode=INSPECT, read_only=False, timeout_s=None)` (`:223-230`).
- Flow: `_maybe_auto_recover` (`:262`) → `_require_accepting_new_work` (`:263`) → `canonical = self._path_policy.validate_path(path, MUTATING)` (`:264`) → approval gate (`:270-279`) → fast reject if already open (`:283-288`) → acquire per-canonical lock held for workbook lifetime (`:294-295`) → `working_path = self._materialize_working_copy(canonical, workbook_id)` (`:306`) → `_open_on_worker(...)` opens **the working copy** (`:309-316`, `:372` `app_adapter.open_workbook(working_path, ...)`) → on failure `_cleanup_working_copy` + `lock.release()` (`:317-323`).
- `WorkbookInfo` is built by `workbook_adapter.describe(wb, source_path=canonical_str)` (`:379-385`): `path` = `wb.FullName` (the working copy), `source_path` = the original canonical (`workbook_adapter.py:97-119`).
- The **original is never opened or written**; only the copy under `work/` is opened (`:231-236` docstring "Copies, never the original"; module docstring `:11-15`).

### 1.2 `_materialize_working_copy` (session_service.py:1199)

```text
1199  def _materialize_working_copy(self, canonical: Path, workbook_id: str) -> str:
1202      self._work_dir.mkdir(parents=True, exist_ok=True)
1203      suffix = canonical.suffix or ".xlsm"
1204      working = self._work_dir / f"{workbook_id}{suffix}"
1205      shutil.copy2(canonical, working)
1206      return str(working)
```

- Working-copy root: `self._work_dir = work_dir or (Path(tempfile.gettempdir()) / "excel_vba_mcp" / "work")` (`:155`); `_WORK_SUBDIR = "work"` (`:80`). Server default also `tempfile.gettempdir()/excel_vba_mcp/work` (`server.py:1252`).
- Filename is the `workbook_id` (uuid4, `:305`) + source suffix — **no source basename, no hash** in the working-copy name.
- Cleanup is best-effort `Path(working_path).unlink(missing_ok=True)` inside `suppress(OSError)` (`:1208-1213`), invoked on close (`:451`), discard (`:1083`), and failed open (`:319`). A leftover working copy is explicitly non-fatal.

### 1.3 `save_workbook` (session_service.py:474)

- Signature `save_workbook(workbook_id, *, timeout_s=60.0)` (`:474-479`). **No path argument.**
- Approval `MUTATING` on `record.canonical` (`:495-498`).
- STA callable `save_fn` (`:504-519`): `_find_open_workbook(app_adapter, record.working_path)` (`:505`) → `with app_adapter.alerts_suppressed(): wb.Save()` (`:511-512`) → `workbook_adapter.describe(...)` (`:513-519`).
- After the STA call: `record.dirty = False; record.info = info` (`:528-529`). `wb.Save()` writes **in place to the working copy** (the file opened at `:372`), never to the original.
- `_find_open_workbook` matches the open proxy by `str(wb.FullName).lower() == working_path.lower()` (`:1236-1252`) — **case-insensitive** match, the only handle-resolution mechanism.

### 1.4 `_SessionRecord` (session_service.py:83-94)

```text
84  class _SessionRecord:            # dataclass
87      workbook_id: str
88      canonical: str                # original allowlisted path (str)
89      working_path: str             # the work/ copy path
90      mode: SessionMode
91      read_only: bool
92      dirty: bool
93      info: WorkbookInfo
94      lock: asyncio.Lock
```

- **No source hash, no source stat/size/mtime, no backup path, no "saved-as target" field.** This is the central gap for source-identity safety (§4).
- The record is created at `:325-335` and stored in `self._records[workbook_id]` (`:338`); `self._active_by_canonical[canonical_str] = workbook_id` (`:339`).

### 1.5 How the original is protected / where the working copy lives (summary)

- **Protection mechanism = structural copy-on-open** (`:306` + `:1205`). The original path is used **only** for `validate_path`, the approval key, the per-path lock, and `WorkbookInfo.source_path`. It is never passed to `Workbooks.Open` and never written.
- **Working copy location:** `%TEMP%\excel_vba_mcp\work\<workbook_id><ext>` by default (`:155`, `server.py:1252`); overridable via `work_dir` (server artifact root). Snapshots/artifacts/runs/jobs each have their own temp subdir (`server.py:1256-1268`).

---

## 2. Verified facts — all persistence / materialization paths (HIGH)

| Path | Where | Mechanism | Atomic? | Hash? |
|---|---|---|---|---|
| Working copy | `session_service._materialize_working_copy:1199-1206` | `shutil.copy2` | No (plain copy) | No |
| Snapshot (closed) | `snapshot_store.create:157-214` | `shutil.copy2` into staging dir (`:201`) then `os.replace(staging, final)` (`:433`) | **Yes** (staging + `os.replace`) | **Yes** `_sha256_file` (`:124-133`, `:402`) |
| Snapshot (open) | `snapshot_service._create_open_branch:192-253` + `snapshot_store.finalize_open_branch:247-309` | `wb.SaveCopyAs(tmp)` on STA (`:219`) → `shutil.move` into staging (`:230-233`) → `os.replace` publish | **Yes** | **Yes** (`:402`) |
| Snapshot manifest | `snapshot_store._atomic_write_json:436-452` | `tempfile.mkstemp` in `path.parent` → `write` → `flush` → `os.fsync` → `os.replace`; temp unlinked on failure | **Yes** | n/a |
| VBA export | `vba_project_service.export_project:192-243` | temp dir (`mkdtemp`, `:228`) or PathPolicy-validated `out_dir` (`:230`); adapter writes `.bas/.cls/.frm/.frx` | Per-file (VBE `Export`) | Per-file `file_sha256` in manifest |
| Run document | `run_repository._atomic_write_json:113-137` | `mkstemp` in `target.parent` → `write`→`flush`→`fsync`→`os.replace`; `finally` closes/unlinks temp | **Yes** | No |
| Artifact | `artifact_store._atomic_write:321-347` | same pattern; wraps `OSError` as `ArtifactWriteError` | **Yes** | `sha256` of payload (`:257`) |
| Owned-pid marker | `owned_pid_store` (docstring `:21`) | flush/`fsync` + `os.replace` | **Yes** | No |
| Job store | `job_store` (docstring `:12`) | temp + `fsync` + `os.replace` | **Yes** | No |

**Reusable atomic-write primitive (HIGH):** the `mkstemp(dir=target.parent) → write → flush → fsync → os.replace` pattern is implemented **four times** (`snapshot_store.py:436-452`, `run_repository.py:113-137`, `artifact_store.py:321-347`, plus job/pid stores). The canonical shape to copy for Save As is `artifact_store._atomic_write` (`:321-347`) because it also wraps `OSError` into a typed error and cleans the temp in `finally`.

**Hashing helpers (HIGH):** `_sha256_file(path)->(hex,size)` streaming 1 MiB chunks (`snapshot_store.py:124-133`); `_sha256_file(path)->hex|None` (`macro_service.py:172-182`); `_sha256(path)->hex` (`snapshot_service.py:594-601`). All use `hashlib.sha256` over `open("rb")` in 1 MiB chunks.

**Backups (HIGH):** **No backup mechanism exists anywhere.** `grep -i "backup|\.bak|preserve.*existing"` over `src` returns only `os.replace` atomic-write hits — there is **no** "copy existing destination aside before replace" helper. Backup-on-overwrite must be designed new (§5).

**Temp/backup file locations (HIGH):** temp files are always created **in the destination directory** (`dir=str(target.parent)` / `dir=str(path.parent)`), never in a separate temp root — this is deliberate so `os.replace` is same-filesystem atomic. Save As must follow the same rule (temp on the **target** filesystem).

---

## 3. Verified facts — `PathPolicy` (security/path_policy.py) (HIGH)

- `validate_path(raw, operation, *, expected_root=None) -> Path` (`:110-185`):
  - Empty → `EMPTY` (`:139-140`); device path `\\?\`/`\\.\` denied unless `allow_unc` (`:144-145`, `_is_device_path:225-232`); UNC `\\`/`//` denied unless `allow_unc` (`:148-149`, `_is_unc_path:235-236`).
  - **Canonicalization:** `candidate=Path(probe)`; `lexical=abspath`; `resolved=_resolve_candidate(candidate)` (`:151-153`). `_resolve_candidate` (`:253-260`) does `candidate.resolve()` **and** `os.path.realpath(candidate)` and **prefers realpath when they disagree** (reparse-point safety).
  - **Containment:** `_is_within(resolved, root)` for each canonical root (`:155`, `_is_within:263-269` via `os.path.commonpath` + `_normalize`). No match → distinguish `SYMLINK_ESCAPE` (lexically inside but resolves outside, `:159-160`) vs `TRAVERSAL` (`..`, `:161-162`) vs `OUTSIDE_ROOTS` (`:163-164`); error carries `allowed_roots` (`:168-174`).
  - `expected_root` mismatch → `ROOT_MISMATCH` (`:176-177`); most-specific root chosen otherwise (`:179-181`).
  - **Extension allowlist:** `candidate.suffix.lower() not in _allowed_extensions(matched_root)` → `EXTENSION_NOT_ALLOWED` (`:182-183`). `WORKBOOK_EXTENSIONS` = `.xlsm/.xlsb/.xlsa/.xlam/.xltm/.xltx/.xlsx/.xls` (`:38-49`); artifact roots add `.json/.sha256` (`:52`, `:187-190`).
  - Returns the **canonical resolved `Path`** (`:185`).
- **Case-insensitivity / aliases (HIGH):** `_normalize = os.path.normcase(os.fspath(path)).rstrip("\\/")` (`:272-273`) — `normcase` lowercases and converts `/`→`\` on Windows, so containment is case-insensitive. **But** `validate_path` returns a `Path` (not normcased); two spellings of the same file (e.g. `C:\Data\a.xlsm` vs `c:/data/a.xlsm`) both canonicalize via `resolve()`/`realpath()` to the same absolute form on Windows, yet the returned `Path` objects compare **case-sensitively** as `PurePath`. **A robust same-file comparison must apply `os.path.normcase` to both canonical strings** (see §4 algorithm).
- **No writes:** the module performs no filesystem writes (`:22`) — it is a pure gate. Save As must call it for the **target** and separately compare target↔source.
- Rejection type `PathNotAllowedError` / code `PATH_NOT_ALLOWED` (`:33`, `:210-213`; errors.py:466-470).

**Reuse for Save As (HIGH):** call `validate_path(target, MUTATING)` for the output path (allowlist + extension + symlink/traversal/UNC/device). For the **source-overwrite** comparison, do **not** rely on `Path` equality — canonicalize both with `validate_path`/`resolve` then compare `os.path.normcase(str(...))` (§4).

---

## 4. Source identity / external-modification detection (design)

**Current state (HIGH):** no source hash/stat is captured at open; `_SessionRecord` (`:83-94`) and `WorkbookInfo` (`models.py:86-99`) carry no hash. There is therefore **no way today** to detect that the on-disk original changed between open and a later write-back.

**Design (MEDIUM — recommended):**

1. **Capture at open.** In `open_workbook`, immediately after `validate_path` and **before** `_materialize_working_copy` (`:306`), compute `source_hash = _sha256_file(canonical)` (reuse `macro_service._sha256_file` shape or `snapshot_store._sha256_file`) plus `source_size = stat().st_size` and `source_mtime_ns = stat().st_mtime_ns`. Store on `_SessionRecord` as new fields `source_sha256: str`, `source_size: int`, `source_mtime_ns: int` (additive dataclass fields; internal only, not a DTO). Rationale: hashing the original **before** the copy proves the bytes the working copy was derived from.
   - Cheap alternative if hashing large files at open is a concern: stat-only (size+mtime) as a fast pre-check, hash only at Save-As time. **Recommendation: hash at open** for a strong identity; files are workbooks (bounded), and the snapshot store already hashes whole files (`snapshot_store.py:402`).
2. **Compare at Save As.** Re-hash the **current** on-disk source (`_sha256_file(canonical)`) and compare to `record.source_sha256`:
   - **Equal** → source unchanged since open → source-overwrite (if requested) is safe.
   - **Different** → source externally modified → **reject** with a stale/conflict error (§7) and suggested actions (re-open to re-baseline; do not blind-overwrite). This mirrors the VBA `RevisionConflictError` discipline (`vba_project_service.py:628-643`, errors.py:585-593) and the `STALE_TARGET` "do NOT blind-replay" guidance (errors.py:196-199).
3. **Identity of the source path itself.** Also re-run `validate_path(record.canonical, ...)` at Save-As time so a source that was **replaced by a symlink/junction** after open is caught (the canonical could now resolve elsewhere). If the re-canonicalized source ≠ the stored `record.canonical` (normcase compare), treat as modified/unsafe.
4. **WorkbookInfo (optional, additive).** Optionally surface `source_sha256` on `WorkbookInfo` (new optional field, `extra="forbid"` respected) so a caller can see the baseline identity. **Recommendation: keep it internal to `_SessionRecord` for now** to avoid an DTO/schema change; expose only if the Auditor wants caller-visible identity (open question §10).

**Why hash, not mtime alone (MEDIUM):** mtime can be preserved by some tools / coarse on some filesystems; size can collide. Hash is the authoritative identity already used elsewhere (`snapshot_store`, `macro_service`, approval `sha256` identity `approval.py:112-117`). Use size+mtime only as a cheap early-out before the hash.

---

## 5. Atomic replacement + backup strategy (design)

**Goal:** overwrite an existing destination (target, or the source in the narrow source-overwrite case) such that the destination is **either the complete new file or the untouched old file** — never a half-written file — and the prior destination is recoverable.

**Algorithm (MEDIUM — recommended), all on the STA worker for the `wb.Save()` step, then pure-FS on the event loop for the copy/replace:**

1. **Validate target** `canonical_t = validate_path(target, MUTATING)` (allowlist + extension). Reject if `canonical_t` equals the **working copy** path (would clobber the live file Excel holds open) — compare normcased.
2. **Source-identity guard** (§4): re-hash source; mismatch → reject.
3. **Existence / overwrite decision:**
   - `exists = canonical_t.is_file()` (or `.exists()`).
   - If `exists and not overwrite_existing` → reject (target-exists, §7).
   - If `canonical_t` is the **source** (normcase-equal to `record.canonical`): require `overwrite_source=True` **and** `overwrite_existing` semantics folded in (source-overwrite implies overwrite); if `overwrite_source=False` → reject (source-protection, §7).
4. **Flush the working copy** on the STA: `wb.Save()` (in place, as `save_workbook` does at `:512`) so the working file on disk reflects the current in-memory state. (Alternative rejected: `wb.SaveAs` — see §6.)
5. **Stage the new bytes on the target filesystem:** `mkstemp(prefix=f".{name}.", suffix=".tmp", dir=canonical_t.parent)`; `shutil.copyfileobj(working_fh, tmp_fh)` (or `shutil.copy2` working→tmp); `flush` + `os.fsync`; **verify** staged size == working size and staged `sha256` == working `sha256`. On any failure → close+unlink temp, **no change to destination**, raise typed error.
6. **Preserve existing destination (backup) — only if `exists`:**
   - Generate a backup name in the **same directory**: `f"{name}.{utcstamp}.bak"` (UTC `YYYYMMDDTHHMMSSZ` + short uuid suffix to avoid collision), e.g. `report.xlsm.20261126T181534Z.a1b2c3.bak`.
   - `os.replace(canonical_t, backup_path)` (atomic move of the old file aside). Now the destination slot is empty and the old bytes are safe at `backup_path`.
   - **Retention:** keep the single newest backup per overwrite; older `.bak` files may be pruned best-effort (or left — product decision §10). Backup location = destination directory (same filesystem, so the later `os.replace` is atomic).
7. **Atomic install:** `os.replace(tmp_path, canonical_t)` (same filesystem → atomic). `tmp_path=None` so `finally` does not unlink it.
8. **Verify final:** `os.path.getsize(canonical_t) == working_size` and `_sha256_file(canonical_t) == working_hash`. If verification fails → **do not report success**; attempt best-effort rollback (`os.replace(backup_path, canonical_t)` if a backup exists), and report `unknown`/error (§7).
9. **Cleanup:** `finally` unlink any leftover temp; on success keep the `.bak`; on pre-install failure the destination was never touched (old file still in place) and temp is removed.

**Why `os.replace` (HIGH):** it is atomic on the same volume on Windows (uses `MoveFileEx` with `MOVEFILE_REPLACE_EXISTING` semantics via Python) and is the established pattern (`snapshot_store.py:433`, `artifact_store.py:337`, `run_repository.py:129`). Staging **in the destination directory** guarantees same-volume.

**Backup naming (MEDIUM):** `f"{name}.{stamp}.bak"` in the destination dir. Deterministic stamp + short random suffix avoids clobbering an existing backup and is self-describing. The backup path is returned in the result DTO (§8) so the caller can restore manually.

**Interrupted-outcome reporting (MEDIUM):** if the process dies between step 6 (`os.replace` old→backup) and step 7 (`os.replace` tmp→dest), the destination is **absent** and the old file is at `backup_path`. This is a **recoverable, reported-as-unknown** state: the result (if the call returns at all) must be `outcome="unknown"` with `backup_path` populated and a suggested action ("restore from backup_path or retry"). On a clean retry, the algorithm must **detect an existing `.bak` for the same target** and either resume (rename backup back) or treat as unknown — never silently create a second backup that orphans the first. (Simplest safe rule: if `canonical_t` is absent but a matching `.bak` exists, report unknown and do not proceed without an explicit retry.)

---

## 6. Excel session identity after Save As — decision (HIGH recommendation)

**Decision: Save As = flush working copy (`wb.Save()`) then atomically copy the working file to the target; the session STAYS bound to the working copy.**

- After Save As, `record.working_path`, `record.canonical`, `record.info`, and `workbook_id` are **unchanged**. A subsequent `excel_save` writes **the working copy**, NOT the save-as target (because `save_workbook` resolves the open proxy by `record.working_path` at `:505` and calls `wb.Save()` at `:512`).
- The returned DTO describes the **new file on disk** (target path + hash), not a new open session (matches `plan_03a` §1.6 option A).

**Justification (preserves invariants):**
- The working-copy model requires exactly one open copy per `workbook_id` (`:307-339`). Rebinding the session to the target would either (a) require closing+reopening (changing `workbook_id`, like `restore_snapshot` `snapshot_service.py:331-343`) or (b) `wb.SaveAs(target)` which **repoints the open Workbook object's FullName to the target** — after which `excel_save` would write the target and `_find_open_workbook` (matching `record.working_path`) would **fail to find the workbook** (`:1242-1247`), breaking every later tool that resolves by working path (`vba_project_service._find_open_workbook:820-836`, `snapshot_service._find_open_workbook:513-529`, `macro_service`). That is a cross-cutting regression.
- `wb.SaveAs` also **rebinds original protection**: the "original never modified" guarantee is structural (only the copy is opened). `SaveAs` to the source path would make the open object point at the original — destroying the invariant for the rest of the session. The flush+`copy` approach keeps Excel's open object pointed at the **working copy** forever.
- Therefore `wb.SaveAs` is **rejected**. Save As is implemented as `wb.Save()` (STA) + pure-filesystem atomic copy (event loop), never `wb.SaveAs`.

**One tool vs two (HIGH recommendation): ONE tool `excel_save_as`.** Rationale:
- `excel_save` (in-place, no path) and `excel_save_as` (new path, overwrite flags) are distinct intents but the same "persist the working copy" family; a single `excel_save_as(workbook_id, path, *, overwrite_existing=False, overwrite_source=False)` is the smallest surface that closes the gap without a vague `confirm`.
- A separate `export_fixed_format`/PDF tool is **out of scope** (PathPolicy rejects `.pdf/.xps`, `path_policy.py:38-49`; `plan_03a` §1.7) — a future, separate decision.
- `vba_export_project` stays VBA-only (`vba_project_service.py:192-243`); Save As is a **workbook-file** write, complementary, not a replacement.

---

## 7. Save As / export contract (the deliverable)

### 7.1 Signature (proposed)

```
excel_save_as(workbook_id: str, path: str,
              *, overwrite_existing: bool = False,
              overwrite_source: bool = False) -> SaveAsResult
```

- Thin async handler in `api/workbook_tools.py` (mirrors `excel_save` `:147-159`, no COM), delegating to `SessionService.save_as_workbook(...)`.
- Registered in `server.py:_register_tools` (`:1845`) with `_DESC_SAVE_AS` and `_annotations("excel_save_as")` (mutating, **not** read-only, **not** idempotent; consider `destructive_hint` when `overwrite_*` — but annotations are static per name, so leave destructive off and rely on the flags + description).
- Approval class **MUTATING** on the **target** canonical (and, for source-overwrite, also the source), consistent with `excel_save` (`:495-498`) and `snapshot_create` (`snapshot_service.py:162-166`). FULL mode auto-allows; RESTRICTED requires allowlist/grant.

### 7.2 Semantics Q1–Q6 (answers)

- **Q1 — Snapshot gate?** **SKIP `require_snapshot`.** Save As does not mutate the open working copy; it reads it and writes a separate file (like the snapshot SaveCopyAs branch, op_class `READ_ONLY`, `snapshot_service.py:210-228`). The original is protected structurally; there is no in-place state to roll back. State this explicitly in the description. (Requiring a snapshot would impose an in-place-mutation precondition on a non-mutating new-file write.)
- **Q2 — Working-copy switch / session rebinding?** **NO switch.** Session stays bound to the working copy (§6). `workbook_id`, `working_path`, `canonical` unchanged; later `excel_save` writes the working copy, not the target.
- **Q3 — Office-trust for the output file?** **Inherit the source session's mode.** The output is *written* (not opened); its bytes derive from an already-open, already-trust-decided working copy. Record the inherited mode in the audit entry. (Re-evaluating trust on a written file is meaningless — trust applies at open.)
- **Q4 — Overwrite error code (no new `ErrorCode` members; `SCHEMA_VERSION=1`).** Reuse existing codes:
  - Target exists & `overwrite_existing=False` → **`PATH_NOT_ALLOWED`** is wrong semantically; prefer **`STALE_TARGET`**? No — `STALE_TARGET` means "identity changed". The cleanest existing fit is **`PROTECTED_TARGET`** ("target is protected/read-only", errors.py:501-505) for "refuse to clobber an existing file", OR reuse **`INVALID_RANGE`**-family. **Recommendation: `PROTECTED_TARGET`** for "existing target not overwritten", with `context={"reason":"target_exists","target_path":...}`. (Open question §10 — Auditor to confirm vs a future dedicated code.)
  - Source-overwrite requested but `overwrite_source=False`, or source==target without the flag → **`PROTECTED_TARGET`** with `reason="source_protected"`.
  - Source externally modified (hash mismatch) → **`REVISION_CONFLICT`** (errors.py:585-593) with `reason="source_modified_since_open"` (the "do NOT blind-replay" family, errors.py:200-203).
  - Target outside roots / bad extension → **`PATH_NOT_ALLOWED`** (from `validate_path`).
  - Target == working copy path → **`PROTECTED_TARGET`** `reason="target_is_working_copy"`.
- **Q5 — Fixed-format/PDF?** **Separate future tool**, not part of `excel_save_as` (PathPolicy rejects `.pdf/.xps`).
- **Q6 — Approval/audit?** **MUTATING**; audit via existing redacted JSONL (`session_service._record_audit:1215-1233`) with `tool="excel_save_as"`, `params={workbook_id, target_path, overwrite_existing, overwrite_source}`, `result={target_path, sha256, size, backup_path, source_overwritten, outcome}`.

### 7.3 Overwrite / source-overwrite flag design (HIGH)

Two **distinct, explicit** booleans — **not** a vague `confirm`:

| Flag | Guards | Default | Effect when true |
|---|---|---|---|
| `overwrite_existing` | target file already exists (and target ≠ source) | `False` | Permit replacing an existing target (old file → `.bak`). |
| `overwrite_source` | target **is** the allowlisted source (`normcase(canonical_t)==normcase(record.canonical)`) | `False` | Permit the narrow source-overwrite; still requires source-identity match (§4) and still produces a `.bak` of the source. |

Decision matrix (target canonical vs source canonical):

| target == source? | `overwrite_source` | `overwrite_existing` | source hash | Result |
|---|---|---|---|---|
| no | any | false (target absent) | any | **write** (new file) |
| no | any | false (target exists) | any | **reject** `PROTECTED_TARGET/target_exists` |
| no | any | true (target exists) | any | **write** + backup old target |
| yes | false | any | any | **reject** `PROTECTED_TARGET/source_protected` |
| yes | true | (implied) | **match** | **write** (source-overwrite) + backup source |
| yes | true | (implied) | **mismatch** | **reject** `REVISION_CONFLICT/source_modified_since_open` |

`overwrite_source=True` implies overwrite of the source slot (no separate `overwrite_existing` needed for the source case). The two flags are orthogonal for the non-source case.

### 7.4 Path / identity safety algorithm (HIGH, end-to-end)

```
1  record = require_record(workbook_id)                     # WORKBOOK_NOT_FOUND if absent
2  require writable (not read_only)                          # WORKBOOK_READ_ONLY
3  canonical_t = path_policy.validate_path(path, MUTATING)   # PATH_NOT_ALLOWED (allowlist/ext/symlink/UNC/device)
4  if normcase(canonical_t) == normcase(record.working_path): reject PROTECTED_TARGET/target_is_working_copy
5  src = Path(record.canonical);  re_canonical_src = validate_path(record.canonical, READ_ONLY)
6  if normcase(re_canonical_src) != normcase(record.canonical): reject REVISION_CONFLICT/source_replaced
7  src_hash_now = sha256(src)
8  if src_hash_now != record.source_sha256: reject REVISION_CONFLICT/source_modified_since_open
9  same = normcase(canonical_t) == normcase(record.canonical)
10 if same and not overwrite_source: reject PROTECTED_TARGET/source_protected
11 exists = canonical_t.exists()
12 if exists and not (overwrite_existing or same): reject PROTECTED_TARGET/target_exists
13 approve MUTATING on canonical_t (and record.canonical if same)   # APPROVAL_DENIED
14 wb.Save() on STA (flush working copy)                             # in place
15 stage tmp in canonical_t.parent; copy working->tmp; fsync; verify size+sha
16 if exists: os.replace(canonical_t -> backup_path)                # preserve old
17 os.replace(tmp -> canonical_t); verify final size+sha
18 audit + return SaveAsResult(outcome="ok", target_path, sha256, size, backup_path, source_overwritten=same)
```

All FS steps (15-17) run on the event loop / a worker thread (no COM); only step 14 touches COM (STA). Steps 3-13 are non-COM guards that fail fast before any write.

### 7.5 Result DTO (proposed `SaveAsResult`, additive, `extra="forbid"`)

New class in `domain/models.py` (do **not** bump `SCHEMA_VERSION`; additive):

| Field | Type | Meaning |
|---|---|---|
| `workbook_id` | str | unchanged session id (session not rebound) |
| `target_path` | str | canonical absolute path written |
| `source_path` | str | original allowlisted canonical (`record.canonical`) |
| `source_overwritten` | bool | `same` (target was the source) |
| `overwrote_existing` | bool | a prior file was replaced |
| `backup_path` | str \| None | `.bak` path when a prior file was preserved, else `None` |
| `sha256` | str | SHA-256 of the written target |
| `size` | int | bytes written |
| `source_sha256` | str | source hash at open (baseline) — optional, see §10 |
| `outcome` | Literal["ok","unknown"] | `ok` only after final verify; `unknown` if interrupted/unprovable |
| `mode` | SessionMode | inherited source session mode (Q3) |
| `notes` | list[str] | diagnostics (e.g. "session remains bound to working copy") |
| `created_at`/`updated_at` | datetime | via `TimestampedModel` |

`outcome="unknown"` is the fail-closed signal (§9): never fabricate `ok`.

---

## 8. Failure / interruption matrix (HIGH)

| Failure | Detection point | Destination preserved? | Reported outcome | Error / code |
|---|---|---|---|---|
| Path outside roots / bad ext / symlink / UNC / device | step 3 `validate_path` | n/a (no write) | reject | `PATH_NOT_ALLOWED` |
| Target == working copy | step 4 | n/a | reject | `PROTECTED_TARGET` |
| Source replaced by symlink after open | step 6 | n/a | reject | `REVISION_CONFLICT` |
| Source externally modified | step 8 | n/a | reject | `REVISION_CONFLICT` |
| Source-overwrite without flag | step 10 | n/a | reject | `PROTECTED_TARGET` |
| Existing target without flag | step 12 | n/a (untouched) | reject | `PROTECTED_TARGET` |
| Approval denied (restricted) | step 13 | n/a | reject | `APPROVAL_DENIED` |
| `wb.Save()` fails (STA/COM) | step 14 | working copy only; target untouched | error | `COM_BUSY`/`EXCEL_HUNG` per worker |
| Disk full / permission denied staging | step 15 | **untouched** (temp unlinked) | error | `ArtifactWriteError`-style (`SNAPSHOT_FAILED` reuse) |
| Staged verify (size/hash) mismatch | step 15 | untouched | error | typed write error |
| Failure after old→backup, before tmp→dest | step 16→17 gap | old file at `backup_path`; dest absent | **unknown** (recoverable) | `outcome="unknown"` + `backup_path` |
| Final verify mismatch | step 17 | attempt rollback from `.bak` | **unknown** / error | `outcome="unknown"` |
| Process crash mid-copy | any | depends on step | **unknown** on next call (detect stray `.tmp`/`.bak`) | reported, never `ok` |

**Rule (HIGH):** success (`outcome="ok"`) is reported **only** after step-17 final size+hash verification. Any unprovable state is `unknown` with a suggested action; the server never claims success from "no exception" (mirrors the truth-signal discipline, glossary "truth signal"; `snapshot_store` atomicity doc `:22-26`).

---

## 9. Test matrix (14 cases)

Unit tests inject a fake STA worker + fake `Application`/`Workbook` (pattern: `tests/unit/test_session_service.py:79-130`), a real `PathPolicy`/`ApprovalPolicy`/`AuditLogger` over a temp dir. Save-As FS steps run against real temp files.

| # | Case | Setup | Expected |
|---|---|---|---|
| 1 | New-path write | target absent, `overwrite_*=False` | `outcome=ok`; target exists; `sha256`==working; `source_overwritten=False`; `backup_path=None`; original unchanged |
| 2 | Existing target, no flag | target exists (≠source) | reject `PROTECTED_TARGET/target_exists`; target bytes unchanged |
| 3 | Existing target, `overwrite_existing=True` | target exists | `outcome=ok`; old target → `.bak` (exists, bytes==old); target==new; `overwrote_existing=True` |
| 4 | Source-overwrite, no flag | target==source, `overwrite_source=False` | reject `PROTECTED_TARGET/source_protected`; source unchanged |
| 5 | Source-overwrite, flag, source unchanged | target==source, `overwrite_source=True`, hash match | `outcome=ok`; `source_overwritten=True`; source replaced; `.bak` holds old source; working copy still open |
| 6 | Source-overwrite, source externally modified | target==source, hash mismatch (mutate file after open) | reject `REVISION_CONFLICT/source_modified_since_open`; source NOT overwritten |
| 7 | Path outside roots | target outside allowlist | reject `PATH_NOT_ALLOWED` (reason `outside_roots`); no file |
| 8 | Bad extension | target `out.pdf` | reject `PATH_NOT_ALLOWED` (`extension_not_allowed`) |
| 9 | Target == working copy | target = `record.working_path` | reject `PROTECTED_TARGET/target_is_working_copy`; working copy intact |
| 10 | Session not rebound | after any successful Save As | `record.workbook_id/working_path/canonical` unchanged; subsequent `excel_save` writes working copy (target hash unchanged by the later save) |
| 11 | Disk-full / permission staging failure | simulate `copy2`/`fsync` OSError at step 15 | error (typed); destination untouched; no leftover `.tmp` |
| 12 | Interrupted between backup and install | simulate failure after step 16 | `outcome=unknown`; `backup_path` set; destination absent; suggested action present; retry does not orphan backup |
| 13 | Final verify mismatch | tamper staged/target before step-17 verify | `outcome=unknown` (or rollback from `.bak`); never `ok` |
| 14 | Approval denied (restricted) | RESTRICTED mode, no grant, `overwrite_source=True` | reject `APPROVAL_DENIED`; no FS change; audit entry `allowed=false` |

Coverage rationale: 1-6 = contract/flags/identity; 7-9 = path safety; 10 = session-binding invariant; 11-13 = atomicity/failure; 14 = approval. Add an integration (`tests/integration_excel`) case for a real `wb.Save()`+copy round-trip if the build env allows (L3).

---

## 10. Unresolved questions for the Auditor / Validator

1. **Overwrite-rejection error code (Q4).** Confirm `PROTECTED_TARGET` for "existing target / source protected" vs minting a dedicated code (forbidden by `SCHEMA_VERSION=1` closed set). Is `PROTECTED_TARGET` acceptable, or should `PATH_NOT_ALLOWED`/`STALE_TARGET` be used?
2. **`source_sha256` on `WorkbookInfo`?** Keep internal to `_SessionRecord` (no DTO change) or expose additively on `WorkbookInfo` for caller-visible identity? (Brief says "optionally".)
3. **Backup retention.** Keep exactly one newest `.bak` per target (prune older) or retain all? Where may backups accumulate (destination dir is user data — is writing `.bak` beside user files acceptable, or must backups go to an artifact root)?
4. **Source hash at open — cost.** Is hashing the whole source at every `excel_open` acceptable (workbooks can be large), or should it be stat-only at open + hash only at Save-As? (Recommendation: hash at open; confirm.)
5. **`overwrite_source` + `overwrite_existing` interaction.** Confirm the matrix in §7.3 (source-overwrite implies overwrite; two orthogonal flags otherwise) matches product intent, and that `overwrite_source` is the *only* sanctioned path to modify an allowlisted original.
6. **Interrupted-state recovery.** Is "destination absent + `.bak` present ⇒ report unknown, require explicit retry" the desired UX, or should the server auto-resume (rename `.bak` back) on the next call? Auto-resume risks masking a partial write.
7. **Concurrency.** Save As to a target while another `excel_open` of the same target is in flight — the per-canonical lock (`:294`) is keyed on the **source**, not the target. Should Save As acquire a target-keyed lock, or is target-write racing out of scope?
8. **Read-only / inspect sessions.** Confirm Save As is allowed from an `inspect`/read-only session (it reads the working copy, writes a new file) or must require a writable session. (Recommendation: allow from read-only — no working-copy mutation — but confirm.)
9. **Fixed-format/PDF** remains a separate future tool (§7.2 Q5) — confirm no PDF/XPS expectation is attached to this work package.
10. **Annotations.** Should `excel_save_as` carry `destructive_hint` (it can overwrite)? Annotations are static per tool name (`server.py:1835-1842`), so the hint cannot depend on the flags — confirm leaving it off and relying on the description + flags.