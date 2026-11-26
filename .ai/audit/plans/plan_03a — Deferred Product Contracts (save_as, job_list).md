# Plan 03a — Deferred Product Contracts (`save_as`, `job_list`)

- **Role:** Researcher (U5 — deferred product decisions). **DOCUMENT ONLY — no production/test/config change.**
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca01fd10136716a598cad1db6b6845fed73`.
- **Parent plan:** `.ai/audit/plans/plan_03 — Agent Usability Remediation.md` §14 (U5), §18 (deferred items).
- **Binding decision:** `decision_log.md` §"D-U9" (`decision_log.md:200-204`) — *investigate only; do NOT implement in this remediation; `save_as` needs a separate product contract (original protection, output path policy, overwrite, result retrieval).*
- **Authority order:** `decision_log` D-U9 > plan_03 §18 > source at HEAD. Source line citations below are authoritative over prose.

> **NOT IMPLEMENTED — requires separate product approval.** Nothing in this document authorizes a code, test, or config change. Both `save_as`/export and `xlvba_job_list` are **contract outlines only**. No tool is added by this work package (plan_03 §14 U5 "Non-scope: implementing any tool; changing `excel_save`"; §12.5 type **L** = tool-lifecycle, deferred).

---

## 0. Grounding: what exists today (source-verified at HEAD)

These are the facts every contract clause below is derived from. Each is cited `file:line`.

### 0.1 `excel_save` is in-place only (no path argument)

- Handler: `excel_save(workbook_id)` → `service.save_workbook(workbook_id)` (`src/excel_vba_mcp/api/workbook_tools.py:147-159`). There is **no `path` parameter** in the signature (`workbook_tools.py:147`).
- Service: `SessionService.save_workbook` calls `wb.Save()` on the open working copy inside `alerts_suppressed()` and returns a refreshed `WorkbookInfo` (`src/excel_vba_mcp/application/session_service.py:456-520`, the `wb.Save()` at `session_service.py:494`). It clears `record.dirty` (`session_service.py:510`) and re-describes the workbook (`session_service.py:495-501`).
- Description today: `_DESC_SAVE = "Explicitly save an open workbook's working copy."` (`src/excel_vba_mcp/server.py:226`).
- **Conclusion:** persistence today is strictly *save-in-place of the working copy*. There is no save-as / new-file path anywhere (`grep job_list|save_as` over `src` → 0 tool hits; the only `SaveCopyAs` uses are the snapshot branch, `snapshot_service.py:219`, and `storage/snapshot_store.py`).

### 0.2 Original protection is already structural (working-copy model)

- `open_workbook` copies the allowlisted original into `work/` and opens **only the copy**; the original is never modified (`session_service.py:205-331`, copy at `_materialize_working_copy` `session_service.py:1153-1160`, `shutil.copy2` at `session_service.py:1159`). Module docstring states "Copies, never the original" (`session_service.py:11-15`).
- Any new-file writer must preserve this invariant (plan_03 §4 constraint 2: "Originals never modified; working-copy model").

### 0.3 PathPolicy is the filesystem boundary (allowlist + canonicalisation)

- `PathPolicy.validate_path(raw, operation, *, expected_root=None)` canonicalises and enforces: allowlisted roots only, symlink/reparse escape rejection, `..` traversal rejection, UNC/device denial, and an **extension allowlist** (`src/excel_vba_mcp/security/path_policy.py:110-185`; extension gate at `path_policy.py:182-183`).
- Permitted workbook extensions are a fixed set `WORKBOOK_EXTENSIONS` = `.xlsm/.xlsb/.xlsa/.xlam/.xltm/.xltx/.xlsx/.xls` (`path_policy.py:38-49`). Artifact roots additionally allow `.json/.sha256` (`path_policy.py:52-56`, `187-190`).
- Rejections raise `PathNotAllowedError` / code `PATH_NOT_ALLOWED` with a safe `reason` and, for outside-root, `context["allowed_roots"]` (`path_policy.py:168-213`). The module performs **no filesystem writes** (`path_policy.py:22`).
- **Office-trust application** is *not* in PathPolicy. Macro trust is applied at **open** time via `automation_security_for(mode)` inside the adapter context manager on the STA thread (`session_service.py:348`, `353`; `security/macro_policy.py`). PathPolicy governs *where* a path may be; Office-trust governs *how Excel opens it*.

### 0.4 Snapshot gate is in-memory and mutation-scoped

- `SnapshotService.require_snapshot(workbook_id)` reads the in-memory `_latest_snapshot` dict and raises `SNAPSHOT_FAILED` when absent (`src/excel_vba_mcp/application/snapshot_service.py:400-425`, the dict at `snapshot_service.py:120`, the raise at `snapshot_service.py:419-424`).
- `has_snapshot(workbook_id)` reads the same in-memory gate (`snapshot_service.py:427-430`).
- The gate exists to force a pre-mutation rollback point for **in-place mutations** of an open workbook (docstring `snapshot_service.py:24-29`). A snapshot create itself is approval-gated but does **not** require a pre-existing snapshot (`snapshot_service.py:162-166`).

### 0.5 Job registry is in-memory, server-lifetime (job store is optional FS persistence)

- `JobRegistry` holds jobs in a plain dict `self._jobs: dict[str, Job] = {}` (`src/excel_vba_mcp/application/job_registry.py:232`), guarded by a `threading.Lock` (`job_registry.py:231`). Lifecycle `queued → running → {succeeded|failed|expired|unknown|interrupted_unknown}` (`job_registry.py:18-27`, states `job_registry.py:71-92`).
- An **optional** `JobStore` may persist records to disk (RUNNING marker + finalize) so a *restarted* process can serve a terminal result by `job_id` (`job_registry.py:33-37`, `218-221`, `299-307`; `storage/job_store.py:1-28`). The store is a *fallback for restart*, not the live registry; the live registry is process-local and in-memory.
- `job_get(job_id)` returns a `JobStatus` DTO (`src/excel_vba_mcp/api/job_tools.py:96-110`; DTO `job_registry.py:173-189`). `JobStatus` fields: `job_id, status, started, cancel_requested, poll_after_s, result, error_code, error_text, retry_after_s, phase, outcome, retryable, events` (`job_registry.py:173-189`). `RunningJob` (the start handle) is `{status, job_id, poll_after_s, result, error_code, error_text, retry_after_s, outcome, retryable}` (`job_registry.py:192-207`).
- The registry is published to the tools via `set_job_registry` (`job_tools.py:27-42`) and wired at server lifespan (`server.py:1078`, cleared `server.py:1097`).

### 0.6 DTO / schema discipline

- `SCHEMA_VERSION = 1` (`src/excel_vba_mcp/domain/models.py:62`); all domain DTOs derive from `DomainModel` with `model_config = ConfigDict(extra="forbid", use_enum_values=False)` (`models.py:65-68`). Any new DTO must be a new class with `extra="forbid"` and additive fields only (plan_03 §4 constraint 7).
- `WorkbookInfo` is the existing per-workbook identity DTO: `workbook_id, name, format, path, source_path, read_only, has_vba, excel_version, session_id, mode, model_version` (`models.py:83-96`).
- Tool registration pattern is `server.add_tool(fn, name=..., description=...)` (`server.py:1479-1539`), e.g. `excel_save` at `server.py:1481`, `job_get`/`job_cancel` at `server.py:1538-1539`.

### 0.7 VBA-only export already exists

- `vba_export_project(workbook_id, out_dir=None)` exports **VBA components only** (`.bas/.cls/.frm/.frx`) and returns a `VbaExportManifest` (`src/excel_vba_mcp/api/vba_tools.py:132`; service `application/vba_project_service.py:192-243`; DTO `models.py:242-251`). A caller `out_dir` is validated against the artifact roots by `PathPolicy` (`vba_project_service.py:227-231`, `530-564`) and is approval-gated (`vba_project_service.py:214-225`). This is **not** a workbook-file export.

---

## 1. `save_as(path)` / export — product contract outline

> Status: **NOT IMPLEMENTED.** This section is the decision scaffold a product owner must ratify before any build. Each clause cites the source invariant it must honor.

### 1.1 Purpose & signature (proposed)

- Proposed shape: `save_as(workbook_id, path, *, overwrite=False, mode=..., read_only=...) -> WorkbookInfo`-like DTO. Thin handler in `api/workbook_tools.py` (composition; no logic in `server.py`), delegating to a new `SessionService.save_as_workbook(...)` — mirroring the thin-handler pattern of `excel_save` (`workbook_tools.py:147-159`) and the no-COM rule (`workbook_tools.py:1-17`).
- **Relationship to `excel_save` (unchanged):** `excel_save` stays *in-place, working-copy only, no path* (`workbook_tools.py:147`; `session_service.py:456-520`). `save_as` is **additive** and must not alter `excel_save` semantics (plan_03 §18.1 "Do not change `excel_save` in-place semantics"; §14 U5 non-scope).

### 1.2 Original protection (mandatory)

- `save_as` writes a **new file**; it must never mutate the original allowlisted source. This is already structurally guaranteed for the open path (original copied to `work/`, `session_service.py:1153-1160`) and must be preserved: the new file is produced from the **working copy**, and the original on disk is untouched.
- The output file's bytes must be hashed before/after to prove the original fixture is unchanged (the existing originals-safe harness pattern, plan_03 §15.2 "Originals-safe: fixtures SHA-256 unchanged").

### 1.3 Output path policy (mandatory)

- The `path` argument MUST pass `PathPolicy.validate_path(path, OperationClass.MUTATING)` **against the OUTPUT path** (not the source). This yields: allowlisted-roots-only, symlink/traversal/UNC/device rejection, and the workbook-extension allowlist (`path_policy.py:110-185`, extensions `path_policy.py:38-49`). A non-workbook extension (e.g. `.pdf`) is rejected by `EXTENSION_NOT_ALLOWED` (`path_policy.py:182-183`) — see §1.7.
- **Office-trust application to the OUTPUT path:** today Office macro trust is chosen by `mode` at *open* (`session_service.py:348`, `353`). For `save_as` the output is *written*, not opened, so the contract must decide explicitly whether the saved file's macro trust is inherited from the source session's mode or re-evaluated. Recommendation: **inherit the source session's mode** (the file is a byte copy of an already-open, already-trust-decided working copy) and record the decision in the audit entry. This is a product decision (see §4 Q3), not a silent default.
- On rejection, the error is `PATH_NOT_ALLOWED` with `reason` + `allowed_roots` (`path_policy.py:168-174`), consistent with `excel_open` (`session_service.py:246`).

### 1.4 Snapshot-gate interaction (decide explicitly)

- The snapshot gate (`require_snapshot`, `snapshot_service.py:400-425`) exists to guarantee a rollback point **before an in-place mutation** of an open workbook (`snapshot_service.py:24-29`).
- **Decision required (do not silently require or skip — plan_03 §18.1).** Recommended contract clause: **`save_as` SKIPS the snapshot gate**, with rationale:
  1. `save_as` does **not** mutate the open working copy; it reads it and writes a *separate* new file (like the snapshot SaveCopyAs branch, which is `READ_ONLY` and does not require a pre-snapshot — `snapshot_service.py:210-228`, op_class `READ_ONLY` at `snapshot_service.py:227`).
  2. The original is protected structurally (§1.2), so there is no in-place state to roll back.
  3. Requiring a snapshot would impose an in-place-mutation precondition on a non-mutating, new-file write, contradicting the gate's stated purpose.
- The alternative (require snapshot) must be rejected in the contract if the product owner agrees `save_as` is non-mutating; the choice and its rationale must be written into the tool description, not left implicit.

### 1.5 Overwrite policy (mandatory)

- Reject an existing target **unless** an explicit `overwrite=True` flag is passed — mirroring the range-write overwrite discipline (plan_03 §18.1 "reject existing targets unless an explicit `overwrite` flag is passed (mirror `range` overwrite discipline)"; the `range` tools default `overwrite_policy='reject-nonempty'`).
- Implementation note for the future build: the existence check is a pure filesystem `Path.exists()` on the canonicalised output (PathPolicy already returns the canonical path, `path_policy.py:185`); a rejected overwrite raises a domain error (additive `ErrorCode` reuse — plan_03 §4 constraint 7 forbids new `ErrorCode` members, so reuse an existing code such as `INVALID_RANGE`/`NOT_FOUND` family as the product owner directs). Never silently clobber.

### 1.6 Result retrieval & working-copy switch (decide explicitly)

- Return a **`WorkbookInfo`-like DTO** for the saved copy (new path + identity), consistent with `excel_save` returning `WorkbookInfo` (`workbook_tools.py:147-159`; DTO `models.py:83-96`). If a new DTO is needed it must be a new `extra="forbid"` class with additive fields (`models.py:65-68`; plan_03 §4 constraint 7).
- **Does the working copy switch to the new path? — product decision (plan_03 §18.1).** Two options, both must be stated in the contract:
  - **(A) No switch (recommended):** the open session keeps its current `workbook_id`/working path; `save_as` is a one-shot export of the current state. Simpler; preserves the "one working copy per open" bookkeeping (`session_service.py:307-321`). The returned DTO describes the *new file on disk*, not a new open session.
  - **(B) Switch:** close the current working copy and reopen from the new path (like `restore_snapshot` does, `snapshot_service.py:331-343`). Heavier; changes `workbook_id`; risks surprising the caller.
- Recommendation: **(A) no switch**; the new file is a persisted artifact, the session is unchanged. Record the decision in the description.

### 1.7 Fixed-format / PDF export (separate decision)

- A fixed-format export (PDF/XPS via `ExportFixedFormat`) is a **distinct product decision**, not part of the workbook `save_as`. Rationale from source: PathPolicy's workbook-extension allowlist rejects `.pdf`/`.xps` (`path_policy.py:38-49`, `182-183`), so a fixed-format export needs its own artifact-root/extension handling and its own approval semantics.
- `vba_export_project` already covers **VBA-only** export (`vba_tools.py:132`; `vba_project_service.py:192-243`) and is the correct tool for code export; it is **not** a workbook or PDF export.
- Contract position: `save_as` (workbook file) and a future fixed-format/PDF export are **two separate tools/contracts**; do not overload `save_as` to emit PDF. Fixed-format export is deferred to its own product decision (plan_03 §18.1 "a fixed-format/PDF export is a separate decision").

### 1.8 Approval / audit

- `save_as` writes a new file → it is a `MUTATING` filesystem write for approval purposes, exactly like `excel_save` (`session_service.py:477-480`) and `snapshot_create` (`snapshot_service.py:162-166`, `456-486`). In default `FULL` mode it is auto-allowed; in `RESTRICTED` it requires an allowlist entry / live grant (plan_03 §2 access model). Audit via the existing redacted JSONL logger (`session_service.py:1169-1187`).

---

## 2. `xlvba_job_list` — product contract outline

> Status: **NOT IMPLEMENTED.** Read-only outline only (plan_03 §18.2; D-U9 `decision_log.md:200-204`).

### 2.1 Scope (read-only, current-process only)

- **Read-only:** `readOnlyHint: true`; never touches COM; never creates/starts/cancels jobs. Mirrors `job_get` (non-COM, `job_tools.py:96-110`) and the read-only tool class (plan_03 §11 read-only list includes `job_get`).
- **Current-process jobs only:** the live registry is the in-memory `self._jobs` dict (`job_registry.py:232`), server-lifetime. `xlvba_job_list` enumerates **that dict**, not the optional on-disk `JobStore` (`job_registry.py:33-37`; `storage/job_store.py`). It MUST NOT present restart-persisted records as if they were live in this process — the store is a `job_get`-by-id fallback after restart (`job_registry.py:299-307`), not a listable live set. (If the product owner later wants restart-visible jobs listed, that is a separate decision; the default contract is current-process only, per plan_03 §18.2 "current-process jobs only (job registry is in-memory, server-lifetime)".)
- **No secrets / no paths beyond `job_get`:** each entry exposes only the fields `JobStatus` already exposes (`job_registry.py:173-189`) — `job_id, status, started, cancel_requested, poll_after_s, result, error_code, error_text, retry_after_s, phase, outcome, retryable, events`. It must not add filesystem paths, raw results beyond what `job_get` returns, or internal keys. This matches plan_03 §18.2 "no secrets/paths beyond what `job_get` returns".

### 2.2 DTO (additive, SCHEMA_VERSION=1)

- Return `list[JobStatus]` (reuse the existing DTO, `job_registry.py:173-189`) — **no new DTO required**, which is the most additive-safe option (`extra="forbid"` respected, `models.py:65-68`). If a wrapper list DTO is preferred, it must be a new `extra="forbid"` class with `schema_version: int = SCHEMA_VERSION` (`models.py:62`) and an additive `jobs: list[JobStatus]` field; `SCHEMA_VERSION` stays `1` (plan_03 §4 constraint 7, §18.2).
- Registration (future): `server.add_tool(xlvba_job_list, name="xlvba_job_list", description=..., annotations=ToolAnnotations(read_only_hint=True))`, alongside `job_get`/`job_cancel` (`server.py:1538-1539`). Handler in `api/job_tools.py` (thin, non-COM), reading `get_job_registry()` (`job_tools.py:34-42`).

### 2.3 Behavior notes

- Ordering: stable, e.g. by `created_at` (the `Job` record carries `created_at`/`created_wall`, `job_registry.py:150-166`) — a product/UX choice to confirm.
- No pagination is required at outline stage (registry is small and process-local); if added later it must be additive.
- A `job_list` entry is **advisory**; `job_get(job_id)` remains the authoritative per-job poll (`job_tools.py:96-110`; plan_03 §8 job-handle contract).

---

## 3. Client-forwarding unknowns to confirm against the target Kilo build

These are **unverifiable from inside this session** (plan_03 §18.3). They gate whether the schema-shape and resource claims in plan_03 U1–U4 matter in practice. Recorded here so the product owner can confirm them against the real Kilo build before relying on them.

1. **`outputSchema` forwarding to the model** — whether Kilo forwards each tool's `outputSchema` to the model (determines whether schema-shape claims / U4 slimming affect model context). Source: plan_03 §18.3; baseline `outputSchema` = 75.05% of `tools/list` (plan_03 §2). **Status: unverified.**
2. **Resource meta-tools** — whether Kilo surfaces `list_mcp_resources` / `read_mcp_resource` (docs say yes; not observed live). The compact glossary mirror inside `xlvba_environment` is the reliable channel regardless; the `xlvba://glossary` resource is the forward-compatible channel. Source: plan_03 §18.3, §10, D-U8. **Status: unverified live.**
3. **`structuredContent` consumers** — whether any external consumer relies on `structuredContent` / specific nested DTO fields (no in-repo inventory). U4 must prefer `$ref` dedupe over field removal until confirmed. Source: plan_03 §18.3, §10, §14 U4. **Status: unknown inventory.**
4. **Tokenization ratios** — char→token ratios are estimates (~4:1 / ~3.5:1 lower bound); the ~113K-token figure is an estimate, not a measurement. Source: plan_03 §2, §18.3. **Status: estimate.**
5. **Real 4-part Excel build source** — only `Application.Version` = "16.0" is read (`workbook_adapter.py:100-106`); no 4-part build source exists, so `build` stays honest `null` and `target_build` is a labelled design constant. Source: plan_03 §18.3, §6.1; `workbook_adapter.py:100-106`. **Status: confirmed absent; `build=null`.**

---

## 4. Decision questions a product owner must answer (before any build)

`save_as` / export:
1. **Snapshot gate:** confirm `save_as` **skips** `require_snapshot` (recommended, §1.4) or requires it — and the rationale recorded in the description.
2. **Working-copy switch:** confirm the session does **not** switch to the new path (recommended, §1.6 A) or does (B).
3. **Office-trust for the output file:** confirm the saved file inherits the source session's mode (§1.3) vs. re-evaluating trust.
4. **Overwrite error code:** which existing `ErrorCode` to reuse for an existing-target rejection (no new members allowed, plan_03 §4.7).
5. **Fixed-format/PDF:** confirm PDF/XPS export is a **separate** future tool (§1.7), not part of `save_as`.
6. **Approval class:** confirm `save_as` is `MUTATING` for approval (auto-allowed in FULL, grant in RESTRICTED), consistent with `excel_save`/`snapshot_create` (§1.8).

`xlvba_job_list`:
7. **Scope:** confirm **current-process, in-memory only** (§2.1) — restart-persisted `JobStore` records are **not** listed (they remain reachable only via `job_get` by id).
8. **DTO:** confirm reuse of `JobStatus` (no new DTO) vs. a new additive list wrapper (§2.2).
9. **Ordering / pagination:** confirm stable ordering (e.g. `created_at`) and whether pagination is needed (§2.3).

Cross-cutting:
10. **Client-forwarding unknowns (§3):** confirm items 1–5 against the target Kilo build before relying on schema-shape/resource claims.

---

## 5. Explicit non-implementation statement

**NOT IMPLEMENTED — requires separate product approval.** This document creates no tool, no DTO, no schema change, no test, and no config change. `excel_save` is unchanged (in-place, working-copy only). `vba_export_project` remains VBA-only. `save_as`/export and `xlvba_job_list` are deferred product contracts (plan_03 §12.5 type **L**, §14 U5, §18; `decision_log.md:200-204` D-U9). Any future build requires its own work package, product-owner answers to §4, and the standard L1/L2/L3 verification path.

---

*Cite anchors @`e2289ca`. DOCUMENT ONLY — authorizes no code change.*