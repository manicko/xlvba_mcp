# Large-Data Architecture — Independent Auditor Review

Status: READ-ONLY source review. Verdict is advisory. Sources cited as `path:line`
against the working tree at commit `41a4a88` plus staged WP-A edits.

## Verdict: PROCEED WITH CHANGES

The layered, additive design is fundamentally sound and reuses the right primitives
(single-STA FIFO, `JobRegistry` async path, `SaveCopyAs` snapshot, atomic artifact store,
path policy). The COM/OOXML split is coherent, the per-response memory bound is real, and
the WP-A DTO/error/config edits are correct and backward-compatible. However, several
design-level gaps must be closed **before** WP-E/WP-F implementation, or the large-data
path will either block the server event loop, silently mislabel a page's version, or claim
reuse of a store that cannot stream. None require re-architecture.

---

## Issues (severity — `path:line` — required fix)

### 1. HIGH — OOXML "async" job will block the single asyncio event loop
`large_data_architecture_proposal.md:78-91`, `large_data_execution_contract.md:102-112`.
The plan schedules parsing as a `JobRegistry` async callable (`job_registry.py:289-292`
`_run_async`) and asserts "no STA". But a synchronous `iterparse` of a 100 MB sheet run
inline in the coroutine is CPU-bound work on the **loop thread**: it stalls all other async
jobs, `job_get` polling, cooperative cancellation, and every MCP response — the STA is idle
but the server is frozen. `asyncio.to_thread` is already used elsewhere
(`session_service.py:1342`, `com_worker.py:560,579`), so the pattern exists.
**Fix:** run the CPU-bound parse in `asyncio.to_thread`/a bounded executor (still no COM,
no second STA), bounded by `max_file_jobs`, checking `cancel_requested` between chunks and
yielding. Add an acceptance test that the loop stays responsive during a parse.

### 2. HIGH — Version capture is TOCTOU relative to the FIFO (single-page mislabel)
`session_service.py:1095-1103` (`get_version` is non-COM), `ranges_service.py:888-907`
(read captures nothing; version would be read before `submit`). All COM calls serialize
through one FIFO (`com_worker.py:376`), but `get_version` runs **outside** that ordering. A
mutation queued after the page's `get_version` but before the read executes will run first
(FIFO), so the page reflects post-mutation data yet is tagged with the pre-mutation version.
A following page detects the mismatch (`STALE_CURSOR`) — safe — but a **single-page** read
(`has_more=false`) is never re-checked and is silently mislabeled.
**Fix:** capture the version **inside** the submitted STA `fn` (atomic with the read), or
bump inside the mutation `fn` on the STA thread, so version and data share one serialization
point. Document the residual single-page window if not fixed.

### 3. MED — `bump_version` is defined but wired nowhere; 8 separate choke points
`session_service.py:1105-1113` defines it; `grep bump_version` shows **zero** callers. Each
service has its own `_submit_mutation` (`ranges_service.py:842`, `tables_service.py:455`,
`sheets_service.py:404`, `names_service.py:337`, `shapes_service.py:242`,
`controls_service.py:248`, `pivot_tables_service.py:823`, `vba_project_service.py:645`), so
wiring is one line in eight places, not one. The contract list
(`large_data_execution_contract.md:76-78`) omits VBA-component mutation and
`snapshot_restore`; restore changes `workbook_id` (`snapshot_service.py:397` →
`session_service.py:376` new uuid) so old cursors fail as `WORKBOOK_NOT_FOUND` (acceptable),
but this should be stated. **Fix:** wire all eight `_submit_mutation` success paths, add the
R2 unit test that mutates via each service and asserts the version increments, and document
the restore/identity behavior.

### 4. MED — WP-F cannot "reuse ArtifactStore": it is bytes-in-memory and `.zip`-only
`large_data_architecture_proposal.md:28,33,93-99` claim reuse of `ArtifactStore` atomic
write. But `artifact_store.py:208-211` `save_artifact(name, data: bytes, ...)` takes the
**whole payload in memory**, and `artifact_store.py:78` maps `ArtifactType.EXPORT → ".zip"`
only. A streamed CSV/JSONL export can do neither. **Fix:** add a streaming artifact writer
(temp file → `fsync` → `os.replace`, reusing `path_for`/`_validate_target`) and a new
`ArtifactType`/extension allowance for `.csv`/`.jsonl`; stop claiming `save_artifact` reuse.

### 5. MED — Temp-file cleanup on `interrupted_unknown` has no owner
`large_data_execution_contract.md:110-111` promises deterministic cleanup, but
`job_registry._run_async` (`job_registry.py:469-507`) performs no filesystem cleanup, and a
disconnect does not cancel the task (registry owns it — good), while a server restart leaves
the job `interrupted_unknown` with its temp file orphaned. **Fix:** the export service owns a
`finally`-unlink plus a startup sweep of stale temp files under the artifact root; name the
owner in the contract.

### 6. MED — ZIP/XML guards: ratio has no size floor and is weaker than the reference
`config.py:121-124` sets `max_compression_ratio=1000`; the reference uses 500 with a 1 MiB
floor (`large_data_reference_implementations.md:26`). A naive ratio check with no floor
false-positives on small highly-compressible XML members, and `max_xml_depth`
(`config.py:133`) is not enforceable with stock `iterparse`. **Fix:** apply the ratio check
only above a size floor (e.g., 1 MiB), reconsider the 1000 default toward 500, specify the
depth-enforcement mechanism (custom target/expat), and enforce limits **per-member during**
streaming, not only up front.

### 7. MED — New page DTOs must each enforce their cap; job results bypass `_ensure_dto`
`RangeValues`/`RangeFormulas` enforce the cap (`models.py:1033-1042,1076-1085`) and
`TableRows` uses `_MAX_RESPONSE_CELLS=2000` (`tables_adapter.py:70`, to be raised to
`max_table_cells`). The new `PivotResultPage` and OOXML page DTOs must replicate a
`model_validator` for `max_pivot_result_cells`/`max_response_cells`. Also the worker
return-boundary guard `_ensure_dto` (`com_worker.py:197`) only covers COM returns; OOXML
pages returned via the job result are not guarded. **Fix:** add cap validators to every new
page DTO and a job-result size guard.

### 8. MED — Snapshot lifetime vs in-flight OOXML job (GC/retention race)
`SnapshotManifest.stored_path` (`models.py:542`) is the immutable input
(`large_data_execution_contract.md:107-108`), but a long paging job can outlive snapshot
retention/GC, deleting the file mid-paging. **Fix:** pin the snapshot until the job reaches a
terminal state; document retention ≥ max job lifetime.

### 9. LOW — `LargeDataConfig` has no `gt=0`; env can set 0/negative
`config.py:85-144`. `EXCEL_VBA_MCP_LARGE_DATA_MAX_RESPONSE_CELLS=0` (or `< max col width`)
makes `_window` compute `max_rows=0` (`ranges_adapter.py:3772`) → non-progressing empty
pages. **Fix:** add `gt=0` to all limits and a guard that `max_response_cells ≥
MAX_RANGE_COLS_HARD` (16,384).

### 10. LOW — `PageInfo.reason` enum overlaps the exception path
`models.py:995-1004` lists `stale_cursor`/`invalid_page`, but the contract raises
`STALE_CURSOR`/`InvalidRangeError` instead of returning a page
(`large_data_execution_contract.md:80-82,94-96`). **Fix:** reconcile doc/DTO so a raised
error and a returned reason are not both implied.

### 11. LOW — `max_range_cells` (100k) "per request window" vs 50k response cap
`large_data_execution_contract.md:18-19`. The effective read window is
`min(max_range_cells, max_response_cells) = 50k`; the table implies 100k. **Fix:** state the
effective window is the response cap, or align the defaults.

### 12. LOW — Traversal check must cover Windows forms
`path_policy.py:182` checks suffix; the OOXML member-traversal rejection must normalize
backslash, absolute, and drive-relative entry names on Windows (reference does `resolve()` in
`opc.py`). **Fix:** normalize before the allowlist check.

### 13. LOW — Benchmark should split queue-wait from execution
`large_data_implementation_plan.md:66-67`. The worker already distinguishes queue-wait vs
execution (`com_worker.py:389-400`); "time-to-first-page" should report both, not a blended
wall-clock. Otherwise honest.

---

## WP-A staged-edit sanity check — PASS (minor notes)

- `LargeDataConfig` (`config.py:77-144`) + env loader (`config.py:188-196`): correct,
  iterates `model_fields`, 1:1 with contract. Note issue #9 (`gt=0`).
- `STALE_CURSOR` (`errors.py:93`): appended last (additive, no reorder); mapped in
  `_CATEGORY_BY_CODE`→`validation` (`:188`), absent from `_RETRYABLE_CODES` (correct),
  has `_ACTIONS_BY_CODE` (`:279-282`), `StaleCursorError` (`:585-595`), and `_ERROR_BY_CODE`
  (`:765`). `category_for`/`retryable_for`/`actions_for` all total. Correct.
- `PageInfo` (`models.py:964-1004`), `RangeValues.page` (`:1028`), `RangeFormulas.page`
  (`:1071`), `TableRows.page` (`tables_adapter.py:120`): additive optional fields with
  defaults; `DomainModel` `extra="forbid"` (`models.py:71`) is inherited safely; Pydantic
  field ordering (required-after-optional) is not a dataclass hazard. Correct.
- `SessionRecord.version` + `get_version`/`bump_version` (`session_service.py:111,1095-1113`):
  correct primitive; wiring is issue #3.

---

## Accepted risks (explicit)

- Single-STA OLAP/Data-Model **member/CubeField enumeration** stays fail-closed this pass;
  only the displayed-result read (cells, no external resolve) is enabled — evidence is
  inferred, not executed on our runtime (`large_data_reference_implementations.md:54-68`).
- `SaveCopyAs` is a point-in-time input; OOXML reads are intentionally snapshot-staleness-
  bound (`source="snapshot"`, `sha256`-pinned), never live.
- ~2–3× transient Python duplication per page (`_normalize_grid`/`_classify_grid` +
  `_ensure_dto` `model_dump`) is accepted as bounded per page.
- Very wide sheets (cols near 16,384) page only a few rows per response — chatty but bounded;
  `total` is reported so the agent can plan.
- CSV/JSONL streaming export (WP-F) and all OOXML reading (WP-E) are unimplemented; this
  review validates the design, not the not-yet-written code.