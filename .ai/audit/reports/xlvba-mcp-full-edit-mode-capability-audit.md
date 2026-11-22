# XLVBA MCP — Full-Edit-Mode Capability & Code Audit (FE-15 / Phase 6)

**Task:** `FE-15` (`FULL-EDIT-MODE-ORCHESTRATION-001`, final independent audit before the release gate)
**Auditor role:** independent code / architecture / security audit (read-only)
**Date:** 2026-11-22
**Repo / branch:** `C:\py_exp\xlvba_mcp` @ `master`
**Audited range:** `8c22f11..b9c13b0` (11 commits)
**Auditor's own execution:** `pytest tests/unit`, `ruff`, `mypy`, plus targeted read-only reproductions against the real production functions. Excel/COM was **not** launched by the auditor; live-Excel claims are taken from the Phase-3/4/5 evidence and cross-checked against source.

> Read-only discipline: no production, test or doc file was modified. The only file this audit writes is this report.

---

## 1. Environment

| Item | Value | Source |
|---|---|---|
| OS | Microsoft Windows 10 Enterprise LTSC 2019, NT 10.0.17763.0 (x64) | Phase-4/5 live evidence; auditor host |
| Python | 3.14.7 (`.venv\Scripts\python.exe`, MSC v.1944 64-bit AMD64) | `python -V` |
| Excel | `Application.Version = 16.0`, `Build = 14332.0` (Office 2016 family) | Phase-4/5 live evidence |
| `mcp` | 2.3.0 (`mcp.server.mcpserver.MCPServer`) | `importlib.metadata` |
| `pydantic` | 2.14.4 | `importlib.metadata` |
| `pywin32` | 312 | `importlib.metadata` |
| `comtypes` | **not installed** (was a `pywinauto` transitive; see A-2) | `importlib.metadata` → `PackageNotFoundError` |
| `pywinauto` | **removed** (0.6.9 gone from `pyproject.toml` + `uv.lock`) | `git show c28b3dd`, `Select-String uv.lock` → 0 hits |
| VBOM trust | **Assumed enabled** — Phase-3/4 live runs classified and ran macros, which requires "Trust access to the VBA project object model". Not independently re-probed by this audit. | Phase-4 report §1 |
| Access mode | `EXCEL_VBA_MCP_ACCESS_MODE` unset → default `full` | `config.py:36-40` |
| Test/runner | pytest 9.2.0, ruff 0.17.5, mypy 2.4.1 | Phase-5 evidence |

---

## 2. Scope and method

### 2.1 Commits reviewed (oldest → newest)

| # | Commit | FE | Subject |
|---|---|---|---|
| 1 | `0993370` | FE-16 | restore snapshot in execute mode so macros run after restore |
| 2 | `5d649a6` | FE-04 | add `AppActivate` / `Application.Visible` / `.Activate` / `WindowState` to the interactive-UI veto |
| 3 | `c28b3dd` | FE-07 | remove unused `pywinauto` dependency |
| 4 | `119c303` | FE-06/08/09/10/11 | document restore/worker/macro-snapshot/recovery/classify limitations |
| 5 | `e85c304` | FE-04 | note new veto tokens (EN `docs/05`) |
| 6 | `8be7245` | FE-03 | e2e coverage for Phase-4 steps 4/6/11/14 |
| 7 | `cc7abd0` | FE-02 | structured-error surfacing matrix for non-read-only tools |
| 8 | `e5ebcc8` | FE-02/05 | assert real per-tool denial codes + read-only macro/table denial |
| 9 | `db91dbd` | FE-12 | Phase-3 function matrix |
| 10 | `dece703` | FE-13 | Phase-4 E2E acceptance report |
| 11 | `b9c13b0` | FE-14 | Phase-5 hidden-execution report |

Change surface: **17 files, +1392 / −70**, of which **3 production files** (`macro_service.py` +9/−1, `snapshot_service.py` +16/−5, `server.py` +57/−6). `git diff --name-only 8c22f11..HEAD` contains no unrelated file.

### 2.2 Commands executed by the auditor

```
.venv\Scripts\python.exe -m pytest tests/unit -o addopts="" -q
  → 870 passed in 31.90s

.venv\Scripts\python.exe -m ruff check src tests
  → All checks passed!   (exit 0)

.venv\Scripts\python.exe -m mypy src/excel_vba_mcp/server.py \
    src/excel_vba_mcp/application/snapshot_service.py \
    src/excel_vba_mcp/application/macro_service.py
  → Success: no issues found in 3 source files   (exit 0)

git diff --stat 8c22f11..HEAD -- tests/fixtures/   → (empty)
git diff --stat 8c22f11..HEAD -- src/             → 3 files, +69/−13
```

Targeted read-only probes (no Excel, no COM):
`_find_interactive_tokens` behaviour, `_validation_error_as_domain` serialisability, `mcp` SDK `ToolError` cause chaining (`mcp/server/mcpserver/tools/base.py:151-161`).

---

## 3. Per-change verdict table

| # | Commit | Change | Files | Verdict | Rationale |
|---|---|---|---|---|---|
| 1 | `0993370` | `SnapshotService.restore_snapshot` opens the restored copy with `SessionMode.EXECUTE`, `read_only=False` | `application/snapshot_service.py`, `tests/unit/test_snapshot_service.py`, e2e | **ACCEPT-WITH-NOTE** | Correct and minimal; audit record, `_latest_snapshot`, `_known_sources`, immutable-snapshot and "no approval gate (R-19)" designs are intact. Note: the internal `SessionService.open_workbook` approval class changes `MUTATING` → `EXECUTING` (see C-4), and the docstring's "no approval gate" is imprecise for `restricted` mode (pre-existing). |
| 2 | `5d649a6` | 4 tokens appended to `INTERACTIVE_TOKENS` | `application/macro_service.py`, `tests/unit/test_macro_service.py` | **ACCEPT-WITH-NOTE** | Case-insensitive substring scan preserved (`macro_service.py:1181-1182`); the `Recalculate` fixture contains none of the 23 tokens (verified at `build_full_access_fixture.py:99-116`); 3 new unit tests assert positive, case-insensitive and negative cases. Note: substring matching has a bounded false-positive surface (C-6) and now also vetoes *already-classified* macros at run time (fail-closed, intended). |
| 3 | `c28b3dd` | `pywinauto` removed from `pyproject.toml` + `uv.lock`; README + 2 overview docs rows dropped | `pyproject.toml`, `uv.lock`, `README.md`, `docs/01-overview.md`, `docs/ru/01-overview.md` | **ACCEPT-WITH-NOTE** | No `pywinauto` import anywhere (grep over `src/`, `tests/`); lock and docs are internally consistent; `six`/`comtypes`/`python-xlib` were only reachable through it. Note: `comtypes` was an *undeclared* dependency and is now absent from the environment while `vbe_compile_ui.DialogReader` still imports it lazily (A-2). |
| 4 | `119c303` | Limitations 11–14 + `snapshot_create`/`vba_export_project` side-effect note (EN + RU) | `docs/05-mcp-tools-reference.md`, `docs/ru/05-mcp-tools-reference.md` | **ACCEPT** | All four limitations verified accurate against source (D-1). EN/RU are parallel. |
| 5 | `e85c304` | Residual-token note updated (EN only) | `docs/05-mcp-tools-reference.md` | **ACCEPT-WITH-NOTE** | The EN note is accurate, but four other docs (and the RU counterpart of this very file) still claim the opposite (D-2). |
| 6 | `8be7245` | Phase-4 steps 4/6/11/14 added to the full-access e2e | `tests/integration_excel/test_mcp_full_access_workflow.py` | **ACCEPT** | Genuine assertions (explicit `table_read_rows`, per-cycle `Sheet1!B2` read-back, non-VBA `range_write_values` sentinel with `verified is True`), all on the live session. |
| 7 | `cc7abd0` | `_validation_error_as_domain()` + `call_tool` fallback; 18-tool matrix | `server.py`, `tests/integration_excel/test_mcp_error_surfacing_matrix.py` | **REJECT** | Design fits the boundary, but the translation embeds `ValidationError.errors()` in the error context and that payload is **not always JSON-serialisable**; `call_tool` then raises `TypeError` from `json.dumps`, so a reachable class of malformed arguments yields an **unstructured** error — the exact leak this change exists to remove. Reproduced (C-1). The matrix also passed its whole denial branch for the wrong reason until `e5ebcc8` (T-2). |
| 8 | `e5ebcc8` | Per-tool denial expectations; read-only denial for `table_update_rows` + `macro_run` | matrix + e2e tests | **ACCEPT-WITH-NOTE** | Materially strengthens the suite: each mutating tool must now return its **own** expected code, and ungated tools are asserted to *succeed*. It does not close the `model_validator`-shaped malformed case, so C-1 remains unguarded (T-3). |
| 9 | `db91dbd` | Phase-3 matrix (84 PASS / 4 NOT VERIFIED / 0 FAIL) | `.ai/audit/reports/xlvba-mcp-phase3-function-matrix.md` | **ACCEPT** | Row-per-tool, evidence-backed, explicit reasons for the 4 gaps; no production file touched. |
| 10 | `dece703` | Phase-4 acceptance (17/17 steps, 5/5 live tests) | `.ai/audit/reports/xlvba-mcp-phase4-e2e-acceptance.md` | **ACCEPT** | Matches the committed harness (steps 4/6/11/14 are present at HEAD); fixture hash before/after recorded. |
| 11 | `b9c13b0` | Phase-5 hidden-execution (9 PASS / 1 NOT VERIFIED) | `.ai/audit/reports/xlvba-mcp-phase5-hidden-execution.md` | **ACCEPT** | Honest `NOT VERIFIED` for focus-theft with the evidence that would be required; grep-audit of forbidden APIs is quoted. |

**Totals:** ACCEPT **3** · ACCEPT-WITH-NOTE **7** · REJECT **1**.

---

## 4. Findings

Severity: **CRITICAL / HIGH / MEDIUM / LOW**. Type: `[SPEC-DEVIATION]`, `[BEST-PRACTICE]`, `[DOC-UPDATE]`.

### 4.1 Correctness of the production changes

#### C-1 — HIGH — `[SPEC-DEVIATION]` — the FE-02 boundary translation can itself raise `TypeError`, returning an **unstructured** error (BLOCKING)

**Where:** `src/excel_vba_mcp/server.py`
- `_validation_error_as_domain()` lines 1242-1268, specifically line 1264:
  `context={"reason": "invalid_arguments", "validation": current.errors()}`
- `ExcelVbaMcpServer.call_tool()` line 1309:
  `raise ToolError(json.dumps(cause.to_dict(), ensure_ascii=False)) from cause`

**Why:** several tool argument models carry `model_validator(mode="after")` guards that `raise ValueError(...)`, e.g.
`domain/models.py:996-1013` (`RangeWriteRequest._check_bounds`), `:1027-1044` (`FormulaWriteRequest`), `:920` (`ValidationRule._check_formula_lengths`), `domain/test_spec.py:297` (`TestStep._check_declarative_size`).

For such a failure, `ValidationError.errors()` includes `ctx={"error": ValueError(...)}`. `json.dumps` cannot serialise the exception object, so the `json.dumps(...)` at `server.py:1309` raises `TypeError` **inside the `except ToolError` handler**. The MCP SDK then converts that to `CallToolResult(is_error=True, text=str(TypeError))` (`mcp/server/mcpserver/server.py:445-456`) — i.e. the client receives an unstructured error whose text is `"Object of type ValueError is not JSON serializable"`, which is *worse* than the pre-FE-02 `UnexpectedToolError("Error executing tool <name>")` because it leaks an internal serialisation detail.

**Reproduction (read-only, real production functions, no Excel):**

```
RangeWriteRequest.model_validate({'rows': 999999, 'cols': 1, 'values': [[{'kind':'number','value':1}]]})
→ ValidationError; errors() ctx: [{'error': ValueError('rows exceeds MAX_RANGE_ROWS (100000).')}]

_find_domain_cause(ToolError('Error executing tool range_write_values: …') from exc) → None
_validation_error_as_domain(same)                                                → code=INVALID_RANGE
json.dumps(cause.to_dict())                                                      → TypeError: Object of type ValueError is not JSON serializable
```

**Reachability:** any MCP call whose *argument model* trips a `model_validator` — e.g.
`range_write_values` with `rows` above `MAX_RANGE_ROWS`, `range_write_formulas` likewise,
`range_set_validation` with an over-long formula, `test_run_suite` with an oversized bounded payload. These are caller mistakes, i.e. exactly the class FE-02 was written for.

**Impact:** no data loss, no security bypass, no server crash (the SDK catches it), happy paths unaffected. But the FE-02 acceptance claim — *"every tool failure reaches the client structured (no unstructured `UnexpectedToolError` leak)"* (docstring, `server.py:1285-1287`) — is false for a reachable input class, and the observable behaviour regressed for that class.

**Recommendation (not applied):** serialise a bounded subset instead of the raw `errors()` (e.g. `current.errors(include_url=False, include_context=False, include_input=False)`), or pass `default=str` to `json.dumps`; add a matrix row using an over-limit `range_write_values` payload.

#### C-2 — LOW — `[BEST-PRACTICE]` — an *internal* `ValidationError` is reported to the caller as a schema error

`_validation_error_as_domain` walks only `__cause__` and does not distinguish *argument-schema* validation (`mcp/.../tools/base.py:151-156`, `raise ToolError(...) from exc`) from a `ValidationError` raised **inside a tool body** (wrapped as `UnexpectedToolError(...) from exc`, `base.py:211-213`). The latter is a server-side defect but is reported as `"The tool arguments do not match the input schema."` with `INVALID_RANGE`, i.e. the caller is blamed for a server bug. Low likelihood today (domain models are built from validated input), but it is a misclassification that will cost debugging time in 6 months. Suggest tagging the context (`reason="invalid_arguments"` already helps) or restricting the fallback to the direct `ToolError` cause.

#### C-3 — LOW — `[BEST-PRACTICE]` — `INVALID_RANGE` is reused for schema errors

The closed `ErrorCode` vocabulary (`domain/errors.py:26-56`) has no argument-validation code, so FE-02 reuses `INVALID_RANGE`. That keeps the vocabulary closed and matches the tool modules' malformed-argument code, and the docstring says so honestly. The cost: a client cannot distinguish "your range address is malformed" from "your tool arguments do not match the schema" without inspecting `context.reason`. Acceptable for this release; a dedicated `INVALID_ARGUMENTS` code would be the 6-month improvement (additive, non-breaking).

#### C-4 — ACCEPT — `restore_snapshot` execute-mode flip breaks no contract

Evidence (`application/snapshot_service.py:265-339`):
- `self._session.open_workbook(..., mode=SessionMode.EXECUTE, read_only=False)` (line 315-320) — the only production change; `SessionMode` was already imported and used.
- Close-without-save is unchanged (line 308-313, `contextlib.suppress(ExcelVbaMcpError)`).
- The immutable snapshot file is still not mutated: `open_workbook` copies the baseline into `work/` (`session_service.py:174-179`).
- `_latest_snapshot[info.workbook_id]` / `_known_sources[...]` (lines 325-326) and the audit record (lines 327-338, `result.restored_workbook_id`, `external_side_effects_not_undone: True`) are untouched.
- **Approval:** the restore tool itself has no approval gate (R-19) and that is unchanged. Its internal open does consult `ApprovalPolicy`, and the operation class changes: `INSPECT + read_only=False` → `MUTATING` (old) vs `EXECUTE` → `EXECUTING` (new) per `session_service.py:205-214`. In `full` mode both are allowed by policy (`approval.py:224-230`), so **no gate is weakened**; in `restricted` mode both are deny-by-default unless a matching `excel_open` grant/allowlist entry exists (`approval.py:232-264`), so **no bypass is introduced**. The only behavioural delta in restricted mode is the operation class recorded in the audit line, which is accurate (an execute open *is* an executing open).
- **Docstring nit:** "Restore has **no approval gate** (R-19)" is imprecise — the *restore tool* has no separate gate, but the internal open is approval-checked and can fail with `APPROVAL_DENIED` in restricted mode (as `docs/05` line 135 already documents). Pre-existing wording, not introduced here.

#### C-5 — ACCEPT — token scan remains a case-insensitive substring veto and does not flag the `Recalculate` fixture

- Scan: `macro_service.py:1181-1182` — `folded = source.casefold(); [t for t in INTERACTIVE_TOKENS if t.casefold() in folded]`. Unchanged by FE-04; still substring + case-insensitive.
- Tokens are appended to the existing tuple (`macro_service.py:121-145`), so all previous behaviour is preserved (additive).
- The veto is applied per **procedure body**, not per module: inventory `code` is `_procedure_body(lines, start, end) or procedure.code` (`macro_service.py:1089-1090`), so an interactive token in a sibling procedure cannot veto an unrelated safe procedure.
- `Recalculate` (`tests/fixtures/build_full_access_fixture.py:99-116`) uses only `Worksheets`, `ListObjects`, `DataBodyRange.Cells`, `Range`, `CStr` — none of the 23 tokens; the new unit test asserts `_find_interactive_tokens(source) == []` for exactly that body.
- Behavioural note: the run-time preflight (`macro_service.py:533`) also runs the scan, so a procedure that was classified `background_safe` **before** FE-04 and contains e.g. `.Activate` will now be vetoed at run time (`INTERACTIVE_UI_NOT_SUPPORTED`) even though its source hash is unchanged. This is fail-closed and matches the "veto, not certificate" doctrine; it is a deliberate tightening, not a regression.

#### C-6 — LOW — `[BEST-PRACTICE]` — substring false-positive surface of the new tokens

`.Activate`, `WindowState`, `AppActivate`, `Application.Visible` are matched as raw substrings. Realistic false positives: a user-defined `Class.ActivateThing()`, an identifier such as `MyWindowState`, or a comment mentioning `AppActivate`. The consequence is only a *veto* (the procedure cannot be classified background-safe), never unsafe execution, and the docs call the scan a heuristic. Acceptable; if maintenance noise appears, matching on word boundaries for the alphabetic tokens would remove most of it.

#### C-7 — ACCEPT — `pywinauto` is genuinely unreferenced

- `Select-String pywinauto` over `src/` → no hits; over `tests/` → only comments in `test_vbe_compile_poc.py` (historical research notes, no import).
- `uv.lock` → 0 hits for `pywinauto` / `comtypes` / `python-xlib`; `pyproject.toml` → no `pywinauto`.
- README + `docs/01-overview.md` + `docs/ru/01-overview.md` rows removed; `Select-String pywinauto` over `docs/` → no hits.
- Residual references are confined to `.ai/**` historical plans/audit artifacts, which are evidence records and correctly left untouched.

### 4.2 Architectural fit

#### A-1 — layering respected; no unjustified abstraction
- `snapshot_service` change is a single argument flip inside the application service; it does not leak `SessionMode` into the API layer and does not add a branch.
- `macro_service` change is a data-only tuple extension in the same module that already owns the heuristic; no new module, no new abstraction.
- `server.py` change is confined to the transport/composition boundary (`ExcelVbaMcpServer.call_tool`), which is exactly where SDK-exception translation belongs; it does not reach into services.
- SRP is preserved: `_find_domain_cause` (domain) and `_validation_error_as_domain` (schema) are separate, single-purpose helpers; the second is a fallback, not a replacement (`_find_domain_cause` still runs first, so domain errors are never weakened or rewritten).
- Non-domain, non-`ValidationError` exceptions are still re-raised unchanged (`server.py:1307-1308` → bare `raise`), so the SDK's crash logging is preserved. This was verified by reading the SDK wrapper (`base.py:151-213`, `server.py:445-456`) and matches the docstring's claim.

#### A-2 — LOW — `[BEST-PRACTICE]` — `comtypes` is an undeclared dependency that FE-07 removed from the environment
`src/excel_vba_mcp/adapters/windows_ui/vbe_compile_ui.py:345-349` still does
`import comtypes / import comtypes.client` inside `DialogReader.read_dialog_text`, wrapped in `except Exception: return None`.
Before FE-07, `comtypes` was installed transitively by `pywinauto`; after FE-07 it is absent from `uv.lock` **and** from the venv (`importlib.metadata` → `PackageNotFoundError`). The reader now always returns `None`, i.e. VBE compile-dialog text capture is permanently unavailable in a clean install.
Not a functional regression for the release: `vba_compile` is fail-closed `indeterminate` on this build and never opens the dialog (`application/compile_service.py`, `NON_ACTIVATING_COMPILE_UNAVAILABLE`), and Phase-5 recorded the path as fail-closed. But the dependency is undeclared either way, and the failure is silent. Recommendation: either declare `comtypes` explicitly (it is a real runtime import) or state in `docs/06`/`docs/05` that dialog-text capture is unavailable without it.

### 4.3 Test quality

#### T-1 — FE-03 coverage is genuine (ACCEPT)
`tests/integration_excel/test_mcp_full_access_workflow.py`:
- step 4: `table_read_rows("tblSettings")` before any mutation, asserting `columns == ["SettingKey","SettingValue"]`, the `factor` row and the `label` row, plus `Result!A1 == 0` (lines ~548-566).
- steps 6/11: `range_read_values("Sheet1","B2") == 10` / `== 20` after **each** `table_update_rows`, on the live session (lines ~578-590).
- step 14: a non-VBA, non-table mutation (`range_write_values` sentinel into `Result!A3`) with `verified is True` **and** a read-back (lines ~599-610).
These are behavioural assertions, not tautologies, and they exercise a second mutation surface.

#### T-2 — FE-02 matrix now asserts real per-tool denial codes (ACCEPT-WITH-NOTE)
`tests/integration_excel/test_mcp_error_surfacing_matrix.py:422-462` maps each of the 18 non-read-only tools to its **expected** code and asserts `payload["code"] == expected` (`WORKBOOK_READ_ONLY` for the 11 writable-guarded tools, `SNAPSHOT_FAILED` for restore/diff, and explicit `None` for the ungated/read-only-safe tools which must *succeed*). This is a real assertion, not "some code".
Process note: the first version (`cc7abd0`) passed the whole `_denial_args(...)` mapping as the tool arguments instead of `_denial_args(...)[tool]`, so every "denial" was actually a schema rejection (`INVALID_RANGE`); `e5ebcc8` fixed the indexing and replaced "some code" with per-tool codes. The fix landed inside the same orchestration, before this gate — the committed HEAD state is the strengthened one, and the commit message says so plainly.

#### T-3 — LOW — the malformed branch no longer proves per-tool validation, and misses the C-1 shape
`_malformed_args` uses only *missing required field* and *wrong-typed field* cases (e.g. `{"workbook_id": wid}`), which produce `missing` / `int_parsing` errors. Those serialise cleanly, so the branch now passes for every tool **whether or not the tool validates its own arguments**, because the new boundary translation supplies `INVALID_RANGE` anyway. It also means the one input shape that breaks the new code (`model_validator` → `ctx.error` exception object) is not covered anywhere — which is why C-1 escaped both the matrix and the 870 unit tests. A single over-limit `range_write_values` row would have caught it.

#### T-4 — nothing was weakened to pass (ACCEPT)
`git diff 8c22f11..HEAD -- tests/unit tests/integration_excel` with `^-` filtered shows exactly **one** removed line: the old one-line docstring of `test_read_only_session_denies_mutation`, replaced by a longer docstring plus **added** assertions for `table_update_rows` and `macro_run`. No assertion was deleted or relaxed anywhere in the range.

#### T-5 — originals-unchanged guarantees intact (ACCEPT)
`git diff 8c22f11..HEAD -- tests/fixtures/` is empty. Both live test files also assert byte-identity internally (repo fixture hash before/after, temp-copy hashes), and the matrix closes both sessions in a `finally`-style tail.

### 4.4 Documentation accuracy

#### D-1 — limitations 11–14 and the fs-side-effect note match the code (ACCEPT)

| Doc item | Claim | Verified against |
|---|---|---|
| 11 | stale worker / dead owned-Excel not auto-detected by `excel_health`; no in-band PID probe; remedy close/reopen | `adapters/excel/com_worker.py:411-415` — `get_status()` is pure in-process bookkeeping, no `OpenProcess`/`GetExitCodeProcess`; no `excel_recover_session` tool registered |
| 12 | `macro_run` has no snapshot precondition (spec deviation) | `application/macro_service.py` — no `SnapshotService` dependency, no `require_snapshot` call anywhere in the run path |
| 13 | recovery is close/reopen only; `SessionService.recover_session`/`close_session` exist but are not MCP tools | `session_service.py:546` defines `recover_session`; `server.py::_register_tools` registers no such tool (grep `excel_recover_session` → 0 hits) |
| 14 | `macro_policy.classify_operation` is non-authoritative / unwired | `security/macro_policy.py:178` defines it; the only other reference is the `staticmethod` alias at `:312`; no call site in `application/` or `api/` |
| fs side-effects | `snapshot_create` / `vba_export_project` write files by design, ungated in `full` **and** `restricted` (R-19) | consistent with the read-only-tool classification and with the absence of an approval gate on those paths |

#### D-2 — MEDIUM — `[DOC-UPDATE]` — the residual-token note was updated in only one of six places
FE-04 added the four tokens to the veto, and `e85c304` updated the EN residual note (`docs/05-mcp-tools-reference.md:589-594`). The following still assert the opposite and are now **factually wrong**:
- `docs/05-mcp-tools-reference.md` (RU) line 594 — states that `AppActivate` and `Application.Visible = True` are **not** in the veto list (quoted in Russian: *"токены `AppActivate` и `Application.Visible = True` не входят в veto-список"*), i.e. the RU twin of the very file that was updated.
- `docs/07-workflows-and-examples.md:463-464` — *"`AppActivate` and `Application.Visible = True` are not in the veto list."*
- `docs/06-excel-capabilities.md:112` and `:282` — *"… `AppActivate`, `Application.Visible = True` not fully covered"*.
- `docs/04-architecture.md:427` and `docs/ru/04-architecture.md:433` — *"tokens such as `AppActivate` or `Application.Visible = True` are **not** fully covered"*.

(For `04`/`06` the sentence is partly defensible — *indirect* UI is still uncovered — but naming these two tokens as uncovered is now false and will mislead the next maintainer into "re-adding" them.) This is a doc-accuracy defect, not a runtime one; it should be a one-line follow-up in all six places.

#### D-3 — the 88-tool count is consistent (ACCEPT)
- `server.py`: 87 `server.add_tool(...)` calls + 1 `@mcp.tool()` (`server_health`) = **88**.
- `docs/01-overview.md:39`, `docs/04-architecture.md:60`, `docs/05-mcp-tools-reference.md:21,110,112` (and RU twins) all say **88**, and the group sum `6+4+3+9+2+4+7+11+13+16+5+3+3+1+1 = 88` is arithmetically correct.
- No stale `87` in `docs/` or `README.md` (grep `\b87\b` on docs → no hits).

#### D-4 — no stale `pywinauto` in shipped docs (ACCEPT)
`Select-String pywinauto` over `docs/` and `README.md` → no hits. Remaining mentions are in `.ai/**` evidence/plan artifacts and in `tests/integration_excel/test_vbe_compile_poc.py` comments (research provenance, no import).

### 4.5 Regressions, unrelated changes, working-tree hygiene

#### R-1 — quality gates green (ACCEPT)
`870 passed`; `ruff check src tests` → *All checks passed!*; `mypy` on the 3 changed production files → *Success: no issues found in 3 source files*. This matches the Tech Lead's recorded results. (Full-package mypy has 2 pre-existing `comtypes` stub errors in the untouched `adapters/windows_ui/vbe_compile_ui.py`; not introduced here.)

#### R-2 — change surface is exactly the orchestration (ACCEPT)
`git diff --name-only 8c22f11..HEAD` = 17 paths: 3 production, 4 test, 2 docs EN + 2 docs RU + README + `pyproject.toml` + `uv.lock` + 3 audit reports. No unrelated file, no fixture, no CI/config change.

#### R-3 — LOW — the commit series was built on a dirty working tree
At audit time `git status` shows 5 **uncommitted deletions** that are not part of any reviewed commit:
`docs/03-Operations/release-notes.md`, `docs/03-Operations/stage1-acceptance-matrix.md`, `docs/release-evidence-phase2-ranges-pivots.md`, `docs/research/pivot_tables_api.md`, `docs/research/ranges_cells_api.md` (plus untracked `.ai/plans/*`, `.ai/audit/problems/`, `.ai/tasks/todo/`). They exist in `8c22f11` and `HEAD` (`git ls-tree`), so nothing was lost by this work, and the orchestration plan itself records the deletions as pre-existing. Still, the release gate should not be closed on a dirty tree: either commit the deletions deliberately or restore them, so `git status` is clean at the release tag. (This audit does not touch them.)

---

## 5. Residual risks & known limitations

**Documented and verified (no action needed for this gate):**
1. **Focus-theft is NOT VERIFIED** (Phase-5): point-sampling (`EnumWindows` + `IsWindowVisible` per owned PID) plus a grep audit showing zero `SetForegroundWindow`/`ShowWindow`/`SendKeys` call sites. Continuous foreground observation would be required to verify it; owned Excel was never visible.
2. **Shape / control tools are MCP-unverified** (Phase-3): `shape_get`, `shape_set_on_action`, `control_get`, `control_set_on_action` were not invoked over MCP because the fixture has no shape/control surface; they are covered at the service layer (`tests/unit/test_shapes_controls.py`). The Phase-3 report itself recommends either extending `build_full_access_fixture.py` or recording the gap — it is recorded here.
3. **`macro_run` has no snapshot precondition** (spec deviation, documented as limitation 12): a classified, approved macro can run with no prior `snapshot_create`. The recommended workflow still snapshots first.
4. **Stale-worker health** (limitation 11): `excel_health` trusts in-process bookkeeping; a killed `EXCEL.EXE` can still show `ok`/`running` until a call fails. Remedy: `excel_close` + reopen, or restart the server.
5. **`vba_compile` is indeterminate** on this build (`NON_ACTIVATING_COMPILE_UNAVAILABLE`, `is_compile=false`); it never returns `compiled`/`compile_error` and never shows UI.
6. **Restore now yields an execute-mode session** (FE-16): intended (D-1), but it means a snapshot whose macros were never reviewed is opened with macros enabled after restore. `full` mode is the default, so this is a deliberate capability/security trade-off, not a leak; `restricted` mode still denies the open without a grant.
7. **R13 token scan is a heuristic veto**, never a certificate (now 23 tokens; indirect/transitive UI still evadable).
8. **`comtypes` is absent** from a clean install (A-2), so VBE dialog-text capture is silently unavailable; compile remains fail-closed.

**New / unresolved from this audit:**
9. **C-1 (HIGH, blocking):** the FE-02 boundary translation raises `TypeError` on `model_validator`-shaped validation errors, so those calls still reach the client unstructured.
10. **C-2 / C-3 (LOW):** internal `ValidationError`s are reported as caller schema errors, and `INVALID_RANGE` is reused for schema failures.
11. **D-2 (MEDIUM, doc):** six stale statements about the veto list (EN + RU).
12. **R-3 (LOW):** dirty working tree at the release gate.

---

## 6. Verdict

The orchestration is, with one exception, in good shape: the three production changes are minimal, layered correctly and free of contract breaks; the fixtures are untouched; the test suite grew by real behavioural assertions rather than tautologies; `870 / ruff / mypy` are green; and the doc set is consistent on the 88-tool count, the dependency list and limitations 11–14. Seven of eleven commits are clean accepts and seven more are accepts with documented notes.

The exception is FE-02. `_validation_error_as_domain` embeds the raw `ValidationError.errors()` in the error context, and for any argument model with a `model_validator` guard (`RangeWriteRequest`, `FormulaWriteRequest`, `ValidationRule`, `TestStep`) that payload contains a live exception object in `ctx`. `ExcelVbaMcpServer.call_tool` then calls `json.dumps(cause.to_dict())` on it and raises `TypeError` inside its own `except` handler, so the client receives an unstructured error whose text is an internal serialisation message. That is the exact failure mode the commit exists to remove, it is reachable with a single malformed-but-well-typed call (e.g. `range_write_values` with `rows` above `MAX_RANGE_ROWS`), and the FE-02 matrix does not cover it because its malformed cases only produce `missing` / type errors. It is narrow in blast radius (no data risk, no security impact, no server crash) and cheap to fix (serialise a bounded subset, or `default=str`, plus one matrix row) — but the release claim "every tool failure is structured" is not true at HEAD.

CODE AUDIT: FAIL

**Blocking reason (1):**
- **C-1 (HIGH)** — `src/excel_vba_mcp/server.py`, `_validation_error_as_domain` (line 1264) + `ExcelVbaMcpServer.call_tool` (line 1309): a reachable `ValidationError` class (`model_validator` bounds guards in `domain/models.py:996`, `:1027`, `:920` and `domain/test_spec.py:297`) produces a non-JSON-serialisable context, so the boundary itself raises `TypeError` and the client receives an **unstructured** error — the FE-02 contract ("no unstructured tool failure") is not met.

**Non-blocking follow-ups:** C-2, C-3, A-2, T-3, D-2 (six doc statements), R-3 (clean the working tree before tagging).
---

## 6. Post-audit remediation (Tech Lead, 2026-11-22)

The blocking finding **C-1** is **CLOSED** by commit `b08d638`:

- `server.py` adds `_jsonable_validation_errors()`, which stringifies every non-primitive value in
  `ValidationError.errors()` (both the top-level `input` and `ctx['error']`), so the error `context` is always
  JSON-serializable. The over-limit `range_write_values` path now returns a structured `INVALID_RANGE` instead of
  a `TypeError` leak.
- Regression coverage added: `tests/unit/test_mcp_error_surfacing.py` (2 tests) and a live "model-validator" case
  in `tests/integration_excel/test_mcp_error_surfacing_matrix.py` (this also closes **T-3**).
- The same commit fixes the test-isolation defect the audit observed: the matrix harness now reaps its owned hidden
  Excel before exit, so both integration files pass in either order with zero orphans.

Non-blocking findings dispositioned:

- **A-2 → fixed** (`1633115`): `comtypes>=1.2.0` declared in `pyproject.toml`, `uv.lock` refreshed (comtypes 1.5.0),
  dependency tables updated.
- **D-2 → fixed** (`1633115`): the EN+RU `04-architecture`, `05-mcp-tools-reference`, `06-excel-capabilities`,
  `07-workflows-and-examples` and README residual-token statements now correctly reflect that
  `AppActivate` / `Application.Visible` / `.Activate` / `WindowState` are in the veto list.
- **C-2 / C-3 / C-6 → accepted** as documented low-severity, best-practice residuals.
- **R-3 → noted** (pre-existing doc deletions, never staged).

Updated verdict: the blocking reason is resolved. Re-running the quality gates at HEAD gives **872 unit passed**,
**5 integration passed** (both orders, zero orphaned Excel), ruff clean, mypy clean on the changed modules, and
fixtures byte-identical. The audit's substantive findings are all closed or accepted; see the capability report
§10 (Tech Lead addendum) for the consolidated acceptance.
