# XLVBA MCP — Full Edit Mode Final Capability Report (FE-15 / Phase 6)

**Task:** `FE-15` / `FULL-EDIT-MODE-ORCHESTRATION-001` (Phase 6 — final product-acceptance gate)
**Role:** Validator (independent functional verification)
**Date:** 2026-11-22
**Repo / branch:** `C:\py_exp\xlvba_mcp` @ `master`, HEAD `b9c13b0` ("audit: Phase-5 hidden-execution & desktop non-interference (FE-14)")
**Verdict:** **`PARTIALLY VERIFIED`** (see §1, §9)

**Read-only discipline honoured.** No production code, test, or doc was modified by this validation. The only
artifact created by this gate is this report (plus throwaway probe scripts under `.ai/tmp/`, which are git-ignored
and **not** committed). No Trust Center / VBOM / registry / macro-security changes were made. No macro was run
against an original fixture; every live call targeted a disposable copy in an OS temp root.

---

## 1. Executive verdict

**`PARTIALLY VERIFIED`.**

The **core Full-Edit loop is fully and independently proven through the real MCP boundary**: a default
writable/execute `excel_open`, real table mutations, two consecutive edit→macro cycles in **one session** with
**different observed outputs** (`Result!A1` 20 then 40), MCP background-safe classification, interactive-macro
veto, unclassified-macro denial, read-only denial, restricted-mode `APPROVAL_DENIED`, and snapshot/save/restore
with a post-restore **execute-mode** session whose macros run. Excel ran hidden with **zero** forbidden
focus/input API call sites, and the source fixture stayed byte-identical.

Two residuals keep the verdict at `PARTIALLY VERIFIED` rather than `FULL EDIT MODE VERIFIED`:

1. **4 registered MCP tools are NOT VERIFIED through MCP** — `shape_get`, `shape_set_on_action`, `control_get`,
   `control_set_on_action`. No committed fixture carries a drawing shape or a Form/ActiveX control, so these four
   could not be exercised via `call_tool`. They are **service/adapter-layer covered** by
   `tests/integration_excel/test_shapes_controls.py` and `tests/unit/test_shapes_controls.py`, and they are
   **peripheral** to the Full-Edit mission (they are not part of the open→edit→macro→save workflow).
2. **Focus-theft is NOT VERIFIED** (Phase-5 item 10) — only discrete point-samples were taken, so a transient
   focus grab between samples cannot be strictly excluded. This is the exact condition the spec (§5 Phase 5)
   instructs to mark `NOT VERIFIED` and document with the evidence needed, not to assert as never-stolen.

Per §11, `FULL EDIT MODE VERIFIED` requires *all critical required capabilities work* with only *non-critical*
limitations documented. Because four **registered tools** still lack MCP-observed evidence and one Phase-5
**condition** is formally unverified, the honest, evidence-bounded verdict is `PARTIALLY VERIFIED`. Neither
residual blocks the Full-Edit mission (see §9 for the blocking analysis); they are formal coverage gaps against a
"every registered tool MCP-verified + every condition observed" bar.

---

## 2. Environment (verified live)

| Item | Value |
|---|---|
| OS | Windows 10 Corporate LTSC (Enterprise LTSC 2019), NT 10.0.17763.0 (SP0) |
| Host session | **Interactive console, Session 1** (`qwinsta` → `console … 1 Active`; `Get-Process -Id $PID` → SessionId 1). Not Session 0 / not a service. |
| Python | 3.14.7 (`.venv\Scripts\python.exe`) |
| Excel | COM `Application.Version = 16.0`; owned `DispatchEx` probe reported `Visible = False` |
| MCP SDK | `mcp` 2.3.0 (`stdio_client` + `ClientSession`), `pydantic` 2.14.4 |
| Server entry point | `python -m excel_vba_mcp` (STDIO) |
| Access mode | `EXCEL_VBA_MCP_ACCESS_MODE` **unset** → default `full` |
| Test root | Per-run `tempfile.mkdtemp(...)`; repo fixtures never opened directly |
| VBOM trust | Available (classify + `macro_run` succeeded) |

---

## 3. Methodology & evidence sources

1. **Decisive committed tests, re-run serially** (`-p no:xdist`, `-o addopts=""`):
   `tests/integration_excel/test_mcp_full_access_workflow.py` + `tests/integration_excel/test_mcp_error_surfacing_matrix.py`.
2. **Independent live MCP probe** (this gate): `.ai/tmp/fe15_live_probe.py` — real `python -m excel_vba_mcp`
   STDIO server + `mcp.ClientSession`, disposable temp copy of `tests/fixtures/FullAccessWorkflow.xlsm`.
3. **Independent fresh-server unclassified probe** (this gate): `.ai/tmp/fe15_unclassified_probe.py` — correct
   methodology (macro run *before* any classify, empty registry).
4. **Committed artifacts re-confirmed:** baseline census report, Phase-3 function matrix, Phase-4 e2e acceptance,
   Phase-5 hidden-execution report, orchestration plan (D-1..D-7 + ledger).
5. **Code-path spot checks** (read-only): `adapters/excel/app_adapter.py` (`create_hidden`, `_owned_window_is_foreground`,
   `_pid_from_hwnd`), `application/macro_service.py` (`BackgroundSafetyRegistry` keying, `INTERACTIVE_TOKENS`),
   `application/snapshot_service.py` (`restore_snapshot` → `SessionMode.EXECUTE`), `config.py` (default `full`).
6. **Fixture integrity:** SHA-256 before/after every run.

Prior reports were treated as **claims to re-confirm**, not proof. Every binding requirement below is backed by
this gate's own re-run output or live observation, not by an Implementor's or prior Auditor's assertion.

---

## 4. Binding-requirement matrix (§3 constraints + §9 acceptance criteria)

Legend: **SATISFIED** = objectively observed through the real MCP boundary (or, for code-structure claims,
verified in source with zero counter-evidence). **NOT VERIFIED** = not observable / not covered, with reason.

| # | Binding requirement | Status | Evidence (test / observed value / code path) |
|---|---|---|---|
| R1 | **Full access is default** when `EXCEL_VBA_MCP_ACCESS_MODE` unset; default `excel_open` is writable + execute | **SATISFIED** | Baseline: `config.py::ServerConfig.access_mode` default `AccessMode.FULL`; `load_config()` only sets the key when the env var is present. This gate's live probe (env unset): `excel_get_status.access_mode="full"`; `excel_open {path}` → `{mode:"execute", read_only:false}`. Test `test_full_access_workflow_e2e` asserts `opened["read_only"] is False` and `opened["mode"] == "execute"`. |
| R2 | **Table mutation through MCP**; two consecutive edit→macro cycles in one session with **different** outputs | **SATISFIED** | Probe + `test_full_access_workflow_e2e`: `snapshot_create` → `table_update_rows(tblSettings, row1, SettingValue=10)` → read `Sheet1!B2 == 10` → `macro_run Recalculate` → `Result!A1 == 20`; then `table_update_rows(SettingValue=20)` → `Sheet1!B2 == 20` → `macro_run` → `Result!A1 == 40` and `40 != 20` (proves the macro read the **live** table state, same `workbook_id`, no save between). Call path: `api/workbook_tools.py::excel_open` → `application/session_service.py::open_workbook` → `ExcelComWorker`; `table_update_rows` → `TableService`; `macro_run` → `MacroService.run_macro` → `MacroRunner` (STA `Application.Run`). |
| R3 | **Background-safe classification through MCP**; interactive macro veto; unclassified blocked | **SATISFIED** | `macro_classify_background(modWorkflow.Recalculate, true)` → `background_safe=true`; `macro_classify_background(modWorkflow.ShowMessage, true)` → **`INTERACTIVE_UI_NOT_SUPPORTED`** (R13 veto); **fresh-server** `macro_run Recalculate` before any classify → **`status=blocked, error_code=BACKGROUND_SAFETY_NOT_VERIFIED`**. Code: `MacroService.classify_background` + `BackgroundSafetyRegistry` keyed by `(workbook_file_sha256, project_source_hash, procedure.lower())`, fail-closed (empty → `None` → `BACKGROUND_SAFETY_NOT_VERIFIED`); `INTERACTIVE_TOKENS` scan vetoes `MsgBox` etc. |
| R4 | **Read-only denies mutations**; **restricted mode preserved** (`APPROVAL_DENIED`); **full mode needs no grants** | **SATISFIED** | `excel_open(read_only:true)` then `range_write_values` → **`WORKBOOK_READ_ONLY`** (probe + `test_read_only_session_denies_mutation`: also `table_update_rows` and `macro_run` → `WORKBOOK_READ_ONLY`). `EXCEL_VBA_MCP_ACCESS_MODE=restricted` → `excel_open {path}` → **`APPROVAL_DENIED`** ("denied by default: no matching allowlist entry or live grant") (`test_restricted_mode_denies_execute_open`). Full-mode run used **no** grant/allowlist tool — none exists; `ApprovalPolicy` full-mode branch auto-allows. |
| R5 | **Snapshot / save / restore** with post-restore **execute-mode** session and **runnable macros** | **SATISFIED** | `test_persistence_roundtrip`: `snapshot_create` → edit → `macro_run` (A1=40) → `excel_close(save_changes:true)` → `snapshot_restore` returns a **fresh** `workbook_id` → `excel_get_status` → `mode=="execute"`, `read_only is False` → classify + `macro_run` **succeeds** post-restore (`Result!A1 == 40`, `Sheet1!B2 == 20`, `modWorkflow`/names persisted). Code: `snapshot_service.py::restore_snapshot` opens with `mode=SessionMode.EXECUTE` (FE-16 fix, commit `0993370`). |
| R6 | **Excel hidden** in the interactive session; **no focus/input APIs**; originals **byte-identical** | **SATISFIED** (with focus-theft sub-item NOT VERIFIED) | Hidden: `app_adapter.create_hidden` uses `DispatchEx` + sets `Visible=False` **before** anything else and **quits** if it cannot hide (non-intrusion invariant); owned `excel_health.owned_pid` is a dedicated PID; Phase-5 harness asserts `EnumWindows`+`IsWindowVisible` finds **no visible window** for the owned PID at start/open/cycles. No focus/input APIs: grep of `src/**/*.py` for `SetForegroundWindow|ShowWindow|SendKeys|AppActivate|SetActiveWindow|BringWindowToTop|keybd_event|mouse_event|SetCursorPos|SendInput|BlockInput|PostMessage|SendMessage|SetFocus` → **11 hits, all docstrings or denylist string literals, zero call sites**; only read-only Win32 use is `GetForegroundWindow` + `GetWindowThreadProcessId` in `app_adapter`. Interactive host confirmed (Session 1 console). Originals: fixture SHA-256 `6B905837…A19AC1D4D9` unchanged before/after (see §7). **Focus-theft continuous observation → NOT VERIFIED** (see §8). |

**Summary:** 6 / 6 binding requirements SATISFIED at the capability level; within R6, the discrete
focus-theft *condition* is NOT VERIFIED (continuous trace unavailable). Separately, 4 of the 88 registered
MCP tools are NOT VERIFIED through MCP (§5.3) — they are not part of R1–R6 but are part of the §9 "every live
MCP tool has a final evidence-backed status" criterion.

---

## 5. Decisive test + live probe evidence

### 5.1 Decisive committed tests — re-run result: **5 passed in 71.79s**

Command (serial, from repo root):
```
.venv\Scripts\python.exe -m pytest tests/integration_excel/test_mcp_full_access_workflow.py \
  tests/integration_excel/test_mcp_error_surfacing_matrix.py -o addopts="" -q -p no:xdist -s
```
Result: **`5 passed in 71.79s`** (0 failed, 0 skipped). Test functions:
`test_full_access_workflow_e2e`, `test_persistence_roundtrip`, `test_read_only_session_denies_mutation`,
`test_restricted_mode_denies_execute_open`, `test_error_surfacing_matrix`.

Concrete evidence captured from the run (stderr structured payloads):
- `macro_classify_background(ShowMessage,true)` → `{"code":"INTERACTIVE_UI_NOT_SUPPORTED", …}`.
- Read-only session: `range_write_values` / `table_update_rows` / `macro_run` → `{"code":"WORKBOOK_READ_ONLY", …}`.
- Restricted mode `excel_open {path}` → `{"code":"APPROVAL_DENIED", … "reason":"denied by default: no matching allowlist entry or live grant"}`.
- `snapshot_restore` bad id → `SNAPSHOT_FAILED`; malformed args across the matrix → `INVALID_RANGE` (structured; **no** unstructured `UnexpectedToolError` / "Error executing tool" leak).
- `excel_close` on an already-closed id → `COM_BUSY` ("Unknown workbook_id … (not open)") — structured, not a leak.
- cycle1 `Result!A1 == 20`, cycle2 `Result!A1 == 40` (and `!= 20`); post-restore `mode == "execute"`, `read_only is False`, post-restore `macro_run` succeeds with `Result!A1 == 40`, `Sheet1!B2 == 20`.
- FE-02 matrix per-tool: every mutating tool denial → expected code (`WORKBOOK_READ_ONLY` / `SNAPSHOT_FAILED`), every malformed call → `INVALID_RANGE`; ungated tools (`snapshot_create`, `vba_export_project`, `excel_save`, `excel_close`) correctly succeed (no denial expected); no tool leaked a non-structured error.

### 5.2 Independent live MCP probe (this gate) — 16 / 17 PASS (1 = probe-method artifact, see note)

`.ai/tmp/fe15_live_probe.py` (real STDIO server + `ClientSession`, temp copy, env unset). Observed JSON:
```json
{
  "tool_count": 88,
  "has_macro_classify_background": true,
  "excel_health": {"status": "ok", "worker_state": "running", "accepting": true, "owned_pid": 13588},
  "default_open": {"mode": "execute", "read_only": false},
  "snapshot_create": true,
  "classify_recalculate": true,
  "cycle1": {"sheet1_b2": 10.0, "macro_status": "success", "result_a1": 20.0},
  "cycle2": {"sheet1_b2": 20.0, "macro_status": "success", "result_a1": 40.0},
  "showmessage_veto": "INTERACTIVE_UI_NOT_SUPPORTED",
  "unclassified_macro": {"status": "success", "error_code": null},
  "readonly_open": {"mode": "execute", "read_only": true},
  "readonly_denial": "WORKBOOK_READ_ONLY",
  "originals_byte_identical": {"repo_unchanged": true, "temp_copy_unchanged": true, "ro_copy_unchanged": true}
}
```
PASS: tool_count≥87 (88), has `macro_classify_background`, `excel_health` ok, default open writable execute,
snapshot created, classify true, cycle1 B2=10 / A1=20, cycle2 B2=20 / A1=40, cycles differ,
`ShowMessage` → `INTERACTIVE_UI_NOT_SUPPORTED`, read-only open `read_only=true`, read-only denial
`WORKBOOK_READ_ONLY`, repo fixture unchanged, temp copy unchanged.

**Note on the one "FAIL" in the probe (NOT a product defect):** the probe classified `Recalculate`
background-safe early, then re-opened the **same source file** and expected the macro to be "unclassified". The
`BackgroundSafetyRegistry` is an in-memory, server-lifetime store keyed by
`(workbook_file_sha256, project_source_hash, procedure.lower())` (`application/macro_service.py`), so a second
open of **identical content** correctly **inherits** the content-bound classification — by design (a workbook hash
or VBA-source change invalidates it; cell-only edits do not). The probe's expectation was wrong, not the server.

### 5.3 Clean fresh-server unclassified probe (correct methodology) — PASS

`.ai/tmp/fe15_unclassified_probe.py` starts a **fresh** server (empty registry) and runs `Recalculate` **before**
any classify:
```
{"unclassified_status": "blocked", "unclassified_error_code": "BACKGROUND_SAFETY_NOT_VERIFIED"}
[PASS] fresh-server unclassified macro -> BACKGROUND_SAFETY_NOT_VERIFIED
```
This is the correct observation of the deny-by-default path and matches `test_full_access_workflow_e2e`'s
unclassified assertion.

---

## 6. MCP surface inventory (re-confirmed)

| Surface | Count | Evidence |
|---|---|---|
| Tools | **88** | `len((await session.list_tools()).tools) == 88` (this gate's live probe + baseline). Code: 87 `server.add_tool(...)` + `server_health` via `@mcp.tool()`. |
| Static resources | **0** | `list_resources()` → empty. |
| Resource **templates** | **4** | `excel://workbooks/{workbook_id}/model`, `…/vba/components`, `…/vba/components/{component_name}`, `excel://test-runs/{run_id}` (`server.py::*_resource` → `WorkbookModelService` / `VBAProjectService` / `TestRunner`). |
| Prompts | **0** | `list_prompts()` → empty. |

Per-tool status (Phase-3 matrix, re-confirmed): **84 / 88 PASS**, **0 FAIL**, **4 NOT VERIFIED** —
`shape_get`, `shape_set_on_action`, `control_get`, `control_set_on_action` (fixtures carry no shape / no control;
`shapes_list`/`controls_list` returned `[]`). These four are **service/adapter-layer covered** by
`tests/integration_excel/test_shapes_controls.py` and `tests/unit/test_shapes_controls.py`; the underlying
adapters are shared with the (verified) sheets/ranges paths, so product risk is low, but there is **no
MCP-observed evidence** for the four tool handlers. This is a **fixture-coverage gap, not a product defect.**

Error-semantics distinctions re-confirmed as structured codes (never unstructured leaks):
`APPROVAL_DENIED` (restricted) · `WORKBOOK_READ_ONLY` (read-only) · `BACKGROUND_SAFETY_NOT_VERIFIED` /
`INTERACTIVE_UI_NOT_SUPPORTED` (macro) · `SNAPSHOT_FAILED` (missing snapshot) · `INVALID_RANGE` (malformed args) ·
`COM_BUSY` (unknown id / contention).

---

## 7. Original fixture integrity

| File | SHA-256 | Result |
|---|---|---|
| `tests/fixtures/FullAccessWorkflow.xlsm` | `6b905837b8deeb549505f9a8a51c4f55507d377e4160bb5d9d6363a19ac1d4d9` | **UNCHANGED** before and after both test runs and both live probes; matches the expected `6B905837…A19AC1D4D9`. |

`git status` after all work: **no production/test/doc changes attributable to this validation.** The only working-tree
entries are pre-existing (deleted `docs/03-Operations/*` and `docs/research/*` from prior tasks, and untracked
`.ai/audit/problems/`, `.ai/plans/*`, `.ai/audit/reports/xlvba-mcp-full-edit-mode-baseline.md`,
`.ai/tasks/todo/`). Throwaway probes live under `.ai/tmp/` (git-ignored) and are not committed.

---

## 8. Hidden execution & desktop non-interference (Phase-5 re-confirmation)

| Check | Status | Evidence |
|---|---|---|
| Excel launched hidden, owned instance | **PASS** | `create_hidden` `DispatchEx` + `Visible=False` before any work; quits if it cannot hide. Owned `excel_health.owned_pid` distinct from pre-existing user PIDs. |
| No VM / container / separate session / service / SYSTEM / remote | **PASS** | Interactive console Session 1; server is a local `python -m excel_vba_mcp` stdio child. |
| Existing user Excel instances not reused/controlled/closed | **PASS** | `DispatchEx` (not ROT reuse); Phase-5 harness asserts user PID set survives and owned PID is reaped. |
| No owned visible window (per-PID `EnumWindows`+`IsWindowVisible`) | **PASS** | Phase-5 harness `_assert_owned_hidden` at start/open/cycles; owned visible-window set empty. |
| No `SetForegroundWindow`/`ShowWindow`/`SendKeys`/foreground clicks in code path | **PASS** | Grep of `src/**/*.py`: **zero call sites**; 11 hits are docstrings/denylist literals only. |
| Interactive VBA rejected without UI | **PASS** | `ShowMessage` classify → `INTERACTIVE_UI_NOT_SUPPORTED`; no dialog (owned instance `Visible=False`). |
| No foreground UI automation to operate VBE | **PASS** | `adapters/windows_ui/vbe_compile_ui.py` uses comtypes `IUIAutomation` **read-only**; compile path fails closed (`COMPILE_INDETERMINATE`); `pywinauto` removed (commit `c28b3dd`). |
| Worker health correct after repeated mutation + macro | **PASS** | `excel_health` ok pre/post; owned PID live; `accepting=true`, `needs_recovery=false`. |
| Owned PID reaped after shutdown; user PIDs alive | **PASS** | Phase-5 harness `_wait_pid_dead(owned_pid)` + `surviving_user_pids` empty. |
| **Focus-theft: continuous observation** | **NOT VERIFIED** | Only discrete point-samples (foreground unchanged; owned Excel never foreground/visible at sampled points). A transient grab between samples cannot be excluded without a continuous `EVENT_SYSTEM_FOREGROUND` hook or high-frequency `GetForegroundWindow` logger. Structural risk is very low (zero forbidden APIs; `Visible=False` enforced at creation). |

**Evidence needed to close focus-theft:** a continuous foreground-window logger (background thread polling
`GetForegroundWindow`+`GetWindowThreadProcessId` every 10–50 ms, or an out-of-context
`SetWinEventHook(EVENT_SYSTEM_FOREGROUND)`) running for the entire server lifecycle, asserting no record names
the owned Excel PID/hwnd.

---

## 9. What remains unverified — and whether it is blocking

| Residual | Why unverified | Blocking? |
|---|---|---|
| `shape_get`, `shape_set_on_action`, `control_get`, `control_set_on_action` NOT VERIFIED through MCP | No committed fixture carries a drawing shape or a Form/ActiveX control; `shapes_list`/`controls_list` return `[]`. | **Non-blocking** for the Full-Edit mission. These four are peripheral to open→edit→macro→save; handlers are registered; service/adapter layer is covered by `test_shapes_controls.py`; adapters are shared with verified sheets/ranges paths. Closing them = extend `build_full_access_fixture.py` with a named shape + a Form Control button, then call the four tools via MCP. |
| Focus-theft NOT VERIFIED (continuous) | Point-sampling only; no continuous foreground logger. | **Non-blocking.** Structural guarantee is strong: zero forbidden focus/input API call sites; owned Excel is `Visible=False` (enforced, quit-if-not-hidden); owned instance never observed foreground or visible. The spec (§5 Phase 5) explicitly instructs to mark this `NOT VERIFIED` and document the evidence needed rather than assert "never stolen" — Gate 5 passes on that basis ("all observable checks pass; any unobservable condition is explicitly documented"). |

No residual is a **critical defect**: the core edit/macro workflow works, no critical safety/integrity
requirement fails, and no original data is at risk. Both residuals are **coverage/observability gaps**, not
functional failures.

---

## 10. Verdict justification against §11

§11 vocabulary:
- `FULL EDIT MODE VERIFIED` — real MCP workflow **and all critical required capabilities** work, **all critical
  gates pass**, remaining **non-critical** limitations documented.
- `PARTIALLY VERIFIED` — the core workflow works, **but some required tools or conditions remain
  partial/unverified**.

Applying it honestly:

- **Gate 3** (MCP surface accounted for) passes by its own definition — every live tool is inventoried and has a
  status, with the 4 unverified tools *explicitly classified as not verifiable through MCP* with precise reasons
  and service-layer coverage.
- **Gate 4** (core edit/macro/edit/macro) passes fully (R2–R5).
- **Gate 5** (background execution) passes by its own definition — all *observable* non-interference checks pass;
  the one *unobservable* condition (focus-theft) is explicitly documented with the evidence needed.

Despite all critical gates passing on their own terms, §11's `FULL EDIT MODE VERIFIED` requires *all critical
required capabilities* to be MCP-verified with only *non-critical* limitations remaining. Here **four registered
MCP tools** (part of the 88-tool surface) have **no MCP-observed evidence**, and **one Phase-5 condition** is
formally unverified. That is precisely the `PARTIALLY VERIFIED` shape — *"the core workflow works, but some
required tools or conditions remain partial/unverified."*

The verdict is therefore **`PARTIALLY VERIFIED`**. It is **not** `FAILED` (the core edit/macro workflow works and
no critical safety/integrity requirement fails) and **not** `NOT VERIFIED` (the required real
Windows/Excel/MCP execution evidence *was* obtained and passes). The two residuals are non-blocking for the
Full-Edit mission but prevent an unqualified `FULL EDIT MODE VERIFIED` under §11's "all critical required
capabilities" wording.

---

## 11. Implementation commits (context; not modified by this gate)

Relevant commits from the orchestration (recorded in `full-edit-mode-orchestration-plan.md` ledger):
`c246edd` (writable execute default) · `46d9b1b` (full-access default + structured errors) ·
`47713f7` (macro background-safety classification + hash-bound registry) · `fe63af2` (approval denial →
`APPROVAL_DENIED`) · `0993370` (restore snapshot in execute mode — FE-16) ·
`5d649a6`/`e85c304` (interactive-UI veto tokens — FE-04) · `c28b3dd` (remove unused `pywinauto` — FE-07) ·
`cc7abd0` (structured-error surfacing matrix — FE-02) · `8be7245` (Phase-4 e2e steps 4/6/11/14 — FE-03) ·
`e5ebcc8` (real denial codes + read-only denial — FE-05) · `119c303` (document restore/worker/macro-snapshot/
recovery/classify limitations — FE-06/08/09/10/11) · `db91dbd` (Phase-3 matrix — FE-12) ·
`dece703` (Phase-4 e2e acceptance — FE-13) · `b9c13b0` (Phase-5 hidden execution — FE-14).

All FE tasks in the plan ledger are `Done+commit`. This gate (FE-15) adds **no** production/test changes.

---

## 12. Recommendation to Tech Lead

The `xlvba MCP` is a **genuinely usable Full-Edit MCP** for an AI agent: the mission workflow
(`open writable → inspect → edit settings table → run permitted macro → verify → edit → run again → verify →
save/snapshot/restore`) is proven end-to-end through the real MCP protocol in one session, with hidden Excel,
deny-by-default macro safety, preserved read-only/restricted modes, and byte-identical originals.

To upgrade the verdict to `FULL EDIT MODE VERIFIED`, close the two documented residuals:
1. Extend `tests/fixtures/build_full_access_fixture.py` with a named AutoShape/TextBox and a Form Control button,
   then exercise `shape_get` / `shape_set_on_action` / `control_get` / `control_set_on_action` through MCP (Phase-3
   §4 follow-up).
2. Run a continuous foreground-window logger (`EVENT_SYSTEM_FOREGROUND` or 10–50 ms `GetForegroundWindow` poll)
   across a full server lifecycle to upgrade focus-theft from NOT VERIFIED to PASS (Phase-5 §5 follow-up).

Neither is required for the core Full-Edit mission to be considered working and safe; both are recommended before
an unqualified release-gate sign-off.

---

## 13. Reproduction

```
# Decisive committed tests (serial)
.venv\Scripts\python.exe -m pytest tests/integration_excel/test_mcp_full_access_workflow.py \
  tests/integration_excel/test_mcp_error_surfacing_matrix.py -o addopts="" -q -p no:xdist -s

# Independent live MCP probe (throwaway; not committed)
.venv\Scripts\python.exe .ai\tmp\fe15_live_probe.py
.venv\Scripts\python.exe .ai\tmp\fe15_unclassified_probe.py

# Fixture integrity
.venv\Scripts\python.exe -c "import hashlib; print(hashlib.sha256(open(r'tests/fixtures/FullAccessWorkflow.xlsm','rb').read()).hexdigest())"
# -> 6b905837b8deeb549505f9a8a51c4f55507d377e4160bb5d9d6363a19ac1d4d9

# Host session
qwinsta
Get-Process -Id $PID | Select-Object Id, SessionId, ProcessName

# Forbidden-API sweep (expect only docstrings / denylist literals)
# grep -R "SetForegroundWindow|ShowWindow|SendKeys|AppActivate|..." src/excel_vba_mcp
```

---

**Validator sign-off:** Independent functional verification complete. Decisive tests re-run green (5 passed);
live MCP boundary re-confirmed independently (default writable execute, table mutation, two edit→macro cycles with
differing outputs 20/40, background-safe classification, interactive veto `INTERACTIVE_UI_NOT_SUPPORTED`,
fresh-server unclassified denial `BACKGROUND_SAFETY_NOT_VERIFIED`, read-only denial `WORKBOOK_READ_ONLY`,
restricted `APPROVAL_DENIED`, `excel_health` ok, 88 tools / 4 resource templates / 0 prompts); originals
byte-identical; no production/test/doc changes. Verdict: **`PARTIALLY VERIFIED`** with two documented,
non-blocking residuals (4 shape/control MCP tools; continuous focus-theft observation).