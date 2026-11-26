# Report 08b — U1 Validator Verdict (Descriptions + Glossary)

- **Role:** Independent Validator for Work Package **U1** (descriptions + glossary). VALIDATION ONLY — no code/test/config modified.
- **Repo / HEAD:** `C:\py_exp\xlvba_mcp` @ `e2289ca` + U1 working-tree changes.
- **Contract:** `plan_03 — Agent Usability Remediation.md` §5 (description standard + glossary §5.3), §9 (routing), §12.1 (description changes), §12.4 (docs), §14 U1 (acceptance/tests/rollback).
- **Method:** source inspection (`server.py`, `models.py`, `macro_runner.py`, `names_adapter.py`, `test_tools.py`), `git diff`/`git show HEAD:` comparison, live `import excel_vba_mcp.server` + `list_tools()`/`list_resources()`, and `pytest tests/unit`.
- **Independence:** I did not implement U1.

---

## Per-item verdict

| # | Item | Verdict |
|---|------|---------|
| 1 | §5.1 section order on all workflow-critical `_DESC_*` | **APPROVE** |
| 2 | R-D3 guard sentence uniform; 3 U3 enum tools untouched | **APPROVE** |
| 3 | Truthfulness (R-D1) vs source | **APPROVE** |
| 4 | `_DESC_ENVIRONMENT` / `_DESC_RES_GLOSSARY` + `xlvba://glossary` (14 terms) + count 91 | **APPROVE** |
| 5 | Docs updated (EN + RU) | **APPROVE** |
| 6 | Scope discipline (no schema/enum/annotation/tool/DTO/save_as/job_list) | **APPROVE** |
| 7 | `pytest tests/unit` green; import + `list_tools()`=91 | **APPROVE** |
| 8 | `test_description_contract.py` quality | **APPROVE (with caveats)** |

**Final U1 verdict: ACCEPT** (no blocking reasons). Advisory recommendations in §9.

---

## Evidence

### Item 1 — §5.1 order (APPROVE)

All 25 workflow-critical descriptions follow the fixed order Purpose → Use when → Do not use when → Preconditions → Parameters → Behavior → Result → Next actions, with genuinely-empty sections omitted and none reordered. Verified by reading each constant:

- `excel_open` 216–233 (Purpose/Use when/Do not use when/Preconditions/Parameters/Behavior/Result/Next actions/Example) — full order.
- `excel_close` 234–243 (Do-not-use + Parameters omitted — empty).
- `excel_save` 244–253 (Parameters omitted).
- `excel_get_status` 254–263 (Do-not-use + Parameters omitted).
- `range_write_values` 651–670, `range_read_values` 632–644, `range_write_formulas` 671–685 — full order.
- `macro_run` 428–448 (Preconditions = numbered gate list), `macro_classify_background` 453–470.
- `table_update_rows` 577–593, `table_delete_rows` 594–606, `table_append_rows` 559–571.
- `name_create` 781–792, `name_update` 793–805.
- `sheet_add` 505–516, `sheet_copy` 522–533.
- `snapshot_restore` 290–299, `excel_recover` 270–284.
- `job_cancel` 487–498, `job_get` 471–486 (Do-not-use omitted; routing carried in Next actions).
- `vba_compile` 405–419, `vba_replace_component_source` 373–390, `vba_insert_code` 391–404.
- `test_run_suite` 301–317, `test_suite_run` 322–346.

No section is out of order. Terse (non-critical) guarded mutations correctly keep the R-D3 sentence without the full 8-section template, as §5.1 scopes the template to "Blocker/Major/workflow-critical" tools.

### Item 2 — R-D3 uniformity + U3 tools untouched (APPROVE)

- Uniform guard sentence `Guarded mutation: writable session + snapshot + approval (full-access default; restricted requires approval) + postcondition re-read.` is applied to every U1-owned guarded mutation. Full-template tools carry the identical body under `Preconditions:` (e.g. 378–379, 510–511, 582–583, 656–658, 677–678, 785–786, 798–799).
- Scan of the working tree: **0** remaining old-style `"requires a writable (working copy|session and a snapshot)"`, **0** `"allowed by default (full-access policy)"`, **0** short `"approval."` forms.
- The three U3-owned enum tools are **byte-identical to HEAD** (`git show HEAD:` vs working tree):
  - `_DESC_RANGE_SET_CONDITIONAL_FORMATTING` 763–770 — IDENTICAL (still `approval + postcondition re-read`, no audit).
  - `_DESC_PIVOT_FIELD_SET_LAYOUT` 895–907 — IDENTICAL (still `approval + audit + postcondition`).
  - `_DESC_PIVOT_FIELD_APPLY_FILTER` 950–961 — IDENTICAL (still `approval + audit + postcondition`).
  - The residual `approval + audit + postcondition` (×2) and the lone `approval + postcondition re-read` without parenthetical (line 769) are exactly these U3 tools — correctly left for U3 (D-U6).

### Item 3 — Truthfulness / R-D1 (APPROVE)

| Claim | Description | Source truth |
|---|---|---|
| `vba_compile` "diagnostic only — does **not** compile" | 406–407 | `compile_service` returns `status='indeterminate'`, `is_compile=false` (desc 412–416); no compiled/compile_error path. |
| `macro_classify_background` "never grants run permission" | 457–459 | `macro_tools.py` classify is "UNGATED bookkeeping"; `macro_run` independently enforces approval+BG06+R13 (desc 434–439). |
| `job_cancel` "intent-only; never interrupts" | 490–492 | sets `cancel_requested` only; single-STA call runs to completion (desc 494–495). |
| Range ceilings `rows≤100000, cols≤512, cells≤100000`; response `50000 cells / 8 MiB` | 659, 637–639 | `models.py:666–672`: `MAX_RANGE_CELLS=100_000`, `MAX_RANGE_ROWS=100_000`, `MAX_RANGE_COLS=512`, `MAX_WRITE_PAYLOAD_CELLS=100_000`, `MAX_RESPONSE_CELLS=50_000`, `MAX_RESPONSE_BYTES=8*1024*1024`. |
| `macro_run` "≤30 — the Excel Application.Run limit; bare `datetime.date` rejected, pass `datetime.datetime`" | 440–441 | 30-arg Excel limit documented `test_spec.py:35–36` (`MAX_MACRO_ARGS=30`), `audit.py:58–59`; bare-date rejection `macro_runner.py:96–103`. |
| `name_create`/`name_update` `refers_to` "with or without a leading `=`, normalized" | 787–788, 800–801 | `names_adapter.py:248–257` `_normalize_refers_to` "accept both forms". |
| `test_run_suite` has **no** `background` | 301–317 (no background; "Do not use when … need background=true — use test_suite_run") | `test_tools.py:163` `async def test_run_suite(suite: TestSuiteSpec)` (no `background`); `test_suite_run` 189–192 has `background: bool = False`. |

No description promises behavior the code lacks. The `≤30` is stated as an Excel constraint (true), not a server-side rejection claim — acceptable.

### Item 4 — Constants + glossary resource + count (APPROVE)

- `_DESC_ENVIRONMENT` exists 997–1012; `_DESC_RES_GLOSSARY` exists 1013–1018.
- `@mcp.resource("xlvba://glossary", name="glossary", description=_DESC_RES_GLOSSARY, mime_type="text/markdown")` registered 2224–2233; body `_GLOSSARY_MD` 2200–2221.
- Glossary contains exactly the **14** §5.3 terms: R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, SUSPECTED_STALLED, interrupted_unknown, indeterminate, working copy, snapshot gate (2202–2220).
- Live check: `list_tools()` → **91**; `list_resources()` → 1 static (`xlvba://glossary`) + 4 URI templates (templates are not returned by `list_resources`); `xlvba_environment` **not** in tools (U2 owns registration). Resource is not a tool — count stays 91.

### Item 5 — Docs (APPROVE)

EN (`docs/05,07,01`) and RU (`docs/ru/05,07,01`) updated:
- `test_run_suite` `background` mis-attribution removed — table now `test_run_suite | suite:TestSuiteSpec` (no background); `test_suite_run | … background:boolean=false?`. Prose: "`test_run_suite` is blocking only and has no `background` parameter."
- Test-tool routing present — EN "**Test-tool routing.**"; RU "**Маршрутизация тест-инструментов.**".
- `xlvba_environment` mentioned as planned forward reference (EN "Forward references (planned, not yet exposed)"; RU "Плановые ссылки (ещё не опубликованы)").
- Tool count accurate — "91 tools (91 exposed) and 5 read-only resources (four per-workbook URI templates plus static `xlvba://glossary`)" / RU "91 инструмент … 5 read-only ресурсов".

### Item 6 — Scope discipline (APPROVE)

- `git status`: only `server.py`, 6 docs, and new `tests/unit/test_description_contract.py` changed. `models.py`, `statuses.py`, `errors.py`, `session_service.py` **untouched** → no schema/enum/DTO/environment changes.
- Diff of `server.py` adds exactly **one** `@mcp.resource(` (glossary) and **zero** `annotations=`; **no** `server.add_tool(` added/removed (90 `add_tool` + 1 `@mcp.tool` = 91, unchanged).
- `save_as`=0, `job_list`=0 occurrences → U5 items not implemented.

### Item 7 — Tests + import (APPROVE)

- `pytest tests/unit -q` → **1081 passed** (0 failed, 0 errors) in ~39s.
- `import excel_vba_mcp.server` succeeds; `asyncio.run(mcp.list_tools())` → **91**; glossary resource resolvable; `xlvba_environment` absent.
- `tests/unit/test_description_contract.py` → 7 passed.

### Item 8 — Test quality (APPROVE with caveats)

`test_description_contract.py` is **not** superficial — it asserts real contract properties:
- `test_tool_count_is_91_exposed` (== 91).
- `test_critical_tools_follow_template` — 25 tools contain `Purpose:`/`Use when:`/`Preconditions`/`Result:`/`Next actions:`.
- `test_truthful_phrases_present` — the 3 R-D1 phrases.
- `test_examples_present` — 4 R-D8 example tools.
- `test_guarded_mutations_use_uniform_guard_sentence` — exact R-D3 sentence across 18 U1-owned guarded mutations (U3 tools correctly excluded).
- `test_glossary_resource_registered_and_complete` — `xlvba://glossary` registered + all 14 terms.
- `test_environment_and_glossary_constants_exist` — constants present; `xlvba_environment` not yet a tool.

**Weaknesses (non-blocking, see §9):** the template test checks only **5 of 8** markers (not `Do not use when` / `Parameters` / `Behavior`) and does **not** assert section **order**; and the §14 "forbidden phrases absent" half is **not** implemented (only positive-presence checks). Because I independently verified order and truthfulness above, these gaps do not change the verdict.

---

## Unsafe / misleading descriptions found

No unsafe description. Two minor accuracy nits (advisory, not blocking):

1. **`_DESC_SAVE` dangling glossary reference** (`server.py:247–248`): "there is no save-as (see `xlvba://glossary` / docs 09)" — the glossary (2200–2221) defines **no** `save-as` term, so the `xlvba://glossary` pointer for save-as resolves to nothing. Reference docs 09 only, or add a save-as term in a later package.
2. **`_DESC_STATUS` forward reference** (`server.py:258–259`): "the additive policy block (U2) co-locates per-workbook approval/snapshot context" is present-tense for a U2 feature not yet shipped. It is explicitly tagged `(U2)` and matches plan §12.1, so acceptable, but a client reading only the description could infer the block exists today.

---

## Rollout / dependency notes

- U1 is description + docs + one static resource only; tool count 91→91; no state; rollback = revert `_DESC_*` constants + glossary resource (plan §14). No dependency on U2/U3 artifacts is violated: the 3 enum descriptions are untouched (D-U6 preserved), and `_DESC_ENVIRONMENT` is a dormant constant until U2 registers the tool.
- U2 dependency (glossary mirror/uri) is satisfied: `xlvba://glossary` exists and `_DESC_ENVIRONMENT` already cites it.

## Required fixes

None blocking.

## Advisory recommendations

1. Strengthen `test_description_contract.py`: assert the **full ordered** marker sequence (including `Do not use when`/`Parameters`/`Behavior` where present) and add the §14 **forbidden-phrase absence** checks (e.g. `vba_compile` must not contain a "compiled successfully" success claim; `job_cancel` must not claim it interrupts the COM call) — implemented so the legitimate "never interrupts" wording is not falsely flagged.
2. Fix the `_DESC_SAVE` → `xlvba://glossary` dangling reference (point to docs 09, or add a `save-as` glossary entry when U5 is scheduled).
3. Re-baseline `_DESC_STATUS` wording once U2 ships so the policy block is described in the present tense only after it exists.

---

*Validator: independent (did not implement U1). Verdict: U1 ACCEPT.*