# Remaining Usability — Results (plan_05)

- **Baseline:** HEAD `40c8b4d` (92 tools; list 481,641 chars; desc 37,125; input 82,373; output 350,758; annotations 92/92).
- **Final:** working tree after U6-A..E + U7 + U6-D + alignment (92 tools; list 464,914; desc 36,575; input 60,480; output 356,455; annotations 92/92).

## 1. Work packages (each Implementor → Validator → Tech Lead accept)
- **U6-B ACCEPT** — error envelope: `to_dict` keys `errorCode`/`details`; additive `INTERNAL_ERROR` + `CONFIRMATION_REQUIRED`; `call_tool` hatch closed (non-domain → structured `INTERNAL_ERROR`, no "Error executing tool" leak). 6 keys on every isError path.
- **U6-A ACCEPT** — `vba_compile`→`vba_compile_diagnostics`; `job_cancel`→`job_request_cancel`; `test_run_suite` removed (sole entry `test_suite_run`); count 92→91; zero residual old names (excl. `TestAction.VBA_COMPILE`).
- **U6-C ACCEPT** — capability-based routing in `xlvba_environment` (prefer-xlvba / use-excel-mcp matrix + examples; explicit no other-server detection).
- **U6-E ACCEPT** — `WorkbookInfo` two-path docs; cross-scope name-conflict docs; `confirm` flags on `snapshot_restore`/`excel_recover` (reject-before-action → `CONFIRMATION_REQUIRED`; discard/preserve + external-side-effect caveat).
- **U7 ACCEPT** (after `.bak` guard fix) — `excel_save_as` with source-hash capture, normcase/realpath compare, atomic temp→replace, single `.bak` (never overwritten if pre-existing), fail-closed `outcome`; session stays bound; 15 unit cases.
- **U6-D ACCEPT** — boilerplate compression: guard sentence → glossary token on 24 tools; approval clause → token; glossary +2 terms; desc −2,412 chars (gross); no safety text removed.

## 2. Token / surface delta (measured, not model-token)
| Metric | Baseline | Final | Delta |
|---|---|---|---|
| Tools | 92 | 92 | 0 (91 after A, +excel_save_as) |
| `tools/list` chars | 481,641 | 464,914 | **−16,727 (−3.5%)** |
| Description chars | 37,125 | 36,575 | −550 net (U6-D −2,412; +routing +save_as desc) |
| Input schema chars | 82,373 | 60,480 | **−21,893** (duplicate test DSL removed) |
| Output schema chars | 350,758 | 356,455 | +5,697 (SaveAsResult) |
| Annotations | 92/92 | 92/92 | — |

The net reduction is real but modest; the dominant lever was removing the duplicate `test_run_suite` input DSL. Character counts are not model-token measurements.

## 3. Benchmark
- Updated `.ai/tmp/benchmark_tasks.yaml` (31 entries): T22 `job_request_cancel`; T27 sole `test_suite_run`; T28 real `excel_save_as` (new path/source-overwrite/rejection/external-change).
- Deterministic proxy: **30/30 real tasks first-tool feasible; 31/31 arg keys valid** (`c1` = no-tool control).
- Full model-in-the-loop run through Kilo/Excel: **not executed here** (no agent runtime/Excel in this session); harness + tasks ready.

## 4. Verification
- `pytest tests/unit` → **1152 passed**; ruff clean on changed files; integration **collects** cleanly; envelope-key + doc alignment applied to integration tests and `docs/08`.
- Per-package Validators: U6-B/C/E (report_10b/10d), U6-A (10c), U7 (10e) all ACCEPT; final Auditor spot-check (report_10f) ACCEPT.

## 5. Remaining / deferred
- Full model-in-the-loop benchmark (needs Kilo+Excel runtime).
- `export_fixed_format`/PDF (out of scope); `xlvba_job_list` (deferred).
- `application_version`/`vbom_available` null server-side (non-COM); `build:null`.
- outputSchema still ~77% of payload; further reduction needs DTO field decisions (product).