# Agent Usability — Real Kilo Runtime Compatibility

- **Role:** Tech Lead (Phase 3). Real Kilo client + live remediated MCP server (`81bcf2c`/`40c8b4d`).
- **Date:** 2026-11-26
- **Method:** live read-only tool calls through the actual Kilo MCP connection (this session is a real model-in-the-loop client), plus source verification by the Auditor (Phase 2 change map). No production changes.
- **Classification:** PASS / FAIL / PARTIAL / BLOCKED / UNSUPPORTED.

## A. Tool descriptions and schemas

| Check | Verdict | Evidence |
|---|---|---|
| Model receives current descriptions | **PASS** | Auditor change map: 53/91 descriptions changed to the 8-section standard; the live catalog is the `81bcf2c` surface. Raw `tools/list` text is not directly introspectable from inside a session (limitation), but behavior is consistent with the remediated contract. |
| Argument schemas match server contract | **PASS** | `range_write_values` and `range_set_conditional_formatting` enforced the live input schema; a schema-violating call returned `INVALID_RANGE` with a pydantic `validation` detail. |
| Removed enum values no longer advertised | **PASS** | `range_set_conditional_formatting` `spec.type="color_scale"` → `INVALID_RANGE`; the enum error lists only `cell_value, expression, text_string, blanks, no_blanks, errors, no_errors, time_period, top10, above_average, unique_values`. `color_scale/data_bar/icon_set/unsupported` are absent. |
| Valid enum values remain available | **PASS** | Same error lists the 11 valid values (read/classify enums retained per Auditor). |
| Defaults and required fields preserved | **PASS** | Auditor: 0 required/default changes between revisions. |

## B. Environment discovery (`xlvba_environment`)

| Check | Verdict | Evidence |
|---|---|---|
| Model can call `xlvba_environment` | **PASS** | Live call returned a full `EnvironmentReport`. |
| Receives effective access mode + policy | **PASS** | `server.access_mode="full"`; `policy.approval/snapshot/path/macro` populated; `policy.classification_is_authorization=false`. |
| Distinguishes server capabilities from live Excel state | **PASS** | `capabilities` (static: vba/macros/pivot/background/recovery, `non_activating_compile=false`, `pivot_filter_kinds=["manual"]`) separated from `excel` (live: `available=true`, `owned_pid=23140`). |
| Handles null / unavailable values honestly | **PASS** | `excel.build=null`, `application_version=null`, `vbom_available=null` (no fabrication); `target_build="16.0.14332"` labelled as design constant. |
| Model does not invent unknown build | **PASS (tool support)** | Tool returns honest `build:null`; model-side non-fabrication is a Phase 4 behavioral check. |
| Additive `SessionStatus.policy` reaches model | **PASS** | `excel_close` response carried `policy{approval_mode, mutating_requires_grant, executing_requires_grant, recovery_auto_allowed, snapshot_required_for_mutations, has_snapshot, classification_is_authorization, r13_preflight}`. |

## C. Glossary resource

| Check | Verdict | Evidence |
|---|---|---|
| Glossary accessible through Kilo | **PASS** | `list_mcp_resources(xlvba)` → 1 resource `glossary` @ `xlvba://glossary` (`text/markdown`). |
| Model can retrieve the terms it needs | **PASS** | `read_mcp_resource(xlvba, xlvba://glossary)` returned all 14 terms (R13, BG06, preflight, truth signal, fail-closed, RECOVERY, STA, Stage 1, normalized source hash, SUSPECTED_STALLED, interrupted_unknown, indeterminate, working copy, snapshot gate). |
| Descriptions do not depend on an unavailable resource | **PASS** | Glossary is also mirrored compactly in `xlvba_environment`/descriptions (Auditor); the resource is additive, not load-bearing. |

## D. Errors and annotations

| Check | Verdict | Evidence |
|---|---|---|
| Execution errors expose structured fields to the client | **PASS** | Write to a read-only workbook → `{code:WORKBOOK_READ_ONLY, message, hresult, hresult_hex, context, errorCategory:"missing_prerequisite", isRetryable:false, suggestedNextActions:["re-open with read_only=false (mode=execute)"]}` — full additive envelope visible to the model. |
| Model can distinguish permission / validation / transient / unknown | **PARTIAL** | Demonstrated live: `missing_prerequisite` (read-only) and `validation` (pruned enum). `permission_denied`, `transient_com_failure`, `interrupted_unknown` categories exist in source + unit tests (Auditor) but were not exercised live here (require restricted mode / induced stall / seeded job). |
| Annotations actually forwarded and available to the model | **UNSUPPORTED (unverifiable in-session)** | Server emits truthful annotations 92/92 (Auditor), but the Kilo client's rendering of `readOnlyHint/destructiveHint/idempotentHint` to the model is not observable from inside a session. Not marked verified. |
| Model does not assume unsupported metadata is present | **PASS (design)** | Descriptions reference the glossary by URI and do not claim client-only metadata; model behavior verified in Phase 4. |

## Summary

- **PASS:** descriptions, schemas, pruned-enum enforcement, environment discovery, honest null handling, `SessionStatus.policy`, glossary retrieval, structured error envelope.
- **PARTIAL:** live error-category coverage (2 of 8 categories exercised; remainder source/test-verified).
- **UNSUPPORTED:** in-session observation of annotation forwarding by the Kilo client.
- **FAIL / BLOCKED:** none.

**Limitations (explicit):** Kilo version and global MCP/model config are unreadable (`external_directory: deny`; `kilo` not on PATH). Raw `tools/list` text and annotation rendering are not introspectable from within a session; these are reported as limitations, not as verified capabilities. No production changes were made.