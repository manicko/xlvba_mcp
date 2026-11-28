# Phase 0 — Evaluation Environment & Baseline Integrity

- **Role:** Tech Lead (Phase 0 record). Read-only inspection. No production/test/config change.
- **Date:** 2026-11-26
- **Repo:** `C:\py_exp\xlvba_mcp`

## 1. Revision integrity (independently verified)

A concurrent/external process committed the working tree during early inspection. The two evaluation surfaces are now **both pinned commits**, which removes the earlier "uncommitted after-state" reproducibility risk.

| Surface | Commit | Tools | Resources | Notes |
|---|---|---|---|---|
| Baseline (pre-remediation) | `e2289ca01fd10136716a598cad1db6b6845fed73` | **91** (90 `add_tool` + 1 `@mcp.tool` `server_health`) | 4 (all URI-templated) | `environment_tools.py` absent; no annotations; 5-key error envelope |
| Remediated (accepted) | `81bcf2c` (`feat(xlvba): agent usability remediation`) | **92** (91 `add_tool` + 1 `@mcp.tool`) | 5 (added `xlvba://glossary`) | `environment_tools.py` present; annotations 92/92; additive error envelope |
| Current HEAD | `40c8b4d` (`docs(xlvba): agent usability audit, plan, and results`) | 92 (code identical to `81bcf2c`) | 5 | **Docs-only**: `git diff 81bcf2c..40c8b4d -- src/` is empty |

- Working tree at inspection end: clean except two untracked legacy plan files (`.ai/audit/plans/plan_01`, `plan_02`) unrelated to xlvba usability.
- A dangling `git stash create` merge commit `2404ed2` was produced during inspection (non-mutating; not on any branch; safe to ignore/GC). It is **not** an evaluation surface.
- Baseline tool count independently reproduced: `git show e2289ca:.../server.py` → `add_tool=90`, `@mcp.tool=1`, `@mcp.resource=4`.

## 2. Runtime

- Python **3.14.7**; `mcp` **2.3.0**; `mcp-types` **2.3.0**; `pydantic` **2.14.4** (`.venv`).
- Server: `excel-vba-mcp` v0.1.0, transport **stdio** (per benchmark meta), backend = local in-process MCP server.
- `xlvba_server_health` → `{status: ok, service: excel-vba-mcp, version: 0.1.0}`.
- Live Excel worker reachable: `xlvba_environment.excel.available=true`, `owned_pid=23140`.

## 3. Model / Kilo client

- **Running model (this Tech Lead session):** `kilo/baidu/ernie-5.5-flash-preview:free`. Subagents inherit the configured subagent model unless overridden (none overridden here).
- **Kilo version:** undetermined — `kilo` CLI not on PATH; global config read-restricted. Recorded limitation (see §6).
- **Global config `~/.config/kilo/kilo.jsonc`:** read blocked by `external_directory: deny` permission. MCP declaration, model params and tool-permission globs could not be inspected here. Recorded limitation (same as baseline §4).

## 4. Live remediated `xlvba_environment` (verbatim key facts)

- `server.access_mode="full"`; timeouts 90/300/600; `log_level=INFO`.
- `excel`: `available=true`, `application_version=null`, `build=null` (honest), `target_build="16.0.14332"`, `owned_pid=23140`, `vbom_available=null`.
- `capabilities`: vba/macros/pivot_tables/background_jobs/recovery=true; `non_activating_compile=false`; `pivot_filter_kinds=["manual"]`; `pivot_source_types=["database"]`.
- `policy`: approval mode full (no grants required), `snapshot.required_for_mutations=true`, allowed_roots = repo + temp + artifacts, `macro.classification_is_authorization=false`, `r13_preflight="veto"`.
- `glossary_uri="xlvba://glossary"`; `workbooks=[]`.

## 5. Glossary resource (reachable through Kilo)

- `list_mcp_resources(xlvba)` returns exactly one resource: `glossary` @ `xlvba://glossary`, `mimeType=text/markdown`, with a description listing the 14 jargon terms. **Reachable via Kilo = PASS** (Phase 3 will confirm the model can retrieve and use it).

## 6. Recorded limitations (carried into Phase 3)

- Kilo version and global MCP/model config unreadable (permission + PATH).
- Whether Kilo forwards `outputSchema`, `structuredContent`, annotations, and resource meta-tools to the model is **not directly introspectable** from inside a session; Phase 3 must infer from model-visible behavior and mark `UNSUPPORTED`/unverified honestly.
- Actual model input tokens are not directly observable from the session; token analysis must separate server `tools/list` size from client-transmitted context and actual model tokens.

## 7. Phase 0 verdict

- Baseline and remediated surfaces are **reproducible, committed revisions** (`e2289ca` vs `81bcf2c`/`40c8b4d`).
- The live runtime is the **remediated** surface and is healthy.
- The prior `agent_usability_results.md` is confirmed to be a **deterministic proxy only**; the model-in-the-loop benchmark was never executed. Phase 4 is genuinely new evidence.
- No production changes made. Evaluation may proceed.