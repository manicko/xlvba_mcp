# Research 2A — Actual MCP/Kilo Discovery Capabilities & the Smallest Practical Progressive-Discovery Solution

- **Role:** Researcher 2A (independent; did **not** read `research_1a_*` / `research_1b_*`).
- **Repo:** `C:\py_exp\xlvba_mcp` @ git HEAD `e2289ca01fd10136716a598cad1db6b6845fed73`.
- **Scope:** Research only. No production code/tests/config modified.
- **Method:** Installed-stack inspection (`.venv`), source reading, empirical measurement of the live tool catalog, Kilo official docs, and direct observation of this live Kilo CLI session's own toolset.

---

## 0. TL;DR verdict

| Option | Verdict | Why (one line) |
|---|---|---|
| MCP `defer_loading` | **needs-client-feature** (and server) | Not present anywhere in installed `mcp`/`mcp_types` 2.3.0; Kilo CLI has no deferral — it loads all MCP tools "alongside built-in tools". |
| Custom `xlvba_search_tools` (standalone) | **not supported as a context reducer** | In Kilo it just adds a 92nd always-loaded definition; the other 91 still load. Net context **increases**. |
| Domain tool groups | **not supported (no context benefit)** | Kilo has no group/defer concept; grouping is cosmetic naming only. |
| Resources for progressive disclosure | **supported by Kilo (docs) — partial** | Kilo exposes `list_mcp_resources` / `read_mcp_resource` on demand. Only covers the 4 existing resources, not the 90 tools. Not observed in this live session (see §4.4). |
| Shorter descriptions + on-demand detail | **low impact** | Descriptions are only **~5%** of the tool payload. The dominant cost is auto-generated `outputSchema` (**~75%**). |
| **Reduce `outputSchema` footprint** (server-side) | **supported today, highest impact** | `mcp` 2.3.0 emits `outputSchema` for all 91 tools; `structured_output=False` / simpler return types cut the 75% driver. This is the smallest high-impact lever. |

**The dominant cost driver is NOT descriptions — it is the auto-derived `outputSchema` (338,547 of 451,108 chars ≈ 75% of the `tools/list` payload). The baseline's "none emitted" assumption is wrong: all 91 tools emit `outputSchema` today.**

---

## 1. Verified installed stack

| Component | Version | Evidence |
|---|---|---|
| Python | 3.14.7 | `.venv\Scripts\python.exe --version` → `Python 3.14.7` |
| `mcp` | 2.3.0 | `importlib.metadata.version('mcp')` → `2.3.0` |
| `mcp_types` | 2.3.0 | `importlib.metadata.version('mcp-types')` → `2.3.0` |
| `pydantic` | 2.14.4 | `importlib.metadata.version('pydantic')` → `2.14.4` |
| Declared floor | `mcp>=2.3,<3`, `pydantic>=2.0.0`, Python `>=3.11,<3.15` | `pyproject.toml:6,9-14` |

Server module: `src/excel_vba_mcp/server.py` (1950 lines), package `excel_vba_mcp` v0.1.0.

---

## 2. What the xlvba server emits today (source + empirical)

### 2.1 Registration mechanism (source)
- Tools: `server.add_tool(fn, name=..., description=...)` — 90 calls in `_register_tools` (`server.py:1476-1704`), plus **1** `@mcp.tool()`-decorated `server_health` (`server.py:1834-1838`). **Total = 91 tools.**
- **No `annotations=` argument is passed on any `add_tool` call** (only `fn`, `name`, `description`). **No `title=` either.**
- **No `structured_output=` override** is passed anywhere → the SDK default (`None` = auto-detect) applies to every tool.
- Resources: 4 `@mcp.resource(uri_template, name=, description=, mime_type=)` (`server.py:1869-1943`). All 4 URIs contain `{...}` variables → all 4 register as **templates**, 0 static resources.
- Prompts: 0.

### 2.2 Empirical measurement (live import, no Excel started)
Script: `.ai/tmp/measure_tools_2a.py` (imports `excel_vba_mcp.server`, calls `list_tools()` / `list_resource_templates()`; lifespan does not run on import). Raw output:

```
TOOL_COUNT=91
RESOURCE_COUNT=0
RESOURCE_TEMPLATE_COUNT=4
TOOLS_JSON_CHARS=451108
LIST_RESULT_JSON_CHARS=451119
TOTAL_DESC_CHARS=24272
TOTAL_INPUT_SCHEMA_CHARS=80656
TOTAL_OUTPUT_SCHEMA_CHARS=338547
TOOLS_WITH_OUTPUT_SCHEMA=91
TOOLS_WITH_ANNOTATIONS=0
TOOLS_WITH_TITLE=0
AVG_DESC_CHARS=266.7
AVG_INPUT_SCHEMA_CHARS=886.3
```

Top definitions by size (total = desc+input+output chars):
```
29433  workbook_get_model   desc=344  input=239   output=28850
28640  test_suite_run       desc=642  input=22314 output=5684
26990  test_run_suite       desc=543  input=22157 output=4290
11439  sheets_list          desc=211  input=155   output=11073
11418  sheet_set_visibility desc=256  input=280   output=10882
... (sheet_* all ~11.1-11.4K, dominated by ~10.9K output schema each)
```

**Conclusion:** `outputSchema` **is emitted today for all 91 tools** (auto-derived from structured return annotations such as `-> CompileResult`, `-> list[SheetInfo]`, `-> dict[str,str]`). `annotations` and `title` are **not** emitted. This **contradicts** the baseline note "outputSchema / structuredContent / annotations currently emitted: none (verify in Phase 2)" — annotations: correct (none); **outputSchema: incorrect (all 91).**

---

## 3. Server-side MCP feature support (installed `mcp` 2.3.0) — inspected modules

| Feature | Supported? | Inspected module / class (evidence) |
|---|---|---|
| Tool `outputSchema` | **Yes** | `mcp/server/mcpserver/tools/base.py` → `Tool.output_schema` (cached_property, lines 60-62) from `FuncMetadata.output_schema`; `mcp/server/mcpserver/utilities/func_metadata.py:122,128-131` (schema built from return annotation); emitted in `mcp/server/mcpserver/server.py` `MCPServer.list_tools()` → `MCPTool(..., output_schema=info.output_schema)` (lines 516-531). Wire type: `mcp_types/_types.py` `Tool.output_schema` (line 1424). |
| `structuredContent` | **Yes** | `mcp_types/_types.py` `CallToolResult.structured_content` (line 1474); produced by `func_metadata.py` `FuncMetadata.convert_result()` (lines 186-226) which builds `CallToolResult(content=..., structured_content=...)`. |
| Tool `annotations` (readOnlyHint / destructiveHint / idempotentHint / openWorldHint) | **Yes** | `mcp_types/_types.py` `ToolAnnotations` (lines 1365-1408: `read_only_hint`, `destructive_hint`, `idempotent_hint`, `open_world_hint`); `Tool.annotations` field (`base.py:56`); `MCPServer.add_tool(..., annotations=ToolAnnotations|None)` (`server.py:618-656`) and `@tool(annotations=...)` (`server.py:669-737`); emitted in `list_tools()` (`server.py:526`). |
| Resources with URI templates (RFC 6570) | **Yes** | `MCPServer.resource(uri, ...)` (`server.py:788-947`) parses `UriTemplate.parse(uri)`; `{var}` ⇒ template via `ResourceManager.add_template`; `MCPServer.list_resource_templates()` → `MCPResourceTemplate` (`server.py:568-582`); type `mcp_types/_types.py` `ResourceTemplate` (line 821). xlvba already uses this for all 4 resources. |
| `defer_loading` (or similar) | **No** | `grep -SimpleMatch defer_loading` / `deferLoading` over the **entire** `.venv/Lib/site-packages` (mcp, mcp_types, pydantic, …) → **zero matches**. No `defer_loading`, no `search_tools` meta-tool, no lazy-load flag in the server SDK. |

---

## 4. Kilo CLIENT capability matrix

### 4.1 Method & limitation
- Global `~/.config/kilo/kilo.jsonc` **exists but is read-restricted**: `read` returned `permission external_directory pattern:* action:deny source:global`. **Live global config contents could not be read.** Recorded limitation; conclusions rely on official Kilo docs + the project `.kilo/` config + direct observation of this live session.
- Project config: no root `kilo.json`; `.kilo/` contains agents/commands/rules only (no `mcp` block). No project-level MCP override present.

### 4.2 Capability matrix

| Capability | Status | Source / evidence |
|---|---|---|
| All MCP tool definitions loaded into context at once | **Yes (confirmed)** | Kilo CLI doc: *"Once added, MCP tools are automatically available to the LLM alongside built-in tools"* + *"MCP servers add to your context… Certain MCP servers with many tools can quickly add up and exceed the context limit"* — `kilo.ai/docs/automate/mcp/using-in-cli`. **Direct observation:** this live session exposes every `xlvba_*` and `excel-mcp_*` tool as an individual first-class tool (no dispatcher). |
| Deferred / lazy tool loading (`defer_loading`) | **Unsupported** | No mention in any Kilo doc; `defer_loading` is an Anthropic Tool-Search feature, not MCP, and not implemented by Kilo. Kilo has no per-server "defer tools" setting (contrast: Copilot CLI `deferTools`, GitHub docs). |
| On-demand tool discovery / `search_tools` meta-tool | **Unsupported (native)** | No Kilo-native search meta-tool in the live toolset. Kilo's *own* remote MCP server implements search+call as **application-level** tools (`blog.kilo.ai/p/mcp`, 2026-10-13) precisely because *"The Kilo platform's API is too large to load as one tool per endpoint"* — i.e., Kilo the client does not provide discovery; a server must build it. |
| `outputSchema` forwarded to model | **Unknown (likely)** | MCP `Tool` object includes `outputSchema`; Kilo loads "MCP tools … alongside built-in tools", so the full definition (incl. outputSchema) is the likely payload — but Kilo does not document its internal serialization. Not verifiable from inside the session. |
| `structuredContent` display/use | **Unknown** | Not documented by Kilo. `use_mcp_tool`/MCP handling processes text/image/resource content; structuredContent handling not specified. |
| MCP tool annotations (readOnly/destructive/idempotent/openWorld) | **Transmitted: likely · Used by Kilo: no** | Annotations are part of the `Tool` object (so transmitted). But Kilo's approval is **name-based** `{server}_{tool}` glob permissions (`using-in-kilo-code` → Auto Approve Tools), not annotation-driven. No doc states Kilo auto-approves by `readOnlyHint`. |
| Resources exposed to the agent | **Yes (per docs) — not observed live** | *"Resource-capable servers also make the `list_mcp_resources`, `list_mcp_resource_templates`, and `read_mcp_resource` tools available to the agent… the agent fills in a template, then reads the resulting resource URI"* — `kilo.ai/docs/automate/mcp/using-in-kilo-code` (§ Server instructions and resources). **Discrepancy:** in this live session those three meta-tools are **not** present in the toolset despite xlvba advertising 4 resource templates and excel-mcp being resource-capable. Cause unresolved (build/version/gating). |
| Server `instructions` injected | **Yes** | *"When a connected MCP server provides instructions, Kilo adds them to the model context"* (same page). Omitted when every tool from the server is denied. |
| Per-tool permission granularity | **Yes, name-glob** | `permission` keys `{server}_{tool}` with globs (skill + docs). |

### 4.3 The decisive architectural fact
Kilo loads **every** MCP tool definition (name + description + inputSchema, and almost certainly outputSchema) into the model context at session start. There is **no** client-side deferral, grouping, or search. Therefore:
- A standalone `xlvba_search_tools` **adds** a 92nd always-loaded definition and removes nothing → **net context increase**.
- The **only** way a search/discovery tool reduces initial context in Kilo is to **collapse the 90 tools behind a generic dispatcher** (e.g., `xlvba_call(tool_name, arguments)` + `xlvba_search_tools(query)`), so the 90 definitions become *data returned by search* rather than *tool schemas*. That is exactly the pattern Kilo's own MCP server uses (2 tools total). It is a **large architectural change** (production), with tradeoffs (see §6).

### 4.4 Resources: the exact progressive-disclosure mechanism
Per Kilo docs, resources are the one progressive mechanism the client supports: the agent is given `list_mcp_resource_templates` / `list_mcp_resources` / `read_mcp_resource` and reads **only** the URIs it needs, on demand, through the normal read-approval flow. xlvba already ships 4 templates (`excel://workbooks/{workbook_id}/model`, `.../vba/components`, `.../vba/components/{component_name}`, `excel://test-runs/{run_id}`). **Caveat:** these three meta-tools were **not observed** in the live session, so the doc-described behavior could not be confirmed here (§7).

---

## 5. Baseline token-cost measurement & dominant driver

Measured from the live catalog (§2.2). Token estimates use ~4 chars/token (English/JSON) with a denser-JSON lower bound of ~3.5.

| Component | Chars | ≈ Tokens (4:1) | Share of `tools/list` |
|---|---|---|---|
| Descriptions | 24,272 | ~6,068 | **5.4%** |
| Input schemas | 80,656 | ~20,164 | **17.9%** |
| **Output schemas** | **338,547** | **~84,637** | **75.0%** |
| **Total `tools/list` JSON** | **451,108** | **~112,777** (≈128.9K at 3.5:1) | 100% |

**Dominant cost driver = auto-generated `outputSchema` (~75%).** Descriptions — the baseline's focus (~28,907 source-region chars ≈ 7,200 tokens) — are the **smallest** component (~5%). The baseline's 28,907 figure is the *source-region* size (Python syntax/quotes/comments); the *runtime* description strings total 24,272 chars.

The single largest definitions are `workbook_get_model` (28.8K output schema), `test_suite_run` / `test_run_suite` (22.1K/22.3K **input** schemas — the declarative suite DSL), and the `sheet_*` family (~10.9K output schema each, repeated 7×).

**Implication:** any remediation that only shortens prose descriptions yields ≤5% savings. The real lever is the 75% `outputSchema` block (and secondarily the 18% input schemas of the two test tools).

---

## 6. Verdict on each discovery option + smallest practical recommendation

### 6.1 Option verdicts
1. **MCP `defer_loading` → needs-client-feature (and server).** Absent from installed `mcp`/`mcp_types` 2.3.0 (zero grep hits) and from Kilo. Dead end for this stack.
2. **Custom `xlvba_search_tools` (standalone) → not supported as a context reducer.** Adds a 92nd always-loaded tool; the other 91 still load. **Net worse.** Only helps if paired with a generic `xlvba_call` dispatcher that collapses 90→2 definitions (large production change; erodes per-tool `{server}_{tool}` permissions and client-side input-schema validation; contradicts the typed-DTO/structured-error design).
3. **Domain tool groups → not supported (no context benefit).** Kilo has no group/deferral concept; grouping is cosmetic (name prefixes). Zero token savings.
4. **Resources for progressive disclosure → supported by Kilo (docs), partial coverage.** The only client-native on-demand mechanism. Covers the 4 existing resources; does **not** cover the 90 tools. Expanding to a resource catalog (e.g., `xlvba://catalog/{domain}`) is the only no-dispatcher progressive option, but is contingent on Kilo actually surfacing the resource meta-tools (docs yes; **not observed live**).
5. **Shorter descriptions + on-demand detail → low impact.** Descriptions are ~5% of cost. Trimming them is cosmetic against a 451KB payload.
6. **Reduce `outputSchema` footprint → supported today, highest impact, smallest high-value change.** `mcp` 2.3.0 supports `structured_output=False` per tool (`server.py add_tool(..., structured_output=...)`, `func_metadata.py:364-365`) and simpler return annotations. Suppressing outputSchema where the structured channel is unused, and/or slimming the giant return DTOs (`WorkbookModel`, `SheetInfo`, the test-suite DSL), directly attacks the 75% driver. This is a **server-side** change compatible with the installed SDK — no client feature required.

### 6.2 Smallest practical recommendation (compatible with installed stack)
Given research-only scope and "smallest practical":
- **Do not add a standalone search tool** — in Kilo it increases context.
- **Smallest high-impact, in-stack change:** cut the `outputSchema` block (the 75% driver) via `structured_output=False` on tools whose structured output is redundant with the text content, and/or slim the oversized return models. This needs no Kilo feature and no dispatcher.
- **If a discovery tool is mandated**, the minimal *effective* form is a **two-tool search+call dispatcher** (Kilo's own MCP pattern), accepting the permission/validation tradeoffs — this is a larger architectural change, not "smallest".
- **Resources** remain the only no-dispatcher progressive option; expanding them is worthwhile **only after** confirming Kilo surfaces `read_mcp_resource` in the target build.

**Does a custom search tool reduce initial context in THIS client?** **No** — not unless it replaces the 90 individual tools with a generic dispatcher. As an added tool it is a net increase.

---

## 7. What could NOT be verified (and why)
1. **Live global Kilo config** (`~/.config/kilo/kilo.jsonc`): read blocked by `external_directory: deny` (global). Relied on docs + project `.kilo/` + live observation instead.
2. **Whether Kilo forwards `outputSchema` to the model / renders `structuredContent`:** Kilo does not document its internal tool serialization; not observable from inside a session. Marked *unknown/likely*.
3. **Whether Kilo applies MCP `annotations`:** transmitted (part of `Tool`) but Kilo approval is name-glob based; no doc of annotation-driven behavior. Marked *used-by-Kilo: no/unknown*.
4. **Resource meta-tools in the live build:** Kilo docs say `list_mcp_resources` / `list_mcp_resource_templates` / `read_mcp_resource` are exposed for resource-capable servers, but **these three tools were absent from this live session's toolset** despite xlvba advertising 4 templates. Cause (build/version/gating) unresolved → resource-based progressive disclosure is *doc-supported but not confirmed here*.
5. **Exact tokenization:** char→token ratios are estimates (~4:1, with 3.5:1 lower bound); the ~113K–129K token range is approximate, not a tokenizer measurement.
6. **Baseline tool-count reconciliation:** baseline said 90 `add_tool`; empirical total is **91** (90 `add_tool` + 1 `@mcp.tool` `server_health`). The baseline's 90 refers to `add_tool` calls only.

---

## 8. Evidence index
- Source: `src/excel_vba_mcp/server.py` (registration 1476-1704; `server_health` 1834-1838; resources 1869-1943; `_DESC_*` 216-770).
- Installed SDK: `mcp/server/mcpserver/tools/base.py` (Tool, output_schema, annotations); `mcp/server/mcpserver/server.py` (add_tool 618-656, list_tools 516-531, resource 788-947, list_resource_templates 568-582); `mcp/server/mcpserver/utilities/func_metadata.py` (output_schema 122-131, convert_result 186-226, structured_output 364-477); `mcp_types/_types.py` (ToolAnnotations 1365, Tool 1411, CallToolResult 1463, ResourceTemplate 821).
- `defer_loading` absence: full-tree grep over `.venv/Lib/site-packages` (0 hits).
- Measurement: `.ai/tmp/measure_tools_2a.py` + raw output in §2.2.
- Kilo docs: `kilo.ai/docs/automate/mcp/using-in-cli` (loads alongside built-ins; context warning), `kilo.ai/docs/automate/mcp/using-in-kilo-code` (resources meta-tools; name-glob permissions; instructions), `kilo.ai/docs/automate/how-tools-work`, `kilo.ai/docs/automate/mcp/overview`, `blog.kilo.ai/p/mcp` (Kilo's own search+call server), `kilocode-legacy` `use-mcp-tool` (legacy dispatcher, archived).
- MCP best practices (context for `defer_loading`/progressive discovery as a client pattern, not a Kilo feature): `modelcontextprotocol.io/docs/2026-07-28/develop/clients/client-best-practices`.
- Baseline: `.ai/audit/reports/agent_usability_baseline.md` (28,907-char description region; "none emitted" assumption corrected here).