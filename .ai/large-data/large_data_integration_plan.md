# Integration Plan — wire new tools + one-time WP0 baseline update

Server wiring and the frozen WP0 baseline update are intentionally deferred to ONE integration step
after WP-E2/F modules are merged, so the consolidation baseline is touched exactly once. Two new
justified tools are added (large-data workflows that cannot be done through existing tools without
full materialization / STA occupancy):

- `workbook_file_read` — non-COM streaming OOXML read (worksheet rows, metadata, pivot-cache records)
  from an immutable snapshot; cannot be done via COM range/pivot tools without materializing the sheet.
- `data_export` — stream a range/table/pivot-result to a CSV/JSONL artifact; the artifact store is
  bytes-in-memory and no existing tool streams to a file.

## A. server.py wiring
1. Imports: `from excel_vba_mcp.application.ooxml_service import OoxmlService`;
   `from excel_vba_mcp.application.data_export_service import DataExportService`;
   `import excel_vba_mcp.api.ooxml_tools as ooxml_tools`;
   `import excel_vba_mcp.api.export_tools as export_tools`; and the tool funcs `workbook_file_read`,
   `data_export` + their description constants.
2. `_Runtime`: add `ooxml: OoxmlService` and `data_export: DataExportService`.
3. `_build_runtime`: construct `ooxml = OoxmlService(session, snapshot, config.large_data)` and
   `data_export = DataExportService(ranges, tables, pivots, artifacts, config.large_data)`; pass into
   `_Runtime(...)`. (Order: after ranges/tables/pivots/artifacts exist.)
4. Lifespan worker-started branch: `ooxml_tools.set_ooxml_service(runtime.ooxml)` and
   `export_tools.set_data_export_service(runtime.data_export)`; in `finally`: set both to `None`.
5. `_register_tools`: `server.add_tool(workbook_file_read, name="workbook_file_read",
   description=_DESC_WORKBOOK_FILE_READ, annotations=_annotations("workbook_file_read"))` and the same
   for `data_export`. Add both names to `_ANN_READ_ONLY`? NO — `data_export` writes a file
   (non-read-only, non-destructive, idempotent); `workbook_file_read` IS read-only. Add
   `"workbook_file_read"` to `_ANN_READ_ONLY` and `_ANN_IDEMPOTENT`; add `"data_export"` to
   `_ANN_IDEMPOTENT` only (not read-only, not destructive).

## B. One-time WP0 baseline update (documented intended surface change)
Recompute the live measurement (the fixture `baseline.command` shows how):
```
& .\.venv\Scripts\python.exe -c "import asyncio,json,excel_vba_mcp.server as s; t=asyncio.run(s.mcp.list_tools()); data=[x.model_dump(by_alias=True) for x in t]; print('count',len(t)); print('bytes',len(json.dumps(data,default=str).encode()))"
```
Live surface is the consolidated 41-tool set (~337077 bytes); after wiring it is 43 tools. Set
`baseline.serialized_bytes` to the **live-recomputed 43-tool value** (the honest current-surface
measurement) and `baseline.tool_count` 92→94 (ledger rows), `measured_at`→now ISO. Keep `command`/`sdk`.

Update `tests/fixtures/wp0_baseline.json` coherently (net-new standalone tools have `old_name == tool`,
matching existing standalone rows like `server_health`):
- `ledger`: append `{"old_name":"workbook_file_read","target":{"kind":"standalone","tool":"workbook_file_read"}}`
  and the same shape for `data_export`.
- `security`: append `{"tool":"workbook_file_read","risk":"read","kilo_policy":"ask (xlvba_*)","new_name":"workbook_file_read","permission_profile":"P_READ"}`
  and `{"tool":"data_export","risk":"additive-mutation","kilo_policy":"ask (xlvba_*)","new_name":"data_export","permission_profile":"P_EXEC"}`
  (copy exact profile strings from comparable rows).
- `truth`: append one standalone row per new tool (REQUIRED — every annotation tool must have a truth
  row so `test_tool_annotations_match_truth_derivation` holds):
  `{"tool":"workbook_file_read","action":"workbook_file_read","read_only":true,"destructive":false,"idempotent":true}`
  and `{"tool":"data_export","action":"data_export","read_only":false,"destructive":false,"idempotent":true}`.
- `tool_annotations`: append `{"tool":"workbook_file_read","read_only":true,"destructive":false,"idempotent":true}`
  and `{"tool":"data_export","read_only":false,"destructive":false,"idempotent":true}`.

Update the frozen-count tests to the new intended totals (authorized files: `test_wp0_baseline.py`,
`test_annotation_matrix.py`, `test_description_contract.py`, **and** `test_mapping_and_approval.py`):
92→94 and 41→43 wherever hard-coded; `serialized_bytes`→the recomputed value. Do NOT weaken any
invariant (per-tool annotation derivation, dual-hint set == {range_merge, pivot_field}, locked triples,
high-risk dispositions must still hold — the 2 new tools are (T,F,T) and (F,F,T), so they do not enter
the dual-hint set and do not disturb locked triples).

## C. Verify
Full `pytest tests/unit` green (incl. WP0/annotation/description). `ruff`/`mypy` clean on changed src
(6 pre-existing ranges_adapter CF errors excepted). Then the Validator runs live MCP tool-call tests
for `workbook_file_read` + `data_export` + paged `range_read`/`pivot_read result-read` on the real
large workbooks (dev-only, originals unchanged).

## D. Documentation
Update `docs/` tool reference + limitations with the two new tools and the large-data limits
(`LargeDataConfig`), and note the WP0 baseline was intentionally extended (92→94) for justified
large-data tools.