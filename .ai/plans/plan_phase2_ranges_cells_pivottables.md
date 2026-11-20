# Excel VBA MCP Server — Phase 2 Implementation Plan
## Scope: Ranges & Cells + PivotTables

**Plan type:** Incremental child plan for the existing `plan.md`  
**Parent plan:** `plan.md` (tasks `task_001`–`task_037`)  
**Normative background requirement:** `plan_addendum_background_execution.md`  
**Specification:** `excel_vba_mcp_technical_spec.docx`  
**New task ID range:** `task_038`–`task_054`

---

## 1. Objective and boundaries

Implement the next Phase 2 increment of the Excel VBA MCP Server: **Ranges & Cells** and **PivotTables**. Extend the existing application/service, Excel adapter, domain DTO, MCP API, workbook model, and bounded test-runner patterns. This plan does not redesign the architecture and does not replace or modify the parent plan.

The work starts only after the parent Stage 1 release gate, `task_037`, is `DONE`. The completed Stage 1 services remain the source of truth for workbook sessions, snapshots, policies, Tables, Names, workbook model, recovery, and test execution.

### Included scope

**Ranges & Cells**
- Resolve explicit worksheet/range targets and return range metadata.
- Read and write cell values and formulas.
- Read and write the supported formatting properties agreed through the research gate.
- Copy, insert, delete, merge, and unmerge ranges.
- Read and manage data validation rules.
- Read and manage conditional formatting rules.
- Support bounded/bulk requests and typed values/errors.

**PivotTables**
- Inventory PivotTables on worksheets.
- Read PivotTable, PivotCache, and PivotField metadata.
- Read field layout, aggregation, and current filter state.
- Refresh supported PivotTables/PivotCaches.
- Change field layout/position/aggregation.
- Apply and clear supported filters.
- Extend the existing workbook model with PivotTable metadata.

### Explicitly out of scope

- Charts, ChartObjects, ChartSheets content management, and Slicers/SlicerCaches.
- Workbook Connections, QueryTables, Power Query, and connection management.
- Data Model / Power Pivot object management. OLAP/Data Model-specific field mutation and refresh are not promised by this increment; expose an explicit unsupported/read-only status where detection is available.
- Creating or deleting PivotTable reports or PivotCaches. The approved specification defines inventory and management of existing tables, fields, caches, filters, and refresh, but does not explicitly require PivotTable creation/deletion.
- New Excel runtime, new COM worker, alternate session, container, VM, separate Windows user/session, Windows Service, or a second UI automation stack.
- Any operation that requires showing Excel/VBE, activating a window, taking foreground focus, clicking, typing, or using `SendKeys`.
- Arbitrary Python, VBA, shell, or formula evaluation on behalf of the MCP caller.

### Inherited execution and safety constraints

- All Excel COM objects are created, accessed, and released only through the existing single STA worker (`com_worker.py`). COM proxies never cross the worker boundary; only serializable DTOs leave it.
- All Excel operations run in the current interactive Windows user session through an automation-owned hidden Excel instance. No isolated environment or VM is permitted.
- Excel and VBE windows must remain hidden and must not steal foreground focus. No `.Activate`, `.Select`, visible navigation, keyboard/mouse injection, coordinate clicking, or blind dialog handling.
- If an unexpected modal window blocks an operation, return a truthful blocked/indeterminate status, stop issuing further COM operations, and use the existing recovery workflow. Never surface a window to continue.
- Temporary Excel application settings must be scoped and restored on success, failure, timeout, and cancellation. Do not use `DisplayAlerts=False` globally.
- Mutations operate only on approved working copies. Create/identify the applicable baseline snapshot before workbook mutations; use the existing approval, audit, lock, recovery, and artifact mechanisms.
- Mutating requests must have explicit target identities and bounded payloads. Ambiguous targets are rejected with a preview/diagnostic rather than guessed.
- Excel Table structure remains owned by `TablesService` / `ListObject` operations. Range structural operations that intersect a ListObject must be rejected or routed through the existing table operations; they must not silently convert or corrupt an Excel Table.
- No implementation may claim success when Excel returns an ambiguous or unsupported result.

---

## 2. Execution DAG

### Groups and ordering

| Group | Tasks | Dependency / gate | Parallelism |
|---|---|---|---|
| G11 — Research and contract decisions | `task_038`, `task_039`, then `task_040` | Stage 1 `task_037` must be `DONE`; freeze DTO/tool contracts only after both research outputs and architecture review | Range and Pivot research can run in parallel |
| G12 — Test fixtures | `task_041` | After the relevant research findings; fixtures must represent the selected supported scope | Fixture preparation can run alongside contract implementation after research decisions |
| G13 — Ranges & Cells | `task_042` → `task_043`, with `task_044`, `task_045`, `task_046` depending on shared range contracts/resolver; then `task_047` | Range contracts, fixtures, existing STA/session/snapshot/policy/Table services | Code work can be split by isolated modules; all Excel integration runs are serialized through the same STA worker |
| G14 — PivotTables | `task_048` → `task_049`, `task_050`, `task_051`; then `task_052` | Pivot contracts, fixtures, existing workbook/session/snapshot/policy/recovery services | Inventory can be developed while range work proceeds; Pivot COM tests remain serialized |
| G15 — Test-runner integration and release | `task_053` → `task_054` | Both feature verification tasks passed | Runner schema work may be prepared earlier, but integrated execution waits for both feature branches |

### Dependency graph

```text
task_037 (Stage 1 complete)
   ├── task_038 (Range research) ─┐
   └── task_039 (Pivot research) ├── task_040 (Phase 2 DTO/tool contracts)
                                 └── task_041 (test fixtures)

Range branch:
task_040 + task_041 + Stage 1 services
   └── task_042 (resolve/read)
         ├── task_043 (write values/formulas)
         ├── task_044 (formatting)
         ├── task_045 (copy/insert/delete/merge)
         └── task_046 (data validation/conditional formatting)
               └── all four implementation tasks ──> task_047 (Ranges & Cells verification)

Pivot branch:
task_040 + task_041 + Stage 1 services
   └── task_048 (PivotTables/PivotCaches/PivotFields inventory)
         ├── task_049 (refresh/cache operations)
         ├── task_050 (field configuration)
         └── task_051 (filters; also depends on task_050)
               └── all three implementation tasks ──> task_052 (PivotTables verification)

Both branches:
task_047 + task_052
   └── task_053 (bounded TestRunner extension)
         └── task_054 (cross-feature/background regression and release gate)
```

### Critical execution rules

- `task_038` and `task_039` are decision gates, not permission to change the approved architecture. Their outputs must compare viable Excel Object Model approaches and document a bounded recommendation.
- `task_040` freezes the shared serializable contracts before feature implementation begins. Range and Pivot implementation must not independently invent overlapping DTOs or status conventions.
- The two feature branches may be developed in parallel only when their file ownership is isolated. Shared files such as `domain/models.py`, `workbook_model_service.py`, `server.py`, and `test_spec.py` must not be edited concurrently by different Implementors.
- One Excel instance / one STA worker is used for all integration tests. Parallel test execution against that instance is prohibited.
- No dependent task may proceed past a failed Research, Architecture, Test, or Validation gate.

---

## 3. Task specifications

### task_038 — Research Range and Cells Object Model behavior

```yaml
id: task_038
title: Research Range and Cells API semantics and compatibility
priority: high
depends_on: [task_037]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — Full Range/Cells API"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Investigate the Excel Range Object Model behavior required by the approved Phase 2 scope.
  Do not modify code or choose an implementation path before documenting the relevant alternatives.
goals:
  - compare Range.Value2 and Value serialization semantics, including dates, empty cells, booleans, numbers, and Excel cell errors
  - compare Formula and Formula2 compatibility and dynamic-array behavior for the supported Excel build
  - recommend bounded bulk read/write and response-size behavior
  - define the supported formatting property contract without expanding beyond the specification
  - investigate Copy, Insert, Delete, Merge, and UnMerge semantics and safe preconditions
  - investigate the supported DataValidation and FormatConditions operations
  - identify multi-area, merged-cell, spill-range, protected-sheet, and ListObject interaction constraints
  - document where an operation must be rejected rather than guessed
extra_context: |
  The parent specification includes values, formulas, formatting, copy/insert/delete/merge operations,
  data validation, and conditional formatting. The Stage 1 range support was intentionally limited to
  metadata and read-only assertions. This task researches only the additional Phase 2 range API.
  Value2 returns Currency/Date-formatted values as floating-point values rather than Currency/Date
  Variants. Formula2 supersedes Formula for dynamic-array-enabled Excel but Formula remains for
  backward compatibility. Date conversion must not be guessed from a raw numeric value alone.
  Range structural operations must not mutate ListObject structure; table row/column changes remain
  the responsibility of TablesService. Preserve the hidden same-session execution contract.
files:
  - path: docs/research/ranges_cells_api.md
    targets:
      - {type: document, name: Range and Cells API decision record}
      - {type: section, name: Value2 and formula representation}
      - {type: section, name: bounded formatting schema}
      - {type: section, name: structural operation preconditions}
      - {type: section, name: validation and conditional-formatting support}
changes:
  - action: research_only
    description: Compare documented Excel Object Model alternatives; do not change source code.
acceptance_criteria:
  - findings cite official Microsoft Excel Object Model documentation
  - alternatives and compatibility constraints are explicit
  - the recommendation preserves existing service/adapter/API layering
  - date and cell-error serialization are documented without lossy assumptions
  - interactions with ListObjects, merged cells, spill ranges, and protected sheets are documented
  - no unsupported requirement is added to the feature scope
tests_to_run: ["source and API review; no code tests"]
agents:
  - researcher: required
  - auditor: review recommended for architecture compatibility
  - validator: review the decision record before task_040
verification_mode: dedicated_gate
```

### task_039 — Research PivotTables, PivotCaches, and PivotFields behavior

```yaml
id: task_039
title: Research PivotTables object model, refresh semantics, and supported field operations
priority: critical
depends_on: [task_037]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — PivotTables/PivotCaches/PivotFields"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Investigate the supported object-model paths for discovering and managing existing PivotTables.
  Resolve the major compatibility and shared-cache uncertainties before implementing the service.
goals:
  - compare worksheet-level PivotTables enumeration with workbook-level discovery and define complete inventory behavior
  - determine the metadata required for PivotTable, PivotCache, and PivotField DTOs
  - investigate shared PivotCache relationships and the effects of refreshing a cache used by multiple tables
  - compare PivotTable.RefreshTable and PivotCache.Refresh for the approved local-source scope
  - catalogue the field-level configuration surface required by the specification's full-field-settings requirement, including Orientation, Position, data-field aggregation, and other supported field properties
  - investigate adding/removing fields from layout areas within existing PivotTables, distinguishing field metadata changes from PivotTable creation/deletion
  - enumerate supported manual, label, value, and date filter operations and their limitations
  - identify regular worksheet/table-backed caches versus external, OLAP, and Data Model caches
  - define explicit unsupported outcomes for features owned by deferred Connections/Power Query and Data Model phases
extra_context: |
  The approved scope is inventory and management of existing PivotTables, PivotCaches, PivotFields,
  filters, and refresh. Data Model/Power Pivot, Connections, and Power Query are separate deferred
  feature areas. Do not silently expand into them. Workbook-level PivotTables enumeration may not
  provide a complete inventory; verify worksheet-level enumeration. Field orientation and filters
  can behave differently for OLAP hierarchies. Shared caches can make cache-level operations affect
  multiple PivotTables. Refresh must be targeted; do not use Workbook.RefreshAll.
  Preserve hidden same-session execution: no UI activation, visible workbook windows, or modal-dialog
  interaction. If a cache requires an unsupported external/data-model action, return an explicit
  unsupported/blocked result rather than invoking another deferred feature.
files:
  - path: docs/research/pivot_tables_api.md
    targets:
      - {type: document, name: PivotTables API decision record}
      - {type: section, name: inventory and cache discovery}
      - {type: section, name: refresh and shared-cache effects}
      - {type: section, name: field layout and aggregation}
      - {type: section, name: filter support matrix}
      - {type: section, name: OLAP/external source limitations}
changes:
  - action: research_only
    description: Compare official Excel Object Model options; do not change source code.
acceptance_criteria:
  - findings cite official Microsoft Excel Object Model documentation
  - inventory method covers all supported worksheet PivotTables
  - refresh alternatives and shared-cache effects are explicit
  - supported field and filter operation matrix is documented
  - Data Model/OLAP and external-source limitations are explicit
  - no PivotTable creation/deletion or deferred feature scope is added implicitly
  - recommendation preserves the existing COM STA worker and application/service/adapter/API layers
agents:
  - researcher: required
  - auditor: required
  - validator: required_after_research_and_auditor_review
verification_mode: dedicated_gate
```

### task_040 — Define Phase 2 Range and Pivot DTOs and tool contracts

```yaml
id: task_040
title: Define serializable Range and PivotTables domain contracts
priority: high
depends_on: [task_038, task_039]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.2. Component responsibilities; 6. MCP contract; 14. Phase 2"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Extend the existing domain models, typed errors, statuses, and MCP tool schemas for Ranges & Cells
  and PivotTables using the two approved research decision records.
goals:
  - define typed RangeAddress, RangeMetadata, cell-value/error, formula, formatting, validation, and conditional-format DTOs
  - define PivotTable, PivotCache, PivotField, field-layout, aggregation, filter, refresh-result, and unsupported-source DTOs
  - define bounded request schemas for range sizes, bulk payloads, filters, and field updates
  - define explicit error/status values for invalid range, protected target, unsupported pivot source, stale target, ambiguous filter, refresh failure, and modal/UI-blocked outcomes where needed
  - define structured outputs for all new MCP tools
  - preserve the existing JSON/Pydantic conventions and DTO-only boundary
extra_context: |
  Extend existing domain modules; do not introduce a second model or error system.
  No COM object, pywin32 proxy, Excel enum object, or arbitrary Python object may escape the STA worker.
  Keep Phase 2 content out of the Stage 1 workbook model except for the PivotTable metadata explicitly
  added by task_048. Never return an unbounded whole-sheet cell matrix.
files:
  - path: src/excel_vba_mcp/domain/models.py
    targets:
      - {type: model_group, name: Range and PivotTables DTOs}
  - path: src/excel_vba_mcp/domain/errors.py
    targets:
      - {type: module, name: existing typed Excel errors}
  - path: src/excel_vba_mcp/domain/statuses.py
    targets:
      - {type: module, name: existing operation statuses}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: tool_schema_group, name: Range and Cells tool contracts}
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - {type: tool_schema_group, name: PivotTables tool contracts}
changes:
  - action: extend_existing_contracts
    description: Add only the DTOs, validators, and structured tool schemas required by tasks_042–task_051.
acceptance_criteria:
  - all DTOs are JSON serializable and contain no COM references
  - bulk range requests have enforceable row/cell/byte limits
  - values and cell errors are represented without lossy string coercion
  - PivotTable and cache identity/relationship fields are stable within a workbook session
  - unsupported PivotTable source types have explicit statuses
  - existing Stage 1 DTOs and MCP tool contracts remain backward compatible
tests_to_run: ["unit tests for DTO validation, serialization, bounds, and backward compatibility"]
agents:
  - auditor: required
  - validator: review shared contracts before implementation
verification_mode: dedicated_gate
```

### task_041 — Prepare Range and PivotTables integration fixtures

```yaml
id: task_041
title: Prepare bounded Phase 2 Range and PivotTables test fixtures
priority: high
depends_on: [task_038, task_039]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "12.1. Test fixtures; 14. Phase 2"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add deterministic workbook fixtures that exercise the researched supported scope without external
  services, UI interaction, or changes to global Office settings.
goals:
  - prepare RangesCellsProject.xlsm with constants, formulas, date-like values, Excel errors, formatting, merged cells, data validation, and conditional formatting
  - include ranges inside and outside a ListObject so Table-structure guards can be tested
  - prepare PivotTablesProject.xlsm with a local worksheet/table-backed cache and representative row, column, page, and data fields
  - include multiple PivotTables and a shared-cache case if supported by the target Excel build
  - include filters, aggregation settings, and a deterministic refresh source
  - record fixture manifest, expected properties, source hashes, and target Excel build
  - optionally include unsupported OLAP/Data Model metadata fixture only when available without adding a Data Model implementation dependency
extra_context: |
  Fixtures must be reproducible and must not rely on network connections, external databases,
  Power Query, or user-specific file paths. Create/read them through the hidden Excel instance in the
  current interactive Windows session. Do not open visible Excel/VBE windows or use UI automation.
files:
  - path: tests/fixtures/RangesCellsProject.xlsm
    targets: [{type: fixture, name: Range and Cells behavior workbook}]
  - path: tests/fixtures/PivotTablesProject.xlsm
    targets: [{type: fixture, name: PivotTables behavior workbook}]
  - path: tests/fixtures/phase2_manifest.json
    targets: [{type: manifest, name: Phase 2 fixture manifest}]
changes:
  - action: add_test_fixtures
    description: Create deterministic workbook fixtures and machine-readable expected metadata.
acceptance_criteria:
  - fixtures open and close through the existing hidden Excel COM session
  - fixture creation is reproducible and independent of external services
  - expected values, formulas, validation, formats, pivot layout, filters, and cache sharing are documented
  - no visible window or foreground-focus change occurs during fixture creation
  - fixture hashes and target Excel build are recorded
tests_to_run: ["fixture smoke test using the existing STA worker"]
agents:
  - auditor: review fixture coverage
  - validator: verify fixture determinism and background behavior
verification_mode: dedicated_verification
```

### task_042 — Implement range resolution, metadata, and read operations

```yaml
id: task_042
title: Implement Range resolution and bounded read operations
priority: high
depends_on: [task_040, task_041, task_012, task_010]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.2–5.3. Component responsibilities and COM boundary; 14. Phase 2 — Range/Cells"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add the read-only Range/Cells path using explicit workbook, worksheet, and range identifiers.
  Reuse the existing Excel session and STA worker.
goals:
  - implement range_get metadata resolution for explicit worksheet and A1-style range addresses
  - implement range_read_values and range_read_formulas using the researched property semantics
  - return address, dimensions, cell count, values/formulas, and supported value/error typing
  - support bounded bulk reads and response-size limits
  - normalize single-cell and multi-cell return shapes consistently
  - reject invalid, ambiguous, oversized, or unsupported range references with typed errors
  - avoid reading the entire UsedRange unless explicitly requested within configured limits
extra_context: |
  All object resolution occurs inside the existing Excel STA worker. Use direct object references;
  do not use Activate, Select, ActiveCell, or visible navigation. Preserve the background contract.
  Date-like serial values must follow the task_038 decision record; do not guess dates from numeric
  values. Excel cell errors must remain typed. This task is read-only and does not add writes.
files:
  - path: src/excel_vba_mcp/application/ranges_service.py
    targets:
      - {type: class, name: RangesService}
      - {type: method, name: get_range}
      - {type: method, name: read_values}
      - {type: method, name: read_formulas}
  - path: src/excel_vba_mcp/adapters/excel/ranges_adapter.py
    targets:
      - {type: class, name: RangesAdapter}
      - {type: method, name: resolve_range}
      - {type: method, name: read_values}
      - {type: method, name: read_formulas}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: function, name: range_get}
      - {type: function, name: range_read_values}
      - {type: function, name: range_read_formulas}
  - path: src/excel_vba_mcp/server.py
    targets: [{type: registration, name: Range read tools}]
changes:
  - action: add_read_operations
    description: Implement bounded Range metadata/value/formula reads through the existing layered architecture.
acceptance_criteria:
  - one-cell and multi-cell ranges return consistent dimensions and typed values
  - formulas are returned according to the task_038 decision record
  - read requests enforce maximum cell and response limits
  - invalid sheet/range references return structured errors
  - no COM proxy leaves the STA worker
  - no Excel/VBE window appears or takes foreground focus
  - existing TablesService and NamesService behavior is unchanged
tests_to_run: ["unit tests for range DTO normalization and limits", "integration tests for range reads on RangesCellsProject.xlsm"]
agents:
  - auditor: recommended
  - validator: recommended for read boundary and data typing
verification_mode: inline
```

### task_043 — Implement range value and formula mutations

```yaml
id: task_043
title: Implement snapshot-backed cell value and formula writes
priority: critical
depends_on: [task_042, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "9. Snapshot model; 10. Security; 14. Phase 2 — Range/Cells values and formulas"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add explicit value-write and formula-write operations with preconditions, snapshots, approval,
  bounded payloads, and postcondition checks.
goals:
  - implement range_write_values for scalar and two-dimensional bulk data
  - implement range_write_formulas using the property selected by task_038
  - validate target dimensions and payload shape before mutating Excel
  - require a working copy, applicable baseline snapshot, and existing approval policy
  - return before/after metadata and a structured operation result
  - verify written values/formulas after assignment
  - preserve typed dates, numbers, booleans, blanks, and Excel errors according to the approved DTO contract
extra_context: |
  A string beginning with '=' is not implicitly interpreted as a formula. Value writes and formula
  writes are separate explicit operations. No formula string is evaluated by Python. Range writes
  that target cells within a ListObject are allowed only as explicit cell-level updates with snapshot
  and postconditions; they must not alter the ListObject row/column structure. If the target is
  ambiguous, protected, read-only, or outside the allowlist, reject it.
files:
  - path: src/excel_vba_mcp/application/ranges_service.py
    targets:
      - {type: method, name: write_values}
      - {type: method, name: write_formulas}
  - path: src/excel_vba_mcp/adapters/excel/ranges_adapter.py
    targets:
      - {type: method, name: write_values}
      - {type: method, name: write_formulas}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: function, name: range_write_values}
      - {type: function, name: range_write_formulas}
changes:
  - action: add_snapshot_backed_mutations
    description: Add explicit cell value and formula mutation operations with policy and postconditions.
acceptance_criteria:
  - wrong-sized payloads are rejected before any workbook mutation
  - writes require the existing snapshot/approval/audit path
  - formula strings are never inferred from ordinary value strings
  - returned postconditions match actual Excel values/formulas
  - a failed/partial COM operation is reported truthfully and enters existing recovery handling
  - no visible UI, focus change, or separate execution environment is used
tests_to_run: ["unit tests for payload validation", "integration tests for value/formula writes, rollback evidence, and table-boundary cases"]
agents:
  - auditor: required; review mutation and table-boundary safety
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_044 — Implement range formatting read/write

```yaml
id: task_044
title: Implement bounded Range formatting inspection and mutation
priority: high
depends_on: [task_040, task_042, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — Range/Cells formatting"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add the bounded formatting operations approved in task_038, using typed formatting DTOs and
  explicit property updates rather than copying arbitrary COM objects.
goals:
  - implement range_get_format for supported formatting metadata
  - implement range_set_format for supported formatting properties
  - support partial updates without overwriting unspecified formatting properties
  - validate supported enum/string/numeric values before COM mutation
  - require snapshot and existing approval/audit policy for formatting writes
  - verify the targeted formatting properties after mutation
extra_context: |
  Do not implement arbitrary serialization of Excel Style, Font, Interior, Borders, or other COM
  objects. Expose only the bounded property schema documented by task_038. Keep unsupported or
  mixed-format properties explicit in the result; do not silently flatten them. No UI interaction
  or window activation is permitted.
files:
  - path: src/excel_vba_mcp/application/ranges_service.py
    targets:
      - {type: method, name: get_format}
      - {type: method, name: set_format}
  - path: src/excel_vba_mcp/adapters/excel/ranges_adapter.py
    targets:
      - {type: method, name: get_format}
      - {type: method, name: set_format}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: function, name: range_get_format}
      - {type: function, name: range_set_format}
changes:
  - action: add_bounded_format_operations
    description: Implement only the format fields accepted by the research decision record.
acceptance_criteria:
  - formatting DTO contains no COM objects and is JSON serializable
  - partial updates do not reset unspecified properties
  - mixed or unsupported formatting is reported explicitly
  - formatting mutations use snapshots, approval, audit, and postcondition checks
  - hidden Excel remains non-activating throughout the operation
tests_to_run: ["unit tests for formatting DTO constraints", "integration tests for supported formatting fields"]
agents:
  - auditor: recommended
  - validator: recommended
verification_mode: inline
```

### task_045 — Implement range copy, insert, delete, merge, and unmerge

```yaml
id: task_045
title: Implement guarded Range copy and structural operations
priority: critical
depends_on: [task_042, task_043, task_014, task_009, task_026]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — Range/Cells copy, insert, delete, and merge"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Implement explicit same-workbook copy and guarded structural range operations. Preserve existing
  TablesService ownership of ListObject structure.
goals:
  - implement range_copy with explicit source and destination ranges
  - implement range_insert and range_delete with explicit shift direction and target
  - implement range_merge and range_unmerge with explicit target ranges
  - require snapshot, approval, and audit for destructive structural operations
  - preflight merged cells, non-empty cells, affected formulas, and range/table intersections as supported by task_038
  - verify resulting addresses, dimensions, and affected table metadata
extra_context: |
  Use direct Range object operations and destination arguments; do not use clipboard/UI navigation.
  Structural operations that intersect or alter a ListObject are rejected and direct callers to the
  existing TablesService ListRows/ListColumns operations. Range writes to explicit cells inside a
  table are a separate cell-level capability. Ambiguous insert/delete shift behavior must be rejected.
  Merging cells can discard non-anchor values; require explicit user intent, snapshot, approval, and
  a preflight that reports the affected cells. No operation may show a dialog or take desktop focus.
files:
  - path: src/excel_vba_mcp/application/ranges_service.py
    targets:
      - {type: method, name: copy_range}
      - {type: method, name: insert_range}
      - {type: method, name: delete_range}
      - {type: method, name: merge_range}
      - {type: method, name: unmerge_range}
  - path: src/excel_vba_mcp/adapters/excel/ranges_adapter.py
    targets:
      - {type: method, name: copy_range}
      - {type: method, name: insert_range}
      - {type: method, name: delete_range}
      - {type: method, name: merge_range}
      - {type: method, name: unmerge_range}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: function, name: range_copy}
      - {type: function, name: range_insert}
      - {type: function, name: range_delete}
      - {type: function, name: range_merge}
      - {type: function, name: range_unmerge}
changes:
  - action: add_guarded_structure_operations
    description: Add structural Range operations with ListObject guards and verified postconditions.
acceptance_criteria:
  - operation targets and shifts are explicit and validated before mutation
  - ListObject structure is never modified through Range.Insert/Delete/Merge
  - merge with non-anchor data is blocked unless explicitly approved after preview
  - all structural mutations create/use a baseline snapshot and record audit evidence
  - copy uses explicit destination and does not depend on clipboard or active selection
  - failed operations return truthful status and trigger existing recovery handling
tests_to_run: ["integration tests for copy/insert/delete/merge/unmerge", "negative tests for ListObject and merged-cell intersections"]
agents:
  - researcher: consult if task_038 left unresolved Excel-version behavior
  - auditor: required
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_046 — Implement data validation and conditional formatting

```yaml
id: task_046
title: Implement Range data validation and conditional formatting operations
priority: high
depends_on: [task_040, task_042, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — Range/Cells data validation and conditional formatting"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add bounded inspection and management of Range.Validation and Range.FormatConditions through
  the existing RangesService and RangesAdapter.
goals:
  - implement range_get_validation and range_set_validation
  - implement range_get_conditional_formatting and range_set_conditional_formatting
  - support the rule types and properties accepted by task_038
  - preserve rule order/priority and distinguish add, modify, and delete behavior
  - validate formulas and operands as data passed to Excel; never evaluate them in Python
  - require snapshot, approval, audit, and postcondition verification for mutations
extra_context: |
  Use the Excel Validation object and FormatConditions collection through the existing COM adapter.
  Do not create a generic arbitrary COM property editor. Unsupported rule types, mixed rules, and
  invalid formulas return explicit structured diagnostics. Respect Excel's rule-specific parameter
  constraints and string length limits. Keep all operations hidden and non-activating.
files:
  - path: src/excel_vba_mcp/application/ranges_service.py
    targets:
      - {type: method, name: get_validation}
      - {type: method, name: set_validation}
      - {type: method, name: get_conditional_formatting}
      - {type: method, name: set_conditional_formatting}
  - path: src/excel_vba_mcp/adapters/excel/ranges_adapter.py
    targets:
      - {type: method, name: get_validation}
      - {type: method, name: set_validation}
      - {type: method, name: get_conditional_formatting}
      - {type: method, name: set_conditional_formatting}
  - path: src/excel_vba_mcp/api/range_tools.py
    targets:
      - {type: function, name: range_get_validation}
      - {type: function, name: range_set_validation}
      - {type: function, name: range_get_conditional_formatting}
      - {type: function, name: range_set_conditional_formatting}
changes:
  - action: add_validation_rule_operations
    description: Add bounded data-validation and conditional-formatting support using research-approved rule types.
acceptance_criteria:
  - supported rule properties serialize to JSON and round-trip through Excel
  - unsupported or mixed rules are represented explicitly, not silently omitted
  - rule order/priority remains correct after updates
  - all mutations are snapshot-backed, approved, audited, and postcondition-checked
  - no visible windows or foreground activation occur
tests_to_run: ["unit tests for rule schemas", "integration tests for validation and conditional formatting fixtures"]
agents:
  - auditor: recommended
  - validator: recommended
verification_mode: inline
```

### task_047 — Verify the complete Ranges & Cells workflow

```yaml
id: task_047
title: Independently verify Ranges and Cells behavior and safety
priority: critical
depends_on: [task_042, task_043, task_044, task_045, task_046]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "11. Reliability; 12. Quality and acceptance; 14. Phase 2 — Range/Cells"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Independently verify the Ranges & Cells API against real Excel fixtures, including read/write,
  formula, formatting, structural, validation, conditional-formatting, and background constraints.
goals:
  - verify value and formula round-trips including errors, blanks, dates, and two-dimensional data
  - verify formatting read/write and mixed-format reporting
  - verify copy, insert, delete, merge, and unmerge postconditions
  - verify data validation and conditional-formatting rule behavior
  - verify snapshot, approval, audit, and recovery on failed mutations
  - verify ListObject structural guards and preservation of Stage 1 TablesService behavior
  - verify response-size limits and invalid/oversized request handling
  - verify no focus theft, visible Excel/VBE window, or UI automation
extra_context: |
  Use RangesCellsProject.xlsm and the existing Stage 1 Table/Names fixtures. Integration tests run
  serially through the single STA worker in the current interactive session. Do not use a VM,
  container, alternate Windows session, visible UI, or mock-only validation for Excel behavior.
files:
  - path: tests/integration_excel/test_ranges_service.py
    targets: [{type: test_module, name: Ranges & Cells integration verification}]
  - path: tests/integration_excel/test_range_background_behavior.py
    targets: [{type: test_module, name: Range desktop non-interference checks}]
changes:
  - action: add_independent_verification
    description: Verify full range workflow and inherited background execution requirements.
acceptance_criteria:
  - all supported Range operations pass real Excel fixture tests
  - destructive mutations require snapshot/approval and verify postconditions
  - no mutation is performed when ListObject structure would be changed through Range APIs
  - serialization and response limits are enforced
  - no Excel/VBE window appears or steals foreground focus
  - Validator returns PASS or PASS WITH WARNINGS with no critical findings
tests_to_run: ["test_ranges_service.py", "test_range_background_behavior.py"]
agents:
  - auditor: required for failed cases and architecture findings
  - validator: required independent verification
verification_mode: dedicated_gate
```

### task_048 — Implement PivotTables, PivotCaches, and PivotFields inventory

```yaml
id: task_048
title: Implement PivotTables inventory and metadata inspection
priority: critical
depends_on: [task_040, task_041, task_012, task_032, task_034]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — PivotTables/PivotCaches/PivotFields"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Implement read-only inventory and metadata inspection for existing worksheet PivotTables,
  their PivotCaches, and their PivotFields.
goals:
  - enumerate PivotTables from each Worksheet.PivotTables collection
  - implement pivot_tables_list and pivot_table_get
  - implement pivot_cache_get and pivot_fields_list
  - return PivotTable identity, worksheet, TableRange1/TableRange2 addresses, cache identity, and supported state metadata
  - return cache source type, source metadata, OLAP indicator, refresh capability, and shared-table relationships where available
  - return each field's source name, caption, orientation, position, data-field function, and filter state where available
  - extend WorkbookModelService.get_sheet_model to include PivotTable inventory metadata without reading every cell
  - return explicit unsupported/read-only statuses for deferred cache/source types
extra_context: |
  Use worksheet-level PivotTables enumeration to avoid an incomplete workbook-level inventory.
  Inventory must be read-only. Do not create/delete PivotTables, edit source ranges, refresh caches,
  access Power Pivot/Data Model internals, or invoke external connections in this task. Do not return
  COM references. Preserve hidden same-session operation and do not use Select/Activate.
files:
  - path: src/excel_vba_mcp/application/pivot_tables_service.py
    targets:
      - {type: class, name: PivotTablesService}
      - {type: method, name: list_pivot_tables}
      - {type: method, name: get_pivot_table}
      - {type: method, name: get_pivot_cache}
      - {type: method, name: list_pivot_fields}
  - path: src/excel_vba_mcp/adapters/excel/pivot_tables_adapter.py
    targets:
      - {type: class, name: PivotTablesAdapter}
      - {type: method, name: enumerate_pivot_tables}
      - {type: method, name: read_pivot_cache}
      - {type: method, name: enumerate_pivot_fields}
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - {type: function, name: pivot_tables_list}
      - {type: function, name: pivot_table_get}
      - {type: function, name: pivot_cache_get}
      - {type: function, name: pivot_fields_list}
  - path: src/excel_vba_mcp/application/workbook_model_service.py
    targets:
      - {type: method, name: get_sheet_model}
  - path: src/excel_vba_mcp/server.py
    targets: [{type: registration, name: PivotTables inventory tools}]
changes:
  - action: add_pivot_inventory
    description: Add read-only PivotTable inventory and connect it to the existing sheet/workbook model aggregation.
acceptance_criteria:
  - inventory finds all PivotTables on supported worksheets
  - multiple PivotTables sharing a cache are represented with a common cache identity
  - cache and field metadata are serializable and contain no COM objects
  - regular range/table-backed caches are identified correctly
  - unsupported OLAP/Data Model/external sources return explicit statuses and do not trigger refresh
  - workbook model adds only PivotTable metadata; it does not materialize the full PivotTable cell output
  - all operations run hidden through the existing STA worker without focus changes
tests_to_run: ["unit tests for pivot DTOs", "integration tests for inventory and workbook model aggregation"]
agents:
  - auditor: required; verify inventory completeness and model consistency
  - validator: recommended
verification_mode: inline
```

### task_049 — Implement targeted PivotTable and PivotCache refresh

```yaml
id: task_049
title: Implement policy-controlled PivotTable and PivotCache refresh
priority: critical
depends_on: [task_048, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — PivotCaches and refresh"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Add targeted refresh operations for the cache/source types approved by task_039, with snapshot,
  approval, audit, timeout, and postcondition reporting.
goals:
  - implement pivot_table_refresh for an explicitly identified PivotTable
  - implement pivot_cache_refresh for an explicitly identified supported cache when justified by the research decision
  - report all PivotTables affected by a shared-cache refresh
  - capture start/end time, result, error details, refresh metadata, and resulting PivotTable ranges
  - reject unsupported external, OLAP, Data Model, or connection-backed refresh paths
  - avoid refreshing unrelated workbook connections or queries
extra_context: |
  Research must choose and document the appropriate distinction between PivotTable.RefreshTable
  and PivotCache.Refresh. Do not use Workbook.RefreshAll. Cache refresh may affect multiple PivotTables
  sharing the cache; callers must receive the affected-object list and must pass the existing policy
  and snapshot requirements. External/Power Query/Data Model refresh remains out of scope. If Excel
  blocks on an unexpected modal dialog or cannot complete without UI, return a truthful blocked or
  indeterminate status and do not surface the window.
files:
  - path: src/excel_vba_mcp/application/pivot_tables_service.py
    targets:
      - {type: method, name: refresh_pivot_table}
      - {type: method, name: refresh_pivot_cache}
  - path: src/excel_vba_mcp/adapters/excel/pivot_tables_adapter.py
    targets:
      - {type: method, name: refresh_pivot_table}
      - {type: method, name: refresh_pivot_cache}
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - {type: function, name: pivot_table_refresh}
      - {type: function, name: pivot_cache_refresh}
changes:
  - action: add_targeted_refresh
    description: Implement only the refresh operations validated by task_039 for supported local caches.
acceptance_criteria:
  - refresh target and cache source type are checked before invocation
  - shared-cache impact is reported before and after the operation
  - snapshot, approval, and audit records are present
  - refresh result is based on Excel's actual return/result state, not absence of a COM exception alone
  - Workbook.RefreshAll and external/data-model refresh are never invoked
  - no Excel/VBE window appears or takes foreground focus
  - unexpected modal/blocked outcomes are represented explicitly and handled by existing recovery
tests_to_run: ["integration tests for local PivotTable and PivotCache refresh", "negative tests for unsupported source types"]
agents:
  - researcher: required if task_039 leaves refresh semantics unresolved
  - auditor: required
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_050 — Implement PivotField layout and aggregation operations

```yaml
id: task_050
title: Implement PivotField layout, ordering, and aggregation changes
priority: critical
depends_on: [task_048, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — PivotFields configuration"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Implement bounded changes to existing PivotField orientation, position, and data-field aggregation
  using the behavior approved in task_039.
goals:
  - implement the complete bounded PivotField configuration matrix approved by task_039
  - include layout/orientation, position, data-field aggregation, and other explicitly approved field properties where supported by the target Excel build
  - validate each property and its prerequisites against the current field type/orientation
  - handle duplicate captions/data-field naming according to the researched identity rules
  - return before/after field configuration and verify the actual PivotTable state
  - require snapshot, approval, audit, and explicit workbook/PivotTable/field identity
extra_context: |
  Do not assume regular PivotField behavior applies to OLAP hierarchies. OLAP/Data Model field
  mutations are outside this increment. Research must establish whether moving a field into the data
  area requires AddDataField or setting Orientation and how the returned data field is identified.
  The implementation matrix must cover the specification's full field-settings scope while remaining
  bounded to existing PivotTables and excluding creation/deletion of PivotTable reports. All operations
  are direct COM calls through the STA worker, with no visible navigation or UI.
files:
  - path: src/excel_vba_mcp/application/pivot_tables_service.py
    targets:
      - {type: method, name: set_pivot_field_layout}
      - {type: method, name: set_pivot_field_position}
      - {type: method, name: set_pivot_field_aggregation}
      - {type: method, name: set_pivot_field_properties}
  - path: src/excel_vba_mcp/adapters/excel/pivot_tables_adapter.py
    targets:
      - {type: method, name: set_field_layout}
      - {type: method, name: set_field_position}
      - {type: method, name: set_field_aggregation}
      - {type: method, name: set_field_properties}
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - {type: function, name: pivot_field_set_layout}
      - {type: function, name: pivot_field_set_position}
      - {type: function, name: pivot_field_set_aggregation}
      - {type: function, name: pivot_field_set_properties}
changes:
  - action: add_field_configuration
    description: Implement the approved existing-field layout and aggregation operations.
acceptance_criteria:
  - every property in the approved PivotField configuration matrix is validated before mutation
  - only valid orientations, positions, and aggregation functions are accepted
  - field positions are validated and resulting order is returned
  - unsupported OLAP/Data Model fields are rejected with explicit status
  - mutations are snapshot-backed, approved, audited, and postcondition-checked
  - partial failures enter existing recovery and do not report success
  - desktop non-interference criteria remain satisfied
tests_to_run: ["unit tests for field operation validation", "integration tests for layout, ordering, and aggregation"]
agents:
  - auditor: required
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_051 — Implement PivotField filters

```yaml
id: task_051
title: Implement PivotField filter inspection, apply, and clear operations
priority: critical
depends_on: [task_048, task_050, task_014, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "14. Phase 2 — PivotFields filters"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Implement the bounded PivotField filter operations agreed through task_039, including inspection,
  applying supported filters, and clearing selected filters.
goals:
  - implement pivot_field_get_filters
  - implement pivot_field_apply_filter for supported manual, label, value, and date filter types selected by research
  - implement pivot_field_clear_filters with explicit field or PivotTable scope
  - validate filter type, target field, data field, operands, and value types before mutation
  - preserve and report filter type, parameters, and resulting field state
  - require snapshot, approval, audit, and postcondition checks for filter mutations
extra_context: |
  Filter API must be bounded; do not accept arbitrary COM filter objects. The research decision record
  must identify which PivotFilters.Add and manual PivotItems operations are supported for the target
  Excel build and local cache types. Do not silently convert an unsupported filter into another type.
  Clear operations must state whether they clear one field or the entire PivotTable. No UI is used.
files:
  - path: src/excel_vba_mcp/application/pivot_tables_service.py
    targets:
      - {type: method, name: get_pivot_field_filters}
      - {type: method, name: apply_pivot_field_filter}
      - {type: method, name: clear_pivot_field_filters}
  - path: src/excel_vba_mcp/adapters/excel/pivot_tables_adapter.py
    targets:
      - {type: method, name: get_field_filters}
      - {type: method, name: apply_field_filter}
      - {type: method, name: clear_field_filters}
  - path: src/excel_vba_mcp/api/pivot_tools.py
    targets:
      - {type: function, name: pivot_field_get_filters}
      - {type: function, name: pivot_field_apply_filter}
      - {type: function, name: pivot_field_clear_filters}
changes:
  - action: add_bounded_pivot_filters
    description: Implement only the supported filter types from the approved research matrix.
acceptance_criteria:
  - filter requests are validated before any mutation
  - applied filters can be read back and match the requested supported filter
  - field-level and PivotTable-level clear behavior are distinct
  - unsupported filter/source combinations return explicit statuses
  - all mutations use snapshots, approvals, audit, and postcondition checks
  - no visible Excel/VBE window or foreground activation occurs
tests_to_run: ["unit tests for filter schemas", "integration tests for manual, label, value, and selected date-filter cases"]
agents:
  - researcher: required if filter matrix remains unresolved after task_039
  - auditor: required
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_052 — Verify the complete PivotTables workflow

```yaml
id: task_052
title: Independently verify PivotTables inventory, refresh, layout, and filters
priority: critical
depends_on: [task_048, task_049, task_050, task_051]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "11. Reliability; 12. Quality and acceptance; 14. Phase 2 — PivotTables"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Independently validate the PivotTables feature against real Excel fixtures, including cache sharing,
  refresh, field configuration, filter behavior, unsupported sources, and background constraints.
goals:
  - verify complete worksheet PivotTables inventory and metadata
  - verify shared PivotCache identity and reported side effects
  - verify targeted local-cache refresh and resulting table output
  - verify field layout, order, aggregation, and filters by reading the state back from Excel
  - verify snapshots, approvals, audit, postconditions, and recovery on failed mutations
  - verify unsupported external/OLAP/Data Model source cases do not invoke deferred APIs
  - verify no visible windows, foreground activation, mouse/keyboard input, or UI automation
extra_context: |
  All integration tests run serially through the existing STA worker in the current interactive
  Windows session. Do not use VM/container/alternate user session. Do not use Workbook.RefreshAll.
  The tests must use local deterministic data sources and must not rely on network connections.
files:
  - path: tests/integration_excel/test_pivot_tables_service.py
    targets: [{type: test_module, name: PivotTables inventory and configuration verification}]
  - path: tests/integration_excel/test_pivot_refresh.py
    targets: [{type: test_module, name: targeted cache/table refresh verification}]
  - path: tests/integration_excel/test_pivot_background_behavior.py
    targets: [{type: test_module, name: Pivot desktop non-interference checks}]
changes:
  - action: add_independent_verification
    description: Verify the full supported PivotTables workflow and inherited background requirements.
acceptance_criteria:
  - all supported PivotTable inventory and mutation operations pass real Excel tests
  - shared-cache side effects are correctly reported
  - unsupported cache/source types are safely rejected
  - mutations are snapshot-backed and auditable
  - no Excel/VBE window appears or steals foreground focus
  - Validator returns PASS or PASS WITH WARNINGS with no critical findings
tests_to_run: ["test_pivot_tables_service.py", "test_pivot_refresh.py", "test_pivot_background_behavior.py"]
agents:
  - auditor: required for failed cases and architecture findings
  - validator: required independent verification
verification_mode: dedicated_gate
```

### task_053 — Extend the bounded TestRunner for Range and Pivot workflows

```yaml
id: task_053
title: Extend declarative workbook tests with Range and Pivot actions/assertions
priority: high
depends_on: [task_047, task_052, task_035, task_036]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "9.2–9.3. Test suite and assertions; 14. Phase 2"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Extend the existing declarative test specification and runner so an agent can verify Range and
  PivotTables behavior as part of repeatable workbook scenarios.
goals:
  - add bounded test actions for supported range reads/writes and the approved Pivot refresh/configuration operations
  - add assertions for cell values, formulas, supported formatting, range structure, validation, and conditional formatting
  - add assertions for PivotTable inventory, field layout/aggregation, filters, refresh status, and output range metadata
  - retain per-test-case baseline and disposable-copy semantics from Stage 1
  - retain structured execution logs, artifacts, and truthful blocked/unsupported statuses
  - reject arbitrary Python, shell, VBA source, and unbounded COM actions from test specifications
extra_context: |
  Extend existing domain/test_spec.py and application/test_runner.py; do not create a parallel test
  runner. Use existing snapshot, artifact, approval, and recovery services. Each test case starts from
  its own baseline copy and is executed sequentially by the single STA worker in the current user
  session. Test-runner actions must call the same application services as the direct MCP tools; do
  not duplicate COM logic in the runner.
files:
  - path: src/excel_vba_mcp/domain/test_spec.py
    targets:
      - {type: schema_group, name: bounded Range and Pivot test actions/assertions}
  - path: src/excel_vba_mcp/application/test_runner.py
    targets:
      - {type: method, name: execute_test_case}
      - {type: method, name: execute_step}
      - {type: method, name: evaluate_assertion}
  - path: src/excel_vba_mcp/api/test_tools.py
    targets:
      - {type: function, name: test_suite_run}
  - path: tests/unit/test_test_spec.py
    targets: [{type: test_module, name: Phase 2 test schema validation}]
  - path: tests/integration_excel/test_test_runner_ranges_pivots.py
    targets: [{type: test_module, name: Range and Pivot end-to-end declarative tests}]
changes:
  - action: extend_existing_test_runner
    description: Add allowlisted Phase 2 actions and assertions through existing service methods.
acceptance_criteria:
  - only documented bounded actions/assertions are accepted
  - all actions reuse the same application services as MCP tools
  - each test case starts from the correct baseline and does not inherit prior case mutations
  - assertions compare actual Excel state, not only requested input values
  - test result logs include operation IDs, statuses, and artifacts
  - no visible Excel/VBE window, focus theft, or separate execution environment occurs
tests_to_run: ["test_test_spec.py", "test_test_runner_ranges_pivots.py"]
agents:
  - auditor: required; verify runner/service separation and bounded schema
  - validator: required after implementation
verification_mode: dedicated_verification
```

### task_054 — Verify cross-feature integration and Phase 2 release readiness

```yaml
id: task_054
title: Verify Ranges and PivotTables integration and background release readiness
priority: critical
depends_on: [task_047, task_052, task_053, task_034, task_036]
source_reference: plan.md; plan_addendum_background_execution.md; excel_vba_mcp_technical_spec.docx
source_section: "Phase 2 — Ranges & Cells; PivotTables; background execution addendum"
source_file: plan.md
description: >
  Perform the final independent integration and regression gate for the Ranges & Cells and PivotTables
  increment, including the inherited no-window/no-focus-steal contract.
goals:
  - verify Range tools, Pivot tools, workbook model, and declarative TestRunner use consistent DTOs and identities
  - verify Range operations do not corrupt ListObjects, Names, Sheets, or existing Stage 1 objects
  - verify PivotTable operations do not invoke deferred Connections, Power Query, Charts, Slicers, or Data Model services
  - run the full unit, integration, and regression suite for the affected features
  - verify snapshot/restore, approval, audit, locks, recovery, and artifact manifests
  - verify desktop non-interference for reads, writes, formatting, structural operations, rule changes, refresh, layout, and filters
  - reconcile all task IDs and evidence before marking the increment complete
extra_context: |
  The background addendum is authoritative. Every operation runs in the current interactive user
  session through the hidden automation-owned Excel instance and existing STA worker. No VM,
  container, alternate desktop, separate user/session, Windows Service, or visible UI fallback is
  allowed. If an operation cannot be completed without UI, classify it as blocked/unsupported and
  do not report the feature as complete for that operation.
files:
  - path: docs/release-evidence-phase2-ranges-pivots.md
    targets: [{type: document, name: Phase 2 Range and Pivot release evidence}]
  - path: tests/integration_excel/test_phase2_ranges_pivots_regression.py
    targets: [{type: test_module, name: Cross-feature regression and desktop non-interference tests}]
changes:
  - action: add_phase2_release_gate
    description: Aggregate feature verification evidence and publish the release decision for this increment.
acceptance_criteria:
  - task_047, task_052, and task_053 are DONE
  - no critical Auditor or Validator findings remain open
  - all documented supported operations pass real Excel integration tests
  - all unsupported cases return explicit statuses and never false success
  - the background BG01–BG09 criteria relevant to this increment are satisfied
  - no Excel/VBE window appears, activates, or steals foreground focus during the test suite
  - Stage 1 tests and previously accepted behavior remain green
  - release evidence records Windows/Office build, fixture hashes, and executed test results
  - final task status is COMPLETE only when every acceptance criterion is evidenced
tests_to_run: ["full affected unit suite", "full affected Excel integration suite", "test_phase2_ranges_pivots_regression.py"]
agents:
  - auditor: required final architecture/code review
  - validator: required independent release verification
verification_mode: dedicated_gate
```

---

## 4. Risk register and unresolved decisions

| ID | Risk / uncertainty | Required investigation or mitigation | Gate / owner task |
|---|---|---|---|
| P2-R01 | Excel date serials returned by `Value2` can be mistaken for ordinary numbers | Define typed conversion using cell format/context; preserve raw numeric serial if date classification is ambiguous | `task_038`, `task_040`, `task_042` |
| P2-R02 | `Formula` and `Formula2` differ for dynamic arrays and compatibility | Research target Excel support and formula round-trip semantics before freezing contracts | `task_038` |
| P2-R03 | Bulk range requests can exceed MCP payload or memory limits | Enforce cell/row/byte limits and return bounded structured outputs | `task_040`, `task_042` |
| P2-R04 | Range structural operations may damage ListObjects, merged ranges, spills, or formulas | Preflight target intersections; reject ListObject structure changes; snapshot and approval for destructive operations | `task_045`, `task_047` |
| P2-R05 | Excel formatting can be mixed or not representable by one scalar value | Preserve mixed/unsupported indicators; never silently flatten formatting | `task_038`, `task_044` |
| P2-R06 | Validation and conditional formatting have heterogeneous rule types and parameter limits | Freeze a bounded rule matrix; reject unsupported types explicitly | `task_038`, `task_046` |
| P2-R07 | A workbook may have several PivotTables sharing a cache | Represent cache identity and all affected tables; do not treat cache refresh as a single-table mutation | `task_039`, `task_048`, `task_049` |
| P2-R08 | PivotTable and PivotCache refresh paths have different scopes and return semantics | Research `RefreshTable` vs `PivotCache.Refresh`; target only approved local caches | `task_039`, `task_049` |
| P2-R09 | OLAP/Data Model PivotFields behave differently from ordinary PivotFields | Keep Data Model/Power Pivot out of scope; expose explicit unsupported/read-only status | `task_039`, `task_048`–`task_052` |
| P2-R10 | PivotFilters and manual item visibility have different semantics and limits | Define supported filter matrix and postcondition checks before implementing mutations | `task_039`, `task_051` |
| P2-R11 | Pivot cache refresh may involve an external connection or produce modal UI | Do not call RefreshAll; restrict refresh to local source types; block UI-dependent/unsupported refresh | `task_049`, `task_052` |
| P2-R12 | New tools can drift from the existing TestRunner and workbook model | Reuse existing services and DTOs; extend `WorkbookModelService` and `TestRunner` rather than duplicating COM logic | `task_048`, `task_053`, `task_054` |
| P2-R13 | Hidden Excel may surface dialogs or steal focus during range/pivot mutations | Assert no visible/foreground windows; report blocked/indeterminate status and stop COM queue on modal block | `task_047`, `task_052`, `task_054` |

### Technical choices intentionally left for research

- `Range.Value2` versus `Value` and exact date/error conversion rules.
- `Range.Formula` versus `Formula2`, including dynamic-array compatibility.
- The bounded formatting property set exposed by the API.
- Which DataValidation and FormatConditions rule types are supported in the first implementation increment.
- Exact copy/insert/delete/merge preconditions and supported multi-area behavior.
- The authoritative PivotTable discovery path and stable cache identity strategy.
- When to use `PivotTable.RefreshTable` versus `PivotCache.Refresh`, including shared-cache behavior.
- Which PivotField layout/aggregation and filter combinations are supported on the target build.

These choices must not be silently decided by Implementor. Researcher produces evidence; Auditor checks architectural fit; Tech Lead records the final decision before implementation.

---

## 5. Quality gates

### GATE P2-1 — Parent and background prerequisites

- Parent `task_037` is `DONE`.
- The background addendum requirements BG01–BG09 are implemented and verified for the existing Excel session/STA/snapshot/recovery stack.
- No alternate execution environment is introduced.

### GATE P2-2 — Research and contract freeze

- `task_038` and `task_039` decision records are complete.
- Auditor has reviewed Pivot architecture compatibility and the shared service boundaries.
- Validator has reviewed the high-risk Pivot decision record.
- `task_040` contracts are stable, bounded, serializable, and backward compatible.

### GATE P2-3 — Feature implementation

- Each feature task passes its inline tests.
- All mutation paths use existing snapshot, approval, audit, locking, and recovery services.
- All API handlers delegate to application services; COM logic remains inside adapters/STA worker.

### GATE P2-4 — Independent feature verification

- `task_047` passes for Ranges & Cells.
- `task_052` passes for PivotTables.
- Real Excel integration tests have been executed in the current interactive Windows session.
- No mock-only result is used to claim Excel behavior was verified.

### GATE P2-5 — TestRunner and cross-feature integration

- `task_053` reuses the existing TestRunner and existing feature services.
- Both feature branches run from independent baseline copies.
- Range operations preserve Stage 1 Tables/Names/Sheets behavior.
- Pivot operations preserve workbook model consistency and do not invoke deferred APIs.

### GATE P2-6 — Phase 2 increment release

- `task_054` passes.
- No critical review/validation finding remains open.
- All supported operations have real integration evidence.
- Unsupported cases produce explicit results, never false success.
- BG01–BG09 desktop non-interference requirements relevant to this increment pass.

---

## 6. Parallel execution and implementation ownership

After `task_040`, the Range and Pivot branches may be assigned to separate Implementors if their worktrees and file scopes are isolated. Recommended ownership:

- **Range Implementor:** `ranges_service.py`, `ranges_adapter.py`, `range_tools.py`, and range-specific tests.
- **Pivot Implementor:** `pivot_tables_service.py`, `pivot_tables_adapter.py`, `pivot_tools.py`, plus the PivotTable section of `WorkbookModelService` after coordination with Tech Lead.
- **Shared-contract Implementor:** `domain/models.py`, `errors.py`, `statuses.py`; complete this before either branch begins.
- **TestRunner Implementor:** `domain/test_spec.py`, `application/test_runner.py`, and `api/test_tools.py`; start after both feature APIs are stable.

Do not allow two agents to edit the same shared file concurrently. Use separate worktrees/branches when supported; otherwise sequence the tasks. Git integration is performed by Tech Lead after individual task commits and required reviews.

All Excel integration tests, fixture generation, and runtime verification must be serialized through the single existing STA worker. Parallel code preparation and documentation/review are permitted; concurrent COM activity against the same Excel instance is not.

---

## 7. Completion definition

This increment is `COMPLETE` only when:

- Every task `task_038`–`task_054` is `DONE`.
- Both feature verification gates pass using real Excel.
- The TestRunner supports bounded Range and Pivot scenarios without duplicating business logic or COM access.
- The full affected regression suite passes.
- There are no unresolved critical Auditor or Validator findings.
- All supported operations meet the hidden same-session/no-focus-steal requirement.
- Unsupported external/OLAP/Data Model/UI-dependent cases are explicitly reported and are not represented as successful operations.
- Release evidence records the Office/Windows build, fixture hashes, test results, and relevant limitations.

---

## 8. Primary technical references

These references support the API behavior to be validated during implementation; the research tasks remain responsible for confirming compatibility on the target Office build.

### Range and Cells

- [Range.Value2 property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2) — value read/write; Date/Currency variants are returned as floating-point values.
- [Range.Formula2 property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2) — dynamic-array formula semantics and backward compatibility with Formula.
- [Validation object](https://learn.microsoft.com/en-us/office/vba/api/excel.validation) and [Validation.Add method](https://learn.microsoft.com/en-us/office/vba/api/excel.validation.add) — data validation operations and parameters.
- [FormatConditions object](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions) and [FormatConditions.Add method](https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.add) — conditional formatting collection and rule creation.

### PivotTables

- [PivotTable object](https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable) — PivotTable members and field configuration.
- [PivotTables collection](https://learn.microsoft.com/en-us/office/vba/api/excel.pivottables) — collection discovery; note the documented workbook-level caveat and verify worksheet-level enumeration.
- [PivotCache object](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache) and [PivotTable.PivotCache property](https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.pivotcache) — cache metadata and relationships.
- [PivotTable.RefreshTable method](https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.refreshtable) and [PivotCache.Refresh method](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcache.refresh) — distinct targeted refresh operations.
- [PivotField.Orientation property](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.orientation) and [PivotField.Position property](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.position) — field placement and ordering.
- [PivotFilters.Add method](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfilters.add) and [PivotField.ClearAllFilters method](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotfield.clearallfilters) — filter operations.

---

**Interpretation rule:** This is an incremental Phase 2 child plan for Ranges & Cells and PivotTables only. The parent `plan.md` and `plan_addendum_background_execution.md` remain authoritative for all inherited architecture, safety, background execution, and execution rules. This plan adds tasks `task_038`–`task_054`; it does not rewrite or supersede existing task specifications.
