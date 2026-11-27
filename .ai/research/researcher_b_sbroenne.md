# Researcher B — `sbroenne/mcp-server-excel` PivotTable / PivotCache / OLAP / Data Model large-data handling

READ-ONLY research. No files in `C:\py_exp\xlvba_mcp` were modified except this report.

## Commit checked
- Repo: `sbroenne/mcp-server-excel` (default branch `main`, C# / .NET, MIT).
- HEAD commit: **`e42059685810775ee5f8277a56972c44546f40b0`** (2026-11-24, "chore(actions)(deps): Bump the github-actions group…", author dependabot[bot]).
- Tree SHA: `4403e8640d518e543263ae4c67ac82ae04328560`. Files fetched from `raw.githubusercontent.com` at that exact SHA.

## Architectural context (critical distinction for xlvba_mcp)
Their server is **multi-process**: `ExcelController` (host) spawns up to **32 private Excel worker processes** (`ExcelController.StartWorkerAsync`, `_workers.Count >= 32` cap), each worker (`ExcelWorkerHost`) hosting **one `ExcelBatch` = ONE dedicated STA thread + ONE hidden Excel instance** (`ExcelBatch` ctor `RunStaThread`, `_staThread.SetApartmentState(ApartmentState.STA)`). Workers talk to the controller over named pipes + JSON-RPC (`WorkerProtocol.cs`, `IExcelWorkerRpc`). The controller polls each worker every 1 s (`ExcelController.PollAsync`) and tracks late/abandoned results.

**Our model is the opposite:** ONE owned hidden Excel COM instance, ONE STA worker, FIFO COM queue. So every technique that relies on "spin up another Excel process / another worker" is REJECT for us; techniques that make a *single* STA call cancellable/bounded are ADAPT/ADOPT.

---

## 1. Separating metadata reads from result-data reads

- **Metadata / configuration read** = `PivotTableCommands.Read(...)` in `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Lifecycle.cs` (lines 136–272). Returns `PivotTableInfoResult` (fields, layout, `FieldSelections`, `ValueFields`, `IsOlap`, `MaxSelectedMembers`, `SnapshotComplete`). It reads **no result cells** — only `TableRange2.Address`, field counts, `CubeFields`/`PivotFields` metadata, member selection state, filters.
- **List (workbook-wide, no data)** = `PivotTableCommands.List(...)` same file (lines 16–131): per-pivot `Name`, `CacheIndex`, `Range`, `LayoutError`, `SourceData`, field counts, `LastRefresh` + `#SPILL!` layout check via `PivotTableHelpers.GetLayoutError`.
- **Result-data read** = `PivotTableCommands.GetData(...)` in `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Analysis.cs` (lines 17–89). Reads the **entire** `TableRange2.Value2` and `TableRange1.Value2` in one shot via `ExcelValueNormalizer.Normalize`. **No paging / no row bound** — see §4.
- **Field metadata only** = strategy `ListFields` (`OlapPivotTableFieldStrategy.ListFields`, `RegularPivotTableFieldStrategy.ListFields`) exposed as `pivottable_field list-fields`.

Verdict: **ADOPT DIRECTLY** the split (metadata `read` vs data `get-data` are separate actions, separate DTOs). It maps cleanly onto our single-STA queue (each is one bounded COM call).

## 2. Bounded PivotTable member enumeration (limits + truncation flags)

- `PivotTableSnapshotHelpers` (`src/ExcelMcp.Core/Commands/PivotTable/PivotTableSnapshotHelpers.cs`): `DefaultMaxSelectedMembers = 100`, `MaximumMaxSelectedMembers = 1000`, `ValidateMaxMembers(1..1000)`.
- `PivotTableCommands.SemanticSnapshot.cs`:
  - `BoundedMemberNames` (lines 600–614): a `SortedSet<string>` that keeps only the **ordinal-smallest `maxMembers`** names (`if (_names.Count > maxMembers) _names.Remove(_names.Max!)`), while still counting the true total.
  - `GetFieldSelectionSnapshots` (lines 30–136) caps the cross-field union the same way and sets `MembersTruncated = SelectedMemberCount > SelectedMemberUniqueNames.Count` (lines 127–129, 217–219, 349–351).
  - DTO flags: `PivotFieldSelectionSnapshot.MembersTruncated`, `SelectedMemberCount`, `MaxSelectedMembers`, and top-level `SnapshotComplete` (`PivotTableTypes.cs` lines 468–471, 520–523).
- Per-property reads are individually wrapped in `TryReadSnapshotProperty` (catches reportable `COMException`/`NotSupportedException` → `Available=false` + message, never throws for a single bad property).

Verdict: **ADOPT DIRECTLY** (the bounded-set + truncation-flag + per-property failure isolation pattern is exactly what a single STA worker needs to avoid unbounded enumeration).

## 3. OLAP / Data Model inspection via COM — exact properties touched

OLAP detection: `PivotTableHelpers.HasOlapCache` reads **`PivotCache.OLAP`**; `IsOlapPivotTable` then reads **`PivotTable.CubeFields.Count`** (`PivotTableHelpers.cs` lines 208–264). Cache type: **`PivotCache.SourceType`** (`xlDatabase` vs external/Data Model) in `Lifecycle.Read`/`List`.

OLAP field metadata (`GetCubeFieldsInfo`, `Lifecycle.cs` 277–304; `OlapPivotTableFieldStrategy.ListFields` 75–127): per `CubeField` reads **`.Name`, `.Caption`, `.Orientation`, `.Position`, `.CubeFieldType`** (and calls `.CreatePivotFields()` only for hidden hierarchies before manipulation — `GetFieldForManipulation` 19–71).

OLAP member selection (`ReadOlapFieldSelection` / `ReadOlapMemberState`, `SemanticSnapshot.cs` 240–373, 534–587):
- `CubeField.Name/.Caption/.Orientation/.Position`
- `CubeField.PivotFields` → level `PivotField`s; per level `.Orientation`, `.Name`, **`.VisibleItemsList`** (the member unique-name array), **`.CurrentPageName`** (single-select report filter), `CubeField.EnableMultiplePageItems`, `CubeField.AllItemsVisible`.
- "All member" disambiguation: `PivotTableHelpers.IsAllOlapReportFilterMember` opens the **ADO schema rowset** (`PivotCache.ADOConnection`, `WorkbookConnection`, `ConnectionCommands.OpenSchemaRowset`, `ALL_MEMBER` column) — NOT a cell read.
- Report filter read/write: `PivotTableCommands.Analysis.GetReportFilter/SetReportFilter` use `CubeField.Orientation`, `CubeField.PivotFields`, `PivotField.CurrentPageName`, `PivotField.VisibleItemsList`, `CubeField.EnableMultiplePageItems`, `CubeField.AllItemsVisible`.
- Value-field calc readback: `ReadValueFieldCalculation` (`Calculation.cs` 259–320) reads `PivotField.Calculation/.SourceName/.Position/.Function` and, **only for non-OLAP**, `.BaseField/.BaseItem`. OLAP explicitly skips BaseField/BaseItem (throws `NotSupportedException` if requested).

How they avoid hangs on OLAP: (a) every property read is wrapped in `TryReadSnapshotProperty`; (b) `VisibleItemsList`/`CurrentPageName` are read as **arrays/scalars**, never by iterating `PivotItems` for OLAP; (c) ADO schema access is bounded (`maximumRows`, `MaximumOlapMemberRowsScanned = 1000`, `MaximumOlapMemberPageSize = 250`); (d) the whole batch runs on the STA with `OleMessageFilter` + a 2-min `CancellationTokenSource` (`ConnectionCommands.ExternalOlapSchema.cs` lines 51, 130).

Verdict:
- **ADOPT DIRECTLY**: `PivotCache.OLAP`, `SourceType`, `CubeFields.Count`, `CubeField.Name/Caption/Orientation/Position`, `CubeField.PivotFields`, `PivotField.VisibleItemsList`, `CurrentPageName`, `EnableMultiplePageItems`, `AllItemsVisible`. These are scalar/array reads, safe on one STA.
- **ADAPT**: `IsAllOlapReportFilterMember` ADO-rowset path — keep the bounded `OpenSchemaRowset(maximumRows)` idea, but it must run on our single STA and tolerate no-ADO-session (return "unknown" + recovery guidance) rather than assuming a second worker.
- **REJECT**: nothing in OLAP *property reads* requires a second instance; the multi-instance part is only their worker hosting, not the property set.

## 4. `pivottablecalc` operations + large result retrieval

- `pivottable_calc` tool (`IPivotTableCalcCommands`, `[McpReadOnlyActions("get-data", "list-calculated-fields", "list-calculated-members", "get-layout-options", "list-calculated-items")]`) hosts `GetData`.
- `GetData` (`Analysis.cs` 17–89) materialises the **whole** pivot body (`TableRange2.Value2` + `TableRange1.Value2`) — **no paging, no row cap, no truncation flag**. Large pivots return one giant 2-D array.
- Calculated fields/items/members (`CalculatedFields.cs`, `CalculatedItems` in `Reporting.cs` 106–199, `CalculatedMembers.cs`) iterate `PivotFields`/`PivotItems`/`CalculatedItems` with `ct.ThrowIfCancellationRequested()` per item but **no count cap**.
- `pivottable_calc set-layout-options` (`DepthLayout.cs`) iterates row fields with per-field cancellation, no cap.

Verdict:
- **ADAPT**: the per-item `ct.ThrowIfCancellationRequested()` loop is good and we already do it; but `GetData` needs a **row/cell bound + truncation flag** added for us (they lack it). Do NOT copy the unbounded full-range read.
- **REJECT**: nothing here needs multiple instances.

## 5. Controller/worker separation, monitoring, late results, lifecycle

- `ExcelController` (`src/ExcelMcp.Service/ExcelController.cs`): `ProcessCoreAsync` routes each command to a worker; `CompleteWorkerRequestAsync` (254–341) owns **late completion after caller cancellation** and records `outcome="unknown"` / "do not repeat edits" (`CreateWorkerError` 343–360). `PollAsync` (737–786) refreshes every 1 s, retires exited workers (`RetireExitedWorker` 665–693), releases idle workers (`ReleaseEmptyWorkerAsync`). `StartWorkerAsync` caps at 32.
- `ExcelWorkerClient.CallAsync` (106–140): per-call timeout, **detects worker process exit while a response is pending** and reports "outcome unconfirmed; no work was repeated"; `TimeoutException` on deadline.
- `ExcelWorkerHost.ExecuteAsync` (239–312): single `_execution` semaphore (1,1); rejects concurrent work with `Busy`; tracks `_pendingOperation/_pendingCommand`; on caller cancel calls `RevealAbandonedRequest` (shows the Excel window so the user can see the still-running op). `MonitorControllerAsync` (403–459) reveals sessions if the controller dies.
- Worker identity is verified by PID + start-time + session id (`WorkerProtocol.WorkerProcess.IsAlive`, `WorkerPipeIdentity.VerifyServer`).

Verdict:
- **ADOPT DIRECTLY (concepts)**: "report unknown outcome, never auto-repeat a mutation on timeout/cancel" and "reveal the window on abandonment" are model-agnostic and we should keep them.
- **REJECT (mechanism)**: the whole controller↔worker pipe fabric, 32-worker pool, per-worker process-exit detection, and 1 s multi-worker poll — these exist *because* they run many Excel instances. Under our single owned instance + FIFO queue we must NOT add a second process; we rely on agent-directed recovery. We keep only the *outcome-unknown / no-repeat* reporting and our existing recovery.

## 6. CLI and MCP shared Core operations

- Both `excelcli` (`src/ExcelMcp.CLI/Commands/BatchCommand.cs`, `Program.cs`) and MCP (`src/ExcelMcp.McpServer/Tools/ExcelBatchTool.cs`) are thin transports that build a `ServiceRequest` and call the **same** `ExcelMcpService`/`ServiceRegistry` → `PivotTableCommands` etc. `ExcelBatchTool.ExecuteBatchAsync` explicitly reuses `ServiceRegistry.ValidateCommandArguments` and the CLI batch entry format (`command` = `group.action`, `args` = camelCase Service keys).
- Command metadata (actions, read-only actions, categories) lives once as attributes on the Core interfaces (`[McpTool]`, `[ServiceAction]`, `[McpReadOnlyActions]`).

Verdict: **ADOPT DIRECTLY** (single Core command layer + attribute-driven action metadata, shared by CLI and MCP). Matches our "Core operations shared by transports" direction.

## 7. PivotCache record reading / large caches

- There is **no API that reads PivotCache *records*** (no `PivotCache.Records` iteration). "RecordCount" is read as a scalar (`PivotCache.RecordCount`) in `Lifecycle.Refresh` (408–413) and `Source.ReadPivotSource` (211).
- Large *source* caches are never enumerated; `GetSource` only reads `SourceData` (a string/address) + `RecordCount` + shared-pivot names (`Source.cs` 203–260).
- OLAP member enumeration over a real cache is bounded via ADO schema rowsets (`ExternalOlapSchema.cs`, `maximumRows`, `MaximumOlapMemberRowsScanned=1000`, page size ≤250) with a 2-min CTS.

Verdict: **ADOPT DIRECTLY** (scalar `RecordCount` + bounded ADO schema rowsets; never iterate records). **REJECT** nothing.

## Timeout / isolation that we lack (single-STA note)
They DO have a same-thread cancellation mechanism we should note: **`OleMessageFilter`** (`src/ExcelMcp.ComInterop/OleMessageFilter.cs`) registered on the STA thread. `SetPendingCancellationToken(token)` makes `IMessageFilter.MessagePending` return **`PENDINGMSG_CANCELCALL` (0)** when the token fires, aborting a blocked synchronous COM dispatch (e.g. `connection.Refresh()`) with `RPC_E_CALL_CANCELLED` **without a second process**. `EnterLongOperation`/`HandleInComingCall` reject re-entrant callbacks with `SERVERCALL_RETRYLATER` to stop CPU spin during Data Model/refresh. `ExcelBatch.WaitWithDialogMonitoringAsync` + `ExcelDialogProbe.Read` detect a modal dialog from another thread and reveal the window. This is the key isolation primitive that survives our single-instance model.

Verdict: **ADAPT** — an `IOleMessageFilter` with `PENDINGMSG_CANCELCALL` on our single STA is the right way to bound a hung COM call without adding a process; pair with agent-directed recovery (we already rely on that). The 32-worker pool / per-worker process-exit detection is **REJECT**.

---

## OLAP / Data Model properties: SAFE vs RISKY under a single STA worker
SAFE (scalar/array reads, no hang risk): `PivotCache.OLAP`, `PivotCache.SourceType`, `PivotCache.RecordCount`, `PivotCache.RefreshDate`, `PivotTable.CubeFields.Count`, `CubeField.Name/.Caption/.Orientation/.Position/.CubeFieldType`, `CubeField.PivotFields`, `CubeField.EnableMultiplePageItems`, `CubeField.AllItemsVisible`, `PivotField.VisibleItemsList` (array), `PivotField.CurrentPageName`, `PivotField.Orientation/.Name/.SourceName/.Position/.Calculation`, `PivotTable.TableRange2.Address`, `PivotField.Function` (non-OLAP), `Model.ModelTables/ModelMeasures/ModelRelationships` counts/names.
RISKY (can block / needs bound + cancellation, still single-STA-safe if wrapped): `PivotTable.RefreshTable()` (use `OleMessageFilter` + CTS, `PivotTableRefreshHelper.RefreshSynchronously`), `PivotCache.ADOConnection` + `OpenSchema` rowsets (bound rows + 2-min CTS), `TableRange2.Value2` full-body read in `GetData` (add a row/cell cap), `Model.DataModelConnection` DAX `Execute` (2-min CTS + recordset close in `finally`). No OLAP/Data Model property requires a second Excel process.

## Confidence
- HIGH: file/function citations above are read verbatim at the pinned SHA; property names are quoted from source.
- MEDIUM: exact runtime hang behaviour of each property is inferred from code structure + their tests (`PersistentServicePivotTableOlapFieldTests`, `PersistentServicePivotTableCacheScaleTests`), not executed here.
- Not verified in code: any claim that a specific OLAP property *always* hangs — the repo wraps everything defensively rather than asserting a hang.

## Source citations (repo-relative paths + functions, commit `e42059685810775ee5f8277a56972c44546f40b0`)
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Lifecycle.cs` — `List`, `Read`, `GetCubeFieldsInfo`, `GetRegularFieldsInfo`, `Refresh`, `GetRefreshDateSafe`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.SemanticSnapshot.cs` — `GetFieldSelectionSnapshots`, `ReadRegularFieldSelection`, `ReadOlapFieldSelection`, `ReadRegularMemberState`, `ReadVisibleRegularMembers`, `ReadAllRegularMembers`, `ReadOlapMemberState`, `BoundedMemberNames`, `ReadOlapLevelFilters`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableSnapshotHelpers.cs` — `DefaultMaxSelectedMembers`, `MaximumMaxSelectedMembers`, `ValidateMaxMembers`, `TryReadSnapshotProperty`, `CreateFieldSelection`, `PivotSnapshotProperty<T>`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableHelpers.cs` — `IsOlapPivotTable`, `TryGetCubeFields`, `HasOlapCache`, `IsAllOlapReportFilterMember`, `GetLayoutError`, `IsValuesPlaceholderField`, `DetectFieldDataType`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Analysis.cs` — `GetData`, `GetReportFilter`, `SetReportFilter`, `ReadReportFilterSelection`, `SortField`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Calculation.cs` — `SetFieldCalculation`, `ReadValueFieldCalculation`, `ReadAllValueFieldCalculations`, `DescribeCalculationReadbackFailure`.
- `src/ExcelMcp.Core/Commands/PivotTable/OlapPivotTableFieldStrategy.cs` — `ListFields`, `GetFieldForManipulation`, `AddRowField`/`AddColumnField` (note: "No RefreshTable() needed … RPC disconnection on rapid operations (issue #426)").
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Cache.cs` — `GetCacheOptions`, `SetCacheOptions`, `ReadCacheOptions`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.Source.cs` — `GetSource`, `SetSource`, `ReadPivotSource`, `ReadSharedPivotTables`, `ConnectedPivotSlicerCaches`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.DepthLayout.cs` — `GetLayoutOptions`, `SetLayoutOptions`, `ReadRowLayouts`.
- `src/ExcelMcp.Core/Commands/PivotTable/PivotTableCommands.DepthFilters.cs` — `GetFieldFilters`, `AddFieldFilter`, `ClearFieldFilters`, `ReadPivotFilters`.
- `src/ExcelMcp.Core/Commands/PivotTable/IPivotTableCommands.cs` / `IPivotTableFieldCommands.cs` / `IPivotTableCalcCommands.cs` — `[McpTool("pivottable"|"pivottable_field"|"pivottable_calc")]`, `[McpReadOnlyActions(...)]`, `Read(..., int maxSelectedMembers)`, `GetData`.
- `src/ExcelMcp.Core/Models/PivotTableTypes.cs` — `PivotTableInfoResult`, `PivotFieldSelectionSnapshot`, `PivotTableInfo`, `PivotTableDataResult`, `PivotTableRefreshResult`.
- `src/ExcelMcp.Core/Commands/Connection/ConnectionCommands.ExternalOlapSchema.cs` — `DiscoverOlapSchema`, `SearchOlapMembers`, `OpenSchemaRowset` (`MaximumOlapMemberPageSize=250`, `MaximumOlapMemberRowsScanned=1000`).
- `src/ExcelMcp.Core/Commands/Connection/ExternalOlapSchemaMapper.cs` — `BuildHierarchyRequest`, `BuildLevelRequest`, `BuildMembersRequest`, `MapSchema`, `SelectMemberPage`, continuation tokens.
- `src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.Read.cs` — `ListTables`, `ListMeasures`, `Read`, `ListRelationships`, `ListColumns`, `ReadTable`, `ReadInfo`.
- `src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.Evaluate.cs` — `Evaluate` (DAX via `Model.DataModelConnection` → ADO `Execute`, 2-min CTS, recordset close in `finally`).
- `src/ExcelMcp.Core/Commands/DataModel/DataModelCommands.Helpers.cs` — `HasDataModelTables`, `ForEachTable/Measure/Relationship`, `TryReadStoredMeasureFormulas` (`$SYSTEM.MDSCHEMA_MEASURES`), `CloseRecordset`.
- `src/ExcelMcp.Core/Utilities/PivotTableRefreshHelper.cs` — `RefreshSynchronously` (BackgroundQuery off + `OleMessageFilter.SetPendingCancellationToken`).
- `src/ExcelMcp.ComInterop/OleMessageFilter.cs` — `Register`, `MessagePending` (`PENDINGMSG_CANCELCALL`), `HandleInComingCall` (`SERVERCALL_RETRYLATER`), `RetryRejectedCall`, `EnterLongOperation`, `SetPendingCancellationToken`.
- `src/ExcelMcp.ComInterop/Session/ExcelBatch.cs` — `RunStaThread` (single STA + `OleMessageFilter.Register`), `Execute<T>` (`_operationTimeout`, `WaitWithDialogMonitoringAsync`), `CreateOperationTimeoutException`.
- `src/ExcelMcp.ComInterop/Session/IExcelBatch.cs` — `Execute`, `OperationTimeout`, `HasTimedOutOperation`.
- `src/ExcelMcp.Service/ExcelController.cs` — `ProcessCoreAsync`, `CompleteWorkerRequestAsync`, `CreateWorkerError`, `PollAsync`, `StartWorkerAsync` (32 cap), `RetireExitedWorker`, `ReleaseEmptyWorkerAsync`.
- `src/ExcelMcp.Service/Workers/ExcelWorkerClient.cs` — `CallAsync`, `ExecuteWithAdmissionAsync`, `RecordSnapshot`.
- `src/ExcelMcp.Service/Workers/ExcelWorkerHost.cs` — `ExecuteAsync`, `ReleaseAsync`, `MonitorControllerAsync`, `CompleteRelease`.
- `src/ExcelMcp.Service/Workers/WorkerProtocol.cs` — `WorkerProcess.IsAlive`, `WorkerRegistration`, `WorkerSnapshot`, `IExcelWorkerRpc`, `WorkerPipeIdentity`.
- `src/ExcelMcp.McpServer/Tools/ExcelBatchTool.cs` — `ExcelBatch`, `ExecuteBatchAsync` (shared Core via `ServiceRegistry`).
- `src/ExcelMcp.CLI/Commands/BatchCommand.cs` — CLI batch transport over the same Core.