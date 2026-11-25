# Architecture Decision Report
## Handling Short and Long-Running Operations in xlvba MCP

**Status:** Recommended — subject to Proof of Concept validation  
**Priority:** Implement a hybrid synchronous/asynchronous execution model.

## 1. Executive Decision

Adopt an application-level **Operation Manager** on top of the existing Excel COM execution architecture.

The MCP server must distinguish between:
- Short operations that can safely complete within a normal MCP request.
- Long-running operations whose execution time is unpredictable or may exceed client request timeouts.

The critical architectural rule is:

**An MCP request timeout must never be interpreted as proof that Excel is hung or as a reason to terminate or recover the Excel Worker.**

A long-running Excel operation must continue independently of the MCP request that initiated it.

Keep a single Excel COM STA execution channel. Do not introduce multiple concurrent Excel workers, virtual machines, containers or separate Windows sessions.

## 2. Architecture for Short Operations

Short operations should retain the existing synchronous MCP interface.

Examples:
- Reading a bounded cell range.
- Reading a small Excel Table.
- Updating report settings.
- Adding or updating a limited number of table rows.
- Reading worksheet and defined-name metadata.
- Simple formatting operations.

### Execution flow

1. The agent calls the existing MCP tool.
2. The MCP handler validates arguments and permissions.
3. The operation is submitted to the existing serialized Excel execution queue.
4. The Excel STA Worker executes the COM operation.
5. The handler waits for completion.
6. The result is returned through the original MCP tool response.

### Requirements

- Preserve existing tool names and response contracts.
- Enforce reasonable input and output size limits for operations classified as short.
- Serialize all Excel COM access through the same STA Worker.
- Do not execute parallel COM calls against the same Excel instance.
- If a long operation currently occupies the worker, return a structured retryable `WORKER_BUSY` response instead of blocking another MCP request indefinitely.
- Keep request timeout handling separate from worker health management.
- Track execution internally so that an abandoned client request does not cause an automatic retry of a potentially completed mutation.

Progress notifications may be used for operations with measurable intermediate stages, but they are optional and do not replace asynchronous execution.

## 3. Architecture for Long-Running Operations

Introduce an **Operation Manager** responsible for accepting, tracking and managing long-running Excel operations independently of MCP request lifetimes.

Examples:
- Opening large Excel workbooks.
- Enumerating PivotTables in large workbooks.
- Reading PivotCache metadata or contents.
- Refreshing PivotTables.
- Full workbook model inspection.
- Running arbitrary VBA macros.
- VBA compilation.
- Large workbook processing and export operations.

### Execution flow

1. The agent calls `xlvba_operation_start`.
2. The MCP server validates the requested operation and its arguments against registered tool schemas and existing permissions.
3. The Operation Manager creates and persists an operation record.
4. The operation is placed in the Excel execution queue.
5. MCP immediately returns an `operation_id`.
6. The Excel STA Worker executes the operation independently of the initiating request.
7. The agent periodically calls `xlvba_operation_status`.
8. When execution completes, the agent retrieves the result through `xlvba_operation_result`.

### Required MCP tools

| Tool | Responsibility |
|---|---|
| `xlvba_operation_start` | Submit a registered long-running operation |
| `xlvba_operation_status` | Retrieve current execution status |
| `xlvba_operation_result` | Retrieve the final result |
| `xlvba_operation_cancel` | Request cancellation |

Do not expose arbitrary internal method execution. Only explicitly registered and schema-validated Excel operations may be submitted.

### Operation state model

The Operation Manager must distinguish:

- `queued`
- `running`
- `long_running`
- `cancel_requested`
- `completed`
- `failed`
- `cancelled`
- `interrupted_unknown`

`long_running` is informational, not an error or recovery state.

Persist operation metadata and status in a lightweight local SQLite database. Store large results and artifacts in the existing filesystem artifact directory.

Persist:
- Operation ID.
- Operation type.
- Workbook ID.
- Submission time.
- Start time.
- Completion time.
- Current status.
- Progress information, when available.
- Result location.
- Error details.
- Idempotency key, where applicable.

Do not automatically retry an interrupted macro or mutation. Its execution outcome may be unknown.

## 4. Excel Worker Architecture

Keep one Excel STA execution channel and one owned hidden Excel instance in the current interactive Windows user session.

The Operation Manager and MCP control plane must remain responsive even when Excel is processing a long COM call.

Target logical architecture:

```text
MCP Client
    |
    v
MCP Server / Control Plane
    |
    +---- Short Operations
    |        |
    |        v
    |    Excel Command Queue
    |
    +---- Long Operations
             |
             v
       Operation Manager
             |
             v
       Excel Command Queue
             |
             v
       Single Excel COM STA Worker
             |
             v
       Hidden Excel Instance
```

The STA Worker must remain the exclusive owner of its COM interfaces. No other thread or process may access its Excel COM objects directly.

The control plane must be able to process operation-status requests without querying Excel through the busy STA channel.

First verify whether the existing worker thread can provide this separation reliably. If it cannot, evaluate hosting the single STA worker in a supervised child process in the same Windows user session. This is process-level supervision, not a VM or separate execution environment.

## 5. Timeout Semantics — Critical Requirement

Three different concepts must be implemented separately.

### A. MCP Request Timeout

Controls how long the client waits for an individual MCP response.

A client-side timeout must not stop the Excel operation or change worker health.

### B. Operation Soft Threshold

Used to identify operations taking longer than expected.

For example, after 60 seconds:

```json
{
  "status": "long_running",
  "elapsed_seconds": 60,
  "worker_state": "busy",
  "recovery_required": false
}
```

This is a status update, not an error.

### C. Actual Execution Failure

Worker recovery is justified only by evidence of an actual failure, such as:

- Excel process termination.
- Confirmed COM server failure.
- Worker process termination.
- Definitive unrecoverable execution error.

Elapsed time alone, high memory usage or missing intermediate progress must never be sufficient evidence.

A hard execution limit may be introduced only when an operation has an explicitly defined safe cancellation policy. It must not be used as a generic replacement for the existing 60-second soft timeout.

## 6. Cancellation and Idempotency

Cancellation is a request, not a guarantee that Excel has stopped executing.

The system must distinguish:

- Client stopped waiting.
- Cancellation was requested.
- Operation actually stopped.
- Operation completed before cancellation was processed.

Do not kill Excel simply because cancellation was requested or a soft threshold was exceeded.

Use idempotency protection for operations that modify workbooks, especially macro execution. Never automatically rerun an operation whose previous execution outcome is unknown.

## 7. MCP Protocol Compatibility

Use standard MCP tools for the initial Operation Manager implementation.

The current official MCP Python SDK v2 supports inline progress reporting through `Context.report_progress()`, but its migration guide explicitly states that the previous experimental Tasks runtime has no direct replacement.

The MCP Tasks Extension is a standardized protocol design for asynchronous task handles and polling, but integration must not depend on unsupported SDK functionality.

Therefore:

- Implement application-level operation management using ordinary MCP tools.
- Use MCP progress notifications as an additional reporting mechanism.
- Keep the internal Operation Manager independent of the transport.
- Design the operation model so it can later be adapted to the standard MCP Tasks Extension when compatible SDK and client support is available.

References:
- https://github.com/modelcontextprotocol/modelcontextprotocol/blob/main/docs/seps/2663-tasks-extension.mdx
- https://py.sdk.modelcontextprotocol.io/handlers/progress/
- https://py.sdk.modelcontextprotocol.io/migration/

## 8. Proof of Concept and Acceptance Criteria

Before full implementation, prove the following:

1. A COM operation can run for more than 60 seconds without triggering `NEEDS_RECOVERY`.
2. The MCP client receives an operation ID promptly.
3. Status and health requests remain responsive while Excel is busy.
4. The Excel Worker does not receive concurrent COM commands.
5. The completed result can be retrieved independently of the original request.
6. A disconnected client does not automatically cancel the operation.
7. Cancellation does not falsely report success while Excel is still running.
8. A genuine Excel process failure is detected and reported correctly.
9. An interrupted mutation is marked `interrupted_unknown` and is never automatically repeated.
10. A real large-workbook scenario, including the reported 1.3 GB PivotCache workload, does not trigger false recovery.

## 9. Final Recommendation

Prioritize the following implementation sequence:

1. Correct timeout semantics and remove elapsed-time-based worker failure classification.
2. Introduce the Operation Manager and operation state model.
3. Separate MCP status handling from the busy Excel STA execution channel.
4. Implement submit, status, result and cancellation tools.
5. Integrate long-running PivotTable, workbook and VBA operations.
6. Validate the full workflow against real large Excel workbooks.

**Final architectural principle:**

Short operations remain synchronous and convenient. Long operations become independently managed jobs. Excel continues to execute through one serialized STA worker, while MCP remains responsive and accurately reports operation status.

The system must never confuse a slow but healthy Excel operation with a failed worker.