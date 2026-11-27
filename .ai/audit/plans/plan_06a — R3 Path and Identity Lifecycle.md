# plan_06a — R3 Path & Identity Lifecycle (implementation prerequisite)

- **Status:** design prerequisite for R3; no approval gate. Preserve: original workbook (never opened/modified) / immutable snapshot (stored baseline) / mutable working copy (`work/<id>.ext`).

## Current (verified, WP3)
- Open: `record.canonical = <original source>`, `working_path = work/<id>.ext`, `source_sha256` captured.
- Snapshot create: stored under `snapshots/<snap>/<file>`; `manifest.source_path = <original>`.
- Restore: `open_workbook(<snapshot file>, EXECUTE, rw)` → new id, `working_path = work/<newid>.ext`, but `record.canonical = <snapshot file path>` (G5). Close previous by stable id (discard). Only new id valid after success.

## Target lifecycle (R3)
1. **Open** (unchanged): canonical = original; working = `work/<id>.ext`; source hash captured.
2. **Snapshot create** (unchanged): immutable file under `snapshots/`; `manifest.source_path = original`.
3. **Restore**:
   a. Load manifest; validate snapshot exists.
   b. Open the snapshot file as a NEW working copy → new id, `working_path = work/<newid>.ext` (the snapshot file is only a source for the copy; it is never mutated).
   c. Validate the restored working copy (read-only inspect).
   d. **Rebind identity:** set restored `record.canonical = manifest.source_path` (the ORIGINAL), and `record.source_sha256 = manifest.source_sha256` (original baseline). The session identity/`list_workbooks`/reopen now reflect the original, not the snapshot file. The snapshot file remains immutable and is NOT the session identity.
   e. Register/switch active to the new id; close previous by stable id (discard).
   f. Return restored `WorkbookInfo` (canonical = original) + `previous_close` status (R1).
4. **Repeated restore**:
   - Same snapshot while its restored copy is active → deterministic `SAME_SNAPSHOT` (reason `same_snapshot`) isError (clear, not `COM_BUSY`).
   - Different snapshot → allowed (new working copy; previous closed).
5. **No-op vs rollback**: a no-op restore (snapshot identical to current working state) still produces a valid restored copy bound to the original; a rollback restore (snapshot older) replaces the working copy content, original untouched.

## Invariants preserved
Original never opened/modified; snapshot immutable; working copy mutable and per-session; cleanup by stable id; never close source/restored/unrelated; failure to create/validate/activate never closes previous; close-previous failure never invalidates restored (R1 surfaces it).

## Changed files (R3)
`application/snapshot_service.py` (rebind canonical + source hash; SAME_SNAPSHOT), `application/session_service.py` (record canonical update), `domain/errors.py` (reuse `COM_BUSY`/validation or add `SAME_SNAPSHOT` reason), tests.