"""Policy-controlled macro execution service (task_022).

:class:`MacroService` is the application-layer orchestrator for the macro surface
(``macro_list``, ``macro_run``, ``macro_get_last_result``). It owns **no COM
proxies**: every Excel/VBIDE access is delegated to the single-STA
:class:`~excel_vba_mcp.adapters.excel.com_worker.ExcelComWorker` via
``await worker.submit(...)``. The synchronous ``Application.Run`` dispatch itself
runs on the STA worker inside :class:`~excel_vba_mcp.adapters.excel.macro_runner.MacroRunner`;
this service builds the STA-local callable, enforces policy, and normalizes the
outcome into a :class:`~excel_vba_mcp.domain.models.MacroRunResult`.

Safety contract (spec §7.3, §10; addendum BG03/BG04/BG05/BG06/BG09/R13):

* **Execute-mode only.** ``run_macro`` requires an execute-mode, writable session;
  inspect-mode (macros disabled) or read-only sessions are refused before any
  dispatch.
* **Deny-by-default approval.** ``ApprovalPolicy.check(key, EXECUTING)`` with
  ``key = canonical_path + sha256 + "macro_run" + procedure``. No grant/allowlist
  entry -> ``MACRO_BLOCKED`` (the macro is **not** run). This is the
  ``authorize_macro`` gate expressed through the shared
  :class:`~excel_vba_mcp.security.approval.ApprovalPolicy` public API.
* **background_safe gate (BG06).** A procedure is eligible only when it has been
  explicitly classified ``background_safe`` for the **current workbook hash**
  *and* the **current VBA source revision** (``project_source_hash``) and the
  **exact** ``Module.Procedure`` by :class:`BackgroundSafetyRegistry`. The registry
  is an in-memory, server-lifetime classification store keyed by the 3-tuple
  ``(workbook_file_sha256, project_source_hash, procedure.strip().lower())``. It
  starts empty and is fail-closed: an unclassified procedure -> ``None`` ->
  ``BACKGROUND_SAFETY_NOT_VERIFIED``; a procedure classified unsafe ->
  ``INTERACTIVE_UI_NOT_SUPPORTED``. A workbook hash change (file edit) **or** a
  VBA source change (``project_source_hash`` recomputed live from module code at
  every run/list/classify) invalidates the prior classification; cell-only edits
  do not.
* **UI preflight (R13) — a veto, never a certificate.** Before dispatch the
  procedure source is scanned (case-insensitively) for known interactive-UI tokens
  (``MsgBox`` / ``InputBox`` / ``.Show`` / ``FileDialog`` / ``Shell`` /
  ``CreateObject`` / ``FollowHyperlink`` / ``SendKeys`` / ``DoEvents`` /
  ``GetOpenFilename`` / ``GetSaveAsFilename`` / ``CallByName`` / ``ExecuteGlobal``
  / ``Execute`` / ``Eval`` / ``OnTime`` / ``DisplayAlerts`` / ``.Run`` /
  ``CommandBars``); any hit blocks the run with ``INTERACTIVE_UI_NOT_SUPPORTED``.
  A procedure whose inventory entry is ``is_heuristic`` (its body could not be
  enumerated precisely) is also blocked (fail-closed): it is never run as safe.
  **Residual boundary (R-15):** this token scan is a heuristic risk filter, not
  static proof. Dynamic/indirect UI (``CallByName`` / ``ExecuteGlobal`` / ``Eval``
  indirection) and UI reached transitively through other procedures the scan does
  not see are **not** fully covered; a ``background_safe`` classification is a
  human review outcome, and the deterministic layer can only **veto**, never
  **certify**. Office/Trust-Center macro-disable surfaces at runtime as
  ``Application.Run`` failing with ``MACRO_NOT_FOUND`` (not ``MACRO_BLOCKED``), so
  the macro error-distinction claim is deliberately not overstated.
* **No blind retry.** The ``Application.Run`` dispatch is submitted once
  (``section="execute"``, ``op_class=EXECUTING``); a timeout / COM-busy result is
  never replayed.
* **No live-pid auto-kill (C1 / Block E).** If the dispatch exceeds its soft
  deadline the run is classified ``SUSPECTED_STALLED`` and the owned Excel PID is
  **left running** — the server **never** auto-terminates a live pid on a
  deadline. The READ-ONLY owned-PID modal probe only enriches the diagnostic
  text; it never dismisses a dialog and never terminates anything. Recovery is
  **agent-directed** (``excel_recover``, Block A — not in this block); the macro
  is never reported as success while it is blocked and is never auto-replayed.
* **Hidden / non-activating.** All work runs on the owned ``Visible=False`` Excel
  instance; the service never activates/selects anything, never sends input, and
  never touches a user-owned Excel process.
"""

from __future__ import annotations

import contextlib
import hashlib
import threading
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from excel_vba_mcp.adapters.excel.app_adapter import ExcelAppAdapter
from excel_vba_mcp.adapters.excel.com_worker import ExcelComWorker
from excel_vba_mcp.adapters.excel.macro_runner import MacroRunner
from excel_vba_mcp.adapters.excel.vbide_adapter import VBIDEAdapter, hash_code
from excel_vba_mcp.application.session_service import SessionService
from excel_vba_mcp.domain.errors import (
    ErrorCode,
    ExcelHungError,
    ExcelVbaMcpError,
    StaleTargetError,
    WorkbookNotFoundError,
)
from excel_vba_mcp.domain.models import DomainModel, MacroRunResult, WorkbookInfo
from excel_vba_mcp.domain.statuses import (
    MacroRunStatus,
    OperationClass,
    ProcedureType,
    ProcedureVisibility,
    SessionMode,
    VbaComponentType,
)
from excel_vba_mcp.security.approval import ApprovalKey, ApprovalPolicy
from excel_vba_mcp.security.audit import AuditLogger

# Default per-call soft ceiling (seconds) for the Application.Run dispatch.
# This is the macro-dispatch deadline hint; it is env-configurable via
# ``EXCEL_VBA_MCP_MACRO_TIMEOUT_S`` (``ServerConfig.macro_timeout_s``). On expiry
# the run is classified ``SUSPECTED_STALLED`` and the owned pid is NOT terminated
# (Block E / C1: the server never auto-kills a live pid).
_DEFAULT_MACRO_TIMEOUT_S = 90.0

# Canonical action name for the macro-run approval key.
_ACTION_MACRO_RUN = "macro_run"

# Structured error codes for a blocked (did-not-run) result. These live in
# ``MacroRunResult.error_code`` alongside the ``ErrorCode`` values so a caller can
# tell a policy block apart from a background-safety / preflight block. The
# ``ErrorCode`` enum is the single source of truth; these aliases keep the literal
# string values that existing callers/tests assert on.
ERROR_BACKGROUND_SAFETY_NOT_VERIFIED = ErrorCode.BACKGROUND_SAFETY_NOT_VERIFIED.value
ERROR_INTERACTIVE_UI_NOT_SUPPORTED = ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED.value

# Tokens that flag a procedure as potentially interactive / not background-safe
# (addendum §2.4 / BG06 / R13). This is a heuristic risk filter, not static proof.
# Matching is case-insensitive (``_find_interactive_tokens``). The list keeps the
# original seven tokens and adds dynamic/indirect-UI and automation tokens so the
# deterministic R13 veto under-approximates UI risk less (R-15). It also adds
# window/focus and visibility tokens (``AppActivate`` / ``Application.Visible`` /
# ``.Activate`` / ``WindowState``) that a bare ``MsgBox``/``.Show`` scan misses; the
# scan stays a veto, not a certificate of non-interactivity.
INTERACTIVE_TOKENS: tuple[str, ...] = (
    "MsgBox",
    "InputBox",
    ".Show",
    "FileDialog",
    "Shell",
    "CreateObject",
    "FollowHyperlink",
    "SendKeys",
    "DoEvents",
    "GetOpenFilename",
    "GetSaveAsFilename",
    "CallByName",
    "ExecuteGlobal",
    "Execute",
    "Eval",
    "OnTime",
    "DisplayAlerts",
    ".Run",
    "CommandBars",
    "AppActivate",
    "Application.Visible",
    ".Activate",
    "WindowState",
)

# Distinct per-component source-hash sentinel for a module whose code could not be
# read (a COM failure), so an UNREADABLE module is never conflated with an EMPTY
# module (``hash_code("")``). A different sentinel value invalidates any
# classification derived while the module was readable.
_UNREADABLE_MODULE_SENTINEL = "<unreadable>"

# VBIDE modal dialog window class (used for READ-ONLY owned-modal detection).
_DIAGLG_CLASS = "#32770"

# Component types whose public procedures are callable macro candidates.
_CALLABLE_COMPONENT_TYPES = (
    VbaComponentType.STANDARD_MODULE,
    VbaComponentType.CLASS_MODULE,
)


def _utc_now() -> datetime:
    return datetime.now(UTC)


def _sha256_file(path: str) -> str | None:
    """Return the SHA-256 hex digest of ``path`` or ``None`` when absent."""

    target = Path(path)
    if not target.is_file():
        return None
    digest = hashlib.sha256()
    with target.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# DTOs / registry (no COM)
# ---------------------------------------------------------------------------
class MacroInfo(DomainModel):
    """One callable macro candidate from the VBA project inventory (spec §6.3).

    ``name`` is ``Module.Procedure`` (the ``Application.Run`` procedure part).
    ``background_safe`` reflects the :class:`BackgroundSafetyRegistry` for the
    current workbook hash **and** ``project_source_hash``; it is ``False`` when the
    procedure has not yet been classified (deny-by-default), in which case
    ``background_safe_classified`` is ``False``. ``project_source_hash`` is the
    live VBA source-revision hash the classification is bound to (annotated by the
    service on every ``list_macros`` call); an empty default keeps existing
    constructors working.
    """

    name: str
    module: str
    procedure: str
    kind: ProcedureType
    visibility: ProcedureVisibility
    background_safe: bool = False
    background_safe_classified: bool = False
    project_source_hash: str = ""


class BackgroundSafetyClassification(DomainModel):
    """Outcome of classifying a procedure's ``background_safe`` state (fix_005).

    Returned by :meth:`MacroService.classify_background` (MCP
    ``macro_classify_background``). It records the workbook identity
    (``workbook_sha256``) and the VBA ``project_source_hash`` the classification is
    bound to, so a later run/list recomputes the same hashes and the classification
    only holds while both are unchanged.
    """

    workbook_id: str
    procedure: str
    background_safe: bool
    workbook_sha256: str
    project_source_hash: str
    classified_at: datetime


class BackgroundSafetyRegistry:
    """Per-(workbook-hash, source-revision) ``background_safe`` registry (BG06).

    A procedure is eligible for strict background execution only when it has been
    explicitly classified ``background_safe`` for the **current** workbook hash,
    the **current** VBA source revision (``project_source_hash``) and the **exact**
    ``Module.Procedure``. The classification is bound to both the workbook content
    hash and the live source hash so a source/hash change invalidates it (the key no
    longer matches). The registry is deny-by-default: an unknown key returns
    ``None`` (unclassified -> not eligible).

    The registry is in-memory, server-lifetime bookkeeping (not disk-persisted), not
    a security boundary; it records a reviewed classification (via the
    ``macro_classify_background`` MCP tool) that a procedure's non-interactive
    behavior was verified on a disposable copy. It is keyed by the 3-tuple
    ``(workbook_sha256, project_source_hash, procedure.strip().lower())``.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        # (workbook_sha256, project_source_hash, "module.procedure" lowercased)
        # -> background_safe
        self._classified: dict[tuple[str, str, str], bool] = {}

    @staticmethod
    def _key(
        workbook_sha256: str, project_source_hash: str, procedure: str
    ) -> tuple[str, str, str]:
        return (workbook_sha256, project_source_hash, procedure.strip().lower())

    def classify(
        self,
        workbook_sha256: str,
        project_source_hash: str,
        procedure: str,
        *,
        background_safe: bool,
    ) -> None:
        """Record whether ``procedure`` is ``background_safe`` for a hash + revision."""

        with self._lock:
            self._classified[self._key(workbook_sha256, project_source_hash, procedure)] = (
                background_safe
            )

    def is_safe(
        self, workbook_sha256: str, project_source_hash: str, procedure: str
    ) -> bool | None:
        """Return the classification, or ``None`` when unclassified (deny-by-default)."""

        with self._lock:
            return self._classified.get(self._key(workbook_sha256, project_source_hash, procedure))

    def clear(self) -> None:
        """Drop all classifications (e.g. on session teardown)."""

        with self._lock:
            self._classified.clear()


# ---------------------------------------------------------------------------
# Service
# ---------------------------------------------------------------------------
class MacroService:
    """Policy-controlled macro list / run / last-result orchestration.

    Args:
        worker: The single-STA COM worker (all Excel/VBIDE access runs there).
        session_service: The workbook session lifecycle service (owns the
            ``workbook_id`` bookkeeping and the owned app adapter).
        approval: Deny-by-default approval gate for the ``macro_run`` action.
        background_safety: Per-workbook-hash ``background_safe`` registry (BG06).
        audit: Append-only JSONL audit trail.
        macro_timeout_s: Soft per-call ceiling (seconds) for the ``Application.Run``
            dispatch. Defaults to :data:`_DEFAULT_MACRO_TIMEOUT_S` (90s); the server
            threads ``ServerConfig.macro_timeout_s`` here.
    """

    def __init__(
        self,
        worker: ExcelComWorker,
        session_service: SessionService,
        approval: ApprovalPolicy,
        background_safety: BackgroundSafetyRegistry,
        audit: AuditLogger,
        *,
        macro_timeout_s: float = _DEFAULT_MACRO_TIMEOUT_S,
    ) -> None:
        self._worker = worker
        self._session = session_service
        self._approval = approval
        self._background = background_safety
        self._audit = audit
        self._macro_timeout_s = macro_timeout_s
        self._last: dict[str, MacroRunResult] = {}

    # ------------------------------------------------------------------
    # Public API (async; no COM here — all COM on the STA worker)
    # ------------------------------------------------------------------
    async def list_macros(self, workbook_id: str) -> list[MacroInfo]:
        """List the callable macro candidates of ``workbook_id`` (public procedures).

        Candidates are the public ``Sub`` / ``Function`` procedures of standard and
        class modules (the procedures ``Application.Run`` can target by
        ``Module.Procedure``). Each is annotated with its ``background_safe``
        classification for the current workbook hash (deny-by-default).

        Args:
            workbook_id: The id returned by :func:`excel_open`.

        Returns:
            One :class:`MacroInfo` per callable public procedure.

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown or the project cannot
                be read (e.g. VBOM disabled / locked).
        """

        record = self._resolve_record(workbook_id)
        workbook_hash = _sha256_file(record.path) or ""
        inventory, project_source_hash = await self._collect_project_state(workbook_id, record)

        macros: list[MacroInfo] = []
        for entry in inventory:
            if entry["visibility"] != ProcedureVisibility.PUBLIC.value:
                continue
            qualified = entry["qualified"]
            classified = self._background.is_safe(workbook_hash, project_source_hash, qualified)
            macros.append(
                MacroInfo(
                    name=qualified,
                    module=entry["module"],
                    procedure=entry["procedure"],
                    kind=ProcedureType(entry["kind"]),
                    visibility=ProcedureVisibility(entry["visibility"]),
                    background_safe=bool(classified),
                    background_safe_classified=classified is not None,
                    project_source_hash=project_source_hash,
                )
            )

        self._record_audit(
            tool="macro_list",
            operation_class="read_only",
            workbook_id=workbook_id,
            record=record,
            workbook_hash=workbook_hash,
            params={"workbook_id": workbook_id},
            approval=None,
            result={"count": len(macros)},
            extra=None,
        )
        return macros

    async def run_macro(
        self,
        workbook_id: str,
        macro_name: str,
        args: list[Any] | None = None,
        *,
        timeout_s: float | None = None,
    ) -> MacroRunResult:
        """Run one approved, background-safe macro of ``workbook_id``.

        Enforces (in order): execute-mode writable session -> resolve the procedure
        -> deny-by-default approval -> background_safe gate (BG06) -> interactive-UI
        preflight (R13) -> a single ``Application.Run`` dispatch on the STA worker.
        A blocked macro is **not** dispatched; a timed-out / modal-hung macro is
        classified honestly and never reported as success.

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            macro_name: The procedure to run, qualified as ``Module.Procedure`` or
                a bare ``Procedure`` name (resolved against the inventory).
            args: Positional arguments (coerced/validated by the runner).
            timeout_s: Soft per-call ceiling for the ``Application.Run`` dispatch.
                Defaults to the configured ``macro_timeout_s`` (90s). A server-side
                soft hint only (no client contract).

        Returns:
            A :class:`MacroRunResult`. ``status`` is one of ``SUCCESS`` /
            ``NOT_FOUND`` / ``RUNTIME_ERROR`` / ``BLOCKED`` / ``SUSPECTED_STALLED``
            (deadline exceeded; the owned pid is left running, never terminated).

        Raises:
            ExcelVbaMcpError: If ``workbook_id`` is unknown.
        """

        effective_timeout_s = self._macro_timeout_s if timeout_s is None else timeout_s
        started_at = _utc_now()
        run_id = f"run-{uuid4().hex}"
        record = self._resolve_record(workbook_id)

        # 1. Execute-mode writable context.
        self._require_execute_context(record)
        workbook_hash = _sha256_file(record.path) or ""

        # 2. Resolve the procedure from the inventory (fail closed on read errors).
        try:
            inventory, project_source_hash = await self._collect_project_state(workbook_id, record)
        except ExcelVbaMcpError as exc:
            # A timeout / lost-Excel while reading the inventory is a structured
            # error that already moved the worker into needs_recovery: report it
            # truthfully (NEEDS_RECOVERY) and never a false success.
            needs_recovery = exc.code in (
                ErrorCode.EXCEL_HUNG,
                ErrorCode.COM_BUSY,
                ErrorCode.RPC_E_CALL_REJECTED,
            )
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=macro_name,
                args=args,
                started_at=started_at,
                status=(
                    MacroRunStatus.NEEDS_RECOVERY if needs_recovery else MacroRunStatus.BLOCKED
                ),
                error_code=exc.code.value,
                error_text=f"Macro inventory unavailable; refusing to run: {exc}",
                approval=None,
                extra={
                    "reason": "inventory_unavailable",
                    "needs_recovery": needs_recovery,
                },
            )

        entry = _resolve_procedure(inventory, macro_name)
        if entry is None:
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=macro_name,
                args=args,
                started_at=started_at,
                status=MacroRunStatus.NOT_FOUND,
                error_code=ErrorCode.MACRO_NOT_FOUND.value,
                error_text=f"Macro {macro_name!r} was not found as a public procedure.",
                approval=None,
                extra={"reason": "macro_not_found"},
            )

        qualified = entry["qualified"]

        # 3. Deny-by-default approval (key = canonical path + sha256 + action + procedure).
        approval_key = ApprovalKey(
            canonical_path=record.source_path or record.path,
            sha256=workbook_hash or None,
            action=_ACTION_MACRO_RUN,
            procedure=qualified,
        )
        decision = self._approval.check(
            approval_key,
            session_id=self._session.session_id,
            operation_class=OperationClass.EXECUTING,
        )
        if not decision.allowed:
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=qualified,
                args=args,
                started_at=started_at,
                status=MacroRunStatus.BLOCKED,
                error_code=ErrorCode.APPROVAL_DENIED.value,
                error_text=f"Macro run denied by approval policy: {decision.reason}",
                approval=decision,
                extra={"reason": "approval_denied"},
            )

        # 4. background_safe gate (BG06) — bound to the live source revision.
        classified = self._background.is_safe(workbook_hash, project_source_hash, qualified)
        if classified is None:
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=qualified,
                args=args,
                started_at=started_at,
                status=MacroRunStatus.BLOCKED,
                error_code=ERROR_BACKGROUND_SAFETY_NOT_VERIFIED,
                error_text=(
                    f"Macro {qualified!r} is not classified background_safe for this "
                    "workbook hash; refusing to run in background mode (BG06)."
                ),
                approval=decision,
                extra={"reason": "background_safety_not_verified", "background_safe": None},
            )
        if not classified:
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=qualified,
                args=args,
                started_at=started_at,
                status=MacroRunStatus.BLOCKED,
                error_code=ERROR_INTERACTIVE_UI_NOT_SUPPORTED,
                error_text=(
                    f"Macro {qualified!r} is classified NOT background_safe; "
                    "refusing to run in background mode (BG06)."
                ),
                approval=decision,
                extra={"reason": "background_unsafe", "background_safe": False},
            )

        # 5. Interactive-UI preflight (R13) — heuristic, best-effort veto. Fail
        # closed on a heuristic (imprecisely enumerated) procedure: it is never run
        # as background-safe.
        is_heuristic = entry.get("is_heuristic") == "true"
        flagged = _find_interactive_tokens(entry.get("code", ""))
        if is_heuristic:
            error_text = (
                f"Macro {qualified!r} could not be enumerated precisely "
                "(is_heuristic); refusing to run it as background_safe (R13)."
            )
            preflight_extra: dict[str, Any] = {
                "reason": "heuristic_procedure",
                "is_heuristic": True,
            }
        elif flagged:
            error_text = (
                f"Macro {qualified!r} references interactive UI "
                f"({', '.join(flagged)}); refusing to run in background mode (R13)."
            )
            preflight_extra = {"reason": "interactive_ui_preflight", "flagged_tokens": flagged}
        else:
            preflight_extra = {}
        if is_heuristic or flagged:
            return self._finish_blocked(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=qualified,
                args=args,
                started_at=started_at,
                status=MacroRunStatus.BLOCKED,
                error_code=ERROR_INTERACTIVE_UI_NOT_SUPPORTED,
                error_text=error_text,
                approval=decision,
                extra=preflight_extra,
            )

        # 6. Single Application.Run dispatch on the STA worker (execute/EXECUTING).
        app_adapter = self._app_adapter()
        # Workbook-qualified target so Application.Run resolves to THIS workbook even
        # when the instance holds several open workbooks (research §4.1).
        run_target = f"'{record.name}'!{qualified}"

        def run_fn() -> MacroRunResult:
            app = app_adapter.app
            if app is None:
                # The owned instance is gone (crash / lost Excel): a structured
                # EXCEL_HUNG that requires recovery, never a false success.
                raise ExcelHungError(
                    "The owned Excel instance is not available.",
                    context={
                        "workbook_id": workbook_id,
                        "needs_recovery": True,
                    },
                )
            if _find_open_workbook(app_adapter, record.path) is None:
                raise ExcelHungError(
                    "The working copy is no longer open in the owned instance.",
                    context={
                        "workbook_id": workbook_id,
                        "working_path": record.path,
                        "needs_recovery": True,
                    },
                )
            return MacroRunner(app).run(run_target, list(args or ()))

        run_fn.__name__ = "macro_run"
        try:
            result = await self._worker.submit(
                run_fn,
                timeout=effective_timeout_s,
                section="execute",
                op_class=OperationClass.EXECUTING,
            )
        except ExcelHungError:
            result = await self._handle_hung(
                workbook_id=workbook_id,
                record=record,
                workbook_hash=workbook_hash,
                run_id=run_id,
                macro_name=qualified,
                args=args,
                started_at=started_at,
                approval=decision,
            )
            self._last[workbook_id] = result
            return result

        self._last[workbook_id] = result
        self._record_audit(
            tool="macro_run",
            operation_class="executing",
            workbook_id=workbook_id,
            record=record,
            workbook_hash=workbook_hash,
            params={"workbook_id": workbook_id, "macro": qualified, "args": list(args or ())},
            approval=decision,
            result=_result_audit(
                result, background_safe=True, dismissed=None, needs_recovery=False
            ),
            extra=None,
        )
        return result

    async def get_last_result(self, workbook_id: str) -> MacroRunResult | None:
        """Return the cached last :class:`MacroRunResult` for ``workbook_id``.

        Never re-runs the macro and never touches COM.

        Args:
            workbook_id: The id returned by :func:`excel_open`.

        Returns:
            The cached result, or ``None`` when no macro has been run yet.
        """

        return self._last.get(workbook_id)

    # ------------------------------------------------------------------
    # Internal helpers (non-COM orchestration)
    # ------------------------------------------------------------------
    def _require_execute_context(self, record: WorkbookInfo) -> None:
        """Reject a session that is not execute-mode and writable."""

        if record.mode != SessionMode.EXECUTE:
            raise ExcelVbaMcpError(
                "Macro execution requires an execute-mode session; macros are "
                "disabled in inspect mode.",
                code=ErrorCode.MACRO_BLOCKED,
                context={"workbook_id": record.workbook_id, "mode": record.mode.value},
            )
        if record.read_only:
            raise ExcelVbaMcpError(
                "Macro execution requires a writable working copy.",
                code=ErrorCode.WORKBOOK_READ_ONLY,
                context={"workbook_id": record.workbook_id, "path": record.path},
            )

    async def _collect_project_state(
        self, workbook_id: str, record: WorkbookInfo
    ) -> tuple[list[dict[str, str]], str]:
        """Read the procedure inventory AND the live ``project_source_hash``.

        This is the SINGLE source-of-truth scan used identically by
        :meth:`list_macros`, :meth:`run_macro` and :meth:`classify_background`: one
        READ_ONLY STA submit returns both the callable inventory and the
        ``project_source_hash`` (a deterministic hash over every VBA component's
        source, including Document modules and UserForm code-behind). Deriving both
        from one scan guarantees classify / run / list never disagree on the hash a
        classification is bound to.
        """

        app_adapter = self._app_adapter()

        def collect_fn() -> tuple[list[dict[str, str]], str]:
            if _find_open_workbook(app_adapter, record.path) is None:
                raise ExcelHungError(
                    "The working copy is no longer open in the owned instance.",
                    context={
                        "workbook_id": workbook_id,
                        "working_path": record.path,
                        "needs_recovery": True,
                    },
                )
            return _scan_project(app_adapter, record.path)

        collect_fn.__name__ = "macro_inventory"
        raw = await self._worker.submit(
            collect_fn,
            timeout=_DEFAULT_MACRO_TIMEOUT_S,
            section="diagnostic",
            op_class=OperationClass.READ_ONLY,
        )
        inventory, project_source_hash = raw
        return list(inventory), str(project_source_hash)

    async def classify_background(
        self,
        workbook_id: str,
        procedure: str,
        background_safe: bool,
        *,
        expected_project_source_hash: str | None = None,
    ) -> BackgroundSafetyClassification:
        """Classify a procedure's ``background_safe`` for the current source revision.

        This is the bookkeeping behind the ``macro_classify_background`` MCP tool
        (approach A). It is UNGATED (a review outcome, not a security boundary —
        :meth:`run_macro` keeps the R13 veto). Steps:

        1. resolve the record and read the inventory + live ``project_source_hash``
           via the SAME single scan path as run/list;
        2. resolve ``procedure`` (``MACRO_NOT_FOUND`` when absent/ambiguous);
        3. when ``expected_project_source_hash`` is supplied and differs from the
           live hash, raise ``StaleTargetError`` (``STALE_TARGET``) and DO NOT write;
        4. when ``background_safe`` is requested for a heuristic procedure or one
           whose body matches an interactive-UI token, raise
           ``INTERACTIVE_UI_NOT_SUPPORTED`` and DO NOT write (the deterministic layer
           can only veto);
        5. record the classification in the registry and audit it.

        Args:
            workbook_id: The id returned by :func:`excel_open`.
            procedure: The procedure to classify (``Module.Procedure`` or a bare
                unique ``Procedure`` name).
            background_safe: The classification to record.
            expected_project_source_hash: Optional optimistic-concurrency guard; when
                given and it no longer matches the live source hash, the write is
                rejected with ``STALE_TARGET``.

        Returns:
            A :class:`BackgroundSafetyClassification` DTO.

        Raises:
            ExcelVbaMcpError: ``MACRO_NOT_FOUND`` (unknown/ambiguous procedure) or
                ``INTERACTIVE_UI_NOT_SUPPORTED`` (heuristic / interactive-UI).
            StaleTargetError: ``expected_project_source_hash`` no longer matches.
        """

        record = self._resolve_record(workbook_id)
        workbook_hash = _sha256_file(record.path) or ""
        inventory, project_source_hash = await self._collect_project_state(workbook_id, record)

        entry = _resolve_procedure(inventory, procedure)
        if entry is None:
            raise ExcelVbaMcpError(
                f"Macro {procedure!r} was not found as a public procedure.",
                code=ErrorCode.MACRO_NOT_FOUND,
                context={"workbook_id": workbook_id, "procedure": procedure},
            )

        qualified = entry["qualified"]
        if (
            expected_project_source_hash is not None
            and expected_project_source_hash != project_source_hash
        ):
            raise StaleTargetError(
                "The VBA source changed since it was read; re-read and retry.",
                context={"workbook_id": workbook_id, "procedure": qualified},
            )

        is_heuristic = entry.get("is_heuristic") == "true"
        if background_safe and (is_heuristic or _find_interactive_tokens(entry.get("code", ""))):
            raise ExcelVbaMcpError(
                f"Macro {qualified!r} references interactive UI or is heuristic; "
                "it cannot be classified background_safe (R13 veto).",
                code=ErrorCode.INTERACTIVE_UI_NOT_SUPPORTED,
                context={"workbook_id": workbook_id, "procedure": qualified},
            )

        self._background.classify(
            workbook_hash,
            project_source_hash,
            qualified,
            background_safe=background_safe,
        )
        self._record_audit(
            tool="macro_classify_background",
            operation_class="config",
            workbook_id=workbook_id,
            record=record,
            workbook_hash=workbook_hash,
            params={
                "workbook_id": workbook_id,
                "procedure": qualified,
                "background_safe": background_safe,
            },
            approval=None,
            result={
                "background_safe": background_safe,
                "project_source_hash": project_source_hash,
            },
            extra=None,
        )
        return BackgroundSafetyClassification(
            workbook_id=workbook_id,
            procedure=qualified,
            background_safe=background_safe,
            workbook_sha256=workbook_hash,
            project_source_hash=project_source_hash,
            classified_at=_utc_now(),
        )

    async def _handle_hung(
        self,
        *,
        workbook_id: str,
        record: WorkbookInfo,
        workbook_hash: str,
        run_id: str,
        macro_name: str,
        args: list[Any] | None,
        started_at: datetime,
        approval: Any,
    ) -> MacroRunResult:
        """Classify a dispatch that exceeded its soft deadline (no kill, no auto-recovery).

        Block E / C1: the server **never** auto-terminates a live pid. On a deadline
        the owned Excel PID is **left running** and the operation is **not**
        cancelled (an uninterruptible COM call cannot be); recovery is
        **agent-directed** (``excel_recover``, Block A — not in this block). The
        READ-ONLY owned-PID modal probe only enriches the diagnostic text — it never
        dismisses a dialog and never terminates anything.

        The result is ``SUSPECTED_STALLED`` (``error_code`` ``MACRO_TIMEOUT``); it is
        never reported as success and never retried. ``session_recovery`` is ``None``
        because no recovery is attempted here.
        """

        owned_pid = self._worker.get_status().owned_pid
        modal_present = _owned_modal_present(owned_pid)

        finished_at = _utc_now()
        if modal_present:
            error_text = (
                "Macro exceeded the soft deadline (SUSPECTED_STALLED) and a visible "
                "modal dialog is present on the owned Excel process. The owned Excel "
                "process was NOT terminated and the dialog was NOT dismissed "
                "(dismissed=false); the operation is left running. Recovery is "
                "agent-directed (excel_recover); the macro is never auto-replayed."
            )
        else:
            error_text = (
                "Macro exceeded the soft deadline (SUSPECTED_STALLED). The owned Excel "
                "process was NOT terminated and the operation is left running. Recovery "
                "is agent-directed (excel_recover); the macro is never auto-replayed."
            )
        result = MacroRunResult(
            run_id=run_id,
            macro_name=macro_name,
            status=MacroRunStatus.SUSPECTED_STALLED,
            started_at=started_at,
            finished_at=finished_at,
            duration_ms=int(round((finished_at - started_at).total_seconds() * 1000.0)),
            error_code=ErrorCode.MACRO_TIMEOUT.value,
            error_text=error_text,
            session_recovery=None,
            created_at=started_at,
            updated_at=finished_at,
        )
        extra: dict[str, Any] = {
            "reason": "deadline_exceeded",
            "verdict": "SUSPECTED_STALLED",
            "dismissed": False if modal_present else None,
            "needs_recovery": False,
            "owned_pid": owned_pid,
            "owned_pid_terminated": False,
            "modal_present": modal_present,
        }

        self._record_audit(
            tool="macro_run",
            operation_class="executing",
            workbook_id=workbook_id,
            record=record,
            workbook_hash=workbook_hash,
            params={"workbook_id": workbook_id, "macro": macro_name, "args": list(args or ())},
            approval=approval,
            result=_result_audit(
                result,
                background_safe=True,
                dismissed=extra.get("dismissed"),
                needs_recovery=False,
                extra=extra,
            ),
            extra=None,
        )
        return result

    def _finish_blocked(
        self,
        *,
        workbook_id: str,
        record: WorkbookInfo,
        workbook_hash: str,
        run_id: str,
        macro_name: str,
        args: list[Any] | None,
        started_at: datetime,
        status: MacroRunStatus,
        error_code: str,
        error_text: str,
        approval: Any,
        extra: dict[str, Any] | None,
    ) -> MacroRunResult:
        """Build, audit, cache and return a blocked/not-found result (never dispatched)."""

        finished_at = _utc_now()
        result = MacroRunResult(
            run_id=run_id,
            macro_name=macro_name,
            status=status,
            started_at=started_at,
            finished_at=finished_at,
            duration_ms=int(round((finished_at - started_at).total_seconds() * 1000.0)),
            error_code=error_code,
            error_text=error_text,
            created_at=started_at,
            updated_at=finished_at,
        )
        self._last[workbook_id] = result
        self._record_audit(
            tool="macro_run",
            operation_class="executing",
            workbook_id=workbook_id,
            record=record,
            workbook_hash=workbook_hash,
            params={"workbook_id": workbook_id, "macro": macro_name, "args": list(args or ())},
            approval=approval,
            result=_result_audit(
                result, background_safe=None, dismissed=None, needs_recovery=False
            ),
            extra=extra,
        )
        return result

    def _resolve_record(self, workbook_id: str) -> WorkbookInfo:
        """Resolve ``workbook_id`` to its open :class:`WorkbookInfo` (no COM)."""

        for info in self._session.list_workbooks_sync():
            if info.workbook_id == workbook_id:
                return info
        raise WorkbookNotFoundError(
            f"Unknown workbook_id {workbook_id!r} (not open).",
            context={"workbook_id": workbook_id},
        )

    def _app_adapter(self) -> ExcelAppAdapter:
        """Return the owned hidden Excel application adapter (never a new one)."""

        adapter: ExcelAppAdapter | None = getattr(self._session, "_app_adapter", None)
        if adapter is None:  # pragma: no cover - defensive
            raise ExcelVbaMcpError(
                "The session service does not expose an Excel application adapter.",
                context={"tool": "macro_run"},
            )
        return adapter

    def _record_audit(
        self,
        *,
        tool: str,
        operation_class: str,
        workbook_id: str,
        record: WorkbookInfo,
        workbook_hash: str,
        params: dict[str, Any],
        approval: Any,
        result: dict[str, Any],
        extra: dict[str, Any] | None,
    ) -> None:
        """Record one audit entry (macro identity, workbook hash, approval, result)."""

        approval_payload: dict[str, Any] | None = None
        if approval is not None:
            approval_payload = {
                "required": True,
                "decision": "allow" if getattr(approval, "allowed", False) else "deny",
                "source": getattr(approval.source, "value", None)
                if getattr(approval, "source", None) is not None
                else None,
                "grant_id": getattr(approval, "grant_id", None),
            }
        if extra:
            result = {**result, **extra}
        self._audit.record(
            tool=tool,
            operation_class=operation_class,
            params=params,
            session_id=getattr(self._session, "session_id", None),
            workbook={"canonical_path": record.source_path or record.path, "sha256": workbook_hash},
            approval=approval_payload,
            result=result,
        )


# ---------------------------------------------------------------------------
# Module-level helpers (no COM at import time)
# ---------------------------------------------------------------------------
def _find_open_workbook(app_adapter: ExcelAppAdapter, working_path: str) -> Any | None:
    """Return the open ``Workbook`` proxy matching ``working_path`` (STA thread)."""

    app = app_adapter.app
    if app is None:
        return None
    target = working_path.lower()
    try:
        for wb in app.Workbooks:
            try:
                if str(wb.FullName).lower() == target:
                    return wb
            except Exception:  # noqa: BLE001 - defensive per-item COM read
                continue
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return None


def _scan_project(
    app_adapter: ExcelAppAdapter, working_path: str
) -> tuple[list[dict[str, str]], str]:
    """Single STA scan: return ``(callable_inventory, project_source_hash)``.

    Iterates **every** ``VBComponent`` (standard/class **and** Document modules /
    UserForm code-behind) once, computing a per-component source hash for the
    project hash, and — for the callable standard/class components — the procedure
    inventory. Both values come from one pass so :meth:`MacroService.run_macro`,
    :meth:`MacroService.list_macros` and :meth:`MacroService.classify_background`
    always derive the identical ``project_source_hash`` a classification is bound
    to (no path-dependent mismatch).

    The callable inventory is a JSON-serializable list of dicts (``module`` /
    ``procedure`` / ``qualified`` / ``kind`` / ``visibility`` / ``code`` /
    ``is_heuristic``). The per-procedure ``code`` used by the UI preflight is
    extracted **precisely** from the module text using the procedure's
    ``start_line`` / ``end_line``. ``VBIDEAdapter`` reads ``CodeModule`` with an
    absolute end-line-as-count that can over-read into later procedures in the same
    module; recomputing the exact body here keeps the preflight scoped to the single
    procedure being considered.
    """

    app = app_adapter.app
    wb = _find_open_workbook(app_adapter, working_path)
    if app is None or wb is None:
        raise WorkbookNotFoundError(
            "The working copy is no longer open in the owned instance.",
            context={"working_path": working_path},
        )

    adapter = VBIDEAdapter()
    project = wb.VBProject
    inventory: list[dict[str, str]] = []
    component_hashes: dict[str, str] = {}
    for component in adapter.list_components(wb):
        module_lines = _read_module_lines(project, component.component_name)
        # Every component (standard/class + Document + UserForm code-behind) feeds
        # the project hash so a doc/form edit invalidates a classification made from
        # a hashed standard/class procedure (MUST-fix 4).
        component_hashes[component.component_name] = _hash_module_lines(module_lines)
        if component.component_type not in _CALLABLE_COMPONENT_TYPES:
            continue
        lines = module_lines or []
        for procedure in adapter.list_procedures(wb, component.component_name):
            inventory.append(
                {
                    "module": procedure.component_name,
                    "procedure": procedure.name,
                    "qualified": f"{procedure.component_name}.{procedure.name}",
                    "kind": procedure.procedure_type.value,
                    "visibility": procedure.visibility.value,
                    "code": _procedure_body(lines, procedure.start_line, procedure.end_line)
                    or procedure.code,
                    "is_heuristic": "true" if procedure.is_heuristic else "false",
                }
            )
    return inventory, _compute_project_source_hash(component_hashes)


def _read_module_lines(project: Any, component_name: str) -> list[str] | None:
    """Read a module's full text as lines (STA thread).

    Distinguishes three states so an empty module is never conflated with an
    unreadable one (MUST-fix 2):

    * ``None`` — the module could not be read (a COM failure / VBOM issue);
    * ``[]`` — the module was read and is genuinely empty;
    * a non-empty ``list[str]`` — the module's lines.
    """

    try:
        code_module = project.VBComponents(component_name).CodeModule
        total = int(code_module.CountOfLines)
        if total <= 0:
            return []
        text = str(code_module.Lines(1, total))
    except Exception:  # noqa: BLE001 - defensive COM read
        return None
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def _hash_module_lines(module_lines: list[str] | None) -> str:
    """Per-component source hash: empty -> ``hash_code("")``; unreadable -> sentinel."""

    if module_lines is None:
        return _UNREADABLE_MODULE_SENTINEL
    return hash_code("\n".join(module_lines))


def _compute_project_source_hash(component_hashes: dict[str, str]) -> str:
    """Deterministic SHA-256 over every component's per-component source hash.

    The single source-of-truth helper used identically by ``classify_background``,
    ``run_macro`` and ``list_macros`` (via :func:`_scan_project`). Components are
    combined in sorted-name order with explicit separators so the digest is stable
    across runs and independent of enumeration order. Includes Document modules and
    UserForm code-behind; an unreadable component contributes its distinct sentinel
    (never the empty-module hash), so any read failure invalidates the digest.
    """

    combined = hashlib.sha256()
    for name in sorted(component_hashes):
        combined.update(name.encode("utf-8"))
        combined.update(b"\x00")
        combined.update(component_hashes[name].encode("utf-8"))
        combined.update(b"\n")
    return combined.hexdigest()


def _procedure_body(module_lines: list[str], start_line: int, end_line: int) -> str:
    """Extract the exact procedure body (1-based, inclusive) from ``module_lines``."""

    if not module_lines or start_line < 1 or end_line < start_line:
        return ""
    body = module_lines[start_line - 1 : end_line]
    return "\n".join(body).strip("\n")


def _resolve_procedure(inventory: list[dict[str, str]], macro_name: str) -> dict[str, str] | None:
    """Resolve ``macro_name`` to an inventory entry (qualified or bare, case-insensitive)."""

    target = macro_name.strip().lower()
    # Exact qualified match first.
    for entry in inventory:
        if entry["qualified"].lower() == target:
            return entry
    # Bare procedure-name match (must be unique among public procedures).
    matches = [entry for entry in inventory if entry["procedure"].lower() == target]
    if len(matches) == 1:
        return matches[0]
    return None


def _find_interactive_tokens(source: str) -> list[str]:
    """Return the interactive-UI tokens present in ``source`` (heuristic, BG06/R13).

    Matching is case-insensitive (VBA identifiers are case-insensitive), so a
    lowercased ``msgbox`` / ``.show`` is still caught (MUST-fix 1a). This is a risk
    filter / veto, not static proof of non-interactivity.
    """

    if not source:
        return []
    folded = source.casefold()
    return [token for token in INTERACTIVE_TOKENS if token.casefold() in folded]


def _result_audit(
    result: MacroRunResult,
    *,
    background_safe: bool | None,
    dismissed: bool | None,
    needs_recovery: bool,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build the audit ``result`` payload for a macro run."""

    payload: dict[str, Any] = {
        "run_id": result.run_id,
        "status": result.status.value,
        "error_code": result.error_code,
        "error_text": result.error_text,
        "duration_ms": result.duration_ms,
        "background_safe": background_safe,
        "dismissed": dismissed,
        "needs_recovery": needs_recovery,
    }
    if extra:
        payload.update(extra)
    return payload


# ---------------------------------------------------------------------------
# Owned-process modal detection + termination (READ-ONLY inspect; owned PID only)
# ---------------------------------------------------------------------------
def _owned_modal_present(pid: int | None) -> bool:
    """Return whether the owned PID has a visible ``#32770`` modal (READ-ONLY).

    Enumerates top-level windows of ``pid`` only and never activates, focuses,
    clicks, types or dismisses anything. Returns ``False`` when pywin32 is
    unavailable (non-Windows) or no owned modal is found.
    """

    if pid is None:
        return False
    try:  # pragma: no cover - platform specific
        import win32gui  # type: ignore[import-untyped]  # noqa: PLC0415
        import win32process  # type: ignore[import-untyped]  # noqa: PLC0415
    except Exception:  # noqa: BLE001 - non-Windows / pywin32 absent
        return False

    found = False

    def _callback(hwnd: int, _ctx: Any) -> bool:
        nonlocal found
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid != pid:
                return True
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != _DIAGLG_CLASS:
                return True
            found = True
        except Exception:  # noqa: BLE001 - defensive per-window read
            return True
        return True

    with contextlib.suppress(Exception):
        win32gui.EnumWindows(_callback, None)
    return found


__all__ = [
    "ERROR_BACKGROUND_SAFETY_NOT_VERIFIED",
    "ERROR_INTERACTIVE_UI_NOT_SUPPORTED",
    "INTERACTIVE_TOKENS",
    "BackgroundSafetyClassification",
    "BackgroundSafetyRegistry",
    "MacroInfo",
    "MacroService",
]
