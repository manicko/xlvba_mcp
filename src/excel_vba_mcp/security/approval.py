"""Deny-by-default approval and allowlist policy (spec §7.1, §7.3, §10).

The approval gate decides whether a ``MUTATING`` or ``EXECUTING`` operation may
proceed. ``READ_ONLY`` operations are allowed by the caller (or auto-allowed by
:meth:`ApprovalPolicy.check` when the operation class is supplied). Anything else
is **denied unless** an explicit allowlist entry or a live grant matches.

Design (research task_009 §2):

* Two allow layers share one :class:`ApprovalKey`: a static **allowlist**
  (config-supplied) and dynamic **grants** (one-time / session, supplied by the
  MCP client at runtime). There is no interactive prompt (spec §10: STDIO stays
  non-blocking).
* **Identity rule (anti-bypass):** an approval key is matched on ``sha256`` when
  *both* sides carry it (content hash is authoritative — a different file with the
  same basename never matches); otherwise it falls back to ``canonical_path``.
  Callers **must** pass a path produced by
  :meth:`excel_vba_mcp.security.path_policy.PathPolicy.validate_path` (symlink /
  reparse / traversal resolved). This module never re-canonicalizes raw strings —
  path-alias equivalence is guaranteed by that upstream contract.
* **Deny-by-default:** the terminal branch of :meth:`ApprovalPolicy.check` returns
  ``allowed=False``.

Grant lifetimes:

* ``ONE_TIME`` — consumed on first successful use (removed from the store).
* ``SESSION`` — valid until ``expires_at`` (if set) or until the session changes;
  a session grant may be bound to a ``session_id`` at grant time and then only
  matches ``check(..., session_id=...)`` for that same session.
"""

from __future__ import annotations

import threading
from collections.abc import Callable, Iterable
from datetime import UTC, datetime
from enum import StrEnum

from excel_vba_mcp.domain.models import DomainModel
from excel_vba_mcp.domain.statuses import AccessMode, OperationClass


class ApprovalScope(StrEnum):
    """Lifetime of an :class:`ApprovalGrant`."""

    ONE_TIME = "one_time"
    SESSION = "session"


class ApprovalSource(StrEnum):
    """Where an approval decision originated."""

    CLIENT = "client"
    ALLOWLIST = "allowlist"
    CONFIG = "config"
    POLICY = "policy"


class ApprovalKey(DomainModel):
    """Identity of an approved action (workbook + action + optional procedure).

    ``canonical_path`` must come from ``PathPolicy.validate_path`` (never raw
    input). ``sha256`` is the authoritative identity when present.
    """

    canonical_path: str
    sha256: str | None = None
    action: str
    procedure: str | None = None


class ApprovalGrant(DomainModel):
    """A one-time or session approval granted by the client/config."""

    key: ApprovalKey
    scope: ApprovalScope
    granted_at: datetime
    expires_at: datetime | None = None
    granted_by: str
    source: ApprovalSource
    grant_id: str


class AllowlistEntry(DomainModel):
    """A static allowlist entry (config-supplied; no secrets)."""

    key: ApprovalKey
    note: str | None = None


class ApprovalDecision(DomainModel):
    """Outcome of an approval check."""

    allowed: bool
    source: ApprovalSource | None = None
    grant_id: str | None = None
    reason: str


def _utcnow() -> datetime:
    return datetime.now(UTC)


def _as_utc(value: datetime) -> datetime:
    """Coerce a datetime to timezone-aware UTC (naive values assumed UTC)."""

    if value.tzinfo is None:
        return value.replace(tzinfo=UTC)
    return value.astimezone(UTC)


def _identity_matches(requested: ApprovalKey, entry: ApprovalKey) -> bool:
    """Match workbook identity: hash when both present, else canonical path."""

    if requested.sha256 is not None and entry.sha256 is not None:
        return requested.sha256 == entry.sha256
    return requested.canonical_path == entry.canonical_path


def _procedure_matches(requested: str | None, entry: str | None) -> bool:
    """Match procedure: ``None`` on the entry is a whole-workbook wildcard."""

    if entry is None:
        return True
    return requested == entry


def _key_matches(requested: ApprovalKey, entry: ApprovalKey) -> bool:
    """True when ``entry`` approves ``requested`` (identity + action + procedure)."""

    if not _identity_matches(requested, entry):
        return False
    if requested.action != entry.action:
        return False
    return _procedure_matches(requested.procedure, entry.procedure)


class ApprovalPolicy:
    """Deny-by-default approval gate over an allowlist and live grants.

    Args:
        allowlist: Static :class:`AllowlistEntry` items (config/allowlist source).
        now: Injectable clock returning the current UTC time (tests override it).
        mode: Server access policy. The object default is :attr:`AccessMode.RESTRICTED`
            (deny-by-default); the server passes :attr:`AccessMode.FULL` to authorize
            mutating/executing operations by default (see :meth:`check`).
    """

    def __init__(
        self,
        allowlist: Iterable[AllowlistEntry] = (),
        *,
        now: Callable[[], datetime] = _utcnow,
        mode: AccessMode = AccessMode.RESTRICTED,
    ) -> None:
        self._allowlist: list[AllowlistEntry] = list(allowlist)
        self._now: Callable[[], datetime] = now
        self._mode = mode
        # Grants are stored with an optional bound session id (used only for
        # SESSION-scoped grants to detect a session change).
        self._grants: list[tuple[ApprovalGrant, str | None]] = []
        self._lock = threading.Lock()

    @property
    def allowlist(self) -> list[AllowlistEntry]:
        """Read-only snapshot of the configured allowlist."""

        return list(self._allowlist)

    @property
    def access_mode(self) -> AccessMode:
        """The configured server access policy."""

        return self._mode

    def grant(self, grant: ApprovalGrant, *, session_id: str | None = None) -> None:
        """Register a grant.

        Args:
            grant: The grant to add.
            session_id: For ``SESSION``-scoped grants, bind the grant to this
                session so a later ``check`` with a different ``session_id`` is
                denied. Ignored for ``ONE_TIME`` grants.
        """

        with self._lock:
            self._grants.append((grant, session_id))

    def check(
        self,
        key: ApprovalKey,
        *,
        session_id: str | None = None,
        operation_class: OperationClass | None = None,
    ) -> ApprovalDecision:
        """Decide whether ``key`` is approved.

        Args:
            key: The requested action identity.
            session_id: Current session id, used to validate SESSION grants.
            operation_class: When ``READ_ONLY``, the operation is auto-allowed
                without consulting the allowlist or grants.

        Returns:
            An :class:`ApprovalDecision`. ``READ_ONLY`` requests are auto-allowed
            with ``source=None``. In :attr:`AccessMode.FULL`, ``MUTATING`` and
            ``EXECUTING`` requests are allowed by policy (source
            ``ApprovalSource.POLICY``, no grant). In the default
            :attr:`AccessMode.RESTRICTED`, requests with no allowlist match and no
            live grant are denied (deny-by-default).
        """

        if operation_class == OperationClass.READ_ONLY:
            return ApprovalDecision(
                allowed=True,
                source=None,
                grant_id=None,
                reason="read-only operations require no approval",
            )

        # Full-access policy: mutations/executions are authorized by default
        # (no per-action grant). This does not touch PathPolicy / Office-trust /
        # snapshot / audit — those gates run independently upstream.
        if self._mode is AccessMode.FULL:
            return ApprovalDecision(
                allowed=True,
                source=ApprovalSource.POLICY,
                grant_id=None,
                reason="allowed by full-access policy",
            )

        now = self._now()

        with self._lock:
            for entry in self._allowlist:
                if _key_matches(key, entry.key):
                    return ApprovalDecision(
                        allowed=True,
                        source=ApprovalSource.ALLOWLIST,
                        grant_id=None,
                        reason="matched an allowlist entry",
                    )

            for index, (grant, bound_session) in enumerate(self._grants):
                if not _grant_is_live(grant, bound_session, now, session_id):
                    continue
                if not _key_matches(key, grant.key):
                    continue
                if grant.scope == ApprovalScope.ONE_TIME:
                    # Consume the one-time grant so it cannot be reused.
                    del self._grants[index]
                return ApprovalDecision(
                    allowed=True,
                    source=grant.source,
                    grant_id=grant.grant_id,
                    reason="matched a live approval grant",
                )

        return ApprovalDecision(
            allowed=False,
            source=None,
            grant_id=None,
            reason="denied by default: no matching allowlist entry or live grant",
        )

    def authorize_vba_write(self, key: ApprovalKey) -> ApprovalDecision:
        """Authorize a MUTATING VBA write operation (deny-by-default).

        A thin wrapper over :meth:`check` that always classifies the request as
        ``MUTATING``. Mutating VBA writes (replace / insert) must never be
        auto-allowed; they require an explicit allowlist entry or a live grant
        that matches the ``key`` (canonical path + sha256 + action + optional
        component / procedure). Returns the :class:`ApprovalDecision`; the
        caller is responsible for raising the appropriate error when
        ``allowed`` is ``False``.

        Args:
            key: The identity of the action to authorize. ``sha256`` is the
                authoritative identity when both the request and the entry carry
                it; otherwise ``canonical_path`` is used (see :func:`_identity_matches`).

        Returns:
            An :class:`ApprovalDecision` whose ``allowed`` is ``True`` only when
            a live allowlist entry or grant matches; ``False`` otherwise.
        """

        return self.check(key, operation_class=OperationClass.MUTATING)

    def check_recovery(self, key: ApprovalKey) -> ApprovalDecision:
        """Authorize the agent-directed ``excel_recover`` RECOVERY operation.

        The RECOVERY class is **auto-allowed in BOTH** :attr:`AccessMode.FULL` and
        :attr:`AccessMode.RESTRICTED` (report_07 §5, plan_02 Work item A): recovery
        is autonomous (no human approval) and is **pid-scoped** — the caller passes a
        ``key`` whose ``canonical_path`` is the owned-pid identity (e.g.
        ``"pid:<owned_pid>"``), never a workbook wildcard. It discards **working
        copies only**; originals are never modified. The decision is still
        **audited** by the caller (the recovery service records an audit entry).

        This is deliberately distinct from :meth:`check` for ``MUTATING`` /
        ``EXECUTING`` operations, which stay deny-by-default in RESTRICTED: a
        mutating op is **not** a recovery and is never auto-allowed here.

        Args:
            key: The pid-scoped recovery identity (audit only; recovery is always
                allowed). ``canonical_path`` should encode the owned pid.

        Returns:
            An allowed :class:`ApprovalDecision` (source ``POLICY``).
        """

        return ApprovalDecision(
            allowed=True,
            source=ApprovalSource.POLICY,
            grant_id=None,
            reason="recovery is auto-allowed (agent-directed, pid-scoped, originals safe)",
        )


def _grant_is_live(
    grant: ApprovalGrant,
    bound_session: str | None,
    now: datetime,
    session_id: str | None,
) -> bool:
    """Return whether a grant is currently usable (not expired / session-matched)."""

    if grant.expires_at is not None and now > _as_utc(grant.expires_at):
        return False
    session_mismatch = (
        grant.scope == ApprovalScope.SESSION
        and bound_session is not None
        and session_id is not None
        and session_id != bound_session
    )
    return not session_mismatch
