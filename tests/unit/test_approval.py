"""Unit tests for the deny-by-default approval policy (task_009)."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from pathlib import Path

from excel_vba_mcp.domain.statuses import AccessMode, OperationClass
from excel_vba_mcp.security.approval import (
    AllowlistEntry,
    ApprovalGrant,
    ApprovalKey,
    ApprovalPolicy,
    ApprovalScope,
    ApprovalSource,
)

RO = OperationClass.READ_ONLY
MU = OperationClass.MUTATING
EX = OperationClass.EXECUTING

_T0 = datetime(2026, 11, 19, 12, 0, 0, tzinfo=UTC)


def _key(
    path: str = "C:/work/Report.xlsm",
    *,
    sha256: str | None = None,
    action: str = "vba_write",
    procedure: str | None = None,
) -> ApprovalKey:
    return ApprovalKey(canonical_path=path, sha256=sha256, action=action, procedure=procedure)


def _grant(
    key: ApprovalKey,
    *,
    scope: ApprovalScope = ApprovalScope.ONE_TIME,
    expires_at: datetime | None = None,
    source: ApprovalSource = ApprovalSource.CLIENT,
    grant_id: str = "g1",
) -> ApprovalGrant:
    return ApprovalGrant(
        key=key,
        scope=scope,
        granted_at=_T0,
        expires_at=expires_at,
        granted_by="operator",
        source=source,
        grant_id=grant_id,
    )


def test_deny_by_default_mutating() -> None:
    policy = ApprovalPolicy(now=lambda: _T0)
    decision = policy.check(_key(), operation_class=MU)
    assert decision.allowed is False
    assert "denied by default" in decision.reason


def test_deny_by_default_executing() -> None:
    policy = ApprovalPolicy(now=lambda: _T0)
    decision = policy.check(_key(action="macro_run", procedure="Mod.Run"), operation_class=EX)
    assert decision.allowed is False


def test_read_only_auto_allowed() -> None:
    policy = ApprovalPolicy(now=lambda: _T0)
    decision = policy.check(_key(), operation_class=RO)
    assert decision.allowed is True
    assert decision.source is None


def test_allowlist_allows() -> None:
    key = _key(action="sheet_delete")
    policy = ApprovalPolicy([AllowlistEntry(key=key, note="trusted")], now=lambda: _T0)
    decision = policy.check(key, operation_class=MU)
    assert decision.allowed is True
    assert decision.source is ApprovalSource.ALLOWLIST


def test_allowlist_action_mismatch_denies() -> None:
    entry = AllowlistEntry(key=_key(action="sheet_delete"))
    policy = ApprovalPolicy([entry], now=lambda: _T0)
    # Same workbook, different action -> not approved.
    decision = policy.check(_key(action="name_delete"), operation_class=MU)
    assert decision.allowed is False


def test_one_time_grant_consumed() -> None:
    key = _key(action="macro_run", procedure="Mod.Run")
    policy = ApprovalPolicy(now=lambda: _T0)
    policy.grant(_grant(key))

    first = policy.check(key, operation_class=EX)
    assert first.allowed is True
    assert first.grant_id == "g1"

    # One-time grant is consumed: the second attempt is denied.
    second = policy.check(key, operation_class=EX)
    assert second.allowed is False


def test_session_grant_valid_within_session() -> None:
    key = _key(action="vba_write")
    policy = ApprovalPolicy(now=lambda: _T0)
    policy.grant(_grant(key, scope=ApprovalScope.SESSION), session_id="s1")

    decision = policy.check(key, session_id="s1", operation_class=MU)
    assert decision.allowed is True
    # Session grants are reusable within the same session.
    assert policy.check(key, session_id="s1", operation_class=MU).allowed is True


def test_session_grant_denied_after_session_change() -> None:
    key = _key(action="vba_write")
    policy = ApprovalPolicy(now=lambda: _T0)
    policy.grant(_grant(key, scope=ApprovalScope.SESSION), session_id="s1")

    assert policy.check(key, session_id="s2", operation_class=MU).allowed is False


def test_session_grant_expiry_denies() -> None:
    key = _key(action="vba_write")
    policy = ApprovalPolicy(now=lambda: _T0)
    policy.grant(
        _grant(
            key,
            scope=ApprovalScope.SESSION,
            expires_at=_T0 + timedelta(seconds=30),
        ),
        session_id="s1",
    )

    assert policy.check(key, session_id="s1", operation_class=MU).allowed is True

    # Advance the clock past expiry.
    later = _T0 + timedelta(seconds=31)
    policy._now = lambda: later  # type: ignore[assignment]
    assert policy.check(key, session_id="s1", operation_class=MU).allowed is False


def test_one_time_grant_expiry_denies() -> None:
    key = _key(action="vba_write")
    policy = ApprovalPolicy(now=lambda: _T0 + timedelta(hours=1))
    policy.grant(_grant(key, expires_at=_T0 + timedelta(minutes=30)))
    assert policy.check(key, operation_class=MU).allowed is False


def test_same_canonical_path_different_spellings_same_key(tmp_path: Path) -> None:
    # Simulate two textual spellings that resolve to the same canonical path; the
    # approval key is built from the resolved path, so they match.
    target = tmp_path / "Report.xlsm"
    target.write_bytes(b"seed")
    spelling_a = str(target.resolve())
    spelling_b = str((tmp_path / "." / "Report.xlsm").resolve())
    assert spelling_a == spelling_b

    key = _key(spelling_a, action="excel_save")
    policy = ApprovalPolicy([AllowlistEntry(key=_key(spelling_b, action="excel_save"))])
    assert policy.check(key, operation_class=MU).allowed is True


def test_same_basename_different_sha256_distinct_keys() -> None:
    # Hash is authoritative when both present: same basename/path but different
    # content hash -> a grant for one does not cover the other.
    grant_key = _key("C:/work/Report.xlsm", sha256="aaa", action="vba_write")
    policy = ApprovalPolicy(now=lambda: _T0)
    policy.grant(_grant(grant_key))

    matching = _key("C:/work/Report.xlsm", sha256="aaa", action="vba_write")
    assert policy.check(matching, operation_class=MU).allowed is True

    other = _key("C:/work/Report.xlsm", sha256="bbb", action="vba_write")
    assert policy.check(other, operation_class=MU).allowed is False


def test_sha256_only_matches_when_both_present() -> None:
    # Request has no hash -> falls back to canonical path; entry keyed by hash only
    # (different path) does not match, and vice versa.
    entry = AllowlistEntry(key=_key("C:/a.xlsm", sha256="aaa", action="vba_write"))
    policy = ApprovalPolicy([entry], now=lambda: _T0)
    # Same path, no hash on request -> canonical-path fallback matches.
    assert policy.check(_key("C:/a.xlsm", action="vba_write"), operation_class=MU).allowed is True
    # Different path -> denied.
    assert policy.check(_key("C:/b.xlsm", action="vba_write"), operation_class=MU).allowed is False


def test_procedure_wildcard_and_exact() -> None:
    # Allowlist entry with procedure=None is a whole-workbook wildcard for the action.
    wildcard = AllowlistEntry(key=_key(action="macro_run", procedure=None))
    policy = ApprovalPolicy([wildcard], now=lambda: _T0)
    assert (
        policy.check(_key(action="macro_run", procedure="Mod.Alpha"), operation_class=EX).allowed
        is True
    )
    assert (
        policy.check(_key(action="macro_run", procedure="Mod.Beta"), operation_class=EX).allowed
        is True
    )

    # An exact-procedure entry only covers that procedure.
    exact = ApprovalPolicy(
        [AllowlistEntry(key=_key(action="macro_run", procedure="Mod.Alpha"))], now=lambda: _T0
    )
    assert (
        exact.check(_key(action="macro_run", procedure="Mod.Alpha"), operation_class=EX).allowed
        is True
    )
    assert (
        exact.check(_key(action="macro_run", procedure="Mod.Beta"), operation_class=EX).allowed
        is False
    )


# ---------------------------------------------------------------------------
# Full-access mode (task_b2 / fix_002)
# ---------------------------------------------------------------------------
def test_full_mode_allows_mutating_and_executing() -> None:
    policy = ApprovalPolicy(mode=AccessMode.FULL, now=lambda: _T0)

    mutating = policy.check(_key(action="sheet_delete"), operation_class=MU)
    assert mutating.allowed is True
    assert mutating.source is ApprovalSource.POLICY
    assert mutating.grant_id is None

    executing = policy.check(
        _key(action="macro_run", procedure="Mod.Run"), operation_class=EX
    )
    assert executing.allowed is True
    assert executing.source is ApprovalSource.POLICY
    assert executing.grant_id is None


def test_full_mode_read_only_still_source_none() -> None:
    policy = ApprovalPolicy(mode=AccessMode.FULL, now=lambda: _T0)

    decision = policy.check(_key(), operation_class=RO)
    assert decision.allowed is True
    assert decision.source is None


def test_explicit_restricted_mode_still_denies() -> None:
    policy = ApprovalPolicy(mode=AccessMode.RESTRICTED, now=lambda: _T0)

    assert policy.check(_key(), operation_class=MU).allowed is False
    assert (
        policy.check(_key(action="macro_run", procedure="Mod.Run"), operation_class=EX).allowed
        is False
    )


def test_full_mode_access_mode_property() -> None:
    assert ApprovalPolicy(mode=AccessMode.FULL).access_mode is AccessMode.FULL
    # The object default stays restricted (deny-by-default).
    assert ApprovalPolicy().access_mode is AccessMode.RESTRICTED
