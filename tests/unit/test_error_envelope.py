"""Unit tests for the structured error envelope (U2; plan_03 §7).

Static, read-only: never touches Excel/COM. Proves that:

* ``ExcelVbaMcpError.to_dict`` is additive — the original five keys are kept and
  ``errorCategory`` / ``isRetryable`` / ``suggestedNextActions`` are added;
* ``category_for`` / ``retryable_for`` / ``actions_for`` are total over every
  ``ErrorCode`` member and the job-state strings (``expired`` /
  ``interrupted_unknown``);
* the D-U10 cases hold: ``STALE_TARGET`` / ``REVISION_CONFLICT`` are NOT retryable
  and carry the "may have partially applied" caveat; only ``COM_BUSY`` /
  ``RPC_E_CALL_REJECTED`` are retryable; a job ``interrupted_unknown`` is never
  retryable.
"""

from __future__ import annotations

import json

import pytest

from excel_vba_mcp.domain.errors import (
    ErrorCode,
    ExcelVbaMcpError,
    actions_for,
    category_for,
    retryable_for,
)

_ORIGINAL_KEYS = {"errorCode", "message", "hresult", "hresult_hex", "details"}
_NEW_KEYS = {"errorCategory", "isRetryable", "suggestedNextActions"}


def test_to_dict_is_additive_and_complete() -> None:
    error = ExcelVbaMcpError("boom", code=ErrorCode.COM_BUSY, context={"workbook_id": "w"})
    payload = error.to_dict()

    # Original five keys preserved + three new keys added.
    assert payload.keys() >= _ORIGINAL_KEYS
    assert payload.keys() >= _NEW_KEYS
    assert payload["errorCode"] == "COM_BUSY"
    assert payload["details"] == {"workbook_id": "w"}
    # Final key order is fixed (U6-B; plan_05 §5.2).
    assert list(payload) == [
        "errorCode",
        "message",
        "hresult",
        "hresult_hex",
        "details",
        "errorCategory",
        "isRetryable",
        "suggestedNextActions",
    ]
    # Must remain JSON-serializable.
    assert json.loads(json.dumps(payload))["errorCategory"] == "transient_com_failure"


@pytest.mark.parametrize("code", list(ErrorCode))
def test_envelope_functions_are_total_over_every_code(code: ErrorCode) -> None:
    category = category_for(code)
    retryable = retryable_for(code)
    actions = actions_for(code, {"workbook_id": "w"})

    assert isinstance(category, str) and category
    assert isinstance(retryable, bool)
    assert isinstance(actions, list) and actions
    assert all(isinstance(action, str) and action for action in actions)

    # to_dict carries the same values.
    payload = ExcelVbaMcpError("x", code=code).to_dict()
    assert payload["errorCategory"] == category
    assert payload["isRetryable"] is retryable
    assert payload["suggestedNextActions"] == actions


# ---------------------------------------------------------------------------
# D-U10 per-code expectations (plan_03 §7.3)
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "code",
    [
        ErrorCode.STALE_TARGET,
        ErrorCode.REVISION_CONFLICT,
    ],
)
def test_stale_and_revision_are_not_retryable_with_caveat(code: ErrorCode) -> None:
    assert retryable_for(code) is False
    actions = actions_for(code, None)
    joined = " ".join(actions).lower()
    assert "partially applied" in joined
    assert "do not blind-replay" in joined


@pytest.mark.parametrize("code", [ErrorCode.COM_BUSY, ErrorCode.RPC_E_CALL_REJECTED])
def test_com_busy_codes_are_retryable(code: ErrorCode) -> None:
    assert retryable_for(code) is True
    assert category_for(code) == "transient_com_failure"


def test_snapshot_failed_is_missing_prerequisite_with_action() -> None:
    payload = ExcelVbaMcpError("x", code=ErrorCode.SNAPSHOT_FAILED).to_dict()
    assert payload["errorCategory"] == "missing_prerequisite"
    assert payload["isRetryable"] is False
    assert any("snapshot_create" in action for action in payload["suggestedNextActions"])


def test_approval_denied_is_permission_denied() -> None:
    payload = ExcelVbaMcpError("x", code=ErrorCode.APPROVAL_DENIED).to_dict()
    assert payload["errorCategory"] == "permission_denied"
    assert payload["isRetryable"] is False


def test_internal_error_is_internal_not_retryable_with_action() -> None:
    # U6-B: guaranteed structured fallback for an unhandled non-domain exception.
    payload = ExcelVbaMcpError("x", code=ErrorCode.INTERNAL_ERROR).to_dict()
    assert payload["errorCategory"] == "internal"
    assert payload["isRetryable"] is False
    assert payload["suggestedNextActions"] == [
        "inspect server logs; the failure was not a domain error"
    ]


def test_confirmation_required_is_validation_not_retryable_with_action() -> None:
    # U6-B: caller-argument gate raised before a destructive action.
    payload = ExcelVbaMcpError("x", code=ErrorCode.CONFIRMATION_REQUIRED).to_dict()
    assert payload["errorCategory"] == "validation"
    assert payload["isRetryable"] is False
    assert payload["suggestedNextActions"] == [
        "re-call with confirm=true after reviewing what is discarded"
    ]


def test_range_too_large_is_validation_not_retryable_with_action() -> None:
    # WP-Range: a valid but over-cap range is a validation error (not retryable)
    # that tells the caller to page the work into smaller ranges.
    payload = ExcelVbaMcpError("x", code=ErrorCode.RANGE_TOO_LARGE).to_dict()
    assert payload["errorCategory"] == "validation"
    assert payload["isRetryable"] is False
    joined = " ".join(payload["suggestedNextActions"]).lower()
    assert "smaller ranges" in joined


def test_job_interrupted_unknown_is_never_retryable() -> None:
    # Job-state string (not an ErrorCode member) still classifies truthfully.
    assert category_for("interrupted_unknown") == "interrupted_unknown"
    assert retryable_for("interrupted_unknown") is False
    assert actions_for("interrupted_unknown", None) == ["never resubmit; report status"]


def test_job_expired_is_deadline_not_retryable() -> None:
    assert category_for("expired") == "deadline"
    assert retryable_for("expired") is False
    assert actions_for("expired", None)


def test_unknown_string_falls_back_to_validation() -> None:
    assert category_for("totally_unknown") == "validation"
    assert retryable_for("totally_unknown") is False
    assert actions_for("totally_unknown", None)