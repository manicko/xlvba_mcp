"""Shared transient-COM detection for the STA-local adapters.

The single-STA worker (``com_worker``) classifies a genuine transient COM
failure — a retryable HRESULT such as ``RPC_E_CALL_REJECTED`` /
``RPC_E_SERVERCALL_RETRYLATER``, or an already-typed :class:`ComBusyError` /
:class:`RpcCallRejectedError` — as retryable (``errorCategory =
transient_com_failure``, ``isRetryable = true``).

The STA-local adapters run *inside* a worker call, so a broad ``except`` around
a single COM call must re-raise such a failure unchanged instead of relabelling
it as a non-retryable domain error (``INVALID_RANGE`` / ``NOT_FOUND``). This
module provides the predicate those wrappers use to keep genuine transient
behaviour identical to the worker path.

It does not import ``pywin32``: a ``com_error``-like exception is recognised by
its ``hresult`` (or ``args[0]``) and classified through the domain retryable
helpers, mirroring ``com_worker._hresult_of`` + ``error_from_hresult``.
"""

from __future__ import annotations

from excel_vba_mcp.domain.errors import (
    ComBusyError,
    code_from_hresult,
    retryable_for,
)


def _hresult_of(exc: BaseException) -> int | None:
    """Extract a raw HRESULT from a ``com_error``-like exception, if present."""

    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        try:
            return int(hresult) & 0xFFFFFFFF
        except (TypeError, ValueError):
            return None
    args = getattr(exc, "args", ())
    if args:
        try:
            return int(args[0]) & 0xFFFFFFFF
        except (TypeError, ValueError):
            return None
    return None


def _is_transient_com(exc: BaseException) -> bool:
    """Return whether ``exc`` is a genuine, retryable transient COM failure.

    True for an already-typed :class:`ComBusyError` /
    :class:`RpcCallRejectedError` and for a ``com_error``-like exception whose
    HRESULT maps to a retryable domain code (``COM_BUSY`` /
    ``RPC_E_CALL_REJECTED``). False for every other exception (including plain
    Python errors and non-retryable domain errors), so the deterministic
    wrapping into a non-retryable domain error is preserved.
    """

    if isinstance(exc, ComBusyError):
        return True
    hresult = _hresult_of(exc)
    if hresult is None:
        return False
    code = code_from_hresult(hresult)
    return code is not None and retryable_for(code)