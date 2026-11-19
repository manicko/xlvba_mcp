"""STA-local ``Application.Run`` adapter with argument coercion (task_022).

:class:`MacroRunner` is the single place that invokes ``Application.Run`` for a
macro. It is a **pure STA-local helper**: it is constructed on the dedicated
single-STA COM worker thread with the owned ``Excel.Application`` proxy and
returns only a :class:`~excel_vba_mcp.domain.models.MacroRunResult` DTO (or
raises), so the result crosses the worker boundary safely
(:func:`~excel_vba_mcp.adapters.excel.com_worker._ensure_dto`). It never imports
``pywin32`` at module import time, never activates/selects anything and never
shows or dismisses a window.

Binding ``Application.Run`` facts (research ``g1-poc-research-2026-11-19.md`` §4):

* A qualified macro name (``'<wb>'!<Module>.<Proc>``) returns a serializable
  Variant for scalar arguments and for a Variant array return.
* ``datetime.datetime`` arguments are accepted by pywin32; a bare
  ``datetime.date`` is **not** (``TypeError: must be a pywintypes time object``)
  and must be rejected before dispatch. Arbitrary Python objects / COM
  references are rejected before any COM call reaches Excel.
* Supported arguments: ``str`` / ``int`` / ``float`` / ``bool`` / ``None`` /
  ``datetime`` and 1-D / 2-D lists of those.
* A missing macro and an unhandled runtime error both surface an **outer**
  ``DISP_E_EXCEPTION`` (``0x80020009``); they are distinguished by the **inner**
  HRESULT (``0x800A03ED`` not-found) and the EXCEPINFO description ("Cannot run
  the macro …" / "… may not be available"). The generic outer code alone is never
  trusted, and a generic failure is never reported as ``NOT_FOUND``.
* A return value of an unsupported type (e.g. ``Object`` / ``Scripting.
  Dictionary``) is rejected as non-serializable rather than crossing the worker
  boundary.

The runner never retries: a single ``Application.Run`` dispatch is issued and its
outcome (success, not-found, runtime error, blocked) is normalized into a
:class:`MacroRunResult`. Timeout / hang handling (worker soft deadline, modal
detection, owned-PID termination) is owned by the application-layer
``MacroService`` because it needs the worker and process bookkeeping; this adapter
only classifies the synchronous return / exception of one dispatch.
"""

from __future__ import annotations

import datetime as _dt
import time
from typing import Any
from uuid import uuid4

from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError, format_hresult
from excel_vba_mcp.domain.models import MacroRunResult
from excel_vba_mcp.domain.statuses import MacroRunStatus

# Outer DISP_E_EXCEPTION wrapper that carries the inner HRESULT (research §4.2).
DISP_E_EXCEPTION = 0x80020009
# Inner HRESULTs observed on the target build.
HR_MACRO_NOT_FOUND = 0x800A03ED  # macro not available in this workbook
HR_MACRO_RUNTIME_ERROR = 0x800A03EC  # VBA runtime error surfaced by Excel
HR_E_FAIL = 0x80004005  # generic E_FAIL from a runtime error

# Structured error codes attached to a blocked (did-not-run / not-serializable)
# result. These are carried in ``MacroRunResult.error_code`` so a caller can tell
# a policy block apart from a bad-argument / bad-return block.
ERROR_UNSUPPORTED_ARGUMENT = "UNSUPPORTED_ARGUMENT"
ERROR_NON_SERIALIZABLE_RETURN = "NON_SERIALIZABLE_RETURN"

# Substrings in the EXCEPINFO description that positively identify a missing
# macro regardless of the inner HRESULT (research §4.2 / task_009 §3).
_NOT_FOUND_MARKERS = ("cannot run the macro", "may not be available")


def utc_now() -> _dt.datetime:
    """Current UTC time (used for the ``MacroRunResult`` timestamps)."""

    return _dt.datetime.now(_dt.UTC)


# ---------------------------------------------------------------------------
# Argument coercion (pure; no COM)
# ---------------------------------------------------------------------------
def coerce_arg(value: Any) -> Any:
    """Validate/normalize one macro argument for ``Application.Run``.

    Accepts ``str`` / ``int`` / ``float`` / ``bool`` / ``None`` / ``datetime`` and
    (recursively) 1-D / 2-D lists of those. A bare ``datetime.date`` (which is
    **not** a ``datetime.datetime``) and any arbitrary object / COM reference are
    rejected with :class:`ExcelVbaMcpError` **before** any COM dispatch.

    Raises:
        ExcelVbaMcpError: When the argument (or a nested element) is not a
            supported COM Variant scalar / array.
    """

    # ``datetime.datetime`` is a subclass of ``datetime.date``; accept it first so
    # a real timestamp is allowed while a bare date is rejected below.
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        raise ExcelVbaMcpError(
            "datetime.date arguments are not supported by Application.Run; "
            "pass a datetime.datetime instead.",
            code=ErrorCode.MACRO_BLOCKED,
            context={"reason": ERROR_UNSUPPORTED_ARGUMENT, "type": type(value).__name__},
        )
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (list, tuple)):
        return [coerce_arg(item) for item in value]
    raise ExcelVbaMcpError(
        f"Unsupported macro argument of type {type(value).__name__!r}; only "
        "str/int/float/bool/None/datetime and 1-D/2-D lists are supported.",
        code=ErrorCode.MACRO_BLOCKED,
        context={"reason": ERROR_UNSUPPORTED_ARGUMENT, "type": type(value).__name__},
    )


def coerce_args(args: list[Any] | tuple[Any, ...] | None) -> list[Any]:
    """Coerce a positional argument sequence for ``Application.Run``.

    Raises:
        ExcelVbaMcpError: When any argument is unsupported (see :func:`coerce_arg`).
    """

    return [coerce_arg(item) for item in (args or ())]


def serialize_return(value: Any) -> Any:
    """Serialize an ``Application.Run`` Variant return into a JSON-safe value.

    Scalars (``None`` / ``bool`` / ``int`` / ``float`` / ``str``), ``datetime``
    (ISO-8601) and (nested) lists/tuples are accepted; any other object (a COM
    proxy, ``dict`` or arbitrary Python object) is rejected as non-serializable
    so it never crosses the worker boundary.

    Raises:
        ExcelVbaMcpError: When the return value is not serializable.
    """

    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [serialize_return(item) for item in value]
    raise ExcelVbaMcpError(
        f"Macro return value of type {type(value).__name__!r} is not serializable; "
        "only scalars, datetime and (nested) arrays are supported.",
        code=ErrorCode.MACRO_BLOCKED,
        context={"reason": ERROR_NON_SERIALIZABLE_RETURN, "type": type(value).__name__},
    )


# ---------------------------------------------------------------------------
# HRESULT extraction (from a pywintypes.com_error-like exception)
# ---------------------------------------------------------------------------
def _outer_hresult(exc: BaseException) -> int | None:
    """Return the outer HRESULT carried on a COM error, if any."""

    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        return hresult & 0xFFFFFFFF
    args: tuple[Any, ...] = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0] & 0xFFFFFFFF
    return None


def _inner_hresult(exc: BaseException) -> int | None:
    """Return the EXCEPINFO inner HRESULT from a ``com_error``, if present.

    ``win32com`` wraps an ``IDispatch`` failure as ``DISP_E_EXCEPTION`` whose
    third arg is the EXCEPINFO tuple; the inner HRESULT is element 5 (verified on
    the target build; research §4.2). Returns ``None`` when absent.
    """

    args: tuple[Any, ...] = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        candidate = args[2][5]
        if isinstance(candidate, int):
            return candidate & 0xFFFFFFFF
    return None


def _exception_description(exc: BaseException) -> str | None:
    """Return the EXCEPINFO description string, if present.

    The EXCEPINFO tuple is ``(hr, source, description, helpfile, helpcontext,
    inner_hresult)``. Element 2 is the human-readable description (e.g. Excel's
    ``Cannot run the macro '...'``). Returns ``None`` when absent.
    """

    args: tuple[Any, ...] = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 6:
        candidate = args[2][2]
        if isinstance(candidate, str):
            return candidate
    return None


def _classify_exception(exc: BaseException) -> tuple[MacroRunStatus, str]:
    """Classify an ``Application.Run`` exception into ``(status, error_code)``.

    The outer ``DISP_E_EXCEPTION`` (``0x80020009``) is shared by every failure;
    the inner HRESULT and the EXCEPINFO description are the decisive evidence
    (research §4.2 / task_009 §3):

    * inner ``0x800A03ED`` -> ``MACRO_NOT_FOUND``;
    * description contains ``Cannot run the macro`` / ``may not be available``
      -> ``MACRO_NOT_FOUND`` (regardless of the inner HRESULT);
    * otherwise -> ``MACRO_RUNTIME_ERROR`` (fail-closed; a generic failure is
      never reported as ``NOT_FOUND``).
    """

    inner = _inner_hresult(exc)
    description = (_exception_description(exc) or "").lower()
    if inner == HR_MACRO_NOT_FOUND or any(marker in description for marker in _NOT_FOUND_MARKERS):
        return MacroRunStatus.NOT_FOUND, ErrorCode.MACRO_NOT_FOUND.value
    return MacroRunStatus.RUNTIME_ERROR, ErrorCode.MACRO_RUNTIME_ERROR.value


# ---------------------------------------------------------------------------
# Result builders
# ---------------------------------------------------------------------------
def _duration_ms(started: float) -> int:
    return int(round((time.monotonic() - started) * 1000.0))


def _base_result(
    run_id: str,
    macro_name: str,
    status: MacroRunStatus,
    started_at: _dt.datetime,
    started: float,
) -> MacroRunResult:
    """Build a ``MacroRunResult`` with timing + timestamps filled in."""

    finished_at = utc_now()
    return MacroRunResult(
        run_id=run_id,
        macro_name=macro_name,
        status=status,
        started_at=started_at,
        finished_at=finished_at,
        duration_ms=_duration_ms(started),
        created_at=started_at,
        updated_at=finished_at,
    )


def _blocked_result(
    run_id: str,
    macro_name: str,
    started_at: _dt.datetime,
    started: float,
    error_code: str,
    error_text: str,
) -> MacroRunResult:
    """Build a ``BLOCKED`` result for an argument/return rejection (did not run)."""

    result = _base_result(run_id, macro_name, MacroRunStatus.BLOCKED, started_at, started)
    result.error_code = error_code
    result.error_text = error_text
    return result


class MacroRunner:
    """STA-local ``Application.Run`` wrapper (single dispatch, no retry).

    Args:
        app: The owned, hidden ``Excel.Application`` COM proxy (STA thread only).
    """

    def __init__(self, app: Any) -> None:
        self._app = app

    def run(self, qualified_name: str, args: list[Any] | None = None) -> MacroRunResult:
        """Invoke ``qualified_name`` with positional ``args`` and normalize the outcome.

        The arguments are coerced/validated **before** dispatch (research §4.1): a
        bare ``datetime.date`` or an arbitrary object is rejected without any COM
        call reaching Excel. A single ``Application.Run`` is issued; its Variant
        return is serialized, and any COM failure is classified by inner HRESULT +
        description (never by the generic outer code alone). No retry is performed.

        Args:
            qualified_name: Fully qualified macro name, e.g.
                ``'Book.xlsm'!modTest.AddNumbers``.
            args: Positional arguments (coerced by :func:`coerce_args`).

        Returns:
            A :class:`MacroRunResult` with ``status`` one of ``SUCCESS`` /
            ``NOT_FOUND`` / ``RUNTIME_ERROR`` / ``BLOCKED`` and, for a blocked
            argument/return, an ``error_code`` of ``UNSUPPORTED_ARGUMENT`` or
            ``NON_SERIALIZABLE_RETURN``.
        """

        run_id = f"run-{uuid4().hex}"
        started_at = utc_now()
        started = time.monotonic()

        try:
            coerced = coerce_args(args)
        except ExcelVbaMcpError as exc:
            return _blocked_result(
                run_id,
                qualified_name,
                started_at,
                started,
                ERROR_UNSUPPORTED_ARGUMENT,
                str(exc),
            )

        try:
            returned = self._app.Run(qualified_name, *coerced)
        except TypeError as exc:
            # pywin32 rejected an argument before dispatch (defensive: coerce_args
            # should already have caught this).
            return _blocked_result(
                run_id,
                qualified_name,
                started_at,
                started,
                ERROR_UNSUPPORTED_ARGUMENT,
                f"Application.Run rejected the arguments: {exc}",
            )
        except Exception as exc:  # noqa: BLE001 - com_error or other dispatch failure
            status, error_code = _classify_exception(exc)
            result = _base_result(run_id, qualified_name, status, started_at, started)
            result.error_code = error_code
            result.error_text = _format_exception_text(exc)
            return result

        try:
            value = serialize_return(returned)
        except ExcelVbaMcpError as exc:
            return _blocked_result(
                run_id,
                qualified_name,
                started_at,
                started,
                ERROR_NON_SERIALIZABLE_RETURN,
                str(exc),
            )

        result = _base_result(run_id, qualified_name, MacroRunStatus.SUCCESS, started_at, started)
        result.return_value = value
        return result


def _format_exception_text(exc: BaseException) -> str:
    """Format a dispatch exception into a compact, HRESULT-annotated message."""

    outer = _outer_hresult(exc)
    inner = _inner_hresult(exc)
    description = _exception_description(exc) or str(exc) or exc.__class__.__name__
    parts: list[str] = []
    if outer is not None:
        parts.append(f"outer={format_hresult(outer)}")
    if inner is not None:
        parts.append(f"inner={format_hresult(inner)}")
    prefix = f"[{', '.join(parts)}] " if parts else ""
    # Excel returns a generic COM error for a runtime error; never promise the
    # original VBA description (spec §11 / research §4.2).
    return f"{prefix}{description[:400]}"


__all__ = [
    "DISP_E_EXCEPTION",
    "ERROR_NON_SERIALIZABLE_RETURN",
    "ERROR_UNSUPPORTED_ARGUMENT",
    "HR_E_FAIL",
    "HR_MACRO_NOT_FOUND",
    "HR_MACRO_RUNTIME_ERROR",
    "MacroRunner",
    "coerce_arg",
    "coerce_args",
    "serialize_return",
]
