"""Unit tests for transient-COM propagation through the STA-local adapters.

A broad ``except Exception`` around a single COM call in
:class:`~excel_vba_mcp.adapters.excel.names_adapter.NamesAdapter` /
:class:`~excel_vba_mcp.adapters.excel.tables_adapter.TablesAdapter` must
**re-raise** a genuine transient COM failure (a retryable ``com_error`` HRESULT
or an already-typed ``ComBusyError`` / ``RpcCallRejectedError``) instead of
relabelling it as a non-retryable ``INVALID_RANGE`` / ``NOT_FOUND``. A non-COM
Python error must still be wrapped deterministically.

These tests drive the adapters directly against a minimal fake ``Workbook`` (no
Excel / COM), injecting the failure at the wrapped single COM call.
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.com_transient import _is_transient_com
from excel_vba_mcp.adapters.excel.names_adapter import NamesAdapter
from excel_vba_mcp.adapters.excel.tables_adapter import TablesAdapter
from excel_vba_mcp.domain.errors import (
    ComBusyError,
    ErrorCode,
    InvalidRangeError,
    RpcCallRejectedError,
    category_for,
    error_from_hresult,
    retryable_for,
)

# Well-known retryable COM HRESULTs (kept as plain ints; no pywin32 import).
RPC_E_CALL_REJECTED = 0x80010001
RPC_E_SERVERCALL_RETRYLATER = 0x8001010A


class _FakeComError(Exception):
    """Minimal stand-in for ``pywintypes.com_error`` carrying an ``hresult``.

    The adapters never import ``pywin32``; ``_is_transient_com`` reads
    ``.hresult`` (and falls back to ``args[0]``), so this stand-in exercises the
    same classification path as a real ``com_error``.
    """

    def __init__(self, hresult: int) -> None:
        super().__init__(f"com_error 0x{hresult & 0xFFFFFFFF:08X}")
        self.hresult = hresult


# ---------------------------------------------------------------------------
# Minimal fakes — only what the guarded single COM call needs.
# ---------------------------------------------------------------------------
class _EmptyCollection:
    """An empty, iterable COM collection stand-in (no ``Count``)."""

    def __iter__(self):
        return iter(())


class _FakeNamesCollection(_EmptyCollection):
    """``wb.Names`` stand-in whose ``Add`` raises the injected exception."""

    def __init__(self, add_exc: BaseException) -> None:
        self._add_exc = add_exc

    def Add(self, **_kwargs: Any) -> Any:  # noqa: N802 - COM casing
        raise self._add_exc


class _FakeNamesWorkbook:
    """Workbook stand-in for ``NamesAdapter.create_name`` (workbook scope)."""

    def __init__(self, add_exc: BaseException) -> None:
        self.Names = _FakeNamesCollection(add_exc)
        self.Worksheets = _EmptyCollection()


class _FakeListObjects:
    """``worksheet.ListObjects`` stand-in whose ``Add`` raises."""

    def __init__(self, add_exc: BaseException) -> None:
        self._add_exc = add_exc

    def Add(self, **_kwargs: Any) -> Any:  # noqa: N802 - COM casing
        raise self._add_exc


class _FakeWorksheet:
    """Worksheet stand-in for ``TablesAdapter.create``."""

    Name = "Sheet1"

    def __init__(self, add_exc: BaseException) -> None:
        self.ListObjects = _FakeListObjects(add_exc)

    def Range(self, _address: str) -> Any:  # noqa: N802 - COM casing
        return object()


class _FakeWorksheets:
    def __init__(self, worksheet: _FakeWorksheet) -> None:
        self._worksheet = worksheet

    def Item(self, _name: str) -> _FakeWorksheet:  # noqa: N802 - COM casing
        return self._worksheet


class _FakeTablesWorkbook:
    """Workbook stand-in for ``TablesAdapter.create``."""

    def __init__(self, add_exc: BaseException) -> None:
        self.Worksheets = _FakeWorksheets(_FakeWorksheet(add_exc))


# ---------------------------------------------------------------------------
# _is_transient_com predicate
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    "exc",
    [
        ComBusyError("busy"),
        RpcCallRejectedError("rejected"),
        _FakeComError(RPC_E_CALL_REJECTED),
        _FakeComError(RPC_E_SERVERCALL_RETRYLATER),
    ],
)
def test_is_transient_com_true_for_retryable(exc: BaseException) -> None:
    assert _is_transient_com(exc) is True


@pytest.mark.parametrize(
    "exc",
    [
        ValueError("not com"),
        InvalidRangeError("bad range"),
        _FakeComError(0x80004005),  # E_FAIL -> MACRO_RUNTIME_ERROR (not retryable)
    ],
)
def test_is_transient_com_false_for_non_retryable(exc: BaseException) -> None:
    assert _is_transient_com(exc) is False


# ---------------------------------------------------------------------------
# Names wrapper: create_name
# ---------------------------------------------------------------------------
def test_names_create_propagates_com_busy_error() -> None:
    wb = _FakeNamesWorkbook(ComBusyError("busy"))

    with pytest.raises(ComBusyError) as excinfo:
        NamesAdapter().create_name(wb, "X", "=Sheet1!$A$1", scope="workbook")

    assert excinfo.value.code in (ErrorCode.COM_BUSY, ErrorCode.RPC_E_CALL_REJECTED)
    assert category_for(excinfo.value.code) == "transient_com_failure"
    assert retryable_for(excinfo.value.code) is True


def test_names_create_propagates_retryable_com_error() -> None:
    wb = _FakeNamesWorkbook(_FakeComError(RPC_E_CALL_REJECTED))

    with pytest.raises(_FakeComError) as excinfo:
        NamesAdapter().create_name(wb, "X", "=Sheet1!$A$1", scope="workbook")

    # Not relabelled into a non-retryable domain error: the raw com_error escapes
    # so the worker classifies it. The worker path yields a retryable domain error.
    assert not isinstance(excinfo.value, InvalidRangeError)
    worker_error = error_from_hresult(excinfo.value.hresult, "boom")
    assert retryable_for(worker_error.code) is True
    assert category_for(worker_error.code) == "transient_com_failure"


def test_names_create_wraps_non_com_value_error() -> None:
    wb = _FakeNamesWorkbook(ValueError("bad value"))

    with pytest.raises(InvalidRangeError) as excinfo:
        NamesAdapter().create_name(wb, "X", "=Sheet1!$A$1", scope="workbook")

    assert excinfo.value.code is ErrorCode.INVALID_RANGE
    assert category_for(excinfo.value.code) == "validation"
    assert retryable_for(excinfo.value.code) is False


# ---------------------------------------------------------------------------
# Tables wrapper: create
# ---------------------------------------------------------------------------
def test_tables_create_propagates_com_busy_error() -> None:
    wb = _FakeTablesWorkbook(ComBusyError("busy"))

    with pytest.raises(ComBusyError) as excinfo:
        TablesAdapter().create_table(wb, "Sheet1", "A1:B2")

    assert excinfo.value.code in (ErrorCode.COM_BUSY, ErrorCode.RPC_E_CALL_REJECTED)
    assert category_for(excinfo.value.code) == "transient_com_failure"
    assert retryable_for(excinfo.value.code) is True


def test_tables_create_propagates_retryable_com_error() -> None:
    wb = _FakeTablesWorkbook(_FakeComError(RPC_E_SERVERCALL_RETRYLATER))

    with pytest.raises(_FakeComError) as excinfo:
        TablesAdapter().create_table(wb, "Sheet1", "A1:B2")

    assert not isinstance(excinfo.value, InvalidRangeError)
    worker_error = error_from_hresult(excinfo.value.hresult, "boom")
    assert retryable_for(worker_error.code) is True
    assert category_for(worker_error.code) == "transient_com_failure"


def test_tables_create_wraps_non_com_value_error() -> None:
    wb = _FakeTablesWorkbook(ValueError("bad value"))

    with pytest.raises(InvalidRangeError) as excinfo:
        TablesAdapter().create_table(wb, "Sheet1", "A1:B2")

    assert excinfo.value.code is ErrorCode.INVALID_RANGE
    assert category_for(excinfo.value.code) == "validation"
    assert retryable_for(excinfo.value.code) is False