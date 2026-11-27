"""Unit tests for the WP1 read-path source-type gate.

``pivot_read(fields-list)`` and ``pivot_read(field-get-filters)`` enumerate
``PivotFields``; on an external / OLAP / Data Model (or undeterminable) cache
that enumeration forces Excel to resolve external data and blocks the
single-STA worker past the client deadline. These tests drive
:class:`~excel_vba_mcp.adapters.excel.pivot_tables_adapter.PivotTablesAdapter`
against a **fake COM graph** and assert:

* an ``EXTERNAL`` / ``OLAP`` / ``UNKNOWN`` cache raises
  ``UnsupportedPivotSourceError`` (``UNSUPPORTED_PIVOT_SOURCE``) from both
  ``enumerate_pivot_fields`` and ``get_field_filters``;
* the fake's ``PivotFields`` accessor is **never** touched (a sentinel raises
  if it is), proving the gate runs before any field enumeration;
* a supported ``DATABASE`` cache still returns its fields.

Run::

    uv run pytest tests/unit/test_pivot_read_source_gate.py -q
"""

from __future__ import annotations

from typing import Any

import pytest

from excel_vba_mcp.adapters.excel.pivot_tables_adapter import PivotTablesAdapter
from excel_vba_mcp.domain.errors import ErrorCode, ExcelVbaMcpError
from excel_vba_mcp.domain.statuses import PivotFieldOrientation

# ``XlPivotTableSourceType`` values used to build the fakes (mirrors the
# adapter's own map). ``99`` is deliberately outside the map -> UNKNOWN.
XL_SOURCE_DATABASE = 1
XL_SOURCE_EXTERNAL = 2
XL_SOURCE_UNKNOWN = 99

XL_HIDDEN = 0
XL_ROW = 1


class PivotFieldsTouchedError(AssertionError):
    """Raised by the sentinel ``PivotFields`` accessor if it is ever touched."""


class FakeCache:
    def __init__(self, *, source_type: int, olap: bool = False) -> None:
        self.SourceType = source_type
        self.OLAP = olap


class _SentinelPivotFields:
    """A ``PivotFields`` handle that fails the test if it is ever obtained."""

    def __init__(self) -> None:
        raise PivotFieldsTouchedError("PivotFields must not be accessed before the gate")


class FakePivot:
    """A PivotTable whose ``PivotFields`` accessor is a tripwire.

    ``Name`` / ``PivotCache`` are readable (the gate needs them); ``PivotFields``
    raises the sentinel so any read past the gate fails loudly. ``DataFields`` is
    also a tripwire because the supported path reads it immediately after the
    gate — an unsupported cache must never reach it.
    """

    def __init__(self, name: str, *, source_type: int, olap: bool = False) -> None:
        self.Name = name
        self._cache = FakeCache(source_type=source_type, olap=olap)

    def PivotCache(self) -> FakeCache:  # noqa: N802 - COM casing
        return self._cache

    def PivotFields(self, *_args: Any) -> Any:  # noqa: N802 - COM casing
        return _SentinelPivotFields()

    @property
    def DataFields(self) -> Any:  # noqa: N802 - COM casing
        raise PivotFieldsTouchedError("DataFields must not be accessed before the gate")


class _Fields:
    def __init__(self, fields: list[Any]) -> None:
        self._fields = fields

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._fields)

    def Item(self, index: int) -> Any:  # noqa: N802 - COM casing
        return self._fields[index - 1]


class SupportedField:
    """A minimal real source field for the supported DATABASE case."""

    def __init__(self, source_name: str, *, orientation: int = XL_HIDDEN) -> None:
        self.SourceName = source_name
        self.Name = source_name
        self.Caption = source_name
        self.Orientation = orientation
        self.Position = 1 if orientation != XL_HIDDEN else -2146826246
        self.ShowAllItems = False
        self.AllItemsVisible = True

    def PivotItems(self) -> _Fields:  # noqa: N802 - COM casing
        return _Fields([])


class SupportedPivot:
    """A DATABASE-backed pivot that returns real fields (no tripwire)."""

    def __init__(self, name: str, fields: list[SupportedField]) -> None:
        self.Name = name
        self._cache = FakeCache(source_type=XL_SOURCE_DATABASE)
        self._fields = fields

    def PivotCache(self) -> FakeCache:  # noqa: N802 - COM casing
        return self._cache

    def PivotFields(self, *_args: Any) -> _Fields:  # noqa: N802 - COM casing
        return _Fields(list(self._fields))

    @property
    def DataFields(self) -> _Fields:  # noqa: N802 - COM casing
        return _Fields([])


class FakePivotTables:
    def __init__(self, pivots: list[Any]) -> None:
        self._pivots = pivots

    @property
    def Count(self) -> int:  # noqa: N802 - COM casing
        return len(self._pivots)

    def Item(self, index: int) -> Any:  # noqa: N802 - COM casing
        return self._pivots[index - 1]


class FakeSheet:
    def __init__(self, pivots: list[Any]) -> None:
        self._pivots = pivots

    def PivotTables(self) -> FakePivotTables:  # noqa: N802 - COM casing
        return FakePivotTables(self._pivots)


class FakeWorksheets:
    def __init__(self, sheets: list[FakeSheet]) -> None:
        self._sheets = sheets

    def __iter__(self):
        return iter(self._sheets)


class FakeWorkbook:
    def __init__(self, pivots: list[Any]) -> None:
        self.Worksheets = FakeWorksheets([FakeSheet(pivots)])


# ---------------------------------------------------------------------------
# Unsupported source types: both read paths raise, PivotFields never touched
# ---------------------------------------------------------------------------
@pytest.mark.parametrize(
    ("source_type", "olap"),
    [
        (XL_SOURCE_EXTERNAL, False),
        (XL_SOURCE_DATABASE, True),  # OLAP indicator overrides -> OLAP
        (XL_SOURCE_UNKNOWN, False),  # undeterminable -> fail closed
    ],
)
def test_fields_list_rejects_unsupported_source(source_type: int, olap: bool) -> None:
    pivot = FakePivot("ptExt", source_type=source_type, olap=olap)
    wb = FakeWorkbook([pivot])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.enumerate_pivot_fields(wb, "ptExt")

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    assert exc_info.value.context["pivot_table"] == "ptExt"
    assert exc_info.value.context["reason"] == "unsupported_source"
    assert exc_info.value.context["source_type"] in {"external", "olap", "unknown"}


@pytest.mark.parametrize(
    ("source_type", "olap"),
    [
        (XL_SOURCE_EXTERNAL, False),
        (XL_SOURCE_DATABASE, True),  # OLAP indicator overrides -> OLAP
        (XL_SOURCE_UNKNOWN, False),  # undeterminable -> fail closed
    ],
)
def test_field_get_filters_rejects_unsupported_source(source_type: int, olap: bool) -> None:
    pivot = FakePivot("ptExt", source_type=source_type, olap=olap)
    wb = FakeWorkbook([pivot])
    adapter = PivotTablesAdapter()

    with pytest.raises(ExcelVbaMcpError) as exc_info:
        adapter.get_field_filters(wb, "ptExt", "Region")

    assert exc_info.value.code is ErrorCode.UNSUPPORTED_PIVOT_SOURCE
    assert exc_info.value.context["pivot_table"] == "ptExt"
    assert exc_info.value.context["reason"] == "unsupported_source"


# ---------------------------------------------------------------------------
# Supported DATABASE source: fields are still returned
# ---------------------------------------------------------------------------
def test_fields_list_returns_fields_for_database_source() -> None:
    pivot = SupportedPivot(
        "ptDb",
        [
            SupportedField("Region", orientation=XL_ROW),
            SupportedField("Sales"),
        ],
    )
    wb = FakeWorkbook([pivot])
    adapter = PivotTablesAdapter()

    fields = adapter.enumerate_pivot_fields(wb, "ptDb")

    names = {f.source_name for f in fields}
    assert names == {"Region", "Sales"}
    region = next(f for f in fields if f.source_name == "Region")
    assert region.orientation is PivotFieldOrientation.ROW


def test_field_get_filters_returns_state_for_database_source() -> None:
    pivot = SupportedPivot("ptDb", [SupportedField("Region", orientation=XL_ROW)])
    wb = FakeWorkbook([pivot])
    adapter = PivotTablesAdapter()

    state = adapter.get_field_filters(wb, "ptDb", "Region")

    assert state.pivot_table == "ptDb"
    assert state.field == "Region"
    assert state.all_items_visible is True