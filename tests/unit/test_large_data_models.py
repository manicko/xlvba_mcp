"""Unit tests for the WP-A large-data additive DTO fields.

Static, read-only: never touches Excel/COM. Proves that:

* ``PageInfo`` round-trips through ``model_dump(mode="json")``;
* ``RangeValues.page`` defaults to ``None`` and accepts a ``PageInfo``;
* ``RangeFormulas.page`` likewise (additive, optional, no behavior change).
"""

from __future__ import annotations

import json

from excel_vba_mcp.domain.models import (
    CellValue,
    CellValueKind,
    FormulaCell,
    PageInfo,
    RangeAddress,
    RangeFormulas,
    RangeShape,
    RangeValues,
)


def _addr() -> RangeAddress:
    return RangeAddress(sheet="Sheet1", address="A1", rows=1, cols=1)


def _page() -> PageInfo:
    return PageInfo(offset=0, returned=1, total=1, complete=True, version=3, source="live")


def test_page_info_json_round_trip() -> None:
    page = _page()

    payload = json.loads(page.model_dump_json())
    restored = PageInfo.model_validate(payload)

    assert restored.model_dump(mode="json") == page.model_dump(mode="json")
    assert restored.source == "live"
    assert restored.version == 3
    assert restored.complete is True


def test_range_values_page_defaults_to_none() -> None:
    values = RangeValues(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        values=[[CellValue(kind=CellValueKind.NUMBER, value=1)]],
    )

    assert values.page is None


def test_range_values_accepts_page_info() -> None:
    page = _page()
    values = RangeValues(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        values=[[CellValue(kind=CellValueKind.NUMBER, value=1)]],
        page=page,
    )

    assert values.page is not None
    assert values.page.version == 3
    payload = json.loads(values.model_dump_json())
    assert payload["page"]["version"] == 3


def test_range_formulas_page_defaults_to_none() -> None:
    formulas = RangeFormulas(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        formulas=[[FormulaCell(address="A1", formula2="=1+1")]],
    )

    assert formulas.page is None


def test_range_formulas_accepts_page_info() -> None:
    page = _page()
    formulas = RangeFormulas(
        range=_addr(),
        shape=RangeShape.BLOCK,
        rows=1,
        cols=1,
        formulas=[[FormulaCell(address="A1", formula2="=1+1")]],
        page=page,
    )

    assert formulas.page is not None
    assert formulas.page.source == "live"
    payload = json.loads(formulas.model_dump_json())
    restored = RangeFormulas.model_validate(payload)
    assert restored.model_dump(mode="json") == formulas.model_dump(mode="json")