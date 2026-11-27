"""Unit tests for the unified large-data configuration (WP-A).

These tests never touch Excel/COM. They prove that:

* ``load_config()`` with no large-data env vars yields the documented defaults;
* each ``LargeDataConfig`` field is overridable via
  ``EXCEL_VBA_MCP_LARGE_DATA_<UPPER_FIELD>`` (pinned on ``MAX_RANGE_CELLS``);
* a non-integer value raises :class:`pydantic.ValidationError`.
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from excel_vba_mcp.config import load_config

_ENV_MAX_RANGE_CELLS = "EXCEL_VBA_MCP_LARGE_DATA_MAX_RANGE_CELLS"
_ENV_MAX_RESPONSE_CELLS = "EXCEL_VBA_MCP_LARGE_DATA_MAX_RESPONSE_CELLS"
_ENV_MAX_FILE_MB = "EXCEL_VBA_MCP_LARGE_DATA_MAX_FILE_MB"


def test_large_data_defaults(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_MAX_RANGE_CELLS, raising=False)
    monkeypatch.delenv(_ENV_MAX_RESPONSE_CELLS, raising=False)
    monkeypatch.delenv(_ENV_MAX_FILE_MB, raising=False)

    config = load_config()

    assert config.large_data.max_range_cells == 100_000
    assert config.large_data.max_response_cells == 50_000
    assert config.large_data.max_file_mb == 100


def test_large_data_max_range_cells_from_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_MAX_RANGE_CELLS, "5")

    config = load_config()

    assert config.large_data.max_range_cells == 5


@pytest.mark.parametrize("value", ["not-an-int", "1.5", ""])
def test_large_data_invalid_value_raises(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv(_ENV_MAX_RANGE_CELLS, value)

    with pytest.raises(ValidationError):
        load_config()