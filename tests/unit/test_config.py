"""Unit tests for typed configuration loading (ARCH-004)."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from excel_vba_mcp.config import ServerConfig, load_config

_ENV_VAR = "EXCEL_VBA_MCP_LOG_LEVEL"


def test_log_level_defaults_to_info(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_VAR, raising=False)

    config = load_config()

    assert config.server.log_level == "INFO"


def test_valid_log_level_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_VAR, "DEBUG")

    config = load_config()

    assert config.server.log_level == "DEBUG"


@pytest.mark.parametrize("value", ["TRACE", "verbose", "10", ""])
def test_invalid_log_level_raises(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv(_ENV_VAR, value)

    with pytest.raises(ValidationError):
        load_config()


def test_server_config_rejects_invalid_literal() -> None:
    with pytest.raises(ValidationError):
        ServerConfig.model_validate({"log_level": "TRACE"})
