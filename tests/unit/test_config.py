"""Unit tests for typed configuration loading (ARCH-004)."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from excel_vba_mcp.config import ServerConfig, load_config
from excel_vba_mcp.domain.statuses import AccessMode

_ENV_VAR = "EXCEL_VBA_MCP_LOG_LEVEL"
_ENV_ACCESS_MODE = "EXCEL_VBA_MCP_ACCESS_MODE"


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


# ---------------------------------------------------------------------------
# Access-mode configuration (task_b2 / fix_002)
# ---------------------------------------------------------------------------
def test_access_mode_defaults_to_full(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_VAR, raising=False)
    monkeypatch.delenv(_ENV_ACCESS_MODE, raising=False)

    config = load_config()

    assert config.server.access_mode == AccessMode.FULL


def test_access_mode_restricted_honored(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_VAR, raising=False)
    monkeypatch.setenv(_ENV_ACCESS_MODE, "restricted")

    config = load_config()

    assert config.server.access_mode == AccessMode.RESTRICTED


def test_access_mode_full_honored(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_VAR, raising=False)
    monkeypatch.setenv(_ENV_ACCESS_MODE, "full")

    config = load_config()

    assert config.server.access_mode == AccessMode.FULL


@pytest.mark.parametrize("value", ["", "FULL", "readonly", "1", "yes"])
def test_invalid_access_mode_raises_with_log_level_unset(
    monkeypatch: pytest.MonkeyPatch, value: str
) -> None:
    # F-01: an invalid access mode must raise even when the log level is unset.
    monkeypatch.delenv(_ENV_VAR, raising=False)
    monkeypatch.setenv(_ENV_ACCESS_MODE, value)

    with pytest.raises(ValidationError):
        load_config()


def test_invalid_access_mode_raises_with_valid_log_level(monkeypatch: pytest.MonkeyPatch) -> None:
    # F-01 companion: the same failure must surface when the log level is valid.
    monkeypatch.setenv(_ENV_VAR, "DEBUG")
    monkeypatch.setenv(_ENV_ACCESS_MODE, "bogus")

    with pytest.raises(ValidationError):
        load_config()
