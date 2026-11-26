"""Unit tests for typed configuration loading (ARCH-004)."""

from __future__ import annotations

import os

import pytest
from pydantic import ValidationError

from excel_vba_mcp.config import ServerConfig, load_config
from excel_vba_mcp.domain.statuses import AccessMode
from excel_vba_mcp.server import _configured_roots

_ENV_VAR = "EXCEL_VBA_MCP_LOG_LEVEL"
_ENV_ACCESS_MODE = "EXCEL_VBA_MCP_ACCESS_MODE"
_ENV_MACRO_TIMEOUT_S = "EXCEL_VBA_MCP_MACRO_TIMEOUT_S"
_ENV_OPEN_TIMEOUT_S = "EXCEL_VBA_MCP_OPEN_TIMEOUT_S"
_ENV_LIVENESS_CEILING_S = "EXCEL_VBA_MCP_LIVENESS_CEILING_S"
_ENV_ROOTS = "EXCEL_VBA_MCP_ROOTS"


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


# ---------------------------------------------------------------------------
# Macro dispatch timeout configuration (F2b)
# ---------------------------------------------------------------------------
def test_macro_timeout_defaults_to_ninety_seconds(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_MACRO_TIMEOUT_S, raising=False)

    config = load_config()

    assert config.server.macro_timeout_s == 90.0


def test_macro_timeout_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_MACRO_TIMEOUT_S, "15.5")

    config = load_config()

    assert config.server.macro_timeout_s == 15.5


@pytest.mark.parametrize("value", ["0", "-1", "nan", "not-a-number"])
def test_invalid_macro_timeout_raises(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv(_ENV_MACRO_TIMEOUT_S, value)

    with pytest.raises(ValidationError):
        load_config()


# ---------------------------------------------------------------------------
# Block E — open / liveness-ceiling configuration
# ---------------------------------------------------------------------------
def test_open_timeout_defaults_to_three_hundred_seconds(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_OPEN_TIMEOUT_S, raising=False)

    config = load_config()

    assert config.server.open_timeout_s == 300.0


def test_open_timeout_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_OPEN_TIMEOUT_S, "45.5")

    config = load_config()

    assert config.server.open_timeout_s == 45.5


@pytest.mark.parametrize("value", ["0", "-1", "nan", "not-a-number"])
def test_invalid_open_timeout_raises(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv(_ENV_OPEN_TIMEOUT_S, value)

    with pytest.raises(ValidationError):
        load_config()


def test_liveness_ceiling_defaults_to_six_hundred_seconds(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv(_ENV_LIVENESS_CEILING_S, raising=False)

    config = load_config()

    assert config.server.liveness_ceiling_s == 600.0


def test_liveness_ceiling_from_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_LIVENESS_CEILING_S, "120")

    config = load_config()

    assert config.server.liveness_ceiling_s == 120.0


@pytest.mark.parametrize("value", ["0", "-1", "nan", "not-a-number"])
def test_invalid_liveness_ceiling_raises(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv(_ENV_LIVENESS_CEILING_S, value)

    with pytest.raises(ValidationError):
        load_config()


# ---------------------------------------------------------------------------
# Roots allowlist environment parsing (F5)
#
# ``EXCEL_VBA_MCP_ROOTS`` is consumed by ``server._configured_roots()`` (the
# typed ``config.py`` model does not own roots). These tests pin the
# os.pathsep-separated parsing that lets an operator allowlist an extra project
# tree (e.g. ``C:\py_exp\mko_treport``) without a source change.
# ---------------------------------------------------------------------------
def test_roots_env_split_on_pathsep(monkeypatch: pytest.MonkeyPatch) -> None:
    first = r"C:\py_exp\mko_treport"
    second = r"C:\data\workbooks"
    monkeypatch.setenv(_ENV_ROOTS, os.pathsep.join([first, second]))

    roots = _configured_roots()

    assert roots["root_0"] == first
    assert roots["root_1"] == second
    # The artifact root is always registered in addition to the generic roots.
    assert "artifacts" in roots


def test_roots_env_empty_falls_back_to_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_ROOTS, "")

    roots = _configured_roots()

    # Default roots are CWD + system temp; only the generic + artifact keys are
    # guaranteed, so assert the shape rather than host-specific paths.
    assert "root_0" in roots
    assert "artifacts" in roots


def test_roots_env_blank_segments_ignored(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(_ENV_ROOTS, f"{os.pathsep}C:\\only{os.pathsep}  ")

    roots = _configured_roots()

    assert roots["root_0"] == r"C:\only"
    assert "root_1" not in roots


# ---------------------------------------------------------------------------
# Block C — resolved roots are logged at startup (no silent fallback)
# ---------------------------------------------------------------------------
def test_configured_roots_logs_resolved_roots(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    first = r"C:\py_exp\mko_treport"
    monkeypatch.setenv(_ENV_ROOTS, first)

    with caplog.at_level("INFO", logger="excel_vba_mcp.server"):
        roots = _configured_roots()

    logged = "\n".join(record.getMessage() for record in caplog.records)
    assert "Allowed roots:" in logged
    # The resolved root appears in the log line (dict repr escapes backslashes,
    # so assert on a path fragment rather than the raw path).
    assert "mko_treport" in logged
    assert roots["root_0"] == first


def test_configured_roots_logs_fallback_roots(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    # Unset roots => cwd+tempdir fallback must still be logged (not silent).
    monkeypatch.setenv(_ENV_ROOTS, "")

    with caplog.at_level("INFO", logger="excel_vba_mcp.server"):
        _configured_roots()

    logged = "\n".join(record.getMessage() for record in caplog.records)
    assert "Allowed roots:" in logged
    assert "root_0" in logged
