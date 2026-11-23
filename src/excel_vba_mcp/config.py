"""Application configuration (typed).

Configuration is loaded from environment variables and validated with Pydantic:

* ``EXCEL_VBA_MCP_LOG_LEVEL`` sets ``ServerConfig.log_level`` (one of ``DEBUG``,
  ``INFO``, ``WARNING``, ``ERROR``, ``CRITICAL``). When unset it defaults to
  ``INFO``; an empty string is invalid.
* ``EXCEL_VBA_MCP_ACCESS_MODE`` sets ``ServerConfig.access_mode`` (``full`` or
  ``restricted``). When unset it defaults to ``full``.
* ``EXCEL_VBA_MCP_MACRO_TIMEOUT_S`` sets ``ServerConfig.macro_timeout_s`` (the
  soft per-call ceiling, in seconds, for the ``Application.Run`` macro dispatch).
  When unset it defaults to ``90.0``; a non-positive value is invalid.

All variables are validated together: an invalid value for any of them raises a
Pydantic ``ValidationError`` regardless of whether the others are set.

Sensitive values must never be logged or written to artifacts.
"""

from __future__ import annotations

import os
from typing import Literal

from pydantic import BaseModel, Field

from excel_vba_mcp.domain.statuses import AccessMode

LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]

_ENV_LOG_LEVEL = "EXCEL_VBA_MCP_LOG_LEVEL"
_ENV_ACCESS_MODE = "EXCEL_VBA_MCP_ACCESS_MODE"
_ENV_MACRO_TIMEOUT_S = "EXCEL_VBA_MCP_MACRO_TIMEOUT_S"


class ServerConfig(BaseModel):
    """Server runtime configuration."""

    log_level: LogLevel = Field(default="INFO", description="Log level for stderr logging.")
    access_mode: AccessMode = Field(
        default=AccessMode.FULL,
        description="Server access policy: 'full' allows mutations/execution by default; "
        "'restricted' is deny-by-default (approval grants required).",
    )
    macro_timeout_s: float = Field(
        default=90.0,
        gt=0,
        description="Soft per-call ceiling (seconds) for the Application.Run macro dispatch "
        "(F2b). Exceeding it classifies the run as MACRO_TIMEOUT, terminates only the "
        "automation-owned Excel PID and recovers the session in-process.",
    )


class Config(BaseModel):
    """Top-level configuration container."""

    server: ServerConfig = Field(default_factory=ServerConfig)


def load_config() -> Config:
    """Load configuration from defaults and the environment.

    Reads ``EXCEL_VBA_MCP_LOG_LEVEL``, ``EXCEL_VBA_MCP_ACCESS_MODE`` and
    ``EXCEL_VBA_MCP_MACRO_TIMEOUT_S`` when present and validates them together
    through the Pydantic model, so that an invalid value for any of them raises a
    ``ValidationError`` even when the others are unset. Unset variables fall back
    to their ``ServerConfig`` defaults.

    Returns:
        Parsed configuration.

    Raises:
        pydantic.ValidationError: If an environment value fails validation.
    """

    present: dict[str, object] = {}
    log_level = os.environ.get(_ENV_LOG_LEVEL)
    if log_level is not None:
        present["log_level"] = log_level
    access_mode = os.environ.get(_ENV_ACCESS_MODE)
    if access_mode is not None:
        present["access_mode"] = access_mode
    macro_timeout_s = os.environ.get(_ENV_MACRO_TIMEOUT_S)
    if macro_timeout_s is not None:
        present["macro_timeout_s"] = macro_timeout_s
    return Config(server=ServerConfig.model_validate(present))
