"""Application configuration (typed).

Configuration is loaded from environment variables and validated with Pydantic:

* ``EXCEL_VBA_MCP_LOG_LEVEL`` sets ``ServerConfig.log_level`` (one of ``DEBUG``,
  ``INFO``, ``WARNING``, ``ERROR``, ``CRITICAL``). When unset it defaults to
  ``INFO``; an empty string is invalid.
* ``EXCEL_VBA_MCP_ACCESS_MODE`` sets ``ServerConfig.access_mode`` (``full`` or
  ``restricted``). When unset it defaults to ``full``.

Both variables are validated together: an invalid value for either raises a
Pydantic ``ValidationError`` regardless of whether the other is set.

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


class ServerConfig(BaseModel):
    """Server runtime configuration."""

    log_level: LogLevel = Field(default="INFO", description="Log level for stderr logging.")
    access_mode: AccessMode = Field(
        default=AccessMode.FULL,
        description="Server access policy: 'full' allows mutations/execution by default; "
        "'restricted' is deny-by-default (approval grants required).",
    )


class Config(BaseModel):
    """Top-level configuration container."""

    server: ServerConfig = Field(default_factory=ServerConfig)


def load_config() -> Config:
    """Load configuration from defaults and the environment.

    Reads ``EXCEL_VBA_MCP_LOG_LEVEL`` and ``EXCEL_VBA_MCP_ACCESS_MODE`` when
    present and validates them together through the Pydantic model, so that an
    invalid value for either raises a ``ValidationError`` even when the other is
    unset. Unset variables fall back to their ``ServerConfig`` defaults.

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
    return Config(server=ServerConfig.model_validate(present))
