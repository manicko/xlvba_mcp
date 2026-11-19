"""Application configuration (typed).

Configuration is loaded from environment variables and validated with Pydantic.
Currently the only supported variable is ``EXCEL_VBA_MCP_LOG_LEVEL``, which sets
``ServerConfig.log_level`` (one of ``DEBUG``, ``INFO``, ``WARNING``, ``ERROR``,
``CRITICAL``). When the variable is unset the level defaults to ``INFO``. An
invalid value raises a Pydantic ``ValidationError`` with a clear message.

Sensitive values must never be logged or written to artifacts.
"""

from __future__ import annotations

import os
from typing import Literal

from pydantic import BaseModel, Field

LogLevel = Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]

_ENV_LOG_LEVEL = "EXCEL_VBA_MCP_LOG_LEVEL"


class ServerConfig(BaseModel):
    """Server runtime configuration."""

    log_level: LogLevel = Field(default="INFO", description="Log level for stderr logging.")


class Config(BaseModel):
    """Top-level configuration container."""

    server: ServerConfig = Field(default_factory=ServerConfig)


def load_config() -> Config:
    """Load configuration from defaults and the environment.

    Reads ``EXCEL_VBA_MCP_LOG_LEVEL`` when present and validates it through the
    Pydantic model so that an invalid value raises a ``ValidationError``. When the
    variable is unset, ``ServerConfig`` falls back to its ``INFO`` default.

    Returns:
        Parsed configuration.

    Raises:
        pydantic.ValidationError: If an environment value fails validation.
    """

    env_log_level = os.environ.get(_ENV_LOG_LEVEL)
    if env_log_level is None:
        return Config()
    server = ServerConfig.model_validate({"log_level": env_log_level})
    return Config(server=server)
