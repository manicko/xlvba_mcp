"""Application configuration (typed).

Configuration is loaded from environment variables. Sensitive values must never be
logged or written to artifacts.
"""

from __future__ import annotations

from pydantic import BaseModel, Field


class ServerConfig(BaseModel):
    """Server runtime configuration."""

    log_level: str = Field(default="INFO", description="Log level for stderr logging.")


class Config(BaseModel):
    """Top-level configuration container."""

    server: ServerConfig = Field(default_factory=ServerConfig)


def load_config() -> Config:
    """Load configuration from defaults and environment.

    Returns:
        Parsed configuration.
    """

    return Config()
