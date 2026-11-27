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
* ``EXCEL_VBA_MCP_OPEN_TIMEOUT_S`` sets ``ServerConfig.open_timeout_s`` (the
  server-side deadline, in seconds, for the ``excel_open`` STA open). When unset
  it defaults to ``300.0``; a non-positive value is invalid (Block E).
* ``EXCEL_VBA_MCP_LIVENESS_CEILING_S`` sets ``ServerConfig.liveness_ceiling_s``
  (the server-wide upper bound, in seconds, applied to every per-call COM
  deadline). When unset it defaults to ``600.0``; a non-positive value is
  invalid (Block E).

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
_ENV_OPEN_TIMEOUT_S = "EXCEL_VBA_MCP_OPEN_TIMEOUT_S"
_ENV_LIVENESS_CEILING_S = "EXCEL_VBA_MCP_LIVENESS_CEILING_S"
_ENV_LARGE_DATA = "EXCEL_VBA_MCP_LARGE_DATA_"


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
        "(F2b). Exceeding it classifies the run as a deadline-exceeded (SUSPECTED_STALLED) "
        "outcome; the server never auto-terminates a live pid (Block E / C1).",
    )
    open_timeout_s: float = Field(
        default=300.0,
        gt=0,
        description="Server-side deadline (seconds) for the excel_open STA open (Block E). "
        "This is the per-op-class server authority for the open path.",
    )
    liveness_ceiling_s: float = Field(
        default=600.0,
        gt=0,
        description="Server-wide upper bound (seconds) applied to every per-call COM "
        "deadline (Block E). A per-call timeout is clamped to this ceiling so no single "
        "operation's deadline exceeds the liveness ceiling.",
    )


class LargeDataConfig(BaseModel):
    """Unified large-data limits (single source; no per-tool limits).

    Defaults are conservative and justified by benchmarks; they are changed only
    via ``EXCEL_VBA_MCP_LARGE_DATA_*`` env vars, never raised to pass a test.
    Capability comes from bounded/streaming processing, not larger allocations.
    """

    max_range_cells: int = Field(
        default=100_000,
        description="Per-request read window / per-write payload cell cap.",
    )
    max_response_cells: int = Field(
        default=50_000,
        description="Max cells returned in one MCP response.",
    )
    max_response_bytes: int = Field(
        default=8 * 1024 * 1024,
        description="Serialized response byte budget for a single page.",
    )
    max_table_cells: int = Field(
        default=50_000,
        description="Max cells returned by table read-rows.",
    )
    max_pivot_result_cells: int = Field(
        default=50_000,
        description="Max cells in one PivotTable rendered-result page.",
    )
    max_pivot_members: int = Field(
        default=10_000,
        description="Max PivotItems enumerated for a field/member read.",
    )
    max_text_len: int = Field(
        default=32_767,
        description="Max characters returned for one text field (Excel cell max).",
    )
    max_file_mb: int = Field(
        default=100,
        description="Max OOXML input file size (MiB) before streaming.",
    )
    max_uncompressed_mb: int = Field(
        default=1024,
        description="Max total uncompressed OOXML size (MiB) — ZIP-bomb guard.",
    )
    max_compression_ratio: int = Field(
        default=1000,
        description="Max per-member compression ratio — ZIP-bomb guard.",
    )
    max_zip_members: int = Field(
        default=10_000,
        description="Max ZIP archive member count.",
    )
    max_shared_strings: int = Field(
        default=1_000_000,
        description="Max sharedStrings entries streamed.",
    )
    max_xml_depth: int = Field(
        default=64,
        description="Max XML nesting depth for the streaming parser.",
    )
    max_export_rows: int = Field(
        default=10_000_000,
        description="Max rows streamed to a file export.",
    )
    max_file_jobs: int = Field(
        default=2,
        description="Max concurrent non-COM file-processing jobs.",
    )


class Config(BaseModel):
    """Top-level configuration container."""

    server: ServerConfig = Field(default_factory=ServerConfig)
    large_data: LargeDataConfig = Field(default_factory=LargeDataConfig)


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
    open_timeout_s = os.environ.get(_ENV_OPEN_TIMEOUT_S)
    if open_timeout_s is not None:
        present["open_timeout_s"] = open_timeout_s
    liveness_ceiling_s = os.environ.get(_ENV_LIVENESS_CEILING_S)
    if liveness_ceiling_s is not None:
        present["liveness_ceiling_s"] = liveness_ceiling_s
    server = ServerConfig.model_validate(present)

    # Large-data limits: each field is overridable via
    # ``EXCEL_VBA_MCP_LARGE_DATA_<UPPER_FIELD>`` (e.g.
    # ``EXCEL_VBA_MCP_LARGE_DATA_MAX_RANGE_CELLS``). Unset fields keep defaults.
    large_present: dict[str, object] = {}
    for field_name in LargeDataConfig.model_fields:
        raw = os.environ.get(f"{_ENV_LARGE_DATA}{field_name.upper()}")
        if raw is not None:
            large_present[field_name] = raw
    large_data = LargeDataConfig.model_validate(large_present)
    return Config(server=server, large_data=large_data)
