"""Basic smoke tests for package bootstrap."""

from excel_vba_mcp import __version__


def test_version() -> None:
    assert __version__ == "0.1.0"
