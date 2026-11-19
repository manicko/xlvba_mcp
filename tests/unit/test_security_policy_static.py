"""Static policy assertions for the security package (task_009).

These tests scan the *source text* of the task_009 modules and assert that no
Trust Center / registry mutation or automation-security assignment can hide in
them (spec §7.2, §10, §10.1). The rule is deliberately textual: the security
policy must never write Office trust settings or set the application automation
security property — deciding the value is pure; applying it belongs to the
session service.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

_SECURITY_DIR = Path(__file__).resolve().parents[2] / "src" / "excel_vba_mcp" / "security"
_SOURCE_FILES = ["macro_policy.py", "approval.py", "audit.py"]

# Patterns that must never appear in the security sources. Each is an assignment
# or a bare token that would indicate a Trust Center / registry / automation
# security mutation.
_FORBIDDEN_PATTERNS: dict[str, re.Pattern[str]] = {
    "winreg import/use": re.compile(r"winreg"),
    "AccessVBOM assignment": re.compile(r"AccessVBOM\s*="),
    "VBAWarnings assignment": re.compile(r"VBAWarnings\s*="),
    "TrustedLocation(s) reference": re.compile(r"TrustedLocations?"),
    "AutomationSecurity assignment": re.compile(r"AutomationSecurity\s*="),
}


@pytest.fixture(scope="module")
def sources() -> dict[str, str]:
    texts: dict[str, str] = {}
    for name in _SOURCE_FILES:
        path = _SECURITY_DIR / name
        assert path.exists(), f"missing security source: {path}"
        texts[name] = path.read_text(encoding="utf-8")
    return texts


@pytest.mark.parametrize("name", _SOURCE_FILES)
def test_no_forbidden_patterns(sources: dict[str, str], name: str) -> None:
    text = sources[name]
    for label, pattern in _FORBIDDEN_PATTERNS.items():
        match = pattern.search(text)
        assert match is None, f"{name} contains forbidden {label}: {match.group(0)!r}"


def test_macro_policy_does_not_import_winreg(sources: dict[str, str]) -> None:
    text = sources["macro_policy.py"]
    for line in text.splitlines():
        stripped = line.strip()
        assert not stripped.startswith("import winreg")
        assert not stripped.startswith("from winreg")
        assert "import winreg" not in stripped
        assert "from winreg" not in stripped
