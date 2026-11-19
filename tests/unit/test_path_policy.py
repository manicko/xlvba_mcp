"""Unit tests for the canonical path policy (task_008, spec §7.1, §10)."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

from excel_vba_mcp.domain.errors import ErrorCode, PathNotAllowedError
from excel_vba_mcp.domain.statuses import OperationClass
from excel_vba_mcp.security.path_policy import PathPolicy

_READ = OperationClass.READ_ONLY


@pytest.fixture
def roots(tmp_path: Path) -> dict[str, Path]:
    """Build fake allowlisted roots under ``tmp_path``."""

    projects = tmp_path / "projects"
    work = tmp_path / "work"
    snapshots = tmp_path / "snapshots"
    for directory in (projects, work, snapshots):
        directory.mkdir()
    return {"projects": projects, "work": work, "snapshots": snapshots}


def _assert_rejected(
    policy: PathPolicy,
    raw: str | Path,
    reason: str,
    *,
    operation: OperationClass | str = _READ,
    expected_root: str | None = None,
) -> None:
    with pytest.raises(PathNotAllowedError) as excinfo:
        policy.validate_path(raw, operation, expected_root=expected_root)
    error = excinfo.value
    assert error.code == ErrorCode.PATH_NOT_ALLOWED
    assert error.code == "PATH_NOT_ALLOWED"
    assert error.context["reason"] == reason


def test_allowed_path_returns_canonical_path(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    book = roots["projects"] / "Report.xlsm"
    book.write_text("fixture")

    result = policy.validate_path(str(book), _READ)

    assert result.is_absolute()
    assert result == Path(os.path.realpath(book))
    assert result.name == "Report.xlsm"


def test_plain_parent_traversal_rejected(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    escaping = str(roots["projects"] / ".." / "evil.xlsm")

    _assert_rejected(policy, escaping, "traversal")


def test_absolute_path_outside_roots_rejected(roots: dict[str, Path], tmp_path: Path) -> None:
    policy = PathPolicy(roots)
    outside = tmp_path / "outside"
    outside.mkdir()

    _assert_rejected(policy, str(outside / "evil.xlsm"), "outside_roots")


def test_unc_path_denied_by_default(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)

    _assert_rejected(policy, r"\\server\share\book.xlsm", "unc_denied")


def test_device_path_denied(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)

    _assert_rejected(policy, r"\\?\C:\Windows\evil.xlsm", "device_path_denied")


def test_disallowed_extension_rejected(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    note = roots["projects"] / "notes.txt"

    _assert_rejected(policy, str(note), "extension_not_allowed")


def test_json_allowed_in_snapshots_but_not_projects(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    manifest = roots["snapshots"] / "manifest.json"

    assert policy.validate_path(str(manifest), _READ) == Path(os.path.realpath(manifest))

    _assert_rejected(policy, str(roots["projects"] / "manifest.json"), "extension_not_allowed")


def test_expected_root_mismatch_rejected(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    book = roots["work"] / "Report.xlsm"

    _assert_rejected(
        policy,
        str(book),
        "root_mismatch",
        expected_root="projects",
    )


def test_expected_root_match_accepts(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    book = roots["projects"] / "Report.xlsm"

    result = policy.validate_path(str(book), _READ, expected_root="projects")

    assert result == Path(os.path.realpath(book))


def test_operation_accepts_string_value(roots: dict[str, Path]) -> None:
    policy = PathPolicy(roots)
    book = roots["work"] / "Report.xlsm"

    assert policy.validate_path(str(book), "mutating") == Path(os.path.realpath(book))


def test_allow_unc_permits_unc_path_under_unc_root() -> None:
    unc_root = Path(r"\\server\share")
    policy = PathPolicy({"unc": unc_root}, allow_unc=True)

    result = policy.validate_path(r"\\server\share\book.xlsm", _READ)

    assert result.is_absolute()
    assert result == Path(os.path.realpath(r"\\server\share\book.xlsm"))


def _make_escape_link(link: Path, target: Path) -> None:
    """Create a directory symlink or junction at ``link`` pointing to ``target``.

    Falls back to a directory junction (``mklink /J``) when symlink creation is
    not permitted, and skips the test when neither mechanism is available.
    """

    try:
        link.symlink_to(target, target_is_directory=True)
        return
    except (OSError, NotImplementedError):
        pass

    completed = subprocess.run(  # noqa: S603 - fixed, trusted executable
        ["cmd", "/c", "mklink", "/J", str(link), str(target)],
        capture_output=True,
        check=False,
    )
    if completed.returncode != 0:
        pytest.skip(
            "cannot create a symlink or directory junction on this host "
            f"(mklink exit {completed.returncode})"
        )


def test_symlink_escape_rejected(roots: dict[str, Path], tmp_path: Path) -> None:
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "evil.xlsm").write_text("fixture")
    escape_link = roots["projects"] / "escape"
    _make_escape_link(escape_link, outside)

    policy = PathPolicy(roots)

    _assert_rejected(policy, str(escape_link / "evil.xlsm"), "symlink_escape")
