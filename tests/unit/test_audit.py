"""Unit tests for the append-only JSONL audit trail (task_009)."""

from __future__ import annotations

import json
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from excel_vba_mcp.security.audit import AUDIT_SCHEMA_VERSION, AuditLogger, redact


def test_redact_reserved_keys_case_insensitive() -> None:
    params = {
        "password": "hunter2",
        "PASSWD": "x",
        "Pwd": "y",
        "Secret": "z",
        "Token": "t",
        "connection_string": "Server=db;Password=pw",
        "Conn_Str": "c",
        "API_KEY": "k",
        "credential": "cred",
        "Auth": "bearer",
        "safe": "visible",
    }
    out = redact(params)
    for key in (
        "password",
        "PASSWD",
        "Pwd",
        "Secret",
        "Token",
        "connection_string",
        "Conn_Str",
        "API_KEY",
        "credential",
        "Auth",
    ):
        assert out[key] == "<redacted>"
    assert out["safe"] == "visible"


def test_redact_does_not_mutate_input() -> None:
    params = {"password": "p", "nested": {"token": "t"}}
    out = redact(params)
    assert params == {"password": "p", "nested": {"token": "t"}}
    assert out["nested"]["token"] == "<redacted>"


def test_redact_paths_to_basename() -> None:
    params = {"canonical_path": "C:/work/sub/Report.xlsm", "path": "/a/b/c.xlsx"}
    out = redact(params)
    assert out["canonical_path"] == "Report.xlsm"
    assert out["path"] == "c.xlsx"


def test_redact_args_summary() -> None:
    params = {"args": ["hello", 42, True, [1, 2, 3], None]}
    out = redact(params)
    summary = out["args"]
    assert summary["arg_count"] == 5
    assert summary["args"][0].startswith("str(")
    assert summary["args"][1] == "int(42)"
    assert summary["args"][2] == "bool(True)"
    assert summary["args"][3] == "list[3]"
    assert summary["args"][4] == "null"


def test_redact_nested_list_of_dicts() -> None:
    params = {"items": [{"api_key": "k", "ok": 1}]}
    out = redact(params)
    assert out["items"][0]["api_key"] == "<redacted>"
    assert out["items"][0]["ok"] == 1


def test_jsonl_lines_parse(tmp_path: Path) -> None:
    path = tmp_path / "audit" / "audit.jsonl"
    with AuditLogger(path) as logger:
        logger.record(tool="excel_open", operation_class="read_only", correlation_id="c1")
        logger.record(
            tool="macro_run",
            operation_class="executing",
            params={"macro": "Mod.Run", "password": "secret", "args": [1, 2]},
            correlation_id="c2",
            result={"status": "success", "hresult": 0, "hresult_hex": "0x00000000"},
        )

    lines = path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == 2
    first = json.loads(lines[0])
    second = json.loads(lines[1])
    assert first["schema_version"] == AUDIT_SCHEMA_VERSION
    assert first["tool"] == "excel_open"
    assert first["correlation_id"] == "c1"
    assert second["params"]["password"] == "<redacted>"
    assert second["result"]["status"] == "success"
    # timestamps are ISO-8601 with a trailing Z
    assert first["timestamp"].endswith("Z")


def test_append_mode_preserves_existing_lines(tmp_path: Path) -> None:
    path = tmp_path / "audit.jsonl"
    with AuditLogger(path) as logger:
        logger.record(tool="a", operation_class="read_only")
    with AuditLogger(path) as logger:
        logger.record(tool="b", operation_class="mutating")
    lines = path.read_text(encoding="utf-8").splitlines()
    assert [json.loads(line)["tool"] for line in lines] == ["a", "b"]


def test_audit_writes_nothing_to_stdout(tmp_path: Path, capfd: pytest.CaptureFixture[str]) -> None:
    path = tmp_path / "audit.jsonl"
    logger = AuditLogger(path)
    capfd.readouterr()  # clear

    sys.stdout.write("MARKER")
    before = capfd.readouterr().out

    logger.record(tool="vba_write", operation_class="mutating")
    logger.close()

    sys.stdout.write("MARKER")
    after = capfd.readouterr().out

    # The audit write must not emit any bytes to stdout.
    assert before == "MARKER"
    assert after == "MARKER"


def test_concurrent_writes_do_not_interleave(tmp_path: Path) -> None:
    path = tmp_path / "audit.jsonl"
    total = 64
    with AuditLogger(path) as logger:

        def emit(index: int) -> None:
            logger.record(
                tool="macro_run",
                operation_class="executing",
                correlation_id=f"corr-{index}",
            )

        with ThreadPoolExecutor(max_workers=16) as pool:
            list(pool.map(emit, range(total)))

    lines = path.read_text(encoding="utf-8").splitlines()
    assert len(lines) == total
    correlation_ids = [json.loads(line)["correlation_id"] for line in lines]
    # Every line is valid JSON and every event appears exactly once.
    assert sorted(correlation_ids) == sorted(f"corr-{i}" for i in range(total))
