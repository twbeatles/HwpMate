from __future__ import annotations

import asyncio
import json
import time
from pathlib import Path

import pytest

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.errors import McpError
from hwpmate.mcp.jobs import JobManager, wait_for_state
from hwpmate.mcp.planner_adapter import get_plan, preview_conversion
from hwpmate.mcp.schemas import CONFIRMATION_TOKEN
from hwpmate.mcp.server import create_server
from hwpmate.mcp.validator import validate_job_outputs

FAKE = str(Path(__file__).parent / "fake_cli.py")


def _cfg(tmp_path: Path, **overrides: object) -> McpConfig:
    base: dict[str, object] = {
        "input_roots": [str(tmp_path / "in")],
        "output_roots": [str(tmp_path / "out")],
        "cli_executable": FAKE,
        "job_dir": str(tmp_path / "jobs"),
        "job_timeout_seconds": 60,
    }
    base.update(overrides)
    return McpConfig.from_mapping(base)


def _run_ok_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    (src / "a.hwp").write_bytes(b"x")
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=out)
    record = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="v1")
    done = wait_for_state(manager, record.job_id, timeout_seconds=60)
    assert done.state == "succeeded"
    return manager, done


def test_validate_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager, done = _run_ok_job(tmp_path, monkeypatch)
    result = validate_job_outputs(manager, done.job_id)
    assert result["ok"] is True
    assert result["all_verified"] is True
    assert result["checked"] == 1
    assert result["invalid"] == []


def test_validate_missing_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager, done = _run_ok_job(tmp_path, monkeypatch)
    for output in done.items[0]["outputs"]:
        Path(output).unlink()
    result = validate_job_outputs(manager, done.job_id)
    assert result["all_verified"] is False
    assert result["invalid"][0]["reasons"] == ["파일 없음"]


def test_validate_bad_signature(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager, done = _run_ok_job(tmp_path, monkeypatch)
    for output in done.items[0]["outputs"]:
        Path(output).write_bytes(b"not a pdf at all")
    result = validate_job_outputs(manager, done.job_id)
    assert result["all_verified"] is False
    assert any("서명" in reason for reason in result["invalid"][0]["reasons"])


def test_validate_unfinished_job(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "10")
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    (src / "a.hwp").write_bytes(b"x")
    cfg = _cfg(tmp_path)
    manager = JobManager(cfg)
    out = str(tmp_path / "out")
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=out)
    record = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="u1")
    with pytest.raises(McpError) as excinfo:
        validate_job_outputs(manager, record.job_id)
    assert excinfo.value.code == "JOB_NOT_DONE"


def test_cancel_tool_via_server(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "sleep")
    monkeypatch.setenv("HWPMATE_FAKECLI_SLEEP", "10")
    src = tmp_path / "in"
    src.mkdir(parents=True, exist_ok=True)
    (src / "a.hwp").write_bytes(b"x")
    (src / "b.hwp").write_bytes(b"x")
    cfg = _cfg(tmp_path)
    server = create_server(cfg)

    async def call(name, args):
        from mcp.types import CallToolResult, TextContent

        result = await server.call_tool(name, args)
        assert isinstance(result, CallToolResult), type(result).__name__
        assert result.content, "empty tool result"
        first = result.content[0]
        assert isinstance(first, TextContent), type(first).__name__
        return json.loads(first.text)

    def run(coro):
        return asyncio.run(coro)

    out = str(tmp_path / "out")
    first = run(call("hwpmate_preview_conversion", {"input_paths": [str(src / "a.hwp")], "output_dir": out}))
    second = run(call("hwpmate_preview_conversion", {"input_paths": [str(src / "b.hwp")], "output_dir": out}))
    job_a = run(call("hwpmate_submit_conversion", {"plan_id": first["plan_id"], "confirmation": CONFIRMATION_TOKEN, "idempotency_key": "ca"}))
    job_b = run(call("hwpmate_submit_conversion", {"plan_id": second["plan_id"], "confirmation": CONFIRMATION_TOKEN, "idempotency_key": "cb"}))
    status: dict = {}
    canceled = run(call("hwpmate_cancel_job", {"job_id": job_b["job_id"], "confirmation": CONFIRMATION_TOKEN}))
    assert canceled["ok"] is True
    assert canceled["state"] == "canceled"
    validated = run(call("hwpmate_validate_artifacts", {"job_id": job_b["job_id"]}))
    assert validated["ok"] is True
    assert validated["checked"] == 0
    deadline = time.time() + 30
    while time.time() < deadline:
        status = run(call("hwpmate_get_job_status", {"job_id": job_a["job_id"]}))
        if status["state"] not in ("queued", "running"):
            break
        time.sleep(0.2)
