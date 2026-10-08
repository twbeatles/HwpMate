from __future__ import annotations

import asyncio
import json
import subprocess
import sys

from hwpmate.constants import FORMAT_TYPES
from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.server import create_server


def _run(coro):
    return asyncio.run(coro)


def _tool_text(result) -> str:
    from mcp.types import CallToolResult, TextContent

    assert isinstance(result, CallToolResult), type(result).__name__
    assert result.content, "empty tool result"
    first = result.content[0]
    assert isinstance(first, TextContent), type(first).__name__
    assert isinstance(first.text, str)
    return first.text


def test_tools_listed() -> None:
    server = create_server(McpConfig())

    async def main():
        return await server.list_tools()

    tools = _run(main())
    names = [tool.name for tool in tools]
    assert names == [
        "hwpmate_get_capabilities",
        "hwpmate_list_supported_formats",
        "hwpmate_preview_conversion",
        "hwpmate_submit_conversion",
        "hwpmate_get_job_status",
        "hwpmate_get_job_result",
        "hwpmate_cancel_job",
        "hwpmate_validate_artifacts",
    ]


def test_tool_annotations_match_design() -> None:
    server = create_server(McpConfig())

    async def main():
        return await server.list_tools()

    hints = {tool.name: tool.annotations for tool in _run(main())}
    read_only_tools = {
        "hwpmate_get_capabilities",
        "hwpmate_list_supported_formats",
        "hwpmate_preview_conversion",
        "hwpmate_get_job_status",
        "hwpmate_get_job_result",
        "hwpmate_validate_artifacts",
    }
    for name in read_only_tools:
        annotation = hints[name]
        assert annotation is not None, name
        assert annotation.read_only_hint is True, name
    submit = hints["hwpmate_submit_conversion"]
    assert submit is not None
    assert submit.read_only_hint is False
    cancel = hints["hwpmate_cancel_job"]
    assert cancel is not None
    assert cancel.read_only_hint is False
    assert cancel.destructive_hint is True


def test_call_list_formats_structured() -> None:
    server = create_server(McpConfig())

    async def main():
        return await server.call_tool("hwpmate_list_supported_formats", {})

    result = _run(main())
    payload = json.loads(_tool_text(result))
    assert payload["ok"] is True
    assert {item["name"] for item in payload["formats"]} == set(FORMAT_TYPES.keys())


def test_submit_status_result_via_tools(tmp_path, monkeypatch) -> None:
    import time
    from pathlib import Path

    from hwpmate.mcp.jobs import TERMINAL_STATES
    from hwpmate.mcp.schemas import CONFIRMATION_TOKEN

    monkeypatch.setenv("HWPMATE_FAKECLI_MODE", "ok")
    src = tmp_path / "in"
    src.mkdir()
    (src / "a.hwp").write_bytes(b"x")
    cfg = McpConfig(
        input_roots=[str(src)],
        output_roots=[str(tmp_path / "out")],
        cli_executable=str(Path(__file__).parent / "fake_cli.py"),
        job_dir=str(tmp_path / "jobs"),
    )
    server = create_server(cfg)

    async def call(name, args):
        result = await server.call_tool(name, args)
        import json as _json

        return _json.loads(_tool_text(result))

    preview = _run(call("hwpmate_preview_conversion", {"input_paths": [str(src)], "output_dir": str(tmp_path / "out")}))
    assert preview["ok"] is True
    submitted = _run(
        call(
            "hwpmate_submit_conversion",
            {"plan_id": preview["plan_id"], "confirmation": CONFIRMATION_TOKEN, "idempotency_key": "srv1"},
        )
    )
    assert submitted["ok"] is True
    job_id = submitted["job_id"]
    deadline = time.time() + 60
    status: dict = {}
    while time.time() < deadline:
        status = _run(call("hwpmate_get_job_status", {"job_id": job_id}))
        if status["state"] in TERMINAL_STATES:
            break
        time.sleep(0.2)
    assert status["state"] == "succeeded"
    result = _run(call("hwpmate_get_job_result", {"job_id": job_id, "limit": 1, "cursor": 0}))
    assert result["total"] == 1
    assert len(result["items"]) == 1
    assert result["next_cursor"] is None


def test_module_help_does_not_start_server() -> None:
    proc = subprocess.run(
        [sys.executable, "-m", "hwpmate.mcp", "--help"],
        capture_output=True,
        text=True,
        timeout=60,
    )
    assert proc.returncode == 0
    assert "stdio" in proc.stdout
    assert proc.stderr == ""
