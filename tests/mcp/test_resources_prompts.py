from __future__ import annotations

import asyncio
import json
import sys

from hwpmate.mcp.config import McpConfig
from hwpmate.mcp.server import create_server


def _run(coro):
    return asyncio.run(coro)


def _resource_text(result) -> str:
    assert isinstance(result, list), type(result).__name__
    assert result, "empty resource result"
    first = result[0]
    content = getattr(first, "content", None)
    if isinstance(content, str):
        return content
    text = getattr(first, "text", None)
    assert isinstance(text, str), type(first).__name__
    return text


def _prompt_text(result) -> str:
    from mcp.types import GetPromptResult, TextContent

    assert isinstance(result, GetPromptResult), type(result).__name__
    assert result.messages, "empty prompt result"
    content = result.messages[0].content
    assert isinstance(content, TextContent), type(content).__name__
    return content.text


def test_mcp_does_not_import_gui_or_lock_modules() -> None:
    import subprocess

    probe = (
        "import sys; "
        "import hwpmate.mcp.server, hwpmate.mcp.jobs, hwpmate.mcp.planner_adapter; "
        "bad = [m for m in ('hwpmate.app', 'hwpmate.app_instance', 'hwpmate.ui.main_window') if m in sys.modules]; "
        "sys.exit('GUI modules imported: ' + ','.join(bad) if bad else 0)"
    )
    proc = subprocess.run([sys.executable, "-c", probe], capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stderr or "GUI modules imported"


def test_resources_list_and_read() -> None:
    server = create_server(McpConfig())

    async def main():
        resources = await server.list_resources()
        caps = await server.read_resource("hwpmate://capabilities")
        formats = await server.read_resource("hwpmate://formats")
        missing = await server.read_resource("hwpmate://jobs/no-such-job")
        return resources, caps, formats, missing

    resources, caps, formats, missing = _run(main())
    uris = {str(r.uri) for r in resources}
    assert "hwpmate://capabilities" in uris
    assert "hwpmate://formats" in uris
    assert json.loads(_resource_text(caps))["ok"] is True
    assert json.loads(_resource_text(formats))["ok"] is True
    payload = json.loads(_resource_text(missing))
    assert payload["ok"] is False
    assert payload["error"]["code"] == "JOB_NOT_FOUND"


def test_job_resource_after_submit(tmp_path, monkeypatch) -> None:
    from pathlib import Path

    from hwpmate.mcp.jobs import JobManager, wait_for_state
    from hwpmate.mcp.planner_adapter import get_plan, preview_conversion
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
    manager = JobManager(cfg)
    preview = preview_conversion(cfg, input_paths=[str(src)], format_type="PDF", output_dir=str(tmp_path / "out"))
    record = manager.submit(get_plan(preview["plan_id"]), confirmation=CONFIRMATION_TOKEN, idempotency_key="res1")
    wait_for_state(manager, record.job_id, timeout_seconds=60)

    server = create_server(cfg)

    async def main():
        return await server.read_resource(f"hwpmate://jobs/{record.job_id}")

    payload = json.loads(_resource_text(_run(main())))
    assert payload["job_id"] == record.job_id
    assert payload["state"] == "succeeded"


def test_prompts_registered() -> None:
    server = create_server(McpConfig())

    async def main():
        prompts = await server.list_prompts()
        first = await server.get_prompt("convert_hwp_to_pdf_safely")
        return prompts, first

    prompts, first = _run(main())
    names = {p.name for p in prompts}
    assert names == {
        "convert_hwp_to_pdf_safely",
        "prepare_documents_for_office",
        "review_conversion_failures",
    }
    assert "hwpmate_preview_conversion" in _prompt_text(first)
