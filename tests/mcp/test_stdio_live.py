"""실제 stdio 채널 검증: 서버 프로세스를 자식으로 띄워 JSON-RPC가 깨지지 않음을 증명한다.

설계서 §9 'stdio 독립성' — stdout에 JSON-RPC 외 문구가 섞이면 핸드셰이크가
실패하므로, 이 테스트 통과 자체가 무오염 증거다.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _params() -> object:
    from mcp.client.stdio import StdioServerParameters

    return StdioServerParameters(
        command=sys.executable,
        args=["-m", "hwpmate.mcp"],
        cwd=str(ROOT),
        env=dict(os.environ),
    )


def test_live_stdio_tools_and_capabilities() -> None:
    import anyio

    from mcp.client.session import ClientSession
    from mcp.client.stdio import stdio_client
    from mcp.types import CallToolResult, TextContent

    completed = False

    async def main() -> None:
        nonlocal completed
        async with stdio_client(_params()) as (read_stream, write_stream):  # type: ignore[arg-type]
            async with ClientSession(read_stream, write_stream) as session:
                await session.initialize()
                tools = await session.list_tools()
                names = [tool.name for tool in tools.tools]
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
                result = await session.call_tool("hwpmate_get_capabilities", {})
                assert isinstance(result, CallToolResult), type(result).__name__
                first = result.content[0]
                assert isinstance(first, TextContent), type(first).__name__
                payload = json.loads(first.text)
                assert payload["schema_version"] == "hwpmate-mcp/v1"
                assert payload["hancom_com_probe"] == "not_run"
                completed = True

    async def bounded() -> None:
        with anyio.fail_after(90):
            await main()

    anyio.run(bounded)
    assert completed
