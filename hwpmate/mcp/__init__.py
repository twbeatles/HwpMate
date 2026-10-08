"""HwpMate MCP 서버 패키지 (로컬 stdio 전용).

기존 GUI/CLI/변환 엔진을 재작성하지 않고, 검증된 HwpMate CLI를
subprocess로 호출하는 orchestration/adapter 계층이다.
MCP 프로세스는 한컴 COM 객체를 직접 보유하지 않는다.
"""

from __future__ import annotations

MCP_SCHEMA_VERSION = "hwpmate-mcp/v1"
MCP_SERVER_NAME = "hwpmate"
