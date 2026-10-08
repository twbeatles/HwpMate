"""MCP stdio 서버: 도구 선언과 진입점 (P0: 읽기 전용 3개 도구).

- stdin/stdout은 MCP JSON-RPC 전용이다. 변환 로그·print를 stdout에 쓰지 않는다.
- hwpmate.app (PyQt GUI/CLI 진입점)을 import하지 않는다.
"""

from __future__ import annotations

import json
from typing import Any

from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

from . import MCP_SERVER_NAME
from .capabilities import gather_capabilities, list_supported_formats
from .config import McpConfig
from .errors import McpError, error_envelope
from .jobs import JobManager
from .logging_utils import setup_stderr_logging
from .planner_adapter import get_plan, preview_conversion
from .validator import validate_job_outputs

from . import MCP_SCHEMA_VERSION
from .schemas import CONFIRMATION_TOKEN, ok_envelope

logger = setup_stderr_logging()

_READ_ONLY = ToolAnnotations(read_only_hint=True)
_WRITABLE = ToolAnnotations(read_only_hint=False, destructive_hint=False, idempotent_hint=False)
_DESTRUCTIVE = ToolAnnotations(read_only_hint=False, destructive_hint=True, idempotent_hint=False)


def create_server(cfg: McpConfig) -> MCPServer:
    server = MCPServer(MCP_SERVER_NAME)
    jobs = JobManager(cfg)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_get_capabilities() -> dict[str, Any]:
        """HwpMate 변환 환경 능력 조회 (읽기 전용)."""
        try:
            return gather_capabilities(cfg)
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_list_supported_formats() -> dict[str, Any]:
        """지원 출력 포맷 목록 조회 (읽기 전용)."""
        try:
            return list_supported_formats()
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_preview_conversion(
        input_paths: list[str],
        format: str = "PDF",
        output_dir: str = "",
        recursive: bool = False,
    ) -> dict[str, Any]:
        """변환 미리보기: 파일 생성·문서 변환 없이 계획과 plan_id 반환 (읽기 전용)."""
        try:
            return preview_conversion(
                cfg,
                input_paths=list(input_paths or []),
                format_type=format,
                output_dir=output_dir,
                recursive=bool(recursive),
            )
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_WRITABLE)
    def hwpmate_submit_conversion(
        plan_id: str,
        confirmation: str = "",
        idempotency_key: str = "",
    ) -> dict[str, Any]:
        """승인된 계획을 job으로 제출한다. 새 파일·백업·로그가 생성된다."""
        try:
            if confirmation != CONFIRMATION_TOKEN:
                raise McpError("CONFIRMATION_REQUIRED", "명시적 승인 토큰이 필요합니다.")
            plan = get_plan(plan_id)
            record = jobs.submit(plan, confirmation=confirmation, idempotency_key=idempotency_key)
            return ok_envelope(job_id=record.job_id, state=record.state, format=record.format)
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_get_job_status(job_id: str) -> dict[str, Any]:
        """job 상태와 성공/실패 건수 조회 (읽기 전용)."""
        try:
            return jobs.get(job_id).to_status_dict()
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_get_job_result(
        job_id: str, limit: int = 50, cursor: int = 0
    ) -> dict[str, Any]:
        """job 개별 파일 결과 페이지 조회 (읽기 전용)."""
        try:
            record = jobs.get(job_id)
            total = len(record.items)
            start = max(0, int(cursor))
            size = max(1, min(200, int(limit)))
            page = record.items[start : start + size]
            next_cursor = start + size if start + size < total else None
            return ok_envelope(
                job_id=record.job_id,
                state=record.state,
                total=total,
                items=page,
                next_cursor=next_cursor,
                warnings=list(record.warnings),
                error=dict(record.error),
            )
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_DESTRUCTIVE)
    def hwpmate_cancel_job(job_id: str, confirmation: str = "") -> dict[str, Any]:
        """대기 중 job 취소. 실행 중 취소는 안전상 지원하지 않는다."""
        try:
            record = jobs.cancel(job_id, confirmation=confirmation)
            return ok_envelope(job_id=record.job_id, state=record.state)
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.tool(annotations=_READ_ONLY)
    def hwpmate_validate_artifacts(job_id: str) -> dict[str, Any]:
        """job 산출물의 존재·크기·서명 검증 (읽기 전용)."""
        try:
            return validate_job_outputs(jobs, job_id)
        except McpError as exc:
            return error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION)

    @server.resource("hwpmate://capabilities", description="현재 서버의 변환 가능 환경")
    def resource_capabilities() -> str:
        return json.dumps(gather_capabilities(cfg), ensure_ascii=False, indent=2)

    @server.resource("hwpmate://formats", description="출력 포맷과 특징")
    def resource_formats() -> str:
        return json.dumps(list_supported_formats(), ensure_ascii=False, indent=2)

    @server.resource("hwpmate://jobs/{job_id}", description="제출한 job 요약")
    def resource_job(job_id: str) -> str:
        try:
            return json.dumps(jobs.get(job_id).to_status_dict(), ensure_ascii=False, indent=2)
        except McpError as exc:
            return json.dumps(
                error_envelope(exc.code, exc.message, schema_version=MCP_SCHEMA_VERSION),
                ensure_ascii=False,
                indent=2,
            )

    @server.prompt("convert_hwp_to_pdf_safely")
    def prompt_convert_safely() -> str:
        return (
            "HWP/HWPX를 PDF로 안전하게 변환하는 절차:\n"
            "1. hwpmate_list_supported_formats로 PDF 지원 확인.\n"
            "2. hwpmate_preview_conversion으로 개수·출력 경로·충돌 미리보기 (plan_id 확보).\n"
            "3. 사용자에게 입력 파일·출력 폴더·건수를 보여주고 명시적 승인 획득.\n"
            "4. hwpmate_submit_conversion 제출 후 hwpmate_get_job_status로 추적.\n"
            "5. hwpmate_validate_artifacts로 산출물 검증 후 결과 보고."
        )

    @server.prompt("prepare_documents_for_office")
    def prompt_prepare_office() -> str:
        return (
            "오피스용 문서 준비: 필요한 출력 형식(DOCX/PDF 등)만 확인하고 "
            "hwpmate_preview_conversion으로 HwpMate 지원 범위 내 변환 계획을 제시한다. "
            "문서 내용 편집·작성은 하지 않는다."
        )

    @server.prompt("review_conversion_failures")
    def prompt_review_failures() -> str:
        return (
            "변환 실패 검토: hwpmate_get_job_result로 실패 파일과 사유를 요약하고 "
            "자동 무한 재시도 없이 다음 조치(원본 확인·형식 변경·수동 확인)를 제안한다."
        )

    return server


def run_stdio(cfg: McpConfig) -> None:
    if not cfg.enabled:
        raise SystemExit("MCP 서버가 비활성화되어 있습니다 (mcp.enabled=false).")
    if cfg.transport != "stdio":
        raise SystemExit(f"지원하지 않는 transport: {cfg.transport} (stdio만 지원)")
    server = create_server(cfg)
    logger.info("HwpMate MCP stdio 서버 시작")
    server.run(transport="stdio")
