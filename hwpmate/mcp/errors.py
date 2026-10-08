"""MCP 에러 코드와 예외 정의 (구현 설계서 §6.3)."""

from __future__ import annotations


class McpError(Exception):
    """MCP 도구 실패를 구조화 오류로 전달하기 위한 예외."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
        self.message = message


def error_envelope(code: str, message: str, *, schema_version: str) -> dict:
    return {
        "schema_version": schema_version,
        "ok": False,
        "error": {"code": code, "message": message},
    }
