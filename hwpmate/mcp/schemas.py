"""공유 스키마 상수와 성공 envelope (구현 설계서 §4)."""

from __future__ import annotations

from . import MCP_SCHEMA_VERSION

CONFIRMATION_TOKEN = "approve_non_destructive_conversion"

__all__ = ["CONFIRMATION_TOKEN", "MCP_SCHEMA_VERSION", "ok_envelope"]


def ok_envelope(**fields) -> dict:
    payload = {"schema_version": MCP_SCHEMA_VERSION, "ok": True}
    payload.update(fields)
    return payload
