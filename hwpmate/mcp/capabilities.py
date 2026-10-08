"""환경 능력 조회 (구현 설계서 §4.1).

read-only이며, 한컴 COM 초기화·잠금 획득·파일 생성을 하지 않는다.
hancom_com_probe는 항상 "not_run"이며, --smoke 결과로 available을
선언하지 않는다. instance_lock은 "unknown"으로 보고한다.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

from ..constants import FORMAT_TYPES, VERSION
from . import MCP_SCHEMA_VERSION
from .config import McpConfig


def _cli_status(cfg: McpConfig) -> tuple[bool, str]:
    raw = str(cfg.cli_executable or "").strip()
    if not raw:
        return False, ""
    try:
        candidate = Path(raw).expanduser()
        if candidate.is_file():
            return True, str(candidate.resolve())
        return False, str(candidate)
    except OSError:
        return False, raw


def gather_capabilities(cfg: McpConfig) -> dict:
    cli_valid, cli_path = _cli_status(cfg)
    warnings: list[str] = []
    if not cli_valid:
        warnings.append("CLI 실행 파일이 설정되지 않았거나 존재하지 않아 변환 실행이 불가합니다.")
    warnings.append("실제 COM 변환 가능 여부는 현장 테스트 필요")
    return {
        "schema_version": MCP_SCHEMA_VERSION,
        "ok": True,
        "platform": sys.platform,
        "app_version": VERSION,
        "cli_path": cli_path,
        "cli_path_valid": cli_valid,
        "pywin32_available": importlib.util.find_spec("win32com") is not None,
        "hancom_com_probe": "not_run",
        "supported_formats": sorted(FORMAT_TYPES.keys()),
        "instance_lock": "unknown",
        "warnings": warnings,
    }


def list_supported_formats() -> dict:
    formats = [
        {"name": name, "ext": spec["ext"], "desc": spec["desc"]}
        for name, spec in FORMAT_TYPES.items()
    ]
    return {
        "schema_version": MCP_SCHEMA_VERSION,
        "ok": True,
        "formats": formats,
        "notes": [
            "ODT는 한글 버전에 따라 ODF 형식 문자열로 저장됩니다.",
            "이미지 형식(PNG/JPG/BMP/GIF)은 페이지별 다중 파일로 저장됩니다.",
        ],
    }
