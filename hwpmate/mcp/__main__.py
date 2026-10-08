"""MCP 서버 진입점: python -m hwpmate.mcp."""

from __future__ import annotations

import argparse
import sys

from . import MCP_SCHEMA_VERSION
from .config import McpConfig, load_config
from .logging_utils import setup_stderr_logging
from .server import run_stdio


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m hwpmate.mcp",
        description="HwpMate 로컬 MCP 서버 (stdio 전용, 읽기 전용 미리보기 + 승인 기반 변환 실행).",
    )
    parser.add_argument("--config", default="", help="MCP config.toml 경로 (기본: %%LOCALAPPDATA%%/HwpMate/mcp/config.toml).")
    parser.add_argument("--version", action="store_true", help="MCP 스키마 버전을 출력하고 종료합니다.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(sys.argv[1:] if argv is None else argv)
    if args.version:
        sys.stdout.write(MCP_SCHEMA_VERSION + "\n")
        return 0
    cfg: McpConfig = load_config(args.config or None)
    setup_stderr_logging()
    try:
        run_stdio(cfg)
    except SystemExit as exc:
        return int(exc.code or 0)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
