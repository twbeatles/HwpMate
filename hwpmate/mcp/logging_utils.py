"""MCP 로그 유틸: stdout 오염 금지 — 로그는 stderr/파일로만 기록한다."""

from __future__ import annotations

import logging
import sys


def setup_stderr_logging(level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger("hwpmate.mcp")
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(name)s: %(message)s"))
        logger.addHandler(handler)
    logger.setLevel(level)
    logger.propagate = False
    return logger
