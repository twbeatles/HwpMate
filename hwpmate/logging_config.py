from __future__ import annotations

import logging
import os
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Iterable

LOG_FILE_NAME = "hwp_converter.log"


def _log_dir_candidates() -> list[Path]:
    candidates = [Path.home() / ".hwp_converter" / "logs"]

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        candidates.append(Path(local_app_data) / "HwpMate" / "logs")

    candidates.append(Path(tempfile.gettempdir()) / "HwpMate" / "logs")

    unique_candidates: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        unique_candidates.append(candidate)
    return unique_candidates


def _resolve_log_file(candidates: Iterable[Path] | None = None) -> tuple[Path | None, str | None]:
    errors: list[str] = []

    for log_dir in candidates or _log_dir_candidates():
        try:
            if log_dir.exists() and not log_dir.is_dir():
                raise NotADirectoryError(f"{log_dir} exists and is not a directory")

            log_dir.mkdir(parents=True, exist_ok=True)
            log_file = log_dir / LOG_FILE_NAME

            with log_file.open("a", encoding="utf-8"):
                pass

            return log_file, None
        except OSError as e:
            errors.append(f"{log_dir}: {e}")

    if not errors:
        errors.append("no writable log directory candidates")
    return None, " | ".join(errors)


LOG_FILE, LOG_INIT_ERROR = _resolve_log_file()
LOG_DIR = LOG_FILE.parent if LOG_FILE is not None else None


def configure_logging() -> logging.Logger:
    logger = logging.getLogger("hwpmate")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(stream_handler)
    logger.propagate = False

    if LOG_FILE is not None:
        try:
            file_handler = RotatingFileHandler(
                LOG_FILE,
                maxBytes=10 * 1024 * 1024,
                backupCount=5,
                encoding="utf-8",
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except OSError as e:
            logger.warning(f"파일 로그 핸들러를 비활성화합니다: {e}")
    elif LOG_INIT_ERROR is not None:
        logger.warning(f"파일 로그를 초기화하지 못했습니다: {LOG_INIT_ERROR}")

    return logger


logger = configure_logging()


def get_logger(name: str) -> logging.Logger:
    return logger.getChild(name)
