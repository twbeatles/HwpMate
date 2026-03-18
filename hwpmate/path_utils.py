from __future__ import annotations

import os
from pathlib import Path
from typing import Callable, Iterable, Optional

from .constants import SUPPORTED_EXTENSIONS
from .logging_config import get_logger

logger = get_logger(__name__)

def is_valid_path_name(path: str) -> bool:
    """Windows 파일 경로에 유효하지 않은 문자가 있는지 검증"""
    invalid_chars = '<>"|?*'
    # 드라이브 문자(:) 제외
    path_without_drive = path[2:] if len(path) > 2 and path[1] == ':' else path
    return not any(char in path_without_drive for char in invalid_chars)


def check_write_permission(folder_path: Path) -> bool:
    """폴더에 쓰기 권한이 있는지 확인"""
    try:
        test_file = folder_path / f".write_test_{os.getpid()}"
        test_file.touch()
        test_file.unlink()
        return True
    except (PermissionError, OSError):
        return False


def canonicalize_path(path: str) -> str:
    """경로를 비교/표시에 일관적인 절대경로 문자열로 정규화"""
    return os.path.abspath(os.path.normpath(str(path)))


def make_path_key(path: str) -> str:
    """Windows 대소문자 비민감 중복 체크용 키 생성"""
    return os.path.normcase(canonicalize_path(path))


def iter_supported_files(
    root_path: Path,
    include_sub: bool = True,
    allowed_exts: Optional[Iterable[str]] = None,
    cancel_checker: Optional[Callable[[], bool]] = None,
) -> Iterable[Path]:
    """단일 패스로 지원 확장자 파일을 순회"""
    allowed = {ext.lower() for ext in (allowed_exts or SUPPORTED_EXTENSIONS)}

    try:
        if root_path.is_file():
            if root_path.suffix.lower() in allowed:
                yield root_path
            return
    except OSError:
        return

    if not root_path.is_dir():
        return

    if include_sub:
        try:
            for dirpath, _, filenames in os.walk(root_path):
                if cancel_checker and cancel_checker():
                    return
                for filename in filenames:
                    if cancel_checker and cancel_checker():
                        return
                    _, ext = os.path.splitext(filename)
                    if ext.lower() in allowed:
                        yield Path(dirpath) / filename
        except OSError as e:
            logger.debug(f"하위 폴더 스캔 실패: {root_path} - {e}")
        return

    try:
        with os.scandir(root_path) as entries:
            for entry in entries:
                if cancel_checker and cancel_checker():
                    return
                if not entry.is_file():
                    continue
                _, ext = os.path.splitext(entry.name)
                if ext.lower() in allowed:
                    yield Path(entry.path)
    except OSError as e:
        logger.debug(f"폴더 스캔 실패: {root_path} - {e}")
