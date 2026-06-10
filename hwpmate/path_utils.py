from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Callable, Iterable, Optional

from .constants import BACKUP_DIR_NAME, SUPPORTED_EXTENSIONS
from .logging_config import get_logger

logger = get_logger(__name__)

def is_valid_path_name(path: str) -> bool:
    """Windows 파일 경로에 유효하지 않은 문자가 있는지 검증"""
    path = str(path).strip()
    if not path:
        return False
    if any(ord(char) < 32 for char in path):
        return False

    normalized = path.replace("/", "\\")
    extended_unc_prefix = "\\\\?\\UNC\\"
    extended_prefix = "\\\\?\\"
    if normalized.startswith(extended_unc_prefix):
        normalized = "\\\\" + normalized[len(extended_unc_prefix):]
    elif normalized.startswith(extended_prefix):
        normalized = normalized[len(extended_prefix):]

    invalid_chars = '<>"|?*'
    if any(char in normalized for char in invalid_chars):
        return False

    path_without_drive = normalized
    if len(normalized) >= 2 and normalized[1] == ":":
        if not normalized[0].isalpha():
            return False
        path_without_drive = normalized[2:]
    if ":" in path_without_drive:
        return False

    reserved_names = {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    }
    for part in re.split(r"[\\]+", path_without_drive):
        if not part or part in {".", ".."}:
            continue
        if part.endswith((" ", ".")):
            return False
        base = part.split(".")[0].upper()
        if base in reserved_names:
            return False
    return True


def check_write_permission(folder_path: Path) -> bool:
    """폴더에 쓰기 권한이 있는지 확인"""
    try:
        with tempfile.NamedTemporaryFile(
            dir=folder_path,
            prefix=".hwpmate_write_test_",
            delete=True,
        ):
            pass
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
    excluded_dir_names: Optional[Iterable[str]] = None,
) -> Iterable[Path]:
    """단일 패스로 지원 확장자 파일을 순회"""
    allowed = {ext.lower() for ext in (allowed_exts or SUPPORTED_EXTENSIONS)}
    excluded_dirs = {name.lower() for name in (excluded_dir_names or (BACKUP_DIR_NAME,))}

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
            for dirpath, dirnames, filenames in os.walk(root_path):
                if cancel_checker and cancel_checker():
                    return
                dirnames[:] = [
                    dirname
                    for dirname in dirnames
                    if dirname.lower() not in excluded_dirs
                ]
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
