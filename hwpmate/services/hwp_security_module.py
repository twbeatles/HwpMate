from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Optional

from ..logging_config import get_logger

logger = get_logger(__name__)

SECURITY_MODULE_DLL_NAME = "FilePathCheckerModuleExample.dll"
SECURITY_MODULE_ALIAS = "FilePathCheckerModuleExample"
# Automation(HWPFrame) 과 Ctrl 경로 모두 등록 (한컴 버전·ProgID 별 조회 위치 차이)
REGISTRY_KEY_PATHS = (
    r"Software\HNC\HwpAutomation\Modules",
    r"Software\HNC\HwpCtrl\Modules",
    r"Software\Hnc\HwpAutomation\Modules",
    r"Software\Hnc\HwpCtrl\Modules",
)
# 하위 호환
REGISTRY_KEY_PATH = REGISTRY_KEY_PATHS[0]


def _runtime_security_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return Path(local) / "HwpMate" / "security"
    return Path.home() / ".hwp_converter" / "security"


def _bundled_dll_candidates() -> list[Path]:
    """개발 트리·PyInstaller onefile 모두에서 번들 DLL 후보 경로.

    onefile 배포 시 레지스트리에는 임시 _MEIPASS 경로를 넣지 않고,
    반드시 LOCALAPPDATA 로 복사한 경로를 등록한다 (한컴 프로세스가 안정적으로 로드).
    """
    candidates: list[Path] = []
    # PyInstaller onefile: 최우선 (datas 추출 위치)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        base = Path(meipass)
        candidates.append(base / "hwpmate" / "resources" / "security" / SECURITY_MODULE_DLL_NAME)
        candidates.append(base / "resources" / "security" / SECURITY_MODULE_DLL_NAME)
        candidates.append(base / SECURITY_MODULE_DLL_NAME)
    # 패키지 상대: hwpmate/resources/security/ (개발 실행)
    try:
        pkg_root = Path(__file__).resolve().parent.parent
        candidates.append(pkg_root / "resources" / "security" / SECURITY_MODULE_DLL_NAME)
    except OSError:
        pass
    # 실행 파일 옆 (onedir 또는 수동 동봉)
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "security" / SECURITY_MODULE_DLL_NAME)
        candidates.append(exe_dir / "hwpmate" / "resources" / "security" / SECURITY_MODULE_DLL_NAME)
        candidates.append(exe_dir / SECURITY_MODULE_DLL_NAME)
    return candidates


def find_bundled_security_dll() -> Optional[Path]:
    for path in _bundled_dll_candidates():
        try:
            if path.is_file() and path.stat().st_size > 0:
                return path
        except OSError:
            continue
    return None


def install_security_dll() -> Path:
    """번들 DLL 을 영구 런타임 디렉터리에 복사하고 경로를 반환.

    한컴 Hwp 프로세스가 로드하므로 _MEIPASS(임시)가 아닌 LOCALAPPDATA 경로를 사용한다.
    """
    dest_dir = _runtime_security_dir()
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / SECURITY_MODULE_DLL_NAME

    source = find_bundled_security_dll()
    if source is None:
        if dest.is_file() and dest.stat().st_size > 0:
            logger.info(f"번들 DLL 없음 — 기존 설치본 사용: {dest}")
            return dest
        raise FileNotFoundError(
            f"보안 모듈 DLL 을 찾을 수 없습니다: {SECURITY_MODULE_DLL_NAME}. "
            "배포 빌드에 hwpmate/resources/security/ 가 포함됐는지 확인하세요."
        )

    try:
        need_copy = (
            not dest.is_file()
            or dest.stat().st_size != source.stat().st_size
            or dest.stat().st_mtime < source.stat().st_mtime
        )
        if need_copy:
            # 실행 중 잠금 대비: 임시 파일 후 교체
            tmp = dest.with_suffix(".dll.tmp")
            shutil.copy2(source, tmp)
            tmp.replace(dest)
            logger.info(f"보안 모듈 DLL 설치: {source} -> {dest}")
    except OSError as e:
        if dest.is_file() and dest.stat().st_size > 0:
            logger.warning(f"보안 모듈 DLL 복사 실패, 기존 파일 사용: {e}")
            return dest
        raise
    return dest


def write_security_module_registry(dll_path: Path, alias: str = SECURITY_MODULE_ALIAS) -> str:
    """HKCU 보안 모듈 경로를 Automation/Ctrl 키에 모두 등록. 등록된 별칭 반환."""
    try:
        import winreg
    except ImportError as e:
        raise RuntimeError("winreg 를 사용할 수 없습니다") from e

    dll_abs = str(dll_path.resolve())
    # 따옴표 없이 경로만 (한컴 문서 요구)
    if dll_abs.startswith('"') and dll_abs.endswith('"'):
        dll_abs = dll_abs[1:-1]

    written = 0
    last_error: Exception | None = None
    for reg_path in REGISTRY_KEY_PATHS:
        try:
            key = winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER,
                reg_path,
                0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
            )
            try:
                winreg.SetValueEx(key, alias, 0, winreg.REG_SZ, dll_abs)
                value, reg_type = winreg.QueryValueEx(key, alias)
                if reg_type != winreg.REG_SZ or str(value).strip().strip('"') != dll_abs:
                    raise RuntimeError(f"레지스트리 검증 실패: {reg_path}\\{alias}={value!r}")
            finally:
                winreg.CloseKey(key)
            written += 1
            logger.info(f"보안 모듈 레지스트리 등록: {reg_path}\\{alias} = {dll_abs}")
        except Exception as e:
            last_error = e
            logger.debug(f"레지스트리 등록 스킵/실패 {reg_path}: {e}")

    if written == 0:
        raise RuntimeError(f"보안 모듈 레지스트리 등록 실패: {last_error}")
    return alias


def read_security_module_registry(alias: str = SECURITY_MODULE_ALIAS) -> Optional[str]:
    try:
        import winreg
    except ImportError:
        return None
    for reg_path in REGISTRY_KEY_PATHS:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_QUERY_VALUE)
            try:
                value, _ = winreg.QueryValueEx(key, alias)
                return str(value).strip().strip('"')
            finally:
                winreg.CloseKey(key)
        except OSError:
            continue
    return None


def ensure_hwp_security_module(alias: str = SECURITY_MODULE_ALIAS) -> tuple[bool, str, Optional[str]]:
    """
    보안 모듈 DLL 설치 + 레지스트리 등록을 보장한다.

    Returns:
        (ok, message, alias_or_none)
    """
    try:
        dll_path = install_security_dll()
        if not dll_path.is_file():
            return False, f"DLL 파일이 없습니다: {dll_path}", None
        registered_alias = write_security_module_registry(dll_path, alias=alias)
        verified = read_security_module_registry(registered_alias)
        if not verified or not Path(verified).is_file():
            return (
                False,
                f"레지스트리 경로 검증 실패: alias={registered_alias}, path={verified!r}",
                None,
            )
        return True, f"DLL={verified}", registered_alias
    except Exception as e:
        logger.exception("보안 모듈 준비 실패")
        return False, str(e), None
