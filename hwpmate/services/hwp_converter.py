from __future__ import annotations

import csv
import io
import subprocess
import time
from pathlib import Path
from typing import Any, Optional, Protocol, Tuple, cast

from ..constants import DOCUMENT_LOAD_DELAY, FORMAT_TYPES, HWP_PROGIDS
from ..logging_config import get_logger

logger = get_logger(__name__)

pythoncom: Optional[Any] = None
win32_client: Optional[Any] = None

try:
    import pythoncom as _pythoncom
    from win32com import client as _win32_client

    pythoncom = _pythoncom
    win32_client = _win32_client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

HWP_PROCESS_NAMES = {"hwp.exe", "hwpctrl.exe"}


class HwpAutomation(Protocol):
    """한글 COM 자동화 객체에서 사용하는 최소 인터페이스."""

    def RegisterModule(self, module_name: str, module_name_alias: str) -> Any: ...
    def SetMessageBoxMode(self, mode: int) -> Any: ...
    def Open(self, path: str, format_name: str, options: str) -> Any: ...
    def SaveAs(self, path: str, format_name: str, options: str = "") -> Any: ...
    def Clear(self, option: int = 0) -> Any: ...
    def Quit(self) -> Any: ...


def require_pywin32() -> Tuple[Any, Any]:
    """pywin32 모듈을 보장하고 반환."""
    if pythoncom is None or win32_client is None:
        raise RuntimeError("pywin32가 필요합니다. `pip install pywin32` 후 다시 실행하세요.")
    return pythoncom, win32_client


def _snapshot_hwp_pids() -> set[int]:
    """현재 실행 중인 한글 관련 프로세스 PID 집합 반환."""
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=False,
        )
        if result.returncode != 0:
            return set()

        reader = csv.reader(io.StringIO(result.stdout))
        pids: set[int] = set()
        for row in reader:
            if len(row) < 2:
                continue
            image_name = row[0].strip().lower()
            if image_name not in HWP_PROCESS_NAMES:
                continue
            try:
                pids.add(int(row[1]))
            except ValueError:
                continue
        return pids
    except Exception as e:
        logger.debug(f"한글 프로세스 스냅샷 수집 실패: {e}")
        return set()


def get_registered_hwp_progids() -> list[str]:
    """레지스트리에서 확인 가능한 한글 COM ProgID 목록을 반환."""
    try:
        import winreg
    except ImportError:
        return []

    registered: list[str] = []
    for progid in HWP_PROGIDS:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                registered.append(progid)
        except OSError:
            continue
    return registered


class HWPConverter:
    """한글 변환 엔진 - 기존 로직 완전 유지."""

    def __init__(self) -> None:
        self.hwp: Optional[HwpAutomation] = None
        self.progid_used: Optional[str] = None
        self.is_initialized = False
        self.owned_pids: set[int] = set()

    def initialize(self) -> bool:
        """COM 초기화 및 한글 객체 생성."""
        if self.is_initialized:
            return True

        pythoncom_module, win32_client_module = require_pywin32()

        try:
            pythoncom_module.CoInitialize()
        except Exception as e:
            logger.debug(f"CoInitialize 오류 (무시 가능): {e}")

        errors = []
        for progid in HWP_PROGIDS:
            before_pids = _snapshot_hwp_pids()
            try:
                self.hwp = cast(HwpAutomation, win32_client_module.Dispatch(progid))
                self.progid_used = progid
                hwp = self.hwp

                try:
                    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                except Exception:
                    pass

                hwp.SetMessageBoxMode(0x00000001)
                time.sleep(0.2)
                self.owned_pids = _snapshot_hwp_pids() - before_pids
                self.is_initialized = True
                logger.info(f"한글 연결 성공: {progid}")
                if self.owned_pids:
                    logger.info(f"앱 소유 한글 프로세스 추적: {sorted(self.owned_pids)}")
                else:
                    logger.info("새로 생성된 한글 프로세스를 추적하지 못했습니다. 강제 종료는 비활성화됩니다.")
                return True

            except Exception as e:
                errors.append(f"{progid}: {str(e)}")
                continue

        error_detail = "\n".join(errors)
        raise Exception(f"한글 COM 객체 생성 실패\n\n시도한 ProgID:\n{error_detail}")

    def convert_file(self, input_path, output_path, format_type="PDF") -> Tuple[bool, Optional[str]]:
        """단일 파일 변환."""
        hwp = self.hwp
        if not self.is_initialized or hwp is None:
            return False, "한글 객체가 초기화되지 않았습니다"

        try:
            input_str = str(input_path)
            output_str = str(output_path)

            open_result = hwp.Open(input_str, "", "forceopen:true")
            if open_result is False:
                return False, "문서 열기 실패: HWP Open이 False를 반환했습니다"
            time.sleep(DOCUMENT_LOAD_DELAY)

            format_info = FORMAT_TYPES.get(format_type, FORMAT_TYPES["PDF"])
            save_format = format_info["save_format"]

            save_error = None

            try:
                save_result = hwp.SaveAs(output_str, save_format)
                if save_result is False:
                    raise RuntimeError("SaveAs 2-param returned False")
                logger.debug(f"SaveAs 2-param 성공: {output_str}")
            except Exception as e1:
                logger.debug(f"SaveAs 2-param 실패: {e1}")

                try:
                    save_result = hwp.SaveAs(output_str, save_format, "")
                    if save_result is False:
                        raise RuntimeError("SaveAs 3-param returned False")
                    logger.debug(f"SaveAs 3-param 성공: {output_str}")
                except Exception as e2:
                    save_error = f"2-param: {e1}, 3-param: {e2}"
                    logger.error(f"모든 SaveAs 방식 실패: {save_error}")

                    try:
                        hwp.Clear(option=1)
                    except Exception:
                        pass
                    return False, save_error

            output_file = Path(output_str)
            if not output_file.exists():
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"출력 파일이 생성되지 않았습니다: {output_file.name}"
            try:
                if output_file.stat().st_size <= 0:
                    try:
                        hwp.Clear(option=1)
                    except Exception:
                        pass
                    return False, f"출력 파일이 비어 있습니다: {output_file.name}"
            except OSError as e:
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"출력 파일 확인 실패: {e}"

            hwp.Clear(option=1)

            return True, None

        except Exception as e:
            error_msg = str(e)
            logger.error(f"변환 실패 ({input_path}): {error_msg}")
            if hwp is not None:
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass

            return False, error_msg

    def has_owned_processes(self) -> bool:
        return bool(self.owned_pids)

    def kill_owned_processes(self) -> bool:
        """앱이 새로 띄운 한글 프로세스만 강제 종료."""
        if not self.owned_pids:
            logger.warning("추적된 한글 프로세스가 없어 강제 종료를 수행하지 않습니다.")
            return False

        killed_any = False
        for pid in sorted(self.owned_pids):
            try:
                result = subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
                if result.returncode == 0:
                    killed_any = True
                    logger.warning(f"앱 소유 한글 프로세스를 강제 종료했습니다: PID={pid}")
                else:
                    logger.debug(f"PID 종료 실패 또는 이미 종료됨: PID={pid}, code={result.returncode}")
            except Exception as e:
                logger.error(f"PID 강제 종료 실패: PID={pid}, 오류={e}")

        if killed_any:
            self.owned_pids.clear()
        return killed_any

    def cleanup(self) -> None:
        """정리."""
        hwp = self.hwp
        if hwp is not None and self.is_initialized:
            try:
                hwp.Clear(3)
            except Exception:
                pass

            try:
                hwp.Quit()
            except Exception:
                pass

            self.hwp = None
            self.is_initialized = False
            self.owned_pids.clear()

        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
