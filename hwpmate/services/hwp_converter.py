from __future__ import annotations

import time
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


class HWPConverter:
    """한글 변환 엔진 - 기존 로직 완전 유지."""

    def __init__(self) -> None:
        self.hwp: Optional[HwpAutomation] = None
        self.progid_used: Optional[str] = None
        self.is_initialized = False

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
            try:
                self.hwp = cast(HwpAutomation, win32_client_module.Dispatch(progid))
                self.progid_used = progid
                hwp = self.hwp

                try:
                    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                except Exception:
                    pass

                hwp.SetMessageBoxMode(0x00000001)
                self.is_initialized = True
                logger.info(f"한글 연결 성공: {progid}")
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

            hwp.Open(input_str, "", "forceopen:true")
            time.sleep(DOCUMENT_LOAD_DELAY)

            format_info = FORMAT_TYPES.get(format_type, FORMAT_TYPES["PDF"])
            save_format = format_info["save_format"]

            save_error = None

            try:
                hwp.SaveAs(output_str, save_format)
                logger.debug(f"SaveAs 2-param 성공: {output_str}")
            except Exception as e1:
                logger.debug(f"SaveAs 2-param 실패: {e1}")

                try:
                    hwp.SaveAs(output_str, save_format, "")
                    logger.debug(f"SaveAs 3-param 성공: {output_str}")
                except Exception as e2:
                    save_error = f"2-param: {e1}, 3-param: {e2}"
                    logger.error(f"모든 SaveAs 방식 실패: {save_error}")

                    try:
                        hwp.Clear(option=1)
                    except Exception:
                        pass
                    return False, save_error

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

        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
