from __future__ import annotations

import ctypes
import subprocess
import time
from ctypes import wintypes
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Protocol, Tuple, cast

from ..constants import DOCUMENT_LOAD_DELAY, FORMAT_TYPES, HWP_PROGIDS
from ..logging_config import get_logger
from .artifact_policy import iter_candidate_artifact_paths
from .hwp_security_module import (
    SECURITY_MODULE_ALIAS,
    ensure_hwp_security_module,
)

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

# 한컴 보안 모듈 RegisterModule 두 번째 인자 후보
# (ensure_hwp_security_module 이 등록하는 별칭을 최우선)
SECURITY_MODULE_ALIASES = (
    SECURITY_MODULE_ALIAS,
    "FilePathCheckerModule",
    "SecurityModule",
)

# Windows: 콘솔 창 없이 자식 프로세스 실행
_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

TH32CS_SNAPPROCESS = 0x00000002


def is_com_failure_result(result: object) -> bool:
    """COM Open/SaveAs 실패 반환값 정규화.

    일부 환경은 False, 일부는 0 을 실패로 돌린다. identity 비교만 쓰면 0 을 놓친다.
    """
    if result is False:
        return True
    if result == 0 and not isinstance(result, bool):
        return True
    return False


class _PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wintypes.DWORD),
        ("cntUsage", wintypes.DWORD),
        ("th32ProcessID", wintypes.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wintypes.DWORD),
        ("cntThreads", wintypes.DWORD),
        ("th32ParentProcessID", wintypes.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wintypes.DWORD),
        ("szExeFile", wintypes.WCHAR * 260),
    ]


@dataclass(frozen=True)
class _FileSnapshot:
    size: int
    mtime_ns: int
    ctime_ns: int


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


# Toolhelp 스냅샷 연속 실패 추적 (모듈 전역, 프로세스 수명 동안)
_snapshot_failure_count = 0
_snapshot_last_error: str | None = None


def get_snapshot_health() -> tuple[int, str | None]:
    """(연속 실패 횟수, 마지막 오류 메시지) — UI/워커 경고용."""
    return _snapshot_failure_count, _snapshot_last_error


def _snapshot_hwp_pids() -> set[int]:
    """현재 실행 중인 한글 관련 프로세스 PID 집합 반환.

    tasklist 서브프로세스 대신 Toolhelp32 를 사용해 콘솔 창 플래시를 원천 차단한다.
    """
    global _snapshot_failure_count, _snapshot_last_error
    try:
        kernel32 = ctypes.windll.kernel32
        snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snapshot in (-1, 0xFFFFFFFF):
            _snapshot_failure_count += 1
            _snapshot_last_error = "CreateToolhelp32Snapshot invalid handle"
            return set()

        pids: set[int] = set()
        try:
            entry = _PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(_PROCESSENTRY32W)
            if not kernel32.Process32FirstW(snapshot, ctypes.byref(entry)):
                _snapshot_failure_count += 1
                _snapshot_last_error = "Process32FirstW failed"
                return set()
            while True:
                image_name = entry.szExeFile.strip().lower()
                if image_name in HWP_PROCESS_NAMES:
                    pids.add(int(entry.th32ProcessID))
                if not kernel32.Process32NextW(snapshot, ctypes.byref(entry)):
                    break
            # 성공 시 연속 실패 카운터 리셋
            _snapshot_failure_count = 0
            _snapshot_last_error = None
            return pids
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception as e:
        _snapshot_failure_count += 1
        _snapshot_last_error = str(e)
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


def _snapshot_file(path: Path) -> _FileSnapshot | None:
    try:
        stat = path.stat()
        if not path.is_file():
            return None
        return _FileSnapshot(
            size=stat.st_size,
            mtime_ns=stat.st_mtime_ns,
            ctime_ns=stat.st_ctime_ns,
        )
    except OSError:
        return None


def _iter_candidate_artifact_files(output_file: Path, format_type: str) -> list[Path]:
    return iter_candidate_artifact_paths(output_file, format_type)


def _snapshot_artifacts(output_file: Path, format_type: str) -> dict[Path, _FileSnapshot]:
    snapshots: dict[Path, _FileSnapshot] = {}
    for path in _iter_candidate_artifact_files(output_file, format_type):
        snapshot = _snapshot_file(path)
        if snapshot is not None:
            snapshots[path] = snapshot
    return snapshots


def _changed_artifacts(
    before: dict[Path, _FileSnapshot],
    after: dict[Path, _FileSnapshot],
) -> list[Path]:
    changed: list[Path] = []
    for path, snapshot in after.items():
        if snapshot.size <= 0:
            continue
        if before.get(path) != snapshot:
            changed.append(path)
    return sorted(changed, key=lambda p: str(p).lower())


class HWPConverter:
    """한글 변환 엔진 - 기존 로직 완전 유지."""

    def __init__(self) -> None:
        self.hwp: Optional[HwpAutomation] = None
        self.progid_used: Optional[str] = None
        self.is_initialized = False
        self.owned_pids: set[int] = set()
        self.security_module_registered: bool | None = None
        self.security_module_error: str | None = None
        self.process_tracking_warning: str | None = None
        self.snapshot_unreliable: bool = False
        self.last_created_files: list[Path] = []
        self.last_output_size: int | None = None
        self.last_output_mtime: float | None = None
        self.last_save_format: str | None = None
        # True only when this instance successfully called CoInitialize itself.
        self._com_apartment_owned = False

    def initialize(self, *, manage_com_apartment: bool = True) -> bool:
        """COM 초기화 및 한글 객체 생성.

        manage_com_apartment=False 이면 호출 스레드가 이미 CoInitialize 한 경우
        (ConversionWorker 등) 중복 초기화/해제를 하지 않는다.
        """
        if self.is_initialized:
            return True

        pythoncom_module, win32_client_module = require_pywin32()

        if manage_com_apartment:
            try:
                pythoncom_module.CoInitialize()
                self._com_apartment_owned = True
            except Exception as e:
                # 이미 동일 스레드에서 초기화된 경우 등은 무시하고 소유하지 않는다.
                self._com_apartment_owned = False
                logger.debug(f"CoInitialize 오류 (무시 가능): {e}")
        else:
            self._com_apartment_owned = False

        # DLL 설치 + HKCU\...\HwpAutomation\Modules 레지스트리 (RegisterModule 사전 조건)
        prep_ok, prep_msg, prep_alias = ensure_hwp_security_module()
        if prep_ok:
            logger.info(f"보안 모듈 사전 준비 완료: {prep_msg}")
        else:
            logger.warning(f"보안 모듈 사전 준비 실패: {prep_msg}")

        errors = []
        for progid in HWP_PROGIDS:
            before_pids = _snapshot_hwp_pids()
            try:
                self.hwp = cast(HwpAutomation, win32_client_module.Dispatch(progid))
                self.progid_used = progid
                hwp = self.hwp

                module_errors: list[str] = []
                self.security_module_registered = False
                self.security_module_error = None
                aliases: list[str] = []
                if prep_alias:
                    aliases.append(prep_alias)
                for name in SECURITY_MODULE_ALIASES:
                    if name not in aliases:
                        aliases.append(name)

                for alias in aliases:
                    try:
                        result = hwp.RegisterModule("FilePathCheckDLL", alias)
                        if result is False or result == 0:
                            module_errors.append(f"{alias}: RegisterModule returned {result!r}")
                            continue
                        # 레지스트리+DLL 사전 준비가 된 경우에만 "완전 성공" (팝업 억제 기대)
                        if prep_ok:
                            self.security_module_registered = True
                            self.security_module_error = None
                            logger.info(
                                f"한글 보안 모듈 등록 성공: alias={alias}, result={result!r}"
                            )
                        else:
                            self.security_module_registered = False
                            self.security_module_error = (
                                f"RegisterModule({alias}) 호출됨(result={result!r})이나 "
                                f"레지스트리/DLL 미준비: {prep_msg}"
                            )
                            logger.warning(self.security_module_error)
                        break
                    except Exception as module_error:
                        module_errors.append(f"{alias}: {module_error}")

                if not self.security_module_registered and self.security_module_error is None:
                    self.security_module_error = (
                        f"prep={prep_msg}; " + ("; ".join(module_errors) or "알 수 없는 오류")
                    )
                    logger.warning(
                        "한글 보안 모듈 등록 실패 (파일 접근 시 '모두 허용' 창이 뜰 수 있음): "
                        f"{self.security_module_error}"
                    )

                hwp.SetMessageBoxMode(0x00000001)
                time.sleep(0.2)
                after_pids = _snapshot_hwp_pids()
                fail_count, fail_msg = get_snapshot_health()
                self.snapshot_unreliable = fail_count > 0 and not after_pids and not before_pids
                self.owned_pids = after_pids - before_pids
                self.is_initialized = True
                logger.info(f"한글 연결 성공: {progid}")
                if self.snapshot_unreliable:
                    detail = f" ({fail_msg})" if fail_msg else ""
                    self.process_tracking_warning = (
                        "한글 프로세스 스냅샷(Toolhelp) 수집에 실패했습니다"
                        f"{detail}. 강제 종료와 창 전면화 범위가 제한될 수 있습니다."
                    )
                    logger.warning(self.process_tracking_warning)
                elif self.owned_pids:
                    self.process_tracking_warning = None
                    logger.info(f"앱 소유 한글 프로세스 추적: {sorted(self.owned_pids)}")
                else:
                    self.process_tracking_warning = (
                        "새로 생성된 한글 프로세스를 추적하지 못했습니다. "
                        "강제 종료는 비활성화됩니다. 변환 전 다른 한글 창을 닫으면 추적이 안정됩니다."
                    )
                    logger.info(self.process_tracking_warning)
                # 메인 편집 창 숨김 + 보안/허용 팝업만 전면화 (실패해도 연결 유지)
                self._suppress_hwp_ui_flash()
                return True

            except Exception as e:
                errors.append(f"{progid}: {str(e)}")
                continue

        error_detail = "\n".join(errors)
        if self._com_apartment_owned and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_apartment_owned = False
        raise Exception(f"한글 COM 객체 생성 실패\n\n시도한 ProgID:\n{error_detail}")

    def _try_set_xhw_windows_visible(self, visible: bool) -> bool:
        """XHwpWindows.Item(i).Visible 로 메인 창 표시 여부 설정 (버전 미지원 시 False)."""
        hwp = self.hwp
        if hwp is None:
            return False
        try:
            xwindows = getattr(hwp, "XHwpWindows", None)
            if xwindows is None:
                return False
            count_raw = getattr(xwindows, "Count", None)
            try:
                count = int(count_raw) if count_raw is not None else 1
            except (TypeError, ValueError):
                count = 1
            if count < 1:
                count = 1
            any_ok = False
            for index in range(count):
                try:
                    window = xwindows.Item(index)
                    window.Visible = visible
                    any_ok = True
                except Exception:
                    if index == 0:
                        return False
                    break
            return any_ok
        except Exception as e:
            logger.debug(f"XHwpWindows.Visible={visible} 실패(무시): {e}")
            return False

    def _suppress_hwp_ui_flash(self) -> None:
        """메인 편집 창 숨김 + 보안 대화상자 전면화 (best-effort, 변환 실패로 전파하지 않음)."""
        try:
            self._try_set_xhw_windows_visible(False)
        except Exception as e:
            logger.debug(f"COM Visible=False 실패(무시): {e}")
        try:
            from ..windows_integration import suppress_hwp_ui_flash

            hidden, raised = suppress_hwp_ui_flash(self.owned_pids or None)
            if hidden or raised:
                logger.debug(
                    f"한글 UI 억제: hidden={hidden}, security_raised={raised}, "
                    f"pids={sorted(self.owned_pids) if self.owned_pids else None}"
                )
        except Exception as e:
            logger.debug(f"한글 UI 억제 실패(무시): {e}")

    def convert_file(self, input_path, output_path, format_type="PDF") -> Tuple[bool, Optional[str]]:
        """단일 파일 변환."""
        hwp = self.hwp
        if not self.is_initialized or hwp is None:
            return False, "한글 객체가 초기화되지 않았습니다"

        try:
            input_str = str(input_path)
            output_str = str(output_path)
            output_file = Path(output_str)
            self.last_created_files = []
            self.last_output_size = None
            self.last_output_mtime = None
            self.last_save_format = None

            # 일부 환경에서 Open 직전 재등록이 파일 경로 승인 훅을 안정화함
            try:
                alias = SECURITY_MODULE_ALIAS
                hwp.RegisterModule("FilePathCheckDLL", alias)
            except Exception as re_reg_error:
                logger.debug(f"Open 전 RegisterModule 재호출 실패(무시): {re_reg_error}")

            open_result = hwp.Open(input_str, "", "forceopen:true")
            if is_com_failure_result(open_result):
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"문서 열기 실패: HWP Open이 실패를 반환했습니다 ({open_result!r})"
            time.sleep(DOCUMENT_LOAD_DELAY)
            # Open 후 메인 창이 다시 뜰 수 있어 best-effort 재숨김
            self._suppress_hwp_ui_flash()

            format_info = FORMAT_TYPES.get(format_type, FORMAT_TYPES["PDF"])
            save_format = format_info["save_format"]
            self.last_save_format = save_format

            save_error = None
            before_artifacts = _snapshot_artifacts(output_file, format_type)

            try:
                save_result = hwp.SaveAs(output_str, save_format)
                if is_com_failure_result(save_result):
                    raise RuntimeError(f"SaveAs 2-param returned failure: {save_result!r}")
                logger.debug(f"SaveAs 2-param 성공: {output_str}")
            except Exception as e1:
                logger.debug(f"SaveAs 2-param 실패: {e1}")

                try:
                    save_result = hwp.SaveAs(output_str, save_format, "")
                    if is_com_failure_result(save_result):
                        raise RuntimeError(f"SaveAs 3-param returned failure: {save_result!r}")
                    logger.debug(f"SaveAs 3-param 성공: {output_str}")
                except Exception as e2:
                    save_error = f"2-param: {e1}, 3-param: {e2}"
                    logger.error(f"모든 SaveAs 방식 실패: {save_error}")

                    try:
                        hwp.Clear(option=1)
                    except Exception:
                        pass
                    return False, save_error

            after_artifacts = _snapshot_artifacts(output_file, format_type)
            primary_snapshot = after_artifacts.get(output_file)

            if not after_artifacts:
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"출력 파일이 생성되지 않았습니다: {output_file.name}"

            if primary_snapshot is not None and primary_snapshot.size <= 0:
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"출력 파일이 비어 있습니다: {output_file.name}"

            changed_files = _changed_artifacts(before_artifacts, after_artifacts)
            if not changed_files:
                try:
                    hwp.Clear(option=1)
                except Exception:
                    pass
                return False, f"출력 파일이 새로 생성되거나 갱신되지 않았습니다: {output_file.name}"

            representative = output_file if output_file in changed_files else changed_files[0]
            representative_snapshot = after_artifacts[representative]
            self.last_created_files = changed_files
            self.last_output_size = representative_snapshot.size
            try:
                self.last_output_mtime = representative.stat().st_mtime
            except OSError:
                self.last_output_mtime = representative_snapshot.mtime_ns / 1_000_000_000

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

        # PID 재사용 방지를 위해 종료 직전 한글 이미지명 프로세스인지 재확인한다.
        live_hwp_pids = _snapshot_hwp_pids()
        killed_any = False
        remaining: set[int] = set()
        for pid in sorted(self.owned_pids):
            if pid not in live_hwp_pids:
                logger.warning(
                    f"PID={pid} 가 현재 한글 관련 프로세스가 아니거나 이미 종료되어 강제 종료를 건너뜁니다."
                )
                continue
            try:
                result = subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    creationflags=_CREATE_NO_WINDOW,
                )
                if result.returncode == 0:
                    killed_any = True
                    logger.warning(f"앱 소유 한글 프로세스를 강제 종료했습니다: PID={pid}")
                else:
                    remaining.add(pid)
                    logger.debug(f"PID 종료 실패 또는 이미 종료됨: PID={pid}, code={result.returncode}")
            except Exception as e:
                remaining.add(pid)
                logger.error(f"PID 강제 종료 실패: PID={pid}, 오류={e}")

        self.owned_pids = remaining
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
            self.process_tracking_warning = None

        if self._com_apartment_owned and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_apartment_owned = False
