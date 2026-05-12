from __future__ import annotations

import csv
import io
import subprocess
import time
from dataclasses import dataclass
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
AUXILIARY_ARTIFACT_FORMATS = {"HTML", "PNG", "JPG", "BMP", "GIF"}


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
    candidates: dict[str, Path] = {str(output_file): output_file}
    if format_type not in AUXILIARY_ARTIFACT_FORMATS:
        return list(candidates.values())

    parent = output_file.parent
    if not parent.exists():
        return list(candidates.values())

    stem_key = output_file.stem.lower()
    try:
        for child in parent.iterdir():
            if child.name.lower().startswith(stem_key):
                if child.is_file():
                    candidates[str(child)] = child
                elif child.is_dir():
                    try:
                        for nested in child.rglob("*"):
                            if nested.is_file():
                                candidates[str(nested)] = nested
                    except OSError:
                        continue
    except OSError:
        return list(candidates.values())

    return list(candidates.values())


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
        self.last_created_files: list[Path] = []
        self.last_output_size: int | None = None
        self.last_output_mtime: float | None = None
        self.last_save_format: str | None = None

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
                    self.security_module_registered = True
                    self.security_module_error = None
                except Exception as module_error:
                    self.security_module_registered = False
                    self.security_module_error = str(module_error)
                    logger.warning(f"한글 보안 모듈 등록 실패: {module_error}")

                hwp.SetMessageBoxMode(0x00000001)
                time.sleep(0.2)
                self.owned_pids = _snapshot_hwp_pids() - before_pids
                self.is_initialized = True
                logger.info(f"한글 연결 성공: {progid}")
                if self.owned_pids:
                    self.process_tracking_warning = None
                    logger.info(f"앱 소유 한글 프로세스 추적: {sorted(self.owned_pids)}")
                else:
                    self.process_tracking_warning = "새로 생성된 한글 프로세스를 추적하지 못했습니다. 강제 종료는 비활성화됩니다."
                    logger.info(self.process_tracking_warning)
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
            output_file = Path(output_str)
            self.last_created_files = []
            self.last_output_size = None
            self.last_output_mtime = None
            self.last_save_format = None

            open_result = hwp.Open(input_str, "", "forceopen:true")
            if open_result is False:
                return False, "문서 열기 실패: HWP Open이 False를 반환했습니다"
            time.sleep(DOCUMENT_LOAD_DELAY)

            format_info = FORMAT_TYPES.get(format_type, FORMAT_TYPES["PDF"])
            save_format = format_info["save_format"]
            self.last_save_format = save_format

            save_error = None
            before_artifacts = _snapshot_artifacts(output_file, format_type)

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
            self.process_tracking_warning = None

        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
