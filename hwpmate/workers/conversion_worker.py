from __future__ import annotations

import subprocess
from pathlib import Path
from typing import List

from PyQt6.QtCore import QThread, pyqtSignal

from ..logging_config import get_logger
from ..models import ConversionTask
from ..services.hwp_converter import HWPConverter, pythoncom

logger = get_logger(__name__)

class ConversionWorker(QThread):
    """변환 작업 워커 스레드"""
    
    # 시그널 정의
    progress_updated = pyqtSignal(int, int, str)  # current, total, filename
    status_updated = pyqtSignal(str)
    task_completed = pyqtSignal(int, int, list)  # success, total, failed_tasks
    error_occurred = pyqtSignal(str)
    
    # 스레드 내 COM 객체를 초기화하기 위한 플래그
    _com_initialized = False
    
    def __init__(self, tasks: List[ConversionTask], format_type: str):
        super().__init__()
        self.tasks = tasks
        self.format_type = format_type
        self.cancel_requested = False
    
    def cancel(self) -> None:
        """취소 요청"""
        self.cancel_requested = True
    
    def run(self) -> None:
        """변환 작업 수행"""
        # 별도 스레드에서 COM 초기화 필수
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
            except Exception as e:
                logger.debug(f"Worker COM 초기화: {e}")
        
        converter = HWPConverter()
        success_count = 0
        total = len(self.tasks)
        failed_tasks = []
        
        try:
            # 초기화
            self.status_updated.emit("한글 프로그램 연결 중...")
            converter.initialize()
            
            self.status_updated.emit(f"연결 성공: {converter.progid_used}")
            
            # 변환 실행
            for idx, task in enumerate(self.tasks):
                if self.cancel_requested:
                    self.status_updated.emit("사용자가 취소했습니다.")
                    break
                
                # 상태 업데이트
                self.progress_updated.emit(idx, total, task.input_file.name)
                
                # 0. 백업 수행 (안전장치)
                try:
                    self._create_backup(task.input_file)
                except Exception as e:
                    logger.warning(f"백업 실패 (계속 진행): {e}")
                    # 백업 실패해도 변환은 계속 진행 (선택사항)
                
                # 출력 폴더 생성
                try:
                    task.output_file.parent.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    task.status = "실패"
                    task.error = f"폴더 생성 실패: {e}"
                    failed_tasks.append(task)
                    continue
                
                # 입력 파일 존재 여부 확인
                if not task.input_file.exists():
                    task.status = "실패"
                    task.error = f"파일을 찾을 수 없음: {task.input_file.name}"
                    failed_tasks.append(task)
                    logger.warning(f"파일 없음: {task.input_file}")
                    continue
                
                # 변환 실행
                task.status = "진행중"
                success, error = converter.convert_file(
                    task.input_file,
                    task.output_file,
                    self.format_type
                )
                
                if success:
                    task.status = "성공"
                    success_count += 1
                else:
                    task.status = "실패"
                    task.error = error
                    failed_tasks.append(task)
            
            # 완료 (취소된 경우도 부분 결과 표시)
            self.progress_updated.emit(total, total, "완료" if not self.cancel_requested else "취소됨")
            
            # 결과 시그널 발생 (취소 시에도 부분 결과 표시)
            self.task_completed.emit(success_count, total, failed_tasks)
            
        except Exception as e:
            logger.exception("변환 중 오류 발생")
            self.error_occurred.emit(str(e))
        
        finally:
            try:
                converter.cleanup()
            except Exception as e:
                logger.error(f"정리 중 오류: {e}")
            
            # COM 해제
            if self._com_initialized:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

    def force_terminate(self) -> None:
        """한글 프로세스 강제 종료 (응답 없음 시)"""
        try:
            # HWP 프로세스 찾아서 종료 (taskkill 사용)
            # HwpCtrl.exe 또는 Hwp.exe
            subprocess.run(["taskkill", "/F", "/IM", "Hwp.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(["taskkill", "/F", "/IM", "HwpCtrl.exe"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            logger.warning("한글 프로세스를 강제로 종료했습니다.")
        except Exception as e:
            logger.error(f"프로세스 강제 종료 실패: {e}")


    def _create_backup(self, file_path: Path) -> None:
        """파일 백업 생성"""
        try:
            # backup 폴더 생성
            backup_dir = file_path.parent / "backup"
            backup_dir.mkdir(exist_ok=True)
            
            # 백업 파일명 생성 (원본이름_시간.확장자)
            # 안전을 위해 덮어쓰지 않고 항상 새 파일 생성
            import shutil
            from datetime import datetime
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
            backup_path = backup_dir / backup_name
            
            shutil.copy2(file_path, backup_path)
            logger.debug(f"백업 생성 완료: {backup_path}")
            
        except Exception as e:
            logger.error(f"백업 생성 중 오류: {e}")
            raise e
