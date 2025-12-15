"""
HWP/HWPX 변환기 v7.0 - PyQt6 현대화 버전
안정성과 사용성에 초점을 맞춘 현대적 GUI 버전
"""

import sys
import json
import ctypes
import logging
import time
from pathlib import Path
from typing import Optional, List, Tuple

# PyQt6 imports
try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QGroupBox, QRadioButton, QCheckBox, QPushButton, QLabel,
        QLineEdit, QFileDialog, QProgressBar, QTableWidget, QTableWidgetItem,
        QHeaderView, QMessageBox, QDialog, QTextEdit, QFrame, QSplitter,
        QSystemTrayIcon, QMenu, QButtonGroup, QScrollArea, QSizePolicy,
        QStyle, QStyleFactory, QComboBox
    )
    from PyQt6.QtCore import (
        Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve,
        QTimer, QSize, QMimeData, QUrl
    )
    from PyQt6.QtGui import (
        QFont, QIcon, QPalette, QColor, QDragEnterEvent, QDropEvent,
        QAction, QPixmap, QPainter, QBrush, QPen
    )
    PYQT6_AVAILABLE = True
except ImportError:
    PYQT6_AVAILABLE = False
    print("오류: PyQt6 라이브러리가 필요합니다.\n\npip install PyQt6")
    sys.exit(1)

# pywin32 import (COM 사용)
try:
    import pythoncom
    import win32com.client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger(__name__)

# 설정 파일
CONFIG_FILE = Path.home() / ".hwp_converter_config.json"

# 한글 ProgID 목록 (우선순위 순)
HWP_PROGIDS = [
    "HWPControl.HwpCtrl.1",
    "HwpObject.HwpObject",
    "HWPFrame.HwpObject",
]


# ============================================================================
# 테마 시스템
# ============================================================================

class ThemeManager:
    """테마 관리자"""
    
    DARK_THEME = """
        /* 메인 윈도우 */
        QMainWindow, QWidget {
            background-color: #1a1a2e;
            color: #eaeaea;
            font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
            font-size: 10pt;
        }
        
        /* 그룹박스 */
        QGroupBox {
            background-color: #16213e;
            border: 1px solid #0f3460;
            border-radius: 10px;
            margin-top: 15px;
            padding: 15px;
            font-weight: bold;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            left: 15px;
            padding: 0 10px;
            color: #e94560;
        }
        
        /* 버튼 */
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #e94560, stop:1 #c73e54);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: bold;
            min-height: 20px;
        }
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #ff5a75, stop:1 #e94560);
        }
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #c73e54, stop:1 #a83245);
        }
        QPushButton:disabled {
            background: #3a3a5c;
            color: #666680;
        }
        
        /* 보조 버튼 */
        QPushButton[secondary="true"] {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #0f3460, stop:1 #0a2540);
            border: 1px solid #1a4a80;
        }
        QPushButton[secondary="true"]:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #1a4a80, stop:1 #0f3460);
        }
        
        /* 입력 필드 */
        QLineEdit {
            background-color: #0f3460;
            border: 2px solid #1a4a80;
            border-radius: 8px;
            padding: 10px;
            color: #eaeaea;
            selection-background-color: #e94560;
        }
        QLineEdit:focus {
            border-color: #e94560;
        }
        QLineEdit:disabled {
            background-color: #252545;
            color: #666680;
        }
        
        /* 라디오 버튼 & 체크박스 */
        QRadioButton, QCheckBox {
            spacing: 10px;
            padding: 5px;
        }
        QRadioButton::indicator, QCheckBox::indicator {
            width: 20px;
            height: 20px;
        }
        QRadioButton::indicator:unchecked, QCheckBox::indicator:unchecked {
            background-color: #0f3460;
            border: 2px solid #1a4a80;
            border-radius: 10px;
        }
        QCheckBox::indicator:unchecked {
            border-radius: 5px;
        }
        QRadioButton::indicator:checked, QCheckBox::indicator:checked {
            background-color: #e94560;
            border: 2px solid #e94560;
            border-radius: 10px;
        }
        QCheckBox::indicator:checked {
            border-radius: 5px;
        }
        
        /* 테이블 */
        QTableWidget {
            background-color: #16213e;
            alternate-background-color: #1a2744;
            border: 1px solid #0f3460;
            border-radius: 8px;
            gridline-color: #0f3460;
        }
        QTableWidget::item {
            padding: 8px;
            border-bottom: 1px solid #0f3460;
        }
        QTableWidget::item:selected {
            background-color: #e94560;
            color: white;
        }
        QHeaderView::section {
            background-color: #0f3460;
            color: #eaeaea;
            padding: 10px;
            border: none;
            font-weight: bold;
        }
        
        /* 진행률 바 */
        QProgressBar {
            background-color: #0f3460;
            border: none;
            border-radius: 10px;
            height: 20px;
            text-align: center;
            color: white;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #e94560, stop:1 #ff7b95);
            border-radius: 10px;
        }
        
        /* 스크롤바 */
        QScrollBar:vertical {
            background-color: #16213e;
            width: 12px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background-color: #0f3460;
            border-radius: 6px;
            min-height: 30px;
        }
        QScrollBar::handle:vertical:hover {
            background-color: #e94560;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        /* 드롭 영역 */
        QFrame[dropZone="true"] {
            background-color: #0f3460;
            border: 2px dashed #1a4a80;
            border-radius: 15px;
        }
        QFrame[dropZone="true"]:hover {
            border-color: #e94560;
            background-color: #162850;
        }
        
        /* 콤보박스 */
        QComboBox {
            background-color: #0f3460;
            border: 2px solid #1a4a80;
            border-radius: 8px;
            padding: 8px 15px;
            color: #eaeaea;
        }
        QComboBox:hover {
            border-color: #e94560;
        }
        QComboBox::drop-down {
            border: none;
            padding-right: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #16213e;
            border: 1px solid #0f3460;
            selection-background-color: #e94560;
        }
        
        /* 레이블 */
        QLabel[heading="true"] {
            font-size: 14pt;
            font-weight: bold;
            color: #e94560;
        }
        QLabel[subheading="true"] {
            font-size: 9pt;
            color: #888899;
        }
    """
    
    LIGHT_THEME = """
        /* 메인 윈도우 */
        QMainWindow, QWidget {
            background-color: #f8f9fa;
            color: #2d3436;
            font-family: 'Malgun Gothic', 'Segoe UI', sans-serif;
            font-size: 10pt;
        }
        
        /* 그룹박스 */
        QGroupBox {
            background-color: #ffffff;
            border: 1px solid #dfe6e9;
            border-radius: 10px;
            margin-top: 15px;
            padding: 15px;
            font-weight: bold;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            left: 15px;
            padding: 0 10px;
            color: #6c5ce7;
        }
        
        /* 버튼 */
        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #6c5ce7, stop:1 #5f4dd0);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: bold;
            min-height: 20px;
        }
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #7d6ff0, stop:1 #6c5ce7);
        }
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #5f4dd0, stop:1 #4e3fc0);
        }
        QPushButton:disabled {
            background: #b2bec3;
            color: #636e72;
        }
        
        /* 보조 버튼 */
        QPushButton[secondary="true"] {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #74b9ff, stop:1 #5a9fea);
            border: none;
        }
        QPushButton[secondary="true"]:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #81c4ff, stop:1 #74b9ff);
        }
        
        /* 입력 필드 */
        QLineEdit {
            background-color: #ffffff;
            border: 2px solid #dfe6e9;
            border-radius: 8px;
            padding: 10px;
            color: #2d3436;
            selection-background-color: #6c5ce7;
        }
        QLineEdit:focus {
            border-color: #6c5ce7;
        }
        QLineEdit:disabled {
            background-color: #f1f2f6;
            color: #b2bec3;
        }
        
        /* 라디오 버튼 & 체크박스 */
        QRadioButton, QCheckBox {
            spacing: 10px;
            padding: 5px;
        }
        QRadioButton::indicator, QCheckBox::indicator {
            width: 20px;
            height: 20px;
        }
        QRadioButton::indicator:unchecked, QCheckBox::indicator:unchecked {
            background-color: #ffffff;
            border: 2px solid #dfe6e9;
            border-radius: 10px;
        }
        QCheckBox::indicator:unchecked {
            border-radius: 5px;
        }
        QRadioButton::indicator:checked, QCheckBox::indicator:checked {
            background-color: #6c5ce7;
            border: 2px solid #6c5ce7;
            border-radius: 10px;
        }
        QCheckBox::indicator:checked {
            border-radius: 5px;
        }
        
        /* 테이블 */
        QTableWidget {
            background-color: #ffffff;
            alternate-background-color: #f8f9fa;
            border: 1px solid #dfe6e9;
            border-radius: 8px;
            gridline-color: #dfe6e9;
        }
        QTableWidget::item {
            padding: 8px;
            border-bottom: 1px solid #dfe6e9;
        }
        QTableWidget::item:selected {
            background-color: #6c5ce7;
            color: white;
        }
        QHeaderView::section {
            background-color: #f1f2f6;
            color: #2d3436;
            padding: 10px;
            border: none;
            font-weight: bold;
        }
        
        /* 진행률 바 */
        QProgressBar {
            background-color: #dfe6e9;
            border: none;
            border-radius: 10px;
            height: 20px;
            text-align: center;
            color: #2d3436;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #6c5ce7, stop:1 #a29bfe);
            border-radius: 10px;
        }
        
        /* 스크롤바 */
        QScrollBar:vertical {
            background-color: #f1f2f6;
            width: 12px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background-color: #b2bec3;
            border-radius: 6px;
            min-height: 30px;
        }
        QScrollBar::handle:vertical:hover {
            background-color: #6c5ce7;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        
        /* 드롭 영역 */
        QFrame[dropZone="true"] {
            background-color: #f8f9fa;
            border: 2px dashed #b2bec3;
            border-radius: 15px;
        }
        QFrame[dropZone="true"]:hover {
            border-color: #6c5ce7;
            background-color: #f0f0ff;
        }
        
        /* 콤보박스 */
        QComboBox {
            background-color: #ffffff;
            border: 2px solid #dfe6e9;
            border-radius: 8px;
            padding: 8px 15px;
            color: #2d3436;
        }
        QComboBox:hover {
            border-color: #6c5ce7;
        }
        QComboBox::drop-down {
            border: none;
            padding-right: 10px;
        }
        QComboBox QAbstractItemView {
            background-color: #ffffff;
            border: 1px solid #dfe6e9;
            selection-background-color: #6c5ce7;
        }
        
        /* 레이블 */
        QLabel[heading="true"] {
            font-size: 14pt;
            font-weight: bold;
            color: #6c5ce7;
        }
        QLabel[subheading="true"] {
            font-size: 9pt;
            color: #636e72;
        }
    """
    
    @staticmethod
    def get_theme(theme_name: str) -> str:
        if theme_name == "dark":
            return ThemeManager.DARK_THEME
        return ThemeManager.LIGHT_THEME


# ============================================================================
# 유틸리티 함수들
# ============================================================================

def is_admin() -> bool:
    """관리자 권한 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logger.warning(f"관리자 권한 확인 실패: {e}")
        return False


def load_config() -> dict:
    """설정 로드"""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        logger.error(f"설정 로드 실패: {e}")
    return {}


def save_config(config: dict) -> None:
    """설정 저장"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error(f"설정 저장 실패: {e}")


# ============================================================================
# 변환 엔진 (수정 없음 - 기존 로직 유지)
# ============================================================================

class HWPConverter:
    """한글 변환 엔진 - 기존 로직 완전 유지"""
    
    def __init__(self):
        self.hwp = None
        self.progid_used = None
        self.is_initialized = False
        
    def initialize(self) -> bool:
        """COM 초기화 및 한글 객체 생성"""
        if self.is_initialized:
            return True
            
        try:
            pythoncom.CoInitialize()
        except Exception as e:
            logger.debug(f"CoInitialize 오류 (무시 가능): {e}")
        
        errors = []
        for progid in HWP_PROGIDS:
            try:
                self.hwp = win32com.client.Dispatch(progid)
                self.progid_used = progid
                
                # 한글 설정
                try:
                    self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                except Exception:
                    pass  # 일부 버전에서는 지원하지 않음
                
                self.hwp.SetMessageBoxMode(0x00000001)  # 메시지 박스 비활성화
                self.is_initialized = True
                logger.info(f"한글 연결 성공: {progid}")
                return True
                
            except Exception as e:
                errors.append(f"{progid}: {str(e)}")
                continue
        
        # 모든 시도 실패
        error_detail = "\n".join(errors)
        raise Exception(f"한글 COM 객체 생성 실패\n\n시도한 ProgID:\n{error_detail}")
    
    def convert_file(self, input_path, output_path, format_type="PDF") -> Tuple[bool, Optional[str]]:
        """단일 파일 변환
        
        Returns:
            (성공여부, 오류메시지) 튜플
        """
        if not self.is_initialized:
            return False, "한글 객체가 초기화되지 않았습니다"
        
        try:
            # 파일 열기
            input_str = str(input_path)
            output_str = str(output_path)
            
            self.hwp.Open(input_str, "HWP", "forceopen:true")
            
            # 저장 형식 결정
            save_format = "PDF" if format_type == "PDF" else "HWPX"
            
            # 저장 시도 (3가지 방식으로 폴백)
            save_success = False
            save_error = None
            
            # 시도 1: SaveAs with 3 parameters (빈 문자열 추가 - 한글 2022 호환)
            try:
                self.hwp.SaveAs(output_str, save_format, "")
                save_success = True
                logger.debug(f"SaveAs 3-param 성공: {output_str}")
            except Exception as e1:
                save_error = str(e1)
                logger.debug(f"SaveAs 3-param 실패: {e1}")
                
                # 시도 2: SaveAs with 2 parameters (기존 방식)
                try:
                    self.hwp.SaveAs(output_str, save_format)
                    save_success = True
                    logger.debug(f"SaveAs 2-param 성공: {output_str}")
                except Exception as e2:
                    save_error = str(e2)
                    logger.debug(f"SaveAs 2-param 실패: {e2}")
                    
                    # 시도 3: HAction 사용 (PDF만)
                    if format_type == "PDF":
                        try:
                            # PDF 저장 액션
                            act = self.hwp.CreateAction("FileSaveAsPdf")
                            pset = act.CreateSet()
                            act.GetDefault(pset)
                            pset.SetItem("filename", output_str)
                            pset.SetItem("Format", "PDF")
                            act.Execute(pset)
                            save_success = True
                            logger.debug(f"HAction PDF 성공: {output_str}")
                        except Exception as e3:
                            save_error = f"모든 저장 방식 실패. 마지막 오류: {e3}"
                            logger.debug(f"HAction 실패: {e3}")
            
            if not save_success:
                # 문서 닫기
                try:
                    self.hwp.Clear(option=1)
                except Exception:
                    pass
                return False, save_error
            
            # 문서 닫기
            self.hwp.Clear(option=1)
            
            return True, None
            
        except Exception as e:
            error_msg = str(e)
            logger.error(f"변환 실패 ({input_path}): {error_msg}")
            # 문서 닫기 시도
            try:
                self.hwp.Clear(option=1)
            except Exception:
                pass
            
            return False, error_msg
    
    def cleanup(self) -> None:
        """정리"""
        if self.hwp and self.is_initialized:
            try:
                self.hwp.Clear(3)  # 모든 문서 닫기
            except Exception:
                pass
            
            try:
                self.hwp.Quit()
            except Exception:
                pass
            
            self.hwp = None
            self.is_initialized = False
        
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class ConversionTask:
    """변환 작업 정보 - 기존 로직 유지"""
    
    def __init__(self, input_file, output_file):
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.status = "대기"  # 대기, 진행중, 성공, 실패
        self.error = None


# ============================================================================
# 워커 스레드
# ============================================================================

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
        try:
            import pythoncom
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
            
            # 완료
            self.progress_updated.emit(total, total, "완료")
            
            if not self.cancel_requested:
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
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


# ============================================================================
# 드래그 앤 드롭 영역
# ============================================================================

class DropArea(QFrame):
    """파일 드래그 앤 드롭 영역"""
    
    files_dropped = pyqtSignal(list)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setProperty("dropZone", True)
        self.setMinimumHeight(100)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.icon_label = QLabel("📂")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.icon_label.font()
        font.setPointSize(24)
        self.icon_label.setFont(font)
        
        self.text_label = QLabel("여기에 파일을 드래그하거나 클릭하여 추가")
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setProperty("subheading", True)
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
    
    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("border-color: #e94560 !important; background-color: #1a3050 !important;")
    
    def dragLeaveEvent(self, event) -> None:
        self.setStyleSheet("")
    
    def dropEvent(self, event: QDropEvent) -> None:
        self.setStyleSheet("")
        files = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.hwp', '.hwpx')):
                files.append(path)
        if files:
            self.files_dropped.emit(files)
    
    def mousePressEvent(self, event) -> None:
        # 클릭 시 파일 선택 다이얼로그
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "파일 선택",
            "",
            "한글 파일 (*.hwp *.hwpx);;모든 파일 (*.*)"
        )
        if files:
            self.files_dropped.emit(files)


# ============================================================================
# 결과 다이얼로그
# ============================================================================

class ResultDialog(QDialog):
    """변환 결과 다이얼로그"""
    
    def __init__(self, success: int, total: int, failed_tasks: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("변환 완료")
        self.setMinimumSize(600, 400)
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # 요약
        summary_frame = QFrame()
        summary_layout = QVBoxLayout(summary_frame)
        
        success_label = QLabel(f"✅ 성공: {success}개")
        success_label.setProperty("heading", True)
        summary_layout.addWidget(success_label)
        
        failed = total - success
        if failed > 0:
            failed_label = QLabel(f"❌ 실패: {failed}개")
            failed_label.setStyleSheet("font-size: 12pt; color: #e94560;")
            summary_layout.addWidget(failed_label)
        
        layout.addWidget(summary_frame)
        
        # 실패 목록
        if failed_tasks:
            failed_group = QGroupBox("실패한 파일")
            failed_layout = QVBoxLayout(failed_group)
            
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            
            for task in failed_tasks:
                text_edit.append(f"📄 {task.input_file.name}")
                text_edit.append(f"   오류: {task.error}\n")
            
            failed_layout.addWidget(text_edit)
            layout.addWidget(failed_group)
        
        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        close_btn.setMaximumWidth(150)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)


# ============================================================================
# 메인 윈도우
# ============================================================================

class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        
        # 설정 로드
        self.config = load_config()
        self.current_theme = self.config.get("theme", "dark")
        
        # 변수 초기화
        self.tasks = []
        self.worker = None
        self.is_converting = False
        self.file_list = []
        
        # UI 초기화
        self._init_ui()
        self._apply_theme()
        self._update_mode_ui()
        self._update_output_ui()
        
        logger.info("HWP 변환기 v7.0 시작")
    
    def _init_ui(self) -> None:
        """UI 초기화"""
        self.setWindowTitle("HWP 변환기 v7.0 - PyQt6")
        self.setMinimumSize(750, 700)
        self.resize(800, 900)
        
        # 스크롤 영역 설정
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.setCentralWidget(scroll_area)
        
        # 스크롤 컨텐츠 위젯
        scroll_content = QWidget()
        scroll_area.setWidget(scroll_content)
        
        main_layout = QVBoxLayout(scroll_content)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(25, 25, 25, 25)
        
        # === 헤더 ===
        header_layout = QHBoxLayout()
        
        title_label = QLabel("HWP / HWPX 변환기")
        title_label.setProperty("heading", True)
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # 테마 전환 버튼
        self.theme_btn = QPushButton("🌙 다크" if self.current_theme == "dark" else "☀️ 라이트")
        self.theme_btn.setProperty("secondary", True)
        self.theme_btn.setFixedWidth(100)
        self.theme_btn.clicked.connect(self._toggle_theme)
        header_layout.addWidget(self.theme_btn)
        
        main_layout.addLayout(header_layout)
        
        # === 모드 선택 ===
        mode_group = QGroupBox("변환 모드")
        mode_layout = QVBoxLayout(mode_group)
        mode_layout.setSpacing(8)
        
        self.mode_group = QButtonGroup(self)
        
        self.folder_radio = QRadioButton("📁 폴더 일괄 변환 (폴더 내 모든 파일)")
        self.files_radio = QRadioButton("📄 파일 개별 선택 (원하는 파일만)")
        
        self.mode_group.addButton(self.folder_radio, 0)
        self.mode_group.addButton(self.files_radio, 1)
        
        mode_layout.addWidget(self.folder_radio)
        mode_layout.addWidget(self.files_radio)
        
        saved_mode = self.config.get("mode", "folder")
        if saved_mode == "folder":
            self.folder_radio.setChecked(True)
        else:
            self.files_radio.setChecked(True)
        
        self.folder_radio.toggled.connect(self._update_mode_ui)
        
        main_layout.addWidget(mode_group)
        
        # === 입력 영역 ===
        input_group = QGroupBox("입력")
        input_layout = QVBoxLayout(input_group)
        input_layout.setSpacing(12)
        
        # 폴더 모드 위젯
        self.folder_widget = QWidget()
        folder_layout = QVBoxLayout(self.folder_widget)
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(10)
        
        folder_row = QHBoxLayout()
        folder_row.setSpacing(10)
        self.folder_entry = QLineEdit()
        self.folder_entry.setPlaceholderText("변환할 폴더를 선택하세요...")
        self.folder_entry.setReadOnly(True)
        self.folder_entry.setMinimumHeight(40)
        folder_row.addWidget(self.folder_entry)
        
        folder_btn = QPushButton("찾아보기")
        folder_btn.setProperty("secondary", True)
        folder_btn.setFixedWidth(100)
        folder_btn.setMinimumHeight(40)
        folder_btn.clicked.connect(self._select_folder)
        folder_row.addWidget(folder_btn)
        
        folder_layout.addLayout(folder_row)
        
        self.include_sub_check = QCheckBox("하위 폴더 포함")
        self.include_sub_check.setChecked(self.config.get("include_sub", True))
        folder_layout.addWidget(self.include_sub_check)
        
        input_layout.addWidget(self.folder_widget)
        
        # 파일 모드 위젯
        self.files_widget = QWidget()
        files_layout = QVBoxLayout(self.files_widget)
        files_layout.setContentsMargins(0, 0, 0, 0)
        files_layout.setSpacing(12)
        
        # 드롭 영역 - 고정 높이
        self.drop_area = DropArea()
        self.drop_area.setFixedHeight(120)
        self.drop_area.files_dropped.connect(self._add_files)
        files_layout.addWidget(self.drop_area)
        
        # 버튼 행
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        
        add_btn = QPushButton("➕ 파일 추가")
        add_btn.setProperty("secondary", True)
        add_btn.setMinimumHeight(36)
        add_btn.clicked.connect(self._browse_files)
        btn_row.addWidget(add_btn)
        
        remove_btn = QPushButton("➖ 선택 제거")
        remove_btn.setProperty("secondary", True)
        remove_btn.setMinimumHeight(36)
        remove_btn.clicked.connect(self._remove_selected)
        btn_row.addWidget(remove_btn)
        
        clear_btn = QPushButton("🗑️ 전체 제거")
        clear_btn.setProperty("secondary", True)
        clear_btn.setMinimumHeight(36)
        clear_btn.clicked.connect(self._clear_all)
        btn_row.addWidget(clear_btn)
        
        btn_row.addStretch()
        files_layout.addLayout(btn_row)
        
        # 파일 테이블 - 고정 높이
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(2)
        self.file_table.setHorizontalHeaderLabels(["파일명", "경로"])
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.file_table.setAlternatingRowColors(True)
        self.file_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.file_table.setFixedHeight(180)
        self.file_table.verticalHeader().setVisible(False)
        files_layout.addWidget(self.file_table)
        
        input_layout.addWidget(self.files_widget)
        
        main_layout.addWidget(input_group)
        
        # === 출력 설정 ===
        output_group = QGroupBox("출력")
        output_layout = QVBoxLayout(output_group)
        output_layout.setSpacing(10)
        
        self.same_location_check = QCheckBox("입력 파일과 같은 위치에 저장")
        self.same_location_check.setChecked(self.config.get("same_location", True))
        self.same_location_check.toggled.connect(self._update_output_ui)
        output_layout.addWidget(self.same_location_check)
        
        output_row = QHBoxLayout()
        output_row.setSpacing(10)
        output_label = QLabel("저장 폴더:")
        output_label.setFixedWidth(70)
        output_row.addWidget(output_label)
        
        self.output_entry = QLineEdit()
        self.output_entry.setPlaceholderText("저장할 폴더를 선택하세요...")
        self.output_entry.setReadOnly(True)
        self.output_entry.setMinimumHeight(40)
        output_row.addWidget(self.output_entry)
        
        self.output_btn = QPushButton("찾아보기")
        self.output_btn.setProperty("secondary", True)
        self.output_btn.setFixedWidth(100)
        self.output_btn.setMinimumHeight(40)
        self.output_btn.clicked.connect(self._select_output)
        output_row.addWidget(self.output_btn)
        
        output_layout.addLayout(output_row)
        
        main_layout.addWidget(output_group)
        
        # === 변환 옵션 ===
        options_group = QGroupBox("변환 옵션")
        options_layout = QVBoxLayout(options_group)
        options_layout.setSpacing(10)
        
        # 변환 형식
        format_layout = QHBoxLayout()
        format_label = QLabel("변환 형식:")
        format_label.setFixedWidth(70)
        format_layout.addWidget(format_label)
        
        self.format_group = QButtonGroup(self)
        
        self.pdf_radio = QRadioButton("📕 PDF")
        self.hwpx_radio = QRadioButton("📘 HWPX")
        
        self.format_group.addButton(self.pdf_radio, 0)
        self.format_group.addButton(self.hwpx_radio, 1)
        
        saved_format = self.config.get("format", "PDF")
        if saved_format == "PDF":
            self.pdf_radio.setChecked(True)
        else:
            self.hwpx_radio.setChecked(True)
        
        format_layout.addWidget(self.pdf_radio)
        format_layout.addWidget(self.hwpx_radio)
        format_layout.addStretch()
        
        options_layout.addLayout(format_layout)
        
        # 덮어쓰기 옵션
        self.overwrite_check = QCheckBox("기존 파일 덮어쓰기 (체크 해제 시 번호 자동 추가)")
        self.overwrite_check.setChecked(self.config.get("overwrite", False))
        options_layout.addWidget(self.overwrite_check)
        
        main_layout.addWidget(options_group)
        
        # === 실행 버튼 ===
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        
        self.start_btn = QPushButton("🚀 변환 시작")
        self.start_btn.setMinimumHeight(55)
        font = self.start_btn.font()
        font.setPointSize(12)
        font.setBold(True)
        self.start_btn.setFont(font)
        self.start_btn.clicked.connect(self._start_conversion)
        btn_layout.addWidget(self.start_btn)
        
        self.cancel_btn = QPushButton("⏹️ 취소")
        self.cancel_btn.setProperty("secondary", True)
        self.cancel_btn.setMinimumHeight(55)
        self.cancel_btn.setFixedWidth(100)
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self._cancel_conversion)
        btn_layout.addWidget(self.cancel_btn)
        
        main_layout.addLayout(btn_layout)
        
        # === 진행 상태 ===
        progress_group = QGroupBox("진행 상태")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setSpacing(8)
        
        self.status_label = QLabel("준비됨")
        self.status_label.setMinimumHeight(25)
        progress_layout.addWidget(self.status_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setMinimumHeight(28)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("0 / 0")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        progress_layout.addWidget(self.progress_label)
        
        main_layout.addWidget(progress_group)
        
        # 하단 여백
        main_layout.addSpacing(20)
    
    def _apply_theme(self) -> None:
        """테마 적용"""
        theme_css = ThemeManager.get_theme(self.current_theme)
        self.setStyleSheet(theme_css)
    
    def _toggle_theme(self) -> None:
        """테마 전환"""
        if self.current_theme == "dark":
            self.current_theme = "light"
            self.theme_btn.setText("☀️ 라이트")
        else:
            self.current_theme = "dark"
            self.theme_btn.setText("🌙 다크")
        
        self._apply_theme()
        self.config["theme"] = self.current_theme
        save_config(self.config)
    
    def _update_mode_ui(self) -> None:
        """모드에 따라 UI 업데이트"""
        is_folder_mode = self.folder_radio.isChecked()
        self.folder_widget.setVisible(is_folder_mode)
        self.files_widget.setVisible(not is_folder_mode)
    
    def _update_output_ui(self) -> None:
        """출력 폴더 UI 상태 업데이트"""
        same_location = self.same_location_check.isChecked()
        self.output_entry.setEnabled(not same_location)
        self.output_btn.setEnabled(not same_location)
    
    def _select_folder(self) -> None:
        """폴더 선택"""
        initial = self.config.get("last_folder", "")
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", initial)
        if folder:
            self.folder_entry.setText(folder)
            self.config["last_folder"] = folder
    
    def _select_output(self) -> None:
        """출력 폴더 선택"""
        initial = self.config.get("last_output", "")
        folder = QFileDialog.getExistingDirectory(self, "출력 폴더 선택", initial)
        if folder:
            self.output_entry.setText(folder)
            self.config["last_output"] = folder
    
    def _browse_files(self) -> None:
        """파일 찾아보기"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "파일 선택",
            "",
            "한글 파일 (*.hwp *.hwpx);;모든 파일 (*.*)"
        )
        if files:
            self._add_files(files)
    
    def _add_files(self, files: list) -> None:
        """파일 추가"""
        added = 0
        for file_path in files:
            if file_path not in self.file_list:
                self.file_list.append(file_path)
                
                row = self.file_table.rowCount()
                self.file_table.insertRow(row)
                
                name = Path(file_path).name
                self.file_table.setItem(row, 0, QTableWidgetItem(name))
                self.file_table.setItem(row, 1, QTableWidgetItem(str(Path(file_path).parent)))
                
                added += 1
        
        if added > 0:
            self.status_label.setText(f"{added}개 파일 추가됨 (총 {len(self.file_list)}개)")
    
    def _remove_selected(self) -> None:
        """선택된 파일 제거"""
        selected = self.file_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "경고", "제거할 파일을 선택하세요.")
            return
        
        rows = set(item.row() for item in selected)
        for row in sorted(rows, reverse=True):
            if row < len(self.file_list):
                del self.file_list[row]
            self.file_table.removeRow(row)
        
        self.status_label.setText(f"선택 파일 제거됨 (총 {len(self.file_list)}개)")
    
    def _clear_all(self) -> None:
        """전체 파일 제거"""
        if not self.file_list:
            return
        
        reply = QMessageBox.question(
            self, "확인",
            f"{len(self.file_list)}개 파일을 모두 제거하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.file_list.clear()
            self.file_table.setRowCount(0)
            self.status_label.setText("모든 파일 제거됨")
    
    def _collect_tasks(self) -> List[ConversionTask]:
        """변환 작업 목록 생성"""
        tasks = []
        is_folder_mode = self.folder_radio.isChecked()
        format_type = "PDF" if self.pdf_radio.isChecked() else "HWPX"
        output_ext = ".pdf" if format_type == "PDF" else ".hwpx"
        
        if is_folder_mode:
            folder_path = self.folder_entry.text()
            if not folder_path:
                raise ValueError("폴더를 선택하세요.")
            
            folder = Path(folder_path)
            if not folder.exists():
                raise ValueError("폴더가 존재하지 않습니다.")
            
            # 검색할 확장자
            if format_type == "PDF":
                patterns = ["*.hwp", "*.hwpx"]
            else:
                patterns = ["*.hwp"]
            
            # 파일 검색
            input_files = []
            if self.include_sub_check.isChecked():
                for pattern in patterns:
                    input_files.extend(folder.rglob(pattern))
            else:
                for pattern in patterns:
                    input_files.extend(folder.glob(pattern))
            
            if not input_files:
                raise ValueError("변환할 파일이 없습니다.")
            
            # 작업 생성
            for input_file in input_files:
                if self.same_location_check.isChecked():
                    output_file = input_file.parent / (input_file.stem + output_ext)
                else:
                    output_folder = Path(self.output_entry.text())
                    if not output_folder:
                        raise ValueError("출력 폴더를 선택하세요.")
                    
                    rel_path = input_file.relative_to(folder)
                    output_file = output_folder / rel_path.parent / (input_file.stem + output_ext)
                
                tasks.append(ConversionTask(input_file, output_file))
        
        else:  # 파일 모드
            if not self.file_list:
                raise ValueError("파일을 추가하세요.")
            
            for file_path in self.file_list:
                input_file = Path(file_path)
                
                if self.same_location_check.isChecked():
                    output_file = input_file.parent / (input_file.stem + output_ext)
                else:
                    output_folder = Path(self.output_entry.text())
                    if not output_folder:
                        raise ValueError("출력 폴더를 선택하세요.")
                    
                    output_file = output_folder / (input_file.stem + output_ext)
                
                tasks.append(ConversionTask(input_file, output_file))
        
        return tasks
    
    def _adjust_output_paths(self, tasks: List[ConversionTask]) -> None:
        """출력 경로 조정 (덮어쓰기 방지)"""
        for task in tasks:
            if task.output_file.exists():
                counter = 1
                stem = task.output_file.stem
                ext = task.output_file.suffix
                parent = task.output_file.parent
                
                while True:
                    new_name = f"{stem} ({counter}){ext}"
                    new_path = parent / new_name
                    if not new_path.exists():
                        task.output_file = new_path
                        break
                    counter += 1
    
    def _save_settings(self) -> None:
        """설정 저장"""
        self.config["mode"] = "folder" if self.folder_radio.isChecked() else "files"
        self.config["format"] = "PDF" if self.pdf_radio.isChecked() else "HWPX"
        self.config["include_sub"] = self.include_sub_check.isChecked()
        self.config["same_location"] = self.same_location_check.isChecked()
        self.config["overwrite"] = self.overwrite_check.isChecked()
        save_config(self.config)
    
    def _start_conversion(self) -> None:
        """변환 시작"""
        try:
            # 작업 목록 생성
            self.tasks = self._collect_tasks()
            
            # 덮어쓰기 확인
            if not self.overwrite_check.isChecked():
                self._adjust_output_paths(self.tasks)
            
            # 설정 저장
            self._save_settings()
            
            # UI 업데이트
            self._set_converting_state(True)
            
            # 진행률 초기화
            self.progress_bar.setMaximum(len(self.tasks))
            self.progress_bar.setValue(0)
            
            # 워커 시작
            format_type = "PDF" if self.pdf_radio.isChecked() else "HWPX"
            self.worker = ConversionWorker(self.tasks, format_type)
            self.worker.progress_updated.connect(self._on_progress_updated)
            self.worker.status_updated.connect(self._on_status_updated)
            self.worker.task_completed.connect(self._on_task_completed)
            self.worker.error_occurred.connect(self._on_error_occurred)
            self.worker.finished.connect(self._on_worker_finished)
            self.worker.start()
            
        except ValueError as e:
            QMessageBox.warning(self, "경고", str(e))
        except Exception as e:
            logger.exception("변환 시작 오류")
            QMessageBox.critical(self, "오류", f"오류 발생: {e}")
    
    def _cancel_conversion(self) -> None:
        """변환 취소"""
        reply = QMessageBox.question(
            self, "확인",
            "변환을 취소하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes and self.worker:
            self.worker.cancel()
            self.status_label.setText("취소 중...")
    
    def _set_converting_state(self, converting: bool) -> None:
        """변환 중 상태 설정"""
        self.is_converting = converting
        self.start_btn.setEnabled(not converting)
        self.cancel_btn.setEnabled(converting)
    
    def _on_progress_updated(self, current: int, total: int, filename: str) -> None:
        """진행률 업데이트"""
        self.progress_bar.setValue(current)
        self.progress_label.setText(f"{current} / {total}")
        self.status_label.setText(f"변환 중: {filename}")
    
    def _on_status_updated(self, text: str) -> None:
        """상태 텍스트 업데이트"""
        self.status_label.setText(text)
    
    def _on_task_completed(self, success: int, total: int, failed_tasks: list) -> None:
        """작업 완료"""
        dialog = ResultDialog(success, total, failed_tasks, self)
        dialog.exec()
    
    def _on_error_occurred(self, error_msg: str) -> None:
        """오류 발생"""
        QMessageBox.critical(self, "오류", f"변환 중 오류 발생:\n{error_msg}")
    
    def _on_worker_finished(self) -> None:
        """워커 종료"""
        self._set_converting_state(False)
        
        # 시그널 연결 해제 (메모리 누수 방지)
        if self.worker:
            try:
                self.worker.progress_updated.disconnect()
                self.worker.status_updated.disconnect()
                self.worker.task_completed.disconnect()
                self.worker.error_occurred.disconnect()
                self.worker.finished.disconnect()
            except (TypeError, RuntimeError):
                pass  # 이미 연결 해제된 경우
        
        self.worker = None
    
    def closeEvent(self, event) -> None:
        """윈도우 닫기 이벤트"""
        if self.is_converting:
            reply = QMessageBox.question(
                self, "확인",
                "변환 작업이 진행 중입니다. 종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            
            if self.worker:
                self.worker.cancel()
                self.worker.wait(3000)  # 최대 3초 대기
        
        save_config(self.config)
        event.accept()


# ============================================================================
# 메인 함수
# ============================================================================

def main():
    """메인 함수"""
    
    # pywin32 확인
    if not PYWIN32_AVAILABLE:
        app = QApplication(sys.argv)
        QMessageBox.critical(
            None, "오류",
            "pywin32 라이브러리가 필요합니다.\n\npip install pywin32"
        )
        return
    
    # 관리자 권한 확인
    if not is_admin():
        app = QApplication(sys.argv)
        QMessageBox.warning(
            None,
            "관리자 권한 필요",
            "이 프로그램은 관리자 권한으로 실행해야 합니다.\n\n"
            "파일을 마우스 오른쪽 버튼으로 클릭하여\n"
            "'관리자 권한으로 실행'을 선택하세요."
        )
        sys.exit(1)
    
    # 애플리케이션 실행
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
