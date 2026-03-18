from __future__ import annotations

from pathlib import Path
from typing import Optional

from PyQt6.QtCore import QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QDragEnterEvent, QDragLeaveEvent, QDragMoveEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import QFileDialog, QFrame, QGraphicsDropShadowEffect, QLabel, QVBoxLayout

from ..constants import FEEDBACK_RESET_DELAY, SUPPORTED_EXTENSIONS
from ..logging_config import get_logger

logger = get_logger(__name__)

class DropArea(QFrame):
    """파일 드래그 앤 드롭 영역
    
    Note: Qt의 OLE 드래그 앤 드롭(setAcceptDrops)을 비활성화합니다.
    관리자 권한으로 실행 시 UIPI가 OLE 드롭을 차단하기 때문에,
    Windows 네이티브 WM_DROPFILES만 사용합니다.
    """
    
    files_dropped = pyqtSignal(list)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # Qt OLE 드래그 앤 드롭 비활성화 (관리자 권한에서 UIPI 차단됨)
        # 대신 MainWindow에서 네이티브 WM_DROPFILES 사용
        self.setAcceptDrops(False)
        self.setProperty("dropZone", True)
        self.setMinimumHeight(100)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("HWP/HWPX 파일을 드래그하여 추가하거나 클릭하여 선택하세요")
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.icon_label = QLabel("📂")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = self.icon_label.font()
        font.setPointSize(28)
        self.icon_label.setFont(font)
        
        self.text_label = QLabel("여기에 파일을 드래그하거나 클릭하여 추가")
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setProperty("subheading", True)
        
        self.hint_label = QLabel("HWP, HWPX 파일 지원")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setStyleSheet("font-size: 8pt; color: #666680;")
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.hint_label)
        
        # 원본 텍스트 저장
        self._original_icon = "📂"
        self._original_text = "여기에 파일을 드래그하거나 클릭하여 추가"
    
    def _get_files_from_urls(self, urls) -> list:
        """URL 목록에서 스캔 대상 경로(지원 파일/폴더) 추출"""
        files = []
        for url in urls:
            path = url.toLocalFile()
            if not path:
                continue
            
            path_obj = Path(path)
            if path_obj.is_dir() or (path_obj.is_file() and path.lower().endswith(SUPPORTED_EXTENSIONS)):
                files.append(path)
        return files
    
    def _has_valid_content(self, mime_data) -> bool:
        """유효한 HWP/HWPX 파일이 있는지 확인"""
        if not mime_data.hasUrls():
            return False
        
        for url in mime_data.urls():
            path = url.toLocalFile()
            if not path:
                continue
            
            path_obj = Path(path)
            if path_obj.is_file() and path.lower().endswith(SUPPORTED_EXTENSIONS):
                return True
            elif path_obj.is_dir():
                # 폴더인 경우에도 허용
                return True
        return False
    
    def dragEnterEvent(self, a0: Optional[QDragEnterEvent]) -> None:
        """드래그 진입 이벤트"""
        if a0 is None:
            return
        mime_data = a0.mimeData()
        if mime_data is None:
            a0.ignore()
            logger.debug("dragEnterEvent - mimeData 없음")
            return

        logger.debug(f"dragEnterEvent 호출됨 - hasUrls: {mime_data.hasUrls()}")
        
        if mime_data.hasUrls():
            urls = mime_data.urls()
            logger.debug(f"URL 개수: {len(urls)}, 첫번째: {urls[0].toLocalFile() if urls else 'N/A'}")
            
            if self._has_valid_content(mime_data):
                a0.acceptProposedAction()
                self.icon_label.setText("📥")
                self.text_label.setText("파일을 놓으세요!")
                self.setStyleSheet("border-color: #e94560 !important; background-color: #1a3050 !important;")
                logger.debug("드래그 수락됨")
            else:
                a0.ignore()
                self.text_label.setText("지원하지 않는 파일 형식입니다")
                logger.debug("유효하지 않은 콘텐츠 - 무시됨")
        else:
            a0.ignore()
            logger.debug("URL 없음 - 무시됨")
    
    def dragMoveEvent(self, a0: Optional[QDragMoveEvent]) -> None:
        """드래그 이동 이벤트 - 드래그 중 계속 호출됨"""
        if a0 is None:
            return
        mime_data = a0.mimeData()
        if mime_data is not None and mime_data.hasUrls():
            a0.acceptProposedAction()
        else:
            a0.ignore()
    
    def dragLeaveEvent(self, a0: Optional[QDragLeaveEvent]) -> None:
        """드래그 이탈 이벤트"""
        del a0
        self._reset_appearance()
    
    def dropEvent(self, a0: Optional[QDropEvent]) -> None:
        """드롭 이벤트"""
        if a0 is None:
            return
        logger.debug("dropEvent 호출됨")
        self._reset_appearance()
        mime_data = a0.mimeData()
        if mime_data is None:
            logger.debug("dropEvent - mimeData 없음")
            a0.ignore()
            return
        
        if not mime_data.hasUrls():
            logger.debug("dropEvent - URL 없음")
            a0.ignore()
            return
        
        files = self._get_files_from_urls(mime_data.urls())
        logger.debug(f"dropEvent - 추출된 파일 수: {len(files)}")
        
        if files:
            a0.acceptProposedAction()
            self.files_dropped.emit(files)
            # 성공 피드백
            self.icon_label.setText("✅")
            self.text_label.setText(f"{len(files)}개 경로 스캔 시작")
            QTimer.singleShot(FEEDBACK_RESET_DELAY, self._reset_appearance)
            logger.debug(f"드래그 앤 드롭 입력 수신: {len(files)}개 경로")
        else:
            a0.ignore()
            self.text_label.setText("HWP/HWPX 파일이 없습니다")
            QTimer.singleShot(FEEDBACK_RESET_DELAY, self._reset_appearance)
            logger.debug("dropEvent - 유효한 HWP/HWPX 파일 없음")
    
    def _reset_appearance(self) -> None:
        """외관 초기화"""
        self.icon_label.setText(self._original_icon)
        self.text_label.setText(self._original_text)
        self.setStyleSheet("")
    
    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        """클릭 시 파일 선택 다이얼로그"""
        del a0
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "파일 선택",
            "",
            "한글 파일 (*.hwp *.hwpx);;모든 파일 (*.*)"
        )
        if files:
            self.files_dropped.emit(files)


# ============================================================================
# 포맷 선택 카드
# ============================================================================

class FormatCard(QFrame):
    """변환 형식 선택 카드"""
    
    clicked = pyqtSignal(str)  # format_type 시그널
    
    def __init__(self, format_type: str, icon: str, title: str, description: str, parent=None):
        super().__init__(parent)
        self.format_type = format_type
        self._selected = False
        
        self.setProperty("formatCard", True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumSize(140, 100)
        self.setMaximumWidth(180)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(8)  # 간격 약간 증가
        layout.setContentsMargins(10, 15, 10, 15)  # 상하 여백 확보
        
        # 아이콘
        self.icon_label = QLabel(icon)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_font = self.icon_label.font()
        icon_font.setPointSize(24)
        self.icon_label.setFont(icon_font)
        layout.addWidget(self.icon_label)
        
        # 타이틀
        self.title_label = QLabel(title)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = self.title_label.font()
        title_font.setPointSize(11)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        layout.addWidget(self.title_label)
        
        # 설명
        self.desc_label = QLabel(description)
        self.desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.desc_label.setProperty("subheading", True)
        self.desc_label.setStyleSheet("font-size: 8pt;")
        layout.addWidget(self.desc_label)
        
        self.setToolTip(f"{title} 형식으로 변환합니다")
        
        # 그림자 효과 추가
        from PyQt6.QtWidgets import QGraphicsDropShadowEffect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)
    
    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        """클릭 이벤트"""
        del a0
        self.clicked.emit(self.format_type)
    
    def setSelected(self, selected: bool) -> None:
        """선택 상태 설정"""
        self._selected = selected
        if selected:
            self.setProperty("formatCard", False)
            self.setProperty("formatCardSelected", True)
        else:
            self.setProperty("formatCard", True)
            self.setProperty("formatCardSelected", False)
        # 스타일 갱신
        style = self.style()
        if style is None:
            return
        style.unpolish(self)
        style.polish(self)
    
    def isSelected(self) -> bool:
        return self._selected
