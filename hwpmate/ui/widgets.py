from __future__ import annotations

from pathlib import Path
from typing import Optional

from PyQt6.QtCore import QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDragLeaveEvent, QDragMoveEvent, QDropEvent, QMouseEvent
from PyQt6.QtWidgets import QFrame, QLabel, QVBoxLayout, QWidget

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
    # 클릭 시 파일 선택 요청 — 메인 창의 browse_files 흐름(입력 잠금·최근 폴더)을 재사용한다.
    browse_requested = pyqtSignal()
    
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
        # 전역 QSS 의 font-size 가 setFont 를 덮어쓰므로 크기는 테마 QSS 속성으로 지정한다.
        self.icon_label.setProperty("dropIcon", True)
        
        self.text_label = QLabel("여기에 파일을 드래그하거나 클릭하여 추가")
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setProperty("subheading", True)
        
        self.hint_label = QLabel("HWP · HWPX 파일 또는 폴더 (폴더는 하위 파일까지 추가)")
        self.hint_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setProperty("caption", True)
        
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
                self._set_drop_active(True)
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
            self.show_feedback("✅", f"{len(files)}개 경로 스캔 시작")
            logger.debug(f"드래그 앤 드롭 입력 수신: {len(files)}개 경로")
        else:
            a0.ignore()
            self.show_feedback("⚠️", "HWP/HWPX 파일이 없습니다")
            logger.debug("dropEvent - 유효한 HWP/HWPX 파일 없음")
    
    def _set_drop_active(self, active: bool) -> None:
        """테마 QSS 의 dropActive 상태를 켜고 끈다 (색상 하드코딩 없이 라이트/다크 공용)."""
        if bool(self.property("dropActive")) == active:
            return
        self.setProperty("dropActive", active)
        style = self.style()
        if style is not None:
            style.unpolish(self)
            style.polish(self)
        self.update()

    def show_feedback(self, icon: str, text: str) -> None:
        """잠시 표시 후 원래 안내로 돌아가는 피드백 메시지."""
        self.icon_label.setText(icon)
        self.text_label.setText(text)
        QTimer.singleShot(FEEDBACK_RESET_DELAY, self._reset_appearance)

    def _reset_appearance(self) -> None:
        """외관 초기화"""
        self.icon_label.setText(self._original_icon)
        self.text_label.setText(self._original_text)
        self._set_drop_active(False)
    
    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        """왼쪽 클릭 시 파일 선택 요청"""
        if a0 is not None and a0.button() != Qt.MouseButton.LeftButton:
            return
        self.browse_requested.emit()


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
        self.setMinimumSize(120, 96)
        self.setMaximumWidth(180)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(4)
        layout.setContentsMargins(10, 12, 10, 12)
        
        # 아이콘
        self.icon_label = QLabel(icon)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_label.setProperty("cardIcon", True)
        layout.addWidget(self.icon_label)
        
        # 타이틀
        self.title_label = QLabel(title)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setProperty("cardTitle", True)
        layout.addWidget(self.title_label)
        
        # 설명
        self.desc_label = QLabel(description)
        self.desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.desc_label.setProperty("caption", True)
        self.desc_label.setWordWrap(True)
        layout.addWidget(self.desc_label)
        
        self.setToolTip(f"{title} ({description}) 형식으로 변환합니다")
    
    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        """클릭 이벤트 (왼쪽 버튼만)"""
        if a0 is not None and a0.button() != Qt.MouseButton.LeftButton:
            return
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
        # 자식 라벨의 [cardTitle] 선택 색도 다시 계산
        for child in self.findChildren(QWidget):
            style.unpolish(child)
            style.polish(child)
        self.update()
    
    def isSelected(self) -> bool:
        return self._selected
