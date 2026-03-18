from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import QEasingCurve, QPropertyAnimation, QTimer, Qt, pyqtSignal
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel

from ..constants import TOAST_DURATION_DEFAULT, TOAST_FADE_DURATION
from ..logging_config import get_logger

logger = get_logger(__name__)

class ToastWidget(QFrame):
    """토스트 알림 위젯"""
    
    closed = pyqtSignal(object)  # 닫힐 때 시그널
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        
        self._setup_ui()
        self._animation = None
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._fade_out)
        
    def _setup_ui(self) -> None:
        self.setFixedSize(320, 65)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)
        
        self.icon_label = QLabel("ℹ️")
        self.icon_label.setFixedWidth(30)
        font = self.icon_label.font()
        font.setPointSize(14)
        self.icon_label.setFont(font)
        layout.addWidget(self.icon_label)
        
        self.message_label = QLabel()
        self.message_label.setWordWrap(True)
        layout.addWidget(self.message_label)
        
        self.setStyleSheet("""
            ToastWidget {
                background-color: rgba(22, 33, 62, 0.95);
                border: 1px solid #0f3460;
                border-radius: 12px;
            }
            QLabel {
                color: #eaeaea;
                font-size: 10pt;
            }
        """)
    
    def show_message(
        self,
        message: str,
        icon: str = "ℹ️",
        duration: int = TOAST_DURATION_DEFAULT,
        position_y: Optional[int] = None,
    ) -> None:
        """토스트 메시지 표시"""
        self.icon_label.setText(icon)
        self.message_label.setText(message)
        
        # 부모 윈도우 기준 위치 계산
        parent_widget = self.parentWidget()
        if parent_widget is not None:
            x = parent_widget.x() + parent_widget.width() - self.width() - 20
            if position_y is not None:
                y = position_y
            else:
                y = parent_widget.y() + parent_widget.height() - self.height() - 20
            self.move(x, y)
        
        self.setWindowOpacity(1.0)
        self.show()
        self.raise_()
        self._timer.start(duration)
    
    def _fade_out(self) -> None:
        """페이드 아웃 애니메이션"""
        self._timer.stop()
        self._animation = QPropertyAnimation(self, b"windowOpacity")
        self._animation.setDuration(TOAST_FADE_DURATION)
        self._animation.setStartValue(1.0)
        self._animation.setEndValue(0.0)
        self._animation.setEasingCurve(QEasingCurve.Type.OutQuad)
        self._animation.finished.connect(self._on_fade_finished)
        self._animation.start()
    
    def _on_fade_finished(self) -> None:
        """페이드 아웃 완료"""
        self.hide()
        self._cleanup()
        self.closed.emit(self)
    
    def _cleanup(self) -> None:
        """리소스 정리"""
        if self._timer:
            self._timer.stop()
        if self._animation:
            self._animation.stop()
            self._animation = None


class ToastManager:
    """Toast 알림 관리자 - 스택 기능 지원"""
    
    MAX_TOASTS = 3
    TOAST_HEIGHT = 70
    TOAST_SPACING = 10
    
    def __init__(self, parent=None):
        self.parent = parent
        self.toasts = []
    
    def show_message(self, message: str, icon: str = "ℹ️", duration: int = TOAST_DURATION_DEFAULT) -> None:
        """새 토스트 메시지 표시"""
        if not self.parent:
            logger.warning("ToastManager: parent가 없어 메시지를 표시할 수 없습니다")
            return
        
        try:
            # 최대 개수 초과 시 가장 오래된 것 제거
            while len(self.toasts) >= self.MAX_TOASTS:
                old_toast = self.toasts.pop(0)
                try:
                    old_toast.hide()
                    old_toast.deleteLater()
                except RuntimeError:
                    pass  # 이미 삭제된 위젯
            
            # 새 토스트 생성
            toast = ToastWidget(self.parent)
            toast.closed.connect(self._on_toast_closed)
            self.toasts.append(toast)
            
            # 위치 계산 및 표시
            self._update_positions()
            position_y = self._get_position_for_toast(len(self.toasts) - 1)
            toast.show_message(message, icon, duration, position_y)
        except Exception as e:
            logger.error(f"Toast 표시 오류: {e}")
    
    def _get_position_for_toast(self, index: int) -> int:
        """토스트 위치 계산"""
        if self.parent:
            base_y = self.parent.y() + self.parent.height() - 20
            return base_y - (index + 1) * (self.TOAST_HEIGHT + self.TOAST_SPACING)
        return 100
    
    def _update_positions(self) -> None:
        """모든 토스트 위치 업데이트"""
        if not self.parent:
            return
        
        for i, toast in enumerate(self.toasts):
            try:
                if toast.isVisible():
                    x = self.parent.x() + self.parent.width() - toast.width() - 20
                    y = self._get_position_for_toast(i)
                    toast.move(x, y)
            except RuntimeError:
                pass  # 이미 삭제된 위젯
    
    def _on_toast_closed(self, toast: ToastWidget) -> None:
        """토스트 닫힘 처리"""
        try:
            if toast in self.toasts:
                self.toasts.remove(toast)
                toast.deleteLater()
                self._update_positions()
        except RuntimeError:
            pass  # 이미 삭제된 위젯
    
    def clear_all(self) -> None:
        """모든 토스트 제거 및 정리"""
        for toast in self.toasts[:]:
            try:
                toast._cleanup()
                toast.hide()
                toast.deleteLater()
            except RuntimeError:
                pass
        self.toasts.clear()
