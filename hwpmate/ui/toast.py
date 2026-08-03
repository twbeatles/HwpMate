from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import QEasingCurve, QPropertyAnimation, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtWidgets import QFrame, QGraphicsDropShadowEffect, QHBoxLayout, QLabel

from ..constants import TOAST_DURATION_DEFAULT, TOAST_FADE_DURATION
from ..logging_config import get_logger

logger = get_logger(__name__)

# 아이콘별 테두리/강조색 (어두운 배경 + 밝은 글씨 대비)
_TOAST_ACCENT_BY_ICON: dict[str, str] = {
    "✅": "#22c55e",
    "🎉": "#22c55e",
    "🚀": "#38bdf8",
    "🔁": "#38bdf8",
    "ℹ️": "#38bdf8",
    "⚠️": "#f59e0b",
    "❌": "#ef4444",
    "🛑": "#ef4444",
    "⏭️": "#94a3b8",
}


def _accent_for_icon(icon: str) -> str:
    return _TOAST_ACCENT_BY_ICON.get(icon, "#94a3b8")


class ToastWidget(QFrame):
    """고대비 토스트 알림 위젯 (어두운 패널 + 밝은 본문 글씨)."""

    closed = pyqtSignal(object)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)

        self._setup_ui()
        self._animation = None
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._fade_out)

    def _setup_ui(self) -> None:
        self.setMinimumSize(360, 72)
        self.setMaximumWidth(440)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(12)

        self.icon_label = QLabel("ℹ️")
        self.icon_label.setFixedWidth(32)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_font = QFont(self.icon_label.font())
        icon_font.setPointSize(16)
        self.icon_label.setFont(icon_font)
        layout.addWidget(self.icon_label)

        self.message_label = QLabel()
        self.message_label.setWordWrap(True)
        self.message_label.setAlignment(
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        )
        msg_font = QFont(self.message_label.font())
        msg_font.setPointSize(11)
        msg_font.setBold(True)
        self.message_label.setFont(msg_font)
        layout.addWidget(self.message_label, stretch=1)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 160))
        self.setGraphicsEffect(shadow)

        self._apply_style("#38bdf8")

    def _apply_style(self, accent: str) -> None:
        """어두운 배경 + 거의 흰색 본문으로 대비를 확보."""
        self.setStyleSheet(
            f"""
            ToastWidget {{
                background-color: rgba(15, 23, 42, 0.97);
                border: 2px solid {accent};
                border-left: 6px solid {accent};
                border-radius: 12px;
            }}
            ToastWidget QLabel {{
                color: #ffffff;
                background: transparent;
                font-size: 11pt;
                font-weight: 700;
            }}
            """
        )

    def show_message(
        self,
        message: str,
        icon: str = "ℹ️",
        duration: int = TOAST_DURATION_DEFAULT,
        position_y: Optional[int] = None,
    ) -> None:
        """토스트 메시지 표시."""
        accent = _accent_for_icon(icon)
        self._apply_style(accent)
        self.icon_label.setText(icon)
        self.message_label.setText(message)
        self.message_label.setStyleSheet(
            "color: #ffffff; background: transparent; font-size: 11pt; font-weight: 700;"
        )
        self.icon_label.setStyleSheet("color: #ffffff; background: transparent;")

        # 긴 메시지에 맞게 높이 조정
        self.adjustSize()
        width = max(360, min(440, self.sizeHint().width()))
        height = max(72, min(140, self.sizeHint().height() + 8))
        self.setFixedSize(width, height)

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
        self._timer.stop()
        self._animation = QPropertyAnimation(self, b"windowOpacity")
        self._animation.setDuration(TOAST_FADE_DURATION)
        self._animation.setStartValue(1.0)
        self._animation.setEndValue(0.0)
        self._animation.setEasingCurve(QEasingCurve.Type.OutQuad)
        self._animation.finished.connect(self._on_fade_finished)
        self._animation.start()

    def _on_fade_finished(self) -> None:
        self.hide()
        self._cleanup()
        self.closed.emit(self)

    def _cleanup(self) -> None:
        if self._timer:
            self._timer.stop()
        if self._animation:
            self._animation.stop()
            self._animation = None


class ToastManager:
    """Toast 알림 관리자 - 스택 기능 지원."""

    MAX_TOASTS = 3
    TOAST_HEIGHT = 80
    TOAST_SPACING = 12

    def __init__(self, parent=None):
        self.parent = parent
        self.toasts: list[ToastWidget] = []

    def show_message(
        self,
        message: str,
        icon: str = "ℹ️",
        duration: int = TOAST_DURATION_DEFAULT,
    ) -> None:
        if not self.parent:
            logger.warning("ToastManager: parent가 없어 메시지를 표시할 수 없습니다")
            return

        try:
            while len(self.toasts) >= self.MAX_TOASTS:
                old_toast = self.toasts.pop(0)
                try:
                    old_toast.hide()
                    old_toast.deleteLater()
                except RuntimeError:
                    pass

            toast = ToastWidget(self.parent)
            toast.closed.connect(self._on_toast_closed)
            self.toasts.append(toast)

            self._update_positions()
            position_y = self._get_position_for_toast(len(self.toasts) - 1)
            toast.show_message(message, icon, duration, position_y)
        except Exception as e:
            logger.error(f"Toast 표시 오류: {e}")

    def _get_position_for_toast(self, index: int) -> int:
        if self.parent:
            base_y = self.parent.y() + self.parent.height() - 20
            return base_y - (index + 1) * (self.TOAST_HEIGHT + self.TOAST_SPACING)
        return 100

    def _update_positions(self) -> None:
        if not self.parent:
            return

        for i, toast in enumerate(self.toasts):
            try:
                if toast.isVisible():
                    x = self.parent.x() + self.parent.width() - toast.width() - 20
                    y = self._get_position_for_toast(i)
                    toast.move(x, y)
            except RuntimeError:
                pass

    def _on_toast_closed(self, toast: ToastWidget) -> None:
        try:
            if toast in self.toasts:
                self.toasts.remove(toast)
                toast.deleteLater()
                self._update_positions()
        except RuntimeError:
            pass

    def clear_all(self) -> None:
        for toast in self.toasts[:]:
            try:
                toast._cleanup()
                toast.hide()
                toast.deleteLater()
            except RuntimeError:
                pass
        self.toasts.clear()
