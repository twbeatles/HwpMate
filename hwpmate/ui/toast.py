from __future__ import annotations

from typing import Optional

from PyQt6.QtCore import QEasingCurve, QEvent, QObject, QPropertyAnimation, QRectF, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QMouseEvent, QPainter, QPainterPath, QPaintEvent
from PyQt6.QtWidgets import QGraphicsOpacityEffect, QHBoxLayout, QLabel, QMainWindow, QWidget

from ..constants import TOAST_DURATION_DEFAULT, TOAST_FADE_DURATION
from ..logging_config import get_logger
from .theme.palette import ThemePalette, palette_for

logger = get_logger(__name__)

# 아이콘별 강조색 (왼쪽 띠·테두리)
_TOAST_ACCENT_BY_ICON: dict[str, str] = {
    "✅": "#22c55e",
    "🎉": "#22c55e",
    "🚀": "#3b82f6",
    "🔁": "#3b82f6",
    "ℹ️": "#3b82f6",
    "⚠️": "#f59e0b",
    "❌": "#ef4444",
    "🛑": "#ef4444",
    "⏭️": "#94a3b8",
}

_SHADOW = 8  # 그림자용 바깥 여백(px)
_RADIUS = 10.0
_ACCENT_BAR = 5.0
_MIN_WIDTH = 300
_MAX_WIDTH = 440
_EDGE_MARGIN = 16


def _accent_for_icon(icon: str) -> str:
    return _TOAST_ACCENT_BY_ICON.get(icon, "#94a3b8")


class ToastWidget(QWidget):
    """메인 창 위에 겹쳐 그리는 불투명 토스트.

    이전 구현은 최상위 반투명 창(WA_TranslucentBackground) + QGraphicsDropShadowEffect +
    QSS rgba 배경 조합이라 Windows 에서 배경이 그려지지 않아 글씨만 떠 보였다.
    이제 부모 창의 자식 오버레이로 띄우고 배경·그림자를 paintEvent 에서 직접 그린다.
    """

    closed = pyqtSignal(object)

    def __init__(self, parent: Optional[QWidget] = None, palette: Optional[ThemePalette] = None):
        super().__init__(parent)
        self._palette = palette or palette_for("dark")
        self._accent = _accent_for_icon("ℹ️")
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, False)
        self.setAutoFillBackground(False)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("클릭하면 닫힙니다")

        self._setup_ui()
        self._animation: Optional[QPropertyAnimation] = None
        self._closing = False
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._fade_out)

        self._opacity = QGraphicsOpacityEffect(self)
        self._opacity.setOpacity(1.0)
        self.setGraphicsEffect(self._opacity)

    def _setup_ui(self) -> None:
        layout = QHBoxLayout(self)
        # 그림자 여백 + 왼쪽 강조 띠 + 내부 여백
        layout.setContentsMargins(
            _SHADOW + int(_ACCENT_BAR) + 12, _SHADOW + 10, _SHADOW + 14, _SHADOW + 10
        )
        layout.setSpacing(10)

        self.icon_label = QLabel("ℹ️")
        self.icon_label.setFixedWidth(26)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        layout.addWidget(self.icon_label)

        self.message_label = QLabel()
        self.message_label.setWordWrap(True)
        self.message_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.message_label.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(self.message_label, stretch=1)

        self._apply_label_colors()

    def set_palette(self, palette: ThemePalette) -> None:
        self._palette = palette
        self._apply_label_colors()
        self.update()

    def _apply_label_colors(self) -> None:
        # 앱 전역 QSS 가 라벨 배경을 칠하지 않도록 명시적으로 투명 처리
        # (전역 QSS 의 font-size 가 setFont 를 덮어쓰므로 크기·굵기도 여기서 지정)
        base = f"color: {self._palette.toast_text}; background: transparent; border: none;"
        self.icon_label.setStyleSheet(base + " font-size: 14pt;")
        self.message_label.setStyleSheet(base + " font-size: 10pt; font-weight: bold;")

    # --- 그리기 -------------------------------------------------------------

    def paintEvent(self, a0: Optional[QPaintEvent]) -> None:
        del a0
        painter = QPainter(self)
        try:
            painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            panel = QRectF(self.rect()).adjusted(_SHADOW, _SHADOW - 2, -_SHADOW, -_SHADOW - 2)

            # 부드러운 그림자: 바깥으로 갈수록 옅어지는 둥근 사각형 여러 겹
            alpha_max = max(0, min(255, self._palette.shadow_alpha))
            steps = _SHADOW
            for i in range(steps, 0, -1):
                alpha = int(alpha_max * (1 - i / (steps + 1)) ** 2 / 3)
                shadow_rect = panel.adjusted(-i, -i + 3, i, i + 3)
                path = QPainterPath()
                path.addRoundedRect(shadow_rect, _RADIUS + i, _RADIUS + i)
                painter.fillPath(path, QColor(0, 0, 0, alpha))

            # 불투명 패널
            panel_path = QPainterPath()
            panel_path.addRoundedRect(panel, _RADIUS, _RADIUS)
            painter.fillPath(panel_path, QColor(self._palette.toast_bg))

            # 왼쪽 강조 띠 (패널 모양으로 클리핑)
            painter.save()
            painter.setClipPath(panel_path)
            painter.fillRect(
                QRectF(panel.left(), panel.top(), _ACCENT_BAR, panel.height()),
                QColor(self._accent),
            )
            painter.restore()

            # 테두리
            border = QColor(self._palette.toast_border)
            painter.setPen(border)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawPath(panel_path)
        finally:
            painter.end()

    # --- 표시/닫기 ----------------------------------------------------------

    def preferred_width(self) -> int:
        parent = self.parentWidget()
        available = (parent.width() - 2 * _EDGE_MARGIN) if parent is not None else _MAX_WIDTH
        return max(_MIN_WIDTH, min(_MAX_WIDTH, available)) + 2 * _SHADOW

    def show_message(
        self,
        message: str,
        icon: str = "ℹ️",
        duration: int = TOAST_DURATION_DEFAULT,
        position_y: Optional[int] = None,
    ) -> None:
        """토스트 메시지 표시. position_y 는 부모 좌표계의 y 위치."""
        self._accent = _accent_for_icon(icon)
        self.icon_label.setText(icon)
        self.message_label.setText(message)

        width = self.preferred_width()
        self.setFixedWidth(width)
        layout = self.layout()
        height = self.heightForWidth(width)
        if height <= 0 and layout is not None:
            height = layout.sizeHint().height()
        self.setFixedHeight(max(56 + 2 * _SHADOW, height))

        parent_widget = self.parentWidget()
        if parent_widget is not None:
            x = parent_widget.width() - self.width() - _EDGE_MARGIN + _SHADOW
            y = position_y if position_y is not None else parent_widget.height() - self.height()
            self.move(max(0, x), max(0, y))

        self._closing = False
        self._opacity.setOpacity(1.0)
        self.show()
        self.raise_()
        self.update()
        self._timer.start(max(500, int(duration)))

    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        del a0
        self._fade_out()

    def _fade_out(self) -> None:
        if self._closing:
            return
        self._closing = True
        try:
            self._timer.stop()
            self._animation = QPropertyAnimation(self._opacity, b"opacity", self)
            self._animation.setDuration(TOAST_FADE_DURATION)
            self._animation.setStartValue(self._opacity.opacity())
            self._animation.setEndValue(0.0)
            self._animation.setEasingCurve(QEasingCurve.Type.OutQuad)
            self._animation.finished.connect(self._on_fade_finished)
            self._animation.start()
        except RuntimeError:
            # 부모 창이 먼저 파괴되는 중이면 애니메이션 없이 닫는다.
            self._on_fade_finished()

    def _on_fade_finished(self) -> None:
        try:
            self.hide()
            self._cleanup()
            self.closed.emit(self)
        except RuntimeError:
            pass

    def _cleanup(self) -> None:
        try:
            if self._timer:
                self._timer.stop()
            if self._animation:
                self._animation.stop()
        except RuntimeError:
            pass
        self._animation = None


class ToastManager(QObject):
    """Toast 알림 관리자 - 메인 창 오른쪽 아래에 스택으로 표시."""

    MAX_TOASTS = 3
    TOAST_SPACING = 10

    def __init__(self, parent=None):
        self.host: Optional[QWidget] = parent if isinstance(parent, QWidget) else None
        self.toasts: list[ToastWidget] = []
        super().__init__(parent)
        if self.host is not None:
            self.host.installEventFilter(self)

    def _palette(self) -> ThemePalette:
        return palette_for(str(getattr(self.host, "current_theme", "dark")))

    def show_message(
        self,
        message: str,
        icon: str = "ℹ️",
        duration: int = TOAST_DURATION_DEFAULT,
    ) -> None:
        if not self.host:
            logger.warning("ToastManager: parent가 없어 메시지를 표시할 수 없습니다")
            return

        try:
            while len(self.toasts) >= self.MAX_TOASTS:
                old_toast = self.toasts.pop(0)
                try:
                    old_toast._cleanup()
                    old_toast.hide()
                    old_toast.deleteLater()
                except RuntimeError:
                    pass

            toast = ToastWidget(self.host, self._palette())
            toast.closed.connect(self._on_toast_closed)
            self.toasts.append(toast)
            toast.show_message(message, icon, duration, self._bottom_limit())
            self._update_positions()
        except Exception as e:
            logger.error(f"Toast 표시 오류: {e}")

    def apply_theme(self) -> None:
        palette = self._palette()
        for toast in self.toasts:
            try:
                toast.set_palette(palette)
            except RuntimeError:
                pass

    def _bottom_limit(self) -> int:
        """토스트를 쌓기 시작할 기준 y (상태바 위)."""
        host = self.host
        if not isinstance(host, QWidget):
            return 100
        bottom = host.height() - _EDGE_MARGIN + _SHADOW
        if isinstance(host, QMainWindow):
            status_bar = host.statusBar()
            if status_bar is not None and status_bar.isVisible():
                bottom -= status_bar.height()
        return bottom

    def _update_positions(self) -> None:
        host = self.host
        if not isinstance(host, QWidget):
            return
        y = self._bottom_limit()
        # 최신 토스트가 가장 아래에 오도록 역순으로 쌓는다.
        for toast in reversed(self.toasts):
            try:
                if not toast.isVisible():
                    continue
                y -= toast.height()
                x = host.width() - toast.width() - _EDGE_MARGIN + _SHADOW
                toast.move(max(0, x), max(0, y))
                toast.raise_()
                # 위젯은 그림자 여백을 포함하므로 패널 사이 간격만 남기고 겹쳐 쌓는다.
                y += 2 * _SHADOW - self.TOAST_SPACING
            except RuntimeError:
                pass

    def eventFilter(self, a0: Optional[QObject], a1: Optional[QEvent]) -> bool:
        host = getattr(self, "host", None)
        if host is not None and a0 is host and a1 is not None and a1.type() == QEvent.Type.Resize:
            self._update_positions()
        return False

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
