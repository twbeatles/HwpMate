from __future__ import annotations

import sys

from PyQt6.QtWidgets import QApplication, QMessageBox, QStyleFactory

from .logging_config import get_logger
from .services.hwp_converter import PYWIN32_AVAILABLE
from .ui.main_window import MainWindow
from .windows_integration import enable_drag_drop_for_admin, is_admin

logger = get_logger(__name__)


def handle_exception(exc_type, exc_value, exc_traceback) -> None:
    """글로벌 예외 핸들러."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical("치명적 오류 발생", exc_info=(exc_type, exc_value, exc_traceback))

    try:
        if QApplication.instance():
            QMessageBox.critical(
                None,
                "치명적 오류",
                f"프로그램에서 예기치 않은 오류가 발생했습니다.\n\n"
                f"오류: {exc_type.__name__}: {exc_value}\n\n"
                f"프로그램을 다시 시작해 주세요.",
            )
    except Exception:
        pass


def main() -> None:
    """메인 함수."""
    sys.excepthook = handle_exception

    if not PYWIN32_AVAILABLE:
        app = QApplication(sys.argv)
        QMessageBox.critical(
            None, "오류",
            "pywin32 라이브러리가 필요합니다.\n\npip install pywin32"
        )
        del app
        return

    if not is_admin():
        app = QApplication(sys.argv)
        QMessageBox.warning(
            None,
            "관리자 권한 필요",
            "이 프로그램은 관리자 권한으로 실행해야 합니다.\n\n"
            "파일을 마우스 오른쪽 버튼으로 클릭하여\n"
            "'관리자 권한으로 실행'을 선택하세요."
        )
        del app
        sys.exit(1)

    try:
        enable_drag_drop_for_admin()

        app = QApplication(sys.argv)
        app.setStyle(QStyleFactory.create("Fusion"))

        window = MainWindow()
        window.show()

        sys.exit(app.exec())
    except Exception as e:
        logger.critical(f"애플리케이션 실행 오류: {e}")
        raise
