from __future__ import annotations

from typing import Any, Callable

from PyQt6.QtGui import QAction, QCloseEvent, QKeySequence, QShortcut
from PyQt6.QtWidgets import QLabel, QMenu, QMessageBox, QStyle, QSystemTrayIcon

from ...constants import FORMAT_GROUPS, FORMAT_TYPES, VERSION, WORKER_WAIT_TIMEOUT
from ...logging_config import get_logger
from ...models import AppConfig
from .state import MainWindowState

logger = get_logger(__name__)


class LifecycleController:
    """Menu, tray, shortcuts, settings persistence, and close handling."""

    def __init__(
        self,
        window: Any,
        state: MainWindowState,
        save_config_func: Callable[[AppConfig], None],
    ) -> None:
        self.window = window
        self.state = state
        self._save_config = save_config_func

    def init_menu_bar(self) -> None:
        menubar = self.window.menuBar()
        assert menubar is not None

        file_menu = menubar.addMenu("파일(&F)")
        assert file_menu is not None

        add_files_action = QAction("파일 추가(&A)", self.window)
        add_files_action.setShortcut("Ctrl+O")
        add_files_action.triggered.connect(self.window._browse_files)
        file_menu.addAction(add_files_action)

        add_folder_action = QAction("폴더 선택(&F)", self.window)
        add_folder_action.setShortcut("Ctrl+Shift+O")
        add_folder_action.triggered.connect(self.window._select_folder)
        file_menu.addAction(add_folder_action)

        file_menu.addSeparator()

        exit_action = QAction("종료(&X)", self.window)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.window.close)
        file_menu.addAction(exit_action)

        edit_menu = menubar.addMenu("편집(&E)")
        assert edit_menu is not None

        remove_selected_action = QAction("선택 파일 제거(&R)", self.window)
        remove_selected_action.setShortcut("Delete")
        remove_selected_action.triggered.connect(self.window._remove_selected)
        edit_menu.addAction(remove_selected_action)

        clear_all_action = QAction("전체 제거(&C)", self.window)
        clear_all_action.setShortcut("Ctrl+Delete")
        clear_all_action.triggered.connect(self.window._clear_all)
        edit_menu.addAction(clear_all_action)

        help_menu = menubar.addMenu("도움말(&H)")
        assert help_menu is not None

        usage_action = QAction("사용법(&U)", self.window)
        usage_action.triggered.connect(self.show_usage)
        help_menu.addAction(usage_action)

        help_menu.addSeparator()

        about_action = QAction("프로그램 정보(&A)", self.window)
        about_action.setShortcut("F1")
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def init_status_bar(self) -> None:
        status_bar = self.window.statusBar()
        assert status_bar is not None
        self.window.status_bar = status_bar

        self.window.version_label = QLabel(f"v{VERSION}")
        status_bar.addPermanentWidget(self.window.version_label)

        self.window.hwp_status_label = QLabel("🔵 한글 대기중")
        status_bar.addPermanentWidget(self.window.hwp_status_label)

        self.window.file_count_label = QLabel("📄 파일: 0개")
        status_bar.addPermanentWidget(self.window.file_count_label)

    def init_shortcuts(self) -> None:
        start_shortcut = QShortcut(QKeySequence("Ctrl+Return"), self.window)
        start_shortcut.activated.connect(self.window._start_conversion)

        cancel_shortcut = QShortcut(QKeySequence("Escape"), self.window)
        cancel_shortcut.activated.connect(self.cancel_conversion_if_running)

    def init_tray_icon(self) -> None:
        self.window.tray_icon = QSystemTrayIcon(self.window)

        style = self.window.style()
        assert style is not None
        self.window.tray_icon.setIcon(style.standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView))
        self.window.tray_icon.setToolTip(f"HWP 변환기 v{VERSION}")

        tray_menu = QMenu()

        show_action = QAction("열기", self.window)
        show_action.triggered.connect(self.show_from_tray)
        tray_menu.addAction(show_action)

        tray_menu.addSeparator()

        quit_action = QAction("종료", self.window)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)

        self.window.tray_icon.setContextMenu(tray_menu)
        self.window.tray_icon.activated.connect(self.on_tray_activated)
        self.window.tray_icon.show()

    def show_from_tray(self) -> None:
        self.window.showNormal()
        self.window.activateWindow()
        self.window.raise_()

    def quit_app(self) -> None:
        self.window.close()

    def on_tray_activated(self, reason: object) -> None:
        try:
            if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
                self.show_from_tray()
        except Exception as e:
            logger.debug(f"트레이 아이콘 이벤트 처리 오류: {e}")

    def cancel_conversion_if_running(self) -> None:
        if self.state.is_converting:
            self.window._cancel_conversion()

    def show_usage(self) -> None:
        format_html = "".join(
            f"<li><b>{group}</b>: {', '.join(f'{key} ({FORMAT_TYPES[key].desc})' for key in keys)}</li>"
            for group, keys in FORMAT_GROUPS.items()
        )
        usage_text = f"""<h3>HWP 변환기 사용법</h3>

<p><b>1. 변환 모드 선택</b></p>
<ul>
<li>폴더 일괄 변환: 폴더 내 실제 변환 가능한 파일만 미리보기 후 변환</li>
<li>파일 개별 선택: 원하는 파일만 선택하여 변환</li>
</ul>

<p><b>2. 변환 형식 선택</b></p>
<ul>
{format_html}
</ul>
<p>이미지 변환은 한글 설치 버전에 따라 저장 방식이 다를 수 있으며, 기본 출력 파일 생성 여부와 파일 크기로 성공을 판단합니다.</p>

<p><b>3. 단축키</b></p>
<ul>
<li>Ctrl+O: 파일 추가</li>
<li>Ctrl+Shift+O: 폴더 선택</li>
<li>Ctrl+Enter: 변환 시작</li>
<li>Esc: 변환 취소</li>
<li>Delete: 선택 파일 제거</li>
</ul>
"""
        QMessageBox.information(self.window, "사용법", usage_text)

    def show_about(self) -> None:
        supported_formats = "".join(
            f"<li><b>{group}</b>: {', '.join(keys)}</li>"
            for group, keys in FORMAT_GROUPS.items()
        )
        about_text = f"""<h2>HWP 변환기 v{VERSION}</h2>
<p>HWP/HWPX 파일을 다양한 문서/이미지 형식으로 변환하는 프로그램</p>

<p><b>주요 기능:</b></p>
<ul>
<li>폴더 일괄 변환 / 파일 개별 선택</li>
<li>모드별 드래그 앤 드롭 지원</li>
<li>다크/라이트 테마</li>
<li>사전 점검, 결과 리포트, 실패 목록 저장</li>
<li>원본 백업 옵션과 실패 자동 재시도</li>
</ul>

<p><b>지원 형식:</b></p>
<ul>
{supported_formats}
</ul>
<p>이미지 변환 결과는 한글 설치 버전에 따라 차이가 있을 수 있으며, 앱은 기본 출력 파일의 존재와 0바이트 초과 크기를 성공 기준으로 사용합니다.</p>

<p><b>요구사항:</b></p>
<ul>
<li>Windows 10/11</li>
<li>한컴오피스 한글 2018 이상</li>
<li>관리자 권한</li>
</ul>

<p>© 2024-2025</p>
"""
        QMessageBox.about(self.window, "프로그램 정보", about_text)

    def save_settings(self) -> None:
        self.window.config["mode"] = "folder" if self.window.folder_radio.isChecked() else "files"
        self.window.config["format"] = self.state.selected_format

        self.window.config["include_sub"] = self.window.include_sub_check.isChecked()
        self.window.config["same_location"] = self.window.same_location_check.isChecked()
        self.window.config["overwrite"] = self.window.overwrite_check.isChecked()
        self.window.config["backup_enabled"] = self.window.backup_check.isChecked()
        self.window.config["retry_count"] = self.window.retry_spin.value()

        self.window.config["folder_path"] = self.window.folder_entry.text().strip()
        self.window.config["output_path"] = self.window.output_entry.text().strip()
        if self.window.folder_entry.text().strip():
            self.window.config["last_folder"] = self.window.folder_entry.text().strip()
        if self.window.output_entry.text().strip():
            self.window.config["last_output"] = self.window.output_entry.text().strip()

        self._save_config(self.window.config)

    def close_event(self, event: QCloseEvent) -> None:
        logger.info("메인 윈도우 종료 이벤트 수신")
        if not self.window._cancel_active_scan(wait_ms=WORKER_WAIT_TIMEOUT):
            self.window.status_label.setText("파일 스캔 종료 대기 중...")
            event.ignore()
            return

        if self.state.is_converting:
            if self.state.close_after_worker:
                self.window.status_label.setText("종료 대기 중...")
                event.ignore()
                return

            reply = QMessageBox.question(
                self.window,
                "확인",
                "변환 작업이 진행 중입니다. 종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return

            self.state.close_after_worker = True
            if not self.window._request_worker_stop("종료 대기 중..."):
                if self.state.worker and self.state.worker.can_force_terminate():
                    reply2 = QMessageBox.question(
                        self.window,
                        "강제 종료 경고",
                        "앱이 소유한 한글 프로세스만 강제 종료합니다.\n열려 있는 문서가 저장되지 않을 수 있습니다.\n\n계속할까요?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    )
                    if reply2 == QMessageBox.StandardButton.Yes:
                        self.window._perform_force_terminate()

                if self.state.worker and self.state.worker.isRunning():
                    self.window.status_label.setText("종료 대기 중...")
                    event.ignore()
                    return

        if hasattr(self.window, "toast") and self.window.toast:
            self.window.toast.clear_all()

        if hasattr(self.window, "tray_icon"):
            self.window.tray_icon.hide()

        self.save_settings()
        logger.info("메인 윈도우 종료 허용")
        event.accept()
