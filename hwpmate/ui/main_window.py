from __future__ import annotations

import ctypes
import logging
import platform
import sys
import time
from pathlib import Path
from typing import Iterable, List, Optional

from PyQt6.QtCore import QSignalBlocker, QTimer
from PyQt6.QtGui import QAction, QCloseEvent, QKeySequence, QShortcut, QShowEvent
from PyQt6.QtWidgets import QApplication, QFileDialog, QMainWindow, QLabel, QMenu, QMessageBox, QStyle, QSystemTrayIcon, QTableWidgetItem, QCheckBox, QDialog, QLineEdit, QPushButton, QProgressBar, QRadioButton, QTabWidget, QTableWidget, QWidget

from ..config_repository import load_config, save_config
from ..constants import FEEDBACK_RESET_DELAY, FORMAT_GROUPS, FORMAT_TYPES, SCAN_BATCH_SIZE, SCAN_CANCEL_WAIT_MS, SUPPORTED_EXTENSIONS, WORKER_WAIT_TIMEOUT, VERSION
from ..logging_config import get_logger
from ..models import ConversionSummary, ConversionTask, PlannedConversion
from ..path_utils import canonicalize_path, check_write_permission, is_valid_path_name
from ..services.file_selection_store import FileSelectionStore
from ..services.task_planner import TaskPlanner
from ..windows_integration import NativeDropFilter, get_native_admin_drag_drop_policy
from ..workers.conversion_worker import ConversionWorker
from ..workers.file_scan_worker import FileScanWorker
from .dialogs import PreflightDialog, ResultDialog
from .main_window_ui import MainWindowWidgets, build_main_window_ui
from .theme import ThemeManager
from .toast import ToastManager
from .widgets import DropArea, FormatCard

logger = get_logger(__name__)

class MainWindow(QMainWindow):
    """메인 윈도우"""

    ui: MainWindowWidgets
    theme_btn: QPushButton
    folder_radio: QRadioButton
    files_radio: QRadioButton
    folder_widget: QWidget
    folder_entry: QLineEdit
    folder_btn: QPushButton
    include_sub_check: QCheckBox
    files_widget: QWidget
    drop_area: DropArea
    add_btn: QPushButton
    remove_btn: QPushButton
    clear_btn: QPushButton
    file_table: QTableWidget
    same_location_check: QCheckBox
    output_entry: QLineEdit
    output_btn: QPushButton
    format_tabs: QTabWidget
    format_cards: dict[str, FormatCard]
    overwrite_check: QCheckBox
    start_btn: QPushButton
    cancel_btn: QPushButton
    status_label: QLabel
    progress_bar: QProgressBar
    progress_label: QLabel
    
    def __init__(self) -> None:
        super().__init__()
        
        # 설정 로드
        self.config = load_config()
        self.current_theme = self.config.get("theme", "dark")
        
        # 변수 초기화
        self.tasks: List[ConversionTask] = []
        self.plan: Optional[PlannedConversion] = None
        self.last_summary: Optional[ConversionSummary] = None
        self.worker: Optional[ConversionWorker] = None
        self.is_converting = False
        self.file_store = FileSelectionStore()
        self.file_list = self.file_store.paths  # 기존 메서드 호환용 뷰
        self._file_set = self.file_store.path_keys  # 기존 메서드 호환용 뷰
        self.task_planner = TaskPlanner()
        self.conversion_start_time: Optional[float] = None
        self.file_scan_worker: Optional[FileScanWorker] = None
        self._scan_mode: Optional[str] = None
        self._scan_new_file_count = 0
        self._scan_preview_count = 0
        self._scan_started_at = None
        self._force_kill_pending = False
        self._close_after_worker = False
        
        # 드래그 앤 드롭 초기화 플래그
        self._drag_drop_initialized = False
        
        # UI 초기화
        self._init_menu_bar()
        self._init_ui()
        self._init_status_bar()
        self._init_shortcuts()
        self._init_tray_icon()
        self._apply_theme()
        self._update_mode_ui()
        self._update_output_ui()
        
        # Toast 관리자 초기화 (스택 지원)
        self.toast = ToastManager(self)
        
        logger.info(f"HWP 변환기 v{VERSION} 시작")
        logger.info(f"시스템 정보: {platform.system()} {platform.release()} ({platform.version()})")
        logger.info(f"Python 버전: {sys.version}")
    
    def showEvent(self, a0: Optional[QShowEvent]) -> None:
        """윈도우 표시 이벤트 - 네이티브 드래그 앤 드롭 활성화"""
        if a0 is None:
            return
        super().showEvent(a0)
        
        # 처음 표시될 때만 실행
        if not self._drag_drop_initialized:
            self._drag_drop_initialized = True
            
            try:
                native_dnd_enabled, native_dnd_reason = get_native_admin_drag_drop_policy()
                if not native_dnd_enabled:
                    logger.warning(f"네이티브 드래그 앤 드롭 초기화 건너뜀: {native_dnd_reason}")
                    return

                # 네이티브 드래그 앤 드롭 필터 설정
                drop_filter = NativeDropFilter.get_instance()
                
                # 메인 윈도우 핸들 가져오기
                main_hwnd = int(self.winId())
                drop_filter.register_window(main_hwnd)
                
                # 모든 자식 윈도우에도 등록 (Qt는 여러 계층의 윈도우를 생성함)
                try:
                    user32 = ctypes.windll.user32
                    
                    # 자식 윈도우 열거를 위한 콜백
                    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
                    
                    def enum_callback(child_hwnd, lParam):
                        try:
                            drop_filter.register_window(child_hwnd)
                        except Exception:
                            pass
                        return True  # 계속 열거
                    
                    callback = WNDENUMPROC(enum_callback)
                    user32.EnumChildWindows(main_hwnd, callback, 0)
                    logger.debug("자식 윈도우 드래그 앤 드롭 등록 완료")
                except Exception as e:
                    logger.debug(f"자식 윈도우 열거 실패 (무시): {e}")
                
                # 파일 드롭 콜백 설정
                drop_filter.files_dropped_callback = self._on_native_files_dropped
                
                # 애플리케이션에 네이티브 이벤트 필터 설치
                app = QApplication.instance()
                if app:
                    app.installNativeEventFilter(drop_filter)
                    logger.info("네이티브 이벤트 필터 설치 완료")
                
                logger.info("네이티브 드래그 앤 드롭 초기화 완료")
            except Exception as e:
                logger.warning(f"네이티브 드래그 앤 드롭 초기화 중 오류: {e}")
                import traceback
                traceback.print_exc()
    
    def _on_native_files_dropped(self, files: List[str]) -> None:
        """네이티브 드래그 앤 드롭 입력 처리 (파일/폴더 경로)"""
        if not files:
            return

        normalized = [canonicalize_path(path) for path in files if str(path).strip()]
        if not normalized:
            return

        if self.folder_radio.isChecked():
            if len(normalized) == 1 and Path(normalized[0]).is_dir():
                folder = normalized[0]
                self.folder_entry.setText(folder)
                self.config["last_folder"] = folder
                self.config["folder_path"] = folder
                self._start_folder_preview_scan(folder)
                if hasattr(self, "toast"):
                    self.toast.show_message("📁 폴더 드롭을 받아 미리보기 스캔을 시작합니다", "✅")
                return

            QMessageBox.warning(
                self,
                "경고",
                "폴더 모드에서는 폴더 1개만 드롭할 수 있습니다.\n파일이나 다중 경로 드롭은 지원하지 않습니다.",
            )
            return

        self._add_files(normalized)
        if hasattr(self, "drop_area") and self.drop_area:
            self.drop_area.icon_label.setText("✅")
            self.drop_area.text_label.setText(f"{len(normalized)}개 경로 스캔 시작")
            QTimer.singleShot(FEEDBACK_RESET_DELAY, self.drop_area._reset_appearance)
        if hasattr(self, "toast"):
            self.toast.show_message(f"📂 {len(normalized)}개 경로를 스캔합니다", "✅")
    
    def _init_menu_bar(self) -> None:
        """메뉴바 초기화"""
        menubar = self.menuBar()
        assert menubar is not None
        
        # 파일 메뉴
        file_menu = menubar.addMenu("파일(&F)")
        assert file_menu is not None
        
        add_files_action = QAction("파일 추가(&A)", self)
        add_files_action.setShortcut("Ctrl+O")
        add_files_action.triggered.connect(self._browse_files)
        file_menu.addAction(add_files_action)
        
        add_folder_action = QAction("폴더 선택(&F)", self)
        add_folder_action.setShortcut("Ctrl+Shift+O")
        add_folder_action.triggered.connect(self._select_folder)
        file_menu.addAction(add_folder_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("종료(&X)", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 편집 메뉴
        edit_menu = menubar.addMenu("편집(&E)")
        assert edit_menu is not None
        
        remove_selected_action = QAction("선택 파일 제거(&R)", self)
        remove_selected_action.setShortcut("Delete")
        remove_selected_action.triggered.connect(self._remove_selected)
        edit_menu.addAction(remove_selected_action)
        
        clear_all_action = QAction("전체 제거(&C)", self)
        clear_all_action.setShortcut("Ctrl+Delete")
        clear_all_action.triggered.connect(self._clear_all)
        edit_menu.addAction(clear_all_action)
        
        # 도움말 메뉴
        help_menu = menubar.addMenu("도움말(&H)")
        assert help_menu is not None
        
        usage_action = QAction("사용법(&U)", self)
        usage_action.triggered.connect(self._show_usage)
        help_menu.addAction(usage_action)
        
        help_menu.addSeparator()
        
        about_action = QAction("프로그램 정보(&A)", self)
        about_action.setShortcut("F1")
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)
    
    def _init_status_bar(self) -> None:
        """상태바 초기화"""
        status_bar = self.statusBar()
        assert status_bar is not None
        self.status_bar = status_bar
        
        # 버전 정보
        self.version_label = QLabel(f"v{VERSION}")
        status_bar.addPermanentWidget(self.version_label)
        
        # 한글 연결 상태
        self.hwp_status_label = QLabel("🔵 한글 대기중")
        status_bar.addPermanentWidget(self.hwp_status_label)
        
        # 파일 수
        self.file_count_label = QLabel("📄 파일: 0개")
        status_bar.addPermanentWidget(self.file_count_label)
    
    def _init_shortcuts(self) -> None:
        """키보드 단축키 초기화"""
        # Ctrl+Enter: 변환 시작
        start_shortcut = QShortcut(QKeySequence("Ctrl+Return"), self)
        start_shortcut.activated.connect(self._start_conversion)
        
        # Esc: 변환 취소
        cancel_shortcut = QShortcut(QKeySequence("Escape"), self)
        cancel_shortcut.activated.connect(self._cancel_conversion_if_running)
    
    def _init_tray_icon(self) -> None:
        """시스템 트레이 아이콘 초기화"""
        self.tray_icon = QSystemTrayIcon(self)
        
        # 기본 아이콘 설정 (앱 아이콘 또는 기본)
        style = self.style()
        assert style is not None
        self.tray_icon.setIcon(style.standardIcon(QStyle.StandardPixmap.SP_FileDialogContentsView))
        self.tray_icon.setToolTip(f"HWP 변환기 v{VERSION}")
        
        # 트레이 메뉴
        tray_menu = QMenu()
        
        show_action = QAction("열기", self)
        show_action.triggered.connect(self._show_from_tray)
        tray_menu.addAction(show_action)
        
        tray_menu.addSeparator()
        
        quit_action = QAction("종료", self)
        quit_action.triggered.connect(self._quit_app)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()  # 시스템 트레이에 아이콘 표시
    
    def _show_from_tray(self) -> None:
        """트레이에서 창 복원"""
        self.showNormal()
        self.activateWindow()
        self.raise_()
    
    def _quit_app(self) -> None:
        """애플리케이션 종료"""
        self.close()
    
    def _on_tray_activated(self, reason) -> None:
        """트레이 아이콘 클릭 이벤트"""
        try:
            if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
                self._show_from_tray()
        except Exception as e:
            logger.debug(f"트레이 아이콘 이벤트 처리 오류: {e}")
    
    def _cancel_conversion_if_running(self) -> None:
        """변환 중일 때만 취소"""
        if self.is_converting:
            self._cancel_conversion()
    
    def _show_usage(self) -> None:
        """사용법 표시"""
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

<p><b>3. 단축키</b></p>
<ul>
<li>Ctrl+O: 파일 추가</li>
<li>Ctrl+Shift+O: 폴더 선택</li>
<li>Ctrl+Enter: 변환 시작</li>
<li>Esc: 변환 취소</li>
<li>Delete: 선택 파일 제거</li>
</ul>
"""
        QMessageBox.information(self, "사용법", usage_text)
    
    def _show_about(self) -> None:
        """프로그램 정보 표시"""
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
</ul>

<p><b>지원 형식:</b></p>
<ul>
{supported_formats}
</ul>

<p><b>요구사항:</b></p>
<ul>
<li>Windows 10/11</li>
<li>한컴오피스 한글 2018 이상</li>
<li>관리자 권한</li>
</ul>

<p>© 2024-2025</p>
"""
        QMessageBox.about(self, "프로그램 정보", about_text)
    
    def _init_ui(self) -> None:
        """UI 초기화"""
        self.ui = build_main_window_ui(self, self.config)
    
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
    
    def _on_format_card_clicked(self, format_type: str) -> None:
        """포맷 카드 클릭 이벤트"""
        self._selected_format = format_type
        self._update_format_cards()
        if self.folder_radio.isChecked() and self.folder_entry.text().strip():
            self._start_folder_preview_scan(self.folder_entry.text().strip())
    
    def _update_format_cards(self) -> None:
        """포맷 카드 선택 상태 업데이트"""
        for fmt_key, card in self.format_cards.items():
            card.setSelected(self._selected_format == fmt_key)
    
    def _update_mode_ui(self) -> None:
        """모드에 따라 UI 업데이트"""
        self._cancel_active_scan()
        is_folder_mode = self.folder_radio.isChecked()
        self.folder_widget.setVisible(is_folder_mode)
        self.files_widget.setVisible(not is_folder_mode)
    
    def _update_output_ui(self) -> None:
        """출력 폴더 UI 상태 업데이트"""
        same_location = self.same_location_check.isChecked()
        self.output_entry.setEnabled(not same_location)
        self.output_btn.setEnabled(not same_location)

    def _on_include_sub_toggled(self, _: bool) -> None:
        """하위 폴더 옵션 변경 시 폴더 미리보기 재스캔"""
        if self.folder_radio.isChecked() and self.folder_entry.text().strip():
            self._start_folder_preview_scan(self.folder_entry.text().strip())

    def _cancel_active_scan(self, wait_ms: int = SCAN_CANCEL_WAIT_MS) -> bool:
        """진행 중인 파일 스캔이 있으면 취소"""
        worker = self.file_scan_worker
        if not worker:
            return True

        if worker.isRunning():
            worker.cancel()
            worker.wait(wait_ms)

        if worker.isRunning():
            return False

        try:
            worker.batch_found.disconnect(self._on_scan_batch_found)
            worker.scan_progress.disconnect(self._on_scan_progress)
            worker.scan_finished.disconnect(self._on_scan_finished)
            worker.scan_error.disconnect(self._on_scan_error)
            worker.finished.disconnect(self._on_scan_worker_finished)
        except (TypeError, RuntimeError):
            pass

        worker.deleteLater()
        self.file_scan_worker = None
        self._scan_mode = None
        self._scan_started_at = None
        self._scan_new_file_count = 0
        self._scan_preview_count = 0
        return True

    def _start_scan(
        self,
        input_paths: List[str],
        mode: str,
        include_sub: bool = True,
        allowed_exts: Optional[Iterable[str]] = None,
    ) -> None:
        """비동기 파일 스캔 시작"""
        cleaned_inputs = [str(p).strip() for p in input_paths if str(p).strip()]
        if not cleaned_inputs:
            return

        if not self._cancel_active_scan():
            logger.warning("이전 파일 스캔이 아직 종료되지 않아 새 스캔을 시작하지 않습니다.")
            return

        self._scan_mode = mode
        self._scan_new_file_count = 0
        self._scan_preview_count = 0
        self._scan_started_at = time.perf_counter()

        self.file_scan_worker = FileScanWorker(
            cleaned_inputs,
            include_sub=include_sub,
            allowed_exts=allowed_exts or SUPPORTED_EXTENSIONS,
            batch_size=SCAN_BATCH_SIZE,
        )
        self.file_scan_worker.batch_found.connect(self._on_scan_batch_found)
        self.file_scan_worker.scan_progress.connect(self._on_scan_progress)
        self.file_scan_worker.scan_finished.connect(self._on_scan_finished)
        self.file_scan_worker.scan_error.connect(self._on_scan_error)
        self.file_scan_worker.finished.connect(self._on_scan_worker_finished)
        self.file_scan_worker.start()

    def _start_folder_preview_scan(self, folder_path: str) -> None:
        """폴더 모드 파일 수 미리보기 스캔 시작"""
        self.status_label.setText("📂 폴더 스캔 중...")
        self._start_scan(
            [folder_path],
            mode="folder_preview",
            include_sub=self.include_sub_check.isChecked(),
            allowed_exts=set(self.task_planner.preview_allowed_extensions(self._selected_format)),
        )

    def _append_files_batch(self, files: List[str]) -> int:
        """파일 목록을 배치로 렌더링"""
        if not files:
            return 0

        unique_files = self.file_store.add_paths(files)
        if not unique_files:
            return 0

        render_start = time.perf_counter()
        start_row = self.file_table.rowCount()
        end_row = start_row + len(unique_files)

        self.file_table.setUpdatesEnabled(False)
        blocker = QSignalBlocker(self.file_table)
        try:
            self.file_table.setRowCount(end_row)
            for row_idx, file_path in enumerate(unique_files, start=start_row):
                file_obj = Path(file_path)
                self.file_table.setItem(row_idx, 0, QTableWidgetItem(file_obj.name))
                self.file_table.setItem(row_idx, 1, QTableWidgetItem(str(file_obj.parent)))
        finally:
            del blocker
            self.file_table.setUpdatesEnabled(True)

        self._update_file_count()

        if logger.isEnabledFor(logging.DEBUG):
            elapsed = time.perf_counter() - render_start
            logger.debug(f"파일 목록 렌더링: batch={len(unique_files)}, 소요={elapsed:.4f}s")
        return len(unique_files)

    def _on_scan_batch_found(self, batch: list) -> None:
        """비동기 스캔 배치 결과 처리"""
        if self.sender() is not self.file_scan_worker:
            return

        if self._scan_mode == "add_files":
            added = self._append_files_batch(batch)
            self._scan_new_file_count += added
            return

        if self._scan_mode == "folder_preview":
            self._scan_preview_count += len(batch)

    def _on_scan_progress(self, current: int, total: int) -> None:
        """비동기 스캔 진행률 처리"""
        if self.sender() is not self.file_scan_worker:
            return

        if self._scan_mode == "add_files":
            self.status_label.setText(
                f"📥 파일 스캔 중... {current}/{total} 경로 처리 (신규 {self._scan_new_file_count}개)"
            )
            return

        if self._scan_mode == "folder_preview":
            self.status_label.setText(
                f"📂 폴더 스캔 중... {current}/{total} 경로 처리 ({self._scan_preview_count}개 발견)"
            )

    def _on_scan_finished(self, total_found: int, canceled: bool) -> None:
        """비동기 스캔 완료 처리"""
        if self.sender() is not self.file_scan_worker:
            return

        elapsed = 0.0
        if self._scan_started_at is not None:
            elapsed = time.perf_counter() - self._scan_started_at

        if self._scan_mode == "add_files":
            if canceled:
                self.status_label.setText("파일 스캔이 취소되었습니다")
            elif self._scan_new_file_count == 0:
                self.status_label.setText("추가할 새 파일이 없습니다")
            else:
                self.status_label.setText(
                    f"{self._scan_new_file_count}개 파일 추가됨 (총 {len(self.file_list)}개)"
                )
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"파일 추가 스캔 완료: 발견={total_found}, 신규={self._scan_new_file_count}, "
                    f"취소={canceled}, 소요={elapsed:.3f}s"
                )
            return

        if self._scan_mode == "folder_preview":
            if canceled:
                self.status_label.setText("폴더 스캔이 취소되었습니다")
            elif self._scan_preview_count == 0:
                self.status_label.setText("⚠️ 현재 포맷으로 변환 가능한 파일이 없습니다")
            else:
                self.status_label.setText(f"📁 {self._scan_preview_count}개 변환 가능 파일 발견")
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"폴더 미리보기 스캔 완료: 발견={self._scan_preview_count}, "
                    f"취소={canceled}, 소요={elapsed:.3f}s"
                )

    def _on_scan_error(self, error_msg: str) -> None:
        """비동기 스캔 오류 처리"""
        if self.sender() is not self.file_scan_worker:
            return
        logger.error(f"파일 스캔 오류: {error_msg}")
        self.status_label.setText("파일 스캔 중 오류가 발생했습니다")

    def _on_scan_worker_finished(self) -> None:
        """스캔 워커 종료 처리"""
        worker = self.file_scan_worker
        if self.sender() is not worker or worker is None:
            return
        worker.deleteLater()
        self.file_scan_worker = None
        self._scan_mode = None
        self._scan_started_at = None
        self._scan_new_file_count = 0
        self._scan_preview_count = 0
    
    def _select_folder(self) -> None:
        """폴더 선택"""
        initial = self.config.get("last_folder", "")
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", initial)
        if folder:
            self.folder_entry.setText(folder)
            self.config["last_folder"] = folder
            self._start_folder_preview_scan(folder)
    
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
        """파일/폴더 입력을 비동기로 스캔해 파일 목록에 추가"""
        if not files:
            return

        requested = [canonicalize_path(p) for p in files if str(p).strip()]
        if not requested:
            return

        scan_enqueue_start = time.perf_counter()
        self.status_label.setText(f"📥 {len(requested)}개 경로 스캔 시작...")
        self._start_scan(
            requested,
            mode="add_files",
            include_sub=True,
            allowed_exts=set(SUPPORTED_EXTENSIONS),
        )
        if logger.isEnabledFor(logging.DEBUG):
            elapsed = time.perf_counter() - scan_enqueue_start
            logger.debug(f"파일 스캔 요청 등록: 입력={len(requested)}, 소요={elapsed:.4f}s")
    
    def _remove_selected(self) -> None:
        """선택된 파일 제거"""
        selected = self.file_table.selectedItems()
        if not selected:
            # 선택된 항목이 없으면 조용히 반환 (단축키 사용 시 불필요한 팝업 방지)
            return
        
        rows = set(item.row() for item in selected)
        self.file_store.remove_rows(rows)
        for row in sorted(rows, reverse=True):
            self.file_table.removeRow(row)
        
        self.status_label.setText(f"선택 파일 제거됨 (총 {len(self.file_list)}개)")
        self._update_file_count()
    
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
            self.file_store.clear()
            self.file_table.setRowCount(0)
            self.status_label.setText("모든 파일 제거됨")
            self._update_file_count()
    
    def _update_file_count(self) -> None:
        """상태바 파일 수 업데이트"""
        count = self.file_store.count
        self.file_count_label.setText(f"📄 파일: {count}개")
    
    def _collect_tasks(self) -> PlannedConversion:
        """변환 작업 목록 생성"""
        return self.task_planner.build_tasks(
            is_folder_mode=self.folder_radio.isChecked(),
            format_type=self._selected_format,
            folder_path=self.folder_entry.text(),
            include_sub=self.include_sub_check.isChecked(),
            same_location=self.same_location_check.isChecked(),
            output_path=self.output_entry.text(),
            file_paths=self.file_store.paths,
        )

    def _adjust_output_paths(self, plan: PlannedConversion) -> int:
        """출력 경로 조정 (덮어쓰기 방지)"""
        return self.task_planner.resolve_output_conflicts(plan.tasks, overwrite=False)

    def _save_settings(self) -> None:
        """설정 저장"""
        self.config["mode"] = "folder" if self.folder_radio.isChecked() else "files"
        self.config["format"] = self._selected_format
        
        self.config["include_sub"] = self.include_sub_check.isChecked()
        self.config["same_location"] = self.same_location_check.isChecked()
        self.config["overwrite"] = self.overwrite_check.isChecked()

        self.config["folder_path"] = self.folder_entry.text().strip()
        self.config["output_path"] = self.output_entry.text().strip()
        if self.folder_entry.text().strip():
            self.config["last_folder"] = self.folder_entry.text().strip()
        if self.output_entry.text().strip():
            self.config["last_output"] = self.output_entry.text().strip()

        save_config(self.config)

    def _validate_output_settings(self) -> None:
        if self.same_location_check.isChecked():
            return

        output_path = self.output_entry.text().strip()
        if not output_path:
            raise ValueError("출력 폴더를 선택하세요.")
        if not is_valid_path_name(output_path):
            raise ValueError(f"출력 경로에 사용할 수 없는 문자가 있습니다:\n{output_path}")

        output_folder = Path(output_path)
        if not output_folder.exists():
            raise ValueError(f"출력 폴더가 존재하지 않습니다:\n{output_folder}")
        if not check_write_permission(output_folder):
            raise ValueError(f"출력 폴더에 쓰기 권한이 없습니다:\n{output_folder}")
    
    def _start_conversion(self) -> None:
        """변환 시작"""
        try:
            if self.file_scan_worker and self.file_scan_worker.isRunning():
                if self._scan_mode == "add_files":
                    QMessageBox.warning(self, "경고", "파일 스캔이 진행 중입니다. 스캔 완료 후 다시 시도하세요.")
                    return
                if not self._cancel_active_scan():
                    QMessageBox.warning(self, "경고", "폴더 스캔이 아직 종료되지 않았습니다. 잠시 후 다시 시도하세요.")
                    return

            self._validate_output_settings()

            task_collect_start = time.perf_counter()
            plan = self._collect_tasks()
            if logger.isEnabledFor(logging.DEBUG):
                logger.debug(
                    f"작업 목록 생성 완료: 실행={plan.runnable_count}개, 건너뜀={plan.skipped_count}개, "
                    f"소요={time.perf_counter() - task_collect_start:.3f}s"
                )

            if not self.overwrite_check.isChecked():
                plan.conflict_renamed_count = self._adjust_output_paths(plan)
                if plan.conflict_renamed_count:
                    plan.warnings.append(
                        f"출력 경로 충돌 {plan.conflict_renamed_count}개는 자동으로 새 이름으로 저장됩니다."
                    )

            if not plan.tasks:
                message = "실행할 변환 대상이 없습니다."
                if plan.skipped_count:
                    message += f"\n동일 형식 {plan.skipped_count}개는 자동으로 건너뜁니다."
                raise ValueError(message)

            preflight = PreflightDialog(plan, self)
            if preflight.exec() != QDialog.DialogCode.Accepted:
                self.status_label.setText("변환 시작이 취소되었습니다")
                return

            self.plan = plan
            self.tasks = plan.tasks
            self._save_settings()

            self._set_converting_state(True)
            self.progress_bar.setMaximum(plan.runnable_count)
            self.progress_bar.setValue(0)
            self.conversion_start_time = time.time()
            self.worker = ConversionWorker(plan)
            self.worker.progress_updated.connect(self._on_progress_updated)
            self.worker.status_updated.connect(self._on_status_updated)
            self.worker.task_completed.connect(self._on_task_completed)
            self.worker.error_occurred.connect(self._on_error_occurred)
            self.worker.finished.connect(self._on_worker_finished)
            self.worker.start()
            self.hwp_status_label.setText("🟡 한글 연결 중...")

            start_message = f"{plan.runnable_count}개 파일 변환 시작"
            if plan.skipped_count:
                start_message += f" (건너뜀 {plan.skipped_count}개)"
            self.toast.show_message(start_message, "🚀")
        except ValueError as e:
            QMessageBox.warning(self, "경고", str(e))
        except Exception as e:
            logger.exception("변환 시작 오류")
            QMessageBox.critical(self, "오류", f"오류 발생: {e}")

    def _request_worker_stop(self, waiting_text: str) -> bool:
        worker = self.worker
        if worker is None:
            return True

        self.status_label.setText(waiting_text)
        worker.cancel()
        if worker.wait(WORKER_WAIT_TIMEOUT):
            return True

        if worker.can_force_terminate():
            self._force_kill_pending = True
            self.cancel_btn.setText("🛑 강제 종료")
            self.status_label.setText("취소 요청됨 (응답 대기)")
        else:
            self._force_kill_pending = False
            self.cancel_btn.setText("⏹️ 취소")
            self.status_label.setText("안전하게 강제 종료할 대상 프로세스를 확인하지 못했습니다. 종료를 기다리는 중입니다.")
        return False

    def _perform_force_terminate(self) -> bool:
        worker = self.worker
        if worker is None:
            return False

        self.status_label.setText("강제 종료 중...")
        QApplication.processEvents()
        killed = worker.force_terminate()
        if not killed:
            self._force_kill_pending = False
            self.cancel_btn.setText("⏹️ 취소")
            QMessageBox.warning(
                self,
                "강제 종료 불가",
                "안전하게 종료할 대상 프로세스를 확인하지 못해 강제 종료를 수행하지 않았습니다.",
            )
            self.status_label.setText("안전한 강제 종료 대상이 없어 종료를 기다리는 중입니다.")
            return False

        worker.wait(1000)
        self._force_kill_pending = False
        self.cancel_btn.setText("⏹️ 취소")
        return True

    def _cancel_conversion(self) -> None:
        """변환 취소"""
        if not self.worker:
            return

        if self._force_kill_pending:
            reply = QMessageBox.question(
                self,
                "강제 종료 경고",
                "앱이 소유한 한글 프로세스만 강제 종료합니다.\n열려 있는 문서가 저장되지 않을 수 있습니다.\n\n계속할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                if self._perform_force_terminate():
                    self.status_label.setText("강제 종료 요청 완료")
            return

        reply = QMessageBox.question(
            self,
            "확인",
            "변환을 취소하시겠습니까?\n응답이 없으면 앱이 소유한 한글 프로세스만 강제 종료할 수 있습니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        if self._request_worker_stop("취소 요청 중..."):
            self.status_label.setText("취소됨")
    
    def _set_converting_state(self, converting: bool) -> None:
        """변환 중 상태 설정 - 입력 위젯 비활성화 포함"""
        if converting:
            self._cancel_active_scan()

        self.is_converting = converting
        self.start_btn.setEnabled(not converting)
        self.cancel_btn.setEnabled(converting)

        # 취소 후 무응답 플래그/버튼 텍스트 정리
        if not converting:
            self._force_kill_pending = False
            self.cancel_btn.setText("⏹️ 취소")
        
        # 변환 중에는 주요 입력 위젯 비활성화
        self.folder_radio.setEnabled(not converting)
        self.files_radio.setEnabled(not converting)
        self.files_radio.setEnabled(not converting)
        
        # 포맷 카드 비활성화
        for card in self.format_cards.values():
            card.setEnabled(not converting)
            
        self.same_location_check.setEnabled(not converting)
        self.overwrite_check.setEnabled(not converting)
        self.include_sub_check.setEnabled(not converting)
        
        # 파일 목록 변경 방지 - 변환 중 파일 추가/제거 차단
        if hasattr(self, 'drop_area'):
            self.drop_area.setEnabled(not converting)
        if hasattr(self, 'add_btn'):
            self.add_btn.setEnabled(not converting)
        if hasattr(self, 'remove_btn'):
            self.remove_btn.setEnabled(not converting)
        if hasattr(self, 'clear_btn'):
            self.clear_btn.setEnabled(not converting)
        
        # 폴더 모드 버튼도 비활성화
        if hasattr(self, 'folder_btn'):
            self.folder_btn.setEnabled(not converting)
        if hasattr(self, 'output_btn'):
            self.output_btn.setEnabled(not converting)
    
    def _on_progress_updated(self, current: int, total: int, filename: str) -> None:
        """진행률 업데이트"""
        self.progress_bar.setValue(current)
        
        # 예상 남은 시간 계산
        if current > 0 and self.conversion_start_time:
            elapsed = time.time() - self.conversion_start_time
            avg_time = elapsed / current
            remaining = avg_time * (total - current)
            remaining_str = f" (남은 시간: {int(remaining)}초)" if remaining > 0 else ""
        else:
            remaining_str = ""
        
        self.progress_label.setText(f"{current} / {total}{remaining_str}")
        self.status_label.setText(f"변환 중: {filename}")
    
    def _on_status_updated(self, text: str) -> None:
        """상태 텍스트 업데이트"""
        self.status_label.setText(text)
    
    def _on_task_completed(self, summary_obj: object) -> None:
        """작업 완료"""
        if not isinstance(summary_obj, ConversionSummary):
            return

        summary = summary_obj
        self.last_summary = summary
        elapsed_str = f"{summary.elapsed_seconds:.1f}초" if summary.elapsed_seconds is not None else "알 수 없음"

        if summary.failed_count == 0 and summary.canceled_count == 0:
            self.toast.show_message(
                f"✅ 성공 {summary.success_count}개, 건너뜀 {summary.skipped_count}개 ({elapsed_str})",
                "🎉",
            )
        else:
            self.toast.show_message(
                f"⚠️ 성공 {summary.success_count} / 실패 {summary.failed_count} / 취소 {summary.canceled_count} ({elapsed_str})",
                "⚠️",
            )

        self.hwp_status_label.setText("🟢 한글 연결됨")
        if self._close_after_worker:
            return
        dialog = ResultDialog(summary, self)
        dialog.exec()
    
    def _on_error_occurred(self, error_msg: str) -> None:
        """오류 발생"""
        self.toast.show_message("변환 중 오류 발생", "❌")
        self.hwp_status_label.setText("🔴 한글 연결 오류")
        QMessageBox.critical(self, "오류", f"변환 중 오류 발생:\n{error_msg}")
    
    def _on_worker_finished(self) -> None:
        """워커 종료"""
        self._set_converting_state(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText("0 / 0")
        self.status_label.setText("대기 중")
        self.hwp_status_label.setText("🟢 한글 대기중")

        if self.worker:
            try:
                self.worker.progress_updated.disconnect()
                self.worker.status_updated.disconnect()
                self.worker.task_completed.disconnect()
                self.worker.error_occurred.disconnect()
                self.worker.finished.disconnect()
            except (TypeError, RuntimeError):
                pass

        self.worker = None
        self.plan = None
        if self._close_after_worker:
            QTimer.singleShot(0, self.close)
    
    def closeEvent(self, a0: Optional[QCloseEvent]) -> None:
        """윈도우 닫기 이벤트"""
        if a0 is None:
            return
        logger.info("메인 윈도우 종료 이벤트 수신")
        if not self._cancel_active_scan(wait_ms=WORKER_WAIT_TIMEOUT):
            self.status_label.setText("파일 스캔 종료 대기 중...")
            a0.ignore()
            return

        if self.is_converting:
            if self._close_after_worker:
                self.status_label.setText("종료 대기 중...")
                a0.ignore()
                return

            reply = QMessageBox.question(
                self, "확인",
                "변환 작업이 진행 중입니다. 종료하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                a0.ignore()
                return

            self._close_after_worker = True
            if not self._request_worker_stop("종료 대기 중..."):
                if self.worker and self.worker.can_force_terminate():
                    reply2 = QMessageBox.question(
                        self,
                        "강제 종료 경고",
                        "앱이 소유한 한글 프로세스만 강제 종료합니다.\n열려 있는 문서가 저장되지 않을 수 있습니다.\n\n계속할까요?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    )
                    if reply2 == QMessageBox.StandardButton.Yes:
                        self._perform_force_terminate()

                if self.worker and self.worker.isRunning():
                    self.status_label.setText("종료 대기 중...")
                    a0.ignore()
                    return

        if hasattr(self, 'toast') and self.toast:
            self.toast.clear_all()

        if hasattr(self, 'tray_icon'):
            self.tray_icon.hide()

        self._save_settings()
        logger.info("메인 윈도우 종료 허용")
        a0.accept()
