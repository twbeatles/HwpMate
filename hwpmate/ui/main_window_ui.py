from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QTableWidget,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ..constants import FORMAT_GROUPS, FORMAT_TYPES, VERSION, WINDOW_DEFAULT_HEIGHT, WINDOW_DEFAULT_WIDTH, WINDOW_MIN_HEIGHT, WINDOW_MIN_WIDTH
from .widgets import DropArea, FormatCard


@dataclass
class MainWindowWidgets:
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
    backup_check: QCheckBox
    retry_spin: QSpinBox
    start_btn: QPushButton
    cancel_btn: QPushButton
    status_label: QLabel
    progress_bar: QProgressBar
    progress_label: QLabel


def build_main_window_ui(window: Any, config: Any) -> MainWindowWidgets:
    window.setWindowTitle(f"HWP 변환기 v{VERSION} - PyQt6")
    window.setMinimumSize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
    window.resize(WINDOW_DEFAULT_WIDTH, WINDOW_DEFAULT_HEIGHT)

    # 스크롤 영역 설정
    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll_area.setFrameShape(QFrame.Shape.NoFrame)
    window.setCentralWidget(scroll_area)

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
    window.theme_btn = QPushButton("🌙 다크" if window.current_theme == "dark" else "☀️ 라이트")
    window.theme_btn.setProperty("secondary", True)
    window.theme_btn.setFixedWidth(100)
    window.theme_btn.setToolTip("다크 모드와 라이트 모드를 전환합니다")
    window.theme_btn.clicked.connect(window._toggle_theme)
    header_layout.addWidget(window.theme_btn)

    main_layout.addLayout(header_layout)

    # === 모드 선택 ===
    mode_group = QGroupBox("변환 모드")
    mode_layout = QVBoxLayout(mode_group)
    mode_layout.setSpacing(8)

    window.mode_group = QButtonGroup(window)

    window.folder_radio = QRadioButton("📁 폴더 일괄 변환 (폴더 내 모든 파일)")
    window.folder_radio.setToolTip("폴더 내 모든 HWP/HWPX 파일을 일괄 변환합니다")
    window.files_radio = QRadioButton("📄 파일 개별 선택 (원하는 파일만)")
    window.files_radio.setToolTip("원하는 파일만 선택하여 변환합니다")

    window.mode_group.addButton(window.folder_radio, 0)
    window.mode_group.addButton(window.files_radio, 1)

    mode_layout.addWidget(window.folder_radio)
    mode_layout.addWidget(window.files_radio)

    saved_mode = window.config.get("mode", "folder")
    if saved_mode == "folder":
        window.folder_radio.setChecked(True)
    else:
        window.files_radio.setChecked(True)

    window.folder_radio.toggled.connect(window._update_mode_ui)

    main_layout.addWidget(mode_group)

    # === 입력 영역 ===
    input_group = QGroupBox("입력")
    input_layout = QVBoxLayout(input_group)
    input_layout.setSpacing(12)

    # 폴더 모드 위젯
    window.folder_widget = QWidget()
    folder_layout = QVBoxLayout(window.folder_widget)
    folder_layout.setContentsMargins(0, 0, 0, 0)
    folder_layout.setSpacing(10)

    folder_row = QHBoxLayout()
    folder_row.setSpacing(10)
    window.folder_entry = QLineEdit()
    window.folder_entry.setPlaceholderText("변환할 폴더를 선택하세요...")
    window.folder_entry.setReadOnly(True)
    window.folder_entry.setMinimumHeight(40)
    folder_row.addWidget(window.folder_entry)

    window.folder_btn = QPushButton("찾아보기")
    window.folder_btn.setProperty("secondary", True)
    window.folder_btn.setFixedWidth(100)
    window.folder_btn.setMinimumHeight(40)
    window.folder_btn.clicked.connect(window._select_folder)
    folder_row.addWidget(window.folder_btn)

    folder_layout.addLayout(folder_row)

    window.include_sub_check = QCheckBox("하위 폴더 포함")
    window.include_sub_check.setToolTip("하위 폴더의 파일도 함께 변환합니다")
    window.include_sub_check.setChecked(window.config.get("include_sub", True))
    window.include_sub_check.toggled.connect(window._on_include_sub_toggled)
    folder_layout.addWidget(window.include_sub_check)

    # 저장된 폴더 경로 복원
    saved_folder = window.config.get("folder_path", "")
    if saved_folder and Path(saved_folder).exists():
        window.folder_entry.setText(saved_folder)

    input_layout.addWidget(window.folder_widget)

    # 파일 모드 위젯
    window.files_widget = QWidget()
    files_layout = QVBoxLayout(window.files_widget)
    files_layout.setContentsMargins(0, 0, 0, 0)
    files_layout.setSpacing(12)

    # 드롭 영역 - 고정 높이
    window.drop_area = DropArea()
    window.drop_area.setFixedHeight(120)
    window.drop_area.files_dropped.connect(window._add_files)
    files_layout.addWidget(window.drop_area)

    # 버튼 행
    btn_row = QHBoxLayout()
    btn_row.setSpacing(8)

    window.add_btn = QPushButton("➕ 파일 추가")
    window.add_btn.setProperty("secondary", True)
    window.add_btn.setMinimumHeight(36)
    window.add_btn.setToolTip("파일 선택 대화상자를 엽니다 (Ctrl+O)")
    window.add_btn.clicked.connect(window._browse_files)
    btn_row.addWidget(window.add_btn)

    window.remove_btn = QPushButton("➖ 선택 제거")
    window.remove_btn.setProperty("secondary", True)
    window.remove_btn.setMinimumHeight(36)
    window.remove_btn.setToolTip("선택한 파일을 목록에서 제거합니다 (Delete)")
    window.remove_btn.clicked.connect(window._remove_selected)
    btn_row.addWidget(window.remove_btn)

    window.clear_btn = QPushButton("🗑️ 전체 제거")
    window.clear_btn.setProperty("secondary", True)
    window.clear_btn.setMinimumHeight(36)
    window.clear_btn.setToolTip("모든 파일을 목록에서 제거합니다 (Ctrl+Delete)")
    window.clear_btn.clicked.connect(window._clear_all)
    btn_row.addWidget(window.clear_btn)

    btn_row.addStretch()
    files_layout.addLayout(btn_row)

    # 파일 테이블 - 고정 높이
    window.file_table = QTableWidget()
    window.file_table.setColumnCount(2)
    window.file_table.setHorizontalHeaderLabels(["파일명", "경로"])
    horizontal_header = window.file_table.horizontalHeader()
    vertical_header = window.file_table.verticalHeader()
    assert horizontal_header is not None
    assert vertical_header is not None
    horizontal_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
    horizontal_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
    window.file_table.setAlternatingRowColors(True)
    window.file_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
    window.file_table.setFixedHeight(180)
    vertical_header.setVisible(False)
    window.file_table.setSortingEnabled(False)  # 정렬 비활성화 - file_list 동기화 문제 방지
    files_layout.addWidget(window.file_table)

    input_layout.addWidget(window.files_widget)

    main_layout.addWidget(input_group)

    # === 출력 설정 ===
    output_group = QGroupBox("출력")
    output_layout = QVBoxLayout(output_group)
    output_layout.setSpacing(10)

    window.same_location_check = QCheckBox("입력 파일과 같은 위치에 저장")
    window.same_location_check.setToolTip("변환된 파일을 원본과 같은 폴더에 저장합니다")
    window.same_location_check.setChecked(window.config.get("same_location", True))
    window.same_location_check.toggled.connect(window._update_output_ui)
    output_layout.addWidget(window.same_location_check)

    output_row = QHBoxLayout()
    output_row.setSpacing(10)
    output_label = QLabel("저장 폴더:")
    output_label.setFixedWidth(70)
    output_row.addWidget(output_label)

    window.output_entry = QLineEdit()
    window.output_entry.setPlaceholderText("저장할 폴더를 선택하세요...")
    window.output_entry.setReadOnly(True)
    window.output_entry.setMinimumHeight(40)
    output_row.addWidget(window.output_entry)

    window.output_btn = QPushButton("찾아보기")
    window.output_btn.setProperty("secondary", True)
    window.output_btn.setFixedWidth(100)
    window.output_btn.setMinimumHeight(40)
    window.output_btn.clicked.connect(window._select_output)
    output_row.addWidget(window.output_btn)

    output_layout.addLayout(output_row)

    # 저장된 출력 경로 복원
    saved_output = window.config.get("output_path", "")
    if saved_output and Path(saved_output).exists():
        window.output_entry.setText(saved_output)

    main_layout.addWidget(output_group)

    # === 변환 옵션 ===
    options_group = QGroupBox("변환 형식")
    options_layout = QVBoxLayout(options_group)
    options_layout.setSpacing(15)

    # 변환 형식 카드 UI (Tab Widget 사용)
    from PyQt6.QtWidgets import QGridLayout, QTabWidget

    window.format_tabs = QTabWidget()
    window.format_cards = {}

    # 탭별 포맷 정의
    tabs_config = FORMAT_GROUPS

    for tab_name, formats in tabs_config.items():
        tab_widget = QWidget()
        tab_layout = QGridLayout(tab_widget)
        tab_layout.setSpacing(15)
        tab_layout.setContentsMargins(15, 15, 15, 15)

        row = 0
        col = 0
        max_cols = 4

        for fmt_key in formats:
            if fmt_key not in FORMAT_TYPES:
                continue

            info = FORMAT_TYPES[fmt_key]
            card = FormatCard(
                fmt_key, 
                info['icon'], 
                fmt_key, 
                info['desc']
            )
            card.clicked.connect(window._on_format_card_clicked)
            card.setMinimumSize(120, 120) # 크기 충분히 확보 (텍스트 잘림 방지)
            card.setMaximumWidth(1000)

            tab_layout.addWidget(card, row, col)
            window.format_cards[fmt_key] = card

            col += 1
            if col >= max_cols:
                col = 0
                row += 1

        # 빈 공간 채우기 (레이아웃 틀어짐 방지)
        if col > 0:
            tab_layout.setColumnStretch(max_cols-1, 1)
        tab_layout.setRowStretch(row+1, 1)

        window.format_tabs.addTab(tab_widget, tab_name)

    # 저장된 형식 복원
    window._selected_format = window.config.get("format", "PDF")
    # 없는 형식이면 기본값 PDF
    if window._selected_format not in FORMAT_TYPES:
        window._selected_format = "PDF"

    # 선택된 포맷이 있는 탭 활성화
    for i in range(window.format_tabs.count()):
        tab_name = window.format_tabs.tabText(i)
        if window._selected_format in tabs_config.get(tab_name, []):
            window.format_tabs.setCurrentIndex(i)
            break

    window._update_format_cards()

    options_layout.addWidget(window.format_tabs)

    # 덮어쓰기 옵션
    window.overwrite_check = QCheckBox("기존 파일 덮어쓰기 (체크 해제 시 번호 자동 추가)")
    window.overwrite_check.setToolTip("같은 이름의 파일이 있으면 덮어씁니다")
    window.overwrite_check.setChecked(window.config.get("overwrite", False))
    options_layout.addWidget(window.overwrite_check)

    window.backup_check = QCheckBox("변환 전 원본 백업")
    window.backup_check.setToolTip("원본 파일을 각 폴더의 backup 폴더에 복사한 뒤 변환합니다")
    window.backup_check.setChecked(window.config.get("backup_enabled", True))
    options_layout.addWidget(window.backup_check)

    retry_row = QHBoxLayout()
    retry_row.setSpacing(10)
    retry_label = QLabel("실패 시 재시도:")
    retry_label.setFixedWidth(100)
    retry_row.addWidget(retry_label)

    window.retry_spin = QSpinBox()
    window.retry_spin.setRange(0, 3)
    window.retry_spin.setValue(int(window.config.get("retry_count", 1)))
    window.retry_spin.setToolTip("파일별 변환 실패 시 재시도 횟수입니다")
    window.retry_spin.setFixedWidth(80)
    retry_row.addWidget(window.retry_spin)
    retry_row.addWidget(QLabel("회"))
    retry_row.addStretch()
    options_layout.addLayout(retry_row)

    main_layout.addWidget(options_group)

    # === 실행 버튼 ===
    btn_layout = QHBoxLayout()
    btn_layout.setSpacing(10)

    window.start_btn = QPushButton("🚀 변환 시작")
    window.start_btn.setMinimumHeight(55)
    window.start_btn.setToolTip("선택한 파일을 변환합니다 (Ctrl+Enter)")
    font = window.start_btn.font()
    font.setPointSize(12)
    font.setBold(True)
    window.start_btn.setFont(font)
    window.start_btn.clicked.connect(window._start_conversion)
    btn_layout.addWidget(window.start_btn)

    window.cancel_btn = QPushButton("⏹️ 취소")
    window.cancel_btn.setProperty("secondary", True)
    window.cancel_btn.setMinimumHeight(55)
    window.cancel_btn.setFixedWidth(100)
    window.cancel_btn.setToolTip("진행 중인 변환을 취소합니다 (Esc)")
    window.cancel_btn.setEnabled(False)
    window.cancel_btn.clicked.connect(window._cancel_conversion)
    btn_layout.addWidget(window.cancel_btn)

    main_layout.addLayout(btn_layout)

    # 팁 메시지
    tip_label = QLabel("💡 Tip: 변환 시작 시 나오는 팝업에서 '모두 허용'을 눌러주셔야 진행됩니다.")
    tip_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    tip_label.setStyleSheet("color: #ff9f43; font-weight: bold; margin-top: 5px;")
    main_layout.addWidget(tip_label)

    # === 진행 상태 ===
    progress_group = QGroupBox("진행 상태")
    progress_layout = QVBoxLayout(progress_group)
    progress_layout.setSpacing(8)

    window.status_label = QLabel("준비됨")
    window.status_label.setMinimumHeight(25)
    progress_layout.addWidget(window.status_label)

    window.progress_bar = QProgressBar()
    window.progress_bar.setValue(0)
    window.progress_bar.setMinimumHeight(28)
    progress_layout.addWidget(window.progress_bar)

    window.progress_label = QLabel("0 / 0")
    window.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    progress_layout.addWidget(window.progress_label)

    main_layout.addWidget(progress_group)

    # 하단 여백
    main_layout.addSpacing(20)

    return MainWindowWidgets(
        theme_btn=window.theme_btn,
        folder_radio=window.folder_radio,
        files_radio=window.files_radio,
        folder_widget=window.folder_widget,
        folder_entry=window.folder_entry,
        folder_btn=window.folder_btn,
        include_sub_check=window.include_sub_check,
        files_widget=window.files_widget,
        drop_area=window.drop_area,
        add_btn=window.add_btn,
        remove_btn=window.remove_btn,
        clear_btn=window.clear_btn,
        file_table=window.file_table,
        same_location_check=window.same_location_check,
        output_entry=window.output_entry,
        output_btn=window.output_btn,
        format_tabs=window.format_tabs,
        format_cards=window.format_cards,
        overwrite_check=window.overwrite_check,
        backup_check=window.backup_check,
        retry_spin=window.retry_spin,
        start_btn=window.start_btn,
        cancel_btn=window.cancel_btn,
        status_label=window.status_label,
        progress_bar=window.progress_bar,
        progress_label=window.progress_label,
    )
