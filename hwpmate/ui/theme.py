from __future__ import annotations

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
        QTableWidget::item:hover {
            background-color: #1e3a5f;
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
            border-radius: 12px;
            height: 24px;
            text-align: center;
            color: white;
            font-weight: bold;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #e94560, stop:0.5 #ff7b95, stop:1 #e94560);
            border-radius: 12px;
        }
        
        /* 메뉴바 */
        QMenuBar {
            background-color: #16213e;
            color: #eaeaea;
            border-bottom: 1px solid #0f3460;
            padding: 5px;
        }
        QMenuBar::item {
            padding: 8px 15px;
            border-radius: 5px;
        }
        QMenuBar::item:selected {
            background-color: #0f3460;
        }
        QMenu {
            background-color: #16213e;
            border: 1px solid #0f3460;
            border-radius: 8px;
            padding: 5px;
        }
        QMenu::item {
            padding: 8px 25px;
            border-radius: 5px;
        }
        QMenu::item:selected {
            background-color: #e94560;
        }
        
        /* 상태바 */
        QStatusBar {
            background-color: #16213e;
            color: #eaeaea;
            border-top: 1px solid #0f3460;
        }
        QStatusBar::item {
            border: none;
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
        QFrame[dropZone="true"] QLabel {
            color: #eaeaea;
            background-color: transparent;
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
            font-size: 16pt;
            font-weight: bold;
            color: #e94560;
        }
        QLabel[subheading="true"] {
            font-size: 9pt;
            color: #888899;
        }
        
        /* 포맷 카드 */
        QFrame[formatCard="true"] {
            background-color: #0f3460;
            border: 2px solid #1a4a80;
            border-radius: 12px;
            padding: 15px;
        }
        QFrame[formatCard="true"] QLabel {
            color: #eaeaea;
            background-color: transparent;
        }
        QFrame[formatCard="true"]:hover {
            border-color: #e94560;
            background-color: #162850;
        }
        QFrame[formatCardSelected="true"] {
            background-color: #1e3a5f;
            border: 2px solid #e94560;
            border-radius: 12px;
            padding: 15px;
        }
        QFrame[formatCardSelected="true"] QLabel {
            color: #ffffff;
            background-color: transparent;
        }

        /* 탭 위젯 */
        QTabWidget::pane {
            border: 1px solid #0f3460;
            background-color: #16213e;
            border-radius: 8px;
        }
        QTabWidget::tab-bar {
            left: 5px;
        }
        QTabBar::tab {
            background: #0f3460;
            color: #888899;
            border: 1px solid #0f3460;
            border-bottom-color: #0f3460;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            min-width: 100px;
            padding: 8px 15px;
            margin-right: 2px;
            font-weight: bold;
        }
        QTabBar::tab:selected, QTabBar::tab:hover {
            background: #16213e;
            color: #e94560;
            border-color: #0f3460;
            border-bottom-color: #16213e;
        }
        QTabBar::tab:selected {
            border-top: 2px solid #e94560;
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
        QTableWidget::item:hover {
            background-color: #f0f0ff;
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
            border-radius: 12px;
            height: 24px;
            text-align: center;
            color: #2d3436;
            font-weight: bold;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #6c5ce7, stop:0.5 #a29bfe, stop:1 #6c5ce7);
            border-radius: 12px;
        }
        
        /* 메뉴바 */
        QMenuBar {
            background-color: #ffffff;
            color: #2d3436;
            border-bottom: 1px solid #dfe6e9;
            padding: 5px;
        }
        QMenuBar::item {
            padding: 8px 15px;
            border-radius: 5px;
        }
        QMenuBar::item:selected {
            background-color: #f0f0ff;
        }
        QMenu {
            background-color: #ffffff;
            border: 1px solid #dfe6e9;
            border-radius: 8px;
            padding: 5px;
        }
        QMenu::item {
            padding: 8px 25px;
            border-radius: 5px;
        }
        QMenu::item:selected {
            background-color: #6c5ce7;
            color: white;
        }
        
        /* 상태바 */
        QStatusBar {
            background-color: #ffffff;
            color: #2d3436;
            border-top: 1px solid #dfe6e9;
        }
        QStatusBar::item {
            border: none;
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
        QFrame[dropZone="true"] QLabel {
            color: #2d3436;
            background-color: transparent;
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
            font-size: 16pt;
            font-weight: bold;
            color: #6c5ce7;
        }
        QLabel[subheading="true"] {
            font-size: 9pt;
            color: #636e72;
        }
        
        /* 포맷 카드 */
        QFrame[formatCard="true"] {
            background-color: #ffffff;
            border: 4px solid #dfe6e9;
            border-radius: 13px;
            padding: 17px;
        }
        QFrame[formatCard="true"] QLabel {
            color: #2d3436;
            background-color: transparent;
        }
        QFrame[formatCard="true"]:hover {
            border-color: #6c5ce7;
            background-color: #f0f0ff;
        }
        QFrame[formatCardSelected="true"] {
            background-color: #f0f0ff;
            border: 4px solid #6c5ce7;
            border-radius: 13px;
            padding: 17px;
        }
        QFrame[formatCardSelected="true"] QLabel {
            color: #2d3436;
            background-color: transparent;
        }

        /* 탭 위젯 */
        QTabWidget::pane {
            border: 1px solid #dfe6e9;
            background-color: #ffffff;
            border-radius: 8px;
        }
        QTabWidget::tab-bar {
            left: 5px;
        }
        QTabBar::tab {
            background: #f1f2f6;
            color: #636e72;
            border: 1px solid #dfe6e9;
            border-bottom-color: #dfe6e9;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            min-width: 100px;
            padding: 8px 15px;
            margin-right: 2px;
            font-weight: bold;
        }
        QTabBar::tab:selected, QTabBar::tab:hover {
            background: #ffffff;
            color: #6c5ce7;
            border-color: #dfe6e9;
            border-bottom-color: #ffffff;
        }
        QTabBar::tab:selected {
            border-top: 2px solid #6c5ce7;
        }
    """
    
    @staticmethod
    def get_theme(theme_name: str) -> str:
        if theme_name == "dark":
            return ThemeManager.DARK_THEME
        return ThemeManager.LIGHT_THEME
