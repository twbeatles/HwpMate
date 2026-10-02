# -*- coding: utf-8 -*-
"""팔레트 토큰으로 앱 전체 QSS 를 만든다 (라이트/다크 공용 템플릿)."""

from __future__ import annotations

from .icons import icon_url
from .palette import ThemePalette

FONT_FAMILY = "'Malgun Gothic', 'Segoe UI', sans-serif"


def _image_rule(url: str) -> str:
    return f"image: url({url});" if url else ""


def build_stylesheet(p: ThemePalette, *, with_icons: bool = True) -> str:
    check = icon_url("check", p.on_accent, 18) if with_icons else ""
    check_disabled = icon_url("check", p.text_disabled, 18) if with_icons else ""
    menu_check = icon_url("check", p.text, 18) if with_icons else ""
    menu_check_selected = icon_url("check", p.on_accent, 18) if with_icons else ""
    down = icon_url("chevron-down", p.text_muted, 14) if with_icons else ""
    down_disabled = icon_url("chevron-down", p.text_disabled, 14) if with_icons else ""
    up = icon_url("chevron-up", p.text_muted, 12) if with_icons else ""
    up_disabled = icon_url("chevron-up", p.text_disabled, 12) if with_icons else ""
    spin_down = icon_url("chevron-down", p.text_muted, 12) if with_icons else ""
    spin_down_disabled = icon_url("chevron-down", p.text_disabled, 12) if with_icons else ""

    return f"""
/* ===== 기본 ===== */
QWidget {{
    color: {p.text};
    font-family: {FONT_FAMILY};
    font-size: 10pt;
}}
QMainWindow, QDialog, QMessageBox {{
    background-color: {p.window_bg};
}}
QScrollArea#mainScroll,
QScrollArea#mainScroll > QWidget#qt_scrollarea_viewport,
QWidget#scrollContent {{
    background-color: {p.window_bg};
    border: none;
}}
QLabel {{
    background: transparent;
}}
QLabel:disabled {{
    color: {p.text_disabled};
}}
QToolTip {{
    background-color: {p.toast_bg};
    color: {p.toast_text};
    border: 1px solid {p.toast_border};
    border-radius: 6px;
    padding: 6px 8px;
}}

/* ===== 그룹(카드) ===== */
QGroupBox {{
    background-color: {p.surface};
    border: 1px solid {p.border};
    border-radius: 10px;
    margin-top: 14px;
    padding: 16px 14px 12px 14px;
    font-weight: bold;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: {p.text};
    background-color: {p.window_bg};
    border-radius: 4px;
}}
QGroupBox QLabel, QGroupBox QCheckBox, QGroupBox QRadioButton {{
    font-weight: normal;
}}

/* ===== 텍스트 역할 ===== */
QLabel[heading="true"] {{
    font-size: 16pt;
    font-weight: bold;
    color: {p.text};
}}
QLabel[subheading="true"] {{
    font-size: 9pt;
    color: {p.text_muted};
}}
QLabel[caption="true"] {{
    font-size: 8pt;
    color: {p.text_muted};
}}
QLabel[fieldLabel="true"] {{
    color: {p.text_muted};
}}
QLabel[statusText="true"] {{
    font-weight: bold;
}}
QLabel[hint="true"] {{
    background-color: {p.warning_bg};
    color: {p.warning_text};
    border: 1px solid {p.warning_border};
    border-radius: 8px;
    padding: 8px 12px;
}}

/* ===== 버튼 ===== */
QPushButton {{
    background-color: {p.accent};
    color: {p.on_accent};
    border: 1px solid {p.accent};
    border-radius: 8px;
    padding: 8px 18px;
    font-weight: bold;
    min-height: 20px;
}}
QPushButton:hover {{
    background-color: {p.accent_hover};
    border-color: {p.accent_hover};
}}
QPushButton:pressed {{
    background-color: {p.accent_pressed};
    border-color: {p.accent_pressed};
}}
QPushButton:focus {{
    outline: none;
}}
QPushButton:disabled {{
    background-color: {p.surface_alt};
    color: {p.text_disabled};
    border-color: {p.border};
}}
QPushButton[secondary="true"] {{
    background-color: {p.surface};
    color: {p.text};
    border: 1px solid {p.border_strong};
    font-weight: normal;
}}
QPushButton[secondary="true"]:hover {{
    background-color: {p.accent_soft};
    border-color: {p.accent};
    color: {p.text};
}}
QPushButton[secondary="true"]:pressed {{
    background-color: {p.accent_soft};
    border-color: {p.accent_pressed};
}}
QPushButton[secondary="true"]:disabled {{
    background-color: {p.surface_alt};
    color: {p.text_disabled};
    border-color: {p.border};
}}
QPushButton[large="true"] {{
    font-size: 12pt;
    font-weight: bold;
    padding: 10px 22px;
}}
QPushButton[danger="true"] {{
    background-color: {p.surface};
    color: {p.danger};
    border: 1px solid {p.danger};
}}
QPushButton[danger="true"]:hover {{
    background-color: {p.danger_soft};
    border-color: {p.danger_hover};
    color: {p.danger_hover};
}}
QPushButton[danger="true"]:disabled {{
    background-color: {p.surface_alt};
    color: {p.text_disabled};
    border-color: {p.border};
}}

/* ===== 입력 ===== */
QLineEdit, QSpinBox, QComboBox, QTextEdit, QPlainTextEdit {{
    background-color: {p.input_bg};
    color: {p.text};
    border: 1px solid {p.border_strong};
    border-radius: 8px;
    selection-background-color: {p.accent};
    selection-color: {p.on_accent};
}}
QLineEdit {{
    padding: 7px 10px;
}}
QLineEdit:read-only {{
    background-color: {p.surface_alt};
}}
QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {{
    border-color: {p.accent};
}}
QLineEdit:disabled, QSpinBox:disabled, QComboBox:disabled {{
    background-color: {p.surface_alt};
    color: {p.text_disabled};
    border-color: {p.border};
}}
QTextEdit, QPlainTextEdit {{
    padding: 6px;
}}

/* 스핀박스 */
QSpinBox {{
    padding: 5px 26px 5px 10px;
    min-height: 20px;
}}
QSpinBox::up-button, QSpinBox::down-button {{
    subcontrol-origin: border;
    width: 22px;
    border: none;
    border-left: 1px solid {p.border};
    background: transparent;
}}
QSpinBox::up-button {{
    subcontrol-position: top right;
    border-top-right-radius: 8px;
}}
QSpinBox::down-button {{
    subcontrol-position: bottom right;
    border-bottom-right-radius: 8px;
}}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{
    background-color: {p.accent_soft};
}}
QSpinBox::up-arrow {{
    {_image_rule(up)}
    width: 12px;
    height: 12px;
}}
QSpinBox::down-arrow {{
    {_image_rule(spin_down)}
    width: 12px;
    height: 12px;
}}
QSpinBox::up-arrow:disabled, QSpinBox::up-arrow:off {{
    {_image_rule(up_disabled)}
}}
QSpinBox::down-arrow:disabled, QSpinBox::down-arrow:off {{
    {_image_rule(spin_down_disabled)}
}}

/* 드롭다운 */
QComboBox {{
    padding: 6px 34px 6px 10px;
    min-height: 22px;
}}
QComboBox:hover {{
    border-color: {p.accent};
}}
QComboBox:on {{
    border-color: {p.accent};
    border-bottom-left-radius: 8px;
    border-bottom-right-radius: 8px;
}}
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 28px;
    border: none;
    border-left: 1px solid {p.border};
}}
QComboBox::down-arrow {{
    {_image_rule(down)}
    width: 14px;
    height: 14px;
}}
QComboBox::down-arrow:disabled {{
    {_image_rule(down_disabled)}
}}
QComboBox QAbstractItemView {{
    background-color: {p.surface};
    color: {p.text};
    border: 1px solid {p.border_strong};
    border-radius: 8px;
    padding: 4px;
    outline: 0;
    selection-background-color: {p.accent_soft};
    selection-color: {p.text};
}}
QComboBox QAbstractItemView::item {{
    min-height: 30px;
    padding: 0 10px;
    border-radius: 6px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {p.accent_soft};
    color: {p.text};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {p.accent};
    color: {p.on_accent};
}}

/* ===== 체크박스 / 라디오 ===== */
QCheckBox, QRadioButton {{
    spacing: 9px;
    padding: 3px 0;
    background: transparent;
}}
QCheckBox:disabled, QRadioButton:disabled {{
    color: {p.text_disabled};
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 18px;
    height: 18px;
}}
QCheckBox::indicator {{
    border: 2px solid {p.border_strong};
    border-radius: 5px;
    background-color: {p.input_bg};
}}
QCheckBox::indicator:hover {{
    border-color: {p.accent};
}}
QCheckBox::indicator:checked {{
    background-color: {p.accent};
    border-color: {p.accent};
    {_image_rule(check)}
}}
QCheckBox::indicator:checked:hover {{
    background-color: {p.accent_hover};
    border-color: {p.accent_hover};
}}
QCheckBox::indicator:disabled {{
    background-color: {p.surface_alt};
    border-color: {p.border};
}}
QCheckBox::indicator:checked:disabled {{
    background-color: {p.border};
    border-color: {p.border};
    {_image_rule(check_disabled)}
}}
QRadioButton::indicator {{
    border: 2px solid {p.border_strong};
    border-radius: 10px;
    background-color: {p.input_bg};
}}
QRadioButton::indicator:hover {{
    border-color: {p.accent};
}}
QRadioButton::indicator:checked {{
    border: 2px solid {p.accent};
    background-color: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
        stop:0 {p.accent}, stop:0.42 {p.accent}, stop:0.52 {p.input_bg}, stop:1 {p.input_bg});
}}
QRadioButton::indicator:disabled {{
    background-color: {p.surface_alt};
    border-color: {p.border};
}}
QRadioButton::indicator:checked:disabled {{
    border-color: {p.border};
    background-color: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
        stop:0 {p.text_disabled}, stop:0.42 {p.text_disabled}, stop:0.52 {p.surface_alt}, stop:1 {p.surface_alt});
}}

/* 변환 모드 선택 (라디오 카드형) */
QRadioButton[modeOption="true"] {{
    background-color: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 8px;
    padding: 9px 12px;
}}
QRadioButton[modeOption="true"]:hover {{
    border-color: {p.accent};
}}
QRadioButton[modeOption="true"]:checked {{
    background-color: {p.accent_soft};
    border-color: {p.accent};
    font-weight: bold;
}}

/* ===== 테이블 ===== */
QTableWidget, QTableView {{
    background-color: {p.surface};
    alternate-background-color: {p.surface_alt};
    color: {p.text};
    border: 1px solid {p.border};
    border-radius: 8px;
    gridline-color: transparent;
    selection-background-color: {p.accent_soft};
    selection-color: {p.text};
    outline: 0;
}}
QTableWidget::item, QTableView::item {{
    padding: 4px 8px;
    border: none;
}}
QTableWidget::item:selected, QTableView::item:selected {{
    background-color: {p.accent_soft};
    color: {p.text};
}}
QHeaderView {{
    background-color: transparent;
}}
QHeaderView::section {{
    background-color: {p.surface_alt};
    color: {p.text_muted};
    padding: 7px 8px;
    border: none;
    border-bottom: 1px solid {p.border};
    font-weight: bold;
}}
QTableCornerButton::section {{
    background-color: {p.surface_alt};
    border: none;
}}

/* ===== 진행률 ===== */
QProgressBar {{
    background-color: {p.surface_alt};
    border: 1px solid {p.border};
    border-radius: 8px;
    min-height: 18px;
    text-align: center;
    color: {p.text};
    font-weight: bold;
}}
QProgressBar::chunk {{
    background-color: {p.accent};
    border-radius: 7px;
}}

/* ===== 메뉴 ===== */
QMenuBar {{
    background-color: {p.surface};
    color: {p.text};
    border-bottom: 1px solid {p.border};
    padding: 2px 6px;
}}
QMenuBar::item {{
    background: transparent;
    padding: 6px 12px;
    border-radius: 6px;
}}
QMenuBar::item:selected, QMenuBar::item:pressed {{
    background-color: {p.accent_soft};
}}
QMenu {{
    background-color: {p.surface};
    color: {p.text};
    border: 1px solid {p.border_strong};
    border-radius: 8px;
    padding: 5px;
}}
QMenu::item {{
    padding: 7px 28px 7px 30px;
    border-radius: 6px;
    background: transparent;
}}
QMenu::item:selected {{
    background-color: {p.accent};
    color: {p.on_accent};
}}
QMenu::item:disabled {{
    color: {p.text_disabled};
    background: transparent;
}}
QMenu::separator {{
    height: 1px;
    background-color: {p.border};
    margin: 5px 8px;
}}
QMenu::indicator {{
    width: 16px;
    height: 16px;
    left: 8px;
    border-radius: 4px;
}}
QMenu::indicator:non-exclusive:unchecked {{
    border: 2px solid {p.border_strong};
    background-color: {p.input_bg};
}}
QMenu::indicator:non-exclusive:checked {{
    border: 2px solid {p.accent};
    background-color: {p.accent};
    {_image_rule(check)}
}}
QMenu::indicator:exclusive:checked {{
    {_image_rule(menu_check)}
}}
QMenu::indicator:exclusive:checked:selected {{
    {_image_rule(menu_check_selected)}
}}

/* ===== 상태바 ===== */
QStatusBar {{
    background-color: {p.surface};
    color: {p.text_muted};
    border-top: 1px solid {p.border};
}}
QStatusBar QLabel {{
    color: {p.text_muted};
    padding: 0 8px;
}}
QStatusBar::item {{
    border: none;
}}

/* ===== 스크롤바 ===== */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 2px;
}}
QScrollBar::handle:vertical {{
    background-color: {p.border_strong};
    border-radius: 3px;
    min-height: 32px;
}}
QScrollBar::handle:horizontal {{
    background-color: {p.border_strong};
    border-radius: 3px;
    min-width: 32px;
}}
QScrollBar::handle:hover {{
    background-color: {p.text_muted};
}}
QScrollBar::add-line, QScrollBar::sub-line {{
    width: 0px;
    height: 0px;
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: transparent;
}}

/* ===== 탭 (형식 그룹) ===== */
QTabWidget::pane {{
    border: 1px solid {p.border};
    background-color: {p.surface};
    border-radius: 8px;
    top: -1px;
}}
QTabWidget::tab-bar {{
    left: 8px;
}}
QTabBar::tab {{
    background: transparent;
    color: {p.text_muted};
    border: 1px solid transparent;
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    min-width: 88px;
    padding: 7px 14px;
    margin-right: 2px;
    font-weight: bold;
}}
QTabBar::tab:hover {{
    color: {p.text};
    background: {p.surface_alt};
}}
QTabBar::tab:selected {{
    background: {p.surface};
    color: {p.accent};
    border-color: {p.border};
    border-top: 2px solid {p.accent};
}}
QTabBar::tab:disabled {{
    color: {p.text_disabled};
}}

/* ===== 드롭 영역 ===== */
QFrame[dropZone="true"] {{
    background-color: {p.surface_alt};
    border: 2px dashed {p.border_strong};
    border-radius: 12px;
}}
QFrame[dropZone="true"]:hover {{
    border-color: {p.accent};
    background-color: {p.accent_soft};
}}
QFrame[dropZone="true"][dropActive="true"] {{
    border: 2px solid {p.accent};
    background-color: {p.accent_soft};
}}
QFrame[dropZone="true"]:disabled {{
    border-color: {p.border};
    background-color: {p.surface_alt};
}}
QFrame[dropZone="true"] QLabel {{
    background: transparent;
}}
QLabel[dropIcon="true"] {{
    font-size: 24pt;
}}

/* ===== 형식 카드 ===== */
QFrame[formatCard="true"], QFrame[formatCardSelected="true"] {{
    background-color: {p.surface};
    border: 1px solid {p.border};
    border-radius: 10px;
}}
QFrame[formatCard="true"]:hover {{
    border-color: {p.accent};
    background-color: {p.surface_alt};
}}
QFrame[formatCardSelected="true"] {{
    background-color: {p.accent_soft};
    border: 2px solid {p.accent};
}}
QFrame[formatCard="true"]:disabled, QFrame[formatCardSelected="true"]:disabled {{
    background-color: {p.surface_alt};
}}
QFrame[formatCard="true"] QLabel, QFrame[formatCardSelected="true"] QLabel {{
    background: transparent;
}}
QLabel[cardIcon="true"] {{
    font-size: 20pt;
}}
QLabel[cardTitle="true"] {{
    font-size: 11pt;
    font-weight: bold;
}}
QFrame[formatCardSelected="true"] QLabel[cardTitle="true"] {{
    color: {p.accent};
}}

/* ===== 구분선 ===== */
QFrame[separator="true"] {{
    background-color: {p.border};
    border: none;
    max-height: 1px;
    min-height: 1px;
}}
"""
