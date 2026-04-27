from __future__ import annotations

from .models import FormatSpec

VERSION = "8.7"
SUPPORTED_EXTENSIONS = (".hwp", ".hwpx")
BACKUP_DIR_NAME = "backup"

FORMAT_TYPES: dict[str, FormatSpec] = {
    "HWP": FormatSpec(ext=".hwp", save_format="HWP", icon="📝", desc="한글 문서"),
    "HWPX": FormatSpec(ext=".hwpx", save_format="HWPX", icon="📘", desc="한글 표준 문서"),
    "PDF": FormatSpec(ext=".pdf", save_format="PDF", icon="📕", desc="PDF 문서"),
    "DOCX": FormatSpec(ext=".docx", save_format="OOXML", icon="📄", desc="MS Word"),
    "ODT": FormatSpec(ext=".odt", save_format="ODT", icon="🌐", desc="ODF 텍스트"),
    "HTML": FormatSpec(ext=".html", save_format="HTML", icon="🌍", desc="웹 문서"),
    "RTF": FormatSpec(ext=".rtf", save_format="RTF", icon="📋", desc="서식있는 텍스트"),
    "TXT": FormatSpec(ext=".txt", save_format="TEXT", icon="📝", desc="텍스트 문서"),
    "PNG": FormatSpec(ext=".png", save_format="PNG", icon="🖼️", desc="PNG 이미지"),
    "JPG": FormatSpec(ext=".jpg", save_format="JPG", icon="📷", desc="JPG 이미지"),
    "BMP": FormatSpec(ext=".bmp", save_format="BMP", icon="🎨", desc="BMP 이미지"),
    "GIF": FormatSpec(ext=".gif", save_format="GIF", icon="🎞️", desc="GIF 이미지"),
}

FORMAT_GROUPS: dict[str, list[str]] = {
    "문서 변환": ["HWP", "HWPX", "PDF", "DOCX", "ODT", "HTML", "RTF", "TXT"],
    "이미지 변환": ["PNG", "JPG", "BMP", "GIF"],
}

WINDOW_MIN_WIDTH = 750
WINDOW_MIN_HEIGHT = 700
WINDOW_DEFAULT_WIDTH = 800
WINDOW_DEFAULT_HEIGHT = 900

TOAST_DURATION_DEFAULT = 3000
TOAST_FADE_DURATION = 300
FEEDBACK_RESET_DELAY = 1500
WORKER_WAIT_TIMEOUT = 3000

DOCUMENT_LOAD_DELAY = 1.0
RETRY_DELAY_SECONDS = 1.0

MAX_FILENAME_COUNTER = 1000
MAX_RETRY_COUNT = 3
CONFIG_VERSION = 2
SCAN_BATCH_SIZE = 100
SCAN_CANCEL_WAIT_MS = 200

HWP_PROGIDS = [
    "HWPControl.HwpCtrl.1",
    "HwpObject.HwpObject",
    "HWPFrame.HwpObject",
]
