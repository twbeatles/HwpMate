from __future__ import annotations

from .models import FormatSpec

VERSION = "9.0"
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

# Open 직후 문서 로드 대기 (초). 과도한 대량 배치 지연을 줄이기 위해 0.5s.
DOCUMENT_LOAD_DELAY = 0.5
RETRY_DELAY_SECONDS = 1.0

MAX_FILENAME_COUNTER = 1000
MAX_RETRY_COUNT = 3
CONFIG_VERSION = 3
SCAN_BATCH_SIZE = 100
SCAN_CANCEL_WAIT_MS = 2000
FOLDER_SCAN_WAIT_MS = 120_000
# 폴더 미리보기 캐시 최대 유효 시간 (초). 초과 시 변환 전 재스캔.
FOLDER_SCAN_CACHE_MAX_AGE_SECONDS = 300
# 캐시 신선도 샘플 검사 개수 (앞·뒤·중간 분산)
FOLDER_SCAN_CACHE_SAMPLE_SIZE = 24

# 사전 점검(Preflight) UI 부하 상한
# 상세 목록에 표시할 최대 작업 수 (초과 분은 요약 한 줄)
PREFLIGHT_DETAIL_MAX_TASKS = 80
# 읽기 가능 여부(open) 심층 검사 상한 — 존재 여부는 전 건 검사
PREFLIGHT_READ_CHECK_MAX_TASKS = 48

HWP_PROGIDS = [
    "HWPControl.HwpCtrl.1",
    "HwpObject.HwpObject",
    "HWPFrame.HwpObject",
]

# 한글 COM 허용/보안 창이 작업 표시줄 뒤로 가려질 때 사용자 안내
HWP_PERMISSION_HINT = (
    "한글 허용·보안 확인 창이 뒤에 가려질 수 있습니다. "
    "작업 표시줄의 한글 창을 확인해 주세요."
)

# PDF·이미지 등 인쇄 경로 산출물 안내 (앱은 1쪽씩 기본 인쇄로 best-effort 초기화)
PRINT_SETTINGS_NOTICE = (
    "PDF·이미지 변환 시 앱이 인쇄 방식을 기본(1쪽씩)으로 맞추려 시도합니다. "
    "문서 편집 용지(대부분 A4)는 그대로 두고, 모아찍기 등만 해제하는 방향입니다. "
    "기본 PDF 경로는 SaveAs(용지 품질 우선)이며, 설정 pdf_export_mode 로 "
    "PrintToPDFEx 우선(모아찍기 완화)을 고를 수 있습니다. "
    "물리 프린터 Print 실행은 사용하지 않습니다. "
    "한글 버전·환경에 따라 일부 설정이 남을 수 있습니다."
)

# 연결 중 한글 창 전면화 폴링 간격 (ms) — Toolhelp 기반이라 콘솔 플래시 없음
HWP_FOREGROUND_POLL_MS = 800
# 보안 모듈 등록 성공 후 전면화만 유지할 때 완화 간격
HWP_FOREGROUND_POLL_MS_RELAXED = 2500

# 「모두 허용」 자동 클릭 보호
SECURITY_AUTO_CLICK_COOLDOWN_SECONDS = 1.5
SECURITY_AUTO_CLICK_MAX_PER_SESSION = 40

# 한컴 FilePathCheckerModuleExample.dll 번들 무결성 (SHA-256)
# 파일 교체 시 이 값을 함께 갱신해야 한다.
SECURITY_MODULE_DLL_SHA256 = (
    "9ac5b97c47ac8aed1e8bca27a3eef39411361d8f68c262509f0c40a8f9d21bb6"
)
