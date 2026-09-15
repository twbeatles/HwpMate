"""한글 변환 중 확인 대화상자 자동 응답 (앱 소유 PID 한정).

한글 2022(12.0) 는 DOCX·RTF 등 호환 형식 저장 시
「호환 문서 — 이 파일을 저장하면 문서 내용의 배치가 변경될 수 있습니다. 저장을 계속할까요?」
WPF 메시지 상자(MessageBoxImpl, 버튼 「계속 : ALT+Y」/「취소 : ALT+N」)를 띄운다.
이 창은 SetMessageBoxMode 로 억제되지 않고 Win32 버튼 컨트롤도 없어 BM_CLICK 이 불가하다.
실측 결과 창 핸들에 Y 키 메시지를 보내면 「계속」, N 이면 「취소」로 처리된다.

안전 정책:
- 소유 PID 가 확정된 경우에만 동작한다 (다른 한글 편집 세션 조작 금지).
- 제목이 허용 목록과 정확히 일치하는 창에만 응답한다.
- 같은 창에는 쿨다운 동안 재전송하지 않는다.
"""

from __future__ import annotations

import ctypes
import threading
import time
from typing import Callable, Optional, Set

from ..logging_config import get_logger

logger = get_logger(__name__)

# 「계속(Y)」 로 응답해도 변환 의도와 일치하는 확인 창 제목 (정확 일치)
COMPAT_DIALOG_TITLES = frozenset({"호환 문서"})

_WM_KEYDOWN = 0x0100
_WM_KEYUP = 0x0101
_WM_CHAR = 0x0102
_VK_Y = 0x59
_SCAN_Y = 0x15

DEFAULT_POLL_INTERVAL_SECONDS = 0.2
DEFAULT_PER_WINDOW_COOLDOWN_SECONDS = 1.0


def _post_continue_key(hwnd: int) -> bool:
    user32 = ctypes.windll.user32
    down_lparam = 1 | (_SCAN_Y << 16)
    up_lparam = down_lparam | 0xC0000000
    ok_down = bool(user32.PostMessageW(hwnd, _WM_KEYDOWN, _VK_Y, down_lparam))
    user32.PostMessageW(hwnd, _WM_CHAR, ord("y"), down_lparam)
    ok_up = bool(user32.PostMessageW(hwnd, _WM_KEYUP, _VK_Y, up_lparam))
    return ok_down and ok_up


def _is_window_visible(hwnd: int) -> bool:
    return bool(ctypes.windll.user32.IsWindowVisible(hwnd))


def find_hwp_compat_dialogs(pids: Optional[Set[int]]) -> list[int]:
    """소유 PID 의 보이는 top-level 창 중 호환 문서 확인 창 HWND 목록."""
    if not pids:
        return []
    from hwpmate import windows_integration as api
    from . import hwp_dialog_responder as module

    found: list[int] = []
    for hwnd in api._list_top_level_hwnds_for_pids(set(pids), security_dialogs_only=False):
        try:
            if not module._is_window_visible(hwnd):
                continue
            if api._window_title(hwnd) in COMPAT_DIALOG_TITLES:
                found.append(hwnd)
        except Exception:
            continue
    return found


def respond_hwp_compat_dialogs(
    pids: Optional[Set[int]],
    skip_hwnds: frozenset[int] = frozenset(),
) -> list[int]:
    """호환 문서 확인 창에 「계속(Y)」 을 보낸다. 응답한 HWND 목록을 반환."""
    responded: list[int] = []
    from . import hwp_dialog_responder as module

    for hwnd in module.find_hwp_compat_dialogs(pids):
        if hwnd in skip_hwnds:
            continue
        try:
            if module._post_continue_key(hwnd):
                responded.append(hwnd)
        except Exception as e:
            logger.debug(f"호환 문서 확인 창 응답 실패: hwnd={hwnd}, {e}")
    return responded


class HwpDialogAutoResponder:
    """COM 호출이 블로킹되는 동안 백그라운드에서 확인 창에 응답하는 스레드.

    Win32 메시지만 사용하므로 COM apartment 와 무관하다.
    """

    def __init__(
        self,
        pids_provider: Callable[[], Set[int]],
        *,
        poll_interval: float = DEFAULT_POLL_INTERVAL_SECONDS,
        cooldown: float = DEFAULT_PER_WINDOW_COOLDOWN_SECONDS,
        responder: Callable[[Optional[Set[int]], frozenset[int]], list[int]] | None = None,
    ) -> None:
        self._pids_provider = pids_provider
        self._poll_interval = max(0.02, float(poll_interval))
        self._cooldown = max(0.0, float(cooldown))
        self._responder = responder or respond_hwp_compat_dialogs
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._last_response_at: dict[int, float] = {}
        self._lock = threading.Lock()
        self.response_count = 0

    def _tick(self) -> None:
        pids = set(self._pids_provider() or set())
        if not pids:
            return
        now = time.monotonic()
        with self._lock:
            recent = frozenset(
                hwnd
                for hwnd, at in self._last_response_at.items()
                if now - at < self._cooldown
            )
        responded = self._responder(pids, recent)
        if not responded:
            return
        with self._lock:
            for hwnd in responded:
                self._last_response_at[hwnd] = now
            self.response_count += len(responded)
        logger.info(
            f"한글 「호환 문서」 확인 창 자동 계속: pids={sorted(pids)}, 응답={len(responded)}"
        )

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                self._tick()
            except Exception as e:
                logger.debug(f"확인 창 자동 응답 루프 오류(무시): {e}")
            self._stop_event.wait(self._poll_interval)

    def start(self) -> "HwpDialogAutoResponder":
        if self._thread is not None:
            return self
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._run,
            name="HwpDialogAutoResponder",
            daemon=True,
        )
        self._thread.start()
        return self

    def stop(self) -> None:
        self._stop_event.set()
        thread = self._thread
        self._thread = None
        if thread is not None and thread is not threading.current_thread():
            thread.join(timeout=2.0)

    def __enter__(self) -> "HwpDialogAutoResponder":
        return self.start()

    def __exit__(self, *_exc: object) -> None:
        self.stop()
