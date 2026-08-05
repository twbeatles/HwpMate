"""변환 세션 동안 보안 모듈·전면화·자동 클릭 정책을 묶는 가벼운 세션 상태."""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Optional

from ..constants import (
    HWP_FOREGROUND_POLL_MS,
    HWP_FOREGROUND_POLL_MS_RELAXED,
    SECURITY_AUTO_CLICK_COOLDOWN_SECONDS,
    SECURITY_AUTO_CLICK_MAX_PER_SESSION,
)


@dataclass
class HwpSecuritySession:
    """UI 폴링과 워커 엔진 상태를 연결하는 세션 정책."""

    owned_pids: set[int] = field(default_factory=set)
    security_module_registered: bool | None = None
    # initialize 결과 수신 전에는 자동 클릭을 허용하지 않는다.
    engine_status_received: bool = False
    auto_accept_enabled: bool = True
    auto_accept_clicks: int = 0
    last_auto_accept_at: float = 0.0
    snapshot_unreliable: bool = False
    max_auto_accepts: int = SECURITY_AUTO_CLICK_MAX_PER_SESSION
    cooldown_seconds: float = SECURITY_AUTO_CLICK_COOLDOWN_SECONDS

    def reset_runtime(self) -> None:
        """변환 시작 시 런타임 카운터만 초기화 (설정 플래그는 유지)."""
        self.owned_pids = set()
        self.security_module_registered = None
        self.engine_status_received = False
        self.auto_accept_clicks = 0
        self.last_auto_accept_at = 0.0
        self.snapshot_unreliable = False

    def apply_engine_status(self, status: dict) -> None:
        """ConversionWorker.engine_status_updated 페이로드 반영."""
        registered = status.get("security_module_registered")
        if registered is None or isinstance(registered, bool):
            self.security_module_registered = registered

        pids = status.get("owned_pids") or []
        try:
            self.owned_pids = {int(pid) for pid in pids if int(pid) > 0}
        except (TypeError, ValueError):
            self.owned_pids = set()

        self.snapshot_unreliable = bool(status.get("snapshot_unreliable", False))
        self.engine_status_received = True

    def allows_window_control(self) -> bool:
        """소유 PID가 확정된 뒤에만 다른 한글 세션을 건드리지 않고 창 조작을 허용한다."""
        return self.engine_status_received and bool(self.owned_pids)

    def target_pids(self) -> Optional[set[int]]:
        """
        전면화/자동 클릭 대상 PID.

        - 소유 PID가 있으면 그 집합만 반환한다.
        - 엔진 상태 수신 후 소유 PID가 비어 있으면 빈 set (전역 HWP 조작 금지).
        - 엔진 상태 수신 전에는 None (호출측에서 폴링 본문 no-op 권장).
        """
        if not self.engine_status_received:
            return None
        if self.owned_pids:
            return set(self.owned_pids)
        return set()

    def should_auto_accept(self) -> bool:
        if not self.auto_accept_enabled:
            return False
        # 워커 initialize 결과 수신 전: 연결 창에서 전역 자동 클릭 금지
        if not self.engine_status_received:
            return False
        # 소유 PID 없으면 다른 한글 세션 오클릭 방지
        if not self.owned_pids:
            return False
        # 모듈 등록 성공(True) 또는 미확인(None) 이면 자동 클릭 안 함.
        # 실패(False) 일 때만 「모두 허용」 보조 클릭.
        if self.security_module_registered is not False:
            return False
        if self.auto_accept_clicks >= self.max_auto_accepts:
            return False
        now = time.monotonic()
        if self.last_auto_accept_at and (now - self.last_auto_accept_at) < self.cooldown_seconds:
            return False
        return True

    def note_auto_accept(self) -> None:
        self.auto_accept_clicks += 1
        self.last_auto_accept_at = time.monotonic()

    def poll_interval_ms(self) -> int:
        if self.security_module_registered is True:
            return HWP_FOREGROUND_POLL_MS_RELAXED
        return HWP_FOREGROUND_POLL_MS
