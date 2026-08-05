"""변환 시작 전 확인 다이얼로그."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from PyQt6.QtWidgets import (
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ...constants import (
    HWP_PERMISSION_HINT,
    PREFLIGHT_DETAIL_MAX_TASKS,
    PREFLIGHT_READ_CHECK_MAX_TASKS,
    PRINT_SETTINGS_NOTICE,
    WINDOWS_PATH_BLOCK_LENGTH,
    WINDOWS_PATH_WARN_LENGTH,
)
from ...models import PlannedConversion
from ...path_utils import (
    check_write_permission,
    is_path_length_blocking,
    is_path_length_risky,
)
from ...services.hwp_converter import get_registered_hwp_progids


class PreflightDialog(QDialog):
    """변환 시작 전 최종 확인 다이얼로그."""

    def __init__(self, plan: PlannedConversion, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("변환 시작 전 확인")
        self.setModal(True)
        self.setMinimumSize(520, 360)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("아래 내용을 확인한 뒤 변환을 시작합니다.")
        title.setProperty("heading", True)
        layout.addWidget(title)

        info_group = QGroupBox("사전 점검")
        info_layout = QVBoxLayout(info_group)
        info_layout.addWidget(QLabel(f"실행 대상 수: {plan.runnable_count}개"))
        info_layout.addWidget(QLabel(f"건너뜀 수: {plan.skipped_count}개"))
        info_layout.addWidget(QLabel(f"덮어쓰기 회피로 이름 변경: {plan.conflict_renamed_count}개"))
        info_layout.addWidget(QLabel(f"선택 형식: {plan.format_type}"))
        info_layout.addWidget(QLabel(f"저장 위치 정책: {plan.output_policy_label}"))
        info_layout.addWidget(QLabel(f"원본 백업: {'사용' if plan.backup_enabled else '사용 안 함'}"))
        info_layout.addWidget(QLabel(f"실패 시 재시도: {plan.retry_count}회"))
        if str(plan.format_type).upper() == "PDF":
            mode = getattr(plan, "pdf_export_mode", "saveas_first") or "saveas_first"
            mode_label = (
                "SaveAs 우선 (용지 품질)"
                if mode == "saveas_first"
                else "PrintToPDFEx 우선 (모아찍기 완화)"
            )
            info_layout.addWidget(QLabel(f"PDF 내보내기: {mode_label}"))
        registered_progids = get_registered_hwp_progids()
        hwp_state = ", ".join(registered_progids) if registered_progids else "레지스트리에서 감지되지 않음"
        info_layout.addWidget(QLabel(f"한글 COM ProgID: {hwp_state}"))
        layout.addWidget(info_group)

        self._deep_read_skipped = 0
        blocking_errors = self._blocking_errors(plan)
        warnings = list(plan.warnings)
        if blocking_errors:
            warnings.extend(f"변환 시작 차단: {error}" for error in blocking_errors)
        long_path_count = self._long_path_warning_count(plan)
        if long_path_count:
            warnings.append(
                f"경로 길이가 {WINDOWS_PATH_WARN_LENGTH}자 이상인 항목이 "
                f"{long_path_count}개 있습니다. 앱이 확장 경로(\\\\?\\)로 재시도합니다. "
                f"{WINDOWS_PATH_BLOCK_LENGTH}자 이상은 시작이 차단될 수 있습니다."
            )
        if self._deep_read_skipped > 0:
            warnings.append(
                f"입력 읽기 심층 검사는 앞쪽 {PREFLIGHT_READ_CHECK_MAX_TASKS}개만 수행했습니다 "
                f"(나머지 {self._deep_read_skipped}개는 변환 시점에 확인)."
            )
        warnings.append(PRINT_SETTINGS_NOTICE)
        warnings.append(HWP_PERMISSION_HINT)
        warnings.append(
            "강제 종료는 앱이 새로 띄운 한글 프로세스에만 적용됩니다. "
            "이미 한글이 실행 중이면 프로세스 추적이 실패할 수 있으니 변환 전 다른 한글 창을 닫는 것을 권장합니다."
        )
        if not warnings:
            warnings = ["추가 경고 없음"]

        warning_group = QGroupBox("주요 경고")
        warning_layout = QVBoxLayout(warning_group)
        warning_text = QTextEdit()
        warning_text.setReadOnly(True)
        warning_text.setPlainText("\n".join(f"- {warning}" for warning in warnings))
        warning_layout.addWidget(warning_text)
        layout.addWidget(warning_group)

        detail_group = QGroupBox("대상 상세")
        detail_layout = QVBoxLayout(detail_group)
        detail_text = QTextEdit()
        detail_text.setReadOnly(True)
        detail_text.setPlainText(self._build_detail_text(plan))
        detail_layout.addWidget(detail_text)
        layout.addWidget(detail_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("취소")
        cancel_btn.setProperty("secondary", True)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        start_btn = QPushButton("변환 시작")
        start_btn.clicked.connect(self.accept)
        if blocking_errors:
            start_btn.setEnabled(False)
            start_btn.setToolTip("입력 파일 또는 출력 폴더의 확정 오류를 먼저 해결해야 합니다")
        btn_layout.addWidget(start_btn)

        layout.addLayout(btn_layout)

    def _build_detail_text(self, plan: PlannedConversion) -> str:
        """대상 상세. 대량 배치에서는 상위 N개만 표시해 UI freeze 를 줄인다.

        상세 목록에서는 존재 여부만 확인하고 open() 심층 읽기는 하지 않는다.
        """
        lines: list[str] = []
        all_tasks = plan.all_tasks
        total = len(all_tasks)
        shown = all_tasks[:PREFLIGHT_DETAIL_MAX_TASKS]

        for index, task in enumerate(shown, start=1):
            try:
                exists = task.input_file.is_file()
            except OSError:
                exists = False
            if task.status == "건너뜀":
                action = "건너뜀"
            else:
                action = "변환 예정"

            lines.append(f"{index}. {task.input_file.name}")
            lines.append(f"   상태: {action}")
            lines.append(f"   입력: {'존재' if exists else '없음'}")
            lines.append(f"   출력: {task.output_file}")
            if task.conflict_original_output_file is not None:
                lines.append(
                    f"   충돌 조정: {task.conflict_original_output_file} -> {task.output_file}"
                )
            if task.status == "건너뜀":
                lines.append(f"   사유: {task.detail}")
            else:
                lines.append(f"   백업: {'사용' if plan.backup_enabled else '사용 안 함'}")
            lines.append("")

        if total > PREFLIGHT_DETAIL_MAX_TASKS:
            lines.append(
                f"... 외 {total - PREFLIGHT_DETAIL_MAX_TASKS}개 항목은 목록에서 생략했습니다 "
                f"(전체 {total}개)."
            )

        return "\n".join(lines).strip() or "대상 없음"

    def _is_readable(self, path: Path) -> bool:
        """심층 읽기 검사 (open + 1바이트). 차단 오류 샘플에만 사용."""
        try:
            if not path.is_file():
                return False
            with path.open("rb") as f:
                f.read(1)
            return True
        except OSError:
            return False

    def _long_path_warning_count(self, plan: PlannedConversion) -> int:
        """입력/출력 경로가 경고 임계를 넘는 실행 대상 개수."""
        count = 0
        for task in plan.tasks:
            if is_path_length_risky(task.input_file) or is_path_length_risky(task.output_file):
                count += 1
        return count

    def _blocking_errors(self, plan: PlannedConversion) -> list[str]:
        """차단 오류.

        - 입력 존재 여부: 실행 대상 전 건 (is_file, 저비용)
        - 읽기 가능 여부: 최대 PREFLIGHT_READ_CHECK_MAX_TASKS 건만 open 심층 검사
        - 출력 쓰기: 출력 폴더 단위 중복 제거 후 검사
        """
        errors: list[str] = []
        checked_output_dirs: set[str] = set()
        deep_read_checked = 0
        deep_read_skipped = 0

        for task in plan.tasks:
            try:
                is_file = task.input_file.is_file()
            except OSError:
                is_file = False

            if not is_file:
                errors.append(f"입력 파일 없음: {task.input_file}")
            elif deep_read_checked < PREFLIGHT_READ_CHECK_MAX_TASKS:
                deep_read_checked += 1
                if not self._is_readable(task.input_file):
                    errors.append(f"입력 파일 읽기 불가: {task.input_file}")
            else:
                deep_read_skipped += 1

            if is_path_length_blocking(task.input_file) or is_path_length_blocking(
                task.output_file
            ):
                errors.append(
                    f"경로가 너무 깁니다 ({WINDOWS_PATH_BLOCK_LENGTH}자 이상): "
                    f"{task.input_file.name}"
                )

            output_dir = task.output_file.parent
            output_key = str(output_dir).lower()
            if output_key in checked_output_dirs:
                continue
            checked_output_dirs.add(output_key)
            if output_dir.exists() and not check_write_permission(output_dir):
                errors.append(f"출력 폴더 쓰기 불가: {output_dir}")

        self._deep_read_skipped = deep_read_skipped
        return errors

