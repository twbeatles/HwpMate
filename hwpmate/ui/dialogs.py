from __future__ import annotations

import csv
import json
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional, TextIO, cast

from PyQt6.QtWidgets import (
    QFileDialog,
    QDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..logging_config import get_logger
from ..constants import (
    HWP_PERMISSION_HINT,
    PREFLIGHT_DETAIL_MAX_TASKS,
    PREFLIGHT_READ_CHECK_MAX_TASKS,
    PRINT_SETTINGS_NOTICE,
)
from ..models import ConversionSummary, ConversionTask, PlannedConversion
from ..path_utils import check_write_permission
from ..services.hwp_converter import get_registered_hwp_progids

logger = get_logger(__name__)


def _write_text_file_atomically(
    path: Path,
    writer: Callable[[TextIO], None],
    *,
    encoding: str,
    newline: str | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding=encoding,
            newline=newline,
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as f:
            temp_path = Path(f.name)
            writer(cast(TextIO, f))
        temp_path.replace(path)
    except Exception:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
        raise


def write_failed_list(path: Path, failed_tasks: list[ConversionTask]) -> None:
    def writer(f: TextIO) -> None:
        f.write(f"HWP 변환 실패 목록 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 50 + "\n\n")
        for task in failed_tasks:
            f.write(f"파일: {task.input_file}\n")
            f.write(f"오류: {task.detail}\n\n")

    _write_text_file_atomically(path, writer, encoding="utf-8")


def write_results_csv(path: Path, summary: ConversionSummary) -> None:
    def writer(f: TextIO) -> None:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "input_file",
                "output_file",
                "status",
                "detail",
                "retry_count",
                "backup_file",
                "backup_error",
                "created_files",
                "output_size",
                "output_mtime",
                "save_format",
                "export_method",
                "progid_used",
            ],
        )
        writer.writeheader()
        for task in summary.sorted_tasks():
            writer.writerow(task.to_record())

    _write_text_file_atomically(path, writer, encoding="utf-8-sig", newline="")


def write_results_json(path: Path, summary: ConversionSummary) -> None:
    def writer(f: TextIO) -> None:
        json.dump(summary.to_json_dict(), f, ensure_ascii=False, indent=2)

    _write_text_file_atomically(path, writer, encoding="utf-8")

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

            output_dir = task.output_file.parent
            output_key = str(output_dir).lower()
            if output_key in checked_output_dirs:
                continue
            checked_output_dirs.add(output_key)
            if output_dir.exists() and not check_write_permission(output_dir):
                errors.append(f"출력 폴더 쓰기 불가: {output_dir}")

        self._deep_read_skipped = deep_read_skipped
        return errors


class ResultDialog(QDialog):
    """변환 결과 다이얼로그"""

    def __init__(
        self,
        summary: ConversionSummary,
        parent: Optional[QWidget] = None,
        *,
        on_retry_failed: Callable[[list[ConversionTask]], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self.summary = summary
        self._on_retry_failed = on_retry_failed
        self.setWindowTitle("변환 결과")
        self.setMinimumSize(640, 480)
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 25, 25, 25)

        summary_frame = QFrame()
        summary_layout = QVBoxLayout(summary_frame)
        summary_layout.addWidget(self._make_heading(f"✅ 성공: {summary.success_count}개"))
        summary_layout.addWidget(QLabel(f"❌ 실패: {summary.failed_count}개"))
        summary_layout.addWidget(QLabel(f"⏭️ 건너뜀: {summary.skipped_count}개"))
        summary_layout.addWidget(QLabel(f"🛑 취소됨: {summary.canceled_count}개"))
        summary_layout.addWidget(QLabel(f"📄 전체 요청: {summary.total_requested}개"))
        if summary.elapsed_seconds is not None:
            summary_layout.addWidget(QLabel(f"⏱️ 소요 시간: {summary.elapsed_seconds:.1f}초"))
        layout.addWidget(summary_frame)

        if summary.warnings:
            warning_group = QGroupBox("경고")
            warning_layout = QVBoxLayout(warning_group)
            warning_text = QTextEdit()
            warning_text.setReadOnly(True)
            warning_text.setPlainText("\n".join(f"- {warning}" for warning in summary.warnings))
            warning_layout.addWidget(warning_text)
            layout.addWidget(warning_group)

        if summary.failed_tasks:
            failed_group = QGroupBox("실패한 파일")
            failed_layout = QVBoxLayout(failed_group)
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            for task in summary.failed_tasks:
                text_edit.append(f"📄 {task.input_file.name}")
                text_edit.append(f"   오류: {task.detail}\n")
            failed_layout.addWidget(text_edit)
            layout.addWidget(failed_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        if summary.failed_tasks:
            export_btn = QPushButton("📋 실패 목록 저장")
            export_btn.setProperty("secondary", True)
            export_btn.setToolTip("실패한 파일 목록을 텍스트 파일로 저장합니다")
            export_btn.clicked.connect(self._export_failed_list)
            export_btn.setMaximumWidth(150)
            btn_layout.addWidget(export_btn)

            if self._on_retry_failed is not None:
                retry_btn = QPushButton("🔁 실패 항목 재변환")
                retry_btn.setProperty("secondary", True)
                retry_btn.setToolTip("실패한 파일만 다시 변환합니다")
                retry_btn.clicked.connect(self._retry_failed)
                retry_btn.setMaximumWidth(160)
                btn_layout.addWidget(retry_btn)

        save_results_btn = QPushButton("💾 결과 저장")
        save_results_btn.setProperty("secondary", True)
        save_results_btn.setToolTip("전체 결과를 CSV 또는 JSON으로 저장합니다")
        save_results_btn.clicked.connect(self._save_results)
        save_results_btn.setMaximumWidth(150)
        btn_layout.addWidget(save_results_btn)

        if summary.output_paths:
            open_folder_btn = QPushButton("📂 폴더 열기")
            open_folder_btn.setProperty("secondary", True)
            open_folder_btn.setToolTip("변환된 파일이 있는 폴더를 엽니다")
            open_folder_btn.clicked.connect(self._open_output_folder)
            open_folder_btn.setMaximumWidth(150)
            btn_layout.addWidget(open_folder_btn)

        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        close_btn.setMaximumWidth(150)
        btn_layout.addWidget(close_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

    def _make_heading(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setProperty("heading", True)
        return label

    def _retry_failed(self) -> None:
        if self._on_retry_failed is None or not self.summary.failed_tasks:
            return
        callback = self._on_retry_failed
        failed = list(self.summary.failed_tasks)
        self.accept()
        callback(failed)

    def _export_failed_list(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "실패 목록 저장",
            f"변환실패_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "텍스트 파일 (*.txt)",
        )

        if file_path:
            try:
                write_failed_list(Path(file_path), self.summary.failed_tasks)
                QMessageBox.information(self, "저장 완료", f"실패 목록이 저장되었습니다:\n{file_path}")
            except Exception as e:
                QMessageBox.warning(self, "저장 실패", f"파일 저장 중 오류 발생:\n{e}")

    def _save_results(self) -> None:
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "결과 저장",
            f"변환결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV 파일 (*.csv);;JSON 파일 (*.json)",
        )

        if not file_path:
            return

        output_path = Path(file_path)
        if output_path.suffix.lower() not in {".csv", ".json"}:
            output_path = output_path.with_suffix(".json" if "JSON" in selected_filter else ".csv")

        try:
            if output_path.suffix.lower() == ".json":
                write_results_json(output_path, self.summary)
            else:
                write_results_csv(output_path, self.summary)
            QMessageBox.information(self, "저장 완료", f"변환 결과가 저장되었습니다:\n{output_path}")
        except Exception as e:
            QMessageBox.warning(self, "저장 실패", f"결과 저장 중 오류 발생:\n{e}")

    def _open_output_folder(self) -> None:
        if self.summary.output_paths:
            first_path = Path(self.summary.output_paths[0])

            if first_path.exists():
                try:
                    subprocess.run(["explorer", "/select,", str(first_path)], check=False)
                    return
                except Exception as e:
                    logger.debug(f"파일 선택 열기 실패: {e}")

            folder = first_path.parent if first_path.is_file() else first_path
            if folder.exists():
                try:
                    subprocess.run(["explorer", str(folder)], check=False)
                except Exception as e:
                    logger.error(f"폴더 열기 실패: {e}")
