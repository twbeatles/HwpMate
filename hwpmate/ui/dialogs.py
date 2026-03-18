from __future__ import annotations

import csv
import json
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional

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
from ..models import ConversionSummary, ConversionTask, PlannedConversion

logger = get_logger(__name__)


def write_failed_list(path: Path, failed_tasks: list[ConversionTask]) -> None:
    with path.open("w", encoding="utf-8") as f:
        f.write(f"HWP 변환 실패 목록 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 50 + "\n\n")
        for task in failed_tasks:
            f.write(f"파일: {task.input_file}\n")
            f.write(f"오류: {task.detail}\n\n")


def write_results_csv(path: Path, summary: ConversionSummary) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["input_file", "output_file", "status", "detail"])
        writer.writeheader()
        for task in summary.sorted_tasks():
            writer.writerow(task.to_record())


def write_results_json(path: Path, summary: ConversionSummary) -> None:
    with path.open("w", encoding="utf-8") as f:
        json.dump(summary.to_json_dict(), f, ensure_ascii=False, indent=2)


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
        layout.addWidget(info_group)

        warnings = list(plan.warnings)
        if not warnings:
            warnings = ["추가 경고 없음"]

        warning_group = QGroupBox("주요 경고")
        warning_layout = QVBoxLayout(warning_group)
        warning_text = QTextEdit()
        warning_text.setReadOnly(True)
        warning_text.setPlainText("\n".join(f"- {warning}" for warning in warnings))
        warning_layout.addWidget(warning_text)
        layout.addWidget(warning_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("취소")
        cancel_btn.setProperty("secondary", True)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        start_btn = QPushButton("변환 시작")
        start_btn.clicked.connect(self.accept)
        btn_layout.addWidget(start_btn)

        layout.addLayout(btn_layout)


class ResultDialog(QDialog):
    """변환 결과 다이얼로그"""

    def __init__(
        self,
        summary: ConversionSummary,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.summary = summary
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
