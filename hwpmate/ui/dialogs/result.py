"""변환 결과 다이얼로그."""

from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

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

from ...logging_config import get_logger
from ...models import ConversionSummary, ConversionTask
from .atomic_io import write_failed_list, write_results_csv, write_results_json

logger = get_logger(__name__)


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
