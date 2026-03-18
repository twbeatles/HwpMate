from __future__ import annotations

import subprocess
from pathlib import Path
from typing import List, Optional

from PyQt6.QtWidgets import QFileDialog, QDialog, QFrame, QGroupBox, QHBoxLayout, QLabel, QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget

from ..logging_config import get_logger
from ..models import ConversionTask

logger = get_logger(__name__)

class ResultDialog(QDialog):
    """변환 결과 다이얼로그"""
    
    def __init__(
        self,
        success: int,
        total: int,
        failed_tasks: List[ConversionTask],
        output_paths: Optional[List[str]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("변환 완료")
        self.setMinimumSize(600, 400)
        self.setModal(True)
        
        # 출력 경로 저장 (폴더 열기용)
        self.output_paths: List[str] = list(output_paths or [])
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 25, 25, 25)
        
        # 요약
        summary_frame = QFrame()
        summary_layout = QVBoxLayout(summary_frame)
        
        success_label = QLabel(f"✅ 성공: {success}개")
        success_label.setProperty("heading", True)
        summary_layout.addWidget(success_label)
        
        failed = total - success
        if failed > 0:
            failed_label = QLabel(f"❌ 실패: {failed}개")
            failed_label.setStyleSheet("font-size: 12pt; color: #e94560;")
            summary_layout.addWidget(failed_label)
        
        layout.addWidget(summary_frame)
        
        # 실패 목록
        if failed_tasks:
            failed_group = QGroupBox("실패한 파일")
            failed_layout = QVBoxLayout(failed_group)
            
            text_edit = QTextEdit()
            text_edit.setReadOnly(True)
            
            for task in failed_tasks:
                text_edit.append(f"📄 {task.input_file.name}")
                text_edit.append(f"   오류: {task.error}\n")
            
            failed_layout.addWidget(text_edit)
            layout.addWidget(failed_group)
        
        # 버튼 영역
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        # 실패 목록 내보내기 버튼 (실패한 파일이 있을 때만)
        if failed_tasks:
            self._failed_tasks = failed_tasks  # 내보내기용 저장
            export_btn = QPushButton("📋 실패 목록 저장")
            export_btn.setProperty("secondary", True)
            export_btn.setToolTip("실패한 파일 목록을 텍스트 파일로 저장합니다")
            export_btn.clicked.connect(self._export_failed_list)
            export_btn.setMaximumWidth(150)
            btn_layout.addWidget(export_btn)
        
        # 폴더 열기 버튼
        if success > 0 and self.output_paths:
            open_folder_btn = QPushButton("📂 폴더 열기")
            open_folder_btn.setProperty("secondary", True)
            open_folder_btn.setToolTip("변환된 파일이 있는 폴더를 엽니다")
            open_folder_btn.clicked.connect(self._open_output_folder)
            open_folder_btn.setMaximumWidth(150)
            btn_layout.addWidget(open_folder_btn)
        
        # 닫기 버튼
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        close_btn.setMaximumWidth(150)
        btn_layout.addWidget(close_btn)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
    
    def _export_failed_list(self) -> None:
        """실패 목록 텍스트 파일로 내보내기"""
        from datetime import datetime
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "실패 목록 저장",
            f"변환실패_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "텍스트 파일 (*.txt)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"HWP 변환 실패 목록 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("=" * 50 + "\n\n")
                    for task in self._failed_tasks:
                        f.write(f"파일: {task.input_file}\n")
                        f.write(f"오류: {task.error}\n\n")
                QMessageBox.information(self, "저장 완료", f"실패 목록이 저장되었습니다:\n{file_path}")
            except Exception as e:
                QMessageBox.warning(self, "저장 실패", f"파일 저장 중 오류 발생:\n{e}")
    
    def _open_output_folder(self) -> None:
        """출력 폴더 열기 (파일 선택)"""
        if self.output_paths:
            # 첫 번째 출력 파일
            first_path = Path(self.output_paths[0])
            
            # 파일이 존재하면 /select로 선택하여 열기
            if first_path.exists():
                try:
                    subprocess.run(['explorer', '/select,', str(first_path)], check=False)
                    return
                except Exception as e:
                    logger.debug(f"파일 선택 열기 실패: {e}")

            # 파일 선택 실패 시 폴더만 열기
            folder = first_path.parent if first_path.is_file() else first_path
            if folder.exists():
                try:
                    subprocess.run(['explorer', str(folder)], check=False)
                except Exception as e:
                    logger.error(f"폴더 열기 실패: {e}")
