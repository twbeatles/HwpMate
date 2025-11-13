"""
HWP/HWPX 변환기 v6.0 - 완전 재설계
안정성과 사용성에 초점을 맞춘 새로운 버전
"""

import os
import sys
import json
import ctypes
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# pywin32 import (COM 사용)
try:
    import pythoncom
    import win32com.client
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    messagebox.showerror("오류", "pywin32 라이브러리가 필요합니다.\n\npip install pywin32")
    sys.exit(1)

# 설정 파일
CONFIG_FILE = Path.home() / ".hwp_converter_config.json"

# 한글 ProgID 목록 (우선순위 순)
HWP_PROGIDS = [
    "HWPControl.HwpCtrl.1",
    "HwpObject.HwpObject",
    "HWPFrame.HwpObject",
]


def is_admin():
    """관리자 권한 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def load_config():
    """설정 로드"""
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except:
        pass
    return {}


def save_config(config):
    """설정 저장"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except:
        pass


class HWPConverter:
    """한글 변환 엔진"""
    
    def __init__(self):
        self.hwp = None
        self.progid_used = None
        self.is_initialized = False
        
    def initialize(self):
        """COM 초기화 및 한글 객체 생성"""
        if self.is_initialized:
            return True
            
        try:
            pythoncom.CoInitialize()
        except Exception as e:
            print(f"CoInitialize 오류 (무시 가능): {e}")
        
        errors = []
        for progid in HWP_PROGIDS:
            try:
                self.hwp = win32com.client.Dispatch(progid)
                self.progid_used = progid
                
                # 한글 설정
                try:
                    self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                except:
                    pass  # 일부 버전에서는 지원하지 않음
                
                self.hwp.SetMessageBoxMode(0x00000001)  # 메시지 박스 비활성화
                self.is_initialized = True
                return True
                
            except Exception as e:
                errors.append(f"{progid}: {str(e)}")
                continue
        
        # 모든 시도 실패
        error_detail = "\n".join(errors)
        raise Exception(f"한글 COM 객체 생성 실패\n\n시도한 ProgID:\n{error_detail}")
    
    def convert_file(self, input_path, output_path, format_type="PDF"):
        """단일 파일 변환
        
        Returns:
            (성공여부, 오류메시지) 튜플
        """
        if not self.is_initialized:
            return False, "한글 객체가 초기화되지 않았습니다"
        
        try:
            # 파일 열기
            input_str = str(input_path)
            output_str = str(output_path)
            
            self.hwp.Open(input_str, "HWP", "forceopen:true")
            
            # 저장
            save_format = "PDF" if format_type == "PDF" else "HWPX"
            self.hwp.SaveAs(output_str, save_format)
            
            # 문서 닫기
            self.hwp.Clear(option=1)
            
            return True, None
            
        except Exception as e:
            error_msg = str(e)
            # 문서 닫기 시도
            try:
                self.hwp.Clear(option=1)
            except:
                pass
            
            return False, error_msg
    
    def cleanup(self):
        """정리"""
        if self.hwp and self.is_initialized:
            try:
                self.hwp.Clear(3)  # 모든 문서 닫기
            except:
                pass
            
            try:
                self.hwp.Quit()
            except:
                pass
            
            self.hwp = None
            self.is_initialized = False
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass


class ConversionTask:
    """변환 작업 정보"""
    
    def __init__(self, input_file, output_file):
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.status = "대기"  # 대기, 진행중, 성공, 실패
        self.error = None


class HWPConverterGUI:
    """메인 GUI"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("HWP 변환기 v6.0 - 재설계판")
        self.root.geometry("700x750")
        self.root.resizable(True, True)
        
        # 변수
        self.config = load_config()
        self.tasks = []  # 변환 작업 목록
        self.is_converting = False
        self.cancel_requested = False
        
        # 설정 변수
        self.mode_var = tk.StringVar(value=self.config.get("mode", "folder"))
        self.format_var = tk.StringVar(value=self.config.get("format", "PDF"))
        self.include_sub_var = tk.BooleanVar(value=self.config.get("include_sub", True))
        self.same_location_var = tk.BooleanVar(value=self.config.get("same_location", True))
        self.overwrite_var = tk.BooleanVar(value=self.config.get("overwrite", False))
        
        # GUI 생성
        self._create_ui()
        
        # 초기 상태 설정
        self._update_mode_ui()
    
    def _create_ui(self):
        """UI 생성"""
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === 1. 모드 선택 ===
        mode_frame = ttk.LabelFrame(main_frame, text="변환 모드", padding=10)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Radiobutton(
            mode_frame, 
            text="폴더 일괄 변환 (폴더 내 모든 파일)", 
            variable=self.mode_var, 
            value="folder",
            command=self._update_mode_ui
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(
            mode_frame, 
            text="파일 개별 선택 (원하는 파일만)", 
            variable=self.mode_var, 
            value="files",
            command=self._update_mode_ui
        ).pack(anchor=tk.W, pady=2)
        
        # === 2. 입력 영역 ===
        input_frame = ttk.LabelFrame(main_frame, text="입력", padding=10)
        input_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 폴더 모드 UI
        self.folder_frame = ttk.Frame(input_frame)
        
        ttk.Label(self.folder_frame, text="폴더 경로:").pack(anchor=tk.W, pady=(0, 5))
        
        folder_entry_frame = ttk.Frame(self.folder_frame)
        folder_entry_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.folder_entry = ttk.Entry(folder_entry_frame, state="readonly")
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(
            folder_entry_frame, 
            text="찾아보기...", 
            command=self._select_folder,
            width=12
        ).pack(side=tk.LEFT)
        
        self.include_sub_check = ttk.Checkbutton(
            self.folder_frame,
            text="하위 폴더 포함",
            variable=self.include_sub_var
        )
        self.include_sub_check.pack(anchor=tk.W)
        
        # 파일 모드 UI
        self.files_frame = ttk.Frame(input_frame)
        
        btn_frame = ttk.Frame(self.files_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(
            btn_frame,
            text="파일 추가...",
            command=self._add_files
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            btn_frame,
            text="선택 제거",
            command=self._remove_selected
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            btn_frame,
            text="전체 제거",
            command=self._clear_all
        ).pack(side=tk.LEFT)
        
        # 파일 리스트
        list_frame = ttk.Frame(self.files_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            height=8
        )
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_listbox.yview)
        
        # === 3. 출력 설정 ===
        output_frame = ttk.LabelFrame(main_frame, text="출력", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.same_location_check = ttk.Checkbutton(
            output_frame,
            text="입력 파일과 같은 위치에 저장",
            variable=self.same_location_var,
            command=self._update_output_ui
        )
        self.same_location_check.pack(anchor=tk.W, pady=(0, 10))
        
        ttk.Label(output_frame, text="저장 폴더:").pack(anchor=tk.W, pady=(0, 5))
        
        output_entry_frame = ttk.Frame(output_frame)
        output_entry_frame.pack(fill=tk.X)
        
        self.output_entry = ttk.Entry(output_entry_frame, state="readonly")
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.output_button = ttk.Button(
            output_entry_frame,
            text="찾아보기...",
            command=self._select_output,
            width=12
        )
        self.output_button.pack(side=tk.LEFT)
        
        # === 4. 변환 옵션 ===
        options_frame = ttk.LabelFrame(main_frame, text="변환 옵션", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 변환 형식
        format_frame = ttk.Frame(options_frame)
        format_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(format_frame, text="변환 형식:").pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Radiobutton(
            format_frame,
            text="PDF",
            variable=self.format_var,
            value="PDF"
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Radiobutton(
            format_frame,
            text="HWPX",
            variable=self.format_var,
            value="HWPX"
        ).pack(side=tk.LEFT)
        
        # 덮어쓰기
        ttk.Checkbutton(
            options_frame,
            text="기존 파일 덮어쓰기 (체크 해제 시 번호 자동 추가)",
            variable=self.overwrite_var
        ).pack(anchor=tk.W)
        
        # === 5. 실행 버튼 ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_button = ttk.Button(
            button_frame,
            text="변환 시작",
            command=self._start_conversion
        )
        self.start_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.cancel_button = ttk.Button(
            button_frame,
            text="취소",
            command=self._cancel_conversion,
            state=tk.DISABLED
        )
        self.cancel_button.pack(side=tk.LEFT)
        
        # === 6. 진행 상태 ===
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상태", padding=10)
        progress_frame.pack(fill=tk.X)
        
        self.status_label = ttk.Label(progress_frame, text="준비됨")
        self.status_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_label = ttk.Label(progress_frame, text="0 / 0")
        self.progress_label.pack(anchor=tk.W)
        
        # 초기 UI 상태
        self._update_output_ui()
    
    def _update_mode_ui(self):
        """모드에 따라 UI 업데이트"""
        mode = self.mode_var.get()
        
        if mode == "folder":
            self.files_frame.pack_forget()
            self.folder_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.folder_frame.pack_forget()
            self.files_frame.pack(fill=tk.BOTH, expand=True)
    
    def _update_output_ui(self):
        """출력 폴더 UI 상태 업데이트"""
        if self.same_location_var.get():
            self.output_entry.config(state=tk.DISABLED)
            self.output_button.config(state=tk.DISABLED)
        else:
            self.output_entry.config(state="readonly")
            self.output_button.config(state=tk.NORMAL)
    
    def _select_folder(self):
        """폴더 선택"""
        initial = self.config.get("last_folder", "")
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            self.folder_entry.config(state=tk.NORMAL)
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            self.folder_entry.config(state="readonly")
            self.config["last_folder"] = folder
    
    def _select_output(self):
        """출력 폴더 선택"""
        initial = self.config.get("last_output", "")
        folder = filedialog.askdirectory(initialdir=initial)
        if folder:
            self.output_entry.config(state=tk.NORMAL)
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)
            self.output_entry.config(state="readonly")
            self.config["last_output"] = folder
    
    def _add_files(self):
        """파일 추가"""
        format_type = self.format_var.get()
        
        if format_type == "PDF":
            filetypes = [("한글 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")]
        else:
            filetypes = [("HWP 파일", "*.hwp"), ("모든 파일", "*.*")]
        
        files = filedialog.askopenfilenames(
            title="파일 선택",
            filetypes=filetypes
        )
        
        if files:
            current = list(self.files_listbox.get(0, tk.END))
            added = 0
            
            for file in files:
                if file not in current:
                    self.files_listbox.insert(tk.END, file)
                    added += 1
            
            if added > 0:
                messagebox.showinfo("완료", f"{added}개 파일이 추가되었습니다.")
    
    def _remove_selected(self):
        """선택된 파일 제거"""
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("경고", "제거할 파일을 선택하세요.")
            return
        
        for idx in reversed(selection):
            self.files_listbox.delete(idx)
    
    def _clear_all(self):
        """전체 파일 제거"""
        count = self.files_listbox.size()
        if count == 0:
            return
        
        if messagebox.askyesno("확인", f"{count}개 파일을 모두 제거하시겠습니까?"):
            self.files_listbox.delete(0, tk.END)
    
    def _collect_tasks(self):
        """변환 작업 목록 생성"""
        tasks = []
        mode = self.mode_var.get()
        format_type = self.format_var.get()
        output_ext = ".pdf" if format_type == "PDF" else ".hwpx"
        
        # 입력 파일 목록 수집
        if mode == "folder":
            folder_path = self.folder_entry.get()
            if not folder_path:
                raise ValueError("폴더를 선택하세요.")
            
            folder = Path(folder_path)
            if not folder.exists():
                raise ValueError("폴더가 존재하지 않습니다.")
            
            # 검색할 확장자
            if format_type == "PDF":
                patterns = ["*.hwp", "*.hwpx"]
            else:
                patterns = ["*.hwp"]
            
            # 파일 검색
            input_files = []
            if self.include_sub_var.get():
                for pattern in patterns:
                    input_files.extend(folder.rglob(pattern))
            else:
                for pattern in patterns:
                    input_files.extend(folder.glob(pattern))
            
            if not input_files:
                raise ValueError("변환할 파일이 없습니다.")
            
            # 작업 생성
            for input_file in input_files:
                if self.same_location_var.get():
                    output_file = input_file.parent / (input_file.stem + output_ext)
                else:
                    output_folder = Path(self.output_entry.get())
                    if not output_folder:
                        raise ValueError("출력 폴더를 선택하세요.")
                    
                    # 상대 경로 유지
                    rel_path = input_file.relative_to(folder)
                    output_file = output_folder / rel_path.parent / (input_file.stem + output_ext)
                
                tasks.append(ConversionTask(input_file, output_file))
        
        else:  # files mode
            file_list = list(self.files_listbox.get(0, tk.END))
            if not file_list:
                raise ValueError("파일을 추가하세요.")
            
            for file_path in file_list:
                input_file = Path(file_path)
                
                if self.same_location_var.get():
                    output_file = input_file.parent / (input_file.stem + output_ext)
                else:
                    output_folder = Path(self.output_entry.get())
                    if not output_folder:
                        raise ValueError("출력 폴더를 선택하세요.")
                    
                    output_file = output_folder / (input_file.stem + output_ext)
                
                tasks.append(ConversionTask(input_file, output_file))
        
        return tasks
    
    def _start_conversion(self):
        """변환 시작"""
        try:
            # 작업 목록 생성
            self.tasks = self._collect_tasks()
            
            # 덮어쓰기 확인
            if not self.overwrite_var.get():
                self._adjust_output_paths()
            
            # 설정 저장
            self._save_settings()
            
            # UI 업데이트
            self._set_converting_state(True)
            
            # 스레드 시작
            thread = threading.Thread(target=self._conversion_worker, daemon=True)
            thread.start()
            
        except ValueError as e:
            messagebox.showwarning("경고", str(e))
        except Exception as e:
            messagebox.showerror("오류", f"오류 발생: {e}")
    
    def _adjust_output_paths(self):
        """출력 경로 조정 (덮어쓰기 방지)"""
        for task in self.tasks:
            if task.output_file.exists():
                counter = 1
                stem = task.output_file.stem
                ext = task.output_file.suffix
                parent = task.output_file.parent
                
                while True:
                    new_name = f"{stem} ({counter}){ext}"
                    new_path = parent / new_name
                    if not new_path.exists():
                        task.output_file = new_path
                        break
                    counter += 1
    
    def _save_settings(self):
        """설정 저장"""
        self.config["mode"] = self.mode_var.get()
        self.config["format"] = self.format_var.get()
        self.config["include_sub"] = self.include_sub_var.get()
        self.config["same_location"] = self.same_location_var.get()
        self.config["overwrite"] = self.overwrite_var.get()
        save_config(self.config)
    
    def _set_converting_state(self, converting):
        """변환 중 상태 설정"""
        self.is_converting = converting
        
        if converting:
            self.start_button.config(state=tk.DISABLED)
            self.cancel_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.DISABLED)
    
    def _cancel_conversion(self):
        """변환 취소"""
        if messagebox.askyesno("확인", "변환을 취소하시겠습니까?"):
            self.cancel_requested = True
    
    def _conversion_worker(self):
        """변환 작업 수행 (별도 스레드)"""
        converter = HWPConverter()
        success_count = 0
        total = len(self.tasks)
        
        try:
            # 초기화
            self._update_status("한글 프로그램 연결 중...")
            converter.initialize()
            
            status_msg = f"연결 성공: {converter.progid_used}"
            self._update_status(status_msg)
            
            # 진행률 초기화
            self.progress_bar["maximum"] = total
            
            # 변환 실행
            for idx, task in enumerate(self.tasks):
                if self.cancel_requested:
                    self._update_status("사용자가 취소했습니다.")
                    break
                
                # 상태 업데이트
                status_text = f"변환 중: {task.input_file.name}"
                self._update_status(status_text)
                self._update_progress(idx, total)
                
                # 출력 폴더 생성
                try:
                    task.output_file.parent.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    task.status = "실패"
                    task.error = f"폴더 생성 실패: {e}"
                    continue
                
                # 변환 실행
                task.status = "진행중"
                format_type = self.format_var.get()
                success, error = converter.convert_file(
                    task.input_file,
                    task.output_file,
                    format_type
                )
                
                if success:
                    task.status = "성공"
                    success_count += 1
                else:
                    task.status = "실패"
                    task.error = error
            
            # 완료
            self._update_progress(total, total)
            
            # 결과 표시
            if not self.cancel_requested:
                self._schedule_show_result(success_count, total)
            
        except Exception as e:
            error_msg = f"변환 중 오류 발생:\n{str(e)}"
            self._schedule_show_error(error_msg)
        
        finally:
            # 정리
            try:
                converter.cleanup()
            except Exception as e:
                print(f"정리 중 오류: {e}")
            
            # UI 복원
            self._schedule_restore_ui()
            self.cancel_requested = False
    
    def _schedule_show_result(self, success_count, total):
        """결과 표시 예약"""
        def show():
            self._show_result(success_count, total)
        self.root.after(0, show)
    
    def _schedule_show_error(self, error_msg):
        """오류 표시 예약"""
        def show():
            messagebox.showerror("오류", error_msg)
        self.root.after(0, show)
    
    def _schedule_restore_ui(self):
        """UI 복원 예약"""
        def restore():
            self._set_converting_state(False)
        self.root.after(0, restore)
    
    def _update_status(self, text):
        """상태 텍스트 업데이트"""
        def update():
            self.status_label.config(text=text)
        self.root.after(0, update)
    
    def _update_progress(self, current, total):
        """진행률 업데이트"""
        def update():
            self.progress_bar.config(value=current)
            self.progress_label.config(text=f"{current} / {total}")
        self.root.after(0, update)
    
    def _show_result(self, success, total):
        """결과 표시"""
        failed = total - success
        
        result_win = tk.Toplevel(self.root)
        result_win.title("변환 완료")
        result_win.geometry("600x400")
        result_win.transient(self.root)
        result_win.grab_set()
        
        # 요약
        summary_frame = ttk.Frame(result_win, padding=20)
        summary_frame.pack(fill=tk.X)
        
        ttk.Label(
            summary_frame,
            text=f"✓ 성공: {success}개",
            font=("", 12, "bold")
        ).pack(anchor=tk.W)
        
        if failed > 0:
            ttk.Label(
                summary_frame,
                text=f"✗ 실패: {failed}개",
                font=("", 12),
                foreground="red"
            ).pack(anchor=tk.W, pady=(5, 0))
        
        # 실패 목록
        if failed > 0:
            ttk.Label(
                result_win,
                text="실패한 파일:",
                font=("", 10, "bold")
            ).pack(anchor=tk.W, padx=20, pady=(10, 5))
            
            list_frame = ttk.Frame(result_win)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))
            
            scrollbar = ttk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            text_widget = tk.Text(
                list_frame,
                yscrollcommand=scrollbar.set,
                wrap=tk.WORD
            )
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=text_widget.yview)
            
            for task in self.tasks:
                if task.status == "실패":
                    text_widget.insert(tk.END, f"파일: {task.input_file.name}\n")
                    text_widget.insert(tk.END, f"오류: {task.error}\n\n")
            
            text_widget.config(state=tk.DISABLED)
        
        # 닫기 버튼
        ttk.Button(
            result_win,
            text="닫기",
            command=result_win.destroy
        ).pack(pady=10)


def main():
    """메인 함수"""
    
    # pywin32 확인
    if not PYWIN32_AVAILABLE:
        return
    
    # 관리자 권한 확인
    if not is_admin():
        messagebox.showwarning(
            "관리자 권한 필요",
            "이 프로그램은 관리자 권한으로 실행해야 합니다.\n\n"
            "파일을 마우스 오른쪽 버튼으로 클릭하여\n"
            "'관리자 권한으로 실행'을 선택하세요."
        )
        sys.exit(1)
    
    # GUI 실행
    root = tk.Tk()
    app = HWPConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
