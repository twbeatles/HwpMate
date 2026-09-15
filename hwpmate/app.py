from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys
import time
from typing import Optional

from PyQt6.QtWidgets import QApplication, QMessageBox, QStyleFactory

from .app_instance import SingleInstanceLock
from .constants import FORMAT_TYPES, SUPPORTED_EXTENSIONS, VERSION
from .logging_config import get_logger
from .models import ConversionSummary
from .services.hwp_converter import HWPConverter, PYWIN32_AVAILABLE, pythoncom
from .services.hwp_print_settings import normalize_pdf_export_mode
from .services.task_planner import TaskPlanner, count_protected_source_renames
from .services.update_installer import (
    UPDATE_STATUS_APPLIED,
    UPDATE_STATUS_ROLLED_BACK,
    apply_staged_update,
    update_status_for_exception,
    wait_for_file_writable,
    wait_for_process_exit,
    write_update_result,
)
from .ui.main_window import MainWindow
from .windows_integration import (
    enable_drag_drop_for_admin,
    get_native_admin_drag_drop_policy,
    is_admin,
)

logger = get_logger(__name__)

_CLI_CONSOLE_ATTACHED = False


def _ensure_cli_console_output() -> None:
    global _CLI_CONSOLE_ATTACHED
    if _CLI_CONSOLE_ATTACHED or os.name != "nt" or not bool(getattr(sys, "frozen", False)):
        return
    _CLI_CONSOLE_ATTACHED = True
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        attached = bool(kernel32.AttachConsole(-1))
        already_attached = int(kernel32.GetLastError()) == 5
        if attached or already_attached:
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
            sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)
    except Exception:
        pass


def _write_json_line(stream: object, line: str) -> bool:
    payload = f"{line}\n"
    write = getattr(stream, "write", None)
    flush = getattr(stream, "flush", None)
    if callable(write):
        try:
            write(payload)
            if callable(flush):
                flush()
            return True
        except (OSError, UnicodeError, ValueError):
            pass
    return False


def _print_json_line(payload: dict[str, object], *, output_path: str = "") -> None:
    line = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    wrote = False
    original_stdout = getattr(sys, "stdout", None)
    original_stderr = getattr(sys, "stderr", None)
    if _write_json_line(original_stdout, line):
        wrote = True
    elif _write_json_line(original_stderr, line):
        wrote = True

    if not wrote:
        _ensure_cli_console_output()
        if _write_json_line(getattr(sys, "stdout", None), line):
            wrote = True
        elif _write_json_line(getattr(sys, "stderr", None), line):
            wrote = True

    if not wrote and os.name == "nt":
        try:
            with open("CONOUT$", "w", encoding="utf-8", buffering=1) as console:
                console.write(f"{line}\n")
                console.flush()
        except OSError:
            pass

    target = str(output_path or "").strip()
    if not target:
        return
    try:
        out_path = Path(target).resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(line + "\n", encoding="utf-8")
    except Exception:
        pass


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=f"HwpMate v{VERSION} - 한글(HWP/HWPX) 일괄 변환기")
    parser.add_argument(
        "--smoke",
        action="store_true",
        help="GUI 없이 import 및 핵심 모듈 무결성을 검증합니다.",
    )
    parser.add_argument(
        "--smoke-output",
        default="",
        help="smoke 결과 JSON을 stdout 외에 추가로 기록할 파일 경로입니다.",
    )
    # 헤드리스 CLI 변환 인자
    parser.add_argument(
        "--input", "-i",
        default="",
        help="변환할 HWP/HWPX 파일 또는 폴더 경로입니다.",
    )
    parser.add_argument(
        "--format", "-f",
        default="PDF",
        help=f"출력 형식 ({', '.join(FORMAT_TYPES.keys())}, 기본값: PDF).",
    )
    parser.add_argument(
        "--output", "-o",
        default="",
        help="변환 결과물이 저장될 출력 폴더 경로입니다. (생략 시 원본과 동일 위치)",
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="폴더 입력 시 하위 폴더의 모든 파일을 포함하여 변환합니다.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="대상 경로에 동일한 이름의 파일이 있으면 덮어씁니다.",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="변환 전 원본 백업 생성을 비활성화합니다.",
    )
    parser.add_argument(
        "--retry",
        type=int,
        default=1,
        help="변환 실패 시 재시도 횟수 (0~3, 기본값: 1).",
    )
    parser.add_argument(
        "--pdf-export-mode",
        default="saveas_first",
        choices=["saveas_first", "print_to_pdf_ex_first"],
        help="PDF 내보내기 우선 모드 (saveas_first | print_to_pdf_ex_first).",
    )
    parser.add_argument(
        "--report",
        default="",
        help="변환 결과를 저장할 CSV 또는 JSON 경로 (확장자로 형식 결정).",
    )
    parser.add_argument(
        "--no-auto-continue",
        action="store_true",
        help="한글 「호환 문서(배치가 변경될 수 있습니다)」 확인 창에 자동으로 계속하지 않습니다.",
    )

    # 자동 업데이트 내부 헬퍼 인자
    parser.add_argument("--apply-update", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--update-target", default="", help=argparse.SUPPRESS)
    parser.add_argument("--update-staged", default="", help=argparse.SUPPRESS)
    parser.add_argument("--update-backup", default="", help=argparse.SUPPRESS)
    parser.add_argument("--update-parent-pid", default=0, type=int, help=argparse.SUPPRESS)
    parser.add_argument("--update-expected-sha256", default="", help=argparse.SUPPRESS)
    parser.add_argument("--update-expected-size", default=0, type=int, help=argparse.SUPPRESS)
    parser.add_argument("--update-result-file", default="", help=argparse.SUPPRESS)
    return parser.parse_args(argv)


def _run_smoke(args: argparse.Namespace) -> int:
    result: dict[str, object] = {
        "status": "ok",
        "version": VERSION,
        "pywin32_available": PYWIN32_AVAILABLE,
    }
    try:
        import cryptography  # noqa: F401
        from .services.update_manifest import verify_release_manifest  # noqa: F401
        from .services.update_installer import apply_staged_update  # noqa: F401
        result["cryptography_available"] = True
    except Exception as exc:
        result["status"] = "error"
        result["error"] = f"cryptography / update 모듈 로드 실패: {exc}"
        _print_json_line(result, output_path=args.smoke_output)
        return 1

    _print_json_line(result, output_path=args.smoke_output)
    return 0


def _wait_for_parent(parent_pid: int, timeout: float = 30.0) -> None:
    """부모 앱 프로세스 종료를 기다린다 (os.kill(pid, 0) 은 Windows 에서 존재 확인이 아님)."""
    if parent_pid <= 0:
        return
    if not wait_for_process_exit(parent_pid, timeout):
        raise TimeoutError("업데이트 적용 전 부모 프로세스가 종료되지 않았습니다.")


def _relaunch(target: Path) -> None:
    try:
        import subprocess

        subprocess.Popen([str(target.resolve())])
    except Exception:
        pass


def _run_apply_update(args: argparse.Namespace) -> int:
    target = Path(args.update_target)
    base_result = {
        "target": str(target.resolve()),
        "backup": str(Path(args.update_backup).resolve()),
        "completed_at": time.time(),
    }
    try:
        _wait_for_parent(args.update_parent_pid)
        # onefile 부트로더가 exe 이미지를 놓을 때까지 대기 (교체 자체도 재시도한다)
        wait_for_file_writable(target.resolve())
        apply_staged_update(
            target=target,
            staged=Path(args.update_staged),
            backup=Path(args.update_backup),
            expected_sha256=args.update_expected_sha256,
            expected_size=args.update_expected_size,
        )
    except Exception as exc:
        status = update_status_for_exception(exc)
        write_update_result(
            args.update_result_file,
            {**base_result, "status": status, "error": str(exc)},
        )
        if status == UPDATE_STATUS_ROLLED_BACK:
            # 이전 버전으로 복구됐으므로 사용자가 앱이 사라졌다고 느끼지 않게 다시 실행한다.
            _relaunch(target)
        return 1

    write_update_result(args.update_result_file, {**base_result, "status": UPDATE_STATUS_APPLIED})
    _relaunch(target)
    return 0


def _cli_error(message: str) -> int:
    print(f"[오류] {message}", file=sys.stderr)
    return 1


def _write_cli_report(report_path: str, summary: ConversionSummary) -> str | None:
    target = str(report_path or "").strip()
    if not target:
        return None
    from .ui.dialogs.atomic_io import write_results_csv, write_results_json

    path = Path(target).resolve()
    if path.suffix.lower() == ".json":
        write_results_json(path, summary)
    else:
        if path.suffix.lower() != ".csv":
            path = path.with_suffix(".csv")
        write_results_csv(path, summary)
    return str(path)


def _run_cli_conversion(args: argparse.Namespace) -> int:
    """헤드리스 CLI 일괄 변환 실행 (GUI 워커와 같은 작업 실행기 사용)."""
    from .workers.conversion_worker.summary import collect_converter_warnings
    from .workers.conversion_worker.task_runner import (
        TaskRunOptions,
        compat_dialog_note,
        execute_task,
        fail_pending_tasks,
        recycle_converter,
    )
    from .constants import CONVERTER_RECYCLE_BATCH_COUNT, MAX_RETRY_COUNT

    _ensure_cli_console_output()
    started = time.perf_counter()
    input_path = Path(args.input).resolve()
    if not input_path.exists():
        return _cli_error(f"입력 경로가 존재하지 않습니다: {input_path}")

    format_type = str(args.format).upper().strip()
    if format_type not in FORMAT_TYPES:
        return _cli_error(
            f"지원하지 않는 출력 형식: {args.format} (가능한 형식: {', '.join(FORMAT_TYPES.keys())})"
        )

    is_folder_mode = input_path.is_dir()
    if not is_folder_mode and input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return _cli_error(
            f"한글 문서(.hwp/.hwpx)만 변환할 수 있습니다: {input_path.name}"
        )

    retry_count = max(0, min(MAX_RETRY_COUNT, int(args.retry)))
    same_location = not bool(args.output)
    output_path = ""
    if args.output:
        try:
            out_dir = Path(args.output).resolve()
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return _cli_error(f"출력 폴더를 만들 수 없습니다: {exc}")
        output_path = str(out_dir)

    planner = TaskPlanner()
    try:
        plan = planner.build_tasks(
            is_folder_mode=is_folder_mode,
            format_type=format_type,
            folder_path=str(input_path) if is_folder_mode else "",
            include_sub=args.recursive,
            same_location=same_location,
            output_path=output_path,
            overwrite=args.overwrite,
            file_paths=[str(input_path)] if not is_folder_mode else [],
            backup_enabled=not args.no_backup,
            retry_count=retry_count,
            pdf_export_mode=args.pdf_export_mode,
        )
    except ValueError as exc:
        return _cli_error(str(exc))

    renamed = planner.resolve_output_conflicts(plan.tasks, overwrite=args.overwrite, format_type=format_type)
    if renamed:
        plan.warnings.append(f"출력 경로 충돌 {renamed}개는 자동으로 새 이름으로 저장됩니다.")
    protected = count_protected_source_renames(plan.tasks)
    if args.overwrite and protected:
        plan.warnings.append(
            f"원본 한글 문서 보호: {protected}개는 덮어쓰지 않고 새 이름으로 저장합니다."
        )

    tasks = plan.tasks
    print(f"=== HwpMate v{VERSION} CLI 일괄 변환 ===")
    print(f"변환 대상: {len(tasks)}개 | 건너뜀: {plan.skipped_count}개 | 목표 형식: {format_type}")
    for warning in plan.warnings:
        print(f"⚠️ [경고] {warning}")
    for skipped in plan.skipped_tasks:
        print(f"⏭️ 건너뜀: {skipped.input_file.name} ({skipped.detail})")

    summary_warnings = list(plan.warnings)
    if not tasks:
        print("[안내] 실행할 변환 대상이 없습니다.")
        summary = ConversionSummary(
            format_type=format_type,
            tasks=[task.snapshot() for task in plan.skipped_tasks],
            warnings=summary_warnings,
            elapsed_seconds=time.perf_counter() - started,
        )
        _write_cli_report(args.report, summary)
        return 0

    if not PYWIN32_AVAILABLE:
        return _cli_error("pywin32 라이브러리가 필요합니다. pip install pywin32 후 다시 실행하세요.")

    # GUI 또는 다른 CLI 와 같은 한글 COM·출력 파일을 동시에 다루지 않게 한다.
    instance_lock = SingleInstanceLock()
    if not instance_lock.try_lock():
        return _cli_error("HwpMate(GUI 또는 다른 CLI)가 이미 실행 중입니다. 종료한 뒤 다시 실행하세요.")

    converter = HWPConverter()
    options = TaskRunOptions.from_plan(plan)
    runtime_warnings: list[str] = []
    try:
        try:
            converter.auto_continue_compat_dialogs = not args.no_auto_continue
            converter.initialize(manage_com_apartment=True)
            converter.pdf_export_mode = normalize_pdf_export_mode(args.pdf_export_mode)
            runtime_warnings.extend(collect_converter_warnings(converter))
        except Exception as exc:
            return _cli_error(f"한글 COM 초기화 실패: {exc}")

        used_output_path_keys: set[str] = set()
        total = len(tasks)
        try:
            for idx, task in enumerate(tasks):
                prefix = f"[{idx + 1}/{total}]"
                print(f"{prefix} 🔄 {task.input_file.name} ➔ {task.output_file.name} ...", end=" ", flush=True)
                runtime_warnings.extend(
                    execute_task(
                        task,
                        converter,
                        options,
                        planner=planner,
                        used_output_path_keys=used_output_path_keys,
                        cancel_check=lambda: False,
                        on_status=lambda text: print(f"({text})", end=" ", flush=True),
                    )
                )
                if task.status == "성공":
                    print("✅ 성공")
                else:
                    print(f"❌ {task.status}: {task.detail}")

                if (idx + 1) % CONVERTER_RECYCLE_BATCH_COUNT == 0 and (idx + 1) < total:
                    recycle_error = recycle_converter(converter, options)
                    if recycle_error is not None:
                        message = f"한글 프로세스 재순환 실패로 남은 작업을 중단했습니다: {recycle_error}"
                        fail_pending_tasks(tasks, message)
                        runtime_warnings.append(message)
                        print(f"[오류] {message}", file=sys.stderr)
                        break
        except KeyboardInterrupt:
            for task in tasks:
                if task.status in {"대기", "진행중"}:
                    task.status = "취소됨"
                    task.error = "사용자 취소 (Ctrl+C)"
            runtime_warnings.append("사용자가 Ctrl+C 로 변환을 중단했습니다.")
        note = compat_dialog_note(converter)
        if note:
            runtime_warnings.append(note)
    finally:
        try:
            converter.cleanup()
        except Exception:
            pass
        instance_lock.release()

    summary = ConversionSummary(
        format_type=format_type,
        tasks=[task.snapshot() for task in tasks] + [task.snapshot() for task in plan.skipped_tasks],
        warnings=summary_warnings + runtime_warnings,
        elapsed_seconds=time.perf_counter() - started,
        progid_used=converter.progid_used,
    )
    print("=" * 45)
    print(
        f"변환 완료 요약: 성공 {summary.success_count}건, 실패 {summary.failed_count}건, "
        f"건너뜀 {summary.skipped_count}건, 취소 {summary.canceled_count}건"
    )
    for warning in runtime_warnings:
        print(f"⚠️ {warning}")
    try:
        report = _write_cli_report(args.report, summary)
        if report:
            print(f"결과 저장: {report}")
    except Exception as exc:
        print(f"[경고] 결과 저장 실패: {exc}", file=sys.stderr)
    return 0 if summary.failed_count == 0 and summary.canceled_count == 0 else 1


def handle_exception(exc_type, exc_value, exc_traceback) -> None:
    """글로벌 예외 핸들러."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical("치명적 오류 발생", exc_info=(exc_type, exc_value, exc_traceback))

    try:
        if QApplication.instance():
            QMessageBox.critical(
                None,
                "치명적 오류",
                f"프로그램에서 예기치 않은 오류가 발생했습니다.\n\n"
                f"오류: {exc_type.__name__}: {exc_value}\n\n"
                f"프로그램을 다시 시작해 주세요.",
            )
    except Exception:
        pass


def main(argv: list[str] | None = None) -> int:
    """메인 함수."""
    sys.excepthook = handle_exception
    args = _parse_args(sys.argv[1:] if argv is None else argv)

    if args.apply_update:
        return _run_apply_update(args)

    if args.smoke:
        return _run_smoke(args)

    if args.input:
        return _run_cli_conversion(args)

    if not PYWIN32_AVAILABLE:
        app = QApplication(sys.argv)
        QMessageBox.critical(
            None, "오류",
            "pywin32 라이브러리가 필요합니다.\n\npip install pywin32"
        )
        del app
        return 1

    if not is_admin():
        app = QApplication(sys.argv)
        QMessageBox.warning(
            None,
            "관리자 권한 필요",
            "이 프로그램은 관리자 권한으로 실행해야 합니다.\n\n"
            "파일을 마우스 오른쪽 버튼으로 클릭하여\n"
            "'관리자 권한으로 실행'을 선택하세요."
        )
        del app
        return 1

    try:
        native_dnd_enabled, native_dnd_reason = get_native_admin_drag_drop_policy()
        if native_dnd_enabled:
            enable_drag_drop_for_admin()
        else:
            logger.warning(f"관리자용 네이티브 드래그 앤 드롭 비활성화: {native_dnd_reason}")

        app = QApplication(sys.argv)
        app.setStyle(QStyleFactory.create("Fusion"))

        instance_lock = SingleInstanceLock()
        if not instance_lock.try_lock():
            QMessageBox.information(
                None,
                "이미 실행 중",
                "HwpMate가 이미 실행 중입니다.\n기존 창을 사용해 주세요.",
            )
            del app
            return 0

        try:
            window = MainWindow()
            window.show()

            exit_code = app.exec()
            logger.info(f"애플리케이션 이벤트 루프 종료: code={exit_code}")
            return int(exit_code)
        finally:
            instance_lock.release()
    except Exception as e:
        logger.critical(f"애플리케이션 실행 오류: {e}")
        raise
