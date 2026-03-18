from __future__ import annotations

import datetime
from pathlib import Path
from typing import Sequence, Set

from ..constants import FORMAT_TYPES, MAX_FILENAME_COUNTER, SUPPORTED_EXTENSIONS
from ..logging_config import get_logger
from ..models import ConversionTask
from ..path_utils import canonicalize_path, iter_supported_files

logger = get_logger(__name__)


class TaskPlanner:
    def build_tasks(
        self,
        *,
        is_folder_mode: bool,
        format_type: str,
        folder_path: str,
        include_sub: bool,
        same_location: bool,
        output_path: str,
        file_paths: Sequence[str],
    ) -> list[ConversionTask]:
        tasks: list[ConversionTask] = []
        format_info = FORMAT_TYPES[format_type]
        output_ext = format_info["ext"]

        if is_folder_mode:
            collect_start_path = folder_path.strip()
            if not collect_start_path:
                raise ValueError("폴더를 선택하세요.")

            folder = Path(canonicalize_path(collect_start_path))
            if not folder.exists():
                raise ValueError("폴더가 존재하지 않습니다.")

            allowed_exts: Set[str] = {".hwp"} if format_type == "HWPX" else set(SUPPORTED_EXTENSIONS)
            input_files = [
                Path(canonicalize_path(str(p)))
                for p in iter_supported_files(
                    folder,
                    include_sub=include_sub,
                    allowed_exts=allowed_exts,
                )
            ]

            if not input_files:
                raise ValueError("변환할 파일이 없습니다.")

            input_files = sorted(input_files, key=lambda p: str(p).lower())
            logger.debug(f"폴더 작업 수집: {len(input_files)}개")

            for input_file in input_files:
                if same_location:
                    output_file = input_file.parent / (input_file.stem + output_ext)
                else:
                    output_folder_text = output_path.strip()
                    if not output_folder_text:
                        raise ValueError("출력 폴더를 선택하세요.")
                    output_folder = Path(output_folder_text)
                    if not output_folder.exists():
                        raise ValueError(f"출력 폴더가 존재하지 않습니다: {output_folder}")

                    rel_path = input_file.relative_to(folder)
                    output_file = output_folder / rel_path.parent / (input_file.stem + output_ext)

                tasks.append(ConversionTask(input_file=input_file, output_file=output_file))
            return tasks

        if not file_paths:
            raise ValueError("파일을 추가하세요.")

        skipped_hwpx = 0
        for file_path in file_paths:
            input_file = Path(file_path)
            if format_type == "HWPX" and input_file.suffix.lower() == ".hwpx":
                skipped_hwpx += 1
                logger.debug(f"HWPX->HWPX 변환 건너뜀: {input_file.name}")
                continue

            if same_location:
                output_file = input_file.parent / (input_file.stem + output_ext)
            else:
                output_folder_text = output_path.strip()
                if not output_folder_text:
                    raise ValueError("출력 폴더를 선택하세요.")
                output_folder = Path(output_folder_text)
                if not output_folder.exists():
                    raise ValueError(f"출력 폴더가 존재하지 않습니다: {output_folder}")
                output_file = output_folder / (input_file.stem + output_ext)

            tasks.append(ConversionTask(input_file=input_file, output_file=output_file))

        if skipped_hwpx > 0 and not tasks:
            raise ValueError(
                f"선택한 모든 파일({skipped_hwpx}개)이 이미 HWPX 형식입니다.\n"
                "HWPX 파일을 다시 HWPX로 변환할 수 없습니다."
            )
        if skipped_hwpx > 0:
            logger.debug(f"{skipped_hwpx}개 HWPX 파일을 건너뛰었습니다 (HWPX->HWPX 변환 불가)")

        return tasks

    def resolve_output_conflicts(self, tasks: list[ConversionTask], overwrite: bool) -> None:
        if overwrite:
            return

        used_paths: set[Path] = set()

        for task in tasks:
            original_path = task.output_file

            if task.output_file.exists() or task.output_file in used_paths:
                counter = 1
                stem = original_path.stem
                ext = original_path.suffix
                parent = original_path.parent

                while counter <= MAX_FILENAME_COUNTER:
                    new_name = f"{stem} ({counter}){ext}"
                    new_path = parent / new_name
                    if (not new_path.exists()) and (new_path not in used_paths):
                        task.output_file = new_path
                        break
                    counter += 1
                else:
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_name = f"{stem}_{timestamp}{ext}"
                    task.output_file = parent / new_name
                    logger.warning(f"파일명 카운터 초과, 타임스탬프 사용: {new_name}")

                if task.output_file != original_path:
                    logger.info(f"출력 경로 조정: {original_path} -> {task.output_file}")

            used_paths.add(task.output_file)
