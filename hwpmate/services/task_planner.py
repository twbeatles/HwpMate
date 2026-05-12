from __future__ import annotations

import datetime
from pathlib import Path
from typing import Iterable, Sequence, Set

from ..constants import FORMAT_TYPES, MAX_FILENAME_COUNTER, SUPPORTED_EXTENSIONS
from ..logging_config import get_logger
from ..models import ConversionTask, PlannedConversion
from ..path_utils import canonicalize_path, check_write_permission, iter_supported_files

logger = get_logger(__name__)


class TaskPlanner:
    def preview_allowed_extensions(self, format_type: str) -> Iterable[str]:
        output_ext = FORMAT_TYPES[format_type]["ext"].lower()
        if output_ext in SUPPORTED_EXTENSIONS:
            return [ext for ext in SUPPORTED_EXTENSIONS if ext != output_ext]
        return SUPPORTED_EXTENSIONS

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
        backup_enabled: bool = True,
        retry_count: int = 1,
    ) -> PlannedConversion:
        tasks: list[ConversionTask] = []
        skipped_tasks: list[ConversionTask] = []
        warnings: list[str] = []
        format_info = FORMAT_TYPES[format_type]
        output_ext = format_info["ext"]

        if is_folder_mode:
            collect_start_path = folder_path.strip()
            if not collect_start_path:
                raise ValueError("폴더를 선택하세요.")

            folder = Path(canonicalize_path(collect_start_path))
            if not folder.exists():
                raise ValueError("폴더가 존재하지 않습니다.")
            if not folder.is_dir():
                raise ValueError("폴더 경로가 올바르지 않습니다.")

            allowed_exts: Set[str] = set(SUPPORTED_EXTENSIONS)
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
                if input_file.suffix.lower() == output_ext.lower():
                    skipped_tasks.append(
                        ConversionTask(
                            input_file=input_file,
                            output_file=input_file,
                            status="건너뜀",
                            error=f"이미 {format_type} 형식입니다.",
                        )
                    )
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

                    rel_path = input_file.relative_to(folder)
                    output_file = output_folder / rel_path.parent / (input_file.stem + output_ext)

                tasks.append(ConversionTask(input_file=input_file, output_file=output_file))

            if skipped_tasks:
                warnings.append(
                    f"동일 형식 {len(skipped_tasks)}개는 자동으로 건너뜁니다."
                )
            self._append_output_warnings(tasks, warnings, same_location=same_location)

            return PlannedConversion(
                format_type=format_type,
                same_location=same_location,
                output_path=output_path.strip(),
                backup_enabled=backup_enabled,
                retry_count=retry_count,
                tasks=tasks,
                skipped_tasks=skipped_tasks,
                warnings=warnings,
            )

        if not file_paths:
            raise ValueError("파일을 추가하세요.")

        for file_path in file_paths:
            input_file = Path(file_path)
            if input_file.suffix.lower() == output_ext.lower():
                skipped_tasks.append(
                    ConversionTask(
                        input_file=input_file,
                        output_file=input_file,
                        status="건너뜀",
                        error=f"이미 {format_type} 형식입니다.",
                    )
                )
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

        if skipped_tasks:
            warnings.append(
                f"동일 형식 {len(skipped_tasks)}개는 자동으로 건너뜁니다."
            )
        self._append_output_warnings(tasks, warnings, same_location=same_location)

        return PlannedConversion(
            format_type=format_type,
            same_location=same_location,
            output_path=output_path.strip(),
            backup_enabled=backup_enabled,
            retry_count=retry_count,
            tasks=tasks,
            skipped_tasks=skipped_tasks,
            warnings=warnings,
        )

    def resolve_output_conflicts(self, tasks: list[ConversionTask], overwrite: bool) -> int:
        used_path_keys: set[str] = set()
        renamed_count = 0

        for task in tasks:
            original_path = task.output_file
            original_key = str(original_path).lower()
            batch_duplicate = original_key in used_path_keys

            if batch_duplicate or ((not overwrite) and task.output_file.exists()):
                counter = 1
                stem = original_path.stem
                ext = original_path.suffix
                parent = original_path.parent

                while counter <= MAX_FILENAME_COUNTER:
                    new_name = f"{stem} ({counter}){ext}"
                    new_path = parent / new_name
                    new_key = str(new_path).lower()
                    exists_conflict = (not overwrite) and new_path.exists()
                    if (not exists_conflict) and (new_key not in used_path_keys):
                        task.output_file = new_path
                        break
                    counter += 1
                else:
                    fallback_counter = 1
                    while True:
                        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                        suffix = "" if fallback_counter == 1 else f"_{fallback_counter}"
                        new_name = f"{stem}_{timestamp}{suffix}{ext}"
                        new_path = parent / new_name
                        new_key = str(new_path).lower()
                        exists_conflict = (not overwrite) and new_path.exists()
                        if (not exists_conflict) and (new_key not in used_path_keys):
                            task.output_file = new_path
                            logger.warning(f"파일명 카운터 초과, 타임스탬프 사용: {new_name}")
                            break
                        fallback_counter += 1

                if task.output_file != original_path:
                    task.conflict_original_output_file = original_path
                    renamed_count += 1
                    logger.info(f"출력 경로 조정: {original_path} -> {task.output_file}")

            used_path_keys.add(str(task.output_file).lower())

        return renamed_count

    def _append_output_warnings(
        self,
        tasks: list[ConversionTask],
        warnings: list[str],
        *,
        same_location: bool,
    ) -> None:
        if not same_location:
            return

        unwritable_dirs: list[Path] = []
        seen: set[str] = set()
        for task in tasks:
            parent = task.output_file.parent
            key = str(parent).lower()
            if key in seen:
                continue
            seen.add(key)
            if not check_write_permission(parent):
                unwritable_dirs.append(parent)

        if unwritable_dirs:
            preview = ", ".join(str(path) for path in unwritable_dirs[:3])
            suffix = "" if len(unwritable_dirs) <= 3 else f" 외 {len(unwritable_dirs) - 3}개"
            warnings.append(f"같은 위치 저장 대상 중 쓰기 권한을 확인하지 못한 폴더가 있습니다: {preview}{suffix}")
