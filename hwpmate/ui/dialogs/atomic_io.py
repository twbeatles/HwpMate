"""결과/실패 목록 원자적 파일 저장."""

from __future__ import annotations

import csv
import json
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Callable, TextIO, cast

from ...models import ConversionSummary, ConversionTask


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
