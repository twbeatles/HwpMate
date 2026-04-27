from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class FormatSpec:
    ext: str
    save_format: str
    icon: str
    desc: str

    def __getitem__(self, key: str) -> Any:
        return getattr(self, key)


@dataclass
class AppConfig:
    config_version: int = 2
    theme: str = "dark"
    mode: str = "folder"
    format: str = "PDF"
    include_sub: bool = True
    same_location: bool = True
    overwrite: bool = False
    backup_enabled: bool = True
    retry_count: int = 1
    folder_path: str = ""
    output_path: str = ""
    last_folder: str = ""
    last_output: str = ""

    def get(self, key: str, default: Any = None) -> Any:
        return getattr(self, key, default)

    def __getitem__(self, key: str) -> Any:
        return getattr(self, key)

    def __setitem__(self, key: str, value: Any) -> None:
        setattr(self, key, value)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_mapping(cls, data: dict[str, Any]) -> "AppConfig":
        known_keys = set(cls.__dataclass_fields__.keys())
        filtered = {key: value for key, value in data.items() if key in known_keys}
        return cls(**filtered)


@dataclass
class ConversionTask:
    input_file: Path
    output_file: Path
    status: str = "대기"
    error: str | None = None
    retry_count: int = 0
    backup_file: Path | None = None
    backup_error: str | None = None
    conflict_original_output_file: Path | None = None

    def __post_init__(self) -> None:
        self.input_file = Path(self.input_file)
        self.output_file = Path(self.output_file)
        if self.backup_file is not None:
            self.backup_file = Path(self.backup_file)
        if self.conflict_original_output_file is not None:
            self.conflict_original_output_file = Path(self.conflict_original_output_file)

    @property
    def detail(self) -> str:
        return self.error or ""

    def to_record(self) -> dict[str, Any]:
        return {
            "input_file": str(self.input_file),
            "output_file": str(self.output_file),
            "status": self.status,
            "detail": self.detail,
            "retry_count": self.retry_count,
            "backup_file": str(self.backup_file) if self.backup_file is not None else "",
            "backup_error": self.backup_error or "",
        }


@dataclass
class PlannedConversion:
    format_type: str
    same_location: bool
    output_path: str
    backup_enabled: bool = True
    retry_count: int = 1
    tasks: list[ConversionTask] = field(default_factory=list)
    skipped_tasks: list[ConversionTask] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    conflict_renamed_count: int = 0

    @property
    def runnable_count(self) -> int:
        return len(self.tasks)

    @property
    def skipped_count(self) -> int:
        return len(self.skipped_tasks)

    @property
    def total_requested(self) -> int:
        return self.runnable_count + self.skipped_count

    @property
    def all_tasks(self) -> list[ConversionTask]:
        return sorted(self.tasks + self.skipped_tasks, key=lambda task: str(task.input_file).lower())

    @property
    def output_policy_label(self) -> str:
        if self.same_location:
            return "입력 파일과 같은 위치"
        return self.output_path or "사용자 지정 출력 폴더"


@dataclass
class ConversionSummary:
    format_type: str
    tasks: list[ConversionTask] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    elapsed_seconds: float | None = None
    progid_used: str | None = None

    @property
    def total_requested(self) -> int:
        return len(self.tasks)

    @property
    def success_count(self) -> int:
        return len([task for task in self.tasks if task.status == "성공"])

    @property
    def failed_count(self) -> int:
        return len([task for task in self.tasks if task.status == "실패"])

    @property
    def skipped_count(self) -> int:
        return len([task for task in self.tasks if task.status == "건너뜀"])

    @property
    def canceled_count(self) -> int:
        return len([task for task in self.tasks if task.status == "취소됨"])

    @property
    def output_paths(self) -> list[str]:
        return [str(task.output_file) for task in self.tasks if task.status == "성공"]

    @property
    def failed_tasks(self) -> list[ConversionTask]:
        return [task for task in self.tasks if task.status == "실패"]

    @property
    def skipped_tasks(self) -> list[ConversionTask]:
        return [task for task in self.tasks if task.status == "건너뜀"]

    @property
    def canceled_tasks(self) -> list[ConversionTask]:
        return [task for task in self.tasks if task.status == "취소됨"]

    def sorted_tasks(self) -> list[ConversionTask]:
        return sorted(self.tasks, key=lambda task: str(task.input_file).lower())

    def to_json_dict(self) -> dict[str, Any]:
        return {
            "summary": {
                "format_type": self.format_type,
                "total_requested": self.total_requested,
                "success_count": self.success_count,
                "failed_count": self.failed_count,
                "skipped_count": self.skipped_count,
                "canceled_count": self.canceled_count,
                "elapsed_seconds": self.elapsed_seconds,
                "progid_used": self.progid_used,
                "warnings": list(self.warnings),
            },
            "tasks": [task.to_record() for task in self.sorted_tasks()],
        }
