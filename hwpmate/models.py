from __future__ import annotations

from dataclasses import asdict, dataclass
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
    config_version: int = 1
    theme: str = "dark"
    mode: str = "folder"
    format: str = "PDF"
    include_sub: bool = True
    same_location: bool = True
    overwrite: bool = False
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

    def __post_init__(self) -> None:
        self.input_file = Path(self.input_file)
        self.output_file = Path(self.output_file)
