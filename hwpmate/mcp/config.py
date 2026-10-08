"""MCP 설정: 허용 루트·quota·로그 정책 (구현 설계서 §6.1).

경로는 예시이며, 루트 설정이 없으면 변환 도구는 fail closed 한다.
설정 파일: %LOCALAPPDATA%/HwpMate/mcp/config.toml (없으면 기본값).
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - Python 3.11+ only
    tomllib = None  # type: ignore[assignment]


def default_config_path() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "HwpMate" / "mcp" / "config.toml"
    return Path.home() / ".hwp_converter" / "mcp" / "config.toml"


def default_job_dir() -> Path:
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "HwpMate" / "mcp" / "jobs"
    return Path.home() / ".hwp_converter" / "mcp" / "jobs"


@dataclass
class McpConfig:
    transport: str = "stdio"
    enabled: bool = True
    max_parallel_conversions: int = 1
    max_queued_jobs: int = 3
    max_files_per_job: int = 50
    max_input_total_mb: int = 500
    max_recursive_depth: int = 5
    max_log_kb: int = 256
    require_preview: bool = True
    require_confirmation: bool = True
    allow_overwrite: bool = False
    allow_disable_backup: bool = False
    allow_remote_http: bool = False
    allow_unc_paths: bool = False
    allow_reparse_points: bool = False
    allow_running_cancel: bool = False
    queue_wait_timeout_seconds: int = 600
    job_timeout_seconds: int = 3600
    plan_ttl_seconds: int = 600
    job_retention_days: int = 7
    input_roots: list[str] = field(default_factory=list)
    output_roots: list[str] = field(default_factory=list)
    cli_executable: str = ""
    job_dir: str = ""

    @classmethod
    def from_mapping(cls, data: dict) -> "McpConfig":
        known = set(cls.__dataclass_fields__.keys())
        filtered = {key: value for key, value in data.items() if key in known}
        return cls(**filtered)

    def effective_job_dir(self) -> Path:
        if str(self.job_dir or "").strip():
            return Path(str(self.job_dir)).expanduser()
        return default_job_dir()


def load_config(path: str | Path | None = None) -> McpConfig:
    """설정 파일 + 환경 변수(HWPMATE_MCP_*)를 읽어 설정을 만든다."""
    data: dict = {}
    config_path = Path(path).expanduser() if path else default_config_path()
    if tomllib is not None and config_path.is_file():
        try:
            with open(config_path, "rb") as handle:
                loaded = tomllib.load(handle)
            if isinstance(loaded, dict):
                section = loaded.get("mcp", loaded)
                paths = loaded.get("paths", {})
                if isinstance(section, dict):
                    data.update(section)
                if isinstance(paths, dict):
                    for key in ("input_roots", "output_roots", "cli_executable"):
                        if key in paths:
                            data[key] = paths[key]
        except OSError:
            pass
    env_roots = os.environ.get("HWPMATE_MCP_INPUT_ROOTS", "").strip()
    if env_roots:
        data["input_roots"] = [part for part in env_roots.split(os.pathsep) if part.strip()]
    env_outputs = os.environ.get("HWPMATE_MCP_OUTPUT_ROOTS", "").strip()
    if env_outputs:
        data["output_roots"] = [part for part in env_outputs.split(os.pathsep) if part.strip()]
    env_cli = os.environ.get("HWPMATE_MCP_CLI", "").strip()
    if env_cli:
        data["cli_executable"] = env_cli
    return McpConfig.from_mapping(data)
