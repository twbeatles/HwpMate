"""MCP 설정: 허용 루트·quota·로그 정책 (구현 설계서 §6.1).

경로는 예시이며, 루트 설정이 없으면 변환 도구는 fail closed 한다.
설정 파일: %LOCALAPPDATA%/HwpMate/mcp/config.toml (없으면 기본값).
값의 타입이 맞지 않으면 추측해서 쓰지 않고 McpConfigError 로 시작을 거부한다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field, fields
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - Python 3.11+ only
    tomllib = None  # type: ignore[assignment]


class McpConfigError(ValueError):
    """설정 파일·환경 변수 값이 올바르지 않아 서버를 시작할 수 없음."""


# 설정 파일에서 읽지만 안전상 고정값으로만 동작하는 키 (다른 값이면 경고)
FIXED_POLICY_VALUES: dict[str, object] = {
    "max_parallel_conversions": 1,
    "require_preview": True,
    "require_confirmation": True,
    "allow_overwrite": False,
    "allow_disable_backup": False,
    "allow_remote_http": False,
    "allow_running_cancel": False,
}

_ROOT_LIST_KEYS = ("input_roots", "output_roots")


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


def is_drive_root(raw: str) -> bool:
    """`C:/`, `\\\\server\\share\\`, `/` 처럼 드라이브(볼륨) 전체를 가리키면 True."""
    absolute = Path(os.path.abspath(os.path.expanduser(str(raw))))
    return absolute == Path(absolute.anchor)


def _normalize_roots(key: str, value: object) -> list[str]:
    # 문자열 하나를 그대로 두면 정책 검사가 문자 단위로 반복되어 드라이브 전체를 허용하게 된다.
    if isinstance(value, str):
        items: list[object] = [value]
    elif isinstance(value, (list, tuple)):
        items = list(value)
    else:
        raise McpConfigError(f"{key} 는 경로 문자열 목록이어야 합니다 (현재: {type(value).__name__}).")
    roots: list[str] = []
    for item in items:
        if not isinstance(item, str):
            raise McpConfigError(f"{key} 항목은 문자열이어야 합니다 (현재: {type(item).__name__}).")
        text = item.strip()
        if not text:
            continue
        if is_drive_root(text):
            raise McpConfigError(f"{key} 에 드라이브 전체({text})는 지정할 수 없습니다. 작업 폴더를 지정하세요.")
        roots.append(text)
    return roots


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
    # CLI 1회(입력 파일 1개) 실행 제한 시간
    job_timeout_seconds: int = 3600
    plan_ttl_seconds: int = 600
    job_retention_days: int = 7
    input_roots: list[str] = field(default_factory=list)
    output_roots: list[str] = field(default_factory=list)
    cli_executable: str = ""
    job_dir: str = ""
    # 적용되지 않는 설정 등 시작 시 알릴 내용 (설정 파일 키가 아님)
    config_warnings: list[str] = field(default_factory=list)

    @classmethod
    def from_mapping(cls, data: dict) -> "McpConfig":
        defaults = cls()
        values: dict[str, object] = {}
        for spec in fields(cls):
            key = spec.name
            if key == "config_warnings" or key not in data:
                continue
            value = data[key]
            if key in _ROOT_LIST_KEYS:
                values[key] = _normalize_roots(key, value)
                continue
            expected = type(getattr(defaults, key))
            if expected is bool:
                valid = isinstance(value, bool)
            elif expected is int:
                valid = isinstance(value, int) and not isinstance(value, bool)
            else:
                valid = isinstance(value, expected)
            if not valid:
                raise McpConfigError(
                    f"{key} 값의 형식이 올바르지 않습니다: {value!r} ({expected.__name__} 필요)"
                )
            values[key] = value
        cfg = cls(**values)  # type: ignore[arg-type]
        for key, fixed in FIXED_POLICY_VALUES.items():
            if getattr(cfg, key) != fixed:
                cfg.config_warnings.append(
                    f"{key}={getattr(cfg, key)!r} 설정은 적용되지 않습니다 (안전상 {fixed!r} 고정)."
                )
        return cfg

    def effective_job_dir(self) -> Path:
        if str(self.job_dir or "").strip():
            return Path(str(self.job_dir)).expanduser()
        return default_job_dir()


def load_config(path: str | Path | None = None) -> McpConfig:
    """설정 파일 + 환경 변수(HWPMATE_MCP_*)를 읽어 설정을 만든다.

    파일 문법 오류·형식 오류는 McpConfigError 로 알린다 (조용히 무시하지 않음).
    """
    data: dict = {}
    config_path = Path(path).expanduser() if path else default_config_path()
    if config_path.is_file():
        if tomllib is None:
            raise McpConfigError(
                f"설정 파일을 읽으려면 Python 3.11 이상이 필요합니다: {config_path}"
            )
        try:
            with open(config_path, "rb") as handle:
                loaded = tomllib.load(handle)
        except tomllib.TOMLDecodeError as exc:
            raise McpConfigError(
                f"설정 파일 문법 오류: {config_path} ({exc}). "
                "Windows 경로는 'C:/Work' 처럼 / 를 쓰거나 작은따옴표로 감싸세요."
            ) from exc
        except OSError as exc:
            raise McpConfigError(f"설정 파일을 읽을 수 없습니다: {config_path} ({exc})") from exc
        section = loaded.get("mcp", loaded)
        paths = loaded.get("paths", {})
        if isinstance(section, dict):
            data.update(section)
        if isinstance(paths, dict):
            for key in ("input_roots", "output_roots", "cli_executable"):
                if key in paths:
                    data[key] = paths[key]
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
