from __future__ import annotations

import json
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from .constants import CONFIG_VERSION, FORMAT_TYPES, MAX_RETRY_COUNT
from .logging_config import get_logger
from .models import AppConfig

logger = get_logger(__name__)

CONFIG_FILE = Path.home() / ".hwp_converter_config.json"


class ConfigRepository:
    def __init__(self, config_file: Path = CONFIG_FILE) -> None:
        self.config_file = config_file

    def default_config(self) -> AppConfig:
        return AppConfig(config_version=CONFIG_VERSION)

    def _normalize_mapping(self, data: dict[str, Any], default_config: AppConfig) -> tuple[AppConfig, list[str]]:
        merged = {**default_config.to_dict(), **data}
        repairs: list[str] = []

        def repair(key: str, value: Any) -> None:
            merged[key] = value
            repairs.append(key)

        string_keys = {"theme", "mode", "format", "folder_path", "output_path", "last_folder", "last_output"}
        for key in string_keys:
            if not isinstance(merged.get(key), str):
                repair(key, getattr(default_config, key))

        if merged["theme"] not in {"dark", "light"}:
            repair("theme", default_config.theme)
        if merged["mode"] not in {"folder", "files"}:
            repair("mode", default_config.mode)
        if merged["format"] not in FORMAT_TYPES:
            repair("format", default_config.format)

        bool_keys = {"include_sub", "same_location", "overwrite", "backup_enabled"}
        for key in bool_keys:
            value = merged.get(key)
            if isinstance(value, bool):
                continue
            if isinstance(value, str):
                normalized = value.strip().lower()
                if normalized in {"1", "true", "yes", "on"}:
                    repair(key, True)
                    continue
                if normalized in {"0", "false", "no", "off"}:
                    repair(key, False)
                    continue
            repair(key, getattr(default_config, key))

        try:
            retry_count = int(merged.get("retry_count", default_config.retry_count))
            if retry_count < 0 or retry_count > MAX_RETRY_COUNT:
                raise ValueError
            merged["retry_count"] = retry_count
        except (TypeError, ValueError):
            repair("retry_count", default_config.retry_count)

        merged["config_version"] = CONFIG_VERSION
        return AppConfig.from_mapping(merged), repairs

    def load(self) -> AppConfig:
        default_config = self.default_config()

        try:
            if self.config_file.exists():
                with self.config_file.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        try:
                            saved_version = int(data.get("config_version", 0))
                        except (TypeError, ValueError):
                            saved_version = 0
                        if saved_version < CONFIG_VERSION:
                            logger.info(f"설정 파일 버전 업그레이드: {saved_version} -> {CONFIG_VERSION}")
                        config, repairs = self._normalize_mapping(data, default_config)
                        if repairs:
                            logger.warning(f"설정 값 타입/범위 복구: {', '.join(sorted(set(repairs)))}")
                        return config
                    logger.warning("설정 파일 형식이 올바르지 않습니다. 기본값 사용")
        except json.JSONDecodeError as e:
            logger.error(f"설정 파일 JSON 파싱 오류: {e}")
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                backup_path = self.config_file.with_name(
                    f"{self.config_file.stem}_{timestamp}{self.config_file.suffix}.bak"
                )
                self.config_file.rename(backup_path)
                logger.info(f"손상된 설정 파일을 {backup_path}로 백업했습니다")
            except Exception:
                pass
        except Exception as e:
            logger.error(f"설정 로드 실패: {e}")
        return default_config

    def save(self, config: AppConfig | dict[str, Any]) -> None:
        temp_path: Path | None = None
        try:
            if isinstance(config, AppConfig):
                config_data = config.to_dict()
            else:
                normalized, repairs = self._normalize_mapping(config, self.default_config())
                if repairs:
                    logger.warning(f"저장 전 설정 값 타입/범위 복구: {', '.join(sorted(set(repairs)))}")
                config_data = normalized.to_dict()
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with tempfile.NamedTemporaryFile(
                "w",
                encoding="utf-8",
                dir=self.config_file.parent,
                prefix=f".{self.config_file.name}.",
                suffix=".tmp",
                delete=False,
            ) as f:
                temp_path = Path(f.name)
                json.dump(config_data, f, ensure_ascii=False, indent=2)
                f.write("\n")
            temp_path.replace(self.config_file)
        except Exception as e:
            logger.error(f"설정 저장 실패: {e}")
            if temp_path is not None:
                try:
                    temp_path.unlink(missing_ok=True)
                except OSError:
                    pass


_DEFAULT_REPOSITORY = ConfigRepository()


def load_config() -> AppConfig:
    return _DEFAULT_REPOSITORY.load()


def save_config(config: AppConfig | dict[str, Any]) -> None:
    _DEFAULT_REPOSITORY.save(config)
