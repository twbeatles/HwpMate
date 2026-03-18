from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .constants import CONFIG_VERSION
from .logging_config import get_logger
from .models import AppConfig

logger = get_logger(__name__)

CONFIG_FILE = Path.home() / ".hwp_converter_config.json"


class ConfigRepository:
    def __init__(self, config_file: Path = CONFIG_FILE) -> None:
        self.config_file = config_file

    def default_config(self) -> AppConfig:
        return AppConfig(config_version=CONFIG_VERSION)

    def load(self) -> AppConfig:
        default_config = self.default_config()

        try:
            if self.config_file.exists():
                with self.config_file.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        saved_version = data.get("config_version", 0)
                        if saved_version < CONFIG_VERSION:
                            logger.info(f"설정 파일 버전 업그레이드: {saved_version} -> {CONFIG_VERSION}")
                        merged = {**default_config.to_dict(), **data}
                        merged["config_version"] = CONFIG_VERSION
                        return AppConfig.from_mapping(merged)
                    logger.warning("설정 파일 형식이 올바르지 않습니다. 기본값 사용")
        except json.JSONDecodeError as e:
            logger.error(f"설정 파일 JSON 파싱 오류: {e}")
            try:
                backup_path = self.config_file.with_suffix(".json.bak")
                self.config_file.rename(backup_path)
                logger.info(f"손상된 설정 파일을 {backup_path}로 백업했습니다")
            except Exception:
                pass
        except Exception as e:
            logger.error(f"설정 로드 실패: {e}")
        return default_config

    def save(self, config: AppConfig | dict[str, Any]) -> None:
        try:
            config_data = config.to_dict() if isinstance(config, AppConfig) else AppConfig.from_mapping(config).to_dict()
            with self.config_file.open("w", encoding="utf-8") as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"설정 저장 실패: {e}")


_DEFAULT_REPOSITORY = ConfigRepository()


def load_config() -> AppConfig:
    return _DEFAULT_REPOSITORY.load()


def save_config(config: AppConfig | dict[str, Any]) -> None:
    _DEFAULT_REPOSITORY.save(config)
