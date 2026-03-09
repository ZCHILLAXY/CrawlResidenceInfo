"""
配置管理
Made with ❤️by Z🐻
"""
import json
import logging
from pathlib import Path
from typing import Any, Optional


class Config:
    """配置模块"""

    def __init__(self, config_file: str = 'config.json'):
        """
        初始化配置管理

        :param
            config_file: 配置文件路径
        """
        self.config_file = Path(config_file)
        self.logger = logging.getLogger(__name__)
        self._config = {}
        self._load()

    def _load(self):
        """加载配置文件"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
                self.logger.debug(f"Config has been loaded from {self.config_file}")
            except Exception as e:
                self.logger.error(f"Load config file failed: {str(e)}")
                self._config = {}
        else:
            # Initialize default config
            self._config = {
                'session_id': '',
                'max_retries': 20,
                'last_input_dir': '',
                'last_output_dir': '',
            }
            self.save()

    def save(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=4)
            self.logger.debug(f"config has been saved to {self.config_file}")
        except Exception as e:
            self.logger.error(f"Save config file failed: {str(e)}")

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get config

        Args
        :param key: key
        :param default: default

        :return:
            config value
        """
        return self._config.get(key, default)

    def set(self, key: str, value: Any):
        """
        Set config

        Args：
        :param key: key
        :param value: value
        """
        self._config[key] = value

    def get_all(self) -> dict:
        """
        Get all configs

        :return:
            config dict
        """
        return self._config.copy()