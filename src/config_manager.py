#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理模块
"""

import os
import configparser
from pathlib import Path


class ConfigManager:
    """配置文件管理器"""

    def __init__(self):
        self.config_file = Path(__file__).parent / "config.ini"
        self.config = configparser.ConfigParser()
        self._load()

    def _load(self):
        """加载配置文件"""
        if self.config_file.exists():
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # 创建默认配置
            self._create_default()

    def _create_default(self):
        """创建默认配置"""
        self.config['paths'] = {
            'tongdaxin_path': ''
        }
        self.config['notify'] = {
            'phone_number': ''
        }
        self.config['schedule'] = {
            'enabled': 'true',
            'morning_time': '12:05',
            'afternoon_time': '15:05'
        }
        self.config['log'] = {
            'keep_days': '30'
        }
        self.save()

    def get(self, section: str, key: str, default: str = "") -> str:
        """获取配置值"""
        try:
            return self.config.get(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError):
            return default

    def set(self, section: str, key: str, value: str):
        """设置配置值"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)

    def save(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception as e:
            print(f"保存配置失败：{e}")

    def get_all(self) -> dict:
        """获取所有配置"""
        return {s: dict(self.config.items(s)) for s in self.config.sections()}
