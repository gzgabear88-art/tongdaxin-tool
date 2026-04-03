#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理器 - Config Manager

负责管理用户的配置信息，包括：
- 账号密码的保存和加载
- 通达信软件路径
- 存储路径
- 定时任务设置

配置数据以JSON格式存储在本地文件中。

版本: v1.0
作者: 小张
日期: 2026-04-04
"""

import json
import os
import base64
from pathlib import Path


class ConfigManager:
    """
    配置管理器类
    
    负责读取和保存用户配置信息，使用简单的加密算法保护敏感数据。
    
    配置存储位置: 当前目录下的 config.json
    
    配置数据结构:
    {
        "username": "用户账号",
        "password": "加密后的密码",
        "tdx_path": "通达信软件路径",
        "save_path": "存储路径",
        "timing_12_35": true/false,
        "timing_15_35": true/false
    }
    """

    def __init__(self):
        """初始化配置管理器，设置配置文件路径"""
        # 配置文件存储在程序运行目录下
        self.config_dir = Path(__file__).parent
        self.config_file = self.config_dir / "config.json"
        
        # 简单的XOR加密密钥（仅用于防止明文存储，不能用于安全目的）
        self.encryption_key = b"tongdaxin_tool_2024"

    def _encrypt_password(self, password):
        """
        简单加密密码
        
        使用XOR算法对密码进行简单加密，防止配置文件被直接查看时密码明文显示。
        注意：这不是真正的加密，仅供防止一般性查看。
        
        Args:
            password: 明文密码
            
        Returns:
            str: Base64编码的加密密码
        """
        if not password:
            return ""
        
        # 将密码转换为字节
        password_bytes = password.encode('utf-8')
        
        # XOR加密
        encrypted = bytes(a ^ b for a, b in zip(password_bytes, self.encryption_key * (len(password_bytes) // len(self.encryption_key) + 1)))
        
        # 返回Base64编码
        return base64.b64encode(encrypted).decode('utf-8')

    def _decrypt_password(self, encrypted_password):
        """
        解密密码
        
        Args:
            encrypted_password: Base64编码的加密密码
            
        Returns:
            str: 明文密码
        """
        if not encrypted_password:
            return ""
        
        try:
            # Base64解码
            encrypted = base64.b64decode(encrypted_password.encode('utf-8'))
            
            # XOR解密
            decrypted = bytes(a ^ b for a, b in zip(encrypted, self.encryption_key * (len(encrypted) // len(self.encryption_key) + 1)))
            
            return decrypted.decode('utf-8')
        except Exception:
            return ""

    def load_config(self):
        """
        加载配置文件
        
        从JSON文件中读取配置数据，包括已保存的账号密码等信息。
        
        Returns:
            dict: 配置数据字典，如果文件不存在或读取失败返回空字典
        """
        if not self.config_file.exists():
            return {}
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            # 解密密码
            if 'password' in config and config['password']:
                config['password'] = self._decrypt_password(config['password'])
            
            return config
            
        except Exception as e:
            print(f"加载配置失败: {e}")
            return {}

    def save_config(self, config_data):
        """
        保存配置到文件
        
        将配置数据加密后写入JSON文件。
        
        Args:
            config_data: 配置数据字典，包含username、password、tdx_path等字段
        """
        # 创建配置副本，避免修改原始数据
        config_to_save = config_data.copy()
        
        # 加密密码
        if 'password' in config_to_save and config_to_save['password']:
            config_to_save['password'] = self._encrypt_password(config_to_save['password'])
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=4)
                
        except Exception as e:
            print(f"保存配置失败: {e}")

    def get_default_tdx_path(self):
        """
        获取通达信默认安装路径
        
        检查常见的通达信安装位置，返回第一个找到的路径。
        
        Returns:
            str: 通达信可执行文件路径，如果未找到返回空字符串
        """
        # 常见的通达信安装路径
        common_paths = [
            r"C:\同花顺\xiadan.exe",
            r"C:\Program Files\TongDaXin\tdx.exe",
            r"C:\Program Files (x86)\TongDaXin\tdx.exe",
            r"C:\通达信\tdx.exe",
            r"C:\Users\Public\桌面\通达信.lnk",
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                return path
        
        return ""

    def get_default_save_path(self):
        """
        获取默认的存储路径
        
        使用用户主目录下的"下载"文件夹作为默认存储位置。
        
        Returns:
            str: 默认存储路径
        """
        return os.path.join(os.path.expanduser("~"), "Downloads")

    def clear_config(self):
        """
        清除所有保存的配置
        
        删除配置文件，用于用户需要重置所有设置的情况。
        """
        try:
            if self.config_file.exists():
                os.remove(self.config_file)
            return True
        except Exception as e:
            print(f"清除配置失败: {e}")
            return False
