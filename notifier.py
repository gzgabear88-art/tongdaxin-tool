#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通知模块 - Notifier

负责向用户发送任务完成通知。

当前支持的通知方式：
- 控制台输出（调试用）
- 日志文件记录
- 未来可扩展：邮件、短信、微信等

版本: v1.0
作者: 小张
日期: 2026-04-04
"""

import logging
import time
from datetime import datetime

logger = logging.getLogger(__name__)


class Notifier:
    """
    通知发送器类
    
    负责向用户发送任务执行结果的通知。
    当前实现为日志记录，后续可扩展其他通知方式。
    
    通知内容格式：
    - 任务类型（成功/失败）
    - 执行时间
    - 详细信息（文件路径、错误原因等）
    """

    def __init__(self):
        """初始化通知器"""
        self.notification_history = []

    def send_success_notification(self, save_path, execute_time=""):
        """
        发送任务成功通知
        
        当自动化任务成功完成后调用此方法发送通知。
        
        Args:
            save_path: 数据文件保存路径
            execute_time: 执行时间（格式：YYYY-MM-DD HH:MM:SS）
        """
        if not execute_time:
            execute_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        message = f"""
╔══════════════════════════════════════════════════════════╗
║           🎉 通达信自动下载任务已完成！                  ║
╠══════════════════════════════════════════════════════════╣
║ 执行时间: {execute_time}
║ 保存路径: {save_path}
║ 任务状态: ✅ 成功完成
╚══════════════════════════════════════════════════════════╝
        """
        
        # 记录到日志
        logger.info(f"任务成功通知: {message}")
        
        # 保存通知历史
        self.notification_history.append({
            'type': 'success',
            'time': execute_time,
            'save_path': save_path,
            'message': message
        })
        
        return message

    def send_failure_notification(self, error, failed_step="", execute_time=""):
        """
        发送任务失败通知
        
        当自动化任务执行失败时调用此方法发送通知。
        
        Args:
            error: 错误描述信息
            failed_step: 失败发生的步骤名称
            execute_time: 执行时间（格式：YYYY-MM-DD HH:MM:SS）
        """
        if not execute_time:
            execute_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        message = f"""
╔══════════════════════════════════════════════════════════╗
║           ❌ 通达信自动下载任务执行失败                   ║
╠══════════════════════════════════════════════════════════╣
║ 执行时间: {execute_time}
║ 失败步骤: {failed_step}
║ 错误原因: {error}
║ 任务状态: ❌ 执行失败
╠══════════════════════════════════════════════════════════╣
║ 建议：请检查通达信软件是否正常运行，或重新执行任务。      ║
╚══════════════════════════════════════════════════════════╝
        """
        
        # 记录到日志
        logger.error(f"任务失败通知: {message}")
        
        # 保存通知历史
        self.notification_history.append({
            'type': 'failure',
            'time': execute_time,
            'failed_step': failed_step,
            'error': error,
            'message': message
        })
        
        return message

    def send_progress_notification(self, step, progress, detail=""):
        """
        发送任务进度通知
        
        Args:
            step: 当前步骤名称
            progress: 进度百分比（0-100）
            detail: 详细描述
        """
        message = f"[进度 {progress}%] {step}"
        if detail:
            message += f" - {detail}"
        
        logger.info(message)
        return message

    def get_notification_history(self):
        """
        获取通知历史记录
        
        Returns:
            list: 通知历史列表，每条记录是一个字典
        """
        return self.notification_history

    def clear_history(self):
        """清空通知历史"""
        self.notification_history = []
        logger.info("通知历史已清空")
