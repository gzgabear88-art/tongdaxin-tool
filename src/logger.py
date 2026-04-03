#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
操作日志模块
记录所有操作步骤和执行结果
"""

import os
import json
import logging
from datetime import datetime, timedelta
from pathlib import Path


class OperationLogger:
    """操作日志记录器"""

    def __init__(self):
        # 日志目录
        self.log_dir = Path(__file__).parent / "logs"
        self.log_dir.mkdir(exist_ok=True)

        # 配置文件目录
        self.config_dir = Path(__file__).parent
        self.operations_file = self.config_dir / "operations.json"

        # 初始化日志记录器
        self.logger = logging.getLogger("TongDaXin")
        self.logger.setLevel(logging.DEBUG)

        # 控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter(
            '%(asctime)s  %(levelname)-8s  %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(console_formatter)

        # 文件处理器
        log_file = self.log_dir / f"tongdaxin_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter(
            '%(asctime)s  %(levelname)-8s  %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(file_formatter)

        # 避免重复添加处理器
        if not self.logger.handlers:
            self.logger.addHandler(console_handler)
            self.logger.addHandler(file_handler)

        # 今日操作记录
        self.today_operations = []

    def info(self, message: str):
        """记录信息"""
        self.logger.info(message)

    def warning(self, message: str):
        """记录警告"""
        self.logger.warning(message)

    def error(self, message: str):
        """记录错误"""
        self.logger.error(message)

    def debug(self, message: str):
        """记录调试信息"""
        self.logger.debug(message)

    def log_operation(self, operation: str, status: str, detail: str = ""):
        """
        记录操作结果

        参数：
            operation: 操作名称
            status: 状态（成功/失败）
            detail: 详细信息
        """
        record = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "operation": operation,
            "status": status,
            "detail": detail
        }

        self.today_operations.append(record)
        self._save_operations()

        # 同时记录到日志文件
        if status == "成功":
            self.logger.info(f"[{operation}] ✅ {status} - {detail}")
        else:
            self.logger.error(f"[{operation}] ❌ {status} - {detail}")

    def _save_operations(self):
        """保存操作记录到文件"""
        try:
            data = {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "operations": self.today_operations
            }

            with open(self.operations_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        except Exception as e:
            self.logger.error(f"保存操作记录失败：{e}")

    def get_today_report(self) -> list:
        """获取今日操作报告"""
        report = []

        if not self.today_operations:
            report.append("今日暂无操作记录")
            return report

        success_count = sum(1 for op in self.today_operations if op["status"] == "成功")
        fail_count = len(self.today_operations) - success_count

        report.append(f"总操作数：{len(self.today_operations)}")
        report.append(f"成功：{success_count}，失败：{fail_count}")

        report.append("")
        report.append("详细记录：")
        for op in self.today_operations:
            status_icon = "✅" if op["status"] == "成功" else "❌"
            report.append(
                f"  {status_icon} [{op['time']}] {op['operation']} - {op['detail']}"
            )

        return report

    def get_operation_summary(self) -> dict:
        """获取操作汇总"""
        return {
            "date": datetime.now().strftime("%Y-%m-%d"),
            "total": len(self.today_operations),
            "success": sum(1 for op in self.today_operations if op["status"] == "成功"),
            "failed": sum(1 for op in self.today_operations if op["status"] == "失败"),
            "last_operation": self.today_operations[-1] if self.today_operations else None
        }

    def cleanup_old_logs(self, keep_days: int = 30):
        """清理旧的日志文件"""
        try:
            cutoff_date = datetime.now() - timedelta(days=keep_days)

            for log_file in self.log_dir.glob("tongdaxin_*.log"):
                if log_file.stat().st_mtime < cutoff_date.timestamp():
                    log_file.unlink()
                    self.logger.info(f"已删除旧日志：{log_file.name}")

        except Exception as e:
            self.logger.error(f"清理日志失败：{e}")
