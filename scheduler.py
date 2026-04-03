#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定时任务调度器 - Task Scheduler

负责管理定时任务的设置和执行。

支持功能：
- 设置多个定时时间点
- 每天在指定时间自动执行任务
- 启动/停止调度器
- 查看调度状态

版本: v1.0
作者: 小张
日期: 2026-04-04
"""

import time
import threading
import logging
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

logger = logging.getLogger(__name__)


class TaskScheduler:
    """
    定时任务调度器类
    
    负责管理定时任务的配置和执行。
    使用APScheduler库实现后台定时调度功能。
    
    使用示例：
        scheduler = TaskScheduler(parent_app)
        scheduler.set_schedule(["12:35", "15:35"])
        scheduler.start()
    """

    def __init__(self, parent=None):
        """
        初始化调度器
        
        Args:
            parent: 父窗口引用，用于回调更新界面
        """
        self.parent = parent
        self.scheduler = None
        self.is_running = False
        self.scheduled_times = []
        self.job_id_prefix = "tongdaxin_task_"

    def set_schedule(self, times):
        """
        设置定时执行时间
        
        Args:
            times: 时间列表，格式为 ["HH:MM", "HH:MM", ...]
                              例如：["12:35", "15:35"]
        """
        self.scheduled_times = times
        logger.info(f"定时任务已设置为: {times}")

    def start(self):
        """
        启动调度器
        
        创建后台调度器，添加定时任务，开始执行。
        """
        if self.is_running:
            logger.warning("调度器已在运行中")
            return
        
        if not self.scheduled_times:
            logger.error("请先设置定时时间")
            return
        
        # 创建后台调度器
        self.scheduler = BackgroundScheduler()
        
        # 为每个定时时间添加任务
        for time_str in self.scheduled_times:
            hour, minute = map(int, time_str.split(':'))
            
            # 添加定时任务（每天执行）
            job_id = f"{self.job_id_prefix}{hour}_{minute}"
            
            self.scheduler.add_job(
                func=self._execute_scheduled_task,
                trigger=CronTrigger(hour=hour, minute=minute),
                id=job_id,
                name=f"通达信自动下载 {time_str}",
                replace_existing=True
            )
            
            logger.info(f"已添加定时任务: 每天 {hour:02d}:{minute:02d}")
        
        # 启动调度器
        self.scheduler.start()
        self.is_running = True
        
        logger.info("定时任务调度器已启动")
        
        if self.parent and hasattr(self.parent, 'log_message'):
            self.parent.log_message(f"✅ 定时调度器已启动，等待执行时间: {', '.join(self.scheduled_times)}")

    def stop(self):
        """
        停止调度器
        
        关闭后台调度器，停止所有定时任务。
        """
        if not self.is_running:
            logger.warning("调度器未在运行")
            return
        
        if self.scheduler:
            self.scheduler.shutdown(wait=False)
            self.scheduler = None
        
        self.is_running = False
        logger.info("定时任务调度器已停止")
        
        if self.parent and hasattr(self.parent, 'log_message'):
            self.parent.log_message("🛑 定时调度器已停止")

    def pause(self):
        """
        暂停调度器
        
        暂停所有定时任务的执行，但保留任务配置。
        """
        if self.scheduler and self.is_running:
            self.scheduler.pause()
            logger.info("调度器已暂停")
            
            if self.parent and hasattr(self.parent, 'log_message'):
                self.parent.log_message("⏸️ 定时调度器已暂停")

    def resume(self):
        """
        恢复调度器
        
        恢复暂停的定时任务执行。
        """
        if self.scheduler and not self.is_running:
            self.scheduler.resume()
            self.is_running = True
            logger.info("调度器已恢复")
            
            if self.parent and hasattr(self.parent, 'log_message'):
                self.parent.log_message("▶️ 定时调度器已恢复")

    def _execute_scheduled_task(self):
        """
        执行定时任务的回调函数
        
        当到达设定的执行时间时，自动调用此函数。
        会通过父窗口的回调函数触发自动化任务执行。
        """
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logger.info(f"定时任务触发: {current_time}")
        
        if self.parent and hasattr(self.parent, 'log_message'):
            self.parent.log_message(f"=" * 50)
            self.parent.log_message(f"⏰ 定时任务触发! 时间: {current_time}")
            
            # 调用父窗口的执行方法
            if hasattr(self.parent, 'execute_immediate'):
                self.parent.log_message("开始执行自动化任务...")
                # 在后台线程中执行，避免阻塞调度器
                thread = threading.Thread(
                    target=self.parent.execute_immediate,
                    daemon=True
                )
                thread.start()

    def get_status(self):
        """
        获取调度器状态
        
        Returns:
            dict: 包含is_running和scheduled_times的字典
        """
        return {
            'is_running': self.is_running,
            'scheduled_times': self.scheduled_times,
            'next_run_times': self._get_next_run_times()
        }

    def _get_next_run_times(self):
        """
        获取下次执行时间列表
        
        Returns:
            list: 下次各定时任务的执行时间
        """
        if not self.scheduler or not self.is_running:
            return []
        
        next_times = []
        for job in self.scheduler.get_jobs():
            next_run = job.next_run_time
            if next_run:
                next_times.append(next_run.strftime("%Y-%m-%d %H:%M:%S"))
        
        return next_times

    def add_time(self, time_str):
        """
        动态添加定时时间
        
        Args:
            time_str: 时间字符串，格式为 "HH:MM"
        """
        if time_str not in self.scheduled_times:
            self.scheduled_times.append(time_str)
            
            if self.is_running and self.scheduler:
                hour, minute = map(int, time_str.split(':'))
                
                job_id = f"{self.job_id_prefix}{hour}_{minute}"
                
                self.scheduler.add_job(
                    func=self._execute_scheduled_task,
                    trigger=CronTrigger(hour=hour, minute=minute),
                    id=job_id,
                    name=f"通达信自动下载 {time_str}",
                    replace_existing=True
                )
                
                logger.info(f"已添加定时时间: {time_str}")

    def remove_time(self, time_str):
        """
        移除定时时间
        
        Args:
            time_str: 要移除的时间字符串，格式为 "HH:MM"
        """
        if time_str in self.scheduled_times:
            self.scheduled_times.remove(time_str)
            
            if self.is_running and self.scheduler:
                hour, minute = map(int, time_str.split(':'))
                job_id = f"{self.job_id_prefix}{hour}_{minute}"
                
                self.scheduler.remove_job(job_id)
                logger.info(f"已移除定时时间: {time_str}")
