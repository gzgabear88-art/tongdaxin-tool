#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TDX 自动下载工具 - 主程序入口
TDX Auto Download Tool - Main Entry Point

版本: v1.0
作者: 小张
日期: 2026-04-04

功能说明:
    本程序用于自动化操作 TDX 金融终端软件，帮助用户自动完成以下工作流程：
    1. 打开 TDX 软件
    2. 自动登录账号
    3. 输入板块代码
    4. 下载盘后数据
    5. 导出数据到Excel文件
    6. 完成通知用户
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import sys
import logging
from datetime import datetime

# 导入各功能模块
from config_manager import ConfigManager
from auto_operation import AutoOperation
from notifier import Notifier
from scheduler import TaskScheduler

# 配置日志系统
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('tongdaxin_tool.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class TongDaXinToolApp:
    """
    TDX 自动下载工具 - 主窗口类

    负责管理整个应用程序的GUI界面和用户交互流程。

    窗口布局:
    ┌─────────────────────────────────────────────┐
    │  TDX 自动下载工具 v1.0                      │
    ├─────────────────────────────────────────────┤
    │  【执行模式选择】                              │
    │  ○ 立即执行  ○ 定时执行  ○ 终止任务           │
    ├─────────────────────────────────────────────┤
    │  【账号设置】                                │
    │  账号: [___________] 密码: [___________]    │
    │  ☑ 保存账号密码                              │
    ├─────────────────────────────────────────────┤
    │  【TDX路径】                                │
    │  [C:/path/to/tdx.exe_______] [浏览...]    │
    ├─────────────────────────────────────────────┤
    │  【存储路径】                                │
    │  [C:/Users/Downloads____] [浏览...]       │
    ├─────────────────────────────────────────────┤
    │  【定时时间设置】                            │
    │  ☑ 12:35  ☑ 15:35                         │
    ├─────────────────────────────────────────────┤
    │              [开始执行]  [查看日志]            │
    └─────────────────────────────────────────────┘
    """

    def __init__(self, root):
        self.root = root
        self.root.title("TDX 自动下载工具 v1.0")
        self.root.geometry("600x680")
        self.root.resizable(False, False)

        # 配色方案
        self.BG_COLOR = "#F0F4F8"
        self.ACCENT_COLOR = "#2B6CB0"
        self.ACCENT_HOVER = "#2C5282"
        self.TEXT_COLOR = "#1A202C"
        self.GRAY_COLOR = "#718096"

        # 初始化模块
        self.config_manager = ConfigManager()
        self.auto_op = AutoOperation(self)
        self.notifier = Notifier()
        self.scheduler = TaskScheduler(self)

        self.is_running = False
        self.current_task = None

        self.load_saved_config()
        self.create_widgets()

        logger.info("应用程序初始化完成")

    def load_saved_config(self):
        saved_config = self.config_manager.load_config()
        if saved_config:
            self.saved_username = saved_config.get('username', '')
            self.saved_password = saved_config.get('password', '')
            self.tdx_path = saved_config.get('tdx_path', '')
            self.save_path = saved_config.get('save_path', '')
            self.timing_12_35 = saved_config.get('timing_12_35', True)
            self.timing_15_35 = saved_config.get('timing_15_35', True)
        else:
            self.saved_username = ''
            self.saved_password = ''
            self.tdx_path = ''
            self.save_path = ''
            self.timing_12_35 = True
            self.timing_15_35 = True

    def save_config(self):
        config_data = {
            'username': self.username_var.get(),
            'password': self.password_var.get(),
            'tdx_path': self.tdx_path_var.get(),
            'save_path': self.save_path_var.get(),
            'timing_12_35': self.timing_12_35_var.get(),
            'timing_15_35': self.timing_15_35_var.get()
        }
        self.config_manager.save_config(config_data)

    def create_widgets(self):
        # 全局背景
        self.root.configure(bg=self.BG_COLOR)

        # ========== 标题区域 ==========
        title_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        title_frame.pack(fill=tk.X, padx=30, pady=(20, 10))

        title_label = tk.Label(
            title_frame,
            text="TDX 自动下载工具 v1.0",
            font=('Microsoft YaHei', 18, 'bold'),
            fg=self.ACCENT_COLOR,
            bg=self.BG_COLOR
        )
        title_label.pack()

        subtitle_label = tk.Label(
            title_frame,
            text="自动化操作 TDX，下载盘后数据",
            font=('Microsoft YaHei', 10),
            fg=self.GRAY_COLOR,
            bg=self.BG_COLOR
        )
        subtitle_label.pack()

        # ========== 执行模式选择区 ==========
        mode_frame = self._make_card(self.root)
        mode_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(
            mode_frame,
            text="执行模式",
            font=('Microsoft YaHei', 10, 'bold'),
            fg=self.TEXT_COLOR,
            bg="white"
        ).pack(anchor='w', padx=15, pady=(10, 5))

        self.mode_var = tk.StringVar(value="immediate")

        mode_inner = tk.Frame(mode_frame, bg="white")
        mode_inner.pack(fill=tk.X, padx=15, pady=(0, 10))

        for text, val in [("立即执行", "immediate"), ("定时执行", "scheduled"), ("终止任务", "terminate")]:
            btn = tk.Radiobutton(
                mode_inner,
                text=text,
                variable=self.mode_var,
                value=val,
                font=('Microsoft YaHei', 10),
                fg=self.TEXT_COLOR,
                bg="white",
                activebackground="white",
                command=self._on_mode_changed
            )
            btn.pack(side=tk.LEFT, padx=(0, 20))

        # ========== 账号设置区 ==========
        account_frame = self._make_card(self.root)
        account_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(
            account_frame,
            text="账号设置",
            font=('Microsoft YaHei', 10, 'bold'),
            fg=self.TEXT_COLOR,
            bg="white"
        ).pack(anchor='w', padx=15, pady=(10, 8))

        input_inner = tk.Frame(account_frame, bg="white")
        input_inner.pack(fill=tk.X, padx=15, pady=(0, 10))

        tk.Label(input_inner, text="账号:", font=('Microsoft YaHei', 10), fg=self.TEXT_COLOR, bg="white").grid(row=0, column=0, sticky='w', pady=5)
        self.username_var = tk.StringVar(value=self.saved_username)
        username_entry = tk.Entry(input_inner, textvariable=self.username_var, font=('Microsoft YaHei', 10), bd=1, relief='groove', width=22)
        username_entry.grid(row=0, column=1, padx=(5, 20), pady=5)

        tk.Label(input_inner, text="密码:", font=('Microsoft YaHei', 10), fg=self.TEXT_COLOR, bg="white").grid(row=0, column=2, sticky='w', pady=5)
        self.password_var = tk.StringVar(value=self.saved_password)
        password_entry = tk.Entry(input_inner, textvariable=self.password_var, show='*', font=('Microsoft YaHei', 10), bd=1, relief='groove', width=22)
        password_entry.grid(row=0, column=3, padx=(5, 0), pady=5)

        self.save_credential_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            account_frame,
            text="保存账号密码",
            variable=self.save_credential_var,
            font=('Microsoft YaHei', 9),
            fg=self.TEXT_COLOR,
            bg="white",
            activebackground="white"
        ).pack(anchor='w', padx=15, pady=(0, 10))

        # ========== TDX路径设置区 ==========
        path_frame = self._make_card(self.root)
        path_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(
            path_frame,
            text="TDX 路径",
            font=('Microsoft YaHei', 10, 'bold'),
            fg=self.TEXT_COLOR,
            bg="white"
        ).pack(anchor='w', padx=15, pady=(10, 8))

        path_inner = tk.Frame(path_frame, bg="white")
        path_inner.pack(fill=tk.X, padx=15, pady=(0, 10))

        tk.Label(path_inner, text="软件:", font=('Microsoft YaHei', 10), fg=self.TEXT_COLOR, bg="white").grid(row=0, column=0, sticky='w', pady=5)
        self.tdx_path_var = tk.StringVar(value=self.tdx_path)
        tdx_entry = tk.Entry(path_inner, textvariable=self.tdx_path_var, font=('Microsoft YaHei', 10), bd=1, relief='groove', width=38)
        tdx_entry.grid(row=0, column=1, padx=(5, 5), pady=5)
        tk.Button(
            path_inner, text="浏览",
            font=('Microsoft YaHei', 9),
            bg=self.BG_COLOR,
            fg=self.TEXT_COLOR,
            relief='groove',
            bd=1,
            command=self.browse_tdx_path
        ).grid(row=0, column=2, pady=5)

        tk.Label(path_inner, text="存储:", font=('Microsoft YaHei', 10), fg=self.TEXT_COLOR, bg="white").grid(row=1, column=0, sticky='w', pady=5)
        self.save_path_var = tk.StringVar(value=self.save_path)
        save_entry = tk.Entry(path_inner, textvariable=self.save_path_var, font=('Microsoft YaHei', 10), bd=1, relief='groove', width=38)
        save_entry.grid(row=1, column=1, padx=(5, 5), pady=5)
        tk.Button(
            path_inner, text="浏览",
            font=('Microsoft YaHei', 9),
            bg=self.BG_COLOR,
            fg=self.TEXT_COLOR,
            relief='groove',
            bd=1,
            command=self.browse_save_path
        ).grid(row=1, column=2, pady=5)

        # ========== 定时时间设置区 ==========
        timing_frame = self._make_card(self.root)
        timing_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(
            timing_frame,
            text="定时执行（勾选后生效）",
            font=('Microsoft YaHei', 10, 'bold'),
            fg=self.TEXT_COLOR,
            bg="white"
        ).pack(anchor='w', padx=15, pady=(10, 8))

        timing_inner = tk.Frame(timing_frame, bg="white")
        timing_inner.pack(fill=tk.X, padx=15, pady=(0, 10))

        self.timing_12_35_var = tk.BooleanVar(value=self.timing_12_35)
        tk.Checkbutton(
            timing_inner,
            text="12:35",
            variable=self.timing_12_35_var,
            font=('Microsoft YaHei', 11),
            fg=self.ACCENT_COLOR,
            bg="white",
            activebackground="white"
        ).pack(side=tk.LEFT, padx=(0, 20))

        self.timing_15_35_var = tk.BooleanVar(value=self.timing_15_35)
        tk.Checkbutton(
            timing_inner,
            text="15:35",
            variable=self.timing_15_35_var,
            font=('Microsoft YaHei', 11),
            fg=self.ACCENT_COLOR,
            bg="white",
            activebackground="white"
        ).pack(side=tk.LEFT, padx=0)

        # ========== 日志显示区 ==========
        log_frame = self._make_card(self.root)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        tk.Label(
            log_frame,
            text="执行日志",
            font=('Microsoft YaHei', 10, 'bold'),
            fg=self.TEXT_COLOR,
            bg="white"
        ).pack(anchor='w', padx=15, pady=(10, 5))

        log_inner = tk.Frame(log_frame, bg="white")
        log_inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        self.log_text = tk.Text(
            log_inner,
            height=6,
            font=('Consolas', 9),
            bd=1,
            relief='groove',
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # ========== 操作按钮区 ==========
        button_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        button_frame.pack(fill=tk.X, padx=20, pady=(5, 15))

        self.start_button = tk.Button(
            button_frame,
            text="▶ 开始执行",
            font=('Microsoft YaHei', 12, 'bold'),
            bg=self.ACCENT_COLOR,
            fg="white",
            relief='flat',
            bd=0,
            padx=20,
            pady=8,
            command=self.start_task,
            cursor='hand2'
        )
        self.start_button.pack(side=tk.LEFT, padx=5)

        tk.Button(
            button_frame,
            text="📄 日志",
            font=('Microsoft YaHei', 10),
            bg=self.BG_COLOR,
            fg=self.TEXT_COLOR,
            relief='groove',
            bd=1,
            padx=15,
            pady=8,
            command=self.open_log_file
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            button_frame,
            text="🗑 清空",
            font=('Microsoft YaHei', 10),
            bg=self.BG_COLOR,
            fg=self.TEXT_COLOR,
            relief='groove',
            bd=1,
            padx=15,
            pady=8,
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=5)

    def _make_card(self, parent):
        """创建白色卡片风格的 Frame"""
        return tk.Frame(parent, bg="white", bd=1, relief='groove')

    def _on_mode_changed(self):
        """执行模式切换回调"""
        mode = self.mode_var.get()
        if mode == "scheduled":
            self.start_button.config(text="▶ 设置定时")
        elif mode == "terminate":
            self.start_button.config(text="■ 终止任务")
        else:
            self.start_button.config(text="▶ 开始执行")

    def browse_tdx_path(self):
        path = filedialog.askopenfilename(
            title="选择 TDX 软件",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if path:
            self.tdx_path_var.set(path)
            self.log_message(f"已选择 TDX 路径: {path}")

    def browse_save_path(self):
        path = filedialog.askdirectory(title="选择存储路径")
        if path:
            self.save_path_var.set(path)
            self.log_message(f"已选择存储路径: {path}")

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, full_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        logger.info(message)

    def start_task(self):
        mode = self.mode_var.get()
        if mode == "immediate":
            self.log_message("=" * 50)
            self.log_message("用户选择：立即执行模式")
            self.execute_immediate()
        elif mode == "scheduled":
            self.log_message("=" * 50)
            self.log_message("用户选择：定时执行模式")
            self.execute_scheduled()
        elif mode == "terminate":
            self.log_message("=" * 50)
            self.log_message("用户选择：终止任务")
            self.terminate_task()

    def execute_immediate(self):
        if self.is_running:
            messagebox.showwarning("警告", "任务正在执行中，请勿重复启动")
            return
        self.save_config()
        if not self.validate_inputs():
            return
        self.is_running = True
        self.start_button.config(state=tk.DISABLED, text="▶ 执行中...")
        self.log_message("开始执行自动化任务...")
        import threading
        self.current_task = threading.Thread(target=self.run_auto_operation, daemon=True)
        self.current_task.start()

    def execute_scheduled(self):
        self.save_config()
        times = []
        if self.timing_12_35_var.get():
            times.append("12:35")
        if self.timing_15_35_var.get():
            times.append("15:35")
        if not times:
            messagebox.showerror("错误", "请至少选择一个定时时间")
            return
        self.scheduler.set_schedule(times)
        self.scheduler.start()
        self.log_message(f"定时任务已设置: {', '.join(times)}")
        messagebox.showinfo("成功", f"定时任务已设置:\n{', '.join(times)}")

    def terminate_task(self):
        if not self.is_running:
            self.log_message("当前没有正在执行的任务")
            return
        self.log_message("正在终止任务...")
        self.auto_op.stop()
        self.is_running = False
        self.start_button.config(state=tk.NORMAL, text="▶ 开始执行")
        self.log_message("任务已终止")

    def run_auto_operation(self):
        try:
            result = self.auto_op.run_full_flow(
                username=self.username_var.get(),
                password=self.password_var.get(),
                tdx_path=self.tdx_path_var.get(),
                save_path=self.save_path_var.get(),
                log_callback=self.log_message
            )
            if result['success']:
                self.log_message("=" * 50)
                self.log_message("任务执行成功！")
                self.log_message(f"数据已保存到: {result['save_path']}")
                self.notifier.send_success_notification(
                    save_path=result['save_path'],
                    execute_time=result.get('execute_time', '')
                )
            else:
                self.log_message("=" * 50)
                self.log_message(f"任务执行失败: {result.get('error', '未知错误')}")
                self.notifier.send_failure_notification(
                    error=result.get('error', '未知错误'),
                    failed_step=result.get('step', '')
                )
        except Exception as e:
            logger.exception("自动化任务执行异常")
            self.log_message(f"异常错误: {str(e)}")
        finally:
            self.is_running = False
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL, text="▶ 开始执行"))

    def validate_inputs(self):
        if not self.username_var.get().strip():
            messagebox.showerror("错误", "请输入账号")
            return False
        if not self.password_var.get().strip():
            messagebox.showerror("错误", "请输入密码")
            return False
        if not self.tdx_path_var.get().strip():
            messagebox.showerror("错误", "请选择 TDX 软件路径")
            return False
        if not os.path.exists(self.tdx_path_var.get().strip()):
            messagebox.showerror("错误", "TDX 软件路径不存在，请检查")
            return False
        return True

    def open_log_file(self):
        log_path = os.path.abspath('tongdaxin_tool.log')
        if os.path.exists(log_path):
            os.startfile(log_path) if sys.platform == 'win32' else os.system(f'open "{log_path}"')
        else:
            messagebox.showinfo("提示", "日志文件不存在")

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = TongDaXinToolApp(root)
    logger.info("应用程序启动")
    root.mainloop()


if __name__ == '__main__':
    main()
