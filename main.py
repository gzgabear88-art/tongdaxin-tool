#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通达信自动下载工具 - 主程序入口
TongDaXin Auto Download Tool - Main Entry Point

版本: v1.0
作者: 小张
日期: 2026-04-04

功能说明:
    本程序用于自动化操作通达信金融终端软件，帮助用户自动完成以下工作流程：
    1. 打开通达信软件
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
    通达信自动下载工具 - 主窗口类
    
    负责管理整个应用程序的GUI界面和用户交互流程。
    
    窗口布局:
    ┌─────────────────────────────────────────────┐
    │  通达信自动下载工具 v1.0                      │
    ├─────────────────────────────────────────────┤
    │  【执行模式选择】                              │
    │  ○ 立即执行  ○ 定时执行  ○ 终止任务           │
    ├─────────────────────────────────────────────┤
    │  【账号设置】                                │
    │  账号: [___________] 密码: [___________]    │
    │  ☑ 保存账号密码                              │
    ├─────────────────────────────────────────────┤
    │  【通达信路径】                              │
    │  [C:/path/to/tdx.exe_______] [浏览...]    │
    ├─────────────────────────────────────────────┤
    │  【存储路径】                                │
    │  [C:/Users/Downloads____] [浏览...]       │
    ├─────────────────────────────────────────────┤
    │  【定时时间设置】                            │
    │  ☑ 12:35  ☑ 15:35                         │
    ├─────────────────────────────────────────────┤
    │           [开始执行]  [查看日志]              │
    └─────────────────────────────────────────────┘
    """

    def __init__(self, root):
        """
        初始化主窗口
        
        Args:
            root: tkinter根窗口对象
        """
        self.root = root
        self.root.title("通达信自动下载工具 v1.0")
        self.root.geometry("600x650")
        self.root.resizable(False, False)
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        
        # 初始化自动化操作引擎
        self.auto_op = AutoOperation(self)
        
        # 初始化通知器
        self.notifier = Notifier()
        
        # 初始化任务调度器
        self.scheduler = TaskScheduler(self)
        
        # 当前任务状态
        self.is_running = False
        self.current_task = None
        
        # 加载保存的配置
        self.load_saved_config()
        
        # 创建界面
        self.create_widgets()
        
        logger.info("应用程序初始化完成")

    def load_saved_config(self):
        """
        加载保存的配置信息
        
        从本地JSON文件读取之前保存的账号、路径等信息，
        如果文件不存在或读取失败，使用空值。
        """
        saved_config = self.config_manager.load_config()
        if saved_config:
            self.saved_username = saved_config.get('username', '')
            self.saved_password = saved_config.get('password', '')
            self.tdx_path = saved_config.get('tdx_path', '')
            self.save_path = saved_config.get('save_path', '')
            self.timing_12_35 = saved_config.get('timing_12_35', True)
            self.timing_15_35 = saved_config.get('timing_15_35', True)
            logger.info("已加载保存的配置")
        else:
            self.saved_username = ''
            self.saved_password = ''
            self.tdx_path = ''
            self.save_path = ''
            self.timing_12_35 = True
            self.timing_15_35 = True
            logger.info("使用默认配置")

    def save_config(self):
        """
        保存当前配置到本地文件
        
        将用户输入的账号、路径、定时设置等信息保存到JSON文件，
        以便下次启动时自动填充。
        """
        config_data = {
            'username': self.username_var.get(),
            'password': self.password_var.get(),
            'tdx_path': self.tdx_path_var.get(),
            'save_path': self.save_path_var.get(),
            'timing_12_35': self.timing_12_35_var.get(),
            'timing_15_35': self.timing_15_35_var.get()
        }
        self.config_manager.save_config(config_data)
        logger.info("配置已保存")

    def create_widgets(self):
        """
        创建所有GUI组件
        
        按照功能分区创建界面组件，包括：
        - 执行模式选择区
        - 账号设置区
        - 路径设置区
        - 定时设置区
        - 操作按钮区
        """
        # ========== 标题区域 ==========
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=20, pady=10)
        
        title_label = ttk.Label(
            title_frame, 
            text="通达信自动下载工具 v1.0",
            font=('Microsoft YaHei', 16, 'bold')
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            title_frame,
            text="自动化操作通达信，下载盘后数据",
            font=('Microsoft YaHei', 9),
            foreground='gray'
        )
        subtitle_label.pack()

        # ========== 执行模式选择区 ==========
        mode_frame = ttk.LabelFrame(self.root, text="执行模式选择", padding=10)
        mode_frame.pack(fill=tk.X, padx=20, pady=5)

        self.mode_var = tk.StringVar(value="immediate")
        
        ttk.Radiobutton(
            mode_frame, 
            text="立即执行", 
            variable=self.mode_var, 
            value="immediate"
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            mode_frame, 
            text="定时执行", 
            variable=self.mode_var, 
            value="scheduled"
        ).pack(side=tk.LEFT, padx=10)
        
        ttk.Radiobutton(
            mode_frame, 
            text="终止任务", 
            variable=self.mode_var, 
            value="terminate"
        ).pack(side=tk.LEFT, padx=10)

        # ========== 账号设置区 ==========
        account_frame = ttk.LabelFrame(self.root, text="账号设置", padding=10)
        account_frame.pack(fill=tk.X, padx=20, pady=5)

        # 账号输入
        ttk.Label(account_frame, text="账号:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar(value=self.saved_username)
        username_entry = ttk.Entry(account_frame, textvariable=self.username_var, width=25)
        username_entry.grid(row=0, column=1, padx=5, pady=5)

        # 密码输入
        ttk.Label(account_frame, text="密码:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar(value=self.saved_password)
        password_entry = ttk.Entry(account_frame, textvariable=self.password_var, show='*', width=25)
        password_entry.grid(row=0, column=3, padx=5, pady=5)

        # 保存选项
        self.save_credential_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            account_frame, 
            text="保存账号密码", 
            variable=self.save_credential_var
        ).grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=5)

        # ========== 通达信路径设置区 ==========
        path_frame = ttk.LabelFrame(self.root, text="通达信路径设置", padding=10)
        path_frame.pack(fill=tk.X, padx=20, pady=5)

        ttk.Label(path_frame, text="软件路径:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.tdx_path_var = tk.StringVar(value=self.tdx_path)
        tdx_entry = ttk.Entry(path_frame, textvariable=self.tdx_path_var, width=45)
        tdx_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        
        ttk.Button(
            path_frame, 
            text="浏览...", 
            command=self.browse_tdx_path
        ).grid(row=0, column=3, padx=5, pady=5)

        # ========== 存储路径设置区 ==========
        save_frame = ttk.LabelFrame(self.root, text="存储路径设置", padding=10)
        save_frame.pack(fill=tk.X, padx=20, pady=5)

        ttk.Label(save_frame, text="保存位置:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.save_path_var = tk.StringVar(value=self.save_path)
        save_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, width=45)
        save_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
        
        ttk.Button(
            save_frame, 
            text="浏览...", 
            command=self.browse_save_path
        ).grid(row=0, column=3, padx=5, pady=5)

        # ========== 定时时间设置区 ==========
        timing_frame = ttk.LabelFrame(self.root, text="定时时间设置", padding=10)
        timing_frame.pack(fill=tk.X, padx=20, pady=5)

        self.timing_12_35_var = tk.BooleanVar(value=self.timing_12_35)
        ttk.Checkbutton(
            timing_frame, 
            text="12:35", 
            variable=self.timing_12_35_var
        ).pack(side=tk.LEFT, padx=20)

        self.timing_15_35_var = tk.BooleanVar(value=self.timing_15_35)
        ttk.Checkbutton(
            timing_frame, 
            text="15:35", 
            variable=self.timing_15_35_var
        ).pack(side=tk.LEFT, padx=20)

        ttk.Label(
            timing_frame,
            text="(勾选后，每天在指定时间自动执行)",
            foreground='gray'
        ).pack(side=tk.LEFT, padx=10)

        # ========== 日志显示区 ==========
        log_frame = ttk.LabelFrame(self.root, text="执行日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        self.log_text = tk.Text(log_frame, height=8, width=60, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # ========== 操作按钮区 ==========
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=20, pady=10)

        self.start_button = ttk.Button(
            button_frame, 
            text="开始执行", 
            command=self.start_task,
            style='Accent.TButton'
        )
        self.start_button.pack(side=tk.LEFT, padx=10)

        ttk.Button(
            button_frame, 
            text="查看日志", 
            command=self.open_log_file
        ).pack(side=tk.LEFT, padx=10)

        ttk.Button(
            button_frame, 
            text="清空日志", 
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=10)

    def browse_tdx_path(self):
        """
        浏览选择通达信软件路径
        
        打开文件对话框，让用户选择通达信可执行文件
        """
        path = filedialog.askopenfilename(
            title="选择通达信软件",
            filetypes=[("可执行文件", "*.exe"), ("所有文件", "*.*")]
        )
        if path:
            self.tdx_path_var.set(path)
            self.log_message(f"已选择通达信路径: {path}")

    def browse_save_path(self):
        """
        浏览选择数据存储路径
        
        打开文件夹对话框，让用户选择Excel文件的保存目录
        """
        path = filedialog.askdirectory(title="选择存储路径")
        if path:
            self.save_path_var.set(path)
            self.log_message(f"已选择存储路径: {path}")

    def log_message(self, message):
        """
        在日志文本框中显示消息
        
        Args:
            message: 要显示的消息内容
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, full_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        logger.info(message)

    def start_task(self):
        """
        根据选择的执行模式启动任务
        
        根据用户选择的执行模式（立即/定时/终止），
        调用相应的处理函数。
        """
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
        else:
            messagebox.showerror("错误", "未知的执行模式")

    def execute_immediate(self):
        """
        立即执行自动化任务
        
        执行完整的通达信自动化操作流程：
        1. 保存配置
        2. 启动自动化引擎
        3. 执行各步骤操作
        """
        if self.is_running:
            messagebox.showwarning("警告", "任务正在执行中，请勿重复启动")
            return
        
        # 保存配置
        self.save_config()
        
        # 验证必填项
        if not self.validate_inputs():
            return
        
        self.is_running = True
        self.start_button.config(state=tk.DISABLED)
        
        self.log_message("开始执行自动化任务...")
        
        # 启动自动化执行线程
        import threading
        self.current_task = threading.Thread(target=self.run_auto_operation, daemon=True)
        self.current_task.start()

    def execute_scheduled(self):
        """
        设置定时任务
        
        根据用户选择的定时时间（12:35和/或15:35），
        配置并启动定时调度器。
        """
        # 保存定时设置
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
        """
        终止正在运行的任务
        
        向自动化引擎发送停止信号，终止当前执行的任务。
        """
        if not self.is_running:
            self.log_message("当前没有正在执行的任务")
            return
        
        self.log_message("正在终止任务...")
        self.auto_op.stop()
        self.is_running = False
        self.start_button.config(state=tk.NORMAL)
        self.log_message("任务已终止")

    def run_auto_operation(self):
        """
        在独立线程中运行自动化操作
        
        这是自动化任务的主执行函数，在后台线程中运行，
        执行完成后通过回调更新界面状态。
        """
        try:
            # 调用自动化引擎执行完整流程
            result = self.auto_op.run_full_flow(
                username=self.username_var.get(),
                password=self.password_var.get(),
                tdx_path=self.tdx_path_var.get(),
                save_path=self.save_path_var.get(),
                log_callback=self.log_message
            )
            
            if result['success']:
                self.log_message("=" * 50)
                self.log_message("🎉 任务执行成功！")
                self.log_message(f"数据已保存到: {result['save_path']}")
                
                # 发送通知
                self.notifier.send_success_notification(
                    save_path=result['save_path'],
                    execute_time=result.get('execute_time', '')
                )
            else:
                self.log_message("=" * 50)
                self.log_message(f"❌ 任务执行失败: {result.get('error', '未知错误')}")
                
                # 发送失败通知
                self.notifier.send_failure_notification(
                    error=result.get('error', '未知错误'),
                    failed_step=result.get('step', '')
                )
                
        except Exception as e:
            logger.exception("自动化任务执行异常")
            self.log_message(f"❌ 异常错误: {str(e)}")
            
        finally:
            self.is_running = False
            # 在主线程中更新界面
            self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))

    def validate_inputs(self):
        """
        验证用户输入的有效性
        
        检查必填字段是否已填写，包括账号、密码、通达信路径等。
        
        Returns:
            bool: 验证通过返回True，否则返回False
        """
        if not self.username_var.get().strip():
            messagebox.showerror("错误", "请输入账号")
            return False
        
        if not self.password_var.get().strip():
            messagebox.showerror("错误", "请输入密码")
            return False
        
        if not self.tdx_path_var.get().strip():
            messagebox.showerror("错误", "请选择通达信软件路径")
            return False
        
        if not os.path.exists(self.tdx_path_var.get().strip()):
            messagebox.showerror("错误", "通达信软件路径不存在，请检查")
            return False
        
        return True

    def open_log_file(self):
        """
        打开日志文件
        
        使用系统默认程序打开日志文件，供用户查看详细日志。
        """
        log_path = os.path.abspath('tongdaxin_tool.log')
        if os.path.exists(log_path):
            os.startfile(log_path) if sys.platform == 'win32' else os.system(f'open "{log_path}"')
        else:
            messagebox.showinfo("提示", "日志文件不存在")

    def clear_log(self):
        """
        清空日志显示区域
        
        将日志文本框中的内容清空，但不删除实际的日志文件。
        """
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)


def main():
    """
    应用程序主入口函数
    
    创建根窗口并启动应用程序的主事件循环。
    """
    root = tk.Tk()
    
    # 设置窗口样式
    style = ttk.Style()
    try:
        style.theme_use('clam')  # 使用现代主题
    except:
        pass
    
    app = TongDaXinToolApp(root)
    
    logger.info("应用程序启动")
    
    root.mainloop()


if __name__ == '__main__':
    main()
