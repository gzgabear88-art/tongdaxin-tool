#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通达信金融终端 - 定时数据下载工具
定时在 12:05 和 15:05 自动执行数据下载操作
作者: AI Assistant
版本: 1.0.0
"""

import sys
import os

# 防止 PyInstaller 打包时控制台窗口闪烁
import ctypes

def is_frozen():
    """检测是否被 PyInstaller 打包"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

if is_frozen():
    # 打包模式下，添加所有依赖到路径
    app_dir = os.path.dirname(sys.executable)
    sys.path.insert(0, app_dir)

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import logging
from datetime import datetime

from config_manager import ConfigManager
from automation import TongDaXinAutomation
from scheduler import TaskScheduler
from notifier import NotificationService
from logger import OperationLogger

# ============================================================
# 全局样式配置
# ============================================================
BG_COLOR = "#1a1a2e"       # 深蓝黑背景
FG_COLOR = "#eaeaea"       # 浅灰白文字
ACCENT_COLOR = "#0f4c75"   # 深蓝强调
BTN_COLOR = "#3282b8"      # 蓝色按钮
BTN_HOVER = "#4a9fd4"      # 按钮悬停
SUCCESS_COLOR = "#00c853"  # 绿色（成功）
ERROR_COLOR = "#ff5252"     # 红色（失败）
WARN_COLOR = "#ffab00"     # 黄色（警告）
TEXT_BG = "#16213e"        # 文本区背景

FONT_MAIN = ("微软雅黑", 10)
FONT_TITLE = ("微软雅黑", 14, "bold")
FONT_BTN = ("微软雅黑", 10)


class ToolTip:
    """鼠标悬停提示"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip, text=self.text, justify=tk.LEFT,
                         background="#2d2d2d", foreground="#ffffff",
                         relief=tk.SOLID, borderwidth=1, font=("微软雅黑", 9))
        label.pack()

    def hide(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


class LogHandler(logging.Handler):
    """自定义日志处理器，将日志输出到 GUI 文本框"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.formatter = logging.Formatter(
            '%(asctime)s  %(levelname)-8s  %(message)s',
            datefmt='%H:%M:%S'
        )

    def emit(self, record):
        if record.levelno >= self.level:
            msg = self.formatter.format(record) + "\n"
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg)

            # 根据日志级别着色
            tag = "INFO"
            if record.levelno >= logging.ERROR:
                tag = "ERROR"
            elif record.levelno >= logging.WARNING:
                tag = "WARNING"

            # 应用标签（如果标签存在）
            try:
                self.text_widget.tag_add(tag, f"end-2l", "end-1l")
                self.text_widget.tag_config(tag, foreground=self._get_color(tag))
            except:
                pass

            self.text_widget.see(tk.END)
            self.text_widget.configure(state='disabled')

    def _get_color(self, tag):
        colors = {
            "INFO": "#eaeaea",
            "WARNING": "#ffab00",
            "ERROR": "#ff5252",
            "SUCCESS": "#00c853"
        }
        return colors.get(tag, "#eaeaea")


class TongDaXinToolApp:
    """主应用程序类"""

    def __init__(self, root):
        self.root = root
        self.root.title("通达信定时下载工具 v1.0")
        self.root.geometry("800x620")
        self.root.minsize(700, 550)
        self.root.configure(bg=BG_COLOR)

        # 居中显示
        self._center_window()

        # 初始化各模块
        self.config = ConfigManager()
        self.logger = OperationLogger()
        self.notifier = NotificationService(self.config)
        self.automation = TongDaXinAutomation(self.logger)
        self.scheduler = TaskScheduler(
            self.automation,
            self.logger,
            self.notifier,
            self.config
        )

        # 状态变量
        self.is_running = False
        self.scheduler_thread = None

        # 构建界面
        self._setup_styles()
        self._create_widgets()
        self._bind_events()
        self._load_config()

        # 启动调度线程
        self._start_scheduler()

        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        self.logger.info("程序启动完成")

    def _center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')

    def _setup_styles(self):
        """配置 ttk 样式"""
        style = ttk.Style()
        style.theme_use('clam')

        # Frame 样式
        style.configure("TFrame", background=BG_COLOR)

        # Label 样式
        style.configure("TLabel", background=BG_COLOR, foreground=FG_COLOR,
                        font=FONT_MAIN)

        # Button 样式
        style.configure("TButton", font=FONT_BTN, padding=(10, 5))
        style.map("TButton",
                  background=[('active', BTN_HOVER)],
                  foreground=[('active', FG_COLOR)])

        # Entry 样式
        style.configure("TEntry", fieldbackground=TEXT_BG, foreground=FG_COLOR,
                        insertcolor=FG_COLOR, borderwidth=0)

    def _create_widgets(self):
        """创建所有控件"""
        # ========== 标题区域 ==========
        title_frame = tk.Frame(self.root, bg=BG_COLOR, height=60)
        title_frame.pack(fill=tk.X, padx=20, pady=(15, 10))

        tk.Label(title_frame, text="📈 通达信定时下载工具",
                 font=FONT_TITLE, fg=FG_COLOR, bg=BG_COLOR).pack(side=tk.LEFT)

        # 状态指示
        self.status_frame = tk.Frame(title_frame, bg=BG_COLOR)
        self.status_frame.pack(side=tk.RIGHT)
        self.status_dot = tk.Label(self.status_frame, text="●",
                                   font=("Arial", 12), fg=SUCCESS_COLOR, bg=BG_COLOR)
        self.status_dot.pack(side=tk.LEFT, padx=(0, 5))
        self.status_label = tk.Label(self.status_frame, text="服务正常",
                                     font=FONT_MAIN, fg=SUCCESS_COLOR, bg=BG_COLOR)
        self.status_label.pack(side=tk.LEFT)

        # ========== 主内容区域 ==========
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # ----- 左侧配置区 -----
        left_frame = tk.Frame(main_frame, bg=BG_COLOR)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 配置卡片
        config_card = tk.Frame(left_frame, bg=TEXT_BG, relief=tk.FLAT)
        config_card.pack(fill=tk.X, pady=(0, 15))

        tk.Label(config_card, text="⚙️ 基础配置",
                 font=("微软雅黑", 11, "bold"), fg=FG_COLOR,
                 bg=TEXT_BG).grid(row=0, column=0, columnspan=2,
                                  sticky='w', padx=15, pady=(12, 8))

        # 通达信路径
        tk.Label(config_card, text="通达信路径：",
                 font=FONT_MAIN, fg=FG_COLOR, bg=TEXT_BG) \
            .grid(row=1, column=0, sticky='w', padx=15, pady=5)
        self.path_var = tk.StringVar()
        path_entry = ttk.Entry(config_card, textvariable=self.path_var, width=28)
        path_entry.grid(row=1, column=1, sticky='ew', padx=(0, 15), pady=5)

        # 通知手机号
        tk.Label(config_card, text="通知手机号：",
                 font=FONT_MAIN, fg=FG_COLOR, bg=TEXT_BG) \
            .grid(row=2, column=0, sticky='w', padx=15, pady=5)
        self.phone_var = tk.StringVar()
        phone_entry = ttk.Entry(config_card, textvariable=self.phone_var, width=28)
        phone_entry.grid(row=2, column=1, sticky='ew', padx=(0, 15), pady=5)

        # 保存按钮
        save_btn = tk.Button(config_card, text="保存配置",
                             command=self._save_config,
                             bg=BTN_COLOR, fg=FG_COLOR,
                             font=FONT_BTN, relief=tk.FLAT,
                             cursor='hand2', padx=15, pady=5)
        save_btn.grid(row=3, column=0, columnspan=2, padx=15, pady=(8, 12))

        config_card.columnconfigure(1, weight=1)

        # ----- 定时任务卡片 -----
        schedule_card = tk.Frame(left_frame, bg=TEXT_BG, relief=tk.FLAT)
        schedule_card.pack(fill=tk.X, pady=(0, 15))

        tk.Label(schedule_card, text="⏰ 定时任务",
                 font=("微软雅黑", 11, "bold"), fg=FG_COLOR,
                 bg=TEXT_BG).grid(row=0, column=0, columnspan=2,
                                  sticky='w', padx=15, pady=(12, 8))

        # 启用定时
        self.enable_schedule_var = tk.BooleanVar(value=True)
        tk.Checkbutton(schedule_card, text="启用定时下载任务",
                       variable=self.enable_schedule_var,
                       bg=TEXT_BG, fg=FG_COLOR, selectcolor=ACCENT_COLOR,
                       activebackground=TEXT_BG, activeforeground=FG_COLOR,
                       font=FONT_MAIN, command=self._on_schedule_toggle) \
            .grid(row=1, column=0, columnspan=2, sticky='w', padx=15, pady=5)

        # 上午时间
        tk.Label(schedule_card, text="上午时间：",
                 font=FONT_MAIN, fg=FG_COLOR, bg=TEXT_BG) \
            .grid(row=2, column=0, sticky='w', padx=15, pady=5)
        self.morning_var = tk.StringVar(value="12:05")
        ttk.Entry(schedule_card, textvariable=self.morning_var, width=12) \
            .grid(row=2, column=1, sticky='w', padx=(0, 15), pady=5)

        # 下午时间
        tk.Label(schedule_card, text="下午时间：",
                 font=FONT_MAIN, fg=FG_COLOR, bg=TEXT_BG) \
            .grid(row=3, column=0, sticky='w', padx=15, pady=5)
        self.afternoon_var = tk.StringVar(value="15:05")
        ttk.Entry(schedule_card, textvariable=self.afternoon_var, width=12) \
            .grid(row=3, column=1, sticky='w', padx=(0, 15), pady=5)

        # 下次执行时间
        self.next_run_label = tk.Label(schedule_card, text="",
                                       font=FONT_MAIN, fg=WARN_COLOR, bg=TEXT_BG)
        self.next_run_label.grid(row=4, column=0, columnspan=2,
                                 sticky='w', padx=15, pady=(5, 12))

        # ----- 操作按钮区 -----
        btn_card = tk.Frame(left_frame, bg=TEXT_BG, relief=tk.FLAT)
        btn_card.pack(fill=tk.X)

        tk.Label(btn_card, text="🖱️ 手动操作",
                 font=("微软雅黑", 11, "bold"), fg=FG_COLOR,
                 bg=TEXT_BG).pack(anchor='w', padx=15, pady=(12, 8))

        # 按钮网格
        btn_frame = tk.Frame(btn_card, bg=TEXT_BG)
        btn_frame.pack(padx=15, pady=(0, 12))

        self.start_btn = tk.Button(btn_frame, text="▶ 执行下载任务",
                                   command=self._manual_execute,
                                   bg=SUCCESS_COLOR, fg="#ffffff",
                                   font=FONT_BTN, relief=tk.FLAT,
                                   cursor='hand2', padx=20, pady=8)
        self.start_btn.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))

        self.stop_btn = tk.Button(btn_frame, text="■ 停止任务",
                                  command=self._stop_task,
                                  bg=ERROR_COLOR, fg="#ffffff",
                                  font=FONT_BTN, relief=tk.FLAT,
                                  state='disabled',
                                  cursor='hand2', padx=20, pady=8)
        self.stop_btn.grid(row=1, column=0, columnspan=2, sticky='ew')

        btn_frame.columnconfigure(0, weight=1)

        # 交互确认提示
        self.confirm_label = tk.Label(btn_card, text="",
                                      font=FONT_MAIN, fg=WARN_COLOR, bg=TEXT_BG,
                                      wraplength=280)
        self.confirm_label.pack(pady=(5, 12))

        # ========== 右侧日志区 ==========
        right_frame = tk.Frame(main_frame, bg=BG_COLOR)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(15, 0))
        right_frame.configure(width=300)

        log_label = tk.Label(right_frame, text="📋 操作日志",
                             font=("微软雅黑", 11, "bold"), fg=FG_COLOR, bg=BG_COLOR)
        log_label.pack(anchor='w', pady=(0, 8))

        # 日志文本框
        log_frame = tk.Frame(right_frame, bg=TEXT_BG)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, bg=TEXT_BG, fg=FG_COLOR,
                                 font=("Consolas", 9), relief=tk.FLAT,
                                 state='disabled', wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 配置日志标签
        self.log_text.tag_config("INFO", foreground="#eaeaea")
        self.log_text.tag_config("WARNING", foreground="#ffab00")
        self.log_text.tag_config("ERROR", foreground="#ff5252")
        self.log_text.tag_config("SUCCESS", foreground="#00c853")

        # 日志滚动条
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 清空日志按钮
        tk.Button(right_frame, text="清空日志",
                  command=self._clear_log,
                  bg=ACCENT_COLOR, fg=FG_COLOR,
                  font=FONT_MAIN, relief=tk.FLAT,
                  cursor='hand2').pack(pady=(8, 0))

    def _bind_events(self):
        """绑定热键事件"""
        self.root.bind('<F5>', lambda e: self._manual_execute())
        self.root.bind('<Escape>', lambda e: self._stop_task())

    def _load_config(self):
        """加载配置"""
        self.path_var.set(self.config.get("paths", "tongdaxin_path", ""))
        self.phone_var.set(self.config.get("notify", "phone_number", ""))
        self.morning_var.set(self.config.get("schedule", "morning_time", "12:05"))
        self.afternoon_var.set(self.config.get("schedule", "afternoon_time", "15:05"))
        self.enable_schedule_var.set(
            self.config.get("schedule", "enabled", "true").lower() == "true"
        )

    def _save_config(self):
        """保存配置"""
        path = self.path_var.get().strip()
        phone = self.phone_var.get().strip()
        morning = self.morning_var.get().strip()
        afternoon = self.afternoon_var.get().strip()
        enabled = self.enable_schedule_var.get()

        # 验证路径
        if not path:
            messagebox.showwarning("配置不完整", "请填写通达信安装路径！")
            return

        # 去除末尾的 \
        path = path.rstrip('\\')

        # 检查通达信主程序是否存在
        exe_path = os.path.join(path, "TongDaXin.exe")
        if not os.path.exists(exe_path):
            result = messagebox.askyesno("路径检查",
                f"未找到 TongDaXin.exe\n\n路径：{exe_path}\n\n"
                "请确认路径是否正确（应指向通达信安装目录）\n\n"
                "是否继续保存？")
            if not result:
                return

        # 保存配置
        self.config.set("paths", "tongdaxin_path", path)
        self.config.set("notify", "phone_number", phone)
        self.config.set("schedule", "morning_time", morning)
        self.config.set("schedule", "afternoon_time", afternoon)
        self.config.set("schedule", "enabled", str(enabled))
        self.config.save()

        self.logger.info("配置已保存")
        messagebox.showinfo("保存成功", "配置已保存，下一任务将使用新配置！")

    def _on_schedule_toggle(self):
        """定时任务开关切换"""
        enabled = self.enable_schedule_var.get()
        self.scheduler.set_enabled(enabled)
        state = "启用" if enabled else "停用"
        self.logger.info(f"定时任务已{state}")

    def _manual_execute(self):
        """手动执行任务"""
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("配置不完整", "请先填写并保存通达信路径！")
            return

        path = path.rstrip('\\')
        exe_path = os.path.join(path, "TongDaXin.exe")
        if not os.path.exists(exe_path):
            messagebox.showerror("路径错误",
                f"找不到可执行文件：\n{exe_path}\n\n"
                "请检查路径是否正确！")
            return

        # 确认对话框
        result = messagebox.askyesno("确认执行",
            "即将开始执行下载任务：\n\n"
            "1. 启动通达信\n"
            "2. 登录账户\n"
            "3. 下载日线和实时数据\n"
            "4. 导出 Excel\n"
            "5. 发送通知\n\n"
            "是否继续？")
        if not result:
            self.logger.warning("用户取消手动执行")
            return

        self._execute_download(path)

    def _execute_download(self, path):
        """执行下载任务（在新线程中运行）"""
        self.is_running = True
        self._update_button_state(running=True)

        def run():
            try:
                self.logger.info("=" * 40)
                self.logger.info(">>> 开始执行下载任务 <<<")
                self.logger.info("=" * 40)

                # 设置通达信路径
                self.automation.set_tongdaxin_path(path)
                self.notifier.set_phone(self.phone_var.get().strip())

                # 执行自动化流程
                success = self.automation.execute_full_flow()

                # 生成日志报告
                report = self.logger.get_today_report()
                self.logger.info("--- 今日任务报告 ---")
                for line in report:
                    self.logger.info(line)

                if success:
                    self.logger.info("✅ 任务执行成功！")
                    # 发送通知
                    if self.phone_var.get().strip():
                        self.notifier.send_download_complete()
                else:
                    self.logger.error("❌ 任务执行失败，请检查日志")
                    messagebox.showerror("执行失败",
                        "任务未完全成功，请查看操作日志了解详情。")

            except Exception as e:
                self.logger.error(f"执行异常：{str(e)}")
                import traceback
                self.logger.error(traceback.format_exc())
            finally:
                self.is_running = False
                # 在主线程中更新 UI
                self.root.after(0, lambda: self._update_button_state(running=False))

        thread = threading.Thread(target=run, daemon=True)
        thread.start()

    def _stop_task(self):
        """停止当前任务"""
        if self.is_running:
            self.logger.warning("用户请求停止任务...")
            self.automation.stop()
            messagebox.showinfo("已停止", "任务已请求停止，请稍候...")
        else:
            self.logger.info("当前没有正在执行的任务")

    def _update_button_state(self, running):
        """更新按钮状态"""
        if running:
            self.start_btn.configure(state='disabled', bg=ACCENT_COLOR)
            self.stop_btn.configure(state='normal')
            self._set_status("● 运行中", WARN_COLOR)
        else:
            self.start_btn.configure(state='normal', bg=SUCCESS_COLOR)
            self.stop_btn.configure(state='disabled')
            self._set_status("● 服务正常", SUCCESS_COLOR)

    def _set_status(self, text, color):
        """设置状态显示"""
        self.status_label.configure(text=text, fg=color)
        self.status_dot.configure(fg=color)

    def _clear_log(self):
        """清空日志"""
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

    def _start_scheduler(self):
        """启动调度线程"""
        self.scheduler_thread = threading.Thread(
            target=self.scheduler.run,
            daemon=True
        )
        self.scheduler_thread.start()
        self._update_next_run_label()

    def _update_next_run_label(self):
        """更新下次执行时间显示"""
        next_time = self.scheduler.get_next_run_time()
        if next_time:
            self.next_run_label.configure(
                text=f"下次执行：{next_time.strftime('%H:%M')}"
            )
        # 每分钟更新一次
        self.root.after(60000, self._update_next_run_label)

    def _on_closing(self):
        """窗口关闭事件"""
        if self.is_running:
            result = messagebox.askyesno("确认退出",
                "任务正在执行中，是否强制退出？")
            if not result:
                return

        self.logger.info("程序关闭")
        self.root.destroy()


def main():
    """主入口"""
    # Windows 下设置高 DPI 支持
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per monitor v2
    except:
        pass

    root = tk.Tk()
    app = TongDaXinToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
