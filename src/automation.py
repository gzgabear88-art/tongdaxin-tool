#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通达信自动化操作模块
负责控制鼠标点击和键盘输入
"""

import os
import sys
import time
import subprocess
import pyautogui
import logging
from datetime import datetime

# 配置 pyautogui
pyautogui.FAILSAFE = True  # 移动鼠标到角落退出
pyautogui.PAUSE = 0.3       # 每次操作后暂停


class TongDaXinAutomation:
    """通达信自动化操作类"""

    # 常用分辨率的基准坐标（1280x720）
    BASE_WIDTH = 1280
    BASE_HEIGHT = 720

    def __init__(self, logger):
        self.logger = logger
        self.tongdaxin_path = None
        self.process = None
        self.running = False

        # 用于计算分辨率缩放比例
        self.scale_x = 1.0
        self.scale_y = 1.0

    def set_tongdaxin_path(self, path):
        """设置通达信安装路径"""
        self.tongdaxin_path = path
        self._calculate_scale()

    def _calculate_scale(self):
        """计算屏幕分辨率缩放比例"""
        try:
            screen_w, screen_h = pyautogui.size()
            self.scale_x = screen_w / self.BASE_WIDTH
            self.scale_y = screen_h / self.BASE_HEIGHT
            self.logger.info(
                f"屏幕分辨率：{screen_w}x{screen_h}，缩放比例：{self.scale_x:.2f}x{self.scale_y:.2f}"
            )
        except Exception as e:
            self.logger.warning(f"无法获取屏幕分辨率，使用默认值：{e}")
            self.scale_x = 1.0
            self.scale_y = 1.0

    def _scale_coords(self, x, y):
        """根据当前分辨率缩放坐标"""
        return int(x * self.scale_x), int(y * self.scale_y)

    def stop(self):
        """请求停止当前任务"""
        self.running = False

    def execute_full_flow(self) -> bool:
        """
        执行完整下载流程
        返回：是否完全成功
        """
        self.running = True
        start_time = time.time()

        try:
            # ========== 步骤 1：启动通达信 ==========
            self.logger.info("[步骤1/7] 启动通达信金融终端...")
            if not self._launch_tongdaxin():
                return False

            # ========== 步骤 2：登录 ==========
            self.logger.info("[步骤2/7] 等待登录界面...")
            if not self._wait_for_login():
                return False

            self.logger.info("[步骤2/7] 执行登录操作...")
            if not self._do_login():
                return False

            # ========== 步骤 3：输入 51 并回车 ==========
            self.logger.info("[步骤3/7] 输入命令：51")
            if not self._input_command_51():
                return False

            # ========== 步骤 4：选择菜单栏选项按钮 ==========
            self.logger.info("[步骤4/7] 选择菜单栏选项...")
            if not self._click_option_button():
                return False

            # ========== 步骤 5：点击盘后数据下载 ==========
            self.logger.info("[步骤5/7] 点击盘后数据下载...")
            if not self._download_data():
                return False

            # ========== 步骤 6：输入 34 并导出 Excel ==========
            self.logger.info("[步骤6/7] 输入命令：34，导出数据...")
            if not self._input_command_34():
                return False

            # ========== 步骤 7：退出系统 ==========
            self.logger.info("[步骤7/7] 关闭通达信...")
            self._close_tongdaxin()

            elapsed = time.time() - start_time
            self.logger.info(f"✅ 全部流程执行完成，耗时：{elapsed:.1f} 秒")

            # 记录成功
            self.logger.log_operation(
                operation="完整下载流程",
                status="成功",
                detail=f"耗时 {elapsed:.1f} 秒"
            )
            return True

        except Exception as e:
            self.logger.error(f"执行异常：{str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            self.logger.log_operation(
                operation="完整下载流程",
                status="失败",
                detail=str(e)
            )
            return False

    def _launch_tongdaxin(self) -> bool:
        """启动通达信"""
        try:
            exe_path = os.path.join(self.tongdaxin_path, "TongDaXin.exe")
            if not os.path.exists(exe_path):
                self.logger.error(f"找不到通达信：{exe_path}")
                return False

            self.logger.info(f"正在启动：{exe_path}")
            self.process = subprocess.Popen(
                [exe_path],
                cwd=self.tongdaxin_path,
                shell=False
            )

            # 等待程序启动
            time.sleep(3)
            self.logger.info("通达信已启动，等待界面加载...")
            return True

        except Exception as e:
            self.logger.error(f"启动失败：{e}")
            return False

    def _wait_for_login(self, timeout=30) -> bool:
        """等待登录界面出现"""
        self.logger.info("等待登录界面...")
        # 等待一段时间让程序完全加载
        for i in range(timeout // 2):
            if not self.running:
                return False
            time.sleep(2)
            self.logger.info(f"  等待中... ({i*2}s/{timeout}s)")
            # 这里可以加入图片识别来检测登录框
            # 但简单等待通常足够
        return True

    def _do_login(self) -> bool:
        """执行登录操作"""
        try:
            # 查找登录按钮并点击
            # 注意：具体坐标需要根据实际界面调整
            # 这里使用相对坐标（基于1280x720）

            # 点击登录按钮（需要根据实际位置调整）
            login_x, login_y = self._scale_coords(640, 450)
            self.logger.info(f"点击登录按钮，坐标：({login_x}, {login_y})")
            pyautogui.click(login_x, login_y)

            time.sleep(2)
            return True

        except Exception as e:
            self.logger.error(f"登录失败：{e}")
            return False

    def _input_command_51(self) -> bool:
        """输入 51 并按回车"""
        try:
            time.sleep(1)
            self.logger.info("键盘输入：51")
            pyautogui.write("51", interval=0.1)
            time.sleep(0.3)
            self.logger.info("按回车")
            pyautogui.press("enter")

            time.sleep(2)
            return True

        except Exception as e:
            self.logger.error(f"输入命令51失败：{e}")
            return False

    def _click_option_button(self) -> bool:
        """点击菜单栏的选项按钮"""
        try:
            # 选项按钮位置（需要根据实际界面调整）
            # 通常在菜单栏上
            option_x, option_y = self._scale_coords(300, 40)
            self.logger.info(f"点击选项按钮，坐标：({option_x}, {option_y})")
            pyautogui.click(option_x, option_y)

            time.sleep(1)
            return True

        except Exception as e:
            self.logger.error(f"点击选项按钮失败：{e}")
            return False

    def _download_data(self) -> bool:
        """执行盘后数据下载"""
        try:
            # 点击盘后数据下载菜单项
            menu_x, menu_y = self._scale_coords(350, 80)
            self.logger.info(f"点击盘后数据下载，坐标：({menu_x}, {menu_y})")
            pyautogui.click(menu_x, menu_y)

            time.sleep(2)

            # 勾选日线数据
            daily_x, daily_y = self._scale_coords(200, 150)
            self.logger.info(f"勾选日线数据，坐标：({daily_x}, {daily_y})")
            pyautogui.click(daily_x, daily_y)

            time.sleep(0.5)

            # 勾选实时行情数据
            realtime_x, realtime_y = self._scale_coords(200, 180)
            self.logger.info(f"勾选实时行情，坐标：({realtime_x}, {realtime_y})")
            pyautogui.click(realtime_x, realtime_y)

            time.sleep(0.5)

            # 点击开始下载
            start_x, start_y = self._scale_coords(400, 350)
            self.logger.info(f"点击开始下载，坐标：({start_x}, {start_y})")
            pyautogui.click(start_x, start_y)

            # 等待下载完成
            self.logger.info("等待下载完成（最多60秒）...")
            for i in range(60):
                if not self.running:
                    return False
                time.sleep(1)
                if i % 10 == 0:
                    self.logger.info(f"  下载中... ({i}s)")

            # 点击关闭对话框
            close_x, close_y = self._scale_coords(500, 200)
            self.logger.info(f"关闭对话框，坐标：({close_x}, {close_y})")
            pyautogui.click(close_x, close_y)

            time.sleep(1)
            return True

        except Exception as e:
            self.logger.error(f"下载数据失败：{e}")
            return False

    def _input_command_34(self) -> bool:
        """输入 34 并导出 Excel"""
        try:
            time.sleep(1)

            # 输入 34
            self.logger.info("键盘输入：34")
            pyautogui.write("34", interval=0.1)
            time.sleep(0.3)
            pyautogui.press("enter")

            time.sleep(2)

            # 点击文件菜单（或使用快捷键）
            file_x, file_y = self._scale_coords(150, 40)
            self.logger.info(f"点击文件菜单，坐标：({file_x}, {file_y})")
            pyautogui.click(file_x, file_y)

            time.sleep(1)

            # 选择导出选项
            export_x, export_y = self._scale_coords(170, 80)
            self.logger.info(f"点击导出，坐标：({export_x}, {export_y})")
            pyautogui.click(export_x, export_y)

            time.sleep(2)

            # 选择 Excel
            excel_x, excel_y = self._scale_coords(300, 200)
            self.logger.info(f"选择Excel文件，坐标：({excel_x}, {excel_y})")
            pyautogui.click(excel_x, excel_y)

            time.sleep(1)

            # 选择所有数据
            all_x, all_y = self._scale_coords(300, 250)
            self.logger.info(f"选择所有数据，坐标：({all_x}, {all_y})")
            pyautogui.click(all_x, all_y)

            time.sleep(1)

            # 点击导出按钮
            do_export_x, do_export_y = self._scale_coords(400, 350)
            self.logger.info(f"执行导出，坐标：({do_export_x}, {do_export_y})")
            pyautogui.click(do_export_x, do_export_y)

            # 等待导出完成
            self.logger.info("等待导出完成...")
            time.sleep(5)

            return True

        except Exception as e:
            self.logger.error(f"导出Excel失败：{e}")
            return False

    def _close_tongdaxin(self):
        """关闭通达信"""
        try:
            if self.process:
                self.process.terminate()
                self.process.wait(timeout=10)
                self.logger.info("通达信已关闭")
            else:
                # 使用快捷键关闭
                pyautogui.hotkey('alt', 'f4')
                time.sleep(1)
                pyautogui.press('enter')
                self.logger.info("通达信已通过快捷键关闭")

        except Exception as e:
            self.logger.warning(f"关闭通达信时出现异常：{e}")


# ============================================================
# 辅助函数
# ============================================================

def find_window_by_title(title: str) -> bool:
    """查找指定标题的窗口"""
    try:
        import win32gui
        import win32con

        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if title in window_title:
                    windows.append(hwnd)
            return True

        windows = []
        win32gui.EnumWindows(callback, windows)
        return len(windows) > 0

    except ImportError:
        return False


def activate_window(title: str) -> bool:
    """激活指定标题的窗口"""
    try:
        import win32gui
        import win32con

        hwnd = win32gui.FindWindow(None, title)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            return True
        return False

    except ImportError:
        return False
