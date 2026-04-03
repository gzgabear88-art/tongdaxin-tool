#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化操作引擎 - Auto Operation Engine

本模块是通达信自动下载工具的核心，负责：
1. 控制鼠标和键盘操作
2. 图像识别定位按钮位置
3. 执行自动化流程的各个步骤

技术方案：
- 使用 PyAutoGUI 进行鼠标键盘控制
- 使用 OpenCV 进行图像识别
- 每步操作前先验证目标存在，然后再执行

版本: v1.0
作者: 小张
日期: 2026-04-04
"""

import os
import sys
import time
import subprocess
import logging
import threading
from typing import Optional, Callable, Dict, Any

# 尝试导入可选依赖
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    import cv2
    import numpy as np
    OPENCV_AVAILABLE = True
except ImportError:
    OPENCV_AVAILABLE = False

logger = logging.getLogger(__name__)


class AutoOperation:
    """
    自动化操作引擎类
    
    负责执行自动化操作的各个步骤，包括：
    - 启动通达信软件
    - 登录账号
    - 导航操作
    - 数据下载和导出
    
    Attributes:
        parent: 父窗口引用，用于更新日志等
        is_stopped: 停止标志，设置为True时终止当前任务
    """

    def __init__(self, parent=None):
        """
        初始化自动化操作引擎
        
        Args:
            parent: 父窗口引用，用于回调更新界面
        """
        self.parent = parent
        self.is_stopped = False
        self.current_step = 0
        
        # PyAutoGUI 安全设置
        if PYAUTOGUI_AVAILABLE:
            # 设置安全区域，鼠标移到角落不会失控
            pyautogui.FAILSAFE = True
            # 每个操作之间的暂停时间
            pyautogui.PAUSE = 0.5

    def stop(self):
        """发送停止信号，终止当前正在执行的任务"""
        self.is_stopped = True
        logger.info("收到停止信号")

    def _check_stop(self):
        """
        检查是否收到停止信号
        
        Returns:
            bool: 如果收到停止信号返回True
        """
        if self.is_stopped:
            logger.info("任务已被用户终止")
            return True
        return False

    def _log(self, message, callback=None):
        """
        记录日志消息
        
        Args:
            message: 日志消息
            callback: 可选的回调函数，用于更新UI
        """
        logger.info(message)
        if callback and callable(callback):
            callback(message)
        if self.parent and hasattr(self.parent, 'log_message'):
            self.parent.log_message(message)

    def _wait_and_check(self, seconds=1):
        """
        等待一段时间，同时检查是否收到停止信号
        
        Args:
            seconds: 等待秒数
            
        Returns:
            bool: 如果收到停止信号返回True
        """
        for _ in range(int(seconds * 10)):
            if self._check_stop():
                return True
            time.sleep(0.1)
        return False

    # ==================== 核心自动化步骤 ====================

    def find_element_on_screen(self, image_path, confidence=0.8, timeout=10):
        """
        在屏幕上查找图像元素
        
        使用图像识别技术在屏幕上查找指定的图像区域。
        这是实现跨版本兼容性的关键，通过图像匹配而非固定坐标来定位按钮。
        
        Args:
            image_path: 要查找的图像文件路径（相对于images目录）
            confidence: 匹配置信度，0-1之间，越高越严格
            timeout: 超时时间（秒）
            
        Returns:
            tuple: (x, y) 匹配到的位置坐标，未找到返回None
        """
        if not OPENCV_AVAILABLE:
            logger.warning("OpenCV未安装，无法使用图像识别功能")
            return None
        
        # 构建完整的图像路径
        base_dir = os.path.dirname(os.path.abspath(__file__))
        full_image_path = os.path.join(base_dir, "images", image_path)
        
        if not os.path.exists(full_image_path):
            logger.error(f"图像文件不存在: {full_image_path}")
            return None
        
        try:
            # 读取目标图像
            target = cv2.imread(full_image_path)
            if target is None:
                logger.error(f"无法读取图像: {full_image_path}")
                return None
            
            # 截取屏幕
            screenshot = pyautogui.screenshot()
            screen = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # 模板匹配
            result = cv2.matchTemplate(screen, target, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= confidence:
                # 计算图像中心点
                h, w = target.shape[:2]
                center_x = max_loc[0] + w // 2
                center_y = max_loc[1] + h // 2
                logger.info(f"找到图像 {image_path}, 位置: ({center_x}, {center_y}), 置信度: {max_val:.2f}")
                return (center_x, center_y)
            else:
                logger.warning(f"未找到图像 {image_path}, 最高置信度: {max_val:.2f}")
                return None
                
        except Exception as e:
            logger.exception(f"图像识别出错: {e}")
            return None

    def find_text_on_screen(self, text, region=None, timeout=10):
        """
        在屏幕上查找文字
        
        使用OCR技术（或简化的颜色匹配）在屏幕上查找文字。
        
        Args:
            text: 要查找的文字
            region: 可选的屏幕区域 (left, top, width, height)
            timeout: 超时时间
            
        Returns:
            tuple: (x, y) 文字位置坐标，未找到返回None
        """
        # TODO: 实现文字识别功能
        # 目前返回None，实际使用时需要配合图像识别
        logger.warning(f"文字识别功能暂未实现，请使用图像识别功能")
        return None

    def click_element(self, x, y, button='left', clicks=1, interval=0.1):
        """
        点击指定坐标位置
        
        Args:
            x: X坐标
            y: Y坐标
            button: 鼠标按钮 ('left', 'right', 'middle')
            clicks: 点击次数
            interval: 点击间隔时间
        """
        if self._check_stop():
            return False
            
        pyautogui.click(x, y, clicks=clicks, button=button, interval=interval)
        self._log(f"点击位置: ({x}, {y})")
        return True

    def type_text(self, text, interval=0.05):
        """
        输入文本
        
        Args:
            text: 要输入的文本
            interval: 每个字符之间的间隔
        """
        if self._check_stop():
            return False
            
        pyautogui.write(text, interval=interval)
        self._log(f"输入文本: {text}")
        return True

    def press_key(self, key):
        """
        按下键盘按键
        
        Args:
            key: 按键名称（如 'enter', 'esc', 'tab'等）
        """
        if self._check_stop():
            return False
            
        pyautogui.press(key)
        self._log(f"按下按键: {key}")
        return True

    def move_to(self, x, y, duration=0.5):
        """
        移动鼠标到指定位置
        
        Args:
            x: 目标X坐标
            y: 目标Y坐标
            duration: 移动时间（秒）
        """
        if self._check_stop():
            return False
            
        pyautogui.moveTo(x, y, duration=duration)
        return True

    # ==================== 自动化流程步骤 ====================

    def step_launch_app(self, app_path, log_callback=None):
        """
        步骤1: 启动通达信软件
        
        Args:
            app_path: 通达信软件路径
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 1
        self._log("=" * 50, log_callback)
        self._log("步骤1: 启动通达信软件...", log_callback)
        
        if self._wait_and_check(1):
            return {'success': False, 'error': '用户终止'}
        
        if not os.path.exists(app_path):
            return {'success': False, 'error': f'软件路径不存在: {app_path}'}
        
        try:
            self._log(f"正在启动: {app_path}", log_callback)
            subprocess.Popen(f'"{app_path}"', shell=True)
            self._log("软件已启动，等待加载...", log_callback)
            
            # 等待软件加载
            time.sleep(5)
            
            if self._wait_and_check(5):
                return {'success': False, 'error': '用户终止'}
            
            self._log("软件启动完成", log_callback)
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': f'启动失败: {str(e)}'}

    def step_login(self, username, password, log_callback=None):
        """
        步骤2: 自动登录
        
        Args:
            username: 账号
            password: 密码
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 2
        self._log("=" * 50, log_callback)
        self._log("步骤2: 自动登录...", log_callback)
        
        if self._check_stop():
            return {'success': False, 'error': '用户终止'}
        
        try:
            # 查找并点击账号输入框
            account_pos = self.find_element_on_screen("account_input.png")
            if account_pos:
                self.click_element(*account_pos)
                self._wait_and_check(0.5)
            
            # 输入账号
            self.type_text(username)
            self._wait_and_check(0.3)
            
            # Tab切换到密码框
            self.press_key('tab')
            self._wait_and_check(0.3)
            
            # 输入密码
            self.type_text(password)
            self._wait_and_check(0.3)
            
            # 点击登录按钮
            login_btn = self.find_element_on_screen("login_button.png")
            if login_btn:
                self.click_element(*login_btn)
            else:
                # 如果找不到登录按钮，尝试按回车
                self._log("未找到登录按钮，尝试按回车", log_callback)
                self.press_key('enter')
            
            self._log("登录信息已提交，等待登录结果...", log_callback)
            
            # 等待登录完成
            time.sleep(3)
            
            if self._wait_and_check(3):
                return {'success': False, 'error': '用户终止'}
            
            self._log("登录步骤完成", log_callback)
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': f'登录失败: {str(e)}'}

    def step_input_code(self, code, log_callback=None):
        """
        步骤3: 输入板块代码
        
        Args:
            code: 板块代码（默认51）
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 3
        self._log("=" * 50, log_callback)
        self._log(f"步骤3: 输入板块代码 {code}...", log_callback)
        
        if self._check_stop():
            return {'success': False, 'error': '用户终止'}
        
        try:
            # 输入代码
            self.type_text(code)
            self._wait_and_check(0.3)
            
            # 按回车确认
            self.press_key('enter')
            self._log("已输入代码并确认", log_callback)
            
            # 等待界面响应
            time.sleep(2)
            
            if self._wait_and_check(2):
                return {'success': False, 'error': '用户终止'}
            
            self._log("代码输入完成", log_callback)
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': f'输入代码失败: {str(e)}'}

    def step_open_menu(self, menu_name, log_callback=None):
        """
        步骤4: 打开菜单
        
        Args:
            menu_name: 菜单名称（如"选项"）
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 4
        self._log("=" * 50, log_callback)
        self._log(f"步骤4: 打开{menu_name}菜单...", log_callback)
        
        if self._check_stop():
            return {'success': False, 'error': '用户终止'}
        
        try:
            # 查找菜单项
            menu_pos = self.find_element_on_screen(f"menu_{menu_name}.png")
            if menu_pos:
                self.click_element(*menu_pos)
            else:
                # 尝试使用快捷键 Alt+O 打开选项菜单
                self._log("未找到菜单图像，尝试快捷键", log_callback)
                pyautogui.keyDown('alt')
                pyautogui.press('o')
                pyautogui.keyUp('alt')
            
            self._wait_and_check(1)
            self._log(f"{menu_name}菜单已打开", log_callback)
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': f'打开菜单失败: {str(e)}'}

    def step_select_download(self, log_callback=None):
        """
        步骤5: 选择盘后数据下载并下载
        
        Args:
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 5
        self._log("=" * 50, log_callback)
        self._log("步骤5: 盘后数据下载...", log_callback)
        
        if self._check_stop():
            return {'success': False, 'error': '用户终止'}
        
        try:
            # 选择盘后数据下载选项
            download_option = self.find_element_on_screen("menu_download.png")
            if download_option:
                self.click_element(*download_option)
            else:
                return {'success': False, 'error': '未找到盘后数据下载选项'}
            
            self._wait_and_check(1)
            
            # 勾选日线
            daily_line = self.find_element_on_screen("checkbox_daily.png")
            if daily_line:
                self.click_element(*daily_line)
            self._wait_and_check(0.3)
            
            # 勾选实时行情
            realtime = self.find_element_on_screen("checkbox_realtime.png")
            if realtime:
                self.click_element(*realtime)
            self._wait_and_check(0.3)
            
            # 点击开始下载
            start_btn = self.find_element_on_screen("button_start_download.png")
            if start_btn:
                self.click_element(*start_btn)
            else:
                return {'success': False, 'error': '未找到开始下载按钮'}
            
            self._log("已开始下载，等待完成...", log_callback)
            
            # 等待下载完成（这里需要根据实际情况调整）
            time.sleep(5)
            
            if self._wait_and_check(5):
                return {'success': False, 'error': '用户终止'}
            
            self._log("数据下载完成", log_callback)
            return {'success': True}
            
        except Exception as e:
            return {'success': False, 'error': f'下载失败: {str(e)}'}

    def step_export_data(self, save_path, log_callback=None):
        """
        步骤6: 导出数据到Excel
        
        Args:
            save_path: 存储路径
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success和可选的error信息
        """
        self.current_step = 6
        self._log("=" * 50, log_callback)
        self._log("步骤6: 导出数据...", log_callback)
        
        if self._check_stop():
            return {'success': False, 'error': '用户终止'}
        
        try:
            # 选择数据导出选项
            export_option = self.find_element_on_screen("menu_export.png")
            if export_option:
                self.click_element(*export_option)
            else:
                return {'success': False, 'error': '未找到数据导出选项'}
            
            self._wait_and_check(1)
            
            # 选择Excel格式
            excel_format = self.find_element_on_screen("format_excel.png")
            if excel_format:
                self.click_element(*excel_format)
            self._wait_and_check(0.3)
            
            # 选择所有数据列
            all_columns = self.find_element_on_screen("option_all_columns.png")
            if all_columns:
                self.click_element(*all_columns)
            self._wait_and_check(0.3)
            
            # 设置保存路径
            path_input = self.find_element_on_screen("input_save_path.png")
            if path_input:
                self.click_element(*path_input)
                self.type_text(save_path)
            self._wait_and_check(0.3)
            
            # 点击导出
            export_btn = self.find_element_on_screen("button_export.png")
            if export_btn:
                self.click_element(*export_btn)
            else:
                return {'success': False, 'error': '未找到导出按钮'}
            
            self._log("数据导出中...", log_callback)
            
            # 等待导出完成
            time.sleep(3)
            
            if self._wait_and_check(3):
                return {'success': False, 'error': '用户终止'}
            
            self._log(f"数据已导出到: {save_path}", log_callback)
            return {'success': True, 'save_path': save_path}
            
        except Exception as e:
            return {'success': False, 'error': f'导出失败: {str(e)}'}

    def run_full_flow(self, username, password, tdx_path, save_path, log_callback=None):
        """
        执行完整的自动化流程
        
        按照顺序执行所有步骤，实现从启动到导出的完整流程。
        
        Args:
            username: 账号
            password: 密码
            tdx_path: 通达信软件路径
            save_path: 数据存储路径
            log_callback: 日志回调函数
            
        Returns:
            dict: 包含success、save_path等信息
        """
        execute_time = time.strftime("%Y-%m-%d %H:%M:%S")
        
        # 步骤1: 启动软件
        result = self.step_launch_app(tdx_path, log_callback)
        if not result['success']:
            return {**result, 'step': '启动软件', 'execute_time': execute_time}
        
        # 步骤2: 登录
        result = self.step_login(username, password, log_callback)
        if not result['success']:
            return {**result, 'step': '登录', 'execute_time': execute_time}
        
        # 步骤3: 输入代码
        result = self.step_input_code('51', log_callback)
        if not result['success']:
            return {**result, 'step': '输入代码', 'execute_time': execute_time}
        
        # 步骤4: 打开选项菜单
        result = self.step_open_menu('选项', log_callback)
        if not result['success']:
            return {**result, 'step': '打开菜单', 'execute_time': execute_time}
        
        # 步骤5: 下载数据
        result = self.step_select_download(log_callback)
        if not result['success']:
            return {**result, 'step': '下载数据', 'execute_time': execute_time}
        
        # 步骤6: 导出数据
        result = self.step_export_data(save_path, log_callback)
        return {**result, 'step': '导出数据', 'execute_time': execute_time}
