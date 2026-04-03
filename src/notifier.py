#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
短信通知模块
支持多种短信发送方式
"""

import logging
import urllib.request
import urllib.parse
import json
import hashlib
import time


class NotificationService:
    """通知服务类"""

    def __init__(self, config):
        self.config = config
        self.phone = ""
        self.logger = logging.getLogger(__name__)

    def set_phone(self, phone: str):
        """设置通知手机号"""
        self.phone = phone

    def send_download_complete(self) -> bool:
        """发送下载完成通知"""
        if not self.phone:
            self.logger.warning("未配置手机号，跳过通知")
            return False

        try:
            return self._send_sms(
                self.phone,
                "✅ 通达信数据下载已完成！\n"
                "时间：{time}\n"
                "日线和实时行情数据已成功下载并导出。".format(
                    time=time.strftime("%Y-%m-%d %H:%M:%S")
                )
            )
        except Exception as e:
            self.logger.error(f"发送通知失败：{e}")
            return False

    def send_error(self, error_msg: str) -> bool:
        """发送错误通知"""
        if not self.phone:
            return False

        try:
            return self._send_sms(
                self.phone,
                f"⚠️ 通达信下载任务异常\n"
                f"时间：{time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"原因：{error_msg}"
            )
        except Exception as e:
            self.logger.error(f"发送错误通知失败：{e}")
            return False

    def _send_sms(self, phone: str, message: str) -> bool:
        """
        发送短信
        注意：这里使用模拟实现，实际使用需要接入真实的短信网关
        常见方案：
        1. 阿里云短信服务
        2. 腾讯云短信服务
        3. Twilio
        4. 企业短信网关
        """
        self.logger.info(f"准备发送短信到 {phone}")
        self.logger.info(f"短信内容：{message}")

        # 方案1：阿里云短信（需要配置 AccessKey）
        # return self._send_aliyun_sms(phone, message)

        # 方案2：企业微信/钉钉 webhook（免费）
        # return self._send_wechat_webhook(message)

        # 方案3：打印到日志（演示用）
        self.logger.info(f"【短信模拟】To: {phone}")
        self.logger.info(f"【短信模拟】Content: {message}")
        return True

    def _send_aliyun_sms(self, phone: str, message: str) -> bool:
        """
        通过阿里云短信服务发送
        需要安装 aliyun-python-sdk-core 并配置 AccessKey
        """
        try:
            from aliyunsdkcore.client import AcsClient
            from aliyunsdkdysmsapi.request.v20170525 import SendSmsRequest

            access_key_id = self.config.get("notify", "aliyun_key_id", "")
            access_key_secret = self.config.get("notify", "aliyun_key_secret", "")
            sign_name = self.config.get("notify", "aliyun_sign", "通达信助手")
            template_code = self.config.get("notify", "aliyun_template", "")

            if not all([access_key_id, access_key_secret, template_code]):
                self.logger.warning("阿里云短信配置不完整")
                return False

            client = AcsClient(access_key_id, access_key_secret, 'default')

            request = SendSmsRequest.SendSmsRequest()
            request.set_accept_format('json')
            request.set_PhoneNumbers(phone)
            request.set_SignName(sign_name)
            request.set_TemplateCode(template_code)
            request.set_TemplateParam(json.dumps({'message': message}))

            response = client.do_action_with_exception(request)
            result = json.loads(response)

            if result.get('Code') == 'OK':
                self.logger.info("阿里云短信发送成功")
                return True
            else:
                self.logger.error(f"阿里云短信发送失败：{result}")
                return False

        except Exception as e:
            self.logger.error(f"阿里云短信异常：{e}")
            return False

    def _send_wechat_webhook(self, message: str) -> bool:
        """
        通过企业微信群机器人发送
        需要配置 webhook 地址
        """
        try:
            webhook_url = self.config.get("notify", "wechat_webhook", "")
            if not webhook_url:
                return False

            data = {
                "msgtype": "text",
                "text": {
                    "content": f"📈 通达信下载通知\n{message}",
                    "mentioned_mobile_list": [self.phone] if self.phone else []
                }
            }

            req = urllib.request.Request(
                webhook_url,
                data=json.dumps(data).encode('utf-8'),
                headers={'Content-Type': 'application/json'}
            )

            with urllib.request.urlopen(req, timeout=10) as response:
                result = json.loads(response.read().decode('utf-8'))

                if result.get('errcode') == 0:
                    self.logger.info("企业微信通知发送成功")
                    return True
                else:
                    self.logger.error(f"企业微信通知失败：{result}")
                    return False

        except Exception as e:
            self.logger.error(f"企业微信通知异常：{e}")
            return False
