# -*- coding: utf-8 -*-
"""
微信消息自动监听与转发工具
=====================================
【实用版本 - 基于坐标点击和剪贴板】

说明：
微信PC客户端使用自定义UI渲染技术，传统的UI自动化工具（如uiautomation）
无法直接获取控件。本脚本使用更实用的方法：
1. 窗口控制 - 操作微信窗口
2. 坐标点击 - 通过屏幕坐标模拟鼠标操作
3. 剪贴板获取 - 选中文本后复制到剪贴板
4. 快捷键操作 - 模拟键盘操作

使用前配置：
1. 修改邮箱配置
2. 配置微信窗口的坐标（运行坐标获取工具）
3. 登录微信PC客户端
"""

import logging
import smtplib
import time
import sys
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

# ==============================================
# 【必须修改】邮箱配置
# ==============================================
EMAIL_CONFIG = {
    # 发件人邮箱
    'sender_email': '3138969266@qq.com',
    # 邮箱授权码（不是登录密码！）
    'sender_password': 'ovyktydfyoivdcci',
    # 收件人邮箱
    'receiver_email': '3138969266@qq.com',
    # SMTP服务器
    'smtp_server': 'smtp.qq.com',
    # SMTP端口
    'smtp_port': 465,
    # 是否使用SSL
    'use_ssl': True,
}

# ==============================================
# 【重要】微信窗口坐标配置
# ==============================================
# 使用方法：
# 1. 运行 python wechat_click.py --get-coords 获取坐标
# 2. 根据输出修改下面的配置
# 3. 或者使用默认配置（假设微信窗口在屏幕左上角）

WECHAT_COORDS = {
    # ==============================
    # 方式一：相对于微信窗口的坐标（推荐）
    # 微信窗口大小：宽 x 高
    # 你需要根据实际情况调整这些坐标
    # ==============================
    
    # 微信窗口的标准大小（根据你的实际窗口调整）
    'window_width': 1000,    # 窗口宽度
    'window_height': 700,    # 窗口高度
    
    # 左侧会话列表区域（相对于窗口左上角）
    'session_list': {
        'left': 0,           # 左边界
        'top': 50,           # 上边界（排除顶部搜索栏）
        'right': 300,        # 右边界
        'bottom': 650,       # 下边界
        'item_height': 70,   # 每个会话项的高度
    },
    
    # 第一个会话项的点击位置（相对于窗口）
    'first_session': {
        'x': 150,            # X坐标（会话列表中间）
        'y': 80,             # Y坐标（第一个会话）
    },
    
    # 第二个会话项
    'second_session': {
        'x': 150,
        'y': 150,
    },
    
    # 聊天内容区域（用于复制消息）
    'chat_area': {
        'left': 320,         # 左边界（会话列表右侧）
        'top': 50,           # 上边界
        'right': 950,        # 右边界
        'bottom': 600,       # 下边界（输入框上方）
    },
    
    # 聊天标题位置（显示当前聊天的联系人/群聊名称）
    'chat_title': {
        'x': 600,            # 窗口顶部中间
        'y': 25,             # 顶部
    },
    
    # ==============================
    # 方式二：屏幕绝对坐标
    # 如果知道微信窗口在屏幕上的绝对位置，可以使用这些坐标
    # ==============================
    'use_absolute_coords': False,  # 是否使用绝对坐标
    
    # 微信窗口在屏幕上的位置（如果使用绝对坐标）
    'window_screen': {
        'left': 100,
        'top': 100,
        'right': 1100,
        'bottom': 800,
    },
}

# ==============================================
# 程序配置
# ==============================================
APP_CONFIG = {
    # 检查新消息的间隔（秒）
    'check_interval': 5,
    
    # 只处理新消息（不重复发送）
    'only_new_messages': True,
    
    # 运行模式
    # 'normal'   - 正常模式：监听消息
    # 'test'     - 测试模式：测试邮件和坐标
    # 'get_coords' - 获取坐标模式：帮助你获取微信窗口坐标
    'run_mode': 'normal',
    
    # 日志级别
    'log_level': logging.INFO,
    
    # 是否显示详细的调试信息
    'debug': True,
}

# ==============================================
# 导入依赖库
# ==============================================

LIBRARIES = {}

# Windows API 库（必须）
try:
    import win32gui
    import win32con
    import win32api
    import win32clipboard
    LIBRARIES['win32'] = True
except ImportError:
    LIBRARIES['win32'] = False
    print("错误: 未安装 pywin32 库")
    print("请运行: pip install pywin32")
    sys.exit(1)

# 截图和图像识别库（可选，用于检测未读消息）
try:
    from PIL import ImageGrab, Image
    LIBRARIES['pil'] = True
except ImportError:
    LIBRARIES['pil'] = False
    print("警告: 未安装 PIL 库，图像识别功能不可用")
    print("请运行: pip install pillow")

# 鼠标键盘模拟库
try:
    import pyautogui
    LIBRARIES['pyautogui'] = True
except ImportError:
    LIBRARIES['pyautogui'] = False
    print("警告: 未安装 pyautogui 库")
    print("请运行: pip install pyautogui")


# ==============================================
# 日志设置
# ==============================================

def setup_logger():
    logger = logging.getLogger('WeChatClick')
    logger.setLevel(APP_CONFIG['log_level'])
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger

logger = setup_logger()


# ==============================================
# 邮件发送模块
# ==============================================

class EmailSender:
    """邮件发送器"""
    
    def __init__(self, config):
        self.config = config
        self._server = None
        self._connected = False
    
    def test_connection(self):
        """测试邮件连接"""
        logger.info("=" * 60)
        logger.info("测试邮件连接...")
        logger.info(f"  发件人: {self.config['sender_email']}")
        logger.info(f"  收件人: {self.config['receiver_email']}")
        logger.info(f"  服务器: {self.config['smtp_server']}:{self.config['smtp_port']}")
        logger.info("=" * 60)
        
        try:
            if self.config['use_ssl']:
                logger.info("建立SSL连接...")
                self._server = smtplib.SMTP_SSL(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=20
                )
            else:
                logger.info("建立普通连接...")
                self._server = smtplib.SMTP(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=20
                )
                self._server.ehlo()
            
            logger.info("登录邮箱...")
            self._server.login(
                self.config['sender_email'],
                self.config['sender_password']
            )
            
            self._connected = True
            logger.info("✓ 邮件连接测试成功！")
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            logger.error("✗ 认证失败！")
            logger.error("")
            logger.error("可能的原因：")
            logger.error("  1. 邮箱地址或授权码错误")
            logger.error("  2. 未开启SMTP服务")
            logger.error("")
            logger.error("QQ邮箱授权码获取方法：")
            logger.error("  1. 登录QQ邮箱网页版")
            logger.error("  2. 设置 -> 账户 -> 开启POP3/SMTP服务")
            logger.error("  3. 按提示发送短信，获取16位授权码")
            logger.error(f"  错误: {e}")
            return False
            
        except Exception as e:
            logger.error(f"✗ 连接出错: {e}")
            import traceback
            logger.debug(traceback.format_exc())
            return False
    
    def send(self, subject, body, receiver=None):
        """发送邮件"""
        if receiver is None:
            receiver = self.config['receiver_email']
        
        try:
            if not self._connected:
                if not self.test_connection():
                    return False
            
            msg = MIMEMultipart()
            msg['From'] = self.config['sender_email']
            msg['To'] = receiver
            msg['Subject'] = Header(subject, 'utf-8')
            
            text_part = MIMEText(body, 'plain', 'utf-8')
            msg.attach(text_part)
            
            self._server.sendmail(
                self.config['sender_email'],
                receiver.split(','),
                msg.as_string()
            )
            
            logger.info(f"✓ 邮件发送成功: {subject}")
            return True
            
        except Exception as e:
            logger.error(f"✗ 邮件发送失败: {e}")
            self._connected = False
            return False
    
    def disconnect(self):
        """断开连接"""
        if self._server:
            try:
                self._server.quit()
            except:
                pass
            self._server = None
        self._connected = False


# ==============================================
# 剪贴板操作模块
# ==============================================

class Clipboard:
    """剪贴板操作"""
    
    @staticmethod
    def get_text():
        """获取剪贴板文本"""
        try:
            win32clipboard.OpenClipboard()
            text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return text
        except Exception as e:
            logger.debug(f"获取剪贴板失败: {e}")
            return ""
    
    @staticmethod
    def set_text(text):
        """设置剪贴板文本"""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return True
        except Exception as e:
            logger.debug(f"设置剪贴板失败: {e}")
            return False
    
    @staticmethod
    def clear():
        """清空剪贴板"""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except:
            pass


# ==============================================
# 微信窗口控制器
# ==============================================

class WeChatWindow:
    """微信窗口控制器"""
    
    def __init__(self, window_title="微信"):
        self.window_title = window_title
        self.hwnd = None
        self._rect = None  # 窗口位置和大小
    
    def find(self):
        """查找微信窗口"""
        # 精确查找
        self.hwnd = win32gui.FindWindow(None, self.window_title)
        
        if self.hwnd == 0:
            # 模糊查找
            def callback(h, _):
                if win32gui.IsWindowVisible(h):
                    title = win32gui.GetWindowText(h)
                    if '微信' in title:
                        self.hwnd = h
                        self.window_title = title
                return True
            
            win32gui.EnumWindows(callback, None)
        
        if self.hwnd == 0:
            logger.error("未找到微信窗口")
            logger.error("请确保：")
            logger.error("  1. 微信PC客户端已登录")
            logger.error("  2. 微信窗口已打开（不要最小化到托盘）")
            return False
        
        logger.info(f"找到微信窗口: HWND={self.hwnd}, Title='{self.window_title}'")
        self._update_rect()
        return True
    
    def _update_rect(self):
        """更新窗口位置和大小"""
        if self.hwnd:
            self._rect = win32gui.GetWindowRect(self.hwnd)
            # 返回 (left, top, right, bottom)
            return self._rect
        return None
    
    def get_rect(self):
        """获取窗口位置和大小"""
        if self.hwnd:
            self._update_rect()
            return self._rect
        return None
    
    def get_size(self):
        """获取窗口大小 (width, height)"""
        rect = self.get_rect()
        if rect:
            return (rect[2] - rect[0], rect[3] - rect[1])
        return None
    
    def restore(self):
        """恢复窗口（从最小化恢复）"""
        if self.hwnd:
            placement = win32gui.GetWindowPlacement(self.hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
                time.sleep(0.3)
                self._update_rect()
    
    def bring_to_front(self):
        """将窗口置于前台"""
        if not self.hwnd:
            return False
        
        try:
            self.restore()
            win32gui.SetForegroundWindow(self.hwnd)
            time.sleep(0.3)
            self._update_rect()
            return True
        except Exception as e:
            logger.debug(f"无法将窗口置于前台: {e}")
            return False
    
    def relative_to_absolute(self, rel_x, rel_y):
        """
        将相对坐标转换为绝对屏幕坐标
        
        Args:
            rel_x: 相对于窗口左上角的X坐标
            rel_y: 相对于窗口左上角的Y坐标
        
        Returns:
            (abs_x, abs_y): 绝对屏幕坐标
        """
        if not self._rect:
            self._update_rect()
        
        if self._rect:
            return (self._rect[0] + rel_x, self._rect[1] + rel_y)
        return (rel_x, rel_y)
    
    def print_info(self):
        """打印窗口信息"""
        rect = self.get_rect()
        if rect:
            logger.info("-" * 60)
            logger.info("微信窗口信息：")
            logger.info(f"  句柄: {self.hwnd}")
            logger.info(f"  标题: {self.window_title}")
            logger.info(f"  位置: 左={rect[0]}, 上={rect[1]}, 右={rect[2]}, 下={rect[3]}")
            logger.info(f"  大小: 宽={rect[2]-rect[0]}, 高={rect[3]-rect[1]}")
            logger.info("-" * 60)
            
            # 建议的坐标配置
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            logger.info("建议的坐标配置：")
            logger.info("")
            logger.info("WECHAT_COORDS = {")
            logger.info(f"    'window_width': {width},")
            logger.info(f"    'window_height': {height},")
            logger.info("    'session_list': {")
            logger.info("        'left': 0,")
            logger.info("        'top': 50,")
            logger.info(f"        'right': {int(width * 0.3)},  # 会话列表约占窗口宽度的30%")
            logger.info(f"        'bottom': {height - 50},")
            logger.info("        'item_height': 70,")
            logger.info("    },")
            logger.info("    'first_session': {")
            logger.info(f"        'x': {int(width * 0.15)},  # 会话列表中间")
            logger.info("        'y': 80,              # 第一个会话的Y坐标")
            logger.info("    },")
            logger.info("    'chat_area': {")
            logger.info(f"        'left': {int(width * 0.32)},")
            logger.info("        'top': 50,")
            logger.info(f"        'right': {int(width * 0.95)},")
            logger.info(f"        'bottom': {height - 100},")
            logger.info("    },")
            logger.info("}")
            logger.info("")


# ==============================================
# 鼠标键盘模拟模块
# ==============================================

class InputSimulator:
    """输入模拟器 - 模拟鼠标和键盘操作"""
    
    def __init__(self, wechat_window):
        self.window = wechat_window
    
    def click(self, x, y, relative=True):
        """
        点击指定位置
        
        Args:
            x: X坐标
            y: Y坐标
            relative: 是否为相对于窗口的坐标
        """
        if relative:
            abs_x, abs_y = self.window.relative_to_absolute(x, y)
        else:
            abs_x, abs_y = x, y
        
        if APP_CONFIG['debug']:
            logger.debug(f"点击位置: ({abs_x}, {abs_y})")
        
        # 使用 pyautogui 点击
        if LIBRARIES['pyautogui']:
            pyautogui.click(abs_x, abs_y)
        else:
            # 使用 win32api 模拟点击
            # 移动鼠标
            win32api.SetCursorPos((abs_x, abs_y))
            # 按下左键
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, abs_x, abs_y, 0, 0)
            time.sleep(0.05)
            # 释放左键
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, abs_x, abs_y, 0, 0)
        
        time.sleep(0.2)
    
    def double_click(self, x, y, relative=True):
        """双击"""
        self.click(x, y, relative)
        time.sleep(0.1)
        self.click(x, y, relative)
    
    def right_click(self, x, y, relative=True):
        """右键点击"""
        if relative:
            abs_x, abs_y = self.window.relative_to_absolute(x, y)
        else:
            abs_x, abs_y = x, y
        
        if LIBRARIES['pyautogui']:
            pyautogui.rightClick(abs_x, abs_y)
        else:
            win32api.SetCursorPos((abs_x, abs_y))
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, abs_x, abs_y, 0, 0)
            time.sleep(0.05)
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, abs_x, abs_y, 0, 0)
        
        time.sleep(0.2)
    
    def drag(self, from_x, from_y, to_x, to_y, relative=True):
        """拖拽（用于选中文本）"""
        if relative:
            abs_from = self.window.relative_to_absolute(from_x, from_y)
            abs_to = self.window.relative_to_absolute(to_x, to_y)
        else:
            abs_from = (from_x, from_y)
            abs_to = (to_x, to_y)
        
        if LIBRARIES['pyautogui']:
            pyautogui.moveTo(abs_from[0], abs_from[1])
            pyautogui.dragTo(abs_to[0], abs_to[1], duration=0.3, button='left')
        else:
            # 移动到起点
            win32api.SetCursorPos(abs_from)
            time.sleep(0.1)
            # 按下左键
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, abs_from[0], abs_from[1], 0, 0)
            time.sleep(0.1)
            # 移动到终点
            win32api.SetCursorPos(abs_to)
            time.sleep(0.2)
            # 释放左键
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, abs_to[0], abs_to[1], 0, 0)
        
        time.sleep(0.2)
    
    def copy(self):
        """模拟 Ctrl+C 复制"""
        if LIBRARIES['pyautogui']:
            pyautogui.hotkey('ctrl', 'c')
        else:
            # 按下 Ctrl
            win32api.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL
            time.sleep(0.05)
            # 按下 C
            win32api.keybd_event(0x43, 0, 0, 0)  # VK_C
            time.sleep(0.05)
            # 释放 C
            win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)
            # 释放 Ctrl
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(0.1)
    
    def select_all(self):
        """模拟 Ctrl+A 全选"""
        if LIBRARIES['pyautogui']:
            pyautogui.hotkey('ctrl', 'a')
        else:
            win32api.keybd_event(0x11, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(0x41, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(0x41, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)
            win32api.keybd_event(0x11, 0, win32con.KEYEVENTF_KEYUP, 0)
        
        time.sleep(0.1)


# ==============================================
# 消息对象和处理器
# ==============================================

class Message:
    """消息对象"""
    
    def __init__(self, sender, content, msg_type="私聊", group_name=None):
        self.sender = sender
        self.content = content
        self.type = msg_type
        self.group_name = group_name
        self.time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    def get_id(self):
        """获取消息唯一ID"""
        return f"{self.sender}|{self.content[:100]}|{self.time[:19]}"
    
    def __str__(self):
        if self.group_name:
            return f"[群聊] {self.group_name} - {self.sender}: {self.content[:30]}..."
        else:
            return f"[私聊] {self.sender}: {self.content[:30]}..."


class MessageProcessor:
    """消息处理器"""
    
    def __init__(self, email_sender):
        self.email_sender = email_sender
        self._processed = set()
        self._max_history = 2000
    
    def format_email(self, messages):
        """格式化邮件内容"""
        if not messages:
            return None, None
        
        if len(messages) == 1:
            msg = messages[0]
            if msg.group_name:
                subject = f"[微信] {msg.group_name} - {msg.sender}"
            else:
                subject = f"[微信] {msg.sender}"
        else:
            subject = f"[微信] 收到 {len(messages)} 条新消息"
        
        lines = [
            "=" * 60,
            "微信消息通知",
            "=" * 60,
            ""
        ]
        
        for i, msg in enumerate(messages, 1):
            lines.append(f"【消息 {i}】")
            lines.append(f"  类型: {msg.type}")
            if msg.group_name:
                lines.append(f"  群聊: {msg.group_name}")
            lines.append(f"  发送人: {msg.sender}")
            lines.append(f"  时间: {msg.time}")
            lines.append(f"  内容:")
            lines.append(f"    {msg.content}")
            lines.append("-" * 60)
            lines.append("")
        
        body = "\n".join(lines)
        return subject, body
    
    def process(self, messages):
        """处理消息列表"""
        if not messages:
            return
        
        new_msgs = []
        for msg in messages:
            if APP_CONFIG['only_new_messages']:
                msg_id = msg.get_id()
                if msg_id not in self._processed:
                    new_msgs.append(msg)
                    self._processed.add(msg_id)
            else:
                new_msgs.append(msg)
        
        if len(self._processed) > self._max_history:
            self._processed = set(list(self._processed)[-1000:])
        
        if not new_msgs:
            logger.debug("没有新消息需要处理")
            return
        
        for msg in new_msgs:
            logger.info(f"✓ 新消息: {msg}")
        
        subject, body = self.format_email(new_msgs)
        if subject and body:
            self.email_sender.send(subject, body)


# ==============================================
# 微信自动化控制器
# ==============================================

class WeChatController:
    """微信自动化控制器"""
    
    def __init__(self):
        self.window = WeChatWindow(WECHAT_COORDS.get('window_title', '微信'))
        self.input_sim = None
        self._current_chat = None  # 当前打开的聊天
    
    def initialize(self):
        """初始化"""
        logger.info("初始化微信控制器...")
        
        # 检查依赖
        if not LIBRARIES['win32']:
            logger.error("缺少必要的依赖库: pywin32")
            return False
        
        # 查找窗口
        if not self.window.find():
            return False
        
        # 置于前台
        if not self.window.bring_to_front():
            logger.warning("无法将微信窗口置于前台")
        
        # 创建输入模拟器
        self.input_sim = InputSimulator(self.window)
        
        # 打印窗口信息
        self.window.print_info()
        
        logger.info("✓ 微信控制器初始化成功")
        return True
    
    def click_session(self, session_index=0):
        """
        点击指定的会话项
        
        Args:
            session_index: 会话索引（0=第一个，1=第二个，...）
        """
        coords = WECHAT_COORDS
        
        # 根据索引计算Y坐标
        base_y = coords['first_session']['y']
        item_height = coords['session_list']['item_height']
        x = coords['first_session']['x']
        y = base_y + session_index * item_height
        
        logger.info(f"点击会话 #{session_index+1} (x={x}, y={y})")
        self.input_sim.click(x, y, relative=True)
        time.sleep(0.5)
    
    def get_chat_title(self):
        """
        获取当前聊天的标题（联系人名称或群聊名称）
        
        方法：点击标题区域，查看窗口标题或使用其他方式
        """
        # 注意：微信窗口标题会随着选中的聊天而变化
        # 我们可以通过获取窗口标题来判断
        title = self.window.window_title
        
        # 如果标题不是"微信"，可能是当前聊天的名称
        if title and title != '微信':
            logger.info(f"当前聊天标题: {title}")
            return title
        
        return None
    
    def select_and_copy_messages(self):
        """
        选中并复制聊天区域的消息
        
        方法：
        1. 点击聊天区域
        2. 全选 (Ctrl+A)
        3. 复制 (Ctrl+C)
        4. 从剪贴板获取文本
        """
        coords = WECHAT_COORDS['chat_area']
        
        # 点击聊天区域中间
        center_x = (coords['left'] + coords['right']) // 2
        center_y = (coords['top'] + coords['bottom']) // 2
        
        logger.info(f"点击聊天区域 (x={center_x}, y={center_y})")
        self.input_sim.click(center_x, center_y, relative=True)
        time.sleep(0.3)
        
        # 全选
        logger.info("全选文本 (Ctrl+A)")
        self.input_sim.select_all()
        time.sleep(0.3)
        
        # 复制
        logger.info("复制文本 (Ctrl+C)")
        Clipboard.clear()
        self.input_sim.copy()
        time.sleep(0.3)
        
        # 获取剪贴板内容
        text = Clipboard.get_text()
        
        if text:
            logger.info(f"获取到 {len(text)} 个字符")
            if APP_CONFIG['debug']:
                logger.debug(f"内容预览: {text[:200]}...")
        else:
            logger.warning("未能获取到文本内容")
        
        return text
    
    def parse_messages(self, text, chat_title=None):
        """
        解析复制的文本，提取消息
        
        注意：微信复制的文本格式通常是：
        发送人
        消息内容
        时间
        
        或者：
        [时间] 发送人: 消息内容
        
        这需要根据实际格式调整
        """
        messages = []
        
        if not text:
            return messages
        
        # 按行分割
        lines = text.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        if not lines:
            return messages
        
        logger.info(f"解析 {len(lines)} 行文本")
        
        # 简单的解析逻辑
        # 这需要根据实际的微信文本格式调整
        # 这里提供一个基础框架
        
        # 假设格式是：
        # 发送人
        # 消息内容
        # 时间
        # 或者更复杂的格式
        
        # 你需要根据实际复制的文本调整这个解析逻辑
        
        # 示例：如果有群聊名称
        if chat_title:
            # 尝试识别发送人
            # 这部分需要根据实际情况定制
            pass
        
        # 简单处理：将所有内容作为一条消息
        # 实际使用时应该解析出发送人、内容等
        if lines:
            # 尝试解析
            # 这里只是一个示例，你需要根据实际格式调整
            
            # 方法1：假设第一行是发送人
            sender = lines[0] if len(lines) > 0 else "未知"
            
            # 方法2：尝试查找常见的时间格式
            # 时间格式通常是: HH:MM 或 YYYY-MM-DD HH:MM
            
            # 简单的合并方式
            content = "\n".join(lines)
            
            # 判断是群聊还是私聊
            # 如果有群聊标题，并且内容中有多个发送人，可能是群聊
            msg_type = "私聊"
            group_name = None
            
            # 简单判断：如果有"群"字在标题中，或者内容格式像群聊
            if chat_title and ('群' in chat_title or '组' in chat_title):
                msg_type = "群聊"
                group_name = chat_title
            
            msg = Message(
                sender=sender,
                content=content,
                msg_type=msg_type,
                group_name=group_name
            )
            messages.append(msg)
        
        return messages
    
    def get_messages_from_chat(self, session_index=0):
        """
        从指定会话获取消息
        
        完整流程：
        1. 点击会话
        2. 获取聊天标题
        3. 选中并复制消息
        4. 解析消息
        """
        logger.info(f"{'='*60}")
        logger.info(f"获取会话 #{session_index+1} 的消息")
        logger.info(f"{'='*60}")
        
        # 1. 点击会话
        self.click_session(session_index)
        
        # 2. 获取聊天标题
        chat_title = self.get_chat_title()
        if not chat_title:
            chat_title = f"会话_{session_index+1}"
        
        # 3. 选中并复制
        text = self.select_and_copy_messages()
        
        # 4. 解析消息
        messages = self.parse_messages(text, chat_title)
        
        # 5. 如果没有解析到消息，创建一条包含全部内容的消息
        if not messages and text:
            msg = Message(
                sender="未知",
                content=text,
                msg_type="私聊",
                group_name=chat_title if '群' in chat_title else None
            )
            messages.append(msg)
        
        return messages


# ==============================================
# 坐标获取工具
# ==============================================

def get_mouse_position():
    """获取鼠标当前位置（用于配置坐标）"""
    print("\n" + "=" * 60)
    print("坐标获取工具")
    print("=" * 60)
    print("")
    print("使用方法：")
    print("  1. 将鼠标移动到你想获取坐标的位置")
    print("  2. 按回车键记录当前坐标")
    print("  3. 按 'q' 键退出")
    print("")
    print("提示：")
    print("  - 先打开微信窗口")
    print("  - 将鼠标移动到会话列表的第一个会话上")
    print("  - 按回车记录坐标，然后修改 WECHAT_COORDS 配置")
    print("")
    
    # 先查找微信窗口
    window = WeChatWindow()
    if window.find():
        window.bring_to_front()
        window.print_info()
    else:
        print("警告: 未找到微信窗口")
    
    print("\n" + "-" * 60)
    print("开始记录坐标（按回车记录，按 q 回车退出）")
    print("-" * 60)
    
    count = 1
    while True:
        user_input = input(f"\n请将鼠标移到位置 #{count}，然后按回车 (q退出): ").strip().lower()
        
        if user_input == 'q':
            break
        
        # 获取鼠标位置
        if LIBRARIES['pyautogui']:
            x, y = pyautogui.position()
        else:
            x, y = win32api.GetCursorPos()
        
        # 获取相对于微信窗口的坐标
        rel_x, rel_y = x, y
        if window.hwnd:
            rect = window.get_rect()
            if rect:
                rel_x = x - rect[0]
                rel_y = y - rect[1]
        
        print(f"  绝对坐标: ({x}, {y})")
        print(f"  相对坐标: ({rel_x}, {rel_y})")
        
        # 建议的配置
        print(f"  建议配置:")
        if count == 1:
            print(f"    'first_session': {{'x': {rel_x}, 'y': {rel_y}}},")
        elif count == 2:
            print(f"    'second_session': {{'x': {rel_x}, 'y': {rel_y}}},")
        
        count += 1


# ==============================================
# 主程序
# ==============================================

def run_get_coords_mode():
    """运行坐标获取模式"""
    logger.info("")
    logger.info("=" * 60)
    logger.info("运行模式: 坐标获取模式")
    logger.info("=" * 60)
    get_mouse_position()


def run_test_mode():
    """运行测试模式"""
    logger.info("")
    logger.info("=" * 60)
    logger.info("运行模式: 测试模式")
    logger.info("=" * 60)
    logger.info("")
    
    # 创建邮件发送器
    email_sender = EmailSender(EMAIL_CONFIG)
    
    # 测试邮件连接
    if not email_sender.test_connection():
        logger.error("邮件连接测试失败，请检查配置")
        return
    
    print("\n" + "=" * 60)
    print("测试选项")
    print("=" * 60)
    print("")
    print("  1. 发送测试邮件")
    print("  2. 测试微信控制（点击会话+复制消息）")
    print("  3. 测试剪贴板操作")
    print("  0. 退出")
    print("")
    
    while True:
        choice = input("请选择 (0-3): ").strip()
        
        if choice == '0':
            break
        
        elif choice == '1':
            # 发送测试邮件
            logger.info("发送测试邮件...")
            
            msg1 = Message(
                sender="张三",
                content="这是一条私聊测试消息。\n\n发送时间: " + datetime.now().strftime('%H:%M:%S'),
                msg_type="私聊"
            )
            
            msg2 = Message(
                sender="李四",
                content="这是一条群聊测试消息。\n\n大家好！",
                msg_type="群聊",
                group_name="测试群聊"
            )
            
            processor = MessageProcessor(email_sender)
            processor.process([msg1, msg2])
            
            logger.info("测试邮件已发送，请查收！")
        
        elif choice == '2':
            # 测试微信控制
            logger.info("测试微信控制...")
            
            controller = WeChatController()
            if not controller.initialize():
                logger.error("初始化失败")
                continue
            
            # 点击第一个会话
            messages = controller.get_messages_from_chat(0)
            
            if messages:
                processor = MessageProcessor(email_sender)
                processor.process(messages)
            else:
                logger.warning("未能获取到消息")
        
        elif choice == '3':
            # 测试剪贴板
            logger.info("测试剪贴板操作...")
            
            test_text = f"测试文本 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            logger.info(f"设置剪贴板: {test_text}")
            Clipboard.set_text(test_text)
            
            time.sleep(0.5)
            
            result = Clipboard.get_text()
            logger.info(f"获取剪贴板: {result}")
            
            if result == test_text:
                logger.info("✓ 剪贴板操作正常")
            else:
                logger.error("✗ 剪贴板操作异常")
        
        else:
            print("无效选项")
    
    email_sender.disconnect()
    logger.info("测试模式结束")


def run_normal_mode():
    """运行正常模式"""
    logger.info("")
    logger.info("=" * 60)
    logger.info("运行模式: 正常模式")
    logger.info("=" * 60)
    logger.info("")
    
    # 检查配置
    if EMAIL_CONFIG['sender_email'] == 'your_email@qq.com':
        logger.error("错误: 请先修改邮箱配置 EMAIL_CONFIG")
        logger.error("  sender_email: 你的邮箱地址")
        logger.error("  sender_password: 你的邮箱授权码")
        logger.error("  receiver_email: 收件人邮箱")
        return
    
    # 创建组件
    email_sender = EmailSender(EMAIL_CONFIG)
    processor = MessageProcessor(email_sender)
    controller = WeChatController()
    
    # 初始化
    if not controller.initialize():
        logger.error("初始化失败")
        return
    
    # 测试邮件连接
    logger.info("测试邮件连接...")
    if not email_sender.test_connection():
        logger.error("邮件连接失败")
        return
    email_sender.disconnect()  # 断开，让processor管理
    
    # 主循环
    logger.info("")
    logger.info("=" * 60)
    logger.info("开始监控微信消息...")
    logger.info(f"检查间隔: {APP_CONFIG['check_interval']} 秒")
    logger.info("按 Ctrl+C 停止")
    logger.info("=" * 60)
    logger.info("")
    
    # 记录已处理的会话
    processed_sessions = set()
    
    try:
        while True:
            logger.info("-" * 60)
            logger.info(f"[{datetime.now().strftime('%H:%M:%S')}] 检查消息...")
            
            # 确保窗口在前台
            controller.window.bring_to_front()
            time.sleep(0.5)
            
            # 获取前几个会话的消息
            # 你可以根据需要调整检查的会话数量
            for i in range(3):  # 检查前3个会话
                try:
                    messages = controller.get_messages_from_chat(i)
                    
                    if messages:
                        processor.process(messages)
                    else:
                        logger.debug(f"会话 #{i+1} 没有新消息")
                        
                except Exception as e:
                    logger.error(f"处理会话 #{i+1} 时出错: {e}")
                    import traceback
                    logger.debug(traceback.format_exc())
            
            # 等待
            logger.info(f"等待 {APP_CONFIG['check_interval']} 秒后检查...")
            time.sleep(APP_CONFIG['check_interval'])
            
    except KeyboardInterrupt:
        logger.info("收到停止信号")
    finally:
        email_sender.disconnect()
    
    logger.info("正常模式结束")


def main():
    """主函数"""
    # 检查命令行参数
    if len(sys.argv) > 1:
        if sys.argv[1] == '--get-coords':
            run_get_coords_mode()
            return
    
    # 根据配置选择模式
    run_mode = APP_CONFIG['run_mode']
    
    print("")
    print("=" * 60)
    print("微信消息自动监听与转发工具")
    print("=" * 60)
    print("")
    print(f"当前运行模式: {run_mode}")
    print("")
    print("可用模式：")
    print("  get_coords - 坐标获取模式：帮助配置微信窗口坐标")
    print("  test       - 测试模式：测试邮件和微信控制")
    print("  normal     - 正常模式：自动监听消息")
    print("")
    print("提示：")
    print("  - 首次使用请先运行 get_coords 模式配置坐标")
    print("  - 或者直接修改 WECHAT_COORDS 配置")
    print("")
    
    if run_mode == 'get_coords':
        run_get_coords_mode()
    elif run_mode == 'test':
        run_test_mode()
    elif run_mode == 'normal':
        run_normal_mode()
    else:
        logger.error(f"未知的运行模式: {run_mode}")
        logger.error("请选择: get_coords, test, 或 normal")


if __name__ == '__main__':
    main()
