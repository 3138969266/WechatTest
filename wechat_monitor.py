# -*- coding: utf-8 -*-
"""
微信消息监听与邮件转发脚本
功能：自动接收微信新消息（私聊和群消息），并转发到指定邮箱
依赖：uiautomation, pywin32

使用方法：
1. 安装依赖: pip install uiautomation pywin32
2. 修改下方的邮箱配置
3. 登录微信PC客户端
4. 运行脚本: python wechat_monitor.py
"""

import logging
import smtplib
import time
import threading
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from queue import Queue

# 尝试导入依赖库
try:
    import uiautomation as auto
    UIAUTO_AVAILABLE = True
except ImportError:
    UIAUTO_AVAILABLE = False
    print("警告: 未安装uiautomation库，请运行: pip install uiautomation")

try:
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("警告: 未安装pywin32库，请运行: pip install pywin32")


# ==============================================
# 配置区域 - 请根据实际情况修改以下配置
# ==============================================

# 邮箱配置
EMAIL_CONFIG = {
    'sender_email': 'your_email@qq.com',
    'sender_password': 'your_authorization_code',
    'receiver_email': 'receiver@example.com',
    'smtp_server': 'smtp.qq.com',
    'smtp_port': 465,
    'use_ssl': True,
}

# 微信配置
WECHAT_CONFIG = {
    'window_title': '微信',
    'check_interval': 2,
    'only_new_messages': True,
    'auto_open_chat': True,
}

# 日志配置
LOG_CONFIG = {
    'level': logging.INFO,
    'format': '%(asctime)s - %(levelname)s - %(message)s',
    'console': True,
    'file': False,
    'file_path': 'wechat_monitor.log',
}

# ==============================================
# 日志初始化
# ==============================================

def setup_logger():
    logger = logging.getLogger('WeChatMonitor')
    logger.setLevel(LOG_CONFIG['level'])
    
    formatter = logging.Formatter(LOG_CONFIG['format'])
    
    if LOG_CONFIG['console']:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    if LOG_CONFIG['file']:
        file_handler = logging.FileHandler(LOG_CONFIG['file_path'], encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger

logger = setup_logger()


# ==============================================
# 邮件发送模块
# ==============================================

class EmailSender:
    """邮件发送器"""
    
    def __init__(self, config):
        self.config = config
        self._connected = False
        self._server = None
    
    def connect(self):
        """连接到SMTP服务器"""
        try:
            if self.config['use_ssl']:
                self._server = smtplib.SMTP_SSL(
                    self.config['smtp_server'], 
                    self.config['smtp_port'],
                    timeout=10
                )
            else:
                self._server = smtplib.SMTP(
                    self.config['smtp_server'], 
                    self.config['smtp_port'],
                    timeout=10
                )
                self._server.ehlo()
            
            self._server.login(
                self.config['sender_email'], 
                self.config['sender_password']
            )
            self._connected = True
            logger.debug("SMTP服务器连接成功")
            return True
            
        except Exception as e:
            logger.error(f"SMTP服务器连接失败: {str(e)}")
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
    
    def send(self, subject, body, receiver=None):
        """发送邮件"""
        if receiver is None:
            receiver = self.config['receiver_email']
        
        try:
            if not self._connected:
                if not self.connect():
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
            
            logger.info(f"邮件发送成功: {subject}")
            return True
            
        except Exception as e:
            logger.error(f"邮件发送失败: {str(e)}")
            self._connected = False
            return False


# ==============================================
# 微信消息处理器
# ==============================================

class MessageProcessor:
    """消息处理器"""
    
    def __init__(self, email_sender):
        self.email_sender = email_sender
        self.processed_messages = set()
        self.max_history = 1000
    
    def format_message(self, sender, content, msg_type="私聊", group_name=None):
        """格式化消息"""
        now = datetime.now()
        return {
            'sender': sender,
            'content': content,
            'type': msg_type,
            'group_name': group_name,
            'time': now.strftime('%Y-%m-%d %H:%M:%S'),
            'timestamp': now.timestamp()
        }
    
    def create_email_content(self, messages):
        """创建邮件内容"""
        if not messages:
            return None, None
        
        if len(messages) == 1:
            msg = messages[0]
            if msg['type'] == '群聊' and msg['group_name']:
                subject = f"[微信] {msg['group_name']} - {msg['sender']}"
            else:
                subject = f"[微信] {msg['sender']}"
        else:
            subject = f"[微信] 收到 {len(messages)} 条新消息"
        
        body_lines = [
            "=" * 60,
            "微信消息通知",
            "=" * 60,
            ""
        ]
        
        for i, msg in enumerate(messages, 1):
            body_lines.append(f"【消息 {i}】")
            body_lines.append(f"  类型: {msg['type']}")
            if msg['group_name']:
                body_lines.append(f"  群聊: {msg['group_name']}")
            body_lines.append(f"  发送人: {msg['sender']}")
            body_lines.append(f"  时间: {msg['time']}")
            body_lines.append(f"  内容:")
            body_lines.append(f"    {msg['content']}")
            body_lines.append("-" * 60)
            body_lines.append("")
        
        body = "\n".join(body_lines)
        return subject, body
    
    def is_new_message(self, msg):
        """检查是否为新消息"""
        msg_id = f"{msg['sender']}|{msg['content']}|{msg['time'][:19]}"
        return msg_id not in self.processed_messages
    
    def mark_processed(self, msg):
        """标记为已处理"""
        msg_id = f"{msg['sender']}|{msg['content']}|{msg['time'][:19]}"
        self.processed_messages.add(msg_id)
        
        if len(self.processed_messages) > self.max_history:
            self.processed_messages = set(list(self.processed_messages)[-500:])
    
    def process(self, messages):
        """处理消息列表"""
        if not messages:
            return
        
        new_messages = []
        for msg in messages:
            if WECHAT_CONFIG['only_new_messages']:
                if self.is_new_message(msg):
                    new_messages.append(msg)
                    self.mark_processed(msg)
            else:
                new_messages.append(msg)
        
        if not new_messages:
            logger.debug("没有新消息需要处理")
            return
        
        for msg in new_messages:
            if msg['type'] == '群聊' and msg['group_name']:
                logger.info(f"[群聊] {msg['group_name']} - {msg['sender']}: {msg['content'][:60]}")
            else:
                logger.info(f"[私聊] {msg['sender']}: {msg['content'][:60]}")
        
        subject, body = self.create_email_content(new_messages)
        if subject and body:
            self.email_sender.send(subject, body)


# ==============================================
# 微信窗口控制器
# ==============================================

class WeChatWindow:
    """微信窗口控制器"""
    
    def __init__(self, window_title="微信"):
        self.window_title = window_title
        self.hwnd = None
    
    def find(self):
        """查找微信窗口"""
        if not WIN32_AVAILABLE:
            logger.error("pywin32库未安装")
            return False
        
        self.hwnd = win32gui.FindWindow(None, self.window_title)
        
        if self.hwnd == 0:
            logger.warning(f"未找到微信窗口，请确保微信已登录，窗口标题为: {self.window_title}")
            return False
        
        logger.info(f"找到微信窗口，句柄: {self.hwnd}")
        return True
    
    def is_minimized(self):
        """检查窗口是否最小化"""
        if not self.hwnd:
            return True
        placement = win32gui.GetWindowPlacement(self.hwnd)
        return placement[1] == win32con.SW_SHOWMINIMIZED
    
    def restore(self):
        """恢复窗口"""
        if self.hwnd:
            win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
            time.sleep(0.1)
    
    def bring_to_front(self):
        """将窗口置于前台"""
        if not self.hwnd:
            return False
        
        try:
            if self.is_minimized():
                self.restore()
            
            win32gui.SetForegroundWindow(self.hwnd)
            time.sleep(0.2)
            logger.debug("微信窗口已置于前台")
            return True
        except Exception as e:
            logger.error(f"无法将窗口置于前台: {str(e)}")
            return False
    
    def get_rect(self):
        """获取窗口位置和大小"""
        if self.hwnd:
            return win32gui.GetWindowRect(self.hwnd)
        return None


# ==============================================
# 微信消息监听器（UI自动化版本）
# ==============================================

class WeChatMessageListener:
    """微信消息监听器"""
    
    def __init__(self):
        self.window = WeChatWindow(WECHAT_CONFIG['window_title'])
        self.running = False
        self.message_queue = Queue()
        
        # 已发现的会话和消息
        self.discovered_sessions = set()
        self.session_messages = {}  # {session_name: last_content}
    
    def initialize(self):
        """初始化"""
        if not UIAUTO_AVAILABLE or not WIN32_AVAILABLE:
            logger.error("缺少必要的依赖库")
            return False
        
        if not self.window.find():
            return False
        
        logger.info("微信消息监听器初始化完成")
        return True
    
    def get_wechat_window(self):
        """获取微信UI自动化窗口"""
        try:
            wechat_win = auto.WindowControl(Name=WECHAT_CONFIG['window_title'])
            if wechat_win.Exists(maxSearchSeconds=2):
                return wechat_win
        except Exception as e:
            logger.debug(f"获取微信窗口时出错: {str(e)}")
        return None
    
    def scan_sessions(self):
        """扫描会话列表（左侧联系人列表）"""
        sessions = []
        
        try:
            wechat_win = self.get_wechat_window()
            if not wechat_win:
                return sessions
            
            # 方法1: 查找所有可能包含未读消息的控件
            # 微信UI结构可能因版本不同而有所差异
            # 这里使用通用的搜索方法
            
            # 查找包含"条未读消息"文本的控件
            def find_unread_controls(control, depth=0, max_depth=5):
                if depth > max_depth:
                    return
                
                try:
                    name = control.Name
                    if name and ("条未读消息" in name or "未读" in name):
                        sessions.append({
                            'control': control,
                            'name': name,
                            'has_unread': True
                        })
                except:
                    pass
                
                try:
                    for child in control.GetChildren():
                        find_unread_controls(child, depth + 1, max_depth)
                except:
                    pass
            
            find_unread_controls(wechat_win)
            
            if sessions:
                logger.debug(f"发现 {len(sessions)} 个会话有未读消息提示")
            
        except Exception as e:
            logger.error(f"扫描会话时出错: {str(e)}")
        
        return sessions
    
    def get_current_chat_messages(self):
        """
        获取当前聊天窗口的消息
        
        注意：由于微信UI结构的复杂性，这个方法需要根据实际微信版本调整
        建议使用 UI Spy 或 uiautomation 的 inspector 工具查看实际的控件结构
        """
        messages = []
        
        try:
            wechat_win = self.get_wechat_window()
            if not wechat_win:
                return messages
            
            # 获取当前聊天窗口的标题（通常是联系人或群聊名称）
            chat_title = wechat_win.Name
            if chat_title == WECHAT_CONFIG['window_title']:
                # 可能没有选中任何聊天
                pass
            
            # 方法：查找消息列表区域
            # 微信的消息通常在一个列表控件中
            # 以下是一个通用的搜索框架
            
            # 尝试查找所有文本控件
            text_controls = []
            
            def find_text_controls(control, depth=0, max_depth=8):
                if depth > max_depth:
                    return
                
                try:
                    control_type = control.ControlTypeName
                    name = control.Name
                    
                    if control_type == 'TextControl' and name and name.strip():
                        text_controls.append({
                            'name': name,
                            'control': control
                        })
                except:
                    pass
                
                try:
                    for child in control.GetChildren():
                        find_text_controls(child, depth + 1, max_depth)
                except:
                    pass
            
            find_text_controls(wechat_win)
            
            # 分析文本控件，尝试识别消息
            # 这部分需要根据实际微信UI结构调整
            logger.debug(f"当前聊天窗口发现 {len(text_controls)} 个文本控件")
            
            # 简单的启发式方法：假设文本控件按顺序排列，包含发送者和消息内容
            # 这里仅做示例，实际使用时需要根据微信版本调整
            
        except Exception as e:
            logger.error(f"获取聊天消息时出错: {str(e)}")
        
        return messages
    
    def create_test_message(self):
        """创建测试消息（用于演示）"""
        now = datetime.now()
        return {
            'sender': '测试用户',
            'content': f'这是一条测试消息，发送时间: {now.strftime("%H:%M:%S")}',
            'type': '私聊',
            'group_name': None,
            'time': now.strftime('%Y-%m-%d %H:%M:%S'),
            'timestamp': now.timestamp()
        }
    
    def check_for_new_messages(self):
        """检查新消息"""
        new_messages = []
        
        # 方法1: 扫描会话列表中的未读消息
        sessions = self.scan_sessions()
        
        for session in sessions:
            session_name = session['name']
            
            # 如果是新发现的会话，或者之前已经处理过
            if session_name not in self.discovered_sessions:
                logger.info(f"发现新会话: {session_name}")
                self.discovered_sessions.add(session_name)
                
                # 尝试打开这个会话获取消息
                if WECHAT_CONFIG['auto_open_chat']:
                    try:
                        # 点击会话（需要调整坐标或控件）
                        # 这里仅做示例
                        pass
                    except:
                        pass
        
        # 方法2: 获取当前聊天窗口的消息
        current_messages = self.get_current_chat_messages()
        
        # 注意：由于微信UI结构的复杂性，
        # 实际使用时你可能需要根据自己的微信版本调整UI自动化代码
        
        # 建议：
        # 1. 使用 uiautomation 自带的 inspector 工具查看微信的控件结构
        # 2. 或者使用 UI Spy 工具
        # 3. 根据实际的控件名称和类型调整代码
        
        return new_messages
    
    def start(self, processor, test_mode=True):
        """开始监听"""
        if not self.initialize():
            logger.error("初始化失败，无法启动监听")
            return
        
        self.running = True
        
        logger.info("=" * 60)
        logger.info("微信消息监听器已启动")
        logger.info("=" * 60)
        
        # 发送启动通知
        startup_msg = self.create_test_message()
        startup_msg['sender'] = '系统'
        startup_msg['content'] = f"""
微信消息监控程序已启动！

启动时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
配置：
- 检查间隔：{WECHAT_CONFIG['check_interval']} 秒
- 收件人：{EMAIL_CONFIG['receiver_email']}

{"【测试模式】每30秒发送一条测试消息" if test_mode else ""}
"""
        processor.process([startup_msg])
        
        test_counter = 0
        
        try:
            while self.running:
                # 检查新消息
                messages = self.check_for_new_messages()
                
                # 处理消息
                if messages:
                    processor.process(messages)
                
                # 测试模式：定时发送测试消息
                if test_mode:
                    test_counter += 1
                    if test_counter >= 15:  # 每30秒（15次检查）
                        test_msg = self.create_test_message()
                        processor.process([test_msg])
                        test_counter = 0
                
                # 等待
                time.sleep(WECHAT_CONFIG['check_interval'])
                
        except KeyboardInterrupt:
            logger.info("收到停止信号")
        finally:
            self.stop()
    
    def stop(self):
        """停止监听"""
        self.running = False
        logger.info("微信消息监听器已停止")


# ==============================================
# 主程序
# ==============================================

def main():
    """主函数"""
    logger.info("=" * 60)
    logger.info("微信消息监控程序")
    logger.info("=" * 60)
    logger.info(f"发件人: {EMAIL_CONFIG['sender_email']}")
    logger.info(f"收件人: {EMAIL_CONFIG['receiver_email']}")
    logger.info(f"SMTP: {EMAIL_CONFIG['smtp_server']}:{EMAIL_CONFIG['smtp_port']}")
    logger.info("=" * 60)
    
    # 检查依赖
    if not UIAUTO_AVAILABLE or not WIN32_AVAILABLE:
        logger.error("缺少必要的依赖库，请运行:")
        logger.error("  pip install uiautomation pywin32")
        return
    
    # 创建组件
    email_sender = EmailSender(EMAIL_CONFIG)
    processor = MessageProcessor(email_sender)
    listener = WeChatMessageListener()
    
    # 启动监听
    # test_mode=True 表示启用测试模式，会定时发送测试消息
    # 实际使用时请设置为 False
    listener.start(processor, test_mode=True)


if __name__ == '__main__':
    main()
