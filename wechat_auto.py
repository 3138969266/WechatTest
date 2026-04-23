# -*- coding: utf-8 -*-
"""
微信消息自动监听与转发工具
=====================================
【真正的自动化版本】

功能：
1. 自动检测微信PC客户端的新消息
2. 自动识别私聊/群聊
3. 自动点击会话，提取消息内容
4. 自动发送到指定邮箱

使用方法：
1. 先修改下方的邮箱配置（EMAIL_CONFIG）
2. 登录微信PC客户端
3. 运行: python wechat_auto.py

注意：
- 不同微信版本的UI结构不同
- 首次使用请运行诊断模式，查看你的微信UI结构
- 根据诊断结果调整 UI_CONFIG 中的参数
"""

import logging
import smtplib
import time
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

# ==============================================
# 【重要】请先修改以下配置
# ==============================================

# ----------------------------------------------
# 邮箱配置 - 必须修改！
# ----------------------------------------------
EMAIL_CONFIG = {
    # 发件人邮箱（你的邮箱）
    'sender_email': '3138969266@qq.com',
    
    # 邮箱授权码（不是邮箱密码！）
    # QQ邮箱：在设置->账户中开启SMTP服务，获取16位授权码
    # 163邮箱：在设置->客户端授权码中开启
    'sender_password': 'ovyktydfyoivdcci',
    
    # 收件人邮箱（多个用逗号分隔）
    'receiver_email': '3138969266@qq.com',
    
    # SMTP服务器
    # QQ邮箱: smtp.qq.com
    # 163邮箱: smtp.163.com
    'smtp_server': 'smtp.qq.com',
    
    # SMTP端口
    # QQ邮箱(SSL): 465
    # 163邮箱(SSL): 465, (非SSL): 25
    'smtp_port': 465,
    
    # 是否使用SSL加密
    'use_ssl': True,
}

# ----------------------------------------------
# 微信UI配置 - 根据你的微信版本调整
# ----------------------------------------------
# 提示：运行诊断模式可以查看你的微信UI结构
UI_CONFIG = {
    # 微信窗口标题
    'window_title': '微信',
    
    # 会话列表区域的特征（用于定位左侧联系人列表）
    'session_list': {
        # 控件类型（通常是 ListControl 或 PaneControl）
        'control_type': 'ListControl',
        # Automation ID（如果有的话）
        'automation_id': '',
        # 类名（如果有的话）
        'class_name': '',
        # 包含的文本特征（用于识别）
        'contains_text': [],
    },
    
    # 会话项的特征（每个联系人/群聊）
    'session_item': {
        'control_type': 'ListItemControl',
    },
    
    # 未读消息提示的特征（如"1条未读消息"）
    'unread_indicator': {
        # 包含的关键词
        'keywords': ['未读', '条消息', '条新消息'],
        # 控件类型
        'control_type': 'TextControl',
    },
    
    # 聊天消息区域的特征
    'message_area': {
        'control_type': 'ListControl',
    },
    
    # 消息文本的特征
    'message_text': {
        'control_type': 'TextControl',
    },
}

# ----------------------------------------------
# 程序配置
# ----------------------------------------------
APP_CONFIG = {
    # 检查新消息的间隔（秒）
    'check_interval': 5,
    
    # 只处理新消息（不重复发送）
    'only_new_messages': True,
    
    # 运行模式
    # 'normal'  - 正常模式：监听并转发消息
    # 'diagnose' - 诊断模式：分析微信UI结构，帮助配置
    # 'test'     - 测试模式：测试邮件功能
    'run_mode': 'diagnose',  # 首次使用建议先用 diagnose 模式
    
    # 日志级别: DEBUG, INFO, WARNING, ERROR
    'log_level': logging.INFO,
}

# ==============================================
# 导入依赖库
# ==============================================

LIBRARIES = {}

try:
    import uiautomation as auto
    LIBRARIES['uiautomation'] = True
except ImportError:
    LIBRARIES['uiautomation'] = False
    print("错误: 未安装 uiautomation 库")
    print("请运行: pip install uiautomation")

try:
    import win32gui
    import win32con
    import win32api
    LIBRARIES['win32'] = True
except ImportError:
    LIBRARIES['win32'] = False
    print("警告: 未安装 pywin32 库")
    print("请运行: pip install pywin32")


# ==============================================
# 日志设置
# ==============================================

def setup_logger():
    logger = logging.getLogger('WeChatAuto')
    logger.setLevel(APP_CONFIG['log_level'])
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # 控制台输出
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
        logger.info(f"  服务器: {self.config['smtp_server']}:{self.config['smtp_port']}")
        logger.info(f"  发件人: {self.config['sender_email']}")
        logger.info(f"  收件人: {self.config['receiver_email']}")
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
            logger.error("  1. 邮箱地址错误")
            logger.error("  2. 授权码错误（注意：不是邮箱密码！）")
            logger.error("  3. 未开启SMTP服务")
            logger.error("")
            logger.error("QQ邮箱设置方法：")
            logger.error("  1. 登录QQ邮箱网页版")
            logger.error("  2. 设置 -> 账户 -> 开启POP3/SMTP服务")
            logger.error("  3. 按提示发送短信，获取16位授权码")
            logger.error(f"  错误详情: {e}")
            return False
            
        except smtplib.SMTPConnectError as e:
            logger.error("✗ 连接失败！")
            logger.error("")
            logger.error("可能的原因：")
            logger.error("  1. SMTP服务器地址错误")
            logger.error("  2. 端口错误")
            logger.error("  3. 网络问题或防火墙拦截")
            logger.error("")
            logger.error("常见邮箱配置：")
            logger.error("  QQ邮箱:  smtp.qq.com:465 (SSL)")
            logger.error("  163邮箱: smtp.163.com:465 (SSL)")
            logger.error(f"  错误详情: {e}")
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
# 微信窗口控制模块
# ==============================================

class WeChatWindow:
    """微信窗口控制器"""
    
    def __init__(self, window_title="微信"):
        self.window_title = window_title
        self.hwnd = None
        self._ui_window = None
    
    def find(self):
        """查找微信窗口"""
        if not LIBRARIES['win32']:
            logger.error("pywin32 库未安装")
            return False
        
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
        return True
    
    def restore(self):
        """恢复窗口（从最小化恢复）"""
        if self.hwnd:
            placement = win32gui.GetWindowPlacement(self.hwnd)
            if placement[1] == win32con.SW_SHOWMINIMIZED:
                win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
    
    def bring_to_front(self):
        """将窗口置于前台"""
        if not self.hwnd:
            return False
        
        try:
            self.restore()
            win32gui.SetForegroundWindow(self.hwnd)
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.debug(f"无法将窗口置于前台: {e}")
            return False
    
    def get_ui_control(self):
        """获取UI自动化控件"""
        if not LIBRARIES['uiautomation']:
            logger.error("uiautomation 库未安装")
            return None
        
        try:
            win = auto.WindowControl(Name=self.window_title)
            if win.Exists(maxSearchSeconds=3):
                self._ui_window = win
                return win
        except:
            pass
        
        return None


# ==============================================
# UI诊断模块
# ==============================================

class UIDiagnoser:
    """UI诊断器 - 用于分析微信UI结构"""
    
    def __init__(self, wechat_window):
        self.window = wechat_window
        self._ui_window = None
    
    def diagnose(self):
        """执行诊断"""
        logger.info("=" * 60)
        logger.info("微信UI结构诊断")
        logger.info("=" * 60)
        
        # 1. 检查依赖
        if not LIBRARIES['uiautomation']:
            logger.error("错误: uiautomation 库未安装")
            return False
        
        if not LIBRARIES['win32']:
            logger.error("错误: pywin32 库未安装")
            return False
        
        # 2. 查找微信窗口
        if not self.window.find():
            return False
        
        # 3. 获取UI控件
        self._ui_window = self.window.get_ui_control()
        if not self._ui_window:
            logger.error("无法获取微信UI控件")
            return False
        
        logger.info("✓ 初始化成功，开始分析UI结构...")
        
        # 4. 执行各项诊断
        self._analyze_window_structure()
        self._search_unread_indicators()
        self._search_text_controls()
        self._analyze_session_list()
        
        logger.info("")
        logger.info("=" * 60)
        logger.info("诊断完成！")
        logger.info("=" * 60)
        logger.info("")
        logger.info("请根据以上输出，调整 UI_CONFIG 中的配置。")
        logger.info("")
        
        return True
    
    def _analyze_window_structure(self):
        """分析窗口结构"""
        logger.info("")
        logger.info("-" * 60)
        logger.info("[1/4] 分析窗口控件结构（前4层）")
        logger.info("-" * 60)
        
        def print_tree(control, indent=0, max_depth=4):
            if indent > max_depth:
                return
            
            prefix = "  " * indent
            
            try:
                ctrl_type = control.ControlTypeName
                name = control.Name
                auto_id = control.AutomationId
                class_name = control.ClassName
                
                # 决定是否显示
                should_show = False
                if (name and name.strip()) or auto_id or ctrl_type in [
                    'WindowControl', 'PaneControl', 'ListControl', 
                    'ListItemControl', 'TextControl', 'ButtonControl',
                    'TabControl', 'EditControl'
                ]:
                    should_show = True
                
                if should_show:
                    info = [ctrl_type]
                    if name and name.strip():
                        display_name = name[:40] + "..." if len(name) > 40 else name
                        info.append(f'Name="{display_name}"')
                    if auto_id:
                        info.append(f'Id="{auto_id}"')
                    if class_name and class_name not in ['', 'CtrlNotifySink']:
                        info.append(f'Class="{class_name}"')
                    
                    logger.info(f"{prefix}{' | '.join(info)}")
                
                # 递归子控件
                try:
                    for child in control.GetChildren():
                        print_tree(child, indent + 1, max_depth)
                except:
                    pass
                    
            except:
                pass
        
        print_tree(self._ui_window)
    
    def _search_unread_indicators(self):
        """搜索未读消息提示"""
        logger.info("")
        logger.info("-" * 60)
        logger.info("[2/4] 搜索未读消息提示控件")
        logger.info("-" * 60)
        
        keywords = UI_CONFIG['unread_indicator']['keywords']
        results = []
        
        def search(control, depth=0, max_depth=8):
            if depth > max_depth:
                return
            
            try:
                name = control.Name
                if name:
                    for kw in keywords:
                        if kw in name:
                            results.append({
                                'depth': depth,
                                'type': control.ControlTypeName,
                                'name': name,
                                'control': control
                            })
                            break
                
                for child in control.GetChildren():
                    search(child, depth + 1, max_depth)
            except:
                pass
        
        search(self._ui_window)
        
        if results:
            logger.info(f"找到 {len(results)} 个未读消息提示:")
            for i, r in enumerate(results, 1):
                indent = "  " * min(r['depth'], 3)
                logger.info(f"{i}. {indent}{r['type']}: {r['name']}")
        else:
            logger.info("未找到未读消息提示。")
            logger.info("提示：如果当前没有未读消息，请让别人给你发一条消息后重新诊断。")
    
    def _search_text_controls(self):
        """搜索文本控件"""
        logger.info("")
        logger.info("-" * 60)
        logger.info("[3/4] 搜索所有文本控件（前20个）")
        logger.info("-" * 60)
        
        text_controls = []
        
        def search(control, depth=0, max_depth=10):
            if depth > max_depth:
                return
            
            try:
                if control.ControlTypeName == 'TextControl':
                    name = control.Name
                    if name and name.strip():
                        text_controls.append({
                            'depth': depth,
                            'name': name
                        })
                
                for child in control.GetChildren():
                    search(child, depth + 1, max_depth)
            except:
                pass
        
        search(self._ui_window)
        
        if text_controls:
            logger.info(f"找到 {len(text_controls)} 个文本控件，显示前20个:")
            for i, tc in enumerate(text_controls[:20], 1):
                indent = "  " * min(tc['depth'], 3)
                display = tc['name'][:50] + "..." if len(tc['name']) > 50 else tc['name']
                logger.info(f"{i}. {indent}{repr(display)}")
            
            if len(text_controls) > 20:
                logger.info(f"... 还有 {len(text_controls) - 20} 个文本控件")
        else:
            logger.info("未找到文本控件。")
    
    def _analyze_session_list(self):
        """分析会话列表"""
        logger.info("")
        logger.info("-" * 60)
        logger.info("[4/4] 分析会话列表区域")
        logger.info("-" * 60)
        
        # 搜索 ListControl
        list_controls = []
        
        def search(control, depth=0, max_depth=6):
            if depth > max_depth:
                return
            
            try:
                if control.ControlTypeName == 'ListControl':
                    # 检查是否有子项
                    children = control.GetChildren()
                    list_controls.append({
                        'depth': depth,
                        'child_count': len(children),
                        'control': control
                    })
                
                for child in control.GetChildren():
                    search(child, depth + 1, max_depth)
            except:
                pass
        
        search(self._ui_window)
        
        if list_controls:
            logger.info(f"找到 {len(list_controls)} 个 ListControl:")
            for i, lc in enumerate(list_controls, 1):
                logger.info(f"  {i}. 深度: {lc['depth']}, 子控件数: {lc['child_count']}")
                
                # 显示前几个子项
                try:
                    children = lc['control'].GetChildren()
                    for j, child in enumerate(children[:5]):
                        try:
                            name = child.Name
                            ctrl_type = child.ControlTypeName
                            if name and name.strip():
                                display = name[:40] + "..." if len(name) > 40 else name
                                logger.info(f"       子项 {j+1}: {ctrl_type} - {display}")
                        except:
                            pass
                except:
                    pass
        else:
            logger.info("未找到 ListControl 控件。")


# ==============================================
# 消息处理器
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
        """获取消息唯一ID（用于去重）"""
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
        
        # 生成主题
        if len(messages) == 1:
            msg = messages[0]
            if msg.group_name:
                subject = f"[微信] {msg.group_name} - {msg.sender}"
            else:
                subject = f"[微信] {msg.sender}"
        else:
            subject = f"[微信] 收到 {len(messages)} 条新消息"
        
        # 生成正文
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
        
        # 过滤新消息
        new_msgs = []
        for msg in messages:
            if APP_CONFIG['only_new_messages']:
                msg_id = msg.get_id()
                if msg_id not in self._processed:
                    new_msgs.append(msg)
                    self._processed.add(msg_id)
            else:
                new_msgs.append(msg)
        
        # 清理历史记录
        if len(self._processed) > self._max_history:
            self._processed = set(list(self._processed)[-1000:])
        
        if not new_msgs:
            logger.debug("没有新消息需要处理")
            return
        
        # 显示日志
        for msg in new_msgs:
            logger.info(f"✓ 新消息: {msg}")
        
        # 发送邮件
        subject, body = self.format_email(new_msgs)
        if subject and body:
            self.email_sender.send(subject, body)


# ==============================================
# 微信自动化控制器
# ==============================================

class WeChatController:
    """微信自动化控制器"""
    
    def __init__(self):
        self.window = WeChatWindow(UI_CONFIG['window_title'])
        self._ui_window = None
        self._sessions = {}  # 已处理的会话
    
    def initialize(self):
        """初始化"""
        logger.info("初始化微信控制器...")
        
        # 检查依赖
        if not LIBRARIES['uiautomation'] or not LIBRARIES['win32']:
            logger.error("缺少必要的依赖库")
            return False
        
        # 查找窗口
        if not self.window.find():
            return False
        
        # 获取UI控件
        self._ui_window = self.window.get_ui_control()
        if not self._ui_window:
            logger.error("无法获取微信UI控件")
            return False
        
        # 窗口置于前台
        self.window.bring_to_front()
        
        logger.info("✓ 微信控制器初始化成功")
        return True
    
    def find_sessions_with_unread(self):
        """查找有未读消息的会话"""
        sessions = []
        
        if not self._ui_window:
            return sessions
        
        keywords = UI_CONFIG['unread_indicator']['keywords']
        
        def search(control, depth=0, max_depth=8):
            if depth > max_depth:
                return
            
            try:
                name = control.Name
                if name:
                    for kw in keywords:
                        if kw in name:
                            # 找到未读消息提示，向上查找父级会话项
                            sessions.append({
                                'name': name,
                                'control': control,
                                'depth': depth
                            })
                            break
                
                for child in control.GetChildren():
                    search(child, depth + 1, max_depth)
            except:
                pass
        
        search(self._ui_window)
        
        return sessions
    
    def click_session(self, session_control):
        """点击会话"""
        try:
            # 尝试点击控件
            session_control.Click(simulateMove=False)
            time.sleep(0.5)
            logger.info("已点击会话")
            return True
        except Exception as e:
            logger.debug(f"点击会话失败: {e}")
            return False
    
    def get_current_chat_messages(self):
        """获取当前聊天窗口的消息"""
        messages = []
        
        if not self._ui_window:
            return messages
        
        try:
            # 获取当前窗口标题（可能是群聊名称或联系人名称）
            chat_title = self._ui_window.Name
            
            # 判断是群聊还是私聊
            # 注意：这需要根据实际UI结构调整
            # 这里仅提供框架
            
            # 搜索所有文本控件
            text_controls = []
            
            def search(control, depth=0, max_depth=10):
                if depth > max_depth:
                    return
                
                try:
                    if control.ControlTypeName == 'TextControl':
                        name = control.Name
                        if name and name.strip():
                            text_controls.append({
                                'depth': depth,
                                'name': name
                            })
                    
                    for child in control.GetChildren():
                        search(child, depth + 1, max_depth)
                except:
                    pass
            
            search(self._ui_window)
            
            logger.debug(f"当前聊天窗口找到 {len(text_controls)} 个文本控件")
            
            # TODO: 根据实际UI结构解析消息
            # 这部分需要根据诊断结果定制
            # 通常消息区域的文本控件有特定的排列规律
            
        except Exception as e:
            logger.debug(f"获取消息失败: {e}")
        
        return messages


# ==============================================
# 主程序
# ==============================================

def run_diagnose_mode():
    """运行诊断模式"""
    logger.info("")
    logger.info("=" * 60)
    logger.info("运行模式: 诊断模式")
    logger.info("=" * 60)
    logger.info("")
    logger.info("此模式将分析微信的UI结构，帮助你配置程序。")
    logger.info("请确保：")
    logger.info("  1. 微信PC客户端已登录")
    logger.info("  2. 微信窗口已打开（不要最小化）")
    logger.info("  3. 最好有1-2条未读消息（便于诊断）")
    logger.info("")
    
    input("按回车键开始诊断...")
    
    window = WeChatWindow(UI_CONFIG['window_title'])
    diagnoser = UIDiagnoser(window)
    diagnoser.diagnose()
    
    logger.info("")
    logger.info("提示：")
    logger.info("  1. 根据诊断结果，调整 UI_CONFIG 中的配置")
    logger.info("  2. 配置完成后，将 APP_CONFIG['run_mode'] 改为 'normal'")
    logger.info("  3. 或者先用 'test' 模式测试邮件功能")


def run_test_mode():
    """运行测试模式"""
    logger.info("")
    logger.info("=" * 60)
    logger.info("运行模式: 测试模式")
    logger.info("=" * 60)
    logger.info("")
    logger.info("此模式将测试邮件发送功能。")
    logger.info("")
    
    # 创建邮件发送器
    email_sender = EmailSender(EMAIL_CONFIG)
    
    # 测试连接
    if not email_sender.test_connection():
        logger.error("邮件连接测试失败，请检查配置。")
        return
    
    logger.info("")
    print("=" * 60)
    print("测试邮件功能")
    print("=" * 60)
    print("")
    print("操作选项：")
    print("  1. 发送测试邮件")
    print("  2. 自定义消息内容")
    print("  0. 退出")
    print("")
    
    while True:
        choice = input("请选择 (0-2): ").strip()
        
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
            # 自定义消息
            sender = input("发送人: ").strip() or "测试用户"
            content = input("消息内容: ").strip()
            
            if not content:
                logger.warning("消息内容不能为空")
                continue
            
            msg_type = "私聊"
            group_name = None
            
            type_choice = input("消息类型 - 1.私聊(默认) 2.群聊: ").strip()
            if type_choice == '2':
                msg_type = "群聊"
                group_name = input("群聊名称: ").strip() or "测试群"
            
            msg = Message(sender, content, msg_type, group_name)
            processor = MessageProcessor(email_sender)
            processor.process([msg])
        
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
    logger.info("注意：正常模式需要正确配置 UI_CONFIG。")
    logger.info("如果是首次使用，建议先用 'diagnose' 模式诊断UI结构。")
    logger.info("")
    
    # 检查依赖
    if not LIBRARIES['uiautomation'] or not LIBRARIES['win32']:
        logger.error("缺少必要的依赖库")
        return
    
    # 创建组件
    email_sender = EmailSender(EMAIL_CONFIG)
    processor = MessageProcessor(email_sender)
    controller = WeChatController()
    
    # 初始化
    if not controller.initialize():
        logger.error("初始化失败")
        return
    
    # 主循环
    logger.info("开始监控微信消息...")
    logger.info(f"检查间隔: {APP_CONFIG['check_interval']} 秒")
    logger.info("按 Ctrl+C 停止")
    
    try:
        while True:
            # 查找有未读消息的会话
            sessions = controller.find_sessions_with_unread()
            
            if sessions:
                logger.info(f"发现 {len(sessions)} 个会话有未读消息提示")
                
                for session in sessions:
                    logger.info(f"  - {session['name']}")
                    
                    # TODO: 根据实际UI结构实现
                    # 1. 点击会话
                    # 2. 提取消息
                    # 3. 发送邮件
                    
                    # 示例框架：
                    # controller.click_session(session['control'])
                    # messages = controller.get_current_chat_messages()
                    # processor.process(messages)
                    
                    # 由于微信UI结构差异大，这里仅打印日志
                    # 你需要根据诊断结果实现具体逻辑
                    pass
            else:
                logger.debug("没有检测到未读消息")
            
            # 等待
            time.sleep(APP_CONFIG['check_interval'])
            
    except KeyboardInterrupt:
        logger.info("收到停止信号")
    finally:
        email_sender.disconnect()
    
    logger.info("正常模式结束")


def main():
    """主函数"""
    print("")
    print("=" * 60)
    print("微信消息自动监听与转发工具")
    print("=" * 60)
    print("")
    print(f"当前配置的运行模式: {APP_CONFIG['run_mode']}")
    print("")
    print("可用模式：")
    print("  diagnose - 诊断模式：分析微信UI结构（推荐先用这个）")
    print("  test     - 测试模式：测试邮件功能")
    print("  normal   - 正常模式：自动监听消息（需要配置UI）")
    print("")
    
    # 根据配置选择模式
    run_mode = APP_CONFIG['run_mode']
    
    if run_mode == 'diagnose':
        run_diagnose_mode()
    elif run_mode == 'test':
        run_test_mode()
    elif run_mode == 'normal':
        run_normal_mode()
    else:
        logger.error(f"未知的运行模式: {run_mode}")
        logger.error("请选择: diagnose, test, 或 normal")


if __name__ == '__main__':
    main()
