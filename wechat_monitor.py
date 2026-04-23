# -*- coding: utf-8 -*-
"""
微信消息监听与邮件转发脚本

功能说明：
1. 自动检测微信PC客户端的新消息
2. 记录发送人、消息内容、时间
3. 自动转发到指定邮箱

运行模式：
- MODE_TEST: 测试模式，手动输入消息内容，验证邮件功能
- MODE_SIMULATE: 模拟模式，自动生成模拟消息
- MODE_UI_AUTO: UI自动化模式，需要根据实际微信版本配置控件

使用前请先：
1. 修改下方的邮箱配置
2. 运行 wechat_inspector.py 查看微信UI结构（如果使用UI自动化）
3. 建议先用 MODE_TEST 测试邮件功能
"""

import logging
import smtplib
import time
import sys
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

# ==============================================
# 配置区域 - 请根据实际情况修改
# ==============================================

# 运行模式
RUN_MODE = 'MODE_TEST'  # 可选: MODE_TEST, MODE_SIMULATE, MODE_UI_AUTO

# ==============================================
# 邮箱配置（必须修改！）
# ==============================================
EMAIL_CONFIG = {
    # 发件人邮箱地址
    'sender_email': '3138969266@qq.com',
    
    # 邮箱授权码（注意：不是邮箱登录密码！）
    # QQ邮箱：在设置->账户中开启SMTP服务，获取16位授权码
    # 163邮箱：在设置->客户端授权码中开启
    'sender_password': 'ovyktydfyoivdcci',
    
    # 收件人邮箱地址（多个收件人用逗号分隔）
    'receiver_email': '3138969266@qq.com',
    
    # SMTP服务器地址
    # QQ邮箱: smtp.qq.com
    # 163邮箱: smtp.163.com
    # Gmail: smtp.gmail.com
    'smtp_server': 'smtp.qq.com',
    
    # SMTP端口
    # QQ邮箱(SSL): 465
    # 163邮箱(SSL): 465, (非SSL): 25
    'smtp_port': 465,
    
    # 是否使用SSL加密连接
    'use_ssl': True,
}

# ==============================================
# 微信配置
# ==============================================
WECHAT_CONFIG = {
    'window_title': '微信',           # 微信窗口标题
    'check_interval': 3,              # 消息检查间隔（秒）
    'only_new_messages': True,        # 只发送新消息（不重复发送）
}

# ==============================================
# 日志配置
# ==============================================
LOG_CONFIG = {
    'level': logging.INFO,
    'format': '%(asctime)s - %(levelname)s - %(message)s',
    'console': True,
    'file': False,
    'file_path': 'wechat_monitor.log',
}

# ==============================================
# 尝试导入依赖库
# ==============================================
LIBRARIES = {}

try:
    import uiautomation as auto
    LIBRARIES['uiautomation'] = True
except ImportError:
    LIBRARIES['uiautomation'] = False

try:
    import win32gui
    import win32con
    LIBRARIES['win32'] = True
except ImportError:
    LIBRARIES['win32'] = False


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
        self._server = None
        self._connected = False
    
    def test_connection(self):
        """测试SMTP连接"""
        logger.info("=" * 60)
        logger.info("正在测试邮箱连接...")
        logger.info(f"  服务器: {self.config['smtp_server']}:{self.config['smtp_port']}")
        logger.info(f"  发件人: {self.config['sender_email']}")
        logger.info(f"  收件人: {self.config['receiver_email']}")
        logger.info(f"  SSL加密: {'是' if self.config['use_ssl'] else '否'}")
        logger.info("=" * 60)
        
        try:
            if self.config['use_ssl']:
                logger.info("正在建立SSL连接...")
                self._server = smtplib.SMTP_SSL(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=20
                )
            else:
                logger.info("正在建立普通连接...")
                self._server = smtplib.SMTP(
                    self.config['smtp_server'],
                    self.config['smtp_port'],
                    timeout=20
                )
                logger.info("发送EHLO...")
                self._server.ehlo()
            
            logger.info("正在登录邮箱...")
            self._server.login(
                self.config['sender_email'],
                self.config['sender_password']
            )
            
            self._connected = True
            logger.info("✓ 邮箱连接测试成功！")
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            logger.error("✗ 认证失败！")
            logger.error("")
            logger.error("可能的原因：")
            logger.error("  1. 邮箱地址或授权码错误")
            logger.error("  2. 未开启SMTP服务")
            logger.error("  3. 授权码已过期（需要重新生成）")
            logger.error("")
            logger.error("提示：")
            logger.error("  - QQ邮箱授权码是16位，不是QQ密码！")
            logger.error("  - 163邮箱需要在设置中开启'客户端授权码'")
            logger.error(f"  - 错误详情: {str(e)}")
            return False
            
        except smtplib.SMTPConnectError as e:
            logger.error("✗ 连接失败！")
            logger.error("")
            logger.error("可能的原因：")
            logger.error("  1. SMTP服务器地址错误")
            logger.error("  2. 端口号错误")
            logger.error("  3. 网络问题或防火墙拦截")
            logger.error("  4. 杀毒软件阻止")
            logger.error("")
            logger.error("常见邮箱配置：")
            logger.error("  QQ邮箱:  smtp.qq.com:465 (SSL)")
            logger.error("  163邮箱: smtp.163.com:465 (SSL) 或 25 (非SSL)")
            logger.error(f"  错误详情: {str(e)}")
            return False
            
        except Exception as e:
            logger.error(f"✗ 连接出错: {str(e)}")
            import traceback
            logger.debug(traceback.format_exc())
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
            logger.error(f"✗ 邮件发送失败: {str(e)}")
            self._connected = False
            return False


# ==============================================
# 消息处理器
# ==============================================
class MessageProcessor:
    """消息处理器"""
    
    def __init__(self, email_sender):
        self.email_sender = email_sender
        self.processed = set()
        self.max_history = 2000
    
    def create_message(self, sender, content, msg_type="私聊", group_name=None):
        """创建消息对象"""
        now = datetime.now()
        return {
            'sender': sender,
            'content': content,
            'type': msg_type,
            'group_name': group_name,
            'time': now.strftime('%Y-%m-%d %H:%M:%S'),
            'timestamp': now.timestamp()
        }
    
    def _get_message_id(self, msg):
        """生成消息唯一ID"""
        return f"{msg['sender']}|{msg['content'][:100]}|{msg['time'][:19]}"
    
    def is_new(self, msg):
        """检查是否为新消息"""
        msg_id = self._get_message_id(msg)
        return msg_id not in self.processed
    
    def mark_processed(self, msg):
        """标记为已处理"""
        msg_id = self._get_message_id(msg)
        self.processed.add(msg_id)
        
        if len(self.processed) > self.max_history:
            self.processed = set(list(self.processed)[-1000:])
    
    def format_email(self, messages):
        """格式化邮件内容"""
        if not messages:
            return None, None
        
        # 生成主题
        if len(messages) == 1:
            msg = messages[0]
            if msg['type'] == '群聊' and msg['group_name']:
                subject = f"[微信] {msg['group_name']} - {msg['sender']}"
            else:
                subject = f"[微信] {msg['sender']}"
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
            lines.append(f"  类型: {msg['type']}")
            if msg['group_name']:
                lines.append(f"  群聊: {msg['group_name']}")
            lines.append(f"  发送人: {msg['sender']}")
            lines.append(f"  时间: {msg['time']}")
            lines.append(f"  内容:")
            lines.append(f"    {msg['content']}")
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
            if WECHAT_CONFIG['only_new_messages']:
                if self.is_new(msg):
                    new_msgs.append(msg)
                    self.mark_processed(msg)
            else:
                new_msgs.append(msg)
        
        if not new_msgs:
            logger.debug("没有新消息需要处理")
            return
        
        # 显示日志
        for msg in new_msgs:
            if msg['type'] == '群聊' and msg['group_name']:
                preview = msg['content'][:50] + "..." if len(msg['content']) > 50 else msg['content']
                logger.info(f"[群聊] {msg['group_name']} - {msg['sender']}: {preview}")
            else:
                preview = msg['content'][:50] + "..." if len(msg['content']) > 50 else msg['content']
                logger.info(f"[私聊] {msg['sender']}: {preview}")
        
        # 发送邮件
        subject, body = self.format_email(new_msgs)
        if subject and body:
            self.email_sender.send(subject, body)


# ==============================================
# 不同运行模式的实现
# ==============================================

class BaseMonitor:
    """监控器基类"""
    
    def __init__(self, processor):
        self.processor = processor
        self.running = False
    
    def start(self):
        raise NotImplementedError
    
    def stop(self):
        self.running = False


class TestModeMonitor(BaseMonitor):
    """测试模式：手动输入消息"""
    
    def start(self):
        logger.info("=" * 60)
        logger.info("运行模式: 测试模式 (MODE_TEST)")
        logger.info("=" * 60)
        logger.info("说明：你可以手动输入消息内容，程序会将其发送到邮箱")
        logger.info("用途：验证邮件配置是否正确，以及邮件内容格式")
        logger.info("-" * 60)
        
        print("\n" + "="*60)
        print("操作说明：")
        print("="*60)
        print("1. 输入 'demo'   - 发送演示消息（私聊+群聊）")
        print("2. 输入 'test'   - 测试邮件连接")
        print("3. 输入任意文本  - 作为消息内容发送")
        print("4. 输入 'q'      - 退出程序")
        print("="*60)
        
        while True:
            try:
                user_input = input("\n请输入: ").strip()
                
                if not user_input:
                    continue
                
                # 退出
                if user_input.lower() in ['q', 'quit', 'exit']:
                    logger.info("用户请求退出")
                    break
                
                # 测试连接
                if user_input.lower() == 'test':
                    self.processor.email_sender.test_connection()
                    continue
                
                # 演示消息
                if user_input.lower() == 'demo':
                    logger.info("发送演示消息...")
                    
                    # 演示私聊
                    demo_msg1 = self.processor.create_message(
                        sender="张三",
                        content="今天晚上有空吗？一起吃个饭？\n地点：老地方\n时间：晚上7点",
                        msg_type="私聊"
                    )
                    self.processor.process([demo_msg1])
                    
                    # 演示群聊
                    demo_msg2 = self.processor.create_message(
                        sender="李四",
                        content="各位同事，明天的会议时间改到下午3点了，请准时参加。\n\n会议地点：会议室A\n会议主题：项目进度汇报",
                        msg_type="群聊",
                        group_name="项目组群聊"
                    )
                    self.processor.process([demo_msg2])
                    
                    logger.info("演示消息已发送，请检查邮箱！")
                    continue
                
                # 自定义消息
                # 询问发送人
                sender = input(f"发送人 (默认: 测试用户): ").strip()
                if not sender:
                    sender = "测试用户"
                
                # 询问消息类型
                msg_type = "私聊"
                group_name = None
                
                type_input = input("消息类型 - 1.私聊(默认) 2.群聊: ").strip()
                if type_input == '2':
                    msg_type = "群聊"
                    group_name = input("群聊名称 (默认: 测试群): ").strip()
                    if not group_name:
                        group_name = "测试群"
                
                # 创建消息
                msg = self.processor.create_message(
                    sender=sender,
                    content=user_input,
                    msg_type=msg_type,
                    group_name=group_name
                )
                
                # 处理消息
                self.processor.process([msg])
                
            except KeyboardInterrupt:
                logger.info("收到停止信号")
                break
            except EOFError:
                break
        
        logger.info("测试模式结束")


class SimulateModeMonitor(BaseMonitor):
    """模拟模式：自动生成模拟消息"""
    
    def __init__(self, processor, interval=30):
        super().__init__(processor)
        self.interval = interval
        self._message_templates = [
            {"sender": "张三", "content": "今天的会议改到下午3点了，请准时参加。", "type": "私聊"},
            {"sender": "李四", "content": "收到，我会准时参加的。", "type": "群聊", "group": "项目组"},
            {"sender": "王五", "content": "文档我已经发到群里了，请查收。\n\n文档名称：项目方案v2.0.docx\n如有问题请及时反馈。", "type": "群聊", "group": "工作群"},
            {"sender": "赵六", "content": "周末有空一起打球吗？", "type": "私聊"},
            {"sender": "系统通知", "content": "您的快递已送达，请注意查收。\n\n快递公司：顺丰速运\n取件码：8848", "type": "私聊"},
            {"sender": "财务部", "content": "本月报销单已审批通过，款项将在3个工作日内到账。", "type": "群聊", "group": "公司群"},
        ]
        self._template_index = 0
    
    def start(self):
        logger.info("=" * 60)
        logger.info("运行模式: 模拟模式 (MODE_SIMULATE)")
        logger.info("=" * 60)
        logger.info(f"说明：每 {self.interval} 秒自动生成一条模拟消息并发送到邮箱")
        logger.info("用途：持续测试邮件功能和消息处理流程")
        logger.info("-" * 60)
        
        # 发送启动通知
        startup_msg = self.processor.create_message(
            sender="系统",
            content=f"微信监控程序已启动！\n\n运行模式：模拟模式\n消息间隔：{self.interval}秒\n启动时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            msg_type="私聊"
        )
        self.processor.process([startup_msg])
        
        self.running = True
        counter = 0
        
        try:
            while self.running:
                counter += 1
                
                # 从模板中选择消息
                template = self._message_templates[self._template_index % len(self._message_templates)]
                self._template_index += 1
                
                # 创建消息（添加时间戳让内容不同）
                msg = self.processor.create_message(
                    sender=template['sender'],
                    content=f"{template['content']}\n\n[模拟消息 #{counter} - {datetime.now().strftime('%H:%M:%S')}]",
                    msg_type=template['type'],
                    group_name=template.get('group')
                )
                
                # 处理消息
                self.processor.process([msg])
                
                # 等待
                logger.info(f"下一条消息将在 {self.interval} 秒后发送... (按 Ctrl+C 停止)")
                time.sleep(self.interval)
                
        except KeyboardInterrupt:
            logger.info("收到停止信号")
        finally:
            self.stop()
        
        logger.info("模拟模式结束")


class UIAutoModeMonitor(BaseMonitor):
    """UI自动化模式：从微信窗口获取真实消息"""
    
    def __init__(self, processor):
        super().__init__(processor)
        self._wechat_win = None
    
    def _find_wechat_window(self):
        """查找微信窗口"""
        if not LIBRARIES['win32']:
            logger.error("pywin32库未安装")
            return False
        
        # 精确查找
        hwnd = win32gui.FindWindow(None, WECHAT_CONFIG['window_title'])
        
        if hwnd == 0:
            # 模糊查找
            def callback(h, _):
                if win32gui.IsWindowVisible(h):
                    title = win32gui.GetWindowText(h)
                    if '微信' in title:
                        nonlocal hwnd
                        hwnd = h
                return True
            
            win32gui.EnumWindows(callback, None)
        
        if hwnd == 0:
            logger.error("未找到微信窗口")
            logger.error("请确保：")
            logger.error("  1. 微信PC客户端已登录")
            logger.error("  2. 微信窗口已打开（不要最小化到托盘）")
            return False
        
        logger.info(f"找到微信窗口: HWND={hwnd}")
        return True
    
    def _get_ui_control(self):
        """获取UI自动化控件"""
        if not LIBRARIES['uiautomation']:
            logger.error("uiautomation库未安装")
            return None
        
        try:
            win = auto.WindowControl(Name=WECHAT_CONFIG['window_title'])
            if win.Exists(maxSearchSeconds=3):
                return win
        except:
            pass
        
        return None
    
    def start(self):
        logger.info("=" * 60)
        logger.info("运行模式: UI自动化模式 (MODE_UI_AUTO)")
        logger.info("=" * 60)
        logger.info("说明：自动从微信PC客户端获取消息并转发")
        logger.info("-" * 60)
        
        # 检查依赖
        if not LIBRARIES['uiautomation'] or not LIBRARIES['win32']:
            logger.error("错误: 缺少必要的依赖库")
            logger.error("请运行: pip install uiautomation pywin32")
            return
        
        # 查找微信窗口
        if not self._find_wechat_window():
            return
        
        # 获取UI控件
        self._wechat_win = self._get_ui_control()
        if not self._wechat_win:
            logger.error("无法获取微信UI自动化控件")
            return
        
        logger.info("✓ 初始化成功！")
        logger.info("")
        logger.info("重要提示：")
        logger.info("  由于微信PC客户端的UI结构在不同版本中差异很大，")
        logger.info("  UI自动化模式需要根据你的具体微信版本进行定制。")
        logger.info("")
        logger.info("建议操作：")
        logger.info("  1. 先运行 wechat_inspector.py 查看微信的UI结构")
        logger.info("  2. 根据诊断结果调整本文件中的控件查找逻辑")
        logger.info("  3. 先用 MODE_TEST 确认邮件功能正常")
        logger.info("")
        
        # 发送启动通知
        startup_msg = self.processor.create_message(
            sender="系统",
            content=f"微信监控程序已启动！\n\n运行模式：UI自动化模式\n检查间隔：{WECHAT_CONFIG['check_interval']}秒\n启动时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n注意：本模式需要根据微信实际版本调整UI控件查找逻辑。",
            msg_type="私聊"
        )
        self.processor.process([startup_msg])
        
        logger.info("UI自动化模式已启动，正在监控微信消息...")
        logger.info("提示：目前仅演示框架，需要根据实际微信版本完善消息获取逻辑")
        
        self.running = True
        
        try:
            while self.running:
                # TODO: 在这里实现具体的消息获取逻辑
                # 你需要：
                # 1. 扫描会话列表，查找有未读消息的会话
                # 2. 点击会话，打开聊天窗口
                # 3. 从聊天窗口中提取消息内容
                # 4. 创建消息对象并调用 processor.process()
                
                # 示例框架：
                # sessions = self._scan_sessions()
                # for session in sessions:
                #     messages = self._get_session_messages(session)
                #     self.processor.process(messages)
                
                time.sleep(WECHAT_CONFIG['check_interval'])
                
        except KeyboardInterrupt:
            logger.info("收到停止信号")
        finally:
            self.stop()
        
        logger.info("UI自动化模式结束")


# ==============================================
# 主程序
# ==============================================

def show_welcome():
    """显示欢迎信息"""
    print("\n" + "=" * 60)
    print("微信消息监控与邮件转发工具")
    print("=" * 60)
    print("")
    print(f"当前运行模式: {RUN_MODE}")
    print("")
    print("模式说明：")
    print("  MODE_TEST     - 测试模式：手动输入消息，验证邮件配置")
    print("  MODE_SIMULATE - 模拟模式：自动生成模拟消息，持续测试")
    print("  MODE_UI_AUTO  - UI自动化模式：从微信窗口获取真实消息")
    print("")
    print("建议步骤：")
    print("  1. 先用 MODE_TEST 测试邮件配置是否正确")
    print("  2. 确认邮件能收到后，再使用其他模式")
    print("")


def main():
    """主函数"""
    show_welcome()
    
    # 创建邮件发送器
    email_sender = EmailSender(EMAIL_CONFIG)
    
    # 先测试邮件连接
    if RUN_MODE in ['MODE_TEST', 'MODE_SIMULATE']:
        print("-" * 60)
        test_email = input("是否先测试邮件连接? (y/n，默认y): ").strip().lower()
        if test_email in ['', 'y', 'yes']:
            if not email_sender.test_connection():
                print("")
                print("!" * 60)
                print("邮件连接测试失败！")
                print("!" * 60)
                print("")
                print("请检查配置：")
                print(f"  sender_email: {EMAIL_CONFIG['sender_email']}")
                print(f"  smtp_server: {EMAIL_CONFIG['smtp_server']}")
                print(f"  smtp_port: {EMAIL_CONFIG['smtp_port']}")
                print("")
                print("注意：")
                print("  - sender_password 应该是邮箱授权码，不是登录密码！")
                print("  - QQ邮箱授权码是16位，需要在设置中开启SMTP服务")
                print("")
                
                retry = input("是否继续运行? (y/n): ").strip().lower()
                if retry not in ['y', 'yes']:
                    logger.info("用户选择退出")
                    return
            else:
                email_sender.disconnect()
    
    # 创建消息处理器
    processor = MessageProcessor(email_sender)
    
    # 根据运行模式选择监控器
    if RUN_MODE == 'MODE_TEST':
        monitor = TestModeMonitor(processor)
    elif RUN_MODE == 'MODE_SIMULATE':
        monitor = SimulateModeMonitor(processor, interval=30)
    elif RUN_MODE == 'MODE_UI_AUTO':
        monitor = UIAutoModeMonitor(processor)
    else:
        logger.error(f"未知的运行模式: {RUN_MODE}")
        logger.error("请选择: MODE_TEST, MODE_SIMULATE, 或 MODE_UI_AUTO")
        return
    
    # 启动监控
    try:
        monitor.start()
    except Exception as e:
        logger.error(f"程序运行出错: {str(e)}")
        import traceback
        logger.debug(traceback.format_exc())
    finally:
        email_sender.disconnect()


if __name__ == '__main__':
    main()
