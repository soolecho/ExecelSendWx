from wxauto4 import WeChat
import logging
import time

logger = logging.getLogger(__name__)


class WeChatSender:
    def __init__(self):
        self.wx = None

    def initialize(self):
        logger.info("Initializing WeChat client...")
        try:
            self.wx = WeChat(ads=False)
            logger.info(f"WeChat client initialized successfully: {self.wx.nickname}")
            return True
        except Exception as e:
            logger.error(f"Failed to initialize WeChat: {e}")
            return False

    def send_message(self, content, recipient):
        if not self.wx:
            self.log(f"初始化微信客户端...")
            if not self.initialize():
                self.log(f"❌ 微信初始化失败")
                return False
        
        self.log(f"发送消息给 {recipient}")
        try:
            self.wx.SendMsg(content, who=recipient)
            self.log(f"✅ 消息发送成功")
            return True
        except Exception as e:
            self.log(f"❌ 发送消息失败: {e}")
            return False

    def send_file(self, file_path, recipient):
        if not self.wx:
            if not self.initialize():
                return False
        
        self.log(f"发送文件 {file_path} 给 {recipient}")
        try:
            self.wx.ChatWith(recipient)
            
            time.sleep(1)
            
            chatinfo = self.wx.ChatInfo()
            chat_name = chatinfo.get('chat_name', '') if chatinfo else ''
            if chat_name != recipient:
                self.log(f"❌ 无法切换到聊天窗口: {recipient}")
                return False
            
            self.wx.SendFiles(file_path)
            self.log(f"✅ 文件发送成功")
            return True
        except Exception as e:
            self.log(f"❌ 发送文件失败: {e}")
            return False

    def send_multiple_messages(self, messages, recipient):
        success_count = 0
        for i, message in enumerate(messages):
            self.log(f"发送消息 {i+1}/{len(messages)} 给 {recipient}")
            if self.send_message(message, recipient):
                success_count += 1
            time.sleep(0.5)
        
        self.log(f"成功发送 {success_count}/{len(messages)} 条消息")
        return success_count == len(messages)

    def get_chat_list(self):
        if not self.wx:
            if not self.initialize():
                return []
        
        try:
            sessions = self.wx.GetSession()
            return [s.name for s in sessions if hasattr(s, 'name') and s.name]
        except Exception as e:
            self.log(f"❌ 获取会话列表失败: {e}")
            return []

    def is_online(self):
        if not self.wx:
            if not self.initialize():
                return False
        
        try:
            return self.wx.IsOnline()
        except Exception as e:
            self.log(f"❌ 检查在线状态失败: {e}")
            return False

    def log(self, message):
        logger.info(message)