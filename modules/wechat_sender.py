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

    def send_message(self, content, recipient, first_send=False, chat_delay=0.3, fast_mode=False):
        if not self.wx:
            self.log(f"初始化微信客户端...")
            if not self.initialize():
                self.log(f"❌ 微信初始化失败")
                return False
        
        self.log(f"发送消息给 {recipient}")
        try:
            if first_send:
                self.log(f"首次发送，确保微信窗口激活...")
                time.sleep(0.5)
            
            if fast_mode:
                chatinfo = self.wx.ChatInfo()
                current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                if current_chat >= recipient:
                    self.log(f"当前窗口正确，直接发送")
                    self.wx.SendMsg(content)
                    
                    message_length = len(content)
                    if message_length > 500:
                        time.sleep(1)
                    elif message_length > 100:
                        time.sleep(0.5)
                    else:
                        time.sleep(0.2)
                    
                    self.log(f"✅ 快速发送成功")
                    return True
                else:
                    self.log(f"当前窗口不正确({current_chat})，需要重新切换")
            
            self.wx.ChatWith(recipient, exact=False)
            time.sleep(chat_delay)
            
            chatinfo = self.wx.ChatInfo()
            current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
            self.log(f"当前窗口: {current_chat}")
            
            if current_chat >= recipient:
                self.wx.SendMsg(content)
                
                message_length = len(content)
                if message_length > 500:
                    self.log(f"消息较长({message_length}字符)，等待发送完成...")
                    time.sleep(1)
                elif message_length > 100:
                    time.sleep(0.5)
                else:
                    time.sleep(0.2)
                
                self.log(f"✅ 消息发送成功")
                return True
            else:
                self.log(f"❌ 窗口切换失败，当前: {current_chat}，目标: {recipient}")
                return False
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

    def send_multiple_messages(self, messages, recipient, chat_delay=0.2):
        success_count = 0
        for i, message in enumerate(messages):
            self.log(f"发送消息 {i+1}/{len(messages)} 给 {recipient}")
            if self.send_message(message, recipient, first_send=(i == 0), chat_delay=chat_delay):
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
