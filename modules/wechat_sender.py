from wxauto4 import WeChat
import logging
import os
import tempfile
import time

from PyQt6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter

logger = logging.getLogger(__name__)


class WeChatSender:
    def __init__(self):
        self.wx = None

    @staticmethod
    def _is_target_chat(current_chat, recipient):
        current_chat = str(current_chat or "").strip()
        recipient = str(recipient or "").strip()
        return bool(recipient and recipient in current_chat)

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
                if self._is_target_chat(current_chat, recipient):
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
            
            max_retries = 3
            for attempt in range(max_retries):
                self.log(f"尝试切换窗口 ({attempt+1}/{max_retries}): {recipient}")
                self.wx.ChatWith(recipient, exact=False)
                time.sleep(chat_delay)
                
                chatinfo = self.wx.ChatInfo()
                current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                self.log(f"当前窗口: {current_chat}")
                
                if self._is_target_chat(current_chat, recipient):
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
                    if attempt < max_retries - 1:
                        self.log(f"窗口切换失败，尝试重新搜索 ({attempt+1}/{max_retries})")
                        time.sleep(0.2)
            
            self.log(f"❌ 窗口切换失败，当前: {current_chat}，目标: {recipient}")
            return False
        except Exception as e:
            self.log(f"❌ 发送消息失败: {e}")
            return False

    def send_file(self, file_path, recipient, chat_delay=0.3, fast_mode=False):
        if not self.wx:
            if not self.initialize():
                return False
        
        self.log(f"发送文件 {file_path} 给 {recipient}")
        try:
            if fast_mode:
                chatinfo = self.wx.ChatInfo()
                current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                if self._is_target_chat(current_chat, recipient):
                    self.wx.SendFiles(file_path)
                    time.sleep(0.3)
                    self.log(f"✅ 图片发送成功")
                    return True

            max_retries = 3
            current_chat = ""
            for attempt in range(max_retries):
                self.log(f"尝试切换窗口 ({attempt + 1}/{max_retries}): {recipient}")
                self.wx.ChatWith(recipient, exact=False)
                time.sleep(chat_delay)

                chatinfo = self.wx.ChatInfo()
                current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                self.log(f"当前窗口: {current_chat}")
                if self._is_target_chat(current_chat, recipient):
                    self.wx.SendFiles(file_path)
                    time.sleep(0.3)
                    self.log(f"✅ 图片发送成功")
                    return True

                if attempt < max_retries - 1:
                    self.log(f"窗口切换失败，尝试重新搜索 ({attempt + 1}/{max_retries})")
                    time.sleep(0.2)

            self.log(f"❌ 窗口切换失败，当前: {current_chat}，目标: {recipient}")
            return False
        except Exception as e:
            self.log(f"❌ 发送文件失败: {e}")
            return False

    @classmethod
    def get_temp_image_dir(cls):
        return os.path.join(tempfile.gettempdir(), "wxauto_images")

    def _remove_temp_image(self, image_path):
        if not image_path or not os.path.exists(image_path):
            return True

        for _ in range(3):
            try:
                os.remove(image_path)
                self.log(f"已删除临时图片: {os.path.basename(image_path)}")
                return True
            except PermissionError:
                time.sleep(0.2)
            except OSError as e:
                self.log(f"⚠ 临时图片删除失败: {e}")
                return False

        self.log(f"⚠ 临时图片正在被占用，暂时无法删除: {image_path}")
        return False

    def cleanup_temp_images(self):
        temp_dir = self.get_temp_image_dir()
        if not os.path.isdir(temp_dir):
            return

        removed_count = 0
        for file_name in os.listdir(temp_dir):
            if not file_name.startswith("wxauto_table_") or not file_name.endswith(".png"):
                continue
            if self._remove_temp_image(os.path.join(temp_dir, file_name)):
                removed_count += 1

        if removed_count:
            self.log(f"已清理 {removed_count} 张临时图片")

    @staticmethod
    def _wrap_cell_text(value, metrics, max_width):
        wrapped_lines = []
        text = "" if value is None else str(value)
        for source_line in text.splitlines() or [""]:
            if not source_line:
                wrapped_lines.append("")
                continue

            current_line = ""
            for char in source_line:
                candidate = current_line + char
                if current_line and metrics.horizontalAdvance(candidate) > max_width:
                    wrapped_lines.append(current_line)
                    current_line = char
                else:
                    current_line = candidate
            wrapped_lines.append(current_line)
        return wrapped_lines or [""]

    def _create_table_images(self, table_data):
        headers = [
            "" if value is None else str(value)
            for value in table_data.get("headers", [])
        ]
        rows = table_data.get("rows", [])
        if not headers:
            raise ValueError("图片数据缺少表头")

        column_count = len(headers)
        normalized_rows = []
        for row in rows:
            values = [
                "" if value is None else str(value)
                for value in list(row)[:column_count]
            ]
            values.extend([""] * (column_count - len(values)))
            normalized_rows.append(values)

        font = QFont("Microsoft YaHei")
        font.setPixelSize(20)
        header_font = QFont(font)
        header_font.setBold(True)
        metrics = QFontMetrics(font)
        header_metrics = QFontMetrics(header_font)

        cell_padding = 12
        line_height = metrics.height() + 4
        header_line_height = header_metrics.height() + 4
        min_column_width = 110
        max_column_width = 360
        max_table_width = 1750

        column_widths = []
        for column_index, header in enumerate(headers):
            values = [header]
            values.extend(row[column_index] for row in normalized_rows)
            widest = max(
                metrics.horizontalAdvance(line)
                for value in values
                for line in (str(value).splitlines() or [""])
            )
            column_widths.append(
                max(
                    min_column_width,
                    min(max_column_width, widest + cell_padding * 2)
                )
            )

        if sum(column_widths) > max_table_width:
            scale = max_table_width / sum(column_widths)
            column_widths = [
                max(90, int(width * scale))
                for width in column_widths
            ]

        header_cells = [
            self._wrap_cell_text(
                header,
                header_metrics,
                column_widths[index] - cell_padding * 2
            )
            for index, header in enumerate(headers)
        ]
        header_height = max(
            52,
            max(len(lines) for lines in header_cells) * header_line_height
            + cell_padding * 2
        )

        prepared_rows = []
        for row in normalized_rows:
            cells = [
                self._wrap_cell_text(
                    value,
                    metrics,
                    column_widths[index] - cell_padding * 2
                )
                for index, value in enumerate(row)
            ]
            row_height = max(
                48,
                max(len(lines) for lines in cells) * line_height
                + cell_padding * 2
            )
            prepared_rows.append((cells, row_height))

        page_margin = 24
        max_image_height = 3600
        pages = []
        current_page = []
        current_height = page_margin * 2 + header_height
        for prepared_row in prepared_rows:
            row_height = prepared_row[1]
            if current_page and current_height + row_height > max_image_height:
                pages.append(current_page)
                current_page = []
                current_height = page_margin * 2 + header_height
            current_page.append(prepared_row)
            current_height += row_height
        pages.append(current_page)

        temp_dir = self.get_temp_image_dir()
        os.makedirs(temp_dir, exist_ok=True)
        image_paths = []
        try:
            for page in pages:
                image_width = page_margin * 2 + sum(column_widths)
                image_height = (
                    page_margin * 2
                    + header_height
                    + sum(row_height for _, row_height in page)
                )
                image = QImage(
                    image_width,
                    image_height,
                    QImage.Format.Format_ARGB32
                )
                image.fill(QColor("#FFFFFF"))

                painter = QPainter(image)
                painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
                try:
                    y = page_margin
                    self._draw_table_row(
                        painter,
                        page_margin,
                        y,
                        column_widths,
                        header_cells,
                        header_height,
                        header_font,
                        header_metrics,
                        header_line_height,
                        QColor("#DCE6F1"),
                        cell_padding
                    )
                    y += header_height

                    for row_index, (cells, row_height) in enumerate(page):
                        background = (
                            QColor("#FFFFFF")
                            if row_index % 2 == 0
                            else QColor("#F7F9FC")
                        )
                        self._draw_table_row(
                            painter,
                            page_margin,
                            y,
                            column_widths,
                            cells,
                            row_height,
                            font,
                            metrics,
                            line_height,
                            background,
                            cell_padding
                        )
                        y += row_height
                finally:
                    painter.end()

                fd, image_path = tempfile.mkstemp(
                    prefix="wxauto_table_",
                    suffix=".png",
                    dir=temp_dir
                )
                os.close(fd)
                if not image.save(image_path, "PNG"):
                    self._remove_temp_image(image_path)
                    raise RuntimeError("表格图片保存失败")
                image_paths.append(image_path)
        except Exception:
            for image_path in image_paths:
                self._remove_temp_image(image_path)
            raise

        return image_paths

    @staticmethod
    def _draw_table_row(
        painter,
        start_x,
        y,
        column_widths,
        cells,
        row_height,
        font,
        metrics,
        line_height,
        background,
        cell_padding
    ):
        x = start_x
        painter.setFont(font)
        for column_index, cell_lines in enumerate(cells):
            column_width = column_widths[column_index]
            painter.fillRect(x, y, column_width, row_height, background)
            painter.setPen(QColor("#AAB4C0"))
            painter.drawRect(x, y, column_width, row_height)

            painter.setPen(QColor("#202124"))
            text_y = y + cell_padding + metrics.ascent()
            for line in cell_lines:
                painter.drawText(x + cell_padding, text_y, line)
                text_y += line_height
            x += column_width

    def send_table_images(
        self,
        table_data,
        recipient,
        chat_delay=0.3,
        should_stop=None
    ):
        image_paths = []
        try:
            image_paths = self._create_table_images(table_data)
            self.log(f"临时图片目录: {self.get_temp_image_dir()}")
            for index, image_path in enumerate(image_paths):
                if should_stop and should_stop():
                    self.log("⏹ 已停止图片发送")
                    return False

                self.log(f"发送图片 {index + 1}/{len(image_paths)} 给 {recipient}")
                if not self.send_file(
                    image_path,
                    recipient,
                    chat_delay=chat_delay,
                    fast_mode=(index > 0)
                ):
                    return False
            return True
        except Exception as e:
            self.log(f"❌ 生成或发送表格图片失败: {e}")
            return False
        finally:
            for image_path in image_paths:
                self._remove_temp_image(image_path)

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
