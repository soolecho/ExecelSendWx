from wxauto4 import WeChat
import logging
import os
import tempfile
import threading
import time

from PyQt6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter

logger = logging.getLogger(__name__)


class WeChatSender:
    """
    微信发送器。

    稳定性优化（v2）：
    1. 进程级共享同一个 WeChatSender 实例（shared_instance），避免每次任务
       都重新创建 WeChat 句柄、重复查找微信窗口，降低 CPU 与句柄开销。
    2. 当 ChatWith/ChatInfo 抛异常或返回无效值时，判定为「微信窗口句柄失效 / 微信重启 /
       切换账号」，自动调用 reconnect() 重新初始化 WeChat 并再给一次重试机会。
    """

    _shared_lock = threading.Lock()
    _shared_instance: "WeChatSender | None" = None

    def __init__(self):
        self.wx = None
        # 最近一次确认句柄可用的时间戳；超时会先做一次轻量探活
        self._last_healthy_ts = 0.0
        self._own_initialize_called = False

    @classmethod
    def shared_instance(cls) -> "WeChatSender":
        """进程级单例：所有 ScheduleSendWorker / SendWorker 共用同一份 WeChat 句柄。"""
        with cls._shared_lock:
            if cls._shared_instance is None:
                cls._shared_instance = cls()
            return cls._shared_instance

    @classmethod
    def reset_shared_instance(cls) -> None:
        """极少数情况下（例如用户要求彻底重启）用于强制回收单例。"""
        with cls._shared_lock:
            cls._shared_instance = None

    @staticmethod
    def _is_target_chat(current_chat, recipient):
        current_chat = str(current_chat or "").strip()
        recipient = str(recipient or "").strip()
        return bool(recipient and recipient in current_chat)

    def initialize(self):
        logger.info("Initializing WeChat client...")
        try:
            self.wx = WeChat(ads=False)
            # 读一次 nickname 确认句柄真实可用
            _ = getattr(self.wx, "nickname", None)
            self._last_healthy_ts = time.time()
            self._own_initialize_called = True
            logger.info("WeChat client initialized successfully.")
            return True
        except Exception as e:
            logger.error(f"Failed to initialize WeChat: {e}")
            self.wx = None
            self._own_initialize_called = False
            return False

    def reconnect(self, log_fn=None) -> bool:
        """当检测到微信句柄失效时重新初始化。失败会写日志但不抛异常。"""
        try:
            # 解除旧引用，便于 GC 回收句柄相关资源
            self.wx = None
            if log_fn:
                try:
                    log_fn("🔄 微信句柄失效，正在重新连接微信...")
                except Exception:
                    pass
            ok = self.initialize()
            if ok and log_fn:
                try:
                    log_fn("✅ 微信重连成功")
                except Exception:
                    pass
            return ok
        except Exception as exc:
            logger.error(f"Reconnect failed: {exc}")
            if log_fn:
                try:
                    log_fn(f"❌ 微信重连失败: {exc}")
                except Exception:
                    pass
            return False

    def _safe_chatinfo(self) -> tuple[bool, dict | None]:
        """
        安全读取 ChatInfo。
        返回 (is_healthy, chatinfo)。
        is_healthy=False 意味着句柄已失效，调用方应主动 reconnect 再试一次。
        """
        if not self.wx:
            return False, None
        try:
            chatinfo = self.wx.ChatInfo()
        except Exception:
            return False, None
        # wxauto4 正常时会返回 dict；异常状态下可能返回 None / {} / '无' 等异常值
        if not isinstance(chatinfo, dict):
            return False, None
        # chat_name 不存在或者为空字符串，通常也表明当前无法读取会话，需要重连
        name = chatinfo.get("chat_name")
        # 注意：刚打开微信、没有任何聊天被选中时，chat_name 可能是空，这并不意味着句柄坏，
        # 所以这里只判断"调用没抛异常且是 dict"就算健康，chat_name 缺省由上层再处理。
        _ = name
        self._last_healthy_ts = time.time()
        return True, chatinfo

    def _safe_chatwith(self, recipient: str, exact: bool = False) -> bool:
        if not self.wx:
            return False
        try:
            self.wx.ChatWith(recipient, exact=exact)
            return True
        except Exception:
            return False

    @staticmethod
    def _stop_requested(stop_event):
        return bool(stop_event and stop_event.is_set())

    @staticmethod
    def _wait_or_stopped(delay, stop_event):
        if not stop_event:
            time.sleep(delay)
            return False
        return stop_event.wait(delay)

    def send_message(
        self,
        content,
        recipient,
        first_send=False,
        chat_delay=0.3,
        fast_mode=False,
        stop_event=None
    ):
        if self._stop_requested(stop_event):
            return False

        if not self.wx:
            self.log(f"初始化微信客户端...")
            if not self.initialize():
                self.log(f"❌ 微信初始化失败")
                return False

        # 轻量探活：距离上次健康超过 60 秒，做一次 ChatInfo 预检
        now = time.time()
        if now - self._last_healthy_ts > 60.0:
            ok, _ = self._safe_chatinfo()
            if not ok and not self.reconnect(self.log):
                self.log("❌ 微信探活失败且重连未成功")
                return False

        self.log(f"发送消息给 {recipient}")

        def attempt_once(allow_reconnect: bool) -> bool:
            try:
                if first_send:
                    self.log(f"首次发送，确保微信窗口激活...")
                    if self._wait_or_stopped(0.5, stop_event):
                        return False

                if fast_mode:
                    if self._stop_requested(stop_event):
                        return False
                    ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        if allow_reconnect and self.reconnect(self.log):
                            ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        return False
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
                    if self._stop_requested(stop_event):
                        return False

                    self.log(f"尝试切换窗口 ({attempt+1}/{max_retries}): {recipient}")
                    switched = self._safe_chatwith(recipient, exact=False)
                    if not switched:
                        if allow_reconnect and self.reconnect(self.log):
                            switched = self._safe_chatwith(recipient, exact=False)
                    if not switched:
                        # 切换直接失败（句柄坏）
                        if attempt < max_retries - 1:
                            if self._wait_or_stopped(0.2, stop_event):
                                return False
                        continue
                    if self._wait_or_stopped(chat_delay, stop_event):
                        return False

                    ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        if allow_reconnect and self.reconnect(self.log):
                            # 重连后重新 ChatWith 再读一次
                            if self._safe_chatwith(recipient, exact=False):
                                self._wait_or_stopped(chat_delay, stop_event)
                                ok, chatinfo = self._safe_chatinfo()
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
                            if self._wait_or_stopped(0.2, stop_event):
                                return False

                self.log(f"❌ 窗口切换失败，当前: {current_chat}，目标: {recipient}")
                return False
            except Exception as e:
                self.log(f"❌ 发送消息失败: {e}")
                return False

        # 第一次尝试：失败若是句柄异常，会在内部 reconnect 一次；
        # 第二次再失败就不再重连，避免"微信没登录"情况下无限重连。
        current_chat = ""
        ok = attempt_once(allow_reconnect=True)
        if ok:
            return True
        return attempt_once(allow_reconnect=False)

    def send_file(
        self,
        file_path,
        recipient,
        chat_delay=0.3,
        fast_mode=False,
        stop_event=None
    ):
        if self._stop_requested(stop_event):
            return False

        if not self.wx:
            if not self.initialize():
                return False

        # 探活
        if time.time() - self._last_healthy_ts > 60.0:
            ok, _ = self._safe_chatinfo()
            if not ok and not self.reconnect(self.log):
                return False

        self.log(f"发送文件 {file_path} 给 {recipient}")

        def attempt_once(allow_reconnect: bool) -> bool:
            try:
                if fast_mode:
                    if self._stop_requested(stop_event):
                        return False
                    ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        if allow_reconnect and self.reconnect(self.log):
                            ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        return False
                    current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                    if self._is_target_chat(current_chat, recipient):
                        self.wx.SendFiles(file_path)
                        time.sleep(0.3)
                        self.log(f"✅ 图片发送成功")
                        return True

                max_retries = 3
                current_chat = ""
                for attempt in range(max_retries):
                    if self._stop_requested(stop_event):
                        return False

                    self.log(f"尝试切换窗口 ({attempt + 1}/{max_retries}): {recipient}")
                    switched = self._safe_chatwith(recipient, exact=False)
                    if not switched:
                        if allow_reconnect and self.reconnect(self.log):
                            switched = self._safe_chatwith(recipient, exact=False)
                    if not switched:
                        if attempt < max_retries - 1:
                            if self._wait_or_stopped(0.2, stop_event):
                                return False
                        continue
                    if self._wait_or_stopped(chat_delay, stop_event):
                        return False

                    ok, chatinfo = self._safe_chatinfo()
                    if not ok:
                        if allow_reconnect and self.reconnect(self.log):
                            if self._safe_chatwith(recipient, exact=False):
                                self._wait_or_stopped(chat_delay, stop_event)
                                ok, chatinfo = self._safe_chatinfo()
                    current_chat = chatinfo.get('chat_name', '') if chatinfo else ''
                    self.log(f"当前窗口: {current_chat}")
                    if self._is_target_chat(current_chat, recipient):
                        self.wx.SendFiles(file_path)
                        time.sleep(0.3)
                        self.log(f"✅ 图片发送成功")
                        return True

                    if attempt < max_retries - 1:
                        self.log(f"窗口切换失败，尝试重新搜索 ({attempt + 1}/{max_retries})")
                        if self._wait_or_stopped(0.2, stop_event):
                            return False

                self.log(f"❌ 窗口切换失败，当前: {current_chat}，目标: {recipient}")
                return False
            except Exception as e:
                self.log(f"❌ 发送文件失败: {e}")
                return False

        if attempt_once(allow_reconnect=True):
            return True
        return attempt_once(allow_reconnect=False)

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

    def send_table_images_progress(
        self,
        table_data,
        recipient,
        chat_delay=0.3,
        start_index=0,
        stop_event=None
    ):
        image_paths = []
        next_index = start_index
        try:
            image_paths = self._create_table_images(table_data)
            if start_index >= len(image_paths):
                return True, len(image_paths)

            self.log(f"临时图片目录: {self.get_temp_image_dir()}")
            for index in range(start_index, len(image_paths)):
                if self._stop_requested(stop_event):
                    self.log("⏹ 已停止图片发送")
                    return False, next_index

                image_path = image_paths[index]
                self.log(f"发送图片 {index + 1}/{len(image_paths)} 给 {recipient}")
                if not self.send_file(
                    image_path,
                    recipient,
                    chat_delay=chat_delay,
                    fast_mode=(index > 0),
                    stop_event=stop_event
                ):
                    return False, next_index
                next_index = index + 1
            return True, next_index
        except Exception as e:
            self.log(f"❌ 生成或发送表格图片失败: {e}")
            return False, next_index
        finally:
            for image_path in image_paths:
                self._remove_temp_image(image_path)

    def send_table_images(
        self,
        table_data,
        recipient,
        chat_delay=0.3,
        stop_event=None
    ):
        success, _ = self.send_table_images_progress(
            table_data,
            recipient,
            chat_delay=chat_delay,
            stop_event=stop_event
        )
        return success

    def send_multiple_messages_progress(
        self,
        messages,
        recipient,
        chat_delay=0.2,
        start_index=0,
        stop_event=None
    ):
        next_index = start_index
        for i in range(start_index, len(messages)):
            if self._stop_requested(stop_event):
                return False, next_index

            message = messages[i]
            self.log(f"发送消息 {i+1}/{len(messages)} 给 {recipient}")
            if not self.send_message(
                message,
                recipient,
                first_send=(i == 0),
                chat_delay=chat_delay,
                stop_event=stop_event
            ):
                return False, next_index

            next_index = i + 1
            if i < len(messages) - 1 and self._wait_or_stopped(0.5, stop_event):
                return False, next_index

        self.log(f"成功发送 {len(messages)}/{len(messages)} 条消息")
        return True, next_index
        
    def send_multiple_messages(
        self,
        messages,
        recipient,
        chat_delay=0.2,
        stop_event=None
    ):
        success, _ = self.send_multiple_messages_progress(
            messages,
            recipient,
            chat_delay=chat_delay,
            stop_event=stop_event
        )
        return success

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
